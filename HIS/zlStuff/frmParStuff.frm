VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmParStuff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���Ĳ�������"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14040
   Icon            =   "frmParStuff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14040
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabDesign 
      Height          =   8295
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Ŀ¼(&0)"
      TabPicture(0)   =   "frmParStuff.frx":74F2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���(&1)"
      TabPicture(1)   =   "frmParStuff.frx":750E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�ڿ�(&2)"
      TabPicture(2)   =   "frmParStuff.frx":752A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "����(&3)"
      TabPicture(3)   =   "frmParStuff.frx":7546
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "�������(&10)"
      TabPicture(4)   =   "frmParStuff.frx":7562
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(10)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�����(&11)"
      TabPicture(5)   =   "frmParStuff.frx":757E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(11)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "����ⷿ(&12)"
      TabPicture(6)   =   "frmParStuff.frx":759A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picPar(12)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "���ľ���(&13)"
      TabPicture(7)   =   "frmParStuff.frx":75B6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(13)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "����(&14)"
      TabPicture(8)   =   "frmParStuff.frx":75D2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(14)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "ͨ��(&4)"
      TabPicture(9)   =   "frmParStuff.frx":75EE
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "picPar(4)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7200
         Index           =   4
         Left            =   135
         ScaleHeight     =   7170
         ScaleWidth      =   10005
         TabIndex        =   141
         Top             =   450
         Width           =   10035
         Begin VB.Frame fraBarCodeStuff 
            Caption         =   "��������ʶ�����"
            ForeColor       =   &H00800000&
            Height          =   1080
            Left            =   165
            TabIndex        =   142
            Top             =   105
            Width           =   5655
            Begin VB.OptionButton optBarcode 
               Caption         =   "ֻ�������������ɨ�����ʶ��"
               Height          =   240
               Index           =   1
               Left            =   75
               TabIndex        =   144
               Top             =   660
               Width           =   4755
            End
            Begin VB.OptionButton optBarcode 
               Caption         =   "����������롢���롢����Ƚ���ʶ��"
               Height          =   240
               Index           =   0
               Left            =   75
               TabIndex        =   143
               Top             =   345
               Value           =   -1  'True
               Width           =   3720
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   14
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   135
         Top             =   360
         Visible         =   0   'False
         Width           =   10455
         Begin VSFlex8Ctl.VSFlexGrid vsf���ݻ��ڿ��� 
            Height          =   6885
            Left            =   240
            TabIndex        =   136
            Top             =   360
            Width           =   10020
            _cx             =   17674
            _cy             =   12144
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
         Begin VB.Label lbl���ݿ��� 
            AutoSize        =   -1  'True
            Caption         =   "���ݻ��ڿ��ƣ��������ĵ������ض�ҵ�񻷽��������޸ĵ���Ŀ"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   300
            TabIndex        =   137
            Top             =   120
            Width           =   5040
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   13
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   131
         Top             =   360
         Width           =   10455
         Begin ZL9BillEdit.BillEdit BillҩƷ���ľ��� 
            Height          =   6180
            Left            =   240
            TabIndex        =   132
            Top             =   360
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   10901
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   2
            RowHeight0      =   315
            RowHeightMin    =   315
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
         Begin VB.Label lbl����˵�� 
            Caption         =   $"frmParStuff.frx":760A
            ForeColor       =   &H00000080&
            Height          =   720
            Left            =   240
            TabIndex        =   134
            Top             =   6600
            Width           =   7995
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "���ľ������ã�����װ��λ�����ü۸���������¼��ľ��ȣ�������С��λ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   240
            TabIndex        =   133
            Top             =   120
            Width           =   6480
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   0
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   10455
         Begin VB.CheckBox chk 
            Caption         =   "���ϸ��������ָ�����ۺ�ָ���ۼ�"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   44
            Top             =   120
            Width           =   3615
         End
         Begin VB.Frame fra���Ķ��۵�λ 
            Caption         =   " ����Ŀ¼���۵�λ"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   240
            TabIndex        =   41
            Top             =   1920
            Width           =   4740
            Begin VB.OptionButton opt���۵�λ 
               Caption         =   "ɢװ��λ"
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   1185
            End
            Begin VB.OptionButton opt���۵�λ 
               Caption         =   "��װ��λ"
               Height          =   285
               Index           =   1
               Left            =   1560
               TabIndex        =   42
               Top             =   360
               Width           =   1425
            End
         End
         Begin VB.Frame fra�������ģʽ 
            Caption         =   " �������ģʽ"
            ForeColor       =   &H00800000&
            Height          =   1275
            Left            =   240
            TabIndex        =   37
            Top             =   480
            Width           =   4740
            Begin VB.OptionButton opt����ģʽ 
               Caption         =   "�����+˳����"
               Height          =   210
               Index           =   2
               Left            =   240
               TabIndex        =   40
               Top             =   960
               Width           =   3420
            End
            Begin VB.OptionButton opt����ģʽ 
               Caption         =   "�������+�����+˳����"
               Height          =   210
               Index           =   1
               Left            =   240
               TabIndex        =   39
               Top             =   660
               Width           =   3420
            End
            Begin VB.OptionButton opt����ģʽ 
               Caption         =   "ͬ��˳����"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   38
               Top             =   360
               Value           =   -1  'True
               Width           =   2655
            End
         End
         Begin VB.Frame fraIncome 
            Caption         =   " �������Ķ�Ӧȱʡ������Ŀ"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   240
            TabIndex        =   34
            Top             =   2760
            Width           =   4740
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   300
               Width           =   1875
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "��������"
               Height          =   180
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " ���ķ��������Զ�����"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   240
            TabIndex        =   29
            Top             =   3720
            Width           =   4740
            Begin VB.OptionButton opt�������� 
               Caption         =   "�ⷿ�ͷ��ϲ��ŷ���"
               Height          =   210
               Index           =   2
               Left            =   240
               TabIndex        =   33
               Top             =   720
               Width           =   1980
            End
            Begin VB.OptionButton opt�������� 
               Caption         =   "���ⷿ����"
               Height          =   210
               Index           =   1
               Left            =   2280
               TabIndex        =   32
               Top             =   360
               Width           =   1500
            End
            Begin VB.OptionButton opt�������� 
               Caption         =   "�ⷿ�ͷ��ϲ��Ŷ�������"
               Height          =   210
               Index           =   3
               Left            =   2280
               TabIndex        =   31
               Top             =   720
               Width           =   2385
            End
            Begin VB.OptionButton opt�������� 
               Caption         =   "�ֹ����÷�������"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   1740
            End
         End
         Begin VB.Frame fra 
            Caption         =   " ���ô洢�ⷿʱ����Ӧ���ڵķ�Χ"
            ForeColor       =   &H00800000&
            Height          =   2385
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   4920
            Width           =   4785
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ���ڷ�����������������"
               Height          =   255
               Index           =   2
               Left            =   270
               TabIndex        =   26
               Top             =   840
               Width           =   2760
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ���ڱ���������������"
               Height          =   255
               Index           =   1
               Left            =   270
               TabIndex        =   25
               Top             =   562
               Width           =   2712
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ����������������"
               Height          =   255
               Index           =   0
               Left            =   270
               TabIndex        =   24
               Top             =   285
               Width           =   2364
            End
            Begin VB.Label lblInfor 
               Caption         =   "   ��:û�й��ϴ���Ŀ�еġ�Ӧ���������������ϡ������ڴ洢�ⷿ���ý����еġ�Ӧ�������С��������ϡ�(4)��������ѡ��"
               ForeColor       =   &H00000080&
               Height          =   615
               Index           =   2
               Left            =   240
               TabIndex        =   28
               Top             =   1680
               Width           =   4260
            End
            Begin VB.Label lblInfor 
               Caption         =   "    ����Ŀ��Ҫ�ǿ����������Ϲ���Ĵ洢�ⷿ���ý����еġ�Ӧ����...�����ܡ�"
               ForeColor       =   &H00000080&
               Height          =   405
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   1200
               Width           =   4350
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   12
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   21
         Top             =   300
         Width           =   10455
         Begin VSFlex8Ctl.VSFlexGrid vsf���� 
            Height          =   6375
            Left            =   240
            TabIndex        =   116
            Top             =   840
            Width           =   8175
            _cx             =   14420
            _cy             =   11245
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
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParStuff.frx":76EB
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   240
            Picture         =   "frmParStuff.frx":7803
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lbl����ⷿ 
            Caption         =   $"frmParStuff.frx":80CD
            ForeColor       =   &H00000080&
            Height          =   540
            Left            =   840
            TabIndex        =   117
            Top             =   120
            Width           =   6255
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   11
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   20
         Top             =   300
         Width           =   10455
         Begin VSFlex8Ctl.VSFlexGrid vsf�ⷿ��� 
            Height          =   6495
            Left            =   240
            TabIndex        =   114
            Top             =   720
            Width           =   8055
            _cx             =   14208
            _cy             =   11456
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
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParStuff.frx":8190
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
         Begin VB.Label lbl��ʾ 
            Caption         =   "  ���������ѡ����ⷿ�Ƿ����漰����鷽ʽ�����ⷿѡ��ʱ˫�����ⷿ��鷽ʽ���пɸı�ⷿ�ļ�鷽ʽ��"
            ForeColor       =   &H00000080&
            Height          =   435
            Left            =   840
            TabIndex        =   115
            Top             =   195
            Width           =   7080
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   240
            Picture         =   "frmParStuff.frx":824D
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   10
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   19
         Top             =   300
         Width           =   10455
         Begin VSFlex8Ctl.VSFlexGrid vsf���� 
            Height          =   6495
            Left            =   240
            TabIndex        =   112
            Top             =   720
            Width           =   8055
            _cx             =   14208
            _cy             =   11456
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
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParStuff.frx":88CE
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
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Ʋ����ڲ�ͬ�ⷿ�����ͨ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Index           =   23
            Left            =   840
            TabIndex        =   113
            Top             =   270
            Width           =   2700
         End
         Begin VB.Image Image1 
            Height          =   495
            Index           =   0
            Left            =   240
            Picture         =   "frmParStuff.frx":89E9
            Stretch         =   -1  'True
            Top             =   120
            Width           =   435
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   3
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   18
         Top             =   300
         Width           =   10455
         Begin VB.ComboBox cbo 
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   1
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   540
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ�Զ������ʷ�������"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   111
            Top             =   960
            Width           =   2880
         End
         Begin VB.Frame fra 
            Caption         =   " �������ϵ��ݹ��˿��� "
            ForeColor       =   &H00800000&
            Height          =   1425
            Index           =   3
            Left            =   240
            TabIndex        =   101
            Top             =   4800
            Width           =   4455
            Begin VB.CheckBox chkDeptType 
               Caption         =   "Ӫ��"
               Enabled         =   0   'False
               Height          =   255
               Index           =   6
               Left            =   2280
               TabIndex        =   110
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "����"
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   1440
               TabIndex        =   109
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "����"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   600
               TabIndex        =   108
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "����"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   107
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "���"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   106
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "����"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   105
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "�ٴ�"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   104
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chk���� 
               Caption         =   "����������ʱ�����ǲ��˿��ҿ����ļ�¼"
               Height          =   255
               Left            =   240
               TabIndex        =   103
               Top             =   360
               Width           =   3690
            End
            Begin VB.TextBox txt 
               Height          =   375
               Index           =   3
               Left            =   3120
               TabIndex        =   102
               Text            =   "�����ֵ"
               Top             =   960
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "������ǩ��"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   100
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ��������ʱ�����"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   99
            Top             =   1680
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����������շѻ���ʺ��Զ�����"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   98
            Top             =   240
            Width           =   3000
         End
         Begin VB.Frame fraδ�շѷ�ҩ 
            Caption         =   " δ�շѻ����ʱ��������"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   240
            TabIndex        =   93
            Top             =   2880
            Width           =   4455
            Begin VB.CheckBox chk 
               Caption         =   "δ�շѵ����ﻮ�۴�������"
               Height          =   180
               Index           =   27
               Left            =   240
               TabIndex        =   96
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "δ��˵ļ��˴�������"
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   95
               Top             =   1200
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ŀִ��ǰ���շѻ����"
               Height          =   195
               Index           =   26
               Left            =   480
               TabIndex        =   94
               Top             =   0
               Visible         =   0   'False
               Width           =   2880
            End
            Begin VB.Label lblδ�շѷ�ҩ 
               Caption         =   "  �������������һ��ͨ����""ִ��ǰ�������շѻ��ȼ������""��������ﲡ�˷���ʱ�����²�����ʧЧ��"
               ForeColor       =   &H00000080&
               Height          =   615
               Left            =   240
               TabIndex        =   97
               Top             =   240
               Width           =   3855
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Ƿ��Զ�ȱ�ϼ��"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   92
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ�������������¼"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   91
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������סԺ���ʺ��Զ����Ϸ�ʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   139
            Top             =   600
            Width           =   2520
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   2
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   17
         Top             =   300
         Width           =   10455
         Begin VB.Frame fra 
            Caption         =   "���Ľ��"
            ForeColor       =   &H00800000&
            Height          =   1860
            Index           =   10
            Left            =   6000
            TabIndex        =   123
            Top             =   1080
            Width           =   3975
            Begin VB.Frame fra�Զ���淽ʽ 
               Caption         =   " ���ý��ʱ��"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   127
               Top             =   1140
               Width           =   3375
               Begin VB.TextBox txt 
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   4
                  Left            =   825
                  TabIndex        =   128
                  Text            =   "25"
                  Top             =   315
                  Width           =   300
               End
               Begin VB.OptionButton opt���ʱ��ģʽ 
                  Caption         =   "ÿ�����һ��"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   130
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton opt���ʱ��ģʽ 
                  Caption         =   "ÿ��    ��"
                  Height          =   180
                  Index           =   1
                  Left            =   120
                  TabIndex        =   129
                  Top             =   315
                  Width           =   1215
               End
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   5
               Left            =   120
               TabIndex        =   126
               Text            =   "Text1"
               Top             =   1200
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.OptionButton opt��淽ʽ 
               Caption         =   "�Զ����(���ⷿ��ͬһ���ڽ��)"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   125
               Top             =   720
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton opt��淽ʽ 
               Caption         =   "�ֹ����(���ⷿ���Բ�ͬ���ڽ��)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   124
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "ʱ�����İ����ε���"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   90
            Top             =   600
            Width           =   3105
         End
         Begin VB.Frame fra����У�� 
            Caption         =   " �ƻ�������У��"
            ForeColor       =   &H00800000&
            Height          =   6255
            Index           =   1
            Left            =   240
            TabIndex        =   83
            Top             =   1080
            Width           =   5295
            Begin VB.Frame fraCheck 
               Caption         =   "ѡ��У�鷽ʽ"
               ForeColor       =   &H00800000&
               Height          =   615
               Index           =   1
               Left            =   120
               TabIndex        =   85
               Top             =   5520
               Width           =   4935
               Begin VB.OptionButton opt�ƻ�����У�� 
                  Caption         =   "У��δͨ��ʱ��ֹ����"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   87
                  Top             =   280
                  Width           =   2175
               End
               Begin VB.OptionButton opt�ƻ�����У�� 
                  Caption         =   "У��δͨ��ʱ����"
                  Height          =   180
                  Index           =   1
                  Left            =   2520
                  TabIndex        =   86
                  Top             =   280
                  Width           =   1935
               End
            End
            Begin VB.TextBox txt 
               Height          =   375
               Index           =   2
               Left            =   4200
               TabIndex        =   84
               Text            =   "�����ֵ"
               Top             =   5760
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
               Height          =   4485
               Index           =   1
               Left            =   120
               TabIndex        =   88
               Top             =   840
               Width           =   4935
               _cx             =   8705
               _cy             =   7911
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
               BackColorSel    =   16711680
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483633
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483632
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   25
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmParStuff.frx":8CF3
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
               VirtualData     =   0   'False
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
            Begin VB.Label lblComment 
               Caption         =   $"frmParStuff.frx":8F9E
               ForeColor       =   &H00000080&
               Height          =   540
               Index           =   1
               Left            =   120
               TabIndex        =   89
               Top             =   240
               Width           =   4980
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����̵�û�����ô洢�ⷿ������"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   82
            Top             =   240
            Width           =   3105
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   1
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   16
         Top             =   300
         Width           =   10455
         Begin VB.Frame fra���� 
            Caption         =   " ��������Ϣȡֵ��ʽ"
            ForeColor       =   &H00800000&
            Height          =   1470
            Left            =   120
            TabIndex        =   119
            Top             =   1800
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "�����������Ų��ؿ���"
               Height          =   255
               Index           =   34
               Left            =   240
               TabIndex        =   138
               Top             =   1080
               Width           =   2895
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "����ȡĿ¼�еĲ���"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   121
               Top             =   480
               Width           =   2295
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "����ȡ�ϴ����Ĳ���"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   120
               Top             =   240
               Value           =   -1  'True
               Width           =   2295
            End
         End
         Begin VB.Frame fra���� 
            Caption         =   " �������̿���"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   4560
            TabIndex        =   78
            Top             =   4680
            Width           =   5565
            Begin VB.CheckBox chk 
               Caption         =   "���������ǰ��Ҫ���в���˲�"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   81
               Top             =   720
               Width           =   3180
            End
            Begin VB.CheckBox chk 
               Caption         =   "���������""��������""���Ե����Ľ�������"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   80
               Top             =   480
               Width           =   3810
            End
            Begin VB.CheckBox chk 
               Caption         =   "�������ϲ���������������"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame fra�ƿ����̿��� 
            Caption         =   " �ƿ⹦�����̿���"
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   4560
            TabIndex        =   76
            Top             =   5880
            Width           =   5565
            Begin VB.CheckBox chk 
               Caption         =   "�ƿ����ʱ������ⷿ��Ҫ���������"
               Height          =   255
               Index           =   16
               Left            =   180
               TabIndex        =   77
               Top             =   240
               Value           =   1  'Checked
               Width           =   3705
            End
         End
         Begin VB.Frame fra�����㷨 
            Caption         =   " �������ȹ���"
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   4560
            TabIndex        =   73
            Top             =   6600
            Width           =   5565
            Begin VB.OptionButton opt���ĳ����㷨 
               Caption         =   "�������Ƚ��ȳ�"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton opt���ĳ����㷨 
               Caption         =   "��Ч������ȳ�"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   74
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame fra������� 
            Caption         =   " �������"
            ForeColor       =   &H00800000&
            Height          =   1332
            Left            =   120
            TabIndex        =   70
            Top             =   5880
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "�������ƿ���������"
               Height          =   255
               Index           =   31
               Left            =   120
               TabIndex        =   122
               Top             =   480
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "������������������"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   118
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "������������������"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "������¿��ÿ��"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame fra������ 
            Caption         =   " ���۸����"
            ForeColor       =   &H00800000&
            Height          =   1575
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "ʱ�������ԼӼ������"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   69
               Top             =   240
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ�����İ��ֶμӳ������"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   68
               Top             =   480
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��������ⰴ��ǰ�ӳ�����"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   67
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ���������ʱ�����ֹ������ۼ�"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   66
               Top             =   1200
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ���������ʱȡ�ϴ��ۼ�"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   2895
            End
         End
         Begin VB.Frame fra�⹺��� 
            Caption         =   " �⹺����������"
            ForeColor       =   &H00800000&
            Height          =   2415
            Left            =   120
            TabIndex        =   53
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "�⹺��ⵥ��Ҫ�˲�"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   62
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����޸Ĳɹ��޼�"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "�б����Ŀ�ѡ����б굥λ���"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   2880
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ֵ���ı�����д��ϸ��Ϣ"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   59
               Top             =   720
               Width           =   2880
            End
            Begin VB.Frame fraBidMess 
               Caption         =   " �ɹ��۳��б�۸�ʱ"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   55
               Top             =   1680
               Width           =   3315
               Begin VB.OptionButton opt�ɹ��� 
                  Caption         =   "��ֹ"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opt�ɹ��� 
                  Caption         =   "��ʾ"
                  Height          =   180
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   57
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton opt�ɹ��� 
                  Caption         =   "������"
                  Height          =   180
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   56
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   0
               Left            =   2760
               MaxLength       =   8
               TabIndex        =   54
               Top             =   1260
               Width           =   945
            End
            Begin VB.Label lbl����ǰ׺��ʾ 
               AutoSize        =   -1  'True
               Caption         =   "��������ǰ׺(2-8λ���ֻ���ĸ)"
               Height          =   180
               Left            =   120
               TabIndex        =   63
               Top             =   1320
               Width           =   2610
            End
         End
         Begin VB.Frame fra����У�� 
            Caption         =   " �⹺�������У��"
            ForeColor       =   &H00800000&
            Height          =   4455
            Index           =   0
            Left            =   4560
            TabIndex        =   45
            Top             =   120
            Width           =   5535
            Begin VB.CheckBox chk 
               Caption         =   "�������ڴ���ע��֤Ч�ڼ��"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   50
               Top             =   4080
               Width           =   2775
            End
            Begin VB.Frame fraCheck 
               Caption         =   "ѡ��У�鷽ʽ"
               ForeColor       =   &H00800000&
               Height          =   615
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   3360
               Width           =   4935
               Begin VB.OptionButton opt�⹺����У�� 
                  Caption         =   "У��δͨ��ʱ����"
                  Height          =   180
                  Index           =   1
                  Left            =   2520
                  TabIndex        =   49
                  Top             =   280
                  Width           =   1935
               End
               Begin VB.OptionButton opt�⹺����У�� 
                  Caption         =   "У��δͨ��ʱ��ֹ����"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   48
                  Top             =   280
                  Width           =   2175
               End
            End
            Begin VB.TextBox txt 
               Height          =   375
               Index           =   1
               Left            =   3360
               TabIndex        =   46
               Text            =   "�����ֵ"
               Top             =   4080
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
               Height          =   2415
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   840
               Width           =   4935
               _cx             =   8705
               _cy             =   4260
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
               BackColorSel    =   16711680
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483633
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483632
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   25
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmParStuff.frx":9026
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
               VirtualData     =   0   'False
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
            Begin VB.Label lblComment 
               Caption         =   $"frmParStuff.frx":92D1
               ForeColor       =   &H00000080&
               Height          =   540
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   4980
            End
         End
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   2415
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   10
         Top             =   120
         Width           =   45
      End
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   7
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   8
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1800
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParStuff.frx":935F
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParStuff.frx":F139
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   14040
      TabIndex        =   0
      Top             =   8640
      Width           =   14040
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11760
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   10605
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   14
         Top             =   165
         Width           =   4455
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ������(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   12
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   165
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmParStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mstrLike As String
Private mrs���� As New ADODB.Recordset
Private mblnLoad As Boolean     '������ؽ���
Private mblnOk As Boolean

Private Enum constCbo
    cbo_������Ŀ = 0
    cbo_סԺ�����Զ����� = 1
End Enum

'������Ƶ�������Ŀ
Private Const cst������Ŀ As String = "�ɹ���,����,�����,������,�ۼ�,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"

Private Enum constDigit
    dig_������� = 0
    dig_�������� = 1
    dig_���ȵ�λ = 2
    dig_���� = 3
    dig_��С���� = 4
    dig_��󾫶� = 5
    dig_ԭʼ���� = 6
    dig_��� = 7
    dig_���� = 8
    dig_��λ = 9
    dig_Cols = 10
End Enum

'ҩƷ���ĵ��ݻ�����Ŀ����
'��������
Private Enum ����
    ҩƷ�⹺ = 1
    �����⹺ = 15
End Enum

'ҵ�񻷽�
Private Enum ����
    �˲� = 1
    ��� = 2
    ������� = 3
End Enum

'
'Private Enum constListBox
'
'End Enum
'
'Private Enum constUd
'
'End Enum
'
Private Enum constTxt
    'ϵͳ����
    txt_����ǰ׺ = 0
    
    '�⹺���
    txt_�⹺����У�� = 1
    
    '�ƻ�
    txt_�ƻ�����У�� = 2
    
    '����
    txt_�������Ϸ�ʽ = 3
    
    txt_���ʱ��ģʽ = 4
    
    '����
    txt_������ֵ = 5
End Enum
'
'Private Enum constBill
'
'End Enum
'
'Private Enum constDigit
'
'End Enum

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    'ϵͳ����
    chk_ʱ�����ļӳ������ = 0
    '�������
    chk_�������ƿ����� = 31
    chk_�������������� = 1
    chk_��¿��ÿ�� = 2
    chk_�������������� = 3
    
    chk_�ֶμӳ���� = 4
    chk_���ϸ����ָ���� = 5
    chk_ʱ�����İ���ǰ�ӳ����� = 6
    chk_�������ϲ������� = 7
    chk_ʱ������ֱ��ȷ���ۼ� = 8
    chk_�⹺�����Ҫ�˲� = 9
    chk_ʱ���������ȡ�ϴ��ۼ� = 10
    
    chk_��Ŀִ��ǰ���շѻ���� = 26
    chk_δ�շѵ����ﻮ�۴������� = 27
    chk_δ��˵ļ��˴������� = 28
    
    chk_���������Զ����� = 29
        
    '�⹺���
    chk_�����޸Ĳɹ��޼� = 11
    chk_�б����Ŀ�ѡ����б굥λ��� = 12
    chk_��ֵ���ı�����д��ϸ��Ϣ = 13
    chk_��������Ч�ڼ�� = 14
    
    chk_�����������Ų��ؿ��� = 34
    
    '�ƿ�
    chk_������� = 16
    
    '����
    chk_�������� = 17
    chk_����������� = 18
    
    '�̵�
    chk_�̵�洢�ⷿ���� = 19
    
    '����
    chk_ʱ�����İ����ε��� = 20
    
    '����
    chk_�Զ����� = 21
    chk_ȱ�ϼ�� = 22
    chk_������ǩ�� = 23
    chk_����ʱ�����������ʼ�¼ = 24
    chk_����ҽ��������ʱ����� = 25
End Enum

Private Enum constVSF
    vsf_�⹺����У�� = 0
    vsf_�ƻ�����У�� = 1
End Enum

Private Enum m�ⷿ����
    mint����id = 0
    mint���ϲ��� = 1
    mint�ⷿid = 2
    mint���Ĳֿ� = 3
    mint����ⷿid
    mint����ⷿ
    mint����
    mintCount = 7
End Enum

Private Enum m�ⷿ���
    mintid = 0
    mint����
    mint����
    mint��鷽ʽ
    minCheck
    mintCount = 5
End Enum
Private Function Get�������Ϸ�ʽ() As String
    Dim n As Integer
    Dim str�������� As String
    
    '������ҩ
    If chk����.Value = 0 Then
        str�������� = ""
    Else
        For n = 0 To chkDeptType.Count - 1
            If chkDeptType(n).Value = 0 Then
                str�������� = IIf(str�������� = "", "", str�������� & ",") & chkDeptType(n).Caption
            End If
        Next
        If str�������� = "" Then
            str�������� = "�ٴ�,����,���,����,����,����,Ӫ��"
        End If
    End If
    
    Get�������Ϸ�ʽ = str��������
End Function

Private Sub Save�����()
    Dim i As Integer
    
    '����ⷿ���
    gstrSQL = ""
    With vsf�ⷿ���
        For i = 1 To .Rows - 1
            gstrSQL = gstrSQL & .TextMatrix(i, m�ⷿ���.mintid) & "," & Switch(.TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "0-�����", "0", .TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "1-��飬��������", "1", .TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "2-��飬�����ֹ", "2") & ","
        Next
    End With

    gstrSQL = "Zl_���ϳ�����_insert('" & gstrSQL & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
End Sub

Private Sub Save�������()
    Dim str����  As String
    Dim i As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim str���ڿⷿid As String
    Dim str�Է��ⷿid As String
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim bln���� As Boolean
    Dim arrSQL  As Variant
    
    arrSQL = Array()
    With vsf����
        For lngRow = 1 To .Rows - 1
            str���� = Left(.TextMatrix(lngRow, .ColIndex("����")), 1)
            If str���� = "" Then str���� = "3"
            
            str���ڿⷿid = ""
            str�Է��ⷿid = ""
            
            If .TextMatrix(lngRow, .ColIndex("���ڿⷿid")) = "" And lngRow <> .Rows - 1 Then
                gstrSQL = "select id from ���ű� where ����=[1]"
                strID = Mid(.TextMatrix(lngRow, .ColIndex("���ڿⷿ")), 1, InStr(1, .TextMatrix(lngRow, .ColIndex("���ڿⷿ")), "-") - 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID)
                If rsTemp.RecordCount > 0 Then
                    str���ڿⷿid = rsTemp!Id
                End If
            Else
                str���ڿⷿid = .TextMatrix(lngRow, .ColIndex("���ڿⷿid"))
            End If
            
            If .TextMatrix(lngRow, .ColIndex("�Է��ⷿid")) = "" And lngRow <> .Rows - 1 Then
                strID = Mid(.TextMatrix(lngRow, .ColIndex("�Է��ⷿ")), 1, InStr(1, .TextMatrix(lngRow, .ColIndex("�Է��ⷿ")), "-") - 1)
                gstrSQL = "select id from ���ű� where ����=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Է��ⷿ��ѯ", strID)
                If rsTemp.RecordCount > 0 Then
                    str�Է��ⷿid = rsTemp!Id
                End If
            Else
                str�Է��ⷿid = .TextMatrix(lngRow, .ColIndex("�Է��ⷿid"))
            End If
            If str���ڿⷿid <> "" Or str�Է��ⷿid <> "" Then
                If LenB(StrConv(strTemp & str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ",", vbFromUnicode)) >= 4000 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strTemp
                    strTemp = str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ","
                    bln���� = True
                Else
                    strTemp = strTemp & str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ","
                End If
            End If
        Next
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTemp
    End With
    
    For i = 0 To UBound(arrSQL)
        If bln���� = True Then
            If i = 0 Then
                Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "ɾ�����ۼ�¼")
            Else
                Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',1" & ")", "ɾ�����ۼ�¼")
            End If
        Else
            Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "ɾ�����ۼ�¼")
        End If
    Next
End Sub

Private Sub Save����ⷿ����()
    Dim strTemp As String
    Dim i As Integer
    Dim str����id As String
    
    '��������ⷿ����
    With vsf����
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, m�ⷿ����.mint����id)) > 0 And Val(.TextMatrix(i, m�ⷿ����.mint�ⷿid)) > 0 And Val(.TextMatrix(i, m�ⷿ����.mint����ⷿid)) > 0 Then
                If InStr(1, "," & str����id & ",", "," & Val(.TextMatrix(i, m�ⷿ����.mint����id)) & ",") = 0 Then
                    str����id = IIf(str����id = "", "", str����id & ",") & .TextMatrix(i, m�ⷿ����.mint����id)
                    strTemp = IIf(strTemp = "", "", strTemp & "|") & .TextMatrix(i, m�ⷿ����.mint����id) & "," & .TextMatrix(i, m�ⷿ����.mint�ⷿid) & "," & .TextMatrix(i, m�ⷿ����.mint����ⷿid)
                End If
            Else
                If Val(.TextMatrix(i, m�ⷿ����.mint����id)) = 0 Or Val(.TextMatrix(i, m�ⷿ����.mint�ⷿid)) = 0 Or Val(.TextMatrix(i, m�ⷿ����.mint����ⷿid)) = 0 Then
                    If Not (.TextMatrix(i, m�ⷿ����.mint���ϲ���) = "" And .TextMatrix(i, m�ⷿ����.mint���Ĳֿ�) = "" And .TextMatrix(i, m�ⷿ����.mint����ⷿ) = "") Then
                        MsgBox "������ⷿ���ա��ڡ�" & i & "����û��������ȷ�Ĳ��Ż�ⷿ�����б���ʧ�ܣ�", vbInformation, gstrSysName
                    End If
                End If
            End If
        Next
    End With

    gstrSQL = "Zl_����ⷿ����_Update('" & strTemp & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
End Sub

Private Sub Set�������Ϸ�ʽ(ByVal str�������� As String)
    Dim BlnSelect As Boolean
    Dim strArr As Variant
    Dim i As Integer
    Dim n As Integer
    
    BlnSelect = False
    If str�������� = "" Then
        BlnSelect = False
    ElseIf str�������� = "�ٴ�,����,���,����,����,����,Ӫ��" Then
        BlnSelect = True
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
    Else
        str�������� = str�������� & ","
        strArr = Split(str��������, ",")
        
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
        
        For i = 0 To UBound(strArr)
            For n = 0 To chkDeptType.Count - 1
                If strArr(i) = chkDeptType(n).Caption Then
                    chkDeptType(n).Value = 0
                    BlnSelect = True
                    Exit For
                End If
            Next
        Next
    End If
    If BlnSelect = True Then
        chk����.Value = 1
        chk����.Tag = 1
    Else
        chk����.Value = 0
        chk����.Tag = 0
    End If
    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = BlnSelect
    Next
End Sub

Private Sub cbo_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(cbo, Index, mrsPar)
    End If
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub

Private Sub chkDeptType_Click(Index As Integer)
    Dim n As Integer
    Dim blnAllUnselect As Boolean
    
    '����Ҫѡ��һ��
    blnAllUnselect = True
    For n = 0 To chkDeptType.Count - 1
        If chkDeptType(n).Value = 1 Then
            blnAllUnselect = False
            Exit For
        End If
    Next
    If blnAllUnselect = True Then
        chkDeptType(Index).Value = 1
    End If
    
    txt(txt_�������Ϸ�ʽ).Text = Get�������Ϸ�ʽ
End Sub

Private Sub chk����_Click()
    Dim n As Integer

    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = (chk����.Value = 1)
        If chk����.Tag = "0" Then
            chkDeptType(n).Value = 1
        End If
    Next
    
    txt(txt_�������Ϸ�ʽ).Text = Get�������Ϸ�ʽ
End Sub

Private Sub chkӦ�÷�Χ_Click(Index As Integer)
    Dim objӦ�÷�Χ As CheckBox
    Dim strValue As String
    
    If mblnLoad = False Then Exit Sub
    
    If chkӦ�÷�Χ(Index).Value <> Val(chkӦ�÷�Χ(Index).Tag) Then
        chkӦ�÷�Χ(Index).ForeColor = &HC0&             '�޸ĺ������ɫǰ��ɫ��ʶ
    Else
        chkӦ�÷�Χ(Index).ForeColor = &H0&
    End If

    For Each objӦ�÷�Χ In chkӦ�÷�Χ
        strValue = IIf(strValue = "", "", strValue) & objӦ�÷�Χ.Value
    Next
    Call SetParChange(chkӦ�÷�Χ, 0, mrsPar, True, strValue)
End Sub

Private Sub chkӦ�÷�Χ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(chkӦ�÷�Χ, 0, mrsPar, "", chkӦ�÷�Χ(Index))
End Sub

Private Sub CmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "��ʼ�ɹ�" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    mblnOk = False
    
    '���ڴ�С��13000,8385
    Me.Width = 13000
    Me.Height = 8385
    
    For Each objPic In picPar
        Set objPic.Container = Me
    Next
    
    tabDesign.Visible = False
    
    strCategory = "��������,������Ŀ"
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "104,4,����ͨ������;100,0,����Ŀ¼����;101,1,�����������;102,2,�����ڿ����;103,3,���ķ��Ź���"
    
    '��������Pickture������10��ʼ��
    marrFunc(1) = "110,10,�����������;111,11,���Ŀ����;112,12,����ⷿ����;113,13,����¼�뾫��;114,14,���ݻ��ڿ���"
    
    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    
    Me.Tag = "��ʼ�ɹ�"
    
    mblnLoad = True
End Sub

Private Sub opt����ģʽ_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt����ģʽ, Index, mrsPar)
    End If
End Sub

Private Sub opt����ģʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt����ģʽ, Index, mrsPar)
End Sub

Private Sub opt�ɹ���_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt�ɹ���, Index, mrsPar)
    End If
End Sub

Private Sub opt�ɹ���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt�ɹ���, Index, mrsPar)
End Sub

Private Sub opt����_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt����, Index, mrsPar)
    End If
End Sub

Private Sub opt����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt����, Index, mrsPar)
End Sub
Private Sub optBarcode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBarcode, Index, mrsPar)
End Sub
Private Sub optBarcode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(optBarcode, Index, mrsPar)
End Sub

Private Sub opt���۵�λ_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt���۵�λ, Index, mrsPar)
    End If
End Sub

Private Sub opt���۵�λ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt���۵�λ, Index, mrsPar)
End Sub

Private Sub opt��������_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt��������, Index, mrsPar)
    End If
End Sub

Private Sub opt��������_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt��������, Index, mrsPar)
End Sub

Private Sub opt�ƻ�����У��_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_�ƻ�����У��, mrsPar, True, Get��Ӧ������У��(vsf_�ƻ�����У��))
    End If
    
    fra����У��(vsf_�ƻ�����У��).ForeColor = txt(txt_�ƻ�����У��).ForeColor
End Sub

Private Sub opt��淽ʽ_Click(Index As Integer)
    Dim strValue As String
    
    If opt��淽ʽ(0).Value = True Then
        opt���ʱ��ģʽ(0).Enabled = False
        opt���ʱ��ģʽ(1).Enabled = False
        txt(txt_���ʱ��ģʽ).Enabled = False
        
        '�ֹ�������ֵΪ-1
        strValue = "-1"
    Else
        opt���ʱ��ģʽ(0).Enabled = True
        opt���ʱ��ģʽ(1).Enabled = True
        txt(txt_���ʱ��ģʽ).Enabled = opt���ʱ��ģʽ(1).Value
        
        strValue = IIf(opt���ʱ��ģʽ(0).Value, 0, Val(txt(txt_���ʱ��ģʽ).Text))
    End If
    
    If Me.Visible Then
        Call SetParChange(txt, txt_������ֵ, mrsPar, True, strValue)
        
        opt��淽ʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt��淽ʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        txt(txt_���ʱ��ģʽ).ForeColor = opt���ʱ��ģʽ(1).ForeColor
    End If
End Sub

Private Sub opt���ʱ��ģʽ_Click(Index As Integer)
    Dim strValue As String
    
    txt(txt_���ʱ��ģʽ).Enabled = opt���ʱ��ģʽ(1).Value
    
    If Me.Visible Then
        strValue = IIf(opt���ʱ��ģʽ(0).Value, 0, Val(txt(txt_���ʱ��ģʽ).Text))
        Call SetParChange(txt, txt_������ֵ, mrsPar, True, strValue)
        
        opt��淽ʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt��淽ʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        txt(txt_���ʱ��ģʽ).ForeColor = opt���ʱ��ģʽ(1).ForeColor
    End If
End Sub

Private Sub opt�⹺����У��_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_�⹺����У��, mrsPar, True, Get��Ӧ������У��(vsf_�⹺����У��))
    End If
    
    fra����У��(vsf_�⹺����У��).ForeColor = txt(txt_�⹺����У��).ForeColor
End Sub

Private Sub opt���ĳ����㷨_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt���ĳ����㷨, Index, mrsPar)
    End If
End Sub

Private Sub opt���ĳ����㷨_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt���ĳ����㷨, Index, mrsPar)
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        objPic.Visible = (objPic.Index = Item.Id)
    Next
        
    lblLocate(txt_Dept).Visible = (Item.Id = GetFuncID("ҩ����ҩ����", marrFunc) Or _
                            Item.Id = GetFuncID("��Һ��������", marrFunc) Or _
                            Item.Id = GetFuncID("ҩƷ�������", marrFunc) Or _
                            Item.Id = GetFuncID("ҩƷ�����", marrFunc) Or _
                            Item.Id = GetFuncID("ҩƷ������λ", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        
        If Item.Id = GetFuncID("��Һ��������", marrFunc) Then
            lblLocate(txt_Dept).Caption = "���Ҳ���(&F)"
        Else
            lblLocate(txt_Dept).Caption = "ҩ������(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.Id   '���ڻ�ȡ��ǰѡ�е�TaskPanelItem
End Sub

Private Sub Form_Resize()
    Dim i As Long
    Dim objPic As PictureBox
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For Each objPic In picPar
        objPic.Top = Me.ScaleTop
        objPic.Left = picFunc.Left + picFunc.ScaleWidth
        objPic.Width = Me.ScaleWidth - objPic.Left
        objPic.Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
    
'    For i = 0 To picPar.UBound
'        If Not picPar(i) Is Nothing Then
'            picPar(i).Top = Me.ScaleTop
'            picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
'            picPar(i).Width = Me.ScaleWidth - picPar(i).Left
'            picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
'        End If
'    Next
End Sub

Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub

Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIf(picVbar.Left + x < 2000, 2000, picVbar.Left + x)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.Id - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub

Public Sub LocateFuncItem(ByVal lngFunc As Long)
'���ܣ�����IDѡ��һ���Ͷ�������
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).Id = lngId
                Next
            End If
        Next
    Next
End Sub

Private Sub InitData()
'���ܣ���ʼ������ؼ�,��ȡ����������
    
    '1.��ʼ������
    
    mlngPreFind = 1
    Call InitSystemPara
    
    mstrLike = IIf(gstrMatchMethod = "0", "%", "")
    
    '2.��ʼ������ؼ�
    Call InitEnv
    
    '����������������
    Call Load��������
    Call Load�ⷿ���
    Call Load����ⷿ����
    Call LoadҩƷ���ľ���
    Call Load���ݻ��ڿ���
    
    '3.����ϵͳ����
    Call LoadPar
    
End Sub

Private Sub Load��������()
    '����:װ�������������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTemp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    With vsf����
        .Rows = 1
        .Rows = 2
        
        '����װ��ⷿ
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('���Ŀ�','�Ƽ���','����ⷿ','���ϲ���') " & _
                   "   and  b.����ID=a.ID and " & Where����ʱ��("A") & _
                   " order by ����"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        If Not rsTemp.EOF Then
            rsTemp.MoveFirst
            For i = 1 To rsTemp.RecordCount
                strTemp = strTemp & rsTemp!���� & "-" & rsTemp!���� & "|"
                rsTemp.MoveNext
            Next
        End If
        .ColComboList(.ColIndex("���ڿⷿ")) = strTemp
        .ColComboList(.ColIndex("�Է��ⷿ")) = strTemp
        .ColComboList(.ColIndex("����")) = "1-���ڿⷿ������Է��ⷿ|2-�Է��ⷿ���������ڿⷿ|3-���ⷿ���˫����ͨ"
        
        'װ�������������
        gstrSQL = "select A.���ڿⷿID,A.�Է��ⷿID,A.����" & _
                ",B.���� as ���ڱ���,B.���� as ��������,C.���� as �Է�����,C.���� as �Է����� " & _
                " from ����������� A,���ű� B,���ű� C " & _
                " where A.���ڿⷿID= B.ID and A.�Է��ⷿID=C.ID " & _
                "   and (b.����ʱ��=to_date('3000-1-1','yyyy-mm-dd') or b.����ʱ�� is null) " & _
                " order by b.����, c.���� "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
                                                          
            .TextMatrix(lngRow, .ColIndex("���ڿⷿ")) = IIf(IsNull(rsTemp!���ڿⷿid), "", rsTemp!���ڱ��� & "-" & rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("���ڿⷿid")) = rsTemp!���ڿⷿid
            .TextMatrix(lngRow, .ColIndex("�Է��ⷿ")) = IIf(IsNull(rsTemp!�Է��ⷿID), "", rsTemp!�Է����� & "-" & rsTemp!�Է�����)
            .TextMatrix(lngRow, .ColIndex("�Է��ⷿid")) = rsTemp!�Է��ⷿID
            .TextMatrix(lngRow, .ColIndex("����")) = Switch(rsTemp("����") = 1, "1-���ڿⷿ������Է��ⷿ", _
                                            rsTemp("����") = 2, "2-�Է��ⷿ���������ڿⷿ", _
                                                          True, "3-���ⷿ���˫����ͨ")
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load�ⷿ���()
    '���ܣ���ʼ���ⷿ
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim objItem As ListItem
    
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) ��鷽ʽ" & vbCrLf & _
        " FROM ��������˵�� A, ���ű� B, ���ϳ����� C" & vbCrLf & _
        " WHERE A.����ID = B.ID AND A.����ID = C.�ⷿID(+) AND" & vbCrLf & _
        "      A.�������� IN" & vbCrLf & _
        "      ('���Ŀ�','�Ƽ���','���ϲ���','����ⷿ') " & vbCrLf & _
        "     And (b.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd') or b.����ʱ�� is null) " & vbCrLf & _
        " GROUP BY B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0)" & vbCrLf & _
        " ORDER BY B.���� "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    Me.vsf�ⷿ���.Rows = 1
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        vsf�ⷿ���.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            With vsf�ⷿ���
                .TextMatrix(i, .ColIndex("id")) = rsTmp!Id
                .TextMatrix(i, .ColIndex("��������")) = rsTmp!����
                .TextMatrix(i, .ColIndex("����")) = rsTmp!����
                .TextMatrix(i, .ColIndex("�ⷿ��鷽ʽ")) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
                .TextMatrix(i, .ColIndex("check")) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
            End With
            rsTmp.MoveNext
        Next
        
        vsf�ⷿ���.Cell(flexcpBackColor, 1, vsf�ⷿ���.ColIndex("�ⷿ��鷽ʽ"), vsf�ⷿ���.Rows - 1, vsf�ⷿ���.ColIndex("�ⷿ��鷽ʽ")) = &HF4F4EA
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load����ⷿ����()
    '����:װ����������ⷿ���չ�ϵ
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    
    With vsf����
        'ȡ���з��ϲ��ţ����Ŀ⣬����ⷿ
        mrs����.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.����, b.�������� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('���Ŀ�','���ϲ���','����ⷿ') " & _
                   " and  b.����ID=a.ID and " & Where����ʱ��("A") & " order by ����"
        zlDatabase.OpenRecordset mrs����, gstrSQL, Me.Caption
        
        'װ��Ŀǰ������ⷿ���չ�ϵ
        gstrSQL = "Select b.Id As ����id, b.���� || '-' || b.���� As ���ϲ���, c.Id As �ⷿid, c.���� || '-' || c.���� As ���Ĳֿ�," & _
                  " d.Id As ����ⷿid,d.���� || '-' || d.���� As ����ⷿ " & _
                  "From ����ⷿ���� A, ���ű� B, ���ű� C, ���ű� D " & _
                  "Where a.����id = b.Id And a.�ⷿid = c.Id And a.����ⷿid = d.Id " & _
                  "  And (b.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd') or b.����ʱ�� is null) " & _
                  "Order by b.���� "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("����id")) = rsTemp!����id
            .TextMatrix(lngRow, .ColIndex("���ϲ���")) = rsTemp!���ϲ���
            .TextMatrix(lngRow, .ColIndex("���Ĳֿ�id")) = rsTemp!�ⷿID
            .TextMatrix(lngRow, .ColIndex("���Ĳֿ�")) = rsTemp!���Ĳֿ�
            .TextMatrix(lngRow, .ColIndex("����ⷿid")) = rsTemp!����ⷿid
            .TextMatrix(lngRow, .ColIndex("����ⷿ")) = rsTemp!����ⷿ
'            .TextMatrix(lngRow, .ColIndex("����")) = "��"
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LoadPar()
'���ܣ���ȡ�����ز���������ؼ�
    Dim strValue As String, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '�������ģ��1,������1,�ؼ�����1,ģ��2,������2,�ؼ�����2,......
    Dim n As Integer
    Dim objӦ�÷�Χ As CheckBox
    
    '��ȡ����(Ĭ�϶�ȡϵͳ��������Ҫ��ģ�����������Ӷ�Ӧ��ģ���)
    Set rsTmp = GetPar(mrsPar, p����Ŀ¼���� & "," & _
            p�����⹺���� & "," & _
            p�����ƿ���� & "," & _
            p�������ù��� & "," & _
            p�����̵���� & "," & _
            p����������� & "," & _
            p���ķ��Ź��� & "," & _
            p���ļƻ����� & "," & _
            p���ĵ��۹���)
    

    '----------------------------------------------------------
    'ϵͳ����
    '1.����CheckBox�����
    strTmp = "0:82:" & chk_ʱ�����ļӳ������ & _
            ",0:280:" & chk_�������ƿ����� & _
            ",0:83:" & chk_�������������� & _
            ",0:95:" & chk_��¿��ÿ�� & _
            ",0:121:" & chk_�ֶμӳ���� & _
            ",0:123:" & chk_���ϸ����ָ���� & _
            ",0:127:" & chk_ʱ�����İ���ǰ�ӳ����� & _
            ",0:132:" & chk_�������ϲ������� & _
            ",0:136:" & chk_ʱ������ֱ��ȷ���ۼ� & _
            ",0:140:" & chk_�⹺�����Ҫ�˲� & _
            ",0:163:" & chk_��Ŀִ��ǰ���շѻ���� & _
            ",0:171:" & chk_δ�շѵ����ﻮ�۴������� & _
            ",0:172:" & chk_δ��˵ļ��˴������� & _
            ",0:229:" & chk_ʱ���������ȡ�ϴ��ۼ� & _
            ",0:92:" & chk_���������Զ����� & _
            ",0:258:" & chk_�������������� & _
            ",0:305:" & chk_�����������Ų��ؿ���
    Call SetParToControl(strTmp, mrsPar, chk)
    
'    chk(chk_��¿��ÿ��).Enabled = (Check���쵥 And Check�ƿⵥ And Check���õ�)
'    chk(chk_��������������).Enabled = Check���쵥
'    chk(chk_�������ƿ�����).Enabled = Check�ƿⵥ
'    chk(chk_��������������).Enabled = Check���õ�

    '���ò�����ϵ
    If chk(chk_��Ŀִ��ǰ���շѻ����).Value = 1 Then
        chk(chk_δ�շѵ����ﻮ�۴�������).Enabled = False
        chk(chk_δ��˵ļ��˴�������).Enabled = False
        lblδ�շѷ�ҩ.Caption = "  ������������һ��ͨ������ִ��ǰ�������շѻ��ȼ�����ˡ���������ﲡ�˷���ʱ�����²������۹�ѡ����ʧЧ��"
    Else
        chk(chk_δ�շѵ����ﻮ�۴�������).Enabled = True
        chk(chk_δ��˵ļ��˴�������).Enabled = True
        lblδ�շѷ�ҩ.Caption = "  �������������һ��ͨ������ִ��ǰ�������շѻ��ȼ�����ˡ���������ﲡ�˷���ʱ�����²�����ʧЧ��"
    End If
    
    
'    '2.����ComboBox�����
'    strTmp = "0:29:" & cbo_���۵�λ & _
'            ",0:64:" & cbo_ҩƷ������� & _
'            ",0:87:" & cbo_ҩƷ����ģʽ & _
'            ",0:149:" & cbo_Ч����ʾ��ʽ & _
'            ",0:150:" & cbo_ҩƷ���������㷨
'
'    Call SetParToControl(strTmp, mrsPar, cbo)
'
'
'    '3.����UpDown�����
'    strTmp = ""
'    'Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
'
'
    '4.����TextBox�����
    strTmp = "0:159:" & txt_����ǰ׺
    Call SetParToControl(strTmp, mrsPar, txt)
'
'    '5.����ListBox�����
''    strTmp = pסԺҽ���´� & ":44:" & lst_��Һ���ķ�ҩ���˿���
''    Call SetParToControl(strTmp, mrsPar, lst, 1)
'
    '6.����OptionButton�����
    arrObj = Array(0, 88, opt���۵�λ, _
                 0, 156, opt���ĳ����㷨, _
                 0, 268, opt����, _
                 0, 320, optBarcode)
    Call SetParToControl("", mrsPar, arrObj)
    
'    '7.����ϵͳ����
    rsTmp.Filter = "ģ��=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        
        Case 281   'ҩƷ���ʱ��ģʽ
            If Val(strValue) = -1 Then
                '����ֵΪ-1��ʾ�ֹ����
                opt��淽ʽ(0).Value = True
                opt��淽ʽ(1).Value = False
                
                opt���ʱ��ģʽ(0).Enabled = False
                opt���ʱ��ģʽ(1).Enabled = False
                txt(txt_���ʱ��ģʽ).Enabled = False
            Else
                '����ֵ��Ϊ-1��ʾ�Զ����
                opt��淽ʽ(0).Value = False
                opt��淽ʽ(1).Value = True
                
                If Val(strValue) = 0 Then
                    '����ֵΪ0��ʾÿ�����һ����
                    opt���ʱ��ģʽ(0).Value = True
                    opt���ʱ��ģʽ(1).Value = False
                    txt(txt_���ʱ��ģʽ).Enabled = False
                Else
                    '����ֵ����0С�ڵ���31��ʾָ�����ڽ��
                    opt���ʱ��ģʽ(0).Value = False
                    opt���ʱ��ģʽ(1).Value = True
                    
                    txt(txt_���ʱ��ģʽ).Enabled = True
                    
                    '���ʱ��ֻ������Ϊ1-31
                    If Val(strValue) > 0 Or Val(strValue) <= 31 Then
                        txt(txt_���ʱ��ģʽ).Text = Val(strValue)
                    Else
                        txt(txt_���ʱ��ģʽ).Text = "25"
                    End If
                End If
            End If
            
            Call SetParRelation(txt, txt_������ֵ, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txt, txt_���ʱ��ģʽ, mrsPar)
            
         Case 320 '���������ϱ�:1-����ͨ��ɨ��¼���¼���������ϱ���������;0-�����ƣ�����ͨ�����롢���롢�����¼�뷽ʽ��ʶ����������
            
            If Val(strValue) = 0 Then
                optBarcode(0).Value = True: optBarcode(1).Value = False
            Else
                optBarcode(0).Value = False: optBarcode(1).Value = True
            End If
        End Select
        rsTmp.MoveNext
    Loop

    '----------------------------------------------------------
    '8.����ģ�����
    '����Ŀ¼���� = 1711
    '����ComboBox�����
    strTmp = p����Ŀ¼���� & ":������Ŀ��Ӧ:" & cbo_������Ŀ & _
            ",0:63:" & cbo_סԺ�����Զ�����
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '����OptionButton�����
    arrObj = Array(p����Ŀ¼����, "�������ģʽ", opt����ģʽ, _
                    p����Ŀ¼����, "���ķ��������Զ�����", opt��������)
    Call SetParToControl("", mrsPar, arrObj)
    
    '��������
    '���������������ֵ��Ӧ����ؼ�(��)���ȵ��ù���������¼�ؼ����ƣ�����ؼ���ʾ��������
    strTmp = p����Ŀ¼���� & ":����Ӧ���ڵķ�Χ:0"
    Call SetParToControl(strTmp, mrsPar, chkӦ�÷�Χ)

    rsTmp.Filter = "ģ��=" & p����Ŀ¼���� & " And ������='����Ӧ���ڵķ�Χ'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ, "111")
    If strValue <> "" Then
        For n = 0 To chkӦ�÷�Χ.Count - 1
            chkӦ�÷�Χ(n).Value = Mid(strValue, n + 1, 1)
            chkӦ�÷�Χ(n).Tag = Mid(strValue, n + 1, 1)
        Next
    End If
    
    
    '----------------------------------------------------------
    '���
    '�����⹺���� = 1712
    '����CheckBox�����
    strTmp = p�����⹺���� & ":�޸Ĳɹ��޼�:" & chk_�����޸Ĳɹ��޼� & _
            "," & p�����⹺���� & ":�б����Ŀ�ѡ����б굥λ���:" & chk_�б����Ŀ�ѡ����б굥λ��� & _
            "," & p�����⹺���� & ":��ֵ���ı�����д��ϸ��Ϣ:" & chk_��ֵ���ı�����д��ϸ��Ϣ & _
            "," & p�����⹺���� & ":��������Ч�ڼ��:" & chk_��������Ч�ڼ��
    Call SetParToControl(strTmp, mrsPar, chk)

    '����OptionButton�����
    arrObj = Array(p�����⹺����, "��ⵥ�۳��б굥��", opt�ɹ���)
    Call SetParToControl("", mrsPar, arrObj)
    
    '���⴦��
    '�ò���ʵ���ñ��������ؼ���ʾ���ر���ö����ı��ؼ���¼ԭʼֵ���������������ʾ
    strTmp = p�����⹺���� & ":����У��:" & txt_�⹺����У��
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "ģ��=" & p�����⹺���� & " And ������='����У��'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    Call Load��Ӧ������У��(vsf_�⹺����У��, strValue)
    
    
    '----------------------------------------------------------
    '����
    '�����ƿ���� = 1716
    '�������ù��� = 1717
    '����CheckBox�����
    strTmp = p�����ƿ���� & ":��������:" & chk_������� & _
            "," & p�������ù��� & ":��������:" & chk_�������� & _
            "," & p�������ù��� & ":�������:" & chk_�����������
    Call SetParToControl(strTmp, mrsPar, chk)

    
    '----------------------------------------------------------
    '�����̵���� = 1719
    '����CheckBox�����
    strTmp = p�����̵���� & ":�洢�ⷿ:" & chk_�̵�洢�ⷿ����
    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    '----------------------------------------------------------
    '���ļƻ����� = 1724
    '���⴦��
    '�ò���ʵ���ñ��������ؼ���ʾ���ر���ö����ı��ؼ���¼ԭʼֵ���������������ʾ
    strTmp = p���ļƻ����� & ":����У��:" & txt_�ƻ�����У��
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "ģ��=" & p���ļƻ����� & " And ������='����У��'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    Call Load��Ӧ������У��(vsf_�ƻ�����У��, strValue)

    
    '----------------------------------------------------------
    '���ĵ��۹��� = 1726
    '����CheckBox�����
    strTmp = p���ĵ��۹��� & ":ʱ�����İ����ε���:" & chk_ʱ�����İ����ε���
    Call SetParToControl(strTmp, mrsPar, chk)


    '----------------------------------------------------------
    '���ķ��Ź��� = 1723
    '����CheckBox�����
    strTmp = p���ķ��Ź��� & ":�Զ�����:" & chk_�Զ����� & _
        "," & p���ķ��Ź��� & ":ȱ�ϼ��:" & chk_ȱ�ϼ�� & _
        "," & p���ķ��Ź��� & ":������ǩ��:" & chk_������ǩ�� & _
        "," & p���ķ��Ź��� & ":����ʱ�����������ʼ�¼:" & chk_����ʱ�����������ʼ�¼ & _
        "," & p���ķ��Ź��� & ":����ҽ��������ʱ�����:" & chk_����ҽ��������ʱ�����
    Call SetParToControl(strTmp, mrsPar, chk)

    '�ò���ʵ���ñ��������ؼ���ʾ���ر���ö����ı��ؼ���¼ԭʼֵ���������������ʾ
    strTmp = p���ķ��Ź��� & ":�������Ϸ�ʽ:" & txt_�������Ϸ�ʽ
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "ģ��=" & p���ķ��Ź��� & " And ������='�������Ϸ�ʽ'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    Call Set�������Ϸ�ʽ(strValue)
      
    
    '----------------------------------------------------------
    '������ϵ����
    If chk(chk_��¿��ÿ��).Value = 1 Then
        chk(chk_��������������).Value = 1
        chk(chk_��������������).Enabled = False
        
        chk(chk_��������������).Value = 1
        chk(chk_��������������).Enabled = False
        
        chk(chk_�������ƿ�����).Value = 1
        chk(chk_�������ƿ�����).Enabled = False
    End If
    
    If chk(chk_ʱ�����ļӳ������).Value = 1 Then
        chk(chk_�ֶμӳ����).Value = 0
        chk(chk_�ֶμӳ����).Enabled = False
        chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 0
        chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = False
    ElseIf chk(chk_�ֶμӳ����).Value = 1 Then
        chk(chk_ʱ�����ļӳ������).Value = 0
        chk(chk_ʱ�����ļӳ������).Enabled = False
        chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 0
        chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = False
    ElseIf chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 1 Then
        chk(chk_�ֶμӳ����).Value = 0
        chk(chk_�ֶμӳ����).Enabled = False
        chk(chk_ʱ�����ļӳ������).Value = 0
        chk(chk_ʱ�����ļӳ������).Enabled = False
    End If
        
End Sub

Private Function Get��Ӧ������У��(ByVal intType As Integer) As String
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    blnAllUnCheck = True
    
    '��������У����Ŀ�ͷ�ʽ����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    With vsfCheck(intType)
        For i = 1 To .Rows - 1
            strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("У����Ŀ")) & "," & _
                IIf(.TextMatrix(i, .ColIndex("У��")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("У��")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If intType = 0 Then
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf opt�⹺����У��(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
    Else
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf opt�ƻ�����У��(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
    End If
        
    Get��Ӧ������У�� = strCheck
End Function
Private Sub Load��Ӧ������У��(ByVal intType As Integer, ByVal strParaValue As String)
    Dim i As Integer
    Dim n As Integer
    Dim intCheckType As Integer
    Dim arrColumn
    
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....

    If strParaValue <> "" Then
        If InStr(1, strParaValue, "|") > 0 Then
            'У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
            intCheckType = Val(Mid(strParaValue, 1, InStr(1, strParaValue, "|") - 1))
            
            If intType = 0 Then
                If intCheckType = 2 Then
                    opt�⹺����У��(0).Value = True
                ElseIf intCheckType = 1 Then
                    opt�⹺����У��(1).Value = True
                End If
            Else
                If intCheckType = 2 Then
                    opt�ƻ�����У��(0).Value = True
                ElseIf intCheckType = 1 Then
                    opt�ƻ�����У��(1).Value = True
                End If
            End If
            
            strParaValue = Mid(strParaValue, InStr(1, strParaValue, "|") + 1)
             
            If strParaValue <> "" Then
                strParaValue = strParaValue & ";"
                arrColumn = Split(strParaValue, ";")
                For n = 0 To UBound(arrColumn)
                    If arrColumn(n) <> "" Then
                        With vsfCheck(intType)
                            For i = 1 To .Rows - 1
                                If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("���")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("У����Ŀ")) Then
                                    If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                        .TextMatrix(i, .ColIndex("У��")) = "��"
                                    End If
                                End If
                            Next
                        End With
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub InitEnv()
'���ܣ���ʼ������ؼ������ػ�������
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    '����Ŀ¼
    gstrSQL = "Select ID,����||'-'||���� ���� From ������Ŀ Where ĩ��=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "InitEnv")
    With rsData
        Do While Not .EOF
            cbo(cbo_������Ŀ).AddItem !����
            cbo(cbo_������Ŀ).ItemData(cbo(cbo_������Ŀ).NewIndex) = !Id
            .MoveNext
        Loop
    End With
    
    'סԺ�����Զ�����
    With cbo(cbo_סԺ�����Զ�����)
        .Clear
        .AddItem "0-���Զ�����"
        .ItemData(.NewIndex) = 0
        
        .AddItem "1-�Զ�����"
        .ItemData(.NewIndex) = 1
        
        .AddItem "2-�����ҿ����Զ�����"
        .ItemData(.NewIndex) = 2
        
        .ListIndex = 0
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mrsPar.Filter = "�޸�״̬=1"
    If mrsPar.RecordCount > 0 Then
    
'        Or Bill(bill_ҩƷ�ⷿ����).Tag = "���޸�" Or Bill(bill_ҩƷ��������).Tag = "���޸�" _
'        Or lvw�����.Tag = "���޸�" Or msf�ⷿ������λ.Tag = "���޸�" Or Billҩ����ҩ����.Tag = "���޸�" _
'        Or BillҩƷ���ľ���.Tag = "���޸�" Or vsf���ݻ��ڿ���.Tag = "���޸�" Then
        
        If Not mblnOk Then
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    Set mrsPar = Nothing
    Set mrs���� = Nothing
    
    mblnLoad = False
End Sub

Private Sub cmdOk_Click()
    Dim objӦ�÷�Χ As CheckBox
    Dim strValue As String
    
    If ValidateData() = False Then Exit Sub
    
    mblnOk = True
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    
    Call zlDatabase.ClearParaCache
    
    '������������
    Call Save�������
    Call Save�����
    Call Save����ⷿ����
    Call SaveҩƷ���ľ���
    Call Save���ݻ��ڿ���
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    
    Unload Me
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)

    Select Case Index
    Case txt_���ʱ��ģʽ
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 31 Then
            txt(Index).Text = 25
        End If
    End Select
    
    If Me.Visible Then
        Call SetParChange(txt, Index, mrsPar)
    End If
    
    If Index = txt_�������Ϸ�ʽ Then
        chk����.ForeColor = txt(txt_�������Ϸ�ʽ).ForeColor
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    
    ElseIf KeyAscii = Asc(gstrParSplit1) Or KeyAscii = Asc(gstrParSplit2) Then
        KeyAscii = 0
    Else
        Select Case Index
        Case txt_���ʱ��ģʽ
            Select Case KeyAscii
            Case vbKeyBack, vbKeyEscape, 3, 22  'С����
                KeyAscii = 0
            Case Else
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
            End Select
        End Select
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
'            If Billҩ����ҩ����.Visible Then
'                Call LocateDept(strFind, Billҩ����ҩ����, 0)
'
'            ElseIf Bill(bill_ҩƷ��������).Visible Then
'                If lblLocate(txt_Dept).Tag = bill_ҩƷ�ⷿ���� Or lblLocate(txt_Dept).Tag = "" Then
'                    Call LocateDept(strFind, Bill(bill_ҩƷ�ⷿ����), IIf(Bill(bill_ҩƷ�ⷿ����).Col = 0, 0, 1))
'                Else
'                    Call LocateDept(strFind, Bill(bill_ҩƷ��������), Bill(bill_ҩƷ��������).Col)
'                End If
'
'            ElseIf lvw�����.Visible Then
'                Call LocateDept(strFind, lvw�����, 1)
'
'            ElseIf msf�ⷿ������λ.Visible Then
'                Call LocateDept(strFind, msf�ⷿ������λ, 0)
        End Select
    End If
End Sub


Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object, ByVal lngCol As Long)
'���ܣ����ҿ���
'������lngCol-���в��ҵ���
    Dim i As Long, lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    With objTmp
        If TypeName(objTmp) = "ListView" Then 'lvw�����
            lngRows = .ListItems.Count
            For i = mlngPreFind To lngRows
                If .ListItems(i).ListSubItems(lngCol).Text Like IIf(mstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst_��Һ���ķ�ҩ���˿���
            With objTmp
                lngRows = .ListCount - 1
                
                lngStart = IIf(mlngPreFind = 1, 0, mlngPreFind)
                For i = lngStart To .ListCount - 1
                    strCode = Split(.List(i), "-")(0)
                    strName = Split(.List(i), "-")(1)
                    If strCode Like strFind & "*" Or strName Like IIf(mstrLike <> "", "*", "") & strFind & "*" Then
                        .ListIndex = i
                        .SetFocus
                        Exit For
                    End If
                Next
            End With
        Else
            lngRows = objTmp.Rows
            For i = mlngPreFind To .Rows - 1
                If InStr(.TextMatrix(i, lngCol), "-") > 0 Then
                    strCode = Split(.TextMatrix(i, lngCol), "-")(0)
                    strName = Split(.TextMatrix(i, lngCol), "-")(1)
                Else
                    strCode = ""
                    strName = .TextMatrix(i, lngCol)
                End If
                
                If strCode Like strFind & "*" Or strName Like IIf(mstrLike <> "", "*", "") & strFind & "*" Then
                    objTmp.SetFocus
                    .Row = i: .Col = lngCol
                    .TopRow = i
                    Exit For
                End If
            Next
        End If
    End With
    If i < lngRows Then
        mlngPreFind = i + 1
    Else
        If mlngPreFind = 1 Then
            MsgBox "û���ҵ�ƥ��ģ�������������ݡ�", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        Else
            MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, Me.Caption
            mlngPreFind = 1
        End If
    End If
End Sub

Private Function ValidateData() As Boolean
    Dim lngRow As Long
    Dim j As Long
    
    With vsf����
        For lngRow = 1 To .Rows - 1
            If (.TextMatrix(lngRow, .ColIndex("���ڿⷿ")) = "" Or .TextMatrix(lngRow, .ColIndex("�Է��ⷿ")) = "" Or .TextMatrix(lngRow, .ColIndex("����")) = "") And lngRow <> .Rows - 1 Then
                MsgBox "���������У���" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("���ڿⷿ")

                Exit Function
            End If

            If .TextMatrix(lngRow, .ColIndex("���ڿⷿ")) = .TextMatrix(lngRow, .ColIndex("�Է��ⷿ")) And lngRow <> .Rows - 1 Then
                MsgBox "���������У���" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("���ڿⷿ")

                Exit Function
            End If

            For j = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("���ڿⷿ")) = .TextMatrix(j, .ColIndex("���ڿⷿ")) And .TextMatrix(lngRow, .ColIndex("�Է��ⷿ")) = .TextMatrix(j, .ColIndex("�Է��ⷿ")) And lngRow <> j Then
                    MsgBox "���������У���" & lngRow & "�����" & j & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = .ColIndex("���ڿⷿ")

                    Exit Function
                End If
            Next
        Next
    End With
    
    ValidateData = True
End Function

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnResulte As Boolean
    
    On Error GoTo ErrHandle
    
    If Me.Visible Then
        Call SetParChange(chk, Index, mrsPar)
    End If
    
    Select Case Index
    Case chk_��¿��ÿ��
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_��¿��ÿ��).Value = Val(GetParOriginalValue(chk, chk_��¿��ÿ��, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = (Check���쵥 And Check�ƿⵥ And Check���õ�)
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵����죬�ƿ⣬�����õ������ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_��¿��ÿ��).Value = Val(GetParOriginalValue(chk, chk_��¿��ÿ��, mrsPar))
        End If
        
        If chk(Index).Value = 1 Then
            chk(chk_��������������).Enabled = False
            If chk(chk_��������������).Value = 0 Then chk(chk_��������������).Value = 1
            
            chk(chk_��������������).Enabled = False
            If chk(chk_��������������).Value = 0 Then chk(chk_��������������).Value = 1
            
            chk(chk_�������ƿ�����).Enabled = False
            If chk(chk_�������ƿ�����).Value = 0 Then chk(chk_�������ƿ�����).Value = 1
            Call SetPrompt(lblPrompt, "�ʱ�¿��ÿ��ʱ��������Ϊ[����ʱ��ȷ��������][����ʱ��ȷ��������]")
        Else
            chk(chk_�������ƿ�����).Enabled = True
            chk(chk_��������������).Enabled = True
            chk(chk_��������������).Enabled = True
        End If
        
    Case chk_��������������
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_��������������).Value = Val(GetParOriginalValue(chk, chk_��������������, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check���쵥
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵����쵥�����ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_��������������).Value = Val(GetParOriginalValue(chk, chk_��������������, mrsPar))
        End If
     Case chk_�������ƿ�����
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_�������ƿ�����).Value = Val(GetParOriginalValue(chk, chk_�������ƿ�����, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check�ƿⵥ
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵��ƿⵥ�����ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_�������ƿ�����).Value = Val(GetParOriginalValue(chk, chk_�������ƿ�����, mrsPar))
        End If
    Case chk_��������������
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_��������������).Value = Val(GetParOriginalValue(chk, chk_��������������, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check���õ�
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵����õ������ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_��������������).Value = Val(GetParOriginalValue(chk, chk_��������������, mrsPar))
        End If
'    chk(chk_��������������).Enabled = Check���쵥
'    chk(chk_�������ƿ�����).Enabled = Check�ƿⵥ
'    chk(chk_��������������).Enabled = Check���õ�
        
    Case chk_ʱ���������ȡ�ϴ��ۼ�
        If chk(Index).Value = 1 Then
            chk(chk_ʱ�����ļӳ������).Enabled = False
            If chk(chk_ʱ�����ļӳ������).Value = 1 Then chk(chk_ʱ�����ļӳ������).Value = 0: Call SetPrompt(lblPrompt, "���������ȡ�ϴ��ۼ۷�ʽ��Ͳ���ѡ��[���ӳ��ʼ����ۼ�]��ʽ��")
            
            chk(chk_�ֶμӳ����).Enabled = False
            If chk(chk_�ֶμӳ����).Value = 1 Then chk(chk_�ֶμӳ����).Value = 0: Call SetPrompt(lblPrompt, "���������ȡ�ϴ��ۼ۷�ʽ��Ͳ���ѡ��[���ֶμӳɼ����ۼ�]��ʽ��")
        Else
            chk(chk_ʱ�����ļӳ������).Enabled = True
            chk(chk_�ֶμӳ����).Enabled = True
        End If
    Case chk_ʱ�����ļӳ������
        If chk(Index).Value = 1 Then
            chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = False
            If chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 1 Then chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 0: Call SetPrompt(lblPrompt, "��������ⰴ�ӳ��ʼ����ۼ۷�ʽ��Ͳ���ѡ��[ȡ�ϴ��ۼ�]��ʽ��")
            
            chk(chk_�ֶμӳ����).Enabled = False
            If chk(chk_�ֶμӳ����).Value = 1 Then chk(chk_�ֶμӳ����).Value = 0: Call SetPrompt(lblPrompt, "��������ⰴ�ӳ��ʼ����ۼ۷�ʽ��Ͳ���ѡ��[���ֶμӳɼ����ۼ�]��ʽ��")
        Else
            chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = True
            chk(chk_�ֶμӳ����).Enabled = True
        End If
    Case chk_�ֶμӳ����
        If chk(Index).Value = 1 Then
            chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = False
            If chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 1 Then chk(chk_ʱ���������ȡ�ϴ��ۼ�).Value = 0: Call SetPrompt(lblPrompt, "�����˰��ֶμӳɼ����ۼ۷�ʽ��Ͳ���ѡ��[ȡ�ϴ��ۼ�]��ʽ��")
            
            chk(chk_ʱ�����ļӳ������).Enabled = False
            If chk(chk_ʱ�����ļӳ������).Value = 1 Then chk(chk_ʱ�����ļӳ������).Value = 0: Call SetPrompt(lblPrompt, "�����˰��ֶμӳɼ����ۼ۷�ʽ��Ͳ���ѡ��[���ӳ��ʼ����ۼ�]��ʽ��")
        Else
            chk(chk_ʱ���������ȡ�ϴ��ۼ�).Enabled = True
            chk(chk_ʱ�����ļӳ������).Enabled = True
        End If
    Case chk_�������
        '����Ϊ����Ҫ����ʱ��Ҫ����Ƿ���δ��˵ĳ������뵥����������ܸı�
        Dim rsTemp As ADODB.Recordset
        
        If chk(Index).Value = 0 And mblnLoad = True Then
            If MsgBox("��������Ƿ����δ��˵ĳ������뵥��������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '�ù�����10.34�汾����������һ�������������ڷ�Χ������ȫ��ɨ�裬��˿��Ǵ�34�汾�޸����ڿ�ʼ
               gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
                    " Where a.���� = 19 And a.�������� Between To_Date('2014/2/20 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Mod(b.��¼״̬, 3) = 2) And Rownum < 2"
                
                
                DoEvents
                zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵ĳ������뵥")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "����δ��˵ĳ������뵥�����ܸı�˲�����", vbInformation, gstrSysName
                    chk(Index).Value = 1
                End If
            Else
                chk(Index).Value = 1
            End If
        End If
    
    End Select

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfCheck_DblClick(Index As Integer)
    With vsfCheck(Index)
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("У��") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "��" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            .TextMatrix(.Row, .Col) = "��"
        End If
    End With
    
    If Me.Visible Then
        Call SetParChange(txt, IIf(Index = 0, txt_�⹺����У��, txt_�ƻ�����У��), mrsPar, True, Get��Ӧ������У��(Index))
    End If
       
    fra����У��(Index).ForeColor = txt(IIf(Index = 0, txt_�⹺����У��, txt_�ƻ�����У��)).ForeColor
End Sub

Private Sub vsfCheck_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(txt, IIf(Index = 0, txt_�⹺����У��, txt_�ƻ�����У��), mrsPar, "", vsfCheck(Index))
End Sub

Private Sub vsf����_ChangeEdit()
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
    
    If InStr(1, vsf����.EditText, "-") <= 0 Then Exit Sub
    strID = Mid(vsf����.EditText, 1, InStr(1, vsf����.EditText, "-") - 1)
    str���� = Mid(vsf����.EditText, InStr(1, vsf����.EditText, "-") + 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ų�ѯ", strID, str����)
    If rsTemp.RecordCount > 0 Then
        With vsf����
            If .Col = m�ⷿ����.mint���ϲ��� Then
                .TextMatrix(.Row, m�ⷿ����.mint����id) = rsTemp!Id
            ElseIf .Col = m�ⷿ����.mint���Ĳֿ� Then
                .TextMatrix(.Row, m�ⷿ����.mint�ⷿid) = rsTemp!Id
            ElseIf .Col = m�ⷿ����.mint����ⷿ Then
                .TextMatrix(.Row, m�ⷿ����.mint����ⷿid) = rsTemp!Id
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'Private Sub vsf����_DblClick()
'    With vsf����
'        If .Col = m�ⷿ����.mint���� Then
'            If .TextMatrix(.Row, m�ⷿ����.mint����) = "" Then
'                .TextMatrix(.Row, m�ⷿ����.mint����) = "��"
'            Else
'                .TextMatrix(.Row, m�ⷿ����.mint����) = ""
'            End If
'        End If
'    End With
'End Sub

Private Sub vsf����_EnterCell()
    Dim strTemp As String
    
    With vsf����
        If .Col = m�ⷿ����.mint���� Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = 1 Then
            mrs����.Filter = "��������='���ϲ���'"
        ElseIf .Col = 3 Then
            mrs����.Filter = "��������='���Ŀ�'"
        ElseIf .Col = 5 Then
            mrs����.Filter = "��������='����ⷿ'"
        End If
        
'        .Clear
        strTemp = ""
        Do While Not mrs����.EOF
            strTemp = strTemp & mrs����("����") & "-" & mrs����("����") & "|"
'            .AddItem mrs����("����") & "-" & mrs����("����")
'            .ItemData(.NewIndex) = mrs����("ID")
            mrs����.MoveNext
        Loop
        .ColComboList(.Col) = strTemp
    End With
End Sub

Private Sub vsf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsf����
            If .Rows > 1 Then
                If .TextMatrix(.Row, m�ⷿ����.mint���ϲ���) <> "" Or .TextMatrix(.Row, m�ⷿ����.mint���Ĳֿ�) <> "" Or .TextMatrix(.Row, m�ⷿ����.mint����ⷿ) <> "" Then
                    If MsgBox("�Ƿ�ȷ��ɾ�����У�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                If .Rows = 2 Then
                    .Rows = 1
                    .Rows = 2
                    .Row = 1
                    .Col = 1
                Else
                    .RemoveItem .Row
                    .Col = 1
                End If
            End If
        End With
    End If
End Sub

Private Sub vsf����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf����
            If .Col = m�ⷿ����.mint���� - 1 And .TextMatrix(.Row, m�ⷿ����.mint���ϲ���) <> "" And _
                .TextMatrix(.Row, m�ⷿ����.mint���Ĳֿ�) <> "" And .TextMatrix(.Row, m�ⷿ����.mint����ⷿ) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m�ⷿ����.mint���� - 1 And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 2
            End If
        End With
    End If
End Sub

Private Sub vsf�ⷿ���_DblClick()
    With vsf�ⷿ���
        If .Col = m�ⷿ���.mint��鷽ʽ Then
            Select Case .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ)
                Case "0-�����"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "1-��飬��������"
                Case "1-��飬��������"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "2-��飬�����ֹ"
                Case "2-��飬�����ֹ"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "0-�����"
                Case Else
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "0-�����"
            End Select
            
            If .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) <> .TextMatrix(.Row, m�ⷿ���.minCheck) Then
                .Cell(flexcpForeColor, .Row, m�ⷿ���.mint��鷽ʽ, .Row, m�ⷿ���.mint��鷽ʽ) = vbRed
            Else
                .Cell(flexcpForeColor, .Row, m�ⷿ���.mint��鷽ʽ, .Row, m�ⷿ���.mint��鷽ʽ) = vbBlack
            End If
        End If
    End With
End Sub

Private Sub vsf����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strID As String
    Dim str���� As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    
    With vsf����
        strTemp = .TextMatrix(Row, Col)
        If strTemp <> "" Then
            If Col = .ColIndex("���ڿⷿ") Then
                gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str���� = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID, str����)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("���ڿⷿid")) = rsTemp!Id
                End If
            ElseIf Col = .ColIndex("�Է��ⷿ") Then
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str���� = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID, str����)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("�Է��ⷿid")) = rsTemp!Id
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf����_DblClick()
    With vsf����
        If .Col = .ColIndex("����") Then
            If .MouseRow = 0 Then Exit Sub
            .Editable = flexEDNone
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
            End Select
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsf����.Rows > 1 Then
        vsf����.RemoveItem vsf����.Row
    End If
End Sub

Private Sub vsf����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf����
            If .Col = .ColIndex("����") And .TextMatrix(.Row, .ColIndex("���ڿⷿ")) <> "" And _
                .TextMatrix(.Row, .ColIndex("�Է��ⷿ")) <> "" And .TextMatrix(.Row, .ColIndex("����")) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = .ColIndex("���ڿⷿ")
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("���ڿⷿ")
                End If
            ElseIf .Col < .ColIndex("����") And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 1
            End If
        End With
    End If
End Sub

Private Sub LoadҩƷ���ľ���()
    Const intMinDigit As Integer = 2
    Dim intMaxCost As Integer
    Dim intMaxPrice As Integer
    Dim intMaxNumber As Integer
    Dim intMaxMoney As Integer
    Dim rs As ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo ErrHandle
    'ȡ��󾫶�
    gstrSQL = "Select �ɱ���, ���ۼ�, ʵ������,���۽�� From ҩƷ�շ���¼ Where Rownum <2"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
    intMaxCost = IIf(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIf(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIf(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIf(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

    With BillҩƷ���ľ���
        .Cols = dig_Cols
        .TextMatrix(0, dig_���) = ""
        .TextMatrix(0, dig_����) = ""
        .TextMatrix(0, dig_��λ) = ""
        .TextMatrix(0, dig_�������) = "���"
        .TextMatrix(0, dig_��������) = "����"
        .TextMatrix(0, dig_���ȵ�λ) = "��λ"
        .TextMatrix(0, dig_����) = "Ŀǰ����"
        .TextMatrix(0, dig_��С����) = "��С����"
        .TextMatrix(0, dig_��󾫶�) = "��󾫶�"
        .TextMatrix(0, dig_ԭʼ����) = ""
        
        .ColWidth(dig_���) = 0
        .ColWidth(dig_����) = 0
        .ColWidth(dig_��λ) = 0
        .ColWidth(dig_�������) = 700
        .ColWidth(dig_��������) = 850
        .ColWidth(dig_���ȵ�λ) = 1000
        .ColWidth(dig_����) = 850
        .ColWidth(dig_��С����) = 850
        .ColWidth(dig_��󾫶�) = 850
        .ColWidth(dig_ԭʼ����) = 0
        
        .ColData(dig_���) = 0
        .ColData(dig_����) = 0
        .ColData(dig_��λ) = 0
        .ColData(dig_�������) = 0
        .ColData(dig_��������) = 0
        .ColData(dig_���ȵ�λ) = 0
        .ColData(dig_����) = 4
        .ColData(dig_��С����) = 0
        .ColData(dig_��󾫶�) = 0
        .ColData(dig_ԭʼ����) = 0
        
        .PrimaryCol = 0
        .MsfObj.MergeCells = flexMergeFree
        .MergeCol dig_�������, True
        .MergeCol dig_��������, True
        .Active = True
    End With
    
    'ȡĿǰ����
    gstrSQL = " Select ����, ���, ����, ��λ, Decode(���, 1, 'ҩƷ', '����') �������, Decode(����, 1, '�ɱ���', 2, '���ۼ�',3, '����','���') ��������," & _
            " Decode(���, 1, Decode(��λ, 1, '�ۼ۵�λ', 2, '���ﵥλ', 3, 'סԺ��λ',4, 'ҩ�ⵥλ','���е�λ')," & _
            " Decode(��λ, 1, 'ɢװ',2, '��װ','���е�λ')) ���ȵ�λ, Nvl(����, 0) ���� " & _
            " From ҩƷ���ľ��� where ���=2 Order By ����, ���, ����, ��λ"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
    With BillҩƷ���ľ���
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            For n = 1 To rs.RecordCount
                .TextMatrix(n, dig_���) = rs!���
                .TextMatrix(n, dig_����) = rs!����
                .TextMatrix(n, dig_��λ) = rs!��λ
                .TextMatrix(n, dig_�������) = rs!�������
                .TextMatrix(n, dig_��������) = rs!��������
                .TextMatrix(n, dig_���ȵ�λ) = rs!���ȵ�λ
                .TextMatrix(n, dig_����) = IIf(rs!���� > 4, 4, rs!����)
                .TextMatrix(n, dig_��С����) = intMinDigit
                Select Case rs!����
                    Case 1
                        .TextMatrix(n, dig_��󾫶�) = intMaxCost
                    Case 2
                        .TextMatrix(n, dig_��󾫶�) = intMaxPrice
                    Case 3
                        .TextMatrix(n, dig_��󾫶�) = intMaxNumber
                    Case 4
                        .TextMatrix(n, dig_��󾫶�) = intMaxMoney
                End Select
                .TextMatrix(n, dig_ԭʼ����) = rs!����
                .RowData(n) = rs!����
                rs.MoveNext
            Next
        End If
    End With
        
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveҩƷ���ľ���()
    Dim n As Integer
    Dim strInput As String
       
    On Error GoTo ErrHandle
    With BillҩƷ���ľ���
        If .Tag = "���޸�" Then
            For n = 1 To .Rows - 1
                strInput = strInput & "0," & _
                    .TextMatrix(n, dig_���) & "," & _
                    .TextMatrix(n, dig_����) & "," & _
                    .TextMatrix(n, dig_��λ) & "," & _
                    .TextMatrix(n, dig_����) & ";"
            Next
        
            gstrSQL = "ZL_ҩƷ���ľ���_Update('" & strInput & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub BillҩƷ���ľ���_EnterCell(Row As Long, Col As Long)
    With BillҩƷ���ľ���
        If Col = dig_���� Then
            .TxtCheck = True
            .TextMask = "123456789"
            .MaxLength = 1
        End If
    End With
End Sub

Private Sub BillҩƷ���ľ���_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With BillҩƷ���ľ���
        If .Col = dig_���� Then
            If .Text = "" Then Exit Sub
            
            .Text = Val(.Text)
            strKey = .Text
            
            If Val(strKey) > .TextMatrix(.Row, dig_��󾫶�) Or Val(strKey) < .TextMatrix(.Row, dig_��С����) Then
                MsgBox "���ȳ�������Χ��", vbInformation, gstrSysName
                .Text = .RowData(.Row)
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            
            .TextMatrix(.Row, .Col) = strKey
            .RowData(.Row) = Val(strKey)
            
            .Tag = "���޸�"
        End If
    End With
End Sub
Private Sub Load���ݻ��ڿ���()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo ErrHandle
    intAllItems = UBound(Split(cst������Ŀ, ",")) + 1
    
    With vsf���ݻ��ڿ���
        .Rows = 4
        .Cols = 2 + intAllItems
        .FixedRows = 1
        .FixedCols = 2
        .RowHeightMin = 400
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
                        
        .ColWidth(0) = 820
        .ColWidth(1) = 820
                        
        For n = 0 To UBound(Split(cst������Ŀ, ","))
            .TextMatrix(0, n + 2) = Split(cst������Ŀ, ",")(n)
            .ColWidth(n + 2) = 820
            .ColAlignment(n + 2) = flexAlignCenterCenter
        Next
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        .TextMatrix(1, 0) = "�����⹺"
        .TextMatrix(2, 0) = "�����⹺"
        .TextMatrix(3, 0) = "�����⹺"

        .TextMatrix(1, 1) = "�˲�"
        .TextMatrix(2, 1) = "���"
        .TextMatrix(3, 1) = "�������"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select ����,����,���� From ���ݻ��ڿ��� where ����=15 Order By ����, ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���ݻ��ڿ���")
        
        If Not rsTmp.EOF Then
            For n = 1 To rsTmp.RecordCount
                For m = 2 To intAllItems + 1
                    If InStr(1, "," & rsTmp!���� & ",", Trim(.TextMatrix(0, m))) > 0 Then
                        Select Case rsTmp!����
                            Case ����.�����⹺
                                Select Case rsTmp!����
                                    Case ����.�˲�
                                        .TextMatrix(1, m) = "��"
                                    Case ����.���
                                        .TextMatrix(2, m) = "��"
                                    Case ����.�������
                                        .TextMatrix(3, m) = "��"
                                End Select
                        End Select
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf���ݻ��ڿ���_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsf���ݻ��ڿ���.Tag = "" Then vsf���ݻ��ڿ���.Tag = "���޸�"
End Sub

Private Sub vsf���ݻ��ڿ���_DblClick()
    With vsf���ݻ��ڿ���
        If .Row < 1 Then Exit Sub
        If .Col < 2 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "��" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            '�˲�ʱ�����޸�"��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"
            If .TextMatrix(.Row, 1) = "�˲�" And InStr(1, "��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���", .TextMatrix(0, .Col)) > 0 Then Exit Sub

            .TextMatrix(.Row, .Col) = "��"

        End If
        
    End With
End Sub

Private Sub Save���ݻ��ڿ���()
    Dim n As Integer
    Dim m As Integer
    Dim strInput As String
    Dim int���� As Integer
    Dim int���� As Integer
    Dim str���� As String
    
    On Error GoTo ErrHandle
    With vsf���ݻ��ڿ���
        If .Tag = "���޸�" Then
            For n = 1 To .Rows - 1
                Select Case .TextMatrix(n, 0)
                    Case "�����⹺"
                        int���� = ����.�����⹺
                End Select
                
                Select Case .TextMatrix(n, 1)
                    Case "�˲�"
                        int���� = ����.�˲�
                    Case "���"
                        int���� = ����.���
                    Case "�������"
                        int���� = ����.�������
                End Select
                
                str���� = ""
                For m = 2 To .Cols - 1
                    If .TextMatrix(n, m) = "��" Then
                        str���� = str���� & IIf(str���� <> "", ",", "") & .TextMatrix(0, m)
                    End If
                Next
                
                If str���� <> "" Then
                    strInput = strInput & IIf(strInput <> "", ";", "") & int���� & "," & int���� & "," & str����
                End If
            Next
        
            gstrSQL = "Zl_���ݻ��ڿ���_Update('" & strInput & "'," & ����.�����⹺ & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check�ƿⵥ() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
        " Where a.���� = 19 And a.�������� > Sysdate - 90 And Exists " & _
        " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Nvl(b.��ҩ��ʽ,0) <> 1) And Rownum < 2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵��ƿⵥ")
    
    Check�ƿⵥ = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check���쵥() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
        " Where a.���� = 19 And a.�������� > Sysdate - 90 And Exists " & _
        " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Nvl(b.��ҩ��ʽ,0) = 1) And Rownum < 2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵����쵥")
    
    Check���쵥 = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check���õ�() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ Where ���� = 20 And �������� > Sysdate - 90 And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵����õ�")
    
    Check���õ� = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

