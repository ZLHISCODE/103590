VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmParMedicine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ��������"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14040
   Icon            =   "frmParMedicine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   14040
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabDesign 
      Height          =   8415
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   14
      Tab             =   6
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "ͨ��(&0)"
      TabPicture(0)   =   "frmParMedicine.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ŀ¼(&1)"
      TabPicture(1)   =   "frmParMedicine.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "���(&2)"
      TabPicture(2)   =   "frmParMedicine.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "�ڿ�(&3)"
      TabPicture(3)   =   "frmParMedicine.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "����(&4)"
      TabPicture(4)   =   "frmParMedicine.frx":68C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "����(&5)"
      TabPicture(5)   =   "frmParMedicine.frx":68DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "����(&6)"
      TabPicture(6)   =   "frmParMedicine.frx":68FA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "picPar(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "��ҩ(&11)"
      TabPicture(7)   =   "frmParMedicine.frx":6916
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(11)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "����(&12)"
      TabPicture(8)   =   "frmParMedicine.frx":6932
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(12)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "��λ(&13)"
      TabPicture(9)   =   "frmParMedicine.frx":694E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "picPar(13)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "����(&14)"
      TabPicture(10)  =   "frmParMedicine.frx":696A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "picPar(14)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "���(&15)"
      TabPicture(11)  =   "frmParMedicine.frx":6986
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "picPar(15)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "����(&16)"
      TabPicture(12)  =   "frmParMedicine.frx":69A2
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "picPar(16)"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "�������(&7)"
      TabPicture(13)  =   "frmParMedicine.frx":69BE
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "picPar(7)"
      Tab(13).ControlCount=   1
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   14
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   41
         Top             =   600
         Width           =   10455
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   6975
            Index           =   4
            Left            =   6225
            TabIndex        =   169
            Top             =   360
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   12303
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   3
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
         Begin ZL9BillEdit.BillEdit Bill 
            Height          =   6975
            Index           =   3
            Left            =   120
            TabIndex        =   170
            Top             =   360
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   12303
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   3
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ÿⷿ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Index           =   23
            Left            =   6225
            TabIndex        =   172
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�֮ⷿ���������"
            ForeColor       =   &H00000080&
            Height          =   180
            Index           =   33
            Left            =   120
            TabIndex        =   171
            Top             =   120
            Width           =   1440
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
         TabIndex        =   40
         Top             =   600
         Width           =   10455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf�ⷿ������λ 
            Height          =   6975
            Left            =   240
            TabIndex        =   167
            Top             =   360
            Width           =   6585
            _ExtentX        =   11615
            _ExtentY        =   12303
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483631
            AllowBigSelection=   0   'False
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   1
            FormatString    =   "ҩƷ�ⷿ|�ۼ۵�λ|���ﵥλ|סԺ��λ|ҩ�ⵥλ"
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label lblUnits 
            Caption         =   "ҩƷ�ⷿ�ļ�����λ��˫��������ã�"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   168
            Top             =   120
            Width           =   3855
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
         TabIndex        =   39
         Top             =   600
         Width           =   10455
         Begin ZL9BillEdit.BillEdit BillҩƷ���ľ��� 
            Height          =   6180
            Left            =   240
            TabIndex        =   164
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ�������ã�����װ��λ�����ü۸���������¼��ľ��ȣ�������С��λ����"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   240
            TabIndex        =   166
            Top             =   120
            Width           =   6480
         End
         Begin VB.Label Label23 
            Caption         =   $"frmParMedicine.frx":69DA
            ForeColor       =   &H00000080&
            Height          =   720
            Left            =   240
            TabIndex        =   165
            Top             =   6600
            Width           =   7995
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
         TabIndex        =   38
         Top             =   600
         Width           =   10455
         Begin ZL9BillEdit.BillEdit Billҩ����ҩ���� 
            Height          =   6870
            Left            =   240
            TabIndex        =   162
            Top             =   390
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   12118
            CellAlignment   =   9
            Text            =   ""
            TextMatrix0     =   ""
            MaxDate         =   2958465
            MinDate         =   -53688
            Value           =   36395
            Cols            =   3
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩ����ҩ����"
            Height          =   180
            Index           =   34
            Left            =   315
            TabIndex        =   163
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7605
         Index           =   6
         Left            =   0
         ScaleHeight     =   7575
         ScaleWidth      =   11265
         TabIndex        =   37
         Top             =   600
         Width           =   11295
         Begin TabDlg.SSTab TabPiva 
            Height          =   7320
            Left            =   120
            TabIndex        =   186
            Top             =   120
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   12912
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "��������(&1)"
            TabPicture(0)   =   "frmParMedicine.frx":6ABB
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra��ҩ����"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fra��Һҽ����Ч"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "frmParMedicine"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "frmType"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "��������(&2)"
            TabPicture(1)   =   "frmParMedicine.frx":6AD7
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "picPRI"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "fra��Һ����"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "fra(0)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "frmMoney"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "�Ա�ҩ����(&3)"
            TabPicture(2)   =   "frmParMedicine.frx":6AF3
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fra�Ա�ҩ�嵥����"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame fra�Ա�ҩ�嵥���� 
               Caption         =   " �Ա�ҩ�嵥���� "
               ForeColor       =   &H00800000&
               Height          =   6855
               Left            =   -74880
               TabIndex        =   264
               Top             =   360
               Width           =   10095
               Begin VSFlex8Ctl.VSFlexGrid vsf�Ա�ҩ�嵥 
                  Height          =   6255
                  Left            =   120
                  TabIndex        =   265
                  Top             =   480
                  Width           =   9825
                  _cx             =   17330
                  _cy             =   11033
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
                  BackColorSel    =   16771280
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   -2147483633
                  BackColorAlternate=   -2147483643
                  GridColor       =   10329501
                  GridColorFixed  =   10329501
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   16777215
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
                  Cols            =   3
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"frmParMedicine.frx":6B0F
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
                  AccessibleDescription=   "200"
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "  ��������ҩƷ�ھ������Ĳ������Ա�ҩ������£�����ͨ�����Ա�ҩ�ķ�ʽ������ҩƷ�������������ġ�"
                  ForeColor       =   &H00000080&
                  Height          =   180
                  Left            =   120
                  TabIndex        =   266
                  Top             =   240
                  Width           =   8460
               End
            End
            Begin VB.Frame frmType 
               Caption         =   " ҽ��ִ������ѡ��"
               ForeColor       =   &H00800000&
               Height          =   705
               Left            =   4680
               TabIndex        =   237
               Top             =   480
               Width           =   6015
               Begin VB.CheckBox chk 
                  Caption         =   "��ȡҩ"
                  Height          =   255
                  Index           =   62
                  Left            =   1740
                  TabIndex        =   240
                  Top             =   330
                  Width           =   885
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Ա�ҩ"
                  Height          =   255
                  Index           =   61
                  Left            =   240
                  TabIndex        =   239
                  Top             =   330
                  Width           =   885
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Ժ��ҩ"
                  Height          =   255
                  Index           =   63
                  Left            =   3240
                  TabIndex        =   238
                  Top             =   330
                  Width           =   1125
               End
            End
            Begin VB.Frame frmParMedicine 
               Caption         =   " �������� "
               ForeColor       =   &H00800000&
               Height          =   5775
               Left            =   120
               TabIndex        =   217
               Top             =   1200
               Width           =   4455
               Begin VB.CheckBox chk 
                  Caption         =   "����ҩƷ��ҩƷ����ָ������"
                  Height          =   255
                  Index           =   51
                  Left            =   240
                  TabIndex        =   226
                  Top             =   2505
                  Width           =   3135
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Һ�������Σ�ҩƷ��������"
                  Height          =   255
                  Index           =   47
                  Left            =   240
                  TabIndex        =   225
                  Top             =   1935
                  Width           =   3375
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����ɨ��һ���Զ�����"
                  Height          =   255
                  Index           =   46
                  Left            =   240
                  TabIndex        =   224
                  Top             =   1365
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����ֹ��������Σ���ҩӡǩ���ڣ�"
                  Height          =   255
                  Index           =   33
                  Left            =   240
                  TabIndex        =   223
                  Top             =   240
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����������״̬����ҩӡǩ����ҩ���ڣ�"
                  Height          =   255
                  Index           =   34
                  Left            =   240
                  TabIndex        =   222
                  Top             =   813
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���÷Ѱ�������ȡ(һ������һ��ֻ��һ������)"
                  Height          =   255
                  Index           =   49
                  Left            =   240
                  TabIndex        =   221
                  Top             =   3075
                  Width           =   4095
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Ժ���˲������÷�"
                  Height          =   255
                  Index           =   50
                  Left            =   240
                  TabIndex        =   220
                  Top             =   3645
                  Width           =   1935
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���ҩƷ�ڷ��ͻ�����ȡ���÷�"
                  Height          =   255
                  Index           =   59
                  Left            =   240
                  TabIndex        =   219
                  Top             =   4230
                  Width           =   3495
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա"
                  Height          =   255
                  Index           =   60
                  Left            =   240
                  TabIndex        =   218
                  Top             =   4800
                  Width           =   3495
               End
            End
            Begin VB.PictureBox picPRI 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               Height          =   2055
               Left            =   -72000
               ScaleHeight     =   2055
               ScaleWidth      =   2535
               TabIndex        =   211
               Top             =   7680
               Visible         =   0   'False
               Width           =   2535
               Begin VB.CommandButton cmdYes 
                  Height          =   360
                  Left            =   720
                  Picture         =   "frmParMedicine.frx":6B9D
                  Style           =   1  'Graphical
                  TabIndex        =   213
                  Top             =   1560
                  Width           =   810
               End
               Begin VB.CommandButton cmdNO 
                  Height          =   360
                  Left            =   1560
                  Picture         =   "frmParMedicine.frx":D3EF
                  Style           =   1  'Graphical
                  TabIndex        =   212
                  Top             =   1560
                  Width           =   810
               End
               Begin MSComctlLib.ListView lvwPRI 
                  Height          =   1305
                  Left            =   120
                  TabIndex        =   214
                  ToolTipText     =   "˫���򰴻س���ȷ��"
                  Top             =   120
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   2302
                  View            =   2
                  Arrange         =   1
                  LabelEdit       =   1
                  MultiSelect     =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  Icons           =   "imgLvwSel"
                  SmallIcons      =   "imgLvwSel"
                  ColHdrIcons     =   "imgLvwSel"
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   0
                  NumItems        =   1
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "����"
                     Object.Width           =   3528
                  EndProperty
               End
            End
            Begin VB.Frame fra��Һ���� 
               Caption         =   " ����Һ�������ķ�ҩ�Ĳ��˿��� "
               ForeColor       =   &H00800000&
               Height          =   6735
               Left            =   -69720
               TabIndex        =   205
               Top             =   360
               Width           =   5175
               Begin VB.CheckBox chk��Դ���� 
                  Caption         =   "������Դ��������"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   209
                  Top             =   720
                  Width           =   2295
               End
               Begin VB.ListBox lst 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000012&
                  Height          =   4440
                  IMEMode         =   3  'DISABLE
                  Index           =   0
                  Left            =   240
                  Style           =   1  'Checkbox
                  TabIndex        =   208
                  Top             =   1020
                  Width           =   4785
               End
               Begin VB.CommandButton cmdlst��Һ���ķ�ҩ���˿��� 
                  Caption         =   "ȫѡ"
                  Height          =   350
                  Index           =   0
                  Left            =   2760
                  TabIndex        =   207
                  Top             =   6240
                  Width           =   1100
               End
               Begin VB.CommandButton cmdlst��Һ���ķ�ҩ���˿��� 
                  Caption         =   "ȫ��"
                  Height          =   350
                  Index           =   1
                  Left            =   3960
                  TabIndex        =   206
                  Top             =   6240
                  Width           =   1100
               End
               Begin VB.Label lbl��Դ���� 
                  Caption         =   "  ����ʱ��ѡ��������Һҽ������ʱ������˵����ڲ���û��ѡ���򲻻������Һ���ݡ�"
                  ForeColor       =   &H00000080&
                  Height          =   420
                  Left            =   240
                  TabIndex        =   210
                  Top             =   360
                  Width           =   4560
               End
            End
            Begin VB.Frame fra 
               Caption         =   " �����ҩ;��ѡ��"
               ForeColor       =   &H00800000&
               Height          =   3135
               Index           =   0
               Left            =   -74760
               TabIndex        =   201
               Top             =   360
               Width           =   4935
               Begin VB.CheckBox chk��ҩ;�� 
                  Caption         =   "������Һ��ҩ;������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   203
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.ListBox lst 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000012&
                  Height          =   870
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   120
                  Style           =   1  'Checkbox
                  TabIndex        =   202
                  Top             =   1155
                  Width           =   4560
               End
               Begin VB.Label lbl��ҩ;�� 
                  Caption         =   "  ����ʱ��ѡ��������Һ��ĸ�ҩ;������Һҽ������ʱ���ҽ���ĸ�ҩ;��û��ѡ���򲻻������Һ���ݡ�"
                  ForeColor       =   &H00000080&
                  Height          =   540
                  Left            =   120
                  TabIndex        =   204
                  Top             =   240
                  Width           =   4080
               End
            End
            Begin VB.Frame frmMoney 
               Caption         =   "���÷�����"
               ForeColor       =   &H00800000&
               Height          =   3495
               Left            =   -74760
               TabIndex        =   199
               Top             =   3600
               Width           =   4935
               Begin TabDlg.SSTab tabPrice 
                  Height          =   2175
                  Left            =   120
                  TabIndex        =   252
                  Top             =   720
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   3836
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   2
                  TabHeight       =   520
                  TabCaption(0)   =   "��ҩ����"
                  TabPicture(0)   =   "frmParMedicine.frx":D539
                  Tab(0).ControlEnabled=   -1  'True
                  Tab(0).Control(0)=   "VSFPrice"
                  Tab(0).Control(0).Enabled=   0   'False
                  Tab(0).ControlCount=   1
                  TabCaption(1)   =   "��ҩ;��(ֻ֧�־���Ӫ������)"
                  TabPicture(1)   =   "frmParMedicine.frx":D555
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "VSFPrice_��ҩ;��"
                  Tab(1).ControlCount=   1
                  Begin VSFlex8Ctl.VSFlexGrid VSFPrice 
                     Height          =   1635
                     Left            =   120
                     TabIndex        =   253
                     Top             =   360
                     Width           =   3480
                     _cx             =   6138
                     _cy             =   2884
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
                     BackColorSel    =   16771280
                     ForeColorSel    =   -2147483640
                     BackColorBkg    =   -2147483633
                     BackColorAlternate=   -2147483643
                     GridColor       =   10329501
                     GridColorFixed  =   10329501
                     TreeColor       =   -2147483632
                     FloodColor      =   192
                     SheetBorder     =   16777215
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
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   0
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"frmParMedicine.frx":D571
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
                     AccessibleDescription=   "200"
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
                  Begin VSFlex8Ctl.VSFlexGrid VSFPrice_��ҩ;�� 
                     Height          =   1395
                     Left            =   -74760
                     TabIndex        =   254
                     Top             =   360
                     Width           =   3600
                     _cx             =   6350
                     _cy             =   2461
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
                     BackColorSel    =   16771280
                     ForeColorSel    =   -2147483640
                     BackColorBkg    =   -2147483633
                     BackColorAlternate=   -2147483643
                     GridColor       =   10329501
                     GridColorFixed  =   10329501
                     TreeColor       =   -2147483632
                     FloodColor      =   192
                     SheetBorder     =   16777215
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
                     Cols            =   4
                     FixedRows       =   1
                     FixedCols       =   0
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"frmParMedicine.frx":D61D
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
                     AccessibleDescription=   "200"
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
               End
               Begin VB.CommandButton cmdNext 
                  Caption         =   "����(&N)"
                  Height          =   350
                  Left            =   2280
                  TabIndex        =   244
                  Top             =   3000
                  Width           =   1100
               End
               Begin VB.CommandButton cmdLast 
                  Caption         =   "����(&S)"
                  Enabled         =   0   'False
                  Height          =   350
                  Left            =   600
                  TabIndex        =   243
                  Top             =   3000
                  Width           =   1100
               End
               Begin VB.Label lblprice 
                  Caption         =   "������Һ����ҩ���Ͷ�Ӧ���շ���Ŀ������ҩ��ʱ��������õĹ�������շѣ������Ȱ���ҩ;����ʽ��ȡ���÷�"
                  DragMode        =   1  'Automatic
                  ForeColor       =   &H00000080&
                  Height          =   420
                  Left            =   240
                  TabIndex        =   200
                  Top             =   240
                  Width           =   4560
               End
            End
            Begin VB.Frame fra��Һҽ����Ч 
               Caption         =   " ������Һ�������ĵ�ҽ����Ч"
               ForeColor       =   &H00800000&
               Height          =   705
               Left            =   120
               TabIndex        =   195
               Top             =   480
               Width           =   4455
               Begin VB.OptionButton opt��Һҽ����Ч 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   1
                  Left            =   240
                  TabIndex        =   198
                  Top             =   330
                  Width           =   680
               End
               Begin VB.OptionButton opt��Һҽ����Ч 
                  Caption         =   "����"
                  Height          =   180
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   197
                  Top             =   330
                  Width           =   680
               End
               Begin VB.OptionButton opt��Һҽ����Ч 
                  Caption         =   "����������"
                  Height          =   180
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   196
                  Top             =   330
                  Value           =   -1  'True
                  Width           =   1200
               End
            End
            Begin VB.Frame fra��ҩ���� 
               Caption         =   " ��ҩ���̡���������"
               ForeColor       =   &H00800000&
               Height          =   5775
               Left            =   4680
               TabIndex        =   187
               Top             =   1200
               Width           =   6015
               Begin VB.CheckBox chk 
                  Caption         =   "��Һ����ҩ���ٴ�������ı���״̬"
                  Height          =   255
                  Index           =   64
                  Left            =   240
                  TabIndex        =   249
                  Top             =   5400
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "���췢�͵�ҽ����������Һ��ȫ������������"
                  Height          =   255
                  Index           =   57
                  Left            =   240
                  TabIndex        =   216
                  Top             =   4849
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�Զ�����ʱ��Һ��������ֻ���������α䶯"
                  Height          =   255
                  Index           =   56
                  Left            =   240
                  TabIndex        =   215
                  Top             =   3753
                  Width           =   5535
               End
               Begin VB.CheckBox chk 
                  Caption         =   "ͬһ��Һ�������ϴ���ҩ����"
                  Height          =   255
                  Index           =   39
                  Left            =   240
                  TabIndex        =   194
                  Top             =   773
                  Width           =   2655
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�������Ĳ����յľ���Ӫ��ҽ���ڲ�������"
                  Height          =   255
                  Index           =   42
                  Left            =   240
                  TabIndex        =   193
                  Top             =   1869
                  Width           =   4125
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Һ��Һ����ҩ��������������"
                  Height          =   255
                  Index           =   41
                  Left            =   240
                  TabIndex        =   192
                  Top             =   1321
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "��Һ���������״�ִ�е�ҽ����Ҫ�������"
                  Height          =   240
                  Index           =   83
                  Left            =   240
                  TabIndex        =   191
                  Top             =   240
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�����Զ������������Զ������󣬽����ٱ����ϴ����Σ�"
                  Height          =   255
                  Index           =   54
                  Left            =   240
                  TabIndex        =   190
                  Top             =   3205
                  Width           =   5535
               End
               Begin VB.CheckBox chk 
                  Caption         =   "����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����"
                  Height          =   495
                  Index           =   52
                  Left            =   240
                  TabIndex        =   189
                  Top             =   2417
                  Width           =   5655
               End
               Begin VB.CheckBox chk 
                  Caption         =   "�������û�ҩ������Һ��������"
                  Height          =   255
                  Index           =   55
                  Left            =   240
                  TabIndex        =   188
                  Top             =   4301
                  Width           =   5535
               End
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   5
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   36
         Top             =   600
         Width           =   10455
         Begin VB.CheckBox chk 
            Caption         =   "�Ƿ�������ʾܾ�"
            Height          =   180
            Index           =   53
            Left            =   240
            TabIndex        =   245
            Top             =   3000
            Width           =   4095
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩ��������Ĭ��Ϊ��ҩ״̬"
            Height          =   180
            Index           =   43
            Left            =   240
            TabIndex        =   242
            Top             =   2640
            Width           =   4095
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩʱ���ҽ��"
            Height          =   180
            Index           =   38
            Left            =   240
            TabIndex        =   236
            Top             =   2280
            Width           =   4095
         End
         Begin VB.TextBox txtud 
            Alignment       =   2  'Center
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   1
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   158
            Text            =   "1"
            Top             =   120
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ʾ�ⷿ��λ�����������ʾ"
            Height          =   180
            Index           =   27
            Left            =   240
            TabIndex        =   157
            Top             =   825
            Width           =   2745
         End
         Begin VB.Frame fraǩ�� 
            Caption         =   " ҩ����Աǩ������"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   240
            TabIndex        =   154
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ��ǩ��"
               Height          =   255
               Index           =   31
               Left            =   150
               TabIndex        =   156
               Top             =   285
               Width           =   1485
            End
            Begin VB.CheckBox chk 
               Caption         =   "��ҩ��ǩ��"
               Height          =   255
               Index           =   32
               Left            =   1710
               TabIndex        =   155
               Top             =   285
               Width           =   1485
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Ƿ��Զ�ȱҩ���"
            Height          =   180
            Index           =   25
            Left            =   240
            TabIndex        =   153
            Top             =   480
            Width           =   1845
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   225
            Index           =   3
            Left            =   2175
            MaxLength       =   2
            TabIndex        =   152
            Text            =   "5"
            Top             =   1170
            Width           =   285
         End
         Begin VB.CheckBox chk�Զ�ˢ�� 
            Caption         =   "�Զ�ˢ��δ��ҩ�嵥"
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   1155
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩʱ������ҩ���ʼ�¼"
            Height          =   180
            Index           =   29
            Left            =   240
            TabIndex        =   150
            Top             =   1575
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩ����ʱ������˳�Ժ���˵���������"
            Height          =   180
            Index           =   30
            Left            =   240
            TabIndex        =   149
            Top             =   1920
            Width           =   4095
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   1
            Left            =   1740
            TabIndex        =   159
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtud(1)"
            BuddyDispid     =   196645
            BuddyIndex      =   1
            OrigLeft        =   1920
            OrigTop         =   360
            OrigRight       =   2175
            OrigBottom      =   660
            Max             =   7
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl��ѯ���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ĭ�ϲ�ѯ����"
            Height          =   180
            Left            =   240
            TabIndex        =   161
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   2520
            TabIndex        =   160
            Top             =   1200
            Width           =   480
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   4
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   35
         Top             =   600
         Width           =   10455
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   2
            Left            =   2280
            TabIndex        =   247
            Top             =   2880
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtud(2)"
            BuddyDispid     =   196645
            BuddyIndex      =   2
            OrigLeft        =   480
            OrigTop         =   5640
            OrigRight       =   735
            OrigBottom      =   5940
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtud 
            Alignment       =   2  'Center
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   2
            Left            =   1980
            MaxLength       =   2
            TabIndex        =   246
            Text            =   "1"
            Top             =   2880
            Width           =   300
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩʱ��δ�շѵĵ��ݽ����շ�"
            Height          =   225
            Index           =   23
            Left            =   240
            TabIndex        =   228
            Top             =   2040
            Width           =   2820
         End
         Begin VB.Frame fraδ�շѷ�ҩ 
            Caption         =   " δ�շѻ����ʱ��ҩ"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   120
            TabIndex        =   144
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "����δ��˵ļ��ʴ�����ҩ"
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   147
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3345
            End
            Begin VB.CheckBox chk 
               Caption         =   "����δ�շѵ����ﻮ�۴�����ҩ"
               Height          =   195
               Index           =   58
               Left            =   120
               TabIndex        =   146
               Top             =   960
               Width           =   2880
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ŀִ��ǰ���շѻ����"
               Height          =   195
               Index           =   74
               Left            =   720
               TabIndex        =   145
               Top             =   0
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.Label lblδ�շѷ�ҩ 
               Caption         =   "  �������������һ��ͨ����""ִ��ǰ�������շѻ��ȼ������""��������ﲡ�˷�ҩʱ�����²�����ʧЧ��"
               ForeColor       =   &H00000080&
               Height          =   615
               Left            =   120
               TabIndex        =   148
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame fra 
            Caption         =   " ��ҩ���ڶ�̬���� "
            ForeColor       =   &H00800000&
            Height          =   825
            Index           =   3
            Left            =   120
            TabIndex        =   141
            Top             =   5160
            Width           =   3975
            Begin VB.OptionButton opt��ҩ���� 
               Caption         =   "��æ��ʽ"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   143
               Top             =   360
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton opt��ҩ���� 
               Caption         =   "ƽ����ʽ"
               Height          =   210
               Index           =   1
               Left            =   1560
               TabIndex        =   142
               Top             =   360
               Width           =   1020
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "�շѻ����ָ��ҩ��ʱ�޶�ҩƷ���"
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   140
            Top             =   435
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ�շ���ɺ��Զ���ҩ"
            Height          =   225
            Index           =   17
            Left            =   240
            TabIndex        =   139
            Top             =   120
            Width           =   2280
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   4
            Left            =   1740
            TabIndex        =   136
            Text            =   "500"
            Top             =   2400
            Width           =   705
         End
         Begin VB.Frame fra��֤��ʽ 
            Caption         =   " ҩ����Ա��֤��ʽ "
            ForeColor       =   &H00800000&
            Height          =   765
            Left            =   5280
            TabIndex        =   133
            Top             =   4440
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "У����ҩ��"
               Height          =   195
               Index           =   20
               Left            =   240
               TabIndex        =   135
               Top             =   360
               Width           =   1200
            End
            Begin VB.CheckBox chk 
               Caption         =   "У�鷢ҩ��"
               Height          =   195
               Index           =   24
               Left            =   1830
               TabIndex        =   134
               Top             =   360
               Width           =   1200
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩʱ�Զ������ʷ�������"
            Height          =   225
            Index           =   13
            Left            =   240
            TabIndex        =   132
            Top             =   750
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩʱˢ���￨��֤"
            Height          =   225
            Index           =   16
            Left            =   240
            TabIndex        =   131
            Top             =   1050
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷҽ��������ʱ�����"
            Height          =   225
            Index           =   18
            Left            =   240
            TabIndex        =   130
            Top             =   1365
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ò���ʵ��ȡҩȷ��ģʽ"
            Height          =   225
            Index           =   19
            Left            =   240
            TabIndex        =   129
            Top             =   1680
            Width           =   2460
         End
         Begin VB.Frame fraSetColor 
            Caption         =   "  ������ɫ����"
            ForeColor       =   &H00800000&
            Height          =   4125
            Left            =   5280
            TabIndex        =   113
            Top             =   120
            Width           =   3915
            Begin VB.CommandButton cmdDefaultColor 
               BackColor       =   &H00000000&
               Caption         =   "�ָ�Ĭ����ɫ(&R)"
               Height          =   300
               Left            =   480
               MaskColor       =   &H00000000&
               TabIndex        =   121
               Top             =   3600
               Width           =   2175
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   120
               Top             =   3090
               Width           =   1335
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0FF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   119
               Top             =   2625
               Width           =   1335
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   118
               Top             =   2175
               Width           =   1335
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   117
               Top             =   1710
               Width           =   1335
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   116
               Top             =   1260
               Width           =   1335
            End
            Begin VB.PictureBox pic������ɫ 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   1080
               ScaleHeight     =   225
               ScaleWidth      =   1305
               TabIndex        =   115
               Top             =   810
               Width           =   1335
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   2
               Left            =   2520
               TabIndex        =   114
               Text            =   "�����ԭʼֵ"
               Top             =   3120
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "˵����˫����ɫ��ǩ���崦����ɫ��"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   240
               TabIndex        =   128
               Top             =   360
               Width           =   2880
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Index           =   5
               Left            =   240
               TabIndex        =   127
               Top             =   3120
               Width           =   360
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "����I��"
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   126
               Top             =   2670
               Width           =   630
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "����II��"
               Height          =   180
               Index           =   3
               Left            =   240
               TabIndex        =   125
               Top             =   2205
               Width           =   720
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   124
               Top             =   1755
               Width           =   360
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   123
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl�������� 
               AutoSize        =   -1  'True
               Caption         =   "��ͨ"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   122
               Top             =   840
               Width           =   360
            End
         End
         Begin VB.Label lbl��ѯδ��ҩ�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ѯδ��ҩ��������"
            Height          =   180
            Left            =   240
            TabIndex        =   248
            Top             =   2940
            Width           =   1620
         End
         Begin VB.Label lblMax 
            AutoSize        =   -1  'True
            Caption         =   "�󴦷���˱�׼ֵ"
            Height          =   180
            Left            =   240
            TabIndex        =   138
            Top             =   2460
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ԫ"
            Height          =   180
            Left            =   2520
            TabIndex        =   137
            Top             =   2460
            Width           =   180
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
         TabIndex        =   34
         Top             =   600
         Width           =   10455
         Begin VB.Frame fra���ۿ��� 
            Caption         =   " ���ۿ���"
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   120
            TabIndex        =   110
            Top             =   2040
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "�ɱ��۰��ⷿ���ε���"
               Height          =   255
               Index           =   66
               Left            =   120
               TabIndex        =   251
               Top             =   660
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "�³ɱ��ۡ����ۼ۳����޼�ʱ��ʾ"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   112
               Top             =   960
               Width           =   3090
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷ�����ε���"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   111
               Top             =   360
               Width           =   2010
            End
         End
         Begin VB.Frame fra�̵���� 
            Caption         =   " �̵����"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   120
            TabIndex        =   106
            Top             =   120
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "�̿��������������"
               Height          =   255
               Index           =   65
               Left            =   120
               TabIndex        =   250
               Top             =   1320
               Width           =   1920
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����̵�ͣ��ҩƷ"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   109
               Top             =   1005
               Width           =   1920
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ҩƷ�������"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   108
               Top             =   682
               Width           =   2040
            End
            Begin VB.CheckBox chk 
               Caption         =   "�����̵�û�����ô洢�ⷿ��ҩƷ"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   107
               Top             =   360
               Width           =   3360
            End
         End
         Begin VB.Frame fra�������� 
            Caption         =   " ��������"
            ForeColor       =   &H00800000&
            Height          =   1365
            Left            =   120
            TabIndex        =   103
            Top             =   3600
            Width           =   4245
            Begin VB.CheckBox chk 
               Caption         =   "ҩƷ�����������ʱͬ�����ٿ��"
               Height          =   180
               Index           =   12
               Left            =   120
               TabIndex        =   104
               Top             =   360
               Width           =   3495
            End
            Begin VB.Label Label4 
               Caption         =   "�����ѡ��ѡ��൱������˺��Զ�����������������Ҫʵ�ָù��ܣ�����ȷ���������������������������"
               ForeColor       =   &H00000080&
               Height          =   540
               Left            =   120
               TabIndex        =   105
               Top             =   600
               Width           =   3780
            End
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
         TabIndex        =   33
         Top             =   600
         Width           =   10455
         Begin VB.Frame fra�ƿ����̿��� 
            Caption         =   " ����ҵ�����̿���"
            ForeColor       =   &H00800000&
            Height          =   2445
            Left            =   120
            TabIndex        =   99
            Top             =   4800
            Width           =   4965
            Begin VB.CheckBox chk 
               Caption         =   "���ó���ʱ����Ҫ���������,����˳���"
               Height          =   180
               Index           =   40
               Left            =   120
               TabIndex        =   241
               Top             =   2160
               Width           =   4095
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ҵ��ҩƷ��������д���ⵥ"
               Height          =   255
               Index           =   37
               Left            =   120
               TabIndex        =   232
               Top             =   720
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ƿ�ҵ��ҩƷ��������д���ⵥ"
               Height          =   255
               Index           =   26
               Left            =   120
               TabIndex        =   231
               Top             =   480
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ƿ���ȷ����ʱ����¼��������"
               Height          =   255
               Index           =   76
               Left            =   120
               TabIndex        =   230
               Top             =   1050
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "����ҵ��ҩƷ��������д���ⵥ"
               Height          =   255
               Index           =   75
               Left            =   120
               TabIndex        =   229
               Top             =   240
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ƿ�ʱ��Ҫ��ҩ�����͡�������һ���̡�"
               Height          =   180
               Index           =   5
               Left            =   120
               TabIndex        =   101
               Top             =   1290
               Width           =   4000
            End
            Begin VB.CheckBox chk 
               Caption         =   "�ƿ����ʱ������ⷿ��Ҫ���������"
               Height          =   180
               Index           =   6
               Left            =   120
               TabIndex        =   100
               Top             =   1920
               Width           =   4095
            End
            Begin VB.Label Label3 
               Caption         =   "�������ѡ����ô����д�ƿⵥ������һ����˲�������˺��Զ���ɱ�ҩ�����͡�������һ����"
               ForeColor       =   &H00000080&
               Height          =   495
               Left            =   360
               TabIndex        =   102
               Top             =   1476
               Width           =   4185
            End
         End
         Begin VB.Frame fra�ɱ��� 
            Caption         =   " �������ɱ�����Դ��ʽ"
            ForeColor       =   &H00800000&
            Height          =   2205
            Left            =   5160
            TabIndex        =   96
            Top             =   5040
            Width           =   4905
            Begin VB.OptionButton opt�������ɱ���Դ 
               Caption         =   $"frmParMedicine.frx":D6C6
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   98
               Top             =   720
               Width           =   3015
            End
            Begin VB.OptionButton opt�������ɱ���Դ 
               Caption         =   "����ԭ��ҩƷ�ĳɱ��ۼ���"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   97
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame fra�⹺������ 
            Caption         =   " �⹺������"
            ForeColor       =   &H00800000&
            Height          =   4575
            Left            =   120
            TabIndex        =   83
            Top             =   120
            Width           =   4935
            Begin VB.Frame fra�ϴβɹ���Ϣ 
               Caption         =   " �ϴβɹ���Ϣ��Դ��ʽ"
               ForeColor       =   &H00800000&
               Height          =   1365
               Left            =   120
               TabIndex        =   93
               Top             =   3120
               Width           =   4605
               Begin VB.CheckBox chk 
                  Caption         =   "����ȡĿ¼�еĲ��ء���׼�ĺ�"
                  Height          =   255
                  Index           =   35
                  Left            =   120
                  TabIndex        =   263
                  Top             =   360
                  Width           =   2880
               End
               Begin VB.OptionButton opt�⹺���ȡ�ɱ��۷�ʽ 
                  Caption         =   "���ȴ���һ�����ҵ����ȡ�ɱ��۵���Ϣ"
                  Height          =   180
                  Index           =   1
                  Left            =   120
                  TabIndex        =   95
                  Top             =   1020
                  Width           =   3615
               End
               Begin VB.OptionButton opt�⹺���ȡ�ɱ��۷�ʽ 
                  Caption         =   "���ȴӵ�ǰ�ⷿ�Ŀ�����������ȡ�ɱ��۵���Ϣ"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   94
                  Top             =   727
                  Value           =   -1  'True
                  Width           =   4335
               End
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷֱ��ȷ���ۼ�"
               Height          =   195
               Index           =   36
               Left            =   120
               TabIndex        =   92
               Top             =   960
               Width           =   2280
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ҫ�����˲����������ⵥ"
               Height          =   195
               Index           =   28
               Left            =   120
               TabIndex        =   91
               Top             =   525
               Width           =   3000
            End
            Begin VB.CheckBox chk 
               Caption         =   "��Ҫ������Ǹ������ܽ��и������"
               Height          =   195
               Index           =   70
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   4440
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷͨ���ӳ������"
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   89
               Top             =   1530
               Width           =   2280
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷ��ⰴ��ǰ�ӳ�����"
               Height          =   195
               Index           =   48
               Left            =   120
               TabIndex        =   88
               Top             =   2070
               Width           =   3090
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷͨ���ֶμӳ����"
               Height          =   180
               Index           =   14
               Left            =   120
               TabIndex        =   87
               Top             =   1815
               Width           =   2775
            End
            Begin VB.CheckBox chk 
               Caption         =   "ʱ��ҩƷ���ʱȡ�ϴ��ۼ�"
               Height          =   195
               Index           =   73
               Left            =   120
               TabIndex        =   86
               Top             =   1260
               Width           =   2760
            End
            Begin VB.CheckBox chk 
               Caption         =   "�⹺��������޸Ĳɹ��޼�"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   85
               Top             =   2475
               Width           =   2835
            End
            Begin VB.CheckBox chk 
               Caption         =   "�б�ҩƷ��ѡ����б굥λ���"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   84
               Top             =   2760
               Width           =   2880
            End
         End
         Begin VB.Frame fra��Ӧ������ 
            Caption         =   " �⹺��⹩Ӧ������У��"
            ForeColor       =   &H00800000&
            Height          =   4815
            Left            =   5160
            TabIndex        =   69
            Top             =   120
            Width           =   5085
            Begin VB.TextBox txt 
               Height          =   375
               Index           =   1
               Left            =   3480
               TabIndex        =   75
               Text            =   "�����ԭʼֵ"
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Frame fraCheck 
               Caption         =   " ѡ��У�鷽ʽ"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   72
               Top             =   4080
               Width           =   4815
               Begin VB.OptionButton optCheck 
                  Caption         =   "У��δͨ��ʱ��ֹ����"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   74
                  Top             =   280
                  Width           =   2175
               End
               Begin VB.OptionButton optCheck 
                  Caption         =   "У��δͨ��ʱ����"
                  Height          =   180
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   73
                  Top             =   280
                  Width           =   1935
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
               Height          =   3165
               Left            =   120
               TabIndex        =   71
               Top             =   720
               Width           =   4815
               _cx             =   8493
               _cy             =   5583
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
               Rows            =   13
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmParMedicine.frx":D6E8
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
            Begin VB.Label Label2 
               Caption         =   "    ҩƷ�⹺���༭����ʱ�Ƿ�У��Ӧ�̵���Ϣ�Ƿ��������������Ƿ���ڡ���˫����У�顱�д�"
               ForeColor       =   &H00000080&
               Height          =   540
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   4860
            End
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
         TabIndex        =   32
         Top             =   600
         Width           =   10455
         Begin VB.Frame fra������ 
            Caption         =   " ����������"
            ForeColor       =   &H00800000&
            Height          =   975
            Left            =   120
            TabIndex        =   65
            Top             =   5400
            Width           =   4095
            Begin MSComCtl2.UpDown ud 
               Height          =   300
               Index           =   0
               Left            =   1666
               TabIndex        =   67
               Top             =   360
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               BuddyControl    =   "txtud(0)"
               BuddyDispid     =   196645
               BuddyIndex      =   0
               OrigLeft        =   1920
               OrigTop         =   360
               OrigRight       =   2175
               OrigBottom      =   660
               Max             =   20
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtud 
               Height          =   300
               Index           =   0
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   66
               Top             =   360
               Width           =   720
            End
            Begin VB.Label lbl������ 
               AutoSize        =   -1  'True
               Caption         =   "�����볤��"
               Height          =   180
               Left            =   240
               TabIndex        =   68
               Top             =   420
               Width           =   900
            End
         End
         Begin VB.Frame fra�ۼ۷�ʽ 
            Caption         =   " ��������ۼۼ��㷽ʽ"
            ForeColor       =   &H00800000&
            Height          =   1215
            Left            =   120
            TabIndex        =   58
            Top             =   1920
            Width           =   4035
            Begin VB.OptionButton opt�ۼۼ��� 
               Caption         =   "��һ��ӳ��ʼ����ۼ�"
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   360
               Width           =   3615
            End
            Begin VB.OptionButton opt�ۼۼ��� 
               Caption         =   "���ֶμӳɼ����ۼ�"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   720
               Width           =   3735
            End
         End
         Begin VB.Frame frmStockRange 
            Caption         =   " ���ô洢�ⷿʱ����Ӧ���ڵķ�Χ"
            ForeColor       =   &H00800000&
            Height          =   3255
            Left            =   4800
            TabIndex        =   50
            Top             =   120
            Width           =   3585
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ�������е�ǰ�����µ�ҩƷ(&6)"
               Height          =   225
               Index           =   5
               Left            =   120
               TabIndex        =   56
               Top             =   2280
               Value           =   1  'Checked
               Width           =   2985
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ��������ͬ����ҩƷ(&5)"
               Height          =   225
               Index           =   4
               Left            =   120
               TabIndex        =   55
               Top             =   1905
               Value           =   1  'Checked
               Width           =   2745
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ�������е�ǰѡ���ͬ����ҩƷ(&4)"
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   54
               Top             =   1530
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ�������е�ǰѡ���ͬ����ҩƷ(&3)"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   53
               Top             =   1155
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "Ӧ�������е�ǰѡ���ͬƷ��ҩƷ(&2)"
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   52
               Top             =   780
               Value           =   1  'Checked
               Width           =   3270
            End
            Begin VB.CheckBox chkӦ�÷�Χ 
               Caption         =   "��Ӧ���ڵ�ǰѡ���ҩƷ(&1)"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   405
               Value           =   1  'Checked
               Width           =   2655
            End
            Begin VB.Label lblComment 
               Caption         =   "��ʾ��û��ѡ�񵽵�Ӧ�÷�Χ�����ô洢�ⷿʱ������ѡ��"
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   120
               TabIndex        =   57
               Top             =   2640
               Width           =   2880
            End
         End
         Begin VB.Frame fraIncome 
            Caption         =   " �����ʶ�Ӧȱʡ������Ŀ"
            ForeColor       =   &H00800000&
            Height          =   1605
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   4035
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   6
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1200
               Width           =   2235
            End
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   5
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   757
               Width           =   2235
            End
            Begin VB.ComboBox cbo 
               ForeColor       =   &H80000012&
               Height          =   300
               Index           =   4
               Left            =   1485
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   315
               Width           =   2235
            End
            Begin VB.Label LblNote 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��ҩ"
               Height          =   180
               Index           =   2
               Left            =   885
               TabIndex        =   62
               Top             =   1260
               Width           =   360
            End
            Begin VB.Label LblNote 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��ҩ"
               Height          =   180
               Index           =   1
               Left            =   885
               TabIndex        =   61
               Top             =   817
               Width           =   360
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   1
               Left            =   60
               Picture         =   "frmParMedicine.frx":D8B6
               Top             =   240
               Width           =   480
            End
            Begin VB.Label LblNote 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��ҩ"
               Height          =   180
               Index           =   0
               Left            =   885
               TabIndex        =   49
               Top             =   390
               Width           =   360
            End
         End
         Begin VB.Frame fra���� 
            Caption         =   " ҩƷ���������Զ�����"
            ForeColor       =   &H00800000&
            Height          =   1800
            Left            =   120
            TabIndex        =   42
            Top             =   3360
            Width           =   4035
            Begin VB.OptionButton opt���� 
               Caption         =   "��ҩ�����"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "�ֹ����÷�������"
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   390
               Width           =   1735
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "ҩ���ҩ������"
               Height          =   200
               Index           =   2
               Left            =   120
               TabIndex        =   44
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "ҩ���ҩ����������"
               Height          =   200
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   1440
               Width           =   2055
            End
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   16
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   10455
         Begin VSFlex8Ctl.VSFlexGrid vsf���ݻ��ڿ��� 
            Height          =   6885
            Left            =   240
            TabIndex        =   175
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "���ݻ��ڿ��ƣ�����ҩƷ�������ض�ҵ�񻷽��������޸ĵ���Ŀ"
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   300
            TabIndex        =   176
            Top             =   120
            Width           =   5040
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   7
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   10455
         Begin VB.CheckBox chk 
            Caption         =   "�������ﴦ��������̿���"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   184
            Top             =   240
            Width           =   2535
         End
         Begin VB.Frame fraOpporunity 
            Caption         =   "����ҩʦ�󷽵Ľ���ʱ��"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   720
            TabIndex        =   181
            Top             =   600
            Width           =   7335
            Begin VB.CheckBox chk 
               Caption         =   "�󷽺ϸ�ȷ�Ϻ������Զ����ʹ�������"
               Height          =   180
               Index           =   22
               Left            =   2880
               TabIndex        =   227
               Top             =   360
               Width           =   3975
            End
            Begin VB.OptionButton optOpporunity 
               Caption         =   "���ﴦ������ǰ"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   183
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton optOpporunity 
               Caption         =   "����ҩ����/��ҩǰ"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   182
               Top             =   720
               Width           =   2055
            End
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   5
            Left            =   2595
            MaxLength       =   2
            TabIndex        =   180
            Top             =   2235
            Width           =   375
         End
         Begin VB.CheckBox chk 
            Caption         =   "����סԺҩ��������̿���"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   179
            Top             =   3000
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "��������ҽ�����ϸ�ҽ��������ҽ������������Ƿ�������ҽ���������ҩ����"
            Height          =   255
            Index           =   44
            Left            =   720
            TabIndex        =   178
            Top             =   1920
            Width           =   6975
         End
         Begin VB.CheckBox chk 
            Caption         =   "����סԺҽ�����ϸ�ҽ����סԺҽ������������Ƿ�������ҽ���������ҩ����"
            Height          =   255
            Index           =   45
            Left            =   720
            TabIndex        =   177
            Top             =   3360
            Width           =   6975
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩʦ�����ʱ��      ���ӣ������趨ʱ��ֵδ���Ĵ�����ҽʦ�ɷ���ͨ�������ⲡ�˳�ʱ�������ٴ����һ�ҩ����"
            Height          =   420
            Index           =   0
            Left            =   720
            TabIndex        =   185
            Top             =   2280
            Width           =   9000
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
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   10455
         Begin VB.Frame fra�������� 
            Caption         =   " ҩƷ����������̬����"
            ForeColor       =   &H00800000&
            Height          =   2415
            Left            =   4320
            TabIndex        =   255
            Top             =   120
            Width           =   5655
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   7
               Left            =   1320
               TabIndex        =   261
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.OptionButton opt�������� 
               Caption         =   "����̬����(ʼ���Ե�ǰ����������Ϊ׼)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   259
               Top             =   360
               Value           =   -1  'True
               Width           =   5175
            End
            Begin VB.OptionButton opt�������� 
               Caption         =   "����ָ���·ݵ�δ��ҩ���ݲ������������(���ø÷���ʱ�����ڿ���ҩƷ������ҽ��ʱ��̬�����������)"
               Height          =   540
               Index           =   1
               Left            =   120
               TabIndex        =   258
               Top             =   600
               Width           =   5295
            End
            Begin VB.TextBox txtM 
               Enabled         =   0   'False
               Height          =   270
               Left            =   720
               TabIndex        =   256
               Text            =   "3"
               Top             =   1200
               Width           =   240
            End
            Begin MSComCtl2.UpDown ud�������� 
               Height          =   270
               Left            =   960
               TabIndex        =   257
               Top             =   1200
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   476
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtM"
               BuddyDispid     =   196699
               OrigLeft        =   960
               OrigTop         =   1440
               OrigRight       =   1215
               OrigBottom      =   1710
               Max             =   12
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   0   'False
            End
            Begin VB.Label Label9 
               Caption         =   $"frmParMedicine.frx":E180
               Height          =   615
               Left            =   360
               TabIndex        =   262
               Top             =   1560
               Width           =   5175
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "�·�"
               Height          =   180
               Left            =   360
               TabIndex        =   260
               Top             =   1245
               Width           =   360
            End
         End
         Begin VB.Frame fra���۹��� 
            Caption         =   " ҩƷ���۹���ģʽ"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   120
            TabIndex        =   233
            Top             =   4920
            Width           =   3975
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   0
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   234
               Top             =   270
               Width           =   2700
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "����ģʽ"
               Height          =   180
               Left            =   120
               TabIndex        =   235
               Top             =   330
               Width           =   900
            End
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   18
            ItemData        =   "frmParMedicine.frx":E21B
            Left            =   1725
            List            =   "frmParMedicine.frx":E21D
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1335
            Width           =   2010
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   17
            ItemData        =   "frmParMedicine.frx":E21F
            Left            =   1725
            List            =   "frmParMedicine.frx":E221
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   950
            Width           =   2010
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   180
            Width           =   2010
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   9
            ItemData        =   "frmParMedicine.frx":E223
            Left            =   1725
            List            =   "frmParMedicine.frx":E225
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   565
            Width           =   2010
         End
         Begin VB.Frame fra 
            Caption         =   " ҩƷ���"
            ForeColor       =   &H00800000&
            Height          =   1860
            Index           =   10
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   3975
            Begin VB.Frame fra�Զ���淽ʽ 
               Caption         =   " ���ý��ʱ��"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   78
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
                  Index           =   0
                  Left            =   825
                  TabIndex        =   79
                  Text            =   "25"
                  Top             =   315
                  Width           =   300
               End
               Begin VB.OptionButton opt���ʱ��ģʽ 
                  Caption         =   "ÿ�����һ��"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   81
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton opt���ʱ��ģʽ 
                  Caption         =   "ÿ��    ��"
                  Height          =   180
                  Index           =   1
                  Left            =   120
                  TabIndex        =   80
                  Top             =   315
                  Width           =   1215
               End
            End
            Begin VB.TextBox txt 
               Height          =   270
               Index           =   6
               Left            =   120
               TabIndex        =   82
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
               TabIndex        =   77
               Top             =   720
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton opt��淽ʽ 
               Caption         =   "�ֹ����(���ⷿ���Բ�ͬ���ڽ��)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.Frame Fraҩ����ͨ 
            Caption         =   " ҩ�ⵥ�����"
            ForeColor       =   &H00800000&
            Height          =   705
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   3975
            Begin VB.ComboBox cbo 
               Height          =   300
               Index           =   7
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   270
               Width           =   1380
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "�������������"
               Height          =   180
               Index           =   26
               Left            =   120
               TabIndex        =   20
               Top             =   330
               Width           =   1260
            End
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩƷ���������㷨"
            Height          =   180
            Index           =   44
            Left            =   120
            TabIndex        =   29
            Top             =   1395
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩƷЧ����ʾ��ʽ"
            Height          =   180
            Index           =   31
            Left            =   120
            TabIndex        =   28
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҩ�۱༭���õ�λ"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩƷ�������ģʽ"
            Height          =   180
            Index           =   32
            Left            =   120
            TabIndex        =   26
            Top             =   630
            Width           =   1440
         End
      End
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7455
         Index           =   15
         Left            =   -75000
         ScaleHeight     =   7425
         ScaleWidth      =   10425
         TabIndex        =   16
         Top             =   720
         Width           =   10455
         Begin MSComctlLib.ListView lvw����� 
            Height          =   6975
            Left            =   240
            TabIndex        =   173
            Top             =   360
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   12303
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ils16"
            SmallIcons      =   "ils16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "��������"
               Object.Width           =   4234
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "����鷽ʽ"
               Object.Width           =   4410
            EndProperty
         End
         Begin VB.Label lbl��ʾ 
            Caption         =   "ҩƷ����飨˫������C�����ã�"
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   240
            TabIndex        =   174
            Top             =   120
            Width           =   5775
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
      Height          =   8430
      Left            =   0
      ScaleHeight     =   8430
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
            Icons           =   "frmParMedicine.frx":E227
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
         Icons           =   "frmParMedicine.frx":1B953
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
      Top             =   8430
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
Attribute VB_Name = "frmParMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnOk As Boolean
Private mblnSecondOpporunity As Boolean
Private mRsWay As Recordset
Private mRsType As Recordset
Private mRsPrice As Recordset

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk
    chk_�޶�ҩƷ�Ŀ�� = 2
    
    chk_�״�ҽ��ִ����Ҫ��� = 83
    
    chk_�շ�ͬʱ��ҩ = 17
    
    chk_��Ŀִ��ǰ���շѻ���� = 74
    chk_δ��˼��ʴ�����ҩ = 15
    chk_δ�շѴ�����ҩ = 58
        
    chk_�⹺�����Ҫ�˲� = 28
    chk_�⹺�����Ҫ������Ǹ������ܽ��и��� = 70
    
    chk_ʱ�۷ֶμӳ���� = 14
    chk_ʱ�ۼӳ������ = 21
    chk_ʱ��ҩƷֱ��ȷ���ۼ� = 36
    chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ����� = 48
    chk_ʱ��ҩƷȡ�ϴ��ۼ� = 73
    
    chk_���ȡĿ¼�в�����Ϣ = 35
                
    'ģ�����
    '�⹺���
    chk_�����޸Ĳɹ��޼� = 3
    chk_�б굥λ = 4
    
    '����
    chk_���찴���γ��� = 75
    
    '�ƿ�
    chk_�ƿ����� = 5
    chk_�ƿ�������� = 6
    chk_�����γ���ʱ����¼���Ų��� = 76
    chk_�ƿⰴ���γ��� = 26
    
    '����
    chk_���ð����γ��� = 37
    chk_���ó������� = 40
    
    '�̵�
    chk_�̵�洢�ⷿ = 7
    chk_�̵���Է������ = 8
    chk_�̵���ͣ��ҩƷ = 9
    chk_�̵��̿�������������� = 65
    
    '����
    chk_ʱ��ҩƷ�����ε��� = 10
    chk_�����޼���ʾ = 11
    chk_�ɱ��۰��ⷿ���ε��� = 66
    
    '����
    chk_���ʱ������ = 12
    
    '������ҩ
    chk_��ҩ�Զ����� = 13
    chk_��ҩˢ�� = 16
    chk_ҽ������ = 18
    chk_ȷ�ϲ���ʵ��ȡҩ = 19
    chk_У����ҩ�� = 20
    chk_У�鷢ҩ�� = 24
    chk_��ҩʱ��δ�շѵĵ��ݽ����շ� = 23
    
    '���ŷ�ҩ
    chk_ȱҩ��� = 25
    chk_������� = 27
    chk_��ҩ������ҩ���� = 29
    chk_��Ժ�������� = 30
    chk_סԺҩ����ҩ��ǩ�� = 31
    chk_סԺҩ����ҩ��ǩ�� = 32
    chk_��ҩʱ���ҽ�� = 38
    chk_��ҩ��������Ĭ��Ϊ��ҩ״̬ = 43
    chk_�Ƿ�������ʾܾ� = 53
    
    '����
    chk_�ֹ��������� = 33
    chk_�ֹ�������� = 34
    chk_�����ϴ����� = 39
    chk_��ҩ���������� = 41
    chk_TPN���� = 42
    chk_ɨ��һ����ɲ��� = 46
    chk_��Һ������ = 47
    chk_���÷���ȡ��ʽ = 49
    chk_��Ժ�����Ƿ������÷� = 50
    chk_����ҩƷ���� = 51
    chk_0���ι��� = 52
    chk_�Զ����� = 54
    chk_�������û�ҩ������Һ�������� = 55
    chk_�Զ�����ֻ���������ε��� = 56
    chk_���췢�͵�ҽ����������Һ��ȫ������������ = 57
    chk_���ҩƷ�ڷ��ͻ�����ȡ���÷� = 59
    chk_��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա = 60
    chk_�Ա�ҩ = 61
    chk_��ȡҩ = 62
    chk_��Ժ��ҩ = 63
    chk_��Һ����ҩ���ٴ�������ı���״̬ = 64
    
    '�������
    chk_���ﴦ����� = 0
    chk_סԺҩ����� = 1
    chk_���ﴦ���Զ����� = 22
    chk_��������ҽ�� = 44
    chk_����סԺҽ�� = 45
End Enum

Private Enum constCbo
    cbo_����ģʽ = 0
    cbo_���۵�λ = 3
    cbo_��ҩ������Ŀ = 4
    cbo_��ҩ������Ŀ = 5
    cbo_��ҩ������Ŀ = 6
    cbo_ҩƷ������� = 7
    cbo_ҩƷ����ģʽ = 9
    cbo_Ч����ʾ��ʽ = 17
    cbo_ҩƷ���������㷨 = 18
End Enum

Private Enum constListBox
    lst_PIVA��Դ���� = 0
    lst_PIVA��ҩ;�� = 1
End Enum

Private Enum constUd
    ud_������ = 0
    ud_���ŷ�ҩ��ѯ���� = 1
    ud_������ҩ��ѯ���� = 2
End Enum

Private Enum constTxt
    txt_���ʱ��ģʽ = 0
        
    'ҩƷ�⹺���
    txt_��Ӧ������ = 1
    
    'ҩƷ������ҩ
    txt_������ɫ = 2
    
    'ҩƷ���ŷ�ҩ
    txt_�Զ�ˢ��ʱ�� = 3
    
    '�󴦷����
    txt_����� = 4
    
    '�������
    txt_����ҩʦ�����ʱ�� = 5
    
    '����
    txt_������ֵ = 6
    
    txt_������������ = 7
End Enum

Private Enum constBill
    bill_ҩƷ�ⷿ���� = 3
    bill_ҩƷ�������� = 4
End Enum

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

'������Ƶ�������Ŀ
Private Const cst������Ŀ As String = "�ɹ���,����,�����,������,�ۼ�,���,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"

'ҩƷ�⹺Ĭ�Ͽ�����Ŀ
Private Const cstҩƷ�⹺��Ŀ_�˲� As String = "�����,�ɹ���,�ۼ�,���"
Private Const cstҩƷ�⹺��Ŀ_��� As String = "��Ʊ��,��Ʊ����,��Ʊ���"
Private Const cstҩƷ�⹺��Ŀ_������� As String = "�ɹ���,����,�����,������,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"

'�����⹺Ĭ�Ͽ�����Ŀ
Private Const cst�����⹺��Ŀ_�˲� As String = "�ۼ�"
Private Const cst�����⹺��Ŀ_��� As String = "�ɹ���,����,�����,������,��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"
Private Const cst�����⹺��Ŀ_������� As String = "�����,������"

Private Sub chk��ҩ;��_Click()
    lst(lst_PIVA��ҩ;��).Enabled = (chk��ҩ;��.value = 1)
    lst(lst_PIVA��ҩ;��).BackColor = IIF(lst(lst_PIVA��ҩ;��).Enabled, &H80000005, &H8000000F)
    
    If Me.Visible And chk��ҩ;��.value = 0 Then
        Call SetParChange(lst, lst_PIVA��ҩ;��, mrsPar, True, "")
    End If
End Sub

Private Sub chk��Դ����_Click()
    lst(lst_PIVA��Դ����).Enabled = (chk��Դ����.value = 1)
    lst(lst_PIVA��Դ����).BackColor = IIF(lst(lst_PIVA��Դ����).Enabled, &H80000005, &H8000000F)
    
    If Me.Visible And chk��Դ����.value = 0 Then
        Call SetParChange(lst, lst_PIVA��Դ����, mrsPar, True, "")
    End If
End Sub


Private Sub chkӦ�÷�Χ_Click(Index As Integer)
    If chkӦ�÷�Χ(Index).value <> Val(chkӦ�÷�Χ(Index).Tag) Then
        chkӦ�÷�Χ(Index).ForeColor = &HC0&             '�޸ĺ������ɫǰ��ɫ��ʶ
    Else
        chkӦ�÷�Χ(Index).ForeColor = &H0&
    End If
End Sub

Private Sub chkӦ�÷�Χ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chkӦ�÷�Χ, 0, mrsPar, "", chkӦ�÷�Χ(Index))
End Sub


Private Sub chk�Զ�ˢ��_Click()
    If chk�Զ�ˢ��.value = 1 Then
        txt(txt_�Զ�ˢ��ʱ��).Enabled = True
    Else
        txt(txt_�Զ�ˢ��ʱ��).Text = "0"
        txt(txt_�Զ�ˢ��ʱ��).Enabled = False
    End If
    
    If Me.Visible Then
        Call SetParChange(txt, txt_�Զ�ˢ��ʱ��, mrsPar)
    End If
End Sub

Private Sub chk�Զ�ˢ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_�Զ�ˢ��ʱ��, mrsPar, "", chk�Զ�ˢ��)
End Sub


Private Sub cmdDefaultColor_Click()
    Dim strColor As String
    Dim n As Integer
    
    Call Get����Ĭ����ɫ
    
    '������ɫ
    For n = 0 To pic������ɫ.UBound
        strColor = IIF(strColor = "", "", strColor & ";") & CStr(pic������ɫ(n).BackColor)
    Next
    
    If Me.Visible Then
        Call SetParChange(txt, txt_������ɫ, mrsPar, True, strColor)
    End If
    
    fraSetColor.ForeColor = txt(txt_������ɫ).ForeColor
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLast_Click()
    Dim intRow As Integer
    Dim str��ҩ���� As String
    Dim str�շ���Ŀ As String
    Dim lng��Ŀid As Long
    
    With VSFPrice
        intRow = .Row
        If intRow < 2 Then Exit Sub
        lng��Ŀid = .TextMatrix(.Row - 1, .ColIndex("��Ŀid"))
        str��ҩ���� = .TextMatrix(.Row - 1, .ColIndex("��ҩ����"))
        str�շ���Ŀ = .TextMatrix(.Row - 1, .ColIndex("�շ���Ŀ"))
        .TextMatrix(.Row - 1, .ColIndex("��Ŀid")) = .TextMatrix(.Row, .ColIndex("��Ŀid"))
        .TextMatrix(.Row - 1, .ColIndex("��ҩ����")) = .TextMatrix(.Row, .ColIndex("��ҩ����"))
        .TextMatrix(.Row - 1, .ColIndex("�շ���Ŀ")) = .TextMatrix(.Row, .ColIndex("�շ���Ŀ"))
        
        
        .TextMatrix(.Row, .ColIndex("��Ŀid")) = lng��Ŀid
        .TextMatrix(.Row, .ColIndex("��ҩ����")) = str��ҩ����
        .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) = str�շ���Ŀ
        
        .Row = intRow - 1
    End With
End Sub

Private Sub cmdNext_Click()
    Dim intRow As Integer
    Dim str��ҩ���� As String
    Dim str�շ���Ŀ As String
    Dim lng��Ŀid As Long
    
    With VSFPrice
        intRow = .Row
        If intRow = .Rows - 1 Then Exit Sub
        lng��Ŀid = .TextMatrix(.Row + 1, .ColIndex("��Ŀid"))
        str��ҩ���� = .TextMatrix(.Row + 1, .ColIndex("��ҩ����"))
        str�շ���Ŀ = .TextMatrix(.Row + 1, .ColIndex("�շ���Ŀ"))
        .TextMatrix(.Row + 1, .ColIndex("��Ŀid")) = .TextMatrix(.Row, .ColIndex("��Ŀid"))
        .TextMatrix(.Row + 1, .ColIndex("��ҩ����")) = .TextMatrix(.Row, .ColIndex("��ҩ����"))
        .TextMatrix(.Row + 1, .ColIndex("�շ���Ŀ")) = .TextMatrix(.Row, .ColIndex("�շ���Ŀ"))
        
        
        .TextMatrix(.Row, .ColIndex("��Ŀid")) = lng��Ŀid
        .TextMatrix(.Row, .ColIndex("��ҩ����")) = str��ҩ����
        .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) = str�շ���Ŀ
        
        .Row = intRow + 1
    End With
End Sub


Private Sub cmdNO_Click()
    picPRI.Visible = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdYes_Click()
    Dim strIds As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(0, strIds)
        
    If picPRI.Tag = 0 Then
        If VSFPrice.Col = VSFPrice.ColIndex("�շ���Ŀ") Then
            Me.VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.Col) = strReturn
            VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.ColIndex("��Ŀid")) = strIds
        Else
            With Me.VSFPrice
                If VSFPrice.Col = .ColIndex("��ҩ����") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "����ҩ�����Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
            End With
        End If
    ElseIf picPRI.Tag = 1 Then
        If VSFPrice_��ҩ;��.Col = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
            Me.VSFPrice_��ҩ;��.TextMatrix(VSFPrice_��ҩ;��.Row, VSFPrice_��ҩ;��.Col) = strReturn
            VSFPrice_��ҩ;��.TextMatrix(VSFPrice_��ҩ;��.Row, VSFPrice_��ҩ;��.ColIndex("��Ŀid")) = strIds
        Else
            With VSFPrice_��ҩ;��
                If VSFPrice_��ҩ;��.Col = .ColIndex("��ҩ;��") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "�ø�ҩ;���Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
                .TextMatrix(.Row, .ColIndex("����id")) = strIds
            End With
        End If
    End If
    
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
    
    '���ڴ�С��13000,8385
    mblnOk = False
    Me.Width = 13000
    Me.Height = 8385
    
    For Each objPic In picPar
        Set objPic.Container = Me
    Next
    
    With VSFPrice
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
    With VSFPrice_��ҩ;��
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
    tabDesign.Visible = False
    
    strCategory = "��������,������Ŀ"
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "100,0,ҩƷͨ������;101,1,ҩƷĿ¼����;110,2,ҩƷ�������;111,3,ҩƷ�ڿ����;112,4,ҩƷ������ҩ;113,5,ҩƷ���ŷ�ҩ;114,6,�������Ĺ���;115,7,����������"
    
    '��������Pickture������11��ʼ��
    marrFunc(1) = "101,11,ҩ����ҩ����;102,12,ҩƷ¼�뾫��;107,13,ҩƷ������λ;105,14,ҩƷ�������;106,15,ҩƷ�����;108,16,���ݻ��ڿ���"
    
    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "��ʼ�ɹ�"
End Sub

Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIds As String) As String
    'intType:0-˫���б�ʱ��1-�б��а��س�ʱ
    Dim n As Integer
    Dim strReturn As String
    
    With lvwPRI
        If .SelectedItem Is Nothing Then Exit Function
        
        strReturn = .SelectedItem.Text
        strIds = Mid(.SelectedItem.Key, 2)
        
        picPRI.Visible = False
        
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        ReturnSelectedPri = strReturn
'        mblnEdit = True
    End With
End Function

Private Sub lvwPRI_DblClick()
    Dim strIds As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(0, strIds)
    
    If picPRI.Tag = 0 Then
        If VSFPrice.Col = VSFPrice.ColIndex("�շ���Ŀ") Then
            Me.VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.Col) = strReturn
            VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.ColIndex("��Ŀid")) = strIds
        Else
            With Me.VSFPrice
                If VSFPrice.Col = .ColIndex("��ҩ����") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "����ҩ�����Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
            End With
        End If
    ElseIf picPRI.Tag = 1 Then
        If VSFPrice_��ҩ;��.Col = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
            Me.VSFPrice_��ҩ;��.TextMatrix(VSFPrice_��ҩ;��.Row, VSFPrice_��ҩ;��.Col) = strReturn
            VSFPrice_��ҩ;��.TextMatrix(VSFPrice_��ҩ;��.Row, VSFPrice_��ҩ;��.ColIndex("��Ŀid")) = strIds
        Else
            With VSFPrice_��ҩ;��
                If VSFPrice_��ҩ;��.Col = .ColIndex("��ҩ;��") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "�ø�ҩ;���Ѿ���ӣ�������ѡ��", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
                .TextMatrix(.Row, .ColIndex("����id")) = strIds
            End With
        End If
    End If
    
End Sub

Private Sub optCheck_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_��Ӧ������, mrsPar, True, Get��Ӧ������У��)
    End If
    
    fra��Ӧ������.ForeColor = txt(txt_��Ӧ������).ForeColor
End Sub

Private Sub optOpporunity_Click(Index As Integer)
    If Me.Visible Then
        '������δ���ļ�¼
        If mblnSecondOpporunity Then        'mblnSecondOpporunity ���ƶ�����ʾMsgbox
            mblnSecondOpporunity = False
            Exit Sub
        End If
        If GetRecipeAuditBills(1) Then
            mblnSecondOpporunity = True
            MsgBox "�������ϵͳ�������δ���ļ�¼�����飡", vbInformation, gstrSysName
            If Val(fraOpporunity.Tag) = 2 Then
                optOpporunity(2).value = True
            Else
                optOpporunity(1).value = True
            End If
            Exit Sub
        Else
            Me.fraOpporunity.Tag = CStr(Index)
        End If
        
        chk(chk_���ﴦ���Զ�����).Enabled = optOpporunity(1).value And optOpporunity(1).Enabled
        If chk(chk_���ﴦ���Զ�����).Enabled = False Then chk(chk_���ﴦ���Զ�����).value = 0
        
        Call SetParChange(optOpporunity, Index, mrsPar, True, IIF(optOpporunity(1).value, 1, 2))
    End If
End Sub

Private Sub optOpporunity_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOpporunity, Index, mrsPar)
End Sub


Private Sub opt��淽ʽ_Click(Index As Integer)
    Dim strValue As String
    
    If opt��淽ʽ(0).value = True Then
        opt���ʱ��ģʽ(0).Enabled = False
        opt���ʱ��ģʽ(1).Enabled = False
        txt(txt_���ʱ��ģʽ).Enabled = False
        
        '�ֹ�������ֵΪ-1
        strValue = "-1"
    Else
        opt���ʱ��ģʽ(0).Enabled = True
        opt���ʱ��ģʽ(1).Enabled = True
        txt(txt_���ʱ��ģʽ).Enabled = opt���ʱ��ģʽ(1).value
        
        strValue = IIF(opt���ʱ��ģʽ(0).value, 0, Val(txt(txt_���ʱ��ģʽ).Text))
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

Private Sub opt��������_Click(Index As Integer)
    If Index = 0 Then
        txtM.Enabled = False
        ud��������.Enabled = False
        Call SetParChange(txt, txt_������������, mrsPar, True, 0)
    Else
        txtM.Enabled = True
        ud��������.Enabled = True
        Call SetParChange(txt, txt_������������, mrsPar, True, Val(txtM.Text))
    End If
End Sub

Private Sub opt�������ɱ���Դ_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt�������ɱ���Դ, Index, mrsPar)
    End If
End Sub

Private Sub opt�������ɱ���Դ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�������ɱ���Դ, Index, mrsPar)
End Sub


Private Sub opt����_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt����, Index, mrsPar)
    End If
End Sub

Private Sub opt����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt����, Index, mrsPar)
End Sub


Private Sub opt�ۼۼ���_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt�ۼۼ���, Index, mrsPar)
    End If
End Sub

Private Sub opt�ۼۼ���_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call SetParTip(opt�ۼۼ���, Index, mrsPar)
End Sub


Private Sub opt�⹺���ȡ�ɱ��۷�ʽ_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt�⹺���ȡ�ɱ��۷�ʽ, Index, mrsPar)
    End If
End Sub

Private Sub opt�⹺���ȡ�ɱ��۷�ʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt�⹺���ȡ�ɱ��۷�ʽ, Index, mrsPar)
End Sub


Private Sub pic������ɫ_Click(Index As Integer)
    Dim strColor As String
    Dim n As Integer
    
    On Error GoTo errHandle
    
    cmdialog.CancelError = True
    cmdialog.ShowColor
    pic������ɫ(Index).BackColor = cmdialog.Color
    
    '������ɫ
    For n = 0 To pic������ɫ.UBound
        strColor = IIF(strColor = "", "", strColor & ";") & CStr(pic������ɫ(n).BackColor)
    Next
    
    If Me.Visible Then
        Call SetParChange(txt, txt_������ɫ, mrsPar, True, strColor)
    End If
    
    fraSetColor.ForeColor = txt(txt_������ɫ).ForeColor
    
    Exit Sub
errHandle:
'    Resume
End Sub

Private Sub Get����Ĭ����ɫ()
    pic������ɫ(��������.��ͨ).BackColor = mconlng��ͨ
    pic������ɫ(��������.����).BackColor = mconlng����
    pic������ɫ(��������.����).BackColor = mconlng����
    pic������ɫ(��������.����).BackColor = mconlng����
    pic������ɫ(��������.��һ).BackColor = mconlng��һ
    pic������ɫ(��������.����).BackColor = mconlng����
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        objPic.Visible = (objPic.Index = Item.ID)
    Next
        
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("ҩ����ҩ����", marrFunc) Or _
                            Item.ID = GetFuncID("��Һ��������", marrFunc) Or _
                            Item.ID = GetFuncID("ҩƷ�������", marrFunc) Or _
                            Item.ID = GetFuncID("ҩƷ�����", marrFunc) Or _
                            Item.ID = GetFuncID("ҩƷ������λ", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        
        If Item.ID = GetFuncID("��Һ��������", marrFunc) Then
            lblLocate(txt_Dept).Caption = "���Ҳ���(&F)"
        Else
            lblLocate(txt_Dept).Caption = "ҩ������(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '���ڻ�ȡ��ǰѡ�е�TaskPanelItem
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



Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    If Me.Visible Then
        
        Call SetParChange(lst, Index, mrsPar)
    End If
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(lst, Index, mrsPar)
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


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
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
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
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
    
    
    
    '2.��ʼ������ؼ�
    Call InitEnv
        
    Call LoadҩƷ�ⷿ����
    Call LoadҩƷ���ÿⷿ
    
    Call Load�ⷿ���
    
    Call LoadҩƷ���ľ���
    Call Load���ݻ��ڿ���
    
    Call LoadOther
    Call LoadVsfPrice
    Call LoadVsfPrice_��ҩ;��
    Call Load��Һ�Ա�ҩ�嵥
    
    '3.����ϵͳ����
    Call LoadPar
    
    
End Sub

Private Sub LoadVsfPrice()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ���,��ҩ����,��Ŀid,�շ���Ŀ from �����շѷ��� where nvl(����id,0) = 0 order by ���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .Rows = 2
            .TextMatrix(1, .ColIndex("���ȼ�")) = 1
        Else
            .Rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            If NVL(rsTemp!��Ŀid) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("���ȼ�")) = i
                .TextMatrix(i, .ColIndex("��ҩ����")) = rsTemp!��ҩ����
                .TextMatrix(i, .ColIndex("��Ŀid")) = rsTemp!��Ŀid
                .TextMatrix(i, .ColIndex("�շ���Ŀ")) = rsTemp!�շ���Ŀ
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPrice_��ҩ;��()
    Dim rsTemp As Recordset
    Dim rsData As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select ����id,��Ŀid,�շ���Ŀ from �����շѷ��� where nvl(����id,0) <> 0 order by ���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice_��ҩ;��
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .Rows = 2
        Else
            .Rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            '��ѯ������Ŀ����
            gstrSQL = "select ���� from ������ĿĿ¼ where id = [1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ������Ŀ����", rsTemp!����id)
            
            If NVL(rsTemp!��Ŀid) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("����id")) = rsTemp!����id
                .TextMatrix(i, .ColIndex("��ҩ;��")) = rsData!����
                .TextMatrix(i, .ColIndex("��Ŀid")) = rsTemp!��Ŀid
                .TextMatrix(i, .ColIndex("�շ���Ŀ")) = rsTemp!�շ���Ŀ
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
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
    Set rsTmp = GetPar(mrsPar, pҩƷĿ¼���� & "," & _
            pҩƷ�⹺���� & "," & _
            pҩƷ������� & "," & _
            pҩƷ�ƿ���� & "," & _
            pҩƷ�̵���� & "," & _
            pҩƷ���۹��� & "," & _
            pҩƷ�������� & "," & _
            pҩƷ������ҩ & "," & _
            pҩƷ���ŷ�ҩ & "," & _
            p�󴦷���� & "," & _
            p��Һ�������� & "," & _
            p���ﴦ����� & "," & _
            pסԺҩ����� & "," & _
            p���������Ŀ & "," & _
            p����������� & "," & _
            p�������ͳ�� & "," & _
            pҩƷ������� & "," & _
            pҩƷ���ù���)
    
    '----------------------------------------------------------
    'ϵͳ����
    '1.����CheckBox�����
    strTmp = "0:6:" & chk_δ��˼��ʴ�����ҩ & _
            ",0:18:" & chk_�޶�ҩƷ�Ŀ�� & _
            ",0:45:" & chk_�շ�ͬʱ��ҩ & _
            ",0:54:" & chk_ʱ�ۼӳ������ & _
            ",0:75:" & chk_�⹺�����Ҫ�˲� & _
            ",0:76:" & chk_ʱ��ҩƷֱ��ȷ���ۼ� & _
            ",0:126:" & chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ����� & _
            ",0:148:" & chk_δ�շѴ�����ҩ & _
            ",0:163:" & chk_��Ŀִ��ǰ���շѻ���� & _
            ",0:173:" & chk_�⹺�����Ҫ������Ǹ������ܽ��и��� & _
            ",0:181:" & chk_ʱ�۷ֶμӳ���� & _
            ",0:183:" & chk_ʱ��ҩƷȡ�ϴ��ۼ� & _
            ",0:214:" & chk_�״�ҽ��ִ����Ҫ��� & _
            ",0:294:" & chk_���ȡĿ¼�в�����Ϣ
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '���ò�����ϵ
    If chk(chk_��Ŀִ��ǰ���շѻ����).value = 1 Then
        chk(chk_δ��˼��ʴ�����ҩ).Enabled = False
        chk(chk_δ�շѴ�����ҩ).Enabled = False
        lblδ�շѷ�ҩ.Caption = "  ������������һ��ͨ������ִ��ǰ�������շѻ��ȼ�����ˡ���������ﲡ�˷�ҩ����ʱ�����²������۹�ѡ����ʧЧ��"
    Else
        chk(chk_δ��˼��ʴ�����ҩ).Enabled = True
        chk(chk_δ�շѴ�����ҩ).Enabled = True
        lblδ�շѷ�ҩ.Caption = "  �������������һ��ͨ������ִ��ǰ�������շѻ��ȼ�����ˡ���������ﲡ�˷�ҩ����ʱ�����²�����ʧЧ��"
    End If
        
    '2.����ComboBox�����
    strTmp = "0:29:" & cbo_���۵�λ & _
            ",0:64:" & cbo_ҩƷ������� & _
            ",0:87:" & cbo_ҩƷ����ģʽ & _
            ",0:149:" & cbo_Ч����ʾ��ʽ & _
            ",0:150:" & cbo_ҩƷ���������㷨
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    '��val(cbo.list)ȡֵ
    strTmp = "0:275:" & cbo_����ģʽ
    Call SetParToControl(strTmp, mrsPar, cbo, 2)
        
    '3.����UpDown�����
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
    
    
    '4.����TextBox�����
    strTmp = ""
'    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.����ListBox�����
'    strTmp = pסԺҽ���´� & ":44:" & lst_��Һ���ķ�ҩ���˿���
'    Call SetParToControl(strTmp, mrsPar, lst, 1)
    
    '6.����OptionButton�����
    arrObj = Array(0, 19, opt��ҩ����)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.����ϵͳ����
    rsTmp.Filter = "ģ��=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        
        Case 221   'ҩƷ���ʱ��ģʽ
            If Val(strValue) = -1 Then
                '����ֵΪ-1��ʾ�ֹ����
                opt��淽ʽ(0).value = True
                opt��淽ʽ(1).value = False
                
                opt���ʱ��ģʽ(0).Enabled = False
                opt���ʱ��ģʽ(1).Enabled = False
                txt(txt_���ʱ��ģʽ).Enabled = False
            Else
                '����ֵ��Ϊ-1��ʾ�Զ����
                opt��淽ʽ(0).value = False
                opt��淽ʽ(1).value = True
                
                If Val(strValue) = 0 Then
                    '����ֵΪ0��ʾÿ�����һ����
                    opt���ʱ��ģʽ(0).value = True
                    opt���ʱ��ģʽ(1).value = False
                    txt(txt_���ʱ��ģʽ).Enabled = False
                Else
                    '����ֵ����0С�ڵ���31��ʾָ�����ڽ��
                    opt���ʱ��ģʽ(0).value = False
                    opt���ʱ��ģʽ(1).value = True
                    
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
            Call zldatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txt, txt_���ʱ��ģʽ, mrsPar)
        
        Case 292   'ҩƷ����������̬���㷽ʽ
            If Val(strValue) = 0 Then
                opt��������(0).value = True
                opt��������(1).value = False
                txtM.Enabled = False
                ud��������.Enabled = False
                txt(txt_������������) = 0
            Else
                opt��������(0).value = False
                opt��������(1).value = True
                txtM.Enabled = True
                ud��������.Enabled = True
                txtM.Text = Val(strValue)
                txt(txt_������������) = Val(strValue)
            End If
            
            Call SetParRelation(txt, txt_������������, mrsPar, rsTmp!������)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '----------------------------------------------------------
    '8.����ģ�����
    'ҩƷĿ¼���� = 1023
    '����ComboBox�����
    strTmp = pҩƷĿ¼���� & ":����ҩ������Ŀ:" & cbo_��ҩ������Ŀ & _
            "," & pҩƷĿ¼���� & ":�г�ҩ������Ŀ:" & cbo_��ҩ������Ŀ & _
            "," & pҩƷĿ¼���� & ":�в�ҩ������Ŀ:" & cbo_��ҩ������Ŀ
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '����OptionButton�����
    arrObj = Array(pҩƷĿ¼����, "�ۼ۰��ӳɼ���", opt�ۼۼ���, _
                    pҩƷĿ¼����, "ҩƷ���������Զ�����", opt����)
    Call SetParToControl("", mrsPar, arrObj)
    
    '����UpDown�����
    strTmp = pҩƷĿ¼���� & ":������:" & ud_������
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
    
    '��������
    '���������������ֵ��Ӧ����ؼ�(��)���ȵ��ù���������¼�ؼ����ƣ�����ؼ���ʾ��������
    strTmp = pҩƷĿ¼���� & ":Ӧ�÷�Χ:0"
    Call SetParToControl(strTmp, mrsPar, chkӦ�÷�Χ)
    
    rsTmp.Filter = "ģ��=" & pҩƷĿ¼���� & " And ������='Ӧ�÷�Χ'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ, "111111")
    If strValue <> "" Then
        For n = 1 To chkӦ�÷�Χ.Count - 1
            chkӦ�÷�Χ(n).value = Mid(strValue, n + 1, 1)
            chkӦ�÷�Χ(n).Tag = Mid(strValue, n + 1, 1)
        Next
    End If
    
    '----------------------------------------------------------
    'ҩƷ�⹺���� = 1300
    '����CheckBox�����
    strTmp = pҩƷ�⹺���� & ":�޸Ĳɹ��޼�:" & chk_�����޸Ĳɹ��޼� & _
            "," & pҩƷ�⹺���� & ":�б�ҩƷ��ѡ����б굥λ���:" & chk_�б굥λ
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '����OptionButton�����
    arrObj = Array(pҩƷ�⹺����, "ȡ�ϴβɹ��۷�ʽ", opt�⹺���ȡ�ɱ��۷�ʽ)
    Call SetParToControl("", mrsPar, arrObj)
    
    '���⴦��
    '�ò���ʵ���ñ��������ؼ���ʾ���ر���ö����ı��ؼ���¼ԭʼֵ���������������ʾ
    strTmp = pҩƷ�⹺���� & ":����У��:" & txt_��Ӧ������
    Call SetParToControl(strTmp, mrsPar, txt)
    
    rsTmp.Filter = "ģ��=" & pҩƷ�⹺���� & " And ������='����У��'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    Call Load��Ӧ������У��(strValue)
    
    '----------------------------------------------------------
    'ҩƷ������� = 1301
    '����OptionButton�����
    arrObj = Array(pҩƷ�������, "ҩƷ�������ɱ��ۼ��㷽ʽ", opt�������ɱ���Դ)
    Call SetParToControl("", mrsPar, arrObj)
    
    '----------------------------------------------------------
    'ҩƷ�ƿ���� = 1304
    '����CheckBox�����
    strTmp = pҩƷ�ƿ���� & ":�ƿ�����:" & chk_�ƿ����� & _
            "," & pҩƷ�ƿ���� & ":��������:" & chk_�ƿ�������� & _
            "," & pҩƷ�ƿ���� & ":�ƿ�ʱ����ҩƷ����¼��������:" & chk_�����γ���ʱ����¼���Ų��� & _
            "," & pҩƷ�ƿ���� & ":ҩƷ�����γ���:" & chk_�ƿⰴ���γ���
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    'ҩƷ������� = 1343
    '����CheckBox�����
    strTmp = pҩƷ������� & ":ҩƷ�����γ���:" & chk_���찴���γ���
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    'ҩƷ���ù��� = 1305
    '����CheckBox�����
    strTmp = pҩƷ���ù��� & ":ҩƷ�����γ���:" & chk_���ð����γ��� & _
            "," & pҩƷ���ù��� & ":��������:" & chk_���ó�������
    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    '----------------------------------------------------------
    'ҩƷ�̵���� = 1307
    '����CheckBox�����
    strTmp = pҩƷ�̵���� & ":�洢�ⷿ:" & chk_�̵�洢�ⷿ & _
            "," & pҩƷ�̵���� & ":����ҩƷ�������:" & chk_�̵���Է������ & _
            "," & pҩƷ�̵���� & ":����ͣ�õ�ҩƷ:" & chk_�̵���ͣ��ҩƷ & _
            "," & pҩƷ�̵���� & ":�̿��������������:" & chk_�̵��̿��������������
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    'ҩƷ���۹��� = 1333
    '����CheckBox�����
    strTmp = pҩƷ���۹��� & ":ʱ��ҩƷ�����ε���:" & chk_ʱ��ҩƷ�����ε��� & _
            "," & pҩƷ���۹��� & ":�޼���ʾ:" & chk_�����޼���ʾ & "," & pҩƷ���۹��� & ":�ɱ��۰��ⷿ���ε���:" & chk_�ɱ��۰��ⷿ���ε���
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    'ҩƷ�������� = 1331
    '����CheckBox�����
    strTmp = pҩƷ�������� & ":���ʱ���ٿ��:" & chk_���ʱ������
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    'ҩƷ������ҩ = 1341
    '����CheckBox�����
    strTmp = pҩƷ������ҩ & ":�Զ�����:" & chk_��ҩ�Զ����� & _
            "," & pҩƷ������ҩ & ":��ҩ��ˢ����֤:" & chk_��ҩˢ�� & _
            "," & pҩƷ������ҩ & ":ҩƷҽ��������ʱ�����:" & chk_ҽ������ & _
            "," & pҩƷ������ҩ & ":���ò���ʵ��ȡҩȷ��ģʽ:" & chk_ȷ�ϲ���ʵ��ȡҩ & _
            "," & pҩƷ������ҩ & ":У����ҩ��:" & chk_У����ҩ�� & _
            "," & pҩƷ������ҩ & ":У�鷢ҩ��:" & chk_У�鷢ҩ�� & _
            "," & pҩƷ������ҩ & ":��ҩʱ��δ�շѵĵ��ݽ����շ�:" & chk_��ҩʱ��δ�շѵĵ��ݽ����շ�
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '����UpDown�����
    strTmp = pҩƷ������ҩ & ":��ѯδ��ҩ��������:" & ud_������ҩ��ѯ����
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
    
    '���⴦��
    '�ò���ʵ���ñ��������ؼ���ʾ���ر���ö����ı��ؼ���¼ԭʼֵ���������������ʾ
    strTmp = pҩƷ������ҩ & ":������ɫ:" & txt_������ɫ
    Call SetParToControl(strTmp, mrsPar, txt)
    
    rsTmp.Filter = "ģ��=" & pҩƷ������ҩ & " And ������='������ɫ'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    Call Get������ɫ(strValue)
    
    '----------------------------------------------------------
    'ҩƷ���ŷ�ҩ = 1342
    '����CheckBox�����
    strTmp = pҩƷ���ŷ�ҩ & ":ȱҩ���:" & chk_ȱҩ��� & _
            "," & pҩƷ���ŷ�ҩ & ":�ⷿ��λ�����������ʾ:" & chk_������� & _
            "," & pҩƷ���ŷ�ҩ & ":��ҩʱ������ҩ���ʼ�¼:" & chk_��ҩ������ҩ���� & _
            "," & pҩƷ���ŷ�ҩ & ":��˳�Ժ���˵���������:" & chk_��Ժ�������� & _
            "," & pҩƷ���ŷ�ҩ & ":��ҩ��ǩ��:" & chk_סԺҩ����ҩ��ǩ�� & _
            "," & pҩƷ���ŷ�ҩ & ":��ҩʱ���ҽ��:" & chk_��ҩʱ���ҽ�� & _
            "," & pҩƷ���ŷ�ҩ & ":��ҩ��������Ĭ��Ϊ��ҩ״̬:" & chk_��ҩ��������Ĭ��Ϊ��ҩ״̬ & _
            "," & pҩƷ���ŷ�ҩ & ":��ҩ��ǩ��:" & chk_סԺҩ����ҩ��ǩ�� & _
            "," & pҩƷ���ŷ�ҩ & ":�Ƿ�������ʾܾ�:" & chk_�Ƿ�������ʾܾ�
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '����UpDown�����
    strTmp = pҩƷ���ŷ�ҩ & ":��ѯ����:" & ud_���ŷ�ҩ��ѯ����
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
    
    '����TextBox�����
    strTmp = pҩƷ���ŷ�ҩ & ":�Զ�ˢ��δ��ҩ�嵥:" & txt_�Զ�ˢ��ʱ��
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '���⴦��
    rsTmp.Filter = "ģ��=" & pҩƷ���ŷ�ҩ & " And ������='�Զ�ˢ��δ��ҩ�嵥'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    chk�Զ�ˢ��.value = IIF(Val(strValue) > 0, 1, 0)
    If chk�Զ�ˢ��.value = 0 Then txt(txt_�Զ�ˢ��ʱ��).Enabled = False
    
    '----------------------------------------------------------
    '�󴦷���� = 1347
    '����TextBox�����
    strTmp = p�󴦷���� & ":����׼:" & txt_�����
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '----------------------------------------------------------
    '��Һ�������� = 1345
    '����OptionButton�����
    arrObj = Array(p��Һ��������, "ҽ������", opt��Һҽ����Ч)
    Call SetParToControl("", mrsPar, arrObj)
    
    '����CheckBox�����
    strTmp = p��Һ�������� & ":��������:" & chk_�ֹ��������� & _
            "," & p��Һ�������� & ":�������:" & chk_�ֹ�������� & _
            "," & p��Һ�������� & ":�����ϴ�����:" & chk_�����ϴ����� & _
            "," & p��Һ�������� & ":��Һ��Һ����ҩ��������������:" & chk_��ҩ���������� & _
            "," & p��Һ�������� & ":�������Ĳ����յľ���Ӫ��ҽ���ڲ�������:" & chk_TPN���� & _
            "," & p��Һ�������� & ":�����Σ�ҩƷ����:" & chk_��Һ������ & _
            "," & p��Һ�������� & ":����ҩƷ��ҩƷ����ָ������:" & chk_����ҩƷ���� & _
            "," & p��Һ�������� & ":����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����:" & chk_0���ι��� & _
            "," & p��Һ�������� & ":�����Զ�����:" & chk_�Զ����� & _
            "," & p��Һ�������� & ":�������û�ҩ������Һ��������:" & chk_�������û�ҩ������Һ�������� & _
            "," & p��Һ�������� & ":���÷Ѱ�������ȡ:" & chk_���÷���ȡ��ʽ & _
            "," & p��Һ�������� & ":��Ժ���˲������÷�:" & chk_��Ժ�����Ƿ������÷� & _
            "," & p��Һ�������� & ":ɨ����ƿǩ���Զ�����:" & chk_ɨ��һ����ɲ��� & _
            "," & p��Һ�������� & ":���췢�͵�ҽ����������Һ��ȫ������������:" & chk_���췢�͵�ҽ����������Һ��ȫ������������ & _
            "," & p��Һ�������� & ":�Զ�����ʱ��Һ��������ֻ���������α䶯:" & chk_�Զ�����ֻ���������ε��� & _
            "," & p��Һ�������� & ":���ҩƷ�ڷ��ͻ�����ȡ���÷�:" & chk_���ҩƷ�ڷ��ͻ�����ȡ���÷� & _
            "," & p��Һ�������� & ":�Ա�ҩ��������������:" & chk_�Ա�ҩ & _
            "," & p��Һ�������� & ":��ȡҩ��������������:" & chk_��ȡҩ & _
            "," & p��Һ�������� & ":��Ժ��ҩ��������������:" & chk_��Ժ��ҩ & _
            "," & p��Һ�������� & ":��Һ����ҩ���ٴ�������ı���״̬:" & chk_��Һ����ҩ���ٴ�������ı���״̬ & _
            "," & p��Һ�������� & ":��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա:" & chk_��ӡƿǩʱ��д�������ڵ�ʵ�ʲ���Ա
            
            
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '���⴦��
    '����ListBox�����
    '��ҩ;��
    strTmp = p��Һ�������� & ":��Һ��ҩ;��:" & lst_PIVA��ҩ;��
    Call SetParToControl(strTmp, mrsPar, lst, 4)
    rsTmp.Filter = "ģ��=" & p��Һ�������� & " And ������='��Һ��ҩ;��'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    If strValue <> "" Then
        chk��ҩ;��.value = 1
    End If
    lst(lst_PIVA��ҩ;��).Enabled = chk��ҩ;��.Enabled And (chk��ҩ;��.value = 1)
    lst(lst_PIVA��ҩ;��).BackColor = IIF(lst(lst_PIVA��ҩ;��).Enabled, &H80000005, &H8000000F)
    
    '��Դ����
    strTmp = p��Һ�������� & ":��Դ����:" & lst_PIVA��Դ����
    Call SetParToControl(strTmp, mrsPar, lst, 4)
    rsTmp.Filter = "ģ��=" & p��Һ�������� & " And ������='��Դ����'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!����ֵ)
    If strValue <> "" Then
        chk��Դ����.value = 1
    End If
    lst(lst_PIVA��Դ����).Enabled = chk��Դ����.Enabled And (chk��Դ����.value = 1)
    lst(lst_PIVA��Դ����).BackColor = IIF(lst(lst_PIVA��Դ����).Enabled, &H80000005, &H8000000F)
    
    '�������
    strTmp = "0:245:44,0:246:45,0:267:22"
    Call SetParToControl(strTmp, mrsPar, Me.chk)
    
    strTmp = "0:241:0"
    Call SetParToControl(strTmp, mrsPar, Me.chk)
    rsTmp.Filter = "ģ��=0 And ������=241"
    If rsTmp.EOF Then
        strValue = "0"
    Else
        strValue = zlCommFun.NVL(rsTmp!����ֵ, "0")
    End If
    Select Case Val(strValue)
    Case 1      '�������ã�סԺ������
        chk(chk_���ﴦ�����).value = 1
        chk(chk_סԺҩ�����).value = 0
        chk(chk_����סԺҽ��).Enabled = False
    Case 2      '���ﲻ���ã�סԺ����
        chk(chk_���ﴦ�����).value = 0
        chk(chk_סԺҩ�����).value = 1
        chk(chk_��������ҽ��).Enabled = False
        Me.optOpporunity(1).Enabled = False
        Me.optOpporunity(2).Enabled = False
        Me.txt(txt_����ҩʦ�����ʱ��).Enabled = False
    Case 3      '���סԺ������
        chk(chk_���ﴦ�����).value = 1
        chk(chk_סԺҩ�����).value = 1
    Case Else   '���סԺ��������
        chk(chk_���ﴦ�����).value = 0
        chk(chk_סԺҩ�����).value = 0
        chk(chk_��������ҽ��).Enabled = False
        chk(chk_����סԺҽ��).Enabled = False
        Me.optOpporunity(1).Enabled = False
        Me.optOpporunity(2).Enabled = False
        Me.txt(txt_����ҩʦ�����ʱ��).Enabled = False
    End Select
    
    strTmp = "0:242:1"
    Call SetParToControl(strTmp, mrsPar, Me.optOpporunity)
    rsTmp.Filter = "ģ��=0 And ������=242"
    If rsTmp.EOF Then
        Me.optOpporunity(1).value = True
        Me.fraOpporunity.Tag = "1"
    Else
        If Val(zlCommFun.NVL(rsTmp!����ֵ)) = 2 Then
            Me.optOpporunity(2).value = True
            Me.fraOpporunity.Tag = "2"
        Else
            Me.optOpporunity(1).value = True
            Me.fraOpporunity.Tag = "1"
        End If
    End If
    
    chk(chk_���ﴦ���Զ�����).Enabled = Me.optOpporunity(1).value
    
    strTmp = "0:243:5"
    Call SetParToControl(strTmp, mrsPar, Me.txt)
    
End Sub

Private Sub Get������ɫ(ByVal strParaValue As String)
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If strParaValue <> "" Then
        For n = 0 To UBound(Split(strParaValue, ";"))
            pic������ɫ(n).BackColor = Val(Split(strParaValue, ";")(n))
        Next
    Else
        Call Get����Ĭ����ɫ
    End If
    
    Exit Sub
errHandle:
    Call Get����Ĭ����ɫ
End Sub
Private Function Get��Ӧ������У��() As String
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean

    blnAllUnCheck = True
    
    '��������У����Ŀ�ͷ�ʽ����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    With vsfCheck
        For i = 1 To .Rows - 1
            strCheck = IIF(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("У����Ŀ")) & "," & _
                IIF(.TextMatrix(i, .ColIndex("У��")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("У��")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If blnAllUnCheck = True Then
        strCheck = "0|" & strCheck
    ElseIf optCheck(0).value = True Then
        strCheck = "2|" & strCheck
    Else
        strCheck = "1|" & strCheck
    End If
        
    Get��Ӧ������У�� = strCheck
End Function
Private Sub Load��Ӧ������У��(ByVal strParaValue As String)
    Dim i As Integer
    Dim n As Integer
    Dim intCheckType As Integer
    Dim arrColumn
    
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....

    If strParaValue <> "" Then
        If InStr(1, strParaValue, "|") > 0 Then
            'У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
            intCheckType = Val(Mid(strParaValue, 1, InStr(1, strParaValue, "|") - 1))
            If intCheckType = 2 Then
                optCheck(0).value = True
            ElseIf intCheckType = 1 Then
                optCheck(1).value = True
            End If
            
            strParaValue = Mid(strParaValue, InStr(1, strParaValue, "|") + 1)
             
            If strParaValue <> "" Then
                strParaValue = strParaValue & ";"
                arrColumn = Split(strParaValue, ";")
                For n = 0 To UBound(arrColumn)
                    If arrColumn(n) <> "" Then
                        With vsfCheck
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
    
    '1.��������
    cbo(cbo_ҩƷ����ģʽ).AddItem "˳����"
    cbo(cbo_ҩƷ����ģʽ).AddItem "����+�����+˳����"
    Call zlControl.CboSetWidth(cbo(cbo_ҩƷ����ģʽ).hwnd, cbo(cbo_ҩƷ����ģʽ).Width * 1.2)
    
    cbo(cbo_Ч����ʾ��ʽ).AddItem "0-��ʾʧЧ��"
    cbo(cbo_Ч����ʾ��ʽ).AddItem "1-��ʾ��Ч��"
    Call zlControl.CboSetWidth(cbo(cbo_Ч����ʾ��ʽ).hwnd, cbo(cbo_Ч����ʾ��ʽ).Width * 1.2)
    
    cbo(cbo_ҩƷ���������㷨).AddItem "0-�������Ƚ��ȳ�"
    cbo(cbo_ҩƷ���������㷨).AddItem "1-��Ч������ȳ�"
    Call zlControl.CboSetWidth(cbo(cbo_ҩƷ���������㷨).hwnd, cbo(cbo_ҩƷ���������㷨).Width * 1.2)
    
    cbo(cbo_���۵�λ).AddItem "0-�ۼ۵�λ"
    cbo(cbo_���۵�λ).AddItem "1-ҩ�ⵥλ"
    cbo(cbo_���۵�λ).ListIndex = 0
    
    cbo(cbo_ҩƷ�������).AddItem "0-������"
    cbo(cbo_ҩƷ�������).AddItem "1-��ͬ��ֹ"
    cbo(cbo_ҩƷ�������).ListIndex = 0
    
    cbo(cbo_����ģʽ).AddItem "0-���������۹���ģʽ"
'    cbo(cbo_����ģʽ).AddItem "1-�����ۻ�����������ģʽ"       '��ʱ���ε�1��ģʽ
    cbo(cbo_����ģʽ).AddItem "2-��ȫ��ͨҵ����������ģʽ"
    cbo(cbo_����ģʽ).ListIndex = 0
    cbo(cbo_����ģʽ).Tag = 2     '��val(list)ֵ��д����
    Call zlControl.CboSetWidth(cbo(cbo_����ģʽ).hwnd, cbo(cbo_����ģʽ).Width * 1.2)
    
    '----------------------------------------------------------
    '2.��������
    'ҩƷĿ¼���� = 1023
    gstrSQL = "Select ID,����||'-'||���� ���� From ������Ŀ Where ĩ��=1"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "InitEnv")
    
    With rsData
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_��ҩ������Ŀ).AddItem !����
            cbo(cbo_��ҩ������Ŀ).ItemData(cbo(cbo_��ҩ������Ŀ).NewIndex) = !ID
            .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_��ҩ������Ŀ).AddItem !����
            cbo(cbo_��ҩ������Ŀ).ItemData(cbo(cbo_��ҩ������Ŀ).NewIndex) = !ID
            .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_��ҩ������Ŀ).AddItem !����
            cbo(cbo_��ҩ������Ŀ).ItemData(cbo(cbo_��ҩ������Ŀ).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '��ʾ��ItemDataƥ�����ֵ
    cbo(cbo_��ҩ������Ŀ).Tag = 1
    cbo(cbo_��ҩ������Ŀ).Tag = 1
    cbo(cbo_��ҩ������Ŀ).Tag = 1
    
    gstrSQL = "select nvl(max(length(����)),0) ���� from �շ���Ŀ���� where ����=3"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "�����볤��")
    
    If rsData!���� = 0 Then
        ud(ud_������).Min = 7
    Else
        ud(ud_������).Min = rsData!����
    End If
    ud(ud_������).Max = 40
    
    
    'ҩƷ�⹺���� = 1300
    'ҩƷ������� = 1301
    'ҩƷ�ƿ���� = 1304
    'ҩƷ�̵���� = 1307
    'ҩƷ���۹��� = 1333
    'ҩƷ�������� = 1331
    'ҩƷ������ҩ = 1341
    'ҩƷ���ŷ�ҩ = 1342
    '�󴦷���� = 1347
    
    '----------------------------------------------------------
    '��Һ��������=1345
    ''��ҩ;��
    gstrSQL = "Select ID, ���� as �÷� ,�걾��λ As ���� From ������ĿĿ¼ Where ���='E' And ��������='2'And (�������=2 Or �������=3) And ִ�з��� = 1 " & _
            " And (����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or ����ʱ�� Is Null) Order by ���� "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ;��")
    
    With lst(lst_PIVA��ҩ;��)
        Do While Not rsData.EOF
            .AddItem rsData!�÷�
            .ItemData(.NewIndex) = rsData!ID
            rsData.MoveNext
        Loop
    End With
    
    ''��Դ����
    gstrSQL = "Select ���� || '-' || ���� ����, Id " & _
            " From ���ű� " & _
            " Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And Id In (Select ����id From ��������˵�� Where �������� = '����' And ������� In (2,3)) And " & _
            " (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By ���� || '-' || ���� "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Load��Դ����")

    With lst(lst_PIVA��Դ����)
        Do While Not rsData.EOF
            .AddItem rsData!����
            .ItemData(.NewIndex) = rsData!ID
            rsData.MoveNext
        Loop
    End With
   
    
    '----------------------------------------------------------
    '3.��������
    With Bill(bill_ҩƷ�ⷿ����)
        .Cols = 4 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "���ڿⷿ"
        .TextMatrix(0, 1) = "�Է��ⷿ"
        .TextMatrix(0, 2) = "�Է��ⷿID"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 1900
        .ColWidth(1) = 1900
        .ColWidth(2) = 0
        .ColWidth(3) = 1900
        .ColData(0) = 3
        .ColData(1) = 3
        .ColData(2) = 5
        .ColData(3) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
    
    With Bill(bill_ҩƷ��������)
        .Cols = 3 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "���ò���"
        .TextMatrix(0, 1) = "���ÿⷿ"
        .TextMatrix(0, 2) = "�ⷿID"
        .ColWidth(0) = 1900
        .ColWidth(1) = 1900
        .ColWidth(2) = 0
        .ColData(0) = 1
        .ColData(1) = 3
        .ColData(2) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    

    '�ⷿ��λ
    With msf�ⷿ������λ
        .AllowUserResizing = flexResizeNone
        .FixedRows = 1
        .Cols = 5
        .MergeCol(0) = True
        .FormatString = "ҩƷ�ⷿ|�������|�ۼ۵�λ|���ﵥλ|סԺ��λ|ҩ�ⵥλ"
        .ColWidth(1) = 900
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColWidth(0) = .Width - 900 * 5 - 400
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
    End With
    
    
    With Billҩ����ҩ����
        
        .Cols = 5 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .TextMatrix(0, 0) = "ҩ��"
        .TextMatrix(0, 1) = "�������"
        .TextMatrix(0, 2) = "��ҩ"
        .TextMatrix(0, 3) = "�Զ���ҩ����"
        .TextMatrix(0, 4) = "��ҩȷ��"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColWidth(2) = 600
        .ColWidth(3) = 1200
        .ColWidth(4) = 1000
        .ColData(0) = 0
        .ColData(1) = 0
        .ColData(2) = 0
        .ColData(3) = 4
        .ColData(4) = 0
        
        .PrimaryCol = 0
        .MsfObj.MergeCells = flexMergeFree
        .MergeCol 0, True
        .Active = True
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOk Then
        mrsPar.Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If mrsPar.RecordCount > 0 Or Bill(bill_ҩƷ�ⷿ����).Tag = "���޸�" Or Bill(bill_ҩƷ��������).Tag = "���޸�" _
            Or lvw�����.Tag = "���޸�" Or msf�ⷿ������λ.Tag = "���޸�" Or Billҩ����ҩ����.Tag = "���޸�" _
            Or BillҩƷ���ľ���.Tag = "���޸�" Or vsf���ݻ��ڿ���.Tag = "���޸�" Then
            
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsPar = Nothing
    
End Sub

Private Sub cmdOK_Click()
    Dim objӦ�÷�Χ As CheckBox
    Dim strValue As String
    
    If ValidateData() = False Then Exit Sub
    
    Call SaveҩƷ�ⷿ����
    Call SaveҩƷ��������
    
    
    Call Save�ⷿ���
    Call Save�ⷿ��λ
    Call Saveҩ����ҩ����
    
    Call SaveҩƷ���ľ���
    Call Save���ݻ��ڿ���
    
    Call Save�����շѷ���
    Call Save��Һ�Ա�ҩ�嵥
    
    'ҩƷĿ¼����
    '�����������
    For Each objӦ�÷�Χ In chkӦ�÷�Χ
        strValue = IIF(strValue = "", "", strValue) & objӦ�÷�Χ.value
    Next
    Call SetParChange(chkӦ�÷�Χ, 0, mrsPar, True, strValue)
    
    'ҩƷ������ҩ
    '�����������
    If chk(chk_��ҩ������ҩ����).value = 1 Then zldatabase.SetPara "�����һ�����ʾ�����嵥", 1, glngSys, 1342
    
    updateChange mrsPar
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub updateChange(ByRef rsPar As ADODB.Recordset)
'��������������Ĳ����ı�
'��Ӧ����δ�޸ĵ��Ǵ��ڽ����ϵ�ֵ�����ݿ��ֵ�Ƿ�һ�£���һ�����Խ���Ϊ׼��Ҫ���±���

    'ȷ���ۼ۵ķ�ʽ(chk_ʱ�ۼӳ�����⡢chk_ʱ�۷ֶμӳ���⡢chk_ʱ��ҩƷȡ�ϴ��ۼ�)
    rsPar.Filter = "(�ؼ����� = 'chk' And �ؼ�������� = " & chk_ʱ�۷ֶμӳ���� & ") or (�ؼ����� = 'chk' And �ؼ�������� = " & chk_ʱ�ۼӳ������ & ") or (�ؼ����� = 'chk' And �ؼ�������� = " & chk_ʱ��ҩƷȡ�ϴ��ۼ� & ")"
    
    With rsPar
        Do While Not .EOF
        
            Select Case rsPar!�ؼ��������
            Case chk_ʱ�۷ֶμӳ����
                If ("" & chk(chk_ʱ�۷ֶμӳ����).value <> "" & rsPar!����ֵ) And NVL(rsPar!�޸�״̬, 0) <> 1 Then
                    rsPar!������ֵ = chk(chk_ʱ�۷ֶμӳ����).value
                    rsPar!�޸�״̬ = 1
                    .Update
                    Call MsgBox("���ѣ�������ʱ��ҩƷͨ���ֶμӳ���⡿δ�����޸ģ������ݿⲻһ�£����Խ���Ϊ׼���棡")
                End If
            Case chk_ʱ�ۼӳ������
                If ("" & chk(chk_ʱ�ۼӳ������).value <> "" & rsPar!����ֵ) And NVL(rsPar!�޸�״̬, 0) <> 1 Then
                    rsPar!������ֵ = chk(chk_ʱ�ۼӳ������).value
                    rsPar!�޸�״̬ = 1
                    .Update
                    Call MsgBox("���ѣ�������ʱ��ҩƷͨ���ӳ�����⡿δ�����޸ģ������ݿⲻһ�£����Խ���Ϊ׼���棡")
                End If
            Case chk_ʱ��ҩƷȡ�ϴ��ۼ�
                If ("" & chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value <> "" & rsPar!����ֵ) And NVL(rsPar!�޸�״̬, 0) <> 1 Then
                    rsPar!������ֵ = chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value
                    rsPar!�޸�״̬ = 1
                    .Update
                    Call MsgBox("���ѣ�������ʱ��ҩƷ���ʱȡ�ϴ��ۼۡ�δ�����޸ģ������ݿⲻһ�£����Խ���Ϊ׼���棡")
                End If
            End Select
            
            .MoveNext
        Loop
    End With

End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub


Private Sub opt���ʱ��ģʽ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt��Һҽ����Ч_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt��Һҽ����Ч, Index, mrsPar)
    End If
End Sub

Private Sub opt��Һҽ����Ч_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt��Һҽ����Ч_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt��Һҽ����Ч, Index, mrsPar)
End Sub

Private Sub opt��ҩ����_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt��ҩ����, Index, mrsPar)
End Sub

Private Sub opt���ʱ��ģʽ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt���ʱ��ģʽ, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case chk_סԺҩ�����
        Call SetParTip(chk, chk_���ﴦ�����, mrsPar, , chk(Index))
    Case Else
        Call SetParTip(chk, Index, mrsPar)
    End Select
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub

Private Sub Load���ݻ��ڿ���()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo errHandle
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
        
'        .CellBorderRange 0, 0, 0, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(1, 0) = "ҩƷ�⹺"
        .TextMatrix(2, 0) = "ҩƷ�⹺"
        .TextMatrix(3, 0) = "ҩƷ�⹺"

        .TextMatrix(1, 1) = "�˲�"
        .TextMatrix(2, 1) = "���"
        .TextMatrix(3, 1) = "�������"
        
'        .CellBorderRange 3, 0, 3, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
'
'        .TextMatrix(4, 0) = "�����⹺"
'        .TextMatrix(5, 0) = "�����⹺"
'        .TextMatrix(6, 0) = "�����⹺"
'
'        .TextMatrix(4, 1) = "�˲�"
'        .TextMatrix(5, 1) = "���"
'        .TextMatrix(6, 1) = "�������"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select ����,����,���� From ���ݻ��ڿ��� where ����=1 Order By ����, ����"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "���ݻ��ڿ���")
        
        If Not rsTmp.EOF Then
            For n = 1 To rsTmp.RecordCount
                For m = 2 To intAllItems + 1
                    If InStr(1, "," & rsTmp!���� & ",", Trim(.TextMatrix(0, m))) > 0 Then
                        Select Case rsTmp!����
                            Case ����.ҩƷ�⹺
                                Select Case rsTmp!����
                                    Case ����.�˲�
                                        .TextMatrix(1, m) = "��"
                                    Case ����.���
                                        .TextMatrix(2, m) = "��"
                                    Case ����.�������
                                        .TextMatrix(3, m) = "��"
                                End Select
'                            Case ����.�����⹺
'                                Select Case rsTmp!����
'                                    Case ����.�˲�
'                                        .TextMatrix(4, m) = "��"
'                                    Case ����.���
'                                        .TextMatrix(5, m) = "��"
'                                    Case ����.�������
'                                        .TextMatrix(6, m) = "��"
'                                End Select
                        End Select
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save���ݻ��ڿ���()
    Dim n As Integer
    Dim m As Integer
    Dim strInput As String
    Dim int���� As Integer
    Dim int���� As Integer
    Dim str���� As String
    
    On Error GoTo errHandle
    With vsf���ݻ��ڿ���
        If .Tag = "���޸�" Then
            For n = 1 To .Rows - 1
                Select Case .TextMatrix(n, 0)
                    Case "ҩƷ�⹺"
                        int���� = ����.ҩƷ�⹺
'                    Case "�����⹺"
'                        int���� = ����.�����⹺
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
                        str���� = str���� & IIF(str���� <> "", ",", "") & .TextMatrix(0, m)
                    End If
                Next
                
                If str���� <> "" Then
                    strInput = strInput & IIF(strInput <> "", ";", "") & int���� & "," & int���� & "," & str����
                End If
            Next
        
            gstrSQL = "Zl_���ݻ��ڿ���_Update('" & strInput & "'," & ����.ҩƷ�⹺ & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Bill_GotFocus(Index As Integer)
    If Index = bill_ҩƷ�ⷿ���� Then
        If Val(lblLocate(txt_Dept).Tag) <> bill_ҩƷ�ⷿ���� Then
            lblLocate(txt_Dept).Tag = bill_ҩƷ�ⷿ����
            mlngPreFind = 1
        End If
    ElseIf Index = bill_ҩƷ�������� Then
        If Val(lblLocate(txt_Dept).Tag) <> bill_ҩƷ�������� Then
            lblLocate(txt_Dept).Tag = bill_ҩƷ��������
            mlngPreFind = 1
        End If
    End If
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
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case txt_����ҩʦ�����ʱ��
        Call zlControl.TxtSelAll(txt(Index))
    End Select
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
        Case txt_�Զ�ˢ��ʱ��, txt_�����
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
            KeyAscii = 0
        Case txt_����ҩʦ�����ʱ��
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case txt_����ҩʦ�����ʱ��
        If Val(txt(Index).Text) < 5 Or Val(txt(Index).Text) > 99 Then
            MsgBox "����ҩʦ�����ʱ����Χ��5-99����", vbInformation, gstrSysName
            txt(Index).Text = 10
        End If
    End Select
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
            If Billҩ����ҩ����.Visible Then
                Call LocateDept(strFind, Billҩ����ҩ����, 0)
                                
            ElseIf Bill(bill_ҩƷ��������).Visible Then
                If lblLocate(txt_Dept).Tag = bill_ҩƷ�ⷿ���� Or lblLocate(txt_Dept).Tag = "" Then
                    Call LocateDept(strFind, Bill(bill_ҩƷ�ⷿ����), IIF(Bill(bill_ҩƷ�ⷿ����).Col = 0, 0, 1))
                Else
                    Call LocateDept(strFind, Bill(bill_ҩƷ��������), Bill(bill_ҩƷ��������).Col)
                End If
                
            ElseIf lvw�����.Visible Then
                Call LocateDept(strFind, lvw�����, 1)
                
            ElseIf msf�ⷿ������λ.Visible Then
                Call LocateDept(strFind, msf�ⷿ������λ, 0)
                
            ElseIf lst(lst_PIVA��Դ����).Visible Then
                Call LocateDept(strFind, lst(lst_PIVA��Դ����), 0)
            End If
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
                If .ListItems(i).ListSubItems(lngCol).Text Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst_��Һ���ķ�ҩ���˿���
            With objTmp
                lngRows = .ListCount - 1
                
                lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
                For i = lngStart To .ListCount - 1
                    strCode = Split(.List(i), "-")(0)
                    strName = Split(.List(i), "-")(1)
                    If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
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
                
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
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








Private Sub txtUD_Change(Index As Integer)
    If Index <> 2 Then Exit Sub
    If Val(txtud.Item(2).Text) > 30 Then
        MsgBox "��ѯδ��ҩ����������󲻳���30��!"
    Else
        ud(2).value = Val(txtud.Item(2).Text)
    End If
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Then Exit Sub     '��ѯδ��ҩ��������
    'ֻ������������
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtud, Index, mrsPar)
End Sub


Private Sub ud_Change(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txtud, Index, mrsPar, True, ud(Index).value)
    End If
End Sub

Private Sub ud_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtud, Index, mrsPar)
End Sub




Private Sub ud��������_Change()
     Call SetParChange(txt, txt_������������, mrsPar, True, Val(txtM.Text))
End Sub


Private Sub vsfCheck_DblClick()
    With vsfCheck
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
        Call SetParChange(txt, txt_��Ӧ������, mrsPar, True, Get��Ӧ������У��)
    End If
    
    fra��Ӧ������.ForeColor = txt(txt_��Ӧ������).ForeColor
End Sub


Private Sub VSFPrice_EnterCell()
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    If Me.VSFPrice.Row < 2 Then
        cmdLast.Enabled = False
    ElseIf Me.VSFPrice.Row = Me.VSFPrice.Rows - 1 Then
        cmdNext.Enabled = False
    End If
    
    VSFPrice.Editable = flexEDNone
    
    If VSFPrice.ColSel <> VSFPrice.ColIndex("���ȼ�") Then
        VSFPrice.Editable = flexEDKbdMouse
    End If
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
            
'            '�����⹺�����ѡ��
'            If .TextMatrix(.Row, 0) = "�����⹺" And .TextMatrix(0, .Col) = "���" Then Exit Sub
            
            .TextMatrix(.Row, .Col) = "��"

        End If
        
    End With
End Sub

Private Sub VSFPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_��ҩ;��_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_��ҩ;��_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice_��ҩ;��.Height
        .Top = frmMoney.Top + tabPrice.Top + VSFPrice_��ҩ;��.Top
        .Left = frmMoney.Left + tabPrice.Left + VSFPrice_��ҩ;��.Left
        .Width = VSFPrice_��ҩ;��.Width
        .Tag = 1
    End With
    
    If Col = VSFPrice_��ҩ;��.ColIndex("��ҩ;��") Then
        If mRsWay Is Nothing Then
            Set mRsWay = DeptSendWork_��ҩ;��
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsWay.RecordCount > 0 Then mRsWay.MoveFirst
            Do While Not mRsWay.EOF
                .ListItems.Add , "_" & mRsWay!ID, mRsWay!����
                mRsWay.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice_��ҩ;��.ColIndex("�շ���Ŀ") Then
        If mRsPrice Is Nothing Then
            Set mRsPrice = DeptSendWork_Get�շ���Ŀ
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!ID, mRsPrice!����
                mRsPrice.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub VSFPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice.Height
        .Top = frmMoney.Top + tabPrice.Top + VSFPrice.Top
        .Left = frmMoney.Left + tabPrice.Left + VSFPrice.Left
        .Width = VSFPrice.Width
        .Tag = 0
    End With
    
    
    If Col = VSFPrice.ColIndex("��ҩ����") Then
        If mRsType Is Nothing Then
            Set mRsType = DeptSendWork_Get��ҩ����
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsType.RecordCount > 0 Then mRsType.MoveFirst
            Do While Not mRsType.EOF
                .ListItems.Add , "_" & mRsType!����, mRsType!����
                mRsType.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice.ColIndex("�շ���Ŀ") Then
        If mRsPrice Is Nothing Then
            Set mRsPrice = DeptSendWork_Get�շ���Ŀ
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!ID, mRsPrice!����
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Function DeptSendWork_��ҩ;��() As Recordset
'��ȡ��ҩ;��,Ŀǰֻ��ԡ�����Ӫ������
    On Error GoTo ErrHand
    gstrSQL = "select ID, ���� from ������ĿĿ¼ where ��� = 'E' and �������� = '2' and ִ�з��� = '1' and ִ�б�� = 2"
    
    Set DeptSendWork_��ҩ;�� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ;��")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeptSendWork_Get��ҩ����() As Recordset
'��ȡҩƷ����ҩ����
    On Error GoTo ErrHand
    gstrSQL = "select ����,���� from ��Һ��ҩ����"
    
    Set DeptSendWork_Get��ҩ���� = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ����")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeptSendWork_Get�շ���Ŀ() As Recordset
'��ȡ�շ���Ŀ
    On Error GoTo ErrHand
    gstrSQL = "select id,����,����,���㵥λ,˵�� from �շ���ĿĿ¼ where ���='Z' and nvl(�Ƿ���,0)=0"
    
    Set DeptSendWork_Get�շ���Ŀ = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ���Ŀ")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub VSFPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice.Row = VSFPrice.Rows - 1 Then
        With Me.VSFPrice
            If .TextMatrix(.Row, .ColIndex("��ҩ����")) <> "" And .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = .ColIndex("��ҩ����")
                .TextMatrix(.Row, .ColIndex("���ȼ�")) = .Row
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice.Row
        If VSFPrice.Rows = 2 Then
           VSFPrice.Rows = 1
           VSFPrice.Rows = 2
        Else
            Me.VSFPrice.RemoveItem VSFPrice.Row
        End If
        
        '�������
        For i = intRow To Me.VSFPrice.Rows - 1
            Me.VSFPrice.TextMatrix(i, Me.VSFPrice.ColIndex("���ȼ�")) = i
        Next
    End If
    
End Sub

Private Sub VSFPrice_��ҩ;��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice_��ҩ;��.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice_��ҩ;��.Row = VSFPrice_��ҩ;��.Rows - 1 Then
        Me.VSFPrice_��ҩ;��.Editable = flexEDNone
        With Me.VSFPrice_��ҩ;��
            If .TextMatrix(.Row, .ColIndex("��ҩ;��")) <> "" And .TextMatrix(.Row, .ColIndex("�շ���Ŀ")) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = .ColIndex("��ҩ;��")
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice_��ҩ;��.Row
        If VSFPrice_��ҩ;��.Rows = 2 Then
           VSFPrice_��ҩ;��.Rows = 1
           VSFPrice_��ҩ;��.Rows = 2
        Else
            Me.VSFPrice_��ҩ;��.RemoveItem VSFPrice_��ҩ;��.Row
        End If
    End If
    Me.VSFPrice_��ҩ;��.Editable = flexEDKbd
    
End Sub

Private Sub picPRI_Resize()
    On Error Resume Next
    
    With lvwPRI
        .Top = 0
        .Left = 0
        .Width = picPRI.Width
        .Height = picPRI.Height - 200 - cmdNO.Height
    End With
    
    With cmdNO
        .Top = picPRI.Height - .Height - 50
        .Left = picPRI.Width - .Width - 50
    End With
    
    With cmdYes
        .Top = cmdNO.Top
        .Left = cmdNO.Left - .Width - 100
    End With
End Sub

Private Sub Save��Һ�Ա�ҩ�嵥()
    '���ܣ�������Һ�Ա�ҩ�嵥
    Dim strSql As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    With Me.vsf�Ա�ҩ�嵥
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, .ColIndex("ҩƷid")) <> "") Or i = 1 Then
                gstrSQL = "Zl_��Һ�Ա�ҩ�嵥_����("
                '���
                gstrSQL = gstrSQL & i
                'ҩƷid
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                '�Ƿ�����
                gstrSQL = gstrSQL & "," & IIF(.TextMatrix(i, .ColIndex("�����")) = "", 0, 1)
                '�Ƿ��һ������
                gstrSQL = gstrSQL & "," & i & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "������Һ�Ա�ҩ�嵥")
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

Private Sub Save�����շѷ���()
    Dim i As Integer
    Dim n As Integer
    
    With Me.VSFPrice
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, .ColIndex("���ȼ�")) <> "" And .TextMatrix(i, .ColIndex("�շ���Ŀ")) <> "" And .TextMatrix(i, .ColIndex("��Ŀid")) <> "" And .TextMatrix(i, .ColIndex("��ҩ����")) <> "") Or i = 1 Then
                gstrSQL = "Zl_�����շѷ���_����("
                '���
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("���ȼ�"))) & ","
                '��ҩ����
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("��ҩ����")) & "',"
                '��Ŀid
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("��Ŀid"))) & ","
                '�շ���Ŀ
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("�շ���Ŀ")) & "',"
                '����id
                gstrSQL = gstrSQL & "NULL" & ","
                '�Ƿ��һ������
                gstrSQL = gstrSQL & i & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "���治����ҩƷ")
            End If
        Next
    End With
    
    n = i - 1
    
    With Me.VSFPrice_��ҩ;��
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("�շ���Ŀ")) <> "" And .TextMatrix(i, .ColIndex("��Ŀid")) <> "" And .TextMatrix(i, .ColIndex("��ҩ;��")) <> "" And .TextMatrix(i, .ColIndex("����id")) <> "" Then
                gstrSQL = "Zl_�����շѷ���_����("
                '���
                gstrSQL = gstrSQL & i + n & ","
                '��ҩ����
                gstrSQL = gstrSQL & "NULL" & ","
                '��Ŀid
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("��Ŀid"))) & ","
                '�շ���Ŀ
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("�շ���Ŀ")) & "',"
                '����id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("����id"))) & ","
                '�Ƿ��һ������
                gstrSQL = gstrSQL & i + n & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "���治����ҩƷ")
            End If
        Next
    End With
    
End Sub
Private Sub SaveҩƷ�ⷿ����()
    Dim strTmp As String
    Dim lngRow As Long
    Dim str���� As String
    
    On Error GoTo errHandle
    With Bill(bill_ҩƷ�ⷿ����)
        If .Tag = "���޸�" Then
            For lngRow = 1 To .Rows - 1
                If .RowData(lngRow) > 0 Then
                    str���� = Left(.TextMatrix(lngRow, 3), 1)
                    If str���� = "" Then str���� = "3"
                    strTmp = strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str���� & ","
                End If
            Next
        
            gstrSQL = "zl_ҩƷ�������_Modify('" & strTmp & "')"
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            .Tag = ""
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume

    Call SaveErrLog
End Sub



Private Function ValidateData() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngIndex As Long, strTmp As String
    Dim i As Integer
    Dim strҩƷ�ۼ۾��� As String, strҩƷ�ɱ��۾��� As String
    
    '���ҩƷ��������
    With Bill(bill_ҩƷ�ⷿ����)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "��" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
          
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
              
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                 
                    Exit Function
                End If
            Next
        Next
    End With
    
    '���ҩƷ������������
    With Bill(bill_ҩƷ��������)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "��" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
        
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
             
                    Exit Function
                End If
            Next
        Next
    End With
    
    
    If CheckParChanged(chk, chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����, mrsPar) Then
        If Check�Ƿ���δ��˵��⹺��ⵥ Then
            MsgBox "����δ��˵��⹺��ⵥ�����ܸı������ʱ��ҩƷ��ⰴ��ǰ�ӳ����ۡ�!", vbInformation, gstrSysName
            chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).value = GetParOriginalValue(chk, chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����, mrsPar)
        
            Exit Function
        End If
    End If
    
    '���۹���ҩƷ���ľ��ȼ��
    If cbo(cbo_����ģʽ).ListIndex > 0 Then
        With BillҩƷ���ľ���
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, dig_�������) = "ҩƷ" And .TextMatrix(lngRow, dig_��������) = "���ۼ�" Then
                    strҩƷ�ۼ۾��� = IIF(strҩƷ�ۼ۾��� = "", "", strҩƷ�ۼ۾���) & .TextMatrix(lngRow, dig_����)
                End If
            Next
            
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, dig_�������) = "ҩƷ" And .TextMatrix(lngRow, dig_��������) = "�ɱ���" Then
                    strҩƷ�ɱ��۾��� = IIF(strҩƷ�ɱ��۾��� = "", "", strҩƷ�ɱ��۾���) & .TextMatrix(lngRow, dig_����)
                End If
            Next
            
            If strҩƷ�ۼ۾��� <> strҩƷ�ɱ��۾��� Then
                MsgBox "������ҩƷ���۹���ҩƷ�ۼۺͳɱ��۸�����λ�ľ���Ӧ����һ��!" & vbCrLf & "����ҩƷ¼�뾫��ҳ�������á�", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    ValidateData = True
End Function


Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim lmX As Integer
    Dim lmY As Integer, blnCancel As Boolean
    Dim strTmp As String
    
    With Bill(Index)
        If Index = bill_ҩƷ�������� Then
            If KeyCode <> vbKeyReturn Then Exit Sub
            
            If .Col = 0 Then
                If .Text = "" Then
                        '����һ���ؼ�
                        zlCommFun.PressKey vbKeyTab
                    
                Else
                    strTmp = Replace(.Text, "'", "''")
                    gstrSQL = "Select a.id,a.����,a.���� From ���ű� a , ��������˵�� b " & _
                              " Where a.id = b.����id " & _
                              " And b.�������� In ('��ҩ����') and (a.���� Like [1] or a.���� like [1] or a.���� like [1])"
                    
                    lmX = picPar(tplFunc.Tag).Left + Me.Bill(bill_ҩƷ��������).Left
                    lmY = picPar(tplFunc.Tag).Top + Me.Bill(bill_ҩƷ��������).Top + Me.Bill(bill_ҩƷ��������).RowHeight(.Row) + 350
                    Set rsTmp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ����", False, "", "", False, False, True, lmX, lmY, 300, blnCancel, False, True, UCase(strTmp) & "%")
                    
                    If rsTmp Is Nothing Then Cancel = True: Exit Sub
                    If rsTmp.State <> 1 Then Cancel = True: Exit Sub
                    If rsTmp.EOF = True Then Cancel = True: Exit Sub
        
                    With Bill(bill_ҩƷ��������)
                        .TextMatrix(.Row, 0) = rsTmp("����") & "-" & rsTmp("����")
                        .Text = rsTmp("����") & "-" & rsTmp("����")
                        .RowData(.Row) = rsTmp("ID")
                    End With
                    
                End If
                .Tag = "���޸�"
            End If
        End If
    End With

End Sub


Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        
        If Index = bill_ҩƷ�ⷿ���� Then
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                            Case "2"
                                .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                        End Select
                        
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                        
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                        
                End Select
                .Tag = "���޸�"
            End If
        End If
    End With

End Sub

Private Sub bill_CommandClick(Index As Integer)
'ͨ����ťѡ��ϸĿ
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_ҩƷ�������� Then
        gstrSQL = "Select Distinct Id,����,����,���� From ���ű� a,��������˵�� b " & _
                  "Where a.id = b.����id And b.�������� In('��ҩ����') " & _
                  "    and (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) " & _
                  "order by ���� "
        Set rsTmp = zldatabase.ShowSelect(Me, gstrSQL, 0, "��ҩ����")
        
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> 1 Then Exit Sub
        If rsTmp.EOF = True Then Exit Sub
        
        With Bill(bill_ҩƷ��������)
            .TextMatrix(.Row, 0) = rsTmp("����") & "-" & rsTmp("����")
            .RowData(.Row) = rsTmp("ID")
            .Tag = "���޸�"
        End With
    End If
    
End Sub


Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If ListIndex < 0 Then Exit Sub
    
    With Bill(Index)
        If Index = bill_ҩƷ�ⷿ���� Then
            If .Col = 0 Then
                .RowData(.Row) = .ItemData(ListIndex)
            ElseIf .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            End If
            
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
        
        ElseIf Index = bill_ҩƷ�������� Then
        
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            .TextMatrix(.Row, .Col) = .CboText
        End If
        .Tag = "���޸�"
    End With
    
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        
        If KeyCode = vbKeyReturn Then
            If Index = bill_ҩƷ�ⷿ���� Then
                If .Col = 1 Then
                    .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
                Else
                     .RowData(.Row) = .ItemData(.ListIndex)
                End If
                
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
                
            ElseIf Index = bill_ҩƷ�������� Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            End If
            .Tag = "���޸�"
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'�������һ�еı仯
    With Bill(Index)
        If .MouseRow = 0 Then Exit Sub
        
        If Index = bill_ҩƷ�ⷿ���� Then
            If .MouseCol <> .Cols - 1 Then Exit Sub
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
            End Select
            
            .Tag = "���޸�"
        End If
    End With
End Sub



Private Sub Billҩ����ҩ����_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Billҩ����ҩ����_DblClick(Cancel As Boolean)
    Dim i As Long
    With Me.Billҩ����ҩ����
        If (.Col = 2 Or .Col = 4) And .Row > 0 And Trim(.TextMatrix(.Row, 0)) <> "" Then
            If .TextMatrix(.Row, .Col) = "" And (.Col = 2 Or (.Col = 4 And .TextMatrix(.Row, 1) = "����")) Then
                .TextMatrix(.Row, .Col) = "��"
                If .Col = 4 Then
                    .TextMatrix(.Row, 2) = "��"
                End If
            Else
                If .Col = 2 And .TextMatrix(.Row, 4) = "��" Then Exit Sub
                .TextMatrix(.Row, .Col) = ""
            End If
            .Tag = "���޸�"
        End If
    End With
End Sub

Private Sub Billҩ����ҩ����_EnterCell(Row As Long, Col As Long)
    With Billҩ����ҩ����
        If Col = 3 Then
            If .TextMatrix(Row, 1) = "סԺ" Then
                .ColData(Col) = 4
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 2
            Else
                .ColData(Col) = 0
            End If
        End If
    End With
End Sub

Private Sub Billҩ����ҩ����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Billҩ����ҩ����
        If .Col = 3 Then
            strKey = Val(.Text)
            If strKey > 30 Then
                MsgBox "�Զ���ҩ�������ܴ���30��", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .TextMatrix(.Row, .Col) = IIF(.Text <> "", strKey, "")
            
            .Tag = "���޸�"
        End If
    End With
End Sub

Private Sub Billҩ����ҩ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        With Billҩ����ҩ����
            If .Col = 2 Then
                Call Billҩ����ҩ����_DblClick(False)
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
    
    On Error GoTo errHandle
    'ȡ��󾫶�
    gstrSQL = "Select �ɱ���, ���ۼ�, ʵ������,���۽�� From ҩƷ�շ���¼ Where Rownum <2"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
    intMaxCost = IIF(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIF(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIF(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIF(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

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
            " From ҩƷ���ľ��� where ���=1 Order By ����, ���, ����, ��λ"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
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
                .TextMatrix(n, dig_����) = IIF(rs!���� > 4, 4, rs!����)
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
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Saveҩ����ҩ����()
    Dim i As Integer, blnTrans As Boolean
    
    On Error GoTo errHandle
    
    With Me.Billҩ����ҩ����
        If .Tag = "���޸�" Then
            gcnOracle.BeginTrans: blnTrans = True
            gstrSQL = "ZL_ҩ����ҩ����_DELETE"
            zldatabase.ExecuteProcedure gstrSQL, Me.Caption
        
            For i = 1 To .Rows - 1
                If .RowData(i) > 0 Then
                    gstrSQL = "ZL_ҩ����ҩ����_INSERT(" & .RowData(i) & "," & IIF(.TextMatrix(i, 1) = "����", 1, 2) & "," & IIF(.TextMatrix(i, 2) <> "", 1, 0) & "," & IIF(Val(.TextMatrix(i, 3)) = 0, "Null", Val(.TextMatrix(i, 3))) & "," & IIF(.TextMatrix(i, 4) <> "", 1, 0) & ")"
                    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
            Next
            gcnOracle.CommitTrans: blnTrans = False
            
            .Tag = ""
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub SaveҩƷ���ľ���()
    Dim n As Integer
    Dim strInput As String
       
    On Error GoTo errHandle
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
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetRSDrugStore(ByVal bytMode As Byte) As ADODB.Recordset
'���ܣ���ȡҩ����¼��
'������0-ҩ����ҩ��,1-����ҩ��
    Dim strSql As String
 
    strSql = "Select b.Id, Nvl(b.����, '') ����, Nvl(b.����, '') ����, a.�������, a.��������" & vbNewLine & _
            "From ��������˵�� A, ���ű� B" & vbNewLine & _
            "Where b.Id = a.����id And a.�������� In (" & _
                IIF(bytMode = 0, "'��ҩ��', '��ҩ��', '��ҩ��',", "") & " '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��') And " & Where����ʱ��("B") & vbNewLine & _
            "Order By ����"

    On Error GoTo errH
    Set GetRSDrugStore = zldatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOther()
'�������ĳ�ʼ������
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng��λ As Long
    Dim strTmp As String, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '����ⷿ��λ
    strCoding = ""
    Set rsTemp = GetRSDrugStore(0)
    msf�ⷿ������λ.Rows = 1
    Do Until rsTemp.EOF
        With msf�ⷿ������λ
            If rsTemp("����") <> strCoding Then
                strTmp = ""
            End If
            If InStr(",��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
                If InStr(1, strTmp & ",", ",ҩ��,") <= 0 Then
                    .Rows = .Rows + 1
                    .RowData(.Rows - 1) = rsTemp("ID")
                    .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                    .TextMatrix(.Rows - 1, 1) = "ҩ��"
                    strTmp = strTmp & "," & "ҩ��"
                End If
            End If
            
            If InStr(",�Ƽ���,��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
            
                Select Case rsTemp("�������")
                    Case 0          '�������ڲ���

                    Case 1          '���������ﲡ��
                        If InStr(1, strTmp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTmp = strTmp & "," & "����"
                        End If
                    Case 2          '������סԺ����
                        If InStr(1, strTmp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTmp = strTmp & "," & "סԺ"
                        End If
                    Case 3          '����������סԺ����
                        If InStr(1, strTmp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTmp = strTmp & "," & "����"
                        End If
                        
                        If InStr(1, strTmp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTmp = strTmp & "," & "סԺ"
                        End If
                End Select
            End If
            If InStr(1, strTmp & ",", ",����,") <= 0 Then
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = rsTemp("ID")
                .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                .TextMatrix(.Rows - 1, 1) = "����"
                strTmp = strTmp & "," & "����"
            End If
            
            strCoding = rsTemp("����")
        End With
        rsTemp.MoveNext
    Loop

    If msf�ⷿ������λ.Rows > 1 Then
        msf�ⷿ������λ.FixedRows = 1
    End If
    gstrSQL = "select �ⷿid, ���÷�Χ, ���� from ҩƷ�ⷿ��λ"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngMaxRow = rsTemp.RecordCount
        For lngRow = 1 To lngMaxRow
            For i = 1 To msf�ⷿ������λ.Rows - 1
                Select Case rsTemp!���÷�Χ
                    Case 1
                        strTmp = "ҩ��"
                    Case 2
                        strTmp = "����"
                    Case 3
                        strTmp = "סԺ"
                    Case 4
                        strTmp = "����"
                End Select
                If rsTemp!�ⷿid = msf�ⷿ������λ.RowData(i) And strTmp = msf�ⷿ������λ.TextMatrix(i, 1) Then
                    msf�ⷿ������λ.TextMatrix(i, 2) = ""
                    msf�ⷿ������λ.TextMatrix(i, 3) = ""
                    msf�ⷿ������λ.TextMatrix(i, 4) = ""
                    msf�ⷿ������λ.TextMatrix(i, 5) = ""
                    msf�ⷿ������λ.TextMatrix(i, rsTemp!���� + 1) = "��"
                End If
            Next
            rsTemp.MoveNext
        Next
    End If
    
    'ҩ����ҩ����
    strCoding = ""
    Set rsTemp = GetRSDrugStore(1)
    Billҩ����ҩ����.Clear
    lngRow = 1
    Do Until rsTemp.EOF
        With Billҩ����ҩ����
            If rsTemp("����") <> strCoding Then
                strTmp = ""
            End If
            
            If InStr(",�Ƽ���,��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
            
                Select Case rsTemp("�������")
                    Case 0          '�������ڲ���
                    Case 1          '���������ﲡ��
                        If InStr(1, strTmp & ",", ",����,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTmp = strTmp & "," & "����"
                        End If
                    Case 2          '������סԺ����
                        If InStr(1, strTmp & ",", ",סԺ,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTmp = strTmp & "," & "סԺ"
                        End If
                    Case 3          '����������סԺ����
                        If InStr(1, strTmp & ",", ",����,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTmp = strTmp & "," & "����"
                        End If
                        
                        If InStr(1, strTmp & ",", ",סԺ,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTmp = strTmp & "," & "סԺ"
                        End If
                End Select
            End If
            strCoding = rsTemp("����")
        End With
        rsTemp.MoveNext
    Loop

    gstrSQL = "select ҩ��id, ����, ��ҩ, �Զ���ҩ����,��ҩȷ�� from ҩ����ҩ����"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Billҩ����ҩ����
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            lngMaxRow = rsTemp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To .Rows - 1
                    Select Case rsTemp!����
                        Case 1
                            strTmp = "����"
                        Case 2
                            strTmp = "סԺ"
                    End Select
                    If rsTemp!ҩ��id = .RowData(i) And strTmp = .TextMatrix(i, 1) Then
                        If IIF(IsNull(rsTemp("��ҩ")), 0, rsTemp("��ҩ")) = 1 Then
                            .TextMatrix(i, 2) = "��"
                        End If
                        
                        If IIF(IsNull(rsTemp("��ҩȷ��")), 0, rsTemp("��ҩȷ��")) = 1 Then
                            .TextMatrix(i, 4) = "��"
                        End If
                        .TextMatrix(i, 3) = IIF(IsNull(rsTemp!�Զ���ҩ����), "", rsTemp!�Զ���ҩ����)
                    End If
                Next
                rsTemp.MoveNext
            Next
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load��Һ�Ա�ҩ�嵥()
    '���ܣ����������õ���Һ�Ա�ҩ�嵥
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strSql = "Select ҩƷid, '��' || b.���� || '��' || b.���� || '(' || b.��� || ')' As ����, �Ƿ�����" & vbNewLine & _
            "From ��Һ�Ա�ҩ�嵥 A, �շ���ĿĿ¼ B" & vbNewLine & _
            "Where a.ҩƷid = b.Id" & vbNewLine & _
            "Order By ���"

    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "Load��Һ�Ա�ҩ�嵥")
    
    vsf�Ա�ҩ�嵥.Rows = rsTemp.RecordCount + 2
    
    For i = 1 To rsTemp.RecordCount
        vsf�Ա�ҩ�嵥.TextMatrix(i, vsf�Ա�ҩ�嵥.ColIndex("ҩƷid")) = rsTemp!ҩƷID
        vsf�Ա�ҩ�嵥.TextMatrix(i, vsf�Ա�ҩ�嵥.ColIndex("ҩƷ���������")) = NVL(rsTemp!����)
        vsf�Ա�ҩ�嵥.TextMatrix(i, vsf�Ա�ҩ�嵥.ColIndex("�����")) = IIF(rsTemp!�Ƿ����� = 0, "", "��")
        
        rsTemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ�ⷿ����()
'����:װ��ҩƷ��������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With Bill(bill_ҩƷ�ⷿ����)
        '����װ���ѡ�ⷿ
        gstrSQL = "Select Distinct b.Id, Nvl(b.����, '') ����, Nvl(b.����, '') ���� " & vbNewLine & _
            " From ��������˵�� A, ���ű� B " & vbNewLine & _
            " Where b.Id = a.����id And a.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��') And " & vbNewLine & _
            " (b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or b.����ʱ�� Is Null) " & vbNewLine & _
            "Order By ���� "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����") & "-" & rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("ID")
            
            rsTemp.MoveNext
        Loop
        
        'װ�������������
        gstrSQL = "select A.���ڿⷿID,A.�Է��ⷿID,A.����" & _
                "    ,B.���� as ���ڱ���,B.���� as ��������,C.���� as �Է�����,C.���� as �Է����� " & _
                " from ҩƷ������� A,���ű� B,���ű� C " & _
                " where A.���ڿⷿID= B.ID and A.�Է��ⷿID=C.ID and " & Where����ʱ��("C") & _
                " order by b.����,c.���� "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("���ڿⷿID")
            .TextMatrix(lngRow, 0) = rsTemp("���ڱ���") & "-" & rsTemp("��������")
            .TextMatrix(lngRow, 1) = rsTemp("�Է�����") & "-" & rsTemp("�Է�����")
            .TextMatrix(lngRow, 2) = rsTemp("�Է��ⷿID")
            .TextMatrix(lngRow, 3) = Switch(rsTemp("����") = 1, "1-���ڿⷿ������Է��ⷿ", _
                                            rsTemp("����") = 2, "2-�Է��ⷿ���������ڿⷿ", _
                                                          True, "3-���ⷿ���˫����ͨ")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load�ⷿ���()
    '���ܣ���ʼ���ⷿ
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim ObjItem As ListItem
    On Error GoTo errHandle
    
    gstrSQL = _
        "SELECT B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) ��鷽ʽ" & vbCrLf & _
        " FROM ��������˵�� A, ���ű� B, ҩƷ������ C" & vbCrLf & _
        " WHERE A.����ID = B.ID AND A.����ID = C.�ⷿID(+) AND" & vbCrLf & _
        "      A.�������� IN" & vbCrLf & _
        "      ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')" & vbCrLf & _
        "     And (b.����ʱ��=to_date('3000-1-1','yyyy-mm-dd') or b.����ʱ�� is null) " & vbCrLf & _
        " GROUP BY B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) " & vbCrLf & _
        " order by B.���� "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.lvw�����.ListItems.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set ObjItem = Me.lvw�����.ListItems.Add(, "C_" & rsTmp!ID, rsTmp!����)
            ObjItem.SubItems(1) = "" & rsTmp!����
            ObjItem.SubItems(2) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
            ObjItem.Tag = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save�ⷿ��λ()
    '����ⷿ��λ����
    Dim i As Long
    Dim lngTmp As Long
    Dim intTmp As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    With msf�ⷿ������λ
        If .Rows > 1 And .Tag = "���޸�" Then
            If Trim(.TextMatrix(1, 0)) <> "" Then
                gstrSQL = ""
                For i = 1 To .Rows - 1
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    lngTmp = 1
                    Select Case True
                        Case .TextMatrix(i, 2) = "��"
                            lngTmp = 1
                        Case .TextMatrix(i, 3) = "��"
                            lngTmp = 2
                        Case .TextMatrix(i, 4) = "��"
                            lngTmp = 3
                        Case .TextMatrix(i, 5) = "��"
                            lngTmp = 4
                    End Select
                    Select Case .TextMatrix(i, 1)
                        Case "ҩ��"
                            intTmp = 1
                        Case "����"
                            intTmp = 2
                        Case "סԺ"
                            intTmp = 3
                        Case "����"
                            intTmp = 4
                    End Select
                    gstrSQL = gstrSQL & lngTmp & "," & intTmp & ","
                Next
                strSql = "ZL_ҩƷ�ⷿ��λ_DELETE"
                Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
                
                gstrSQL = "ZL_ҩƷ�ⷿ��λ_INSERT('" & gstrSQL & "')"
                Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            .Tag = ""
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Save�ⷿ���()
    '���ܣ�����ⷿ���
    Dim i As Long
    On Error GoTo errHandle
    
    If lvw�����.Tag = "���޸�" Then
        gstrSQL = ""
        For i = 1 To Me.lvw�����.ListItems.Count
            gstrSQL = gstrSQL & Me.lvw�����.ListItems(i).Tag & "," & Switch(Me.lvw�����.ListItems(i).SubItems(2) = "0-�����", "0", Me.lvw�����.ListItems(i).SubItems(2) = "1-��飬��������", "1", Me.lvw�����.ListItems(i).SubItems(2) = "2-��飬�����ֹ", "2") & ","
        Next
        gstrSQL = "Zl_ҩƷ������_insert('" & gstrSQL & "')"
        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        lvw�����.Tag = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SaveҩƷ��������()
    Dim strTmp As String
    Dim lngRow As Long
    Dim bln���� As Boolean
    
    On Error GoTo ErrHand
    With Bill(bill_ҩƷ��������)
        If .Tag = "���޸�" Then
            For lngRow = 1 To .Rows - 1
                If .RowData(lngRow) > 0 Then
                    If LenB(StrConv(strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ",", vbFromUnicode)) >= 4000 Then
                        If bln���� = True Then
                            gstrSQL = "zl_ҩƷ�����������_Modify('" & strTmp & "'," & 1 & ")"
                        Else
                            gstrSQL = "zl_ҩƷ�����������_Modify('" & strTmp & "'," & 0 & ")"
                        End If
                        bln���� = True
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        
                        strTmp = .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                    Else
                        strTmp = strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                    End If
                End If
            Next
    
            If bln���� = True Then
                gstrSQL = "zl_ҩƷ�����������_Modify('" & strTmp & "'," & 1 & ")"
            Else
                gstrSQL = "zl_ҩƷ�����������_Modify('" & strTmp & "'," & 0 & ")"
            End If
            bln���� = True
            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    Call SaveErrLog
    End If
End Sub

Sub LoadҩƷ���ÿⷿ()
'����:����ҩƷ���ò���
     
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With Bill(bill_ҩƷ��������)
        'װ�������������
        gstrSQL = "Select Distinct b.Id, Nvl(b.����, '') ����, Nvl(b.����, '') ���� " & vbNewLine & _
            " From ��������˵�� A, ���ű� B " & vbNewLine & _
            " Where b.Id = a.����id And a.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��') And " & vbNewLine & _
            " (b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or b.����ʱ�� Is Null) " & vbNewLine & _
            "Order By ���� "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����") & "-" & rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("ID")
            rsTemp.MoveNext
        Loop
        
        'װ�������������
        gstrSQL = "select A.���ò���ID,A.�Է��ⷿID" & _
                ",B.���� as ���ò��ű���,B.���� as ���ò�������,C.���� as �ⷿ����,C.���� as �ⷿ���� " & _
                " from ҩƷ���ÿ��� A,���ű� B,���ű� C " & _
                " where A.���ò���ID= B.ID and A.�Է��ⷿID=C.ID order by b.����,c.���� "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("���ò���ID")
            .TextMatrix(lngRow, 0) = rsTemp("���ò��ű���") & "-" & rsTemp("���ò�������")
            .TextMatrix(lngRow, 1) = rsTemp("�ⷿ����") & "-" & rsTemp("�ⷿ����")
            .TextMatrix(lngRow, 2) = rsTemp("�Է��ⷿID")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub opt���ʱ��ģʽ_Click(Index As Integer)
    Dim strValue As String
    
    txt(txt_���ʱ��ģʽ).Enabled = opt���ʱ��ģʽ(1).value
    
    If Me.Visible Then
        strValue = IIF(opt���ʱ��ģʽ(0).value, 0, Val(txt(txt_���ʱ��ģʽ).Text))
        Call SetParChange(txt, txt_������ֵ, mrsPar, True, strValue)
        
        opt��淽ʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt��淽ʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(0).ForeColor = txt(txt_������ֵ).ForeColor
        opt���ʱ��ģʽ(1).ForeColor = txt(txt_������ֵ).ForeColor
        txt(txt_���ʱ��ģʽ).ForeColor = opt���ʱ��ģʽ(1).ForeColor
    End If
End Sub

Private Sub msf�ⷿ������λ_DblClick()
    Dim i As Long
    
    With msf�ⷿ������λ
    If .Col > 1 And .Row > 0 And Trim(.TextMatrix(.Row, 0)) <> "" Then
        .TextMatrix(.Row, 2) = ""
        .TextMatrix(.Row, 3) = ""
        .TextMatrix(.Row, 4) = ""
        .TextMatrix(.Row, 5) = ""
        .TextMatrix(.Row, .Col) = "��"
        
        .Tag = "���޸�"
    End If
    
    End With
End Sub

Private Sub msf�ⷿ������λ_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn Or KeyAscii = Asc(" ")) Then
        msf�ⷿ������λ_DblClick
    End If
End Sub

Private Sub opt��ҩ����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvw�����_DblClick()
    With lvw�����
        If Not .SelectedItem Is Nothing Then
            .SelectedItem.SubItems(2) = Switch(.SelectedItem.SubItems(2) = "0-�����", "1-��飬��������", _
                .SelectedItem.SubItems(2) = "1-��飬��������", "2-��飬�����ֹ", .SelectedItem.SubItems(2) = "2-��飬�����ֹ", "0-�����")
            .Tag = "���޸�"
        End If
    End With
End Sub

Private Sub lvw�����_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = "C" Then
        Call lvw�����_DblClick
    End If
End Sub

Private Function Check�Ƿ���δ��˵��⹺��ⵥ() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From δ��ҩƷ��¼ Where ���� = 1 And Rownum < 2"
    Call zldatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check�Ƿ���δ��˵��⹺��ⵥ = (rs.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub BillҩƷ���ľ���_EnterCell(Row As Long, Col As Long)
    With BillҩƷ���ľ���
        If Col = dig_���� Then
            .TxtCheck = True
            .TextMask = "123456789"
            .MaxLength = 1
        End If
    End With
End Sub

Private Sub opt��ҩ����_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt��ҩ����, Index, mrsPar)
    End If
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
            
            If cbo(cbo_����ģʽ).ListIndex > 0 Then
                If .TextMatrix(.Row, dig_�������) = "ҩƷ" Then
                    MsgBox "ע�⣬���������۹���ģʽ������������ȿ��ܽ�Ӱ���۽����㣡", vbInformation, gstrSysName
                End If
            End If
            
            .TextMatrix(.Row, .Col) = strKey
            .RowData(.Row) = Val(strKey)
            
            .Tag = "���޸�"
        End If
    End With
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(cbo, Index, mrsPar)
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Dim strVar As String
    Dim rsTemp As ADODB.Recordset
    Dim blnResulte As Boolean
    
    If Me.Visible Then
        Call SetParChange(chk, Index, mrsPar)
    End If
    
    Select Case Index
    Case chk_ʱ�۷ֶμӳ����
        If chk(Index).value = 1 Then
            If chk(chk_ʱ�ۼӳ������).value = 1 Then chk(chk_ʱ�ۼӳ������).value = 0
            If chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value = 1 Then chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value = 0
        End If
    Case chk_ʱ�ۼӳ������
        If chk(Index).value = 1 Then
            If chk(chk_ʱ�۷ֶμӳ����).value = 1 Then chk(chk_ʱ�۷ֶμӳ����).value = 0
            If chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value = 1 Then chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).value = 0
        End If
    Case chk_ʱ��ҩƷȡ�ϴ��ۼ�
        If chk(Index).value = 1 Then
            If chk(chk_ʱ�۷ֶμӳ����).value = 1 Then chk(chk_ʱ�۷ֶμӳ����).value = 0
            If chk(chk_ʱ�ۼӳ������).value = 1 Then chk(chk_ʱ�ۼӳ������).value = 0
        End If
    Case chk_���찴���γ���
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_���찴���γ���).value = Val(GetParOriginalValue(chk, chk_���찴���γ���, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check���쵥
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵����쵥�����ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_���찴���γ���).value = Val(GetParOriginalValue(chk, chk_���찴���γ���, mrsPar))
        End If
    Case chk_���ð����γ���
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_���ð����γ���).value = Val(GetParOriginalValue(chk, chk_���ð����γ���, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check���õ�
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵����õ������ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_���ð����γ���).value = Val(GetParOriginalValue(chk, chk_���ð����γ���, mrsPar))
        End If
    Case chk_�ƿⰴ���γ���
        '���ڼ���ʱ�������������
        If Me.Visible = False Then Exit Sub
        
        '��ǰѡ��ĵ���ԭʼ����ֵʱ������������䣬�������ѭ��
        If chk(chk_�ƿⰴ���γ���).value = Val(GetParOriginalValue(chk, chk_�ƿⰴ���γ���, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
        blnResulte = Check�ƿⵥ
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "���ڽ���δ��˵��ƿⵥ�����ܸı�˲�����", vbInformation, gstrSysName
            chk(chk_�ƿⰴ���γ���).value = Val(GetParOriginalValue(chk, chk_�ƿⰴ���γ���, mrsPar))
        End If
    Case chk_�ƿ��������
        '����Ϊ����Ҫ����ʱ��Ҫ����Ƿ���δ��˵ĳ������뵥����������ܸı�
        
        On Error GoTo errHandle
        If chk(chk_�ƿ��������).value = 0 Then
            If MsgBox("��������Ƿ����δ��˵ĳ������뵥��������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '�ù�����10.20�汾����������һ�������������ڷ�Χ������ȫ��ɨ��
                gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
                    " Where a.���� = 6 And a.�������� Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Mod(b.��¼״̬, 3) = 2) And Rownum < 2"
                
                DoEvents
                zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
                
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵ĳ������뵥")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "����δ��˵ĳ������뵥�����ܸı�˲�����", vbInformation, gstrSysName
                    chk(chk_�ƿ��������).value = 1
                End If
            Else
                chk(chk_�ƿ��������).value = 1
            End If
        End If
    Case chk_���ó�������
        '����Ϊ����Ҫ����ʱ��Ҫ����Ƿ���δ��˵ĳ������뵥����������ܸı�
        
        On Error GoTo errHandle
        If chk(chk_���ó�������).value = 0 Then
            If MsgBox("��������Ƿ����δ��˵ĳ������뵥��������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
                    " Where a.���� = 7 And a.�������� Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Mod(b.��¼״̬, 3) = 2) And Rownum < 2"
                
                DoEvents
                zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
                
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵ĳ������뵥")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "����δ��˵ĳ������뵥�����ܸı�˲�����", vbInformation, gstrSysName
                    chk(chk_���ó�������).value = 1
                End If
            Else
                chk(chk_���ó�������).value = 1
            End If
        End If
    Case chk_���ﴦ�����
        If Me.Visible = False Then Exit Sub
        
        If chk(Index).value = 0 Then
            '������δ���ļ�¼
            If GetRecipeAuditBills(1) Then
                MsgBox "�������ϵͳ�������δ���ļ�¼�����飡", vbInformation, gstrSysName
                chk(Index).value = 1
            End If
        End If
        
        optOpporunity(1).Enabled = chk(Index).value = 1
        optOpporunity(2).Enabled = chk(Index).value = 1
        
        If chk(Index).value = 0 Then chk(chk_��������ҽ��).value = 0
        chk(chk_��������ҽ��).Enabled = chk(Index).value = 1
        
        txt(txt_����ҩʦ�����ʱ��).Enabled = chk(Index).value = 1
        
        If chk(chk_���ﴦ�����).value = 1 And chk(chk_סԺҩ�����).value = 1 Then
            strVar = "3"
        ElseIf chk(chk_���ﴦ�����).value = 0 And chk(chk_סԺҩ�����).value = 1 Then
            strVar = "2"
        ElseIf chk(chk_���ﴦ�����).value = 1 And chk(chk_סԺҩ�����).value = 0 Then
            strVar = "1"
        Else
            strVar = "0"
        End If
        Call SetParChange(chk, Index, mrsPar, True, strVar)
        chk(chk_סԺҩ�����).ForeColor = chk(Index).ForeColor
            
    Case chk_סԺҩ�����
        If Me.Visible = False Then Exit Sub
        
        If chk(Index).value = 0 Then
            '������δ���ļ�¼
            If GetRecipeAuditBills(2) Then
                MsgBox "�������ϵͳ�������δ���ļ�¼�����飡", vbInformation, gstrSysName
                chk(Index).value = 1
            End If
        End If
        
        If chk(Index).value = 0 Then chk(chk_����סԺҽ��).value = 0
        chk(chk_����סԺҽ��).Enabled = chk(Index).value = 1
        
        If chk(chk_���ﴦ�����).value = 1 And chk(chk_סԺҩ�����).value = 1 Then
            strVar = "3"
        ElseIf chk(chk_���ﴦ�����).value = 0 And chk(chk_סԺҩ�����).value = 1 Then
            strVar = "2"
        ElseIf chk(chk_���ﴦ�����).value = 1 And chk(chk_סԺҩ�����).value = 0 Then
            strVar = "1"
        Else
            strVar = "0"
        End If
        Call SetParChange(chk, chk_���ﴦ�����, mrsPar, True, strVar)
        chk(Index).ForeColor = chk(chk_���ﴦ�����).ForeColor
        
    Case chk_���ﴦ���Զ�����
        If Me.Visible = False Then Exit Sub
        
        Call SetParChange(chk, chk_���ﴦ���Զ�����, mrsPar)
    End Select

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check���쵥() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
        " Where a.���� = 6 And a.�������� > Sysdate - 90 And Exists " & _
        " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Nvl(b.��ҩ��ʽ,0) = 1) And Rownum < 2"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵����쵥")
    
    Check���쵥 = rsTemp.RecordCount = 0
End Function

Private Function Check�ƿⵥ() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ A " & _
        " Where a.���� = 6 And a.�������� > Sysdate - 90 And Exists " & _
        " (Select 1 From ҩƷ�շ���¼ B Where a.�շ�id = b.Id And Nvl(b.��ҩ��ʽ,0) <> 1) And Rownum < 2"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵��ƿⵥ")
    
    Check�ƿⵥ = rsTemp.RecordCount = 0
End Function

Private Function Check���õ�() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From δ��ҩƷ��¼ Where ���� = 7 And �������� > Sysdate - 90 And Rownum < 2"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵����õ�")
    
    Check���õ� = rsTemp.RecordCount = 0
End Function
Private Sub cmdlst��Һ���ķ�ҩ���˿���_Click(Index As Integer)
    Dim i As Long
    
    If chk��Դ����.value = 0 Then Exit Sub
    
    With lst(lst_PIVA��Դ����)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '������lst_ItemCheck�¼�
        Next
    End With
End Sub

Private Function GetRecipeAuditBills(ByVal bytType As Byte) As Boolean
'���ܣ������������סԺ�ġ���������¼���Ƿ����δ���ļ�¼
'������
'  bytType��1-���2-סԺ
'���أ�True����δ���ļ�¼��False������δ���ļ�¼

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If bytType = 1 Then
        '����
        strSql = "Select ID From ��������¼ Where ״̬ = 0 And �ύʱ�� >= Trunc(Sysdate - [1]) And �Һ�Id Is Not Null And Rownum < 2 "
    Else
        'סԺ
        strSql = "Select ID From ��������¼ Where ״̬ = 0 And �ύʱ�� >= Trunc(Sysdate - [1]) And ��ҳId Is Not Null And Rownum < 2 "
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "���δ���Ĵ�������¼", IIF(bytType = 1, 3, 5))
    GetRecipeAuditBills = rsTemp.EOF = False
    rsTemp.Close
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Sub vsf�Ա�ҩ�嵥_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsf�Ա�ҩ�嵥.ColIndex("ҩƷ���������") Then Cancel = True
End Sub

Private Sub vsf�Ա�ҩ�嵥_Click()
    With vsf�Ա�ҩ�嵥
        If .Row < 1 Then Exit Sub
        
        If .Col = .ColIndex("�����") And .TextMatrix(.Row, .ColIndex("ҩƷid")) <> "" Then
            If .TextMatrix(.Row, .ColIndex("�����")) = "" Then
                .TextMatrix(.Row, .ColIndex("�����")) = "��"
            Else
                .TextMatrix(.Row, .ColIndex("�����")) = ""
            End If
        End If
    End With
End Sub

Private Sub vsf�Ա�ҩ�嵥_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        With vsf�Ա�ҩ�嵥
            If .Rows = 2 Then
                .TextMatrix(.Row, .ColIndex("ҩƷid")) = ""
                .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = ""
            Else
                .RemoveItem vsf�Ա�ҩ�嵥.Row
            End If
        End With
    End If
End Sub

Private Sub vsf�Ա�ҩ�嵥_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim strCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsf�Ա�ҩ�嵥.hwnd)
        dblLeft = vRect.Left + vsf�Ա�ҩ�嵥.CellLeft
        dblTop = vRect.Top + vsf�Ա�ҩ�嵥.CellTop + vsf�Ա�ҩ�嵥.CellHeight + 3200
        
        With vsf�Ա�ҩ�嵥
            If Col = .ColIndex("ҩƷ���������") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '������
                    strCode = " d.���� like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '����ĸ
                    strCode = " n.���� Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '������
                    strCode = " d.���� like [1] "
                Else
                    strCode = " (n.���� Like [1] Or d.���� Like [1] Or n.���� Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'��' || d.���� || '��' || d.���� || '(' || d.��� || ')' As ͨ����" & vbNewLine & _
                    " From ҩƷ��� T, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    " Where t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And D.��� In ('5', '6') And" & strCode & vbNewLine & _
                    " And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '��' || d.���� || '��' || d.���� || '(' || d.��� || ')'"
                Set rsRecord = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ���������", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .Rows - 1
                        If rsRecord!ID = Val(.TextMatrix(i, .ColIndex("ҩƷID"))) Then
                            MsgBox rsRecord!ͨ���� & "�Ѿ�¼�룬������ѡ��", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("ҩƷID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("ҩƷ���������")) = rsRecord!ͨ����
                    .EditText = rsRecord!ͨ����
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

