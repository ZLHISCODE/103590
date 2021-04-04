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
   Caption         =   "药品参数设置"
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
   StartUpPosition =   1  '所有者中心
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
      TabCaption(0)   =   "通用(&0)"
      TabPicture(0)   =   "frmParMedicine.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "目录(&1)"
      TabPicture(1)   =   "frmParMedicine.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "入出(&2)"
      TabPicture(2)   =   "frmParMedicine.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "在库(&3)"
      TabPicture(3)   =   "frmParMedicine.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "处方(&4)"
      TabPicture(4)   =   "frmParMedicine.frx":68C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "部门(&5)"
      TabPicture(5)   =   "frmParMedicine.frx":68DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "静配(&6)"
      TabPicture(6)   =   "frmParMedicine.frx":68FA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "picPar(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "配药(&11)"
      TabPicture(7)   =   "frmParMedicine.frx":6916
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(11)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "精度(&12)"
      TabPicture(8)   =   "frmParMedicine.frx":6932
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(12)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "单位(&13)"
      TabPicture(9)   =   "frmParMedicine.frx":694E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "picPar(13)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "流向(&14)"
      TabPicture(10)  =   "frmParMedicine.frx":696A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "picPar(14)"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "库检(&15)"
      TabPicture(11)  =   "frmParMedicine.frx":6986
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "picPar(15)"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "环节(&16)"
      TabPicture(12)  =   "frmParMedicine.frx":69A2
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "picPar(16)"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "处方审查(&7)"
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
            Caption         =   "领用库房控制"
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
            Caption         =   "库房之间流向控制"
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf库房计量单位 
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
            FormatString    =   "药品库房|售价单位|门诊单位|住院单位|药库单位"
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VB.Label lblUnits 
            Caption         =   "药品库房的计量单位（双击鼠标设置）"
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
         Begin ZL9BillEdit.BillEdit Bill药品卫材精度 
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
            Caption         =   "药品精度设置：按包装单位来设置价格、数量允许录入的精度（保留的小数位数）"
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
         Begin ZL9BillEdit.BillEdit Bill药房配药控制 
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
            Caption         =   "药房配药控制"
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
            TabCaption(0)   =   "基础设置(&1)"
            TabPicture(0)   =   "frmParMedicine.frx":6ABB
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fra配药控制"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "fra输液医嘱期效"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "frmParMedicine"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "frmType"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "其他设置(&2)"
            TabPicture(1)   =   "frmParMedicine.frx":6AD7
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "picPRI"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "fra输液配置"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "fra(0)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "frmMoney"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "自备药设置(&3)"
            TabPicture(2)   =   "frmParMedicine.frx":6AF3
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fra自备药清单设置"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame fra自备药清单设置 
               Caption         =   " 自备药清单设置 "
               ForeColor       =   &H00800000&
               Height          =   6855
               Left            =   -74880
               TabIndex        =   264
               Top             =   360
               Width           =   10095
               Begin VSFlex8Ctl.VSFlexGrid vsf自备药清单 
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
                     Name            =   "宋体"
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
                  Caption         =   "  允许下列药品在静配中心不接受自备药的情况下，可以通过发自备药的方式将下列药品发送至静配中心。"
                  ForeColor       =   &H00000080&
                  Height          =   180
                  Left            =   120
                  TabIndex        =   266
                  Top             =   240
                  Width           =   8460
               End
            End
            Begin VB.Frame frmType 
               Caption         =   " 医嘱执行性质选择"
               ForeColor       =   &H00800000&
               Height          =   705
               Left            =   4680
               TabIndex        =   237
               Top             =   480
               Width           =   6015
               Begin VB.CheckBox chk 
                  Caption         =   "不取药"
                  Height          =   255
                  Index           =   62
                  Left            =   1740
                  TabIndex        =   240
                  Top             =   330
                  Width           =   885
               End
               Begin VB.CheckBox chk 
                  Caption         =   "自备药"
                  Height          =   255
                  Index           =   61
                  Left            =   240
                  TabIndex        =   239
                  Top             =   330
                  Width           =   885
               End
               Begin VB.CheckBox chk 
                  Caption         =   "离院带药"
                  Height          =   255
                  Index           =   63
                  Left            =   3240
                  TabIndex        =   238
                  Top             =   330
                  Width           =   1125
               End
            End
            Begin VB.Frame frmParMedicine 
               Caption         =   " 操作控制 "
               ForeColor       =   &H00800000&
               Height          =   5775
               Left            =   120
               TabIndex        =   217
               Top             =   1200
               Width           =   4455
               Begin VB.CheckBox chk 
                  Caption         =   "特殊药品按药品类型指定批次"
                  Height          =   255
                  Index           =   51
                  Left            =   240
                  TabIndex        =   226
                  Top             =   2505
                  Width           =   3135
               End
               Begin VB.CheckBox chk 
                  Caption         =   "输液单按批次，药品规则排序"
                  Height          =   255
                  Index           =   47
                  Left            =   240
                  TabIndex        =   225
                  Top             =   1935
                  Width           =   3375
               End
               Begin VB.CheckBox chk 
                  Caption         =   "条码扫描一次自动发送"
                  Height          =   255
                  Index           =   46
                  Left            =   240
                  TabIndex        =   224
                  Top             =   1365
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "允许手工调整批次（排药印签环节）"
                  Height          =   255
                  Index           =   33
                  Left            =   240
                  TabIndex        =   223
                  Top             =   240
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "允许调整打包状态（排药印签、配药环节）"
                  Height          =   255
                  Index           =   34
                  Left            =   240
                  TabIndex        =   222
                  Top             =   813
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "配置费按病人收取(一个病人一天只收一个费用)"
                  Height          =   255
                  Index           =   49
                  Left            =   240
                  TabIndex        =   221
                  Top             =   3075
                  Width           =   4095
               End
               Begin VB.CheckBox chk 
                  Caption         =   "出院病人不收配置费"
                  Height          =   255
                  Index           =   50
                  Left            =   240
                  TabIndex        =   220
                  Top             =   3645
                  Width           =   1935
               End
               Begin VB.CheckBox chk 
                  Caption         =   "打包药品在发送环节收取配置费"
                  Height          =   255
                  Index           =   59
                  Left            =   240
                  TabIndex        =   219
                  Top             =   4230
                  Width           =   3495
               End
               Begin VB.CheckBox chk 
                  Caption         =   "打印瓶签时填写各个环节的实际操作员"
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
                  ToolTipText     =   "双击或按回车键确认"
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
                     Text            =   "名称"
                     Object.Width           =   3528
                  EndProperty
               End
            End
            Begin VB.Frame fra输液配置 
               Caption         =   " 在输液配制中心发药的病人科室 "
               ForeColor       =   &H00800000&
               Height          =   6735
               Left            =   -69720
               TabIndex        =   205
               Top             =   360
               Width           =   5175
               Begin VB.CheckBox chk来源科室 
                  Caption         =   "启用来源病区控制"
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
               Begin VB.CommandButton cmdlst输液中心发药病人科室 
                  Caption         =   "全选"
                  Height          =   350
                  Index           =   0
                  Left            =   2760
                  TabIndex        =   207
                  Top             =   6240
                  Width           =   1100
               End
               Begin VB.CommandButton cmdlst输液中心发药病人科室 
                  Caption         =   "全清"
                  Height          =   350
                  Index           =   1
                  Left            =   3960
                  TabIndex        =   206
                  Top             =   6240
                  Width           =   1100
               End
               Begin VB.Label lbl来源科室 
                  Caption         =   "  启用时可选择病区。输液医嘱发送时如果病人的所在病区没有选择，则不会产生输液单据。"
                  ForeColor       =   &H00000080&
                  Height          =   420
                  Left            =   240
                  TabIndex        =   210
                  Top             =   360
                  Width           =   4560
               End
            End
            Begin VB.Frame fra 
               Caption         =   " 静配给药途径选择"
               ForeColor       =   &H00800000&
               Height          =   3135
               Index           =   0
               Left            =   -74760
               TabIndex        =   201
               Top             =   360
               Width           =   4935
               Begin VB.CheckBox chk给药途径 
                  Caption         =   "启用输液给药途径控制"
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
               Begin VB.Label lbl给药途径 
                  Caption         =   "  启用时可选择下列输液类的给药途径。输液医嘱发送时如果医嘱的给药途径没有选择，则不会产生输液单据。"
                  ForeColor       =   &H00000080&
                  Height          =   540
                  Left            =   120
                  TabIndex        =   204
                  Top             =   240
                  Width           =   4080
               End
            End
            Begin VB.Frame frmMoney 
               Caption         =   "配置费设置"
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
                  TabCaption(0)   =   "配药类型"
                  TabPicture(0)   =   "frmParMedicine.frx":D539
                  Tab(0).ControlEnabled=   -1  'True
                  Tab(0).Control(0)=   "VSFPrice"
                  Tab(0).Control(0).Enabled=   0   'False
                  Tab(0).ControlCount=   1
                  TabCaption(1)   =   "给药途径(只支持静脉营养类型)"
                  TabPicture(1)   =   "frmParMedicine.frx":D555
                  Tab(1).ControlEnabled=   0   'False
                  Tab(1).Control(0)=   "VSFPrice_给药途径"
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
                        Name            =   "宋体"
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
                  Begin VSFlex8Ctl.VSFlexGrid VSFPrice_给药途径 
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
                        Name            =   "宋体"
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
                  Caption         =   "向下(&N)"
                  Height          =   350
                  Left            =   2280
                  TabIndex        =   244
                  Top             =   3000
                  Width           =   1100
               End
               Begin VB.CommandButton cmdLast 
                  Caption         =   "向上(&S)"
                  Enabled         =   0   'False
                  Height          =   350
                  Left            =   600
                  TabIndex        =   243
                  Top             =   3000
                  Width           =   1100
               End
               Begin VB.Label lblprice 
                  Caption         =   "设置输液单配药类型对应的收费项目，在配药的时候根据设置的规则进行收费，且优先按给药途径方式收取配置费"
                  DragMode        =   1  'Automatic
                  ForeColor       =   &H00000080&
                  Height          =   420
                  Left            =   240
                  TabIndex        =   200
                  Top             =   240
                  Width           =   4560
               End
            End
            Begin VB.Frame fra输液医嘱期效 
               Caption         =   " 启用输液配制中心的医嘱期效"
               ForeColor       =   &H00800000&
               Height          =   705
               Left            =   120
               TabIndex        =   195
               Top             =   480
               Width           =   4455
               Begin VB.OptionButton opt输液医嘱期效 
                  Caption         =   "长嘱"
                  Height          =   180
                  Index           =   1
                  Left            =   240
                  TabIndex        =   198
                  Top             =   330
                  Width           =   680
               End
               Begin VB.OptionButton opt输液医嘱期效 
                  Caption         =   "临嘱"
                  Height          =   180
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   197
                  Top             =   330
                  Width           =   680
               End
               Begin VB.OptionButton opt输液医嘱期效 
                  Caption         =   "长嘱和临嘱"
                  Height          =   180
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   196
                  Top             =   330
                  Value           =   -1  'True
                  Width           =   1200
               End
            End
            Begin VB.Frame fra配药控制 
               Caption         =   " 配药流程、操作控制"
               ForeColor       =   &H00800000&
               Height          =   5775
               Left            =   4680
               TabIndex        =   187
               Top             =   1200
               Width           =   6015
               Begin VB.CheckBox chk 
                  Caption         =   "输液单摆药后临床不允许改变打包状态"
                  Height          =   255
                  Index           =   64
                  Left            =   240
                  TabIndex        =   249
                  Top             =   5400
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "当天发送的医嘱产生的输液单全部到备用批次"
                  Height          =   255
                  Index           =   57
                  Left            =   240
                  TabIndex        =   216
                  Top             =   4849
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "自动排批时输液单的批次只往后面批次变动"
                  Height          =   255
                  Index           =   56
                  Left            =   240
                  TabIndex        =   215
                  Top             =   3753
                  Width           =   5535
               End
               Begin VB.CheckBox chk 
                  Caption         =   "同一输液单保持上次配药批次"
                  Height          =   255
                  Index           =   39
                  Left            =   240
                  TabIndex        =   194
                  Top             =   773
                  Width           =   2655
               End
               Begin VB.CheckBox chk 
                  Caption         =   "配制中心不接收的静脉营养医嘱在病区配制"
                  Height          =   255
                  Index           =   42
                  Left            =   240
                  TabIndex        =   193
                  Top             =   1869
                  Width           =   4125
               End
               Begin VB.CheckBox chk 
                  Caption         =   "配液输液单配药后允许销帐申请"
                  Height          =   255
                  Index           =   41
                  Left            =   240
                  TabIndex        =   192
                  Top             =   1321
                  Width           =   3975
               End
               Begin VB.CheckBox chk 
                  Caption         =   "输液配置中心首次执行的医嘱需要进行审核"
                  Height          =   240
                  Index           =   83
                  Left            =   240
                  TabIndex        =   191
                  Top             =   240
                  Width           =   3855
               End
               Begin VB.CheckBox chk 
                  Caption         =   "启用自动排批（启用自动排批后，将不再保持上次批次）"
                  Height          =   255
                  Index           =   54
                  Left            =   240
                  TabIndex        =   190
                  Top             =   3205
                  Width           =   5535
               End
               Begin VB.CheckBox chk 
                  Caption         =   "单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包"
                  Height          =   495
                  Index           =   52
                  Left            =   240
                  TabIndex        =   189
                  Top             =   2417
                  Width           =   5655
               End
               Begin VB.CheckBox chk 
                  Caption         =   "不允许置换药房到输液配置中心"
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
            Caption         =   "是否可以销帐拒绝"
            Height          =   180
            Index           =   53
            Left            =   240
            TabIndex        =   245
            Top             =   3000
            Width           =   4095
         End
         Begin VB.CheckBox chk 
            Caption         =   "退药待发单据默认为发药状态"
            Height          =   180
            Index           =   43
            Left            =   240
            TabIndex        =   242
            Top             =   2640
            Width           =   4095
         End
         Begin VB.CheckBox chk 
            Caption         =   "发药时审核医嘱"
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
            Caption         =   "显示库房货位及库存限量提示"
            Height          =   180
            Index           =   27
            Left            =   240
            TabIndex        =   157
            Top             =   825
            Width           =   2745
         End
         Begin VB.Frame fra签名 
            Caption         =   " 药房人员签名设置"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   240
            TabIndex        =   154
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "领药人签名"
               Height          =   255
               Index           =   31
               Left            =   150
               TabIndex        =   156
               Top             =   285
               Width           =   1485
            End
            Begin VB.CheckBox chk 
               Caption         =   "退药人签名"
               Height          =   255
               Index           =   32
               Left            =   1710
               TabIndex        =   155
               Top             =   285
               Width           =   1485
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "是否自动缺药检查"
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
               Name            =   "宋体"
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
         Begin VB.CheckBox chk自动刷新 
            Caption         =   "自动刷新未发药清单"
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   1155
            Width           =   1935
         End
         Begin VB.CheckBox chk 
            Caption         =   "发药时汇总退药销帐记录"
            Height          =   180
            Index           =   29
            Left            =   240
            TabIndex        =   150
            Top             =   1575
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "退药销账时允许审核出院病人的销账申请"
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
         Begin VB.Label lbl查询天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "默认查询天数"
            Height          =   180
            Left            =   240
            TabIndex        =   161
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "分钟"
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
            Caption         =   "配药时对未收费的单据进行收费"
            Height          =   225
            Index           =   23
            Left            =   240
            TabIndex        =   228
            Top             =   2040
            Width           =   2820
         End
         Begin VB.Frame fra未收费发药 
            Caption         =   " 未收费或审核时发药"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   120
            TabIndex        =   144
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "允许未审核的记帐处方发药"
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   147
               Top             =   1320
               Value           =   1  'Checked
               Width           =   3345
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许未收费的门诊划价处方发药"
               Height          =   195
               Index           =   58
               Left            =   120
               TabIndex        =   146
               Top             =   960
               Width           =   2880
            End
            Begin VB.CheckBox chk 
               Caption         =   "项目执行前先收费或审核"
               Height          =   195
               Index           =   74
               Left            =   720
               TabIndex        =   145
               Top             =   0
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.Label lbl未收费发药 
               Caption         =   "  如果启用了门诊一卡通参数""执行前必须先收费或先记帐审核""，则对门诊病人发药时，以下参数将失效。"
               ForeColor       =   &H00000080&
               Height          =   615
               Left            =   120
               TabIndex        =   148
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame fra 
            Caption         =   " 发药窗口动态分配 "
            ForeColor       =   &H00800000&
            Height          =   825
            Index           =   3
            Left            =   120
            TabIndex        =   141
            Top             =   5160
            Width           =   3975
            Begin VB.OptionButton opt发药窗口 
               Caption         =   "闲忙方式"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   143
               Top             =   360
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton opt发药窗口 
               Caption         =   "平均方式"
               Height          =   210
               Index           =   1
               Left            =   1560
               TabIndex        =   142
               Top             =   360
               Width           =   1020
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "收费或记账指定药房时限定药品库存"
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   140
            Top             =   435
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品收费完成后自动发药"
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
         Begin VB.Frame fra验证方式 
            Caption         =   " 药房人员验证方式 "
            ForeColor       =   &H00800000&
            Height          =   765
            Left            =   5280
            TabIndex        =   133
            Top             =   4440
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "校验配药人"
               Height          =   195
               Index           =   20
               Left            =   240
               TabIndex        =   135
               Top             =   360
               Width           =   1200
            End
            Begin VB.CheckBox chk 
               Caption         =   "校验发药人"
               Height          =   195
               Index           =   24
               Left            =   1830
               TabIndex        =   134
               Top             =   360
               Width           =   1200
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "退药时自动将记帐费用销帐"
            Height          =   225
            Index           =   13
            Left            =   240
            TabIndex        =   132
            Top             =   750
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "发药时刷就诊卡验证"
            Height          =   225
            Index           =   16
            Left            =   240
            TabIndex        =   131
            Top             =   1050
            Width           =   3540
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品医嘱按发生时间过滤"
            Height          =   225
            Index           =   18
            Left            =   240
            TabIndex        =   130
            Top             =   1365
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用病人实际取药确认模式"
            Height          =   225
            Index           =   19
            Left            =   240
            TabIndex        =   129
            Top             =   1680
            Width           =   2460
         End
         Begin VB.Frame fraSetColor 
            Caption         =   "  处方颜色设置"
            ForeColor       =   &H00800000&
            Height          =   4125
            Left            =   5280
            TabIndex        =   113
            Top             =   120
            Width           =   3915
            Begin VB.CommandButton cmdDefaultColor 
               BackColor       =   &H00000000&
               Caption         =   "恢复默认颜色(&R)"
               Height          =   300
               Left            =   480
               MaskColor       =   &H00000000&
               TabIndex        =   121
               Top             =   3600
               Width           =   2175
            End
            Begin VB.PictureBox pic处方颜色 
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
            Begin VB.PictureBox pic处方颜色 
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
            Begin VB.PictureBox pic处方颜色 
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
            Begin VB.PictureBox pic处方颜色 
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
            Begin VB.PictureBox pic处方颜色 
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
            Begin VB.PictureBox pic处方颜色 
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
               Text            =   "存参数原始值"
               Top             =   3120
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "说明：双击颜色标签定义处方颜色！"
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   240
               TabIndex        =   128
               Top             =   360
               Width           =   2880
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "麻醉"
               Height          =   180
               Index           =   5
               Left            =   240
               TabIndex        =   127
               Top             =   3120
               Width           =   360
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "精神I类"
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   126
               Top             =   2670
               Width           =   630
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "精神II类"
               Height          =   180
               Index           =   3
               Left            =   240
               TabIndex        =   125
               Top             =   2205
               Width           =   720
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "急诊"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   124
               Top             =   1755
               Width           =   360
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "儿科"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   123
               Top             =   1290
               Width           =   360
            End
            Begin VB.Label lbl处方类型 
               AutoSize        =   -1  'True
               Caption         =   "普通"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   122
               Top             =   840
               Width           =   360
            End
         End
         Begin VB.Label lbl查询未发药单据天数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查询未发药单据天数"
            Height          =   180
            Left            =   240
            TabIndex        =   248
            Top             =   2940
            Width           =   1620
         End
         Begin VB.Label lblMax 
            AutoSize        =   -1  'True
            Caption         =   "大处方审核标准值"
            Height          =   180
            Left            =   240
            TabIndex        =   138
            Top             =   2460
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "元"
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
         Begin VB.Frame fra调价控制 
            Caption         =   " 调价控制"
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   120
            TabIndex        =   110
            Top             =   2040
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "成本价按库房批次调整"
               Height          =   255
               Index           =   66
               Left            =   120
               TabIndex        =   251
               Top             =   660
               Width           =   2850
            End
            Begin VB.CheckBox chk 
               Caption         =   "新成本价、新售价超过限价时提示"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   112
               Top             =   960
               Width           =   3090
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价药品按批次调价"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   111
               Top             =   360
               Width           =   2010
            End
         End
         Begin VB.Frame fra盘点控制 
            Caption         =   " 盘点控制"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   120
            TabIndex        =   106
            Top             =   120
            Width           =   4215
            Begin VB.CheckBox chk 
               Caption         =   "盘亏减可用数量检查"
               Height          =   255
               Index           =   65
               Left            =   120
               TabIndex        =   250
               Top             =   1320
               Width           =   1920
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许盘点停用药品"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   109
               Top             =   1005
               Width           =   1920
            End
            Begin VB.CheckBox chk 
               Caption         =   "忽略药品服务对象"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   108
               Top             =   682
               Width           =   2040
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许盘点没有设置存储库房的药品"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   107
               Top             =   360
               Width           =   3360
            End
         End
         Begin VB.Frame fra质量控制 
            Caption         =   " 质量控制"
            ForeColor       =   &H00800000&
            Height          =   1365
            Left            =   120
            TabIndex        =   103
            Top             =   3600
            Width           =   4245
            Begin VB.CheckBox chk 
               Caption         =   "药品质量管理审核时同步减少库存"
               Height          =   180
               Index           =   12
               Left            =   120
               TabIndex        =   104
               Top             =   360
               Width           =   3495
            End
            Begin VB.Label Label4 
               Caption         =   "如果勾选此选项，相当于在审核后自动完成其他出库操作；要实现该功能，必须确保已先设置了其他出库的入出类别"
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
         Begin VB.Frame fra移库流程控制 
            Caption         =   " 出库业务，流程控制"
            ForeColor       =   &H00800000&
            Height          =   2445
            Left            =   120
            TabIndex        =   99
            Top             =   4800
            Width           =   4965
            Begin VB.CheckBox chk 
               Caption         =   "领用冲销时，需要先申请冲销,再审核冲销"
               Height          =   180
               Index           =   40
               Left            =   120
               TabIndex        =   241
               Top             =   2160
               Width           =   4095
            End
            Begin VB.CheckBox chk 
               Caption         =   "领用业务药品按批次填写出库单"
               Height          =   255
               Index           =   37
               Left            =   120
               TabIndex        =   232
               Top             =   720
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "移库业务药品按批次填写出库单"
               Height          =   255
               Index           =   26
               Left            =   120
               TabIndex        =   231
               Top             =   480
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "移库明确批次时允许补录产地批号"
               Height          =   255
               Index           =   76
               Left            =   120
               TabIndex        =   230
               Top             =   1050
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "申领业务药品按批次填写出库单"
               Height          =   255
               Index           =   75
               Left            =   120
               TabIndex        =   229
               Top             =   240
               Width           =   4080
            End
            Begin VB.CheckBox chk 
               Caption         =   "移库时需要备药、发送、接收这一过程。"
               Height          =   180
               Index           =   5
               Left            =   120
               TabIndex        =   101
               Top             =   1290
               Width           =   4000
            End
            Begin VB.CheckBox chk 
               Caption         =   "移库冲销时，移入库房需要先申请冲销"
               Height          =   180
               Index           =   6
               Left            =   120
               TabIndex        =   100
               Top             =   1920
               Width           =   4095
            End
            Begin VB.Label Label3 
               Caption         =   "如果不勾选，那么在填写移库单后，增加一个审核操作，审核后自动完成备药、发送、接收这一过程"
               ForeColor       =   &H00000080&
               Height          =   495
               Left            =   360
               TabIndex        =   102
               Top             =   1476
               Width           =   4185
            End
         End
         Begin VB.Frame fra成本价 
            Caption         =   " 自制入库成本价来源方式"
            ForeColor       =   &H00800000&
            Height          =   2205
            Left            =   5160
            TabIndex        =   96
            Top             =   5040
            Width           =   4905
            Begin VB.OptionButton opt自制入库成本来源 
               Caption         =   $"frmParMedicine.frx":D6C6
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   98
               Top             =   720
               Width           =   3015
            End
            Begin VB.OptionButton opt自制入库成本来源 
               Caption         =   "根据原料药品的成本价计算"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   97
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Frame fra外购入库参数 
            Caption         =   " 外购入库控制"
            ForeColor       =   &H00800000&
            Height          =   4575
            Left            =   120
            TabIndex        =   83
            Top             =   120
            Width           =   4935
            Begin VB.Frame fra上次采购信息 
               Caption         =   " 上次采购信息来源方式"
               ForeColor       =   &H00800000&
               Height          =   1365
               Left            =   120
               TabIndex        =   93
               Top             =   3120
               Width           =   4605
               Begin VB.CheckBox chk 
                  Caption         =   "优先取目录中的产地、批准文号"
                  Height          =   255
                  Index           =   35
                  Left            =   120
                  TabIndex        =   263
                  Top             =   360
                  Width           =   2880
               End
               Begin VB.OptionButton opt外购入库取成本价方式 
                  Caption         =   "优先从上一次入库业务中取成本价等信息"
                  Height          =   180
                  Index           =   1
                  Left            =   120
                  TabIndex        =   95
                  Top             =   1020
                  Width           =   3615
               End
               Begin VB.OptionButton opt外购入库取成本价方式 
                  Caption         =   "优先从当前库房的库存最近批次中取成本价等信息"
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
               Caption         =   "时价药品直接确定售价"
               Height          =   195
               Index           =   36
               Left            =   120
               TabIndex        =   92
               Top             =   960
               Width           =   2280
            End
            Begin VB.CheckBox chk 
               Caption         =   "需要经过核查后才能审核入库单"
               Height          =   195
               Index           =   28
               Left            =   120
               TabIndex        =   91
               Top             =   525
               Width           =   3000
            End
            Begin VB.CheckBox chk 
               Caption         =   "需要经过标记付款后才能进行付款管理"
               Height          =   195
               Index           =   70
               Left            =   120
               TabIndex        =   90
               Top             =   240
               Width           =   4440
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价药品通过加成率入库"
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   89
               Top             =   1530
               Width           =   2280
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价药品入库按扣前加成销售"
               Height          =   195
               Index           =   48
               Left            =   120
               TabIndex        =   88
               Top             =   2070
               Width           =   3090
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价药品通过分段加成入库"
               Height          =   180
               Index           =   14
               Left            =   120
               TabIndex        =   87
               Top             =   1815
               Width           =   2775
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价药品入库时取上次售价"
               Height          =   195
               Index           =   73
               Left            =   120
               TabIndex        =   86
               Top             =   1260
               Width           =   2760
            End
            Begin VB.CheckBox chk 
               Caption         =   "外购入库允许修改采购限价"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   85
               Top             =   2475
               Width           =   2835
            End
            Begin VB.CheckBox chk 
               Caption         =   "招标药品可选择非中标单位入库"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   84
               Top             =   2760
               Width           =   2880
            End
         End
         Begin VB.Frame fra供应商资质 
            Caption         =   " 外购入库供应商资质校验"
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
               Text            =   "存参数原始值"
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Frame fraCheck 
               Caption         =   " 选择校验方式"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   72
               Top             =   4080
               Width           =   4815
               Begin VB.OptionButton optCheck 
                  Caption         =   "校验未通过时禁止保存"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   74
                  Top             =   280
                  Width           =   2175
               End
               Begin VB.OptionButton optCheck 
                  Caption         =   "校验未通过时提醒"
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
                  Name            =   "宋体"
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
               Caption         =   "    药品外购入库编辑单据时是否校供应商的信息是否完整，及资质是否过期。请双击“校验”列打勾"
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
         Begin VB.Frame fra数字码 
            Caption         =   " 数字码设置"
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
            Begin VB.Label lbl数字码 
               AutoSize        =   -1  'True
               Caption         =   "数字码长度"
               Height          =   180
               Left            =   240
               TabIndex        =   68
               Top             =   420
               Width           =   900
            End
         End
         Begin VB.Frame fra售价方式 
            Caption         =   " 新增规格售价计算方式"
            ForeColor       =   &H00800000&
            Height          =   1215
            Left            =   120
            TabIndex        =   58
            Top             =   1920
            Width           =   4035
            Begin VB.OptionButton opt售价计算 
               Caption         =   "按一般加成率计算售价"
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   360
               Width           =   3615
            End
            Begin VB.OptionButton opt售价计算 
               Caption         =   "按分段加成计算售价"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   59
               Top             =   720
               Width           =   3735
            End
         End
         Begin VB.Frame frmStockRange 
            Caption         =   " 设置存储库房时允许应用于的范围"
            ForeColor       =   &H00800000&
            Height          =   3255
            Left            =   4800
            TabIndex        =   50
            Top             =   120
            Width           =   3585
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有当前分类下的药品(&6)"
               Height          =   225
               Index           =   5
               Left            =   120
               TabIndex        =   56
               Top             =   2280
               Value           =   1  'Checked
               Width           =   2985
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有同级的药品(&5)"
               Height          =   225
               Index           =   4
               Left            =   120
               TabIndex        =   55
               Top             =   1905
               Value           =   1  'Checked
               Width           =   2745
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有当前选择的同剂型药品(&4)"
               Height          =   225
               Index           =   3
               Left            =   120
               TabIndex        =   54
               Top             =   1530
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有当前选择的同材质药品(&3)"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   53
               Top             =   1155
               Value           =   1  'Checked
               Width           =   3285
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有当前选择的同品种药品(&2)"
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   52
               Top             =   780
               Value           =   1  'Checked
               Width           =   3270
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "仅应用于当前选择的药品(&1)"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   51
               Top             =   405
               Value           =   1  'Checked
               Width           =   2655
            End
            Begin VB.Label lblComment 
               Caption         =   "提示：没有选择到的应用范围在设置存储库房时将不能选择。"
               ForeColor       =   &H00000080&
               Height          =   405
               Left            =   120
               TabIndex        =   57
               Top             =   2640
               Width           =   2880
            End
         End
         Begin VB.Frame fraIncome 
            Caption         =   " 各材质对应缺省收入项目"
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
               Caption         =   "中药"
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
               Caption         =   "成药"
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
               Caption         =   "西药"
               Height          =   180
               Index           =   0
               Left            =   885
               TabIndex        =   49
               Top             =   390
               Width           =   360
            End
         End
         Begin VB.Frame fra分批 
            Caption         =   " 药品分批属性自动设置"
            ForeColor       =   &H00800000&
            Height          =   1800
            Left            =   120
            TabIndex        =   42
            Top             =   3360
            Width           =   4035
            Begin VB.OptionButton opt分批 
               Caption         =   "仅药库分批"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton opt分批 
               Caption         =   "手工设置分批属性"
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   390
               Width           =   1735
            End
            Begin VB.OptionButton opt分批 
               Caption         =   "药库和药房分批"
               Height          =   200
               Index           =   2
               Left            =   120
               TabIndex        =   44
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton opt分批 
               Caption         =   "药库和药房都不分批"
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
         Begin VSFlex8Ctl.VSFlexGrid vsf单据环节控制 
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
               Name            =   "宋体"
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
            Caption         =   "单据环节控制：设置药品单据在特定业务环节中允许修改的项目"
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
            Caption         =   "启用门诊处方审查流程控制"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   184
            Top             =   240
            Width           =   2535
         End
         Begin VB.Frame fraOpporunity 
            Caption         =   "门诊药师审方的介入时机"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   720
            TabIndex        =   181
            Top             =   600
            Width           =   7335
            Begin VB.CheckBox chk 
               Caption         =   "审方合格确认后，启用自动发送处方功能"
               Height          =   180
               Index           =   22
               Left            =   2880
               TabIndex        =   227
               Top             =   360
               Width           =   3975
            End
            Begin VB.OptionButton optOpporunity 
               Caption         =   "门诊处方发送前"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   183
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton optOpporunity 
               Caption         =   "门诊药房配/发药前"
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
            Caption         =   "启用住院药嘱审查流程控制"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   179
            Top             =   3000
            Width           =   2535
         End
         Begin VB.CheckBox chk 
            Caption         =   "提醒门诊医生不合格医嘱（门诊医生开方保存后，是否开启提醒医生有问题的药嘱）"
            Height          =   255
            Index           =   44
            Left            =   720
            TabIndex        =   178
            Top             =   1920
            Width           =   6975
         End
         Begin VB.CheckBox chk 
            Caption         =   "提醒住院医生不合格医嘱（住院医生开方保存后，是否开启提醒医生有问题的药嘱）"
            Height          =   255
            Index           =   45
            Left            =   720
            TabIndex        =   177
            Top             =   3360
            Width           =   6975
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "门诊药师审方离岗时长      分钟（超出设定时长值未审查的处方，医师可发送通过，避免病人长时间滞留临床科室或药房）"
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
         Begin VB.Frame fra可用数量 
            Caption         =   " 药品可用数量动态计算"
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
            Begin VB.OptionButton opt可用数量 
               Caption         =   "不动态计算(始终以当前库存可用数量为准)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   259
               Top             =   360
               Value           =   -1  'True
               Width           =   5175
            End
            Begin VB.OptionButton opt可用数量 
               Caption         =   "超过指定月份的未发药数据不计算可用数量(启用该方案时，将在开具药品处方或医嘱时动态计算可用数量)"
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
            Begin MSComCtl2.UpDown ud可用数量 
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
               Caption         =   "月份"
               Height          =   180
               Left            =   360
               TabIndex        =   260
               Top             =   1245
               Width           =   360
            End
         End
         Begin VB.Frame fra零差价管理 
            Caption         =   " 药品零差价管理模式"
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
               Caption         =   "零差价模式"
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
            Caption         =   " 药品结存"
            ForeColor       =   &H00800000&
            Height          =   1860
            Index           =   10
            Left            =   120
            TabIndex        =   21
            Top             =   2760
            Width           =   3975
            Begin VB.Frame fra自动结存方式 
               Caption         =   " 设置结存时点"
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
                     Name            =   "宋体"
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
               Begin VB.OptionButton opt结存时间模式 
                  Caption         =   "每月最后一天"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   81
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton opt结存时间模式 
                  Caption         =   "每月    日"
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
            Begin VB.OptionButton opt结存方式 
               Caption         =   "自动结存(各库房按同一日期结存)"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   77
               Top             =   720
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton opt结存方式 
               Caption         =   "手工结存(各库房可以不同日期结存)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.Frame Fra药库流通 
            Caption         =   " 药库单据审核"
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
               Caption         =   "开单人与审核人"
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
            Caption         =   "药品出库优先算法"
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
            Caption         =   "药品效期显示方式"
            Height          =   180
            Index           =   31
            Left            =   120
            TabIndex        =   28
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "药价编辑设置单位"
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
            Caption         =   "药品编码递增模式"
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
         Begin MSComctlLib.ListView lvw库存检查 
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
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "编码"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "部门名称"
               Object.Width           =   4234
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "库存检查方式"
               Object.Width           =   4410
            EndProperty
         End
         Begin VB.Label lbl提示 
            Caption         =   "药品库存检查（双击鼠标或按C键设置）"
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
               Name            =   "宋体"
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
         Caption         =   "帮助(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11760
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
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
         Caption         =   "药房查找(&F)"
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
         Caption         =   "参数查找(&S)"
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

Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
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
    chk_限定药品的库存 = 2
    
    chk_首次医嘱执行需要审核 = 83
    
    chk_收费同时发药 = 17
    
    chk_项目执行前先收费或审核 = 74
    chk_未审核记帐处方发药 = 15
    chk_未收费处方发药 = 58
        
    chk_外购入库需要核查 = 28
    chk_外购入库需要经过标记付款后才能进行付款 = 70
    
    chk_时价分段加成入库 = 14
    chk_时价加成率入库 = 21
    chk_时价药品直接确定售价 = 36
    chk_时价入库按折扣前采购价加成销售 = 48
    chk_时价药品取上次售价 = 73
    
    chk_入库取目录中产地信息 = 35
                
    '模块参数
    '外购入库
    chk_允许修改采购限价 = 3
    chk_中标单位 = 4
    
    '申领
    chk_申领按批次出库 = 75
    
    '移库
    chk_移库流程 = 5
    chk_移库冲销申请 = 6
    chk_按批次出库时允许补录批号产地 = 76
    chk_移库按批次出库 = 26
    
    '领用
    chk_领用按批次出库 = 37
    chk_领用冲销申请 = 40
    
    '盘点
    chk_盘点存储库房 = 7
    chk_盘点忽略服务对象 = 8
    chk_盘点已停用药品 = 9
    chk_盘点盘亏减可用数量检查 = 65
    
    '调价
    chk_时价药品按批次调价 = 10
    chk_超过限价提示 = 11
    chk_成本价按库房批次调整 = 66
    
    '质量
    chk_审核时处理库存 = 12
    
    '处方发药
    chk_退药自动销账 = 13
    chk_发药刷卡 = 16
    chk_医嘱过滤 = 18
    chk_确认病人实际取药 = 19
    chk_校验配药人 = 20
    chk_校验发药人 = 24
    chk_配药时对未收费的单据进行收费 = 23
    
    '部门发药
    chk_缺药检查 = 25
    chk_库存限量 = 27
    chk_发药汇总退药销账 = 29
    chk_出院病人销账 = 30
    chk_住院药房领药人签名 = 31
    chk_住院药房退药人签名 = 32
    chk_发药时审核医嘱 = 38
    chk_退药待发单据默认为发药状态 = 43
    chk_是否可以销帐拒绝 = 53
    
    '静配
    chk_手工调整批次 = 33
    chk_手工调整打包 = 34
    chk_保持上次批次 = 39
    chk_配药后销账申请 = 41
    chk_TPN配置 = 42
    chk_扫描一次完成操作 = 46
    chk_输液单排序 = 47
    chk_配置费收取方式 = 49
    chk_出院病人是否收配置费 = 50
    chk_特殊药品批次 = 51
    chk_0批次规则 = 52
    chk_自动排批 = 54
    chk_不允许置换药房到输液配置中心 = 55
    chk_自动排批只往后面批次调整 = 56
    chk_当天发送的医嘱产生的输液单全部到备用批次 = 57
    chk_打包药品在发送环节收取配置费 = 59
    chk_打印瓶签时填写各个环节的实际操作员 = 60
    chk_自备药 = 61
    chk_不取药 = 62
    chk_离院带药 = 63
    chk_输液单摆药后临床不允许改变打包状态 = 64
    
    '处方审查
    chk_门诊处方审查 = 0
    chk_住院药嘱审查 = 1
    chk_门诊处方自动发送 = 22
    chk_提醒门诊医生 = 44
    chk_提醒住院医生 = 45
End Enum

Private Enum constCbo
    cbo_零差价模式 = 0
    cbo_定价单位 = 3
    cbo_西药收入项目 = 4
    cbo_成药收入项目 = 5
    cbo_中药收入项目 = 6
    cbo_药品单据审核 = 7
    cbo_药品编码模式 = 9
    cbo_效期显示方式 = 17
    cbo_药品出库优先算法 = 18
End Enum

Private Enum constListBox
    lst_PIVA来源科室 = 0
    lst_PIVA给药途径 = 1
End Enum

Private Enum constUd
    ud_数字码 = 0
    ud_部门发药查询天数 = 1
    ud_处方发药查询天数 = 2
End Enum

Private Enum constTxt
    txt_结存时间模式 = 0
        
    '药品外购入库
    txt_供应商资质 = 1
    
    '药品处方发药
    txt_处方颜色 = 2
    
    '药品部门发药
    txt_自动刷新时间 = 3
    
    '大处方审查
    txt_审查金额 = 4
    
    '处方审查
    txt_门诊药师审方离岗时长 = 5
    
    '基础
    txt_结存参数值 = 6
    
    txt_可用数量处理 = 7
End Enum

Private Enum constBill
    bill_药品库房流向 = 3
    bill_药品领用流向 = 4
End Enum

Private Enum constDigit
    dig_精度类别 = 0
    dig_精度内容 = 1
    dig_精度单位 = 2
    dig_精度 = 3
    dig_最小精度 = 4
    dig_最大精度 = 5
    dig_原始精度 = 6
    dig_类别 = 7
    dig_内容 = 8
    dig_单位 = 9
    dig_Cols = 10
End Enum

'药品卫材单据环节项目控制
'单据类型
Private Enum 单据
    药品外购 = 1
    卫材外购 = 15
End Enum

'业务环节
Private Enum 环节
    核查 = 1
    审核 = 2
    财务审核 = 3
End Enum

'处方类型：普通、急诊、儿科、麻醉、精一、精二
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

'默认处方颜色：普通－白色；急诊－淡黄色；儿科－淡绿色；麻醉、精一－淡红色；精二－白色
Private Const mconlng普通 = &HFFFFFF
Private Const mconlng儿科 = &HC0FFC0
Private Const mconlng急诊 = &HC0FFFF
Private Const mconlng精二 = &HFFFFFF
Private Const mconlng精一 = &HC0C0FF
Private Const mconlng麻醉 = &HC0C0FF

'允许控制的所有项目
Private Const cst所有项目 As String = "采购价,扣率,结算价,结算金额,售价,外观,发票号,发票代码,发票日期,发票金额"

'药品外购默认控制项目
Private Const cst药品外购项目_核查 As String = "结算价,采购价,售价,外观"
Private Const cst药品外购项目_审核 As String = "发票号,发票日期,发票金额"
Private Const cst药品外购项目_财务审核 As String = "采购价,扣率,结算价,结算金额,发票号,发票代码,发票日期,发票金额"

'卫材外购默认控制项目
Private Const cst卫材外购项目_核查 As String = "售价"
Private Const cst卫材外购项目_审核 As String = "采购价,扣率,结算价,结算金额,发票号,发票代码,发票日期,发票金额"
Private Const cst卫材外购项目_财务审核 As String = "结算价,结算金额"

Private Sub chk给药途径_Click()
    lst(lst_PIVA给药途径).Enabled = (chk给药途径.value = 1)
    lst(lst_PIVA给药途径).BackColor = IIF(lst(lst_PIVA给药途径).Enabled, &H80000005, &H8000000F)
    
    If Me.Visible And chk给药途径.value = 0 Then
        Call SetParChange(lst, lst_PIVA给药途径, mrsPar, True, "")
    End If
End Sub

Private Sub chk来源科室_Click()
    lst(lst_PIVA来源科室).Enabled = (chk来源科室.value = 1)
    lst(lst_PIVA来源科室).BackColor = IIF(lst(lst_PIVA来源科室).Enabled, &H80000005, &H8000000F)
    
    If Me.Visible And chk来源科室.value = 0 Then
        Call SetParChange(lst, lst_PIVA来源科室, mrsPar, True, "")
    End If
End Sub


Private Sub chk应用范围_Click(Index As Integer)
    If chk应用范围(Index).value <> Val(chk应用范围(Index).Tag) Then
        chk应用范围(Index).ForeColor = &HC0&             '修改后用朱红色前景色标识
    Else
        chk应用范围(Index).ForeColor = &H0&
    End If
End Sub

Private Sub chk应用范围_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk应用范围, 0, mrsPar, "", chk应用范围(Index))
End Sub


Private Sub chk自动刷新_Click()
    If chk自动刷新.value = 1 Then
        txt(txt_自动刷新时间).Enabled = True
    Else
        txt(txt_自动刷新时间).Text = "0"
        txt(txt_自动刷新时间).Enabled = False
    End If
    
    If Me.Visible Then
        Call SetParChange(txt, txt_自动刷新时间, mrsPar)
    End If
End Sub

Private Sub chk自动刷新_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, txt_自动刷新时间, mrsPar, "", chk自动刷新)
End Sub


Private Sub cmdDefaultColor_Click()
    Dim strColor As String
    Dim n As Integer
    
    Call Get处方默认颜色
    
    '处方颜色
    For n = 0 To pic处方颜色.UBound
        strColor = IIF(strColor = "", "", strColor & ";") & CStr(pic处方颜色(n).BackColor)
    Next
    
    If Me.Visible Then
        Call SetParChange(txt, txt_处方颜色, mrsPar, True, strColor)
    End If
    
    fraSetColor.ForeColor = txt(txt_处方颜色).ForeColor
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLast_Click()
    Dim intRow As Integer
    Dim str配药类型 As String
    Dim str收费项目 As String
    Dim lng项目id As Long
    
    With VSFPrice
        intRow = .Row
        If intRow < 2 Then Exit Sub
        lng项目id = .TextMatrix(.Row - 1, .ColIndex("项目id"))
        str配药类型 = .TextMatrix(.Row - 1, .ColIndex("配药类型"))
        str收费项目 = .TextMatrix(.Row - 1, .ColIndex("收费项目"))
        .TextMatrix(.Row - 1, .ColIndex("项目id")) = .TextMatrix(.Row, .ColIndex("项目id"))
        .TextMatrix(.Row - 1, .ColIndex("配药类型")) = .TextMatrix(.Row, .ColIndex("配药类型"))
        .TextMatrix(.Row - 1, .ColIndex("收费项目")) = .TextMatrix(.Row, .ColIndex("收费项目"))
        
        
        .TextMatrix(.Row, .ColIndex("项目id")) = lng项目id
        .TextMatrix(.Row, .ColIndex("配药类型")) = str配药类型
        .TextMatrix(.Row, .ColIndex("收费项目")) = str收费项目
        
        .Row = intRow - 1
    End With
End Sub

Private Sub cmdNext_Click()
    Dim intRow As Integer
    Dim str配药类型 As String
    Dim str收费项目 As String
    Dim lng项目id As Long
    
    With VSFPrice
        intRow = .Row
        If intRow = .Rows - 1 Then Exit Sub
        lng项目id = .TextMatrix(.Row + 1, .ColIndex("项目id"))
        str配药类型 = .TextMatrix(.Row + 1, .ColIndex("配药类型"))
        str收费项目 = .TextMatrix(.Row + 1, .ColIndex("收费项目"))
        .TextMatrix(.Row + 1, .ColIndex("项目id")) = .TextMatrix(.Row, .ColIndex("项目id"))
        .TextMatrix(.Row + 1, .ColIndex("配药类型")) = .TextMatrix(.Row, .ColIndex("配药类型"))
        .TextMatrix(.Row + 1, .ColIndex("收费项目")) = .TextMatrix(.Row, .ColIndex("收费项目"))
        
        
        .TextMatrix(.Row, .ColIndex("项目id")) = lng项目id
        .TextMatrix(.Row, .ColIndex("配药类型")) = str配药类型
        .TextMatrix(.Row, .ColIndex("收费项目")) = str收费项目
        
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
        If VSFPrice.Col = VSFPrice.ColIndex("收费项目") Then
            Me.VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.Col) = strReturn
            VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.ColIndex("项目id")) = strIds
        Else
            With Me.VSFPrice
                If VSFPrice.Col = .ColIndex("配药类型") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "该配药类型已经添加，请重新选择！", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
            End With
        End If
    ElseIf picPRI.Tag = 1 Then
        If VSFPrice_给药途径.Col = VSFPrice_给药途径.ColIndex("收费项目") Then
            Me.VSFPrice_给药途径.TextMatrix(VSFPrice_给药途径.Row, VSFPrice_给药途径.Col) = strReturn
            VSFPrice_给药途径.TextMatrix(VSFPrice_给药途径.Row, VSFPrice_给药途径.ColIndex("项目id")) = strIds
        Else
            With VSFPrice_给药途径
                If VSFPrice_给药途径.Col = .ColIndex("给药途径") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "该给药途径已经添加，请重新选择！", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
                .TextMatrix(.Row, .ColIndex("诊疗id")) = strIds
            End With
        End If
    End If
    
End Sub



Private Sub Form_Activate()
    If Me.Tag = "初始成功" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    Dim objPic As PictureBox
    
    '窗口大小：13000,8385
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
    
    With VSFPrice_给药途径
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
    tabDesign.Visible = False
    
    strCategory = "参数设置,基础项目"
    
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "100,0,药品通用设置;101,1,药品目录管理;110,2,药品入出管理;111,3,药品在库管理;112,4,药品处方发药;113,5,药品部门发药;114,6,配置中心管理;115,7,处方审查管理"
    
    '二级分类Pickture索引从11开始排
    marrFunc(1) = "101,11,药房配药控制;102,12,药品录入精度;107,13,药品计量单位;105,14,药品流向控制;106,15,药品库存检查;108,16,单据环节控制"
    
    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    Me.Tag = "初始成功"
End Sub

Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIds As String) As String
    'intType:0-双击列表时；1-列表中按回车时
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
        If VSFPrice.Col = VSFPrice.ColIndex("收费项目") Then
            Me.VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.Col) = strReturn
            VSFPrice.TextMatrix(VSFPrice.Row, VSFPrice.ColIndex("项目id")) = strIds
        Else
            With Me.VSFPrice
                If VSFPrice.Col = .ColIndex("配药类型") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "该配药类型已经添加，请重新选择！", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
            End With
        End If
    ElseIf picPRI.Tag = 1 Then
        If VSFPrice_给药途径.Col = VSFPrice_给药途径.ColIndex("收费项目") Then
            Me.VSFPrice_给药途径.TextMatrix(VSFPrice_给药途径.Row, VSFPrice_给药途径.Col) = strReturn
            VSFPrice_给药途径.TextMatrix(VSFPrice_给药途径.Row, VSFPrice_给药途径.ColIndex("项目id")) = strIds
        Else
            With VSFPrice_给药途径
                If VSFPrice_给药途径.Col = .ColIndex("给药途径") Then
                    For i = 1 To .Rows - 1
                        If strReturn = .TextMatrix(i, .Col) Then
                            MsgBox "该给药途径已经添加，请重新选择！", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(.Row, .Col) = strReturn
                .TextMatrix(.Row, .ColIndex("诊疗id")) = strIds
            End With
        End If
    End If
    
End Sub

Private Sub optCheck_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_供应商资质, mrsPar, True, Get供应商资质校验)
    End If
    
    fra供应商资质.ForeColor = txt(txt_供应商资质).ForeColor
End Sub

Private Sub optOpporunity_Click(Index As Integer)
    If Me.Visible Then
        '检查最近未审查的记录
        If mblnSecondOpporunity Then        'mblnSecondOpporunity 控制二次显示Msgbox
            mblnSecondOpporunity = False
            Exit Sub
        End If
        If GetRecipeAuditBills(1) Then
            mblnSecondOpporunity = True
            MsgBox "处方审查系统最近存在未审查的记录，请检查！", vbInformation, gstrSysName
            If Val(fraOpporunity.Tag) = 2 Then
                optOpporunity(2).value = True
            Else
                optOpporunity(1).value = True
            End If
            Exit Sub
        Else
            Me.fraOpporunity.Tag = CStr(Index)
        End If
        
        chk(chk_门诊处方自动发送).Enabled = optOpporunity(1).value And optOpporunity(1).Enabled
        If chk(chk_门诊处方自动发送).Enabled = False Then chk(chk_门诊处方自动发送).value = 0
        
        Call SetParChange(optOpporunity, Index, mrsPar, True, IIF(optOpporunity(1).value, 1, 2))
    End If
End Sub

Private Sub optOpporunity_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optOpporunity, Index, mrsPar)
End Sub


Private Sub opt结存方式_Click(Index As Integer)
    Dim strValue As String
    
    If opt结存方式(0).value = True Then
        opt结存时间模式(0).Enabled = False
        opt结存时间模式(1).Enabled = False
        txt(txt_结存时间模式).Enabled = False
        
        '手工结存参数值为-1
        strValue = "-1"
    Else
        opt结存时间模式(0).Enabled = True
        opt结存时间模式(1).Enabled = True
        txt(txt_结存时间模式).Enabled = opt结存时间模式(1).value
        
        strValue = IIF(opt结存时间模式(0).value, 0, Val(txt(txt_结存时间模式).Text))
    End If
    
    If Me.Visible Then
        Call SetParChange(txt, txt_结存参数值, mrsPar, True, strValue)
        
        opt结存方式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存方式(1).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(1).ForeColor = txt(txt_结存参数值).ForeColor
        txt(txt_结存时间模式).ForeColor = opt结存时间模式(1).ForeColor
    End If
End Sub

Private Sub opt可用数量_Click(Index As Integer)
    If Index = 0 Then
        txtM.Enabled = False
        ud可用数量.Enabled = False
        Call SetParChange(txt, txt_可用数量处理, mrsPar, True, 0)
    Else
        txtM.Enabled = True
        ud可用数量.Enabled = True
        Call SetParChange(txt, txt_可用数量处理, mrsPar, True, Val(txtM.Text))
    End If
End Sub

Private Sub opt自制入库成本来源_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt自制入库成本来源, Index, mrsPar)
    End If
End Sub

Private Sub opt自制入库成本来源_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt自制入库成本来源, Index, mrsPar)
End Sub


Private Sub opt分批_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt分批, Index, mrsPar)
    End If
End Sub

Private Sub opt分批_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt分批, Index, mrsPar)
End Sub


Private Sub opt售价计算_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt售价计算, Index, mrsPar)
    End If
End Sub

Private Sub opt售价计算_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call SetParTip(opt售价计算, Index, mrsPar)
End Sub


Private Sub opt外购入库取成本价方式_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt外购入库取成本价方式, Index, mrsPar)
    End If
End Sub

Private Sub opt外购入库取成本价方式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt外购入库取成本价方式, Index, mrsPar)
End Sub


Private Sub pic处方颜色_Click(Index As Integer)
    Dim strColor As String
    Dim n As Integer
    
    On Error GoTo errHandle
    
    cmdialog.CancelError = True
    cmdialog.ShowColor
    pic处方颜色(Index).BackColor = cmdialog.Color
    
    '处方颜色
    For n = 0 To pic处方颜色.UBound
        strColor = IIF(strColor = "", "", strColor & ";") & CStr(pic处方颜色(n).BackColor)
    Next
    
    If Me.Visible Then
        Call SetParChange(txt, txt_处方颜色, mrsPar, True, strColor)
    End If
    
    fraSetColor.ForeColor = txt(txt_处方颜色).ForeColor
    
    Exit Sub
errHandle:
'    Resume
End Sub

Private Sub Get处方默认颜色()
    pic处方颜色(处方类型.普通).BackColor = mconlng普通
    pic处方颜色(处方类型.急诊).BackColor = mconlng急诊
    pic处方颜色(处方类型.儿科).BackColor = mconlng儿科
    pic处方颜色(处方类型.麻醉).BackColor = mconlng麻醉
    pic处方颜色(处方类型.精一).BackColor = mconlng精一
    pic处方颜色(处方类型.精二).BackColor = mconlng精二
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        objPic.Visible = (objPic.Index = Item.ID)
    Next
        
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("药房配药控制", marrFunc) Or _
                            Item.ID = GetFuncID("输液配制中心", marrFunc) Or _
                            Item.ID = GetFuncID("药品流向控制", marrFunc) Or _
                            Item.ID = GetFuncID("药品库存检查", marrFunc) Or _
                            Item.ID = GetFuncID("药品计量单位", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        
        If Item.ID = GetFuncID("输液配制中心", marrFunc) Then
            lblLocate(txt_Dept).Caption = "科室查找(&F)"
        Else
            lblLocate(txt_Dept).Caption = "药房查找(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '用于获取当前选中的TaskPanelItem
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
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub

Public Sub LocateFuncItem(ByVal lngFunc As Long)
'功能：根据ID选中一级和二级分类
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
'功能：初始化界面控件,读取并加载数据
    
    '1.初始化变量
    
    mlngPreFind = 1
    Call InitSystemPara
    
    
    
    '2.初始化界面控件
    Call InitEnv
        
    Call Load药品库房流向
    Call Load药品领用库房
    
    Call Load库房检查
    
    Call Load药品卫材精度
    Call Load单据环节控制
    
    Call LoadOther
    Call LoadVsfPrice
    Call LoadVsfPrice_给药途径
    Call Load输液自备药清单
    
    '3.加载系统参数
    Call LoadPar
    
    
End Sub

Private Sub LoadVsfPrice()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 序号,配药类型,项目id,收费项目 from 配置收费方案 where nvl(诊疗id,0) = 0 order by 序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .Rows = 2
            .TextMatrix(1, .ColIndex("优先级")) = 1
        Else
            .Rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            If NVL(rsTemp!项目id) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("优先级")) = i
                .TextMatrix(i, .ColIndex("配药类型")) = rsTemp!配药类型
                .TextMatrix(i, .ColIndex("项目id")) = rsTemp!项目id
                .TextMatrix(i, .ColIndex("收费项目")) = rsTemp!收费项目
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

Private Sub LoadVsfPrice_给药途径()
    Dim rsTemp As Recordset
    Dim rsData As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 诊疗id,项目id,收费项目 from 配置收费方案 where nvl(诊疗id,0) <> 0 order by 序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice_给药途径
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .Rows = 1
            .Rows = 2
        Else
            .Rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            '查询诊疗项目名称
            gstrSQL = "select 名称 from 诊疗项目目录 where id = [1]"
            Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "查询诊疗项目名称", rsTemp!诊疗id)
            
            If NVL(rsTemp!项目id) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("诊疗id")) = rsTemp!诊疗id
                .TextMatrix(i, .ColIndex("给药途径")) = rsData!名称
                .TextMatrix(i, .ColIndex("项目id")) = rsTemp!项目id
                .TextMatrix(i, .ColIndex("收费项目")) = rsTemp!收费项目
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
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    Dim n As Integer
    Dim obj应用范围 As CheckBox
    
    '读取参数(默认读取系统参数，需要的模块参数单独添加对应的模块号)
    Set rsTmp = GetPar(mrsPar, p药品目录管理 & "," & _
            p药品外购管理 & "," & _
            p药品自制入库 & "," & _
            p药品移库管理 & "," & _
            p药品盘点管理 & "," & _
            p药品调价管理 & "," & _
            p药品质量管理 & "," & _
            p药品处方发药 & "," & _
            p药品部门发药 & "," & _
            p大处方审查 & "," & _
            p输液配置中心 & "," & _
            p门诊处方审查 & "," & _
            p住院药嘱审查 & "," & _
            p处方审查项目 & "," & _
            p处方审查条件 & "," & _
            p处方审查统计 & "," & _
            p药品申领管理 & "," & _
            p药品领用管理)
    
    '----------------------------------------------------------
    '系统参数
    '1.设置CheckBox类参数
    strTmp = "0:6:" & chk_未审核记帐处方发药 & _
            ",0:18:" & chk_限定药品的库存 & _
            ",0:45:" & chk_收费同时发药 & _
            ",0:54:" & chk_时价加成率入库 & _
            ",0:75:" & chk_外购入库需要核查 & _
            ",0:76:" & chk_时价药品直接确定售价 & _
            ",0:126:" & chk_时价入库按折扣前采购价加成销售 & _
            ",0:148:" & chk_未收费处方发药 & _
            ",0:163:" & chk_项目执行前先收费或审核 & _
            ",0:173:" & chk_外购入库需要经过标记付款后才能进行付款 & _
            ",0:181:" & chk_时价分段加成入库 & _
            ",0:183:" & chk_时价药品取上次售价 & _
            ",0:214:" & chk_首次医嘱执行需要审核 & _
            ",0:294:" & chk_入库取目录中产地信息
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '设置参数关系
    If chk(chk_项目执行前先收费或审核).value = 1 Then
        chk(chk_未审核记帐处方发药).Enabled = False
        chk(chk_未收费处方发药).Enabled = False
        lbl未收费发药.Caption = "  已启用了门诊一卡通参数“执行前必须先收费或先记帐审核”，则对门诊病人发药或发料时，以下参数无论勾选都将失效。"
    Else
        chk(chk_未审核记帐处方发药).Enabled = True
        chk(chk_未收费处方发药).Enabled = True
        lbl未收费发药.Caption = "  如果启用了门诊一卡通参数“执行前必须先收费或先记帐审核”，则对门诊病人发药或发料时，以下参数将失效。"
    End If
        
    '2.设置ComboBox类参数
    strTmp = "0:29:" & cbo_定价单位 & _
            ",0:64:" & cbo_药品单据审核 & _
            ",0:87:" & cbo_药品编码模式 & _
            ",0:149:" & cbo_效期显示方式 & _
            ",0:150:" & cbo_药品出库优先算法
    Call SetParToControl(strTmp, mrsPar, cbo)
    
    '按val(cbo.list)取值
    strTmp = "0:275:" & cbo_零差价模式
    Call SetParToControl(strTmp, mrsPar, cbo, 2)
        
    '3.设置UpDown类参数
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
    
    
    '4.设置TextBox类参数
    strTmp = ""
'    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.设置ListBox类参数
'    strTmp = p住院医嘱下达 & ":44:" & lst_输液中心发药病人科室
'    Call SetParToControl(strTmp, mrsPar, lst, 1)
    
    '6.设置OptionButton类参数
    arrObj = Array(0, 19, opt发药窗口)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        
        Case 221   '药品结存时间模式
            If Val(strValue) = -1 Then
                '参数值为-1表示手工结存
                opt结存方式(0).value = True
                opt结存方式(1).value = False
                
                opt结存时间模式(0).Enabled = False
                opt结存时间模式(1).Enabled = False
                txt(txt_结存时间模式).Enabled = False
            Else
                '参数值不为-1表示自动结存
                opt结存方式(0).value = False
                opt结存方式(1).value = True
                
                If Val(strValue) = 0 Then
                    '参数值为0表示每月最后一天结存
                    opt结存时间模式(0).value = True
                    opt结存时间模式(1).value = False
                    txt(txt_结存时间模式).Enabled = False
                Else
                    '参数值大于0小于等于31表示指定日期结存
                    opt结存时间模式(0).value = False
                    opt结存时间模式(1).value = True
                    
                    txt(txt_结存时间模式).Enabled = True
                    
                    '结存时点只能设置为1-31
                    If Val(strValue) > 0 Or Val(strValue) <= 31 Then
                        txt(txt_结存时间模式).Text = Val(strValue)
                    Else
                        txt(txt_结存时间模式).Text = "25"
                    End If
                End If
            End If
            
            Call SetParRelation(txt, txt_结存参数值, mrsPar, rsTmp!参数号)
            Call zldatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txt, txt_结存时间模式, mrsPar)
        
        Case 292   '药品可用数量动态计算方式
            If Val(strValue) = 0 Then
                opt可用数量(0).value = True
                opt可用数量(1).value = False
                txtM.Enabled = False
                ud可用数量.Enabled = False
                txt(txt_可用数量处理) = 0
            Else
                opt可用数量(0).value = False
                opt可用数量(1).value = True
                txtM.Enabled = True
                ud可用数量.Enabled = True
                txtM.Text = Val(strValue)
                txt(txt_可用数量处理) = Val(strValue)
            End If
            
            Call SetParRelation(txt, txt_可用数量处理, mrsPar, rsTmp!参数号)
        End Select
        
        rsTmp.MoveNext
    Loop
    
    '----------------------------------------------------------
    '8.其他模块参数
    '药品目录管理 = 1023
    '设置ComboBox类参数
    strTmp = p药品目录管理 & ":西成药收入项目:" & cbo_西药收入项目 & _
            "," & p药品目录管理 & ":中成药收入项目:" & cbo_成药收入项目 & _
            "," & p药品目录管理 & ":中草药收入项目:" & cbo_中药收入项目
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '设置OptionButton类参数
    arrObj = Array(p药品目录管理, "售价按加成计算", opt售价计算, _
                    p药品目录管理, "药品分批属性自动设置", opt分批)
    Call SetParToControl("", mrsPar, arrObj)
    
    '设置UpDown类参数
    strTmp = p药品目录管理 & ":数字码:" & ud_数字码
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
    
    '其他参数
    '特殊参数处理：参数值对应多个控件(组)，先调用公共方法记录控件名称，界面控件显示单独处理
    strTmp = p药品目录管理 & ":应用范围:0"
    Call SetParToControl(strTmp, mrsPar, chk应用范围)
    
    rsTmp.Filter = "模块=" & p药品目录管理 & " And 参数名='应用范围'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值, "111111")
    If strValue <> "" Then
        For n = 1 To chk应用范围.Count - 1
            chk应用范围(n).value = Mid(strValue, n + 1, 1)
            chk应用范围(n).Tag = Mid(strValue, n + 1, 1)
        Next
    End If
    
    '----------------------------------------------------------
    '药品外购管理 = 1300
    '设置CheckBox类参数
    strTmp = p药品外购管理 & ":修改采购限价:" & chk_允许修改采购限价 & _
            "," & p药品外购管理 & ":招标药品可选择非中标单位入库:" & chk_中标单位
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '设置OptionButton类参数
    arrObj = Array(p药品外购管理, "取上次采购价方式", opt外购入库取成本价方式)
    Call SetParToControl("", mrsPar, arrObj)
    
    '特殊处理
    '该参数实际用表格和其他控件显示，特别的用额外文本控件记录原始值，单独处理界面显示
    strTmp = p药品外购管理 & ":资质校验:" & txt_供应商资质
    Call SetParToControl(strTmp, mrsPar, txt)
    
    rsTmp.Filter = "模块=" & p药品外购管理 & " And 参数名='资质校验'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    Call Load供应商资质校验(strValue)
    
    '----------------------------------------------------------
    '药品自制入库 = 1301
    '设置OptionButton类参数
    arrObj = Array(p药品自制入库, "药品自制入库成本价计算方式", opt自制入库成本来源)
    Call SetParToControl("", mrsPar, arrObj)
    
    '----------------------------------------------------------
    '药品移库管理 = 1304
    '设置CheckBox类参数
    strTmp = p药品移库管理 & ":移库流程:" & chk_移库流程 & _
            "," & p药品移库管理 & ":冲销申请:" & chk_移库冲销申请 & _
            "," & p药品移库管理 & ":移库时分批药品允许补录产地批号:" & chk_按批次出库时允许补录批号产地 & _
            "," & p药品移库管理 & ":药品按批次出库:" & chk_移库按批次出库
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    '药品申领管理 = 1343
    '设置CheckBox类参数
    strTmp = p药品申领管理 & ":药品按批次出库:" & chk_申领按批次出库
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    '药品领用管理 = 1305
    '设置CheckBox类参数
    strTmp = p药品领用管理 & ":药品按批次出库:" & chk_领用按批次出库 & _
            "," & p药品领用管理 & ":冲销申请:" & chk_领用冲销申请
    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    '----------------------------------------------------------
    '药品盘点管理 = 1307
    '设置CheckBox类参数
    strTmp = p药品盘点管理 & ":存储库房:" & chk_盘点存储库房 & _
            "," & p药品盘点管理 & ":忽略药品服务对象:" & chk_盘点忽略服务对象 & _
            "," & p药品盘点管理 & ":盘已停用的药品:" & chk_盘点已停用药品 & _
            "," & p药品盘点管理 & ":盘亏减可用数量检查:" & chk_盘点盘亏减可用数量检查
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    '药品调价管理 = 1333
    '设置CheckBox类参数
    strTmp = p药品调价管理 & ":时价药品按批次调价:" & chk_时价药品按批次调价 & _
            "," & p药品调价管理 & ":限价提示:" & chk_超过限价提示 & "," & p药品调价管理 & ":成本价按库房批次调整:" & chk_成本价按库房批次调整
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    '药品质量管理 = 1331
    '设置CheckBox类参数
    strTmp = p药品质量管理 & ":审核时减少库存:" & chk_审核时处理库存
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '----------------------------------------------------------
    '药品处方发药 = 1341
    '设置CheckBox类参数
    strTmp = p药品处方发药 & ":自动销帐:" & chk_退药自动销账 & _
            "," & p药品处方发药 & ":发药后刷卡验证:" & chk_发药刷卡 & _
            "," & p药品处方发药 & ":药品医嘱按发生时间过滤:" & chk_医嘱过滤 & _
            "," & p药品处方发药 & ":启用病人实际取药确认模式:" & chk_确认病人实际取药 & _
            "," & p药品处方发药 & ":校验配药人:" & chk_校验配药人 & _
            "," & p药品处方发药 & ":校验发药人:" & chk_校验发药人 & _
            "," & p药品处方发药 & ":配药时对未收费的单据进行收费:" & chk_配药时对未收费的单据进行收费
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '设置UpDown类参数
    strTmp = p药品处方发药 & ":查询未发药单据天数:" & ud_处方发药查询天数
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
    
    '特殊处理
    '该参数实际用表格和其他控件显示，特别的用额外文本控件记录原始值，单独处理界面显示
    strTmp = p药品处方发药 & ":处方颜色:" & txt_处方颜色
    Call SetParToControl(strTmp, mrsPar, txt)
    
    rsTmp.Filter = "模块=" & p药品处方发药 & " And 参数名='处方颜色'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    Call Get处方颜色(strValue)
    
    '----------------------------------------------------------
    '药品部门发药 = 1342
    '设置CheckBox类参数
    strTmp = p药品部门发药 & ":缺药检查:" & chk_缺药检查 & _
            "," & p药品部门发药 & ":库房货位及库存限量提示:" & chk_库存限量 & _
            "," & p药品部门发药 & ":发药时汇总退药销帐记录:" & chk_发药汇总退药销账 & _
            "," & p药品部门发药 & ":审核出院病人的销账申请:" & chk_出院病人销账 & _
            "," & p药品部门发药 & ":领药人签名:" & chk_住院药房领药人签名 & _
            "," & p药品部门发药 & ":发药时审核医嘱:" & chk_发药时审核医嘱 & _
            "," & p药品部门发药 & ":退药待发单据默认为发药状态:" & chk_退药待发单据默认为发药状态 & _
            "," & p药品部门发药 & ":退药人签名:" & chk_住院药房退药人签名 & _
            "," & p药品部门发药 & ":是否可以销帐拒绝:" & chk_是否可以销帐拒绝
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '设置UpDown类参数
    strTmp = p药品部门发药 & ":查询天数:" & ud_部门发药查询天数
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
    
    '设置TextBox类参数
    strTmp = p药品部门发药 & ":自动刷新未发药清单:" & txt_自动刷新时间
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '特殊处理
    rsTmp.Filter = "模块=" & p药品部门发药 & " And 参数名='自动刷新未发药清单'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    chk自动刷新.value = IIF(Val(strValue) > 0, 1, 0)
    If chk自动刷新.value = 0 Then txt(txt_自动刷新时间).Enabled = False
    
    '----------------------------------------------------------
    '大处方审查 = 1347
    '设置TextBox类参数
    strTmp = p大处方审查 & ":审查标准:" & txt_审查金额
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '----------------------------------------------------------
    '输液配置中心 = 1345
    '设置OptionButton类参数
    arrObj = Array(p输液配置中心, "医嘱类型", opt输液医嘱期效)
    Call SetParToControl("", mrsPar, arrObj)
    
    '设置CheckBox类参数
    strTmp = p输液配置中心 & ":批次设置:" & chk_手工调整批次 & _
            "," & p输液配置中心 & ":打包设置:" & chk_手工调整打包 & _
            "," & p输液配置中心 & ":保持上次批次:" & chk_保持上次批次 & _
            "," & p输液配置中心 & ":配液输液单配药后允许销帐申请:" & chk_配药后销账申请 & _
            "," & p输液配置中心 & ":配置中心不接收的静脉营养医嘱在病区配置:" & chk_TPN配置 & _
            "," & p输液配置中心 & ":按批次，药品排序:" & chk_输液单排序 & _
            "," & p输液配置中心 & ":特殊药品按药品类型指定批次:" & chk_特殊药品批次 & _
            "," & p输液配置中心 & ":单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包:" & chk_0批次规则 & _
            "," & p输液配置中心 & ":启动自动排批:" & chk_自动排批 & _
            "," & p输液配置中心 & ":不允许置换药房到输液配置中心:" & chk_不允许置换药房到输液配置中心 & _
            "," & p输液配置中心 & ":配置费按病人收取:" & chk_配置费收取方式 & _
            "," & p输液配置中心 & ":出院病人不收配置费:" & chk_出院病人是否收配置费 & _
            "," & p输液配置中心 & ":扫两次瓶签号自动发送:" & chk_扫描一次完成操作 & _
            "," & p输液配置中心 & ":当天发送的医嘱产生的输液单全部到备用批次:" & chk_当天发送的医嘱产生的输液单全部到备用批次 & _
            "," & p输液配置中心 & ":自动排批时输液单的批次只往后面批次变动:" & chk_自动排批只往后面批次调整 & _
            "," & p输液配置中心 & ":打包药品在发送环节收取配置费:" & chk_打包药品在发送环节收取配置费 & _
            "," & p输液配置中心 & ":自备药允许发往静配中心:" & chk_自备药 & _
            "," & p输液配置中心 & ":不取药允许发往静配中心:" & chk_不取药 & _
            "," & p输液配置中心 & ":离院带药允许发往静配中心:" & chk_离院带药 & _
            "," & p输液配置中心 & ":输液单摆药后临床不允许改变打包状态:" & chk_输液单摆药后临床不允许改变打包状态 & _
            "," & p输液配置中心 & ":打印瓶签时填写各个环节的实际操作员:" & chk_打印瓶签时填写各个环节的实际操作员
            
            
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '特殊处理
    '设置ListBox类参数
    '给药途径
    strTmp = p输液配置中心 & ":输液给药途径:" & lst_PIVA给药途径
    Call SetParToControl(strTmp, mrsPar, lst, 4)
    rsTmp.Filter = "模块=" & p输液配置中心 & " And 参数名='输液给药途径'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    If strValue <> "" Then
        chk给药途径.value = 1
    End If
    lst(lst_PIVA给药途径).Enabled = chk给药途径.Enabled And (chk给药途径.value = 1)
    lst(lst_PIVA给药途径).BackColor = IIF(lst(lst_PIVA给药途径).Enabled, &H80000005, &H8000000F)
    
    '来源科室
    strTmp = p输液配置中心 & ":来源病区:" & lst_PIVA来源科室
    Call SetParToControl(strTmp, mrsPar, lst, 4)
    rsTmp.Filter = "模块=" & p输液配置中心 & " And 参数名='来源病区'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    If strValue <> "" Then
        chk来源科室.value = 1
    End If
    lst(lst_PIVA来源科室).Enabled = chk来源科室.Enabled And (chk来源科室.value = 1)
    lst(lst_PIVA来源科室).BackColor = IIF(lst(lst_PIVA来源科室).Enabled, &H80000005, &H8000000F)
    
    '处方审查
    strTmp = "0:245:44,0:246:45,0:267:22"
    Call SetParToControl(strTmp, mrsPar, Me.chk)
    
    strTmp = "0:241:0"
    Call SetParToControl(strTmp, mrsPar, Me.chk)
    rsTmp.Filter = "模块=0 And 参数号=241"
    If rsTmp.EOF Then
        strValue = "0"
    Else
        strValue = zlCommFun.NVL(rsTmp!参数值, "0")
    End If
    Select Case Val(strValue)
    Case 1      '门诊启用，住院不启用
        chk(chk_门诊处方审查).value = 1
        chk(chk_住院药嘱审查).value = 0
        chk(chk_提醒住院医生).Enabled = False
    Case 2      '门诊不启用，住院启用
        chk(chk_门诊处方审查).value = 0
        chk(chk_住院药嘱审查).value = 1
        chk(chk_提醒门诊医生).Enabled = False
        Me.optOpporunity(1).Enabled = False
        Me.optOpporunity(2).Enabled = False
        Me.txt(txt_门诊药师审方离岗时长).Enabled = False
    Case 3      '门诊、住院都启用
        chk(chk_门诊处方审查).value = 1
        chk(chk_住院药嘱审查).value = 1
    Case Else   '门诊、住院都不启用
        chk(chk_门诊处方审查).value = 0
        chk(chk_住院药嘱审查).value = 0
        chk(chk_提醒门诊医生).Enabled = False
        chk(chk_提醒住院医生).Enabled = False
        Me.optOpporunity(1).Enabled = False
        Me.optOpporunity(2).Enabled = False
        Me.txt(txt_门诊药师审方离岗时长).Enabled = False
    End Select
    
    strTmp = "0:242:1"
    Call SetParToControl(strTmp, mrsPar, Me.optOpporunity)
    rsTmp.Filter = "模块=0 And 参数号=242"
    If rsTmp.EOF Then
        Me.optOpporunity(1).value = True
        Me.fraOpporunity.Tag = "1"
    Else
        If Val(zlCommFun.NVL(rsTmp!参数值)) = 2 Then
            Me.optOpporunity(2).value = True
            Me.fraOpporunity.Tag = "2"
        Else
            Me.optOpporunity(1).value = True
            Me.fraOpporunity.Tag = "1"
        End If
    End If
    
    chk(chk_门诊处方自动发送).Enabled = Me.optOpporunity(1).value
    
    strTmp = "0:243:5"
    Call SetParToControl(strTmp, mrsPar, Me.txt)
    
End Sub

Private Sub Get处方颜色(ByVal strParaValue As String)
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If strParaValue <> "" Then
        For n = 0 To UBound(Split(strParaValue, ";"))
            pic处方颜色(n).BackColor = Val(Split(strParaValue, ";")(n))
        Next
    Else
        Call Get处方默认颜色
    End If
    
    Exit Sub
errHandle:
    Call Get处方默认颜色
End Sub
Private Function Get供应商资质校验() As String
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean

    blnAllUnCheck = True
    
    '保存资质校验项目和方式，格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    With vsfCheck
        For i = 1 To .Rows - 1
            strCheck = IIF(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("类别")) & "," & .TextMatrix(i, .ColIndex("校验项目")) & "," & _
                IIF(.TextMatrix(i, .ColIndex("校验")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("校验")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If blnAllUnCheck = True Then
        strCheck = "0|" & strCheck
    ElseIf optCheck(0).value = True Then
        strCheck = "2|" & strCheck
    Else
        strCheck = "1|" & strCheck
    End If
        
    Get供应商资质校验 = strCheck
End Function
Private Sub Load供应商资质校验(ByVal strParaValue As String)
    Dim i As Integer
    Dim n As Integer
    Dim intCheckType As Integer
    Dim arrColumn
    
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....

    If strParaValue <> "" Then
        If InStr(1, strParaValue, "|") > 0 Then
            '校验方式：0-不检查；1－提醒；2－禁止
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
                                If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("类别")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("校验项目")) Then
                                    If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                        .TextMatrix(i, .ColIndex("校验")) = "√"
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
'功能：初始化界面控件，加载基础数据
    Dim rsData As ADODB.Recordset
    
    '1.基础参数
    cbo(cbo_药品编码模式).AddItem "顺序编号"
    cbo(cbo_药品编码模式).AddItem "种类+分类号+顺序编号"
    Call zlControl.CboSetWidth(cbo(cbo_药品编码模式).hwnd, cbo(cbo_药品编码模式).Width * 1.2)
    
    cbo(cbo_效期显示方式).AddItem "0-显示失效期"
    cbo(cbo_效期显示方式).AddItem "1-显示有效期"
    Call zlControl.CboSetWidth(cbo(cbo_效期显示方式).hwnd, cbo(cbo_效期显示方式).Width * 1.2)
    
    cbo(cbo_药品出库优先算法).AddItem "0-按批次先进先出"
    cbo(cbo_药品出库优先算法).AddItem "1-按效期最近先出"
    Call zlControl.CboSetWidth(cbo(cbo_药品出库优先算法).hwnd, cbo(cbo_药品出库优先算法).Width * 1.2)
    
    cbo(cbo_定价单位).AddItem "0-售价单位"
    cbo(cbo_定价单位).AddItem "1-药库单位"
    cbo(cbo_定价单位).ListIndex = 0
    
    cbo(cbo_药品单据审核).AddItem "0-不处理"
    cbo(cbo_药品单据审核).AddItem "1-相同禁止"
    cbo(cbo_药品单据审核).ListIndex = 0
    
    cbo(cbo_零差价模式).AddItem "0-不启用零差价管理模式"
'    cbo(cbo_零差价模式).AddItem "1-在销售环节启用零差价模式"       '暂时屏蔽第1种模式
    cbo(cbo_零差价模式).AddItem "2-在全流通业务启用零差价模式"
    cbo(cbo_零差价模式).ListIndex = 0
    cbo(cbo_零差价模式).Tag = 2     '按val(list)值读写参数
    Call zlControl.CboSetWidth(cbo(cbo_零差价模式).hwnd, cbo(cbo_零差价模式).Width * 1.2)
    
    '----------------------------------------------------------
    '2.其他参数
    '药品目录管理 = 1023
    gstrSQL = "Select ID,编码||'-'||名称 名称 From 收入项目 Where 末级=1"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "InitEnv")
    
    With rsData
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_西药收入项目).AddItem !名称
            cbo(cbo_西药收入项目).ItemData(cbo(cbo_西药收入项目).NewIndex) = !ID
            .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_成药收入项目).AddItem !名称
            cbo(cbo_成药收入项目).ItemData(cbo(cbo_成药收入项目).NewIndex) = !ID
            .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            cbo(cbo_中药收入项目).AddItem !名称
            cbo(cbo_中药收入项目).ItemData(cbo(cbo_中药收入项目).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '表示按ItemData匹配参数值
    cbo(cbo_西药收入项目).Tag = 1
    cbo(cbo_成药收入项目).Tag = 1
    cbo(cbo_中药收入项目).Tag = 1
    
    gstrSQL = "select nvl(max(length(简码)),0) 长度 from 收费项目别名 where 码类=3"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "数字码长度")
    
    If rsData!长度 = 0 Then
        ud(ud_数字码).Min = 7
    Else
        ud(ud_数字码).Min = rsData!长度
    End If
    ud(ud_数字码).Max = 40
    
    
    '药品外购管理 = 1300
    '药品自制入库 = 1301
    '药品移库管理 = 1304
    '药品盘点管理 = 1307
    '药品调价管理 = 1333
    '药品质量管理 = 1331
    '药品处方发药 = 1341
    '药品部门发药 = 1342
    '大处方审查 = 1347
    
    '----------------------------------------------------------
    '输液配置中心=1345
    ''给药途径
    gstrSQL = "Select ID, 名称 as 用法 ,标本部位 As 分类 From 诊疗项目目录 Where 类别='E' And 操作类型='2'And (服务对象=2 Or 服务对象=3) And 执行分类 = 1 " & _
            " And (撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or 撤档时间 Is Null) Order by 编码 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "提取给药途径")
    
    With lst(lst_PIVA给药途径)
        Do While Not rsData.EOF
            .AddItem rsData!用法
            .ItemData(.NewIndex) = rsData!ID
            rsData.MoveNext
        Loop
    End With
    
    ''来源科室
    gstrSQL = "Select 编码 || '-' || 名称 科室, Id " & _
            " From 部门表 " & _
            " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And Id In (Select 部门id From 部门性质说明 Where 工作性质 = '护理' And 服务对象 In (2,3)) And " & _
            " (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By 编码 || '-' || 名称 "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "Load来源科室")

    With lst(lst_PIVA来源科室)
        Do While Not rsData.EOF
            .AddItem rsData!科室
            .ItemData(.NewIndex) = rsData!ID
            rsData.MoveNext
        Loop
    End With
   
    
    '----------------------------------------------------------
    '3.其他设置
    With Bill(bill_药品库房流向)
        .Cols = 4 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "所在库房"
        .TextMatrix(0, 1) = "对方库房"
        .TextMatrix(0, 2) = "对方库房ID"
        .TextMatrix(0, 3) = "流向"
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
    
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
    
    With Bill(bill_药品领用流向)
        .Cols = 3 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "领用部门"
        .TextMatrix(0, 1) = "领用库房"
        .TextMatrix(0, 2) = "库房ID"
        .ColWidth(0) = 1900
        .ColWidth(1) = 1900
        .ColWidth(2) = 0
        .ColData(0) = 1
        .ColData(1) = 3
        .ColData(2) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    

    '库房单位
    With msf库房计量单位
        .AllowUserResizing = flexResizeNone
        .FixedRows = 1
        .Cols = 5
        .MergeCol(0) = True
        .FormatString = "药品库房|服务对象|售价单位|门诊单位|住院单位|药库单位"
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
    
    
    With Bill药房配药控制
        
        .Cols = 5 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .TextMatrix(0, 0) = "药房"
        .TextMatrix(0, 1) = "服务对象"
        .TextMatrix(0, 2) = "配药"
        .TextMatrix(0, 3) = "自动发药天数"
        .TextMatrix(0, 4) = "配药确认"
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
        mrsPar.Filter = "(修改状态=1 ANd ErrType =Null) OR  (修改状态=1 And ErrType=" & PET_值超限 & ")"
        If mrsPar.RecordCount > 0 Or Bill(bill_药品库房流向).Tag = "已修改" Or Bill(bill_药品领用流向).Tag = "已修改" _
            Or lvw库存检查.Tag = "已修改" Or msf库房计量单位.Tag = "已修改" Or Bill药房配药控制.Tag = "已修改" _
            Or Bill药品卫材精度.Tag = "已修改" Or vsf单据环节控制.Tag = "已修改" Then
            
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mrsPar = Nothing
    
End Sub

Private Sub cmdOK_Click()
    Dim obj应用范围 As CheckBox
    Dim strValue As String
    
    If ValidateData() = False Then Exit Sub
    
    Call Save药品库房流向
    Call Save药品领用流向
    
    
    Call Save库房检查
    Call Save库房单位
    Call Save药房配药控制
    
    Call Save药品卫材精度
    Call Save单据环节控制
    
    Call Save配置收费方案
    Call Save输液自备药清单
    
    '药品目录管理
    '特殊参数处理
    For Each obj应用范围 In chk应用范围
        strValue = IIF(strValue = "", "", strValue) & obj应用范围.value
    Next
    Call SetParChange(chk应用范围, 0, mrsPar, True, strValue)
    
    '药品处方发药
    '特殊参数处理
    If chk(chk_发药汇总退药销账).value = 1 Then zldatabase.SetPara "按科室汇总显示汇总清单", 1, glngSys, 1342
    
    updateChange mrsPar
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub updateChange(ByRef rsPar As ADODB.Recordset)
'检查由升级带来的参数改变
'对应参数未修改但是存在界面上的值与数据库的值是否不一致，不一致则以界面为准需要重新保存

    '确定售价的方式(chk_时价加成率入库、chk_时价分段加成入库、chk_时价药品取上次售价)
    rsPar.Filter = "(控件名称 = 'chk' And 控件数组序号 = " & chk_时价分段加成入库 & ") or (控件名称 = 'chk' And 控件数组序号 = " & chk_时价加成率入库 & ") or (控件名称 = 'chk' And 控件数组序号 = " & chk_时价药品取上次售价 & ")"
    
    With rsPar
        Do While Not .EOF
        
            Select Case rsPar!控件数组序号
            Case chk_时价分段加成入库
                If ("" & chk(chk_时价分段加成入库).value <> "" & rsPar!参数值) And NVL(rsPar!修改状态, 0) <> 1 Then
                    rsPar!参数新值 = chk(chk_时价分段加成入库).value
                    rsPar!修改状态 = 1
                    .Update
                    Call MsgBox("提醒：参数【时价药品通过分段加成入库】未经过修改，与数据库不一致，将以界面为准保存！")
                End If
            Case chk_时价加成率入库
                If ("" & chk(chk_时价加成率入库).value <> "" & rsPar!参数值) And NVL(rsPar!修改状态, 0) <> 1 Then
                    rsPar!参数新值 = chk(chk_时价加成率入库).value
                    rsPar!修改状态 = 1
                    .Update
                    Call MsgBox("提醒：参数【时价药品通过加成率入库】未经过修改，与数据库不一致，将以界面为准保存！")
                End If
            Case chk_时价药品取上次售价
                If ("" & chk(chk_时价药品取上次售价).value <> "" & rsPar!参数值) And NVL(rsPar!修改状态, 0) <> 1 Then
                    rsPar!参数新值 = chk(chk_时价药品取上次售价).value
                    rsPar!修改状态 = 1
                    .Update
                    Call MsgBox("提醒：参数【时价药品入库时取上次售价】未经过修改，与数据库不一致，将以界面为准保存！")
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


Private Sub opt结存时间模式_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt输液医嘱期效_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt输液医嘱期效, Index, mrsPar)
    End If
End Sub

Private Sub opt输液医嘱期效_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt输液医嘱期效_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt输液医嘱期效, Index, mrsPar)
End Sub

Private Sub opt发药窗口_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt发药窗口, Index, mrsPar)
End Sub

Private Sub opt结存时间模式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(opt结存时间模式, Index, mrsPar)
End Sub

Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case chk_住院药嘱审查
        Call SetParTip(chk, chk_门诊处方审查, mrsPar, , chk(Index))
    Case Else
        Call SetParTip(chk, Index, mrsPar)
    End Select
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    Call SetParTip(cbo, Index, mrsPar)
End Sub

Private Sub Load单据环节控制()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo errHandle
    intAllItems = UBound(Split(cst所有项目, ",")) + 1
    
    With vsf单据环节控制
        .Rows = 4
        .Cols = 2 + intAllItems
        .FixedRows = 1
        .FixedCols = 2
        .RowHeightMin = 400
        
        .TextMatrix(0, 0) = "单据"
        .TextMatrix(0, 1) = "环节"
                        
        .ColWidth(0) = 820
        .ColWidth(1) = 820
                        
        For n = 0 To UBound(Split(cst所有项目, ","))
            .TextMatrix(0, n + 2) = Split(cst所有项目, ",")(n)
            .ColWidth(n + 2) = 820
            .ColAlignment(n + 2) = flexAlignCenterCenter
        Next
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        
'        .CellBorderRange 0, 0, 0, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(1, 0) = "药品外购"
        .TextMatrix(2, 0) = "药品外购"
        .TextMatrix(3, 0) = "药品外购"

        .TextMatrix(1, 1) = "核查"
        .TextMatrix(2, 1) = "审核"
        .TextMatrix(3, 1) = "财务审核"
        
'        .CellBorderRange 3, 0, 3, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
'
'        .TextMatrix(4, 0) = "卫材外购"
'        .TextMatrix(5, 0) = "卫材外购"
'        .TextMatrix(6, 0) = "卫材外购"
'
'        .TextMatrix(4, 1) = "核查"
'        .TextMatrix(5, 1) = "审核"
'        .TextMatrix(6, 1) = "财务审核"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select 单据,环节,内容 From 单据环节控制 where 单据=1 Order By 单据, 环节"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "单据环节控制")
        
        If Not rsTmp.EOF Then
            For n = 1 To rsTmp.RecordCount
                For m = 2 To intAllItems + 1
                    If InStr(1, "," & rsTmp!内容 & ",", Trim(.TextMatrix(0, m))) > 0 Then
                        Select Case rsTmp!单据
                            Case 单据.药品外购
                                Select Case rsTmp!环节
                                    Case 环节.核查
                                        .TextMatrix(1, m) = "√"
                                    Case 环节.审核
                                        .TextMatrix(2, m) = "√"
                                    Case 环节.财务审核
                                        .TextMatrix(3, m) = "√"
                                End Select
'                            Case 单据.卫材外购
'                                Select Case rsTmp!环节
'                                    Case 环节.核查
'                                        .TextMatrix(4, m) = "√"
'                                    Case 环节.审核
'                                        .TextMatrix(5, m) = "√"
'                                    Case 环节.财务审核
'                                        .TextMatrix(6, m) = "√"
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

Private Sub Save单据环节控制()
    Dim n As Integer
    Dim m As Integer
    Dim strInput As String
    Dim int单据 As Integer
    Dim int环节 As Integer
    Dim str内容 As String
    
    On Error GoTo errHandle
    With vsf单据环节控制
        If .Tag = "已修改" Then
            For n = 1 To .Rows - 1
                Select Case .TextMatrix(n, 0)
                    Case "药品外购"
                        int单据 = 单据.药品外购
'                    Case "卫材外购"
'                        int单据 = 单据.卫材外购
                End Select
                
                Select Case .TextMatrix(n, 1)
                    Case "核查"
                        int环节 = 环节.核查
                    Case "审核"
                        int环节 = 环节.审核
                    Case "财务审核"
                        int环节 = 环节.财务审核
                End Select
                
                str内容 = ""
                For m = 2 To .Cols - 1
                    If .TextMatrix(n, m) = "√" Then
                        str内容 = str内容 & IIF(str内容 <> "", ",", "") & .TextMatrix(0, m)
                    End If
                Next
                
                If str内容 <> "" Then
                    strInput = strInput & IIF(strInput <> "", ";", "") & int单据 & "," & int环节 & "," & str内容
                End If
            Next
        
            gstrSQL = "Zl_单据环节控制_Update('" & strInput & "'," & 单据.药品外购 & ")"
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
    If Index = bill_药品库房流向 Then
        If Val(lblLocate(txt_Dept).Tag) <> bill_药品库房流向 Then
            lblLocate(txt_Dept).Tag = bill_药品库房流向
            mlngPreFind = 1
        End If
    ElseIf Index = bill_药品领用流向 Then
        If Val(lblLocate(txt_Dept).Tag) <> bill_药品领用流向 Then
            lblLocate(txt_Dept).Tag = bill_药品领用流向
            mlngPreFind = 1
        End If
    End If
End Sub


Private Sub txt_Change(Index As Integer)
    Select Case Index
    Case txt_结存时间模式
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
    Case txt_门诊药师审方离岗时长
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
        Case txt_结存时间模式
            Select Case KeyAscii
            Case vbKeyBack, vbKeyEscape, 3, 22  '小数点
                KeyAscii = 0
            Case Else
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
            End Select
        Case txt_自动刷新时间, txt_审查金额
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
            KeyAscii = 0
        Case txt_门诊药师审方离岗时长
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case txt_门诊药师审方离岗时长
        If Val(txt(Index).Text) < 5 Or Val(txt(Index).Text) > 99 Then
            MsgBox "门诊药师审方离岗时长范围（5-99）！", vbInformation, gstrSysName
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
            If Bill药房配药控制.Visible Then
                Call LocateDept(strFind, Bill药房配药控制, 0)
                                
            ElseIf Bill(bill_药品领用流向).Visible Then
                If lblLocate(txt_Dept).Tag = bill_药品库房流向 Or lblLocate(txt_Dept).Tag = "" Then
                    Call LocateDept(strFind, Bill(bill_药品库房流向), IIF(Bill(bill_药品库房流向).Col = 0, 0, 1))
                Else
                    Call LocateDept(strFind, Bill(bill_药品领用流向), Bill(bill_药品领用流向).Col)
                End If
                
            ElseIf lvw库存检查.Visible Then
                Call LocateDept(strFind, lvw库存检查, 1)
                
            ElseIf msf库房计量单位.Visible Then
                Call LocateDept(strFind, msf库房计量单位, 0)
                
            ElseIf lst(lst_PIVA来源科室).Visible Then
                Call LocateDept(strFind, lst(lst_PIVA来源科室), 0)
            End If
        End Select
    End If
End Sub


Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object, ByVal lngCol As Long)
'功能：查找科室
'参数：lngCol-进行查找的列
    Dim i As Long, lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    With objTmp
        If TypeName(objTmp) = "ListView" Then 'lvw库存检查
            lngRows = .ListItems.Count
            For i = mlngPreFind To lngRows
                If .ListItems(i).ListSubItems(lngCol).Text Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst_输液中心发药病人科室
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
            MsgBox "没有找到匹配的，请检查输入的内容。", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        Else
            MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
            mlngPreFind = 1
        End If
    End If
End Sub








Private Sub txtUD_Change(Index As Integer)
    If Index <> 2 Then Exit Sub
    If Val(txtud.Item(2).Text) > 30 Then
        MsgBox "查询未发药单据天数最大不超过30天!"
    Else
        ud(2).value = Val(txtud.Item(2).Text)
    End If
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Then Exit Sub     '查询未发药单据天数
    '只允许输入数字
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




Private Sub ud可用数量_Change()
     Call SetParChange(txt, txt_可用数量处理, mrsPar, True, Val(txtM.Text))
End Sub


Private Sub vsfCheck_DblClick()
    With vsfCheck
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("校验") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            .TextMatrix(.Row, .Col) = "√"
        End If
    End With
    
    If Me.Visible Then
        Call SetParChange(txt, txt_供应商资质, mrsPar, True, Get供应商资质校验)
    End If
    
    fra供应商资质.ForeColor = txt(txt_供应商资质).ForeColor
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
    
    If VSFPrice.ColSel <> VSFPrice.ColIndex("优先级") Then
        VSFPrice.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub vsf单据环节控制_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsf单据环节控制.Tag = "" Then vsf单据环节控制.Tag = "已修改"
End Sub

Private Sub vsf单据环节控制_DblClick()
    With vsf单据环节控制
        If .Row < 1 Then Exit Sub
        If .Col < 2 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            '核查时不能修改"发票号,发票代码,发票日期,发票金额"
            If .TextMatrix(.Row, 1) = "核查" And InStr(1, "发票号,发票代码,发票日期,发票金额", .TextMatrix(0, .Col)) > 0 Then Exit Sub
            
'            '卫材外购无外观选项
'            If .TextMatrix(.Row, 0) = "卫材外购" And .TextMatrix(0, .Col) = "外观" Then Exit Sub
            
            .TextMatrix(.Row, .Col) = "√"

        End If
        
    End With
End Sub

Private Sub VSFPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_给药途径_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_给药途径_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice_给药途径.Height
        .Top = frmMoney.Top + tabPrice.Top + VSFPrice_给药途径.Top
        .Left = frmMoney.Left + tabPrice.Left + VSFPrice_给药途径.Left
        .Width = VSFPrice_给药途径.Width
        .Tag = 1
    End With
    
    If Col = VSFPrice_给药途径.ColIndex("给药途径") Then
        If mRsWay Is Nothing Then
            Set mRsWay = DeptSendWork_给药途径
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsWay.RecordCount > 0 Then mRsWay.MoveFirst
            Do While Not mRsWay.EOF
                .ListItems.Add , "_" & mRsWay!ID, mRsWay!名称
                mRsWay.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice_给药途径.ColIndex("收费项目") Then
        If mRsPrice Is Nothing Then
            Set mRsPrice = DeptSendWork_Get收费项目
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!ID, mRsPrice!名称
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
    
    
    If Col = VSFPrice.ColIndex("配药类型") Then
        If mRsType Is Nothing Then
            Set mRsType = DeptSendWork_Get配药类型
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsType.RecordCount > 0 Then mRsType.MoveFirst
            Do While Not mRsType.EOF
                .ListItems.Add , "_" & mRsType!编码, mRsType!名称
                mRsType.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice.ColIndex("收费项目") Then
        If mRsPrice Is Nothing Then
            Set mRsPrice = DeptSendWork_Get收费项目
        End If
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!ID, mRsPrice!名称
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Function DeptSendWork_给药途径() As Recordset
'获取给药途径,目前只针对“静脉营养”类
    On Error GoTo ErrHand
    gstrSQL = "select ID, 名称 from 诊疗项目目录 where 类别 = 'E' and 操作类型 = '2' and 执行分类 = '1' and 执行标记 = 2"
    
    Set DeptSendWork_给药途径 = zldatabase.OpenSQLRecord(gstrSQL, "获取给药途径")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeptSendWork_Get配药类型() As Recordset
'获取药品的配药类型
    On Error GoTo ErrHand
    gstrSQL = "select 编码,名称 from 输液配药类型"
    
    Set DeptSendWork_Get配药类型 = zldatabase.OpenSQLRecord(gstrSQL, "获取配药类型")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeptSendWork_Get收费项目() As Recordset
'获取收费项目
    On Error GoTo ErrHand
    gstrSQL = "select id,编码,名称,计算单位,说明 from 收费项目目录 where 类别='Z' and nvl(是否变价,0)=0"
    
    Set DeptSendWork_Get收费项目 = zldatabase.OpenSQLRecord(gstrSQL, "获取收费项目")
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
            If .TextMatrix(.Row, .ColIndex("配药类型")) <> "" And .TextMatrix(.Row, .ColIndex("收费项目")) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = .ColIndex("配药类型")
                .TextMatrix(.Row, .ColIndex("优先级")) = .Row
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
        
        '调整序号
        For i = intRow To Me.VSFPrice.Rows - 1
            Me.VSFPrice.TextMatrix(i, Me.VSFPrice.ColIndex("优先级")) = i
        Next
    End If
    
End Sub

Private Sub VSFPrice_给药途径_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice_给药途径.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice_给药途径.Row = VSFPrice_给药途径.Rows - 1 Then
        Me.VSFPrice_给药途径.Editable = flexEDNone
        With Me.VSFPrice_给药途径
            If .TextMatrix(.Row, .ColIndex("给药途径")) <> "" And .TextMatrix(.Row, .ColIndex("收费项目")) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = .ColIndex("给药途径")
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice_给药途径.Row
        If VSFPrice_给药途径.Rows = 2 Then
           VSFPrice_给药途径.Rows = 1
           VSFPrice_给药途径.Rows = 2
        Else
            Me.VSFPrice_给药途径.RemoveItem VSFPrice_给药途径.Row
        End If
    End If
    Me.VSFPrice_给药途径.Editable = flexEDKbd
    
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

Private Sub Save输液自备药清单()
    '功能：保存输液自备药清单
    Dim strSql As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    With Me.vsf自备药清单
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, .ColIndex("药品id")) <> "") Or i = 1 Then
                gstrSQL = "Zl_输液自备药清单_设置("
                '序号
                gstrSQL = gstrSQL & i
                '药品id
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(i, .ColIndex("药品id")))
                '是否检查库存
                gstrSQL = gstrSQL & "," & IIF(.TextMatrix(i, .ColIndex("检查库存")) = "", 0, 1)
                '是否第一次重置
                gstrSQL = gstrSQL & "," & i & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "保存输液自备药清单")
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

Private Sub Save配置收费方案()
    Dim i As Integer
    Dim n As Integer
    
    With Me.VSFPrice
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, .ColIndex("优先级")) <> "" And .TextMatrix(i, .ColIndex("收费项目")) <> "" And .TextMatrix(i, .ColIndex("项目id")) <> "" And .TextMatrix(i, .ColIndex("配药类型")) <> "") Or i = 1 Then
                gstrSQL = "Zl_配置收费方案_设置("
                '序号
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("优先级"))) & ","
                '配药类型
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("配药类型")) & "',"
                '项目id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("项目id"))) & ","
                '收费项目
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("收费项目")) & "',"
                '诊疗id
                gstrSQL = gstrSQL & "NULL" & ","
                '是否第一次重置
                gstrSQL = gstrSQL & i & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "保存不接受药品")
            End If
        Next
    End With
    
    n = i - 1
    
    With Me.VSFPrice_给药途径
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("收费项目")) <> "" And .TextMatrix(i, .ColIndex("项目id")) <> "" And .TextMatrix(i, .ColIndex("给药途径")) <> "" And .TextMatrix(i, .ColIndex("诊疗id")) <> "" Then
                gstrSQL = "Zl_配置收费方案_设置("
                '序号
                gstrSQL = gstrSQL & i + n & ","
                '配药类型
                gstrSQL = gstrSQL & "NULL" & ","
                '项目id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("项目id"))) & ","
                '收费项目
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("收费项目")) & "',"
                '诊疗id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("诊疗id"))) & ","
                '是否第一次重置
                gstrSQL = gstrSQL & i + n & ")"
                
                Call zldatabase.ExecuteProcedure(gstrSQL, "保存不接受药品")
            End If
        Next
    End With
    
End Sub
Private Sub Save药品库房流向()
    Dim strTmp As String
    Dim lngRow As Long
    Dim str流向 As String
    
    On Error GoTo errHandle
    With Bill(bill_药品库房流向)
        If .Tag = "已修改" Then
            For lngRow = 1 To .Rows - 1
                If .RowData(lngRow) > 0 Then
                    str流向 = Left(.TextMatrix(lngRow, 3), 1)
                    If str流向 = "" Then str流向 = "3"
                    strTmp = strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str流向 & ","
                End If
            Next
        
            gstrSQL = "zl_药品流向控制_Modify('" & strTmp & "')"
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
    Dim str药品售价精度 As String, str药品成本价精度 As String
    
    '检查药品流向设置
    With Bill(bill_药品库房流向)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
          
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
              
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                 
                    Exit Function
                End If
            Next
        Next
    End With
    
    '检查药品领用流向设置
    With Bill(bill_药品领用流向)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
        
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
             
                    Exit Function
                End If
            Next
        Next
    End With
    
    
    If CheckParChanged(chk, chk_时价入库按折扣前采购价加成销售, mrsPar) Then
        If Check是否有未审核的外购入库单 Then
            MsgBox "还有未审核的外购入库单，不能改变参数“时价药品入库按扣前加成销售”!", vbInformation, gstrSysName
            chk(chk_时价入库按折扣前采购价加成销售).value = GetParOriginalValue(chk, chk_时价入库按折扣前采购价加成销售, mrsPar)
        
            Exit Function
        End If
    End If
    
    '零差价管理：药品卫材精度检查
    If cbo(cbo_零差价模式).ListIndex > 0 Then
        With Bill药品卫材精度
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, dig_精度类别) = "药品" And .TextMatrix(lngRow, dig_精度内容) = "零售价" Then
                    str药品售价精度 = IIF(str药品售价精度 = "", "", str药品售价精度) & .TextMatrix(lngRow, dig_精度)
                End If
            Next
            
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, dig_精度类别) = "药品" And .TextMatrix(lngRow, dig_精度内容) = "成本价" Then
                    str药品成本价精度 = IIF(str药品成本价精度 = "", "", str药品成本价精度) & .TextMatrix(lngRow, dig_精度)
                End If
            Next
            
            If str药品售价精度 <> str药品成本价精度 Then
                MsgBox "已启用药品零差价管理，药品售价和成本价各级单位的精度应保持一致!" & vbCrLf & "请在药品录入精度页面中设置。", vbInformation, gstrSysName
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
        If Index = bill_药品领用流向 Then
            If KeyCode <> vbKeyReturn Then Exit Sub
            
            If .Col = 0 Then
                If .Text = "" Then
                        '到下一个控件
                        zlCommFun.PressKey vbKeyTab
                    
                Else
                    strTmp = Replace(.Text, "'", "''")
                    gstrSQL = "Select a.id,a.编码,a.名称 From 部门表 a , 部门性质说明 b " & _
                              " Where a.id = b.部门id " & _
                              " And b.工作性质 In ('领药部门') and (a.编码 Like [1] or a.名称 like [1] or a.简码 like [1])"
                    
                    lmX = picPar(tplFunc.Tag).Left + Me.Bill(bill_药品领用流向).Left
                    lmY = picPar(tplFunc.Tag).Top + Me.Bill(bill_药品领用流向).Top + Me.Bill(bill_药品领用流向).RowHeight(.Row) + 350
                    Set rsTmp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "领药部门", False, "", "", False, False, True, lmX, lmY, 300, blnCancel, False, True, UCase(strTmp) & "%")
                    
                    If rsTmp Is Nothing Then Cancel = True: Exit Sub
                    If rsTmp.State <> 1 Then Cancel = True: Exit Sub
                    If rsTmp.EOF = True Then Cancel = True: Exit Sub
        
                    With Bill(bill_药品领用流向)
                        .TextMatrix(.Row, 0) = rsTmp("编码") & "-" & rsTmp("名称")
                        .Text = rsTmp("编码") & "-" & rsTmp("名称")
                        .RowData(.Row) = rsTmp("ID")
                    End With
                    
                End If
                .Tag = "已修改"
            End If
        End If
    End With

End Sub


Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        
        If Index = bill_药品库房流向 Then
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                            Case "2"
                                .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                        End Select
                        
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                        
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                        
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                        
                End Select
                .Tag = "已修改"
            End If
        End If
    End With

End Sub

Private Sub bill_CommandClick(Index As Integer)
'通过按钮选择细目
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_药品领用流向 Then
        gstrSQL = "Select Distinct Id,编码,名称,简码 From 部门表 a,部门性质说明 b " & _
                  "Where a.id = b.部门id And b.工作性质 In('领药部门') " & _
                  "    and (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) " & _
                  "order by 编码 "
        Set rsTmp = zldatabase.ShowSelect(Me, gstrSQL, 0, "领药部门")
        
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> 1 Then Exit Sub
        If rsTmp.EOF = True Then Exit Sub
        
        With Bill(bill_药品领用流向)
            .TextMatrix(.Row, 0) = rsTmp("编码") & "-" & rsTmp("名称")
            .RowData(.Row) = rsTmp("ID")
            .Tag = "已修改"
        End With
    End If
    
End Sub


Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If ListIndex < 0 Then Exit Sub
    
    With Bill(Index)
        If Index = bill_药品库房流向 Then
            If .Col = 0 Then
                .RowData(.Row) = .ItemData(ListIndex)
            ElseIf .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            End If
            
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
        
        ElseIf Index = bill_药品领用流向 Then
        
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            .TextMatrix(.Row, .Col) = .CboText
        End If
        .Tag = "已修改"
    End With
    
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        
        If KeyCode = vbKeyReturn Then
            If Index = bill_药品库房流向 Then
                If .Col = 1 Then
                    .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
                Else
                     .RowData(.Row) = .ItemData(.ListIndex)
                End If
                
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
                
            ElseIf Index = bill_药品领用流向 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            End If
            .Tag = "已修改"
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'处理最后一列的变化
    With Bill(Index)
        If .MouseRow = 0 Then Exit Sub
        
        If Index = bill_药品库房流向 Then
            If .MouseCol <> .Cols - 1 Then Exit Sub
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
            End Select
            
            .Tag = "已修改"
        End If
    End With
End Sub



Private Sub Bill药房配药控制_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill药房配药控制_DblClick(Cancel As Boolean)
    Dim i As Long
    With Me.Bill药房配药控制
        If (.Col = 2 Or .Col = 4) And .Row > 0 And Trim(.TextMatrix(.Row, 0)) <> "" Then
            If .TextMatrix(.Row, .Col) = "" And (.Col = 2 Or (.Col = 4 And .TextMatrix(.Row, 1) = "门诊")) Then
                .TextMatrix(.Row, .Col) = "√"
                If .Col = 4 Then
                    .TextMatrix(.Row, 2) = "√"
                End If
            Else
                If .Col = 2 And .TextMatrix(.Row, 4) = "√" Then Exit Sub
                .TextMatrix(.Row, .Col) = ""
            End If
            .Tag = "已修改"
        End If
    End With
End Sub

Private Sub Bill药房配药控制_EnterCell(Row As Long, Col As Long)
    With Bill药房配药控制
        If Col = 3 Then
            If .TextMatrix(Row, 1) = "住院" Then
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

Private Sub Bill药房配药控制_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Bill药房配药控制
        If .Col = 3 Then
            strKey = Val(.Text)
            If strKey > 30 Then
                MsgBox "自动发药天数不能大于30！", vbInformation, gstrSysName
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .TextMatrix(.Row, .Col) = IIF(.Text <> "", strKey, "")
            
            .Tag = "已修改"
        End If
    End With
End Sub

Private Sub Bill药房配药控制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        With Bill药房配药控制
            If .Col = 2 Then
                Call Bill药房配药控制_DblClick(False)
            End If
        End With
    End If
End Sub

Private Sub Load药品卫材精度()
    Const intMinDigit As Integer = 2
    Dim intMaxCost As Integer
    Dim intMaxPrice As Integer
    Dim intMaxNumber As Integer
    Dim intMaxMoney As Integer
    Dim rs As ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    '取最大精度
    gstrSQL = "Select 成本价, 零售价, 实际数量,零售金额 From 药品收发记录 Where Rownum <2"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
    intMaxCost = IIF(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIF(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIF(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIF(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

    With Bill药品卫材精度
        .Cols = dig_Cols
        .TextMatrix(0, dig_类别) = ""
        .TextMatrix(0, dig_内容) = ""
        .TextMatrix(0, dig_单位) = ""
        .TextMatrix(0, dig_精度类别) = "类别"
        .TextMatrix(0, dig_精度内容) = "内容"
        .TextMatrix(0, dig_精度单位) = "单位"
        .TextMatrix(0, dig_精度) = "目前精度"
        .TextMatrix(0, dig_最小精度) = "最小精度"
        .TextMatrix(0, dig_最大精度) = "最大精度"
        .TextMatrix(0, dig_原始精度) = ""
        
        .ColWidth(dig_类别) = 0
        .ColWidth(dig_内容) = 0
        .ColWidth(dig_单位) = 0
        .ColWidth(dig_精度类别) = 700
        .ColWidth(dig_精度内容) = 850
        .ColWidth(dig_精度单位) = 1000
        .ColWidth(dig_精度) = 850
        .ColWidth(dig_最小精度) = 850
        .ColWidth(dig_最大精度) = 850
        .ColWidth(dig_原始精度) = 0
        
        .ColData(dig_类别) = 0
        .ColData(dig_内容) = 0
        .ColData(dig_单位) = 0
        .ColData(dig_精度类别) = 0
        .ColData(dig_精度内容) = 0
        .ColData(dig_精度单位) = 0
        .ColData(dig_精度) = 4
        .ColData(dig_最小精度) = 0
        .ColData(dig_最大精度) = 0
        .ColData(dig_原始精度) = 0
        
        .PrimaryCol = 0
        .MsfObj.MergeCells = flexMergeFree
        .MergeCol dig_精度类别, True
        .MergeCol dig_精度内容, True
        .Active = True
    End With
    
    '取目前精度
    gstrSQL = " Select 性质, 类别, 内容, 单位, Decode(类别, 1, '药品', '卫材') 精度类别, Decode(内容, 1, '成本价', 2, '零售价',3, '数量','金额') 精度内容," & _
            " Decode(类别, 1, Decode(单位, 1, '售价单位', 2, '门诊单位', 3, '住院单位',4, '药库单位','所有单位')," & _
            " Decode(单位, 1, '散装',2, '包装','所有单位')) 精度单位, Nvl(精度, 0) 精度 " & _
            " From 药品卫材精度 where 类别=1 Order By 性质, 类别, 内容, 单位"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
    With Bill药品卫材精度
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            For n = 1 To rs.RecordCount
                .TextMatrix(n, dig_类别) = rs!类别
                .TextMatrix(n, dig_内容) = rs!内容
                .TextMatrix(n, dig_单位) = rs!单位
                .TextMatrix(n, dig_精度类别) = rs!精度类别
                .TextMatrix(n, dig_精度内容) = rs!精度内容
                .TextMatrix(n, dig_精度单位) = rs!精度单位
                .TextMatrix(n, dig_精度) = IIF(rs!精度 > 4, 4, rs!精度)
                .TextMatrix(n, dig_最小精度) = intMinDigit
                Select Case rs!内容
                    Case 1
                        .TextMatrix(n, dig_最大精度) = intMaxCost
                    Case 2
                        .TextMatrix(n, dig_最大精度) = intMaxPrice
                    Case 3
                        .TextMatrix(n, dig_最大精度) = intMaxNumber
                    Case 4
                        .TextMatrix(n, dig_最大精度) = intMaxMoney
                End Select
                .TextMatrix(n, dig_原始精度) = rs!精度
                .RowData(n) = rs!精度
                rs.MoveNext
            Next
        End If
    End With
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save药房配药控制()
    Dim i As Integer, blnTrans As Boolean
    
    On Error GoTo errHandle
    
    With Me.Bill药房配药控制
        If .Tag = "已修改" Then
            gcnOracle.BeginTrans: blnTrans = True
            gstrSQL = "ZL_药房配药控制_DELETE"
            zldatabase.ExecuteProcedure gstrSQL, Me.Caption
        
            For i = 1 To .Rows - 1
                If .RowData(i) > 0 Then
                    gstrSQL = "ZL_药房配药控制_INSERT(" & .RowData(i) & "," & IIF(.TextMatrix(i, 1) = "门诊", 1, 2) & "," & IIF(.TextMatrix(i, 2) <> "", 1, 0) & "," & IIF(Val(.TextMatrix(i, 3)) = 0, "Null", Val(.TextMatrix(i, 3))) & "," & IIF(.TextMatrix(i, 4) <> "", 1, 0) & ")"
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

Private Sub Save药品卫材精度()
    Dim n As Integer
    Dim strInput As String
       
    On Error GoTo errHandle
    With Bill药品卫材精度
        If .Tag = "已修改" Then
            For n = 1 To .Rows - 1
                strInput = strInput & "0," & _
                    .TextMatrix(n, dig_类别) & "," & _
                    .TextMatrix(n, dig_内容) & "," & _
                    .TextMatrix(n, dig_单位) & "," & _
                    .TextMatrix(n, dig_精度) & ";"
            Next
        
            gstrSQL = "ZL_药品卫材精度_Update('" & strInput & "')"
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
'功能：获取药房记录集
'参数：0-药房和药库,1-不含药库
    Dim strSql As String
 
    strSql = "Select b.Id, Nvl(b.编码, '') 编码, Nvl(b.名称, '') 名称, a.服务对象, a.工作性质" & vbNewLine & _
            "From 部门性质说明 A, 部门表 B" & vbNewLine & _
            "Where b.Id = a.部门id And a.工作性质 In (" & _
                IIF(bytMode = 0, "'中药库', '西药库', '成药库',", "") & " '制剂室', '中药房', '西药房', '成药房') And " & Where撤档时间("B") & vbNewLine & _
            "Order By 编码"

    On Error GoTo errH
    Set GetRSDrugStore = zldatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOther()
'完成其余的初始化工作
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng单位 As Long
    Dim strTmp As String, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '输出库房单位
    strCoding = ""
    Set rsTemp = GetRSDrugStore(0)
    msf库房计量单位.Rows = 1
    Do Until rsTemp.EOF
        With msf库房计量单位
            If rsTemp("编码") <> strCoding Then
                strTmp = ""
            End If
            If InStr(",中药库,西药库,成药库,", "," & rsTemp("工作性质") & ",") Then
                If InStr(1, strTmp & ",", ",药库,") <= 0 Then
                    .Rows = .Rows + 1
                    .RowData(.Rows - 1) = rsTemp("ID")
                    .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                    .TextMatrix(.Rows - 1, 1) = "药库"
                    strTmp = strTmp & "," & "药库"
                End If
            End If
            
            If InStr(",制剂室,中药房,西药房,成药房,", "," & rsTemp("工作性质") & ",") Then
            
                Select Case rsTemp("服务对象")
                    Case 0          '不服务于病人

                    Case 1          '服务于门诊病人
                        If InStr(1, strTmp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTmp = strTmp & "," & "门诊"
                        End If
                    Case 2          '服务于住院病人
                        If InStr(1, strTmp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTmp = strTmp & "," & "住院"
                        End If
                    Case 3          '服务于门诊住院病人
                        If InStr(1, strTmp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTmp = strTmp & "," & "门诊"
                        End If
                        
                        If InStr(1, strTmp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTmp = strTmp & "," & "住院"
                        End If
                End Select
            End If
            If InStr(1, strTmp & ",", ",其他,") <= 0 Then
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = rsTemp("ID")
                .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                .TextMatrix(.Rows - 1, 1) = "其他"
                strTmp = strTmp & "," & "其他"
            End If
            
            strCoding = rsTemp("编码")
        End With
        rsTemp.MoveNext
    Loop

    If msf库房计量单位.Rows > 1 Then
        msf库房计量单位.FixedRows = 1
    End If
    gstrSQL = "select 库房id, 适用范围, 性质 from 药品库房单位"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngMaxRow = rsTemp.RecordCount
        For lngRow = 1 To lngMaxRow
            For i = 1 To msf库房计量单位.Rows - 1
                Select Case rsTemp!适用范围
                    Case 1
                        strTmp = "药库"
                    Case 2
                        strTmp = "门诊"
                    Case 3
                        strTmp = "住院"
                    Case 4
                        strTmp = "其他"
                End Select
                If rsTemp!库房id = msf库房计量单位.RowData(i) And strTmp = msf库房计量单位.TextMatrix(i, 1) Then
                    msf库房计量单位.TextMatrix(i, 2) = ""
                    msf库房计量单位.TextMatrix(i, 3) = ""
                    msf库房计量单位.TextMatrix(i, 4) = ""
                    msf库房计量单位.TextMatrix(i, 5) = ""
                    msf库房计量单位.TextMatrix(i, rsTemp!性质 + 1) = "√"
                End If
            Next
            rsTemp.MoveNext
        Next
    End If
    
    '药房配药控制
    strCoding = ""
    Set rsTemp = GetRSDrugStore(1)
    Bill药房配药控制.Clear
    lngRow = 1
    Do Until rsTemp.EOF
        With Bill药房配药控制
            If rsTemp("编码") <> strCoding Then
                strTmp = ""
            End If
            
            If InStr(",制剂室,中药房,西药房,成药房,", "," & rsTemp("工作性质") & ",") Then
            
                Select Case rsTemp("服务对象")
                    Case 0          '不服务于病人
                    Case 1          '服务于门诊病人
                        If InStr(1, strTmp & ",", ",门诊,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTmp = strTmp & "," & "门诊"
                        End If
                    Case 2          '服务于住院病人
                        If InStr(1, strTmp & ",", ",住院,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTmp = strTmp & "," & "住院"
                        End If
                    Case 3          '服务于门诊住院病人
                        If InStr(1, strTmp & ",", ",门诊,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTmp = strTmp & "," & "门诊"
                        End If
                        
                        If InStr(1, strTmp & ",", ",住院,") <= 0 Then
                            .Rows = lngRow + 1: lngRow = lngRow + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTmp = strTmp & "," & "住院"
                        End If
                End Select
            End If
            strCoding = rsTemp("编码")
        End With
        rsTemp.MoveNext
    Loop

    gstrSQL = "select 药房id, 门诊, 配药, 自动发药天数,配药确认 from 药房配药控制"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill药房配药控制
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            lngMaxRow = rsTemp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To .Rows - 1
                    Select Case rsTemp!门诊
                        Case 1
                            strTmp = "门诊"
                        Case 2
                            strTmp = "住院"
                    End Select
                    If rsTemp!药房id = .RowData(i) And strTmp = .TextMatrix(i, 1) Then
                        If IIF(IsNull(rsTemp("配药")), 0, rsTemp("配药")) = 1 Then
                            .TextMatrix(i, 2) = "√"
                        End If
                        
                        If IIF(IsNull(rsTemp("配药确认")), 0, rsTemp("配药确认")) = 1 Then
                            .TextMatrix(i, 4) = "√"
                        End If
                        .TextMatrix(i, 3) = IIF(IsNull(rsTemp!自动发药天数), "", rsTemp!自动发药天数)
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

Private Sub Load输液自备药清单()
    '功能：加载已设置的输液自备药清单
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strSql = "Select 药品id, '【' || b.编码 || '】' || b.名称 || '(' || b.规格 || ')' As 名称, 是否检查库存" & vbNewLine & _
            "From 输液自备药清单 A, 收费项目目录 B" & vbNewLine & _
            "Where a.药品id = b.Id" & vbNewLine & _
            "Order By 序号"

    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "Load输液自备药清单")
    
    vsf自备药清单.Rows = rsTemp.RecordCount + 2
    
    For i = 1 To rsTemp.RecordCount
        vsf自备药清单.TextMatrix(i, vsf自备药清单.ColIndex("药品id")) = rsTemp!药品ID
        vsf自备药清单.TextMatrix(i, vsf自备药清单.ColIndex("药品名称与编码")) = NVL(rsTemp!名称)
        vsf自备药清单.TextMatrix(i, vsf自备药清单.ColIndex("检查库存")) = IIF(rsTemp!是否检查库存 = 0, "", "√")
        
        rsTemp.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load药品库房流向()
'功能:装入药品流向数据
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With Bill(bill_药品库房流向)
        '首向装入可选库房
        gstrSQL = "Select Distinct b.Id, Nvl(b.编码, '') 编码, Nvl(b.名称, '') 名称 " & vbNewLine & _
            " From 部门性质说明 A, 部门表 B " & vbNewLine & _
            " Where b.Id = a.部门id And a.工作性质 In ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房') And " & vbNewLine & _
            " (b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or b.撤档时间 Is Null) " & vbNewLine & _
            "Order By 编码 "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("编码") & "-" & rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("ID")
            
            rsTemp.MoveNext
        Loop
        
        '装入流向控制数据
        gstrSQL = "select A.所在库房ID,A.对方库房ID,A.流向" & _
                "    ,B.编码 as 所在编码,B.名称 as 所在名称,C.编码 as 对方编码,C.名称 as 对方名称 " & _
                " from 药品流向控制 A,部门表 B,部门表 C " & _
                " where A.所在库房ID= B.ID and A.对方库房ID=C.ID and " & Where撤档时间("C") & _
                " order by b.编码,c.编码 "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("所在库房ID")
            .TextMatrix(lngRow, 0) = rsTemp("所在编码") & "-" & rsTemp("所在名称")
            .TextMatrix(lngRow, 1) = rsTemp("对方编码") & "-" & rsTemp("对方名称")
            .TextMatrix(lngRow, 2) = rsTemp("对方库房ID")
            .TextMatrix(lngRow, 3) = Switch(rsTemp("流向") = 1, "1-所在库房可流向对方库房", _
                                            rsTemp("流向") = 2, "2-对方库房可流向所在库房", _
                                                          True, "3-两库房间可双向流通")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load库房检查()
    '功能：初始化库房
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim ObjItem As ListItem
    On Error GoTo errHandle
    
    gstrSQL = _
        "SELECT B.ID,B.编码, B.名称, NVL(C.检查方式, 0) 检查方式" & vbCrLf & _
        " FROM 部门性质说明 A, 部门表 B, 药品出库检查 C" & vbCrLf & _
        " WHERE A.部门ID = B.ID AND A.部门ID = C.库房ID(+) AND" & vbCrLf & _
        "      A.工作性质 IN" & vbCrLf & _
        "      ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房')" & vbCrLf & _
        "     And (b.撤档时间=to_date('3000-1-1','yyyy-mm-dd') or b.撤档时间 is null) " & vbCrLf & _
        " GROUP BY B.ID,B.编码, B.名称, NVL(C.检查方式, 0) " & vbCrLf & _
        " order by B.编码 "
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.lvw库存检查.ListItems.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set ObjItem = Me.lvw库存检查.ListItems.Add(, "C_" & rsTmp!ID, rsTmp!编码)
            ObjItem.SubItems(1) = "" & rsTmp!名称
            ObjItem.SubItems(2) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
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

Private Sub Save库房单位()
    '保存库房单位设置
    Dim i As Long
    Dim lngTmp As Long
    Dim intTmp As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    With msf库房计量单位
        If .Rows > 1 And .Tag = "已修改" Then
            If Trim(.TextMatrix(1, 0)) <> "" Then
                gstrSQL = ""
                For i = 1 To .Rows - 1
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    lngTmp = 1
                    Select Case True
                        Case .TextMatrix(i, 2) = "√"
                            lngTmp = 1
                        Case .TextMatrix(i, 3) = "√"
                            lngTmp = 2
                        Case .TextMatrix(i, 4) = "√"
                            lngTmp = 3
                        Case .TextMatrix(i, 5) = "√"
                            lngTmp = 4
                    End Select
                    Select Case .TextMatrix(i, 1)
                        Case "药库"
                            intTmp = 1
                        Case "门诊"
                            intTmp = 2
                        Case "住院"
                            intTmp = 3
                        Case "其他"
                            intTmp = 4
                    End Select
                    gstrSQL = gstrSQL & lngTmp & "," & intTmp & ","
                Next
                strSql = "ZL_药品库房单位_DELETE"
                Call zldatabase.ExecuteProcedure(strSql, Me.Caption)
                
                gstrSQL = "ZL_药品库房单位_INSERT('" & gstrSQL & "')"
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


Private Sub Save库房检查()
    '功能：保存库房检查
    Dim i As Long
    On Error GoTo errHandle
    
    If lvw库存检查.Tag = "已修改" Then
        gstrSQL = ""
        For i = 1 To Me.lvw库存检查.ListItems.Count
            gstrSQL = gstrSQL & Me.lvw库存检查.ListItems(i).Tag & "," & Switch(Me.lvw库存检查.ListItems(i).SubItems(2) = "0-不检查", "0", Me.lvw库存检查.ListItems(i).SubItems(2) = "1-检查，不足提醒", "1", Me.lvw库存检查.ListItems(i).SubItems(2) = "2-检查，不足禁止", "2") & ","
        Next
        gstrSQL = "Zl_药品出库检查_insert('" & gstrSQL & "')"
        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        lvw库存检查.Tag = ""
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Save药品领用流向()
    Dim strTmp As String
    Dim lngRow As Long
    Dim bln次数 As Boolean
    
    On Error GoTo ErrHand
    With Bill(bill_药品领用流向)
        If .Tag = "已修改" Then
            For lngRow = 1 To .Rows - 1
                If .RowData(lngRow) > 0 Then
                    If LenB(StrConv(strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ",", vbFromUnicode)) >= 4000 Then
                        If bln次数 = True Then
                            gstrSQL = "zl_药品领用流向控制_Modify('" & strTmp & "'," & 1 & ")"
                        Else
                            gstrSQL = "zl_药品领用流向控制_Modify('" & strTmp & "'," & 0 & ")"
                        End If
                        bln次数 = True
                        Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        
                        strTmp = .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                    Else
                        strTmp = strTmp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                    End If
                End If
            Next
    
            If bln次数 = True Then
                gstrSQL = "zl_药品领用流向控制_Modify('" & strTmp & "'," & 1 & ")"
            Else
                gstrSQL = "zl_药品领用流向控制_Modify('" & strTmp & "'," & 0 & ")"
            End If
            bln次数 = True
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

Sub Load药品领用库房()
'功能:读入药品领用部门
     
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With Bill(bill_药品领用流向)
        '装入流向控制数据
        gstrSQL = "Select Distinct b.Id, Nvl(b.编码, '') 编码, Nvl(b.名称, '') 名称 " & vbNewLine & _
            " From 部门性质说明 A, 部门表 B " & vbNewLine & _
            " Where b.Id = a.部门id And a.工作性质 In ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房') And " & vbNewLine & _
            " (b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or b.撤档时间 Is Null) " & vbNewLine & _
            "Order By 编码 "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("编码") & "-" & rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("ID")
            rsTemp.MoveNext
        Loop
        
        '装入流向控制数据
        gstrSQL = "select A.领用部门ID,A.对方库房ID" & _
                ",B.编码 as 领用部门编码,B.名称 as 领用部门名称,C.编码 as 库房编码,C.名称 as 库房名称 " & _
                " from 药品领用控制 A,部门表 B,部门表 C " & _
                " where A.领用部门ID= B.ID and A.对方库房ID=C.ID order by b.编码,c.编码 "
        Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("领用部门ID")
            .TextMatrix(lngRow, 0) = rsTemp("领用部门编码") & "-" & rsTemp("领用部门名称")
            .TextMatrix(lngRow, 1) = rsTemp("库房编码") & "-" & rsTemp("库房名称")
            .TextMatrix(lngRow, 2) = rsTemp("对方库房ID")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub opt结存时间模式_Click(Index As Integer)
    Dim strValue As String
    
    txt(txt_结存时间模式).Enabled = opt结存时间模式(1).value
    
    If Me.Visible Then
        strValue = IIF(opt结存时间模式(0).value, 0, Val(txt(txt_结存时间模式).Text))
        Call SetParChange(txt, txt_结存参数值, mrsPar, True, strValue)
        
        opt结存方式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存方式(1).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(1).ForeColor = txt(txt_结存参数值).ForeColor
        txt(txt_结存时间模式).ForeColor = opt结存时间模式(1).ForeColor
    End If
End Sub

Private Sub msf库房计量单位_DblClick()
    Dim i As Long
    
    With msf库房计量单位
    If .Col > 1 And .Row > 0 And Trim(.TextMatrix(.Row, 0)) <> "" Then
        .TextMatrix(.Row, 2) = ""
        .TextMatrix(.Row, 3) = ""
        .TextMatrix(.Row, 4) = ""
        .TextMatrix(.Row, 5) = ""
        .TextMatrix(.Row, .Col) = "√"
        
        .Tag = "已修改"
    End If
    
    End With
End Sub

Private Sub msf库房计量单位_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn Or KeyAscii = Asc(" ")) Then
        msf库房计量单位_DblClick
    End If
End Sub

Private Sub opt发药窗口_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvw库存检查_DblClick()
    With lvw库存检查
        If Not .SelectedItem Is Nothing Then
            .SelectedItem.SubItems(2) = Switch(.SelectedItem.SubItems(2) = "0-不检查", "1-检查，不足提醒", _
                .SelectedItem.SubItems(2) = "1-检查，不足提醒", "2-检查，不足禁止", .SelectedItem.SubItems(2) = "2-检查，不足禁止", "0-不检查")
            .Tag = "已修改"
        End If
    End With
End Sub

Private Sub lvw库存检查_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = "C" Then
        Call lvw库存检查_DblClick
    End If
End Sub

Private Function Check是否有未审核的外购入库单() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 未审药品记录 Where 单据 = 1 And Rownum < 2"
    Call zldatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check是否有未审核的外购入库单 = (rs.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Bill药品卫材精度_EnterCell(Row As Long, Col As Long)
    With Bill药品卫材精度
        If Col = dig_精度 Then
            .TxtCheck = True
            .TextMask = "123456789"
            .MaxLength = 1
        End If
    End With
End Sub

Private Sub opt发药窗口_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt发药窗口, Index, mrsPar)
    End If
End Sub

Private Sub Bill药品卫材精度_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Bill药品卫材精度
        If .Col = dig_精度 Then
            If .Text = "" Then Exit Sub
            
            .Text = Val(.Text)
            strKey = .Text
            
            If Val(strKey) > .TextMatrix(.Row, dig_最大精度) Or Val(strKey) < .TextMatrix(.Row, dig_最小精度) Then
                MsgBox "精度超过允许范围！", vbInformation, gstrSysName
                .Text = .RowData(.Row)
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            
            If cbo(cbo_零差价模式).ListIndex > 0 Then
                If .TextMatrix(.Row, dig_精度类别) = "药品" Then
                    MsgBox "注意，已启用零差价管理模式，如果调整精度可能将影响差价金额计算！", vbInformation, gstrSysName
                End If
            End If
            
            .TextMatrix(.Row, .Col) = strKey
            .RowData(.Row) = Val(strKey)
            
            .Tag = "已修改"
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
    Case chk_时价分段加成入库
        If chk(Index).value = 1 Then
            If chk(chk_时价加成率入库).value = 1 Then chk(chk_时价加成率入库).value = 0
            If chk(chk_时价药品取上次售价).value = 1 Then chk(chk_时价药品取上次售价).value = 0
        End If
    Case chk_时价加成率入库
        If chk(Index).value = 1 Then
            If chk(chk_时价分段加成入库).value = 1 Then chk(chk_时价分段加成入库).value = 0
            If chk(chk_时价药品取上次售价).value = 1 Then chk(chk_时价药品取上次售价).value = 0
        End If
    Case chk_时价药品取上次售价
        If chk(Index).value = 1 Then
            If chk(chk_时价分段加成入库).value = 1 Then chk(chk_时价分段加成入库).value = 0
            If chk(chk_时价加成率入库).value = 1 Then chk(chk_时价加成率入库).value = 0
        End If
    Case chk_申领按批次出库
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_申领按批次出库).value = Val(GetParOriginalValue(chk, chk_申领按批次出库, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check申领单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的申领单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_申领按批次出库).value = Val(GetParOriginalValue(chk, chk_申领按批次出库, mrsPar))
        End If
    Case chk_领用按批次出库
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_领用按批次出库).value = Val(GetParOriginalValue(chk, chk_领用按批次出库, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check领用单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的领用单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_领用按批次出库).value = Val(GetParOriginalValue(chk, chk_领用按批次出库, mrsPar))
        End If
    Case chk_移库按批次出库
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_移库按批次出库).value = Val(GetParOriginalValue(chk, chk_移库按批次出库, mrsPar)) Then Exit Sub
        
        On Error GoTo errHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check移库单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的移库单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_移库按批次出库).value = Val(GetParOriginalValue(chk, chk_移库按批次出库, mrsPar))
        End If
    Case chk_移库冲销申请
        '当变为不需要申请时，要检查是否有未审核的冲销申请单，如果有则不能改变
        
        On Error GoTo errHandle
        If chk(chk_移库冲销申请).value = 0 Then
            If MsgBox("即将检查是否存在未审核的冲销申请单，可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '该功能是10.20版本新增，增加一个条件填制日期范围，避免全表扫描
                gstrSQL = "Select 1 From 未审药品记录 A " & _
                    " Where a.单据 = 6 And a.填制日期 Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Mod(b.记录状态, 3) = 2) And Rownum < 2"
                
                DoEvents
                zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
                
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的冲销申请单")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "存在未审核的冲销申请单，不能改变此参数！", vbInformation, gstrSysName
                    chk(chk_移库冲销申请).value = 1
                End If
            Else
                chk(chk_移库冲销申请).value = 1
            End If
        End If
    Case chk_领用冲销申请
        '当变为不需要申请时，要检查是否有未审核的冲销申请单，如果有则不能改变
        
        On Error GoTo errHandle
        If chk(chk_领用冲销申请).value = 0 Then
            If MsgBox("即将检查是否存在未审核的冲销申请单，可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSQL = "Select 1 From 未审药品记录 A " & _
                    " Where a.单据 = 7 And a.填制日期 Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Mod(b.记录状态, 3) = 2) And Rownum < 2"
                
                DoEvents
                zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
                
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的冲销申请单")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "存在未审核的冲销申请单，不能改变此参数！", vbInformation, gstrSysName
                    chk(chk_领用冲销申请).value = 1
                End If
            Else
                chk(chk_领用冲销申请).value = 1
            End If
        End If
    Case chk_门诊处方审查
        If Me.Visible = False Then Exit Sub
        
        If chk(Index).value = 0 Then
            '检查最近未审查的记录
            If GetRecipeAuditBills(1) Then
                MsgBox "处方审查系统最近存在未审查的记录，请检查！", vbInformation, gstrSysName
                chk(Index).value = 1
            End If
        End If
        
        optOpporunity(1).Enabled = chk(Index).value = 1
        optOpporunity(2).Enabled = chk(Index).value = 1
        
        If chk(Index).value = 0 Then chk(chk_提醒门诊医生).value = 0
        chk(chk_提醒门诊医生).Enabled = chk(Index).value = 1
        
        txt(txt_门诊药师审方离岗时长).Enabled = chk(Index).value = 1
        
        If chk(chk_门诊处方审查).value = 1 And chk(chk_住院药嘱审查).value = 1 Then
            strVar = "3"
        ElseIf chk(chk_门诊处方审查).value = 0 And chk(chk_住院药嘱审查).value = 1 Then
            strVar = "2"
        ElseIf chk(chk_门诊处方审查).value = 1 And chk(chk_住院药嘱审查).value = 0 Then
            strVar = "1"
        Else
            strVar = "0"
        End If
        Call SetParChange(chk, Index, mrsPar, True, strVar)
        chk(chk_住院药嘱审查).ForeColor = chk(Index).ForeColor
            
    Case chk_住院药嘱审查
        If Me.Visible = False Then Exit Sub
        
        If chk(Index).value = 0 Then
            '检查最近未审查的记录
            If GetRecipeAuditBills(2) Then
                MsgBox "处方审查系统最近存在未审查的记录，请检查！", vbInformation, gstrSysName
                chk(Index).value = 1
            End If
        End If
        
        If chk(Index).value = 0 Then chk(chk_提醒住院医生).value = 0
        chk(chk_提醒住院医生).Enabled = chk(Index).value = 1
        
        If chk(chk_门诊处方审查).value = 1 And chk(chk_住院药嘱审查).value = 1 Then
            strVar = "3"
        ElseIf chk(chk_门诊处方审查).value = 0 And chk(chk_住院药嘱审查).value = 1 Then
            strVar = "2"
        ElseIf chk(chk_门诊处方审查).value = 1 And chk(chk_住院药嘱审查).value = 0 Then
            strVar = "1"
        Else
            strVar = "0"
        End If
        Call SetParChange(chk, chk_门诊处方审查, mrsPar, True, strVar)
        chk(Index).ForeColor = chk(chk_门诊处方审查).ForeColor
        
    Case chk_门诊处方自动发送
        If Me.Visible = False Then Exit Sub
        
        Call SetParChange(chk, chk_门诊处方自动发送, mrsPar)
    End Select

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check申领单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From 未审药品记录 A " & _
        " Where a.单据 = 6 And a.填制日期 > Sysdate - 90 And Exists " & _
        " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Nvl(b.发药方式,0) = 1) And Rownum < 2"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的申领单")
    
    Check申领单 = rsTemp.RecordCount = 0
End Function

Private Function Check移库单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From 未审药品记录 A " & _
        " Where a.单据 = 6 And a.填制日期 > Sysdate - 90 And Exists " & _
        " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Nvl(b.发药方式,0) <> 1) And Rownum < 2"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的移库单")
    
    Check移库单 = rsTemp.RecordCount = 0
End Function

Private Function Check领用单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select 1 From 未审药品记录 Where 单据 = 7 And 填制日期 > Sysdate - 90 And Rownum < 2"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的领用单")
    
    Check领用单 = rsTemp.RecordCount = 0
End Function
Private Sub cmdlst输液中心发药病人科室_Click(Index As Integer)
    Dim i As Long
    
    If chk来源科室.value = 0 Then Exit Sub
    
    With lst(lst_PIVA来源科室)
        For i = 0 To .ListCount - 1
            .Selected(i) = Index = 0    '将触发lst_ItemCheck事件
        Next
    End With
End Sub

Private Function GetRecipeAuditBills(ByVal bytType As Byte) As Boolean
'功能：检查最近门诊或住院的“处方审查记录”是否存在未审查的记录
'参数：
'  bytType：1-门诊；2-住院
'返回：True存在未审查的记录；False不存在未审查的记录

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    If bytType = 1 Then
        '门诊
        strSql = "Select ID From 处方审查记录 Where 状态 = 0 And 提交时间 >= Trunc(Sysdate - [1]) And 挂号Id Is Not Null And Rownum < 2 "
    Else
        '住院
        strSql = "Select ID From 处方审查记录 Where 状态 = 0 And 提交时间 >= Trunc(Sysdate - [1]) And 主页Id Is Not Null And Rownum < 2 "
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "检查未审查的处方审查记录", IIF(bytType = 1, 3, 5))
    GetRecipeAuditBills = rsTemp.EOF = False
    rsTemp.Close
    
    Exit Function

hErr:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Sub vsf自备药清单_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsf自备药清单.ColIndex("药品名称与编码") Then Cancel = True
End Sub

Private Sub vsf自备药清单_Click()
    With vsf自备药清单
        If .Row < 1 Then Exit Sub
        
        If .Col = .ColIndex("检查库存") And .TextMatrix(.Row, .ColIndex("药品id")) <> "" Then
            If .TextMatrix(.Row, .ColIndex("检查库存")) = "" Then
                .TextMatrix(.Row, .ColIndex("检查库存")) = "√"
            Else
                .TextMatrix(.Row, .ColIndex("检查库存")) = ""
            End If
        End If
    End With
End Sub

Private Sub vsf自备药清单_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        With vsf自备药清单
            If .Rows = 2 Then
                .TextMatrix(.Row, .ColIndex("药品id")) = ""
                .TextMatrix(.Row, .ColIndex("药品名称与编码")) = ""
            Else
                .RemoveItem vsf自备药清单.Row
            End If
        End With
    End If
End Sub

Private Sub vsf自备药清单_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strKey As String
    Dim strCode As String
    
    If KeyCode = 13 Then
        vRect = zlControl.GetControlRect(vsf自备药清单.hwnd)
        dblLeft = vRect.Left + vsf自备药清单.CellLeft
        dblTop = vRect.Top + vsf自备药清单.CellTop + vsf自备药清单.CellHeight + 3200
        
        With vsf自备药清单
            If Col = .ColIndex("药品名称与编码") Then
                strKey = Trim(.EditText)
                If strKey = "" Then Exit Sub
                
                If IsNumeric(strKey) Then
                    '纯数字
                    strCode = " d.编码 like [1] "
                ElseIf zlCommFun.IsCharAlpha(strKey) Then
                    '纯字母
                    strCode = " n.简码 Like [1] "
                ElseIf zlCommFun.IsCharChinese(strKey) Then
                    '纯汉字
                    strCode = " d.名称 like [1] "
                Else
                    strCode = " (n.简码 Like [1] Or d.编码 Like [1] Or n.名称 Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')' As 通用名" & vbNewLine & _
                    " From 药品规格 T, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    " Where t.药品id = d.Id And t.药品id = n.收费细目id And D.类别 In ('5', '6') And" & strCode & vbNewLine & _
                    " And (d.撤档时间 Is Null Or To_Char(d.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')'"
                Set rsRecord = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品名称与编码", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .Rows - 1
                        If rsRecord!ID = Val(.TextMatrix(i, .ColIndex("药品ID"))) Then
                            MsgBox rsRecord!通用名 & "已经录入，请重新选择！", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!ID
                    .TextMatrix(.Row, .ColIndex("药品名称与编码")) = rsRecord!通用名
                    .EditText = rsRecord!通用名
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

