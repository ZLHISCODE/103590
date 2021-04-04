VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmParStuff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "卫材参数设置"
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
   StartUpPosition =   1  '所有者中心
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
      TabCaption(0)   =   "目录(&0)"
      TabPicture(0)   =   "frmParStuff.frx":74F2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picPar(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "入出(&1)"
      TabPicture(1)   =   "frmParStuff.frx":750E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPar(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "在库(&2)"
      TabPicture(2)   =   "frmParStuff.frx":752A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picPar(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "发放(&3)"
      TabPicture(3)   =   "frmParStuff.frx":7546
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picPar(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "流向控制(&10)"
      TabPicture(4)   =   "frmParStuff.frx":7562
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picPar(10)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "库存检查(&11)"
      TabPicture(5)   =   "frmParStuff.frx":757E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picPar(11)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "虚拟库房(&12)"
      TabPicture(6)   =   "frmParStuff.frx":759A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picPar(12)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "卫材精度(&13)"
      TabPicture(7)   =   "frmParStuff.frx":75B6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "picPar(13)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "环节(&14)"
      TabPicture(8)   =   "frmParStuff.frx":75D2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "picPar(14)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "通用(&4)"
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
            Caption         =   "条码卫材识别控制"
            ForeColor       =   &H00800000&
            Height          =   1080
            Left            =   165
            TabIndex        =   142
            Top             =   105
            Width           =   5655
            Begin VB.OptionButton optBarcode 
               Caption         =   "只允许输入条码或扫码进行识别"
               Height          =   240
               Index           =   1
               Left            =   75
               TabIndex        =   144
               Top             =   660
               Width           =   4755
            End
            Begin VB.OptionButton optBarcode 
               Caption         =   "允许输入简码、编码、条码等进行识别"
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
         Begin VSFlex8Ctl.VSFlexGrid vsf单据环节控制 
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
         Begin VB.Label lbl单据控制 
            AutoSize        =   -1  'True
            Caption         =   "单据环节控制：设置卫材单据在特定业务环节中允许修改的项目"
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
         Begin ZL9BillEdit.BillEdit Bill药品卫材精度 
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
         Begin VB.Label lbl精度说明 
            Caption         =   $"frmParStuff.frx":760A
            ForeColor       =   &H00000080&
            Height          =   720
            Left            =   240
            TabIndex        =   134
            Top             =   6600
            Width           =   7995
         End
         Begin VB.Label lbl精度 
            AutoSize        =   -1  'True
            Caption         =   "卫材精度设置：按包装单位来设置价格、数量允许录入的精度（保留的小数位数）"
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
            Caption         =   "不严格控制卫材指导批价和指导售价"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   44
            Top             =   120
            Width           =   3615
         End
         Begin VB.Frame fra卫材定价单位 
            Caption         =   " 卫材目录定价单位"
            ForeColor       =   &H00800000&
            Height          =   735
            Left            =   240
            TabIndex        =   41
            Top             =   1920
            Width           =   4740
            Begin VB.OptionButton opt定价单位 
               Caption         =   "散装单位"
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   43
               Top             =   360
               Value           =   -1  'True
               Width           =   1185
            End
            Begin VB.OptionButton opt定价单位 
               Caption         =   "包装单位"
               Height          =   285
               Index           =   1
               Left            =   1560
               TabIndex        =   42
               Top             =   360
               Width           =   1425
            End
         End
         Begin VB.Frame fra编码递增模式 
            Caption         =   " 编码递增模式"
            ForeColor       =   &H00800000&
            Height          =   1275
            Left            =   240
            TabIndex        =   37
            Top             =   480
            Width           =   4740
            Begin VB.OptionButton opt编码模式 
               Caption         =   "分类号+顺序编号"
               Height          =   210
               Index           =   2
               Left            =   240
               TabIndex        =   40
               Top             =   960
               Width           =   3420
            End
            Begin VB.OptionButton opt编码模式 
               Caption         =   "材料类别+分类号+顺序编号"
               Height          =   210
               Index           =   1
               Left            =   240
               TabIndex        =   39
               Top             =   660
               Width           =   3420
            End
            Begin VB.OptionButton opt编码模式 
               Caption         =   "同类顺序编号"
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
            Caption         =   " 设置卫材对应缺省收入项目"
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
               Caption         =   "卫生材料"
               Height          =   180
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   " 卫材分批属性自动设置"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   240
            TabIndex        =   29
            Top             =   3720
            Width           =   4740
            Begin VB.OptionButton opt分批属性 
               Caption         =   "库房和发料部门分批"
               Height          =   210
               Index           =   2
               Left            =   240
               TabIndex        =   33
               Top             =   720
               Width           =   1980
            End
            Begin VB.OptionButton opt分批属性 
               Caption         =   "仅库房分批"
               Height          =   210
               Index           =   1
               Left            =   2280
               TabIndex        =   32
               Top             =   360
               Width           =   1500
            End
            Begin VB.OptionButton opt分批属性 
               Caption         =   "库房和发料部门都不分批"
               Height          =   210
               Index           =   3
               Left            =   2280
               TabIndex        =   31
               Top             =   720
               Width           =   2385
            End
            Begin VB.OptionButton opt分批属性 
               Caption         =   "手工设置分批属性"
               Height          =   210
               Index           =   0
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   1740
            End
         End
         Begin VB.Frame fra 
            Caption         =   " 设置存储库房时允许应用于的范围"
            ForeColor       =   &H00800000&
            Height          =   2385
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   4920
            Width           =   4785
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于分类下所有卫生材料"
               Height          =   255
               Index           =   2
               Left            =   270
               TabIndex        =   26
               Top             =   840
               Width           =   2760
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于本级所有卫生材料"
               Height          =   255
               Index           =   1
               Left            =   270
               TabIndex        =   25
               Top             =   562
               Width           =   2712
            End
            Begin VB.CheckBox chk应用范围 
               Caption         =   "应用于所有卫生材料"
               Height          =   255
               Index           =   0
               Left            =   270
               TabIndex        =   24
               Top             =   285
               Width           =   2364
            End
            Begin VB.Label lblInfor 
               Caption         =   "   如:没有勾上此栏目中的“应用于所有卫生材料”，则在存储库房设置界面中的『应用于所有“卫生材料”(4)』将不能选择！"
               ForeColor       =   &H00000080&
               Height          =   615
               Index           =   2
               Left            =   240
               TabIndex        =   28
               Top             =   1680
               Width           =   4260
            End
            Begin VB.Label lblInfor 
               Caption         =   "    本栏目主要是控制卫生材料管理的存储库房设置界面中的“应用于...”功能。"
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
         Begin VSFlex8Ctl.VSFlexGrid vsf对照 
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
         Begin VB.Label lbl虚拟库房 
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
         Begin VSFlex8Ctl.VSFlexGrid vsf库房检查 
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
         Begin VB.Label lbl提示 
            Caption         =   "  在这里可以选择各库房是否检查库存及库存检查方式。当库房选中时双击“库房检查方式”列可改变库房的检查方式。"
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
         Begin VSFlex8Ctl.VSFlexGrid vsf流向 
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
         Begin VB.Label lbl流向 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "控制材料在不同库房间的流通方向"
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
            Caption         =   "退料时自动将记帐费用销帐"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   111
            Top             =   960
            Width           =   2880
         End
         Begin VB.Frame fra 
            Caption         =   " 病区发料单据过滤控制 "
            ForeColor       =   &H00800000&
            Height          =   1425
            Index           =   3
            Left            =   240
            TabIndex        =   101
            Top             =   4800
            Width           =   4455
            Begin VB.CheckBox chkDeptType 
               Caption         =   "营养"
               Enabled         =   0   'False
               Height          =   255
               Index           =   6
               Left            =   2280
               TabIndex        =   110
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "治疗"
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   1440
               TabIndex        =   109
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "手术"
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   600
               TabIndex        =   108
               Top             =   960
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "检验"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   3120
               TabIndex        =   107
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "检查"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   106
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "护理"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   105
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chkDeptType 
               Caption         =   "临床"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   104
               Top             =   660
               Width           =   735
            End
            Begin VB.CheckBox chk病区 
               Caption         =   "按病区发料时包含非病人科室开单的记录"
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
               Text            =   "存参数值"
               Top             =   960
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "领料人签名"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   100
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CheckBox chk 
            Caption         =   "卫材医嘱按发生时间过滤"
            Height          =   255
            Index           =   25
            Left            =   240
            TabIndex        =   99
            Top             =   1680
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "卫材在门诊收费或记帐后自动发料"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   98
            Top             =   240
            Width           =   3000
         End
         Begin VB.Frame fra未收费发药 
            Caption         =   " 未收费或审核时处方发料"
            ForeColor       =   &H00800000&
            Height          =   1695
            Left            =   240
            TabIndex        =   93
            Top             =   2880
            Width           =   4455
            Begin VB.CheckBox chk 
               Caption         =   "未收费的门诊划价处方发料"
               Height          =   180
               Index           =   27
               Left            =   240
               TabIndex        =   96
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "未审核的记账处方发料"
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   95
               Top             =   1200
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "项目执行前先收费或审核"
               Height          =   195
               Index           =   26
               Left            =   480
               TabIndex        =   94
               Top             =   0
               Visible         =   0   'False
               Width           =   2880
            End
            Begin VB.Label lbl未收费发药 
               Caption         =   "  如果启用了门诊一卡通参数""执行前必须先收费或先记帐审核""，则对门诊病人发料时，以下参数将失效。"
               ForeColor       =   &H00000080&
               Height          =   615
               Left            =   240
               TabIndex        =   97
               Top             =   240
               Width           =   3855
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "是否自动缺料检查"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   92
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chk 
            Caption         =   "发料时汇总销帐申请记录"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   91
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "卫材在住院记帐后自动发料方式"
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
            Caption         =   "卫材结存"
            ForeColor       =   &H00800000&
            Height          =   1860
            Index           =   10
            Left            =   6000
            TabIndex        =   123
            Top             =   1080
            Width           =   3975
            Begin VB.Frame fra自动结存方式 
               Caption         =   " 设置结存时点"
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
                     Name            =   "宋体"
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
               Begin VB.OptionButton opt结存时间模式 
                  Caption         =   "每月最后一天"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   130
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   1455
               End
               Begin VB.OptionButton opt结存时间模式 
                  Caption         =   "每月    日"
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
            Begin VB.OptionButton opt结存方式 
               Caption         =   "自动结存(各库房按同一日期结存)"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   125
               Top             =   720
               Value           =   -1  'True
               Width           =   3495
            End
            Begin VB.OptionButton opt结存方式 
               Caption         =   "手工结存(各库房可以不同日期结存)"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   124
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价卫材按批次调价"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   90
            Top             =   600
            Width           =   3105
         End
         Begin VB.Frame fra资质校验 
            Caption         =   " 计划单资质校验"
            ForeColor       =   &H00800000&
            Height          =   6255
            Index           =   1
            Left            =   240
            TabIndex        =   83
            Top             =   1080
            Width           =   5295
            Begin VB.Frame fraCheck 
               Caption         =   "选择校验方式"
               ForeColor       =   &H00800000&
               Height          =   615
               Index           =   1
               Left            =   120
               TabIndex        =   85
               Top             =   5520
               Width           =   4935
               Begin VB.OptionButton opt计划资质校验 
                  Caption         =   "校验未通过时禁止保存"
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   87
                  Top             =   280
                  Width           =   2175
               End
               Begin VB.OptionButton opt计划资质校验 
                  Caption         =   "校验未通过时提醒"
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
               Text            =   "存参数值"
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
            Caption         =   "允许盘点没有设置存储库房的卫材"
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
         Begin VB.Frame fra产地 
            Caption         =   " 入库产地信息取值方式"
            ForeColor       =   &H00800000&
            Height          =   1470
            Left            =   120
            TabIndex        =   119
            Top             =   1800
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "分批卫材批号产地控制"
               Height          =   255
               Index           =   34
               Left            =   240
               TabIndex        =   138
               Top             =   1080
               Width           =   2895
            End
            Begin VB.OptionButton opt产地 
               Caption         =   "优先取目录中的产地"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   121
               Top             =   480
               Width           =   2295
            End
            Begin VB.OptionButton opt产地 
               Caption         =   "优先取上次入库的产地"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   120
               Top             =   240
               Value           =   -1  'True
               Width           =   2295
            End
         End
         Begin VB.Frame fra领用 
            Caption         =   " 领用流程控制"
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   4560
            TabIndex        =   78
            Top             =   4680
            Width           =   5565
            Begin VB.CheckBox chk 
               Caption         =   "在领用审核前需要进行财务核查"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   81
               Top             =   720
               Width           =   3180
            End
            Begin VB.CheckBox chk 
               Caption         =   "不允许具有""跟踪在用""属性的卫材进行领用"
               Height          =   255
               Index           =   17
               Left            =   120
               TabIndex        =   80
               Top             =   480
               Width           =   3810
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许向发料部门领用卫生材料"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame fra移库流程控制 
            Caption         =   " 移库功能流程控制"
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   4560
            TabIndex        =   76
            Top             =   5880
            Width           =   5565
            Begin VB.CheckBox chk 
               Caption         =   "移库冲销时，移入库房需要先申请冲销"
               Height          =   255
               Index           =   16
               Left            =   180
               TabIndex        =   77
               Top             =   240
               Value           =   1  'Checked
               Width           =   3705
            End
         End
         Begin VB.Frame fra出库算法 
            Caption         =   " 出库优先规则"
            ForeColor       =   &H00800000&
            Height          =   615
            Left            =   4560
            TabIndex        =   73
            Top             =   6600
            Width           =   5565
            Begin VB.OptionButton opt卫材出库算法 
               Caption         =   "按批次先进先出"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton opt卫材出库算法 
               Caption         =   "按效期最近先出"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   74
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame fra出库控制 
            Caption         =   " 出库控制"
            ForeColor       =   &H00800000&
            Height          =   1332
            Left            =   120
            TabIndex        =   70
            Top             =   5880
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "按批次移库卫生材料"
               Height          =   255
               Index           =   31
               Left            =   120
               TabIndex        =   122
               Top             =   480
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "按批次领用卫生材料"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   118
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "按批次申领卫生材料"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "卫材填单下可用库存"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   2895
            End
         End
         Begin VB.Frame fra入库控制 
            Caption         =   " 入库价格控制"
            ForeColor       =   &H00800000&
            Height          =   1575
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "时价卫材以加价率入库"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   69
               Top             =   240
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价卫材按分段加成率入库"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   68
               Top             =   480
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价卫材入库按扣前加成销售"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   67
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价卫材入库时允许手工调整售价"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   66
               Top             =   1200
               Width           =   3135
            End
            Begin VB.CheckBox chk 
               Caption         =   "时价卫材入库时取上次售价"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   65
               Top             =   720
               Width           =   2895
            End
         End
         Begin VB.Frame fra外购入库 
            Caption         =   " 外购入库操作控制"
            ForeColor       =   &H00800000&
            Height          =   2415
            Left            =   120
            TabIndex        =   53
            Top             =   3360
            Width           =   3975
            Begin VB.CheckBox chk 
               Caption         =   "外购入库单需要核查"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   62
               Top             =   960
               Width           =   2895
            End
            Begin VB.CheckBox chk 
               Caption         =   "允许修改采购限价"
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   2700
            End
            Begin VB.CheckBox chk 
               Caption         =   "招标卫材可选择非中标单位入库"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   2880
            End
            Begin VB.CheckBox chk 
               Caption         =   "高值卫材必须填写详细信息"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   59
               Top             =   720
               Width           =   2880
            End
            Begin VB.Frame fraBidMess 
               Caption         =   " 采购价超中标价格时"
               ForeColor       =   &H00800000&
               Height          =   615
               Left            =   120
               TabIndex        =   55
               Top             =   1680
               Width           =   3315
               Begin VB.OptionButton opt采购价 
                  Caption         =   "禁止"
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opt采购价 
                  Caption         =   "提示"
                  Height          =   180
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   57
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton opt采购价 
                  Caption         =   "不限制"
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
            Begin VB.Label lbl条码前缀提示 
               AutoSize        =   -1  'True
               Caption         =   "卫材条码前缀(2-8位数字或字母)"
               Height          =   180
               Left            =   120
               TabIndex        =   63
               Top             =   1320
               Width           =   2610
            End
         End
         Begin VB.Frame fra资质校验 
            Caption         =   " 外购入库资质校验"
            ForeColor       =   &H00800000&
            Height          =   4455
            Index           =   0
            Left            =   4560
            TabIndex        =   45
            Top             =   120
            Width           =   5535
            Begin VB.CheckBox chk 
               Caption         =   "生产日期大于注册证效期检查"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   50
               Top             =   4080
               Width           =   2775
            End
            Begin VB.Frame fraCheck 
               Caption         =   "选择校验方式"
               ForeColor       =   &H00800000&
               Height          =   615
               Index           =   0
               Left            =   120
               TabIndex        =   47
               Top             =   3360
               Width           =   4935
               Begin VB.OptionButton opt外购资质校验 
                  Caption         =   "校验未通过时提醒"
                  Height          =   180
                  Index           =   1
                  Left            =   2520
                  TabIndex        =   49
                  Top             =   280
                  Width           =   1935
               End
               Begin VB.OptionButton opt外购资质校验 
                  Caption         =   "校验未通过时禁止保存"
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
               Text            =   "存参数值"
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
Attribute VB_Name = "frmParStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset '参数与控件对应记录集（同一个参数可能对应一组多个控件）
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mstrLike As String
Private mrs对照 As New ADODB.Recordset
Private mblnLoad As Boolean     '窗体加载结束
Private mblnOk As Boolean

Private Enum constCbo
    cbo_收入项目 = 0
    cbo_住院卫材自动发料 = 1
End Enum

'允许控制的所有项目
Private Const cst所有项目 As String = "采购价,扣率,结算价,结算金额,售价,发票号,发票代码,发票日期,发票金额"

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
    '系统参数
    txt_条码前缀 = 0
    
    '外购入库
    txt_外购资质校验 = 1
    
    '计划
    txt_计划资质校验 = 2
    
    '发放
    txt_病区发料方式 = 3
    
    txt_结存时间模式 = 4
    
    '基础
    txt_结存参数值 = 5
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
    '系统参数
    chk_时价卫材加成率入库 = 0
    '出库控制
    chk_按批次移库卫材 = 31
    chk_按批次申领卫材 = 1
    chk_填单下可用库存 = 2
    chk_按批次领用卫材 = 3
    
    chk_分段加成入库 = 4
    chk_不严格控制指导价 = 5
    chk_时价卫材按扣前加成销售 = 6
    chk_允许向发料部门领用 = 7
    chk_时价卫材直接确定售价 = 8
    chk_外购入库需要核查 = 9
    chk_时价卫材入库取上次售价 = 10
    
    chk_项目执行前先收费或审核 = 26
    chk_未收费的门诊划价处方发料 = 27
    chk_未审核的记账处方发料 = 28
    
    chk_门诊卫材自动发料 = 29
        
    '外购入库
    chk_允许修改采购限价 = 11
    chk_招标卫材可选择非中标单位入库 = 12
    chk_高值卫材必须填写详细信息 = 13
    chk_生产日期效期检查 = 14
    
    chk_分批卫材批号产地控制 = 34
    
    '移库
    chk_申请冲销 = 16
    
    '领用
    chk_跟踪在用 = 17
    chk_领用审核流程 = 18
    
    '盘点
    chk_盘点存储库房限制 = 19
    
    '调价
    chk_时价卫材按批次调价 = 20
    
    '发放
    chk_自动销账 = 21
    chk_缺料检查 = 22
    chk_领料人签名 = 23
    chk_发料时汇总退料销帐记录 = 24
    chk_卫材医嘱按发生时间过滤 = 25
End Enum

Private Enum constVSF
    vsf_外购资质校验 = 0
    vsf_计划资质校验 = 1
End Enum

Private Enum m库房对照
    mint科室id = 0
    mint发料部门 = 1
    mint库房id = 2
    mint卫材仓库 = 3
    mint虚拟库房id
    mint虚拟库房
    mint启用
    mintCount = 7
End Enum

Private Enum m库房检查
    mintid = 0
    mint编码
    mint名称
    mint检查方式
    minCheck
    mintCount = 5
End Enum
Private Function Get病区发料方式() As String
    Dim n As Integer
    Dim str病区发料 As String
    
    '病区发药
    If chk病区.Value = 0 Then
        str病区发料 = ""
    Else
        For n = 0 To chkDeptType.Count - 1
            If chkDeptType(n).Value = 0 Then
                str病区发料 = IIf(str病区发料 = "", "", str病区发料 & ",") & chkDeptType(n).Caption
            End If
        Next
        If str病区发料 = "" Then
            str病区发料 = "临床,护理,检查,检验,手术,治疗,营养"
        End If
    End If
    
    Get病区发料方式 = str病区发料
End Function

Private Sub Save库存检查()
    Dim i As Integer
    
    '保存库房检查
    gstrSQL = ""
    With vsf库房检查
        For i = 1 To .Rows - 1
            gstrSQL = gstrSQL & .TextMatrix(i, m库房检查.mintid) & "," & Switch(.TextMatrix(i, m库房检查.mint检查方式) = "0-不检查", "0", .TextMatrix(i, m库房检查.mint检查方式) = "1-检查，不足提醒", "1", .TextMatrix(i, m库房检查.mint检查方式) = "2-检查，不足禁止", "2") & ","
        Next
    End With

    gstrSQL = "Zl_材料出库检查_insert('" & gstrSQL & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
End Sub

Private Sub Save流向控制()
    Dim str流向  As String
    Dim i As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim str所在库房id As String
    Dim str对方库房id As String
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim bln次数 As Boolean
    Dim arrSQL  As Variant
    
    arrSQL = Array()
    With vsf流向
        For lngRow = 1 To .Rows - 1
            str流向 = Left(.TextMatrix(lngRow, .ColIndex("流向")), 1)
            If str流向 = "" Then str流向 = "3"
            
            str所在库房id = ""
            str对方库房id = ""
            
            If .TextMatrix(lngRow, .ColIndex("所在库房id")) = "" And lngRow <> .Rows - 1 Then
                gstrSQL = "select id from 部门表 where 编码=[1]"
                strID = Mid(.TextMatrix(lngRow, .ColIndex("所在库房")), 1, InStr(1, .TextMatrix(lngRow, .ColIndex("所在库房")), "-") - 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID)
                If rsTemp.RecordCount > 0 Then
                    str所在库房id = rsTemp!Id
                End If
            Else
                str所在库房id = .TextMatrix(lngRow, .ColIndex("所在库房id"))
            End If
            
            If .TextMatrix(lngRow, .ColIndex("对方库房id")) = "" And lngRow <> .Rows - 1 Then
                strID = Mid(.TextMatrix(lngRow, .ColIndex("对方库房")), 1, InStr(1, .TextMatrix(lngRow, .ColIndex("对方库房")), "-") - 1)
                gstrSQL = "select id from 部门表 where 编码=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "对方库房查询", strID)
                If rsTemp.RecordCount > 0 Then
                    str对方库房id = rsTemp!Id
                End If
            Else
                str对方库房id = .TextMatrix(lngRow, .ColIndex("对方库房id"))
            End If
            If str所在库房id <> "" Or str对方库房id <> "" Then
                If LenB(StrConv(strTemp & str所在库房id & "," & str对方库房id & "," & str流向 & ",", vbFromUnicode)) >= 4000 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strTemp
                    strTemp = str所在库房id & "," & str对方库房id & "," & str流向 & ","
                    bln次数 = True
                Else
                    strTemp = strTemp & str所在库房id & "," & str对方库房id & "," & str流向 & ","
                End If
            End If
        Next
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTemp
    End With
    
    For i = 0 To UBound(arrSQL)
        If bln次数 = True Then
            If i = 0 Then
                Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "删除调价记录")
            Else
                Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',1" & ")", "删除调价记录")
            End If
        Else
            Call zlDatabase.ExecuteProcedure("zl_材料流向控制_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "删除调价记录")
        End If
    Next
End Sub

Private Sub Save虚拟库房对照()
    Dim strTemp As String
    Dim i As Integer
    Dim str科室id As String
    
    '保存虚拟库房对照
    With vsf对照
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, m库房对照.mint科室id)) > 0 And Val(.TextMatrix(i, m库房对照.mint库房id)) > 0 And Val(.TextMatrix(i, m库房对照.mint虚拟库房id)) > 0 Then
                If InStr(1, "," & str科室id & ",", "," & Val(.TextMatrix(i, m库房对照.mint科室id)) & ",") = 0 Then
                    str科室id = IIf(str科室id = "", "", str科室id & ",") & .TextMatrix(i, m库房对照.mint科室id)
                    strTemp = IIf(strTemp = "", "", strTemp & "|") & .TextMatrix(i, m库房对照.mint科室id) & "," & .TextMatrix(i, m库房对照.mint库房id) & "," & .TextMatrix(i, m库房对照.mint虚拟库房id)
                End If
            Else
                If Val(.TextMatrix(i, m库房对照.mint科室id)) = 0 Or Val(.TextMatrix(i, m库房对照.mint库房id)) = 0 Or Val(.TextMatrix(i, m库房对照.mint虚拟库房id)) = 0 Then
                    If Not (.TextMatrix(i, m库房对照.mint发料部门) = "" And .TextMatrix(i, m库房对照.mint卫材仓库) = "" And .TextMatrix(i, m库房对照.mint虚拟库房) = "") Then
                        MsgBox "【虚拟库房对照】第【" & i & "】行没有设置正确的部门或库房，该行保存失败！", vbInformation, gstrSysName
                    End If
                End If
            End If
        Next
    End With

    gstrSQL = "Zl_虚拟库房对照_Update('" & strTemp & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
End Sub

Private Sub Set病区发料方式(ByVal str病区发料 As String)
    Dim BlnSelect As Boolean
    Dim strArr As Variant
    Dim i As Integer
    Dim n As Integer
    
    BlnSelect = False
    If str病区发料 = "" Then
        BlnSelect = False
    ElseIf str病区发料 = "临床,护理,检查,检验,手术,治疗,营养" Then
        BlnSelect = True
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
    Else
        str病区发料 = str病区发料 & ","
        strArr = Split(str病区发料, ",")
        
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
        chk病区.Value = 1
        chk病区.Tag = 1
    Else
        chk病区.Value = 0
        chk病区.Tag = 0
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
    
    '至少要选择一个
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
    
    txt(txt_病区发料方式).Text = Get病区发料方式
End Sub

Private Sub chk病区_Click()
    Dim n As Integer

    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = (chk病区.Value = 1)
        If chk病区.Tag = "0" Then
            chkDeptType(n).Value = 1
        End If
    Next
    
    txt(txt_病区发料方式).Text = Get病区发料方式
End Sub

Private Sub chk应用范围_Click(Index As Integer)
    Dim obj应用范围 As CheckBox
    Dim strValue As String
    
    If mblnLoad = False Then Exit Sub
    
    If chk应用范围(Index).Value <> Val(chk应用范围(Index).Tag) Then
        chk应用范围(Index).ForeColor = &HC0&             '修改后用朱红色前景色标识
    Else
        chk应用范围(Index).ForeColor = &H0&
    End If

    For Each obj应用范围 In chk应用范围
        strValue = IIf(strValue = "", "", strValue) & obj应用范围.Value
    Next
    Call SetParChange(chk应用范围, 0, mrsPar, True, strValue)
End Sub

Private Sub chk应用范围_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(chk应用范围, 0, mrsPar, "", chk应用范围(Index))
End Sub

Private Sub CmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
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
    
    mblnOk = False
    
    '窗口大小：13000,8385
    Me.Width = 13000
    Me.Height = 8385
    
    For Each objPic In picPar
        Set objPic.Container = Me
    Next
    
    tabDesign.Visible = False
    
    strCategory = "参数设置,基础项目"
    '图标编号,TaskPanelItem的ID(同时也是参数容器Picture控件数组号),TaskPanelItem的标题;......
    marrFunc(0) = "104,4,卫材通用设置;100,0,卫材目录管理;101,1,卫材入出管理;102,2,卫材在库管理;103,3,卫材发放管理"
    
    '二级分类Pickture索引从10开始排
    marrFunc(1) = "110,10,卫材流向控制;111,11,卫材库存检查;112,12,虚拟库房对照;113,13,卫材录入精度;114,14,单据环节控制"
    
    '1.初始化快捷面板的一级分类列表,缺省选中第一个
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.初始化任务面板的二级分类列表,缺省选中第一个
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    
    Me.Tag = "初始成功"
    
    mblnLoad = True
End Sub

Private Sub opt编码模式_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt编码模式, Index, mrsPar)
    End If
End Sub

Private Sub opt编码模式_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt编码模式, Index, mrsPar)
End Sub

Private Sub opt采购价_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt采购价, Index, mrsPar)
    End If
End Sub

Private Sub opt采购价_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt采购价, Index, mrsPar)
End Sub

Private Sub opt产地_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt产地, Index, mrsPar)
    End If
End Sub

Private Sub opt产地_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt产地, Index, mrsPar)
End Sub
Private Sub optBarcode_Click(Index As Integer)
    If Not Me.Visible Then Exit Sub
    Call SetParChange(optBarcode, Index, mrsPar)
End Sub
Private Sub optBarcode_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(optBarcode, Index, mrsPar)
End Sub

Private Sub opt定价单位_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt定价单位, Index, mrsPar)
    End If
End Sub

Private Sub opt定价单位_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt定价单位, Index, mrsPar)
End Sub

Private Sub opt分批属性_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt分批属性, Index, mrsPar)
    End If
End Sub

Private Sub opt分批属性_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt分批属性, Index, mrsPar)
End Sub

Private Sub opt计划资质校验_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_计划资质校验, mrsPar, True, Get供应商资质校验(vsf_计划资质校验))
    End If
    
    fra资质校验(vsf_计划资质校验).ForeColor = txt(txt_计划资质校验).ForeColor
End Sub

Private Sub opt结存方式_Click(Index As Integer)
    Dim strValue As String
    
    If opt结存方式(0).Value = True Then
        opt结存时间模式(0).Enabled = False
        opt结存时间模式(1).Enabled = False
        txt(txt_结存时间模式).Enabled = False
        
        '手工结存参数值为-1
        strValue = "-1"
    Else
        opt结存时间模式(0).Enabled = True
        opt结存时间模式(1).Enabled = True
        txt(txt_结存时间模式).Enabled = opt结存时间模式(1).Value
        
        strValue = IIf(opt结存时间模式(0).Value, 0, Val(txt(txt_结存时间模式).Text))
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

Private Sub opt结存时间模式_Click(Index As Integer)
    Dim strValue As String
    
    txt(txt_结存时间模式).Enabled = opt结存时间模式(1).Value
    
    If Me.Visible Then
        strValue = IIf(opt结存时间模式(0).Value, 0, Val(txt(txt_结存时间模式).Text))
        Call SetParChange(txt, txt_结存参数值, mrsPar, True, strValue)
        
        opt结存方式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存方式(1).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(0).ForeColor = txt(txt_结存参数值).ForeColor
        opt结存时间模式(1).ForeColor = txt(txt_结存参数值).ForeColor
        txt(txt_结存时间模式).ForeColor = opt结存时间模式(1).ForeColor
    End If
End Sub

Private Sub opt外购资质校验_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(txt, txt_外购资质校验, mrsPar, True, Get供应商资质校验(vsf_外购资质校验))
    End If
    
    fra资质校验(vsf_外购资质校验).ForeColor = txt(txt_外购资质校验).ForeColor
End Sub

Private Sub opt卫材出库算法_Click(Index As Integer)
    If Me.Visible Then
        Call SetParChange(opt卫材出库算法, Index, mrsPar)
    End If
End Sub

Private Sub opt卫材出库算法_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(opt卫材出库算法, Index, mrsPar)
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim objPic As PictureBox
    
    For Each objPic In picPar
        objPic.Visible = (objPic.Index = Item.Id)
    Next
        
    lblLocate(txt_Dept).Visible = (Item.Id = GetFuncID("药房配药控制", marrFunc) Or _
                            Item.Id = GetFuncID("输液配制中心", marrFunc) Or _
                            Item.Id = GetFuncID("药品流向控制", marrFunc) Or _
                            Item.Id = GetFuncID("药品库存检查", marrFunc) Or _
                            Item.Id = GetFuncID("药品计量单位", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
        
        If Item.Id = GetFuncID("输液配制中心", marrFunc) Then
            lblLocate(txt_Dept).Caption = "科室查找(&F)"
        Else
            lblLocate(txt_Dept).Caption = "药房查找(&F)"
        End If
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOK.Left - lblPrompt.Left - 120
    
    mlngPreFind = 1
    
    tplFunc.Tag = Item.Id   '用于获取当前选中的TaskPanelItem
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
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.Id - 1)) 'ID是从1开始的（因为同时为图标序号）,数组是从0开始
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
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).Id = lngId
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
    
    mstrLike = IIf(gstrMatchMethod = "0", "%", "")
    
    '2.初始化界面控件
    Call InitEnv
    
    '加载其他参数设置
    Call Load材料流向
    Call Load库房检查
    Call Load虚拟库房对照
    Call Load药品卫材精度
    Call Load单据环节控制
    
    '3.加载系统参数
    Call LoadPar
    
End Sub

Private Sub Load材料流向()
    '功能:装入材料流向数据
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTemp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    With vsf流向
        .Rows = 1
        .Rows = 2
        
        '首向装入库房
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('卫材库','制剂室','虚拟库房','发料部门') " & _
                   "   and  b.部门ID=a.ID and " & Where撤档时间("A") & _
                   " order by 编码"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        If Not rsTemp.EOF Then
            rsTemp.MoveFirst
            For i = 1 To rsTemp.RecordCount
                strTemp = strTemp & rsTemp!编码 & "-" & rsTemp!名称 & "|"
                rsTemp.MoveNext
            Next
        End If
        .ColComboList(.ColIndex("所在库房")) = strTemp
        .ColComboList(.ColIndex("对方库房")) = strTemp
        .ColComboList(.ColIndex("流向")) = "1-所在库房可流向对方库房|2-对方库房可流向所在库房|3-两库房间可双向流通"
        
        '装入流向控制数据
        gstrSQL = "select A.所在库房ID,A.对方库房ID,A.流向" & _
                ",B.编码 as 所在编码,B.名称 as 所在名称,C.编码 as 对方编码,C.名称 as 对方名称 " & _
                " from 材料流向控制 A,部门表 B,部门表 C " & _
                " where A.所在库房ID= B.ID and A.对方库房ID=C.ID " & _
                "   and (b.撤档时间=to_date('3000-1-1','yyyy-mm-dd') or b.撤档时间 is null) " & _
                " order by b.编码, c.编码 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
                                                          
            .TextMatrix(lngRow, .ColIndex("所在库房")) = IIf(IsNull(rsTemp!所在库房id), "", rsTemp!所在编码 & "-" & rsTemp!所在名称)
            .TextMatrix(lngRow, .ColIndex("所在库房id")) = rsTemp!所在库房id
            .TextMatrix(lngRow, .ColIndex("对方库房")) = IIf(IsNull(rsTemp!对方库房ID), "", rsTemp!对方编码 & "-" & rsTemp!对方名称)
            .TextMatrix(lngRow, .ColIndex("对方库房id")) = rsTemp!对方库房ID
            .TextMatrix(lngRow, .ColIndex("流向")) = Switch(rsTemp("流向") = 1, "1-所在库房可流向对方库房", _
                                            rsTemp("流向") = 2, "2-对方库房可流向所在库房", _
                                                          True, "3-两库房间可双向流通")
            
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

Private Sub Load库房检查()
    '功能：初始化库房
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim objItem As ListItem
    
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.编码, B.名称, NVL(C.检查方式, 0) 检查方式" & vbCrLf & _
        " FROM 部门性质说明 A, 部门表 B, 材料出库检查 C" & vbCrLf & _
        " WHERE A.部门ID = B.ID AND A.部门ID = C.库房ID(+) AND" & vbCrLf & _
        "      A.工作性质 IN" & vbCrLf & _
        "      ('卫材库','制剂室','发料部门','虚拟库房') " & vbCrLf & _
        "     And (b.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd') or b.撤档时间 is null) " & vbCrLf & _
        " GROUP BY B.ID,B.编码, B.名称, NVL(C.检查方式, 0)" & vbCrLf & _
        " ORDER BY B.编码 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    Me.vsf库房检查.Rows = 1
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        vsf库房检查.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            With vsf库房检查
                .TextMatrix(i, .ColIndex("id")) = rsTmp!Id
                .TextMatrix(i, .ColIndex("部门名称")) = rsTmp!名称
                .TextMatrix(i, .ColIndex("编码")) = rsTmp!编码
                .TextMatrix(i, .ColIndex("库房检查方式")) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
                .TextMatrix(i, .ColIndex("check")) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
            End With
            rsTmp.MoveNext
        Next
        
        vsf库房检查.Cell(flexcpBackColor, 1, vsf库房检查.ColIndex("库房检查方式"), vsf库房检查.Rows - 1, vsf库房检查.ColIndex("库房检查方式")) = &HF4F4EA
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load虚拟库房对照()
    '功能:装入卫材虚拟库房对照关系
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    
    With vsf对照
        '取所有发料部门，卫材库，虚拟库房
        mrs对照.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码, b.工作性质 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('卫材库','发料部门','虚拟库房') " & _
                   " and  b.部门ID=a.ID and " & Where撤档时间("A") & " order by 编码"
        zlDatabase.OpenRecordset mrs对照, gstrSQL, Me.Caption
        
        '装入目前的虚拟库房对照关系
        gstrSQL = "Select b.Id As 科室id, b.编码 || '-' || b.名称 As 发料部门, c.Id As 库房id, c.编码 || '-' || c.名称 As 卫材仓库," & _
                  " d.Id As 虚拟库房id,d.编码 || '-' || d.名称 As 虚拟库房 " & _
                  "From 虚拟库房对照 A, 部门表 B, 部门表 C, 部门表 D " & _
                  "Where a.科室id = b.Id And a.库房id = c.Id And a.虚拟库房id = d.Id " & _
                  "  And (b.撤档时间=to_date('3000-1-1', 'yyyy-mm-dd') or b.撤档时间 is null) " & _
                  "Order by b.编码 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("科室id")) = rsTemp!科室id
            .TextMatrix(lngRow, .ColIndex("发料部门")) = rsTemp!发料部门
            .TextMatrix(lngRow, .ColIndex("卫材仓库id")) = rsTemp!库房ID
            .TextMatrix(lngRow, .ColIndex("卫材仓库")) = rsTemp!卫材仓库
            .TextMatrix(lngRow, .ColIndex("虚拟库房id")) = rsTemp!虚拟库房id
            .TextMatrix(lngRow, .ColIndex("虚拟库房")) = rsTemp!虚拟库房
'            .TextMatrix(lngRow, .ColIndex("启用")) = "√"
            
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
'功能：读取并加载参数到界面控件
    Dim strValue As String, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim arrObj As Variant  '数组对象：模块1,参数号1,控件对象1,模块2,参数号2,控件对象2,......
    Dim n As Integer
    Dim obj应用范围 As CheckBox
    
    '读取参数(默认读取系统参数，需要的模块参数单独添加对应的模块号)
    Set rsTmp = GetPar(mrsPar, p卫材目录管理 & "," & _
            p卫材外购管理 & "," & _
            p卫材移库管理 & "," & _
            p卫材领用管理 & "," & _
            p卫材盘点管理 & "," & _
            p卫材申领管理 & "," & _
            p卫材发放管理 & "," & _
            p卫材计划管理 & "," & _
            p卫材调价管理)
    

    '----------------------------------------------------------
    '系统参数
    '1.设置CheckBox类参数
    strTmp = "0:82:" & chk_时价卫材加成率入库 & _
            ",0:280:" & chk_按批次移库卫材 & _
            ",0:83:" & chk_按批次申领卫材 & _
            ",0:95:" & chk_填单下可用库存 & _
            ",0:121:" & chk_分段加成入库 & _
            ",0:123:" & chk_不严格控制指导价 & _
            ",0:127:" & chk_时价卫材按扣前加成销售 & _
            ",0:132:" & chk_允许向发料部门领用 & _
            ",0:136:" & chk_时价卫材直接确定售价 & _
            ",0:140:" & chk_外购入库需要核查 & _
            ",0:163:" & chk_项目执行前先收费或审核 & _
            ",0:171:" & chk_未收费的门诊划价处方发料 & _
            ",0:172:" & chk_未审核的记账处方发料 & _
            ",0:229:" & chk_时价卫材入库取上次售价 & _
            ",0:92:" & chk_门诊卫材自动发料 & _
            ",0:258:" & chk_按批次领用卫材 & _
            ",0:305:" & chk_分批卫材批号产地控制
    Call SetParToControl(strTmp, mrsPar, chk)
    
'    chk(chk_填单下可用库存).Enabled = (Check申领单 And Check移库单 And Check领用单)
'    chk(chk_按批次申领卫材).Enabled = Check申领单
'    chk(chk_按批次移库卫材).Enabled = Check移库单
'    chk(chk_按批次领用卫材).Enabled = Check领用单

    '设置参数关系
    If chk(chk_项目执行前先收费或审核).Value = 1 Then
        chk(chk_未收费的门诊划价处方发料).Enabled = False
        chk(chk_未审核的记账处方发料).Enabled = False
        lbl未收费发药.Caption = "  已启用了门诊一卡通参数“执行前必须先收费或先记帐审核”，则对门诊病人发料时，以下参数无论勾选都将失效。"
    Else
        chk(chk_未收费的门诊划价处方发料).Enabled = True
        chk(chk_未审核的记账处方发料).Enabled = True
        lbl未收费发药.Caption = "  如果启用了门诊一卡通参数“执行前必须先收费或先记帐审核”，则对门诊病人发料时，以下参数将失效。"
    End If
    
    
'    '2.设置ComboBox类参数
'    strTmp = "0:29:" & cbo_定价单位 & _
'            ",0:64:" & cbo_药品单据审核 & _
'            ",0:87:" & cbo_药品编码模式 & _
'            ",0:149:" & cbo_效期显示方式 & _
'            ",0:150:" & cbo_药品出库优先算法
'
'    Call SetParToControl(strTmp, mrsPar, cbo)
'
'
'    '3.设置UpDown类参数
'    strTmp = ""
'    'Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar存储的控件名是txtUD
'
'
    '4.设置TextBox类参数
    strTmp = "0:159:" & txt_条码前缀
    Call SetParToControl(strTmp, mrsPar, txt)
'
'    '5.设置ListBox类参数
''    strTmp = p住院医嘱下达 & ":44:" & lst_输液中心发药病人科室
''    Call SetParToControl(strTmp, mrsPar, lst, 1)
'
    '6.设置OptionButton类参数
    arrObj = Array(0, 88, opt定价单位, _
                 0, 156, opt卫材出库算法, _
                 0, 268, opt产地, _
                 0, 320, optBarcode)
    Call SetParToControl("", mrsPar, arrObj)
    
'    '7.其他系统参数
    rsTmp.Filter = "模块=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!参数值
        Select Case rsTmp!参数号
        
        Case 281   '药品结存时间模式
            If Val(strValue) = -1 Then
                '参数值为-1表示手工结存
                opt结存方式(0).Value = True
                opt结存方式(1).Value = False
                
                opt结存时间模式(0).Enabled = False
                opt结存时间模式(1).Enabled = False
                txt(txt_结存时间模式).Enabled = False
            Else
                '参数值不为-1表示自动结存
                opt结存方式(0).Value = False
                opt结存方式(1).Value = True
                
                If Val(strValue) = 0 Then
                    '参数值为0表示每月最后一天结存
                    opt结存时间模式(0).Value = True
                    opt结存时间模式(1).Value = False
                    txt(txt_结存时间模式).Enabled = False
                Else
                    '参数值大于0小于等于31表示指定日期结存
                    opt结存时间模式(0).Value = False
                    opt结存时间模式(1).Value = True
                    
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
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(txt, txt_结存时间模式, mrsPar)
            
         Case 320 '卫材条码认别:1-必须通过扫码录入或录入条码来认别卫生材料;0-不控制，可以通过简码、编码、条码等录入方式来识别卫生材料
            
            If Val(strValue) = 0 Then
                optBarcode(0).Value = True: optBarcode(1).Value = False
            Else
                optBarcode(0).Value = False: optBarcode(1).Value = True
            End If
        End Select
        rsTmp.MoveNext
    Loop

    '----------------------------------------------------------
    '8.其他模块参数
    '卫材目录管理 = 1711
    '设置ComboBox类参数
    strTmp = p卫材目录管理 & ":收入项目对应:" & cbo_收入项目 & _
            ",0:63:" & cbo_住院卫材自动发料
    Call SetParToControl(strTmp, mrsPar, cbo, 1)
    
    '设置OptionButton类参数
    arrObj = Array(p卫材目录管理, "编码递增模式", opt编码模式, _
                    p卫材目录管理, "卫材分批属性自动设置", opt分批属性)
    Call SetParToControl("", mrsPar, arrObj)
    
    '其他参数
    '特殊参数处理：参数值对应多个控件(组)，先调用公共方法记录控件名称，界面控件显示单独处理
    strTmp = p卫材目录管理 & ":允许应用于的范围:0"
    Call SetParToControl(strTmp, mrsPar, chk应用范围)

    rsTmp.Filter = "模块=" & p卫材目录管理 & " And 参数名='允许应用于的范围'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值, "111")
    If strValue <> "" Then
        For n = 0 To chk应用范围.Count - 1
            chk应用范围(n).Value = Mid(strValue, n + 1, 1)
            chk应用范围(n).Tag = Mid(strValue, n + 1, 1)
        Next
    End If
    
    
    '----------------------------------------------------------
    '入库
    '卫材外购管理 = 1712
    '设置CheckBox类参数
    strTmp = p卫材外购管理 & ":修改采购限价:" & chk_允许修改采购限价 & _
            "," & p卫材外购管理 & ":招标卫材可选择非中标单位入库:" & chk_招标卫材可选择非中标单位入库 & _
            "," & p卫材外购管理 & ":高值卫材必须填写详细信息:" & chk_高值卫材必须填写详细信息 & _
            "," & p卫材外购管理 & ":生产日期效期检查:" & chk_生产日期效期检查
    Call SetParToControl(strTmp, mrsPar, chk)

    '设置OptionButton类参数
    arrObj = Array(p卫材外购管理, "入库单价超中标单价", opt采购价)
    Call SetParToControl("", mrsPar, arrObj)
    
    '特殊处理
    '该参数实际用表格和其他控件显示，特别的用额外文本控件记录原始值，单独处理界面显示
    strTmp = p卫材外购管理 & ":资质校验:" & txt_外购资质校验
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "模块=" & p卫材外购管理 & " And 参数名='资质校验'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    Call Load供应商资质校验(vsf_外购资质校验, strValue)
    
    
    '----------------------------------------------------------
    '出库
    '卫材移库管理 = 1716
    '卫材领用管理 = 1717
    '设置CheckBox类参数
    strTmp = p卫材移库管理 & ":冲销申请:" & chk_申请冲销 & _
            "," & p卫材领用管理 & ":跟踪在用:" & chk_跟踪在用 & _
            "," & p卫材领用管理 & ":审核流程:" & chk_领用审核流程
    Call SetParToControl(strTmp, mrsPar, chk)

    
    '----------------------------------------------------------
    '卫材盘点管理 = 1719
    '设置CheckBox类参数
    strTmp = p卫材盘点管理 & ":存储库房:" & chk_盘点存储库房限制
    Call SetParToControl(strTmp, mrsPar, chk)
    
    
    '----------------------------------------------------------
    '卫材计划管理 = 1724
    '特殊处理
    '该参数实际用表格和其他控件显示，特别的用额外文本控件记录原始值，单独处理界面显示
    strTmp = p卫材计划管理 & ":资质校验:" & txt_计划资质校验
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "模块=" & p卫材计划管理 & " And 参数名='资质校验'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    Call Load供应商资质校验(vsf_计划资质校验, strValue)

    
    '----------------------------------------------------------
    '卫材调价管理 = 1726
    '设置CheckBox类参数
    strTmp = p卫材调价管理 & ":时价卫材按批次调价:" & chk_时价卫材按批次调价
    Call SetParToControl(strTmp, mrsPar, chk)


    '----------------------------------------------------------
    '卫材发放管理 = 1723
    '设置CheckBox类参数
    strTmp = p卫材发放管理 & ":自动销帐:" & chk_自动销账 & _
        "," & p卫材发放管理 & ":缺料检查:" & chk_缺料检查 & _
        "," & p卫材发放管理 & ":领料人签名:" & chk_领料人签名 & _
        "," & p卫材发放管理 & ":发料时汇总退料销帐记录:" & chk_发料时汇总退料销帐记录 & _
        "," & p卫材发放管理 & ":卫材医嘱按发生时间过滤:" & chk_卫材医嘱按发生时间过滤
    Call SetParToControl(strTmp, mrsPar, chk)

    '该参数实际用表格和其他控件显示，特别的用额外文本控件记录原始值，单独处理界面显示
    strTmp = p卫材发放管理 & ":病区发料方式:" & txt_病区发料方式
    Call SetParToControl(strTmp, mrsPar, txt)

    rsTmp.Filter = "模块=" & p卫材发放管理 & " And 参数名='病区发料方式'"
    If Not rsTmp.EOF Then strValue = NVL(rsTmp!参数值)
    Call Set病区发料方式(strValue)
      
    
    '----------------------------------------------------------
    '参数关系控制
    If chk(chk_填单下可用库存).Value = 1 Then
        chk(chk_按批次申领卫材).Value = 1
        chk(chk_按批次申领卫材).Enabled = False
        
        chk(chk_按批次领用卫材).Value = 1
        chk(chk_按批次领用卫材).Enabled = False
        
        chk(chk_按批次移库卫材).Value = 1
        chk(chk_按批次移库卫材).Enabled = False
    End If
    
    If chk(chk_时价卫材加成率入库).Value = 1 Then
        chk(chk_分段加成入库).Value = 0
        chk(chk_分段加成入库).Enabled = False
        chk(chk_时价卫材入库取上次售价).Value = 0
        chk(chk_时价卫材入库取上次售价).Enabled = False
    ElseIf chk(chk_分段加成入库).Value = 1 Then
        chk(chk_时价卫材加成率入库).Value = 0
        chk(chk_时价卫材加成率入库).Enabled = False
        chk(chk_时价卫材入库取上次售价).Value = 0
        chk(chk_时价卫材入库取上次售价).Enabled = False
    ElseIf chk(chk_时价卫材入库取上次售价).Value = 1 Then
        chk(chk_分段加成入库).Value = 0
        chk(chk_分段加成入库).Enabled = False
        chk(chk_时价卫材加成率入库).Value = 0
        chk(chk_时价卫材加成率入库).Enabled = False
    End If
        
End Sub

Private Function Get供应商资质校验(ByVal intType As Integer) As String
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    blnAllUnCheck = True
    
    '保存资质校验项目和方式，格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    With vsfCheck(intType)
        For i = 1 To .Rows - 1
            strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("类别")) & "," & .TextMatrix(i, .ColIndex("校验项目")) & "," & _
                IIf(.TextMatrix(i, .ColIndex("校验")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("校验")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If intType = 0 Then
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf opt外购资质校验(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
    Else
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf opt计划资质校验(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
    End If
        
    Get供应商资质校验 = strCheck
End Function
Private Sub Load供应商资质校验(ByVal intType As Integer, ByVal strParaValue As String)
    Dim i As Integer
    Dim n As Integer
    Dim intCheckType As Integer
    Dim arrColumn
    
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....

    If strParaValue <> "" Then
        If InStr(1, strParaValue, "|") > 0 Then
            '校验方式：0-不检查；1－提醒；2－禁止
            intCheckType = Val(Mid(strParaValue, 1, InStr(1, strParaValue, "|") - 1))
            
            If intType = 0 Then
                If intCheckType = 2 Then
                    opt外购资质校验(0).Value = True
                ElseIf intCheckType = 1 Then
                    opt外购资质校验(1).Value = True
                End If
            Else
                If intCheckType = 2 Then
                    opt计划资质校验(0).Value = True
                ElseIf intCheckType = 1 Then
                    opt计划资质校验(1).Value = True
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
    
    On Error GoTo ErrHandle
    
    '卫材目录
    gstrSQL = "Select ID,编码||'-'||名称 名称 From 收入项目 Where 末级=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "InitEnv")
    With rsData
        Do While Not .EOF
            cbo(cbo_收入项目).AddItem !名称
            cbo(cbo_收入项目).ItemData(cbo(cbo_收入项目).NewIndex) = !Id
            .MoveNext
        Loop
    End With
    
    '住院卫材自动发料
    With cbo(cbo_住院卫材自动发料)
        .Clear
        .AddItem "0-不自动发料"
        .ItemData(.NewIndex) = 0
        
        .AddItem "1-自动发料"
        .ItemData(.NewIndex) = 1
        
        .AddItem "2-本科室开单自动发料"
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
    mrsPar.Filter = "修改状态=1"
    If mrsPar.RecordCount > 0 Then
    
'        Or Bill(bill_药品库房流向).Tag = "已修改" Or Bill(bill_药品领用流向).Tag = "已修改" _
'        Or lvw库存检查.Tag = "已修改" Or msf库房计量单位.Tag = "已修改" Or Bill药房配药控制.Tag = "已修改" _
'        Or Bill药品卫材精度.Tag = "已修改" Or vsf单据环节控制.Tag = "已修改" Then
        
        If Not mblnOk Then
            If MsgBox("你已修改部分参数，如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    
    Set mrsPar = Nothing
    Set mrs对照 = Nothing
    
    mblnLoad = False
End Sub

Private Sub cmdOk_Click()
    Dim obj应用范围 As CheckBox
    Dim strValue As String
    
    If ValidateData() = False Then Exit Sub
    
    mblnOk = True
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    
    Call zlDatabase.ClearParaCache
    
    '保存其他参数
    Call Save流向控制
    Call Save库存检查
    Call Save虚拟库房对照
    Call Save药品卫材精度
    Call Save单据环节控制
    
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
    Case txt_结存时间模式
        If Val(txt(Index).Text) < 0 Or Val(txt(Index).Text) > 31 Then
            txt(Index).Text = 25
        End If
    End Select
    
    If Me.Visible Then
        Call SetParChange(txt, Index, mrsPar)
    End If
    
    If Index = txt_病区发料方式 Then
        chk病区.ForeColor = txt(txt_病区发料方式).ForeColor
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
        Case txt_结存时间模式
            Select Case KeyAscii
            Case vbKeyBack, vbKeyEscape, 3, 22  '小数点
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
'            If Bill药房配药控制.Visible Then
'                Call LocateDept(strFind, Bill药房配药控制, 0)
'
'            ElseIf Bill(bill_药品领用流向).Visible Then
'                If lblLocate(txt_Dept).Tag = bill_药品库房流向 Or lblLocate(txt_Dept).Tag = "" Then
'                    Call LocateDept(strFind, Bill(bill_药品库房流向), IIf(Bill(bill_药品库房流向).Col = 0, 0, 1))
'                Else
'                    Call LocateDept(strFind, Bill(bill_药品领用流向), Bill(bill_药品领用流向).Col)
'                End If
'
'            ElseIf lvw库存检查.Visible Then
'                Call LocateDept(strFind, lvw库存检查, 1)
'
'            ElseIf msf库房计量单位.Visible Then
'                Call LocateDept(strFind, msf库房计量单位, 0)
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
                If .ListItems(i).ListSubItems(lngCol).Text Like IIf(mstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst_输液中心发药病人科室
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
            MsgBox "没有找到匹配的，请检查输入的内容。", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        Else
            MsgBox "全部找完了，后面没有了。", vbInformation, Me.Caption
            mlngPreFind = 1
        End If
    End If
End Sub

Private Function ValidateData() As Boolean
    Dim lngRow As Long
    Dim j As Long
    
    With vsf流向
        For lngRow = 1 To .Rows - 1
            If (.TextMatrix(lngRow, .ColIndex("所在库房")) = "" Or .TextMatrix(lngRow, .ColIndex("对方库房")) = "" Or .TextMatrix(lngRow, .ColIndex("流向")) = "") And lngRow <> .Rows - 1 Then
                MsgBox "流向设置中，第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("所在库房")

                Exit Function
            End If

            If .TextMatrix(lngRow, .ColIndex("所在库房")) = .TextMatrix(lngRow, .ColIndex("对方库房")) And lngRow <> .Rows - 1 Then
                MsgBox "流向设置中，第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("所在库房")

                Exit Function
            End If

            For j = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("所在库房")) = .TextMatrix(j, .ColIndex("所在库房")) And .TextMatrix(lngRow, .ColIndex("对方库房")) = .TextMatrix(j, .ColIndex("对方库房")) And lngRow <> j Then
                    MsgBox "流向设置中，第" & lngRow & "行与第" & j & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = .ColIndex("所在库房")

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
    Case chk_填单下可用库存
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_填单下可用库存).Value = Val(GetParOriginalValue(chk, chk_填单下可用库存, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = (Check申领单 And Check移库单 And Check领用单)
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的申领，移库，或领用单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_填单下可用库存).Value = Val(GetParOriginalValue(chk, chk_填单下可用库存, mrsPar))
        End If
        
        If chk(Index).Value = 1 Then
            chk(chk_按批次申领卫材).Enabled = False
            If chk(chk_按批次申领卫材).Value = 0 Then chk(chk_按批次申领卫材).Value = 1
            
            chk(chk_按批次领用卫材).Enabled = False
            If chk(chk_按批次领用卫材).Value = 0 Then chk(chk_按批次领用卫材).Value = 1
            
            chk(chk_按批次移库卫材).Enabled = False
            If chk(chk_按批次移库卫材).Value = 0 Then chk(chk_按批次移库卫材).Value = 1
            Call SetPrompt(lblPrompt, "填单时下可用库存时必须设置为[申领时明确卫材批次][领用时明确卫材批次]")
        Else
            chk(chk_按批次移库卫材).Enabled = True
            chk(chk_按批次申领卫材).Enabled = True
            chk(chk_按批次领用卫材).Enabled = True
        End If
        
    Case chk_按批次申领卫材
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_按批次申领卫材).Value = Val(GetParOriginalValue(chk, chk_按批次申领卫材, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check申领单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的申领单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_按批次申领卫材).Value = Val(GetParOriginalValue(chk, chk_按批次申领卫材, mrsPar))
        End If
     Case chk_按批次移库卫材
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_按批次移库卫材).Value = Val(GetParOriginalValue(chk, chk_按批次移库卫材, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check移库单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的移库单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_按批次移库卫材).Value = Val(GetParOriginalValue(chk, chk_按批次移库卫材, mrsPar))
        End If
    Case chk_按批次领用卫材
        '窗口加载时不运行下面语句
        If Me.Visible = False Then Exit Sub
        
        '当前选择的等于原始参数值时不进行下面语句，否则会死循环
        If chk(chk_按批次领用卫材).Value = Val(GetParOriginalValue(chk, chk_按批次领用卫材, mrsPar)) Then Exit Sub
        
        On Error GoTo ErrHandle
        
        DoEvents
        zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
        blnResulte = Check领用单
        DoEvents
        zlCommFun.StopFlash
                
        If blnResulte = False Then
            MsgBox "存在近期未审核的领用单，不能改变此参数！", vbInformation, gstrSysName
            chk(chk_按批次领用卫材).Value = Val(GetParOriginalValue(chk, chk_按批次领用卫材, mrsPar))
        End If
'    chk(chk_按批次申领卫材).Enabled = Check申领单
'    chk(chk_按批次移库卫材).Enabled = Check移库单
'    chk(chk_按批次领用卫材).Enabled = Check领用单
        
    Case chk_时价卫材入库取上次售价
        If chk(Index).Value = 1 Then
            chk(chk_时价卫材加成率入库).Enabled = False
            If chk(chk_时价卫材加成率入库).Value = 1 Then chk(chk_时价卫材加成率入库).Value = 0: Call SetPrompt(lblPrompt, "设置了入库取上次售价方式后就不能选择[按加成率计算售价]方式了")
            
            chk(chk_分段加成入库).Enabled = False
            If chk(chk_分段加成入库).Value = 1 Then chk(chk_分段加成入库).Value = 0: Call SetPrompt(lblPrompt, "设置了入库取上次售价方式后就不能选择[按分段加成计算售价]方式了")
        Else
            chk(chk_时价卫材加成率入库).Enabled = True
            chk(chk_分段加成入库).Enabled = True
        End If
    Case chk_时价卫材加成率入库
        If chk(Index).Value = 1 Then
            chk(chk_时价卫材入库取上次售价).Enabled = False
            If chk(chk_时价卫材入库取上次售价).Value = 1 Then chk(chk_时价卫材入库取上次售价).Value = 0: Call SetPrompt(lblPrompt, "设置了入库按加成率计算售价方式后就不能选择[取上次售价]方式了")
            
            chk(chk_分段加成入库).Enabled = False
            If chk(chk_分段加成入库).Value = 1 Then chk(chk_分段加成入库).Value = 0: Call SetPrompt(lblPrompt, "设置了入库按加成率计算售价方式后就不能选择[按分段加成计算售价]方式了")
        Else
            chk(chk_时价卫材入库取上次售价).Enabled = True
            chk(chk_分段加成入库).Enabled = True
        End If
    Case chk_分段加成入库
        If chk(Index).Value = 1 Then
            chk(chk_时价卫材入库取上次售价).Enabled = False
            If chk(chk_时价卫材入库取上次售价).Value = 1 Then chk(chk_时价卫材入库取上次售价).Value = 0: Call SetPrompt(lblPrompt, "设置了按分段加成计算售价方式后就不能选择[取上次售价]方式了")
            
            chk(chk_时价卫材加成率入库).Enabled = False
            If chk(chk_时价卫材加成率入库).Value = 1 Then chk(chk_时价卫材加成率入库).Value = 0: Call SetPrompt(lblPrompt, "设置了按分段加成计算售价方式后就不能选择[按加成率计算售价]方式了")
        Else
            chk(chk_时价卫材入库取上次售价).Enabled = True
            chk(chk_时价卫材加成率入库).Enabled = True
        End If
    Case chk_申请冲销
        '当变为不需要申请时，要检查是否有未审核的冲销申请单，如果有则不能改变
        Dim rsTemp As ADODB.Recordset
        
        If chk(Index).Value = 0 And mblnLoad = True Then
            If MsgBox("即将检查是否存在未审核的冲销申请单，可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '该功能是10.34版本新增，增加一个条件填制日期范围，避免全表扫描，因此考虑从34版本修改日期开始
               gstrSQL = "Select 1 From 未审药品记录 A " & _
                    " Where a.单据 = 19 And a.填制日期 Between To_Date('2014/2/20 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Exists " & _
                    " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Mod(b.记录状态, 3) = 2) And Rownum < 2"
                
                
                DoEvents
                zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的冲销申请单")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "存在未审核的冲销申请单，不能改变此参数！", vbInformation, gstrSysName
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
        If .Col <> .ColIndex("校验") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            .TextMatrix(.Row, .Col) = "√"
        End If
    End With
    
    If Me.Visible Then
        Call SetParChange(txt, IIf(Index = 0, txt_外购资质校验, txt_计划资质校验), mrsPar, True, Get供应商资质校验(Index))
    End If
       
    fra资质校验(Index).ForeColor = txt(IIf(Index = 0, txt_外购资质校验, txt_计划资质校验)).ForeColor
End Sub

Private Sub vsfCheck_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetParTip(txt, IIf(Index = 0, txt_外购资质校验, txt_计划资质校验), mrsPar, "", vsfCheck(Index))
End Sub

Private Sub vsf对照_ChangeEdit()
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim str名称 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
    
    If InStr(1, vsf对照.EditText, "-") <= 0 Then Exit Sub
    strID = Mid(vsf对照.EditText, 1, InStr(1, vsf对照.EditText, "-") - 1)
    str名称 = Mid(vsf对照.EditText, InStr(1, vsf对照.EditText, "-") + 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门查询", strID, str名称)
    If rsTemp.RecordCount > 0 Then
        With vsf对照
            If .Col = m库房对照.mint发料部门 Then
                .TextMatrix(.Row, m库房对照.mint科室id) = rsTemp!Id
            ElseIf .Col = m库房对照.mint卫材仓库 Then
                .TextMatrix(.Row, m库房对照.mint库房id) = rsTemp!Id
            ElseIf .Col = m库房对照.mint虚拟库房 Then
                .TextMatrix(.Row, m库房对照.mint虚拟库房id) = rsTemp!Id
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

'Private Sub vsf对照_DblClick()
'    With vsf对照
'        If .Col = m库房对照.mint启用 Then
'            If .TextMatrix(.Row, m库房对照.mint启用) = "" Then
'                .TextMatrix(.Row, m库房对照.mint启用) = "√"
'            Else
'                .TextMatrix(.Row, m库房对照.mint启用) = ""
'            End If
'        End If
'    End With
'End Sub

Private Sub vsf对照_EnterCell()
    Dim strTemp As String
    
    With vsf对照
        If .Col = m库房对照.mint启用 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = 1 Then
            mrs对照.Filter = "工作性质='发料部门'"
        ElseIf .Col = 3 Then
            mrs对照.Filter = "工作性质='卫材库'"
        ElseIf .Col = 5 Then
            mrs对照.Filter = "工作性质='虚拟库房'"
        End If
        
'        .Clear
        strTemp = ""
        Do While Not mrs对照.EOF
            strTemp = strTemp & mrs对照("编码") & "-" & mrs对照("名称") & "|"
'            .AddItem mrs对照("编码") & "-" & mrs对照("名称")
'            .ItemData(.NewIndex) = mrs对照("ID")
            mrs对照.MoveNext
        Loop
        .ColComboList(.Col) = strTemp
    End With
End Sub

Private Sub vsf对照_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsf对照
            If .Rows > 1 Then
                If .TextMatrix(.Row, m库房对照.mint发料部门) <> "" Or .TextMatrix(.Row, m库房对照.mint卫材仓库) <> "" Or .TextMatrix(.Row, m库房对照.mint虚拟库房) <> "" Then
                    If MsgBox("是否确定删除这行？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
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

Private Sub vsf对照_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf对照
            If .Col = m库房对照.mint启用 - 1 And .TextMatrix(.Row, m库房对照.mint发料部门) <> "" And _
                .TextMatrix(.Row, m库房对照.mint卫材仓库) <> "" And .TextMatrix(.Row, m库房对照.mint虚拟库房) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m库房对照.mint启用 - 1 And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 2
            End If
        End With
    End If
End Sub

Private Sub vsf库房检查_DblClick()
    With vsf库房检查
        If .Col = m库房检查.mint检查方式 Then
            Select Case .TextMatrix(.Row, m库房检查.mint检查方式)
                Case "0-不检查"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "1-检查，不足提醒"
                Case "1-检查，不足提醒"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "2-检查，不足禁止"
                Case "2-检查，不足禁止"
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "0-不检查"
                Case Else
                    .TextMatrix(.Row, m库房检查.mint检查方式) = "0-不检查"
            End Select
            
            If .TextMatrix(.Row, m库房检查.mint检查方式) <> .TextMatrix(.Row, m库房检查.minCheck) Then
                .Cell(flexcpForeColor, .Row, m库房检查.mint检查方式, .Row, m库房检查.mint检查方式) = vbRed
            Else
                .Cell(flexcpForeColor, .Row, m库房检查.mint检查方式, .Row, m库房检查.mint检查方式) = vbBlack
            End If
        End If
    End With
End Sub

Private Sub vsf流向_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strID As String
    Dim str名称 As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    
    With vsf流向
        strTemp = .TextMatrix(Row, Col)
        If strTemp <> "" Then
            If Col = .ColIndex("所在库房") Then
                gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str名称 = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID, str名称)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("所在库房id")) = rsTemp!Id
                End If
            ElseIf Col = .ColIndex("对方库房") Then
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str名称 = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                gstrSQL = "select id from 部门表 where 编码=[1] and 名称=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "所在库房查询", strID, str名称)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("对方库房id")) = rsTemp!Id
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

Private Sub vsf流向_DblClick()
    With vsf流向
        If .Col = .ColIndex("流向") Then
            If .MouseRow = 0 Then Exit Sub
            .Editable = flexEDNone
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
            End Select
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsf流向_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsf流向.Rows > 1 Then
        vsf流向.RemoveItem vsf流向.Row
    End If
End Sub

Private Sub vsf流向_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf流向
            If .Col = .ColIndex("流向") And .TextMatrix(.Row, .ColIndex("所在库房")) <> "" And _
                .TextMatrix(.Row, .ColIndex("对方库房")) <> "" And .TextMatrix(.Row, .ColIndex("流向")) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = .ColIndex("所在库房")
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("所在库房")
                End If
            ElseIf .Col < .ColIndex("流向") And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 1
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
    
    On Error GoTo ErrHandle
    '取最大精度
    gstrSQL = "Select 成本价, 零售价, 实际数量,零售金额 From 药品收发记录 Where Rownum <2"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
    intMaxCost = IIf(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIf(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIf(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIf(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

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
            " From 药品卫材精度 where 类别=2 Order By 性质, 类别, 内容, 单位"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
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
                .TextMatrix(n, dig_精度) = IIf(rs!精度 > 4, 4, rs!精度)
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
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save药品卫材精度()
    Dim n As Integer
    Dim strInput As String
       
    On Error GoTo ErrHandle
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
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Bill药品卫材精度_EnterCell(Row As Long, Col As Long)
    With Bill药品卫材精度
        If Col = dig_精度 Then
            .TxtCheck = True
            .TextMask = "123456789"
            .MaxLength = 1
        End If
    End With
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
            
            .TextMatrix(.Row, .Col) = strKey
            .RowData(.Row) = Val(strKey)
            
            .Tag = "已修改"
        End If
    End With
End Sub
Private Sub Load单据环节控制()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo ErrHandle
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
        
        .TextMatrix(1, 0) = "卫材外购"
        .TextMatrix(2, 0) = "卫材外购"
        .TextMatrix(3, 0) = "卫材外购"

        .TextMatrix(1, 1) = "核查"
        .TextMatrix(2, 1) = "审核"
        .TextMatrix(3, 1) = "财务审核"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select 单据,环节,内容 From 单据环节控制 where 单据=15 Order By 单据, 环节"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "单据环节控制")
        
        If Not rsTmp.EOF Then
            For n = 1 To rsTmp.RecordCount
                For m = 2 To intAllItems + 1
                    If InStr(1, "," & rsTmp!内容 & ",", Trim(.TextMatrix(0, m))) > 0 Then
                        Select Case rsTmp!单据
                            Case 单据.卫材外购
                                Select Case rsTmp!环节
                                    Case 环节.核查
                                        .TextMatrix(1, m) = "√"
                                    Case 环节.审核
                                        .TextMatrix(2, m) = "√"
                                    Case 环节.财务审核
                                        .TextMatrix(3, m) = "√"
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

            .TextMatrix(.Row, .Col) = "√"

        End If
        
    End With
End Sub

Private Sub Save单据环节控制()
    Dim n As Integer
    Dim m As Integer
    Dim strInput As String
    Dim int单据 As Integer
    Dim int环节 As Integer
    Dim str内容 As String
    
    On Error GoTo ErrHandle
    With vsf单据环节控制
        If .Tag = "已修改" Then
            For n = 1 To .Rows - 1
                Select Case .TextMatrix(n, 0)
                    Case "卫材外购"
                        int单据 = 单据.卫材外购
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
                        str内容 = str内容 & IIf(str内容 <> "", ",", "") & .TextMatrix(0, m)
                    End If
                Next
                
                If str内容 <> "" Then
                    strInput = strInput & IIf(strInput <> "", ";", "") & int单据 & "," & int环节 & "," & str内容
                End If
            Next
        
            gstrSQL = "Zl_单据环节控制_Update('" & strInput & "'," & 单据.卫材外购 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            .Tag = ""
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check移库单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From 未审药品记录 A " & _
        " Where a.单据 = 19 And a.填制日期 > Sysdate - 90 And Exists " & _
        " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Nvl(b.发药方式,0) <> 1) And Rownum < 2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的移库单")
    
    Check移库单 = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check申领单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From 未审药品记录 A " & _
        " Where a.单据 = 19 And a.填制日期 > Sysdate - 90 And Exists " & _
        " (Select 1 From 药品收发记录 B Where a.收发id = b.Id And Nvl(b.发药方式,0) = 1) And Rownum < 2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的申领单")
    
    Check申领单 = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check领用单() As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From 未审药品记录 Where 单据 = 20 And 填制日期 > Sysdate - 90 And Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的领用单")
    
    Check领用单 = rsTemp.RecordCount = 0
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

