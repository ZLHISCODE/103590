VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmParameters 
   BackColor       =   &H80000005&
   Caption         =   "系统参数管理"
   ClientHeight    =   6288
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10728
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmParameters.frx":0000
   ScaleHeight     =   6288
   ScaleWidth      =   10728
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExp 
      Caption         =   "导出(&E)"
      Height          =   350
      Left            =   7560
      TabIndex        =   21
      Top             =   935
      Width           =   1100
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "导入(&I)"
      Height          =   350
      Left            =   7560
      TabIndex        =   20
      Top             =   560
      Width           =   1100
   End
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   3360
      Width           =   9975
   End
   Begin VB.ComboBox cboParType 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2445
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6600
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":04F9
            Key             =   "本机公共模块"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":0A93
            Key             =   "本机私有模块"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":102D
            Key             =   "公共模块"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":15C7
            Key             =   "私有模块"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":1B61
            Key             =   "私有全局"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":20FB
            Key             =   "公共全局"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":2695
            Key             =   "部门参数"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "改变参数性质(&M)"
      Height          =   350
      Left            =   5640
      TabIndex        =   7
      Top             =   935
      Width           =   1665
   End
   Begin VB.PictureBox picPara 
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   60
      ScaleHeight     =   2028
      ScaleWidth      =   9828
      TabIndex        =   9
      Top             =   1320
      Width           =   9825
      Begin VSFlex8Ctl.VSFlexGrid vsPara 
         Height          =   1890
         Left            =   30
         TabIndex        =   10
         Top             =   90
         Width           =   7470
         _cx             =   13176
         _cy             =   3334
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParameters.frx":8EF7
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
         ExplorerBar     =   7
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
   Begin VB.CheckBox chkShowFixed 
      BackColor       =   &H80000005&
      Caption         =   "固定参数(&H)"
      Height          =   330
      Left            =   4185
      TabIndex        =   6
      Top             =   945
      Width           =   1410
   End
   Begin VB.ComboBox cboModule 
      Height          =   300
      Left            =   4605
      TabIndex        =   5
      Text            =   "cmbModule"
      Top             =   585
      Width           =   2700
   End
   Begin VB.PictureBox picPage 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   0
      ScaleHeight     =   2616
      ScaleWidth      =   10500
      TabIndex        =   11
      Top             =   3720
      Width           =   10500
      Begin VB.PictureBox picDetailParas 
         BorderStyle     =   0  'None
         Height          =   2220
         Left            =   240
         ScaleHeight     =   2220
         ScaleWidth      =   10212
         TabIndex        =   16
         Top             =   120
         Width           =   10215
         Begin VB.Frame fraDetaisModi 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   10095
            Begin VB.CommandButton cmdDel 
               Caption         =   "删除参数设置(&D)"
               Height          =   350
               Left            =   5535
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   120
               Width           =   1600
            End
            Begin VB.CommandButton cmdAddNew 
               Caption         =   "新增参数设置(&N)"
               Height          =   350
               Left            =   7320
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   120
               Width           =   1600
            End
            Begin VB.CommandButton cmdSearch 
               Height          =   240
               Left            =   3435
               Picture         =   "frmParameters.frx":9154
               Style           =   1  'Graphical
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   180
               Width           =   240
            End
            Begin VB.CommandButton cmdModValue 
               Caption         =   "修改参数值(&B)"
               Height          =   350
               Left            =   3960
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   120
               Width           =   1400
            End
            Begin VB.TextBox txtSearch 
               Height          =   300
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   25
               Top             =   150
               Width           =   1545
            End
            Begin VB.Label lblTip 
               AutoSize        =   -1  'True
               Caption         =   "参数查找："
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   28
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblSearch 
               AutoSize        =   -1  'True
               Caption         =   "用户名(&U)↓"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   960
               TabIndex        =   26
               Tag             =   "1"
               Top             =   210
               Width           =   1095
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDetailParas 
            Height          =   840
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   7710
            _cx             =   13600
            _cy             =   1482
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
            BackColor       =   -2147483628
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483628
            GridColor       =   12632256
            GridColorFixed  =   -2147483630
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   300
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":9468
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
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   2
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   0
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picParInfo 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3120
         ScaleHeight     =   2292
         ScaleWidth      =   6132
         TabIndex        =   18
         Top             =   1080
         Width           =   6135
         Begin VSFlex8Ctl.VSFlexGrid vsParaInfo 
            Height          =   2160
            Left            =   60
            TabIndex        =   19
            Top             =   120
            Width           =   6015
            _cx             =   10610
            _cy             =   3810
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483633
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483633
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483633
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":9604
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
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
            ForeColorFrozen =   -2147483633
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picParaChangeLog 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1692
         ScaleWidth      =   7692
         TabIndex        =   14
         Top             =   1080
         Width           =   7695
         Begin VSFlex8Ctl.VSFlexGrid vsChangeLog 
            Height          =   1320
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   7470
            _cx             =   13176
            _cy             =   2328
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
            BackColor       =   -2147483628
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483628
            GridColor       =   12632256
            GridColorFixed  =   -2147483630
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":967F
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
            ExplorerBar     =   7
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
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   960
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   900
         _Version        =   589884
         _ExtentX        =   1587
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.ComboBox cboSys 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   585
      Width           =   2445
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7050
      Top             =   6300
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":9766
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   7200
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":B4F8
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":BA92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":C02C
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":C37E
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":12BE0
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":19442
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":1990A
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl参数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数类型"
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lbl模块 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模块"
      Height          =   180
      Left            =   4185
      TabIndex        =   4
      Top             =   645
      Width           =   360
   End
   Begin VB.Image imgMain 
      Height          =   384
      Left            =   180
      Picture         =   "frmParameters.frx":19DD2
      Top             =   576
      Width           =   384
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "系统参数管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   8
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "用户名(&U)"
         Index           =   0
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "人  员(&P)"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "部  门(&W)"
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "机器名(&T)"
         Index           =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "院  区(&S)"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "参数值(&R)"
         Index           =   5
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=========================================================
'============                         模块变量                   ======================
'=========================================================
Private mrsPars As ADODB.Recordset '参数信息记录集
Private mrsModule As ADODB.Recordset '模块信息记录集
Private mrsSys As New ADODB.Recordset '系统信息
Private mrsDetailParas As ADODB.Recordset '各部门、各用户、各机器参数值
Private mlngSys As Long '上次的选择系统
Private mlngModule As Long '上次的模块
Private mlngParID As Long '上次的参数ID
Private mstrParType As String '上次的参数类型
Private mlngModulePreIdx As Long '又来恢复输入匹配后恢复原来选择的模块
Private mblnNotClick As Boolean '不触发Click事件的临时变量
Private mblnMultiSta As Boolean '是否启用多站点
Private mstrOwner As String '当前系统所有者

Private Enum ChangeCtrl '发生改变的控件
    CT_Sys = 0
    CT_Module = 1
End Enum

Private Enum mPageNum
    Pag_ParaInfo = 0
    Pag_Computer = 1
    Pag_ChangeLog = 2
End Enum

Private Enum ParaInfoRow '参数说明行枚举
    PR_影响控制说明 = 0
    PR_参数值含义 = 1
    PR_关联说明 = 2
    PR_适用说明 = 3
    PR_警告说明 = 4
End Enum

'搜索标签
Private Enum mnuIndex
    MI_用户名 = 0
    MI_人员 = 1
    MI_部门 = 2
    MI_机器名 = 3
    MI_站点 = 4
    MI_参数值 = 5
End Enum

Private Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '进入控件时,选择显示颜色
Private Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '离开焦点时,选择的显示颜色
'=========================================================
'============                         公共接口                   ======================
'=========================================================
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Me.ActiveControl Is vsDetailParas Then
        objPrint.Title.Text = "分院区或用户的参数清单打印"
    Else
        objPrint.Title.Text = "参数清单打印"
    End If
    
    objRow.Add "应用系统：" & cboSys.Text
    objPrint.UnderAppRows.Add objRow
    If Me.ActiveControl Is vsDetailParas Then
        Set objRow = New zlTabAppRow
        objRow.Add "参数类型：" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数类型"))
        objRow.Add "参数名称：" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数号")) & "-" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数名"))
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "参数说明：" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数说明"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    If Me.ActiveControl Is vsDetailParas Then
        Set objPrint.Body = vsDetailParas
    Else
        Set objPrint.Body = vsPara
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cboModule_LostFocus()
    Dim strKey As String
    strKey = cboModule.Text
    
    If cboModule.ListIndex >= 0 Then
        If cboModule.List(cboModule.ListIndex) = strKey Then
            Exit Sub
        End If
        cboModule.Text = cboModule.List(cboModule.ListIndex)
     Else
        If mlngModulePreIdx >= 0 Then
            cboModule.ListIndex = mlngModulePreIdx
        Else
            cboModule.ListIndex = 0
        End If
    End If
End Sub

'=========================================================
'============                         控件事件                   ======================
'=========================================================
Private Sub cboSys_Click()
    '切换系统，后才刷新数据
    If mblnNotClick Then Exit Sub
    If mlngSys <> cboSys.ItemData(cboSys.ListIndex) Or cboSys.Tag = "强制刷新" Then
        mlngSys = cboSys.ItemData(cboSys.ListIndex)
        mrsSys.Filter = "编号=" & mlngSys
        If Not mrsSys.EOF Then mstrOwner = mrsSys!所有者 & ""
        If cboSys.Tag <> "强制刷新" Then
            mlngModule = -1
            mstrParType = ""
        End If
        Call GetParasInfo(mlngSys)
        Call LoadParas
        Call ResetCtrl
        Call SetParas
    End If
End Sub

Private Sub cmdAddNew_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strDetails As String
    Dim StrValue As String, strUsers As String, strPCs As String
    Dim objfrmParaModiSet As New frmParaModiSet
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数类型"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub
    If Not objfrmParaModiSet.ShowMe(Me, 1, fraDetaisModi.Tag, "", mstrOwner, lngParID, StrValue, strUsers, strPCs) Then Exit Sub
    
    Call ExecuteProcedure("Zlparameters_Add_Details(" & lngParID & ",'" & UCase(strUsers) & "','" & strPCs & "','" & StrValue & "')", "批量修改参数值")
    Set mrsDetailParas = Nothing
    '刷新参数
    Call LoadDetailParas(lngParID, False)
End Sub

Private Sub cmdDel_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strValues As String
    Dim strInfo As String
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数类型"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub

    With vsDetailParas
        If .Tag > 1 Then
            If MsgBox("是否删除已经选中的" & .Tag & "条记录？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        If .Tag <> 0 Then '处理选择行
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                    If strParsType = "部门参数" Then
                        strTmp = .TextMatrix(i, .ColIndex("部门ID"))
                        strInfo = """" & .TextMatrix(i, .ColIndex("部门")) & """"
                    Else
                        strTmp = .TextMatrix(i, .ColIndex("用户名")) & "^" & .TextMatrix(i, .ColIndex("机器名"))
                        If .ColHidden(.ColIndex("人员")) Then
                            strInfo = """" & .TextMatrix(i, .ColIndex("机器名")) & """"
                        ElseIf .ColHidden(.ColIndex("机器名")) Then
                            strInfo = """" & .TextMatrix(i, .ColIndex("人员")) & """"
                        Else
                            strInfo = """" & .TextMatrix(i, .ColIndex("人员")) & """在""" & .TextMatrix(i, .ColIndex("机器名")) & """上"
                        End If
                    End If
                    If ActualLen(strValues & "#" & strTmp) >= 2000 Then
                        Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "删除参数设置")
                        strValues = strTmp
                    Else
                        strValues = IIf(strValues = "", strTmp, strValues & "#" & strTmp)
                    End If
                End If
            Next
            If .Tag = 1 Then
                If MsgBox("是否删除" & strInfo & "的参数设置？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If strValues <> "" Then
                Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "删除参数设置")
            End If
        Else '只处理当前行
            If .RowData(.Row) <> 0 Then
                If strParsType = "部门参数" Then
                    strValues = .TextMatrix(.Row, .ColIndex("部门ID"))
                    strInfo = """" & .TextMatrix(.Row, .ColIndex("部门")) & """"
                Else
                    strValues = .TextMatrix(.Row, .ColIndex("用户名")) & "^" & .TextMatrix(.Row, .ColIndex("机器名"))
                    If .ColHidden(.ColIndex("人员")) Then
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("机器名")) & """"
                    ElseIf .ColHidden(.ColIndex("机器名")) Then
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("人员")) & """"
                    Else
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("人员")) & """在""" & .TextMatrix(.Row, .ColIndex("机器名")) & """上"
                    End If
                End If
                If MsgBox("是否删除" & strInfo & "的参数设置？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "删除参数设置")
            End If
        End If
    End With
    Set mrsDetailParas = Nothing
    '刷新参数
    Call LoadDetailParas(lngParID, txtSearch.Text <> "")
End Sub

Private Sub cmdExp_Click()
    Dim strSets As String
    Dim arrTmp As Variant

    Dim rsParas As ADODB.Recordset
    strSets = frmParaInOut.ShowMe(PST_Exp, mlngSys)
    If strSets = "" Then Exit Sub
    arrTmp = Split(strSets, "|")
    Set rsParas = GetALLPars(IIf(Val(arrTmp(1)) = 0, mlngSys, -1), Val(arrTmp(2)) = 0, True)
'    Set rsParas = CopyNewRec(rsParas) '改变为可变记录集
    If gobjFile.FileExists(arrTmp(0)) Then
        Call gobjFile.DeleteFile(arrTmp(0), True)
    End If
    rsParas.Save arrTmp(0), adPersistXML
    MsgBox "参数导出成功！", vbInformation, gstrSysName
End Sub

Private Sub cmdImp_Click()
    Dim strTmp As String, arrSets As Variant, strValues As String
    Dim rsParas As ADODB.Recordset, rsOldPars As ADODB.Recordset, rsComInfo As ADODB.Recordset
    Dim strSQL As String, arrTmp As Variant, arrCols As Variant, strTmpSQL As String
    Dim i As Long, j As Long
    Dim strPre As String, strCur As String, strMsg As String
    Dim strFilter As String, strFilterEx As String, strFilterOld As String, strFilterTmp As String, strFilterNew As String
    Dim lngSys As Long, blnMultiSys As Boolean, blnDetails As Boolean
    Dim strDeptParas As String, strUserParas As String
    Dim blnTrans As Boolean
    Dim dtStart As Date
    Dim cllErrSQL As Collection '错误SQL
    
    On Error GoTo errH
    strTmp = frmParaInOut.ShowMe(PST_Imp, mlngSys)
    If strTmp = "" Then Exit Sub
    dtStart = Now
    arrTmp = Split(strTmp, "|")
    lngSys = IIf(Val(arrTmp(1)) = 0, mlngSys, -1): blnDetails = Val(arrTmp(2)) = 0
    '获取对应配置的数据
    Set rsOldPars = CopyNewRec(GetALLPars(lngSys, False))  '不获取详细的参数配置，因为数据的不必要
    Set rsParas = New ADODB.Recordset
    rsParas.Open arrTmp(0), , adOpenStatic, adLockOptimistic, adCmdFile
    '获取导入条件
    If lngSys <> -1 Then strFilter = "系统=" & mlngSys: strFilterEx = "系统<>" & mlngSys '只导入当前
    If Not blnDetails Then strFilter = strFilter & IIf(strFilter = "", "", " And ") & " 类型<1":  strFilterEx = strFilterEx & IIf(strFilterEx = "", "", " OR  ") & " 类型>0"
    strFilter = IIf(strFilter = "", "", strFilter & " And  类型<>-99"):  strFilterEx = IIf(strFilterEx = "", "", strFilterEx & " OR  类型=-99")
    '系统版本号对比  名称 参数名, 版本号 参数值, User 缺省值,To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss') 影响控制说明
    rsOldPars.Filter = strFilter & IIf(strFilter = "", "", " And ") & "类型=-9": rsParas.Filter = strFilter & IIf(strFilter = "", "", " And ") & "类型=-9"
    Set rsComInfo = GetCompareRec(rsOldPars, rsParas, "系统", "参数值", "参数名")
    Debug.Print "系统比对=" & DateDiff("s", dtStart, Now)
    strTmp = "": rsComInfo.Filter = "State=-1" '数据中存在的系统，但是导入文件中没有该系统
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!参数名, 18)
        strFilterOld = strFilterOld & " OR 系统=" & Val(rsComInfo!MainKey)
        rsComInfo.MoveNext
    Loop
    If strTmp <> "" Then '导入文件中没有的系统，不进行比较
        strMsg = "参数文件缺失如下系统：" & _
                            strTmp & vbNewLine & _
                        "这些系统的参数将不进行导入操作。"
    End If
    strTmp = "": rsComInfo.Filter = "State=1" '数据中不存在的系统，但是导入文件中有该系统
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!参数名_NEW & "", 18)
        strFilterNew = strFilterNew & " OR 系统=" & Val(rsComInfo!MainKey)
        rsComInfo.MoveNext
    Loop
    strTmp = "": strFilterTmp = "": rsComInfo.Filter = "State=2" '导入文件与数据库均存在该系统
    blnMultiSys = rsComInfo.RecordCount > 1
    Do While Not rsComInfo.EOF
        strFilterTmp = strFilterTmp & " OR 系统=" & Val(rsComInfo!MainKey)
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!参数名, 18) & " 当前版本:" & VerPAD(rsComInfo!参数值) & "<>导入版本：" & VerPAD(rsComInfo!参数值_New)
        rsComInfo.MoveNext
    Loop
    
    If strTmp <> "" Then
        If blnMultiSys Then  '多个系统
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "如下系统：" & _
                                strTmp & vbNewLine & _
                            "版本存在差异，导入可能会影响系统功能使用。是否导入这些系统的参数？"
        Else
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "系统：" & Mid(strTmp, 2) & vbNewLine & _
                            "版本存在差异，导入可能会影响该系统功能使用。是否导入该系统的参数？"
        End If
        '版本有差异的系统，询问是否导入。不导入，则不比较这些系统
        If MsgBox(strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            rsComInfo.Filter = "State=0"
            If rsComInfo.RecordCount = 0 Then '没有可导入的系统
                MsgBox "没有可导入的系统！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName
                Exit Sub
            End If
        Else
            strFilterTmp = ""
        End If
        strMsg = ""
    Else
        rsComInfo.Filter = "State=0"
        If rsComInfo.RecordCount = 0 Then '没有可导入的系统
            MsgBox "没有可导入的系统！" & IIf(strMsg = "", "", "具体情况如下：" & vbNewLine) & strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName
            Exit Sub
        End If
    End If
    strTmp = "": rsComInfo.Filter = IIf(strFilterTmp = "", "State=2 OR State=0", "State=0") '导入文件与数据库均存在该系统
    blnMultiSys = rsComInfo.RecordCount > 1
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!参数名_NEW, 18)
        rsComInfo.MoveNext
    Loop
    If strTmp <> "" Then
        If blnMultiSys Then '多个系统
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "如下系统将会被导入：" & _
                                strTmp & vbNewLine & _
                            IIf(blnDetails, "这些系统的部门、本机、私有参数将被清空，重新导入。", "") & "是否继续？"
        Else
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "系统：" & Mid(strTmp, 2) & vbNewLine & _
                             "将会被导入。" & IIf(blnDetails, "该系统部门、本机、私有参数将被清空，重新导入。", "") & "是否继续？"
        End If
        '再次询问是否继续导入，防止误操作
        If MsgBox(strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call ShowFlash("正在进行参数导入，请稍候！")
    Debug.Print "系统确认=" & DateDiff("s", dtStart, Now)
    '删除不导入的的
    If strFilterOld <> "" Or strFilterTmp <> "" Or strFilterEx <> "" Then
        Call RecDelete(rsOldPars, strFilterEx & IIf(strFilterEx = "", "", IIf(strFilterNew & strFilterTmp <> "", " OR ", "")) & Mid(strFilterOld & strFilterTmp, Len(" OR ") + 1))
    End If
    If strFilterNew <> "" Or strFilterTmp <> "" Or strFilterEx <> "" Then
        Call RecDelete(rsParas, strFilterEx & IIf(strFilterEx = "", "", IIf(strFilterNew & strFilterTmp <> "", " OR ", "")) & Mid(strFilterNew & strFilterTmp, Len(" OR ") + 1))
    End If
    Debug.Print "删除=" & DateDiff("s", dtStart, Now)
    rsOldPars.Filter = "类型=0": rsParas.Filter = "类型=0"
    Set rsComInfo = GetCompareRec(rsOldPars, rsParas, "MAINKEY", "-SORTKEY,系统,模块,参数名,类型,部门id,用户名,机器名,详细参数值", "", Array("SQL", adVarWChar, 20000, Empty))
    Debug.Print "类型变动=" & DateDiff("s", dtStart, Now)
    
    Set cllErrSQL = New Collection
    gcnOracle.BeginTrans: blnTrans = True
    On Error Resume Next
    With rsComInfo
        .Filter = "State<>0"
        .Sort = "Sort,MainKey"
        Do While Not .EOF
            strSQL = "": arrTmp = Split(!MainKey, "#")
            '表格导入测试代码
            Select Case rsComInfo!State
                Case 2 '修正
                    strTmpSQL = "Set ": arrCols = Split(!DifInfo & "", ",")
                    For i = LBound(arrCols) To UBound(arrCols)
                        If IsType(.Fields(arrCols(i) & "_New").Type, adVarChar) Then '字符串类型，则需要转换
                            strTmpSQL = strTmpSQL & IIf(i = 0, " ", " , ") & arrCols(i) & " = " & SQLAdjust(.Fields(arrCols(i) & "_New").value)
                        Else
                            strTmpSQL = strTmpSQL & IIf(i = 0, " ", " , ") & arrCols(i) & " = " & Val(.Fields(arrCols(i) & "_New").value & "")
                        End If
                    Next
                    strSQL = "Update zlParameters " & _
                                strTmpSQL & _
                                " Where Nvl(系统, 0) = " & Val(arrTmp(0)) & " And Nvl(模块, 0) = " & Val(arrTmp(1)) & " And 参数名 = '" & arrTmp(2) & "'"
                Case 1 '新增
                    strSQL = "Insert Into zlParameters (ID, 系统, 模块, 私有, 本机, 授权, 固定, 部门, 性质, 参数号, 参数名, 参数值, 缺省值,影响控制说明, 参数值含义, 关联说明, 适用说明, 警告说明) " & _
                                    "  Select zlParameters_ID.NextVal, " & ZVal(arrTmp(0)) & ", " & ZVal(arrTmp(1)) & ", " & Val(!私有_New & "") & ", " & Val(!本机_New & "") & " , " & Val(!授权_New & "") & ", " & Val(!固定_New & "") & " , " & _
                                    Val(!部门_New & "") & " ," & Val(!性质_New & "") & " ," & !参数号_New & " , '" & arrTmp(2) & "', " & SQLAdjust(!参数值_New) & " , " & SQLAdjust(!缺省值_New) & ", " & SQLAdjust(!影响控制说明_New) & _
                                    "  ," & SQLAdjust(!参数值含义_New) & "  , " & SQLAdjust(!关联说明_New) & " ," & SQLAdjust(!适用说明_New) & " , " & SQLAdjust(!警告说明_New) & " From Dual" & _
                                    " Where Not Exists (Select 1 From zlParameters Where 参数名 =" & SQLAdjust(arrTmp(2)) & " And Nvl(模块,0) = " & Val(arrTmp(1)) & " And Nvl(系统,0) = " & Val(arrTmp(0)) & ")"
                Case -1 '删除
                    strSQL = "Delete zlParameters Where Nvl(系统, 0) = " & Val(arrTmp(0)) & " And Nvl(模块, 0) = " & Val(arrTmp(1)) & " And 参数名 = '" & arrTmp(2) & "'"
            End Select
            strSQL = Replace(Trim(strSQL), ChrW(-3979), "")
            'ChrW(-3979),看起像chr(63)问号，但是不一样，会导致WriteLine 方法报错
            If strSQL <> "" Then
                gcnOracle.Execute strSQL
                If err.Number <> 0 Then
                    err.Clear: gcnOracle.Errors.Clear
                    cllErrSQL.Add strSQL
                End If
            End If
            .MoveNext
        Loop
        '错误SQL可能是由于参数号错位导致的，因此再重复执行一次
        For i = 0 To 1
            For j = 1 To cllErrSQL.Count
                If cllErrSQL(j) <> "" Then
                    gcnOracle.Execute cllErrSQL(j)
                    If err.Number <> 0 Then
                        err.Clear: gcnOracle.Errors.Clear
                    Else
                        cllErrSQL(j) = ""
                    End If
                End If
            Next
        Next
        On Error GoTo errH
        Debug.Print "导入列表清单=" & DateDiff("s", dtStart, Now)
        '删除参数详情
        If blnDetails Then
            .Filter = ""
            strDeptParas = "": strUserParas = ""
            Do While Not .EOF
                If (!部门 = 1 Or !私有 = 1 Or !本机 = 1) Or (!部门_New = 1 Or !私有_New = 1 Or !本机_New = 1) Then          '删除原来有参数详细配置，现在没有的参数配置
                    If !部门 = 1 Or !部门_New = 1 Then
                        strTmp = Replace(!MainKey, "#", "^")
                         If ActualLen(strDeptParas & "#" & strTmp) >= 2000 Then
                            strSQL = "Zlparameters_Delall_Details('" & strDeptParas & "',1)"
                            Call ExecuteProcedure(strSQL, "清空原来的参数配置")
                            strDeptParas = strTmp
                        Else
                            strDeptParas = strDeptParas & IIf(strDeptParas <> "", "#", "") & strTmp
                        End If
                    End If
                    If (!私有 = 1 Or !本机 = 1) Or (!私有_New = 1 Or !本机_New = 1) Then
                        strTmp = Replace(!MainKey, "#", "^")
                         If ActualLen(strUserParas & "#" & strTmp) >= 2000 Then
                            strSQL = "Zlparameters_Delall_Details('" & strUserParas & "',0)"
                            Call ExecuteProcedure(strSQL, "清空原来的参数配置")
                            strUserParas = strTmp
                        Else
                            strUserParas = strUserParas & IIf(strUserParas <> "", "#", "") & strTmp
                        End If
                    End If
                End If
                .MoveNext
            Loop
            If strDeptParas <> "" Then
                strSQL = "Zlparameters_Delall_Details('" & strDeptParas & "',1)"
                Call ExecuteProcedure(strSQL, "清空原来的参数配置")
            End If
            If strUserParas <> "" Then
                strSQL = "Zlparameters_Delall_Details('" & strUserParas & "',0)"
                Call ExecuteProcedure(strSQL, "清空原来的参数配置")
            End If
        End If
    End With
    Debug.Print "清空原有参数=" & DateDiff("s", dtStart, Now)
    If blnDetails Then
        '导入新的参数详情
        strPre = "": strCur = "": strValues = ""
        With rsParas
            .Filter = "类型>0" '现存的参数详情都是需要导入的
            .Sort = "SortKey" '进行排序
            Do While Not rsParas.EOF
                strCur = rsParas!系统 & "," & rsParas!模块 & ",'" & Trim(rsParas!参数名 & "") & "'"
                If strCur <> strPre Then
                    '增加上一个参数的参数详情
                    If strPre <> "" And strValues <> "^^" And strValues <> "^^#^^" Then
                        strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                        Call ExecuteProcedure(strSQL, "导入参数详情")
                        strValues = ""
                    End If
                    strPre = strCur
                End If
                
                If Val(!类型 & "") = 1 Then
                    strTmp = rsParas!部门id & "^^" & rsParas!详细参数值
                Else
                    strTmp = rsParas!用户名 & "^" & rsParas!机器名 & "^" & rsParas!详细参数值
                End If
                If ActualLen(strValues & "#" & strTmp) >= 2000 Then
                    strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                    Call ExecuteProcedure(strSQL, "导入参数详情")
                    strValues = strTmp
                Else
                    strValues = IIf(strValues = "", strTmp, strValues & "#" & strTmp)
                End If
                rsParas.MoveNext
            Loop
            If strValues <> "" Then
                strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                Call ExecuteProcedure(strSQL, "导入参数详情")
            End If
        End With
    End If
    gcnOracle.CommitTrans:  blnTrans = False
    Debug.Print "导入成功=" & DateDiff("s", dtStart, Now)
    ShowFlash ("")
    MsgBox "参数导入成功！", vbInformation, gstrSysName
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    ShowFlash ("")
    If blnTrans Then
        MsgBox "参数导入失败！错误信息如下：" & vbNewLine & gcnOracle.Errors(0).Description, vbInformation, "参数导入"
    Else
        MsgBox "参数导入失败！错误信息如下：" & vbNewLine & err.Description, vbInformation, "参数导入"
    End If
End Sub

Private Sub cmdModValue_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strDetails As String
    Dim StrValue As String
    Dim objfrmParaModiSet As New frmParaModiSet
    Dim strInfo As String
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("参数类型"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub
    
    With vsDetailParas
        If .Tag > 1 Then
            strInfo = "调整选中的" & .Tag & "条参数设置。"
            i = .FindRow(-1, , .ColIndex("选择"))
            If i <> -1 Then
                StrValue = .TextMatrix(i, .ColIndex("参数值"))
            End If
        Else
            If .Tag = 0 Then
                i = .Row
            Else
                i = .FindRow(-1, , .ColIndex("选择"))
            End If
            If i = -1 Then
                MsgBox "当前未选中任何参数设置！", vbInformation, gstrSysName
                Exit Sub
            Else
                If strParsType = "部门参数" Then
                    strInfo = """" & .TextMatrix(i, .ColIndex("部门")) & """"
                Else
                    If .ColHidden(.ColIndex("人员")) Then
                        strInfo = """" & .TextMatrix(i, .ColIndex("机器名")) & """"
                    ElseIf .ColHidden(.ColIndex("机器名")) Then
                        strInfo = """" & .TextMatrix(i, .ColIndex("人员")) & """"
                    Else
                        strInfo = """" & .TextMatrix(i, .ColIndex("人员")) & """在""" & .TextMatrix(i, .ColIndex("机器名")) & """上"
                    End If
                End If
                strInfo = "调整" & strInfo & "的参数值。"
                StrValue = .TextMatrix(i, .ColIndex("参数值"))
            End If
        End If
        If Not objfrmParaModiSet.ShowMe(Me, 0, fraDetaisModi.Tag, strInfo, mstrOwner, lngParID, StrValue) Then Exit Sub

        If .Tag <> 0 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                    If strParsType = "部门参数" Then
                        strTmp = .TextMatrix(i, .ColIndex("部门ID")) & "^^" & Trim(StrValue)
                    Else
                        strTmp = .TextMatrix(i, .ColIndex("用户名")) & "^" & .TextMatrix(i, .ColIndex("机器名")) & "^" & Trim(StrValue)
                    End If
                    If ActualLen(strDetails & "#" & strTmp) >= 2000 Then
                        Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "批量修改参数值")
                        strDetails = strTmp
                    Else
                        strDetails = IIf(strDetails = "", strTmp, strDetails & "#" & strTmp)
                    End If
                End If
            Next
            If strDetails <> "" Then
                Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "批量修改参数值")
            End If
        Else '只处理当前行
            If .RowData(.Row) <> 0 Then
                If strParsType = "部门参数" Then
                    strDetails = .TextMatrix(.Row, .ColIndex("部门ID")) & "^^" & Trim(StrValue)
                Else
                    strDetails = .TextMatrix(.Row, .ColIndex("用户名")) & "^" & .TextMatrix(.Row, .ColIndex("机器名")) & "^" & Trim(StrValue)
                End If
                Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "批量修改参数值")
            End If
        End If
    End With

    Set mrsDetailParas = Nothing
    '刷新参数
    Call LoadDetailParas(lngParID, txtSearch.Text <> "")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Call ShowFlash("正在加载参数！")
    '判断是否弃用多站点
    strSQL = "Select 1 From zlClients Where 站点 Is Not Null and rownum<2"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    mblnMultiSta = Not rsTmp.EOF
    mlngSys = -1: mlngModule = -1
    Call RestoreVsGridWidth(vsChangeLog, Me.Caption, "参数变动日志")
    Call RestoreVsGridWidth(vsDetailParas, Me.Caption, "站点及用户")
    Call RestoreVsGridWidth(vsPara, Me.Caption, "系统参数列表")
    With vsDetailParas
        .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, .ColIndex("选择")) = flexAlignCenterCenter
    End With
    Call InitFace '界面初始化
    mblnNotClick = True
    Call LoadSystems '加载应用系统
    mblnNotClick = False
    Call cboSys_Click
    Call vsPara_AfterRowColChange(-1, -1, 1, 1)
    ShowFlash ("")
End Sub
'=========================================================
'============                         私有方法                   ======================
'=========================================================
Private Sub InitFace()
'功能：初始化界面
    Dim objItem As TabControlItem
    '页面控件设置
    Set objItem = tbPage.InsertItem(Pag_ParaInfo, "参数说明信息", picParInfo.hwnd, 0)
    objItem.Tag = Pag_ParaInfo
    Set objItem = tbPage.InsertItem(Pag_Computer, "院区及用户", picDetailParas.hwnd, 0)
    objItem.Tag = Pag_Computer
    Set objItem = tbPage.InsertItem(Pag_ChangeLog, "参数变动日志", picParaChangeLog.hwnd, 0)
    objItem.Tag = Pag_ChangeLog
    With tbPage
         tbPage.Item(Pag_ParaInfo).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
    End With
     '20个中文字符宽度可以显示系统下拉列表内容
    If cboSys.Width < 20 * Me.TextWidth("宽") Then
        Call CboSetWidth(cboSys.hwnd, 20 * Me.TextWidth("宽"))
    End If
    mblnNotClick = True
    cboParType.AddItem "所有类型": cboParType.ItemData(cboParType.NewIndex) = 0
    cboParType.ListIndex = 0
    
    cboModule.AddItem "所有参数"
    cboModule.ItemData(cboModule.NewIndex) = -1
    cboModule.ListIndex = 0
    mblnNotClick = False
End Sub

Private Sub LoadSystems()
'功能：加载系统
    Dim strSQL As String
    Dim strVer As String
    '获取管理工具版本号
    strVer = GetToolsVersion
    '增加共享号排序，主要是将主系统排在前面
    Set mrsSys = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    If gblnDBA Then
        mrsSys.Filter = ""
    Else
        mrsSys.Filter = "所有者='" & UCase(gstrUserName) & "'"
    End If
    mrsSys.Sort = "编号"
    With mrsSys
        '添加管理工具历史记录查看。
        cboSys.Clear
        cboSys.AddItem String(5, " ") & RPAD("服务器管理工具", 18) & " v" & VerPAD(strVer)
        cboSys.ItemData(cboSys.NewIndex) = 0
        cboSys.ListIndex = cboSys.NewIndex
        Do While Not .EOF
            cboSys.AddItem Lpad(!编号, 4) & "-" & RPAD(!名称 & "", 18) & " v" & VerPAD(!版本号 & "")
            cboSys.ItemData(cboSys.NewIndex) = !编号
            .MoveNext
        Loop
        '缺省为管理工具
        If cboSys.ListCount <= 1 Then
            cboSys.Locked = True
        End If
    End With
End Sub

Private Function GetParasInfo(ByVal lngSys As Long) As Boolean
'功能：获取参数相关信息
    Dim rsTmp As New ADODB.Recordset
    Dim strKey As String, strName As String, strCode As String, arrTmp As Variant
    Dim intTotal As Integer, intFixed As Integer
    Dim int公共全局 As Integer, int私有全局 As Integer, int公共模块 As Integer, int私有模块 As Integer, int本机公共 As Integer, int本机私有 As Integer, int部门参数 As Integer
    Dim strParType As String, blnFixed As Boolean
    On Error GoTo errH
    Set mrsPars = New ADODB.Recordset
    Set mrsModule = New ADODB.Recordset
    'Id, 系统,模块,私有,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明,本机,授权,固定,部门,模块名称,模块简码
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameters", lngSys)
    rsTmp.Filter = "性质=0" '不显示程序内部控制所使用的参数
    Set mrsPars = CopyNewRec(rsTmp, , , Array("Fixed", adInteger, 1, 0, "ParType", adVarChar, 50, Empty))
    Set mrsModule = CopyNewRec(rsTmp, True, "系统,模块,模块名称,模块简码", Array("Key", adVarChar, 200, Empty, "Fixed", adInteger, 3, 0, "Total", adInteger, 5, 0, _
                                                                                                                                "公共全局", adInteger, 3, 0, "私有全局", adInteger, 3, 0, _
                                                                                                                                "公共模块", adInteger, 3, 0, "私有模块", adInteger, 3, 0, _
                                                                                                                                "本机公共模块", adInteger, 3, 0, "本机私有模块", adInteger, 3, 0, _
                                                                                                                                "部门参数", adInteger, 3, 0, "Index", adInteger, 5, 0))
    mrsPars.Filter = ""
    mrsPars.Sort = "系统,模块,参数号"
    With mrsPars
        Do While Not mrsPars.EOF
            If strKey <> !系统 & "_" & !模块 Then
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    mrsModule.AddNew Array("Key", "系统", "模块", "模块名称", "模块简码", "Fixed", "Total", "公共全局", "私有全局", "公共模块", "私有模块", "本机公共模块", "本机私有模块", "部门参数"), _
                                                    Array(strKey, Val(arrTmp(0)), Val(arrTmp(1)), strName, strCode, intFixed, intTotal, int公共全局, int私有全局, int公共模块, int私有模块, int本机公共, int本机私有, int部门参数)
                End If
                strKey = !系统 & "_" & !模块
                strName = !模块名称 & "": strCode = !模块简码 & ""
                intFixed = 0: intTotal = 0
                int公共全局 = 0: int私有全局 = 0: int公共模块 = 0: int私有模块 = 0: int本机公共 = 0: int本机私有 = 0
            End If
            strParType = GetParaType(Val(Nvl(!模块)), Val(Nvl(!私有)), Val(Nvl(!本机)), Val(Nvl(!部门)))
            Select Case strParType
                Case "公共全局"
                    int公共全局 = int公共全局 + 1
                Case "私有全局"
                    int私有全局 = int私有全局 + 1
                Case "公共模块"
                    int公共模块 = int公共模块 + 1
                Case "私有模块"
                    int私有模块 = int私有模块 + 1
                Case "本机公共模块"
                    int本机公共 = int本机公共 + 1
                Case "本机私有模块"
                    int本机私有 = int本机私有 + 1
                Case "部门参数"
                    int部门参数 = int部门参数 + 1
            End Select
            intTotal = intTotal + 1 '总计数+1
            '不能调整参数类型的参数
            If !固定 = 1 Or strParType = "公共全局" Or strParType = "私有全局" Then
                intFixed = intFixed + 1: blnFixed = True
            Else
                blnFixed = False
            End If
            .Update Array("Fixed", "ParType"), Array(IIf(blnFixed, 1, 0), strParType)
            .MoveNext
        Loop
        If strKey <> "" Then
            arrTmp = Split(strKey, "_")
            mrsModule.AddNew Array("Key", "系统", "模块", "模块名称", "模块简码", "Fixed", "Total", "公共全局", "私有全局", "公共模块", "私有模块", "本机公共模块", "本机私有模块", "部门参数"), _
                                            Array(strKey, Val(arrTmp(0)), Val(arrTmp(1)), strName, strCode, intFixed, intTotal, int公共全局, int私有全局, int公共模块, int私有模块, int本机公共, int本机私有, int部门参数)
        End If
    End With
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub ResetCtrl(Optional ByVal ctInput As ChangeCtrl = CT_Sys)
'参数类型、参数模块，固定参数三者相互设置，以及
'参数：intChangeCtrl，发生改变的控件
    Dim strParTypes As String, strOldFilter As String
    Dim arrParType As Variant, i As Long, blnMuiltRow As Boolean
    Dim blnShowFixed As Boolean
    
    With mrsModule
        If ctInput = CT_Sys Then
            mblnNotClick = True
            .Filter = ""
            .Sort = "模块"
            cboModule.Clear
            If .RecordCount <> 1 Then
                mrsPars.Filter = "": mrsPars.Sort = "模块,参数号,ID"
                cboModule.AddItem "所有参数" & "(" & mrsPars.RecordCount & ")"
                cboModule.ItemData(cboModule.NewIndex) = -1
                If .RecordCount = 0 Then
                    cboModule.ListIndex = cboModule.NewIndex
                    Call ResetCtrl(CT_Module)
                    mblnNotClick = False
                    Exit Sub
                End If
            End If
            '加载模块
            Do While Not .EOF
                If Val(Nvl(!模块)) = 0 Then
                    cboModule.AddItem "系统参数" & "(" & !Total & ")"
                    cboModule.ItemData(cboModule.NewIndex) = Nvl(!模块)
                Else
                    cboModule.AddItem Nvl(!模块) & "-" & Nvl(!模块名称) & "(" & !Total & ")"
                    cboModule.ItemData(cboModule.NewIndex) = Nvl(!模块)
                End If
                mrsModule.Update "Index", cboModule.NewIndex '记录索引
                If mlngModule = Val(Nvl(!模块)) Then cboModule.ListIndex = cboModule.NewIndex
                .MoveNext
            Loop
            If cboModule.ListIndex < 0 Then cboModule.ListIndex = 0
            Call ResetCtrl(CT_Module)
            mblnNotClick = False
        ElseIf ctInput = CT_Module Then
            chkShowFixed.value = 0
            If cboModule.ItemData(cboModule.ListIndex) = -1 Then
                .Filter = ""
            Else
                .Filter = "模块=" & cboModule.ItemData(cboModule.ListIndex)
            End If
            .Sort = "模块"
            If .RecordCount = 0 Then
                strParTypes = ""
            Else
                blnMuiltRow = .RecordCount > 1
                strOldFilter = .Filter
                If strOldFilter = "0" Then strOldFilter = ""
                If Not blnMuiltRow Then
                    blnShowFixed = Val(!Fixed) <> 0
                Else
                    .Filter = IIf(strOldFilter <> "", strOldFilter & " And ", "") & "Fixed<>0"
                    blnShowFixed = .RecordCount <> 0
                    .Filter = strOldFilter
                End If
                arrParType = Array("公共全局", "私有全局", "公共模块", "私有模块", "本机公共模块", "本机私有模块", "部门参数")
                For i = LBound(arrParType) To UBound(arrParType)
                    If blnMuiltRow Then
                        .Filter = IIf(strOldFilter <> "" And strOldFilter <> "0", strOldFilter & " And ", "") & arrParType(i) & "<>0"
                        If .RecordCount <> 0 Then
                            strParTypes = strParTypes & "," & arrParType(i)
                        End If
                    Else
                        If Val(.Fields(arrParType(i))) <> 0 Then
                            strParTypes = strParTypes & "," & arrParType(i)
                        End If
                    End If
                Next
                chkShowFixed.Visible = blnShowFixed
                If blnShowFixed And Not blnMuiltRow Then '全部是固定参数，则固定参数显示勾选
                    If !Fixed = !Total Then
                        chkShowFixed.value = 1
                    End If
                End If
                '所有参数类型,"公共全局","私有全局","公共模块","私有模块","本机公共模块","本机私有模块", "部门参数"
                arrParType = Split(strParTypes, ",")
                cboParType.Clear
                '格式为,类型，只有一种可选类型，则没有所有参数类型
                If UBound(arrParType) - LBound(arrParType) + 1 = 2 Then
                    cboParType.AddItem arrParType(UBound(arrParType)): cboParType.ItemData(cboParType.NewIndex) = 0
                    cboParType.ListIndex = 0
                    Exit Sub
                End If
                For i = LBound(arrParType) To UBound(arrParType)
                    If arrParType(i) = "" Then arrParType(i) = "所有类型"
                    cboParType.AddItem arrParType(i): cboParType.ItemData(cboParType.NewIndex) = i
                    If arrParType(i) = mstrParType Then cboParType.ListIndex = cboParType.NewIndex
                Next
                If cboParType.ListIndex < 0 Then cboParType.ListIndex = 0
            End If
        End If
    End With
End Sub

Private Function LoadParas() As Boolean
    Dim i As Long
    'Id, 系统,模块,私有,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明,本机,授权,固定,模块名称,模块简码
    With vsPara
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = 0: .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpPicture, 1, .ColIndex("标志")) = Nothing
        mrsPars.Filter = "": mrsPars.Sort = "模块,参数号,ID"
        .Rows = IIf(mrsPars.RecordCount = 0, 1, mrsPars.RecordCount) + 1
        For i = 1 To mrsPars.RecordCount
            .RowData(i) = Nvl(mrsPars!Id)
             ' 私有,本机,授权,固定,模块名称,模块简码
            .TextMatrix(i, .ColIndex("参数类型")) = mrsPars!ParType
            .TextMatrix(i, .ColIndex("模块名称")) = Nvl(mrsPars!模块名称)
            .TextMatrix(i, .ColIndex("参数号")) = Nvl(mrsPars!参数号)
            .TextMatrix(i, .ColIndex("参数名")) = Nvl(mrsPars!参数名)
            .TextMatrix(i, .ColIndex("参数值")) = Nvl(mrsPars!参数值)
            .TextMatrix(i, .ColIndex("授权")) = IIf(Val(Nvl(mrsPars!授权)) = 1, "√", "")
            .TextMatrix(i, .ColIndex("缺省值")) = Nvl(mrsPars!缺省值)
            .TextMatrix(i, .ColIndex("影响控制说明")) = Nvl(mrsPars!影响控制说明)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(mrsPars!部门)
            .TextMatrix(i, .ColIndex("Fixed")) = Nvl(mrsPars!Fixed)
            .TextMatrix(i, .ColIndex("参数值含义")) = Nvl(mrsPars!参数值含义)
            .TextMatrix(i, .ColIndex("关联说明")) = Nvl(mrsPars!关联说明)
            .TextMatrix(i, .ColIndex("适用说明")) = Nvl(mrsPars!适用说明)
            .TextMatrix(i, .ColIndex("警告说明")) = Nvl(mrsPars!警告说明)
            .TextMatrix(i, .ColIndex("模块")) = Nvl(mrsPars!模块)
            .TextMatrix(i, .ColIndex("模块简码")) = Nvl(mrsPars!模块简码)
            If mlngParID = Val(Nvl(mrsPars!Id)) Then .Row = i: .TopRow = .Row
            Set .Cell(flexcpPicture, i, .ColIndex("标志")) = imgList.ListImages(mrsPars!ParType & "").Picture
            .Cell(flexcpPictureAlignment, i, .ColIndex("标志")) = 4
            If Val(mrsPars!Fixed & "") = 1 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000011 ' &H8000000F
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000012
            End If
            mrsPars.MoveNext
        Next
        .Redraw = flexRDBuffered
    End With
    LoadParas = True
End Function

Private Sub chkShowFixed_Click()
    Call SetParas
End Sub

Private Sub chkShowFixed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cboModule_Click()
    mlngModulePreIdx = cboModule.ListIndex
    If mblnNotClick Then Exit Sub
    If mlngModule <> cboModule.ItemData(cboModule.ListIndex) Then
        Call ResetCtrl(CT_Module)
    End If
    Call SetParas
    mlngModule = cboModule.ItemData(cboModule.ListIndex)
End Sub

Private Sub SetParas()
    '-----------------------------------------------------------------------------------------------------------
    '功能:显示和隐藏相关的参数信息
    '编制:刘兴洪
    '日期:2009-02-19 12:05:34
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim lngModule As Long
    Dim arrData As Variant
    Dim j As Long, lngRow As Long, blnOldRowShow As Boolean
    Dim strParType As String, blnShow As Boolean
    
    
    strParType = cboParType.List(cboParType.ListIndex)
    With vsPara
        .Redraw = flexRDNone
        If cboModule.ListIndex < 0 Then
            cboModule.ListIndex = 0: Exit Sub
        Else
            lngModule = cboModule.ItemData(cboModule.ListIndex)
        End If
        '展示列设置
        For j = 0 To .Cols - 1
           .ColHidden(j) = False
           If j = .ColIndex("模块名称") Then
               .ColHidden(j) = lngModule >= 0
           ElseIf j = .ColIndex("参数类型") Then
               .ColHidden(j) = strParType <> "所有类型"
            ElseIf j >= .ColIndex("影响控制说明") Then
                .ColHidden(j) = True
           End If
        Next
        lngRow = -1
        For i = 1 To .Rows - 1
           blnShow = True '
            If chkShowFixed.value = 0 And Val(.TextMatrix(i, .ColIndex("Fixed"))) = 1 Then
                '不显示固定参数
                blnShow = False
            End If
            If lngModule > -1 Then '只显示当前模块
                If Val(.TextMatrix(i, .ColIndex("模块"))) <> lngModule Then
                    blnShow = False
                End If
            End If
            '只显示当前类型
            If strParType <> "所有类型" Then
                If Trim(.TextMatrix(i, .ColIndex("参数类型"))) <> strParType Then
                    blnShow = False
                End If
            End If
            .RowHidden(i) = Not blnShow
            If lngRow <= 0 And .RowHidden(i) = False Then lngRow = i
        Next
        If lngRow > 0 Then
           If .RowHidden(.Row) = True Then .Row = lngRow
           If .Row = 0 Then .Row = lngRow
        End If
        If .RowHidden(.Row) Then vsPara.Row = 0
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetDetailPara(ByVal strParType As String)
'功能：展示分参数列
'参数：strParType=参数类型
    Dim i As Long, lngCol As Long, intDefIdx As Integer
    With vsDetailParas
        '展示列设置
        For i = 0 To .Cols - 1
            .ColHidden(i) = False
            If i = .ColIndex("机器名") Then '非本机参数隐藏机器名列
                .ColHidden(i) = Not strParType Like "*本机*"
            ElseIf i = .ColIndex("人员") Or i = .ColIndex("用户名") Then '部门参数与本机公共模块隐藏人员
                .ColHidden(i) = Not strParType Like "*私有*" Or strParType = "部门参数"
            ElseIf i = .ColIndex("站点") Then '不存在多站点，则不显示站点列
                .ColHidden(i) = Not mblnMultiSta
            ElseIf i >= .ColIndex("人员id") Then
                .ColHidden(i) = True
            End If
        Next
        If lblSearch.Tag = "" Then
            intDefIdx = MI_部门 '默认按部门搜索，因为部门列一直可见
            If strParType = "本机公共模块" Then
                intDefIdx = MI_机器名
            ElseIf strParType = "私有全局" Or strParType = "私有模块" Or strParType = "本机私有模块" Then
                intDefIdx = MI_人员
            End If
        Else
            intDefIdx = Val(lblSearch.Tag)
        End If
        '设置搜索属性
        For i = 0 To MI_参数值
            lngCol = Decode(i, MI_用户名, .ColIndex("用户名"), MI_人员, .ColIndex("人员"), MI_部门, .ColIndex("部门"), _
                                        MI_机器名, .ColIndex("机器名"), MI_站点, .ColIndex("站点"), MI_参数值, .ColIndex("参数值"))
            mnuPopuMenuSerch(i).Enabled = Not .ColHidden(lngCol)
            mnuPopuMenuSerch(i).Visible = Not .ColHidden(lngCol)
            mnuPopuMenuSerch(i).Checked = i = intDefIdx
        Next
        mnuPopuMenuSerch_Click (intDefIdx)
    End With
End Sub

Private Sub cboModule_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim i As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    strKey = Replace(cboModule.Text, "'", "")
    i = cboModule.ListIndex
    
    If cboModule.ListIndex >= 0 Then
        If cboModule.List(cboModule.ListIndex) = strKey Then
            SendKeys "{tab}"
            cboModule.ListIndex = i
            Exit Sub
        End If
    End If
    If strKey = "" Then SendKeys "{tab}": Exit Sub
    If ShowSelect(strKey) = False Then
        cboModule.SetFocus
        cboModule.ListIndex = i
        Exit Sub
    End If
End Sub

Private Function ShowSelect(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择相应的数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-20 16:50:12
    '-----------------------------------------------------------------------------------------------------------
  
    Dim lngLeft As Long, lngTop As Long, i As Long
    
    Dim vRect  As RECT
    Dim strSelect As String
    Dim sngHight As Single
    If mrsModule Is Nothing Then Exit Function
    
    
    mrsModule.Filter = 0
    mrsModule.Filter = "模块=" & IIf(Val(strKey) = 0, -22, Val(strKey)) & " Or 模块名称 like '%" & strKey & "%' or 模块简码 like '%" & UCase(strKey) & "%' "
    If mrsModule.RecordCount = 0 Then
        MsgBox "注意:" & vbCrLf & _
               "    没有找到满足条件的模块,请检查！", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If cboModule.Visible Then cboModule.SetFocus
        mrsModule.Filter = 0
        Exit Function
    End If
    If mrsModule.RecordCount = 1 Then GoTo SelOk:
    vRect = GetControlRect(cboModule.hwnd)
    sngHight = (IIf(mrsModule.RecordCount <= 2, 5, mrsModule.RecordCount) + 1) * 300
    If sngHight > Screen.Height - (vRect.Top + txtSearch.Height) Then
       If sngHight > vRect.Top Then
          sngHight = vRect.Top
          vRect.Top = 0
       Else
          vRect.Top = vRect.Top - sngHight
       End If
    Else
        vRect.Top = vRect.Top + cboModule.Height
    End If
    If frmSelectList.ShowSelect(Nothing, mrsModule, "模块,800,0,1;模块名称,2400,0,1;模块简码,1440,0,0", vRect.Left, vRect.Top, cboModule.Width * 2, sngHight, "", "系统模块", , strSelect, True) = False Then
        mrsModule.Filter = 0
        Exit Function
    End If
    If mrsModule.EOF Then
        mrsModule.Filter = 0
        Exit Function
    End If
SelOk:
    cboModule.ListIndex = mrsModule!Index
    mrsModule.Filter = 0
    ShowSelect = True
End Function

Private Sub cboSys_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cboParType_Click()
    If mblnNotClick Then Exit Sub
    If (cboParType.Text = "公共全局" Or cboParType.Text = "私有全局") Then
        chkShowFixed.value = 1
    ElseIf cboParType.Text = "所有类型" And chkShowFixed.value = 0 And chkShowFixed.Visible Then
        chkShowFixed.value = 1
    Else
        chkShowFixed.value = 0
    End If
    Call SetParas
End Sub

Private Sub cboParType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdModify_Click()
    '先进行登陆,
    Dim strUserName As String
    Dim strSystem As String, lng参数id As Long
    
    If cboSys.ListIndex < 0 Then Exit Sub
    strSystem = cboSys.ItemData(cboSys.ListIndex)
    With vsPara
        lng参数id = .RowData(.Row)
        If lng参数id = 0 Then Exit Sub
        If Val(vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("Fixed"))) = 1 Then Exit Sub
    End With
    If frmUserCheckLogin.ShowLogin(UCT_NormalUser, , strUserName, gstrServer, strSystem) = False Then Exit Sub
    If frmParaChangeSet.ShowEdit(Me, lng参数id, strUserName) = False Then Exit Sub
    mlngParID = lng参数id
    If cboModule.ListIndex <> 0 Then
        mlngModule = cboModule.ItemData(cboModule.ListIndex)
    End If
    mstrParType = cboParType.Text
    '需要重新设置当前行的参数信息
    cboSys.Tag = "强制刷新"
    Call cboSys_Click
    mlngParID = Val(vsPara.RowData(vsPara.Row))
    If cboModule.ListIndex <> 0 Then
        mlngModule = cboModule.ItemData(cboModule.ListIndex)
    End If
    mstrParType = cboParType.Text
    cboSys.Tag = ""
End Sub
 
Private Sub cmdSearch_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strDistinct As String, strCols As String, strFields As String
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    Dim strSelect As String
    Dim sngHight As Single, sgnWidth As Single
    
    Select Case Val(lblSearch.Tag)
        Case MI_用户名, MI_人员
            strDistinct = "用户名,人员,人员简码"
            strCols = "用户名,1000,0,1;人员,1500,0,1;人员简码,1000,0,1"
            strFields = IIf(Val(lblSearch.Tag) = MI_用户名, "用户名", "人员")
            sgnWidth = 3530
        Case MI_部门
            strDistinct = "部门,部门简码"
            strCols = "部门,1200,0,1;部门简码,1000,0,1"
            strFields = "部门"
            sgnWidth = 2230
        Case MI_机器名
            strDistinct = "机器名,机器名简码"
            strCols = "机器名,2000,0,1;机器名简码,1200,0,1"
            strFields = "机器名"
            sgnWidth = 3230
        Case MI_站点
            strDistinct = "站点"
            strCols = "站点,2000,0,1"
            strFields = "站点"
            sgnWidth = 2030
        Case MI_参数值
            strDistinct = "参数值"
            strCols = "参数值,2000,0,1"
            strFields = "参数值"
            sgnWidth = 2030
    End Select
    mrsDetailParas.Filter = ""
    Set rsTmp = RecDistinct(mrsDetailParas, strDistinct)
    If rsTmp.RecordCount = 0 Then
        MsgBox "注意:" & vbCrLf & _
               "    该参数无相关的用户、机器或部门参数设置,请检查！", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtSearch.Visible Then txtSearch.SetFocus
        Exit Sub
    End If
    
    vRect = GetControlRect(txtSearch.hwnd)
    sngHight = (IIf(rsTmp.RecordCount <= 7, 7, rsTmp.RecordCount) + 1) * 300
    If sngHight > Screen.Height - (vRect.Top + txtSearch.Height) Then
       If sngHight > vRect.Top Then
          sngHight = vRect.Top
          vRect.Top = 0
       Else
          vRect.Top = vRect.Top - sngHight
       End If
    Else
        vRect.Top = vRect.Top + txtSearch.Height
    End If
    If sgnWidth > Screen.Width - vRect.Left Then
        sgnWidth = Screen.Width - vRect.Left
    End If
    If frmSelectList.ShowSelect(Nothing, rsTmp, strCols, vRect.Left, vRect.Top, sgnWidth, sngHight, "", strFields & "选择", , strSelect, True) = False Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    txtSearch.Text = Nvl(rsTmp.Fields(strFields).value)
    If txtSearch.Visible Then txtSearch.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsModule Is Nothing Then
        If mrsModule.State = 1 Then
            mrsModule.Filter = 0
            mrsModule.Close
        End If
    End If
    If Not mrsDetailParas Is Nothing Then
        If mrsDetailParas.State = 1 Then mrsDetailParas.Close
    End If
    Set mrsModule = Nothing
    Set mrsDetailParas = Nothing
    Call SaveVsGridWidth(vsChangeLog, Me.Caption, "参数变动日志")
    Call SaveVsGridWidth(vsDetailParas, Me.Caption, "分站点及用户")
    Call SaveVsGridWidth(vsPara, Me.Caption, "系统参数列表")
    
End Sub

Private Function LoadDetailParas(ByVal lng参数id As Long, Optional blnSearch As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载指定参数的用户参数信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-17 14:58:37
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim int部门 As Integer, int本机 As Integer, int私有 As Integer, strOwner As String
    
    If blnSearch And mrsDetailParas Is Nothing Then
        mstrOwner = ""
        Set mrsDetailParas = GetDetailParas(lng参数id, mrsSys, int部门, int本机, int私有, mstrOwner)
        fraDetaisModi.Tag = int本机 & "," & int私有 & "," & int部门
    End If
    If blnSearch = False Then
        mstrOwner = ""
        Set mrsDetailParas = GetDetailParas(lng参数id, mrsSys, int部门, int本机, int私有, mstrOwner)
        fraDetaisModi.Tag = int本机 & "," & int私有 & "," & int部门
        If UCase(txtSearch.Text) <> "" Then blnSearch = True
    End If
    cmdAddNew.Visible = int部门 = 0
    cmdDel.Visible = int部门 = 0
    If blnSearch Then
        If UCase(txtSearch.Text) = "" Then
            mrsDetailParas.Filter = 0
        Else
            Select Case Val(lblSearch.Tag)
                Case MI_用户名, MI_人员
                    If Val(lblSearch.Tag) = MI_用户名 Then
                        mrsDetailParas.Filter = "用户名 like '" & UCase(txtSearch.Text) & "%'"
                    Else
                        mrsDetailParas.Filter = "人员 like '" & UCase(txtSearch.Text) & "%' OR 人员简码 like '" & UCase(txtSearch.Text) & "%'"
                    End If
                Case MI_部门
                    mrsDetailParas.Filter = "部门 like '" & UCase(txtSearch.Text) & "%' OR 部门简码 like '" & UCase(txtSearch.Text) & "%'"
                Case MI_机器名
                    mrsDetailParas.Filter = "机器名 like '" & UCase(txtSearch.Text) & "%' or 机器名简码 like '" & UCase(txtSearch.Text) & "%'"
                Case MI_站点
                    mrsDetailParas.Filter = "站点 like '" & UCase(txtSearch.Text) & "%'"
                Case MI_参数值
                    mrsDetailParas.Filter = "参数值 like '" & UCase(txtSearch.Text) & "%'"
            End Select
        End If
    End If
    mrsDetailParas.Sort = "站点,部门,人员,机器名"
    With vsDetailParas
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = ""
        .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgEdit.ListImages("UnCheck").Picture
        .ColData(.ColIndex("选择")) = 0
        .Cell(flexcpPictureAlignment, 0, .ColIndex("选择")) = flexAlignCenterCenter
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(mrsDetailParas.RecordCount = 0, 1, mrsDetailParas.RecordCount) + 1
        .Tag = 0 '记录选择条数
        .RowData(0) = mrsDetailParas.RecordCount '记录总条数
        i = 1
        Do While Not mrsDetailParas.EOF  '
            .RowData(i) = Nvl(mrsDetailParas!参数id)
            .TextMatrix(i, .ColIndex("站点")) = Nvl(mrsDetailParas!站点)
            .TextMatrix(i, .ColIndex("部门id")) = Nvl(mrsDetailParas!部门id)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(mrsDetailParas!部门)
            .TextMatrix(i, .ColIndex("部门简码")) = Nvl(mrsDetailParas!部门简码)
            .TextMatrix(i, .ColIndex("用户名")) = Nvl(mrsDetailParas!用户名)
            .TextMatrix(i, .ColIndex("人员id")) = Nvl(mrsDetailParas!人员id)
            .TextMatrix(i, .ColIndex("人员")) = Nvl(mrsDetailParas!人员)
            .TextMatrix(i, .ColIndex("人员简码")) = Nvl(mrsDetailParas!人员简码)
            .TextMatrix(i, .ColIndex("机器名")) = Nvl(mrsDetailParas!机器名)
            .TextMatrix(i, .ColIndex("机器名简码")) = Nvl(mrsDetailParas!机器名简码)
            .TextMatrix(i, .ColIndex("参数值")) = Nvl(mrsDetailParas!参数值)
            i = i + 1
            mrsDetailParas.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Call SetModValue
    LoadDetailParas = True
End Function

Private Function LoadChangeLog(ByVal lng参数id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载指定参数的用户参数信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-17 14:58:37
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
     
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parachangedlog", lng参数id)
    '参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因
    With vsChangeLog
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = ""
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF  '
            .RowData(i) = Nvl(rsTemp!参数id)
            .TextMatrix(i, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .TextMatrix(i, .ColIndex("变动说明")) = Nvl(rsTemp!变动说明)
            .TextMatrix(i, .ColIndex("变动内容")) = Nvl(rsTemp!变动内容)
            .TextMatrix(i, .ColIndex("变动人")) = Nvl(rsTemp!变动人)
            .TextMatrix(i, .ColIndex("变动时间")) = Format(rsTemp!变动时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("变动原因")) = Nvl(rsTemp!变动原因)
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadChangeLog = True
End Function

Private Function GetParaType(ByVal lng模块 As Long, ByVal int私有 As Integer, ByVal int本机 As Integer, ByVal int部门 As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取参数类型
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-17 16:44:21
    '-----------------------------------------------------------------------------------------------------------
    If int部门 = 1 Then GetParaType = "部门参数": Exit Function
    If lng模块 = 0 Then
        '不存模块,证明只有两种类型:公共全局和私有全局
        GetParaType = IIf(int私有 = 0, "公共全局", "私有全局")
        Exit Function
    End If
    '对模块的处理
    If int本机 = 0 Then
        '不是本机的情况,只有两种类型:公共模块和私有模块
         GetParaType = IIf(int私有 = 0, "公共模块", "私有模块")
         Exit Function
    End If
    '对本机的模块进行处理也有两种情况:
    GetParaType = IIf(int私有 = 0, "本机公共模块", "本机私有模块")
End Function

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    If fraSplit.Tag = "" Then '默认上六下四
        fraSplit.Top = (Me.ScaleHeight - picPara.Top) * 0.6 + picPara.Top
    End If
    fraSplit.Width = Me.ScaleWidth - fraSplit.Left + 100
    picPara.Height = fraSplit.Top - picPara.Top - 30
    picPara.Width = Me.ScaleWidth - picPara.Left
    picPage.Top = fraSplit.Top + fraSplit.Height + 30
    picPage.Height = Me.ScaleHeight - picPage.Top
    picPage.Width = Me.ScaleWidth - picPage.Left
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then fraSplit.Top = fraSplit.Top + y
End Sub

Private Sub fraSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If fraSplit.Top - picPara.Top < 1000 Then fraSplit.Top = picPara.Top + 1000
    If fraSplit.Top > picPage.Height + picPage.Top - 1500 Then fraSplit.Top = picPage.Height + picPage.Top - 1500
    fraSplit.Tag = "拖动"
    Call Form_Resize
End Sub

Private Sub lblSearch_Click()
    Dim i As Long
    '设置搜索属性
    For i = 0 To MI_参数值 '更新勾选状态
        mnuPopuMenuSerch(i).Checked = i = Val(lblSearch.Tag)
    Next
    PopupMenu Me.mnuPopuMenu, , picPage.Left + 30, picPage.Top + picDetailParas.Top + lblSearch.Top + lblSearch.Height + 30
End Sub

Private Sub mnuPopuMenuSerch_Click(Index As Integer)
    lblSearch.Caption = mnuPopuMenuSerch(Index).Caption & "↓"
    lblSearch.Tag = Index
    mnuPopuMenuSerch(Index).Checked = True
End Sub

Private Sub picDetailParas_Resize()
    err = 0: On Error Resume Next
    With picDetailParas
        fraDetaisModi.Left = 30: fraDetaisModi.Top = .ScaleTop
        fraDetaisModi.Width = .ScaleWidth - fraDetaisModi.Left
        vsDetailParas.Move .ScaleLeft + 20, fraDetaisModi.Top + fraDetaisModi.Height + 50, .ScaleWidth - 20
        vsDetailParas.Height = .ScaleHeight - vsDetailParas.Top
    End With
End Sub

Private Sub picPage_Resize()
    err = 0: On Error Resume Next
    With picPage
        tbPage.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub picPara_Resize()
    err = 0: On Error Resume Next
    With picPara
        vsPara.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub
Private Sub picParaChangeLog_Resize()
    err = 0: On Error Resume Next
    With picParaChangeLog
        vsChangeLog.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub picParInfo_Resize()
    Dim lngWith As Long
    vsParaInfo.Width = picParInfo.ScaleWidth - vsParaInfo.Left
    vsParaInfo.Height = picParInfo.ScaleHeight - vsParaInfo.Top
    lngWith = vsParaInfo.Width - vsParaInfo.ColWidth(0) - 120
    If lngWith < 10 * Me.TextWidth("字") Then lngWith = 10 * Me.TextWidth("字")
    vsParaInfo.ColWidth(1) = lngWith
    Call vsParaInfo.AutoSize(1)
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    tbPage.Tag = Item.Index
End Sub

Private Sub txtSearch_Change()
    Call LoadDetailParas(0, True)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtSearch.Text = ""
     End If
End Sub

Private Sub vsDetailParas_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intDefIdx As Integer, strParType As String
    
    If OldCol = NewCol Then Exit Sub
    With vsDetailParas
        Select Case NewCol
            Case .ColIndex("用户名")
                Call mnuPopuMenuSerch_Click(MI_用户名)
            Case .ColIndex("人员")
                Call mnuPopuMenuSerch_Click(MI_人员)
            Case .ColIndex("部门")
                Call mnuPopuMenuSerch_Click(MI_部门)
            Case .ColIndex("机器名")
                Call mnuPopuMenuSerch_Click(MI_机器名)
            Case .ColIndex("站点")
                Call mnuPopuMenuSerch_Click(MI_站点)
            Case .ColIndex("参数值")
                Call mnuPopuMenuSerch_Click(MI_参数值)
            Case Else
                strParType = cboParType.Text
                intDefIdx = MI_部门 '默认按部门搜索，因为部门列一直可见
                If strParType = "本机公共模块" Then
                    intDefIdx = MI_机器名
                ElseIf strParType = "私有全局" Or strParType = "私有模块" Or strParType = "本机私有模块" Then
                    intDefIdx = MI_人员
                End If
                Call mnuPopuMenuSerch_Click(intDefIdx)
        End Select
        SetModValue
    End With
End Sub

Private Sub vsDetailParas_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsDetailParas.ColIndex("选择") Then
        Cancel = True
    End If
End Sub

Private Sub vsDetailParas_Click()
    If vsDetailParas.Col = 0 Then
        vsDetailParas.ExplorerBar = flexExNone
    Else
        vsDetailParas.ExplorerBar = flexExSort
    End If
    If vsDetailParas.Col = vsDetailParas.ColIndex("选择") Then
        Call SelDetailParas(vsDetailParas.MouseRow)
    End If
End Sub

Private Sub vsDetailParas_GotFocus()
    Call zl_VsGridGotFocus(vsDetailParas)
End Sub

Private Sub vsDetailParas_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsDetailParas)
End Sub

Private Sub vsDetailParas_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsPara_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng参数id As Long
    Dim blnShowDetail As Boolean
    
    With vsPara
        cmdModify.Enabled = Not (Val(.TextMatrix(.Row, .ColIndex("Fixed"))) = 1 Or Val(.TextMatrix(.Row, .ColIndex("部门"))) = 1)
        If OldRow = NewRow Then Exit Sub
        txtSearch.Text = ""
        lblSearch.Tag = ""
        lng参数id = .RowData(.Row)
        If .RowHidden(.Row) Then lng参数id = 0
        If .Row > 0 Then
            blnShowDetail = Not (InStr(1, "|公共全局|公共模块|", "|" & Trim(.TextMatrix(.Row, .ColIndex("参数类型")) & "|")) > 0)
        Else
            blnShowDetail = Not (InStr(1, "|公共全局|公共模块|", "|" & Trim(cboParType.Text) & "|") > 0)
        End If
        '填充参数说明信息
        vsParaInfo.Cell(flexcpFontBold, PR_影响控制说明, 0, PR_警告说明, 0) = True
        vsParaInfo.Cell(flexcpText, PR_影响控制说明, 1, PR_警告说明, 1) = "" '清空上次信息
        vsParaInfo.TextMatrix(PR_影响控制说明, 1) = .TextMatrix(.Row, .ColIndex("影响控制说明"))
        vsParaInfo.TextMatrix(PR_参数值含义, 1) = .TextMatrix(.Row, .ColIndex("参数值含义"))
        vsParaInfo.TextMatrix(PR_关联说明, 1) = .TextMatrix(.Row, .ColIndex("关联说明"))
        vsParaInfo.TextMatrix(PR_适用说明, 1) = .TextMatrix(.Row, .ColIndex("适用说明"))
        vsParaInfo.TextMatrix(PR_警告说明, 1) = .TextMatrix(.Row, .ColIndex("警告说明"))
        Call vsParaInfo.AutoSize(1)
        If blnShowDetail Then
            Call SetDetailPara(Trim(.TextMatrix(.Row, .ColIndex("参数类型"))))
            Call LoadDetailParas(lng参数id)
        End If
    End With
    Call LoadChangeLog(lng参数id)
    
    tbPage.Item(Pag_Computer).Visible = blnShowDetail
    If tbPage.Item(Val(tbPage.Tag)).Visible Then
        tbPage.Item(Val(tbPage.Tag)).Selected = True
    Else
        tbPage.Item(Pag_ParaInfo).Selected = True
    End If
    If vsPara.Visible And vsPara.Enabled Then vsPara.SetFocus
End Sub
 
Private Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid)
    '进入控件
    With vsGrid
         .SelectionMode = flexSelectionByRow
         .HighLight = flexHighlightAlways
         .BackColorSel = GRD_GOTFOCUS_COLORSEL
    End With
End Sub
Private Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid)
    '离开控件
    With vsGrid
         .SelectionMode = flexSelectionByRow
         .FocusRect = flexFocusHeavy
         .HighLight = flexHighlightAlways
         .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub
Private Sub vsPara_GotFocus()
    Call zl_VsGridGotFocus(vsPara)
End Sub

Private Sub vsPara_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsPara)
End Sub
Private Sub vsChangeLog_GotFocus()
     
    Call zl_VsGridGotFocus(vsChangeLog)
End Sub

Private Sub vsChangeLog_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsChangeLog)
End Sub

Private Sub SelDetailParas(Optional ByVal lngRow As Long)
'功能：批量选择vsDetailParas，或取消选择
'          lngRow=0-选择或取消选择所有行，>0选择或取消选择指定行
    Dim blnSel As Boolean, i As Long
    
    With vsDetailParas
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngRow = 0 Then
            blnSel = Val(.ColData(.ColIndex("选择"))) = 0
            .Cell(flexcpPicture, lngRow, .ColIndex("选择")) = imgEdit.ListImages(IIf(blnSel, "AllCheck", "UnCheck")).Picture
            .ColData(.ColIndex("选择")) = IIf(blnSel, 1, 0) '标记图标状态
            For i = .FixedRows To .Rows - 1
                If Val(.RowData(i)) <> 0 Then
                    .TextMatrix(i, .ColIndex("选择")) = IIf(blnSel, -1, 0)
                End If
            Next
            If blnSel Then
                .Tag = Val(.RowData(0))
            Else
                .Tag = 0
            End If
        Else
            If Val(.RowData(lngRow)) <> 0 Then
                blnSel = Val(.TextMatrix(lngRow, .ColIndex("选择"))) = 0
                .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSel, -1, 0)
                .Tag = (Val(.Tag) + IIf(blnSel, 1, -1))
                If Val(.Tag) = 0 Then '所有的都未选择，则将图标更新为批量未勾选
                    .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgEdit.ListImages("UnCheck").Picture
                    .ColData(.ColIndex("选择")) = 0
                ElseIf Val(.Tag) = Val(.RowData(0)) Then '所有的都选择，则将图标更新为批量勾选
                    .Cell(flexcpPicture, 0, .ColIndex("选择")) = imgEdit.ListImages("AllCheck").Picture
                    .ColData(.ColIndex("选择")) = 1
                End If
            End If
        End If
    End With
    Call SetModValue
End Sub

Private Sub SetModValue()
'设置修改参数值可见性
    Dim blnVisible As Boolean
    blnVisible = Val(vsDetailParas.Tag) <> 0
    If Not blnVisible And vsDetailParas.Row >= vsDetailParas.FixedRows Then
        blnVisible = Val(vsDetailParas.RowData(vsDetailParas.Row)) <> 0
    End If
    
    cmdModValue.Enabled = blnVisible
    cmdDel.Enabled = blnVisible
End Sub


