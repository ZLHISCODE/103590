VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicalStation 
   Caption         =   "体检工作管理"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmMedicalStation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11400
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "体检部门"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   2100
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   9210
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   30
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "接受"
               Key             =   "接受"
               Object.ToolTipText     =   "接受体检"
               Object.Tag             =   "接受"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "完成"
               Key             =   "完成"
               Object.ToolTipText     =   "完成"
               Object.Tag             =   "完成"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "填写"
               Key             =   "填写"
               Object.ToolTipText     =   "填写"
               Object.Tag             =   "填写"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "总检"
               Key             =   "总检"
               Object.ToolTipText     =   "总检"
               Object.Tag             =   "总检"
               ImageIndex      =   7
               Style           =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "主费"
               Key             =   "主费"
               Object.ToolTipText     =   "主费"
               Object.Tag             =   "主费"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "附费"
               Key             =   "附费"
               Object.ToolTipText     =   "附费"
               Object.Tag             =   "附费"
               ImageIndex      =   9
               Style           =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   1605
      Left            =   3510
      TabIndex        =   34
      Top             =   3150
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6330
      Top             =   4335
   End
   Begin VB.PictureBox picClass 
      Height          =   5220
      Left            =   135
      ScaleHeight     =   5160
      ScaleWidth      =   2730
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   2790
      Begin VB.CommandButton cmdKind 
         Caption         =   "&Z.自定义查询"
         Height          =   300
         Index           =   3
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1005
         Width           =   1785
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&C.完成体检"
         Height          =   300
         Index           =   2
         Left            =   135
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   690
         Width           =   1785
      End
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4545
         Left            =   105
         ScaleHeight     =   4515
         ScaleWidth      =   2640
         TabIndex        =   7
         Top             =   1335
         Width           =   2670
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   3
            Left            =   4875
            TabIndex        =   8
            Top             =   255
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnX3 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
            Begin VB.Line lnY3 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   2
            Left            =   4305
            TabIndex        =   9
            Top             =   -60
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnY2 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
            Begin VB.Line lnX2 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   1
            Left            =   5085
            TabIndex        =   10
            Top             =   570
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnX1 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
            Begin VB.Line lnY1 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1395
            Index           =   0
            Left            =   225
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1830
            _cx             =   3228
            _cy             =   2461
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
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
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnY0 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
            Begin VB.Line lnX0 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
         End
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&B.正在体检"
         Height          =   300
         Index           =   1
         Left            =   135
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   390
         Width           =   1785
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&A.等待体检"
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   1785
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   4485
      ScaleHeight     =   1725
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   5085
      Width           =   3975
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7035
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStation.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   300
      Left            =   3435
      TabIndex        =   12
      Top             =   2595
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   529
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.报告"
            Key             =   "报告"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.总检"
            Key             =   "总检"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3.费用"
            Key             =   "费用"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&4.历次"
            Key             =   "历次"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9405
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":29CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":38C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":403C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":4256
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":49D0
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":514A
            Key             =   "附费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":58C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5ADE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5CFE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8790
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5F1E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":613E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":635E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":6AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":7252
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":79CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":7BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":8360
            Key             =   "主费"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":8ADA
            Key             =   "附费"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9254
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":946E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":968E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBack 
      Height          =   1605
      Left            =   3360
      TabIndex        =   13
      Top             =   690
      Width           =   7485
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   1380
         Index           =   1
         Left            =   6240
         ScaleHeight     =   1380
         ScaleWidth      =   6090
         TabIndex        =   47
         Top             =   90
         Visible         =   0   'False
         Width           =   6090
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "团体名称:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   45
            TabIndex        =   55
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重庆中联信息产业公司"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   870
            TabIndex        =   54
            Top             =   60
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联 系 人:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   45
            TabIndex        =   53
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系电话:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   45
            TabIndex        =   52
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电子邮件:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   45
            TabIndex        =   51
            Top             =   915
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重庆中联信息产业公司"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   11
            Left            =   870
            TabIndex        =   50
            Top             =   330
            Width           =   1800
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重庆中联信息产业公司"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   12
            Left            =   870
            TabIndex        =   49
            Top             =   630
            Width           =   1800
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重庆中联信息产业公司"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   13
            Left            =   870
            TabIndex        =   48
            Top             =   915
            Width           =   1800
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1425
         Index           =   0
         Left            =   30
         ScaleHeight     =   1425
         ScaleWidth      =   7050
         TabIndex        =   14
         Top             =   135
         Width           =   7050
         Begin VB.PictureBox picPhoto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   4965
            ScaleHeight     =   1425
            ScaleWidth      =   1020
            TabIndex        =   42
            Top             =   0
            Width           =   1018
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   6030
            ScaleHeight     =   450
            ScaleWidth      =   540
            TabIndex        =   15
            Top             =   45
            Visible         =   0   'False
            Width           =   570
            Begin VB.Shape shpState 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   360
               Left            =   60
               Top             =   45
               Width           =   450
            End
            Begin VB.Label lblState 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "收"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   18
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   90
               TabIndex        =   16
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12345678"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   1
            Left            =   3750
            TabIndex        =   46
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "健康号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   3105
            TabIndex        =   45
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "139"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   10
            Left            =   2475
            TabIndex        =   44
            Top             =   345
            Width           =   270
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "联系电话:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   1650
            TabIndex        =   43
            Top             =   345
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "像片:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   4470
            TabIndex        =   41
            Top             =   45
            Width           =   450
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   9
            Left            =   870
            TabIndex        =   40
            Top             =   915
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体检套餐:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   39
            Top             =   915
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已婚"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   870
            TabIndex        =   33
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婚姻状况:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   32
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "工作单位:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   75
            TabIndex        =   28
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门 诊 号:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   45
            TabIndex        =   27
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体检日期:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   1650
            TabIndex        =   26
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "受检人员:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   25
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性    别:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   1650
            TabIndex        =   24
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "年  龄:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3105
            TabIndex        =   23
            Top             =   45
            Width           =   630
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓无名"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   870
            TabIndex        =   22
            Top             =   60
            Width           =   540
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "男"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   2475
            TabIndex        =   21
            Top             =   60
            Width           =   180
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "30"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   5
            Left            =   3750
            TabIndex        =   20
            Top             =   45
            Width           =   180
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2004-12-20"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   6
            Left            =   2475
            TabIndex        =   19
            Top             =   600
            Width           =   900
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "666666"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   7
            Left            =   870
            TabIndex        =   18
            Top             =   630
            Width           =   540
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重庆中联信息产业公司"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   8
            Left            =   870
            TabIndex        =   17
            Top             =   1170
            Width           =   1800
         End
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   3540
      Top             =   4995
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":98AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9C48
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9FE2
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A37C
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A716
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A9AC
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":AF46
            Key             =   "新开"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":B4E0
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":BA7A
            Key             =   "取消"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C014
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C5AE
            Key             =   "up"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C770
            Key             =   "down"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C932
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":CBC8
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   225
      TabIndex        =   35
      Top             =   6270
      Width           =   2925
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   75
         Picture         =   "frmMedicalStation.frx":CE5E
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   37
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.姓名"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   36
         Tag             =   "姓名"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   3255
      MousePointer    =   9  'Size W E
      Top             =   1410
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintList 
         Caption         =   "体检指引单(&L)"
      End
      Begin VB.Menu mnuFileRequest 
         Caption         =   "体检申请单(&H)"
      End
      Begin VB.Menu mnuFilePrintRequest 
         Caption         =   "项目申请单(&R)"
      End
      Begin VB.Menu mnuFilePrintBook 
         Caption         =   "体检报告书(&B)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileSendMail 
         Caption         =   "发送报告书(&I)"
      End
      Begin VB.Menu mnuFile_11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRptDesign 
         Caption         =   "报告设计(&D)"
         Begin VB.Menu mnuReportDesign 
            Caption         =   "报告封面(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuReportDesign 
            Caption         =   "报告结果(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuReportDesign 
            Caption         =   "报告总检(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&M)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuMedical 
      Caption         =   "体检(&T)"
      Begin VB.Menu mnuMedicalNew 
         Caption         =   "体检登记(&R)"
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "新增个人(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "新增团体(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "修改登记(&3)"
            Index           =   3
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "删除登记(&4)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuMedicalPhoto 
         Caption         =   "照片采集(&P)"
      End
      Begin VB.Menu mnuMedical_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalBegin 
         Caption         =   "接受体检(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuMedicalBeginCancel 
         Caption         =   "取消接受(&C)"
      End
      Begin VB.Menu mnuMedicalGroupIn 
         Caption         =   "人员报到(&I)"
      End
      Begin VB.Menu mnuMedicalGroupOut 
         Caption         =   "取消报到(&X)"
      End
      Begin VB.Menu mnuMedical_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalComplete 
         Caption         =   "完成体检(&E)"
      End
      Begin VB.Menu mnuMedicalCompleteCancel 
         Caption         =   "取消完成(&R)"
      End
      Begin VB.Menu mnuMedical_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalItems 
         Caption         =   "组别项目(&D)"
      End
      Begin VB.Menu mnuMedicalItemsAddtion 
         Caption         =   "人员项目(&A)"
      End
      Begin VB.Menu mnuMedical_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalGroupAdd 
         Caption         =   "添加人员(&N)"
      End
      Begin VB.Menu mnuMedicalGroupDelete 
         Caption         =   "移除人员(&D)"
      End
      Begin VB.Menu mnuMedical_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalDept 
         Caption         =   "执行调整(&M)"
      End
      Begin VB.Menu mnuMedicalCallBack 
         Caption         =   "复查项目(&K)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报告(&E)"
      Begin VB.Menu mnuReportWrite 
         Caption         =   "填写报告(&W)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuReportView 
         Caption         =   "查看报告(&V)"
      End
      Begin VB.Menu mnuReportWriteMuli 
         Caption         =   "批量调整(&B)"
      End
      Begin VB.Menu mnuReport_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportAddOutLine 
         Caption         =   "添加总检(&A)"
         Begin VB.Menu mnuReportAddOutLineCase 
            Caption         =   "<无可用总检>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuReportModifyOutLine 
         Caption         =   "修改总检(&M)"
      End
      Begin VB.Menu mnuReportDelOutLine 
         Caption         =   "删除总检(&D)"
      End
   End
   Begin VB.Menu mnuCharge 
      Caption         =   "费用(&C)"
      Begin VB.Menu mnuChargeMain 
         Caption         =   "生成主费用(&G)"
      End
      Begin VB.Menu mnuCharge_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChargeAdd 
         Caption         =   "增加附加费(&A)"
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "收费单据(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "记帐单据(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "零费耗用登记(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuChargeModify 
         Caption         =   "修改附加费(&M)"
      End
      Begin VB.Menu mnuChargeDelete 
         Caption         =   "删除附加费(&D)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowResult 
         Caption         =   "显示报告(&H)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPatientBrowse 
         Caption         =   "人员信息(&B)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "综合查询(&F)"
         Begin VB.Menu mnuViewFilterItem 
            Caption         =   "自定义..."
            Index           =   0
         End
         Begin VB.Menu mnuViewFilterItem 
            Caption         =   "-"
            Index           =   1
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mlngLoop As Long
Private mintIndex As Integer                                '当前的区域
Private mlngSvrKey(0 To 3)  As Long                     '用于保存各个区域选中的行关键字
Private mfrmActive As Object                            '子窗体对象
Private mlngDept As Long
Private mlngHideRows As Long
Private mblnNoAllowChange As Boolean
Private mobjCls As New clsCISWork
Private mclsCore As New clsCISCore
Private mlngCountTmr As Long
Private mstrPrivilege As String
Private mblnDataMoved As Boolean
Private mintSort As Integer

Public WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Public mbytPopMenu As Byte

Private mlng体检病历id As Long
Private mint正体检查询依据 As Integer
Private mstr正体检团体时间范围 As String

Private Type usrSaveInfo
    lng登记id As Long
    lng病人id As Long
    str组别 As String
End Type

Private usrSave As usrSaveInfo
Private mrsFind As New ADODB.Recordset
Private mstrSvrFind As String

'（２）自定义过程或函数************************************************************************************************

Private Function SelectPerson(ByVal blnSingle As Boolean) As Boolean
    '选中某一个受检人员
    Dim lngLoop As Long
    Dim lngTotal As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim blnFirst As Boolean
    
    On Error Resume Next
    
    blnFirst = True
    lngStart = 1
    lngEnd = vsf(mintIndex).Rows - 1
    
ReStart:

    For lngLoop = lngStart To lngEnd
    
        If blnSingle = True And Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "标志"))) = 0 Then
            vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
        End If
        
        If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "登记id"))) = usrSave.lng登记id Then
            
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "标志"))) = 1 Then
                vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
            End If
            
            If usrSave.str组别 = "" Then
                vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
                vsf(mintIndex).Row = lngLoop
                Exit For
            End If
            
            '展开
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "标志"))) = 2 Then
                
                '组别
                If vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "姓名")) = usrSave.str组别 Then
                    vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
                    
                    lngTotal = vsf(mintIndex).Rows - 1
                    If usrSave.lng病人id = 0 Then
                        vsf(mintIndex).Row = lngLoop
'                        Exit For
                    End If
                End If
            End If
        
        
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "标志"))) >= 98 Then
                If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "病人id"))) = usrSave.lng病人id Then
                    vsf(mintIndex).Row = lngLoop
                    Exit For
                End If
            End If
        
        End If
    Next
    
    If lngEnd <> lngTotal And blnFirst Then
        lngStart = lngEnd + 1
        lngEnd = lngTotal
        blnFirst = False
        GoTo ReStart
    End If
    
    vsf(mintIndex).ShowCell vsf(mintIndex).Row, vsf(mintIndex).Col
    
    SelectPerson = True
    
End Function

Private Function Collapsed(ByVal intIndex As Integer, ByVal bytMode As CollapsedSettings)
    Dim lngLoop As Long

    With vsf(intIndex)
        
        For lngLoop = 1 To .Rows - 1
            
            If bytMode = flexOutlineCollapsed Then
                If .IsCollapsed(lngLoop) = flexOutlineExpanded Then
                    .IsCollapsed(lngLoop) = flexOutlineCollapsed
                End If
            Else
                If .IsCollapsed(lngLoop) = flexOutlineCollapsed Then
                    .IsCollapsed(lngLoop) = flexOutlineExpanded
                End If
            End If
        
        Next
        
    End With
    
    Call InheritAppendSpaceRows(intIndex)
End Function

Public Sub FindLocation(ByVal str姓名 As String)
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '--------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    
    lngRow = vsf(mintIndex).FindRow(str姓名, , GetCol(vsf(mintIndex), "姓名"), , False)
    If lngRow <= 0 Then
        ShowSimpleMsg "没有找到符合要求的信息！"
    Else
        vsf(mintIndex).Row = lngRow
        vsf(mintIndex).ShowCell vsf(mintIndex).Row, vsf(mintIndex).Col
        
    End If
End Sub

Public Sub ActiveFormEnabled()
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    Call AdjustEnableState
    
errHand:
    
End Sub

Private Property Let AutoRefresh(vData As Boolean)
    '
    '功能:自动刷新
    '
    tmr.Enabled = vData
    
    If vData = True Then
        mlngCountTmr = 0
        tmr.Tag = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "自动刷新间隔", 5))
        tmr.Enabled = (Val(tmr.Tag) > 0)
    End If
End Property

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    picState.BorderStyle = 0
    
    mintIndex = 0
    mlngDept = 0
    mlngSvrKey(0) = 0
    mlngSvrKey(1) = 0
    mlngSvrKey(2) = 0
    mlngSvrKey(3) = 0
    mlng体检病历id = 0
    mstrSvrFind = ""
    
    Call ResetActiveForm
        
    picShow.BorderStyle = 0
    
    For mlngLoop = 0 To cmdKind.UBound
        cmdKind(mlngLoop).Left = 15
        cmdKind(mlngLoop).Height = 300
    Next
    
    strVsf = ",450,1,1,1,[性质];,255,4,1,1,[状态];,255,4,1,1,[报告];姓名,900,1,1,1,;门诊号,900,7,1,1,;健康号,810,1,1,1,;就诊卡号,900,1,1,1,;体检编号,990,1,1,1,;性别,450,1,1,1,;年龄,450,1,1,1,;婚姻状况,0,1,1,0,;体检单号,0,1,1,0,;次数,0,7,1,0,;上级id,0,1,1,0,;病人id,0,1,1,0,;登记id,0,1,1,0,;标志,0,1,1,0,;是否装载,0,1,1,0,"
    
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Set vsf(0).Cell(flexcpPicture, 0, 1) = ils13.ListImages("状态").Picture
    Set vsf(0).Cell(flexcpPicture, 0, 2) = ils13.ListImages("单据").Picture
    
    With vsf(0)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
    
    Call CreateVsf(vsf(1), strVsf)
    vsf(1).Cols = vsf(1).Cols + 1
    vsf(1).ColWidth(vsf(1).Cols - 1) = 15
    Set vsf(1).Cell(flexcpPicture, 0, 1) = ils13.ListImages("状态").Picture
    Set vsf(1).Cell(flexcpPicture, 0, 2) = ils13.ListImages("单据").Picture
    With vsf(1)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
        
    Call CreateVsf(vsf(2), strVsf)
    vsf(2).Cols = vsf(2).Cols + 1
    vsf(2).ColWidth(vsf(2).Cols - 1) = 15
    Set vsf(2).Cell(flexcpPicture, 0, 1) = ils13.ListImages("状态").Picture
    Set vsf(2).Cell(flexcpPicture, 0, 2) = ils13.ListImages("单据").Picture
    With vsf(2)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
    
    strVsf = ",255,4,1,1,[状态];登记id,0,1,1,0,;体检单号,990,1,1,1,;姓名,900,1,1,1,;性别,450,1,1,1,;年龄,450,1,1,1,;婚姻状况,0,1,1,0,;门诊号,900,7,1,1,;健康号,810,1,1,1,;出生日期,1500,1,1,0,;团体,1500,1,1,1,;病人id,0,1,1,0,"
    Call CreateVsf(vsf(3), strVsf)
    vsf(3).Cols = vsf(3).Cols + 1
    vsf(3).ColWidth(vsf(3).Cols - 1) = 15
    Set vsf(3).Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture

    Dim strStart As String
    Dim strEnd As String
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "待体检时间范围", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "待体检时间范围", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    vsf(0).Tag = strStart & "|" & strEnd
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检时间范围", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检时间范围", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    vsf(1).Tag = strStart & "|" & strEnd
    
    '团体缺省时间范围
    strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检团体时间范围", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检团体时间范围", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    mstr正体检团体时间范围 = strStart & "|" & strEnd
    
    '团体查找依据
    mint正体检查询依据 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检查询依据", "0"))
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "已完体检时间范围", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "已完体检时间范围", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    vsf(2).Tag = strStart & "|" & strEnd
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActive() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Active事件
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    tmr.Tag = "5"
    mlngCountTmr = 0
    
    gstrSQL = GetPublicSQL(SQL.体检部门清单, IIf(InStr(gstrPrivs, "所有科室") > 0, "所有", ""))
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    If rs.BOF Then
        ShowSimpleMsg "没有体检性质的部门，请在部门管理中设置！"
        Exit Function
    End If
    
    '绑定数据到控件中
    Call AddComboData(cboDept, rs)
    zlControl.CboLocate cboDept, UserInfo.部门ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    '3.读取注册表保存的数据
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
        
        mnuViewShowResult.Checked = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示报告", 0)) = 1)
        
        tmr.Tag = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "自动刷新间隔", 5))
        lbl(1).Caption = "&6." & (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", "姓名"))
        lbl(1).Tag = Mid(lbl(1).Caption, 4)

    End If
    
    '4.最后更新控件属性
    tmr.Enabled = (Val(tmr.Tag) > 0)
    
    Call RefreshQueryMenu
    
    InitActive = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InheritResetVsf(ByVal intIndex As Integer)
    '--------------------------------------------------------------------------------------------------------
    '继承ResetVsf过程
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next

    Call ResetVsf(vsf(intIndex))
    vsf(intIndex).Cell(flexcpFontBold, 1, 0, 1, vsf(intIndex).Cols - 1) = False
    
    Call InheritAppendSpaceRows(intIndex)
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";体检信息;") > 0 Then
        Call InheritResetVsf(0)
        Call InheritResetVsf(1)
        Call InheritResetVsf(2)
    End If
    
    If InStr(strMenuItem, ";结果;") > 0 Then
        
        On Error Resume Next
        
        For mlngLoop = 0 To lblValue.UBound
            lblValue(mlngLoop).Caption = ""
        Next
        
        picPhoto.Cls
        
'        mlng病人id = 0
'        mlng主页id = 0
'        mlng医嘱id = 0
'        mlng发送号 = 0
        
        picState.Visible = False
        
        On Error Resume Next
 
        Call mfrmActive.zlClearData
    End If
        
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 应用权限处理
    '参数： strPrivilege                    权限
    '------------------------------------------------------------------------------------------------------------------
'    strPrivilege = "所有科室;体检登记;开始体检;取消开始;完成体检;取消完成;体检项目;附加项目;添加成员;移除成员;填写报告;科室小结;填写总结;打印报告;综合查询;费用处理;未收费体检"
    
    mstrPrivilege = strPrivilege
    
    If InStr(strPrivilege, "打印报告") = 0 Then mnuFilePrintBook.Visible = False
    
    If InStr(strPrivilege, "报告设计") = 0 Then mnuFileRptDesign.Visible = False
    
    If InStr(strPrivilege, "体检登记") = 0 And _
        InStr(strPrivilege, "开始体检") = 0 And _
        InStr(strPrivilege, "取消开始") = 0 And _
        InStr(strPrivilege, "完成体检") = 0 And _
        InStr(strPrivilege, "取消完成") = 0 And _
        InStr(strPrivilege, "体检项目") = 0 And _
        InStr(strPrivilege, "附加项目") = 0 And _
        InStr(strPrivilege, "添加成员") = 0 And _
        InStr(strPrivilege, "移除成员") = 0 Then
        
        mnuMedical.Visible = False
    Else
        
        If InStr(strPrivilege, "体检登记") = 0 Then mnuMedicalNew.Visible = False
                
        If InStr(strPrivilege, "开始体检") = 0 Then
            mnuMedicalBegin.Visible = False
            mnuMedicalGroupIn.Visible = False
            mnuMedicalGroupOut.Visible = False
        End If
        
        If InStr(strPrivilege, "取消开始") = 0 Then mnuMedicalBeginCancel.Visible = False
        If InStr(strPrivilege, "完成体检") = 0 Then mnuMedicalComplete.Visible = False
        If InStr(strPrivilege, "取消完成") = 0 Then mnuMedicalCompleteCancel.Visible = False
        If InStr(strPrivilege, "体检项目") = 0 Then mnuMedicalItems.Visible = False
        If InStr(strPrivilege, "附加项目") = 0 Then mnuMedicalItemsAddtion.Visible = False
        If InStr(strPrivilege, "添加成员") = 0 Then mnuMedicalGroupAdd.Visible = False
        If InStr(strPrivilege, "移除成员") = 0 Then mnuMedicalGroupDelete.Visible = False
        
        Dim aryMenu As Variant
        
        aryMenu = Array(mnuMedicalNew, mnuMedical_0, mnuMedicalBegin, mnuMedicalBeginCancel, mnuMedicalGroupIn, mnuMedicalGroupOut, mnuMedical_1, mnuMedicalComplete, mnuMedicalCompleteCancel, mnuMedical_2, mnuMedicalItems, mnuMedicalItemsAddtion, mnuMedical_3, mnuMedicalGroupAdd, mnuMedicalGroupDelete)
        
        Call AdjustSplit(aryMenu)
        
    End If
    
    If InStr(strPrivilege, "填写报告") = 0 Then
        mnuReportWrite.Visible = False
        mnuReportWriteMuli.Visible = False
    End If
    
    If InStr(strPrivilege, "填写总结") = 0 Then
        mnuReportAddOutLine.Visible = False
'        mnuReportAgain.Visible = False
    End If
    
    mnuReport_1.Visible = mnuReportAddOutLine.Visible
    'mnuReport_3.Visible = mnuReportAddOutLine.Visible
    
    If InStr(strPrivilege, "费用处理") = 0 Then mnuCharge.Visible = False
    
    If InStr(strPrivilege, "综合查询") = 0 Then mnuViewFind.Visible = False
    
    mnuReportModifyOutLine.Visible = mnuReportAddOutLine.Visible
    mnuReportDelOutLine.Visible = mnuReportAddOutLine.Visible
            
    tbrThis.Buttons("接受").Visible = mnuMedicalBegin.Visible And mnuMedical.Visible
    tbrThis.Buttons("完成").Visible = mnuMedicalComplete.Visible And mnuMedical.Visible
    tbrThis.Buttons("填写").Visible = mnuReportWrite.Visible And mnuReport.Visible
    tbrThis.Buttons("总检").Visible = mnuReportAddOutLine.Visible And mnuReport.Visible
    tbrThis.Buttons("主费").Visible = mnuCharge.Visible
    tbrThis.Buttons("附费").Visible = mnuCharge.Visible
    
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("接受").Visible Or tbrThis.Buttons("完成").Visible
    tbrThis.Buttons("Split_3").Visible = tbrThis.Buttons("总检").Visible Or tbrThis.Buttons("填写").Visible
    tbrThis.Buttons("Split_4").Visible = tbrThis.Buttons("主费").Visible Or tbrThis.Buttons("附费").Visible
    tbrThis.Buttons("Split_5").Visible = tbrThis.Buttons("过滤").Visible

End Sub

Private Sub AdjustSplit(ByVal aryMenu As Variant)
    
    Dim lngLoop As Long
    Dim lngPos As Long
    Dim lngSvrPos As Long
        
    For lngLoop = 0 To UBound(aryMenu)
        
        If aryMenu(lngLoop).Visible Then
            
            lngPos = lngPos + 1
            
            If aryMenu(lngLoop).Caption = "-" Then
                If lngPos = 1 Then
                    aryMenu(lngLoop).Visible = False
                    lngPos = 0
                Else
                    
                    If lngSvrPos + 1 = lngPos Then
                        aryMenu(lngLoop).Visible = False
                    End If
                    
                    lngSvrPos = lngPos
                End If
                
            End If
        End If
    Next
    
    If lngSvrPos = lngPos And lngPos > 0 Then
        aryMenu(lngSvrPos).Visible = False
    End If
    
        
End Sub


Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mnuFilePrintView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileOutExcel.Enabled = True
    mnuFilePrintRequest.Enabled = True
    mnuFilePrintList.Enabled = True
    mnuFilePrintBook.Enabled = True
    mnuFileSendMail.Enabled = True
    
    mnuMedicalNew.Enabled = True
    
    mnuMedicalBegin.Enabled = True
    mnuMedicalBeginCancel.Enabled = True
    mnuMedicalComplete.Enabled = True
    mnuMedicalCompleteCancel.Enabled = True
    
    mnuMedicalGroupIn.Enabled = True
    mnuMedicalGroupOut.Enabled = True
    mnuMedicalGroupAdd.Enabled = True
    mnuMedicalGroupDelete.Enabled = True
    
    mnuMedicalItems.Enabled = True
    mnuMedicalItemsAddtion.Enabled = True
    
    mnuReportWrite.Enabled = True
    
    mnuReportAddOutLine.Enabled = True
    mnuReportModifyOutLine.Enabled = True
    mnuReportDelOutLine.Enabled = True
    
'    mnuReportAgain.Enabled = True
    mnuReportView.Enabled = True
    
    mnuReportWriteMuli.Enabled = True
        
    
    mnuChargeMain.Enabled = True
    mnuChargeAdd.Enabled = True
    mnuChargeModify.Enabled = True
    mnuChargeDelete.Enabled = True
    
    mnuViewPatientBrowse.Enabled = True
       
    Select Case mintIndex
    Case 0
        mnuFilePrintBook.Enabled = False
        mnuFilePrintRequest.Enabled = False
'        mnuFilePrintList.Enabled = False
        
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuMedicalCompleteCancel.Enabled = False
        
        mnuReportWrite.Enabled = False
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        mnuReportView.Enabled = False
            
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
            
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
    
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalBegin.Enabled = False
            mnuFilePrintList.Enabled = False
        Else
            '0-个人分组项;1-团体名称项;2-团体组别项;99-受检人员项
            Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
            Case 0
                mnuMedicalBegin.Enabled = False
                mnuFilePrintList.Enabled = False
            Case 2
                mnuMedicalBegin.Enabled = False
            Case 99
                If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "上级id"))) > 0 Then
                    mnuMedicalBegin.Enabled = False
                End If
            End Select
        End If
        
    Case 1
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        
        If vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[状态]")) = "完成" Then
            mnuReportWrite.Enabled = False
            mnuReportAddOutLine.Enabled = False
            mnuReportModifyOutLine.Enabled = False
            mnuReportDelOutLine.Enabled = False
            mnuReportView.Enabled = False
            mnuReportWriteMuli.Enabled = False
            'mnuReportAgain.Enabled = False
            
            mnuMedicalGroupIn.Enabled = False
            mnuMedicalGroupAdd.Enabled = False
            mnuMedicalGroupDelete.Enabled = False
            mnuMedicalItems.Enabled = False
            mnuMedicalItemsAddtion.Enabled = False
            
            mnuChargeMain.Enabled = False
            mnuChargeAdd.Enabled = False
            mnuChargeModify.Enabled = False
            mnuChargeDelete.Enabled = False
            
            mnuMedicalComplete.Enabled = False
        Else
            mnuMedicalCompleteCancel.Enabled = False
        End If
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalBeginCancel.Enabled = False
            mnuMedicalComplete.Enabled = False
            
            mnuReportWrite.Enabled = False
            mnuReportAddOutLine.Enabled = False
            mnuReportModifyOutLine.Enabled = False
            mnuReportDelOutLine.Enabled = False
            mnuReportView.Enabled = False
            mnuReportWriteMuli.Enabled = False
'            mnuReportAgain.Enabled = False
            
            mnuFilePrintBook.Enabled = False
            mnuFilePrintRequest.Enabled = False
            mnuFilePrintList.Enabled = False
            
            mnuViewPatientBrowse.Enabled = False
                
            mnuMedicalGroupIn.Enabled = False
            mnuMedicalGroupAdd.Enabled = False
            mnuMedicalGroupDelete.Enabled = False
            mnuMedicalItems.Enabled = False
            mnuMedicalItemsAddtion.Enabled = False
            
            mnuChargeMain.Enabled = False
            mnuChargeAdd.Enabled = False
            mnuChargeModify.Enabled = False
            mnuChargeDelete.Enabled = False
        End If
        
        Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
            Case 0               '0-个人分组项;1-团体名称项;2-团体组别项;98-非团体受检人员项;99-团体受检人员
                mnuMedicalBeginCancel.Enabled = False
                mnuMedicalComplete.Enabled = False
                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                mnuReportWriteMuli.Enabled = False
'                mnuReportAgain.Enabled = False
                
                mnuFilePrintBook.Enabled = False
                mnuFilePrintRequest.Enabled = False
                mnuFilePrintList.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                mnuMedicalItems.Enabled = False
                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 1
                                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupDelete.Enabled = False
'                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 2
                mnuMedicalBeginCancel.Enabled = False
                mnuMedicalComplete.Enabled = False
                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                mnuReportWriteMuli.Enabled = False
                
                mnuFilePrintBook.Enabled = False
                mnuFilePrintRequest.Enabled = False
                mnuFilePrintList.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                mnuMedicalItems.Enabled = False
'                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 98
            
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                'mnuMedicalItemsAddtion.Enabled = False
                mnuMedicalItems.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                
                If InStr(mstrPrivilege, "未收费体检") = 0 Then
                    If picState.Visible = False Then
                        mnuReportWrite.Enabled = False
                        mnuReportWriteMuli.Enabled = False
                        
                        mnuReportAddOutLine.Enabled = False
                        mnuReportModifyOutLine.Enabled = False
                        mnuReportDelOutLine.Enabled = False
                        
'                        mnuReportAgain.Enabled = False
                    End If
                End If
                
            Case 99
                mnuMedicalBeginCancel.Enabled = False
                
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                
                mnuMedicalItems.Enabled = False
        End Select
                
        Select Case tbs.SelectedItem.Key
            Case "报告"
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                
                On Error Resume Next
                Select Case Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "报告来源")))
                Case 1, 2
                    If mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "执行状态")) = "正在执行" Then
                        mnuReportWrite.Enabled = False
                        mnuReportView.Enabled = False
                    End If
                End Select
                
                If mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "[状态]")) = "" Then
                    mnuReportWrite.Enabled = False
                    mnuReportView.Enabled = False
                End If
                
                If Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "报告id"))) = 0 Then
                    mnuReportView.Enabled = False
                End If
                
                On Error GoTo 0
                
            Case "总检"
                
                mnuReportWrite.Enabled = False
                mnuReportView.Enabled = False
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                                
            Case "费用"
                mnuReportView.Enabled = False
                mnuReportWrite.Enabled = False
                                
            Case "概况"
                
                mnuReportWrite.Enabled = False
                mnuReportView.Enabled = False
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                mnuReportAddOutLine.Enabled = False
'                mnuReportAgain.Enabled = False
                
                For mlngLoop = 1 To mfrmActive.vsf.Rows - 1
                    If Val(mfrmActive.vsf.RowData(mlngLoop)) > 0 Then
                        If Val(mfrmActive.vsf.TextMatrix(mlngLoop, 4)) = 0 Then
                            Exit For
                        End If
                    End If
                Next
                
                If mlngLoop = mfrmActive.vsf.Rows Then mnuMedicalGroupIn.Enabled = False
                
        End Select
        
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id"))) > 0 Then
            mnuReportAddOutLine.Enabled = (mlng体检病历id = 0)
        End If
        
        If mnuReportModifyOutLine.Enabled Then
            mnuReportModifyOutLine.Enabled = (mlng体检病历id > 0)
            mnuReportDelOutLine.Enabled = mnuReportModifyOutLine.Enabled
        End If
        
    Case 2
        mnuFilePrintRequest.Enabled = False
        mnuFilePrintList.Enabled = False
        
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuReportWrite.Enabled = False
        
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
        
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
                        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalCompleteCancel.Enabled = False
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End If
        
        Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
        Case 0, 2
            mnuMedicalCompleteCancel.Enabled = False
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End Select
        
    Case 3
        mnuFilePrintRequest.Enabled = False
        mnuFilePrintList.Enabled = False
        
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuMedicalCompleteCancel.Enabled = False
        mnuReportWrite.Enabled = False
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
        
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
                        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End If
    End Select
    
    If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    If mintIndex <> 3 Then
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id"))) = 0 Then
            mnuViewPatientBrowse.Enabled = False
        End If
    End If
    Select Case tbs.SelectedItem.Key
        Case "报告"
                        
            On Error Resume Next
            
            If Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "报告id"))) = 0 Then
                mnuReportView.Enabled = False
            End If
            
            On Error GoTo 0
    Case Else
        mnuReportView.Enabled = False
    End Select
    
    mnuFileSendMail.Enabled = mnuFilePrintBook.Enabled
    mnuMedicalGroupOut.Enabled = mnuMedicalGroupDelete.Enabled
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("接受").Enabled = mnuMedicalBegin.Enabled
    tbrThis.Buttons("完成").Enabled = mnuMedicalComplete.Enabled
    tbrThis.Buttons("填写").Enabled = mnuReportWrite.Enabled
        
    tbrThis.Buttons("总检").Enabled = mnuReportAddOutLine.Enabled Or mnuReportModifyOutLine.Enabled Or mnuReportDelOutLine.Enabled
    
    tbrThis.Buttons("主费").Enabled = mnuChargeMain.Enabled
    tbrThis.Buttons("附费").Enabled = mnuChargeAdd.Enabled Or mnuChargeModify.Enabled Or mnuChargeDelete.Enabled

End Sub

Private Function ShowDeptCase(ByVal lngDept As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能;
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHand
    
    
    ShowDeptCase = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim lngIndex As Long
    Dim lngLoop As Long
    Dim lngCount(0 To 3) As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error Resume Next
    
    
    
    strSQL = GetPublicSQL(SQL.体检人数统计, "2'" & vsf(0).Tag & "'0")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 2, CDate(Split(vsf(0).Tag, "|")(0)), CDate(Split(vsf(0).Tag, "|")(1)), 0)
    If rs.BOF = False Then lngCount(0) = rs.Fields(0).Value

    
    strSQL = GetPublicSQL(SQL.体检人数统计, "4'" & vsf(1).Tag & "'1")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 4, CDate(Split(vsf(1).Tag, "|")(0)), CDate(Split(vsf(1).Tag, "|")(1)), 1)
    If rs.BOF = False Then lngCount(1) = rs.Fields(0).Value
    
    strSQL = GetPublicSQL(SQL.体检人数统计, "5'" & vsf(2).Tag & "'1")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 5, CDate(Split(vsf(2).Tag, "|")(0)), CDate(Split(vsf(2).Tag, "|")(1)), 1)
    If rs.BOF = False Then lngCount(2) = rs.Fields(0).Value

    
    Select Case mintIndex
    Case 0
        
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "没有等待的体检。"
        Else
            strInfo = strInfo & "有" & lngCount(mintIndex) & "个等待的体检。"
        End If
    Case 1
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "没有人正在体检。"
        Else
            strInfo = strInfo & "有" & lngCount(mintIndex) & "个人正在体检。"
        End If
    Case 2
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "没有已完成的体检。"
        Else
            strInfo = strInfo & "有" & lngCount(mintIndex) & "个已完成的体检。"
        End If
    Case 3
        If vsf(mintIndex).Rows = 2 And Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            strInfo = "没有查询到符合要求的体检。"
        Else
            strInfo = "共查询到" & vsf(mintIndex).Rows - 1 & "个符合要求的体检。"
        End If
    End Select
 
    cmdKind(0).Caption = "&A.等待体检(" & Lpad(lngCount(0), 4, " ") & " 人)"
    cmdKind(1).Caption = "&B.正在体检(" & Lpad(lngCount(1), 4, " ") & " 人)"
    cmdKind(2).Caption = "&C.完成体检(" & Lpad(lngCount(2), 4, " ") & " 人)"
    
    stbThis.Panels(2).Text = strInfo
End Sub

Private Function SaveRow(ByVal objVsf As Object) As String
    SaveRow = objVsf.RowData(objVsf.Row)
End Function

Private Sub InheritRestoreRow(ByVal objVsf As Object, ByVal strKey As String)
    '--------------------------------------------------------------------------------------------------------
    '功能:继承RestoreRow过程
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim int病人idCol As Integer
    Dim intRow As Integer
    
    On Error Resume Next
        
    Call RestoreRow(objVsf, Val(strKey))
    
    int病人idCol = GetCol(objVsf, "病人id")
    
    If int病人idCol > 0 Then
        For intRow = objVsf.Row To 1 Step -1
            If Val(objVsf.TextMatrix(intRow, int病人idCol)) <= 0 Then
                objVsf.IsCollapsed(objVsf.Row - 1) = flexOutlineExpanded
                Exit For
            End If
        Next
    End If

End Sub

Private Function CancelMedical(ByVal str体检号 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：取消开始体检
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    gstrSQL = "SELECT ID FROM 病人医嘱记录 WHERE 相关ID IS NULL AND 病人来源=4 AND 挂号单='" & str体检号 & "'"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    If rs.BOF Then Exit Function
    
    Do While Not rs.EOF
        
        '门诊病人作废的同时也回退
        strSQL(ReDimArray(strSQL)) = "ZL_体检登记记录_Cancel(" & rs("ID").Value & ")"
                
        rs.MoveNext
    Loop
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    CancelMedical = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function CheckExecuteState(ByVal str体检号 As String, ByVal lng病人id As Long) As Byte
    
    '------------------------------------------------------------------------------------------------------------------
    '功能：检查执行状态
    '返回:  1                   表示有正在执行的项目
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    If lng病人id = 0 Then
        '团体病人
        gstrSQL = " IN (SELECT A.病人id FROM 体检人员档案 A,体检登记记录 B WHERE A.登记id=B.ID AND B.体检号=[1])"
    Else
        '单个病人
        gstrSQL = "=[2]"
    End If
    
    gstrSQL = "SELECT 1 FROM 病人医嘱发送 WHERE 执行状态 = 3 AND  医嘱id IN (select ID from 病人医嘱记录 where 病人id " & gstrSQL & " and 病人来源 = 4 and 挂号单 = [1] and 医嘱状态 <> 4)"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str体检号, lng病人id)
    
    CheckExecuteState = IIf(rs.BOF, 0, 1)
    
End Function

Private Function MenuClick(ByVal strMenuItem As String, Optional ByVal lng文件种类id As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：数据编辑/处理
    '******************************************************************************************************************
    Dim lngKey As Long
    
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL As String
    Dim lng登记id As Long
    Dim str体检号 As String
    Dim lng病人id As Long
    Dim strPrompt As String
    Dim blnGroup As Boolean
    Dim strGroup As String
    Dim lngStop As Long
    Dim rsItems As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsNo As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim lngTmp As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    AutoRefresh = False

    Call SQLRecord(rsSQL)
        
    '退出处理(一般外层菜单可用状态应该是屏蔽了的)
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "参数设置", "体检过滤"
        
        '无处理
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
        If lngKey = 0 Then GoTo pointEnd
        
        '读取可能要用的结果值
        lng登记id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "登记id")))
        str体检号 = vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "体检单号"))
        If mintIndex <> 3 Then
            If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) = "" Then
                '具体体检人员
                lng病人id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id")))
            Else
                
                '团体显示行
                If mintIndex = 0 Then
                    If vsf(mintIndex).Row + 1 > vsf(mintIndex).Rows - 1 Then
                        ShowSimpleMsg "当前团体没有设置体检人员！"
                        GoTo pointEnd
                    End If
                End If
            End If
        
            blnGroup = False
            blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id"))) = 0)
            If blnGroup = False Then
                blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志"))) = 1)
            End If
        End If
    End Select
    
    '第一步处理
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "体检登记"
        
        Select Case lng文件种类id
        Case 0      '新增个人
            
            If Not frmScheduleEdit.ShowEdit(Me, 0, mlngDept, , 2) Then GoTo pointEnd
            
        Case 1      '新增团体
            
            If Not frmScheduleEdit.ShowEdit(Me, 0, mlngDept, True, 2) Then GoTo pointEnd
            
        Case 3      '修改登记
            
            If lng登记id = 0 Then Exit Function
            
            If blnGroup = False Then
                blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "上级id"))) > 0)
            End If
            
            If Not frmScheduleEdit.ShowEdit(Me, lng登记id, mlngDept, blnGroup, 2) Then GoTo pointEnd
            
        Case 4      '删除登记
            
            If lng登记id = 0 Then Exit Function
            
            If MsgBox("你真的要删除当前体检登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
            strSQL = "ZL_体检登记记录_DELETE(" & lng登记id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
            
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case "照片采集"
        
        If lng登记id > 0 Then
            Call frmPersonPhoto.ShowEdit(Me, lng登记id, lng病人id)
            GoTo pointEnd
        End If
     '------------------------------------------------------------------------------------------------------------------
    Case "体检指引单", "项目申请单"
        
        If lng登记id > 0 Then
            Call frmMedicalStationPrintRpt.ShowEdit(Me, lng登记id, lng病人id, strMenuItem)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "体检申请单"
    
        If lng登记id > 0 Then
            Call frmMedicalStationRequest.ShowEdit(Me, lng登记id, lng病人id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "体检报告单"
    
        If lng登记id > 0 Then
            Call frmMedicalStationRptBook.ShowEdit(Me, lng登记id, lng病人id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "发送邮件"
        
        If lng登记id = 0 Then GoTo pointEnd
        
        If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) <> "" Then
            Call frmMedicalStationSendMail.ShowEdit(Me, lng登记id)
        Else
            Call frmMedicalStationSendMail.ShowEdit(Me, lng登记id, lng病人id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "参数设置"
    
        If Not frmMedicalStationPara.ShowPara(Me) Then GoTo pointEnd
    
            Dim strStart As String
        Dim strEnd As String
        
        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "待体检时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "待体检时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        vsf(0).Tag = strStart & "|" & strEnd
        
        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        vsf(1).Tag = strStart & "|" & strEnd
        
        '团体缺省时间范围
        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检团体时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检团体时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        mstr正体检团体时间范围 = strStart & "|" & strEnd
        
        '团体查找依据
        mint正体检查询依据 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "正体检查询依据", "0"))
                
        strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "已完体检时间范围", "今  天"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "已完体检时间范围", "今  天"), 2)
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        vsf(2).Tag = strStart & "|" & strEnd
        
    '------------------------------------------------------------------------------------------------------------------
    Case "体检过滤"
        
        strTmp = vsf(0).Tag & "'" & vsf(1).Tag & "'" & vsf(2).Tag & "'" & mstr正体检团体时间范围 & "'" & mint正体检查询依据
        If frmMedicalStationSearch.ShowFilter(Me, strTmp) = False Then GoTo pointEnd
        
        vsf(0).Tag = Split(strTmp, "'")(0)
        vsf(1).Tag = Split(strTmp, "'")(1)
        vsf(2).Tag = Split(strTmp, "'")(2)
        
        mstr正体检团体时间范围 = Split(strTmp, "'")(3)
        mint正体检查询依据 = Val(Split(strTmp, "'")(4))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "接受体检"             '接受体检
                        
        Select Case CheckAllowMedical(lng登记id)
        Case 1
            strPrompt = "当前体检还没有设置体检团体！"
        Case 2
            strPrompt = "当前体检还没有设置体检人员！"
        Case 3
            strPrompt = "当前体检的体检项目不完整（每种组别必须有体检项目）！"
        Case 4
            strPrompt = "存在没有分组的体检人员，请先在预约管理中进行人员组别划分！"
        End Select
        
        If strPrompt <> "" Then
            ShowSimpleMsg strPrompt
            GoTo pointEnd
        End If
        
        If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) <> "" Then
            '团体
            If Not frmMedicalStationBegin.ShowEdit(Me, lng登记id) Then GoTo pointEnd
        Else
            If Not frmMedicalStationBegin.ShowEdit(Me, lng登记id, lng病人id) Then GoTo pointEnd
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "人员报到"
        
        If str体检号 = "" Then GoTo pointEnd
        
        If blnGroup Then
            '团体
            If Not frmMedicalStationBegin.ShowEdit(Me, lng登记id, , True) Then GoTo pointEnd
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "取消开始"
        If str体检号 = "" Then GoTo pointEnd
        
        If MsgBox("确认要取消当前正在的体检吗？" & vbCrLf & "如果有附加项目，在重新开始体检后需要重新添加。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
        strSQL = "ZL_体检登记记录_Cancel('" & str体检号 & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "填写报告"
        
        Call mfrmActive.zlMenuClick(Me, "填写报告", CStr(lngKey) & "'1")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "查看报告"
        
        Call mfrmActive.zlMenuClick(Me, "查看报告", CStr(lngKey) & "'1")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "填写总结"             '
        If lng病人id = 0 Or lng登记id = 0 Then GoTo pointEnd
        
        If mlng体检病历id > 0 Then GoTo pointEnd
        
        mlng体检病历id = EditPatientFile("", lng病人id, str体检号, 0, lng文件种类id, False, Me, , True, 2, 1)
        If mlng体检病历id = 0 Then GoTo pointEnd
        
        strSQL = "ZL_体检人员档案_总结(" & lng登记id & "," & lng病人id & "," & mlng体检病历id & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "修改总结"
        
        If mlng体检病历id = 0 Then GoTo pointEnd
        
        Call EditPatientFile(mlng体检病历id, lng病人id, str体检号, 0, , False, Me, , True, 2, 1)
    '------------------------------------------------------------------------------------------------------------------
    Case "删除总结"
                
        If mlng体检病历id = 0 Then GoTo pointEnd
        
        If MsgBox("是否删除该人员的体检总结？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then GoTo pointEnd
        strSQL = "ZL_体检人员档案_总结(" & lng登记id & "," & lng病人id & ",null)"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        strSQL = "zl_病人病历_DELETE(" & mlng体检病历id & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "批量填写"
        
        If Not frmMedicalStationAdjust.ShowEdit(Me, mstrPrivilege) Then GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "完成体检"
        
        If blnGroup Then

            If str体检号 = "" Then GoTo pointEnd
            
            '检查是否有正在完成的或没有填写报告的
            If CheckExecuteState(str体检号, 0) = 1 Then
                ShowSimpleMsg "此团体的体检人员还有正在执行的项目！"
                GoTo pointEnd
            End If
            
            If MsgBox("当前团体成员的体检都完成了吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_体检登记记录_Finish('" & str体检号 & "',0)"
            Call SQLRecordAdd(rsSQL, strSQL)
        Else
            If str体检号 = "" Or lng病人id = 0 Then GoTo pointEnd
                                    
            '检查是否有正在完成的或没有填写报告的
            If CheckExecuteState(str体检号, lng病人id) = 1 Then
                ShowSimpleMsg "此体检人员还有正在执行的项目！"
                GoTo pointEnd
            End If
            
            If MsgBox("当前体检人员的体检都完成了吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_体检登记记录_Finish('" & str体检号 & "'," & lng病人id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "取消完成"
                
        If blnGroup Then
            If str体检号 = "" Then GoTo pointEnd
                        
            If MsgBox("真的要取消当前团体已完成的体检？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_体检登记记录_CancelFinish('" & str体检号 & "',0)"
            Call SQLRecordAdd(rsSQL, strSQL)
        Else
            If str体检号 = "" Or lng病人id = 0 Then GoTo pointEnd
   
            If MsgBox("真的要取消当前人员已完成的体检？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_体检登记记录_CancelFinish('" & str体检号 & "'," & lng病人id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "复查项目"
        lngTmp = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
        
        Select Case lngTmp
        Case 1, 2
            With mfrmActive.Body
                lng病人id = Val(.RowData(.Row))
                lngTmp = Abs(Val(.TextMatrix(.Row, 4)))
            End With
        Case Else
            lngTmp = 1
        End Select
        If str体检号 = "" Or lng病人id = 0 Or lng登记id = 0 Then GoTo pointEnd
        
        Dim rsData As New ADODB.Recordset
        
        gstrSQL = GetPublicSQL(SQL.人员原始项目)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id, lng病人id)
        If ShowGrdFilter(Me, vsf(mintIndex), "名称,2700,0,0;类别,900,0,1;执行科室,1500,0,0;采集方式,1200,0,0;检验标本,1200,0,0;检查部位,1200,0,0", Me.Name & "\复查项目选择", "请从列表中选择要复查的体检项目。", rsData, rs, 8790, 4500, , , True) Then
            If rs.RecordCount > 0 Then
                
                '生成项目并发送项目
                Call InsertItems(rsSQL, rs, lng登记id, lng病人id, True)

            End If
        End If
            
    '------------------------------------------------------------------------------------------------------------------
    Case "个人项目"
        
        lngTmp = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
        
        Select Case lngTmp
        Case 1, 2
            With mfrmActive.Body
                lng病人id = Val(.RowData(.Row))
                lngTmp = Abs(Val(.TextMatrix(.Row, 4)))
            End With
        Case Else
            lngTmp = 1
        End Select
        If str体检号 = "" Or lng病人id = 0 Or lng登记id = 0 Then GoTo pointEnd

        Call MedicalItemsRecord(rsItems)
        
        gstrSQL = GetPublicSQL(SQL.人员体检项目)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id, lng病人id)
        Call WriteItems(rs, rsItems, 2)
       
        Select Case lngTmp
        Case 0
            If Not frmItemsEdit.ShowEdit(Me, lng登记id, rsItems, mlngDept, False, 1, lng病人id) Then GoTo pointEnd
            '处理已经删除的体检项目
            Call FilterRecord(rsItems, "删除='1'")
            Call DeleteItem(rsSQL, rsItems, str体检号, lng登记id, lng病人id)
    
            '处理新添加的体检项目
            Call FilterRecord(rsItems, "新加<>'1'")
            Call NewItem(rsSQL, rsItems, lng登记id, lng病人id)

        Case Else
            If Not frmItemsEdit.ShowEdit(Me, lng登记id, rsItems, mlngDept, False, 2, lng病人id) Then GoTo pointEnd
        
            '处理已经删除的体检项目
            Call FilterRecord(rsItems, "删除='1'")
            Call DeleteItems(rsSQL, rsItems, str体检号, lng登记id, lng病人id)
            
            '处理新添加的体检项目
            Call FilterRecord(rsItems, "新加<>'1'")
            Call InsertItems(rsSQL, rsItems, lng登记id, lng病人id)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case "体检项目"

        If str体检号 = "" Or lng登记id = 0 Then GoTo pointEnd
        
        Call MedicalItemsRecord(rsItems)
        
        '读取体检项目
        gstrSQL = GetPublicSQL(SQL.团体体检项目)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id)
        
        Call WriteItems(rs, rsItems, 1)
        If Not frmItemsEdit.ShowEdit(Me, lng登记id, rsItems, mlngDept, blnGroup, 2) Then GoTo pointEnd
            
        '处理已经删除的体检项目
        Call FilterRecord(rsItems, "删除='1'")
        Call DeleteItems(rsSQL, rsItems, str体检号, lng登记id)
                
        '处理新添加的体检项目
        Call FilterRecord(rsItems, "新加<>'1'")
        Call InsertItems(rsSQL, rsItems, lng登记id)
    '------------------------------------------------------------------------------------------------------------------
    Case "添加成员"
    
        If str体检号 = "" Or lng登记id = 0 Then GoTo pointEnd
        
        Dim intCount2 As Integer
        Dim str门诊号 As String
        
        Call MedicalItemsRecord(rsItems, 2)
        
        gstrSQL = GetPublicSQL(SQL.体检人员档案)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id)
        If WriteItems(rs, rsItems, 1, 2) = False Then Exit Function
        
        gstrSQL = "Select 合约单位id From 体检登记记录 Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id)
        If rs.BOF Then Exit Function
        
        If Not frmPersonEdit.ShowEdit(Me, lng登记id, rsItems, True, 2, zlCommFun.NVL(rs("合约单位id"), 0)) Then Exit Function
        
        '处理新添加的体检人员
        Call FilterRecord(rsItems, "新加<>'1'")
        If rsItems.RecordCount > 0 Then rsItems.MoveFirst
        
        Dim intCount As Integer
        Dim intCount1 As Integer
        Dim bytNew As Byte
        Dim lngCount As Long
        
        intCount = -1
        Do While Not rsItems.EOF
            
            Call SQLRecord(rsSQL)
            
            '检查出生日期
            If rsItems("出生日期") <> "" Then
                
                If CheckStrValid(rsItems("出生日期"), CHECKFORMAT.日期) = False Then
                    ShowSimpleMsg rsItems("姓名").Value & "的出生日期无效！"
                    Exit Function
                End If
            End If
            bytNew = 0
            lng病人id = rsItems("病人ID").Value
            If lng病人id = 0 Then
                bytNew = 1
                intCount = intCount + 1
                lng病人id = GetNextNo(1) + intCount
                
                rsItems("病人ID").Value = lng病人id
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(rsItems("门诊号").Value, 0) < 1 Then
                'lng门诊号 = GetNextNo(3) + intCount2
                str门诊号 = CStr(GetNextNo(3) + intCount2)
                intCount2 = intCount2 + 1
            Else
                str门诊号 = CStr(zlCommFun.NVL(rsItems("门诊号").Value, 0))
            End If
                        
            strSQL = "ZL_体检人员档案_INSERT(" & lng登记id & "," & _
                                                                IIf(lng病人id = 0, "NULL", lng病人id) & ",'" & _
                                                                rsItems("组别").Value & "','" & _
                                                                rsItems("姓名").Value & "','" & _
                                                                rsItems("身份证").Value & "','" & _
                                                                rsItems("性别").Value & "'," & _
                                                                IIf(rsItems("出生日期").Value = "", "NULL", "TO_DATE('" & rsItems("出生日期").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                                rsItems("婚姻状况").Value & "','" & _
                                                                rsItems("民族").Value & "','" & _
                                                                rsItems("国籍").Value & "','" & _
                                                                rsItems("学历").Value & "','" & _
                                                                rsItems("职业").Value & "','" & _
                                                                rsItems("联系人姓名").Value & "','" & _
                                                                rsItems("联系人电话").Value & "','" & _
                                                                rsItems("电子邮件").Value & "','" & _
                                                                rsItems("联系人地址").Value & "','" & _
                                                                rsItems("工作单位").Value & "','" & _
                                                                rsItems("年龄").Value & "'," & _
                                                                Val(str门诊号) & ",'" & _
                                                                rsItems("IC卡号").Value & "','" & _
                                                                rsItems("健康号").Value & "','" & _
                                                                rsItems("就诊卡号").Value & "',0,0,0," & bytNew & _
                                                                ",Null)"
            
            Call SQLRecordAdd(rsSQL, strSQL)

            Dim lngSendNo As Long
            Dim str采集No As String
            Dim strNO As String
            
            lngSendNo = GetNextNo(10)
                        
            '产生费用单据号
            strSQL = "Select b.ID,b.结算途径,b.采集方式id From 体检项目清单 b Where b.组别名称=[1] and b.登记id=[2]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsItems("组别").Value, lng登记id)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    
                    str采集No = ""
                    strNO = ""
                    
                    If zlCommFun.NVL(rs("结算途径").Value, 1) = 1 Then
                        '记帐
                        strNO = GetNextNo(14)
                    Else
                        strNO = GetNextNo(13)
                    End If
                    
                    If zlCommFun.NVL(rs("采集方式id").Value, 0) > 0 Then
                        '采集
                        If zlCommFun.NVL(rs("结算途径").Value, 1) = 1 Then
                            '记帐
                            str采集No = GetNextNo(14)
                        Else
                            str采集No = GetNextNo(13)
                        End If
                    End If
                    
                    strSQL = "ZL_体检项目医嘱_NO(" & zlCommFun.NVL(rs("ID").Value, 0) & "," & lng病人id & ",'" & strNO & "','" & str采集No & "')"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    
                    rs.MoveNext
                Loop
            End If
            
            blnTran = True
            gcnOracle.BeginTrans
            If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
            For lngCount = 1 To rsSQL.RecordCount
                Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                rsSQL.MoveNext
            Next
            Call SQLRecord(rsSQL)
    
            strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & "," & lng病人id & "," & mlngDept & ",NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '产生相关费用
            If MakeMedicalCharge(rsSQL, lng登记id) = False Then
                gcnOracle.RollbackTrans
                blnTran = False
                Exit Function
            End If
            
            strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & "," & lng病人id & "," & mlngDept & ",NULL,2)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gcnOracle.CommitTrans
            blnTran = False
            
            
            rsItems.MoveNext
        Loop
    '------------------------------------------------------------------------------------------------------------------
    Case "移除成员"
        
        If lng病人id = 0 Or lng登记id = 0 Then GoTo pointEnd

        If MsgBox("移除此体检人员的同时也将作废体检报告，确认吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd

        strSQL = "ZL_体检人员档案_DELETE(" & lng登记id & "," & lng病人id & ",1)"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "取消报到"
    
        If lng病人id = 0 Or lng登记id = 0 Then GoTo pointEnd

        If MsgBox("取消人员报到的同时作废体检报告，确认吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd

        strSQL = "ZL_体检人员档案_DELETE(" & lng登记id & "," & lng病人id & ",1,1)"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "执行调整"
        '
        If Not frmMedicalStationDept.ShowEdit(Me, lng登记id) Then GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "人员信息"
        If lng病人id = 0 Then GoTo pointEnd
        
        Dim strParam As String
        Dim varParam As Variant
        
        gstrSQL = GetPublicSQL(SQL.体检人员档案_单个)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id, lng病人id)
        
        If rs.BOF = False Then
            strParam = zlCommFun.NVL(rs("病人id").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("姓名").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("身份证").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("性别").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("出生日期").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("婚姻状况").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("民族").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("国籍").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("学历").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("职业").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("身份").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("联系人姓名").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("联系人电话").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("电子邮件").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("联系人地址").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("工作单位").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("年龄").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("健康号").Value)
                        
            If frmPatientEdit.ShowEdit(Me, strParam, (mintIndex = 1)) Then
                
                If mintIndex = 1 Then
                    varParam = Split(strParam, "'")
                    
                    strSQL = "ZL_体检人员档案_INSERT(" & lng登记id & "," & _
                                                    Val(varParam(0)) & "," & _
                                                    "NULL,'" & _
                                                    varParam(1) & "','" & _
                                                    varParam(2) & "','" & _
                                                    varParam(3) & "'," & _
                                                    IIf(varParam(4) = "", "NULL", "TO_DATE('" & varParam(4) & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                    varParam(5) & "','" & _
                                                    varParam(6) & "','" & _
                                                    varParam(7) & "','" & _
                                                    varParam(8) & "','" & _
                                                    varParam(9) & "','" & _
                                                    varParam(11) & "','" & _
                                                    varParam(12) & "','" & _
                                                    varParam(13) & "','" & _
                                                    varParam(14) & "','" & _
                                                    varParam(15) & "','" & _
                                                    varParam(16) & "'," & _
                                                    "NULL," & _
                                                    "NULL,'" & _
                                                    varParam(17) & "'," & _
                                                    "NULL," & _
                                                    "1,0,1,0,Null)"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            End If
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "生成主费"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "生成主费用")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "增加收费单据"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "收费单据")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "增加记帐单据"
            
        Call mfrmActive.zlMenuClick(Me, lngKey, "记帐单据")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "增加零费耗用登记"
        
        Call mfrmActive.zlMenuClick(Me, lngKey, "零费耗用登记")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "修改附费"
        
        Call mfrmActive.zlMenuClick(Me, lngKey, "修改附加费用")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "删除附费"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "删除附加费用")
        GoTo pointEnd
    End Select
    
    '第二步处理
    
    blnTran = True
    
    gcnOracle.BeginTrans
    
    If rsSQL.RecordCount > 0 Then
        zlCommFun.ShowFlash "正在处理...", Me
        DoEvents
        rsSQL.MoveFirst
    End If
    
    For lngLoop = 1 To rsSQL.RecordCount
                    
        If lngLoop > lngStop And lngStop > 0 Then
            lngStop = 0
            '必须停顿一秒钟,否则会导致病人医嘱状态主键重复(医嘱ID,操作时间)
            Sleep 1006
        End If
        
        Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
        rsSQL.MoveNext
    Next
    gcnOracle.CommitTrans
    
    zlCommFun.StopFlash
    DoEvents
        
    blnTran = False
    
    If strMenuItem = "删除总结" Then mlng体检病历id = 0
    
    '刷新处理
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "取消完成"
        
        If CheckState(lng登记id, 2) And mintIndex = 2 Then
            Call mnuViewRefresh_Click
        Else
            Set vsf(mintIndex).Cell(flexcpPicture, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[状态]")) = ils13.ListImages("开始").Picture
            vsf(mintIndex).Cell(flexcpText, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[状态]")) = "开始"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "完成体检"
        
        '检查是否全部完成
        If CheckState(lng登记id) Then
            Call mnuViewRefresh_Click
        Else
            Set vsf(mintIndex).Cell(flexcpPicture, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[状态]")) = ils13.ListImages("完成").Picture
            vsf(mintIndex).Cell(flexcpText, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[状态]")) = "完成"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "取消报到", "移除成员"
        
        vsf(mintIndex).RemoveItem vsf(mintIndex).Row
        Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    '------------------------------------------------------------------------------------------------------------------
    Case "体检项目", "个人项目", "批量填写", "填写报告"
    
        Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    '------------------------------------------------------------------------------------------------------------------
    Case "体检复查"
    
        Call RefreshData("基本")
    '------------------------------------------------------------------------------------------------------------------
    Case "删除登记", "修改登记"
        
        Call mnuViewRefresh_Click
    '------------------------------------------------------------------------------------------------------------------
    Case "人员信息"
        
        If mintIndex = 1 Then
            Call RefreshData("体检")
            Call RefreshData("基本")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case Else
    
        Call mnuViewRefresh_Click
        
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    AutoRefresh = True
    Exit Function
    
pointEnd:
    AutoRefresh = True
    
    
    Exit Function
    
errHand:
    zlCommFun.StopFlash
    DoEvents
    
    AutoRefresh = True
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Function CheckState(ByVal lng登记id As Long, Optional ByVal bytMode As Byte = 1) As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    Select Case bytMode
    Case 1
        strSQL = "Select 1 From 体检登记记录 Where 体检状态<>5 And ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng登记id)
        CheckState = rs.BOF
    Case 2
        strSQL = "Select 1  From 体检登记记录 Where 体检状态=4 And ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng登记id)
        CheckState = (rs.BOF = False)
    End Select
    
    
    
    
End Function

Private Function DeleteItems(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal str体检号 As String, ByVal lng登记id As Long, Optional ByVal lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '作废此体检项目所产生的医嘱

            If lng病人id > 0 Then
            
                strSQL = "ZL_体检登记记录_ItemCancel('" & str体检号 & "'," & Val(rs("清单id").Value) & ",NULL," & lng病人id & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                strSQL = "ZL_体检项目清单_DELETE(" & lng登记id & ",NULL," & Val(rs("清单id").Value) & "," & lng病人id & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            Else
                
                strSQL = "ZL_体检登记记录_ItemCancel('" & str体检号 & "'," & Val(rs("清单id").Value) & ",'" & rs("组别").Value & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
                strSQL = "ZL_体检项目清单_DELETE(" & lng登记id & ",'" & rs("组别").Value & "'," & Val(rs("清单id").Value) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
            
            rs.MoveNext
        Loop
    End If
    
    DeleteItems = True
    
End Function

Private Function InsertItems(ByRef rsSQL As ADODB.Recordset, _
                            ByVal rs As ADODB.Recordset, _
                            ByVal lng登记id As Long, _
                            Optional ByVal lng病人id As Long = 0, _
                            Optional ByVal blnCallBack As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  新加入体检项目
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim lngSendNo As Long
    Dim lngKey As Long
    Dim str采集No As String
    Dim strNO  As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim lngCount As Long
    Dim rsSQLTmp As New ADODB.Recordset
    Dim lng清单id As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        lngSendNo = GetNextNo(10)
        
        Do While Not rs.EOF
            
            Call SQLRecord(rsSQLTmp)
            
            
            strTmp = ""
            
            If blnCallBack = False Then
                varRow = Split(rs("计费明细").Value, ";")
                For lngLoop = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngLoop), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                    
                Next
            End If
            
            '将此体检项目产生为医嘱
            If lng病人id > 0 Then
                
                lngKey = zlDatabase.GetNextId("体检项目清单")
                
                If blnCallBack Then
                    lng清单id = Val(rs("清单id").Value)
                Else
                    lng清单id = 0
                End If
                
                strSQL = "ZL_体检项目清单_INSERT(" & lng登记id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    zlCommFun.NVL(rs("体检类型").Value) & "'," & _
                                                    Val(rs("基本价格").Value) & "," & _
                                                    Val(rs("体检价格").Value) & "," & _
                                                    Val(rs("执行科室id").Value) & "," & _
                                                    IIf(zlCommFun.NVL(rs("采集方式id")) = "", "NULL", rs("采集方式id")) & "," & _
                                                    IIf(zlCommFun.NVL(rs("采集科室id")) = "", "NULL", rs("采集科室id")) & ",'" & _
                                                    zlCommFun.NVL(rs("检验标本").Value) & "','" & _
                                                    zlCommFun.NVL(rs("检查部位").Value) & "','" & _
                                                    zlCommFun.NVL(rs("检查部位id").Value) & "'," & lng病人id & "," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "'," & _
                                                    lngKey & "," & _
                                                    lng清单id & ")"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                '产生费用单据号
                str采集No = ""
                strNO = ""
                If zlCommFun.NVL(rs("结算方式").Value) = "记帐" Then
                    '记帐
                    strNO = GetNextNo(14)
                Else
                    strNO = GetNextNo(13)
                End If
                
                If Val(zlCommFun.NVL(rs("采集方式id").Value, 0)) > 0 Then
                    '采集
                    If zlCommFun.NVL(rs("结算方式").Value) = "记帐" Then
                        '记帐
                        str采集No = GetNextNo(14)
                    Else
                        str采集No = GetNextNo(13)
                    End If
                End If
                
                strSQL = "ZL_体检项目医嘱_NO(" & lngKey & "," & lng病人id & ",'" & strNO & "','" & str采集No & "')"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & "," & lng病人id & "," & mlngDept & "," & lngKey & ",1)"
                Call SQLRecordAdd(rsSQLTmp, strSQL)

                blnTran = True
                gcnOracle.BeginTrans
                
                If rsSQLTmp.RecordCount > 0 Then rsSQLTmp.MoveFirst
                For lngCount = 1 To rsSQLTmp.RecordCount
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQLTmp("SQL").Value), Me.Caption)
                    rsSQLTmp.MoveNext
                Next
                
                '产生相关费用
                If MakeMedicalCharge(rsSQL, lng登记id) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If
                
                strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & "," & lng病人id & "," & mlngDept & "," & lngKey & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
                
            Else
                
                lngKey = zlDatabase.GetNextId("体检项目清单")
                
                strSQL = "ZL_体检项目清单_INSERT(" & lng登记id & ",'" & _
                                                    rs("组别").Value & "'," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("体检类型").Value & "'," & _
                                                    Val(rs("基本价格").Value) & "," & _
                                                    Val(rs("体检价格").Value) & "," & _
                                                    Val(rs("执行科室id").Value) & "," & _
                                                    IIf(rs("采集方式id") = "", "NULL", rs("采集方式id")) & "," & _
                                                    IIf(rs("采集科室id") = "", "NULL", rs("采集科室id")) & ",'" & _
                                                    zlCommFun.NVL(rs("检验标本").Value) & "','" & _
                                                    rs("检查部位").Value & "','" & _
                                                    rs("检查部位id").Value & "',0," & IIf(rs("结算方式").Value = "记帐", "1", "2") & ",'" & strTmp & "'," & lngKey & ")"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                
                '产生费用单据号
                strSQL = "Select a.病人id From 体检人员档案 a Where a.组别名称=[1] and a.登记id=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rs("组别").Value, lng登记id)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        str采集No = ""
                        strNO = ""
                        
                        If zlCommFun.NVL(rs("结算方式").Value) = "记帐" Then
                            '记帐
                            strNO = GetNextNo(14)
                        Else
                            strNO = GetNextNo(13)
                        End If
                        
                        If Val(zlCommFun.NVL(rs("采集方式id").Value, 0)) > 0 Then
                            '采集
                            If zlCommFun.NVL(rs("结算方式").Value) = "记帐" Then
                                '记帐
                                str采集No = GetNextNo(14)
                            Else
                                str采集No = GetNextNo(13)
                            End If
                        End If
                        
                        strSQL = "ZL_体检项目医嘱_NO(" & lngKey & "," & rsTmp("病人id").Value & ",'" & strNO & "','" & str采集No & "')"
                        Call SQLRecordAdd(rsSQLTmp, strSQL)
                        
                        rsTmp.MoveNext
                    Loop
                End If
                
                strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & ",NULL," & mlngDept & "," & lngKey & ",1)"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                blnTran = True
                gcnOracle.BeginTrans
                If rsSQLTmp.RecordCount > 0 Then rsSQLTmp.MoveFirst
                For lngCount = 1 To rsSQLTmp.RecordCount
                    
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQLTmp("SQL").Value), Me.Caption)
                    rsSQLTmp.MoveNext
                Next
                
                '产生相关费用
                If MakeMedicalCharge(rsSQL, lng登记id) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If

                strSQL = "zl_体检人员档案_Accept(" & lng登记id & "," & lngSendNo & ",NULL," & mlngDept & "," & lngKey & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
                
            End If

            rs.MoveNext
        Loop
    End If
    
    InsertItems = True
    
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "", Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngKey As Long
    Dim lngLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim varParam As Variant
    Dim strSQL As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strField As String
    Dim objStd As New StdPicture
    Dim strTmpFile As String
    Dim strTmp As String
        
    If strParam = "" Then strParam = "'"
    varParam = Split(strParam, "'")
    
    On Error GoTo errHand
    
    Select Case strMenuItem
        Case "基本"             '读取体检人员的基本信息
            lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
            If lngKey <= 0 Then Exit Function

            
            If mintIndex <> 3 Then
                If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[性质]"))) <> "" Then
                                                
                    Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
                    Case 1, 2              '0-个人分组项;1-团体名称项;2-团体组别项;99-受检人员项
                        pic(1).Visible = True
                        pic(0).Visible = False
                        
                        strSQL = "Select a.* From 合约单位 a,体检登记记录 b Where a.ID=b.合约单位ID And b.ID=[1]"
                        
                        '数据转储处理
                        '----------------------------------------------------------------------------------------------
                        mblnDataMoved = False
                        If mintIndex = 2 Then mblnDataMoved = DataMove(lngKey)
                        If mblnDataMoved Then
                            strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
                        End If
                        
                        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "登记id"))))
                        If rs.BOF = False Then
                            lblValue(2).Caption = zlCommFun.NVL(rs("名称").Value)
                            lblValue(11).Caption = zlCommFun.NVL(rs("联系人").Value)
                            lblValue(12).Caption = zlCommFun.NVL(rs("电话").Value)
                            lblValue(13).Caption = zlCommFun.NVL(rs("电子邮件").Value)
                        End If
                    End Select
                    Exit Function
                End If
            End If
            
            pic(0).Visible = True
            pic(1).Visible = False

            If mintIndex <> 3 Then
                If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[性质]"))) <> "" Then Exit Function
            End If
            
            strSQL = GetPublicSQL(SQL.病人基本信息)
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            If mintIndex >= 2 Then mblnDataMoved = DataMove(Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)), 2)
            If mblnDataMoved Then
                strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
                strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
            If rs.BOF = False Then
                lblValue(3).Caption = zlCommFun.NVL(rs("姓名").Value)
                lblValue(4).Caption = zlCommFun.NVL(rs("性别").Value)
                lblValue(5).Caption = zlCommFun.NVL(rs("年龄").Value)
                lblValue(0).Caption = zlCommFun.NVL(rs("婚姻状况").Value)
                lblValue(6).Caption = Format(zlCommFun.NVL(rs("体检时间").Value), "yyyy-MM-dd")
                lblValue(7).Caption = zlCommFun.NVL(rs("门诊号").Value)
                lblValue(8).Caption = zlCommFun.NVL(rs("工作单位").Value)
                lblValue(1).Caption = zlCommFun.NVL(rs("健康号").Value)
                lblValue(9).Caption = zlCommFun.NVL(rs("体检类型").Value)
                lblValue(10).Caption = zlCommFun.NVL(rs("联系人电话").Value)
                mlng体检病历id = zlCommFun.NVL(rs("体检病历id").Value, 0)
                                
            End If
                                            
            picState.Visible = True
            strSQL = GetPublicSQL(SQL.个人费用概况)
                        
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            If mblnDataMoved Then
                '此时费用应是完全转出（对于抽回的都不考虑）
                strSQL = Replace(strSQL, "病人费用记录", "H病人费用记录")
                strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
                strSQL = Replace(strSQL, "体检登记记录", "H体检登记记录")
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            Else
                '此时可能费用已部份或完全转出
                gstrSQL = "Select a.体检时间 From 体检登记记录 a,体检人员档案 b Where a.ID=b.登记id And b.ID=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
                If rs.BOF = False Then
                    If zlDatabase.DateMoved(Format(rs("体检时间").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption) Then
                        strTmp = strSQL
                        strTmp = Replace(strTmp, "病人费用记录", "H病人费用记录")
                        strSQL = strSQL & " Union All " & strTmp
                    End If
                End If
            End If
            
            
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
            If CalcCharge(rsData, rs) Then
                picState.Visible = (Val(Format(zlCommFun.NVL(rs("未收金额").Value, 0), "0.00")) = 0)
            End If
                                
            '病人照片
            picPhoto.Cls
            strSQL = "Select B.* From 体检人员档案 A,病人照片 B Where A.病人id=B.病人id AND A.ID=[1]"
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            If mblnDataMoved Then
                strSQL = Replace(strSQL, "体检人员档案", "H体检人员档案")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                strTmpFile = ""
                strTmpFile = ReadPicture(rs, "照片", strTmpFile)
                
                If strTmpFile <> "" Then
                    Set objStd = VB.LoadPicture(strTmpFile)
                    Call DrawPicture(picPhoto, objStd, objStd.Width, objStd.Height)
                End If
            End If
            
            
        Case "预约"             '读取确认的预约体检的人员信息
                                    
            strStart = Split(vsf(0).Tag, "|")(0)
            strEnd = Split(vsf(0).Tag, "|")(1)
            
            strSQL = GetPublicSQL(SQL.体检登记单据)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 2)
            If rs.BOF = False Then Call LoadOutLineGrid(0, rs, , , ils13)
                        
        Case "体检"             '读取正在体检的人员信息
             
            strStart = Split(mstr正体检团体时间范围, "|")(0)
            strEnd = Split(mstr正体检团体时间范围, "|")(1)
            
'            strStart = Split(vsf(1).Tag, "|")(0)
'            strEnd = Split(vsf(1).Tag, "|")(1)
            
            strSQL = GetPublicSQL(SQL.体检登记单据, mint正体检查询依据)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 4)
            If rs.BOF = False Then Call LoadOutLineGrid(1, rs, , , ils13)
            
        Case "完成"             '读取体检完成的人员信息(段时间)
            
            strStart = Split(vsf(2).Tag, "|")(0)
            strEnd = Split(vsf(2).Tag, "|")(1)
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            mblnDataMoved = zlDatabase.DateMoved(Format(strStart, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
            mblnDataMoved = True
            
            strSQL = GetPublicSQL(SQL.体检登记单据, , mblnDataMoved)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 5)
            If rs.BOF = False Then Call LoadOutLineGrid(2, rs, , , ils13)
                        
        Case "查询"             '根据条件读取体检的人员信息
            
            Call InheritResetVsf(3)
            DoEvents
            
            Select Case Split(vsf(3).Tag, "^")(1)
            Case "指  定"
                strStart = Split(vsf(3).Tag, "^")(2)
            Case "所  有"
                strStart = ""
            Case Else
                strStart = GetDateTime(Split(vsf(3).Tag, "^")(1), 1)
            End Select
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            If strStart = "" Then
                mblnDataMoved = True
            Else
                mblnDataMoved = False
                mblnDataMoved = zlDatabase.DateMoved(Format(strStart, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
            End If
            
            strSQL = "SELECT B.ID," & _
                            "DECODE(B.体检状态,4,'开始',5,'完成') AS 状态," & _
                            "A.病人id," & _
                            "C.体检号 AS 体检单号," & _
                            "A.门诊号," & _
                            "A.姓名," & _
                            "A.性别," & _
                            "A.年龄," & _
                            "D.名称 AS 团体," & _
                            "to_char(A.出生日期,'yyyy-mm-dd') AS 出生日期," & _
                            "A.婚姻状况,B.登记id " & _
                        "FROM 病人信息 A,体检人员档案 B,体检登记记录 C,合约单位 D  " & _
                        "WHERE C.合约单位id=D.ID(+) AND B.体检状态=5 AND B.体检报到=1 AND A.病人ID=B.病人ID AND C.ID=B.登记id "
            
            strSQL = strSQL & GetQueryCondition(vsf(3).Tag)
            
            If mblnDataMoved Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "体检登记记录", "H体检登记记录")
                strTmp = Replace(strTmp, "体检人员档案", "H体检人员档案")
                strTmp = Replace(strTmp, "病人病历内容", "H病人病历内容")
                strTmp = Replace(strTmp, "病人医嘱记录", "H病人医嘱记录")
                strTmp = Replace(strTmp, "病人医嘱发送", "H病人医嘱发送")
                strTmp = Replace(strTmp, "病人病历所见单", "H病人病历所见单")
                strSQL = strSQL & " Union All " & strTmp
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then Call LoadGrid(vsf(3), rs, , , ils13)
            
        Case "组别人员"
                                    
            Dim blnField As Boolean
            Dim strIcon As String
            Dim intTmp As Integer
            
            '数据转储处理
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            If Val(varParam(3)) = 5 Then
                '已完成的体检业务
                
                If Val(varParam(1)) = 0 Then
                    '个人的体检业务
                    mblnDataMoved = zlDatabase.DateMoved(Format(Split(varParam(5), "|")(0), "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
                Else
                    '团体的某次体检业务
                    mblnDataMoved = DataMove(Val(varParam(1)))
                End If
            End If
            
            intTmp = 0
            If mintIndex = 0 Then
                strSQL = GetPublicSQL(SQL.体检组别人员1, Val(varParam(1)) & "'" & intTmp, mblnDataMoved)
            Else
                If mintIndex = 1 Then
                    intTmp = mint正体检查询依据
                End If
                strSQL = GetPublicSQL(SQL.体检组别人员, Val(varParam(1)) & "'" & intTmp, mblnDataMoved)
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(1)), CStr(varParam(2)), Val(varParam(3)), Val(varParam(4)), CDate(Split(varParam(5), "|")(0)), CDate(Split(varParam(5), "|")(1)))
            If rs.BOF = False Then
                
                intRow = Val(varParam(0))
                
                Do While Not rs.EOF
                    
                    intRow = intRow + 1
                    
                    vsf(mintIndex).AddItem "", intRow
                    
                    vsf(mintIndex).RowData(intRow) = rs("ID").Value
                    For intCol = 0 To vsf(mintIndex).Cols - 1
                    
                        strField = vsf(mintIndex).Cell(flexcpData, 0, intCol)
                        If strField <> "" And strField <> "是否装载" Then
                            If Left(strField, 1) <> "[" Then
                                vsf(mintIndex).TextMatrix(intRow, intCol) = zlCommFun.NVL(rs(strField))
                            Else
                                strField = Mid(strField, 2, Len(strField) - 2)
                                strIcon = ""
                                
                                On Error Resume Next
                                blnField = False
                                blnField = (UCase(rs(strField).Name) = UCase(strField))
                                If blnField Then
                                
                                    On Error GoTo errHand
                            
                                    strIcon = zlCommFun.NVL(rs(strField))
                                    If strIcon <> "" Then
                                        Set vsf(mintIndex).Cell(flexcpPicture, intRow, intCol) = ils13.ListImages(strIcon).Picture
                                    End If
                                    
                                    
                                    vsf(mintIndex).Cell(flexcpData, intRow, intCol) = strIcon
                                    vsf(mintIndex).TextMatrix(intRow, intCol) = strIcon
                                End If
                            End If
                        End If
                    Next
                    rs.MoveNext
                Loop
            End If
            
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetQueryCondition(ByVal strCondition As String, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strResult As String
     
   
    '以下是根据设置条件构成的条件语句
    
    If strCondition = "" Then Exit Function
    
    varTmp = Split(strCondition, "^")
    
    '体检部门
    If Val(varTmp(0)) > 0 Then strResult = strResult & " AND C.体检部门id + 0 = " & Val(varTmp(0))

    '体检时间
    If Trim(varTmp(1)) <> "所  有" Then
        Select Case Trim(varTmp(1))
        Case "指  定"
            strResult = strResult & " AND C.体检时间 BETWEEN TO_DATE('" & Format(varTmp(2), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(3), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            strResult = strResult & " AND C.体检时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(1), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(1), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    
    
    varTmp2 = Split(Trim(varTmp(4)), ",")
    strTmp = ""
    For mlngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(mlngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR C.体检号='" & varTmp2(mlngLoop) & "'"
        Else
            strTmp = strTmp & "  OR C.体检号 BETWEEN '" & Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1) & "' AND '" & Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1) & "'"
        End If
    Next
    If strTmp <> "" Then strResult = strResult & " AND (1=2 " & strTmp & ")"

    If Trim(varTmp(5)) <> "所  有" Then
        
        Select Case Trim(varTmp(5))
        Case "指  定"
            strResult = strResult & " AND B.完成时间 BETWEEN TO_DATE('" & Format(varTmp(6), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(7), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            strResult = strResult & " AND B.完成时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(5), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(5), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
        
    End If
    
    '病人姓名
    If Trim(varTmp(8)) <> "" Then strResult = strResult & " AND A.姓名 LIKE '%" & varTmp(8) & "%'"
    
    '体检团体
    If Val(varTmp(9)) > 0 Then strResult = strResult & " AND C.合约单位id = " & Val(varTmp(9))
    
    '体检项目及对比结果
    If Val(varTmp(11)) > 0 Then
        strResult = strResult & _
                    " AND (C.体检号,B.病人id) IN (SELECT E.挂号单,E.病人id " & _
                        "FROM 病人病历所见单 A, " & _
                             "诊治所见项目 B, " & _
                             "病人病历内容 C, " & _
                             "病人医嘱发送 D, " & _
                             "病人医嘱记录 E  " & _
                        "Where A.所见项ID = B.ID " & _
                              "AND A.病历id=C.ID " & _
                              "AND D.报告id=C.病历记录id " & _
                              "AND E.ID=D.医嘱ID " & _
                              "AND E.病人来源=4 " & _
                              "AND B.ID=" & Val(varTmp(11))
                
        If Val(varTmp(12)) = 0 Then
            strResult = strResult & " AND A.数值类型=0 AND DECODE(A.数值类型,0,TO_NUMBER(A.所见内容),0)"
            strTmp = Val(varTmp(15))
        Else
            strResult = strResult & " AND A.所见内容"
            strTmp = "'" & varTmp(15) & "'"
        End If
        
        Select Case varTmp(14)
        Case "大于"
            strResult = strResult & ">" & strTmp
        Case "小于"
            strResult = strResult & "<" & strTmp
        Case "大于等于"
            strResult = strResult & ">=" & strTmp
        Case "小于等于"
            strResult = strResult & "<=" & strTmp
        Case "不等于"
            strResult = strResult & "<>" & strTmp
        Case "包含"
            strResult = strResult & " LIKE '%" & varTmp(15) & "%'"
        Case "在范围内"
            If Val(varTmp(12)) = 0 Then
                strResult = strResult & " BETWEEN " & strTmp & " AND " & Val(varTmp(16))
            Else
                strResult = strResult & " BETWEEN " & strTmp & " AND '" & varTmp(16) & "'"
            End If
        Case Else
            strResult = strResult & "=" & strTmp
        End Select
        strResult = strResult & ")"
    End If
    
    GetQueryCondition = strResult
    
End Function


Private Sub InheritAppendSpaceRows(ByVal intIndex As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能：补齐表格空行
    '------------------------------------------------------------------------------------------------------------------
    Select Case intIndex
    Case 0
        Call AppendRows(vsf(intIndex), lnX0, lnY0)
    Case 1
        Call AppendRows(vsf(intIndex), lnX1, lnY1, mlngHideRows)
    Case 2
        Call AppendRows(vsf(intIndex), lnX2, lnY2)
    Case 3
        Call AppendRows(vsf(intIndex), lnX3, lnY3)
    End Select
End Sub

Private Sub ResetActiveForm()
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    
    If Not (mfrmActive Is Nothing) Then
        Unload mfrmActive
        Set mfrmActive = Nothing
    End If
        
End Sub

Private Sub PrintData(ByVal bytMode As Byte)
    '--------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '--------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
            
    mblnNoAllowChange = True
    
    If UserInfo.姓名 = "" Then Call GetUserInfo
        
    Select Case mintIndex
    Case 0
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "预约体检单"
    Case 1
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "正在体检的体检单"
    Case 2
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "已完成的体检单"
    Case 3
        Call CopyGrid(vsf(mintIndex), vsfPrint, 1)
        objPrint.Title = "查询体检单"
    End Select
    
    If mintIndex <> 3 Then
        Set objRow = New zlTabAppRow
        objRow.Add "体检部门:" & zlCommFun.GetNeedName(cboDept.Text)
        objRow.Add ""
        objPrint.UnderAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfPrint
        
    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)

    mblnNoAllowChange = False
End Sub

Private Sub RefreshQueryMenu()
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim strSectoin  As String
    Dim strTmp As String
    Dim lngLoop As Long
    
    For lngLoop = mnuViewFilterItem.UBound To 2 Step -1
        Unload mnuViewFilterItem(lngLoop)
    Next
    mnuViewFilterItem(1).Visible = False
    
    strSectoin = "私有模块\" & App.ProductName & "\过滤查找"
    
    For lngLoop = 1 To CLng(Val(GetSetting("ZLSOFT", strSectoin, "查找项数", "0")))
        
        strTmp = GetSetting("ZLSOFT", strSectoin, "过滤查找" & lngLoop, "")
        
        If Trim(strTmp) <> "" And InStr(strTmp, "|") > 0 Then
            mnuViewFilterItem(1).Visible = True
            Load mnuViewFilterItem(lngLoop + 1)
        
            mnuViewFilterItem(lngLoop + 1).Caption = Mid(strTmp, 1, InStr(strTmp, "|") - 1) & "(&" & lngLoop & ")"
            mnuViewFilterItem(lngLoop + 1).Tag = Mid(strTmp, InStr(strTmp, "|") + 1)
                        
        End If
    Next
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub cboDept_Click()
    Dim intIndex As Integer
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp Then Exit Sub
    If mlngDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngDept = cboDept.ItemData(cboDept.ListIndex)
    
    gstrSQL = "SELECT A.*,ROWNUM AS 序号 from " & _
                "(SELECT ID, 名称,种类 " & _
                "From 病历文件目录 " & _
                "Where 种类 = 1 " & _
                    "AND 应用 = 2 and ',' || 科室ID || ',' like '%," & mlngDept & ",%' " & _
            ") A " & _
            "ORDER BY A.ID"
    
    For intIndex = 1 To mnuReportAddOutLineCase.UBound
        Unload mnuReportAddOutLineCase(intIndex)
    Next
    mnuReportAddOutLineCase(0).Caption = "<无可用病历>"
    mnuReportAddOutLineCase(0).Tag = ""
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rs.BOF = False Then
        intIndex = 0
        Do While Not rs.EOF
            If intIndex > 0 Then
                Load mnuReportAddOutLineCase(intIndex)
                mnuReportAddOutLineCase(intIndex).Visible = True
            End If
            
            mnuReportAddOutLineCase(intIndex).Caption = zlCommFun.NVL(rs("名称").Value) & "(&" & rs("序号").Value & ")"
            mnuReportAddOutLineCase(intIndex).Tag = zlCommFun.NVL(rs("ID").Value)
            
            intIndex = intIndex + 1
            rs.MoveNext
        Loop
    End If
    
    Call mnuViewRefresh_Click

End Sub

Private Sub cmdKind_Click(Index As Integer)

    If mintIndex = Index Then
        vsf(Index).SetFocus
        Exit Sub
    End If
    
    mstrSvrFind = ""
    picShow.Tag = Index
    
    mintIndex = Index
    
    '1.调整界面布局
    For mlngLoop = cmdKind.LBound To cmdKind.UBound
        cmdKind(mlngLoop).Tag = IIf(mlngLoop <= Index, 0, 1)
    Next

    Call picClass_Resize
    Call picShow_Resize
    
    vsf(Index).SetFocus
    
    DoEvents
    
    mlngSvrKey(Index) = 0
    
    '清除右边区域
    Call ClearData("结果")
    Call vsf_AfterRowColChange(Index, 0, 0, vsf(Index).Row, vsf(Index).Col)
    
    Call AdjustEnableState
    Call RefreshStateInfo
End Sub

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 11 - 300)
    
    txt(1).Text = ""
    LocationObj txt(1)
    
End Sub

Private Sub mnuFilePrintBook_Click()
    Call MenuClick("体检报告单")
End Sub

Private Sub mnuFilePrintList_Click()
    Call MenuClick("体检指引单")
End Sub


Private Sub mnuFilePrintRequest_Click()
    Call MenuClick("项目申请单")
End Sub

Private Sub mnuFileRequest_Click()
    Call MenuClick("体检申请单")
End Sub

Private Sub mnuFileSendMail_Click()
    Call MenuClick("发送邮件")
End Sub

Private Sub mnuMedicalCallBack_Click()
    Call MenuClick("复查项目")
End Sub

Private Sub mnuMedicalDept_Click()
    Call MenuClick("执行调整")
End Sub

Private Sub mnuMedicalGroupDelete_Click()
    Call MenuClick("移除成员")
End Sub

Private Sub mnuMedicalGroupIn_Click()
    Call MenuClick("人员报到")
End Sub

Private Sub mnuMedicalGroupOut_Click()
    Call MenuClick("取消报到")
End Sub

Private Sub mnuMedicalNewType_Click(Index As Integer)
    Call MenuClick("体检登记", Index)
End Sub

Private Sub mnuMedicalPhoto_Click()
    Call MenuClick("照片采集")
End Sub

Private Sub mnuReportAddOutLineCase_Click(Index As Integer)
    Call MenuClick("填写总结", Val(mnuReportAddOutLineCase(Index).Tag))
End Sub

'Private Sub mnuReportAgain_Click()
'    Call MenuClick("体检复查")
'End Sub


Private Sub mnuReportDelOutLine_Click()
    Call MenuClick("删除总结")
End Sub

Private Sub mnuReportDesign_Click(Index As Integer)
        
    Select Case Index
    Case 0
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2_1", Me, True)
    Case 1
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2", Me, True)
    Case 2
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2_2", Me, True)
    End Select
End Sub

Private Sub mnuReportModifyOutLine_Click()
    Call MenuClick("修改总结")
End Sub

Private Sub mnuReportView_Click()
    Call MenuClick("查看报告")
End Sub


Private Sub mnuViewFilter_Click()
    Call MenuClick("体检过滤")
End Sub

Private Sub mnuViewFilterItem_Click(Index As Integer)
    Dim strCondition As String
        
    AutoRefresh = False
    If Index = 0 Then
        '定义查询
            
        strCondition = vsf(3).Tag
        If frmMedicalStationFilter.ShowEdit(Me, strCondition) Then
            cmdKind(3).Caption = "&Z.自定义查询"
            vsf(3).Tag = strCondition
            
            Call cmdKind_Click(3)
            
            zlCommFun.ShowFlash "请稍候，正在查询...", Me
            DoEvents
            
            Call RefreshData("查询")
            
            zlCommFun.StopFlash
            
            mintIndex = 1
            Call cmdKind_Click(3)
        End If
        
        Call RefreshQueryMenu
    Else
        '查询数据
        
        If mnuViewFilterItem(Index).Tag <> "" Then
            cmdKind(3).Caption = "&Z." & Mid(mnuViewFilterItem(Index).Caption, 1, Len(mnuViewFilterItem(Index).Caption) - 4)
            vsf(3).Tag = mnuViewFilterItem(Index).Tag
            
            Call cmdKind_Click(3)
            
            zlCommFun.ShowFlash "请稍候，正在查询...", Me
            DoEvents
            
            Call RefreshData("查询")
            
            zlCommFun.StopFlash
            
            mintIndex = 1
            Call cmdKind_Click(3)
        End If
        
    End If
    AutoRefresh = True
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    
    Call InitLoad
    Call InitSysPara
    Call ApplyPrivilege(gstrPrivs)
    
    Call InheritAppendSpaceRows(0)
    Call InheritAppendSpaceRows(1)
    Call InheritAppendSpaceRows(2)
    Call InheritAppendSpaceRows(3)
    
    DoEvents
    
    If InitActive = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call cboDept_Click
    Call tbs_Click          '此调用是为了刷新数据
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then imgY_S.Left = Me.ScaleWidth - 1000
        
    With picClass
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fraInfo.Height + 120
    End With
    
    With fraInfo
        .Left = 0
        .Top = picClass.Top + picClass.Height - 120
        .Width = picClass.Width
    End With
    
    With txt(1)
        .Width = fraInfo.Width - .Left - 75
    End With
    
    
    With imgY_S
        .Top = picClass.Top
        .Height = picClass.Height
    End With
    
    With fraBack
        .Left = imgY_S.Left + imgY_S.Width
        .Top = picClass.Top - 90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With tbs
        .Left = fraBack.Left
        .Top = fraBack.Top + fraBack.Height + 30
        .Width = fraBack.Width
    End With
    
    With picContainer
        .Left = tbs.Left
        .Top = tbs.Top + tbs.Height + 15
        .Width = tbs.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With pic(0)
        .Width = fraBack.Width - .Left - 30
    End With
    
    With picState
        .Left = pic(0).Width - .Width - 30
    End With
    
    pic(1).Move pic(0).Left, pic(0).Top, pic(0).Width, pic(0).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnStartUp Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(1).Tag)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示报告", IIf(mnuViewShowResult.Checked, 1, 0))
        
    Call SaveWinState(Me, App.ProductName)
    
    If mrsFind.State = adStateOpen Then mrsFind.Close
    Set mrsFind = Nothing
    
'    Set mobjCls = Nothing
'    Set mclsCore = Nothing
    
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 3000 Then imgY_S.Left = 3000
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub


Private Sub mnuChargeDelete_Click()
    Call MenuClick("删除附费")
End Sub

Private Sub mnuChargeMain_Click()
    Call MenuClick("生成主费")
End Sub

Private Sub mnuChargeModify_Click()
    Call MenuClick("修改附费")
End Sub

Private Sub mnuChargeAddType_Click(Index As Integer)
    Select Case Index
    Case 0
        Call MenuClick("增加收费单据")
    Case 1
        Call MenuClick("增加记帐单据")
    Case 2
        Call MenuClick("增加零费耗用登记")
    End Select
End Sub


Private Sub mnuMedicalCompleteCancel_Click()
    Call MenuClick("取消完成")
End Sub

Private Sub mnuMedicalComplete_Click()
    Call MenuClick("完成体检")
End Sub

Private Sub mnuMedicalBeginCancel_Click()
    Call MenuClick("取消开始")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuFilePara_Click()
    Call MenuClick("参数设置")
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuMedicalBegin_Click()
    Call MenuClick("接受体检")
End Sub


Private Sub mnuMedicalGroupAdd_Click()
    Call MenuClick("添加成员")
End Sub

Private Sub mnuMedicalItems_Click()
    Call MenuClick("体检项目")
End Sub

Private Sub mnuMedicalItemsAddtion_Click()
    
    If mintIndex <> 1 Then Exit Sub
    
    If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志"))) = 98 Then
        Call MenuClick("体检项目")
    Else
        Call MenuClick("个人项目")
    End If
    
End Sub


Private Sub mnuViewPatientBrowse_Click()
    Call MenuClick("人员信息")
End Sub


Private Sub mnuReportWrite_Click()
    Call MenuClick("填写报告")
End Sub

Private Sub mnuReportWriteMuli_Click()
    Call MenuClick("批量填写")
End Sub

Private Sub mnuViewRefresh_Click()
    
    Dim intRow As Integer
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim blnSingle As Boolean
    
    If mintIndex >= 3 Then Exit Sub
    
    zlCommFun.ShowFlash "请稍候，正在刷新数据...", Me
    DoEvents
    
    mblnNoAllowChange = True
    
    intRow = vsf(mintIndex).Row
    
    usrSave.lng登记id = Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "登记id")))
    usrSave.lng病人id = Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "病人id")))
    usrSave.str组别 = ""
    
    Select Case Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "标志")))
    Case 0, 98
        blnSingle = True
    End Select
    
    If Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "标志"))) > 1 Then
        
        gstrSQL = "Select 组别名称 From 体检人员档案 Where 登记id=[1] And 病人id=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, usrSave.lng登记id, usrSave.lng病人id)
        
        If rs.BOF = False Then
            usrSave.str组别 = rs("组别名称").Value
        End If
        
    End If
    
    LockWindowUpdate vsf(mintIndex).hWnd
    
    Call ClearData("体检信息;结果")
    
    Call RefreshData("预约")
    Call RefreshData("体检")
    Call RefreshData("完成")
    
    '选中某一个受检人员
    Call SelectPerson(blnSingle)
    
    Call InheritAppendSpaceRows(mintIndex)
    
    LockWindowUpdate 0
    
    zlCommFun.StopFlash
    
    mblnNoAllowChange = False
    
    mlngSvrKey(mintIndex) = -1
    Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub


Private Sub mnuViewShowResult_Click()
    
    On Error Resume Next
    
    mnuViewShowResult.Checked = Not mnuViewShowResult.Checked
    
    If tbs.SelectedItem.Key = "报告" Then
        mfrmActive.ShowResult = mnuViewShowResult.Checked
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        If mnuChargeAddType(0).Visible Then mobjPopMenu.Add 1, "增加" & mnuChargeAddType(0).Caption, , , mnuChargeAddType(0).Enabled
        If mnuChargeAddType(1).Visible Then mobjPopMenu.Add 2, "增加" & mnuChargeAddType(1).Caption, , , mnuChargeAddType(1).Enabled
        If mnuChargeAddType(2).Visible Then mobjPopMenu.Add 3, "增加" & mnuChargeAddType(2).Caption, , , mnuChargeAddType(2).Enabled

        mobjPopMenu.Add 4, "-", , 2, True
        
        mobjPopMenu.Add 5, mnuChargeModify.Caption, , , mnuChargeModify.Enabled
        mobjPopMenu.Add 6, mnuChargeDelete.Caption, , , mnuChargeDelete.Enabled
    Case 2
        
        For mlngLoop = 0 To mnuReportAddOutLineCase.UBound
            If mnuReportAddOutLineCase(mlngLoop).Caption = "<无可用总检>" Then
                If mnuReportAddOutLineCase(mlngLoop).Visible Then mobjPopMenu.Add mlngLoop + 1, mnuReportAddOutLineCase(mlngLoop).Caption, , , mnuReportAddOutLineCase(mlngLoop).Enabled And mnuReportAddOutLine.Enabled
            Else
                If mnuReportAddOutLineCase(mlngLoop).Visible Then mobjPopMenu.Add mlngLoop + 1, "增加" & mnuReportAddOutLineCase(mlngLoop).Caption, , , mnuReportAddOutLineCase(mlngLoop).Enabled And mnuReportAddOutLine.Enabled
            End If
        Next
        
        mobjPopMenu.Add mnuReportAddOutLineCase.UBound + 1, "-", , 2, True
        
        If mnuReportModifyOutLine.Visible Then mobjPopMenu.Add 101, mnuReportModifyOutLine.Caption, , , mnuReportModifyOutLine.Enabled
        If mnuReportDelOutLine.Visible Then mobjPopMenu.Add 102, mnuReportDelOutLine.Caption, , , mnuReportDelOutLine.Enabled
        
    Case 3
        
        mobjPopMenu.Add 1, "&1.姓名", , , True, , (lbl(1).Tag = "姓名")
        mobjPopMenu.Add 2, "&2.门诊号", , , True, , (lbl(1).Tag = "门诊号")
        mobjPopMenu.Add 3, "&3.健康号", , , True, , (lbl(1).Tag = "健康号")
        mobjPopMenu.Add 4, "&4.就诊卡号", , , True, , (lbl(1).Tag = "就诊卡号")
        mobjPopMenu.Add 5, "&5.姓名拼音", , , True, , (lbl(1).Tag = "姓名拼音")
        mobjPopMenu.Add 6, "&6.姓名五笔", , , True, , (lbl(1).Tag = "姓名五笔")
        mobjPopMenu.Add 7, "&7.身份证号", , , True, , (lbl(1).Tag = "身份证号")
            
        mobjPopMenu.Add 8, "-", , 2, True
        mobjPopMenu.Add 9, "&8.体检单号", , , True, , (lbl(1).Tag = "体检单号")
        mobjPopMenu.Add 10, "&9.体检编号", , , True, , (lbl(1).Tag = "体检编号")
        mobjPopMenu.Add 11, "&A.团体简码", , , True, , (lbl(1).Tag = "团体简码")
        
    Case 4          '费用
        If mnuCharge.Visible Then
            If mnuChargeMain.Visible Then mobjPopMenu.Add 1, mnuChargeMain.Caption, , , mnuChargeMain.Enabled
            
            If mnuChargeAddType(0).Visible Or mnuChargeAddType(1).Visible Or mnuChargeAddType(2).Visible Then
                mobjPopMenu.Add 2, "-", , 2, True
            End If
            
            If mnuChargeAddType(0).Visible Then mobjPopMenu.Add 3, "增加" & mnuChargeAddType(0).Caption, , , mnuChargeAddType(0).Enabled
            If mnuChargeAddType(1).Visible Then mobjPopMenu.Add 4, "增加" & mnuChargeAddType(1).Caption, , , mnuChargeAddType(1).Enabled
            If mnuChargeAddType(2).Visible Then mobjPopMenu.Add 5, "增加" & mnuChargeAddType(2).Caption, , , mnuChargeAddType(2).Enabled
            
        End If
    Case 5
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuChargeAddType_Click(0)
        Case 2
            Call mnuChargeAddType_Click(1)
        Case 3
            Call mnuChargeAddType_Click(2)
        Case 5
            Call mnuChargeModify_Click
        Case 6
            Call mnuChargeDelete_Click
        End Select
    Case 2
        If Key <= mnuReportAddOutLineCase.UBound + 1 Then
            Call mnuReportAddOutLineCase_Click(Key - 1)
            Exit Sub
        End If
        
        Select Case Key
        Case 101
            Call mnuReportModifyOutLine_Click
        Case 102
            Call mnuReportDelOutLine_Click
        End Select
    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    Case 4          '费用
        Select Case Key
        Case 1
            Call mnuChargeMain_Click
        Case 3
            Call mnuChargeAddType_Click(0)
        Case 4
            Call mnuChargeAddType_Click(1)
        Case 5
            Call mnuChargeAddType_Click(2)
        End Select
        
    Case 5
        
    End Select
End Sub

Private Sub picClass_Resize()
    Dim lngCount As Long
    Dim lngLoop As Long
    
    On Error Resume Next
    
    LockWindowUpdate picClass.hWnd
    
    lngCount = cmdKind.UBound - 1
    If cmdKind(3).Visible Then lngCount = cmdKind.UBound
    
    For lngLoop = cmdKind.LBound To lngCount
        cmdKind(lngLoop).Width = picClass.ScaleWidth
        If Val(cmdKind(lngLoop).Tag) = 0 Then
            cmdKind(lngLoop).Top = picClass.ScaleTop + 285 * lngLoop
            picShow.Top = picClass.ScaleTop + 285 * (lngLoop + 1)
        Else
            cmdKind(lngLoop).Top = picClass.ScaleHeight - 285 * (lngCount - lngLoop + 1)
        End If
    Next
    
    picShow.Left = picClass.ScaleLeft - 30
    picShow.Width = picClass.ScaleWidth + 60
    picShow.Height = picClass.ScaleHeight - 285 * (lngCount + 1) + 15
    
    LockWindowUpdate 0
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    If Not (mfrmActive Is Nothing) Then
        mfrmActive.Width = picContainer.Width
        mfrmActive.Height = picContainer.Height
    End If
End Sub

Private Sub picShow_Resize()
    
    On Error Resume Next
    
    vsf(0).Visible = False
    vsf(1).Visible = False
    vsf(2).Visible = False
    vsf(3).Visible = False
           
    vsf(Val(picShow.Tag)).Visible = True
    
    With vsf(Val(picShow.Tag))
        
        .Left = 0
        .Top = -15
        .Width = picShow.Width
        .Height = picShow.Height + 15
        
        Call InheritAppendSpaceRows(Val(picShow.Tag))
        
    End With
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "接受"
        Call mnuMedicalBegin_Click
    Case "完成"
        Call mnuMedicalComplete_Click
    Case "填写"
        Call mnuReportWrite_Click
 
    Case "总检"
        
        mbytPopMenu = 2
        
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
     Case "主费"
     
        Call mnuChargeMain_Click
        
    Case "附费"
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
                
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "过滤"
        
        If mnuViewFilter.Visible And mnuViewFilter.Enabled Then Call mnuViewFilter_Click
        
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    
    Call tbrThis_ButtonClick(Button)
    
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()
    Dim lngKey As Long
    Dim lngStyle As Long
    Dim blnShowed As Boolean
    Dim lng登记id As Long
    Dim str组别 As String
    
    blnShowed = False
    picContainer.BorderStyle = 0
    
    Select Case tbs.SelectedItem.Key
    Case "报告"
        If TypeName(mfrmActive) = "frmMedicalStationReport" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationReport
        End If
        
        mfrmActive.ShowResult = mnuViewShowResult.Checked
    Case "总检"
        
        Call ResetActiveForm
        
        picContainer.BorderStyle = 1
        
        Set mfrmActive = mclsCore.ShowFileObject(Me, picContainer, 0, 0, gcnOracle, "", glngSys, "", "")
        
        Call mfrmActive.zlMenuClick(Me, mlng体检病历id, "刷新")
        
        Call AdjustEnableState
        Exit Sub
    Case "费用"

        If TypeName(mfrmActive) = "frmMedicalStationCharge" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationCharge
        End If
    Case "历次"
        If TypeName(mfrmActive) = "frmMedicalStationHistory" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationHistory
        End If
    Case "概况"
        If TypeName(mfrmActive) = "frmMedicalStationGroup" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationGroup
        End If
    End Select
    
    '加载当前活动窗口
    
    If Not mfrmActive Is Nothing Then
        If blnShowed = False Then
            
            Load mfrmActive
            
            lngStyle = GetWindowLong(mfrmActive.hWnd, GWL_STYLE)
            Call SetWindowLong(mfrmActive.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
            Call SetParent(mfrmActive.hWnd, picContainer.hWnd)
            Call MoveWindow(mfrmActive.hWnd, 0, 0, picContainer.ScaleWidth / Screen.TwipsPerPixelX, picContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
            mfrmActive.Show
            DoEvents
            
        End If
        
        '刷新数据
        On Error Resume Next
        
        str组别 = ""
        lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
        
        If mintIndex <> 3 Then
            If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[性质]"))) <> "" Then
                lngKey = 0
                lng登记id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "登记id")))
            End If
            
            Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "标志")))
            Case 2               '0-个人分组项;1-团体名称项;2-团体组别项;99-受检人员项
            
                str组别 = vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "姓名"))
            End Select
            
        End If
        
        Select Case UCase(tbs.SelectedItem.Key)
        Case "报告"
            Call mfrmActive.zlMenuClick(Me, "刷新", CStr(lngKey) & "'" & mintIndex)
        Case "历次"
            
            Dim strStart As String
            Dim strEnd As String
            
            strStart = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "历次体检范围", "今  天"), 1)
            strEnd = GetDateTime(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "历次体检范围", "今  天"), 2)
            If strStart = "" Then strStart = GetDateTime("今  天", 1)
            If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
            
            Call mfrmActive.zlMenuClick(Me, "刷新", CStr(Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id")))) & "'" & strStart & "'" & strEnd)
            
        Case "总检", "费用"
            Call mfrmActive.zlMenuClick(Me, lngKey, "刷新")
        Case "概况"
            Call mfrmActive.zlMenuClick(Me, "刷新", CStr(lng登记id) & "'" & str组别)
        Case Else
            
        End Select
                
    End If
    
    Call AdjustEnableState
End Sub

Private Sub tmr_Timer()
    Dim strSvrKey As String
    
    mlngCountTmr = mlngCountTmr + 1
    
    If mlngCountTmr >= Val(tmr.Tag) Then
    
        '时间到了，开始触发
        mlngCountTmr = 0
        
        mblnNoAllowChange = True
        strSvrKey = SaveRow(vsf(mintIndex))
        
        LockWindowUpdate vsf(mintIndex).hWnd
        
        If mintIndex < 2 Then Call ClearData("体检信息;结果")
                
        Call RefreshData("预约")
        Call RefreshData("体检")
        
        Call InheritAppendSpaceRows(mintIndex)
                
        LockWindowUpdate 0
        
        mblnNoAllowChange = False
        
        If mintIndex < 2 Then
            
            Call InheritRestoreRow(vsf(mintIndex), Val(strSvrKey))
            
            mlngSvrKey(mintIndex) = -1
            Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
            
        End If
                
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    Dim strSQL As String
    Dim lngRow As Long
    Dim blnCard As Boolean
    Dim strStart As String
    Dim strEnd As String
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    strCol = Mid(lbl(1).Caption, 4)
    lngCol = GetCol(vsf(mintIndex), strCol)
            
    If strCol = "就诊卡号" And mintIndex <> 3 And KeyAscii <> vbKeyReturn Then
        '就诊卡号，自动识别

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn

        End If

    End If
            
    If KeyAscii = vbKeyReturn And mintIndex <> 3 Then
        
        If mintIndex = 1 And mint正体检查询依据 = 1 Then
            
            Select Case strCol
            Case "团体简码"
                strSQL = "Select * From (Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号,D.简码 As 团体简码 From 体检人员档案 A,体检登记记录 B,病人信息 C,合约单位 D " & _
                                    "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=1 AND a.体检时间 BETWEEN [5] AND [6] And b.合约单位id=D.ID(+) " & _
                                    ")"
            Case Else
                strSQL = "Select * From (Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号 From 体检人员档案 A,体检登记记录 B,病人信息 C " & _
                                    "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=1 AND a.体检时间 BETWEEN [5] AND [6] " & _
                                    "Union All " & _
                                    "Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号 From 体检人员档案 A,体检登记记录 B,病人信息 C " & _
                                    "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=0 AND a.体检时间 BETWEEN [1] AND [2])"
            End Select
        Else
            Select Case strCol
            Case "团体简码"
                strSQL = "Select * From (Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号,D.简码 As 团体简码 From 体检人员档案 A,体检登记记录 B,病人信息 C,合约单位 D " & _
                                        "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=1 And b.体检时间 BETWEEN [5] AND [6] And b.合约单位id=D.ID(+) " & _
                                        ") "
            Case Else
                strSQL = "Select * From (Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号 From 体检人员档案 A,体检登记记录 B,病人信息 C " & _
                                        "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=1 And b.体检时间 BETWEEN [5] AND [6]  " & _
                                        "Union All " & _
                                        "Select a.体检编号,A.ID,A.组别名称,A.登记id,A.病人id,B.是否团体,b.体检号,c.姓名,c.门诊号,c.健康号,c.身份证号,c.就诊卡号 From 体检人员档案 A,体检登记记录 B,病人信息 C " & _
                                        "Where B.体检状态=Decode([4],0,2,1,4,2,5) AND C.病人id=A.病人id AND A.登记id=B.ID AND Nvl(b.是否团体,0)=0 And b.体检时间 BETWEEN [1] AND [2]) "
            End Select
            
        End If
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            Call txt_LostFocus(Index)
            
            If mintIndex = 1 Then
                strStart = Split(mstr正体检团体时间范围, "|")(0)
                strEnd = Split(mstr正体检团体时间范围, "|")(1)
            Else
                strStart = Split(vsf(mintIndex).Tag, "|")(0)
                strEnd = Split(vsf(mintIndex).Tag, "|")(1)
            End If
            
            
            If mstrSvrFind <> txt(Index).Text Then
                
                mstrSvrFind = txt(Index).Text

                Select Case strCol
                    Case "体检单号"
                    
                        strSQL = strSQL & " Where 体检号 Like [3] Order By 是否团体,登记id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & txt(Index).Text & "%", mintIndex, CDate(strStart), CDate(strEnd))
                                                
                    Case "门诊号", "健康号", "就诊卡号"
                        strSQL = strSQL & " Where " & strCol & " = [3] Order By 是否团体,登记id "
                        
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), Val(txt(Index).Text), mintIndex, CDate(strStart), CDate(strEnd))
            
                    Case "身份证号"
                        
                        strSQL = strSQL & " Where 身份证号=[3] Order By 是否团体,登记id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), txt(Index).Text, mintIndex, CDate(strStart), CDate(strEnd))
                        
                    Case "姓名拼音"
                        
                        strSQL = strSQL & " Where zlSpellCode(姓名) Like [3] Order By 是否团体,登记id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                    Case "姓名五笔"
                        strSQL = strSQL & " Where zlWBCode(姓名) Like [3] Order By 是否团体,登记id "
                
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                    
                    Case "团体简码"
                        
                        strSQL = strSQL & " Where 团体简码 Like [3] Order By 是否团体,登记id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                                                           
                                                           
                    Case Else
                    
                        strSQL = strSQL & " Where " & strCol & " Like [3] Order By 是否团体,登记id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & txt(Index).Text & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                End Select

                If mrsFind.BOF Then
                    ShowSimpleMsg "没有找到符合要求的信息！"
                    txt(Index).Text = ""
                    Exit Sub
                End If
            End If
            
            If mrsFind.EOF And mrsFind.RecordCount > 0 Then mrsFind.MoveFirst
            If Not mrsFind.EOF Then
                
                usrSave.lng登记id = mrsFind("登记id").Value
                usrSave.lng病人id = mrsFind("病人id").Value
                usrSave.str组别 = mrsFind("组别名称").Value
                
                Call SelectPerson(IIf(mrsFind("是否团体") = 1, False, True))
                
            End If
            
            On Error Resume Next
            Err = 0
            mrsFind.MoveNext
            If Err <> 0 Then ShowSimpleMsg "已经查找完，如再查找将重新搜索一次！"
            
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    Else
        If Index = 1 Then
            Select Case lbl(1).Tag
            Case "体检单号", "就诊卡号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
            
        End If
        
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    
    If lbl(1).Tag = "体检单号" Then
        Dim intYear As Integer
        Dim strYear As String
        '自动补齐单据号
        If (UCase(Left(txt(Index).Text, 1)) < "A" Or UCase(Left(txt(Index).Text, 1)) > "Z") And Trim(txt(Index).Text) <> "" Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txt(Index).Text = strYear & Right("0000000" & txt(Index).Text, 7)
        End If
    End If
End Sub

Private Sub vsf_AfterCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoAllowChange Then Exit Sub
    
    If OldRow = NewRow Then Exit Sub
    
    Call ClearData("结果")
        
    mlngSvrKey(Index) = Val(vsf(Index).RowData(NewRow))
    
    '读取细节
    Call RefreshData("基本")
    
    If mintIndex = 1 Or mintIndex = 2 Then
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "病人id"))) > 0 Then
            If tbs.Tabs(1).Key <> "报告" Then
                tbs.Tabs.Clear
                tbs.Tabs.Add , "报告", "&1.报告"
                tbs.Tabs.Add , "总检", "&2.总检"
                tbs.Tabs.Add , "费用", "&3.费用"
                tbs.Tabs.Add , "历次", "&4.历次"
            End If
        Else
            If tbs.Tabs(1).Key <> "概况" Then
                tbs.Tabs.Clear
                tbs.Tabs.Add , "概况", "&1.概况"
            End If
        End If
    Else
        If tbs.Tabs(1).Key <> "报告" Then
            tbs.Tabs.Clear
            tbs.Tabs.Add , "报告", "&1.报告"
            tbs.Tabs.Add , "总检", "&2.总检"
            tbs.Tabs.Add , "费用", "&3.费用"
            tbs.Tabs.Add , "历次", "&4.历次"
        End If
    End If
    
    Call tbs_Click
    
    Call AdjustEnableState
    
    On Error Resume Next
    vsf(Index).SetFocus
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_BeforeCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    Dim lng登记id As Long
    Dim str组别 As String
    Dim int标志 As Integer

    If Index > 2 Then Exit Sub
    If mblnStartUp Then Exit Sub

    On Error GoTo errHand

    If State = 0 Then
        '展开,如果没有装载,则装载人员数据
        int标志 = Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "标志")))
        Select Case int标志
            Case 0, 2              '0-个人分组项;1-团体名称项;2-团体组别项;99-受检人员项
                '展开的是组别
                If Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "是否装载"))) = 0 Then
                    '没有装载过
                    vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "是否装载")) = 1

                    '1.删除空行
                    vsf(Index).RemoveItem Row + 1

                    '2.装载此组的人员清单
                    lng登记id = Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "登记id")))

                    If int标志 = 0 Then
                        str组别 = "缺省"
                    Else
                        str组别 = vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "姓名"))
                    End If

                    Select Case Index
                    Case 0
                        Call RefreshData("组别人员", Row & "'" & lng登记id & "'" & str组别 & "'2'0'" & vsf(Index).Tag)
                    Case 1
                        Call RefreshData("组别人员", Row & "'" & lng登记id & "'" & str组别 & "'4'1'" & vsf(Index).Tag)
                    Case 2
                        Call RefreshData("组别人员", Row & "'" & lng登记id & "'" & str组别 & "'5'1'" & vsf(Index).Tag)
                    End Select

                End If
        End Select

    End If

errHand:

End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 3)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    
    Dim r As Long
    
    On Error GoTo errHand
    
    With vsf(Index)

        r = .Row

        If .IsCollapsed(r) = flexOutlineCollapsed Then

            .IsCollapsed(r) = flexOutlineExpanded

        Else

            .IsCollapsed(r) = flexOutlineCollapsed

        End If

    End With
    Call InheritAppendSpaceRows(Index)
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vsf_DblClick (Index)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 3 Then Exit Sub
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        If mnuMedical.Visible Then Me.PopupMenu mnuMedical
    End If
    
    If Button = 1 Then
        If vsf(Index).MouseRow = 0 And vsf(Index).MouseCol > 2 Then
        
            mintSort = IIf(mintSort = flexSortGenericAscending, flexSortGenericDescending, flexSortGenericAscending)
            vsf(Index).Sort = mintSort
            
            Set vsf(Index).Cell(flexcpPicture, 0, 3, 0, vsf(Index).Cols - 1) = Nothing
            
            If mintSort = flexSortGenericAscending Then
                vsf(Index).Cell(flexcpPicture, 0, vsf(Index).Col) = ils13.ListImages("up").Picture
            Else
                vsf(Index).Cell(flexcpPicture, 0, vsf(Index).Col) = ils13.ListImages("down").Picture
            End If
        End If
    End If
    
End Sub
    
Public Function LoadOutLineGrid(ByVal intIndex As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
        '------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    Dim objMsf As Object

    Dim lngCol病人id As Long
    Dim lngCol上级id As Long
    
    vsf(intIndex).Redraw = False
    
    On Error Resume Next
    
    blnForeColor = (rsData("前景色").Name = "前景色")
    
    On Error GoTo 0
    
    Set objMsf = vsf(intIndex)
    
    lngCol病人id = GetCol(objMsf, "病人id")
    lngCol上级id = GetCol(objMsf, "上级id")
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = zlCommFun.NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '下一列
        Next

        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("前景色").Value)
        
        If Val(vsf(intIndex).TextMatrix(lngRow, lngCol病人id)) = 0 Then

            vsf(intIndex).MergeRow(lngRow) = True
            vsf(intIndex).IsSubtotal(lngRow) = True
            
            Select Case Val(vsf(intIndex).TextMatrix(lngRow, GetCol(vsf(intIndex), "标志")))
                Case 0               '0-个人分组项;1-团体名称项;2-团体组别项;99-受检人员项
                    vsf(intIndex).Cell(flexcpFontBold, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = True
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.灰色
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
                    
                    vsf(intIndex).AddItem ""
                    lngRow = lngRow + 1
                Case 2
                    vsf(intIndex).RowOutlineLevel(lngRow) = 1
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.浅黄色
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
                    vsf(intIndex).AddItem ""
                    lngRow = lngRow + 1
                Case Else
                    vsf(intIndex).Cell(flexcpFontBold, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = True
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.灰色
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
            End Select
            

        End If
        
        rsData.MoveNext
    Loop

    vsf(intIndex).Redraw = True

    Call InheritAppendSpaceRows(intIndex)
    
    vsf(intIndex).Outline 1
    vsf(intIndex).Outline 0
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

