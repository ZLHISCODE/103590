VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSystemPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "基础参数设置"
   ClientHeight    =   9465
   ClientLeft      =   2565
   ClientTop       =   1485
   ClientWidth     =   10230
   Icon            =   "frmSystemPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   2
      Left            =   240
      TabIndex        =   283
      Top             =   480
      Width           =   9690
      Begin VB.CheckBox chk 
         Caption         =   "新开医嘱签名时一组医嘱签名一次"
         Height          =   195
         Index           =   91
         Left            =   5760
         TabIndex        =   345
         Top             =   120
         Width           =   3540
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   8760
         TabIndex        =   296
         Top             =   935
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   7180
         TabIndex        =   295
         Top             =   985
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "PACS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   29
         Left            =   5280
         TabIndex        =   293
         Top             =   720
         Width           =   660
      End
      Begin VB.CheckBox chk 
         Caption         =   "LIS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   292
         Top             =   720
         Width           =   660
      End
      Begin VB.CheckBox chk 
         Caption         =   "药品发药"
         Enabled         =   0   'False
         Height          =   195
         Index           =   60
         Left            =   2880
         TabIndex        =   291
         Top             =   720
         Width           =   1020
      End
      Begin VB.CheckBox chk 
         Caption         =   "门诊医嘱,病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   44
         Left            =   960
         TabIndex        =   290
         Top             =   480
         Width           =   1620
      End
      Begin VB.CheckBox chk 
         Caption         =   "住院医嘱,病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   45
         Left            =   2880
         TabIndex        =   289
         Top             =   480
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "医技医嘱,报告"
         Enabled         =   0   'False
         Height          =   195
         Index           =   46
         Left            =   4440
         TabIndex        =   288
         Top             =   480
         Width           =   1860
      End
      Begin VB.CheckBox chk 
         Caption         =   "护理记录,护理病历"
         Enabled         =   0   'False
         Height          =   195
         Index           =   47
         Left            =   960
         TabIndex        =   287
         Top             =   720
         Width           =   1860
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         Index           =   11
         ItemData        =   "frmSystemPara.frx":000C
         Left            =   960
         List            =   "frmSystemPara.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   284
         Top             =   97
         Width           =   3540
      End
      Begin TabDlg.SSTab sstSign 
         Height          =   6690
         Left            =   120
         TabIndex        =   286
         Top             =   1320
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   11800
         _Version        =   393216
         Style           =   1
         Tabs            =   8
         Tab             =   7
         TabsPerRow      =   8
         TabHeight       =   520
         TabCaption(0)   =   "门诊医嘱,病历"
         TabPicture(0)   =   "frmSystemPara.frx":0010
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "vsDept(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "住院医生医嘱,病历"
         TabPicture(1)   =   "frmSystemPara.frx":002C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vsDept(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "住院护士医嘱"
         TabPicture(2)   =   "frmSystemPara.frx":0048
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vsDept(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "医技医嘱,报告"
         TabPicture(3)   =   "frmSystemPara.frx":0064
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "vsDept(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "护理记录,护理病历"
         TabPicture(4)   =   "frmSystemPara.frx":0080
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "vsDept(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "药品发药"
         TabPicture(5)   =   "frmSystemPara.frx":009C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "vsDept(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "LIS"
         TabPicture(6)   =   "frmSystemPara.frx":00B8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "vsDept(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "PACS"
         TabPicture(7)   =   "frmSystemPara.frx":00D4
         Tab(7).ControlEnabled=   -1  'True
         Tab(7).Control(0)=   "vsDept(7)"
         Tab(7).Control(0).Enabled=   0   'False
         Tab(7).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   1
            Left            =   -74880
            TabIndex        =   297
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":00F0
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5865
            Index           =   0
            Left            =   -74880
            TabIndex        =   298
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10345
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":0183
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   2
            Left            =   -74880
            TabIndex        =   299
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":0216
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   3
            Left            =   -74880
            TabIndex        =   300
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":02A9
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   4
            Left            =   -74880
            TabIndex        =   301
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":033C
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   5
            Left            =   -74880
            TabIndex        =   302
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":03CF
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   6
            Left            =   -74880
            TabIndex        =   303
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":0462
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   6015
            Index           =   7
            Left            =   120
            TabIndex        =   304
            Top             =   360
            Width           =   9255
            _cx             =   16325
            _cy             =   10610
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":04F5
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
      Begin VB.Label Label15 
         Caption         =   "说明：启用场合后，未勾选任何部门，表示不按科室控制。"
         Height          =   255
         Left            =   240
         TabIndex        =   305
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "场合"
         Height          =   180
         Left            =   480
         TabIndex        =   294
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "认证中心"
         Height          =   255
         Left            =   120
         TabIndex        =   285
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   3
      Left            =   270
      TabIndex        =   151
      Top             =   540
      Width           =   9690
      Begin VB.Frame Fra 
         Height          =   75
         Index           =   11
         Left            =   7800
         TabIndex        =   225
         Top             =   4815
         Width           =   510
      End
      Begin VB.CommandButton cmdOneCard 
         Height          =   345
         Index           =   0
         Left            =   5040
         Picture         =   "frmSystemPara.frx":0588
         Style           =   1  'Graphical
         TabIndex        =   224
         ToolTipText     =   "增加一种前缀"
         Top             =   7605
         Width           =   345
      End
      Begin VB.CommandButton cmdOneCard 
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   5400
         Picture         =   "frmSystemPara.frx":0B12
         Style           =   1  'Graphical
         TabIndex        =   223
         ToolTipText     =   "修改当前前缀"
         Top             =   7605
         Width           =   345
      End
      Begin VB.CommandButton cmdOneCard 
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   5760
         Picture         =   "frmSystemPara.frx":109C
         Style           =   1  'Graphical
         TabIndex        =   222
         ToolTipText     =   "删除当前前缀"
         Top             =   7605
         Width           =   345
      End
      Begin VB.Frame Fra 
         Height          =   75
         Index           =   15
         Left            =   1065
         TabIndex        =   220
         Top             =   135
         Width           =   5070
      End
      Begin VB.Frame Fra 
         Height          =   75
         Index           =   9
         Left            =   7530
         TabIndex        =   153
         Top             =   165
         Width           =   2070
      End
      Begin VB.ListBox lst 
         Height          =   2370
         Index           =   3
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   78
         Top             =   5160
         Width           =   1785
      End
      Begin VB.CheckBox chk 
         Caption         =   "严格控制"
         Height          =   285
         Index           =   13
         Left            =   8040
         TabIndex        =   77
         ToolTipText     =   "表示各个输入就诊卡号码处是否为密文显示"
         Top             =   435
         Width           =   1020
      End
      Begin VB.TextBox txtUD 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   7125
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "7"
         Top             =   420
         Width           =   390
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   300
         Index           =   4
         Left            =   7515
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   7
         BuddyControl    =   "txtUD(4)"
         BuddyDispid     =   196631
         BuddyIndex      =   4
         OrigLeft        =   3795
         OrigTop         =   3630
         OrigRight       =   4035
         OrigBottom      =   3915
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   2790
         Index           =   0
         Left            =   6360
         TabIndex        =   74
         Top             =   750
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "票据类型"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "号码长度"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "严格控制"
            Object.Width           =   1588
         EndProperty
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   7125
         Index           =   3
         Left            =   165
         TabIndex        =   219
         Top             =   405
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   12568
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "NO"
            Text            =   "编号"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Name"
            Text            =   "名称"
            Object.Width           =   3882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "PayType"
            Text            =   "结算方式"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "OrgCode"
            Text            =   "医院编码"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Enable"
            Text            =   "启用"
            Object.Width           =   970
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一卡通接口"
         Height          =   180
         Index           =   45
         Left            =   165
         TabIndex        =   221
         Top             =   75
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "刷卡要求输密码"
         Height          =   180
         Index           =   41
         Left            =   6360
         TabIndex        =   201
         Top             =   4755
         Width           =   1260
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "号码长度"
         Height          =   180
         Index           =   19
         Left            =   6360
         TabIndex        =   154
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号码控制"
         Height          =   180
         Index           =   9
         Left            =   6360
         TabIndex        =   152
         Top             =   75
         Width           =   1080
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8205
      Index           =   1
      Left            =   270
      TabIndex        =   193
      Top             =   465
      Width           =   9660
      Begin VB.Frame fraCLKS 
         Height          =   1650
         Left            =   3690
         TabIndex        =   338
         Top             =   5025
         Width           =   5895
         Begin VSFlex8Ctl.VSFlexGrid vsUnWriteDept 
            Height          =   1005
            Left            =   90
            TabIndex        =   340
            Top             =   525
            Width           =   5700
            _cx             =   10054
            _cy             =   1773
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
            Rows            =   4
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":1626
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
         Begin VB.CheckBox chk 
            Caption         =   "医嘱超量时必须输入原因"
            Height          =   240
            Index           =   86
            Left            =   105
            TabIndex        =   339
            ToolTipText     =   "勾选时属于表格中科室的病人下达医嘱可不写超量说明"
            Top             =   0
            Width           =   2280
         End
         Begin VB.Label Label16 
            Caption         =   "请设置可不录入超量原因的科室，例如：精神科。"
            Height          =   255
            Left            =   360
            TabIndex        =   341
            Top             =   300
            Width           =   4815
         End
      End
      Begin VB.Frame fraBlood 
         Height          =   555
         Left            =   3675
         TabIndex        =   273
         Top             =   4425
         Width           =   5895
         Begin VB.CheckBox chk 
            Caption         =   "启用输血申请三级审核"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   276
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血申请只能由中级及以上医师提出"
            Enabled         =   0   'False
            Height          =   200
            Index           =   85
            Left            =   2400
            TabIndex        =   275
            Top             =   285
            Width           =   3375
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用输血分级管理"
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   274
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 医嘱发送执行 "
         Height          =   2805
         Index           =   13
         Left            =   3675
         TabIndex        =   263
         Top             =   1005
         Width           =   5895
         Begin VB.CheckBox chk 
            Caption         =   "允许取消"
            Height          =   200
            Index           =   87
            Left            =   120
            TabIndex        =   311
            Top             =   2505
            Width           =   1035
         End
         Begin VB.TextBox txtUNExecLimit 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   310
            Text            =   "999"
            Top             =   2475
            Width           =   525
         End
         Begin VB.CheckBox chk 
            Caption         =   "检验医嘱发送时生成样本条码"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   271
            ToolTipText     =   "是否在执行操作后将划价单审核为记帐单"
            Top             =   2040
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "执行之后自动审核记帐划价单"
            Height          =   210
            Index           =   32
            Left            =   120
            TabIndex        =   270
            ToolTipText     =   "是否在执行操作后将划价单审核为记帐单"
            Top             =   2280
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "执行之后对卫生材料自动发料"
            Height          =   210
            Index           =   61
            Left            =   2880
            TabIndex        =   269
            ToolTipText     =   "是否在执行操作后将划价单审核为记帐单"
            Top             =   2280
            Width           =   2640
         End
         Begin VB.CommandButton cmdSendPriceType 
            Caption         =   "全选(&A)"
            Height          =   350
            Index           =   0
            Left            =   3480
            TabIndex        =   265
            Top             =   420
            Width           =   1100
         End
         Begin VB.CommandButton cmdSendPriceType 
            Caption         =   "全清(&U)"
            Height          =   350
            Index           =   1
            Left            =   4680
            TabIndex        =   264
            Top             =   420
            Width           =   1100
         End
         Begin TabDlg.SSTab SendPriceType 
            Height          =   1545
            Left            =   120
            TabIndex        =   266
            Top             =   435
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   2725
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabMaxWidth     =   882
            TabCaption(0)   =   "门诊"
            TabPicture(0)   =   "frmSystemPara.frx":16DC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lst(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "住院"
            TabPicture(1)   =   "frmSystemPara.frx":16F8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lst(2)"
            Tab(1).ControlCount=   1
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   4
               ItemData        =   "frmSystemPara.frx":1714
               Left            =   75
               List            =   "frmSystemPara.frx":1716
               Style           =   1  'Checkbox
               TabIndex        =   268
               Top             =   360
               Width           =   5475
            End
            Begin VB.ListBox lst 
               Columns         =   3
               Height          =   1110
               IMEMode         =   3  'DISABLE
               Index           =   2
               ItemData        =   "frmSystemPara.frx":1718
               Left            =   -74925
               List            =   "frmSystemPara.frx":171A
               Style           =   1  'Checkbox
               TabIndex        =   267
               Top             =   360
               Width           =   5475
            End
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "天内的医嘱执行操作"
            Height          =   180
            Index           =   1
            Left            =   1800
            TabIndex        =   312
            Top             =   2520
            Width           =   1620
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发送为记帐划价单的诊疗类别："
            Height          =   180
            Left            =   240
            TabIndex        =   272
            Top             =   225
            Width           =   2520
         End
      End
      Begin VB.Frame fraKSSStrict 
         Height          =   525
         Index           =   14
         Left            =   3675
         TabIndex        =   259
         Top             =   3855
         Width           =   2895
         Begin VB.CheckBox chk 
            Caption         =   "启用抗菌药物分级管理"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   261
            Top             =   0
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "抗菌药物允许使用自备药"
            Enabled         =   0   'False
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   260
            Top             =   225
            Width           =   2295
         End
      End
      Begin VB.Frame frmOPS 
         Height          =   495
         Left            =   6675
         TabIndex        =   256
         Top             =   3855
         Width           =   2895
         Begin VB.CheckBox chk 
            Caption         =   "启用手术医师授权管理"
            Enabled         =   0   'False
            Height          =   240
            Index           =   49
            Left            =   120
            TabIndex        =   258
            Top             =   225
            Width           =   2220
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用手术分级管理"
            Height          =   240
            Index           =   80
            Left            =   120
            TabIndex        =   257
            Top             =   0
            Width           =   1740
         End
      End
      Begin VB.Frame fraCheckDrug 
         Height          =   1740
         Left            =   150
         TabIndex        =   247
         Top             =   6405
         Width           =   3345
         Begin VB.OptionButton optPASSVer 
            Caption         =   "美康4.0"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   330
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "美康3.0"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   329
            Top             =   1320
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许使用系统设置"
            Height          =   240
            Index           =   89
            Left            =   120
            TabIndex        =   325
            Top             =   1080
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许下达院外执行的禁忌药品医嘱"
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   324
            Top             =   600
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "禁止下达超极量药品医嘱"
            Height          =   240
            Index           =   63
            Left            =   120
            TabIndex        =   323
            Top             =   840
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许下达禁忌药品医嘱"
            Height          =   240
            Index           =   65
            Left            =   120
            TabIndex        =   322
            Top             =   350
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "启用接口调用日志"
            Height          =   240
            Index           =   88
            Left            =   120
            TabIndex        =   321
            Top             =   1080
            Width           =   2940
         End
         Begin VB.ComboBox cmb 
            Enabled         =   0   'False
            Height          =   300
            Index           =   27
            ItemData        =   "frmSystemPara.frx":171C
            Left            =   1260
            List            =   "frmSystemPara.frx":171E
            Style           =   2  'Dropdown List
            TabIndex        =   313
            Top             =   1335
            Width           =   1770
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   20
            ItemData        =   "frmSystemPara.frx":1720
            Left            =   1380
            List            =   "frmSystemPara.frx":1722
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   37
            Width           =   1410
         End
         Begin VB.Label lblPassVer 
            Caption         =   "当前版本："
            Height          =   255
            Left            =   120
            TabIndex        =   328
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "过敏输入来源"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   314
            Top             =   1395
            Width           =   1080
         End
         Begin VB.Label lbl合理用药接口 
            Caption         =   "合理用药接口"
            Height          =   255
            Left            =   270
            TabIndex        =   252
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmd社区参数 
         Caption         =   "设置(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8470
         TabIndex        =   73
         ToolTipText     =   "对当前选择的社区接口的参数进行设置"
         Top             =   6690
         Width           =   1100
      End
      Begin VB.Frame Fra 
         Caption         =   " 诊断输入 "
         Height          =   900
         Index           =   12
         Left            =   3675
         TabIndex        =   196
         Top             =   75
         Width           =   5895
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   16
            ItemData        =   "frmSystemPara.frx":1724
            Left            =   3480
            List            =   "frmSystemPara.frx":1726
            Style           =   2  'Dropdown List
            TabIndex        =   72
            ToolTipText     =   "影响范围：入出院，医生工作站"
            Top             =   540
            Width           =   2310
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   8
            ItemData        =   "frmSystemPara.frx":1728
            Left            =   720
            List            =   "frmSystemPara.frx":172A
            Style           =   2  'Dropdown List
            TabIndex        =   71
            ToolTipText     =   "影响范围：入出院，医生工作站"
            Top             =   540
            Width           =   2310
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   1
            ItemData        =   "frmSystemPara.frx":172C
            Left            =   720
            List            =   "frmSystemPara.frx":172E
            Style           =   2  'Dropdown List
            TabIndex        =   70
            ToolTipText     =   "影响范围：医生工作站"
            Top             =   210
            Width           =   2310
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院"
            Height          =   180
            Index           =   51
            Left            =   3045
            TabIndex        =   210
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊"
            Height          =   180
            Index           =   27
            Left            =   285
            TabIndex        =   198
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "来源"
            Height          =   180
            Index           =   39
            Left            =   285
            TabIndex        =   197
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.Frame Fra 
         Caption         =   "医嘱相关"
         Height          =   6255
         Index           =   0
         Left            =   165
         TabIndex        =   194
         Top             =   75
         Width           =   3345
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   2010
            MaxLength       =   4
            TabIndex        =   306
            Text            =   "10"
            Top             =   895
            Width           =   495
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   26
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   254
            Top             =   5130
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   $"frmSystemPara.frx":1730
            Height          =   255
            Index           =   79
            Left            =   240
            TabIndex        =   250
            Top             =   5460
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "回退出院医嘱才能撤销预出院"
            Height          =   240
            Index           =   81
            Left            =   240
            TabIndex        =   67
            Top             =   4020
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "输血和皮试医嘱执行后需要核对"
            Height          =   240
            Index           =   74
            Left            =   240
            TabIndex        =   246
            Top             =   4860
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达医嘱时显示产地"
            Height          =   240
            Index           =   66
            Left            =   240
            TabIndex        =   240
            Top             =   4575
            Value           =   1  'Checked
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "指定医嘱在其他科室执行"
            Height          =   240
            Index           =   62
            Left            =   240
            TabIndex        =   63
            Top             =   2940
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   2010
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   211
            Text            =   "12"
            Top             =   1455
            Width           =   495
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   2010
            MaxLength       =   4
            TabIndex        =   53
            Text            =   "30"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2010
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   56
            Text            =   "5"
            Top             =   1775
            Width           =   495
         End
         Begin VB.CheckBox chk 
            Caption         =   "下达出院医嘱才能出院"
            Height          =   240
            Index           =   50
            Left            =   240
            TabIndex        =   66
            Top             =   3750
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "长期医嘱缺省为次日生效"
            Height          =   240
            Index           =   24
            Left            =   240
            TabIndex        =   62
            Top             =   2655
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品长期医嘱按规格下达"
            Height          =   240
            Index           =   3
            Left            =   240
            TabIndex        =   61
            ToolTipText     =   " "
            Top             =   2370
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊药嘱先作废后退药"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   68
            Top             =   4305
            Width           =   2820
         End
         Begin VB.TextBox txtUD 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   270
            Index           =   7
            Left            =   2010
            MaxLength       =   3
            TabIndex        =   59
            Text            =   "1"
            Top             =   2085
            Width           =   495
         End
         Begin VB.CheckBox chk 
            Caption         =   "一次申请多个检验项目"
            Height          =   240
            Index           =   34
            Left            =   240
            TabIndex        =   65
            Top             =   3480
            Width           =   2820
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   10
            ItemData        =   "frmSystemPara.frx":1754
            Left            =   1020
            List            =   "frmSystemPara.frx":1756
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   255
            Width           =   1740
         End
         Begin VB.CommandButton cmdAdvice 
            Caption         =   "医嘱内容定义(&F)"
            Height          =   405
            Left            =   240
            TabIndex        =   69
            Top             =   5730
            Width           =   1680
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   7
            Left            =   2520
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   2070
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(7)"
            BuddyDispid     =   196631
            BuddyIndex      =   7
            OrigLeft        =   2250
            OrigTop         =   1665
            OrigRight       =   2490
            OrigBottom      =   1965
            Max             =   365
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   0   'False
         End
         Begin VB.CheckBox chk 
            Caption         =   "过敏登记有效天数"
            Height          =   240
            Index           =   11
            Left            =   240
            TabIndex        =   58
            Top             =   2100
            Width           =   1740
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   3
            Left            =   2520
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1770
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txtUD(3)"
            BuddyDispid     =   196631
            BuddyIndex      =   3
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   999
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   0   'False
         End
         Begin VB.CheckBox chk 
            Caption         =   "处方药品条数限制"
            Height          =   240
            Index           =   52
            Left            =   240
            TabIndex        =   55
            Top             =   1790
            Width           =   1740
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   8
            Left            =   2520
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txtUD(8)"
            BuddyDispid     =   196631
            BuddyIndex      =   8
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   9999
            Min             =   10
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   9
            Left            =   2520
            TabIndex        =   212
            TabStop         =   0   'False
            Top             =   1455
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(9)"
            BuddyDispid     =   196631
            BuddyIndex      =   9
            OrigLeft        =   2520
            OrigTop         =   1365
            OrigRight       =   2760
            OrigBottom      =   1665
            Max             =   20
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "住院药嘱发送产生领药号"
            Height          =   240
            Index           =   64
            Left            =   240
            TabIndex        =   64
            Top             =   3210
            Width           =   2280
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   270
            Index           =   11
            Left            =   2520
            TabIndex        =   307
            TabStop         =   0   'False
            Top             =   895
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Value           =   10
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   9999
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "则以当前时间作为开始时间"
            Height          =   180
            Index           =   55
            Left            =   990
            TabIndex        =   309
            Top             =   1200
            Width           =   2160
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊新开医嘱间隔         分钟"
            Height          =   180
            Index           =   25
            Left            =   540
            TabIndex        =   308
            Top             =   940
            Width           =   2610
         End
         Begin VB.Label lbl 
            Caption         =   "中药配方每行"
            Height          =   255
            Index           =   54
            Left            =   240
            TabIndex        =   253
            Top             =   5160
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "儿童年龄界定上限         岁"
            Height          =   180
            Index           =   47
            Left            =   540
            TabIndex        =   213
            Top             =   1500
            Width           =   2430
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "补录医嘱识别间隔         分钟"
            Height          =   180
            Index           =   43
            Left            =   540
            TabIndex        =   202
            Top             =   660
            Width           =   2610
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "诊疗编码"
            Height          =   180
            Index           =   36
            Left            =   240
            TabIndex        =   195
            Top             =   315
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvw社区 
         Height          =   1065
         Left            =   3675
         TabIndex        =   262
         Top             =   7065
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1879
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "说明"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "部件"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "启用"
            Object.Width           =   952
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "社区档案接口："
         Height          =   180
         Left            =   3675
         TabIndex        =   277
         Top             =   6780
         Width           =   1260
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   0
      Left            =   225
      TabIndex        =   132
      Top             =   525
      Width           =   9720
      Begin VB.Frame Fra 
         Caption         =   " 药品结存时点 "
         Height          =   1080
         Index           =   10
         Left            =   6840
         TabIndex        =   278
         Top             =   6840
         Width           =   2775
         Begin VB.OptionButton optAccountTime 
            Caption         =   "每月"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   281
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optAccountTime 
            Caption         =   "每月最后一天"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   280
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.TextBox txtAccountTime 
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
            Left            =   1100
            TabIndex        =   279
            Text            =   "25"
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label14 
            Caption         =   "日"
            Height          =   255
            Left            =   1560
            TabIndex        =   282
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 入院时允许 "
         Height          =   1005
         Index           =   5
         Left            =   6840
         TabIndex        =   140
         Top             =   1275
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "办理就诊卡"
            Height          =   195
            Index           =   5
            Left            =   1500
            TabIndex        =   40
            ToolTipText     =   "表示在办理入院时是否允许同时办理就诊卡"
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "收取预交款"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   39
            ToolTipText     =   "表示在办理入院时是否同时收取预交款"
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "分配床位号"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   41
            Top             =   600
            Width           =   1200
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 药品与卫材 "
         Height          =   7845
         Index           =   2
         Left            =   3310
         TabIndex        =   186
         Top             =   75
         Width           =   3480
         Begin VB.CheckBox chk 
            Caption         =   "输液配置中心首次执行的医嘱需要进行审核"
            Height          =   375
            Index           =   83
            Left            =   225
            TabIndex        =   255
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价药品入库时取上次售价"
            Height          =   195
            Index           =   73
            Left            =   225
            TabIndex        =   245
            Top             =   5880
            Width           =   2760
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价药品通过分段加成入库"
            Height          =   180
            Index           =   14
            Left            =   225
            TabIndex        =   244
            Top             =   6600
            Width           =   2775
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品领用时明确药品批次"
            Height          =   195
            Index           =   72
            Left            =   240
            TabIndex        =   243
            Top             =   4440
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品移库时明确药品批次"
            Height          =   195
            Index           =   71
            Left            =   225
            TabIndex        =   242
            Top             =   4200
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品外购入库需要经过标记付款后才能进行付款管理"
            Height          =   360
            Index           =   70
            Left            =   225
            TabIndex        =   241
            Top             =   5040
            Width           =   2520
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   18
            ItemData        =   "frmSystemPara.frx":1758
            Left            =   1605
            List            =   "frmSystemPara.frx":175A
            Style           =   2  'Dropdown List
            TabIndex        =   217
            Top             =   1200
            Width           =   1780
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   17
            ItemData        =   "frmSystemPara.frx":175C
            Left            =   1605
            List            =   "frmSystemPara.frx":175E
            Style           =   2  'Dropdown List
            TabIndex        =   215
            Top             =   872
            Width           =   1780
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价药品入库按扣前加成销售"
            Height          =   195
            Index           =   48
            Left            =   225
            TabIndex        =   32
            ToolTipText     =   "时价药品外购入库时售价计算方式：不选择-按折扣后的采购价计算售价;选择－按折扣前的采购价计算售价。"
            Top             =   6840
            Width           =   3090
         End
         Begin VB.CheckBox chk 
            Caption         =   "填写药品出库类单据减少药品出库库房可用库存"
            Height          =   375
            Index           =   40
            Left            =   225
            TabIndex        =   27
            Top             =   3360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "卫材在门诊收费或记帐后自动发料"
            Height          =   195
            Index           =   38
            Left            =   225
            TabIndex        =   33
            ToolTipText     =   "影响的范围:门诊收费,记帐,医技工作站补费(不含[收费单据])"
            Top             =   7200
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊收费与药房发药分离"
            Height          =   240
            Index           =   22
            Left            =   225
            TabIndex        =   23
            Top             =   2160
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "住院记帐与药房发药分离"
            Height          =   225
            Index           =   23
            Left            =   225
            TabIndex        =   24
            Top             =   2400
            Width           =   2280
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   3
            Left            =   1605
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   220
            Width           =   1780
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   9
            ItemData        =   "frmSystemPara.frx":1760
            Left            =   1605
            List            =   "frmSystemPara.frx":1762
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   546
            Width           =   1780
         End
         Begin VB.CheckBox chk 
            Caption         =   "卫材在住院记帐后自动发料"
            Height          =   195
            Index           =   37
            Left            =   225
            TabIndex        =   34
            ToolTipText     =   "影响范围:住院记帐,医技工作站(补费:记帐单据)"
            Top             =   7440
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "指定药房时限定药品库存"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   25
            Top             =   2745
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品收费完成后自动发药"
            Height          =   195
            Index           =   17
            Left            =   225
            TabIndex        =   26
            Top             =   3000
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价药品通过加成率入库"
            Height          =   195
            Index           =   21
            Left            =   225
            TabIndex        =   31
            Top             =   6360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "药品申领时明确药品批次"
            Height          =   195
            Index           =   26
            Left            =   225
            TabIndex        =   28
            Top             =   3960
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "外购入库单需要经过核查"
            Height          =   195
            Index           =   28
            Left            =   225
            TabIndex        =   29
            Top             =   4800
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "时价药品直接确定售价"
            Height          =   195
            Index           =   36
            Left            =   225
            TabIndex        =   30
            Top             =   5640
            Width           =   2280
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药品出库优先算法"
            Height          =   180
            Index           =   44
            Left            =   120
            TabIndex        =   218
            Top             =   1260
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药品效期显示方式"
            Height          =   180
            Index           =   31
            Left            =   120
            TabIndex        =   216
            Top             =   932
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "药价编辑设置单位"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   188
            Top             =   280
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药品编码递增模式"
            Height          =   180
            Index           =   32
            Left            =   120
            TabIndex        =   187
            Top             =   606
            Width           =   1440
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 发药窗口动态分配 "
         Height          =   765
         Index           =   3
         Left            =   6840
         TabIndex        =   147
         Top             =   3675
         Width           =   2775
         Begin VB.OptionButton opt 
            Caption         =   "平均方式"
            Height          =   210
            Index           =   3
            Left            =   1425
            TabIndex        =   47
            Top             =   360
            Width           =   1020
         End
         Begin VB.OptionButton opt 
            Caption         =   "闲忙方式"
            Height          =   210
            Index           =   2
            Left            =   315
            TabIndex        =   46
            Top             =   360
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 门诊收费时输入 "
         Height          =   975
         Index           =   6
         Left            =   6840
         TabIndex        =   141
         Top             =   2475
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "病人姓名"
            Height          =   210
            Index           =   7
            Left            =   330
            TabIndex        =   42
            Top             =   315
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "挂号单号"
            Height          =   210
            Index           =   10
            Left            =   1470
            TabIndex        =   45
            Top             =   585
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "病人标识"
            Height          =   225
            Index           =   8
            Left            =   330
            TabIndex        =   44
            Top             =   570
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chk 
            Caption         =   "刷就诊卡"
            Height          =   210
            Index           =   9
            Left            =   1470
            TabIndex        =   43
            Top             =   315
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 对外上下班时间 "
         Height          =   1035
         Index           =   1
         Left            =   6840
         TabIndex        =   142
         Top             =   75
         Width           =   2775
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   645
            TabIndex        =   35
            Top             =   300
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105250819
            UpDown          =   -1  'True
            CurrentDate     =   36526.3541666667
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   1635
            TabIndex        =   36
            Top             =   300
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105250819
            UpDown          =   -1  'True
            CurrentDate     =   36526.5
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   2
            Left            =   645
            TabIndex        =   37
            Top             =   675
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105250819
            UpDown          =   -1  'True
            CurrentDate     =   36526.5625
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   3
            Left            =   1635
            TabIndex        =   38
            Top             =   675
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105250819
            UpDown          =   -1  'True
            CurrentDate     =   36526.75
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "上午"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   143
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   4
            Left            =   1500
            TabIndex        =   144
            Top             =   375
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "下午"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   145
            Top             =   735
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   5
            Left            =   1500
            TabIndex        =   146
            Top             =   750
            Width           =   90
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 挂号、门诊、住院 "
         Height          =   7845
         Index           =   4
         Left            =   165
         TabIndex        =   133
         Top             =   75
         Width           =   3105
         Begin VB.CheckBox chk 
            Caption         =   "允许处理超过有效天数的病人"
            Height          =   195
            Index           =   82
            Left            =   165
            TabIndex        =   248
            Top             =   3600
            Width           =   2800
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   10
            Left            =   2150
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "1"
            Top             =   3240
            Width           =   520
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   14
            Text            =   "5"
            Top             =   4695
            Width           =   930
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊退费须先申请"
            Height          =   195
            Index           =   16
            Left            =   165
            TabIndex        =   17
            Top             =   5640
            Width           =   1920
         End
         Begin VB.CheckBox chk 
            Caption         =   "病人每次住院使用新的住院号"
            Height          =   195
            Index           =   57
            Left            =   285
            TabIndex        =   208
            Top             =   700
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "输入费用项目首位当类别简码"
            Height          =   195
            Index           =   56
            Left            =   165
            TabIndex        =   207
            Top             =   6300
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   $"frmSystemPara.frx":1764
            Height          =   195
            Index           =   31
            Left            =   165
            TabIndex        =   20
            Top             =   7410
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "输入全是数字时只查找编码"
            Height          =   195
            Index           =   30
            Left            =   165
            TabIndex        =   19
            Top             =   7125
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   14
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "四舍六入五成双,即:银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一"
            Top             =   2310
            Width           =   1755
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   13
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "四舍六入五成双,即:银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一"
            Top             =   1995
            Width           =   1755
         End
         Begin VB.CheckBox chk 
            Caption         =   "从属项目汇总计算折扣额"
            Height          =   195
            Index           =   39
            Left            =   165
            TabIndex        =   16
            Top             =   5280
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "输入费用项目时先输类别"
            Height          =   195
            Index           =   25
            Left            =   165
            TabIndex        =   18
            Top             =   5970
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   5
            Left            =   2670
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   4335
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(5)"
            BuddyDispid     =   196631
            BuddyIndex      =   5
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   4
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   10
            Text            =   "15"
            Top             =   3975
            Width           =   930
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1740
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   12
            Text            =   "0"
            Top             =   4335
            Width           =   930
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   4
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   1755
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   2
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   315
            Width           =   1755
         End
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   2150
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "1"
            Top             =   2895
            Width           =   520
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   12
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "挂号暂不支持四舍五入六成双"
            Top             =   1680
            Width           =   1755
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   1
            Left            =   2670
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2895
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(1)"
            BuddyDispid     =   196631
            BuddyIndex      =   1
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   0
            Left            =   2670
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3975
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   15
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196631
            BuddyIndex      =   0
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   365
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   6
            Left            =   2670
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   4695
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtUD(6)"
            BuddyDispid     =   196631
            BuddyIndex      =   6
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   5
            Min             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   10
            Left            =   2670
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(10)"
            BuddyDispid     =   196631
            BuddyIndex      =   10
            OrigLeft        =   1845
            OrigTop         =   810
            OrigRight       =   2085
            OrigBottom      =   1110
            Max             =   7
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "急诊挂号单有效的天数"
            Height          =   180
            Index           =   49
            Left            =   195
            TabIndex        =   238
            Top             =   3285
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用单价保留位数"
            Height          =   180
            Index           =   35
            Left            =   195
            TabIndex        =   235
            Top             =   4755
            Width           =   1440
         End
         Begin VB.Label Label9 
            Caption         =   "零钞处理规则"
            Height          =   255
            Left            =   195
            TabIndex        =   200
            Top             =   1470
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费项目和诊疗项目输入匹配方式"
            Height          =   180
            Index           =   40
            Left            =   165
            TabIndex        =   199
            ToolTipText     =   "影响收费,记帐的收费项目输入,医生,护士输入医嘱"
            Top             =   6840
            Width           =   2700
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "结帐"
            Height          =   180
            Index           =   38
            Left            =   675
            TabIndex        =   192
            Top             =   2355
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "收费"
            Height          =   180
            Index           =   37
            Left            =   675
            TabIndex        =   191
            Top             =   2055
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "挂号允许预约天数"
            Height          =   180
            Index           =   30
            Left            =   195
            TabIndex        =   138
            Top             =   4035
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费用金额保留位数"
            Height          =   180
            Index           =   28
            Left            =   195
            TabIndex        =   139
            Top             =   4395
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "门诊号规则"
            Height          =   180
            Index           =   22
            Left            =   195
            TabIndex        =   135
            Top             =   1140
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "住院号规则"
            Height          =   180
            Index           =   10
            Left            =   195
            TabIndex        =   134
            Top             =   390
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "普通挂号单有效的天数"
            Height          =   180
            Index           =   16
            Left            =   195
            TabIndex        =   137
            Top             =   2955
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "挂号"
            Height          =   180
            Index           =   15
            Left            =   675
            TabIndex        =   136
            Top             =   1740
            Width           =   360
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " 特定收费项目 "
         Height          =   1850
         Index           =   7
         Left            =   6840
         TabIndex        =   148
         Top             =   4755
         Width           =   2775
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Height          =   240
            Index           =   4
            Left            =   2415
            TabIndex        =   320
            TabStop         =   0   'False
            Top             =   1350
            Width           =   255
         End
         Begin VB.TextBox txtCmd 
            Height          =   300
            Index           =   4
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   318
            Top             =   1320
            Width           =   1710
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Height          =   240
            Index           =   3
            Left            =   2415
            TabIndex        =   317
            TabStop         =   0   'False
            Top             =   1005
            Width           =   255
         End
         Begin VB.TextBox txtCmd 
            Height          =   300
            Index           =   3
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   315
            Top             =   975
            Width           =   1710
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Height          =   240
            Index           =   0
            Left            =   2415
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   310
            Width           =   255
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "…"
            Height          =   240
            Index           =   1
            Left            =   2415
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   660
            Width           =   255
         End
         Begin VB.TextBox txtCmd 
            Height          =   300
            Index           =   0
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   280
            Width           =   1710
         End
         Begin VB.TextBox txtCmd 
            Height          =   300
            Index           =   1
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   630
            Width           =   1710
         End
         Begin VB.Label lbl 
            Caption         =   "肿瘤配置费"
            Height          =   225
            Index           =   56
            Left            =   45
            TabIndex        =   319
            Top             =   1365
            Width           =   1050
         End
         Begin VB.Label lbl 
            Caption         =   "普通配置费"
            Height          =   225
            Index           =   18
            Left            =   45
            TabIndex        =   316
            Top             =   1005
            Width           =   1050
         End
         Begin VB.Label lbl 
            Caption         =   "病历费"
            Height          =   225
            Index           =   6
            Left            =   285
            TabIndex        =   149
            Top             =   318
            Width           =   585
         End
         Begin VB.Label lbl 
            Caption         =   "工本费"
            Height          =   225
            Index           =   7
            Left            =   285
            TabIndex        =   150
            Top             =   668
            Width           =   585
         End
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   5
      Left            =   240
      TabIndex        =   158
      Top             =   450
      Width           =   9690
      Begin VB.CommandButton cmdWarnDel 
         Caption         =   "删除报警方案(&D)"
         Height          =   350
         Left            =   7920
         TabIndex        =   92
         Top             =   7635
         Width           =   1710
      End
      Begin VB.CommandButton cmdWarnNew 
         Caption         =   "增加报警方案(&A)"
         Height          =   350
         Left            =   7920
         TabIndex        =   91
         Top             =   7275
         Width           =   1710
      End
      Begin VB.CheckBox chk 
         Caption         =   "记帐报警包含划价费用"
         Height          =   255
         Index           =   41
         Left            =   7515
         TabIndex        =   90
         ToolTipText     =   "在记帐报警判断时,病人费用累计或当日费用累计中是否包含未审核的划价单费用"
         Top             =   90
         Width           =   2175
      End
      Begin VB.ListBox lst类别 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2130
         Left            =   2745
         Style           =   1  'Checkbox
         TabIndex        =   89
         Top             =   900
         Visible         =   0   'False
         Width           =   1530
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   6345
         Index           =   1
         Left            =   90
         TabIndex        =   88
         Top             =   765
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   11192
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
      Begin MSComctlLib.TabStrip tab报警 
         Height          =   6795
         Left            =   15
         TabIndex        =   87
         Top             =   420
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   11986
         HotTracking     =   -1  'True
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "普通病人"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSystemPara.frx":1782
         Height          =   555
         Left            =   135
         TabIndex        =   174
         Top             =   7425
         Width           =   7740
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "报警方案：每种方案包括各病区报警线及报警方式，需和 zl_PatiWarnScheme 函数配合使用"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   14
         Left            =   105
         TabIndex        =   159
         Top             =   120
         Width           =   7290
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   4
      Left            =   240
      TabIndex        =   155
      Top             =   495
      Width           =   9690
      Begin VB.OptionButton opt护理 
         Caption         =   "以价格最高的护理等级为标准"
         Height          =   255
         Index           =   1
         Left            =   5385
         TabIndex        =   86
         Top             =   7695
         Width           =   2670
      End
      Begin VB.OptionButton opt护理 
         Caption         =   "以最后一次护理等级为标准"
         Height          =   255
         Index           =   0
         Left            =   2850
         TabIndex        =   85
         Top             =   7695
         Value           =   -1  'True
         Width           =   2625
      End
      Begin VB.CheckBox chk 
         Caption         =   "下午算半天模式 (指以半天为计算单位,上午入院算1天,下午算半天,上午出院当天不算费用,下午算半天)"
         Height          =   225
         Index           =   43
         Left            =   210
         TabIndex        =   83
         ToolTipText     =   "表示是否自动修改上一核算期间的自动费用计算数据"
         Top             =   7470
         Width           =   8775
      End
      Begin VB.TextBox txtDateInput 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Left            =   8265
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
         Height          =   6675
         Index           =   0
         Left            =   165
         TabIndex        =   79
         Top             =   405
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   11774
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   315
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   6660
         Index           =   0
         Left            =   4980
         TabIndex        =   80
         Top             =   405
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   11748
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
      Begin VB.CheckBox chk 
         Caption         =   "修正上期自动计费(表示是否自动修改上一核算期间的自动费用计算数据。)"
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   81
         ToolTipText     =   "表示是否自动修改上一核算期间的自动费用计算数据"
         Top             =   7185
         Width           =   6510
      End
      Begin VB.Label lbl护理 
         AutoSize        =   -1  'True
         Caption         =   "同天不同护理等级的护理费计算"
         Height          =   180
         Left            =   210
         TabIndex        =   84
         Top             =   7770
         Width           =   2520
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按病区对指定费用进行自动计算"
         Height          =   180
         Index           =   13
         Left            =   5025
         TabIndex        =   157
         Top             =   105
         Width           =   2520
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对床位费或护理费进行自动计算"
         Height          =   180
         Index           =   12
         Left            =   210
         TabIndex        =   156
         Top             =   120
         Width           =   2520
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   15
      Left            =   285
      TabIndex        =   203
      Top             =   570
      Width           =   9690
      Begin TabDlg.SSTab sstabDigit 
         Height          =   8010
         Left            =   0
         TabIndex        =   228
         Top             =   0
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   14129
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "  录入精度"
         TabPicture(0)   =   "frmSystemPara.frx":1863
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label10"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label23"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "billDigit(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin ZL9BillEdit.BillEdit billDigit 
            Height          =   6420
            Index           =   0
            Left            =   120
            TabIndex        =   229
            Top             =   720
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   11324
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
         Begin VB.Label Label23 
            Caption         =   $"frmSystemPara.frx":187F
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   195
            TabIndex        =   231
            Top             =   7230
            Width           =   9195
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "按药品、卫材的包装单位来设置价格、数量允许录入的精度（保留的小数位数）"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   230
            Top             =   480
            Width           =   7350
         End
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   14
      Left            =   210
      TabIndex        =   182
      Top             =   420
      Width           =   9690
      Begin ZL9BillEdit.BillEdit Bill药房配药控制 
         Height          =   7590
         Left            =   165
         TabIndex        =   128
         Top             =   420
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   13388
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
         Left            =   240
         TabIndex        =   183
         Top             =   150
         Width           =   1080
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   13
      Left            =   315
      TabIndex        =   184
      Top             =   450
      Width           =   9675
      Begin ZL9BillEdit.BillEdit mshBillEdit 
         Height          =   7080
         Left            =   165
         TabIndex        =   126
         Top             =   480
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   12488
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
      Begin ZL9BillEdit.BillEdit mshBillEditStuff 
         Height          =   7080
         Left            =   4680
         TabIndex        =   127
         Top             =   480
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   12488
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "注意：科室编号可选范围A-Z、1-9，同组中科室编号不能重复。"
         Height          =   285
         Left            =   195
         TabIndex        =   190
         Top             =   7680
         Width           =   5040
      End
      Begin VB.Label Label2 
         Caption         =   "请填写卫材科室对应的编号"
         Height          =   285
         Left            =   4680
         TabIndex        =   189
         Top             =   180
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "请填写药品科室对应的编号"
         Height          =   285
         Left            =   165
         TabIndex        =   185
         Top             =   180
         Width           =   4335
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   12
      Left            =   225
      TabIndex        =   180
      Top             =   420
      Width           =   9690
      Begin MSComctlLib.ListView lvwNo 
         Height          =   7470
         Left            =   165
         TabIndex        =   125
         Top             =   480
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   13176
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iltC32"
         SmallIcons      =   "imgC16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "单据类型"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "编码规则"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "请选择单据对应的编码方式(双击行可以修改编码方式)"
         Height          =   165
         Left            =   120
         TabIndex        =   181
         Top             =   180
         Width           =   8115
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   11
      Left            =   165
      TabIndex        =   178
      Top             =   450
      Width           =   9690
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   7590
         Index           =   4
         Left            =   210
         TabIndex        =   124
         Top             =   420
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   13388
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
         Caption         =   "控制药品的领用库房"
         Height          =   180
         Index           =   33
         Left            =   240
         TabIndex        =   179
         Top             =   150
         Width           =   1620
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   10
      Left            =   255
      TabIndex        =   177
      Top             =   435
      Width           =   9600
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf库房单位 
         Height          =   7815
         Left            =   165
         TabIndex        =   123
         Top             =   195
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   13785
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
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   9
      Left            =   285
      TabIndex        =   175
      Top             =   480
      Width           =   9690
      Begin MSComctlLib.ListView lvwCheckMed 
         Height          =   7335
         Left            =   165
         TabIndex        =   122
         Top             =   660
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   12938
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "编码"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "部门名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "库存检查方式"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmSystemPara.frx":1908
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lbl提示 
         Caption         =   "    在这里可以选择各库房是否检查库存及库存检查方式。当库房选中时双击或按“C”键可改变库房的检查方式。"
         Height          =   435
         Left            =   1455
         TabIndex        =   176
         Top             =   165
         Width           =   5775
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   8
      Left            =   210
      TabIndex        =   172
      Top             =   450
      Width           =   9690
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   7560
         Index           =   3
         Left            =   165
         TabIndex        =   121
         Top             =   420
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13335
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
         Caption         =   "控制药品在不同库房间的流通方向"
         Height          =   180
         Index           =   23
         Left            =   240
         TabIndex        =   173
         Top             =   150
         Width           =   2700
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   7
      Left            =   360
      TabIndex        =   170
      Top             =   480
      Width           =   9570
      Begin VB.TextBox txtMaxMoney 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2220
         MaxLength       =   12
         TabIndex        =   120
         ToolTipText     =   "用于对输入的单笔费用金额进行检查，当超过设置的金额时就进行提醒，以防止输入错误"
         Top             =   7665
         Width           =   1350
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   8280
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystemPara.frx":1F89
               Key             =   "Limit"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystemPara.frx":23DB
               Key             =   "bm"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystemPara.frx":2975
               Key             =   "社区"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystemPara.frx":2F0F
               Key             =   "UnCheck"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSystemPara.frx":34A9
               Key             =   "AllCheck"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "清除(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   3
         Left            =   8175
         TabIndex        =   118
         Top             =   1965
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "删除(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   2
         Left            =   8175
         TabIndex        =   117
         Top             =   1485
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "修改(&M)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   1
         Left            =   8175
         TabIndex        =   116
         Top             =   1005
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "增加(&A)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   0
         Left            =   8175
         TabIndex        =   115
         Top             =   525
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   7095
         Index           =   1
         Left            =   165
         TabIndex        =   114
         Top             =   480
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   12515
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "操作人"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "单据类型"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "历史天数"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "允许操作他人单据"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "金额上限"
            Object.Width           =   2187
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单笔费用最大提醒金额："
         Height          =   180
         Left            =   270
         TabIndex        =   119
         Top             =   7725
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "按操作员对不同单据的操作权限，针对单据的历史天数和最初操作人进行限制"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   171
         Top             =   225
         Width           =   6120
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   129
      Top             =   8955
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8430
      TabIndex        =   130
      Top             =   8955
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   435
      TabIndex        =   131
      Top             =   8955
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   8610
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   15187
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      TabMinWidth     =   883
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   17
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "临床应用"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "电子签名"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "票据和卡"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "自动计算"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "记帐报警"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "权限"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "单据操作"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品流向"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "库房检查"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品库房单位"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品领用流向"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "单据编码规则"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "科室编号"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药房配药控制"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "药品卫材精度"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "单据环节控制"
            ImageVarType    =   2
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
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   16
      Left            =   240
      TabIndex        =   204
      Top             =   600
      Width           =   9690
      Begin VSFlex8Ctl.VSFlexGrid vsfControlItem 
         Height          =   7605
         Left            =   165
         TabIndex        =   205
         Top             =   360
         Width           =   9420
         _cx             =   16616
         _cy             =   13414
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
         Caption         =   "设置药品卫材单据在特定业务环节中允许修改的项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   206
         Top             =   15
         Width           =   4620
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8080
      Index           =   6
      Left            =   240
      TabIndex        =   160
      Top             =   480
      Width           =   9690
      Begin VB.Frame FraChangeDept 
         Caption         =   "病人转科或出院"
         Height          =   1080
         Left            =   165
         TabIndex        =   326
         Top             =   2790
         Width           =   4305
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   29
            ItemData        =   "frmSystemPara.frx":3A43
            Left            =   1905
            List            =   "frmSystemPara.frx":3A45
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   630
            Width           =   2205
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   28
            ItemData        =   "frmSystemPara.frx":3A47
            Left            =   1905
            List            =   "frmSystemPara.frx":3A49
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   255
            Width           =   2205
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(出院)超期护理数据"
            Height          =   180
            Index           =   8
            Left            =   210
            TabIndex        =   344
            Top             =   690
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(转科)未审销帐单据"
            Height          =   180
            Index           =   57
            Left            =   210
            TabIndex        =   327
            Top             =   315
            Width           =   1620
         End
      End
      Begin VB.ComboBox cboPatiVerfy 
         Height          =   300
         ItemData        =   "frmSystemPara.frx":3A4B
         Left            =   5925
         List            =   "frmSystemPara.frx":3A4D
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   2825
         Width           =   3660
      End
      Begin VB.Frame fra补充录入 
         Caption         =   "补充录入限制"
         Height          =   1100
         Left            =   4725
         TabIndex        =   236
         Top             =   6870
         Width           =   4890
         Begin VB.CheckBox chk 
            Caption         =   "转病区病人只允许补录临嘱"
            Height          =   210
            Index           =   78
            Left            =   240
            TabIndex        =   108
            Top             =   720
            Value           =   1  'Checked
            Width           =   2520
         End
         Begin VB.TextBox txtInputHours 
            Height          =   300
            Left            =   1700
            MaxLength       =   4
            TabIndex        =   107
            Top             =   300
            Width           =   765
         End
         Begin VB.Label lbl补充录入 
            AutoSize        =   -1  'True
            Caption         =   "补录时限(0-9999)"
            Height          =   180
            Left            =   240
            TabIndex        =   239
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label lblInputHours 
            AutoSize        =   -1  'True
            Caption         =   "小时"
            Height          =   180
            Left            =   2520
            TabIndex        =   237
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame fra出院检查副 
         Caption         =   "病人转科或出院(未执行诊疗项目)"
         Height          =   2520
         Left            =   165
         TabIndex        =   232
         Top             =   5460
         Width           =   4305
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   6
            ItemData        =   "frmSystemPara.frx":3A4F
            Left            =   1080
            List            =   "frmSystemPara.frx":3A51
            Style           =   2  'Dropdown List
            TabIndex        =   106
            ToolTipText     =   "在病人结帐以及病人入出管理中出院时检查"
            Top             =   675
            Width           =   3015
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   19
            ItemData        =   "frmSystemPara.frx":3A53
            Left            =   1080
            List            =   "frmSystemPara.frx":3A55
            Style           =   2  'Dropdown List
            TabIndex        =   105
            ToolTipText     =   "在病人入出管理中转科时检查"
            Top             =   315
            Width           =   3015
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUnCheckItem 
            Height          =   1125
            Left            =   240
            TabIndex        =   343
            Top             =   1320
            Width           =   3900
            _cx             =   6879
            _cy             =   1984
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSystemPara.frx":3A57
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
         Begin VB.Label Label17 
            Caption         =   "不检查以下未执行诊疗项目："
            Height          =   255
            Left            =   240
            TabIndex        =   342
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "出院时"
            Height          =   180
            Index           =   50
            Left            =   255
            TabIndex        =   234
            Top             =   705
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "转科时"
            Height          =   180
            Index           =   17
            Left            =   255
            TabIndex        =   233
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "门诊一卡通"
         Height          =   2240
         Left            =   4710
         TabIndex        =   214
         Top             =   75
         Width           =   4875
         Begin VB.CheckBox chk 
            Caption         =   "项目开单后立即收费或记帐审核"
            Height          =   210
            Index           =   90
            Left            =   150
            TabIndex        =   337
            Top             =   830
            Width           =   3120
         End
         Begin VB.CheckBox chk 
            Caption         =   "病人消费减少剩余款额时需要刷卡进行验证"
            Height          =   210
            Index           =   59
            Left            =   150
            TabIndex        =   336
            Top             =   270
            Width           =   4080
         End
         Begin VB.CheckBox chk 
            Caption         =   "项目执行前必须先收费或先记帐审核"
            Height          =   210
            Index           =   67
            Left            =   150
            TabIndex        =   335
            Top             =   550
            Width           =   4080
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许未收费的门诊划价处方发药"
            Height          =   195
            Index           =   58
            Left            =   150
            TabIndex        =   334
            Top             =   1110
            Width           =   2880
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许未审核的记帐处方发药"
            Height          =   195
            Index           =   15
            Left            =   150
            TabIndex        =   333
            Top             =   1375
            Value           =   1  'Checked
            Width           =   4425
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许未收费的门诊划价处方发料"
            Height          =   180
            Index           =   68
            Left            =   150
            TabIndex        =   332
            Top             =   1640
            Width           =   2895
         End
         Begin VB.CheckBox chk 
            Caption         =   "允许未审核的记账处方发料"
            Height          =   255
            Index           =   69
            Left            =   150
            TabIndex        =   331
            Top             =   1890
            Width           =   4455
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "入科时必须确定护理等级"
         Height          =   180
         Index           =   42
         Left            =   375
         TabIndex        =   102
         ToolTipText     =   "一个科室可以存在于多个病区,病人入院不分配床位时不确定病区,主要影响病人信息管理,入院管理,入科管理等模块"
         Top             =   3990
         Width           =   2280
      End
      Begin VB.ComboBox cmb 
         Height          =   300
         Index           =   15
         Left            =   5925
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   2440
         Width           =   3660
      End
      Begin VB.Frame fra出院检查 
         Caption         =   "病人转科或出院(未发药品)"
         Height          =   1065
         Left            =   165
         TabIndex        =   165
         Top             =   4290
         Width           =   4305
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   23
            ItemData        =   "frmSystemPara.frx":3A95
            Left            =   1080
            List            =   "frmSystemPara.frx":3A97
            Style           =   2  'Dropdown List
            TabIndex        =   103
            ToolTipText     =   "在病人入出管理中转科时检查"
            Top             =   300
            Width           =   3015
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   22
            ItemData        =   "frmSystemPara.frx":3A99
            Left            =   1080
            List            =   "frmSystemPara.frx":3A9B
            Style           =   2  'Dropdown List
            TabIndex        =   104
            ToolTipText     =   "在病人结帐以及病人入出管理中出院时检查"
            Top             =   660
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "转科时"
            Height          =   180
            Index           =   48
            Left            =   270
            TabIndex        =   226
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "出院时"
            Height          =   180
            Index           =   46
            Left            =   255
            TabIndex        =   227
            Top             =   720
            Width           =   540
         End
      End
      Begin VB.Frame Fra药库流通 
         Caption         =   "药库单据审核"
         Height          =   735
         Left            =   4710
         TabIndex        =   166
         Top             =   3230
         Width           =   4890
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   7
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   270
            Width           =   1380
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "开单人与审核人"
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   167
            Top             =   330
            Width           =   1260
         End
      End
      Begin VB.Frame fra结帐 
         Caption         =   "住院结帐"
         Height          =   960
         Left            =   165
         TabIndex        =   163
         Top             =   1725
         Width           =   4305
         Begin VB.CheckBox chk 
            Caption         =   "在院病人不允许出院结帐"
            Height          =   210
            Index           =   55
            Left            =   240
            TabIndex        =   99
            Top             =   675
            Width           =   2520
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   5
            ItemData        =   "frmSystemPara.frx":3A9D
            Left            =   1410
            List            =   "frmSystemPara.frx":3A9F
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   270
            Width           =   2715
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "未审单据结帐"
            Height          =   180
            Index           =   24
            Left            =   225
            TabIndex        =   164
            Top             =   330
            Width           =   1080
         End
      End
      Begin VB.Frame frm住院记帐 
         Caption         =   "住院记帐"
         Height          =   1545
         Left            =   165
         TabIndex        =   161
         Top             =   75
         Width           =   4305
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   0
            Left            =   1425
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1140
            Width           =   2700
         End
         Begin VB.CheckBox chk 
            Caption         =   "操作员只限以本人身份登记"
            Height          =   210
            Index           =   20
            Left            =   165
            TabIndex        =   94
            Top             =   570
            Width           =   2520
         End
         Begin VB.CheckBox chk 
            Caption         =   "可以输入其它科室的开单人"
            Height          =   210
            Index           =   19
            Left            =   165
            TabIndex        =   93
            Top             =   285
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "必须输入开单人"
            Height          =   210
            Index           =   18
            Left            =   2610
            TabIndex        =   96
            Top             =   855
            Width           =   1590
         End
         Begin VB.CheckBox chk 
            Caption         =   "病人未入科禁止记账操作"
            Height          =   210
            Index           =   84
            Left            =   165
            TabIndex        =   95
            Top             =   855
            Width           =   2340
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "已结单据操作"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   162
            Top             =   1200
            Width           =   1080
         End
      End
      Begin VB.ListBox lst 
         Height          =   2370
         Index           =   1
         Left            =   7395
         Style           =   1  'Checkbox
         TabIndex        =   113
         Top             =   4425
         Width           =   2220
      End
      Begin VB.ListBox lst 
         Height          =   2370
         Index           =   0
         Left            =   4725
         Style           =   1  'Checkbox
         TabIndex        =   112
         Top             =   4425
         Width           =   2220
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "病人审核方式"
         Height          =   180
         Index           =   52
         Left            =   4725
         TabIndex        =   249
         Top             =   2885
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医保对码检查"
         Height          =   180
         Index           =   42
         Left            =   4725
         TabIndex        =   209
         Top             =   2495
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "公费病人适用费用类型"
         Height          =   180
         Index           =   21
         Left            =   7395
         TabIndex        =   169
         Top             =   4140
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医保病人适用费用类型"
         Height          =   180
         Index           =   20
         Left            =   4770
         TabIndex        =   168
         Top             =   4140
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmSystemPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum const数
    ud_挂号预约天数 = 0
    ud_挂号单 = 1
    'ud_收费收据 = 2:56963
    ud_门诊处方条数限制 = 3
    ud_号码长度 = 4
    ud_费用金额保留位数 = 5
    ud_费用单价保留位数 = 6
    ud_过敏登记有效天数 = 7
    ud_补录医嘱识别间隔 = 8
    ud_儿童年龄界定上限 = 9
    ud_急诊挂号单 = 10
    ud_门诊新开医嘱间隔 = 11
End Enum

Private Enum constChk
    chk_未作废临嘱禁止退药 = 0
    'chk_加收工本费 = 1:56963
    chk_限定药品的库存 = 2
    chk_药品按规格下医嘱 = 3
    chk_收取预交款 = 4
    chk_时办理就诊卡 = 5
    chk_分配床位号 = 6
    chk_病人姓名 = 7
    chk_病人ID = 8
    chk_刷就诊卡 = 9
    chk_挂号单号 = 10
    chk_过敏登记有效天数 = 11
    chk_自动修正 = 12
    chk_票号控制 = 13
    'chk_密文显示 = 14
    chk_时价分段加成入库 = 14
    chk_未审核记帐处方发药 = 15
    chk_门诊退费须先申请 = 16
    chk_未收费处方发药 = 58
    
    chk_收费同时发药 = 17
    chk_输入开单人 = 18
    chk_它科开单人 = 19
    chk_本人执行登记 = 20
    chk_时价药品入库 = 21
    chk_门诊收费与发药分离 = 22
    chk_住院记帐与发药分离 = 23
    chk_长期医嘱次日生效 = 24
    chk_首先输入收费类别 = 25
    chk_明确申领药品批次 = 26
    chk_配置中心 = 27
    chk_外购入库需要核查 = 28
    chk_外购入库需要经过标记付款后才能进行付款 = 70
    chk_药品移库明确批次 = 71
    chk_药品领用明确批次 = 72
    'chk_多张单据收费分别打印 = 29:56963
    chk_全数字只查编码 = 30
    chk_全字母只查简码 = 31
    chk_执行后自动审核划价单 = 32
    chk_一次申请多个检验项目 = 34
    'chk_误差项不使用票据 = 35  :56963
    chk_时价药品直接确定售价 = 36
    chk_住院卫材自动发料 = 37
    chk_门诊卫材自动发料 = 38
    chk_执行之后自动发料 = 61
    chk_指定医嘱在其他科室执行 = 62
    
    chk_从属项目汇总计算折扣 = 39
    
    chk_药品填单时下可用库存 = 40
    chk_记帐报警包含划价费用 = 41
    chk_入科确定护理等级 = 42
    chk_下午算半天模式 = 43
    
    chk_电子签名控制_门诊 = 44
    chk_电子签名控制_住院 = 45
    chk_电子签名控制_医技 = 46
    chk_电子签名控制_护理 = 47
    chk_电子签名控制_药品 = 60
    chk_电子签名控制_lis = 1
    chk_电子签名控制_pacs = 29
    
    chk_时价入库按折扣前采购价加成销售 = 48
    'chk_按执行科室分别打印 = 49:56963
    chk_下达出院医嘱才允许出院 = 50
    chk_门诊处方条数限制 = 52
    'chk_收费每次只用一张票据 = 53:56963
    chk_检验医嘱发送生成条形码 = 54
    chk_在院病人不准出院结帐 = 55
    chk_收费项目首位当类别简码 = 56
    chk_每次住院使用新住院号 = 57
    chk_门诊病人消费时需要刷卡验证 = 59
    'chk_就诊卡重复使用 = 63
    chk_住院药嘱发送产生领药号 = 64
    chk_禁忌药嘱 = 65
    chk_下达医嘱时显示产地 = 66
    chk_项目执行前必须收费或审核 = 67
    chk_项目开单后立即收费或记帐审核 = 90 '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
    chk_允许未收费的门诊划价处方发料 = 68
    chk_允许未审核的记账处方发料 = 69
    chk_禁止下达超极量药品医嘱 = 63
    chk_时价药品取上次售价 = 73
    chk_输血和皮试医嘱执行后需要核对 = 74
    chk抗菌药物分级管理 = 75
    chk抗菌药物使用自备药 = 76
    chk允许下达院外执行的禁忌药品医嘱 = 77
    chk只允许补录临嘱 = 78
    chk临床工作站必须使用zlPlugIn部件 = 79
    chk启用手术分级管理 = 80
    chk_回退出院医嘱才允许撤销出院 = 81
    chk_允许处理超过挂号有效天数的病人 = 82
    chk_首次医嘱执行需要审核 = 83
    chk_未入科禁止记账 = 84    '51612
    chk_输血分级管理 = 35
    chk_手术授权管理 = 49
    chk_输血申请三级审核 = 53
    chk_输血申请只能由中级及以上医师提出 = 85
    chk_医嘱执行有效天数 = 87
    chk_启用接口调用日志 = 88   '大通接口日志调用 65522
    chk_允许使用系统设置 = 89   '美康接口系统设置功能控制参数 65198
    chk_医嘱超量时必须输入原因 = 86
    chk_新开医嘱签名时一组医嘱签名一次 = 91
End Enum

Private Enum const日期
    dtp_上午上班 = 0
    dtp_上午下班 = 1
    dtp_下午上班 = 2
    dtp_下午下班 = 3
End Enum

Private Enum constSign
    sst_门诊 = 0
    sst_住院医生 = 1
    sst_住院护士 = 2
    sst_医技 = 3
    sst_护理 = 4
    sst_药品 = 5
    sst_lis = 6
    sst_Pacs = 7
End Enum

Private Enum constDeptCol
    col_选择 = 0
    col_站点 = 1
    col_编码 = 2
    col_名称 = 3
    col_简码 = 4
End Enum

Private Enum constBill
    bill_自动计算 = 0
    bill_记帐报警 = 1
    bill_药品流向 = 3
    bill_药品领用流向 = 4
End Enum

Private Enum constCmb
    cmb_已结单据 = 0
    cmb_诊断输入来源 = 1
    cmb_住院号规则 = 2
    cmb_定价单位 = 3
    cmb_门诊号规则 = 4
    cmb_未审单据结帐 = 5
    cmb_出院时未执行项目检查 = 6
    cmb_药品单据审核 = 7
    cmb_门诊诊断输入 = 8
    cmb_药品编码模式 = 9
    cmb_诊疗编码模式 = 10
    cmb_电子签名认证中心 = 11
    cmb_挂号零钱处理 = 12
    cmb_收费零钱处理 = 13
    cmb_结帐零钱处理 = 14
    cmb_医保对码检查 = 15
    cmb_住院诊断输入 = 16
    cmb_效期显示方式 = 17
    cmb_药品出库优先算法 = 18
    cmb_转科时未执行项目检查 = 19
    cmb_合理用药接口 = 20
    cmb_配置中心 = 21
    cmb_出院时未发药项目检查 = 22
    cmb_转科时未发药项目检查 = 23
    cmd_中药配方 = 26
    cmd_过敏输入来源 = 27
    cmd_转科时未审核销帐单据 = 28
    cmd_出院时超期护理数据 = 29
End Enum

'对应lblINFO
Private Enum lblEnum
    lbl_过敏输入来源 = 0
End Enum

Private Enum constLvw
    lvw_票据 = 0
    lvw_单据 = 1
    lvw_一卡通 = 3
End Enum

Private Enum constListBox
    lst_医保病人 = 0
    lst_公费病人 = 1
    lst_住院发送类别 = 2
    lst_门诊发送类别 = 4    '发送为划价单的诊疗类别
    lst_刷卡密码 = 3
End Enum

Private Enum constOpt
    opt_闲忙方式 = 2
    opt_平均方式 = 3
End Enum

Private Enum mGrdCol
    选择 = 0
    科室
    号码
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

'变量声明
Private mrsWarn As ADODB.Recordset
Private mrs类别 As ADODB.Recordset
Private mblnChange As Boolean     '是否改变了
Private mblnInit As Boolean       '是否初始化失败
Private mblnLoad As Boolean
Private mintColumn As Integer '
Private mDecimal As Integer       '判断费用金额保留小数位是否改变
Private pDecimal As Integer       '判断费用单价保留小数位是否改变
Private mlngFindItem As Long

Private mrsAdvice As New ADODB.Recordset '记录医嘱内容定义
Private mblnJRaiseByDate As Boolean     '判断床位类项目及从属项目是否按日调价
Private mblnHRaiseByDate As Boolean     '判断护理类项目及从属项目是否按日调价
Private mblnMin As Boolean
Private mstrDel适用病人 As String           '记录记帐报警中删除的适用病人类型
Private mcol科室 As Collection '不用填写超量说明的科室

'记录最后编辑的科室编号所在行、列和编号值
Private mintLastRow_Drug As Integer          '行
Private mintLastCol_Drug As Integer          '列
Private mstrLastCode_Drug As String          '编号

Private mintLastRow_Stuff As Integer          '行
Private mintLastCol_Stuff As Integer          '列
Private mstrLastCode_Stuff As String          '编号

'自动计算设置中保存当前行和列
Private mintCurRow As Integer
Private mintCurCol As Integer

''''''药品卫材单据环节项目控制
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

Private Function Check是否有未审核的药品单据() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Id From 药品收发记录 Where (单据 In(6,7,11) Or (单据 In(1,2,3,4,12) And 入出系数*实际数量<0)) And 审核日期 Is Null And ROWNUM<2"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check是否有未审核的药品单据 = (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check是否有未审核的外购入库单() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Id From 药品收发记录 Where 单据=1 And 审核日期 Is Null And ROWNUM<2"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check是否有未审核的外购入库单 = (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsDrugOrStuff(ByVal strID As String) As Boolean
    '判断是否为药品类别
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select id From 收费细目 Where 类别 In('4','5','6','7') and id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    
    IsDrugOrStuff = rs.RecordCount > 0
    rs.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Public Function IsRaiseByDate(ByVal strID As String) As Boolean
    '判断该收费项目是否是按日调价
    '返回True-是按天条件
    '返回False-不是按天调价
    'strID='J' -床位项目
    'strID='H' -护理项目
    'strID=数字 -其他指定的项目
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    If strID = "J" Then
        strSQL = "Select ID" & _
              " From 收费价目 " & _
              " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And 执行日期 <> Trunc(执行日期, 'dd') And " & _
              " 收费细目id In " & _
              " (Select ID " & _
              " From 收费项目目录 " & _
              " Where 类别 = [1] " & _
              " Union All " & _
              " Select 从项id From 收费从属项目 Where 主项id In (Select ID From 收费项目目录 Where 类别 = [1])) "
    ElseIf strID = "H" Then
            strSQL = "Select ID" & _
              " From 收费价目 " & _
              " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And 执行日期 <> Trunc(执行日期, 'dd') And " & _
              " 收费细目id In " & _
              " (Select ID " & _
              " From 收费项目目录 " & _
              " Where 类别 = [1] " & _
              " Union All " & _
              " Select 从项id From 收费从属项目 Where 主项id In (Select ID From 收费项目目录 Where 类别 = [1])) "
    ElseIf Val(strID) <> 0 Then
        strSQL = "Select Id" & _
                " From 收费价目 " & _
                " Where Nvl(终止日期, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
                " And 执行日期<>trunc(执行日期,'dd') And (收费细目id = [2] or 收费细目id in (Select 从项id From 收费从属项目 Where 主项id = [2])) "
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strID, Val(strID))
    
    IsRaiseByDate = Not (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Load单据环节控制()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo ErrHandle
    intAllItems = UBound(Split(cst所有项目, ",")) + 1
    
    With vsfControlItem
        .Rows = 7
        .Cols = 2 + intAllItems
        .FixedRows = 1
        .FixedCols = 2
        .RowHeightMin = 500
        
        .TextMatrix(0, 0) = "单据"
        .TextMatrix(0, 1) = "环节"
                        
        .ColWidth(0) = 950
        .ColWidth(1) = 950
                        
        For n = 0 To UBound(Split(cst所有项目, ","))
            .TextMatrix(0, n + 2) = Split(cst所有项目, ",")(n)
            .ColWidth(n + 2) = 920
            .ColAlignment(n + 2) = flexAlignCenterCenter
        Next
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        .CellBorderRange 0, 0, 0, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(1, 0) = "药品外购"
        .TextMatrix(2, 0) = "药品外购"
        .TextMatrix(3, 0) = "药品外购"

        .TextMatrix(1, 1) = "核查"
        .TextMatrix(2, 1) = "审核"
        .TextMatrix(3, 1) = "财务审核"
        
        .CellBorderRange 3, 0, 3, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(4, 0) = "卫材外购"
        .TextMatrix(5, 0) = "卫材外购"
        .TextMatrix(6, 0) = "卫材外购"

        .TextMatrix(4, 1) = "核查"
        .TextMatrix(5, 1) = "审核"
        .TextMatrix(6, 1) = "财务审核"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select 单据,环节,内容 From 单据环节控制 Order By 单据, 环节"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "单据环节控制")
        
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
                            Case 单据.卫材外购
                                Select Case rsTmp!环节
                                    Case 环节.核查
                                        .TextMatrix(4, m) = "√"
                                    Case 环节.审核
                                        .TextMatrix(5, m) = "√"
                                    Case 环节.财务审核
                                        .TextMatrix(6, m) = "√"
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
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
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
    gstrSQL = "Select 成本价, 零售价, 实际数量,零售金额 From 药品收发记录 Where Rownum = 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
    mblnMin = (rs.RecordCount > 0)
    
    intMaxCost = IIF(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIF(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIF(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIF(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

    With billDigit(0)
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
        .ColWidth(dig_精度类别) = 1000
        .ColWidth(dig_精度内容) = 1000
        .ColWidth(dig_精度单位) = 1000
        .ColWidth(dig_精度) = 1100
        .ColWidth(dig_最小精度) = 1000
        .ColWidth(dig_最大精度) = 1000
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
            " From 药品卫材精度 Order By 性质, 类别, 内容, 单位"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取药品卫材最大精度")
    
    With billDigit(0)
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
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save药房配药控制()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "ZL_药房配药控制_DELETE"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With Me.Bill药房配药控制
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                gstrSQL = "ZL_药房配药控制_INSERT(" & .RowData(i) & "," & IIF(.TextMatrix(i, 1) = "门诊", 1, 2) & "," & IIF(.TextMatrix(i, 2) <> "", 1, 0) & "," & IIF(Val(.TextMatrix(i, 3)) = 0, "Null", Val(.TextMatrix(i, 3))) & "," & IIF(.TextMatrix(i, 4) <> "", 1, 0) & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
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
    With billDigit(0)
        For n = 1 To .Rows - 1
            strInput = strInput & "0," & _
                .TextMatrix(n, dig_类别) & "," & _
                .TextMatrix(n, dig_内容) & "," & _
                .TextMatrix(n, dig_单位) & "," & _
                .TextMatrix(n, dig_精度) & ";"
        Next
    End With
    
    gstrSQL = "ZL_药品卫材精度_Update('" & strInput & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
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
    
    On Error GoTo ErrHandle
    With vsfControlItem
        For n = 1 To .Rows - 1
            Select Case .TextMatrix(n, 0)
                Case "药品外购"
                    int单据 = 单据.药品外购
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
                    str内容 = str内容 & IIF(str内容 <> "", ",", "") & .TextMatrix(0, m)
                End If
            Next
            
            If str内容 <> "" Then
                strInput = strInput & IIF(strInput <> "", ";", "") & int单据 & "," & int环节 & "," & str内容
            End If
        Next
    End With
    
    gstrSQL = "Zl_单据环节控制_Update('" & strInput & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub bill_AfterAddRow(Index As Integer, Row As Long)
    If Index = bill_记帐报警 Then
        With Bill(Index)
            .TextMatrix(Row, 3) = " "
            .TextMatrix(Row, 4) = " "
            .TextMatrix(Row, 5) = " "
            .TextMatrix(Row, 6) = ""
            .TextMatrix(Row, 7) = ""
        End With
    End If
    
    If Index = bill_自动计算 Then
        With Bill(Index)
            .TextMatrix(Row, 3) = "0-按收治日"
            .TextMatrix(Row, 4) = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
        End With
    End If
End Sub

Private Sub Bill_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    If Index = bill_记帐报警 Then
        With Bill(Index)
            If .TextMatrix(Row, 0) <> "" And .TextMatrix(Row, 2) <> "" Then mblnChange = True
        End With
    End If
End Sub

Private Sub Bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            End If
        End If
    End With
End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    '禁止输入报警类别
    With Bill(Index)
        If Index = bill_记帐报警 And .Col >= 3 Then
            If .Col = 6 Or .Col = 7 Then
                .TxtEnable = True
            Else
                .TxtEnable = False
            End If
        Else
            .TxtEnable = True
        End If
        
        If Index = bill_记帐报警 And .Col = 4 Then  '报警方式2
            If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                .ColData(4) = 5 '每日费用不能编辑报警方式2
            Else
                .ColData(4) = 1
            End If
        End If
        If Index = bill_记帐报警 Then
            Select Case .Col
            Case 6, 7
                .ColData(.Col) = 4
            Case Else
            End Select
        End If
    End With
    
End Sub

Private Sub bill_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    With Bill(Index)
        If Index = bill_记帐报警 And .MouseCol >= 3 And .MouseRow > 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub bill_Validate(Index As Integer, Cancel As Boolean)
    Dim lngRow As Long
    
    If Index = bill_记帐报警 Then
        If Not mblnChange Then Exit Sub
        If MouseInRect(cmdCancel.hwnd) Then Exit Sub
        
        '检查记帐报警设置
        If Not Check记帐报警 Then Cancel = True: Exit Sub
        
        '收集记帐报警数据
        With mrsWarn
            .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        
        With Bill(bill_记帐报警)
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!适用病人 = tab报警.SelectedItem.Caption
                    
                    If .RowData(lngRow) <> 0 Then
                        mrsWarn!病区id = .RowData(lngRow)
                        mrsWarn!病区码 = Split(.TextMatrix(lngRow, 0), "-")(0)
                        mrsWarn!病区名 = Split(.TextMatrix(lngRow, 0), "-")(1)
                    End If
                    
                    mrsWarn!报警方法 = CInt(Left(.TextMatrix(lngRow, 1), 1))
                    mrsWarn!报警值 = CCur(.TextMatrix(lngRow, 2))
                    
                    mrsWarn!报警标志1 = Get类别编码串(.TextMatrix(lngRow, 3))
                    mrsWarn!报警标志2 = Get类别编码串(.TextMatrix(lngRow, 4))
                    mrsWarn!报警标志3 = Get类别编码串(.TextMatrix(lngRow, 5))
                    
                    mrsWarn!催款下限 = Round(Val(.TextMatrix(lngRow, 6)), 2)
                    mrsWarn!催款标准 = Round(Val(.TextMatrix(lngRow, 7)), 2)
                    
                    mrsWarn.Update
                End If
            Next
        End With
    End If
End Sub

Private Sub billDigit_EnterCell(Index As Integer, Row As Long, Col As Long)
    With billDigit(Index)
        If Col = dig_精度 Then
            .TxtCheck = True
            .TextMask = "123456789"
            .MaxLength = 1
        End If
    End With
End Sub


Private Sub billDigit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With billDigit(0)
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

Private Sub cboPatiVerfy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'Private Sub chk补充录入_Click()
'    If chk补充录入.Value = 1 Then
'        txtInputHours.Enabled = True
'        txtInputHours.BackColor = vbWhite
'        On Error Resume Next
'        txtInputHours.SetFocus
'    Else
'        txtInputHours.Enabled = False
'        txtInputHours.BackColor = &H8000000F
'    End If
'End Sub

Private Sub cmdAdvice_Click()
    If frmAdviceDefine.ShowMe(Me, mrsAdvice) Then
        '标记为已变化,需要保存
        cmdAdvice.Tag = "1"
        mblnChange = True
    End If
End Sub

Private Sub cmdFind_Click()
    Dim i As Long, strFind As String
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    strFind = UCase(Trim(txtFind.Text))
    
    With vsDept(sstSign.Tab)
        For i = mlngFindItem To .Rows - 1
            If .RowHidden(i) = False Then
                If UCase(.TextMatrix(i, col_名称)) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_简码) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_编码) = strFind Then
                    .Row = i: .ShowCell i, col_名称
                    Exit For
                End If
            End If
        Next
        If i < .Rows Then
            mlngFindItem = i + 1
        Else
            If mlngFindItem = 1 Then
                MsgBox "没有找到匹配的部门。", vbInformation, Me.Caption
            Else
                MsgBox "已经查找到最后一个部门了。", vbInformation, Me.Caption
                mlngFindItem = 1
            End If
        End If
    End With
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOneCard_Click(Index As Integer)
    
    Select Case Index
        Case 0
            frmOneCard.mbytInFun = 0
            Call frmOneCard.ShowMe(Me)
            Call LoadOneCard
        Case 1
            If lvw(lvw_一卡通).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_一卡通).SelectedItem
                frmOneCard.mbytInFun = 1
                Call frmOneCard.ShowMe(Me, Mid(.Key, 2), .SubItems(1), .SubItems(2), .SubItems(3), IIF(.SubItems(4) = "启用:标准一卡通", 2, IIF(.SubItems(4) = "启用:仅涉及扣卡", 1, 0)))
                Call LoadOneCard
            End With
        Case 2
            If lvw(lvw_一卡通).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_一卡通).SelectedItem
                If MsgBox("你确实要删除“" & .SubItems(1) & "”吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call frmOneCard.DelOneCardRec(Val(Mid(.Key, 2)))
                    Call LoadOneCard
                End If
            End With
    End Select
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim str姓名 As String, str人员ID As String, str单据 As String
    Dim lng单据 As Long, lng天数 As Long, bln修改他人 As Boolean
    Dim dbl金额上限 As Double
    Dim lst As ListItem
    
    
    Select Case Index
        Case 0 '新增
            If frmBillPrivilege.编辑权限(str姓名, str人员ID, str单据, lng单据, lng天数, bln修改他人, dbl金额上限, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_单据).ListItems
                If lst.Tag = str人员ID And lst.ListSubItems(1).Tag = lng单据 Then
                    MsgBox "本次新增的操作限制已经存在。", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        Case 1 '修改
            If lvw(lvw_单据).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_单据).SelectedItem
                str姓名 = .Text
                str单据 = .SubItems(1)
                lng天数 = Val(.SubItems(2))
                bln修改他人 = (.SubItems(3) = "是")
                dbl金额上限 = Val(.SubItems(4))
                str人员ID = .Tag
                lng单据 = .ListSubItems(1).Tag
            End With
            If frmBillPrivilege.编辑权限(str姓名, str人员ID, str单据, lng单据, lng天数, bln修改他人, dbl金额上限, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_单据).ListItems
                If Not lst Is lvw(lvw_单据).SelectedItem Then
                    If lst.Tag = str人员ID And lst.ListSubItems(1).Tag = lng单据 Then
                        MsgBox "本次改变的操作限制已经存在。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 '删除
            If lvw(lvw_单据).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_单据).SelectedItem
                If MsgBox("你确实要删除“" & .Text & "”对“" & .SubItems(1) & "”的操作限制？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                lvw(lvw_单据).ListItems.Remove .Index
            End With
        Case 3 '清除
            If MsgBox("你确实要删除所有的操作限制？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            lvw(lvw_单据).ListItems.Clear
    End Select
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_单据).ListItems.Add(, , str姓名, , "Limit")
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_单据).SelectedItem
            lst.Text = str姓名
        End If
        lst.SubItems(1) = str单据
        lst.SubItems(2) = lng天数
        lst.SubItems(3) = IIF(bln修改他人 = True, "是", "否")
        lst.SubItems(4) = IIF(Val(dbl金额上限) = 0, "", Format(Val(dbl金额上限), "0.00"))
        lst.Tag = str人员ID
        lst.ListSubItems(1).Tag = lng单据
    End If
    mblnChange = True
End Sub

Private Sub cmdSendPriceType_Click(Index As Integer)
    Dim i As Long, j As Long
    
    If SendPriceType.Tab = 0 Then
        j = lst_门诊发送类别
    Else
        j = lst_住院发送类别
    End If
    With lst(j)
        For i = 0 To .ListCount - 1
            .Selected(i) = IIF(Index = 0, True, False)
        Next
    End With
End Sub

Private Sub cmdWarnDel_Click()
    If tab报警.SelectedItem.Caption = "普通病人" Then
        MsgBox """" & tab报警.SelectedItem.Caption & """报警方案不允许删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要删除""" & tab报警.SelectedItem.Caption & """报警方案吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
        
        '记录删除的适用病人类型
        If InStr(1, mstrDel适用病人, tab报警.SelectedItem.Caption) = 0 Then
            mstrDel适用病人 = IIF(mstrDel适用病人 = "", "", mstrDel适用病人 & ";") & tab报警.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab报警.Tabs.Remove tab报警.SelectedItem.Index
    tab报警.Tabs(1).Selected = True
    
    mblnChange = True
End Sub

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab报警.Tabs.Count
        strSchemes = strSchemes & "," & tab报警.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '复制内容
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "适用病人='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!适用病人 = strName
        mrsWarn!病区id = rsCopy!病区id
        mrsWarn!病区码 = rsCopy!病区码
        mrsWarn!病区名 = rsCopy!病区名
        mrsWarn!报警方法 = rsCopy!报警方法
        mrsWarn!报警值 = rsCopy!报警值
        mrsWarn!报警标志1 = rsCopy!报警标志1
        mrsWarn!报警标志2 = rsCopy!报警标志2
        mrsWarn!报警标志3 = rsCopy!报警标志3
        mrsWarn!催款下限 = rsCopy!催款下限
        mrsWarn!催款标准 = rsCopy!催款标准
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab报警.Tabs.Add , , strName
    tab报警.Tabs(tab报警.Tabs.Count).Selected = True
    
    mblnChange = True
End Sub

Private Sub cmd社区参数_Click()
    Dim objCommunity As Object
    
    If lvw社区.SelectedItem Is Nothing Then Exit Sub
    If lvw社区.SelectedItem.SubItems(4) = "" Then
        MsgBox lvw社区.SelectedItem.SubItems(1) & "没有启用。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '先保存设置数据：因为接口初始化要判断是否启用
    If lvw社区.Tag <> "" Then
        On Error GoTo errH
        gcnOracle.BeginTrans
        Call Save社区接口
        gcnOracle.CommitTrans
        lvw社区.Tag = ""
    End If
    
    '创建部件
    Err.Clear: On Error Resume Next
    Set objCommunity = CreateObject("zlCommunity.clsCommunity")
    Err.Clear: On Error GoTo 0
    
    '调用功能
    If Not objCommunity Is Nothing Then
        If objCommunity.Initialize(gcnOracle) Then
            Call objCommunity.Setup(Val(Mid(lvw社区.SelectedItem.Key, 2)))
        End If
    Else
        MsgBox "社区公共接口没有正确安装！", vbExclamation, gstrSysName
    End If
    
    Set objCommunity = Nothing
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnLoad = False Then Exit Sub
    '以下部分只运行一次
    mblnLoad = False
    If mblnInit = False Then Unload Me
    Call tabMain_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lst类别.Visible Then
            lst类别.Visible = False
            Bill(bill_记帐报警).SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
'    On Error GoTo ErrHandle
    
    mblnLoad = True
    '进行初始化
    Set mcol科室 = New Collection
    vsUnCheckItem.ComboList = "..."
    vsUnWriteDept.ComboList = "..."
    Call InitSystemPara
    Call InitEnv
    Call LoadPara
    
    Call LoadOneCard
    Call Load社区接口
    Call Load单据操作
    Call Load病区
    Call LoadTable
    Call Load药品流向
    Call Load库房检查
    Call Load药品领用流向
    Call Load单据编码规则
    Call InitFace
    Call Load部门
    Call Load药品卫材精度
    Call Load单据环节控制
    
    Call CheckExist
    
    '恢复列宽
    RestoreFlexState msh(0), App.ProductName & "\" & Me.Name
    RestoreFlexState Bill(bill_自动计算), App.ProductName & "\" & Me.Name & bill_自动计算
    RestoreFlexState Bill(bill_记帐报警), App.ProductName & "\" & Me.Name & bill_记帐报警
    RestoreFlexState Bill(bill_药品流向), App.ProductName & "\" & Me.Name & bill_药品流向
    RestoreFlexState Bill(bill_药品领用流向), App.ProductName & "\" & Me.Name & bill_药品领用流向
    '初始化成功
    mblnChange = False
    mblnInit = True
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub CheckExist()
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "Select Rownum From 未发药品记录 Where 单据 In (8,9,10) and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "CheckExist")
    
    If Not rsTemp.EOF Then
        Me.chk(chk_门诊收费与发药分离).Enabled = False
        Me.chk(chk_住院记帐与发药分离).Enabled = False
    Else
        Me.chk(chk_门诊收费与发药分离).Enabled = True
        Me.chk(chk_住院记帐与发药分离).Enabled = True
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub InitEnv()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim blnTmp As Boolean
    
    '初始化窗口，这些是不需要读数据库的
    Dim lngIndex As Long
    
    On Error GoTo ErrHandle
    cmb(cmb_药品编码模式).AddItem "顺序编号"
    cmb(cmb_药品编码模式).AddItem "种类+分类号+顺序编号"
    Call zlControl.CboSetWidth(cmb(cmb_药品编码模式).hwnd, cmb(cmb_药品编码模式).Width * 1.2)
    
    cmb(cmb_效期显示方式).AddItem "0-显示失效期"
    cmb(cmb_效期显示方式).AddItem "1-显示有效期"
    Call zlControl.CboSetWidth(cmb(cmb_效期显示方式).hwnd, cmb(cmb_效期显示方式).Width * 1.2)
    
    cmb(cmb_药品出库优先算法).AddItem "0-按批次先进先出"
    cmb(cmb_药品出库优先算法).AddItem "1-按效期最近先出"
    Call zlControl.CboSetWidth(cmb(cmb_药品出库优先算法).hwnd, cmb(cmb_药品出库优先算法).Width * 1.2)

    cmb(cmb_诊疗编码模式).AddItem "顺序编号"
    cmb(cmb_诊疗编码模式).AddItem "种类+分类号+顺序编号"
    Call zlControl.CboSetWidth(cmb(cmb_诊疗编码模式).hwnd, cmb(cmb_诊疗编码模式).Width * 1.2)
    
    cmb(cmb_诊断输入来源).AddItem "1-可选择输入来源"
    cmb(cmb_诊断输入来源).AddItem "2-按诊断标准输入"
    cmb(cmb_诊断输入来源).AddItem "3-按疾病编码输入"
    cmb(cmb_诊断输入来源).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_诊断输入来源).hwnd, cmb(cmb_诊断输入来源).Width * 1.2)
    
    cmb(cmb_门诊诊断输入).AddItem "1-允许自由输入"
    cmb(cmb_门诊诊断输入).AddItem "2-从数据库提取输入"
    cmb(cmb_门诊诊断输入).AddItem "3-仅医保病人从数据库输入"
    cmb(cmb_门诊诊断输入).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_门诊诊断输入).hwnd, cmb(cmb_门诊诊断输入).Width * 1.4)
    cmb(cmb_住院诊断输入).AddItem "1-允许自由输入"
    cmb(cmb_住院诊断输入).AddItem "2-从数据库提取输入"
    cmb(cmb_住院诊断输入).AddItem "3-仅医保病人从数据库输入"
    cmb(cmb_住院诊断输入).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_住院诊断输入).hwnd, cmb(cmb_住院诊断输入).Width * 1.4)
    
    cmb(cmb_已结单据).AddItem "0-允许"
    cmb(cmb_已结单据).AddItem "1-提示"
    cmb(cmb_已结单据).AddItem "2-禁止"
    cmb(cmb_已结单据).ListIndex = 0
    
    cmb(cmb_医保对码检查).AddItem "0-不进行检查"
    cmb(cmb_医保对码检查).AddItem "1-检查并提醒未对码项目"
    cmb(cmb_医保对码检查).AddItem "2-检查并禁止未对码项目"
    cmb(cmb_医保对码检查).ListIndex = 1
    zlControl.CboSetWidth cmb(cmb_医保对码检查).hwnd, 2100
    
    cmb(cmb_合理用药接口).AddItem "0-未使用"
    cmb(cmb_合理用药接口).AddItem "1-四川美康"
    cmb(cmb_合理用药接口).AddItem "2-上海大通"
    cmb(cmb_合理用药接口).AddItem "3-北京太元通"
    cmb(cmb_合理用药接口).ListIndex = 0
    
    cmb(cmd_中药配方).AddItem "0-三味中药"
    cmb(cmd_中药配方).AddItem "1-四味中药"
    cmb(cmd_中药配方).ListIndex = 0
    
    cmb(cmd_过敏输入来源).AddItem "0-可选择输入来源"
    cmb(cmd_过敏输入来源).AddItem "1-按药品目录输入"
    cmb(cmd_过敏输入来源).AddItem "2-按过敏源输入"
    cmb(cmd_过敏输入来源).ListIndex = 0
    '------------------------------------------------------------------------------------------------------------------
    '6-分币五舍六入:34519
    strTmp = "0-不处理|1-分币四舍五入|2-分币补整收取|3-分币舍分收取|4-分币四舍六入五成双|5-角币三七作五、二舍八入|6-分币五舍六入"
    For i = 0 To UBound(Split(strTmp, "|"))
        '挂号不支持四舍六入五成双,因挂号是使用医保的结算修正过程处理分币,Oracle中没有四舍六入五成双函数
        If i <> 4 Then cmb(cmb_挂号零钱处理).AddItem Split(strTmp, "|")(i)
        cmb(cmb_收费零钱处理).AddItem Split(strTmp, "|")(i)
        cmb(cmb_结帐零钱处理).AddItem Split(strTmp, "|")(i)
    Next
    cmb(cmb_挂号零钱处理).ListIndex = 0
    cmb(cmb_收费零钱处理).ListIndex = 0
    cmb(cmb_结帐零钱处理).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_挂号零钱处理).hwnd, 2300
    zlControl.CboSetWidth cmb(cmb_收费零钱处理).hwnd, 2300
    zlControl.CboSetWidth cmb(cmb_结帐零钱处理).hwnd, 2300
    
    cmb(cmb_住院号规则).AddItem "0-顺序编号"
    cmb(cmb_住院号规则).AddItem "1-年月(YYMM)+顺序号(0000)"
    cmb(cmb_住院号规则).AddItem "2-年(YYYY)+顺序号(00000)"
    cmb(cmb_住院号规则).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_住院号规则).hwnd, 2500
    
    cmb(cmb_门诊号规则).AddItem "0-顺序编号"
    cmb(cmb_门诊号规则).AddItem "1-年月日(YYMMDD)+顺序号(0000)"
    cmb(cmb_门诊号规则).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_门诊号规则).hwnd, 3000

    
    cmb(cmb_定价单位).AddItem "0-售价单位"
    cmb(cmb_定价单位).AddItem "1-药库单位"
    cmb(cmb_定价单位).ListIndex = 0
    
    cmb(cmb_未审单据结帐).AddItem "0-不检查"
    cmb(cmb_未审单据结帐).AddItem "1-检查并提示"
    cmb(cmb_未审单据结帐).AddItem "2-检查并禁止"
    cmb(cmb_未审单据结帐).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_未审单据结帐).hwnd, 2000
    
    cmb(cmb_出院时未执行项目检查).AddItem "0-不检查"
    cmb(cmb_出院时未执行项目检查).AddItem "1-检查并提示"
    cmb(cmb_出院时未执行项目检查).AddItem "2-检查并禁止"
    cmb(cmb_出院时未执行项目检查).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_出院时未执行项目检查).hwnd, 2000
    
    cmb(cmb_转科时未执行项目检查).AddItem "0-不检查"
    cmb(cmb_转科时未执行项目检查).AddItem "1-检查并提示"
    cmb(cmb_转科时未执行项目检查).AddItem "2-检查并禁止"
    cmb(cmb_转科时未执行项目检查).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_转科时未执行项目检查).hwnd, 2000
    
    cmb(cmb_出院时未发药项目检查).AddItem "0-不检查"
    cmb(cmb_出院时未发药项目检查).AddItem "1-检查并提示"
    cmb(cmb_出院时未发药项目检查).AddItem "2-检查并禁止"
    cmb(cmb_出院时未发药项目检查).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_出院时未发药项目检查).hwnd, 2000
    
    cmb(cmb_转科时未发药项目检查).AddItem "0-不检查"
    cmb(cmb_转科时未发药项目检查).AddItem "1-检查并提示"
    cmb(cmb_转科时未发药项目检查).AddItem "2-检查并禁止"
    cmb(cmb_转科时未发药项目检查).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_转科时未发药项目检查).hwnd, 2000
    
    '61429:刘鹏飞,2013-11-11
    cmb(cmd_转科时未审核销帐单据).AddItem "0-不检查"
    cmb(cmd_转科时未审核销帐单据).AddItem "1-检查并提示"
    cmb(cmd_转科时未审核销帐单据).AddItem "2-检查并禁止"
    cmb(cmd_转科时未审核销帐单据).ListIndex = 0
    zlControl.CboSetWidth cmb(cmd_转科时未审核销帐单据).hwnd, 2000
    
    '68953:刘鹏飞,2014-08-12
    cmb(cmd_出院时超期护理数据).AddItem "0-不检查"
    cmb(cmd_出院时超期护理数据).AddItem "1-检查并提示"
    cmb(cmd_出院时超期护理数据).AddItem "2-检查并禁止"
    cmb(cmd_出院时超期护理数据).ListIndex = 0
    zlControl.CboSetWidth cmb(cmd_出院时超期护理数据).hwnd, 2000
    
    cmb(cmb_药品单据审核).AddItem "0-不处理"
    cmb(cmb_药品单据审核).AddItem "1-相同禁止"
    cmb(cmb_药品单据审核).ListIndex = 0
    
    '病人审核方式:49501
    With cboPatiVerfy
        .Clear
        .AddItem "0-未审核不允许结帐": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-审核时不许调整费用和医嘱": .ItemData(.NewIndex) = 1
    End With
    
    '电子签名认证中心
    cmb(cmb_电子签名认证中心).AddItem "不使用电子签名"
    cmb(cmb_电子签名认证中心).AddItem "1-辽宁省数字证书认证中心"
    cmb(cmb_电子签名认证中心).AddItem "2-广西省数字证书认证中心"
    cmb(cmb_电子签名认证中心).AddItem "3-重庆市数字证书认证中心"
    cmb(cmb_电子签名认证中心).AddItem "4-山东省数字证书认证中心"
    cmb(cmb_电子签名认证中心).AddItem "5-吉大正元数字证书认证中心" '-- 原名称叫 吉林中心医院 数字证书认证中心
    cmb(cmb_电子签名认证中心).AddItem "6-国投安信数字证书认证中心" '-- 原名称叫 吉林省医院 数字证书认证中心
    cmb(cmb_电子签名认证中心).AddItem "7-国投安信证书认证中心(内蒙)"     '-- 原名称叫 准格尔医院 数字证书认证中心,11年12月改成用安信的了
    'cmb(cmb_电子签名认证中心).AddItem "9-广东数字证书认证中心(海南)"    ' 还没定用不用
    cmb(cmb_电子签名认证中心).AddItem "10-北京数字证书认证中心(河南)"
    cmb(cmb_电子签名认证中心).AddItem "11-北京数字证书认证中心(四川)"
    cmb(cmb_电子签名认证中心).AddItem "12-北京数字证书认证中心(广西)"    '有时间戳
    cmb(cmb_电子签名认证中心).AddItem "13-北京数字证书认证中心(湖北)"    '有时间戳
    cmb(cmb_电子签名认证中心).AddItem "14-北京数字证书认证中心(辽宁)"
    cmb(cmb_电子签名认证中心).AddItem "15-上海数字证书认证中心(上海)"
    cmb(cmb_电子签名认证中心).AddItem "16-江苏数字证书认证中心(江苏)"   '--江宁医院
    cmb(cmb_电子签名认证中心).AddItem "17-新疆数字证书认证中心(新疆)"   '无时间戳
    cmb(cmb_电子签名认证中心).ListIndex = 0
    
    mlngFindItem = 1
    For i = 0 To sstSign.Tabs - 1
        sstSign.TabVisible(i) = False
        If i = sst_门诊 Then
            strTmp = " And t.服务对象 IN (1,3)  and T.工作性质 IN ('临床','手术','治疗')"
        ElseIf i = sst_住院医生 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质 IN ('临床','手术','治疗')"
        ElseIf i = sst_住院护士 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质='护理'"
        ElseIf i = sst_医技 Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质 IN('检查','检验','手术','治疗','营养')"
        ElseIf i = sst_护理 Then
            strTmp = " And t.服务对象 IN (2,3)  and T.工作性质='护理'"
        ElseIf i = sst_药品 Then
            strTmp = " and T.工作性质 in('西药房','中药房','成药房')"
        ElseIf i = sst_lis Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质='检验'"
        ElseIf i = sst_Pacs Then
            strTmp = " And t.服务对象 <> 0  and T.工作性质='检查'"
        End If
         '加载默认部门选择
        gstrSQL = "Select Distinct D.ID, d.站点,D.编码, D.名称,D.简码" & vbNewLine & _
                "From 部门表 D, 部门性质说明 T" & vbNewLine & _
                "Where d.Id = t.部门id And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & strTmp & vbNewLine & _
                "order by 站点,名称"
    
        Call OpenRecordset(rsTmp, Me.Caption)
        With vsDept(i)
            .Rows = 1
            .MergeCells = flexMergeFree
            .MergeCol(col_站点) = True
            .AllowUserResizing = flexResizeBoth
            .SelectionMode = flexSelectionByRow
            .Editable = flexEDKbdMouse
            .ExplorerBar = flexExSortShowAndMove
            .ColSort(col_选择) = flexSortNone
            .Cell(flexcpPicture, 0, col_选择) = ils16.ListImages("UnCheck").Picture
            .Cell(flexcpPictureAlignment, 0, col_选择) = flexAlignCenterCenter
            blnTmp = False
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID & "")
                .TextMatrix(.Rows - 1, col_站点) = rsTmp!站点 & ""
                If rsTmp!站点 & "" <> "" Then
                    blnTmp = True
                Else
                    .TextMatrix(.Rows - 1, col_站点) = " "
                End If
                .TextMatrix(.Rows - 1, col_编码) = rsTmp!编码 & ""
                .TextMatrix(.Rows - 1, col_名称) = rsTmp!名称 & ""
                .TextMatrix(.Rows - 1, col_简码) = rsTmp!简码 & ""
                
                rsTmp.MoveNext
            Loop
            .ColHidden(col_站点) = Not blnTmp
        End With
    Next
   
    '票据种类
    lvw(lvw_票据).ListItems.Add , "C1", "收费收据"
    lvw(lvw_票据).ListItems.Add , "C2", "预交收据"
    lvw(lvw_票据).ListItems.Add , "C3", "结帐收据"
    lvw(lvw_票据).ListItems.Add , "C4", "挂号收据"
    'lvw(lvw_票据).ListItems.Add , "C5", "就诊卡"
    
    With lvw(lvw_一卡通)
        .ColumnHeaders(1).Width = 549.9213
        .ColumnHeaders(2).Width = 1200.189
        .ColumnHeaders(3).Width = 975.1182
        .ColumnHeaders(4).Width = 950.7402
        .ColumnHeaders(5).Width = 2204.788
    End With
    
    '刷卡要求输入密码的场合
    With lst(lst_刷卡密码)
        .AddItem "门诊挂号"
        .AddItem "门诊划价"
        .AddItem "门诊收费"
        .AddItem "门诊记帐"
        .AddItem "入院登记"
        .AddItem "住院记帐"
        .AddItem "病人结帐"
        .AddItem "病人预交款"
        .AddItem "检验技师站"
        .AddItem "影像医技站"
        .ListIndex = 0
    End With
    
    msh(0).Cols = 7
    msh(0).TextMatrix(0, 0) = "病区"
    msh(0).TextMatrix(0, 1) = "床位费"
    msh(0).TextMatrix(0, 2) = " 启用日期"
    msh(0).TextMatrix(0, 3) = "护理费"
    msh(0).TextMatrix(0, 4) = " 启用日期"
    msh(0).TextMatrix(0, 5) = " 床位费原始启用日期"
    msh(0).TextMatrix(0, 6) = " 护理费原始启用日期"
    
    msh(0).ColWidth(0) = 1300
    msh(0).ColWidth(1) = 600
    msh(0).ColWidth(2) = 1000
    msh(0).ColWidth(3) = 600
    msh(0).ColWidth(4) = 1000
    msh(0).ColWidth(5) = 0
    msh(0).ColWidth(6) = 0
    msh(0).ColAlignmentFixed(0) = 1
    msh(0).ColAlignment(1) = 4
    msh(0).ColAlignment(2) = 1
    msh(0).ColAlignment(3) = 4
    msh(0).ColAlignment(4) = 1
    msh(0).Col = 0
    msh(0).Row = 0
    msh(0).ColSel = 2
    msh(0).RowSel = 0
    msh(0).FillStyle = flexFillRepeat
    msh(0).CellAlignment = 4
    msh(0).FillStyle = flexFillSingle
    msh(0).AllowBigSelection = False
    msh(0).Row = 1
    
    With Bill(bill_自动计算)
        .Cols = 5 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .TextMatrix(0, 0) = "病区"
        .TextMatrix(0, 1) = "收费细目ID"
        .TextMatrix(0, 2) = "收费项目"
        .TextMatrix(0, 3) = "计算方式"
        .TextMatrix(0, 4) = "启用日期"
        .ColWidth(0) = 1300
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColData(0) = 3
        .ColData(1) = 5
        .ColData(2) = 1
        .ColData(3) = 0
        .ColData(4) = 4
        .PrimaryCol = 0
        .Active = True
    End With
    
    With Bill(bill_药品流向)
        .Cols = 4 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "所在库房"
        .TextMatrix(0, 1) = "对方库房"
        .TextMatrix(0, 2) = "对方库房ID"
        .TextMatrix(0, 3) = "流向"
        .ColWidth(0) = 1700
        .ColWidth(1) = 1700
        .ColWidth(2) = 0
        .ColWidth(3) = 3600
        .ColData(0) = 3
        .ColData(1) = 3
        .ColData(2) = 5
        .ColData(3) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    
    With Bill(bill_药品领用流向)
        .Cols = 3 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "领用部门"
        .TextMatrix(0, 1) = "领用库房"
        .TextMatrix(0, 2) = "库房ID"
        .ColWidth(0) = 3500
        .ColWidth(1) = 3500
        .ColWidth(2) = 0
        .ColData(0) = 1
        .ColData(1) = 3
        .ColData(2) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    
    lngIndex = bill_记帐报警
    Bill(lngIndex).Cols = 8
    Bill(lngIndex).ColAlignment(0) = 1 '病区
    Bill(lngIndex).ColAlignment(1) = 1 '报警方法
    Bill(lngIndex).ColAlignment(2) = 7 '报警值
    Bill(lngIndex).ColAlignment(3) = 1 '报警标志1
    Bill(lngIndex).ColAlignment(4) = 1 '报警标志2
    Bill(lngIndex).ColAlignment(5) = 1 '报警标志3
    '刘兴洪 问题:34770    日期:2010-12-21 10:52:49
    Bill(lngIndex).ColAlignment(6) = 7 '催款下限
    Bill(lngIndex).ColAlignment(7) = 7 '催款标准
    
    Bill(lngIndex).TextMatrix(0, 0) = "病区"
    Bill(lngIndex).TextMatrix(0, 1) = "报警方法"
    Bill(lngIndex).TextMatrix(0, 2) = "报警值"
    Bill(lngIndex).TextMatrix(0, 3) = "报警方式1"
    Bill(lngIndex).TextMatrix(0, 4) = "报警方式2"
    Bill(lngIndex).TextMatrix(0, 5) = "报警方式3"
    Bill(lngIndex).TextMatrix(0, 6) = "催款下限"
    Bill(lngIndex).TextMatrix(0, 7) = "催款标准"
    
    Bill(lngIndex).ColWidth(0) = 1300
    Bill(lngIndex).ColWidth(1) = 1000
    Bill(lngIndex).ColWidth(2) = 800
    Bill(lngIndex).ColWidth(3) = 1500
    Bill(lngIndex).ColWidth(4) = 1500
    Bill(lngIndex).ColWidth(5) = 1500
    Bill(lngIndex).ColWidth(6) = 1000
    Bill(lngIndex).ColWidth(7) = 1000
    
    Bill(lngIndex).ColData(0) = 3
    Bill(lngIndex).ColData(1) = 0
    Bill(lngIndex).ColData(2) = 4
    Bill(lngIndex).ColData(3) = 1
    Bill(lngIndex).ColData(4) = 1
    Bill(lngIndex).ColData(5) = 1
    Bill(lngIndex).ColData(6) = 4
    Bill(lngIndex).ColData(7) = 4
    
    Bill(lngIndex).PrimaryCol = 0
    Bill(lngIndex).Active = True

    '库房单位
    msf库房单位.AllowUserResizing = flexResizeNone
    msf库房单位.FixedRows = 1
    msf库房单位.Cols = 5
    msf库房单位.MergeCol(0) = True
    msf库房单位.FormatString = "药品库房|服务对象|售价单位|门诊单位|住院单位|药库单位"
    msf库房单位.ColWidth(1) = 900
    msf库房单位.ColWidth(2) = 900
    msf库房单位.ColWidth(3) = 900
    msf库房单位.ColWidth(4) = 900
    msf库房单位.ColWidth(5) = 900
    msf库房单位.ColAlignment(1) = 4
    msf库房单位.ColAlignment(2) = 4
    msf库房单位.ColAlignment(3) = 4
    msf库房单位.ColAlignment(4) = 4
    msf库房单位.ColAlignment(5) = 4
    msf库房单位.ColWidth(0) = msf库房单位.Width - 900 * 6 - 27 * Screen.TwipsPerPixelX
    msf库房单位.MergeCells = flexMergeFree
    msf库房单位.MergeCol(0) = True
    
    
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
        .ColWidth(1) = 1500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1500
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
    
    '读取医嘱发送为划价类别
    gstrSQL = "Select 编码,名称 From 诊疗项目类别 Where 编码 Not IN('5','6','7','8','9')" & _
        " Union All Select '5','药品' From Dual Order by 编码"
    Call OpenRecordset(rsTmp, Me.Caption)
  
    Do While Not rsTmp.EOF
        lst(lst_门诊发送类别).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_门诊发送类别).ItemData(lst(lst_门诊发送类别).NewIndex) = Asc(rsTmp!编码)
        
        lst(lst_住院发送类别).AddItem rsTmp!编码 & "-" & rsTmp!名称
        lst(lst_住院发送类别).ItemData(lst(lst_住院发送类别).NewIndex) = Asc(rsTmp!编码)
        
        rsTmp.MoveNext
    Loop
    
    '读取医嘱内容定义
    gstrSQL = "Select 诊疗类别,医嘱内容 From 医嘱内容定义 Order by 诊疗类别"
    Call OpenRecordset(mrsAdvice, Me.Caption)
    
    '读取具有“配制中心”和“药房”属性的部门
    gstrSQL = "Select Distinct A.ID, A.名称" & _
        " From 部门表 A, 部门性质说明 B " & _
        " Where A.ID = B.部门id And B.工作性质 = '配制中心' And " & _
        " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') " & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) "
    Call OpenRecordset(rsTmp, Me.Caption)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Load社区接口() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    strSQL = "Select 序号, 名称, 说明, 启用, 部件名 From 社区目录 Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw社区.ListItems.Add(, "_" & rsTmp!序号, rsTmp!序号, , "社区")
        ObjItem.SubItems(1) = rsTmp!名称
        ObjItem.SubItems(2) = Nvl(rsTmp!说明)
        ObjItem.SubItems(3) = rsTmp!部件名
        ObjItem.SubItems(4) = IIF(Nvl(rsTmp!启用, 0) = 1, "√", "")
        rsTmp.MoveNext
    Loop
    
    If Not lvw社区.SelectedItem Is Nothing Then
        Call lvw社区_ItemClick(lvw社区.SelectedItem)
    End If
    Load社区接口 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function LoadOneCard() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    lvw(lvw_一卡通).ListItems.Clear
    
    strSQL = "Select 编号,名称,结算方式,医院编码,启用 From 一卡通目录 Order by 编号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw(lvw_一卡通).ListItems.Add(, "_" & rsTmp!编号, rsTmp!编号)
        ObjItem.SubItems(1) = rsTmp!名称
        ObjItem.SubItems(2) = rsTmp!结算方式
        ObjItem.SubItems(3) = rsTmp!医院编码
        ObjItem.SubItems(4) = IIF(Nvl(rsTmp!启用, 0) = 2, "启用:标准一卡通", IIF(Nvl(rsTmp!启用, 0) = 1, "启用:仅涉及扣卡", "停用"))
        rsTmp.MoveNext
    Loop
    
    If Not lvw(lvw_一卡通).SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw_一卡通, lvw(lvw_一卡通).SelectedItem)
    End If
    LoadOneCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load部门()
'提取数据并显示出来
    Dim lng序号 As Long, str库房ID As String
    Dim rsTemp As New ADODB.Recordset
    Dim strType As String
    Dim strSequence As String
    
'    StrType = "('中药库','西药库','成药库','制剂室', '中药房', '西药房', '成药房','卫材库','发料部门')"

    '药品科室
    On Error GoTo ErrHandle
    strType = "('中药库','西药库','成药库','制剂室', '中药房', '西药房', '成药房')"
    strSequence = "(21,22,23,24,25,26,27,28,29,32,62)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.编码,a.名称,b.编号 " & _
        "   From 部门表 A,科室号码表 b" & _
        "   Where a.id=b.科室id and a.ID in (select distinct 部门id from 部门性质说明 where 工作性质 in " & strType & ")" & _
        "   And b.项目序号 In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.编码,a.名称,'' As 编号 " & _
        "   From 部门表 A " & _
        "   Where a.ID in (select distinct 部门id from 部门性质说明 " & _
        "   where 工作性质 in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct 科室id From 科室号码表 Where 科室id Is Not null " & _
        "   And 项目序号 In " & strSequence & ") " & _
        "   ORDER BY 编码 "
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关的科室"
    
    With rsTemp
        str库房ID = ""
        Do While Not .EOF
            'mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.选择) = IIf(Nvl(!选择, 0) = 1, "√", "")
            mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.科室) = Nvl(!名称)
            mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.号码) = Nvl(!编号)
            mshBillEdit.RowData(mshBillEdit.Rows - 1) = !ID
            mshBillEdit.Rows = mshBillEdit.Rows + 1
            str库房ID = str库房ID & "," & rsTemp!ID
            .MoveNext
        Loop
    End With
    
    If str库房ID <> "" Then
        str库房ID = Mid(str库房ID, 2)
        mshBillEdit.Rows = mshBillEdit.Rows - 1
        mshBillEdit.Active = True
    Else
        mshBillEdit.Active = False
    End If
    
    rsTemp.Close
    
    '卫材科室
    strType = "('制剂室','卫材库','虚拟库房')"
    strSequence = "(68,69,70,71,72,73,74,75,76,77)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.编码,a.名称,b.编号 " & _
        "   From 部门表 A,科室号码表 b" & _
        "   Where a.id=b.科室id and a.ID in (select distinct 部门id from 部门性质说明 where 工作性质 in " & strType & ")" & _
        "   And b.项目序号 In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.编码,a.名称,'' As 编号 " & _
        "   From 部门表 A " & _
        "   Where a.ID in (select distinct 部门id from 部门性质说明 " & _
        "   where 工作性质 in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct 科室id From 科室号码表 Where 科室id Is Not null " & _
        "   And 项目序号 In " & strSequence & ") " & _
        "   ORDER BY 编码 "

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取相关的科室"
    
    With rsTemp
        str库房ID = ""
        Do While Not .EOF
            'mshBillEditstuff.TextMatrix(mshBillEditstuff.Rows - 1, mGrdCol.选择) = IIf(Nvl(!选择, 0) = 1, "√", "")
            mshBillEditStuff.TextMatrix(mshBillEditStuff.Rows - 1, mGrdCol.科室) = Nvl(!名称)
            mshBillEditStuff.TextMatrix(mshBillEditStuff.Rows - 1, mGrdCol.号码) = Nvl(!编号)
            mshBillEditStuff.RowData(mshBillEditStuff.Rows - 1) = !ID
            mshBillEditStuff.Rows = mshBillEditStuff.Rows + 1
            str库房ID = str库房ID & "," & rsTemp!ID
            .MoveNext
        Loop
    End With
    
    If str库房ID <> "" Then
        str库房ID = Mid(str库房ID, 2)
        mshBillEditStuff.Rows = mshBillEditStuff.Rows - 1
        mshBillEditStuff.Active = True
    Else
        mshBillEditStuff.Active = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPara()
'系统参数表
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer, blnFind As Boolean
    Dim n As Integer

    '首先对费用类型进行初始化
    On Error GoTo ErrHandle
    Call Load费用类型

    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select 参数号,参数值,缺省值 From Zlparameters Where 系统 = " & glngSys & " And Nvl(私有, 0) = 0 And 模块 Is Null Order By 参数号"
    Call OpenRecordset(rsTemp, Me.Caption)

    Do Until rsTemp.EOF
        Select Case rsTemp("参数号")
        Case 1    '上午上下班时间
            i = InStr(UCase(rsTemp("参数值")), "AND")
            strTemp = Mid(rsTemp("参数值"), 1, i - 2)
            dtp(dtp_上午上班).Value = CDate(strTemp)
            strTemp = Mid(rsTemp("参数值"), i + 4)
            dtp(dtp_上午下班).Value = CDate(strTemp)
        Case 2    '下午上下班时间
            i = InStr(UCase(rsTemp("参数值")), "AND")
            strTemp = Mid(rsTemp("参数值"), 1, i - 2)
            dtp(dtp_下午上班).Value = CDate(strTemp)
            strTemp = Mid(rsTemp("参数值"), i + 4)
            dtp(dtp_下午下班).Value = CDate(strTemp)
            '            Case 3 '收据加收工本费  '56963
            '                chk(chk_加收工本费).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            '                Call chk_Click(chk_加收工本费)
            '            Case 4 '收费收据总行次
            '                If Not IsNull(rsTemp("参数值")) Then
            '                    ud(ud_收费收据).Value = rsTemp("参数值")
            '                End If
        Case 5    '补录医嘱识别间隔
            ud(ud_补录医嘱识别间隔).Value = Nvl(rsTemp!参数值, 30)
        Case 6    '未审核记帐处方发药
            chk(chk_未审核记帐处方发药) = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 148    '未收费处方发药
            chk(chk_未收费处方发药) = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 7    '修正上期自动计费
            chk(chk_自动修正).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 9    '费用金额保留位数
            Me.ud(ud_费用金额保留位数).Value = IIF(IsNumeric(zlCommFun.Nvl(rsTemp("参数值"), 2)), zlCommFun.Nvl(rsTemp("参数值"), 2), 2)
            Me.txtUD(ud_费用金额保留位数).Text = Me.ud(ud_费用金额保留位数).Value
            mDecimal = Me.txtUD(ud_费用金额保留位数).Text
        Case 157    '费用单价保留位数
            Me.ud(ud_费用单价保留位数).Value = IIF(IsNumeric(zlCommFun.Nvl(rsTemp("参数值"), 5)), zlCommFun.Nvl(rsTemp("参数值"), 5), 5)
            Me.txtUD(ud_费用单价保留位数).Text = Me.ud(ud_费用单价保留位数).Value
            pDecimal = Me.txtUD(ud_费用单价保留位数).Text

        Case 10    '入院时收预交款
            chk(chk_收取预交款).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 11    '入院时办就诊卡
            chk(chk_时办理就诊卡).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
            '            Case 12 '就诊卡号密文显示
            '                chk(chk_密文显示).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 13    '入院同时入科
            chk(chk_分配床位号).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 14    '零钱处理
            strTemp = IIF(IsNull(rsTemp("参数值")), "000", rsTemp("参数值"))
            n = Val(Mid(strTemp, 1, 1))
            For i = 0 To cmb(cmb_挂号零钱处理).ListCount
                If Val(Split(cmb(cmb_挂号零钱处理).List(i) & "-", "-")(0)) = n Then cmb(cmb_挂号零钱处理).ListIndex = i: Exit For
            Next
            cmb(cmb_收费零钱处理).ListIndex = Val(Mid(strTemp, 2, 1))
            cmb(cmb_结帐零钱处理).ListIndex = Val(Mid(strTemp, 3, 1))
        Case 15    '门诊收费与发药分离
            chk(chk_门诊收费与发药分离).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 0, 0, 1)
        Case 16    '住院记帐与发药分离
            chk(chk_住院记帐与发药分离).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 0, 0, 1)
        Case 17    '病人输入方式，分别为姓名、就诊卡、挂号单、病人ID
            strTemp = IIF(IsNull(rsTemp("参数值")), "1111", rsTemp("参数值"))
            chk(chk_病人姓名).Value = Val(Mid(strTemp, 1, 1))
            chk(chk_刷就诊卡).Value = Val(Mid(strTemp, 2, 1))
            chk(chk_挂号单号).Value = Val(Mid(strTemp, 3, 1))
            chk(chk_病人ID).Value = Val(Mid(strTemp, 4, 1))
        Case 18    '指定药房时限制库存
            chk(chk_限定药品的库存).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 19    '窗口分配方式
            '该组第一个控件的Index值是2
            opt(CInt(IIF(IsNull(rsTemp("参数值")), "0", rsTemp("参数值"))) + 2).Value = True
        Case 20    '表示各种票据的号码长度，各位分别为1-收费,2-预交,3-结帐,4-挂号
            strTemp = IIF(IsNull(rsTemp("参数值")), "7|7|7|7", rsTemp("参数值"))
            lvw(lvw_票据).ListItems("C1").SubItems(1) = Split(strTemp, "|")(0)
            lvw(lvw_票据).ListItems("C2").SubItems(1) = Split(strTemp, "|")(1)
            lvw(lvw_票据).ListItems("C3").SubItems(1) = Split(strTemp, "|")(2)
            lvw(lvw_票据).ListItems("C4").SubItems(1) = Split(strTemp, "|")(3)
            'lvw(lvw_票据).ListItems("C5").SubItems(1) = Split(strTemp, "|")(4)
        Case 21  '挂号有效天数
            '普通号
            ud(ud_挂号单).Value = IIF(Left(zlCommFun.Nvl(rsTemp("参数值"), 0), 1) = 0, 1, Left(zlCommFun.Nvl(rsTemp("参数值"), 0), 1))
            '急诊号
            ud(ud_急诊挂号单).Value = IIF(Mid(zlCommFun.Nvl(rsTemp("参数值"), 0), 2, 1) = 0, 1, Mid(zlCommFun.Nvl(rsTemp("参数值"), 0), 2, 1))
        Case 22    '出院时未执行项目检查
            cmb(cmb_出院时未执行项目检查).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 23    '已结帐单据操作
            cmb(cmb_已结单据).ListIndex = IIF(IsNull(rsTemp("参数值")), 0, rsTemp("参数值"))
        Case 24    '表示是否严格控制管理对票据的使用，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
            strTemp = IIF(IsNull(rsTemp("参数值")), "1111", rsTemp("参数值"))
            lvw(lvw_票据).ListItems("C1").SubItems(2) = IIF(Mid(strTemp, 1, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C2").SubItems(2) = IIF(Mid(strTemp, 2, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C3").SubItems(2) = IIF(Mid(strTemp, 3, 1) = "1", "√", "")
            lvw(lvw_票据).ListItems("C4").SubItems(2) = IIF(Mid(strTemp, 4, 1) = "1", "√", "")
            ' lvw(lvw_票据).ListItems("C5").SubItems(2) = IIF(Mid(strTemp, 5, 1) = "1", "√", "")
        Case 25    '电子签名认证中心
            With cmb(cmb_电子签名认证中心)
                blnFind = False
                For i = 0 To .ListCount - 1
                    If Val(.List(i)) = Val("" & rsTemp!参数值) Then
                        .ListIndex = i
                        blnFind = True
                        Exit For
                    End If
                Next
                If .ListCount > 0 And Not blnFind Then .ListIndex = 0
            End With

        Case 185    '病人审核方式   ' 49501
            With cboPatiVerfy
                blnFind = False
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = Val("" & rsTemp!参数值) Then
                        .ListIndex = i
                        blnFind = True
                        Exit For
                    End If
                Next
                If .ListCount > 0 And Not blnFind Then .ListIndex = 0
            End With
        Case 26    '电子签名使用场合
            chk(chk_电子签名控制_门诊).Value = Val(Mid(Nvl(rsTemp!参数值), 1, 1))
            chk(chk_电子签名控制_住院).Value = Val(Mid(Nvl(rsTemp!参数值), 2, 1))
            chk(chk_电子签名控制_医技).Value = Val(Mid(Nvl(rsTemp!参数值), 3, 1))
            chk(chk_电子签名控制_护理).Value = Val(Mid(Nvl(rsTemp!参数值), 4, 1))
            chk(chk_电子签名控制_药品).Value = Val(Mid(Nvl(rsTemp!参数值), 5, 1))
            chk(chk_电子签名控制_lis).Value = Val(Mid(Nvl(rsTemp!参数值), 6, 1))
            chk(chk_电子签名控制_pacs).Value = Val(Mid(Nvl(rsTemp!参数值), 7, 1))
        Case 27    '住院药嘱发送产生领药号
            chk(chk_住院药嘱发送产生领药号).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 0, 0, 1)
        Case 28    '门诊病人消费时需要刷卡验证
            chk(chk_门诊病人消费时需要刷卡验证).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 163    '项目执行前必须先收费或先记帐审核
            chk(chk_项目执行前必须收费或审核).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
        Case 232    '项目开单后立即收费或记帐审核
            chk(chk_项目开单后立即收费或记帐审核).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 29    '指导批发价定价单位
            cmb(cmb_定价单位).ListIndex = IIF(rsTemp("参数值") = "1", 1, 0)
        Case 31    '在院病人不准出院结帐
            chk(chk_在院病人不准出院结帐).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 32    '转科时未执行项目检查
            cmb(cmb_转科时未执行项目检查).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 33    '执行之后卫材自动发料
            chk(chk_执行之后自动发料).Value = IIF(Val(rsTemp!参数值) <> 0, 1, 0)
        Case 34    '指定医嘱在其他科室执行
            chk(chk_指定医嘱在其他科室执行).Value = IIF(Val(rsTemp!参数值) <> 0, 1, 0)
        Case 41    '医保病人适用费用类型
            SetListByText lst(lst_医保病人), Replace(IIF(IsNull(rsTemp("参数值")), "", rsTemp("参数值")), "|", ",")
        Case 42    '公费病人适用费用类型
            SetListByText lst(lst_公费病人), Replace(IIF(IsNull(rsTemp("参数值")), "", rsTemp("参数值")), "|", ",")
        Case 43    '下达出院医嘱才允许出院
            chk(chk_下达出院医嘱才允许出院).Value = IIF(Val("" & rsTemp!参数值) <> 0, 1, 0)
        Case 44    '收费项目和诊疗项目的输入匹配方式
            chk(chk_全数字只查编码).Value = IIF(Mid(IIF(IsNull(rsTemp!参数值), "00", rsTemp!参数值), 1, 1) = "1", 1, 0)
            chk(chk_全字母只查简码).Value = IIF(Mid(IIF(IsNull(rsTemp!参数值), "00", rsTemp!参数值), 2, 1) = "1", 1, 0)
        Case 45    '收费同时发药
            chk(chk_收费同时发药).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 46    '刷卡要求输入密码
            With lst(lst_刷卡密码)
                For i = 1 To Len(Nvl(rsTemp!参数值))
                    If Mid(rsTemp!参数值, i, 1) = "1" And i - 1 <= .ListCount - 1 Then
                        .Selected(i - 1) = True
                    End If
                Next
            End With
        Case 51    '本人执行登记
            chk(chk_本人执行登记).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 52    '必须输入开单人
            chk(chk_输入开单人).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 53    '输入它科开单人
            chk(chk_它科开单人).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 54    '时价药品以加价率入库
            chk(chk_时价药品入库).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 55    '诊断输入来源
            cmb(cmb_诊断输入来源).ListIndex = CLng(zlCommFun.Nvl(rsTemp("参数值"), 1)) - 1
        Case 56    '门诊处方条数限制
            If zlCommFun.Nvl(rsTemp("参数值"), 0) = 0 Then
                ud(ud_门诊处方条数限制).Value = 5
                chk(chk_门诊处方条数限制).Value = 0
            Else
                ud(ud_门诊处方条数限制).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
                chk(chk_门诊处方条数限制).Value = 1
            End If
            '            Case 57 '收费每次只用一张票据   '56963
            '                chk(chk_收费每次只用一张票据).Value = IIF(rsTemp("参数值") <> 0, 1, 0)
        Case 58    '未审单据结帐处理
            cmb(cmb_未审单据结帐).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 59    '医保对码检查
            cmb(cmb_医保对码检查).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 60    '单笔费用最大提醒金额
            txtMaxMoney.Text = zlCommFun.Nvl(rsTemp("参数值"))
            Call txtMaxMoney_Validate(False)
        Case 61    '诊疗编码递增模式
            cmb(cmb_诊疗编码模式).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 63    '住院卫材自动发料
            chk(chk_住院卫材自动发料).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 64    '药品单据审核规则
            cmb(cmb_药品单据审核).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 65    '诊断输入方式
            cmb(cmb_门诊诊断输入).ListIndex = Val(Mid(Nvl(rsTemp!参数值, "11"), 1, 1)) - 1
            cmb(cmb_住院诊断输入).ListIndex = Val(Mid(Nvl(rsTemp!参数值, "11"), 2, 1)) - 1
        Case 66    '挂号预约天数
            ud(ud_挂号预约天数).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 68    '未作废临嘱禁止退药
            chk(chk_未作废临嘱禁止退药).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 69    '药品按规格下医嘱
            chk(chk_药品按规格下医嘱).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 70    '过敏登记有效天数
            If zlCommFun.Nvl(rsTemp("参数值"), 0) = 0 Then
                ud(ud_过敏登记有效天数).Value = 1
                chk(chk_过敏登记有效天数).Value = 0
            Else
                ud(ud_过敏登记有效天数).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
                chk(chk_过敏登记有效天数).Value = 1
            End If
        Case 71    '长期医嘱次日生效
            chk(chk_长期医嘱次日生效).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 72    '首先输入收费类别
            chk(chk_首先输入收费类别).Value = zlCommFun.Nvl(rsTemp("参数值"), 1)
        Case 73    '明确申领药品批次
            chk(chk_明确申领药品批次).Value = zlCommFun.Nvl(rsTemp("参数值"), 1)
            chk(chk_明确申领药品批次).Tag = chk(chk_明确申领药品批次).Value
        Case 75    '外购入库需要核查
            chk(chk_外购入库需要核查).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 76    'chk_时价药品直接确定售价
            chk(chk_时价药品直接确定售价).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
            '            Case 78 '多张单据收费分别打印 56963
            '                chk(chk_多张单据收费分别打印).Value = IIF(Val("" & rsTemp("参数值")) = 1, 1, 0)
        Case 80    '住院医嘱发送为划价单
            strTemp = zlCommFun.Nvl(rsTemp("参数值"))
            If strTemp <> "" Then
                With lst(lst_住院发送类别)
                    For i = 0 To .ListCount - 1
                        If InStr(strTemp, Chr(.ItemData(i))) > 0 Then
                            .Selected(i) = True
                        End If
                    Next
                    .ListIndex = 0
                End With
            End If
        Case 81    '执行后自动审核划价单
            chk(chk_执行后自动审核划价单).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 84    '一次申请多个检验项目
            chk(chk_一次申请多个检验项目).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 86    '门诊医嘱发送为划价单
            strTemp = zlCommFun.Nvl(rsTemp("参数值"))
            If strTemp <> "" Then
                With lst(lst_门诊发送类别)
                    For i = 0 To .ListCount - 1
                        If InStr(strTemp, Chr(.ItemData(i))) > 0 Then
                            .Selected(i) = True
                        End If
                    Next
                    .ListIndex = 0
                End With
            End If
        Case 87    '药品编码递增模式
            cmb(cmb_药品编码模式).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
            '            Case 89 '56963
            '                chk(chk_误差项不使用票据).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 92    '门诊卫材自动发料
            chk(chk_门诊卫材自动发料).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 30    '在医护工作站，药房等模块使用合理用药类型
            cmb(cmb_合理用药接口).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 93    '是否允许从属项目汇总计算折扣
            chk(chk_从属项目汇总计算折扣).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 96
            chk(chk_药品填单时下可用库存).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
            chk(chk_药品填单时下可用库存).Tag = chk(chk_药品填单时下可用库存).Value
            '            Case 97 '收费票据生成方式 '56963
            '                opt收费票据生成方式(Val("" & rsTemp("参数值")) Mod 10).Value = True
            '                chk(chk_按执行科室分别打印).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) >= 10, 1, 0)
        Case 98
            chk(chk_记帐报警包含划价费用).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 99
            chk(chk_入科确定护理等级).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 100
            chk(chk_下午算半天模式).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 126
            chk(chk_时价入库按折扣前采购价加成销售).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
            chk(chk_时价入库按折扣前采购价加成销售).Tag = chk(chk_时价入库按折扣前采购价加成销售).Value
        Case 143
            chk(chk_检验医嘱发送生成条形码).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 144
            chk(chk_收费项目首位当类别简码).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 1, 1, 0)
        Case 145
            chk(chk_每次住院使用新住院号).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 1, 1, 0)
        Case 147
            txtUD(ud_儿童年龄界定上限).Text = zlCommFun.Nvl(rsTemp("参数值"), 12)
        Case 149    '药品效期显示方式
            cmb(cmb_效期显示方式).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 150    '药品出库优先算法
            cmb(cmb_药品出库优先算法).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 151    '门诊退费须先申请
            chk(chk_门诊退费须先申请).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 1, 1, 0)
            '            Case 152 '就诊卡重复使用，刘兴洪：24357
            '                chk(chk_就诊卡重复使用).Value = IIF(Val(zlCommFun.Nvl(rsTemp("参数值"))) = 1, 1, 0)
        Case 154    '出院时检查未发药项目
            cmb(cmb_出院时未发药项目检查).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 155    '转科时检查未发药项目
            cmb(cmb_转科时未发药项目检查).ListIndex = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 158    '补充录入时限
            If IsNull(rsTemp!参数值) Then
                txtInputHours.Text = zlCommFun.Nvl(rsTemp("缺省值"), 0)
            Else
                txtInputHours.Text = rsTemp!参数值
            End If
        Case 160    '护理费计算标准:34741
            If zlCommFun.Nvl(rsTemp("参数值"), 0) = 1 Then
                opt护理(1).Value = True
            Else
                opt护理(0).Value = True
            End If

        Case 161    '是否允许使用禁忌药嘱
            chk(chk_禁忌药嘱).Value = Val("" & rsTemp("参数值"))

        Case 162    '下达医嘱时显示产地
            chk(chk_下达医嘱时显示产地).Value = Val("" & rsTemp("参数值"))

        Case 171
            chk(chk_允许未收费的门诊划价处方发料).Value = Val("" & rsTemp("参数值"))
        Case 172
            chk(chk_允许未审核的记账处方发料).Value = Val("" & rsTemp("参数值"))
        Case 173    '外购入库需要经过标记付款后才能进行付款管理
            chk(chk_外购入库需要经过标记付款后才能进行付款).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 174    '药品移库时明确药品批次
            chk(chk_药品移库明确批次).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 175    '药品领用时明确药品批次
            chk(chk_药品领用明确批次).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 181    '药品分段加成入库
            chk(chk_时价分段加成入库).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 1, 1, 0)
        Case 182    '禁止下达超极量药品医嘱
            chk(chk_禁止下达超极量药品医嘱).Value = IIF(zlCommFun.Nvl(rsTemp("参数值"), 0) = 1, 0, 1)
        Case 183    '时价药品入库按取上次售价
            chk(chk_时价药品取上次售价).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 186  '输血和皮试医嘱执行后需要核对
            chk(chk_输血和皮试医嘱执行后需要核对).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 187
            chk(chk抗菌药物分级管理).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 188
            chk(chk抗菌药物使用自备药).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
            chk(chk抗菌药物使用自备药).Enabled = chk(chk抗菌药物分级管理).Value = 1
        Case 189
            chk(chk允许下达院外执行的禁忌药品医嘱).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 191
            chk(chk只允许补录临嘱).Value = zlCommFun.Nvl(rsTemp("参数值"), 1)
        Case 192
            chk(chk_回退出院医嘱才允许撤销出院).Value = IIF(Val("" & rsTemp!参数值) <> 0, 1, 0)
        Case 208
            chk(chk临床工作站必须使用zlPlugIn部件).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 209
            chk(chk启用手术分级管理).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 210
            chk(chk_允许处理超过挂号有效天数的病人).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 213
            cmb(cmd_中药配方).ListIndex = IIF(Val("" & rsTemp!参数值) = 4, 1, 0)
        Case 214
            chk(chk_首次医嘱执行需要审核).Value = zlCommFun.Nvl(rsTemp("参数值"), 0)
        Case 215    '51612
            chk(chk_未入科禁止记账).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        Case 216
            chk(chk_输血分级管理).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        Case 217
            chk(chk_手术授权管理).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        Case 218
            chk(chk_输血申请三级审核).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        Case 219
            chk(chk_输血申请只能由中级及以上医师提出).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        Case 220    '允许修改n天内登记的医嘱执行记录
            txtUNExecLimit.Text = zlCommFun.Nvl(rsTemp("参数值"), 999)
            If txtUNExecLimit.Text = "999" Then
                chk(chk_医嘱执行有效天数).Value = 0
                txtUNExecLimit.Enabled = False
            Else
                chk(chk_医嘱执行有效天数).Value = 1
            End If
        Case 221
            If Val(zlCommFun.Nvl(rsTemp("参数值"), 0)) = 0 Then
                optAccountTime(1).Value = True
                txtAccountTime.Enabled = False
            Else
                optAccountTime(0).Value = True
                txtAccountTime.Enabled = True
                txtAccountTime.Text = Val(zlCommFun.Nvl(rsTemp("参数值"), 0))
            End If
        Case 223  '门诊新开医嘱间隔
            ud(ud_门诊新开医嘱间隔).Value = Nvl(rsTemp!参数值, 1)
        Case 224    '过敏输入来源
            '太元通合理用药接口，因为已经按参数号排序，因此可以使用控件的值
            If cmb(cmb_合理用药接口).ListIndex = 3 Then
                cmb(cmd_过敏输入来源).ListIndex = Val(zlCommFun.Nvl(rsTemp("参数值"), 0))
            End If
        Case 225  '启用大通接口日志调用65522
            chk(chk_启用接口调用日志).Value = Val(zlCommFun.Nvl(rsTemp("参数值"), 0))
        Case 226 '美康接口参数
            chk(chk_允许使用系统设置).Value = Val(zlCommFun.Nvl(rsTemp("参数值"), 1))
        Case 227 '转科时检查未审核销帐单据
            cmb(cmd_转科时未审核销帐单据).ListIndex = Val(zlCommFun.Nvl(rsTemp("参数值"), 0))
        Case 228
            strTemp = zlCommFun.Nvl(rsTemp("参数值"), "3.0")
            If strTemp = "3.0" Then
                optPASSVer(0).Value = True
            Else
                optPASSVer(1).Value = True
            End If
        Case 230
            chk(chk_医嘱超量时必须输入原因).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
            Call Set不写超量科室(chk(chk_医嘱超量时必须输入原因).Value = 1)
        Case 233
            strTemp = zlCommFun.Nvl(rsTemp("参数值"))
            Call Init不填超量说明(strTemp)
        Case 234
            strTemp = zlCommFun.Nvl(rsTemp("参数值"))
            Call Init转科出院不检查项目(strTemp)
        Case 235
            cmb(cmd_出院时超期护理数据).ListIndex = Val(zlCommFun.Nvl(rsTemp("参数值"), 0))
        Case 239
            chk(chk_新开医嘱签名时一组医嘱签名一次).Value = IIF(Val(zlCommFun.Nvl(rsTemp!参数值)) = 1, 1, 0)
        End Select
        rsTemp.MoveNext
    Loop

    '显示当前票据的情况
    lvw(lvw_票据).ListItems("C1").Selected = True
    lvw_ItemClick lvw_票据, lvw(lvw_票据).SelectedItem

    '电子签名控制
    Call cmb_Click(cmb_电子签名认证中心)
    Call LoadSign
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load单据操作()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str单据 As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.人员ID,B.姓名,A.单据,A.时间限制,A.他人单据,A.金额上限 from 单据操作控制 A,人员表 B where A.人员ID=B.ID"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    lvw(lvw_单据).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_单据).ListItems.Add(, , rsTemp("姓名"), , "Limit")
        
        str单据 = Switch(rsTemp("单据") = 1, "挂号单据", rsTemp("单据") = 2, "收费单", rsTemp("单据") = 3, "划价单", rsTemp("单据") = 4, "门诊记帐", _
                       rsTemp("单据") = 5, "住院记帐", rsTemp("单据") = 6, "预交款", rsTemp("单据") = 7, "结帐单据", rsTemp("单据") = 8, "就诊卡", rsTemp("单据") = 9, "处方")
        lst.SubItems(1) = str单据
        lst.SubItems(2) = rsTemp("时间限制")
        lst.SubItems(3) = IIF(rsTemp("他人单据") = 1, "是", "否")
        lst.SubItems(4) = IIF(IsNull(rsTemp("金额上限")), "", Format(rsTemp("金额上限"), "0.00"))
        lst.Tag = rsTemp("人员ID")
        lst.ListSubItems(1).Tag = rsTemp("单据")
        
        rsTemp.MoveNext
    Loop
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load病区()
    Dim rs病区 As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    rs病区.CursorLocation = adUseClient
    gstrSQL = "select A.ID,A.名称,A.编码 " & _
               " from  部门性质说明 b,部门表 a " & _
               " where B.服务对象 in(1,2,3) And B.工作性质='护理' and  b.部门ID=a.ID and " & _
               Where撤档时间("A") & " order by 编码"
    Call OpenRecordset(rs病区, Me.Caption)
    
    Bill(bill_自动计算).Clear
    Bill(bill_记帐报警).Clear
    
    If rs病区.RecordCount > 0 Then
        msh(0).Rows = rs病区.RecordCount + 1
        lngRow = 1
        Do Until rs病区.EOF
            Bill(bill_自动计算).AddItem rs病区("编码") & "-" & rs病区("名称")
            Bill(bill_自动计算).ItemData(Bill(bill_自动计算).NewIndex) = rs病区("ID")
            Bill(bill_记帐报警).AddItem rs病区("编码") & "-" & rs病区("名称")
            Bill(bill_记帐报警).ItemData(Bill(bill_自动计算).NewIndex) = rs病区("ID")
            msh(0).TextMatrix(lngRow, 0) = rs病区("编码") & "-" & rs病区("名称")
            msh(0).RowData(lngRow) = rs病区("ID")
            lngRow = lngRow + 1
            rs病区.MoveNext
        Loop
        Bill(bill_自动计算).ListIndex = 0
    End If
    Bill(bill_记帐报警).AddItem "*门诊*"
    Bill(bill_记帐报警).ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadSign()
'功能：加载电子签名启用部门
    Dim rsTmp As New Recordset
    Dim i As Long, lngTmp As Long
    
    gstrSQL = "select 部门ID,场合 from 电子签名启用部门"
    On Error GoTo ErrHandle
    Call OpenRecordset(rsTmp, Me.Caption)
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            rsTmp.Filter = "场合=" & i
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                lngTmp = .FindRow(Val(rsTmp!部门id & ""))
                If lngTmp <> -1 Then
                    .Cell(flexcpChecked, lngTmp, col_选择) = 1
                End If
                rsTmp.MoveNext
            Loop
            
        End With
    Next
    For i = 0 To sstSign.Tabs - 1
        If sstSign.TabVisible(i) = True Then sstSign.Tab = i: Exit For
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadTable()
'完成其余的初始化工作
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng单位 As Long
    Dim strTemp As String, lngTemp As Long, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '收费特定项目
    On Error GoTo ErrHandle
    gstrSQL = "select a.特定项目 ,c.ID,c.名称  " & _
            " from 收费特定项目 a,收费细目 c " & _
            " where a.收费细目ID =c.id"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("特定项目")
            Case "病历费"
                txtCmd(0).Tag = rsTemp("ID")
                txtCmd(0).Text = rsTemp("名称")
            Case "工本费"
                txtCmd(1).Tag = rsTemp("ID")
                txtCmd(1).Text = rsTemp("名称")
            Case "普通配置费"
                txtCmd(3).Tag = rsTemp("ID")
                txtCmd(3).Text = rsTemp("名称")
            Case "肿瘤配置费"
                txtCmd(4).Tag = rsTemp("ID")
                txtCmd(4).Text = rsTemp("名称")
        End Select
        rsTemp.MoveNext
    Loop
    
    '病区自动计帐程序
    gstrSQL = "select A.病区ID,B.编码,b.名称 as 病区 ,a.收费细目ID,c.名称 as 收费细目 ,a.计算标志,a.启用日期 " & _
            " from 自动计价项目 A,部门表 B,收费细目 C " & _
            " where A.病区ID= B.id and A.收费细目ID =C.id(+) " & _
            " order by b.编码 "
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill(bill_自动计算)
        lngRow = 1
        Do Until rsTemp.EOF
            If IsNull(rsTemp("收费细目ID")) Then
                '床位费或护理费
                For lngTemp = 1 To msh(0).Rows - 1
                    If msh(0).RowData(lngTemp) = rsTemp("病区ID") Then
                        If rsTemp("计算标志") = 1 Then
                            '床位费
                            msh(0).TextMatrix(lngTemp, 1) = "√"
                            msh(0).TextMatrix(lngTemp, 2) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                            msh(0).TextMatrix(lngTemp, 5) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                        Else
                            '护理费
                            msh(0).TextMatrix(lngTemp, 3) = "√"
                            msh(0).TextMatrix(lngTemp, 4) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                            msh(0).TextMatrix(lngTemp, 6) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                        End If
                    End If
                Next
            Else
                '其它费用
                .Rows = lngRow + 1
                .RowData(lngRow) = rsTemp("病区ID")
                .TextMatrix(lngRow, 0) = rsTemp("编码") & "-" & rsTemp("病区")
                .TextMatrix(lngRow, 1) = rsTemp("收费细目ID")
                .TextMatrix(lngRow, 2) = rsTemp("收费细目")
                .TextMatrix(lngRow, 3) = Switch(rsTemp("计算标志") = 6, "1-按床日", rsTemp("计算标志") = 8, "2-计算一次", True, "0-按收治日")
                .TextMatrix(lngRow, 4) = Format(IIF(IsNull(rsTemp!启用日期), "", rsTemp!启用日期), "yyyy-mm-dd")
                lngRow = lngRow + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '记帐报警类别
    gstrSQL = "Select 编码,类别 From 收费类别 Order by 编码"
    Set mrs类别 = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs类别, gstrSQL, Me.Caption)
    
    lst类别.Clear
    lst类别.AddItem "所有类别"
    Do While Not mrs类别.EOF
        lst类别.AddItem mrs类别!类别
        lst类别.ItemData(lst类别.NewIndex) = Asc(mrs类别!编码)
        mrs类别.MoveNext
    Loop
    
    '病区记帐报警线
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "病区ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "病区码", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "病区名", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "适用病人", adVarChar, 100
    mrsWarn.Fields.Append "报警方法", adSmallInt
    mrsWarn.Fields.Append "报警值", adCurrency
    mrsWarn.Fields.Append "报警标志1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "报警标志3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "催款下限", adCurrency
    mrsWarn.Fields.Append "催款标准", adCurrency
    
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    
    gstrSQL = "" & _
    "   Select a.病区ID,B.编码,b.名称 as 病区,a.适用病人,nvl(a.报警方法,1) as 报警方法, " & _
    "               a.报警值,a.报警标志1,a.报警标志2,a.报警标志3,A.催款下限,a.催款标准 " & _
    "   From 记帐报警线 a,部门表 b " & _
    "   Where a.病区ID= b.id(+)  " & _
    "   Order by Decode(a.适用病人,'普通病人',1,'医保病人',2,3),a.适用病人,B.编码 Desc"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    strCoding = ",普通病人" '至少有一个普通病人
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!病区id = rsTemp!病区id
        mrsWarn!病区码 = rsTemp!编码
        mrsWarn!病区名 = rsTemp!病区
        mrsWarn!适用病人 = rsTemp!适用病人
        mrsWarn!报警方法 = rsTemp!报警方法
        mrsWarn!报警值 = rsTemp!报警值
        mrsWarn!报警标志1 = rsTemp!报警标志1
        mrsWarn!报警标志2 = rsTemp!报警标志2
        mrsWarn!报警标志3 = rsTemp!报警标志3
        mrsWarn!催款下限 = Val(Nvl(rsTemp!催款下限))
        mrsWarn!催款标准 = Val(Nvl(rsTemp!催款标准))
        mrsWarn.Update
        
        If InStr(strCoding & ",", "," & rsTemp!适用病人 & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!适用病人
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab报警.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab报警.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab报警.Tabs(1).Selected = True '之前不会激活Click事件,人为激活
   
    '输出库房单位
    strCoding = ""
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.编码,'') 编码,nvl(b.名称,'') 名称,a.服务对象,a.工作性质" & vbCrLf & _
            "          FROM 部门性质说明 A, 部门表 B" & vbCrLf & _
            " WHERE B.ID=A.部门ID AND A.工作性质 IN ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房')  order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    msf库房单位.Rows = 1
    Do Until rsTemp.EOF
        With msf库房单位
            If rsTemp("编码") <> strCoding Then
                strTemp = ""
            End If
            If InStr(",中药库,西药库,成药库,", "," & rsTemp("工作性质") & ",") Then
                If InStr(1, strTemp & ",", ",药库,") <= 0 Then
                    .Rows = .Rows + 1
                    .RowData(.Rows - 1) = rsTemp("ID")
                    .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                    .TextMatrix(.Rows - 1, 1) = "药库"
                    strTemp = strTemp & "," & "药库"
                End If
            End If
            
            If InStr(",制剂室,中药房,西药房,成药房,", "," & rsTemp("工作性质") & ",") Then
            
                Select Case rsTemp("服务对象")
                    Case 0          '不服务于病人
'                        .Rows = .Rows + 1
'                        .RowData(.Rows - 1) = rsTemp("ID")
'                        .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
'                        .TextMatrix(.Rows - 1, 1) = "其他"
                    Case 1          '服务于门诊病人
                        If InStr(1, strTemp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTemp = strTemp & "," & "门诊"
                        End If
                    Case 2          '服务于住院病人
                        If InStr(1, strTemp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTemp = strTemp & "," & "住院"
                        End If
                    Case 3          '服务于门诊住院病人
                        If InStr(1, strTemp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTemp = strTemp & "," & "门诊"
                        End If
                        
                        If InStr(1, strTemp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTemp = strTemp & "," & "住院"
                        End If
                End Select
            End If
            If InStr(1, strTemp & ",", ",其他,") <= 0 Then
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = rsTemp("ID")
                .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                .TextMatrix(.Rows - 1, 1) = "其他"
                strTemp = strTemp & "," & "其他"
            End If
            
            strCoding = rsTemp("编码")
        End With
        rsTemp.MoveNext
    Loop

    If msf库房单位.Rows > 1 Then
        msf库房单位.FixedRows = 1
    End If
    gstrSQL = "select 库房id, 适用范围, 性质 from 药品库房单位"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngMaxRow = rsTemp.RecordCount
        For lngRow = 1 To lngMaxRow
            For i = 1 To msf库房单位.Rows - 1
                Select Case rsTemp!适用范围
                    Case 1
                        strTemp = "药库"
                    Case 2
                        strTemp = "门诊"
                    Case 3
                        strTemp = "住院"
                    Case 4
                        strTemp = "其他"
                End Select
                If rsTemp!库房id = msf库房单位.RowData(i) And strTemp = msf库房单位.TextMatrix(i, 1) Then
                    msf库房单位.TextMatrix(i, 2) = ""
                    msf库房单位.TextMatrix(i, 3) = ""
                    msf库房单位.TextMatrix(i, 4) = ""
                    msf库房单位.TextMatrix(i, 5) = ""
                    msf库房单位.TextMatrix(i, rsTemp!性质 + 1) = "√"
                End If
            Next
            rsTemp.MoveNext
        Next
    End If
    
    '药房配药控制
    strCoding = ""
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.编码,'') 编码,nvl(b.名称,'') 名称,a.服务对象,a.工作性质" & vbCrLf & _
            "          FROM 部门性质说明 A, 部门表 B" & vbCrLf & _
            " WHERE B.ID=A.部门ID AND A.工作性质 IN ('制剂室', '中药房', '西药房', '成药房')  order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    Bill药房配药控制.Rows = 1
    Do Until rsTemp.EOF
        With Bill药房配药控制
            If rsTemp("编码") <> strCoding Then
                strTemp = ""
            End If
            
            If InStr(",制剂室,中药房,西药房,成药房,", "," & rsTemp("工作性质") & ",") Then
            
                Select Case rsTemp("服务对象")
                    Case 0          '不服务于病人
'                        .Rows = .Rows + 1
'                        .RowData(.Rows - 1) = rsTemp("ID")
'                        .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
'                        .TextMatrix(.Rows - 1, 1) = "其他"
                    Case 1          '服务于门诊病人
                        If InStr(1, strTemp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTemp = strTemp & "," & "门诊"
                        End If
                    Case 2          '服务于住院病人
                        If InStr(1, strTemp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTemp = strTemp & "," & "住院"
                        End If
                    Case 3          '服务于门诊住院病人
                        If InStr(1, strTemp & ",", ",门诊,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "门诊"
                            strTemp = strTemp & "," & "门诊"
                        End If
                        
                        If InStr(1, strTemp & ",", ",住院,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("名称")
                            .TextMatrix(.Rows - 1, 1) = "住院"
                            strTemp = strTemp & "," & "住院"
                        End If
                End Select
            End If
            strCoding = rsTemp("编码")
        End With
        rsTemp.MoveNext
    Loop

    gstrSQL = "select 药房id, 门诊, 配药, 自动发药天数,配药确认 from 药房配药控制"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill药房配药控制
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            lngMaxRow = rsTemp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To .Rows - 1
                    Select Case rsTemp!门诊
                        Case 1
                            strTemp = "门诊"
                        Case 2
                            strTemp = "住院"
                    End Select
                    If rsTemp!药房id = .RowData(i) And strTemp = .TextMatrix(i, 1) Then
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load药品流向()
'功能:装入药品流向数据
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_药品流向)
        '首向装入库房
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('中药库','西药库','成药库','制剂室','中药房','西药房','成药房') " & _
                   " and  b.部门ID=a.ID and " & Where撤档时间("A") & " order by 编码"
        Call OpenRecordset(rsTemp, Me.Caption)
        
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
        Call OpenRecordset(rsTemp, Me.Caption)
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load费用类型()
'功能：初始化费用类型
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select 编码,名称 From 费用类型 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    lst(lst_医保病人).Clear
    lst(lst_公费病人).Clear
    Do Until rsTemp.EOF
        lst(lst_医保病人).AddItem rsTemp("编码") & "." & rsTemp("名称")
        lst(lst_公费病人).AddItem rsTemp("编码") & "." & rsTemp("名称")
        
        rsTemp.MoveNext
    Loop
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
    Dim ObjItem As ListItem
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.编码, B.名称, NVL(C.检查方式, 0) 检查方式" & vbCrLf & _
        " FROM 部门性质说明 A, 部门表 B, 药品出库检查 C" & vbCrLf & _
        " WHERE A.部门ID = B.ID AND A.部门ID = C.库房ID(+) AND" & vbCrLf & _
        "      A.工作性质 IN" & vbCrLf & _
        "      ('中药库', '西药库', '成药库', '制剂室', '中药房', '西药房', '成药房')" & vbCrLf & _
        "     And (b.撤档时间=to_date('3000-1-1','yyyy-mm-dd') or b.撤档时间 is null) " & vbCrLf & _
        " GROUP BY B.ID,B.编码, B.名称, NVL(C.检查方式, 0) " & vbCrLf & _
        " order by B.编码 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.lvwCheckMed.ListItems.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set ObjItem = Me.lvwCheckMed.ListItems.Add(, "C_" & rsTmp!ID, "[" & zlCommFun.Nvl(rsTmp!编码) & "]", "bm", "bm")
            ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!名称)
            ObjItem.SubItems(2) = Switch(rsTmp!检查方式 = 0, "0-不检查", rsTmp!检查方式 = 1, "1-检查，不足提醒", rsTmp!检查方式 = 2, "2-检查，不足禁止")
            ObjItem.Tag = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
ErrHandle:
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
    Dim strSQL As String
    On Error GoTo ErrHandle
    If msf库房单位.Rows > 1 Then
        If Trim(msf库房单位.TextMatrix(1, 0)) <> "" Then
            gstrSQL = ""
            For i = 1 To msf库房单位.Rows - 1
                gstrSQL = gstrSQL & msf库房单位.RowData(i) & ","
                lngTmp = 1
                Select Case True
                    Case msf库房单位.TextMatrix(i, 2) = "√"
                        lngTmp = 1
                    Case msf库房单位.TextMatrix(i, 3) = "√"
                        lngTmp = 2
                    Case msf库房单位.TextMatrix(i, 4) = "√"
                        lngTmp = 3
                    Case msf库房单位.TextMatrix(i, 5) = "√"
                        lngTmp = 4
                End Select
                Select Case msf库房单位.TextMatrix(i, 1)
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
            strSQL = "ZL_药品库房单位_DELETE"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gstrSQL = "ZL_药品库房单位_INSERT('" & gstrSQL & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Save库房检查() As Boolean
    '功能：保存库房检查
    Dim i As Long
    On Error GoTo ErrHandle
    
    gstrSQL = ""
    For i = 1 To Me.lvwCheckMed.ListItems.Count
        gstrSQL = gstrSQL & Me.lvwCheckMed.ListItems(i).Tag & "," & Switch(Me.lvwCheckMed.ListItems(i).SubItems(2) = "0-不检查", "0", Me.lvwCheckMed.ListItems(i).SubItems(2) = "1-检查，不足提醒", "1", Me.lvwCheckMed.ListItems(i).SubItems(2) = "2-检查，不足禁止", "2") & ","
    Next
    gstrSQL = "Zl_药品出库检查_insert('" & gstrSQL & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save库房检查 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Or lvw社区.Tag <> "" Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    '保存列宽
    SaveFlexState msh(0), App.ProductName & "\" & Me.Name
    SaveFlexState Bill(bill_自动计算), App.ProductName & "\" & Me.Name & bill_自动计算
    SaveFlexState Bill(bill_记帐报警), App.ProductName & "\" & Me.Name & bill_记帐报警
    SaveFlexState Bill(bill_药品流向), App.ProductName & "\" & Me.Name & bill_药品流向
    SaveFlexState Bill(bill_药品领用流向), App.ProductName & "\" & Me.Name & bill_药品领用流向
    
    Set mrsWarn = Nothing
    Set mrs类别 = Nothing
    Set mcol科室 = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If CheckDataValid() = False Then Exit Sub
    If Save数据() = False Then Exit Sub
    mblnChange = False
    lvw社区.Tag = ""
    Unload Me
End Sub

Private Function Check记帐报警() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr类别() As String
        
    With Bill(bill_记帐报警)
        For lngRow = 1 To .Rows - 2
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "病区“" & .TextMatrix(lngTemp, 0) & "”出现多次。", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = 0: .SetFocus: Exit Function
                    End If
                Next
                '刘兴洪 问题: 34770   日期:2010-12-21 10:54:02
                If Val(.TextMatrix(lngRow, 6)) > 999999999 Or Val(.TextMatrix(lngRow, 6)) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”中的催款下限设置有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 6: .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, 7)) > 999999999 Or Val(.TextMatrix(lngRow, 7)) < 0 Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”中的催款标准有误(应该在0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 7: .SetFocus: Exit Function
                End If
                
            End If
        Next
        
        '检查同一病区不同报警方式的类别是否一个都没有设置或重复
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                If Trim(.TextMatrix(lngRow, 3)) = "" And Trim(.TextMatrix(lngRow, 4)) = "" And Trim(.TextMatrix(lngRow, 5)) = "" Then
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”未设置要报警的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If (.TextMatrix(lngRow, 3) = "所有类别" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 4) = "所有类别" And (Trim(.TextMatrix(lngRow, 3)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 5) = "所有类别" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 3)) <> "")) Then
                    
                    MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, 3) <> "所有类别" And Trim(.TextMatrix(lngRow, 4)) <> "所有类别" And Trim(.TextMatrix(lngRow, 5)) <> "所有类别" Then
                    For lngCol1 = 3 To 5
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = 3 To 5
                                If lngCol1 <> lngCol2 Then
                                    arr类别 = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr类别)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr类别(lngTemp) & ",") > 0 Then
                                            MsgBox "病区“" & .TextMatrix(lngRow, 0) & "”不同的报警方式包含相同的收费类别。", vbExclamation, gstrSysName
                                            .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Next
    End With
    
    Check记帐报警 = True
End Function

Private Function CheckDataValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngIndex As Long, strTmp As String
    Dim i As Integer
    
    
    '自动对科室编号最后一个编辑操作进行校验
    If mintLastRow_Drug > 0 And Len(Trim(mstrLastCode_Drug)) > 0 Then
        With mshBillEdit
            If .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) <> UCase(mstrLastCode_Drug) Then
                .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) = UCase(mstrLastCode_Drug)
            End If
        End With
    End If
    If mintLastRow_Stuff > 0 And Len(Trim(mstrLastCode_Stuff)) > 0 Then
        With mshBillEditStuff
            If .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) <> UCase(mstrLastCode_Stuff) Then
                .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) = UCase(mstrLastCode_Stuff)
            End If
        End With
    End If
    
    '检查自动计算项目是否重复
    With Bill(bill_自动计算)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .RowData(lngRow) = .RowData(lngTemp) And .TextMatrix(lngRow, 1) = .TextMatrix(lngTemp, 1) Then
                        MsgBox "病区为“" & .TextMatrix(lngTemp, 0) & "”、收费细目为“" & _
                            .TextMatrix(lngTemp, 2) & "”" & vbCrLf & "这种组合出现多次。", vbExclamation, gstrSysName
                        .Row = lngTemp
                        .Col = 0
                        Call ShowTab(4)
                        .SetFocus
                        Exit Function
                    End If
                Next
            End If
        Next
    End With
    
    '检查自动计算项目的启用日期
    With Bill(bill_自动计算)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                If Not IsDate(.TextMatrix(lngRow, 4)) Then
                    MsgBox "自动计算项目的启用日期未设置或日期格式不正确。", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = 4
                    Call ShowTab(4)
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    '检查药品流向设置
    With Bill(bill_药品流向)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "第" & lngRow & "行信息不完整。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(8)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(8)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(8)
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
                Call ShowTab(11)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "第" & lngRow & "行中所在库房与对方库房相同。", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(11)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "第" & lngRow & "行与第" & lngTemp & "行信息库房相同了。", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(11)
                    Exit Function
                End If
            Next
        Next
    End With
    
    If txtUD(ud_费用金额保留位数).Text <> mDecimal Then
        If MsgBox("你已调整了费用金额保留小数位，可能会引起小数计算误差！是否继续？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    If txtUD(ud_费用单价保留位数).Text <> pDecimal Then
        If MsgBox("你已调整了费用单价保留小数位，可能会引起小数计算误差！是否继续？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            Call ShowTab(1)
            Exit Function
        End If
    End If
      
    If CheckNumberRule_Drug = True Then
'        With mshBillEdit
'            If Len(Trim(.TextMatrix(1, 1))) > 0 Then
'                For i = 1 To .Rows - 1
'                    If Len(Trim(.TextMatrix(i, 2))) <= 0 Then
'                        MsgBox "药品科室编号不能为空!", vbInformation, gstrSysName
'                        Call ShowTab(13)
'                        Exit Function
'                    End If
'                Next
'            End If
'        End With
        
        '同一个GRID里的科室编号不能重复
        With mshBillEdit
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "药品科室第" & i & "行编号重复！", vbQuestion, gstrSysName
                        Call ShowTab(13)
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
        strTmp = ""
    Else
        With mshBillEdit
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
        End With
    End If
    
    If CheckNumberRule_Stuff = True Then
'        With mshBillEditStuff
'            If Len(Trim(.TextMatrix(1, 1))) > 0 Then
'                For i = 1 To .Rows - 1
'                    If Len(Trim(.TextMatrix(i, 2))) <= 0 Then
'                        MsgBox "卫材科室编号不能为空!", vbInformation, gstrSysName
'                        Call ShowTab(13)
'                        Exit Function
'                    End If
'                Next
'            End If
'        End With
        
        '同一个GRID里的科室编号不能重复
        With mshBillEditStuff
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "卫材科室第" & i & "行编号重复！", vbQuestion, gstrSysName
                        Call ShowTab(13)
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
    Else
        With mshBillEditStuff
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
        End With
    End If
    
    If chk(chk_药品填单时下可用库存).Value <> chk(chk_药品填单时下可用库存).Tag Or chk(chk_明确申领药品批次).Value <> chk(chk_明确申领药品批次).Tag Then
        If Check是否有未审核的药品单据 Then
            MsgBox "还有未审核的药品单据，不能改变参数!", vbInformation, gstrSysName
            chk(chk_药品填单时下可用库存).Value = chk(chk_药品填单时下可用库存).Tag
            chk(chk_明确申领药品批次).Value = chk(chk_明确申领药品批次).Tag
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    If chk(chk_时价入库按折扣前采购价加成销售).Value <> chk(chk_时价入库按折扣前采购价加成销售).Tag Then
        If Check是否有未审核的外购入库单 Then
            MsgBox "还有未审核的外购入库单，不能改变参数“时价药品入库按扣前加成销售”!", vbInformation, gstrSysName
            chk(chk_时价入库按折扣前采购价加成销售).Value = chk(chk_时价入库按折扣前采购价加成销售).Tag
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    CheckDataValid = True
End Function

Private Function Save数据() As Boolean
    On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    Call SavePara
    Call Save电子签名
    Call Save社区接口
    Call Save单据操作
    Call Save收费特定项目
    Call Save自动计价项目
    Call Save药品流向
    Call Save记帐报警线
    Call SaveRegister
    Call Save库房检查
    Call Save库房单位
    Call Save药品领用流向
    Call Save药房配药控制
    Call Save单据编码规则
    Call Save科室
    Call Save医嘱内容
    Call Save药品卫材精度
    Call Save单据环节控制
    
    '保存完毕，事务提交
    gcnOracle.CommitTrans
    Call zlDatabase.ClearParaCache
    Save数据 = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call zlDatabase.ClearParaCache
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SavePara()
    Dim strTemp As String, lngTemp As Long
    Dim str类别 As String, i As Long

    On Error GoTo ErrHandle
    '逐个对参数进行保存
    strTemp = "1," & Format(dtp(dtp_上午上班).Value, "HH:mm") & " AND " & Format(dtp(dtp_上午下班).Value, "HH:mm") & ","
    strTemp = strTemp & "2," & Format(dtp(dtp_下午上班).Value, "HH:mm") & " AND " & Format(dtp(dtp_下午下班).Value, "HH:mm") & ","
    strTemp = strTemp & "5," & ud(ud_补录医嘱识别间隔).Value & ","
    strTemp = strTemp & "6," & chk(chk_未审核记帐处方发药).Value & ","
    strTemp = strTemp & "7," & chk(chk_自动修正).Value & ","
    strTemp = strTemp & "9," & Val(Me.txtUD(ud_费用金额保留位数).Text) & ","
    strTemp = strTemp & "157," & Val(Me.txtUD(ud_费用单价保留位数).Text) & ","
    strTemp = strTemp & "10," & chk(chk_收取预交款).Value & ","
    strTemp = strTemp & "11," & chk(chk_时办理就诊卡).Value & ","
    'strTemp = strTemp & "12," & chk(chk_密文显示).Value & ","
    strTemp = strTemp & "13," & chk(chk_分配床位号).Value & ","
    strTemp = strTemp & "14," & Split(cmb(cmb_挂号零钱处理).Text & "-", "-")(0) & cmb(cmb_收费零钱处理).ListIndex & cmb(cmb_结帐零钱处理).ListIndex & ","
    strTemp = strTemp & "15," & chk(chk_门诊收费与发药分离).Value & ","
    strTemp = strTemp & "16," & chk(chk_住院记帐与发药分离).Value & ","
    strTemp = strTemp & "17," & chk(chk_病人姓名).Value & chk(chk_刷就诊卡).Value & chk(chk_挂号单号).Value & chk(chk_病人ID).Value & ","
    strTemp = strTemp & "18," & chk(chk_限定药品的库存).Value & ","
    strTemp = strTemp & "45," & chk(chk_收费同时发药).Value & ","
    strTemp = strTemp & "51," & chk(chk_本人执行登记).Value & ","
    strTemp = strTemp & "52," & chk(chk_输入开单人).Value & ","
    strTemp = strTemp & "53," & chk(chk_它科开单人).Value & ","
    strTemp = strTemp & "54," & chk(chk_时价药品入库).Value & ","

    strTemp = strTemp & "19," & IIF(opt(opt_闲忙方式).Value = True, "0", "1") & ","
    strTemp = strTemp & "20,"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C1").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C2").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C3").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_票据).ListItems("C4").SubItems(1) & "|,"
    strTemp = strTemp & "21," & ud(ud_挂号单).Value & ud(ud_急诊挂号单).Value & ","
    strTemp = strTemp & "22," & (cmb(cmb_出院时未执行项目检查).ListIndex) & ","
    strTemp = strTemp & "23," & cmb(cmb_已结单据).ListIndex & ","
    strTemp = strTemp & "24,"
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C1").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C2").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C3").SubItems(2) = "√", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_票据).ListItems("C4").SubItems(2) = "√", "1", "0") & ","
    '电子签名系统参数
    strTemp = strTemp & "25," & Val(cmb(cmb_电子签名认证中心).List(cmb(cmb_电子签名认证中心).ListIndex)) & ","
    strTemp = strTemp & "26," & chk(chk_电子签名控制_门诊).Value & chk(chk_电子签名控制_住院).Value & chk(chk_电子签名控制_医技).Value & chk(chk_电子签名控制_护理).Value & chk(chk_电子签名控制_药品).Value & chk(chk_电子签名控制_lis).Value & chk(chk_电子签名控制_pacs).Value & ","
    strTemp = strTemp & "27," & chk(chk_住院药嘱发送产生领药号).Value & ","
    strTemp = strTemp & "28," & chk(chk_门诊病人消费时需要刷卡验证).Value & ","
    strTemp = strTemp & "29," & cmb(cmb_定价单位).ListIndex & ","
    strTemp = strTemp & "30," & cmb(cmb_合理用药接口).ListIndex & ","
    strTemp = strTemp & "31," & chk(chk_在院病人不准出院结帐).Value & ","
    strTemp = strTemp & "32," & cmb(cmb_转科时未执行项目检查).ListIndex & ","
    strTemp = strTemp & "33," & chk(chk_执行之后自动发料).Value & ","
    strTemp = strTemp & "34," & chk(chk_指定医嘱在其他科室执行).Value & ","
    '注意返回值是以,分隔，且外面有引号。保存时要转变一下
    strTemp = strTemp & "41," & Replace(Replace(GetTextFromList(lst(lst_医保病人)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "42," & Replace(Replace(GetTextFromList(lst(lst_公费病人)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "43," & chk(chk_下达出院医嘱才允许出院).Value & ","
    strTemp = strTemp & "44," & chk(chk_全数字只查编码).Value & chk(chk_全字母只查简码).Value & ","
    strTemp = strTemp & "163," & chk(chk_项目执行前必须收费或审核).Value & ","
    '74231,冉俊明,2014-6-24,项目开单后立即收费或记帐审核
    strTemp = strTemp & "232," & chk(chk_项目开单后立即收费或记帐审核).Value & ","
    strTemp = strTemp & "171," & chk(chk_允许未收费的门诊划价处方发料).Value & ","
    strTemp = strTemp & "172," & chk(chk_允许未审核的记账处方发料).Value & ","
    strTemp = strTemp & "185," & cboPatiVerfy.ItemData(cboPatiVerfy.ListIndex) & ","

    '刷卡要求输入密码的场合
    With lst(lst_刷卡密码)
        str类别 = ""
        For i = 0 To .ListCount - 1
            str类别 = str类别 & IIF(.Selected(i), 1, 0)
        Next
    End With
    strTemp = strTemp & "46," & str类别 & ","

    strTemp = strTemp & "55," & (cmb(cmb_诊断输入来源).ListIndex + 1) & ","
    strTemp = strTemp & "56," & IIF(chk(chk_门诊处方条数限制).Value = 0, 0, ud(ud_门诊处方条数限制).Value) & ","
    strTemp = strTemp & "58," & (cmb(cmb_未审单据结帐).ListIndex) & ","
    strTemp = strTemp & "59," & (cmb(cmb_医保对码检查).ListIndex) & ","
    strTemp = strTemp & "60," & IIF(Val(txtMaxMoney.Text) = 0, "", Val(txtMaxMoney.Text)) & ","
    strTemp = strTemp & "61," & (cmb(cmb_诊疗编码模式).ListIndex) & ","
    strTemp = strTemp & "63," & chk(chk_住院卫材自动发料).Value & ","
    strTemp = strTemp & "64," & (cmb(cmb_药品单据审核).ListIndex) & ","
    strTemp = strTemp & "65," & cmb(cmb_门诊诊断输入).ListIndex + 1 & cmb(cmb_住院诊断输入).ListIndex + 1 & ","
    strTemp = strTemp & "66," & ud(ud_挂号预约天数).Value & ","
    strTemp = strTemp & "68," & chk(chk_未作废临嘱禁止退药).Value & ","
    strTemp = strTemp & "69," & chk(chk_药品按规格下医嘱).Value & ","
    If chk(chk_过敏登记有效天数).Value = 0 Then
        strTemp = strTemp & "70," & chk(chk_过敏登记有效天数).Value & ","
    Else
        strTemp = strTemp & "70," & ud(ud_过敏登记有效天数).Value & ","
    End If
    strTemp = strTemp & "71," & chk(chk_长期医嘱次日生效).Value & ","
    strTemp = strTemp & "72," & chk(chk_首先输入收费类别).Value & ","
    strTemp = strTemp & "73," & chk(chk_明确申领药品批次).Value & ","
    strTemp = strTemp & "75," & chk(chk_外购入库需要核查).Value & ","
    strTemp = strTemp & "76," & chk(chk_时价药品直接确定售价).Value & ","

    With lst(lst_住院发送类别)
        str类别 = ""
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                str类别 = str类别 & Chr(.ItemData(i))
            End If
        Next
    End With
    strTemp = strTemp & "80," & str类别 & ","

    strTemp = strTemp & "81," & chk(chk_执行后自动审核划价单).Value & ","
    strTemp = strTemp & "84," & chk(chk_一次申请多个检验项目).Value & ","

    With lst(lst_门诊发送类别)
        str类别 = ""
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                str类别 = str类别 & Chr(.ItemData(i))
            End If
        Next
    End With
    strTemp = strTemp & "86," & str类别 & ","

    strTemp = strTemp & "87," & (cmb(cmb_药品编码模式).ListIndex) & ","
    strTemp = strTemp & "92," & chk(chk_门诊卫材自动发料).Value & ","
    strTemp = strTemp & "93," & chk(chk_从属项目汇总计算折扣).Value & ","
    strTemp = strTemp & "96," & chk(chk_药品填单时下可用库存).Value & ","
    'strTemp = strTemp & "97," & CStr(IIF(opt收费票据生成方式(1).Value, 1, 0) + Val(chk(chk_按执行科室分别打印).Value) * 10) & ","
    strTemp = strTemp & "98," & chk(chk_记帐报警包含划价费用).Value & ","
    strTemp = strTemp & "99," & chk(chk_入科确定护理等级).Value & ","
    strTemp = strTemp & "100," & chk(chk_下午算半天模式).Value & ","
    strTemp = strTemp & "126," & chk(chk_时价入库按折扣前采购价加成销售).Value & ","
    strTemp = strTemp & "143," & chk(chk_检验医嘱发送生成条形码).Value & ","
    strTemp = strTemp & "144," & chk(chk_收费项目首位当类别简码).Value & ","
    strTemp = strTemp & "145," & chk(chk_每次住院使用新住院号).Value & ","
    strTemp = strTemp & "147," & Val(txtUD(ud_儿童年龄界定上限).Text) & ","
    strTemp = strTemp & "148," & chk(chk_未收费处方发药).Value & ","
    strTemp = strTemp & "149," & (cmb(cmb_效期显示方式).ListIndex) & ","
    strTemp = strTemp & "150," & (cmb(cmb_药品出库优先算法).ListIndex) & ","
    strTemp = strTemp & "151," & chk(chk_门诊退费须先申请).Value & ","
    'strTemp = strTemp & "152," & chk(chk_就诊卡重复使用).Value & ","
    strTemp = strTemp & "154," & (cmb(cmb_出院时未发药项目检查).ListIndex) & ","
    strTemp = strTemp & "155," & (cmb(cmb_转科时未发药项目检查).ListIndex) & ","
    strTemp = strTemp & "158," & Val(txtInputHours.Text) & ","
    strTemp = strTemp & "160," & IIF(opt护理(1).Value, 1, 0) & ","
    strTemp = strTemp & "161," & chk(chk_禁忌药嘱).Value & ","
    strTemp = strTemp & "162," & chk(chk_下达医嘱时显示产地).Value & ","
    strTemp = strTemp & "173," & chk(chk_外购入库需要经过标记付款后才能进行付款).Value & ","
    strTemp = strTemp & "174," & chk(chk_药品移库明确批次).Value & ","
    strTemp = strTemp & "175," & chk(chk_药品领用明确批次).Value & ","
    strTemp = strTemp & "181," & chk(chk_时价分段加成入库).Value & ","
    strTemp = strTemp & "182," & IIF(chk(chk_禁止下达超极量药品医嘱).Value = 0, 1, 0) & ","
    strTemp = strTemp & "183," & chk(chk_时价药品取上次售价).Value & ","
    strTemp = strTemp & "186," & chk(chk_输血和皮试医嘱执行后需要核对).Value & ","
    strTemp = strTemp & "187," & chk(chk抗菌药物分级管理).Value & ","
    strTemp = strTemp & "188," & chk(chk抗菌药物使用自备药).Value & ","
    strTemp = strTemp & "189," & chk(chk允许下达院外执行的禁忌药品医嘱).Value & ","
    strTemp = strTemp & "191," & chk(chk只允许补录临嘱).Value & ","
    '55791:刘鹏飞,2012-11-13,回退出院医嘱才能撤销出院
    strTemp = strTemp & "192," & chk(chk_回退出院医嘱才允许撤销出院).Value & ","

    strTemp = strTemp & "208," & chk(chk临床工作站必须使用zlPlugIn部件).Value & ","
    strTemp = strTemp & "209," & chk(chk启用手术分级管理).Value & ","
    strTemp = strTemp & "210," & chk(chk_允许处理超过挂号有效天数的病人).Value & ","
    strTemp = strTemp & "213," & (IIF(cmb(cmd_中药配方).ListIndex = 1, 4, 3)) & ","

    strTemp = strTemp & "214," & chk(chk_首次医嘱执行需要审核).Value & ","
    '51612
    strTemp = strTemp & "215," & chk(chk_未入科禁止记账).Value & ","
    strTemp = strTemp & "216," & chk(chk_输血分级管理).Value & ","
    strTemp = strTemp & "217," & chk(chk_手术授权管理).Value & ","
    strTemp = strTemp & "218," & chk(chk_输血申请三级审核).Value & ","
    strTemp = strTemp & "219," & chk(chk_输血申请只能由中级及以上医师提出).Value & ","
    '允许修改n天内登记的医嘱执行记录
    strTemp = strTemp & "220," & IIF(chk(chk_医嘱执行有效天数).Value = 0, 999, Val(txtUNExecLimit.Text)) & ","
    strTemp = strTemp & "221," & IIF(optAccountTime(1).Value = True, 0, Val(txtAccountTime.Text)) & ","
    strTemp = strTemp & "223," & ud(ud_门诊新开医嘱间隔).Value & ","
    '过敏输入来源，启用太元通合理用药接口才使用。
    strTemp = strTemp & "224," & IIF(cmb(cmb_合理用药接口).ListIndex = 3, cmb(cmd_过敏输入来源).ListIndex, -1) & ","
    '启用大通合理用药接口才保存
    strTemp = strTemp & "225," & IIF(cmb(cmb_合理用药接口).ListIndex = 2, chk(chk_启用接口调用日志).Value, 0) & ","
    strTemp = strTemp & "226," & IIF(cmb(cmb_合理用药接口).ListIndex = 1, chk(chk_允许使用系统设置).Value, 1) & ","
    strTemp = strTemp & "227," & (cmb(cmd_转科时未审核销帐单据).ListIndex) & ","
    If cmb(cmb_合理用药接口).ListIndex = 1 Then
        strTemp = strTemp & "228," & IIF(optPASSVer(0).Value, "3.0", "4.0") & ","
    End If
    
    strTemp = strTemp & "230," & chk(chk_医嘱超量时必须输入原因).Value & ","
    
    If chk(chk_医嘱超量时必须输入原因).Value = 1 Then
        strTemp = strTemp & "233," & Get不写超量科室 & ","
    End If
    strTemp = strTemp & "234," & Get转科出院不检查项目 & ","
    
    strTemp = strTemp & "235," & (cmb(cmd_出院时超期护理数据).ListIndex) & ","
    strTemp = strTemp & "239," & chk(chk_新开医嘱签名时一组医嘱签名一次).Value & ","
    
    gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveRegister()
'保存到注册表中的信息

End Sub

Private Sub Save社区接口()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    For i = 1 To lvw社区.ListItems.Count
        With lvw社区.ListItems(i)
            gstrSQL = "Zl_社区目录_启用(" & Mid(.Key, 2) & "," & IIF(.SubItems(4) <> "", 1, 0) & ")"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save电子签名()
    Dim i As Integer, j As Long
    Dim strDept As String
    
    On Error GoTo ErrHandle
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            strDept = ""
            For j = 1 To .Rows - 1
                If .Cell(flexcpChecked, j, col_选择) = 1 Then
                    strDept = strDept & "," & .RowData(j)
                End If
            Next
            gstrSQL = "Zl_电子签名启用部门_Update(" & i & ",'" & Mid(strDept, 2) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End With
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save医嘱内容()
'保存医嘱内容定义
    On Error GoTo ErrHandle
    If cmdAdvice.Tag = "1" Then
        gstrSQL = "zl_医嘱内容定义_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        mrsAdvice.Filter = 0
        Do While Not mrsAdvice.EOF
            If Not IsNull(mrsAdvice!医嘱内容) Then
                gstrSQL = "zl_医嘱内容定义_Insert('" & mrsAdvice!诊疗类别 & "','" & Replace(mrsAdvice!医嘱内容, "'", "''") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            mrsAdvice.MoveNext
        Loop
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save单据操作()
    Dim lst As ListItem
    Dim i As Integer
    
    '首先删除以前的所有单据操作
    On Error GoTo ErrHandle
    gstrSQL = "zl_单据操作控制_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '再增加新的
    For Each lst In lvw(lvw_单据).ListItems
        gstrSQL = "zl_单据操作控制_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                    "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "是", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save收费特定项目()
    Dim strTemp As String
    
    '逐个对参数进行保存
    On Error GoTo ErrHandle
    If txtCmd(0).Text <> "" Then
        strTemp = "病历费," & txtCmd(0).Tag & ","
    End If
    If txtCmd(1).Text <> "" Then
        strTemp = strTemp & "工本费," & txtCmd(1).Tag & ","
    End If
    
    If txtCmd(3).Text <> "" Then
        strTemp = strTemp & "普通配置费," & txtCmd(3).Tag & ","
    End If
    
    If txtCmd(4).Text <> "" Then
        strTemp = strTemp & "肿瘤配置费," & txtCmd(4).Tag & ","
    End If
    
    If strTemp <> "" Then
        gstrSQL = "zl_收费特定项目_Modify('" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save自动计价项目old()
    Dim str病区ID As String
    Dim str细目ID As String
    Dim str计算标志 As String
    Dim str启用日期 As String
    Dim lngRow As Long
    Dim lngTemp As Long
    
    On Error GoTo ErrHandle
    With msh(0)
        For lngRow = 1 To .Rows - 1
            lngTemp = .RowData(lngRow)
            
            If lngTemp <> 0 Then
                If .TextMatrix(lngRow, 1) <> "" Then
                    str病区ID = str病区ID & lngTemp & ","
                    str细目ID = str细目ID & ","
                    str计算标志 = str计算标志 & "1,"
                    str启用日期 = str启用日期 & .TextMatrix(lngRow, 2) & ","
                End If
                If .TextMatrix(lngRow, 3) <> "" Then
                    str病区ID = str病区ID & lngTemp & ","
                    str细目ID = str细目ID & ","
                    str计算标志 = str计算标志 & "2,"
                    str启用日期 = str启用日期 & .TextMatrix(lngRow, 4) & ","
                End If
            End If
        Next
    End With
    With Bill(bill_自动计算)
        For lngRow = 1 To .Rows - 1
            lngTemp = .RowData(lngRow)
            
            If lngTemp <> 0 And .TextMatrix(lngRow, 1) <> "" Then
                str病区ID = str病区ID & lngTemp & ","
                str细目ID = str细目ID & .TextMatrix(lngRow, 1) & ","
                str计算标志 = str计算标志 & Switch(Left(.TextMatrix(lngRow, 3), 1) = "1", "6", Left(.TextMatrix(lngRow, 3), 1) = "2", "8", True, "7") & ","
                str启用日期 = str启用日期 & .TextMatrix(lngRow, 4) & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_自动计价项目_Modify('" & str病区ID & "','" & str细目ID & "','" & str计算标志 & "','" & str启用日期 & "' )"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save自动计价项目()
    Dim str病区ID As String
    Dim str细目ID As String
    Dim str计算标志 As String
    Dim str启用日期 As String
    Dim lngTemp As Long, i As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Zl_自动计价项目_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "删除自动计价项目")
    
    '按床位
    For i = 1 To msh(0).Rows - 1
        lngTemp = msh(0).RowData(i)
        If lngTemp <> 0 Then
            If msh(0).TextMatrix(i, 1) <> "" Then
                str病区ID = str病区ID & lngTemp & ","
                str细目ID = str细目ID & ","
                str计算标志 = str计算标志 & "1,"
                str启用日期 = str启用日期 & msh(0).TextMatrix(i, 2) & ","
            End If
            If msh(0).TextMatrix(i, 3) <> "" Then
                str病区ID = str病区ID & lngTemp & ","
                str细目ID = str细目ID & ","
                str计算标志 = str计算标志 & "2,"
                str启用日期 = str启用日期 & msh(0).TextMatrix(i, 4) & ","
            End If
        End If
        If (i Mod 100) = 0 Or i >= msh(0).Rows - 1 Then
            gstrSQL = "zl_自动计价项目_Modify('" & str病区ID & "','" & str细目ID & "','" & str计算标志 & "','" & str启用日期 & "' )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            str病区ID = ""
            str细目ID = ""
            str计算标志 = ""
            str启用日期 = ""
        End If
    Next
    '按病区
    For i = 1 To Bill(bill_自动计算).Rows - 1
        lngTemp = Bill(bill_自动计算).RowData(i)
        If lngTemp <> 0 And Bill(bill_自动计算).TextMatrix(i, 1) <> "" Then
            If Bill(bill_自动计算).TextMatrix(i, 1) <> "" Then
                str病区ID = str病区ID & lngTemp & ","
                str细目ID = str细目ID & Bill(bill_自动计算).TextMatrix(i, 1) & ","
                str计算标志 = str计算标志 & Switch(Left(Bill(bill_自动计算).TextMatrix(i, 3), 1) = "1", "6", Left(Bill(bill_自动计算).TextMatrix(i, 3), 1) = "2", "8", True, "7") & ","
                str启用日期 = str启用日期 & Bill(bill_自动计算).TextMatrix(i, 4) & ","
            End If
        End If
        If (i Mod 100) = 0 Or i >= Bill(bill_自动计算).Rows - 1 Then
            gstrSQL = "zl_自动计价项目_Modify('" & str病区ID & "','" & str细目ID & "','" & str计算标志 & "','" & str启用日期 & "' )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            str病区ID = ""
            str细目ID = ""
            str计算标志 = ""
            str启用日期 = ""
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save药品流向()
    Dim strTemp As String
    Dim lngRow As Long
    Dim str流向 As String
    
    On Error GoTo ErrHandle
    With Bill(bill_药品流向)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                str流向 = Left(.TextMatrix(lngRow, 3), 1)
                If str流向 = "" Then str流向 = "3"
                strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str流向 & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_药品流向控制_Modify('" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume

    Call SaveErrLog
End Sub

Private Sub Save记帐报警线()
    Dim strTemp As String
    Dim i As Integer
    Dim strArr
    Dim str适用病人 As String
    
    '先处理删除的适用病人记帐报警
    On Error GoTo ErrHandle
    If mstrDel适用病人 <> "" Then
        mstrDel适用病人 = mstrDel适用病人 & ";"
        strArr = Split(mstrDel适用病人, ";")
        For i = 0 To UBound(strArr) - 1
            If strArr(i) <> "" Then
                str适用病人 = strArr(i)
                strTemp = str适用病人 & "|"
                gstrSQL = "zl_记帐报警线_Modify('" & strTemp & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End If
    
    '按适用病人分批保存
    mrsWarn.Filter = 0
    For i = 1 To tab报警.Tabs.Count
        strTemp = ""
        str适用病人 = tab报警.Tabs.Item(i).Caption
        
        mrsWarn.Filter = "适用病人='" & str适用病人 & "'"
        Do While Not mrsWarn.EOF
            strTemp = strTemp & Nvl(mrsWarn!病区id) & "," & mrsWarn!报警方法 & "," & _
                mrsWarn!报警值 & "," & Nvl(mrsWarn!报警标志1) & "," & Nvl(mrsWarn!报警标志2) & "," & Nvl(mrsWarn!报警标志3) & "," & Nvl(mrsWarn!催款下限) & "," & Nvl(mrsWarn!催款标准) & ","
            mrsWarn.MoveNext
        Loop
        
        strTemp = str适用病人 & "|" & strTemp
        
        gstrSQL = "zl_记帐报警线_Modify('" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chk_Click(Index As Integer)
    Dim blnIsChange As Boolean
    
    mblnChange = True
    
    Select Case True
'        Case Index = chk_加收工本费 '56963
'            fra(10).Enabled = (chk(chk_加收工本费).Value = 1)
            
        Case Index = chk_票号控制
            lvw(lvw_票据).SelectedItem.SubItems(2) = IIF(chk(Index).Value = 1, "√", "")
        Case Index = chk_门诊收费与发药分离
            If chk(chk_门诊收费与发药分离).Value <> 0 Then
                chk(chk_收费同时发药).Enabled = False
                chk(chk_收费同时发药).Value = 0
            Else
                chk(chk_收费同时发药).Enabled = True
            End If
        Case Index = chk_过敏登记有效天数
            If chk(Index).Value = 0 Then
                ud(ud_过敏登记有效天数).Enabled = False
                txtUD(ud_过敏登记有效天数).Enabled = False
                txtUD(ud_过敏登记有效天数).BackColor = Me.BackColor
            Else
                ud(ud_过敏登记有效天数).Enabled = True
                txtUD(ud_过敏登记有效天数).Enabled = True
                txtUD(ud_过敏登记有效天数).BackColor = RGB(255, 255, 255)
            End If
        Case Index = chk_门诊处方条数限制
            If chk(Index).Value = 0 Then
                ud(ud_门诊处方条数限制).Enabled = False
                txtUD(ud_门诊处方条数限制).Enabled = False
                txtUD(ud_门诊处方条数限制).BackColor = Me.BackColor
            Else
                ud(ud_门诊处方条数限制).Enabled = True
                txtUD(ud_门诊处方条数限制).Enabled = True
                txtUD(ud_门诊处方条数限制).BackColor = RGB(255, 255, 255)
            End If
        Case Index = chk_药品填单时下可用库存
            If chk(chk_药品填单时下可用库存).Value = 1 Then
                chk(chk_明确申领药品批次).Value = 1
                chk(chk_药品移库明确批次).Value = 1
                chk(chk_药品领用明确批次).Value = 1
            End If
        Case Index = chk_明确申领药品批次
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_药品填单时下可用库存).Value = 1 Then
                chk(chk_明确申领药品批次).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_药品移库明确批次
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_药品填单时下可用库存).Value = 1 Then
                chk(chk_药品移库明确批次).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_药品领用明确批次
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_药品填单时下可用库存).Value = 1 Then
                chk(chk_药品领用明确批次).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_电子签名控制_门诊 Or Index = chk_电子签名控制_住院 _
                Or Index = chk_电子签名控制_医技 Or Index = chk_电子签名控制_护理 Or Index = chk_电子签名控制_药品 _
                Or Index = chk_电子签名控制_lis Or Index = chk_电子签名控制_pacs
            '在使用电子签名的情况下，至少有一个场合需要控制签名
            If cmb(cmb_电子签名认证中心).ListIndex <> 0 Then
                If chk(chk_电子签名控制_门诊).Value = 0 And chk(chk_电子签名控制_住院).Value = 0 _
                    And chk(chk_电子签名控制_医技).Value = 0 And chk(chk_电子签名控制_护理).Value = 0 And chk(chk_电子签名控制_药品).Value = 0 _
                    And chk(chk_电子签名控制_lis).Value = 0 And chk(chk_电子签名控制_pacs).Value = 0 Then
                        If Index = chk_电子签名控制_护理 Then
                            chk(chk_电子签名控制_药品).Value = 1
                        ElseIf Index = chk_电子签名控制_药品 Then
                             chk(chk_电子签名控制_lis).Value = 1
                        ElseIf Index = chk_电子签名控制_lis Then
                             chk(chk_电子签名控制_pacs).Value = 1
                        ElseIf Index = chk_电子签名控制_pacs Then
                             chk(chk_电子签名控制_门诊).Value = 1
                        Else
                            chk(((Index - chk_电子签名控制_门诊 + 1) Mod 4) + chk_电子签名控制_门诊).Value = 1
                        End If
                End If
            End If
            If Index = chk_电子签名控制_护理 Then
                sstSign.TabVisible(sst_护理) = chk(chk_电子签名控制_护理).Value = 1
            ElseIf Index = chk_电子签名控制_药品 Then
                 sstSign.TabVisible(sst_药品) = chk(chk_电子签名控制_药品).Value = 1
            ElseIf Index = chk_电子签名控制_lis Then
                 sstSign.TabVisible(sst_lis) = chk(chk_电子签名控制_lis).Value = 1
            ElseIf Index = chk_电子签名控制_pacs Then
                 sstSign.TabVisible(sst_Pacs) = chk(chk_电子签名控制_pacs).Value = 1
            ElseIf Index = chk_电子签名控制_门诊 Then
                sstSign.TabVisible(sst_门诊) = chk(chk_电子签名控制_门诊).Value = 1
            ElseIf Index = chk_电子签名控制_住院 Then
                sstSign.TabVisible(sst_住院护士) = chk(chk_电子签名控制_住院).Value = 1
                sstSign.TabVisible(sst_住院医生) = chk(chk_电子签名控制_住院).Value = 1
            ElseIf Index = chk_电子签名控制_医技 Then
                sstSign.TabVisible(sst_医技) = chk(chk_电子签名控制_医技).Value = 1
            End If
        Case Index = chk_首先输入收费类别 And Visible
            If chk(Index).Value = 1 Then
                chk(chk_收费项目首位当类别简码).Value = 0
            End If
        Case Index = chk_收费项目首位当类别简码 And Visible
            If chk(Index).Value = 1 Then
                chk(chk_首先输入收费类别).Value = 0
            End If
       Case Index = chk_项目执行前必须收费或审核
            If chk(Index).Value = 1 Then
                chk(chk_未收费处方发药).Enabled = False
                chk(chk_允许未收费的门诊划价处方发料).Enabled = False
                
                chk(chk_未审核记帐处方发药).Caption = "允许未审核的记帐处方发药(只对住院有效)"
                chk(chk_允许未审核的记账处方发料).Caption = "允许未审核的记账处方发料(只对住院有效)"
            
            Else
                chk(chk_未收费处方发药).Enabled = True
                chk(chk_未审核记帐处方发药).Caption = "允许未审核的记帐处方发药"
                
                chk(chk_允许未审核的记账处方发料).Caption = "允许未收费的门诊划价处方发料"
                chk(chk_允许未收费的门诊划价处方发料).Enabled = True
            End If
        Case Index = chk_时价分段加成入库
            If chk(Index).Value = 1 Then
                chk(chk_时价药品入库).Value = 0
                chk(chk_时价药品入库).Enabled = False
            Else
                chk(chk_时价药品入库).Value = 0
                chk(chk_时价药品入库).Enabled = True
            End If
        Case Index = chk抗菌药物分级管理
            chk(chk抗菌药物使用自备药).Enabled = chk(Index).Value = 1
        Case Index = chk_禁忌药嘱
            If chk(Index).Value = 1 Then
                chk(chk允许下达院外执行的禁忌药品医嘱).Value = 0
                chk(chk允许下达院外执行的禁忌药品医嘱).Enabled = False
            Else
                chk(chk允许下达院外执行的禁忌药品医嘱).Value = 0
                chk(chk允许下达院外执行的禁忌药品医嘱).Enabled = True
            End If
        Case Index = chk启用手术分级管理
            If chk(Index).Value = 1 Then
                chk(chk_手术授权管理).Value = 0
                chk(chk_手术授权管理).Enabled = True
            Else
                chk(chk_手术授权管理).Value = 0
                chk(chk_手术授权管理).Enabled = False
            End If
        Case Index = chk_输血分级管理
            If chk(Index).Value = 1 Then
                chk(chk_输血申请三级审核).Value = 0
                chk(chk_输血申请三级审核).Enabled = True
                chk(chk_输血申请只能由中级及以上医师提出).Value = 0
                chk(chk_输血申请只能由中级及以上医师提出).Enabled = True
            Else
                chk(chk_输血申请三级审核).Value = 0
                chk(chk_输血申请三级审核).Enabled = False
                chk(chk_输血申请只能由中级及以上医师提出).Value = 0
                chk(chk_输血申请只能由中级及以上医师提出).Enabled = False
            End If
        Case Index = chk_医嘱执行有效天数
            txtUNExecLimit.Enabled = chk(Index).Value = 1
        Case Index = chk_医嘱超量时必须输入原因
            Call Set不写超量科室(chk(Index).Value = 1)
    End Select
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    mblnChange = True
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmb_Click(Index As Integer)
    mblnChange = True

    If Index = cmb_电子签名认证中心 Then
        If cmb(Index).ListIndex = 0 Then
            chk(chk_电子签名控制_门诊).Value = 0
            chk(chk_电子签名控制_住院).Value = 0
            chk(chk_电子签名控制_医技).Value = 0
            chk(chk_电子签名控制_护理).Value = 0
            chk(chk_电子签名控制_药品).Value = 0
            chk(chk_电子签名控制_lis).Value = 0
            chk(chk_电子签名控制_pacs).Value = 0
            chk(chk_电子签名控制_门诊).Enabled = False
            chk(chk_电子签名控制_住院).Enabled = False
            chk(chk_电子签名控制_医技).Enabled = False
            chk(chk_电子签名控制_护理).Enabled = False
            chk(chk_电子签名控制_药品).Enabled = False
            chk(chk_电子签名控制_lis).Enabled = False
            chk(chk_电子签名控制_pacs).Enabled = False
            sstSign.Enabled = False
            sstSign.TabVisible(sst_门诊) = True
            txtFind.Enabled = False
            cmdFind.Enabled = False
        Else
            If Not chk(chk_电子签名控制_门诊).Enabled Then
                chk(chk_电子签名控制_门诊).Value = 1
            End If
            chk(chk_电子签名控制_门诊).Enabled = True
            chk(chk_电子签名控制_住院).Enabled = True
            chk(chk_电子签名控制_医技).Enabled = True
            chk(chk_电子签名控制_护理).Enabled = True
            chk(chk_电子签名控制_药品).Enabled = True
            chk(chk_电子签名控制_lis).Enabled = True
            chk(chk_电子签名控制_pacs).Enabled = True
            sstSign.Enabled = True
            txtFind.Enabled = True
            cmdFind.Enabled = True
        End If
    ElseIf Index = cmb_合理用药接口 Then
        '美康时可见
        lblPassVer.Visible = cmb(Index).ListIndex = 1
        optPASSVer(0).Visible = cmb(Index).ListIndex = 1
        optPASSVer(1).Visible = cmb(Index).ListIndex = 1
        optPASSVer(1).Enabled = False  '美康4.0不成熟，暂时禁用
            
        If cmb(Index).ListIndex = 0 Then    '未启用接口
            chk(chk_禁忌药嘱).Enabled = False
            chk(chk_禁忌药嘱).Value = 0
            chk(chk_禁止下达超极量药品医嘱).Enabled = False
            chk(chk_禁止下达超极量药品医嘱).Value = 0
            chk(chk允许下达院外执行的禁忌药品医嘱).Enabled = False
            chk(chk允许下达院外执行的禁忌药品医嘱).Value = 0

            chk(chk_启用接口调用日志).Visible = False  '大通时可见
            chk(chk_允许使用系统设置).Visible = False  '美康时可见
            '太元通时可见
            cmb(cmd_过敏输入来源).Visible = False
            lblInfo(lbl_过敏输入来源).Visible = False
        Else
            chk(chk_禁忌药嘱).Enabled = True
            chk(chk允许下达院外执行的禁忌药品医嘱).Enabled = True

            If cmb(Index).ListIndex = 1 Then  '美康
                chk(chk_允许使用系统设置).Visible = True
                chk(chk_允许使用系统设置).Enabled = True
            Else
                chk(chk_允许使用系统设置).Visible = False
                chk(chk_允许使用系统设置).Enabled = False
            End If

            If cmb(Index).ListIndex = 2 Then  '大通
                chk(chk_禁止下达超极量药品医嘱).Enabled = True
                chk(chk_启用接口调用日志).Visible = True
            Else
                chk(chk_禁止下达超极量药品医嘱).Enabled = False
                chk(chk_禁止下达超极量药品医嘱).Value = 0
                chk(chk_启用接口调用日志).Visible = False
            End If
            If cmb(Index).ListIndex = 3 Then    '太元通
                cmb(cmd_过敏输入来源).ListIndex = 0
                cmb(cmd_过敏输入来源).Visible = True
                lblInfo(lbl_过敏输入来源).Visible = True
                cmb(cmd_过敏输入来源).Enabled = True
                lblInfo(lbl_过敏输入来源).Enabled = True
            Else
                cmb(cmd_过敏输入来源).Visible = False
                lblInfo(lbl_过敏输入来源).Visible = False
            End If
        End If


    End If
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle

    '病历费和工本费限定为定价项目
    strSQL = "select id,编码,名称,计算单位,说明 from 收费项目目录 where 类别='Z' and nvl(是否变价,0)=0"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If IsNumeric(txtCmd(Index).Tag) = False Then txtCmd(Index).Tag = 0
        strSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "id,0,0,2;编号,1000,0,2;名称,1800,0,1;单位,800,0,2;说明,2300,0,2", -1, "定价项目选择", , CStr(txtCmd(Index).Tag), 0, 3)
        If strSQL <> "" Then
            txtCmd(Index).Tag = CLng(Split(strSQL, ";")(0))
            txtCmd(Index).Text = Trim(Split(strSQL, ";")(2))
            txtCmd(Index).SetFocus
            mblnChange = True
        End If
    Else
        MsgBox "无任何项目可用！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp_Change(Index As Integer)
    Dim intNext As Integer

    mblnChange = True
    If Index < dtp_下午下班 Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).Value
        If dtp(intNext).Value < dtp(intNext).MinDate Then
            dtp(intNext).Value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
End Sub

Private Sub lst类别_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 And lst类别.Selected(Item) Then
        For i = 1 To lst类别.ListCount - 1
            lst类别.Selected(i) = False
        Next
    ElseIf Item > 0 And lst类别.Selected(Item) Then
        lst类别.Selected(0) = False
    End If
End Sub

Private Sub lst类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst类别_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst类别_LostFocus()
    lst类别.Visible = False
End Sub

Private Sub lst类别_Validate(Cancel As Boolean)
    Dim objGrid As Object, i As Integer
    
    Set objGrid = Bill(bill_记帐报警)
    
    With objGrid
        .TextMatrix(.Row, .Col) = Get类别选择
        If .TextMatrix(.Row, .Col) = "所有类别" Then
            For i = 3 To 5
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    mblnChange = True
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     If Index = lvw_单据 Then
        If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
            lvw(lvw_单据).SortOrder = IIF(lvw(lvw_单据).SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            mintColumn = ColumnHeader.Index - 1
            lvw(lvw_单据).SortKey = mintColumn
            lvw(lvw_单据).SortOrder = lvwAscending
        End If
     End If
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Index = lvw_单据 Then
        Call cmdOperate_Click(1)
    End If
End Sub

Private Sub lvw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    mblnChange = True
    
    Dim itemTemp As MSComctlLib.ListItem
    For Each itemTemp In lvw(Index).ListItems
        If Not itemTemp Is Item Then
            itemTemp.Checked = False
        End If
    Next
End Sub

Private Sub lvw_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim lng原值 As Long
    
    If Index = lvw_票据 Then
        lng原值 = Val(Item.SubItems(1))
        ud(ud_号码长度).Max = 20
        
        '设置最大值时，可能已经更改了列表中的值
        ud(ud_号码长度).Value = lng原值
        chk(chk_票号控制).Value = IIF(Item.SubItems(2) = "√", 1, 0)
    ElseIf Index = lvw_一卡通 Then
        cmdOneCard(1).Enabled = Item.Text <> ""
        cmdOneCard(2).Enabled = cmdOneCard(1).Enabled
    End If
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_票据 Then
        If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    ElseIf Index = lvw_单据 Then
        If KeyAscii = vbKeyReturn Then cmdOperate_Click (1)
    End If
End Sub

Private Sub lvwCheckMed_DblClick()
    If Not Me.lvwCheckMed.SelectedItem Is Nothing Then
        lvwCheckMed.SelectedItem.SubItems(2) = Switch(lvwCheckMed.SelectedItem.SubItems(2) = "0-不检查", "1-检查，不足提醒", lvwCheckMed.SelectedItem.SubItems(2) = "1-检查，不足提醒", "2-检查，不足禁止", lvwCheckMed.SelectedItem.SubItems(2) = "2-检查，不足禁止", "0-不检查")
    End If
End Sub

Private Sub lvwCheckMed_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = "C" Then
        Call lvwCheckMed_DblClick
    End If
End Sub

Private Sub lvwNo_DblClick()
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    Call 改变规则
End Sub

Private Sub lvwNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If lvwNo.SelectedItem Is Nothing Then Exit Sub
        Call 改变规则
    End If
End Sub

Private Sub lvw社区_DblClick()
    If Not lvw社区.SelectedItem Is Nothing Then
        If lvw社区.SelectedItem.SubItems(4) <> "" Then
            lvw社区.SelectedItem.SubItems(4) = ""
        Else
            lvw社区.SelectedItem.SubItems(4) = "√"
        End If
        lvw社区.Tag = "1"
        Call lvw社区_ItemClick(lvw社区.SelectedItem)
    End If
End Sub

Private Sub lvw社区_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmd社区参数.Enabled = Item.SubItems(4) <> ""
End Sub

Private Sub lvw社区_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call lvw社区_DblClick
    End If
End Sub

Private Sub msf库房单位_DblClick()
    Dim i As Long
    
    If msf库房单位.Col > 1 And msf库房单位.Row > 0 And Trim(msf库房单位.TextMatrix(msf库房单位.Row, 0)) <> "" Then
        msf库房单位.TextMatrix(msf库房单位.Row, 2) = ""
        msf库房单位.TextMatrix(msf库房单位.Row, 3) = ""
        msf库房单位.TextMatrix(msf库房单位.Row, 4) = ""
        msf库房单位.TextMatrix(msf库房单位.Row, 5) = ""
        msf库房单位.TextMatrix(msf库房单位.Row, msf库房单位.Col) = "√"
    End If
End Sub

Private Sub msf库房单位_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn Or KeyAscii = Asc(" ")) Then
        msf库房单位_DblClick
    End If
End Sub

Private Sub msh_Click(Index As Integer)
    With Me.msh(0)
        If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
            mintCurRow = .Row
            mintCurCol = .Col
            txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
            If .TextMatrix(.Row, .Col) <> "" Then
                txtDateInput.Text = .TextMatrix(.Row, .Col)
            End If
            txtDateInput.Visible = True
            txtDateInput.SetFocus
        End If
    End With
End Sub

Private Sub msh_DblClick(Index As Integer)
    With msh(Index)
        If .MouseRow > 0 And .MouseCol > 0 And .RowData(.MouseRow) <> 0 Then
            If .Col = 1 Or .Col = 3 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "床位类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "护理类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "√", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                mblnChange = True
            End If
        End If
    End With
End Sub

Private Sub msh_KeyPress(Index As Integer, KeyAscii As Integer)
    With msh(Index)
        If KeyAscii = vbKeyReturn Then
            If .Col = 1 Then
                .Col = 2
            ElseIf .Col = 4 Then
                If .Row = .Rows - 1 Then
                    Bill(bill_自动计算).SetFocus
                Else
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - .TopRow > 8 Then .TopRow = .Row - 8
                End If
            End If
        ElseIf KeyAscii = Asc(" ") Then
            If .Row > 0 And (.Col = 1 Or .Col = 3) And .RowData(.Row) <> 0 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "床位类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "护理类项目或其从属项目的价格调整不是按天来执行的，请检查。", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "√", "")
                .TextMatrix(.Row, IIF(.Col = 1, 2, 4)) = IIF(.Text = "", "", IIF(.TextMatrix(.Row, IIF(.Col = 1, 5, 6)) = "", Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd"), .TextMatrix(.Row, IIF(.Col = 1, 5, 6))))
                mblnChange = True
            End If
        Else
            If .Row > 0 And (.Col = 2 Or .Col = 4) And .TextMatrix(.Row, IIF(.Col = 2, 1, 3)) <> "" Then
                mintCurRow = .Row
                mintCurCol = .Col
                txtDateInput.Move (.Left + .CellLeft - 10), (.Top + .CellTop - 10), .CellWidth, .CellHeight
                If .TextMatrix(.Row, .Col) <> "" Then
                    txtDateInput.Text = .TextMatrix(.Row, .Col)
                End If
                txtDateInput.Visible = True
                txtDateInput.SetFocus
            End If
        End If
    End With
End Sub

Private Sub bill_CommandClick(Index As Integer)
'通过按钮选择收费细目
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_记帐报警 Then
        With Bill(Index)
            Call Set类别选择(.TextMatrix(.Row, .Col))
            
            lst类别.Left = .Left + .MsfObj.CellLeft
            If .Top + .MsfObj.CellTop + .MsfObj.CellHeight + lst类别.Height <= .Container.Height Then
                lst类别.Top = .Top + .MsfObj.CellTop + .MsfObj.CellHeight
            Else
                lst类别.Top = .Top + .MsfObj.CellTop - lst类别.Height - 30
            End If
            lst类别.Width = .MsfObj.CellWidth
            lst类别.ZOrder
            lst类别.Visible = True
            lst类别.SetFocus
        End With
    End If
    
    If Index = bill_自动计算 Then
        With Bill(bill_自动计算)
            If .TextMatrix(.Row, 3) <> "2-计算一次" Then
                blnRe = frmChargeListSel.ShowTree(strID, str名称, False)
            Else
                blnRe = frmChargeListSel.ShowTree(strID, str名称, True)
            End If
            If blnRe And strID <> "" Then
                If .TextMatrix(.Row, 3) <> "2-计算一次" Then
                    If Not IsRaiseByDate(strID) Then
                        MsgBox "项目[" & str名称 & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .SetFocus
                .TextMatrix(.Row, 1) = strID
                .TextMatrix(.Row, 2) = str名称
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-按收治日"
                mblnChange = True
            End If
        End With
    End If
    
    If Index = bill_药品领用流向 Then
        gstrSQL = "Select Distinct Id,编码,名称,简码 From 部门表 a,部门性质说明 b " & _
                  "Where a.id = b.部门id And b.工作性质 In('领药部门') " & _
                  "    and (a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.撤档时间 Is Null) " & _
                  "order by 编码 "
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "领药部门")
        
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> 1 Then Exit Sub
        If rsTmp.EOF = True Then Exit Sub
        
        With Bill(bill_药品领用流向)
            .TextMatrix(.Row, 0) = rsTmp("编码") & "-" & rsTmp("名称")
            .RowData(.Row) = rsTmp("ID")
        End With
        
    End If
    
End Sub
Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If Index = bill_药品流向 Then
        With Bill(bill_药品流向)
            If ListIndex < 0 Then Exit Sub
            If .Col = 0 Then
                .RowData(.Row) = .ItemData(ListIndex)
            'BUG 29812
            ElseIf .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            Else
                
            End If
            'BUG 29812
            '.TextMatrix(.Row, .Col) = .CboText
            
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
        End With
    End If
    
    If Index = bill_药品领用流向 Then
        With Bill(bill_药品领用流向)
            If ListIndex < 0 Then Exit Sub
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
            .TextMatrix(.Row, .Col) = .CboText
        End With
    End If
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If Index = bill_药品流向 And .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                If Index <> bill_药品领用流向 Then
                    .RowData(.Row) = .ItemData(.ListIndex)
                End If
            End If
            If Index = bill_记帐报警 Then
                If .TextMatrix(.Row, 1) = "" Then .TextMatrix(.Row, 1) = "1-累计费用"
            ElseIf Index = bill_药品流向 Then
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-两库房间可双向流通"
            End If
            If .Index = bill_药品领用流向 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            End If
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'处理最后一列的变化
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_自动计算 Then
        If .MouseCol <> 3 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "0"
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或者其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "1-按床日"
            Case "1"
                .TextMatrix(.Row, .Col) = "2-计算一次"
            Case Else
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                        MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或者其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "0-按收治日"
        End Select
    ElseIf Index = bill_药品流向 Then
        If .MouseCol <> .Cols - 1 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "1"
                .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
            Case "2"
                .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
            Case Else
                .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
        End Select
    ElseIf Index = bill_记帐报警 Then
        If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, 1) = IIF(Left(.TextMatrix(.Row, 1), 1) = "1", "2-每日费用", "1-累计费用")
            If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                .TextMatrix(.Row, 4) = ""  '每日费用无报警方式2
                
                '为“每日费用”时判断一下金额不能为负数
                If IsNumeric(.TextMatrix(.Row, 2)) Then
                    If Val(.TextMatrix(.Row, 2)) < 0 Then
                        .TextMatrix(.Row, 2) = "0.00"
                    End If
                Else
                    .TextMatrix(.Row, 2) = "0.00"
                End If
            End If
        End If
    End If
    mblnChange = True
End With
    
End Sub

Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim lmX As Integer
    Dim lmY As Integer
    Dim strTmp As String
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 And KeyCode = vbKeyReturn Then
                If .Text <> "" And Not IsDate(.Text) Then
                    If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                        .Text = ""
                        MsgBox "请输入正确的日期格式(yyyy-mm-dd或者yyyymmdd)。", vbInformation, gstrSysName
                    Else
                        .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                    End If
                    .TextMatrix(.Row, .Col) = .Text
                End If
            End If
                
            If .Col = 2 Then
                '收费细目列只处理回车键
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '到下一个控件
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '选择收费细目
                    If IsRecord(.Text) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = .TextMatrix(.Row, 2)
                    If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-按收治日"
                    mblnChange = True
                End If
            End If
        End If
        
        If Index = bill_记帐报警 Then
            If .Col = 2 Then
                '报警值列只处理回车键
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '到下一个控件
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '判断输入的合法性
                    .Text = Format(.Text, "##########0.00;-##########0.00;0.00;0,00")
                    mblnChange = True
                End If
            ElseIf .Col = 3 Then
                '禁止输入报警类别
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            ElseIf .Col = 6 Or .Col = 6 Then
                .Text = Format(.Text, "###0.00;-###.00;0.00;0,00")
                mblnChange = True
            End If
        End If
  
        
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
                              " And b.工作性质 In ('领药部门') and (a.编码 Like '" & UCase(strTmp) & "%' or a.名称 like '" & UCase(strTmp) & "%' or a.简码 like '" & UCase(strTmp) & "%')"
                    
                    lmX = Me.Left + Me.tabMain.Left + Me.fraMain(9).Left + Me.Bill(bill_药品领用流向).Left
                    lmY = Me.Top + Me.tabMain.Top + Me.fraMain(9).Top + Me.Bill(bill_药品领用流向).Top + Me.Bill(bill_药品领用流向).RowHeight(.Row) + 350
                    Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "领药部门", , , , , , True, lmX, lmY, 300, , , True)
                    
                    If rsTmp Is Nothing Then Cancel = True: Exit Sub
                    If rsTmp.State <> 1 Then Cancel = True: Exit Sub
                    If rsTmp.EOF = True Then Cancel = True: Exit Sub
        
                    With Bill(bill_药品领用流向)
                        .TextMatrix(.Row, 0) = rsTmp("编码") & "-" & rsTmp("名称")
                        .Text = rsTmp("编码") & "-" & rsTmp("名称")
                        .RowData(.Row) = rsTmp("ID")
                    End With
                    mblnChange = True
                End If
            
            End If
            
        End If
    End With

End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_自动计算 Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            Else
                .TxtCheck = False
            End If
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "0"
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "1-按床日"
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-计算一次"
                            Case Else
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                        MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "0-按收治日"
                        End Select
                        mblnChange = True
                    Case vbKey0
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "药品类和卫材类的自动计算方式不能改变。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "0-按收治日"
                        mblnChange = True
                    Case vbKey1
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "药品类和卫材类的自动计算类型不能改变。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "项目[" & .TextMatrix(.Row, 2) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价后再选择其他自动计算方式。", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "1-按床日"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-计算一次"
                        mblnChange = True
                End Select
            End If
        ElseIf Index = bill_药品流向 Then
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
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-所在库房可流向对方库房"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-对方库房可流向所在库房"
                        mblnChange = True
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-两库房间可双向流通"
                        mblnChange = True
                End Select
            End If
        ElseIf Index = bill_记帐报警 Then
            .TxtCheck = False
            If .Col = 1 Then
                
                '切换报警方法
                Select Case KeyAscii
                    Case Asc(" ")
                        '切换计算标志
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-每日费用"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-累计费用"
                        End Select
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-累计费用"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-每日费用"
                        mblnChange = True
                End Select
                If InStr(.TextMatrix(.Row, 1), "每日费用") > 0 Then
                    .TextMatrix(.Row, 4) = ""  '每日费用无报警方式2
                End If
            ElseIf InStr(1, "267", .Col) > 0 Then
                    .TxtCheck = True
                    .TextMask = "0123456789-"
                    .MaxLength = 10
            End If
        End If
    End With

End Sub

Private Sub mshBillEdit_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub mshBillEdit_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub mshBillEdit_EnterCell(Row As Long, Col As Long)
    With mshBillEdit
        Select Case .Col
            Case mGrdCol.号码
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Drug = Row
                mintLastCol_Drug = Col
            End Select
    End With
End Sub
Private Sub mshBillEdit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Drug = ""
    
    With mshBillEdit
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.号码
'                If CheckNumberRule_Drug = True Then
'                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
'                        MsgBox "号码必需输入！", vbOKOnly + vbInformation, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
'                End If
                
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
'                    zlCommFun.PressKey vbKeyTab
                    mshBillEditStuff.SetFocus
                End If
            Case mGrdCol.科室
        End Select
    End With
End Sub

Private Sub mshBillEdit_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Drug = Chr(KeyAscii)
    End If
End Sub

Private Sub mshBillEditStuff_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub mshBillEditStuff_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub mshBillEditStuff_EnterCell(Row As Long, Col As Long)
    With mshBillEditStuff
        Select Case .Col
            Case mGrdCol.号码
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Stuff = Row
                mintLastCol_Stuff = Col
            End Select
    End With
End Sub

Private Sub mshBillEditStuff_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Stuff = ""
    
    With mshBillEditStuff
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.号码
'                If CheckNumberRule_Stuff = True Then
'                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
'                        MsgBox "号码必需输入！", vbOKOnly + vbInformation, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
'                End If
                
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
                    zlCommFun.PressKey vbKeyTab
                End If
            Case mGrdCol.科室
        End Select
    End With
End Sub

Private Sub mshBillEditStuff_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "请先设置科室编码规则为“按执行科室分月编号”后，再设置科室编码。", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Stuff = Chr(KeyAscii)
    End If
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub optAccountTime_Click(Index As Integer)
    If optAccountTime(0).Value = True Then
        txtAccountTime.Enabled = True
    Else
        txtAccountTime.Enabled = False
    End If
End Sub

Private Sub opt护理_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub sstSign_Click(PreviousTab As Integer)
    mlngFindItem = 1
End Sub

Private Sub tab报警_Click()
    Dim lngRow As Long
    
    mrsWarn.Filter = "适用病人='" & tab报警.SelectedItem.Caption & "'"
    
    With Bill(bill_记帐报警)
        If mrsWarn.RecordCount = 0 Then
            .ClearBill
            .Rows = 2: .Row = 1: .Col = 1
            .RowData(1) = 0
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
        Else
            .ClearBill
            .Rows = mrsWarn.RecordCount + 1: .Row = 1: .Col = 1
            lngRow = 1
            Do Until mrsWarn.EOF
                .RowData(lngRow) = Nvl(mrsWarn!病区id, 0)
                .TextMatrix(lngRow, 0) = IIF(IsNull(mrsWarn!病区id), "*门诊*", mrsWarn!病区码 & "-" & mrsWarn!病区名)
                .TextMatrix(lngRow, 1) = IIF(mrsWarn!报警方法 = 1, "1-累计费用", "2-每日费用")
                .TextMatrix(lngRow, 2) = Format(mrsWarn!报警值, "##########0.00;-##########0.00;0.00;0.00")
                
                .TextMatrix(lngRow, 3) = Get类别名称串(Nvl(mrsWarn!报警标志1), mrs类别)
                .TextMatrix(lngRow, 4) = Get类别名称串(Nvl(mrsWarn!报警标志2), mrs类别)
                .TextMatrix(lngRow, 5) = Get类别名称串(Nvl(mrsWarn!报警标志3), mrs类别)
                .TextMatrix(lngRow, 6) = Format(mrsWarn!催款下限, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, 7) = Format(mrsWarn!催款标准, "###0.00;-###0.00;0.00;0.00")
                
                lngRow = lngRow + 1
                mrsWarn.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub txtAccountTime_Change()
    If Val(txtAccountTime.Text) < 0 Or Val(txtAccountTime.Text) > 31 Then
        txtAccountTime.Text = 25
    End If
End Sub

Private Sub txtAccountTime_KeyPress(KeyAscii As Integer)
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


Private Sub txtCmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtCmd(Index).Tag = ""
        txtCmd(Index).Text = ""
        mblnChange = True
    End If
End Sub

Private Sub txtCmd_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = Asc("*") Then
        Call cmdSelect_Click(Index)
    End If
End Sub

Private Sub txtDateInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtDateInput
            If Not IsDate(.Text) Then
                If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                    MsgBox "请输入正确的日期格式(yyyy-mm-dd或者yyyymmdd)。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                End If
            End If
            msh(0).TextMatrix(mintCurRow, mintCurCol) = .Text
            .Visible = False
        End With
    End If
End Sub

Private Sub txtDateInput_LostFocus()
    txtDateInput.Text = ""
    txtDateInput.Visible = False
    
End Sub

Private Sub txtFind_Change()
    mlngFindItem = 1
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Call cmdFind_Click
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdFind_Click
End Sub

Private Sub txtInputHours_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtInputHours_Validate(Cancel As Boolean)
    If Trim(txtInputHours.Text) = "" Or Val(txtInputHours.Text) < 0 Or Val(txtInputHours.Text) > 9999 Then
        MsgBox "请录入0-9999的数值范围。", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtMaxMoney_GotFocus()
    zlControl.TxtSelAll txtMaxMoney
End Sub

Private Sub txtMaxMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtMaxMoney_Validate(Cancel As Boolean)
    If Val(txtMaxMoney.Text) = 0 Then txtMaxMoney.Text = ""
End Sub

Private Sub txtUD_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).Value
    Else
        If Index = ud_门诊新开医嘱间隔 Then
            ud(Index).Value = Val(txtUD(Index).Text)
        End If
    End If
End Sub

Private Sub txtUNExecLimit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtUNExecLimit_Validate(Cancel As Boolean)
    If Trim(txtUNExecLimit.Text) = "" Or Val(txtUNExecLimit.Text) < 0 Or Val(txtUNExecLimit.Text) > 999 Then
        MsgBox "请录入0-999的数值范围。", vbInformation, gstrSysName
        Cancel = True
    Else
        txtUNExecLimit.Text = Val(txtUNExecLimit.Text)
    End If
End Sub

Private Sub ud_Change(Index As Integer)
    mblnChange = True
    '动态改变票号长度
    If Index = ud_号码长度 Then
        lvw(lvw_票据).SelectedItem.SubItems(1) = ud(ud_号码长度).Value
    End If
    If Index = ud_门诊新开医嘱间隔 Then
        txtUD(ud_门诊新开医嘱间隔).Text = ud(ud_门诊新开医嘱间隔).Value
    End If
End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    
    For i = fraMain.LBound To fraMain.UBound
        fraMain(i).Visible = False
    Next
    
    i = tabMain.SelectedItem.Index - 1
    fraMain(i).Visible = True
    fraMain(i).Move 270, 500
    fraMain(i).ZOrder 0
    
    Select Case tabMain.SelectedItem.Index
        Case 1 '常规
            cmb(cmb_住院号规则).SetFocus
        Case 2 '临床应用
            cmb(cmb_诊疗编码模式).SetFocus
        Case 4 '票据管理
            If lvw(lvw_一卡通).Enabled Then lvw(lvw_一卡通).SetFocus
        Case 5 '自动计算
            mblnJRaiseByDate = IsRaiseByDate("J")
            mblnHRaiseByDate = IsRaiseByDate("H")
            msh(0).SetFocus
        Case 6 '记帐报警
            tab报警.SetFocus
        Case 7 '权限
            If chk(chk_它科开单人).Enabled Then chk(chk_它科开单人).SetFocus
        Case 8 '单据操作
            lvw(lvw_单据).SetFocus
        Case 9 '药品流向
            Bill(bill_药品流向).SetFocus
        Case 10  '库房检查
            Me.lvwCheckMed.SetFocus
        Case 11  '药品库房单位
        Case 12 '药品领用流向
            Bill(bill_药品领用流向).SetFocus
        Case 13 '单据编码规则
        Case 14 '科室编号
        Case 15 '药房配药控制
    End Select
End Sub

Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub

Private Function IsRecord(ByVal strFind As String) As Boolean
'功能:分析输入内容是否是有效的数据库中表的记录
'参数:strFind SQL语句的条件
'返回值:有效返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    
    rsTemp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strFind, "'") > 0 Then
        MsgBox "输入了非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    gstrSQL = "select distinct A.编码,A.名称,A.规格,A.计算单位 ,A.id from 收费细目 A,收费别名 B,收费类别 C " & _
         " where A.ID=B.收费细目ID and A.是否变价 <> 1 and A.末级=1 and  A.类别=C.编码 and  (A.编码 like [1] or B.名称 like [2] " & _
         " or  upper(B.简码) like [2]) and " & Where撤档时间("A")
          
    With Bill(bill_自动计算)
        If .TextMatrix(.Row, 3) <> "2-计算一次" Then
            gstrSQL = gstrSQL & " and C.编码 Not In('4','5','6','7') "
        End If
    End With
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind & "%", "%" & UCase(strFind) & "%")
    
    If rsTemp.RecordCount < 1 Then Exit Function
    If rsTemp.RecordCount > 1 Then
        gstrSQL = ""
        gstrSQL = frmSelCurr.ShowCurrSel(Me, rsTemp, "编码,1000,0,2;名称,1800,0,1;规格,2300,0,2;计算单位,1000,0,2;id,0,0,2", -1, "选择收费细目")
        If gstrSQL = "" Then
            Exit Function
        End If
        If Bill(bill_自动计算).TextMatrix(Bill(bill_自动计算).Row, 3) <> "2-计算一次" Then
            If Not IsRaiseByDate(Val(Split(gstrSQL, ";")(4))) Then
                MsgBox "项目[" & Split(gstrSQL, ";")(1) & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_自动计算)
            .TextMatrix(.Row, 1) = Split(gstrSQL, ";")(4) ' rsTemp("ID")
            .TextMatrix(.Row, 2) = Split(gstrSQL, ";")(1) 'rsTemp("名称")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-按收治日"
            End If
        End With
    Else
        rsTemp.MoveFirst
        If Bill(bill_自动计算).TextMatrix(Bill(bill_自动计算).Row, 3) <> "2-计算一次" Then
            If Not IsRaiseByDate(Val(rsTemp!ID)) Then
                MsgBox "项目[" & rsTemp!名称 & "]" & "或其从属项目的价格调整不是按天来执行的，请重新调价。", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_自动计算)
            .TextMatrix(.Row, 1) = rsTemp("ID")
            .TextMatrix(.Row, 2) = rsTemp("名称")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-按收治日"
            End If
        End With
    End If
    IsRecord = True
End Function

Private Function NumIsValid(ByVal lngIndex As Long, ByVal strNumber As String) As Boolean
'功能:分析输入内容是否是一个有效的数字
'参数:strNumber  输入内容
'返回值:有效返回True,否则为False
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "请输入一个数值。", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "这个数太大了。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '江磊修改编号:2500
    '只处理医保病人和非医病人的表格
    If (lngIndex = 1 Or lngIndex = 2) And Left(Bill(lngIndex).TextMatrix(Bill(lngIndex).Row, 1), 1) = "1" Then
        If Val(strNumber) < -9999999999.999 Then
            MsgBox "这个数太小了。", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Val(strNumber) < 0 Then
            MsgBox "不能为负数。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    NumIsValid = True
End Function

Private Sub Set类别选择(str类别 As String)
'功能：根据类似"检查,治疗..."的串设置列表的选择情况
    Dim i As Integer, j As Integer
    Dim arr类别() As String
    
    For i = 0 To lst类别.ListCount - 1
        lst类别.Selected(i) = False
    Next
    
    If Trim(str类别) = "" Then
        Exit Sub
    ElseIf str类别 = "所有类别" Then
        For i = 0 To lst类别.ListCount - 1
            lst类别.Selected(i) = (i = 0)
        Next
    Else
        lst类别.Selected(0) = False
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    lst类别.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst类别.ListCount - 1
        If lst类别.Selected(i) Then
            lst类别.TopIndex = i: Exit For
        End If
    Next
End Sub

Private Function Get类别选择() As String
'功能：根据类别选择框选择的情况返回类似"检查,治疗..."的串
    Dim i As Integer, strTmp As String
    
    If lst类别.Selected(0) Then
        Get类别选择 = "所有类别"
    Else
        For i = 1 To lst类别.ListCount - 1
            If lst类别.Selected(i) Then
                strTmp = strTmp & "," & lst类别.List(i)
            End If
        Next
        Get类别选择 = Mid(strTmp, 2)
        If Get类别选择 = "" Then Get类别选择 = " " '为了能回车新增行
    End If
End Function

Private Function Get类别名称串(str类别 As String, rs类别 As ADODB.Recordset) As String
'功能：将类似"CDEFG"的类别转换为类似"检查,检验..."串
    Dim i As Integer, strTmp As String
    
    If str类别 = "" Then
        Get类别名称串 = " " '为了能按回车新增行
        Exit Function
    End If
    
    If str类别 = "-" Then
        Get类别名称串 = "所有类别"
        Exit Function
    End If
    
    For i = 1 To Len(str类别)
        rs类别.Filter = "编码='" & Mid(str类别, i, 1) & "'"
        If Not rs类别.EOF Then strTmp = strTmp & "," & rs类别!类别
    Next
    Get类别名称串 = Mid(strTmp, 2)
End Function

Private Function Get类别编码串(str类别 As String) As String
'功能：根据类似"检查,治疗"的串返回类似"CDEFG"的串
    Dim i As Integer, j As Integer
    Dim arr类别() As String, strTmp As String
    
    If Trim(str类别) = "" Then Exit Function
    
    If str类别 = "所有类别" Then
        Get类别编码串 = "-"
    Else
        arr类别 = Split(str类别, ",")
        For i = 0 To UBound(arr类别)
            For j = 1 To lst类别.ListCount - 1
                If lst类别.List(j) = arr类别(i) Then
                    strTmp = strTmp & Chr(lst类别.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get类别编码串 = strTmp
    End If
End Function

Sub Load药品领用流向()
    '''''''''''''''''''''''''''''''''''''''''
    '功能           读入药品领用部门
    '''''''''''''''''''''''''''''''''''''''''
    
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_药品领用流向)
        '装入流向控制数据
        gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('中药库','西药库','成药库','制剂室','中药房','西药房','成药房') " & _
                   " and  b.部门ID=a.ID and " & Where撤档时间("A") & " order by 编码"
        Call OpenRecordset(rsTemp, Me.Caption)
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
        Call OpenRecordset(rsTemp, Me.Caption)
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save药品领用流向()
    Dim strTemp As String
    Dim lngRow As Long
    Dim bln次数 As Boolean
    
    On Error GoTo ErrHand
    With Bill(bill_药品领用流向)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                If LenB(StrConv(strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ",", vbFromUnicode)) >= 4000 Then
                    If bln次数 = True Then
                        gstrSQL = "zl_药品领用流向控制_Modify('" & strTemp & "'," & 1 & ")"
                    Else
                        gstrSQL = "zl_药品领用流向控制_Modify('" & strTemp & "'," & 0 & ")"
                    End If
                    bln次数 = True
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    strTemp = .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                Else
                    strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                End If
            End If
        Next
    End With
    
    If bln次数 = True Then
        gstrSQL = "zl_药品领用流向控制_Modify('" & strTemp & "'," & 1 & ")"
    Else
        gstrSQL = "zl_药品领用流向控制_Modify('" & strTemp & "'," & 0 & ")"
    End If
    bln次数 = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    Call SaveErrLog
    End If
End Sub

Private Function Load单据编码规则() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lst As ListItem
    
    gstrSQL = "" & _
        "   Select 项目序号,项目名称,编号规则,decode(编号规则,2,'2-按执行科室分月编号',0,'0-按年顺序编号',1,'1-按日顺序编号','0-按年顺序编号') as 编号规则说明 " & _
        "   From 号码控制表 " & _
        "   where 项目序号 in ( 11,12,13,14,15,16,21,22,23,24,25,26,27,28,29,32,62,68,69,70,71,72,73,74,75,76,77) order by 项目序号 "
    
    Err = 0: On Error GoTo ErrHand:
    Load单据编码规则 = False
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    With rsTmp
        lvwNo.ListItems.Clear
        Do While Not rsTmp.EOF
            Set lst = lvwNo.ListItems.Add(, "K" & Nvl(!项目序号, 0), Nvl(!项目名称))
            lst.SubItems(1) = Nvl(!编号规则说明)
            If Nvl(!项目序号) >= 1 And Nvl(!项目序号) <= 16 Then
                lst.ForeColor = &HC85422
                lvwNo.ListItems("K" & Nvl(!项目序号, 0)).ListSubItems(1).ForeColor = &HC85422
            End If
            If Nvl(!项目序号) >= 21 And Nvl(!项目序号) <= 62 Then
                lst.ForeColor = &H68588
                lvwNo.ListItems("K" & Nvl(!项目序号, 0)).ListSubItems(1).ForeColor = &H68588
            End If
            If Nvl(!项目序号) >= 68 And Nvl(!项目序号) <= 77 Then
                lst.ForeColor = &H856701
                lvwNo.ListItems("K" & Nvl(!项目序号, 0)).ListSubItems(1).ForeColor = &H856701
            End If
            lst.Tag = Nvl(!编号规则, 0)
            If lvwNo.SelectedItem Is Nothing Then
                lst.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '2-住院号，3-门诊号
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select 项目序号,编号规则 as 参数值 From 号码控制表 Where 项目序号 in (2,3)"
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    rsTmp.Filter = "项目序号=2"
    If rsTmp.RecordCount > 0 Then cmb(cmb_住院号规则).ListIndex = Val("" & rsTmp!参数值)
    rsTmp.Filter = "项目序号=3"
    If rsTmp.RecordCount > 0 Then cmb(cmb_门诊号规则).ListIndex = Val("" & rsTmp!参数值)
    
    Load单据编码规则 = True
    Call SetEdit
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SetEdit
    
End Function
Private Function SetEdit()
    '功能:设置编辑属性
    Dim blnEdit As Boolean
    Dim blnData As Boolean
End Function
Private Sub 改变规则()
    '改变编码规则
    Dim strNo As String
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    strNo = lvwNo.SelectedItem.SubItems(1) & "-"
    Select Case Split(strNo, "-")(0)
        Case 0
            If Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16 Then
                strNo = "1-按日顺序编号"
                lvwNo.SelectedItem.Tag = "1"
            Else
                strNo = "2-按执行科室分月编号"
                lvwNo.SelectedItem.Tag = "2"
            End If
        Case 1
            strNo = "0-按年顺序编号"
            lvwNo.SelectedItem.Tag = "0"
        Case 2
            strNo = "0-按年顺序编号"
            lvwNo.SelectedItem.Tag = "0"
    End Select
    lvwNo.SelectedItem.SubItems(1) = strNo
        
'    If Split(StrNo, "-")(0) = 2 Then
'        StrNo = "1-按年顺序编号"
'        lvwNo.SelectedItem.Tag = "0"
'    Else
'        StrNo = "2-按执行科室分月编号"
'        lvwNo.SelectedItem.Tag = "2"
'    End If
'    lvwNo.SelectedItem.SubItems(1) = StrNo
End Sub
Sub Save单据编码规则()
    Dim lst As ListItem
    
    On Error GoTo ErrHandle
    For Each lst In lvwNo.ListItems
        gstrSQL = "ZL_号码控制表_Rule(" & Mid(lst.Key, 2) & "," & Val(lst.Tag) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Next
    
    '2-住院号,3-门诊号
    gstrSQL = "ZL_号码控制表_Rule(2," & cmb(cmb_住院号规则).ListIndex & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    gstrSQL = "ZL_号码控制表_Rule(3," & cmb(cmb_门诊号规则).ListIndex & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub InitFace()
    '初始化控件
    With mshBillEdit
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.选择) = "选择"
        .TextMatrix(0, mGrdCol.科室) = "科室"
        .TextMatrix(0, mGrdCol.号码) = "号码"

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mGrdCol.选择) = 5
        .ColData(mGrdCol.科室) = 5
        .ColData(mGrdCol.号码) = 4


        .ColWidth(mGrdCol.选择) = 0
        .ColWidth(mGrdCol.科室) = 2000
        .ColWidth(mGrdCol.号码) = 1600
        
        .ColAlignment(mGrdCol.号码) = 1
        
        .PrimaryCol = mGrdCol.科室
        .LocateCol = mGrdCol.号码
        .AllowAddRow = False
        .Active = True
    End With
    
    With mshBillEditStuff
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.选择) = "选择"
        .TextMatrix(0, mGrdCol.科室) = "科室"
        .TextMatrix(0, mGrdCol.号码) = "号码"

        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mGrdCol.选择) = 5
        .ColData(mGrdCol.科室) = 5
        .ColData(mGrdCol.号码) = 4


        .ColWidth(mGrdCol.选择) = 0
        .ColWidth(mGrdCol.科室) = 2000
        .ColWidth(mGrdCol.号码) = 1600
        
        .ColAlignment(mGrdCol.号码) = 1
        
        .PrimaryCol = mGrdCol.科室
        .LocateCol = mGrdCol.号码
        .AllowAddRow = False
        .Active = True
    End With
End Sub
Sub Save科室()
    '保存科室编号
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With mshBillEdit
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mGrdCol.科室)) <> "" Then  'And Trim(.TextMatrix(i, mGrdCol.号码)) <> "" Then
                '科室ID_IN   IN 科室编号.科室ID%TYPE,
                '编号_IN     IN 科室编号.编号%TYPE
                gstrSQL = "ZL_科室号码表_UPDATE("
                gstrSQL = gstrSQL & .RowData(i) & ","
                gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.号码)) & "',1)"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With

    With mshBillEditStuff
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mGrdCol.科室)) <> "" Then 'And Trim(.TextMatrix(i, mGrdCol.号码)) <> "" Then
                '科室ID_IN   IN 科室编号.科室ID%TYPE,
                '编号_IN     IN 科室编号.编号%TYPE
                gstrSQL = "ZL_科室号码表_UPDATE("
                gstrSQL = gstrSQL & .RowData(i) & ","
                gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.号码)) & "',2)"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Function CheckNumberRule() As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '功能       检查单据编码规则是否有"2"的
    '返回       有=True 无=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If .ListItems(i).SubItems(1) = "2-按执行科室分月编号" Then
                CheckNumberRule = True
                Exit For
            End If
        Next
    End With
    'Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16
End Function

Function CheckNumberRule_Drug() As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '功能       检查单据编码规则是否有"2"的
    '返回       有=True 无=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 21 And Mid(.ListItems(i).Key, 2) <= 62 Then
                If .ListItems(i).SubItems(1) = "2-按执行科室分月编号" Then
                    CheckNumberRule_Drug = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Function CheckNumberRule_Stuff() As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '功能       检查单据编码规则是否有"2"的
    '返回       有=True 无=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 68 And Mid(.ListItems(i).Key, 2) <= 77 Then
                If .ListItems(i).SubItems(1) = "2-按执行科室分月编号" Then
                    CheckNumberRule_Stuff = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Private Sub vsDept_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsDept_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_选择 Then Cancel = True
End Sub

Private Sub vsDept_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    If Col = col_选择 Then
        Order = 0
        With vsDept(Index)
            If .MouseCol = col_选择 And .MouseRow = .FixedRows - 1 Then
                If sstSign.Enabled = False Then Exit Sub
                If .ColData(col_选择) = "Check" Then
                    .Cell(flexcpPicture, 0, col_选择) = ils16.ListImages("UnCheck").Picture
                    .ColData(col_选择) = ""
                Else
                    .Cell(flexcpPicture, 0, col_选择) = ils16.ListImages("AllCheck").Picture
                    .ColData(col_选择) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .ColData(col_选择) = "Check" Then
                        .Cell(flexcpChecked, i, col_选择) = 1
                    Else
                        .Cell(flexcpChecked, i, col_选择) = 2
                    End If
                    
                Next
            End If
        End With
    End If
End Sub

Private Sub vsDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Call cmdFind_Click
    End If
End Sub

Private Sub vsDept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If vsDept(Index).Row > 0 Then
            vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_选择) = IIF(vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_选择) = 1, 2, 1)
        End If
    End If
End Sub

Private Sub vsfControlItem_DblClick()
    With vsfControlItem
        If .Row < 1 Then Exit Sub
        If .Col < 2 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            '核查时不能修改"发票号,发票代码,发票日期,发票金额"
            If .TextMatrix(.Row, 1) = "核查" And InStr(1, "发票号,发票代码,发票日期,发票金额", .TextMatrix(0, .Col)) > 0 Then Exit Sub
            
            '卫材外购无外观选项
            If .TextMatrix(.Row, 0) = "卫材外购" And .TextMatrix(0, .Col) = "外观" Then Exit Sub
            
            .TextMatrix(.Row, .Col) = "√"

        End If
    End With
End Sub

Private Sub Init不填超量说明(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    vsUnWriteDept.Clear
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,名称 from 部门表 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnWriteDept
        lngRow = (rsTmp.RecordCount + 3) \ 4
        If lngRow > 5 Then .Rows = lngRow
        
        For i = 1 To rsTmp.RecordCount
            Call mcol科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
                lngRow = (i - 1) \ 4
                lngCol = (i - 1) Mod 4
                .TextMatrix(lngRow, lngCol) = rsTmp!名称
                .Cell(flexcpData, lngRow, lngCol) = rsTmp!名称 & ""
                .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init转科出院不检查项目(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,名称 from 诊疗项目目录 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnCheckItem
        .Row = 0: .Col = 0
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(.Row, .Col) = rsTmp!名称 & ""
            .Cell(flexcpData, .Row, .Col) = rsTmp!ID & ""
            Call EnterNextCell(vsUnCheckItem)
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get不写超量科室() As String
    Dim i As Integer
    Dim j As Integer
    Dim strIDs As String
    
    With vsUnWriteDept
        For i = 0 To .Rows - 1
            For j = 0 To 3
                If .TextMatrix(i, j) <> "" Then
                    strIDs = strIDs & "," & Val(.TextMatrix(i, j + 4))
                End If
            Next
        Next
    End With
    strIDs = Replace(strIDs, ",", "|")
    Get不写超量科室 = Mid(strIDs, 2)
End Function

Private Function Get转科出院不检查项目() As String
    Dim i As Integer
    Dim j As Integer
    Dim strIDs As String
    
    With vsUnCheckItem
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    strIDs = strIDs & "|" & Val(.Cell(flexcpData, i, j))
                End If
            Next
        Next
    End With
    Get转科出院不检查项目 = Mid(strIDs, 2)
End Function

Private Sub Set不写超量科室(ByVal blnEdit As Boolean)
'功能：可不录入超量原因的科室（表格）可能性
    With vsUnWriteDept
        .Enabled = blnEdit
        .Editable = IIF(blnEdit, flexEDKbdMouse, flexEDNone)
        .ForeColor = IIF(blnEdit, Me.ForeColor, &H808080)
        .BackColor = IIF(blnEdit, &H80000005, Me.BackColor)
    End With
End Sub

Private Sub vsUnCheckItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsUnCheckItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsUnCheckItem.ComboList = "..."
End Sub

Private Sub vsUnCheckItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = "select A.ID,A.编码,A.名称 from 诊疗项目目录 A Where A.类别 not in('4','5','6','7') and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order By 编码"
    With vsUnCheckItem
        vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "诊疗项目", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetItemInput(Row, Col, rsTmp)
            Call vsUnCheckItem_AfterRowColChange(-1, -1, Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "没有可用的诊疗项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnCheckItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    With vsUnCheckItem
        If KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsUnCheckItem_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
        ElseIf KeyCode = vbKeyReturn Then
            Call EnterNextCell(vsUnCheckItem)
        End If
        
    End With
End Sub

Private Sub vsUnCheckItem_KeyPress(KeyAscii As Integer)
    With vsUnCheckItem
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnCheckItem_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsUnCheckItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsUnCheckItem
        If .EditText = CStr(.TextMatrix(Row, Col)) Then
            Call EnterNextCell(vsUnCheckItem)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select DISTINCT A.ID,A.编码,A.名称 from 诊疗项目目录 A, 诊疗项目别名 B where " & _
            " a.Id = b.诊疗项目id And B.码类=1 And B.性质=1 And A.类别 not in('4','5','6','7') and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(B.简码) Like [2])" & _
            " Order by A.编码"
        With vsUnCheckItem
            vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗项目", False, "", "", False, False, True, _
                vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                Call SetItemInput(Row, Col, rsTmp)
                .EditText = .TextMatrix(Row, Col)
                mblnChange = True
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnCheckItem_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
    End With
End Sub

Private Sub vsUnWriteDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If vsUnWriteDept.Editable = flexEDNone Then
        vsUnWriteDept.FocusRect = flexFocusLight
        vsUnWriteDept.ComboList = ""
    Else
        vsUnWriteDept.FocusRect = flexFocusSolid
        vsUnWriteDept.ComboList = "..."
    End If
End Sub

Private Sub vsUnWriteDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUnWriteDept.Editable = flexEDNone Then
        vsUnWriteDept.FocusRect = flexFocusLight
        vsUnWriteDept.ComboList = ""
    Else
        vsUnWriteDept.FocusRect = flexFocusSolid
        vsUnWriteDept.ComboList = "..."
    End If
End Sub

Private Sub vsUnWriteDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
        " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) Order by A.简码"
    With vsUnWriteDept
        vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "临床科室", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsUnWriteDept_AfterRowColChange(-1, -1, Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "没有可用的科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnWriteDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsUnWriteDept
        If KeyAscii = 13 Then
            KeyAscii = 0
            If .EditText = CStr(.Cell(flexcpData, Row, Col)) Then
                Call EnterNextCell(vsUnWriteDept)
                Exit Sub
            End If
            strInput = UCase(.EditText)
            strSQL = "select Distinct a.id,a.编码,a.名称,a.简码 from 部门表 a,部门性质说明 b where a.id=b.部门id" & _
                " and b.工作性质='临床' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
                " Order by A.简码"
            With vsUnWriteDept
                vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "临床科室", False, "", "", False, False, True, _
                    vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
                If Not rsTmp Is Nothing Then
                    Call SetDeptInput(Row, Col, rsTmp)
                    .EditText = .TextMatrix(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                End If
            End With
            Call vsUnWriteDept_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        End If
    End With
End Sub

Private Sub vsUnWriteDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
    If vsUnWriteDept.TextMatrix(Row, Col + 4) = "" Then vsUnWriteDept.TextMatrix(Row, Col) = ""
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset)
    '先检查下表格中是否存在
    Dim strTmp As String
    With vsUnWriteDept
        On Error Resume Next
        strTmp = mcol科室("_" & rsTmp!ID)
        If Err.Number = 0 Then
            MsgBox "该科室已经存在，请重新输入。", vbInformation, gstrSysName
            .TextMatrix(lngRow, lngCol) = CStr(.Cell(flexcpData, lngRow, lngCol))
            Exit Sub
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If .TextMatrix(lngRow, lngCol + 4) <> "" Then
            Call mcol科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
        End If
        
        .TextMatrix(lngRow, lngCol) = rsTmp!名称 & ""
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!名称 & ""
        .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID & ""
        Call mcol科室.Add(rsTmp!ID & "", "_" & rsTmp!ID)
    End With
End Sub

Private Function SetItemInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset)
    '先检查下表格中是否存在
    Dim i As Long, j As Long
    
    With vsUnCheckItem
        For i = .FixedCols To .Cols - 1
            For j = .FixedRows To .Rows - 1
                If .Cell(flexcpData, j, i) = rsTmp!ID & "" Then
                    MsgBox "该诊疗项目已经加入列表中，请查看。", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
        
        .TextMatrix(lngRow, lngCol) = rsTmp!名称 & ""
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!ID & ""
        SetItemInput = True
        mblnChange = True
    End With
End Function

Private Sub vsUnWriteDept_KeyPress(KeyAscii As Integer)
    If vsUnWriteDept.Editable = flexEDNone Then Exit Sub

    With vsUnWriteDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnWriteDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsUnWriteDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    With vsUnWriteDept
        If KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsUnWriteDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcol科室.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
            .TextMatrix(.Row, .Col + 4) = ""
        End If
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Call EnterNextCell(vsUnWriteDept)
    End With
End Sub

Private Sub EnterNextCell(ByVal vsobj As VSFlexGrid)
'功能：输框定位到下一个
    With vsobj
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then .AddItem ""
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '如果是隐藏行则递归再定位到下一个位置
        If .ColHidden(.Col) = True Then Call EnterNextCell(vsobj)
        .ShowCell .Row, .Col
    End With
End Sub
