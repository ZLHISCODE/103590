VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSystemPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   8050
      Index           =   2
      Left            =   240
      TabIndex        =   283
      Top             =   480
      Width           =   9690
      Begin VB.CheckBox chk 
         Caption         =   "�¿�ҽ��ǩ��ʱһ��ҽ��ǩ��һ��"
         Height          =   195
         Index           =   91
         Left            =   5760
         TabIndex        =   345
         Top             =   120
         Width           =   3540
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
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
         Caption         =   "ҩƷ��ҩ"
         Enabled         =   0   'False
         Height          =   195
         Index           =   60
         Left            =   2880
         TabIndex        =   291
         Top             =   720
         Width           =   1020
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   44
         Left            =   960
         TabIndex        =   290
         Top             =   480
         Width           =   1620
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   45
         Left            =   2880
         TabIndex        =   289
         Top             =   480
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "ҽ��ҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   46
         Left            =   4440
         TabIndex        =   288
         Top             =   480
         Width           =   1860
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����¼,������"
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
         TabCaption(0)   =   "����ҽ��,����"
         TabPicture(0)   =   "frmSystemPara.frx":0010
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "vsDept(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "סԺҽ��ҽ��,����"
         TabPicture(1)   =   "frmSystemPara.frx":002C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vsDept(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "סԺ��ʿҽ��"
         TabPicture(2)   =   "frmSystemPara.frx":0048
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vsDept(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "ҽ��ҽ��,����"
         TabPicture(3)   =   "frmSystemPara.frx":0064
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "vsDept(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "�����¼,������"
         TabPicture(4)   =   "frmSystemPara.frx":0080
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "vsDept(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "ҩƷ��ҩ"
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
         Caption         =   "˵�������ó��Ϻ�δ��ѡ�κβ��ţ���ʾ�������ҿ��ơ�"
         Height          =   255
         Left            =   240
         TabIndex        =   305
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   480
         TabIndex        =   294
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label6 
         Caption         =   "��֤����"
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
         ToolTipText     =   "����һ��ǰ׺"
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
         ToolTipText     =   "�޸ĵ�ǰǰ׺"
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
         ToolTipText     =   "ɾ����ǰǰ׺"
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
         Caption         =   "�ϸ����"
         Height          =   285
         Index           =   13
         Left            =   8040
         TabIndex        =   77
         ToolTipText     =   "��ʾ����������￨���봦�Ƿ�Ϊ������ʾ"
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
            Text            =   "Ʊ������"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "���볤��"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "�ϸ����"
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
            Text            =   "���"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Name"
            Text            =   "����"
            Object.Width           =   3882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Key             =   "PayType"
            Text            =   "���㷽ʽ"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "OrgCode"
            Text            =   "ҽԺ����"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Enable"
            Text            =   "����"
            Object.Width           =   970
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ��ͨ�ӿ�"
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
         Caption         =   "ˢ��Ҫ��������"
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
         Caption         =   "���볤��"
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
         Caption         =   "Ʊ�ݺ������"
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
            Caption         =   "ҽ������ʱ��������ԭ��"
            Height          =   240
            Index           =   86
            Left            =   105
            TabIndex        =   339
            ToolTipText     =   "��ѡʱ���ڱ���п��ҵĲ����´�ҽ���ɲ�д����˵��"
            Top             =   0
            Width           =   2280
         End
         Begin VB.Label Label16 
            Caption         =   "�����ÿɲ�¼�볬��ԭ��Ŀ��ң����磺����ơ�"
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
            Caption         =   "������Ѫ�����������"
            Enabled         =   0   'False
            Height          =   255
            Index           =   53
            Left            =   120
            TabIndex        =   276
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ����ֻ�����м�������ҽʦ���"
            Enabled         =   0   'False
            Height          =   200
            Index           =   85
            Left            =   2400
            TabIndex        =   275
            Top             =   285
            Width           =   3375
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ѫ�ּ�����"
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   274
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " ҽ������ִ�� "
         Height          =   2805
         Index           =   13
         Left            =   3675
         TabIndex        =   263
         Top             =   1005
         Width           =   5895
         Begin VB.CheckBox chk 
            Caption         =   "����ȡ��"
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
            Caption         =   "����ҽ������ʱ������������"
            Height          =   195
            Index           =   54
            Left            =   120
            TabIndex        =   271
            ToolTipText     =   "�Ƿ���ִ�в����󽫻��۵����Ϊ���ʵ�"
            Top             =   2040
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "ִ��֮���Զ���˼��ʻ��۵�"
            Height          =   210
            Index           =   32
            Left            =   120
            TabIndex        =   270
            ToolTipText     =   "�Ƿ���ִ�в����󽫻��۵����Ϊ���ʵ�"
            Top             =   2280
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "ִ��֮������������Զ�����"
            Height          =   210
            Index           =   61
            Left            =   2880
            TabIndex        =   269
            ToolTipText     =   "�Ƿ���ִ�в����󽫻��۵����Ϊ���ʵ�"
            Top             =   2280
            Width           =   2640
         End
         Begin VB.CommandButton cmdSendPriceType 
            Caption         =   "ȫѡ(&A)"
            Height          =   350
            Index           =   0
            Left            =   3480
            TabIndex        =   265
            Top             =   420
            Width           =   1100
         End
         Begin VB.CommandButton cmdSendPriceType 
            Caption         =   "ȫ��(&U)"
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
            TabCaption(0)   =   "����"
            TabPicture(0)   =   "frmSystemPara.frx":16DC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lst(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "סԺ"
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
            Caption         =   "���ڵ�ҽ��ִ�в���"
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
            Caption         =   "����Ϊ���ʻ��۵����������"
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
            Caption         =   "���ÿ���ҩ��ּ�����"
            Height          =   255
            Index           =   75
            Left            =   120
            TabIndex        =   261
            Top             =   0
            Width           =   2175
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩ������ʹ���Ա�ҩ"
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
            Caption         =   "��������ҽʦ��Ȩ����"
            Enabled         =   0   'False
            Height          =   240
            Index           =   49
            Left            =   120
            TabIndex        =   258
            Top             =   225
            Width           =   2220
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������ּ�����"
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
            Caption         =   "����4.0"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   330
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton optPASSVer 
            Caption         =   "����3.0"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   329
            Top             =   1320
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʹ��ϵͳ����"
            Height          =   240
            Index           =   89
            Left            =   120
            TabIndex        =   325
            Top             =   1080
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����´�Ժ��ִ�еĽ���ҩƷҽ��"
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   324
            Top             =   600
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ֹ�´ﳬ����ҩƷҽ��"
            Height          =   240
            Index           =   63
            Left            =   120
            TabIndex        =   323
            Top             =   840
            Width           =   2940
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����´����ҩƷҽ��"
            Height          =   240
            Index           =   65
            Left            =   120
            TabIndex        =   322
            Top             =   350
            Width           =   2100
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ýӿڵ�����־"
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
            Caption         =   "��ǰ�汾��"
            Height          =   255
            Left            =   120
            TabIndex        =   328
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Դ"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   314
            Top             =   1395
            Width           =   1080
         End
         Begin VB.Label lbl������ҩ�ӿ� 
            Caption         =   "������ҩ�ӿ�"
            Height          =   255
            Left            =   270
            TabIndex        =   252
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmd�������� 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8470
         TabIndex        =   73
         ToolTipText     =   "�Ե�ǰѡ��������ӿڵĲ�����������"
         Top             =   6690
         Width           =   1100
      End
      Begin VB.Frame Fra 
         Caption         =   " ������� "
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
            ToolTipText     =   "Ӱ�췶Χ�����Ժ��ҽ������վ"
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
            ToolTipText     =   "Ӱ�췶Χ�����Ժ��ҽ������վ"
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
            ToolTipText     =   "Ӱ�췶Χ��ҽ������վ"
            Top             =   210
            Width           =   2310
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ"
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
            Caption         =   "����"
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
            Caption         =   "��Դ"
            Height          =   180
            Index           =   39
            Left            =   285
            TabIndex        =   197
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.Frame Fra 
         Caption         =   "ҽ�����"
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
            Caption         =   "���˳�Ժҽ�����ܳ���Ԥ��Ժ"
            Height          =   240
            Index           =   81
            Left            =   240
            TabIndex        =   67
            Top             =   4020
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�"
            Height          =   240
            Index           =   74
            Left            =   240
            TabIndex        =   246
            Top             =   4860
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "�´�ҽ��ʱ��ʾ����"
            Height          =   240
            Index           =   66
            Left            =   240
            TabIndex        =   240
            Top             =   4575
            Value           =   1  'Checked
            Width           =   2820
         End
         Begin VB.CheckBox chk 
            Caption         =   "ָ��ҽ������������ִ��"
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
            Caption         =   "�´��Ժҽ�����ܳ�Ժ"
            Height          =   240
            Index           =   50
            Left            =   240
            TabIndex        =   66
            Top             =   3750
            Width           =   2340
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҽ��ȱʡΪ������Ч"
            Height          =   240
            Index           =   24
            Left            =   240
            TabIndex        =   62
            Top             =   2655
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ����ҽ��������´�"
            Height          =   240
            Index           =   3
            Left            =   240
            TabIndex        =   61
            ToolTipText     =   " "
            Top             =   2370
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ҩ�������Ϻ���ҩ"
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
            Caption         =   "һ��������������Ŀ"
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
            Caption         =   "ҽ�����ݶ���(&F)"
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
            Caption         =   "�����Ǽ���Ч����"
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
            Caption         =   "����ҩƷ��������"
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
            Caption         =   "סԺҩ�����Ͳ�����ҩ��"
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
            Caption         =   "���Ե�ǰʱ����Ϊ��ʼʱ��"
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
            Caption         =   "�����¿�ҽ�����         ����"
            Height          =   180
            Index           =   25
            Left            =   540
            TabIndex        =   308
            Top             =   940
            Width           =   2610
         End
         Begin VB.Label lbl 
            Caption         =   "��ҩ�䷽ÿ��"
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
            Caption         =   "��ͯ����綨����         ��"
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
            Caption         =   "��¼ҽ��ʶ����         ����"
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
            Caption         =   "���Ʊ���"
            Height          =   180
            Index           =   36
            Left            =   240
            TabIndex        =   195
            Top             =   315
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvw���� 
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
            Text            =   "���"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "˵��"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "����"
            Object.Width           =   952
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "���������ӿڣ�"
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
         Caption         =   " ҩƷ���ʱ�� "
         Height          =   1080
         Index           =   10
         Left            =   6840
         TabIndex        =   278
         Top             =   6840
         Width           =   2775
         Begin VB.OptionButton optAccountTime 
            Caption         =   "ÿ��"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   281
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optAccountTime 
            Caption         =   "ÿ�����һ��"
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
               Name            =   "����"
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
            Caption         =   "��"
            Height          =   255
            Left            =   1560
            TabIndex        =   282
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " ��Ժʱ���� "
         Height          =   1005
         Index           =   5
         Left            =   6840
         TabIndex        =   140
         Top             =   1275
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "������￨"
            Height          =   195
            Index           =   5
            Left            =   1500
            TabIndex        =   40
            ToolTipText     =   "��ʾ�ڰ�����Ժʱ�Ƿ�����ͬʱ������￨"
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ȡԤ����"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   39
            ToolTipText     =   "��ʾ�ڰ�����Ժʱ�Ƿ�ͬʱ��ȡԤ����"
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "���䴲λ��"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   41
            Top             =   600
            Width           =   1200
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " ҩƷ������ "
         Height          =   7845
         Index           =   2
         Left            =   3310
         TabIndex        =   186
         Top             =   75
         Width           =   3480
         Begin VB.CheckBox chk 
            Caption         =   "��Һ���������״�ִ�е�ҽ����Ҫ�������"
            Height          =   375
            Index           =   83
            Left            =   225
            TabIndex        =   255
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox chk 
            Caption         =   "ʱ��ҩƷ���ʱȡ�ϴ��ۼ�"
            Height          =   195
            Index           =   73
            Left            =   225
            TabIndex        =   245
            Top             =   5880
            Width           =   2760
         End
         Begin VB.CheckBox chk 
            Caption         =   "ʱ��ҩƷͨ���ֶμӳ����"
            Height          =   180
            Index           =   14
            Left            =   225
            TabIndex        =   244
            Top             =   6600
            Width           =   2775
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ����ʱ��ȷҩƷ����"
            Height          =   195
            Index           =   72
            Left            =   240
            TabIndex        =   243
            Top             =   4440
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ�ƿ�ʱ��ȷҩƷ����"
            Height          =   195
            Index           =   71
            Left            =   225
            TabIndex        =   242
            Top             =   4200
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ�⹺�����Ҫ������Ǹ������ܽ��и������"
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
            Caption         =   "ʱ��ҩƷ��ⰴ��ǰ�ӳ�����"
            Height          =   195
            Index           =   48
            Left            =   225
            TabIndex        =   32
            ToolTipText     =   "ʱ��ҩƷ�⹺���ʱ�ۼۼ��㷽ʽ����ѡ��-���ۿۺ�Ĳɹ��ۼ����ۼ�;ѡ�񣭰��ۿ�ǰ�Ĳɹ��ۼ����ۼۡ�"
            Top             =   6840
            Width           =   3090
         End
         Begin VB.CheckBox chk 
            Caption         =   "��дҩƷ�����൥�ݼ���ҩƷ����ⷿ���ÿ��"
            Height          =   375
            Index           =   40
            Left            =   225
            TabIndex        =   27
            Top             =   3360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����������շѻ���ʺ��Զ�����"
            Height          =   195
            Index           =   38
            Left            =   225
            TabIndex        =   33
            ToolTipText     =   "Ӱ��ķ�Χ:�����շ�,����,ҽ������վ����(����[�շѵ���])"
            Top             =   7200
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����շ���ҩ����ҩ����"
            Height          =   240
            Index           =   22
            Left            =   225
            TabIndex        =   23
            Top             =   2160
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "סԺ������ҩ����ҩ����"
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
            Caption         =   "������סԺ���ʺ��Զ�����"
            Height          =   195
            Index           =   37
            Left            =   225
            TabIndex        =   34
            ToolTipText     =   "Ӱ�췶Χ:סԺ����,ҽ������վ(����:���ʵ���)"
            Top             =   7440
            Width           =   3000
         End
         Begin VB.CheckBox chk 
            Caption         =   "ָ��ҩ��ʱ�޶�ҩƷ���"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   25
            Top             =   2745
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ�շ���ɺ��Զ���ҩ"
            Height          =   195
            Index           =   17
            Left            =   225
            TabIndex        =   26
            Top             =   3000
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ʱ��ҩƷͨ���ӳ������"
            Height          =   195
            Index           =   21
            Left            =   225
            TabIndex        =   31
            Top             =   6360
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ҩƷ����ʱ��ȷҩƷ����"
            Height          =   195
            Index           =   26
            Left            =   225
            TabIndex        =   28
            Top             =   3960
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "�⹺��ⵥ��Ҫ�����˲�"
            Height          =   195
            Index           =   28
            Left            =   225
            TabIndex        =   29
            Top             =   4800
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "ʱ��ҩƷֱ��ȷ���ۼ�"
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
            Caption         =   "ҩƷ���������㷨"
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
            Caption         =   "ҩƷЧ����ʾ��ʽ"
            Height          =   180
            Index           =   31
            Left            =   120
            TabIndex        =   216
            Top             =   932
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҩ�۱༭���õ�λ"
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
            Caption         =   "ҩƷ�������ģʽ"
            Height          =   180
            Index           =   32
            Left            =   120
            TabIndex        =   187
            Top             =   606
            Width           =   1440
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " ��ҩ���ڶ�̬���� "
         Height          =   765
         Index           =   3
         Left            =   6840
         TabIndex        =   147
         Top             =   3675
         Width           =   2775
         Begin VB.OptionButton opt 
            Caption         =   "ƽ����ʽ"
            Height          =   210
            Index           =   3
            Left            =   1425
            TabIndex        =   47
            Top             =   360
            Width           =   1020
         End
         Begin VB.OptionButton opt 
            Caption         =   "��æ��ʽ"
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
         Caption         =   " �����շ�ʱ���� "
         Height          =   975
         Index           =   6
         Left            =   6840
         TabIndex        =   141
         Top             =   2475
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "��������"
            Height          =   210
            Index           =   7
            Left            =   330
            TabIndex        =   42
            Top             =   315
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "�Һŵ���"
            Height          =   210
            Index           =   10
            Left            =   1470
            TabIndex        =   45
            Top             =   585
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "���˱�ʶ"
            Height          =   225
            Index           =   8
            Left            =   330
            TabIndex        =   44
            Top             =   570
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chk 
            Caption         =   "ˢ���￨"
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
         Caption         =   " �������°�ʱ�� "
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
            Caption         =   "����"
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
            Caption         =   "����"
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
         Caption         =   " �Һš����סԺ "
         Height          =   7845
         Index           =   4
         Left            =   165
         TabIndex        =   133
         Top             =   75
         Width           =   3105
         Begin VB.CheckBox chk 
            Caption         =   "����������Ч�����Ĳ���"
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
            Caption         =   "�����˷���������"
            Height          =   195
            Index           =   16
            Left            =   165
            TabIndex        =   17
            Top             =   5640
            Width           =   1920
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ÿ��סԺʹ���µ�סԺ��"
            Height          =   195
            Index           =   57
            Left            =   285
            TabIndex        =   208
            Top             =   700
            Width           =   2640
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������Ŀ��λ��������"
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
            Caption         =   "����ȫ������ʱֻ���ұ���"
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
            ToolTipText     =   "�����������˫,��:���м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ"
            Top             =   2310
            Width           =   1755
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   13
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "�����������˫,��:���м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ"
            Top             =   1995
            Width           =   1755
         End
         Begin VB.CheckBox chk 
            Caption         =   "������Ŀ���ܼ����ۿ۶�"
            Height          =   195
            Index           =   39
            Left            =   165
            TabIndex        =   16
            Top             =   5280
            Width           =   2280
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������Ŀʱ�������"
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
            ToolTipText     =   "�Һ��ݲ�֧��������������˫"
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
            Caption         =   "����Һŵ���Ч������"
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
            Caption         =   "���õ��۱���λ��"
            Height          =   180
            Index           =   35
            Left            =   195
            TabIndex        =   235
            Top             =   4755
            Width           =   1440
         End
         Begin VB.Label Label9 
            Caption         =   "�㳮�������"
            Height          =   255
            Left            =   195
            TabIndex        =   200
            Top             =   1470
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շ���Ŀ��������Ŀ����ƥ�䷽ʽ"
            Height          =   180
            Index           =   40
            Left            =   165
            TabIndex        =   199
            ToolTipText     =   "Ӱ���շ�,���ʵ��շ���Ŀ����,ҽ��,��ʿ����ҽ��"
            Top             =   6840
            Width           =   2700
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   38
            Left            =   675
            TabIndex        =   192
            Top             =   2355
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�շ�"
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
            Caption         =   "�Һ�����ԤԼ����"
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
            Caption         =   "���ý���λ��"
            Height          =   180
            Index           =   28
            Left            =   195
            TabIndex        =   139
            Top             =   4395
            Width           =   1440
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����Ź���"
            Height          =   180
            Index           =   22
            Left            =   195
            TabIndex        =   135
            Top             =   1140
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ�Ź���"
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
            Caption         =   "��ͨ�Һŵ���Ч������"
            Height          =   180
            Index           =   16
            Left            =   195
            TabIndex        =   137
            Top             =   2955
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�Һ�"
            Height          =   180
            Index           =   15
            Left            =   675
            TabIndex        =   136
            Top             =   1740
            Width           =   360
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " �ض��շ���Ŀ "
         Height          =   1850
         Index           =   7
         Left            =   6840
         TabIndex        =   148
         Top             =   4755
         Width           =   2775
         Begin VB.CommandButton cmdSelect 
            Caption         =   "��"
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
            Caption         =   "��"
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
            Caption         =   "��"
            Height          =   240
            Index           =   0
            Left            =   2415
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   310
            Width           =   255
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "��"
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
            Caption         =   "�������÷�"
            Height          =   225
            Index           =   56
            Left            =   45
            TabIndex        =   319
            Top             =   1365
            Width           =   1050
         End
         Begin VB.Label lbl 
            Caption         =   "��ͨ���÷�"
            Height          =   225
            Index           =   18
            Left            =   45
            TabIndex        =   316
            Top             =   1005
            Width           =   1050
         End
         Begin VB.Label lbl 
            Caption         =   "������"
            Height          =   225
            Index           =   6
            Left            =   285
            TabIndex        =   149
            Top             =   318
            Width           =   585
         End
         Begin VB.Label lbl 
            Caption         =   "������"
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
         Caption         =   "ɾ����������(&D)"
         Height          =   350
         Left            =   7920
         TabIndex        =   92
         Top             =   7635
         Width           =   1710
      End
      Begin VB.CommandButton cmdWarnNew 
         Caption         =   "���ӱ�������(&A)"
         Height          =   350
         Left            =   7920
         TabIndex        =   91
         Top             =   7275
         Width           =   1710
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ʱ����������۷���"
         Height          =   255
         Index           =   41
         Left            =   7515
         TabIndex        =   90
         ToolTipText     =   "�ڼ��ʱ����ж�ʱ,���˷����ۼƻ��շ����ۼ����Ƿ����δ��˵Ļ��۵�����"
         Top             =   90
         Width           =   2175
      End
      Begin VB.ListBox lst��� 
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
      Begin MSComctlLib.TabStrip tab���� 
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
               Caption         =   "��ͨ����"
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
         Caption         =   "����������ÿ�ַ������������������߼�������ʽ����� zl_PatiWarnScheme �������ʹ��"
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
      Begin VB.OptionButton opt���� 
         Caption         =   "�Լ۸���ߵĻ���ȼ�Ϊ��׼"
         Height          =   255
         Index           =   1
         Left            =   5385
         TabIndex        =   86
         Top             =   7695
         Width           =   2670
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "�����һ�λ���ȼ�Ϊ��׼"
         Height          =   255
         Index           =   0
         Left            =   2850
         TabIndex        =   85
         Top             =   7695
         Value           =   -1  'True
         Width           =   2625
      End
      Begin VB.CheckBox chk 
         Caption         =   "���������ģʽ (ָ�԰���Ϊ���㵥λ,������Ժ��1��,���������,�����Ժ���첻�����,���������)"
         Height          =   225
         Index           =   43
         Left            =   210
         TabIndex        =   83
         ToolTipText     =   "��ʾ�Ƿ��Զ��޸���һ�����ڼ���Զ����ü�������"
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
         Caption         =   "���������Զ��Ʒ�(��ʾ�Ƿ��Զ��޸���һ�����ڼ���Զ����ü������ݡ�)"
         Height          =   285
         Index           =   12
         Left            =   210
         TabIndex        =   81
         ToolTipText     =   "��ʾ�Ƿ��Զ��޸���һ�����ڼ���Զ����ü�������"
         Top             =   7185
         Width           =   6510
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "ͬ�첻ͬ����ȼ��Ļ���Ѽ���"
         Height          =   180
         Left            =   210
         TabIndex        =   84
         Top             =   7770
         Width           =   2520
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ָ�����ý����Զ�����"
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
         Caption         =   "�Դ�λ�ѻ���ѽ����Զ�����"
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
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "  ¼�뾫��"
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
               Name            =   "����"
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
            Caption         =   "��ҩƷ�����ĵİ�װ��λ�����ü۸���������¼��ľ��ȣ�������С��λ����"
            BeginProperty Font 
               Name            =   "����"
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
      Begin ZL9BillEdit.BillEdit Billҩ����ҩ���� 
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
         Caption         =   "ҩ����ҩ����"
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
         Caption         =   "ע�⣺���ұ�ſ�ѡ��ΧA-Z��1-9��ͬ���п��ұ�Ų����ظ���"
         Height          =   285
         Left            =   195
         TabIndex        =   190
         Top             =   7680
         Width           =   5040
      End
      Begin VB.Label Label2 
         Caption         =   "����д���Ŀ��Ҷ�Ӧ�ı��"
         Height          =   285
         Left            =   4680
         TabIndex        =   189
         Top             =   180
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "����дҩƷ���Ҷ�Ӧ�ı��"
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
            Text            =   "��������"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�������"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "��ѡ�񵥾ݶ�Ӧ�ı��뷽ʽ(˫���п����޸ı��뷽ʽ)"
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
         Caption         =   "����ҩƷ�����ÿⷿ"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf�ⷿ��λ 
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
         FormatString    =   "ҩƷ�ⷿ|�ۼ۵�λ|���ﵥλ|סԺ��λ|ҩ�ⵥλ"
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
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��������"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����鷽ʽ"
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
      Begin VB.Label lbl��ʾ 
         Caption         =   "    ���������ѡ����ⷿ�Ƿ����漰����鷽ʽ�����ⷿѡ��ʱ˫���򰴡�C�����ɸı�ⷿ�ļ�鷽ʽ��"
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
         Caption         =   "����ҩƷ�ڲ�ͬ�ⷿ�����ͨ����"
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
         ToolTipText     =   "���ڶ�����ĵ��ʷ��ý����м�飬���������õĽ��ʱ�ͽ������ѣ��Է�ֹ�������"
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
               Key             =   "����"
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
         Caption         =   "���(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   3
         Left            =   8175
         TabIndex        =   118
         Top             =   1965
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "ɾ��(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   2
         Left            =   8175
         TabIndex        =   117
         Top             =   1485
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�޸�(&M)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   1
         Left            =   8175
         TabIndex        =   116
         Top             =   1005
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&A)"
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
            Text            =   "������"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "��ʷ����"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����������˵���"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "�������"
            Object.Width           =   2187
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʷ���������ѽ�"
         Height          =   180
         Left            =   270
         TabIndex        =   119
         Top             =   7725
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ա�Բ�ͬ���ݵĲ���Ȩ�ޣ���Ե��ݵ���ʷ��������������˽�������"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   171
         Top             =   225
         Width           =   6120
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   129
      Top             =   8955
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8430
      TabIndex        =   130
      Top             =   8955
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
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
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ٴ�Ӧ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ǩ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ʊ�ݺͿ�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�Զ�����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ʱ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ȩ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ݲ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ⷿ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ�ⷿ��λ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ��������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ݱ������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab14 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ұ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab15 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩ����ҩ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab16 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ���ľ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab17 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ݻ��ڿ���"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Caption         =   "����ҩƷ���ĵ������ض�ҵ�񻷽��������޸ĵ���Ŀ"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ת�ƻ��Ժ"
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
            Caption         =   "(��Ժ)���ڻ�������"
            Height          =   180
            Index           =   8
            Left            =   210
            TabIndex        =   344
            Top             =   690
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(ת��)δ�����ʵ���"
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
      Begin VB.Frame fra����¼�� 
         Caption         =   "����¼������"
         Height          =   1100
         Left            =   4725
         TabIndex        =   236
         Top             =   6870
         Width           =   4890
         Begin VB.CheckBox chk 
            Caption         =   "ת��������ֻ����¼����"
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
         Begin VB.Label lbl����¼�� 
            AutoSize        =   -1  'True
            Caption         =   "��¼ʱ��(0-9999)"
            Height          =   180
            Left            =   240
            TabIndex        =   239
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label lblInputHours 
            AutoSize        =   -1  'True
            Caption         =   "Сʱ"
            Height          =   180
            Left            =   2520
            TabIndex        =   237
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame fra��Ժ��鸱 
         Caption         =   "����ת�ƻ��Ժ(δִ��������Ŀ)"
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
            ToolTipText     =   "�ڲ��˽����Լ�������������г�Ժʱ���"
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
            ToolTipText     =   "�ڲ������������ת��ʱ���"
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
            Caption         =   "���������δִ��������Ŀ��"
            Height          =   255
            Left            =   240
            TabIndex        =   342
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ժʱ"
            Height          =   180
            Index           =   50
            Left            =   255
            TabIndex        =   234
            Top             =   705
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ת��ʱ"
            Height          =   180
            Index           =   17
            Left            =   255
            TabIndex        =   233
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����һ��ͨ"
         Height          =   2240
         Left            =   4710
         TabIndex        =   214
         Top             =   75
         Width           =   4875
         Begin VB.CheckBox chk 
            Caption         =   "��Ŀ�����������շѻ�������"
            Height          =   210
            Index           =   90
            Left            =   150
            TabIndex        =   337
            Top             =   830
            Width           =   3120
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������Ѽ���ʣ����ʱ��Ҫˢ��������֤"
            Height          =   210
            Index           =   59
            Left            =   150
            TabIndex        =   336
            Top             =   270
            Width           =   4080
         End
         Begin VB.CheckBox chk 
            Caption         =   "��Ŀִ��ǰ�������շѻ��ȼ������"
            Height          =   210
            Index           =   67
            Left            =   150
            TabIndex        =   335
            Top             =   550
            Width           =   4080
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ�շѵ����ﻮ�۴�����ҩ"
            Height          =   195
            Index           =   58
            Left            =   150
            TabIndex        =   334
            Top             =   1110
            Width           =   2880
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ��˵ļ��ʴ�����ҩ"
            Height          =   195
            Index           =   15
            Left            =   150
            TabIndex        =   333
            Top             =   1375
            Value           =   1  'Checked
            Width           =   4425
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ�շѵ����ﻮ�۴�������"
            Height          =   180
            Index           =   68
            Left            =   150
            TabIndex        =   332
            Top             =   1640
            Width           =   2895
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ��˵ļ��˴�������"
            Height          =   255
            Index           =   69
            Left            =   150
            TabIndex        =   331
            Top             =   1890
            Width           =   4455
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ʱ����ȷ������ȼ�"
         Height          =   180
         Index           =   42
         Left            =   375
         TabIndex        =   102
         ToolTipText     =   "һ�����ҿ��Դ����ڶ������,������Ժ�����䴲λʱ��ȷ������,��ҪӰ�첡����Ϣ����,��Ժ����,��ƹ����ģ��"
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
      Begin VB.Frame fra��Ժ��� 
         Caption         =   "����ת�ƻ��Ժ(δ��ҩƷ)"
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
            ToolTipText     =   "�ڲ������������ת��ʱ���"
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
            ToolTipText     =   "�ڲ��˽����Լ�������������г�Ժʱ���"
            Top             =   660
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ת��ʱ"
            Height          =   180
            Index           =   48
            Left            =   270
            TabIndex        =   226
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ժʱ"
            Height          =   180
            Index           =   46
            Left            =   255
            TabIndex        =   227
            Top             =   720
            Width           =   540
         End
      End
      Begin VB.Frame Fraҩ����ͨ 
         Caption         =   "ҩ�ⵥ�����"
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
            Caption         =   "�������������"
            Height          =   180
            Index           =   26
            Left            =   120
            TabIndex        =   167
            Top             =   330
            Width           =   1260
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "סԺ����"
         Height          =   960
         Left            =   165
         TabIndex        =   163
         Top             =   1725
         Width           =   4305
         Begin VB.CheckBox chk 
            Caption         =   "��Ժ���˲������Ժ����"
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
            Caption         =   "δ�󵥾ݽ���"
            Height          =   180
            Index           =   24
            Left            =   225
            TabIndex        =   164
            Top             =   330
            Width           =   1080
         End
      End
      Begin VB.Frame frmסԺ���� 
         Caption         =   "סԺ����"
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
            Caption         =   "����Աֻ���Ա�����ݵǼ�"
            Height          =   210
            Index           =   20
            Left            =   165
            TabIndex        =   94
            Top             =   570
            Width           =   2520
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������������ҵĿ�����"
            Height          =   210
            Index           =   19
            Left            =   165
            TabIndex        =   93
            Top             =   285
            Width           =   2460
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������뿪����"
            Height          =   210
            Index           =   18
            Left            =   2610
            TabIndex        =   96
            Top             =   855
            Width           =   1590
         End
         Begin VB.CheckBox chk 
            Caption         =   "����δ��ƽ�ֹ���˲���"
            Height          =   210
            Index           =   84
            Left            =   165
            TabIndex        =   95
            Top             =   855
            Width           =   2340
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ѽᵥ�ݲ���"
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
         Caption         =   "������˷�ʽ"
         Height          =   180
         Index           =   52
         Left            =   4725
         TabIndex        =   249
         Top             =   2885
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��������"
         Height          =   180
         Index           =   42
         Left            =   4725
         TabIndex        =   209
         Top             =   2495
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���Ѳ������÷�������"
         Height          =   180
         Index           =   21
         Left            =   7395
         TabIndex        =   169
         Top             =   4140
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽ���������÷�������"
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

Private Enum const��
    ud_�Һ�ԤԼ���� = 0
    ud_�Һŵ� = 1
    'ud_�շ��վ� = 2:56963
    ud_���ﴦ���������� = 3
    ud_���볤�� = 4
    ud_���ý���λ�� = 5
    ud_���õ��۱���λ�� = 6
    ud_�����Ǽ���Ч���� = 7
    ud_��¼ҽ��ʶ���� = 8
    ud_��ͯ����綨���� = 9
    ud_����Һŵ� = 10
    ud_�����¿�ҽ����� = 11
End Enum

Private Enum constChk
    chk_δ����������ֹ��ҩ = 0
    'chk_���չ����� = 1:56963
    chk_�޶�ҩƷ�Ŀ�� = 2
    chk_ҩƷ�������ҽ�� = 3
    chk_��ȡԤ���� = 4
    chk_ʱ������￨ = 5
    chk_���䴲λ�� = 6
    chk_�������� = 7
    chk_����ID = 8
    chk_ˢ���￨ = 9
    chk_�Һŵ��� = 10
    chk_�����Ǽ���Ч���� = 11
    chk_�Զ����� = 12
    chk_Ʊ�ſ��� = 13
    'chk_������ʾ = 14
    chk_ʱ�۷ֶμӳ���� = 14
    chk_δ��˼��ʴ�����ҩ = 15
    chk_�����˷��������� = 16
    chk_δ�շѴ�����ҩ = 58
    
    chk_�շ�ͬʱ��ҩ = 17
    chk_���뿪���� = 18
    chk_���ƿ����� = 19
    chk_����ִ�еǼ� = 20
    chk_ʱ��ҩƷ��� = 21
    chk_�����շ��뷢ҩ���� = 22
    chk_סԺ�����뷢ҩ���� = 23
    chk_����ҽ��������Ч = 24
    chk_���������շ���� = 25
    chk_��ȷ����ҩƷ���� = 26
    chk_�������� = 27
    chk_�⹺�����Ҫ�˲� = 28
    chk_�⹺�����Ҫ������Ǹ������ܽ��и��� = 70
    chk_ҩƷ�ƿ���ȷ���� = 71
    chk_ҩƷ������ȷ���� = 72
    'chk_���ŵ����շѷֱ��ӡ = 29:56963
    chk_ȫ����ֻ����� = 30
    chk_ȫ��ĸֻ����� = 31
    chk_ִ�к��Զ���˻��۵� = 32
    chk_һ��������������Ŀ = 34
    'chk_����ʹ��Ʊ�� = 35  :56963
    chk_ʱ��ҩƷֱ��ȷ���ۼ� = 36
    chk_סԺ�����Զ����� = 37
    chk_���������Զ����� = 38
    chk_ִ��֮���Զ����� = 61
    chk_ָ��ҽ������������ִ�� = 62
    
    chk_������Ŀ���ܼ����ۿ� = 39
    
    chk_ҩƷ�ʱ�¿��ÿ�� = 40
    chk_���ʱ����������۷��� = 41
    chk_���ȷ������ȼ� = 42
    chk_���������ģʽ = 43
    
    chk_����ǩ������_���� = 44
    chk_����ǩ������_סԺ = 45
    chk_����ǩ������_ҽ�� = 46
    chk_����ǩ������_���� = 47
    chk_����ǩ������_ҩƷ = 60
    chk_����ǩ������_lis = 1
    chk_����ǩ������_pacs = 29
    
    chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ����� = 48
    'chk_��ִ�п��ҷֱ��ӡ = 49:56963
    chk_�´��Ժҽ���������Ժ = 50
    chk_���ﴦ���������� = 52
    'chk_�շ�ÿ��ֻ��һ��Ʊ�� = 53:56963
    chk_����ҽ���������������� = 54
    chk_��Ժ���˲�׼��Ժ���� = 55
    chk_�շ���Ŀ��λ�������� = 56
    chk_ÿ��סԺʹ����סԺ�� = 57
    chk_���ﲡ������ʱ��Ҫˢ����֤ = 59
    'chk_���￨�ظ�ʹ�� = 63
    chk_סԺҩ�����Ͳ�����ҩ�� = 64
    chk_����ҩ�� = 65
    chk_�´�ҽ��ʱ��ʾ���� = 66
    chk_��Ŀִ��ǰ�����շѻ���� = 67
    chk_��Ŀ�����������շѻ������� = 90 '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
    chk_����δ�շѵ����ﻮ�۴������� = 68
    chk_����δ��˵ļ��˴������� = 69
    chk_��ֹ�´ﳬ����ҩƷҽ�� = 63
    chk_ʱ��ҩƷȡ�ϴ��ۼ� = 73
    chk_��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶� = 74
    chk����ҩ��ּ����� = 75
    chk����ҩ��ʹ���Ա�ҩ = 76
    chk�����´�Ժ��ִ�еĽ���ҩƷҽ�� = 77
    chkֻ����¼���� = 78
    chk�ٴ�����վ����ʹ��zlPlugIn���� = 79
    chk���������ּ����� = 80
    chk_���˳�Ժҽ������������Ժ = 81
    chk_���������Һ���Ч�����Ĳ��� = 82
    chk_�״�ҽ��ִ����Ҫ��� = 83
    chk_δ��ƽ�ֹ���� = 84    '51612
    chk_��Ѫ�ּ����� = 35
    chk_������Ȩ���� = 49
    chk_��Ѫ����������� = 53
    chk_��Ѫ����ֻ�����м�������ҽʦ��� = 85
    chk_ҽ��ִ����Ч���� = 87
    chk_���ýӿڵ�����־ = 88   '��ͨ�ӿ���־���� 65522
    chk_����ʹ��ϵͳ���� = 89   '�����ӿ�ϵͳ���ù��ܿ��Ʋ��� 65198
    chk_ҽ������ʱ��������ԭ�� = 86
    chk_�¿�ҽ��ǩ��ʱһ��ҽ��ǩ��һ�� = 91
End Enum

Private Enum const����
    dtp_�����ϰ� = 0
    dtp_�����°� = 1
    dtp_�����ϰ� = 2
    dtp_�����°� = 3
End Enum

Private Enum constSign
    sst_���� = 0
    sst_סԺҽ�� = 1
    sst_סԺ��ʿ = 2
    sst_ҽ�� = 3
    sst_���� = 4
    sst_ҩƷ = 5
    sst_lis = 6
    sst_Pacs = 7
End Enum

Private Enum constDeptCol
    col_ѡ�� = 0
    col_վ�� = 1
    col_���� = 2
    col_���� = 3
    col_���� = 4
End Enum

Private Enum constBill
    bill_�Զ����� = 0
    bill_���ʱ��� = 1
    bill_ҩƷ���� = 3
    bill_ҩƷ�������� = 4
End Enum

Private Enum constCmb
    cmb_�ѽᵥ�� = 0
    cmb_���������Դ = 1
    cmb_סԺ�Ź��� = 2
    cmb_���۵�λ = 3
    cmb_����Ź��� = 4
    cmb_δ�󵥾ݽ��� = 5
    cmb_��Ժʱδִ����Ŀ��� = 6
    cmb_ҩƷ������� = 7
    cmb_����������� = 8
    cmb_ҩƷ����ģʽ = 9
    cmb_���Ʊ���ģʽ = 10
    cmb_����ǩ����֤���� = 11
    cmb_�Һ���Ǯ���� = 12
    cmb_�շ���Ǯ���� = 13
    cmb_������Ǯ���� = 14
    cmb_ҽ�������� = 15
    cmb_סԺ������� = 16
    cmb_Ч����ʾ��ʽ = 17
    cmb_ҩƷ���������㷨 = 18
    cmb_ת��ʱδִ����Ŀ��� = 19
    cmb_������ҩ�ӿ� = 20
    cmb_�������� = 21
    cmb_��Ժʱδ��ҩ��Ŀ��� = 22
    cmb_ת��ʱδ��ҩ��Ŀ��� = 23
    cmd_��ҩ�䷽ = 26
    cmd_����������Դ = 27
    cmd_ת��ʱδ������ʵ��� = 28
    cmd_��Ժʱ���ڻ������� = 29
End Enum

'��ӦlblINFO
Private Enum lblEnum
    lbl_����������Դ = 0
End Enum

Private Enum constLvw
    lvw_Ʊ�� = 0
    lvw_���� = 1
    lvw_һ��ͨ = 3
End Enum

Private Enum constListBox
    lst_ҽ������ = 0
    lst_���Ѳ��� = 1
    lst_סԺ������� = 2
    lst_���﷢����� = 4    '����Ϊ���۵����������
    lst_ˢ������ = 3
End Enum

Private Enum constOpt
    opt_��æ��ʽ = 2
    opt_ƽ����ʽ = 3
End Enum

Private Enum mGrdCol
    ѡ�� = 0
    ����
    ����
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

'��������
Private mrsWarn As ADODB.Recordset
Private mrs��� As ADODB.Recordset
Private mblnChange As Boolean     '�Ƿ�ı���
Private mblnInit As Boolean       '�Ƿ��ʼ��ʧ��
Private mblnLoad As Boolean
Private mintColumn As Integer '
Private mDecimal As Integer       '�жϷ��ý���С��λ�Ƿ�ı�
Private pDecimal As Integer       '�жϷ��õ��۱���С��λ�Ƿ�ı�
Private mlngFindItem As Long

Private mrsAdvice As New ADODB.Recordset '��¼ҽ�����ݶ���
Private mblnJRaiseByDate As Boolean     '�жϴ�λ����Ŀ��������Ŀ�Ƿ��յ���
Private mblnHRaiseByDate As Boolean     '�жϻ�������Ŀ��������Ŀ�Ƿ��յ���
Private mblnMin As Boolean
Private mstrDel���ò��� As String           '��¼���ʱ�����ɾ�������ò�������
Private mcol���� As Collection '������д����˵���Ŀ���

'��¼���༭�Ŀ��ұ�������С��кͱ��ֵ
Private mintLastRow_Drug As Integer          '��
Private mintLastCol_Drug As Integer          '��
Private mstrLastCode_Drug As String          '���

Private mintLastRow_Stuff As Integer          '��
Private mintLastCol_Stuff As Integer          '��
Private mstrLastCode_Stuff As String          '���

'�Զ����������б��浱ǰ�к���
Private mintCurRow As Integer
Private mintCurCol As Integer

''''''ҩƷ���ĵ��ݻ�����Ŀ����
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

Private Function Check�Ƿ���δ��˵�ҩƷ����() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Id From ҩƷ�շ���¼ Where (���� In(6,7,11) Or (���� In(1,2,3,4,12) And ���ϵ��*ʵ������<0)) And ������� Is Null And ROWNUM<2"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check�Ƿ���δ��˵�ҩƷ���� = (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check�Ƿ���δ��˵��⹺��ⵥ() As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Id From ҩƷ�շ���¼ Where ����=1 And ������� Is Null And ROWNUM<2"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    
    Check�Ƿ���δ��˵��⹺��ⵥ = (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsDrugOrStuff(ByVal strID As String) As Boolean
    '�ж��Ƿ�ΪҩƷ���
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select id From �շ�ϸĿ Where ��� In('4','5','6','7') and id=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    
    IsDrugOrStuff = rs.RecordCount > 0
    rs.Close
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Public Function IsRaiseByDate(ByVal strID As String) As Boolean
    '�жϸ��շ���Ŀ�Ƿ��ǰ��յ���
    '����True-�ǰ�������
    '����False-���ǰ������
    'strID='J' -��λ��Ŀ
    'strID='H' -������Ŀ
    'strID=���� -����ָ������Ŀ
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    If strID = "J" Then
        strSQL = "Select ID" & _
              " From �շѼ�Ŀ " & _
              " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And ִ������ <> Trunc(ִ������, 'dd') And " & _
              " �շ�ϸĿid In " & _
              " (Select ID " & _
              " From �շ���ĿĿ¼ " & _
              " Where ��� = [1] " & _
              " Union All " & _
              " Select ����id From �շѴ�����Ŀ Where ����id In (Select ID From �շ���ĿĿ¼ Where ��� = [1])) "
    ElseIf strID = "H" Then
            strSQL = "Select ID" & _
              " From �շѼ�Ŀ " & _
              " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate And ִ������ <> Trunc(ִ������, 'dd') And " & _
              " �շ�ϸĿid In " & _
              " (Select ID " & _
              " From �շ���ĿĿ¼ " & _
              " Where ��� = [1] " & _
              " Union All " & _
              " Select ����id From �շѴ�����Ŀ Where ����id In (Select ID From �շ���ĿĿ¼ Where ��� = [1])) "
    ElseIf Val(strID) <> 0 Then
        strSQL = "Select Id" & _
                " From �շѼ�Ŀ " & _
                " Where Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
                " And ִ������<>trunc(ִ������,'dd') And (�շ�ϸĿid = [2] or �շ�ϸĿid in (Select ����id From �շѴ�����Ŀ Where ����id = [2])) "
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


Private Sub Load���ݻ��ڿ���()
    Dim n As Integer
    Dim rsTmp As ADODB.Recordset
    Dim m As Integer
    Dim intAllItems As Integer
    
    On Error GoTo ErrHandle
    intAllItems = UBound(Split(cst������Ŀ, ",")) + 1
    
    With vsfControlItem
        .Rows = 7
        .Cols = 2 + intAllItems
        .FixedRows = 1
        .FixedCols = 2
        .RowHeightMin = 500
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "����"
                        
        .ColWidth(0) = 950
        .ColWidth(1) = 950
                        
        For n = 0 To UBound(Split(cst������Ŀ, ","))
            .TextMatrix(0, n + 2) = Split(cst������Ŀ, ",")(n)
            .ColWidth(n + 2) = 920
            .ColAlignment(n + 2) = flexAlignCenterCenter
        Next
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        .CellBorderRange 0, 0, 0, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(1, 0) = "ҩƷ�⹺"
        .TextMatrix(2, 0) = "ҩƷ�⹺"
        .TextMatrix(3, 0) = "ҩƷ�⹺"

        .TextMatrix(1, 1) = "�˲�"
        .TextMatrix(2, 1) = "���"
        .TextMatrix(3, 1) = "�������"
        
        .CellBorderRange 3, 0, 3, .Cols - 1, vbBlue, -1, -1, -1, 1, 0, 0
        
        .TextMatrix(4, 0) = "�����⹺"
        .TextMatrix(5, 0) = "�����⹺"
        .TextMatrix(6, 0) = "�����⹺"

        .TextMatrix(4, 1) = "�˲�"
        .TextMatrix(5, 1) = "���"
        .TextMatrix(6, 1) = "�������"
        
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .Refresh
        
        gstrSQL = "Select ����,����,���� From ���ݻ��ڿ��� Order By ����, ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���ݻ��ڿ���")
        
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
                            Case ����.�����⹺
                                Select Case rsTmp!����
                                    Case ����.�˲�
                                        .TextMatrix(4, m) = "��"
                                    Case ����.���
                                        .TextMatrix(5, m) = "��"
                                    Case ����.�������
                                        .TextMatrix(6, m) = "��"
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
    gstrSQL = "Select �ɱ���, ���ۼ�, ʵ������,���۽�� From ҩƷ�շ���¼ Where Rownum = 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
    mblnMin = (rs.RecordCount > 0)
    
    intMaxCost = IIF(rs.Fields(0).NumericScale > 4, 4, rs.Fields(0).NumericScale)
    intMaxPrice = IIF(rs.Fields(1).NumericScale > 4, 4, rs.Fields(1).NumericScale)
    intMaxNumber = IIF(rs.Fields(2).NumericScale > 4, 4, rs.Fields(2).NumericScale)
    intMaxMoney = IIF(rs.Fields(3).NumericScale > 4, 4, rs.Fields(3).NumericScale)

    With billDigit(0)
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
        .ColWidth(dig_�������) = 1000
        .ColWidth(dig_��������) = 1000
        .ColWidth(dig_���ȵ�λ) = 1000
        .ColWidth(dig_����) = 1100
        .ColWidth(dig_��С����) = 1000
        .ColWidth(dig_��󾫶�) = 1000
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
            " From ҩƷ���ľ��� Order By ����, ���, ����, ��λ"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������󾫶�")
    
    With billDigit(0)
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
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Saveҩ����ҩ����()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "ZL_ҩ����ҩ����_DELETE"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With Me.Billҩ����ҩ����
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                gstrSQL = "ZL_ҩ����ҩ����_INSERT(" & .RowData(i) & "," & IIF(.TextMatrix(i, 1) = "����", 1, 2) & "," & IIF(.TextMatrix(i, 2) <> "", 1, 0) & "," & IIF(Val(.TextMatrix(i, 3)) = 0, "Null", Val(.TextMatrix(i, 3))) & "," & IIF(.TextMatrix(i, 4) <> "", 1, 0) & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
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
    With billDigit(0)
        For n = 1 To .Rows - 1
            strInput = strInput & "0," & _
                .TextMatrix(n, dig_���) & "," & _
                .TextMatrix(n, dig_����) & "," & _
                .TextMatrix(n, dig_��λ) & "," & _
                .TextMatrix(n, dig_����) & ";"
        Next
    End With
    
    gstrSQL = "ZL_ҩƷ���ľ���_Update('" & strInput & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
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
    
    On Error GoTo ErrHandle
    With vsfControlItem
        For n = 1 To .Rows - 1
            Select Case .TextMatrix(n, 0)
                Case "ҩƷ�⹺"
                    int���� = ����.ҩƷ�⹺
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
                    str���� = str���� & IIF(str���� <> "", ",", "") & .TextMatrix(0, m)
                End If
            Next
            
            If str���� <> "" Then
                strInput = strInput & IIF(strInput <> "", ";", "") & int���� & "," & int���� & "," & str����
            End If
        Next
    End With
    
    gstrSQL = "Zl_���ݻ��ڿ���_Update('" & strInput & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub bill_AfterAddRow(Index As Integer, Row As Long)
    If Index = bill_���ʱ��� Then
        With Bill(Index)
            .TextMatrix(Row, 3) = " "
            .TextMatrix(Row, 4) = " "
            .TextMatrix(Row, 5) = " "
            .TextMatrix(Row, 6) = ""
            .TextMatrix(Row, 7) = ""
        End With
    End If
    
    If Index = bill_�Զ����� Then
        With Bill(Index)
            .TextMatrix(Row, 3) = "0-��������"
            .TextMatrix(Row, 4) = Format(DateAdd("d", 1, zlDatabase.Currentdate), "yyyy-mm-dd")
        End With
    End If
End Sub

Private Sub Bill_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    If Index = bill_���ʱ��� Then
        With Bill(Index)
            If .TextMatrix(Row, 0) <> "" And .TextMatrix(Row, 2) <> "" Then mblnChange = True
        End With
    End If
End Sub

Private Sub Bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    With Bill(Index)
        If Index = bill_�Զ����� Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            End If
        End If
    End With
End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    '��ֹ���뱨�����
    With Bill(Index)
        If Index = bill_���ʱ��� And .Col >= 3 Then
            If .Col = 6 Or .Col = 7 Then
                .TxtEnable = True
            Else
                .TxtEnable = False
            End If
        Else
            .TxtEnable = True
        End If
        
        If Index = bill_���ʱ��� And .Col = 4 Then  '������ʽ2
            If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                .ColData(4) = 5 'ÿ�շ��ò��ܱ༭������ʽ2
            Else
                .ColData(4) = 1
            End If
        End If
        If Index = bill_���ʱ��� Then
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
        If Index = bill_���ʱ��� And .MouseCol >= 3 And .MouseRow > 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub bill_Validate(Index As Integer, Cancel As Boolean)
    Dim lngRow As Long
    
    If Index = bill_���ʱ��� Then
        If Not mblnChange Then Exit Sub
        If MouseInRect(cmdCancel.hwnd) Then Exit Sub
        
        '�����ʱ�������
        If Not Check���ʱ��� Then Cancel = True: Exit Sub
        
        '�ռ����ʱ�������
        With mrsWarn
            .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
            Do While Not .EOF
                .Delete
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        
        With Bill(bill_���ʱ���)
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                    mrsWarn.AddNew
                    mrsWarn!���ò��� = tab����.SelectedItem.Caption
                    
                    If .RowData(lngRow) <> 0 Then
                        mrsWarn!����id = .RowData(lngRow)
                        mrsWarn!������ = Split(.TextMatrix(lngRow, 0), "-")(0)
                        mrsWarn!������ = Split(.TextMatrix(lngRow, 0), "-")(1)
                    End If
                    
                    mrsWarn!�������� = CInt(Left(.TextMatrix(lngRow, 1), 1))
                    mrsWarn!����ֵ = CCur(.TextMatrix(lngRow, 2))
                    
                    mrsWarn!������־1 = Get�����봮(.TextMatrix(lngRow, 3))
                    mrsWarn!������־2 = Get�����봮(.TextMatrix(lngRow, 4))
                    mrsWarn!������־3 = Get�����봮(.TextMatrix(lngRow, 5))
                    
                    mrsWarn!�߿����� = Round(Val(.TextMatrix(lngRow, 6)), 2)
                    mrsWarn!�߿��׼ = Round(Val(.TextMatrix(lngRow, 7)), 2)
                    
                    mrsWarn.Update
                End If
            Next
        End With
    End If
End Sub

Private Sub billDigit_EnterCell(Index As Integer, Row As Long, Col As Long)
    With billDigit(Index)
        If Col = dig_���� Then
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

Private Sub cboPatiVerfy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'Private Sub chk����¼��_Click()
'    If chk����¼��.Value = 1 Then
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
        '���Ϊ�ѱ仯,��Ҫ����
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
                If UCase(.TextMatrix(i, col_����)) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_����) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_����) = strFind Then
                    .Row = i: .ShowCell i, col_����
                    Exit For
                End If
            End If
        Next
        If i < .Rows Then
            mlngFindItem = i + 1
        Else
            If mlngFindItem = 1 Then
                MsgBox "û���ҵ�ƥ��Ĳ��š�", vbInformation, Me.Caption
            Else
                MsgBox "�Ѿ����ҵ����һ�������ˡ�", vbInformation, Me.Caption
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
            If lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_һ��ͨ).SelectedItem
                frmOneCard.mbytInFun = 1
                Call frmOneCard.ShowMe(Me, Mid(.Key, 2), .SubItems(1), .SubItems(2), .SubItems(3), IIF(.SubItems(4) = "����:��׼һ��ͨ", 2, IIF(.SubItems(4) = "����:���漰�ۿ�", 1, 0)))
                Call LoadOneCard
            End With
        Case 2
            If lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_һ��ͨ).SelectedItem
                If MsgBox("��ȷʵҪɾ����" & .SubItems(1) & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    Call frmOneCard.DelOneCardRec(Val(Mid(.Key, 2)))
                    Call LoadOneCard
                End If
            End With
    End Select
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim str���� As String, str��ԱID As String, str���� As String
    Dim lng���� As Long, lng���� As Long, bln�޸����� As Boolean
    Dim dbl������� As Double
    Dim lst As ListItem
    
    
    Select Case Index
        Case 0 '����
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                    MsgBox "���������Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        Case 1 '�޸�
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                str���� = .Text
                str���� = .SubItems(1)
                lng���� = Val(.SubItems(2))
                bln�޸����� = (.SubItems(3) = "��")
                dbl������� = Val(.SubItems(4))
                str��ԱID = .Tag
                lng���� = .ListSubItems(1).Tag
            End With
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If Not lst Is lvw(lvw_����).SelectedItem Then
                    If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                        MsgBox "���θı�Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 'ɾ��
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                If MsgBox("��ȷʵҪɾ����" & .Text & "���ԡ�" & .SubItems(1) & "���Ĳ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                lvw(lvw_����).ListItems.Remove .Index
            End With
        Case 3 '���
            If MsgBox("��ȷʵҪɾ�����еĲ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            lvw(lvw_����).ListItems.Clear
    End Select
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_����).ListItems.Add(, , str����, , "Limit")
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_����).SelectedItem
            lst.Text = str����
        End If
        lst.SubItems(1) = str����
        lst.SubItems(2) = lng����
        lst.SubItems(3) = IIF(bln�޸����� = True, "��", "��")
        lst.SubItems(4) = IIF(Val(dbl�������) = 0, "", Format(Val(dbl�������), "0.00"))
        lst.Tag = str��ԱID
        lst.ListSubItems(1).Tag = lng����
    End If
    mblnChange = True
End Sub

Private Sub cmdSendPriceType_Click(Index As Integer)
    Dim i As Long, j As Long
    
    If SendPriceType.Tab = 0 Then
        j = lst_���﷢�����
    Else
        j = lst_סԺ�������
    End If
    With lst(j)
        For i = 0 To .ListCount - 1
            .Selected(i) = IIF(Index = 0, True, False)
        Next
    End With
End Sub

Private Sub cmdWarnDel_Click()
    If tab����.SelectedItem.Caption = "��ͨ����" Then
        MsgBox """" & tab����.SelectedItem.Caption & """��������������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷʵҪɾ��""" & tab����.SelectedItem.Caption & """����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    With mrsWarn
        .Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
        
        '��¼ɾ�������ò�������
        If InStr(1, mstrDel���ò���, tab����.SelectedItem.Caption) = 0 Then
            mstrDel���ò��� = IIF(mstrDel���ò��� = "", "", mstrDel���ò��� & ";") & tab����.SelectedItem.Caption
        End If
        
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    tab����.Tabs.Remove tab����.SelectedItem.Index
    tab����.Tabs(1).Selected = True
    
    mblnChange = True
End Sub

Private Sub cmdWarnNew_Click()
    Dim strName As String, strCopy As String
    Dim strSchemes As String, i As Integer
    Dim rsCopy As ADODB.Recordset
    
    For i = 1 To tab����.Tabs.Count
        strSchemes = strSchemes & "," & tab����.Tabs(i).Caption
    Next
    
    strName = frmWarnEdit.ShowMe(Me, Mid(strSchemes, 2), strCopy)
    If strName = "" Then Exit Sub
    
    '��������
    Set rsCopy = mrsWarn.Clone
    rsCopy.Filter = "���ò���='" & strCopy & "'"
    Do While Not rsCopy.EOF
        mrsWarn.AddNew
        mrsWarn!���ò��� = strName
        mrsWarn!����id = rsCopy!����id
        mrsWarn!������ = rsCopy!������
        mrsWarn!������ = rsCopy!������
        mrsWarn!�������� = rsCopy!��������
        mrsWarn!����ֵ = rsCopy!����ֵ
        mrsWarn!������־1 = rsCopy!������־1
        mrsWarn!������־2 = rsCopy!������־2
        mrsWarn!������־3 = rsCopy!������־3
        mrsWarn!�߿����� = rsCopy!�߿�����
        mrsWarn!�߿��׼ = rsCopy!�߿��׼
        mrsWarn.Update
        rsCopy.MoveNext
    Loop
    
    tab����.Tabs.Add , , strName
    tab����.Tabs(tab����.Tabs.Count).Selected = True
    
    mblnChange = True
End Sub

Private Sub cmd��������_Click()
    Dim objCommunity As Object
    
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    If lvw����.SelectedItem.SubItems(4) = "" Then
        MsgBox lvw����.SelectedItem.SubItems(1) & "û�����á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�ȱ����������ݣ���Ϊ�ӿڳ�ʼ��Ҫ�ж��Ƿ�����
    If lvw����.Tag <> "" Then
        On Error GoTo errH
        gcnOracle.BeginTrans
        Call Save�����ӿ�
        gcnOracle.CommitTrans
        lvw����.Tag = ""
    End If
    
    '��������
    Err.Clear: On Error Resume Next
    Set objCommunity = CreateObject("zlCommunity.clsCommunity")
    Err.Clear: On Error GoTo 0
    
    '���ù���
    If Not objCommunity Is Nothing Then
        If objCommunity.Initialize(gcnOracle) Then
            Call objCommunity.Setup(Val(Mid(lvw����.SelectedItem.Key, 2)))
        End If
    Else
        MsgBox "���������ӿ�û����ȷ��װ��", vbExclamation, gstrSysName
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
    '���²���ֻ����һ��
    mblnLoad = False
    If mblnInit = False Then Unload Me
    Call tabMain_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lst���.Visible Then
            lst���.Visible = False
            Bill(bill_���ʱ���).SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
'    On Error GoTo ErrHandle
    
    mblnLoad = True
    '���г�ʼ��
    Set mcol���� = New Collection
    vsUnCheckItem.ComboList = "..."
    vsUnWriteDept.ComboList = "..."
    Call InitSystemPara
    Call InitEnv
    Call LoadPara
    
    Call LoadOneCard
    Call Load�����ӿ�
    Call Load���ݲ���
    Call Load����
    Call LoadTable
    Call LoadҩƷ����
    Call Load�ⷿ���
    Call LoadҩƷ��������
    Call Load���ݱ������
    Call InitFace
    Call Load����
    Call LoadҩƷ���ľ���
    Call Load���ݻ��ڿ���
    
    Call CheckExist
    
    '�ָ��п�
    RestoreFlexState msh(0), App.ProductName & "\" & Me.Name
    RestoreFlexState Bill(bill_�Զ�����), App.ProductName & "\" & Me.Name & bill_�Զ�����
    RestoreFlexState Bill(bill_���ʱ���), App.ProductName & "\" & Me.Name & bill_���ʱ���
    RestoreFlexState Bill(bill_ҩƷ����), App.ProductName & "\" & Me.Name & bill_ҩƷ����
    RestoreFlexState Bill(bill_ҩƷ��������), App.ProductName & "\" & Me.Name & bill_ҩƷ��������
    '��ʼ���ɹ�
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
    
    strSQL = "Select Rownum From δ��ҩƷ��¼ Where ���� In (8,9,10) and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "CheckExist")
    
    If Not rsTemp.EOF Then
        Me.chk(chk_�����շ��뷢ҩ����).Enabled = False
        Me.chk(chk_סԺ�����뷢ҩ����).Enabled = False
    Else
        Me.chk(chk_�����շ��뷢ҩ����).Enabled = True
        Me.chk(chk_סԺ�����뷢ҩ����).Enabled = True
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub InitEnv()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strTmp As String
    Dim blnTmp As Boolean
    
    '��ʼ�����ڣ���Щ�ǲ���Ҫ�����ݿ��
    Dim lngIndex As Long
    
    On Error GoTo ErrHandle
    cmb(cmb_ҩƷ����ģʽ).AddItem "˳����"
    cmb(cmb_ҩƷ����ģʽ).AddItem "����+�����+˳����"
    Call zlControl.CboSetWidth(cmb(cmb_ҩƷ����ģʽ).hwnd, cmb(cmb_ҩƷ����ģʽ).Width * 1.2)
    
    cmb(cmb_Ч����ʾ��ʽ).AddItem "0-��ʾʧЧ��"
    cmb(cmb_Ч����ʾ��ʽ).AddItem "1-��ʾ��Ч��"
    Call zlControl.CboSetWidth(cmb(cmb_Ч����ʾ��ʽ).hwnd, cmb(cmb_Ч����ʾ��ʽ).Width * 1.2)
    
    cmb(cmb_ҩƷ���������㷨).AddItem "0-�������Ƚ��ȳ�"
    cmb(cmb_ҩƷ���������㷨).AddItem "1-��Ч������ȳ�"
    Call zlControl.CboSetWidth(cmb(cmb_ҩƷ���������㷨).hwnd, cmb(cmb_ҩƷ���������㷨).Width * 1.2)

    cmb(cmb_���Ʊ���ģʽ).AddItem "˳����"
    cmb(cmb_���Ʊ���ģʽ).AddItem "����+�����+˳����"
    Call zlControl.CboSetWidth(cmb(cmb_���Ʊ���ģʽ).hwnd, cmb(cmb_���Ʊ���ģʽ).Width * 1.2)
    
    cmb(cmb_���������Դ).AddItem "1-��ѡ��������Դ"
    cmb(cmb_���������Դ).AddItem "2-����ϱ�׼����"
    cmb(cmb_���������Դ).AddItem "3-��������������"
    cmb(cmb_���������Դ).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_���������Դ).hwnd, cmb(cmb_���������Դ).Width * 1.2)
    
    cmb(cmb_�����������).AddItem "1-������������"
    cmb(cmb_�����������).AddItem "2-�����ݿ���ȡ����"
    cmb(cmb_�����������).AddItem "3-��ҽ�����˴����ݿ�����"
    cmb(cmb_�����������).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_�����������).hwnd, cmb(cmb_�����������).Width * 1.4)
    cmb(cmb_סԺ�������).AddItem "1-������������"
    cmb(cmb_סԺ�������).AddItem "2-�����ݿ���ȡ����"
    cmb(cmb_סԺ�������).AddItem "3-��ҽ�����˴����ݿ�����"
    cmb(cmb_סԺ�������).ListIndex = 0
    Call zlControl.CboSetWidth(cmb(cmb_סԺ�������).hwnd, cmb(cmb_סԺ�������).Width * 1.4)
    
    cmb(cmb_�ѽᵥ��).AddItem "0-����"
    cmb(cmb_�ѽᵥ��).AddItem "1-��ʾ"
    cmb(cmb_�ѽᵥ��).AddItem "2-��ֹ"
    cmb(cmb_�ѽᵥ��).ListIndex = 0
    
    cmb(cmb_ҽ��������).AddItem "0-�����м��"
    cmb(cmb_ҽ��������).AddItem "1-��鲢����δ������Ŀ"
    cmb(cmb_ҽ��������).AddItem "2-��鲢��ֹδ������Ŀ"
    cmb(cmb_ҽ��������).ListIndex = 1
    zlControl.CboSetWidth cmb(cmb_ҽ��������).hwnd, 2100
    
    cmb(cmb_������ҩ�ӿ�).AddItem "0-δʹ��"
    cmb(cmb_������ҩ�ӿ�).AddItem "1-�Ĵ�����"
    cmb(cmb_������ҩ�ӿ�).AddItem "2-�Ϻ���ͨ"
    cmb(cmb_������ҩ�ӿ�).AddItem "3-����̫Ԫͨ"
    cmb(cmb_������ҩ�ӿ�).ListIndex = 0
    
    cmb(cmd_��ҩ�䷽).AddItem "0-��ζ��ҩ"
    cmb(cmd_��ҩ�䷽).AddItem "1-��ζ��ҩ"
    cmb(cmd_��ҩ�䷽).ListIndex = 0
    
    cmb(cmd_����������Դ).AddItem "0-��ѡ��������Դ"
    cmb(cmd_����������Դ).AddItem "1-��ҩƷĿ¼����"
    cmb(cmd_����������Դ).AddItem "2-������Դ����"
    cmb(cmd_����������Դ).ListIndex = 0
    '------------------------------------------------------------------------------------------------------------------
    '6-�ֱ���������:34519
    strTmp = "0-������|1-�ֱ���������|2-�ֱҲ�����ȡ|3-�ֱ������ȡ|4-�ֱ������������˫|5-�Ǳ��������塢�������|6-�ֱ���������"
    For i = 0 To UBound(Split(strTmp, "|"))
        '�ҺŲ�֧�������������˫,��Һ���ʹ��ҽ���Ľ����������̴���ֱ�,Oracle��û�������������˫����
        If i <> 4 Then cmb(cmb_�Һ���Ǯ����).AddItem Split(strTmp, "|")(i)
        cmb(cmb_�շ���Ǯ����).AddItem Split(strTmp, "|")(i)
        cmb(cmb_������Ǯ����).AddItem Split(strTmp, "|")(i)
    Next
    cmb(cmb_�Һ���Ǯ����).ListIndex = 0
    cmb(cmb_�շ���Ǯ����).ListIndex = 0
    cmb(cmb_������Ǯ����).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_�Һ���Ǯ����).hwnd, 2300
    zlControl.CboSetWidth cmb(cmb_�շ���Ǯ����).hwnd, 2300
    zlControl.CboSetWidth cmb(cmb_������Ǯ����).hwnd, 2300
    
    cmb(cmb_סԺ�Ź���).AddItem "0-˳����"
    cmb(cmb_סԺ�Ź���).AddItem "1-����(YYMM)+˳���(0000)"
    cmb(cmb_סԺ�Ź���).AddItem "2-��(YYYY)+˳���(00000)"
    cmb(cmb_סԺ�Ź���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_סԺ�Ź���).hwnd, 2500
    
    cmb(cmb_����Ź���).AddItem "0-˳����"
    cmb(cmb_����Ź���).AddItem "1-������(YYMMDD)+˳���(0000)"
    cmb(cmb_����Ź���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_����Ź���).hwnd, 3000

    
    cmb(cmb_���۵�λ).AddItem "0-�ۼ۵�λ"
    cmb(cmb_���۵�λ).AddItem "1-ҩ�ⵥλ"
    cmb(cmb_���۵�λ).ListIndex = 0
    
    cmb(cmb_δ�󵥾ݽ���).AddItem "0-�����"
    cmb(cmb_δ�󵥾ݽ���).AddItem "1-��鲢��ʾ"
    cmb(cmb_δ�󵥾ݽ���).AddItem "2-��鲢��ֹ"
    cmb(cmb_δ�󵥾ݽ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_δ�󵥾ݽ���).hwnd, 2000
    
    cmb(cmb_��Ժʱδִ����Ŀ���).AddItem "0-�����"
    cmb(cmb_��Ժʱδִ����Ŀ���).AddItem "1-��鲢��ʾ"
    cmb(cmb_��Ժʱδִ����Ŀ���).AddItem "2-��鲢��ֹ"
    cmb(cmb_��Ժʱδִ����Ŀ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_��Ժʱδִ����Ŀ���).hwnd, 2000
    
    cmb(cmb_ת��ʱδִ����Ŀ���).AddItem "0-�����"
    cmb(cmb_ת��ʱδִ����Ŀ���).AddItem "1-��鲢��ʾ"
    cmb(cmb_ת��ʱδִ����Ŀ���).AddItem "2-��鲢��ֹ"
    cmb(cmb_ת��ʱδִ����Ŀ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_ת��ʱδִ����Ŀ���).hwnd, 2000
    
    cmb(cmb_��Ժʱδ��ҩ��Ŀ���).AddItem "0-�����"
    cmb(cmb_��Ժʱδ��ҩ��Ŀ���).AddItem "1-��鲢��ʾ"
    cmb(cmb_��Ժʱδ��ҩ��Ŀ���).AddItem "2-��鲢��ֹ"
    cmb(cmb_��Ժʱδ��ҩ��Ŀ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_��Ժʱδ��ҩ��Ŀ���).hwnd, 2000
    
    cmb(cmb_ת��ʱδ��ҩ��Ŀ���).AddItem "0-�����"
    cmb(cmb_ת��ʱδ��ҩ��Ŀ���).AddItem "1-��鲢��ʾ"
    cmb(cmb_ת��ʱδ��ҩ��Ŀ���).AddItem "2-��鲢��ֹ"
    cmb(cmb_ת��ʱδ��ҩ��Ŀ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmb_ת��ʱδ��ҩ��Ŀ���).hwnd, 2000
    
    '61429:������,2013-11-11
    cmb(cmd_ת��ʱδ������ʵ���).AddItem "0-�����"
    cmb(cmd_ת��ʱδ������ʵ���).AddItem "1-��鲢��ʾ"
    cmb(cmd_ת��ʱδ������ʵ���).AddItem "2-��鲢��ֹ"
    cmb(cmd_ת��ʱδ������ʵ���).ListIndex = 0
    zlControl.CboSetWidth cmb(cmd_ת��ʱδ������ʵ���).hwnd, 2000
    
    '68953:������,2014-08-12
    cmb(cmd_��Ժʱ���ڻ�������).AddItem "0-�����"
    cmb(cmd_��Ժʱ���ڻ�������).AddItem "1-��鲢��ʾ"
    cmb(cmd_��Ժʱ���ڻ�������).AddItem "2-��鲢��ֹ"
    cmb(cmd_��Ժʱ���ڻ�������).ListIndex = 0
    zlControl.CboSetWidth cmb(cmd_��Ժʱ���ڻ�������).hwnd, 2000
    
    cmb(cmb_ҩƷ�������).AddItem "0-������"
    cmb(cmb_ҩƷ�������).AddItem "1-��ͬ��ֹ"
    cmb(cmb_ҩƷ�������).ListIndex = 0
    
    '������˷�ʽ:49501
    With cboPatiVerfy
        .Clear
        .AddItem "0-δ��˲��������": .ItemData(.NewIndex) = 0: .ListIndex = .NewIndex
        .AddItem "1-���ʱ����������ú�ҽ��": .ItemData(.NewIndex) = 1
    End With
    
    '����ǩ����֤����
    cmb(cmb_����ǩ����֤����).AddItem "��ʹ�õ���ǩ��"
    cmb(cmb_����ǩ����֤����).AddItem "1-����ʡ����֤����֤����"
    cmb(cmb_����ǩ����֤����).AddItem "2-����ʡ����֤����֤����"
    cmb(cmb_����ǩ����֤����).AddItem "3-����������֤����֤����"
    cmb(cmb_����ǩ����֤����).AddItem "4-ɽ��ʡ����֤����֤����"
    cmb(cmb_����ǩ����֤����).AddItem "5-������Ԫ����֤����֤����" '-- ԭ���ƽ� ��������ҽԺ ����֤����֤����
    cmb(cmb_����ǩ����֤����).AddItem "6-��Ͷ��������֤����֤����" '-- ԭ���ƽ� ����ʡҽԺ ����֤����֤����
    cmb(cmb_����ǩ����֤����).AddItem "7-��Ͷ����֤����֤����(����)"     '-- ԭ���ƽ� ׼���ҽԺ ����֤����֤����,11��12�¸ĳ��ð��ŵ���
    'cmb(cmb_����ǩ����֤����).AddItem "9-�㶫����֤����֤����(����)"    ' ��û���ò���
    cmb(cmb_����ǩ����֤����).AddItem "10-��������֤����֤����(����)"
    cmb(cmb_����ǩ����֤����).AddItem "11-��������֤����֤����(�Ĵ�)"
    cmb(cmb_����ǩ����֤����).AddItem "12-��������֤����֤����(����)"    '��ʱ���
    cmb(cmb_����ǩ����֤����).AddItem "13-��������֤����֤����(����)"    '��ʱ���
    cmb(cmb_����ǩ����֤����).AddItem "14-��������֤����֤����(����)"
    cmb(cmb_����ǩ����֤����).AddItem "15-�Ϻ�����֤����֤����(�Ϻ�)"
    cmb(cmb_����ǩ����֤����).AddItem "16-��������֤����֤����(����)"   '--����ҽԺ
    cmb(cmb_����ǩ����֤����).AddItem "17-�½�����֤����֤����(�½�)"   '��ʱ���
    cmb(cmb_����ǩ����֤����).ListIndex = 0
    
    mlngFindItem = 1
    For i = 0 To sstSign.Tabs - 1
        sstSign.TabVisible(i) = False
        If i = sst_���� Then
            strTmp = " And t.������� IN (1,3)  and T.�������� IN ('�ٴ�','����','����')"
        ElseIf i = sst_סԺҽ�� Then
            strTmp = " And t.������� IN (2,3)  and T.�������� IN ('�ٴ�','����','����')"
        ElseIf i = sst_סԺ��ʿ Then
            strTmp = " And t.������� IN (2,3)  and T.��������='����'"
        ElseIf i = sst_ҽ�� Then
            strTmp = " And t.������� <> 0  and T.�������� IN('���','����','����','����','Ӫ��')"
        ElseIf i = sst_���� Then
            strTmp = " And t.������� IN (2,3)  and T.��������='����'"
        ElseIf i = sst_ҩƷ Then
            strTmp = " and T.�������� in('��ҩ��','��ҩ��','��ҩ��')"
        ElseIf i = sst_lis Then
            strTmp = " And t.������� <> 0  and T.��������='����'"
        ElseIf i = sst_Pacs Then
            strTmp = " And t.������� <> 0  and T.��������='���'"
        End If
         '����Ĭ�ϲ���ѡ��
        gstrSQL = "Select Distinct D.ID, d.վ��,D.����, D.����,D.����" & vbNewLine & _
                "From ���ű� D, ��������˵�� T" & vbNewLine & _
                "Where d.Id = t.����id And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & strTmp & vbNewLine & _
                "order by վ��,����"
    
        Call OpenRecordset(rsTmp, Me.Caption)
        With vsDept(i)
            .Rows = 1
            .MergeCells = flexMergeFree
            .MergeCol(col_վ��) = True
            .AllowUserResizing = flexResizeBoth
            .SelectionMode = flexSelectionByRow
            .Editable = flexEDKbdMouse
            .ExplorerBar = flexExSortShowAndMove
            .ColSort(col_ѡ��) = flexSortNone
            .Cell(flexcpPicture, 0, col_ѡ��) = ils16.ListImages("UnCheck").Picture
            .Cell(flexcpPictureAlignment, 0, col_ѡ��) = flexAlignCenterCenter
            blnTmp = False
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID & "")
                .TextMatrix(.Rows - 1, col_վ��) = rsTmp!վ�� & ""
                If rsTmp!վ�� & "" <> "" Then
                    blnTmp = True
                Else
                    .TextMatrix(.Rows - 1, col_վ��) = " "
                End If
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                
                rsTmp.MoveNext
            Loop
            .ColHidden(col_վ��) = Not blnTmp
        End With
    Next
   
    'Ʊ������
    lvw(lvw_Ʊ��).ListItems.Add , "C1", "�շ��վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C2", "Ԥ���վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C3", "�����վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C4", "�Һ��վ�"
    'lvw(lvw_Ʊ��).ListItems.Add , "C5", "���￨"
    
    With lvw(lvw_һ��ͨ)
        .ColumnHeaders(1).Width = 549.9213
        .ColumnHeaders(2).Width = 1200.189
        .ColumnHeaders(3).Width = 975.1182
        .ColumnHeaders(4).Width = 950.7402
        .ColumnHeaders(5).Width = 2204.788
    End With
    
    'ˢ��Ҫ����������ĳ���
    With lst(lst_ˢ������)
        .AddItem "����Һ�"
        .AddItem "���ﻮ��"
        .AddItem "�����շ�"
        .AddItem "�������"
        .AddItem "��Ժ�Ǽ�"
        .AddItem "סԺ����"
        .AddItem "���˽���"
        .AddItem "����Ԥ����"
        .AddItem "���鼼ʦվ"
        .AddItem "Ӱ��ҽ��վ"
        .ListIndex = 0
    End With
    
    msh(0).Cols = 7
    msh(0).TextMatrix(0, 0) = "����"
    msh(0).TextMatrix(0, 1) = "��λ��"
    msh(0).TextMatrix(0, 2) = " ��������"
    msh(0).TextMatrix(0, 3) = "�����"
    msh(0).TextMatrix(0, 4) = " ��������"
    msh(0).TextMatrix(0, 5) = " ��λ��ԭʼ��������"
    msh(0).TextMatrix(0, 6) = " �����ԭʼ��������"
    
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
    
    With Bill(bill_�Զ�����)
        .Cols = 5 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "�շ�ϸĿID"
        .TextMatrix(0, 2) = "�շ���Ŀ"
        .TextMatrix(0, 3) = "���㷽ʽ"
        .TextMatrix(0, 4) = "��������"
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
    
    With Bill(bill_ҩƷ����)
        .Cols = 4 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "���ڿⷿ"
        .TextMatrix(0, 1) = "�Է��ⷿ"
        .TextMatrix(0, 2) = "�Է��ⷿID"
        .TextMatrix(0, 3) = "����"
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
    
    With Bill(bill_ҩƷ��������)
        .Cols = 3 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .TextMatrix(0, 0) = "���ò���"
        .TextMatrix(0, 1) = "���ÿⷿ"
        .TextMatrix(0, 2) = "�ⷿID"
        .ColWidth(0) = 3500
        .ColWidth(1) = 3500
        .ColWidth(2) = 0
        .ColData(0) = 1
        .ColData(1) = 3
        .ColData(2) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    
    lngIndex = bill_���ʱ���
    Bill(lngIndex).Cols = 8
    Bill(lngIndex).ColAlignment(0) = 1 '����
    Bill(lngIndex).ColAlignment(1) = 1 '��������
    Bill(lngIndex).ColAlignment(2) = 7 '����ֵ
    Bill(lngIndex).ColAlignment(3) = 1 '������־1
    Bill(lngIndex).ColAlignment(4) = 1 '������־2
    Bill(lngIndex).ColAlignment(5) = 1 '������־3
    '���˺� ����:34770    ����:2010-12-21 10:52:49
    Bill(lngIndex).ColAlignment(6) = 7 '�߿�����
    Bill(lngIndex).ColAlignment(7) = 7 '�߿��׼
    
    Bill(lngIndex).TextMatrix(0, 0) = "����"
    Bill(lngIndex).TextMatrix(0, 1) = "��������"
    Bill(lngIndex).TextMatrix(0, 2) = "����ֵ"
    Bill(lngIndex).TextMatrix(0, 3) = "������ʽ1"
    Bill(lngIndex).TextMatrix(0, 4) = "������ʽ2"
    Bill(lngIndex).TextMatrix(0, 5) = "������ʽ3"
    Bill(lngIndex).TextMatrix(0, 6) = "�߿�����"
    Bill(lngIndex).TextMatrix(0, 7) = "�߿��׼"
    
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

    '�ⷿ��λ
    msf�ⷿ��λ.AllowUserResizing = flexResizeNone
    msf�ⷿ��λ.FixedRows = 1
    msf�ⷿ��λ.Cols = 5
    msf�ⷿ��λ.MergeCol(0) = True
    msf�ⷿ��λ.FormatString = "ҩƷ�ⷿ|�������|�ۼ۵�λ|���ﵥλ|סԺ��λ|ҩ�ⵥλ"
    msf�ⷿ��λ.ColWidth(1) = 900
    msf�ⷿ��λ.ColWidth(2) = 900
    msf�ⷿ��λ.ColWidth(3) = 900
    msf�ⷿ��λ.ColWidth(4) = 900
    msf�ⷿ��λ.ColWidth(5) = 900
    msf�ⷿ��λ.ColAlignment(1) = 4
    msf�ⷿ��λ.ColAlignment(2) = 4
    msf�ⷿ��λ.ColAlignment(3) = 4
    msf�ⷿ��λ.ColAlignment(4) = 4
    msf�ⷿ��λ.ColAlignment(5) = 4
    msf�ⷿ��λ.ColWidth(0) = msf�ⷿ��λ.Width - 900 * 6 - 27 * Screen.TwipsPerPixelX
    msf�ⷿ��λ.MergeCells = flexMergeFree
    msf�ⷿ��λ.MergeCol(0) = True
    
    
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
    
    '��ȡҽ������Ϊ�������
    gstrSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('5','6','7','8','9')" & _
        " Union All Select '5','ҩƷ' From Dual Order by ����"
    Call OpenRecordset(rsTmp, Me.Caption)
  
    Do While Not rsTmp.EOF
        lst(lst_���﷢�����).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_���﷢�����).ItemData(lst(lst_���﷢�����).NewIndex) = Asc(rsTmp!����)
        
        lst(lst_סԺ�������).AddItem rsTmp!���� & "-" & rsTmp!����
        lst(lst_סԺ�������).ItemData(lst(lst_סԺ�������).NewIndex) = Asc(rsTmp!����)
        
        rsTmp.MoveNext
    Loop
    
    '��ȡҽ�����ݶ���
    gstrSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
    Call OpenRecordset(mrsAdvice, Me.Caption)
    
    '��ȡ���С��������ġ��͡�ҩ�������ԵĲ���
    gstrSQL = "Select Distinct A.ID, A.����" & _
        " From ���ű� A, ��������˵�� B " & _
        " Where A.ID = B.����id And B.�������� = '��������' And " & _
        " B.����id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') " & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) "
    Call OpenRecordset(rsTmp, Me.Caption)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Load�����ӿ�() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    strSQL = "Select ���, ����, ˵��, ����, ������ From ����Ŀ¼ Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw����.ListItems.Add(, "_" & rsTmp!���, rsTmp!���, , "����")
        ObjItem.SubItems(1) = rsTmp!����
        ObjItem.SubItems(2) = Nvl(rsTmp!˵��)
        ObjItem.SubItems(3) = rsTmp!������
        ObjItem.SubItems(4) = IIF(Nvl(rsTmp!����, 0) = 1, "��", "")
        rsTmp.MoveNext
    Loop
    
    If Not lvw����.SelectedItem Is Nothing Then
        Call lvw����_ItemClick(lvw����.SelectedItem)
    End If
    Load�����ӿ� = True
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
    
    lvw(lvw_һ��ͨ).ListItems.Clear
    
    strSQL = "Select ���,����,���㷽ʽ,ҽԺ����,���� From һ��ͨĿ¼ Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw(lvw_һ��ͨ).ListItems.Add(, "_" & rsTmp!���, rsTmp!���)
        ObjItem.SubItems(1) = rsTmp!����
        ObjItem.SubItems(2) = rsTmp!���㷽ʽ
        ObjItem.SubItems(3) = rsTmp!ҽԺ����
        ObjItem.SubItems(4) = IIF(Nvl(rsTmp!����, 0) = 2, "����:��׼һ��ͨ", IIF(Nvl(rsTmp!����, 0) = 1, "����:���漰�ۿ�", "ͣ��"))
        rsTmp.MoveNext
    Loop
    
    If Not lvw(lvw_һ��ͨ).SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw_һ��ͨ, lvw(lvw_һ��ͨ).SelectedItem)
    End If
    LoadOneCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load����()
'��ȡ���ݲ���ʾ����
    Dim lng��� As Long, str�ⷿID As String
    Dim rsTemp As New ADODB.Recordset
    Dim strType As String
    Dim strSequence As String
    
'    StrType = "('��ҩ��','��ҩ��','��ҩ��','�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��','���Ŀ�','���ϲ���')"

    'ҩƷ����
    On Error GoTo ErrHandle
    strType = "('��ҩ��','��ҩ��','��ҩ��','�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')"
    strSequence = "(21,22,23,24,25,26,27,28,29,32,62)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.����,a.����,b.��� " & _
        "   From ���ű� A,���Һ���� b" & _
        "   Where a.id=b.����id and a.ID in (select distinct ����id from ��������˵�� where �������� in " & strType & ")" & _
        "   And b.��Ŀ��� In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.����,a.����,'' As ��� " & _
        "   From ���ű� A " & _
        "   Where a.ID in (select distinct ����id from ��������˵�� " & _
        "   where �������� in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct ����id From ���Һ���� Where ����id Is Not null " & _
        "   And ��Ŀ��� In " & strSequence & ") " & _
        "   ORDER BY ���� "
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��صĿ���"
    
    With rsTemp
        str�ⷿID = ""
        Do While Not .EOF
            'mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.ѡ��) = IIf(Nvl(!ѡ��, 0) = 1, "��", "")
            mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.����) = Nvl(!����)
            mshBillEdit.TextMatrix(mshBillEdit.Rows - 1, mGrdCol.����) = Nvl(!���)
            mshBillEdit.RowData(mshBillEdit.Rows - 1) = !ID
            mshBillEdit.Rows = mshBillEdit.Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTemp!ID
            .MoveNext
        Loop
    End With
    
    If str�ⷿID <> "" Then
        str�ⷿID = Mid(str�ⷿID, 2)
        mshBillEdit.Rows = mshBillEdit.Rows - 1
        mshBillEdit.Active = True
    Else
        mshBillEdit.Active = False
    End If
    
    rsTemp.Close
    
    '���Ŀ���
    strType = "('�Ƽ���','���Ŀ�','����ⷿ')"
    strSequence = "(68,69,70,71,72,73,74,75,76,77)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.����,a.����,b.��� " & _
        "   From ���ű� A,���Һ���� b" & _
        "   Where a.id=b.����id and a.ID in (select distinct ����id from ��������˵�� where �������� in " & strType & ")" & _
        "   And b.��Ŀ��� In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.����,a.����,'' As ��� " & _
        "   From ���ű� A " & _
        "   Where a.ID in (select distinct ����id from ��������˵�� " & _
        "   where �������� in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct ����id From ���Һ���� Where ����id Is Not null " & _
        "   And ��Ŀ��� In " & strSequence & ") " & _
        "   ORDER BY ���� "

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��صĿ���"
    
    With rsTemp
        str�ⷿID = ""
        Do While Not .EOF
            'mshBillEditstuff.TextMatrix(mshBillEditstuff.Rows - 1, mGrdCol.ѡ��) = IIf(Nvl(!ѡ��, 0) = 1, "��", "")
            mshBillEditStuff.TextMatrix(mshBillEditStuff.Rows - 1, mGrdCol.����) = Nvl(!����)
            mshBillEditStuff.TextMatrix(mshBillEditStuff.Rows - 1, mGrdCol.����) = Nvl(!���)
            mshBillEditStuff.RowData(mshBillEditStuff.Rows - 1) = !ID
            mshBillEditStuff.Rows = mshBillEditStuff.Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTemp!ID
            .MoveNext
        Loop
    End With
    
    If str�ⷿID <> "" Then
        str�ⷿID = Mid(str�ⷿID, 2)
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
'ϵͳ������
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer, blnFind As Boolean
    Dim n As Integer

    '���ȶԷ������ͽ��г�ʼ��
    On Error GoTo ErrHandle
    Call Load��������

    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select ������,����ֵ,ȱʡֵ From Zlparameters Where ϵͳ = " & glngSys & " And Nvl(˽��, 0) = 0 And ģ�� Is Null Order By ������"
    Call OpenRecordset(rsTemp, Me.Caption)

    Do Until rsTemp.EOF
        Select Case rsTemp("������")
        Case 1    '�������°�ʱ��
            i = InStr(UCase(rsTemp("����ֵ")), "AND")
            strTemp = Mid(rsTemp("����ֵ"), 1, i - 2)
            dtp(dtp_�����ϰ�).Value = CDate(strTemp)
            strTemp = Mid(rsTemp("����ֵ"), i + 4)
            dtp(dtp_�����°�).Value = CDate(strTemp)
        Case 2    '�������°�ʱ��
            i = InStr(UCase(rsTemp("����ֵ")), "AND")
            strTemp = Mid(rsTemp("����ֵ"), 1, i - 2)
            dtp(dtp_�����ϰ�).Value = CDate(strTemp)
            strTemp = Mid(rsTemp("����ֵ"), i + 4)
            dtp(dtp_�����°�).Value = CDate(strTemp)
            '            Case 3 '�վݼ��չ�����  '56963
            '                chk(chk_���չ�����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            '                Call chk_Click(chk_���չ�����)
            '            Case 4 '�շ��վ����д�
            '                If Not IsNull(rsTemp("����ֵ")) Then
            '                    ud(ud_�շ��վ�).Value = rsTemp("����ֵ")
            '                End If
        Case 5    '��¼ҽ��ʶ����
            ud(ud_��¼ҽ��ʶ����).Value = Nvl(rsTemp!����ֵ, 30)
        Case 6    'δ��˼��ʴ�����ҩ
            chk(chk_δ��˼��ʴ�����ҩ) = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 148    'δ�շѴ�����ҩ
            chk(chk_δ�շѴ�����ҩ) = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 7    '���������Զ��Ʒ�
            chk(chk_�Զ�����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 9    '���ý���λ��
            Me.ud(ud_���ý���λ��).Value = IIF(IsNumeric(zlCommFun.Nvl(rsTemp("����ֵ"), 2)), zlCommFun.Nvl(rsTemp("����ֵ"), 2), 2)
            Me.txtUD(ud_���ý���λ��).Text = Me.ud(ud_���ý���λ��).Value
            mDecimal = Me.txtUD(ud_���ý���λ��).Text
        Case 157    '���õ��۱���λ��
            Me.ud(ud_���õ��۱���λ��).Value = IIF(IsNumeric(zlCommFun.Nvl(rsTemp("����ֵ"), 5)), zlCommFun.Nvl(rsTemp("����ֵ"), 5), 5)
            Me.txtUD(ud_���õ��۱���λ��).Text = Me.ud(ud_���õ��۱���λ��).Value
            pDecimal = Me.txtUD(ud_���õ��۱���λ��).Text

        Case 10    '��Ժʱ��Ԥ����
            chk(chk_��ȡԤ����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 11    '��Ժʱ����￨
            chk(chk_ʱ������￨).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            '            Case 12 '���￨��������ʾ
            '                chk(chk_������ʾ).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 13    '��Ժͬʱ���
            chk(chk_���䴲λ��).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 14    '��Ǯ����
            strTemp = IIF(IsNull(rsTemp("����ֵ")), "000", rsTemp("����ֵ"))
            n = Val(Mid(strTemp, 1, 1))
            For i = 0 To cmb(cmb_�Һ���Ǯ����).ListCount
                If Val(Split(cmb(cmb_�Һ���Ǯ����).List(i) & "-", "-")(0)) = n Then cmb(cmb_�Һ���Ǯ����).ListIndex = i: Exit For
            Next
            cmb(cmb_�շ���Ǯ����).ListIndex = Val(Mid(strTemp, 2, 1))
            cmb(cmb_������Ǯ����).ListIndex = Val(Mid(strTemp, 3, 1))
        Case 15    '�����շ��뷢ҩ����
            chk(chk_�����շ��뷢ҩ����).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 0, 0, 1)
        Case 16    'סԺ�����뷢ҩ����
            chk(chk_סԺ�����뷢ҩ����).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 0, 0, 1)
        Case 17    '�������뷽ʽ���ֱ�Ϊ���������￨���Һŵ�������ID
            strTemp = IIF(IsNull(rsTemp("����ֵ")), "1111", rsTemp("����ֵ"))
            chk(chk_��������).Value = Val(Mid(strTemp, 1, 1))
            chk(chk_ˢ���￨).Value = Val(Mid(strTemp, 2, 1))
            chk(chk_�Һŵ���).Value = Val(Mid(strTemp, 3, 1))
            chk(chk_����ID).Value = Val(Mid(strTemp, 4, 1))
        Case 18    'ָ��ҩ��ʱ���ƿ��
            chk(chk_�޶�ҩƷ�Ŀ��).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 19    '���ڷ��䷽ʽ
            '�����һ���ؼ���Indexֵ��2
            opt(CInt(IIF(IsNull(rsTemp("����ֵ")), "0", rsTemp("����ֵ"))) + 2).Value = True
        Case 20    '��ʾ����Ʊ�ݵĺ��볤�ȣ���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�
            strTemp = IIF(IsNull(rsTemp("����ֵ")), "7|7|7|7", rsTemp("����ֵ"))
            lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) = Split(strTemp, "|")(0)
            lvw(lvw_Ʊ��).ListItems("C2").SubItems(1) = Split(strTemp, "|")(1)
            lvw(lvw_Ʊ��).ListItems("C3").SubItems(1) = Split(strTemp, "|")(2)
            lvw(lvw_Ʊ��).ListItems("C4").SubItems(1) = Split(strTemp, "|")(3)
            'lvw(lvw_Ʊ��).ListItems("C5").SubItems(1) = Split(strTemp, "|")(4)
        Case 21  '�Һ���Ч����
            '��ͨ��
            ud(ud_�Һŵ�).Value = IIF(Left(zlCommFun.Nvl(rsTemp("����ֵ"), 0), 1) = 0, 1, Left(zlCommFun.Nvl(rsTemp("����ֵ"), 0), 1))
            '�����
            ud(ud_����Һŵ�).Value = IIF(Mid(zlCommFun.Nvl(rsTemp("����ֵ"), 0), 2, 1) = 0, 1, Mid(zlCommFun.Nvl(rsTemp("����ֵ"), 0), 2, 1))
        Case 22    '��Ժʱδִ����Ŀ���
            cmb(cmb_��Ժʱδִ����Ŀ���).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 23    '�ѽ��ʵ��ݲ���
            cmb(cmb_�ѽᵥ��).ListIndex = IIF(IsNull(rsTemp("����ֵ")), 0, rsTemp("����ֵ"))
        Case 24    '��ʾ�Ƿ��ϸ���ƹ����Ʊ�ݵ�ʹ�ã���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
            strTemp = IIF(IsNull(rsTemp("����ֵ")), "1111", rsTemp("����ֵ"))
            lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = IIF(Mid(strTemp, 1, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C2").SubItems(2) = IIF(Mid(strTemp, 2, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C3").SubItems(2) = IIF(Mid(strTemp, 3, 1) = "1", "��", "")
            lvw(lvw_Ʊ��).ListItems("C4").SubItems(2) = IIF(Mid(strTemp, 4, 1) = "1", "��", "")
            ' lvw(lvw_Ʊ��).ListItems("C5").SubItems(2) = IIF(Mid(strTemp, 5, 1) = "1", "��", "")
        Case 25    '����ǩ����֤����
            With cmb(cmb_����ǩ����֤����)
                blnFind = False
                For i = 0 To .ListCount - 1
                    If Val(.List(i)) = Val("" & rsTemp!����ֵ) Then
                        .ListIndex = i
                        blnFind = True
                        Exit For
                    End If
                Next
                If .ListCount > 0 And Not blnFind Then .ListIndex = 0
            End With

        Case 185    '������˷�ʽ   ' 49501
            With cboPatiVerfy
                blnFind = False
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = Val("" & rsTemp!����ֵ) Then
                        .ListIndex = i
                        blnFind = True
                        Exit For
                    End If
                Next
                If .ListCount > 0 And Not blnFind Then .ListIndex = 0
            End With
        Case 26    '����ǩ��ʹ�ó���
            chk(chk_����ǩ������_����).Value = Val(Mid(Nvl(rsTemp!����ֵ), 1, 1))
            chk(chk_����ǩ������_סԺ).Value = Val(Mid(Nvl(rsTemp!����ֵ), 2, 1))
            chk(chk_����ǩ������_ҽ��).Value = Val(Mid(Nvl(rsTemp!����ֵ), 3, 1))
            chk(chk_����ǩ������_����).Value = Val(Mid(Nvl(rsTemp!����ֵ), 4, 1))
            chk(chk_����ǩ������_ҩƷ).Value = Val(Mid(Nvl(rsTemp!����ֵ), 5, 1))
            chk(chk_����ǩ������_lis).Value = Val(Mid(Nvl(rsTemp!����ֵ), 6, 1))
            chk(chk_����ǩ������_pacs).Value = Val(Mid(Nvl(rsTemp!����ֵ), 7, 1))
        Case 27    'סԺҩ�����Ͳ�����ҩ��
            chk(chk_סԺҩ�����Ͳ�����ҩ��).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 0, 0, 1)
        Case 28    '���ﲡ������ʱ��Ҫˢ����֤
            chk(chk_���ﲡ������ʱ��Ҫˢ����֤).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 163    '��Ŀִ��ǰ�������շѻ��ȼ������
            chk(chk_��Ŀִ��ǰ�����շѻ����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
        Case 232    '��Ŀ�����������շѻ�������
            chk(chk_��Ŀ�����������շѻ�������).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 29    'ָ�������۶��۵�λ
            cmb(cmb_���۵�λ).ListIndex = IIF(rsTemp("����ֵ") = "1", 1, 0)
        Case 31    '��Ժ���˲�׼��Ժ����
            chk(chk_��Ժ���˲�׼��Ժ����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 32    'ת��ʱδִ����Ŀ���
            cmb(cmb_ת��ʱδִ����Ŀ���).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 33    'ִ��֮�������Զ�����
            chk(chk_ִ��֮���Զ�����).Value = IIF(Val(rsTemp!����ֵ) <> 0, 1, 0)
        Case 34    'ָ��ҽ������������ִ��
            chk(chk_ָ��ҽ������������ִ��).Value = IIF(Val(rsTemp!����ֵ) <> 0, 1, 0)
        Case 41    'ҽ���������÷�������
            SetListByText lst(lst_ҽ������), Replace(IIF(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ")), "|", ",")
        Case 42    '���Ѳ������÷�������
            SetListByText lst(lst_���Ѳ���), Replace(IIF(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ")), "|", ",")
        Case 43    '�´��Ժҽ���������Ժ
            chk(chk_�´��Ժҽ���������Ժ).Value = IIF(Val("" & rsTemp!����ֵ) <> 0, 1, 0)
        Case 44    '�շ���Ŀ��������Ŀ������ƥ�䷽ʽ
            chk(chk_ȫ����ֻ�����).Value = IIF(Mid(IIF(IsNull(rsTemp!����ֵ), "00", rsTemp!����ֵ), 1, 1) = "1", 1, 0)
            chk(chk_ȫ��ĸֻ�����).Value = IIF(Mid(IIF(IsNull(rsTemp!����ֵ), "00", rsTemp!����ֵ), 2, 1) = "1", 1, 0)
        Case 45    '�շ�ͬʱ��ҩ
            chk(chk_�շ�ͬʱ��ҩ).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 46    'ˢ��Ҫ����������
            With lst(lst_ˢ������)
                For i = 1 To Len(Nvl(rsTemp!����ֵ))
                    If Mid(rsTemp!����ֵ, i, 1) = "1" And i - 1 <= .ListCount - 1 Then
                        .Selected(i - 1) = True
                    End If
                Next
            End With
        Case 51    '����ִ�еǼ�
            chk(chk_����ִ�еǼ�).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 52    '�������뿪����
            chk(chk_���뿪����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 53    '�������ƿ�����
            chk(chk_���ƿ�����).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 54    'ʱ��ҩƷ�ԼӼ������
            chk(chk_ʱ��ҩƷ���).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 55    '���������Դ
            cmb(cmb_���������Դ).ListIndex = CLng(zlCommFun.Nvl(rsTemp("����ֵ"), 1)) - 1
        Case 56    '���ﴦ����������
            If zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 0 Then
                ud(ud_���ﴦ����������).Value = 5
                chk(chk_���ﴦ����������).Value = 0
            Else
                ud(ud_���ﴦ����������).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
                chk(chk_���ﴦ����������).Value = 1
            End If
            '            Case 57 '�շ�ÿ��ֻ��һ��Ʊ��   '56963
            '                chk(chk_�շ�ÿ��ֻ��һ��Ʊ��).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
        Case 58    'δ�󵥾ݽ��ʴ���
            cmb(cmb_δ�󵥾ݽ���).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 59    'ҽ��������
            cmb(cmb_ҽ��������).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 60    '���ʷ���������ѽ��
            txtMaxMoney.Text = zlCommFun.Nvl(rsTemp("����ֵ"))
            Call txtMaxMoney_Validate(False)
        Case 61    '���Ʊ������ģʽ
            cmb(cmb_���Ʊ���ģʽ).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 63    'סԺ�����Զ�����
            chk(chk_סԺ�����Զ�����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 64    'ҩƷ������˹���
            cmb(cmb_ҩƷ�������).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 65    '������뷽ʽ
            cmb(cmb_�����������).ListIndex = Val(Mid(Nvl(rsTemp!����ֵ, "11"), 1, 1)) - 1
            cmb(cmb_סԺ�������).ListIndex = Val(Mid(Nvl(rsTemp!����ֵ, "11"), 2, 1)) - 1
        Case 66    '�Һ�ԤԼ����
            ud(ud_�Һ�ԤԼ����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 68    'δ����������ֹ��ҩ
            chk(chk_δ����������ֹ��ҩ).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 69    'ҩƷ�������ҽ��
            chk(chk_ҩƷ�������ҽ��).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 70    '�����Ǽ���Ч����
            If zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 0 Then
                ud(ud_�����Ǽ���Ч����).Value = 1
                chk(chk_�����Ǽ���Ч����).Value = 0
            Else
                ud(ud_�����Ǽ���Ч����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
                chk(chk_�����Ǽ���Ч����).Value = 1
            End If
        Case 71    '����ҽ��������Ч
            chk(chk_����ҽ��������Ч).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 72    '���������շ����
            chk(chk_���������շ����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 1)
        Case 73    '��ȷ����ҩƷ����
            chk(chk_��ȷ����ҩƷ����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 1)
            chk(chk_��ȷ����ҩƷ����).Tag = chk(chk_��ȷ����ҩƷ����).Value
        Case 75    '�⹺�����Ҫ�˲�
            chk(chk_�⹺�����Ҫ�˲�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 76    'chk_ʱ��ҩƷֱ��ȷ���ۼ�
            chk(chk_ʱ��ҩƷֱ��ȷ���ۼ�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
            '            Case 78 '���ŵ����շѷֱ��ӡ 56963
            '                chk(chk_���ŵ����շѷֱ��ӡ).Value = IIF(Val("" & rsTemp("����ֵ")) = 1, 1, 0)
        Case 80    'סԺҽ������Ϊ���۵�
            strTemp = zlCommFun.Nvl(rsTemp("����ֵ"))
            If strTemp <> "" Then
                With lst(lst_סԺ�������)
                    For i = 0 To .ListCount - 1
                        If InStr(strTemp, Chr(.ItemData(i))) > 0 Then
                            .Selected(i) = True
                        End If
                    Next
                    .ListIndex = 0
                End With
            End If
        Case 81    'ִ�к��Զ���˻��۵�
            chk(chk_ִ�к��Զ���˻��۵�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 84    'һ��������������Ŀ
            chk(chk_һ��������������Ŀ).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 86    '����ҽ������Ϊ���۵�
            strTemp = zlCommFun.Nvl(rsTemp("����ֵ"))
            If strTemp <> "" Then
                With lst(lst_���﷢�����)
                    For i = 0 To .ListCount - 1
                        If InStr(strTemp, Chr(.ItemData(i))) > 0 Then
                            .Selected(i) = True
                        End If
                    Next
                    .ListIndex = 0
                End With
            End If
        Case 87    'ҩƷ�������ģʽ
            cmb(cmb_ҩƷ����ģʽ).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
            '            Case 89 '56963
            '                chk(chk_����ʹ��Ʊ��).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 92    '���������Զ�����
            chk(chk_���������Զ�����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 30    '��ҽ������վ��ҩ����ģ��ʹ�ú�����ҩ����
            cmb(cmb_������ҩ�ӿ�).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 93    '�Ƿ����������Ŀ���ܼ����ۿ�
            chk(chk_������Ŀ���ܼ����ۿ�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 96
            chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
            chk(chk_ҩƷ�ʱ�¿��ÿ��).Tag = chk(chk_ҩƷ�ʱ�¿��ÿ��).Value
            '            Case 97 '�շ�Ʊ�����ɷ�ʽ '56963
            '                opt�շ�Ʊ�����ɷ�ʽ(Val("" & rsTemp("����ֵ")) Mod 10).Value = True
            '                chk(chk_��ִ�п��ҷֱ��ӡ).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) >= 10, 1, 0)
        Case 98
            chk(chk_���ʱ����������۷���).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 99
            chk(chk_���ȷ������ȼ�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 100
            chk(chk_���������ģʽ).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 126
            chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
            chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Tag = chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Value
        Case 143
            chk(chk_����ҽ����������������).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 144
            chk(chk_�շ���Ŀ��λ��������).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1, 1, 0)
        Case 145
            chk(chk_ÿ��סԺʹ����סԺ��).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1, 1, 0)
        Case 147
            txtUD(ud_��ͯ����綨����).Text = zlCommFun.Nvl(rsTemp("����ֵ"), 12)
        Case 149    'ҩƷЧ����ʾ��ʽ
            cmb(cmb_Ч����ʾ��ʽ).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 150    'ҩƷ���������㷨
            cmb(cmb_ҩƷ���������㷨).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 151    '�����˷���������
            chk(chk_�����˷���������).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1, 1, 0)
            '            Case 152 '���￨�ظ�ʹ�ã����˺飺24357
            '                chk(chk_���￨�ظ�ʹ��).Value = IIF(Val(zlCommFun.Nvl(rsTemp("����ֵ"))) = 1, 1, 0)
        Case 154    '��Ժʱ���δ��ҩ��Ŀ
            cmb(cmb_��Ժʱδ��ҩ��Ŀ���).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 155    'ת��ʱ���δ��ҩ��Ŀ
            cmb(cmb_ת��ʱδ��ҩ��Ŀ���).ListIndex = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 158    '����¼��ʱ��
            If IsNull(rsTemp!����ֵ) Then
                txtInputHours.Text = zlCommFun.Nvl(rsTemp("ȱʡֵ"), 0)
            Else
                txtInputHours.Text = rsTemp!����ֵ
            End If
        Case 160    '����Ѽ����׼:34741
            If zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1 Then
                opt����(1).Value = True
            Else
                opt����(0).Value = True
            End If

        Case 161    '�Ƿ�����ʹ�ý���ҩ��
            chk(chk_����ҩ��).Value = Val("" & rsTemp("����ֵ"))

        Case 162    '�´�ҽ��ʱ��ʾ����
            chk(chk_�´�ҽ��ʱ��ʾ����).Value = Val("" & rsTemp("����ֵ"))

        Case 171
            chk(chk_����δ�շѵ����ﻮ�۴�������).Value = Val("" & rsTemp("����ֵ"))
        Case 172
            chk(chk_����δ��˵ļ��˴�������).Value = Val("" & rsTemp("����ֵ"))
        Case 173    '�⹺�����Ҫ������Ǹ������ܽ��и������
            chk(chk_�⹺�����Ҫ������Ǹ������ܽ��и���).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 174    'ҩƷ�ƿ�ʱ��ȷҩƷ����
            chk(chk_ҩƷ�ƿ���ȷ����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 175    'ҩƷ����ʱ��ȷҩƷ����
            chk(chk_ҩƷ������ȷ����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 181    'ҩƷ�ֶμӳ����
            chk(chk_ʱ�۷ֶμӳ����).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1, 1, 0)
        Case 182    '��ֹ�´ﳬ����ҩƷҽ��
            chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Value = IIF(zlCommFun.Nvl(rsTemp("����ֵ"), 0) = 1, 0, 1)
        Case 183    'ʱ��ҩƷ��ⰴȡ�ϴ��ۼ�
            chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 186  '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
            chk(chk_��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 187
            chk(chk����ҩ��ּ�����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 188
            chk(chk����ҩ��ʹ���Ա�ҩ).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
            chk(chk����ҩ��ʹ���Ա�ҩ).Enabled = chk(chk����ҩ��ּ�����).Value = 1
        Case 189
            chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 191
            chk(chkֻ����¼����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 1)
        Case 192
            chk(chk_���˳�Ժҽ������������Ժ).Value = IIF(Val("" & rsTemp!����ֵ) <> 0, 1, 0)
        Case 208
            chk(chk�ٴ�����վ����ʹ��zlPlugIn����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 209
            chk(chk���������ּ�����).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 210
            chk(chk_���������Һ���Ч�����Ĳ���).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 213
            cmb(cmd_��ҩ�䷽).ListIndex = IIF(Val("" & rsTemp!����ֵ) = 4, 1, 0)
        Case 214
            chk(chk_�״�ҽ��ִ����Ҫ���).Value = zlCommFun.Nvl(rsTemp("����ֵ"), 0)
        Case 215    '51612
            chk(chk_δ��ƽ�ֹ����).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        Case 216
            chk(chk_��Ѫ�ּ�����).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        Case 217
            chk(chk_������Ȩ����).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        Case 218
            chk(chk_��Ѫ�����������).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        Case 219
            chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        Case 220    '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
            txtUNExecLimit.Text = zlCommFun.Nvl(rsTemp("����ֵ"), 999)
            If txtUNExecLimit.Text = "999" Then
                chk(chk_ҽ��ִ����Ч����).Value = 0
                txtUNExecLimit.Enabled = False
            Else
                chk(chk_ҽ��ִ����Ч����).Value = 1
            End If
        Case 221
            If Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0)) = 0 Then
                optAccountTime(1).Value = True
                txtAccountTime.Enabled = False
            Else
                optAccountTime(0).Value = True
                txtAccountTime.Enabled = True
                txtAccountTime.Text = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0))
            End If
        Case 223  '�����¿�ҽ�����
            ud(ud_�����¿�ҽ�����).Value = Nvl(rsTemp!����ֵ, 1)
        Case 224    '����������Դ
            '̫Ԫͨ������ҩ�ӿڣ���Ϊ�Ѿ���������������˿���ʹ�ÿؼ���ֵ
            If cmb(cmb_������ҩ�ӿ�).ListIndex = 3 Then
                cmb(cmd_����������Դ).ListIndex = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0))
            End If
        Case 225  '���ô�ͨ�ӿ���־����65522
            chk(chk_���ýӿڵ�����־).Value = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0))
        Case 226 '�����ӿڲ���
            chk(chk_����ʹ��ϵͳ����).Value = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 1))
        Case 227 'ת��ʱ���δ������ʵ���
            cmb(cmd_ת��ʱδ������ʵ���).ListIndex = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0))
        Case 228
            strTemp = zlCommFun.Nvl(rsTemp("����ֵ"), "3.0")
            If strTemp = "3.0" Then
                optPASSVer(0).Value = True
            Else
                optPASSVer(1).Value = True
            End If
        Case 230
            chk(chk_ҽ������ʱ��������ԭ��).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
            Call Set��д��������(chk(chk_ҽ������ʱ��������ԭ��).Value = 1)
        Case 233
            strTemp = zlCommFun.Nvl(rsTemp("����ֵ"))
            Call Init�����˵��(strTemp)
        Case 234
            strTemp = zlCommFun.Nvl(rsTemp("����ֵ"))
            Call Initת�Ƴ�Ժ�������Ŀ(strTemp)
        Case 235
            cmb(cmd_��Ժʱ���ڻ�������).ListIndex = Val(zlCommFun.Nvl(rsTemp("����ֵ"), 0))
        Case 239
            chk(chk_�¿�ҽ��ǩ��ʱһ��ҽ��ǩ��һ��).Value = IIF(Val(zlCommFun.Nvl(rsTemp!����ֵ)) = 1, 1, 0)
        End Select
        rsTemp.MoveNext
    Loop

    '��ʾ��ǰƱ�ݵ����
    lvw(lvw_Ʊ��).ListItems("C1").Selected = True
    lvw_ItemClick lvw_Ʊ��, lvw(lvw_Ʊ��).SelectedItem

    '����ǩ������
    Call cmb_Click(cmb_����ǩ����֤����)
    Call LoadSign
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load���ݲ���()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.��ԱID,B.����,A.����,A.ʱ������,A.���˵���,A.������� from ���ݲ������� A,��Ա�� B where A.��ԱID=B.ID"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    lvw(lvw_����).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_����).ListItems.Add(, , rsTemp("����"), , "Limit")
        
        str���� = Switch(rsTemp("����") = 1, "�Һŵ���", rsTemp("����") = 2, "�շѵ�", rsTemp("����") = 3, "���۵�", rsTemp("����") = 4, "�������", _
                       rsTemp("����") = 5, "סԺ����", rsTemp("����") = 6, "Ԥ����", rsTemp("����") = 7, "���ʵ���", rsTemp("����") = 8, "���￨", rsTemp("����") = 9, "����")
        lst.SubItems(1) = str����
        lst.SubItems(2) = rsTemp("ʱ������")
        lst.SubItems(3) = IIF(rsTemp("���˵���") = 1, "��", "��")
        lst.SubItems(4) = IIF(IsNull(rsTemp("�������")), "", Format(rsTemp("�������"), "0.00"))
        lst.Tag = rsTemp("��ԱID")
        lst.ListSubItems(1).Tag = rsTemp("����")
        
        rsTemp.MoveNext
    Loop
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load����()
    Dim rs���� As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    rs����.CursorLocation = adUseClient
    gstrSQL = "select A.ID,A.����,A.���� " & _
               " from  ��������˵�� b,���ű� a " & _
               " where B.������� in(1,2,3) And B.��������='����' and  b.����ID=a.ID and " & _
               Where����ʱ��("A") & " order by ����"
    Call OpenRecordset(rs����, Me.Caption)
    
    Bill(bill_�Զ�����).Clear
    Bill(bill_���ʱ���).Clear
    
    If rs����.RecordCount > 0 Then
        msh(0).Rows = rs����.RecordCount + 1
        lngRow = 1
        Do Until rs����.EOF
            Bill(bill_�Զ�����).AddItem rs����("����") & "-" & rs����("����")
            Bill(bill_�Զ�����).ItemData(Bill(bill_�Զ�����).NewIndex) = rs����("ID")
            Bill(bill_���ʱ���).AddItem rs����("����") & "-" & rs����("����")
            Bill(bill_���ʱ���).ItemData(Bill(bill_�Զ�����).NewIndex) = rs����("ID")
            msh(0).TextMatrix(lngRow, 0) = rs����("����") & "-" & rs����("����")
            msh(0).RowData(lngRow) = rs����("ID")
            lngRow = lngRow + 1
            rs����.MoveNext
        Loop
        Bill(bill_�Զ�����).ListIndex = 0
    End If
    Bill(bill_���ʱ���).AddItem "*����*"
    Bill(bill_���ʱ���).ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadSign()
'���ܣ����ص���ǩ�����ò���
    Dim rsTmp As New Recordset
    Dim i As Long, lngTmp As Long
    
    gstrSQL = "select ����ID,���� from ����ǩ�����ò���"
    On Error GoTo ErrHandle
    Call OpenRecordset(rsTmp, Me.Caption)
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            rsTmp.Filter = "����=" & i
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                lngTmp = .FindRow(Val(rsTmp!����id & ""))
                If lngTmp <> -1 Then
                    .Cell(flexcpChecked, lngTmp, col_ѡ��) = 1
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
'�������ĳ�ʼ������
    Dim rsTemp As New ADODB.Recordset
    Dim lngMaxRow As Long, lngRow As Long, lng��λ As Long
    Dim strTemp As String, lngTemp As Long, i As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean, strCoding As String
    
    
    '�շ��ض���Ŀ
    On Error GoTo ErrHandle
    gstrSQL = "select a.�ض���Ŀ ,c.ID,c.����  " & _
            " from �շ��ض���Ŀ a,�շ�ϸĿ c " & _
            " where a.�շ�ϸĿID =c.id"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("�ض���Ŀ")
            Case "������"
                txtCmd(0).Tag = rsTemp("ID")
                txtCmd(0).Text = rsTemp("����")
            Case "������"
                txtCmd(1).Tag = rsTemp("ID")
                txtCmd(1).Text = rsTemp("����")
            Case "��ͨ���÷�"
                txtCmd(3).Tag = rsTemp("ID")
                txtCmd(3).Text = rsTemp("����")
            Case "�������÷�"
                txtCmd(4).Tag = rsTemp("ID")
                txtCmd(4).Text = rsTemp("����")
        End Select
        rsTemp.MoveNext
    Loop
    
    '�����Զ����ʳ���
    gstrSQL = "select A.����ID,B.����,b.���� as ���� ,a.�շ�ϸĿID,c.���� as �շ�ϸĿ ,a.�����־,a.�������� " & _
            " from �Զ��Ƽ���Ŀ A,���ű� B,�շ�ϸĿ C " & _
            " where A.����ID= B.id and A.�շ�ϸĿID =C.id(+) " & _
            " order by b.���� "
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Bill(bill_�Զ�����)
        lngRow = 1
        Do Until rsTemp.EOF
            If IsNull(rsTemp("�շ�ϸĿID")) Then
                '��λ�ѻ����
                For lngTemp = 1 To msh(0).Rows - 1
                    If msh(0).RowData(lngTemp) = rsTemp("����ID") Then
                        If rsTemp("�����־") = 1 Then
                            '��λ��
                            msh(0).TextMatrix(lngTemp, 1) = "��"
                            msh(0).TextMatrix(lngTemp, 2) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                            msh(0).TextMatrix(lngTemp, 5) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                        Else
                            '�����
                            msh(0).TextMatrix(lngTemp, 3) = "��"
                            msh(0).TextMatrix(lngTemp, 4) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                            msh(0).TextMatrix(lngTemp, 6) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                        End If
                    End If
                Next
            Else
                '��������
                .Rows = lngRow + 1
                .RowData(lngRow) = rsTemp("����ID")
                .TextMatrix(lngRow, 0) = rsTemp("����") & "-" & rsTemp("����")
                .TextMatrix(lngRow, 1) = rsTemp("�շ�ϸĿID")
                .TextMatrix(lngRow, 2) = rsTemp("�շ�ϸĿ")
                .TextMatrix(lngRow, 3) = Switch(rsTemp("�����־") = 6, "1-������", rsTemp("�����־") = 8, "2-����һ��", True, "0-��������")
                .TextMatrix(lngRow, 4) = Format(IIF(IsNull(rsTemp!��������), "", rsTemp!��������), "yyyy-mm-dd")
                lngRow = lngRow + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '���ʱ������
    gstrSQL = "Select ����,��� From �շ���� Order by ����"
    Set mrs��� = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs���, gstrSQL, Me.Caption)
    
    lst���.Clear
    lst���.AddItem "�������"
    Do While Not mrs���.EOF
        lst���.AddItem mrs���!���
        lst���.ItemData(lst���.NewIndex) = Asc(mrs���!����)
        mrs���.MoveNext
    Loop
    
    '�������ʱ�����
    Set mrsWarn = New ADODB.Recordset
    mrsWarn.Fields.Append "����ID", adBigInt, , adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "���ò���", adVarChar, 100
    mrsWarn.Fields.Append "��������", adSmallInt
    mrsWarn.Fields.Append "����ֵ", adCurrency
    mrsWarn.Fields.Append "������־1", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־2", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "������־3", adVarChar, 100, adFldIsNullable
    mrsWarn.Fields.Append "�߿�����", adCurrency
    mrsWarn.Fields.Append "�߿��׼", adCurrency
    
    mrsWarn.CursorLocation = adUseClient
    mrsWarn.LockType = adLockOptimistic
    mrsWarn.CursorType = adOpenStatic
    mrsWarn.Open
    
    gstrSQL = "" & _
    "   Select a.����ID,B.����,b.���� as ����,a.���ò���,nvl(a.��������,1) as ��������, " & _
    "               a.����ֵ,a.������־1,a.������־2,a.������־3,A.�߿�����,a.�߿��׼ " & _
    "   From ���ʱ����� a,���ű� b " & _
    "   Where a.����ID= b.id(+)  " & _
    "   Order by Decode(a.���ò���,'��ͨ����',1,'ҽ������',2,3),a.���ò���,B.���� Desc"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    strCoding = ",��ͨ����" '������һ����ͨ����
    Do Until rsTemp.EOF
        mrsWarn.AddNew
        mrsWarn!����id = rsTemp!����id
        mrsWarn!������ = rsTemp!����
        mrsWarn!������ = rsTemp!����
        mrsWarn!���ò��� = rsTemp!���ò���
        mrsWarn!�������� = rsTemp!��������
        mrsWarn!����ֵ = rsTemp!����ֵ
        mrsWarn!������־1 = rsTemp!������־1
        mrsWarn!������־2 = rsTemp!������־2
        mrsWarn!������־3 = rsTemp!������־3
        mrsWarn!�߿����� = Val(Nvl(rsTemp!�߿�����))
        mrsWarn!�߿��׼ = Val(Nvl(rsTemp!�߿��׼))
        mrsWarn.Update
        
        If InStr(strCoding & ",", "," & rsTemp!���ò��� & ",") = 0 Then
            strCoding = strCoding & "," & rsTemp!���ò���
        End If
        rsTemp.MoveNext
    Loop
    strCoding = Mid(strCoding, 2)
    tab����.Tabs.Clear
    For i = 0 To UBound(Split(strCoding, ","))
        tab����.Tabs.Add , , Split(strCoding, ",")(i)
    Next
    tab����.Tabs(1).Selected = True '֮ǰ���ἤ��Click�¼�,��Ϊ����
   
    '����ⷿ��λ
    strCoding = ""
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.����,'') ����,nvl(b.����,'') ����,a.�������,a.��������" & vbCrLf & _
            "          FROM ��������˵�� A, ���ű� B" & vbCrLf & _
            " WHERE B.ID=A.����ID AND A.�������� IN ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')  order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    msf�ⷿ��λ.Rows = 1
    Do Until rsTemp.EOF
        With msf�ⷿ��λ
            If rsTemp("����") <> strCoding Then
                strTemp = ""
            End If
            If InStr(",��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
                If InStr(1, strTemp & ",", ",ҩ��,") <= 0 Then
                    .Rows = .Rows + 1
                    .RowData(.Rows - 1) = rsTemp("ID")
                    .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                    .TextMatrix(.Rows - 1, 1) = "ҩ��"
                    strTemp = strTemp & "," & "ҩ��"
                End If
            End If
            
            If InStr(",�Ƽ���,��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
            
                Select Case rsTemp("�������")
                    Case 0          '�������ڲ���
'                        .Rows = .Rows + 1
'                        .RowData(.Rows - 1) = rsTemp("ID")
'                        .TextMatrix(.Rows - 1, 0) = rsTemp("����")
'                        .TextMatrix(.Rows - 1, 1) = "����"
                    Case 1          '���������ﲡ��
                        If InStr(1, strTemp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTemp = strTemp & "," & "����"
                        End If
                    Case 2          '������סԺ����
                        If InStr(1, strTemp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTemp = strTemp & "," & "סԺ"
                        End If
                    Case 3          '����������סԺ����
                        If InStr(1, strTemp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTemp = strTemp & "," & "����"
                        End If
                        
                        If InStr(1, strTemp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTemp = strTemp & "," & "סԺ"
                        End If
                End Select
            End If
            If InStr(1, strTemp & ",", ",����,") <= 0 Then
                .Rows = .Rows + 1
                .RowData(.Rows - 1) = rsTemp("ID")
                .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                .TextMatrix(.Rows - 1, 1) = "����"
                strTemp = strTemp & "," & "����"
            End If
            
            strCoding = rsTemp("����")
        End With
        rsTemp.MoveNext
    Loop

    If msf�ⷿ��λ.Rows > 1 Then
        msf�ⷿ��λ.FixedRows = 1
    End If
    gstrSQL = "select �ⷿid, ���÷�Χ, ���� from ҩƷ�ⷿ��λ"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngMaxRow = rsTemp.RecordCount
        For lngRow = 1 To lngMaxRow
            For i = 1 To msf�ⷿ��λ.Rows - 1
                Select Case rsTemp!���÷�Χ
                    Case 1
                        strTemp = "ҩ��"
                    Case 2
                        strTemp = "����"
                    Case 3
                        strTemp = "סԺ"
                    Case 4
                        strTemp = "����"
                End Select
                If rsTemp!�ⷿid = msf�ⷿ��λ.RowData(i) And strTemp = msf�ⷿ��λ.TextMatrix(i, 1) Then
                    msf�ⷿ��λ.TextMatrix(i, 2) = ""
                    msf�ⷿ��λ.TextMatrix(i, 3) = ""
                    msf�ⷿ��λ.TextMatrix(i, 4) = ""
                    msf�ⷿ��λ.TextMatrix(i, 5) = ""
                    msf�ⷿ��λ.TextMatrix(i, rsTemp!���� + 1) = "��"
                End If
            Next
            rsTemp.MoveNext
        Next
    End If
    
    'ҩ����ҩ����
    strCoding = ""
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.����,'') ����,nvl(b.����,'') ����,a.�������,a.��������" & vbCrLf & _
            "          FROM ��������˵�� A, ���ű� B" & vbCrLf & _
            " WHERE B.ID=A.����ID AND A.�������� IN ('�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')  order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    Billҩ����ҩ����.Rows = 1
    Do Until rsTemp.EOF
        With Billҩ����ҩ����
            If rsTemp("����") <> strCoding Then
                strTemp = ""
            End If
            
            If InStr(",�Ƽ���,��ҩ��,��ҩ��,��ҩ��,", "," & rsTemp("��������") & ",") Then
            
                Select Case rsTemp("�������")
                    Case 0          '�������ڲ���
'                        .Rows = .Rows + 1
'                        .RowData(.Rows - 1) = rsTemp("ID")
'                        .TextMatrix(.Rows - 1, 0) = rsTemp("����")
'                        .TextMatrix(.Rows - 1, 1) = "����"
                    Case 1          '���������ﲡ��
                        If InStr(1, strTemp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTemp = strTemp & "," & "����"
                        End If
                    Case 2          '������סԺ����
                        If InStr(1, strTemp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTemp = strTemp & "," & "סԺ"
                        End If
                    Case 3          '����������סԺ����
                        If InStr(1, strTemp & ",", ",����,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "����"
                            strTemp = strTemp & "," & "����"
                        End If
                        
                        If InStr(1, strTemp & ",", ",סԺ,") <= 0 Then
                            .Rows = .Rows + 1
                            .RowData(.Rows - 1) = rsTemp("ID")
                            .TextMatrix(.Rows - 1, 0) = rsTemp("����")
                            .TextMatrix(.Rows - 1, 1) = "סԺ"
                            strTemp = strTemp & "," & "סԺ"
                        End If
                End Select
            End If
            strCoding = rsTemp("����")
        End With
        rsTemp.MoveNext
    Loop

    gstrSQL = "select ҩ��id, ����, ��ҩ, �Զ���ҩ����,��ҩȷ�� from ҩ����ҩ����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With Billҩ����ҩ����
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            lngMaxRow = rsTemp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To .Rows - 1
                    Select Case rsTemp!����
                        Case 1
                            strTemp = "����"
                        Case 2
                            strTemp = "סԺ"
                    End Select
                    If rsTemp!ҩ��id = .RowData(i) And strTemp = .TextMatrix(i, 1) Then
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ����()
'����:װ��ҩƷ��������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_ҩƷ����)
        '����װ��ⷿ
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') " & _
                   " and  b.����ID=a.ID and " & Where����ʱ��("A") & " order by ����"
        Call OpenRecordset(rsTemp, Me.Caption)
        
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
        Call OpenRecordset(rsTemp, Me.Caption)
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load��������()
'���ܣ���ʼ����������
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select ����,���� From �������� Order by ����"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    lst(lst_ҽ������).Clear
    lst(lst_���Ѳ���).Clear
    Do Until rsTemp.EOF
        lst(lst_ҽ������).AddItem rsTemp("����") & "." & rsTemp("����")
        lst(lst_���Ѳ���).AddItem rsTemp("����") & "." & rsTemp("����")
        
        rsTemp.MoveNext
    Loop
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
    Dim ObjItem As ListItem
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) ��鷽ʽ" & vbCrLf & _
        " FROM ��������˵�� A, ���ű� B, ҩƷ������ C" & vbCrLf & _
        " WHERE A.����ID = B.ID AND A.����ID = C.�ⷿID(+) AND" & vbCrLf & _
        "      A.�������� IN" & vbCrLf & _
        "      ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')" & vbCrLf & _
        "     And (b.����ʱ��=to_date('3000-1-1','yyyy-mm-dd') or b.����ʱ�� is null) " & vbCrLf & _
        " GROUP BY B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) " & vbCrLf & _
        " order by B.���� "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.lvwCheckMed.ListItems.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set ObjItem = Me.lvwCheckMed.ListItems.Add(, "C_" & rsTmp!ID, "[" & zlCommFun.Nvl(rsTmp!����) & "]", "bm", "bm")
            ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!����)
            ObjItem.SubItems(2) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
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

Private Sub Save�ⷿ��λ()
    '����ⷿ��λ����
    Dim i As Long
    Dim lngTmp As Long
    Dim intTmp As Integer
    Dim strSQL As String
    On Error GoTo ErrHandle
    If msf�ⷿ��λ.Rows > 1 Then
        If Trim(msf�ⷿ��λ.TextMatrix(1, 0)) <> "" Then
            gstrSQL = ""
            For i = 1 To msf�ⷿ��λ.Rows - 1
                gstrSQL = gstrSQL & msf�ⷿ��λ.RowData(i) & ","
                lngTmp = 1
                Select Case True
                    Case msf�ⷿ��λ.TextMatrix(i, 2) = "��"
                        lngTmp = 1
                    Case msf�ⷿ��λ.TextMatrix(i, 3) = "��"
                        lngTmp = 2
                    Case msf�ⷿ��λ.TextMatrix(i, 4) = "��"
                        lngTmp = 3
                    Case msf�ⷿ��λ.TextMatrix(i, 5) = "��"
                        lngTmp = 4
                End Select
                Select Case msf�ⷿ��λ.TextMatrix(i, 1)
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
            strSQL = "ZL_ҩƷ�ⷿ��λ_DELETE"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gstrSQL = "ZL_ҩƷ�ⷿ��λ_INSERT('" & gstrSQL & "')"
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


Private Function Save�ⷿ���() As Boolean
    '���ܣ�����ⷿ���
    Dim i As Long
    On Error GoTo ErrHandle
    
    gstrSQL = ""
    For i = 1 To Me.lvwCheckMed.ListItems.Count
        gstrSQL = gstrSQL & Me.lvwCheckMed.ListItems(i).Tag & "," & Switch(Me.lvwCheckMed.ListItems(i).SubItems(2) = "0-�����", "0", Me.lvwCheckMed.ListItems(i).SubItems(2) = "1-��飬��������", "1", Me.lvwCheckMed.ListItems(i).SubItems(2) = "2-��飬�����ֹ", "2") & ","
    Next
    gstrSQL = "Zl_ҩƷ������_insert('" & gstrSQL & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save�ⷿ��� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Or lvw����.Tag <> "" Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    '�����п�
    SaveFlexState msh(0), App.ProductName & "\" & Me.Name
    SaveFlexState Bill(bill_�Զ�����), App.ProductName & "\" & Me.Name & bill_�Զ�����
    SaveFlexState Bill(bill_���ʱ���), App.ProductName & "\" & Me.Name & bill_���ʱ���
    SaveFlexState Bill(bill_ҩƷ����), App.ProductName & "\" & Me.Name & bill_ҩƷ����
    SaveFlexState Bill(bill_ҩƷ��������), App.ProductName & "\" & Me.Name & bill_ҩƷ��������
    
    Set mrsWarn = Nothing
    Set mrs��� = Nothing
    Set mcol���� = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If CheckDataValid() = False Then Exit Sub
    If Save����() = False Then Exit Sub
    mblnChange = False
    lvw����.Tag = ""
    Unload Me
End Sub

Private Function Check���ʱ���() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim arr���() As String
        
    With Bill(bill_���ʱ���)
        For lngRow = 1 To .Rows - 2
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .TextMatrix(lngRow, 0) = .TextMatrix(lngTemp, 0) And .TextMatrix(lngTemp, 2) <> "" Then
                        MsgBox "������" & .TextMatrix(lngTemp, 0) & "�����ֶ�Ρ�", vbExclamation, gstrSysName
                        .Row = lngTemp: .Col = 0: .SetFocus: Exit Function
                    End If
                Next
                '���˺� ����: 34770   ����:2010-12-21 10:54:02
                If Val(.TextMatrix(lngRow, 6)) > 999999999 Or Val(.TextMatrix(lngRow, 6)) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "���еĴ߿�������������(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 6: .SetFocus: Exit Function
                End If
                If Val(.TextMatrix(lngRow, 7)) > 999999999 Or Val(.TextMatrix(lngRow, 7)) < 0 Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "���еĴ߿��׼����(Ӧ����0~999999999)!", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 7: .SetFocus: Exit Function
                End If
                
            End If
        Next
        
        '���ͬһ������ͬ������ʽ������Ƿ�һ����û�����û��ظ�
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 2) <> "" Then
                If Trim(.TextMatrix(lngRow, 3)) = "" And Trim(.TextMatrix(lngRow, 4)) = "" And Trim(.TextMatrix(lngRow, 5)) = "" Then
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "��δ����Ҫ�������շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If (.TextMatrix(lngRow, 3) = "�������" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 4) = "�������" And (Trim(.TextMatrix(lngRow, 3)) <> "" Or Trim(.TextMatrix(lngRow, 5)) <> "")) _
                    Or (.TextMatrix(lngRow, 5) = "�������" And (Trim(.TextMatrix(lngRow, 4)) <> "" Or Trim(.TextMatrix(lngRow, 3)) <> "")) Then
                    
                    MsgBox "������" & .TextMatrix(lngRow, 0) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
                    .Row = lngRow: .Col = 3: .SetFocus: Exit Function
                End If
                If .TextMatrix(lngRow, 3) <> "�������" And Trim(.TextMatrix(lngRow, 4)) <> "�������" And Trim(.TextMatrix(lngRow, 5)) <> "�������" Then
                    For lngCol1 = 3 To 5
                        If Trim(.TextMatrix(lngRow, lngCol1)) <> "" Then
                            For lngCol2 = 3 To 5
                                If lngCol1 <> lngCol2 Then
                                    arr��� = Split(.TextMatrix(lngRow, lngCol1), ",")
                                    For lngTemp = 0 To UBound(arr���)
                                        If InStr("," & .TextMatrix(lngRow, lngCol2) & ",", "," & arr���(lngTemp) & ",") > 0 Then
                                            MsgBox "������" & .TextMatrix(lngRow, 0) & "����ͬ�ı�����ʽ������ͬ���շ����", vbExclamation, gstrSysName
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
    
    Check���ʱ��� = True
End Function

Private Function CheckDataValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngIndex As Long, strTmp As String
    Dim i As Integer
    
    
    '�Զ��Կ��ұ�����һ���༭��������У��
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
    
    '����Զ�������Ŀ�Ƿ��ظ�
    With Bill(bill_�Զ�����)
        For lngRow = 1 To .Rows - 2
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                For lngTemp = lngRow + 1 To .Rows - 1
                    If .RowData(lngRow) = .RowData(lngTemp) And .TextMatrix(lngRow, 1) = .TextMatrix(lngTemp, 1) Then
                        MsgBox "����Ϊ��" & .TextMatrix(lngTemp, 0) & "�����շ�ϸĿΪ��" & _
                            .TextMatrix(lngTemp, 2) & "��" & vbCrLf & "������ϳ��ֶ�Ρ�", vbExclamation, gstrSysName
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
    
    '����Զ�������Ŀ����������
    With Bill(bill_�Զ�����)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 And .TextMatrix(lngRow, 1) <> "" Then
                If Not IsDate(.TextMatrix(lngRow, 4)) Then
                    MsgBox "�Զ�������Ŀ����������δ���û����ڸ�ʽ����ȷ��", vbInformation, gstrSysName
                    .Row = lngRow
                    .Col = 4
                    Call ShowTab(4)
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    '���ҩƷ��������
    With Bill(bill_ҩƷ����)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "��" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(8)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(8)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(8)
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
                Call ShowTab(11)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(11)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(11)
                    Exit Function
                End If
            Next
        Next
    End With
    
    If txtUD(ud_���ý���λ��).Text <> mDecimal Then
        If MsgBox("���ѵ����˷��ý���С��λ�����ܻ�����С���������Ƿ������", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    If txtUD(ud_���õ��۱���λ��).Text <> pDecimal Then
        If MsgBox("���ѵ����˷��õ��۱���С��λ�����ܻ�����С���������Ƿ������", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
            Call ShowTab(1)
            Exit Function
        End If
    End If
      
    If CheckNumberRule_Drug = True Then
'        With mshBillEdit
'            If Len(Trim(.TextMatrix(1, 1))) > 0 Then
'                For i = 1 To .Rows - 1
'                    If Len(Trim(.TextMatrix(i, 2))) <= 0 Then
'                        MsgBox "ҩƷ���ұ�Ų���Ϊ��!", vbInformation, gstrSysName
'                        Call ShowTab(13)
'                        Exit Function
'                    End If
'                Next
'            End If
'        End With
        
        'ͬһ��GRID��Ŀ��ұ�Ų����ظ�
        With mshBillEdit
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "ҩƷ���ҵ�" & i & "�б���ظ���", vbQuestion, gstrSysName
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
'                        MsgBox "���Ŀ��ұ�Ų���Ϊ��!", vbInformation, gstrSysName
'                        Call ShowTab(13)
'                        Exit Function
'                    End If
'                Next
'            End If
'        End With
        
        'ͬһ��GRID��Ŀ��ұ�Ų����ظ�
        With mshBillEditStuff
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "���Ŀ��ҵ�" & i & "�б���ظ���", vbQuestion, gstrSysName
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
    
    If chk(chk_ҩƷ�ʱ�¿��ÿ��).Value <> chk(chk_ҩƷ�ʱ�¿��ÿ��).Tag Or chk(chk_��ȷ����ҩƷ����).Value <> chk(chk_��ȷ����ҩƷ����).Tag Then
        If Check�Ƿ���δ��˵�ҩƷ���� Then
            MsgBox "����δ��˵�ҩƷ���ݣ����ܸı����!", vbInformation, gstrSysName
            chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = chk(chk_ҩƷ�ʱ�¿��ÿ��).Tag
            chk(chk_��ȷ����ҩƷ����).Value = chk(chk_��ȷ����ҩƷ����).Tag
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    If chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Value <> chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Tag Then
        If Check�Ƿ���δ��˵��⹺��ⵥ Then
            MsgBox "����δ��˵��⹺��ⵥ�����ܸı������ʱ��ҩƷ��ⰴ��ǰ�ӳ����ۡ�!", vbInformation, gstrSysName
            chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Value = chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Tag
            Call ShowTab(1)
            Exit Function
        End If
    End If
    
    CheckDataValid = True
End Function

Private Function Save����() As Boolean
    On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    Call SavePara
    Call Save����ǩ��
    Call Save�����ӿ�
    Call Save���ݲ���
    Call Save�շ��ض���Ŀ
    Call Save�Զ��Ƽ���Ŀ
    Call SaveҩƷ����
    Call Save���ʱ�����
    Call SaveRegister
    Call Save�ⷿ���
    Call Save�ⷿ��λ
    Call SaveҩƷ��������
    Call Saveҩ����ҩ����
    Call Save���ݱ������
    Call Save����
    Call Saveҽ������
    Call SaveҩƷ���ľ���
    Call Save���ݻ��ڿ���
    
    '������ϣ������ύ
    gcnOracle.CommitTrans
    Call zlDatabase.ClearParaCache
    Save���� = True
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
    Dim str��� As String, i As Long

    On Error GoTo ErrHandle
    '����Բ������б���
    strTemp = "1," & Format(dtp(dtp_�����ϰ�).Value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).Value, "HH:mm") & ","
    strTemp = strTemp & "2," & Format(dtp(dtp_�����ϰ�).Value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).Value, "HH:mm") & ","
    strTemp = strTemp & "5," & ud(ud_��¼ҽ��ʶ����).Value & ","
    strTemp = strTemp & "6," & chk(chk_δ��˼��ʴ�����ҩ).Value & ","
    strTemp = strTemp & "7," & chk(chk_�Զ�����).Value & ","
    strTemp = strTemp & "9," & Val(Me.txtUD(ud_���ý���λ��).Text) & ","
    strTemp = strTemp & "157," & Val(Me.txtUD(ud_���õ��۱���λ��).Text) & ","
    strTemp = strTemp & "10," & chk(chk_��ȡԤ����).Value & ","
    strTemp = strTemp & "11," & chk(chk_ʱ������￨).Value & ","
    'strTemp = strTemp & "12," & chk(chk_������ʾ).Value & ","
    strTemp = strTemp & "13," & chk(chk_���䴲λ��).Value & ","
    strTemp = strTemp & "14," & Split(cmb(cmb_�Һ���Ǯ����).Text & "-", "-")(0) & cmb(cmb_�շ���Ǯ����).ListIndex & cmb(cmb_������Ǯ����).ListIndex & ","
    strTemp = strTemp & "15," & chk(chk_�����շ��뷢ҩ����).Value & ","
    strTemp = strTemp & "16," & chk(chk_סԺ�����뷢ҩ����).Value & ","
    strTemp = strTemp & "17," & chk(chk_��������).Value & chk(chk_ˢ���￨).Value & chk(chk_�Һŵ���).Value & chk(chk_����ID).Value & ","
    strTemp = strTemp & "18," & chk(chk_�޶�ҩƷ�Ŀ��).Value & ","
    strTemp = strTemp & "45," & chk(chk_�շ�ͬʱ��ҩ).Value & ","
    strTemp = strTemp & "51," & chk(chk_����ִ�еǼ�).Value & ","
    strTemp = strTemp & "52," & chk(chk_���뿪����).Value & ","
    strTemp = strTemp & "53," & chk(chk_���ƿ�����).Value & ","
    strTemp = strTemp & "54," & chk(chk_ʱ��ҩƷ���).Value & ","

    strTemp = strTemp & "19," & IIF(opt(opt_��æ��ʽ).Value = True, "0", "1") & ","
    strTemp = strTemp & "20,"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C2").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C3").SubItems(1) & "|"
    strTemp = strTemp & lvw(lvw_Ʊ��).ListItems("C4").SubItems(1) & "|,"
    strTemp = strTemp & "21," & ud(ud_�Һŵ�).Value & ud(ud_����Һŵ�).Value & ","
    strTemp = strTemp & "22," & (cmb(cmb_��Ժʱδִ����Ŀ���).ListIndex) & ","
    strTemp = strTemp & "23," & cmb(cmb_�ѽᵥ��).ListIndex & ","
    strTemp = strTemp & "24,"
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C2").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C3").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C4").SubItems(2) = "��", "1", "0") & ","
    '����ǩ��ϵͳ����
    strTemp = strTemp & "25," & Val(cmb(cmb_����ǩ����֤����).List(cmb(cmb_����ǩ����֤����).ListIndex)) & ","
    strTemp = strTemp & "26," & chk(chk_����ǩ������_����).Value & chk(chk_����ǩ������_סԺ).Value & chk(chk_����ǩ������_ҽ��).Value & chk(chk_����ǩ������_����).Value & chk(chk_����ǩ������_ҩƷ).Value & chk(chk_����ǩ������_lis).Value & chk(chk_����ǩ������_pacs).Value & ","
    strTemp = strTemp & "27," & chk(chk_סԺҩ�����Ͳ�����ҩ��).Value & ","
    strTemp = strTemp & "28," & chk(chk_���ﲡ������ʱ��Ҫˢ����֤).Value & ","
    strTemp = strTemp & "29," & cmb(cmb_���۵�λ).ListIndex & ","
    strTemp = strTemp & "30," & cmb(cmb_������ҩ�ӿ�).ListIndex & ","
    strTemp = strTemp & "31," & chk(chk_��Ժ���˲�׼��Ժ����).Value & ","
    strTemp = strTemp & "32," & cmb(cmb_ת��ʱδִ����Ŀ���).ListIndex & ","
    strTemp = strTemp & "33," & chk(chk_ִ��֮���Զ�����).Value & ","
    strTemp = strTemp & "34," & chk(chk_ָ��ҽ������������ִ��).Value & ","
    'ע�ⷵ��ֵ����,�ָ��������������š�����ʱҪת��һ��
    strTemp = strTemp & "41," & Replace(Replace(GetTextFromList(lst(lst_ҽ������)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "42," & Replace(Replace(GetTextFromList(lst(lst_���Ѳ���)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "43," & chk(chk_�´��Ժҽ���������Ժ).Value & ","
    strTemp = strTemp & "44," & chk(chk_ȫ����ֻ�����).Value & chk(chk_ȫ��ĸֻ�����).Value & ","
    strTemp = strTemp & "163," & chk(chk_��Ŀִ��ǰ�����շѻ����).Value & ","
    '74231,Ƚ����,2014-6-24,��Ŀ�����������շѻ�������
    strTemp = strTemp & "232," & chk(chk_��Ŀ�����������շѻ�������).Value & ","
    strTemp = strTemp & "171," & chk(chk_����δ�շѵ����ﻮ�۴�������).Value & ","
    strTemp = strTemp & "172," & chk(chk_����δ��˵ļ��˴�������).Value & ","
    strTemp = strTemp & "185," & cboPatiVerfy.ItemData(cboPatiVerfy.ListIndex) & ","

    'ˢ��Ҫ����������ĳ���
    With lst(lst_ˢ������)
        str��� = ""
        For i = 0 To .ListCount - 1
            str��� = str��� & IIF(.Selected(i), 1, 0)
        Next
    End With
    strTemp = strTemp & "46," & str��� & ","

    strTemp = strTemp & "55," & (cmb(cmb_���������Դ).ListIndex + 1) & ","
    strTemp = strTemp & "56," & IIF(chk(chk_���ﴦ����������).Value = 0, 0, ud(ud_���ﴦ����������).Value) & ","
    strTemp = strTemp & "58," & (cmb(cmb_δ�󵥾ݽ���).ListIndex) & ","
    strTemp = strTemp & "59," & (cmb(cmb_ҽ��������).ListIndex) & ","
    strTemp = strTemp & "60," & IIF(Val(txtMaxMoney.Text) = 0, "", Val(txtMaxMoney.Text)) & ","
    strTemp = strTemp & "61," & (cmb(cmb_���Ʊ���ģʽ).ListIndex) & ","
    strTemp = strTemp & "63," & chk(chk_סԺ�����Զ�����).Value & ","
    strTemp = strTemp & "64," & (cmb(cmb_ҩƷ�������).ListIndex) & ","
    strTemp = strTemp & "65," & cmb(cmb_�����������).ListIndex + 1 & cmb(cmb_סԺ�������).ListIndex + 1 & ","
    strTemp = strTemp & "66," & ud(ud_�Һ�ԤԼ����).Value & ","
    strTemp = strTemp & "68," & chk(chk_δ����������ֹ��ҩ).Value & ","
    strTemp = strTemp & "69," & chk(chk_ҩƷ�������ҽ��).Value & ","
    If chk(chk_�����Ǽ���Ч����).Value = 0 Then
        strTemp = strTemp & "70," & chk(chk_�����Ǽ���Ч����).Value & ","
    Else
        strTemp = strTemp & "70," & ud(ud_�����Ǽ���Ч����).Value & ","
    End If
    strTemp = strTemp & "71," & chk(chk_����ҽ��������Ч).Value & ","
    strTemp = strTemp & "72," & chk(chk_���������շ����).Value & ","
    strTemp = strTemp & "73," & chk(chk_��ȷ����ҩƷ����).Value & ","
    strTemp = strTemp & "75," & chk(chk_�⹺�����Ҫ�˲�).Value & ","
    strTemp = strTemp & "76," & chk(chk_ʱ��ҩƷֱ��ȷ���ۼ�).Value & ","

    With lst(lst_סԺ�������)
        str��� = ""
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                str��� = str��� & Chr(.ItemData(i))
            End If
        Next
    End With
    strTemp = strTemp & "80," & str��� & ","

    strTemp = strTemp & "81," & chk(chk_ִ�к��Զ���˻��۵�).Value & ","
    strTemp = strTemp & "84," & chk(chk_һ��������������Ŀ).Value & ","

    With lst(lst_���﷢�����)
        str��� = ""
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                str��� = str��� & Chr(.ItemData(i))
            End If
        Next
    End With
    strTemp = strTemp & "86," & str��� & ","

    strTemp = strTemp & "87," & (cmb(cmb_ҩƷ����ģʽ).ListIndex) & ","
    strTemp = strTemp & "92," & chk(chk_���������Զ�����).Value & ","
    strTemp = strTemp & "93," & chk(chk_������Ŀ���ܼ����ۿ�).Value & ","
    strTemp = strTemp & "96," & chk(chk_ҩƷ�ʱ�¿��ÿ��).Value & ","
    'strTemp = strTemp & "97," & CStr(IIF(opt�շ�Ʊ�����ɷ�ʽ(1).Value, 1, 0) + Val(chk(chk_��ִ�п��ҷֱ��ӡ).Value) * 10) & ","
    strTemp = strTemp & "98," & chk(chk_���ʱ����������۷���).Value & ","
    strTemp = strTemp & "99," & chk(chk_���ȷ������ȼ�).Value & ","
    strTemp = strTemp & "100," & chk(chk_���������ģʽ).Value & ","
    strTemp = strTemp & "126," & chk(chk_ʱ����ⰴ�ۿ�ǰ�ɹ��ۼӳ�����).Value & ","
    strTemp = strTemp & "143," & chk(chk_����ҽ����������������).Value & ","
    strTemp = strTemp & "144," & chk(chk_�շ���Ŀ��λ��������).Value & ","
    strTemp = strTemp & "145," & chk(chk_ÿ��סԺʹ����סԺ��).Value & ","
    strTemp = strTemp & "147," & Val(txtUD(ud_��ͯ����綨����).Text) & ","
    strTemp = strTemp & "148," & chk(chk_δ�շѴ�����ҩ).Value & ","
    strTemp = strTemp & "149," & (cmb(cmb_Ч����ʾ��ʽ).ListIndex) & ","
    strTemp = strTemp & "150," & (cmb(cmb_ҩƷ���������㷨).ListIndex) & ","
    strTemp = strTemp & "151," & chk(chk_�����˷���������).Value & ","
    'strTemp = strTemp & "152," & chk(chk_���￨�ظ�ʹ��).Value & ","
    strTemp = strTemp & "154," & (cmb(cmb_��Ժʱδ��ҩ��Ŀ���).ListIndex) & ","
    strTemp = strTemp & "155," & (cmb(cmb_ת��ʱδ��ҩ��Ŀ���).ListIndex) & ","
    strTemp = strTemp & "158," & Val(txtInputHours.Text) & ","
    strTemp = strTemp & "160," & IIF(opt����(1).Value, 1, 0) & ","
    strTemp = strTemp & "161," & chk(chk_����ҩ��).Value & ","
    strTemp = strTemp & "162," & chk(chk_�´�ҽ��ʱ��ʾ����).Value & ","
    strTemp = strTemp & "173," & chk(chk_�⹺�����Ҫ������Ǹ������ܽ��и���).Value & ","
    strTemp = strTemp & "174," & chk(chk_ҩƷ�ƿ���ȷ����).Value & ","
    strTemp = strTemp & "175," & chk(chk_ҩƷ������ȷ����).Value & ","
    strTemp = strTemp & "181," & chk(chk_ʱ�۷ֶμӳ����).Value & ","
    strTemp = strTemp & "182," & IIF(chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Value = 0, 1, 0) & ","
    strTemp = strTemp & "183," & chk(chk_ʱ��ҩƷȡ�ϴ��ۼ�).Value & ","
    strTemp = strTemp & "186," & chk(chk_��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�).Value & ","
    strTemp = strTemp & "187," & chk(chk����ҩ��ּ�����).Value & ","
    strTemp = strTemp & "188," & chk(chk����ҩ��ʹ���Ա�ҩ).Value & ","
    strTemp = strTemp & "189," & chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Value & ","
    strTemp = strTemp & "191," & chk(chkֻ����¼����).Value & ","
    '55791:������,2012-11-13,���˳�Ժҽ�����ܳ�����Ժ
    strTemp = strTemp & "192," & chk(chk_���˳�Ժҽ������������Ժ).Value & ","

    strTemp = strTemp & "208," & chk(chk�ٴ�����վ����ʹ��zlPlugIn����).Value & ","
    strTemp = strTemp & "209," & chk(chk���������ּ�����).Value & ","
    strTemp = strTemp & "210," & chk(chk_���������Һ���Ч�����Ĳ���).Value & ","
    strTemp = strTemp & "213," & (IIF(cmb(cmd_��ҩ�䷽).ListIndex = 1, 4, 3)) & ","

    strTemp = strTemp & "214," & chk(chk_�״�ҽ��ִ����Ҫ���).Value & ","
    '51612
    strTemp = strTemp & "215," & chk(chk_δ��ƽ�ֹ����).Value & ","
    strTemp = strTemp & "216," & chk(chk_��Ѫ�ּ�����).Value & ","
    strTemp = strTemp & "217," & chk(chk_������Ȩ����).Value & ","
    strTemp = strTemp & "218," & chk(chk_��Ѫ�����������).Value & ","
    strTemp = strTemp & "219," & chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Value & ","
    '�����޸�n���ڵǼǵ�ҽ��ִ�м�¼
    strTemp = strTemp & "220," & IIF(chk(chk_ҽ��ִ����Ч����).Value = 0, 999, Val(txtUNExecLimit.Text)) & ","
    strTemp = strTemp & "221," & IIF(optAccountTime(1).Value = True, 0, Val(txtAccountTime.Text)) & ","
    strTemp = strTemp & "223," & ud(ud_�����¿�ҽ�����).Value & ","
    '����������Դ������̫Ԫͨ������ҩ�ӿڲ�ʹ�á�
    strTemp = strTemp & "224," & IIF(cmb(cmb_������ҩ�ӿ�).ListIndex = 3, cmb(cmd_����������Դ).ListIndex, -1) & ","
    '���ô�ͨ������ҩ�ӿڲű���
    strTemp = strTemp & "225," & IIF(cmb(cmb_������ҩ�ӿ�).ListIndex = 2, chk(chk_���ýӿڵ�����־).Value, 0) & ","
    strTemp = strTemp & "226," & IIF(cmb(cmb_������ҩ�ӿ�).ListIndex = 1, chk(chk_����ʹ��ϵͳ����).Value, 1) & ","
    strTemp = strTemp & "227," & (cmb(cmd_ת��ʱδ������ʵ���).ListIndex) & ","
    If cmb(cmb_������ҩ�ӿ�).ListIndex = 1 Then
        strTemp = strTemp & "228," & IIF(optPASSVer(0).Value, "3.0", "4.0") & ","
    End If
    
    strTemp = strTemp & "230," & chk(chk_ҽ������ʱ��������ԭ��).Value & ","
    
    If chk(chk_ҽ������ʱ��������ԭ��).Value = 1 Then
        strTemp = strTemp & "233," & Get��д�������� & ","
    End If
    strTemp = strTemp & "234," & Getת�Ƴ�Ժ�������Ŀ & ","
    
    strTemp = strTemp & "235," & (cmb(cmd_��Ժʱ���ڻ�������).ListIndex) & ","
    strTemp = strTemp & "239," & chk(chk_�¿�ҽ��ǩ��ʱһ��ҽ��ǩ��һ��).Value & ","
    
    gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveRegister()
'���浽ע����е���Ϣ

End Sub

Private Sub Save�����ӿ�()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    For i = 1 To lvw����.ListItems.Count
        With lvw����.ListItems(i)
            gstrSQL = "Zl_����Ŀ¼_����(" & Mid(.Key, 2) & "," & IIF(.SubItems(4) <> "", 1, 0) & ")"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save����ǩ��()
    Dim i As Integer, j As Long
    Dim strDept As String
    
    On Error GoTo ErrHandle
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            strDept = ""
            For j = 1 To .Rows - 1
                If .Cell(flexcpChecked, j, col_ѡ��) = 1 Then
                    strDept = strDept & "," & .RowData(j)
                End If
            Next
            gstrSQL = "Zl_����ǩ�����ò���_Update(" & i & ",'" & Mid(strDept, 2) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End With
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Saveҽ������()
'����ҽ�����ݶ���
    On Error GoTo ErrHandle
    If cmdAdvice.Tag = "1" Then
        gstrSQL = "zl_ҽ�����ݶ���_Delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        mrsAdvice.Filter = 0
        Do While Not mrsAdvice.EOF
            If Not IsNull(mrsAdvice!ҽ������) Then
                gstrSQL = "zl_ҽ�����ݶ���_Insert('" & mrsAdvice!������� & "','" & Replace(mrsAdvice!ҽ������, "'", "''") & "')"
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

Private Sub Save���ݲ���()
    Dim lst As ListItem
    Dim i As Integer
    
    '����ɾ����ǰ�����е��ݲ���
    On Error GoTo ErrHandle
    gstrSQL = "zl_���ݲ�������_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�������µ�
    For Each lst In lvw(lvw_����).ListItems
        gstrSQL = "zl_���ݲ�������_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                    "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "��", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save�շ��ض���Ŀ()
    Dim strTemp As String
    
    '����Բ������б���
    On Error GoTo ErrHandle
    If txtCmd(0).Text <> "" Then
        strTemp = "������," & txtCmd(0).Tag & ","
    End If
    If txtCmd(1).Text <> "" Then
        strTemp = strTemp & "������," & txtCmd(1).Tag & ","
    End If
    
    If txtCmd(3).Text <> "" Then
        strTemp = strTemp & "��ͨ���÷�," & txtCmd(3).Tag & ","
    End If
    
    If txtCmd(4).Text <> "" Then
        strTemp = strTemp & "�������÷�," & txtCmd(4).Tag & ","
    End If
    
    If strTemp <> "" Then
        gstrSQL = "zl_�շ��ض���Ŀ_Modify('" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save�Զ��Ƽ���Ŀold()
    Dim str����ID As String
    Dim strϸĿID As String
    Dim str�����־ As String
    Dim str�������� As String
    Dim lngRow As Long
    Dim lngTemp As Long
    
    On Error GoTo ErrHandle
    With msh(0)
        For lngRow = 1 To .Rows - 1
            lngTemp = .RowData(lngRow)
            
            If lngTemp <> 0 Then
                If .TextMatrix(lngRow, 1) <> "" Then
                    str����ID = str����ID & lngTemp & ","
                    strϸĿID = strϸĿID & ","
                    str�����־ = str�����־ & "1,"
                    str�������� = str�������� & .TextMatrix(lngRow, 2) & ","
                End If
                If .TextMatrix(lngRow, 3) <> "" Then
                    str����ID = str����ID & lngTemp & ","
                    strϸĿID = strϸĿID & ","
                    str�����־ = str�����־ & "2,"
                    str�������� = str�������� & .TextMatrix(lngRow, 4) & ","
                End If
            End If
        Next
    End With
    With Bill(bill_�Զ�����)
        For lngRow = 1 To .Rows - 1
            lngTemp = .RowData(lngRow)
            
            If lngTemp <> 0 And .TextMatrix(lngRow, 1) <> "" Then
                str����ID = str����ID & lngTemp & ","
                strϸĿID = strϸĿID & .TextMatrix(lngRow, 1) & ","
                str�����־ = str�����־ & Switch(Left(.TextMatrix(lngRow, 3), 1) = "1", "6", Left(.TextMatrix(lngRow, 3), 1) = "2", "8", True, "7") & ","
                str�������� = str�������� & .TextMatrix(lngRow, 4) & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_�Զ��Ƽ���Ŀ_Modify('" & str����ID & "','" & strϸĿID & "','" & str�����־ & "','" & str�������� & "' )"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save�Զ��Ƽ���Ŀ()
    Dim str����ID As String
    Dim strϸĿID As String
    Dim str�����־ As String
    Dim str�������� As String
    Dim lngTemp As Long, i As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Zl_�Զ��Ƽ���Ŀ_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ���Զ��Ƽ���Ŀ")
    
    '����λ
    For i = 1 To msh(0).Rows - 1
        lngTemp = msh(0).RowData(i)
        If lngTemp <> 0 Then
            If msh(0).TextMatrix(i, 1) <> "" Then
                str����ID = str����ID & lngTemp & ","
                strϸĿID = strϸĿID & ","
                str�����־ = str�����־ & "1,"
                str�������� = str�������� & msh(0).TextMatrix(i, 2) & ","
            End If
            If msh(0).TextMatrix(i, 3) <> "" Then
                str����ID = str����ID & lngTemp & ","
                strϸĿID = strϸĿID & ","
                str�����־ = str�����־ & "2,"
                str�������� = str�������� & msh(0).TextMatrix(i, 4) & ","
            End If
        End If
        If (i Mod 100) = 0 Or i >= msh(0).Rows - 1 Then
            gstrSQL = "zl_�Զ��Ƽ���Ŀ_Modify('" & str����ID & "','" & strϸĿID & "','" & str�����־ & "','" & str�������� & "' )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            str����ID = ""
            strϸĿID = ""
            str�����־ = ""
            str�������� = ""
        End If
    Next
    '������
    For i = 1 To Bill(bill_�Զ�����).Rows - 1
        lngTemp = Bill(bill_�Զ�����).RowData(i)
        If lngTemp <> 0 And Bill(bill_�Զ�����).TextMatrix(i, 1) <> "" Then
            If Bill(bill_�Զ�����).TextMatrix(i, 1) <> "" Then
                str����ID = str����ID & lngTemp & ","
                strϸĿID = strϸĿID & Bill(bill_�Զ�����).TextMatrix(i, 1) & ","
                str�����־ = str�����־ & Switch(Left(Bill(bill_�Զ�����).TextMatrix(i, 3), 1) = "1", "6", Left(Bill(bill_�Զ�����).TextMatrix(i, 3), 1) = "2", "8", True, "7") & ","
                str�������� = str�������� & Bill(bill_�Զ�����).TextMatrix(i, 4) & ","
            End If
        End If
        If (i Mod 100) = 0 Or i >= Bill(bill_�Զ�����).Rows - 1 Then
            gstrSQL = "zl_�Զ��Ƽ���Ŀ_Modify('" & str����ID & "','" & strϸĿID & "','" & str�����־ & "','" & str�������� & "' )"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            str����ID = ""
            strϸĿID = ""
            str�����־ = ""
            str�������� = ""
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveҩƷ����()
    Dim strTemp As String
    Dim lngRow As Long
    Dim str���� As String
    
    On Error GoTo ErrHandle
    With Bill(bill_ҩƷ����)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                str���� = Left(.TextMatrix(lngRow, 3), 1)
                If str���� = "" Then str���� = "3"
                strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str���� & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_ҩƷ�������_Modify('" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume

    Call SaveErrLog
End Sub

Private Sub Save���ʱ�����()
    Dim strTemp As String
    Dim i As Integer
    Dim strArr
    Dim str���ò��� As String
    
    '�ȴ���ɾ�������ò��˼��ʱ���
    On Error GoTo ErrHandle
    If mstrDel���ò��� <> "" Then
        mstrDel���ò��� = mstrDel���ò��� & ";"
        strArr = Split(mstrDel���ò���, ";")
        For i = 0 To UBound(strArr) - 1
            If strArr(i) <> "" Then
                str���ò��� = strArr(i)
                strTemp = str���ò��� & "|"
                gstrSQL = "zl_���ʱ�����_Modify('" & strTemp & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End If
    
    '�����ò��˷�������
    mrsWarn.Filter = 0
    For i = 1 To tab����.Tabs.Count
        strTemp = ""
        str���ò��� = tab����.Tabs.Item(i).Caption
        
        mrsWarn.Filter = "���ò���='" & str���ò��� & "'"
        Do While Not mrsWarn.EOF
            strTemp = strTemp & Nvl(mrsWarn!����id) & "," & mrsWarn!�������� & "," & _
                mrsWarn!����ֵ & "," & Nvl(mrsWarn!������־1) & "," & Nvl(mrsWarn!������־2) & "," & Nvl(mrsWarn!������־3) & "," & Nvl(mrsWarn!�߿�����) & "," & Nvl(mrsWarn!�߿��׼) & ","
            mrsWarn.MoveNext
        Loop
        
        strTemp = str���ò��� & "|" & strTemp
        
        gstrSQL = "zl_���ʱ�����_Modify('" & strTemp & "')"
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
'        Case Index = chk_���չ����� '56963
'            fra(10).Enabled = (chk(chk_���չ�����).Value = 1)
            
        Case Index = chk_Ʊ�ſ���
            lvw(lvw_Ʊ��).SelectedItem.SubItems(2) = IIF(chk(Index).Value = 1, "��", "")
        Case Index = chk_�����շ��뷢ҩ����
            If chk(chk_�����շ��뷢ҩ����).Value <> 0 Then
                chk(chk_�շ�ͬʱ��ҩ).Enabled = False
                chk(chk_�շ�ͬʱ��ҩ).Value = 0
            Else
                chk(chk_�շ�ͬʱ��ҩ).Enabled = True
            End If
        Case Index = chk_�����Ǽ���Ч����
            If chk(Index).Value = 0 Then
                ud(ud_�����Ǽ���Ч����).Enabled = False
                txtUD(ud_�����Ǽ���Ч����).Enabled = False
                txtUD(ud_�����Ǽ���Ч����).BackColor = Me.BackColor
            Else
                ud(ud_�����Ǽ���Ч����).Enabled = True
                txtUD(ud_�����Ǽ���Ч����).Enabled = True
                txtUD(ud_�����Ǽ���Ч����).BackColor = RGB(255, 255, 255)
            End If
        Case Index = chk_���ﴦ����������
            If chk(Index).Value = 0 Then
                ud(ud_���ﴦ����������).Enabled = False
                txtUD(ud_���ﴦ����������).Enabled = False
                txtUD(ud_���ﴦ����������).BackColor = Me.BackColor
            Else
                ud(ud_���ﴦ����������).Enabled = True
                txtUD(ud_���ﴦ����������).Enabled = True
                txtUD(ud_���ﴦ����������).BackColor = RGB(255, 255, 255)
            End If
        Case Index = chk_ҩƷ�ʱ�¿��ÿ��
            If chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = 1 Then
                chk(chk_��ȷ����ҩƷ����).Value = 1
                chk(chk_ҩƷ�ƿ���ȷ����).Value = 1
                chk(chk_ҩƷ������ȷ����).Value = 1
            End If
        Case Index = chk_��ȷ����ҩƷ����
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = 1 Then
                chk(chk_��ȷ����ҩƷ����).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_ҩƷ�ƿ���ȷ����
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = 1 Then
                chk(chk_ҩƷ�ƿ���ȷ����).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_ҩƷ������ȷ����
            If blnIsChange Then
                Exit Sub
            End If
            If chk(chk_ҩƷ�ʱ�¿��ÿ��).Value = 1 Then
                chk(chk_ҩƷ������ȷ����).Value = 1
                blnIsChange = True
            End If
        Case Index = chk_����ǩ������_���� Or Index = chk_����ǩ������_סԺ _
                Or Index = chk_����ǩ������_ҽ�� Or Index = chk_����ǩ������_���� Or Index = chk_����ǩ������_ҩƷ _
                Or Index = chk_����ǩ������_lis Or Index = chk_����ǩ������_pacs
            '��ʹ�õ���ǩ��������£�������һ��������Ҫ����ǩ��
            If cmb(cmb_����ǩ����֤����).ListIndex <> 0 Then
                If chk(chk_����ǩ������_����).Value = 0 And chk(chk_����ǩ������_סԺ).Value = 0 _
                    And chk(chk_����ǩ������_ҽ��).Value = 0 And chk(chk_����ǩ������_����).Value = 0 And chk(chk_����ǩ������_ҩƷ).Value = 0 _
                    And chk(chk_����ǩ������_lis).Value = 0 And chk(chk_����ǩ������_pacs).Value = 0 Then
                        If Index = chk_����ǩ������_���� Then
                            chk(chk_����ǩ������_ҩƷ).Value = 1
                        ElseIf Index = chk_����ǩ������_ҩƷ Then
                             chk(chk_����ǩ������_lis).Value = 1
                        ElseIf Index = chk_����ǩ������_lis Then
                             chk(chk_����ǩ������_pacs).Value = 1
                        ElseIf Index = chk_����ǩ������_pacs Then
                             chk(chk_����ǩ������_����).Value = 1
                        Else
                            chk(((Index - chk_����ǩ������_���� + 1) Mod 4) + chk_����ǩ������_����).Value = 1
                        End If
                End If
            End If
            If Index = chk_����ǩ������_���� Then
                sstSign.TabVisible(sst_����) = chk(chk_����ǩ������_����).Value = 1
            ElseIf Index = chk_����ǩ������_ҩƷ Then
                 sstSign.TabVisible(sst_ҩƷ) = chk(chk_����ǩ������_ҩƷ).Value = 1
            ElseIf Index = chk_����ǩ������_lis Then
                 sstSign.TabVisible(sst_lis) = chk(chk_����ǩ������_lis).Value = 1
            ElseIf Index = chk_����ǩ������_pacs Then
                 sstSign.TabVisible(sst_Pacs) = chk(chk_����ǩ������_pacs).Value = 1
            ElseIf Index = chk_����ǩ������_���� Then
                sstSign.TabVisible(sst_����) = chk(chk_����ǩ������_����).Value = 1
            ElseIf Index = chk_����ǩ������_סԺ Then
                sstSign.TabVisible(sst_סԺ��ʿ) = chk(chk_����ǩ������_סԺ).Value = 1
                sstSign.TabVisible(sst_סԺҽ��) = chk(chk_����ǩ������_סԺ).Value = 1
            ElseIf Index = chk_����ǩ������_ҽ�� Then
                sstSign.TabVisible(sst_ҽ��) = chk(chk_����ǩ������_ҽ��).Value = 1
            End If
        Case Index = chk_���������շ���� And Visible
            If chk(Index).Value = 1 Then
                chk(chk_�շ���Ŀ��λ��������).Value = 0
            End If
        Case Index = chk_�շ���Ŀ��λ�������� And Visible
            If chk(Index).Value = 1 Then
                chk(chk_���������շ����).Value = 0
            End If
       Case Index = chk_��Ŀִ��ǰ�����շѻ����
            If chk(Index).Value = 1 Then
                chk(chk_δ�շѴ�����ҩ).Enabled = False
                chk(chk_����δ�շѵ����ﻮ�۴�������).Enabled = False
                
                chk(chk_δ��˼��ʴ�����ҩ).Caption = "����δ��˵ļ��ʴ�����ҩ(ֻ��סԺ��Ч)"
                chk(chk_����δ��˵ļ��˴�������).Caption = "����δ��˵ļ��˴�������(ֻ��סԺ��Ч)"
            
            Else
                chk(chk_δ�շѴ�����ҩ).Enabled = True
                chk(chk_δ��˼��ʴ�����ҩ).Caption = "����δ��˵ļ��ʴ�����ҩ"
                
                chk(chk_����δ��˵ļ��˴�������).Caption = "����δ�շѵ����ﻮ�۴�������"
                chk(chk_����δ�շѵ����ﻮ�۴�������).Enabled = True
            End If
        Case Index = chk_ʱ�۷ֶμӳ����
            If chk(Index).Value = 1 Then
                chk(chk_ʱ��ҩƷ���).Value = 0
                chk(chk_ʱ��ҩƷ���).Enabled = False
            Else
                chk(chk_ʱ��ҩƷ���).Value = 0
                chk(chk_ʱ��ҩƷ���).Enabled = True
            End If
        Case Index = chk����ҩ��ּ�����
            chk(chk����ҩ��ʹ���Ա�ҩ).Enabled = chk(Index).Value = 1
        Case Index = chk_����ҩ��
            If chk(Index).Value = 1 Then
                chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Value = 0
                chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Enabled = False
            Else
                chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Value = 0
                chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Enabled = True
            End If
        Case Index = chk���������ּ�����
            If chk(Index).Value = 1 Then
                chk(chk_������Ȩ����).Value = 0
                chk(chk_������Ȩ����).Enabled = True
            Else
                chk(chk_������Ȩ����).Value = 0
                chk(chk_������Ȩ����).Enabled = False
            End If
        Case Index = chk_��Ѫ�ּ�����
            If chk(Index).Value = 1 Then
                chk(chk_��Ѫ�����������).Value = 0
                chk(chk_��Ѫ�����������).Enabled = True
                chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Value = 0
                chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Enabled = True
            Else
                chk(chk_��Ѫ�����������).Value = 0
                chk(chk_��Ѫ�����������).Enabled = False
                chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Value = 0
                chk(chk_��Ѫ����ֻ�����м�������ҽʦ���).Enabled = False
            End If
        Case Index = chk_ҽ��ִ����Ч����
            txtUNExecLimit.Enabled = chk(Index).Value = 1
        Case Index = chk_ҽ������ʱ��������ԭ��
            Call Set��д��������(chk(Index).Value = 1)
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

    If Index = cmb_����ǩ����֤���� Then
        If cmb(Index).ListIndex = 0 Then
            chk(chk_����ǩ������_����).Value = 0
            chk(chk_����ǩ������_סԺ).Value = 0
            chk(chk_����ǩ������_ҽ��).Value = 0
            chk(chk_����ǩ������_����).Value = 0
            chk(chk_����ǩ������_ҩƷ).Value = 0
            chk(chk_����ǩ������_lis).Value = 0
            chk(chk_����ǩ������_pacs).Value = 0
            chk(chk_����ǩ������_����).Enabled = False
            chk(chk_����ǩ������_סԺ).Enabled = False
            chk(chk_����ǩ������_ҽ��).Enabled = False
            chk(chk_����ǩ������_����).Enabled = False
            chk(chk_����ǩ������_ҩƷ).Enabled = False
            chk(chk_����ǩ������_lis).Enabled = False
            chk(chk_����ǩ������_pacs).Enabled = False
            sstSign.Enabled = False
            sstSign.TabVisible(sst_����) = True
            txtFind.Enabled = False
            cmdFind.Enabled = False
        Else
            If Not chk(chk_����ǩ������_����).Enabled Then
                chk(chk_����ǩ������_����).Value = 1
            End If
            chk(chk_����ǩ������_����).Enabled = True
            chk(chk_����ǩ������_סԺ).Enabled = True
            chk(chk_����ǩ������_ҽ��).Enabled = True
            chk(chk_����ǩ������_����).Enabled = True
            chk(chk_����ǩ������_ҩƷ).Enabled = True
            chk(chk_����ǩ������_lis).Enabled = True
            chk(chk_����ǩ������_pacs).Enabled = True
            sstSign.Enabled = True
            txtFind.Enabled = True
            cmdFind.Enabled = True
        End If
    ElseIf Index = cmb_������ҩ�ӿ� Then
        '����ʱ�ɼ�
        lblPassVer.Visible = cmb(Index).ListIndex = 1
        optPASSVer(0).Visible = cmb(Index).ListIndex = 1
        optPASSVer(1).Visible = cmb(Index).ListIndex = 1
        optPASSVer(1).Enabled = False  '����4.0�����죬��ʱ����
            
        If cmb(Index).ListIndex = 0 Then    'δ���ýӿ�
            chk(chk_����ҩ��).Enabled = False
            chk(chk_����ҩ��).Value = 0
            chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = False
            chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Value = 0
            chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Enabled = False
            chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Value = 0

            chk(chk_���ýӿڵ�����־).Visible = False  '��ͨʱ�ɼ�
            chk(chk_����ʹ��ϵͳ����).Visible = False  '����ʱ�ɼ�
            '̫Ԫͨʱ�ɼ�
            cmb(cmd_����������Դ).Visible = False
            lblInfo(lbl_����������Դ).Visible = False
        Else
            chk(chk_����ҩ��).Enabled = True
            chk(chk�����´�Ժ��ִ�еĽ���ҩƷҽ��).Enabled = True

            If cmb(Index).ListIndex = 1 Then  '����
                chk(chk_����ʹ��ϵͳ����).Visible = True
                chk(chk_����ʹ��ϵͳ����).Enabled = True
            Else
                chk(chk_����ʹ��ϵͳ����).Visible = False
                chk(chk_����ʹ��ϵͳ����).Enabled = False
            End If

            If cmb(Index).ListIndex = 2 Then  '��ͨ
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = True
                chk(chk_���ýӿڵ�����־).Visible = True
            Else
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Enabled = False
                chk(chk_��ֹ�´ﳬ����ҩƷҽ��).Value = 0
                chk(chk_���ýӿڵ�����־).Visible = False
            End If
            If cmb(Index).ListIndex = 3 Then    '̫Ԫͨ
                cmb(cmd_����������Դ).ListIndex = 0
                cmb(cmd_����������Դ).Visible = True
                lblInfo(lbl_����������Դ).Visible = True
                cmb(cmd_����������Դ).Enabled = True
                lblInfo(lbl_����������Դ).Enabled = True
            Else
                cmb(cmd_����������Դ).Visible = False
                lblInfo(lbl_����������Դ).Visible = False
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

    '�����Ѻ͹������޶�Ϊ������Ŀ
    strSQL = "select id,����,����,���㵥λ,˵�� from �շ���ĿĿ¼ where ���='Z' and nvl(�Ƿ���,0)=0"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If IsNumeric(txtCmd(Index).Tag) = False Then txtCmd(Index).Tag = 0
        strSQL = frmSelCurr.ShowCurrSel(Me, rsTmp, "id,0,0,2;���,1000,0,2;����,1800,0,1;��λ,800,0,2;˵��,2300,0,2", -1, "������Ŀѡ��", , CStr(txtCmd(Index).Tag), 0, 3)
        If strSQL <> "" Then
            txtCmd(Index).Tag = CLng(Split(strSQL, ";")(0))
            txtCmd(Index).Text = Trim(Split(strSQL, ";")(2))
            txtCmd(Index).SetFocus
            mblnChange = True
        End If
    Else
        MsgBox "���κ���Ŀ���ã�", vbInformation, gstrSysName
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
    If Index < dtp_�����°� Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).Value
        If dtp(intNext).Value < dtp(intNext).MinDate Then
            dtp(intNext).Value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
End Sub

Private Sub lst���_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 And lst���.Selected(Item) Then
        For i = 1 To lst���.ListCount - 1
            lst���.Selected(i) = False
        Next
    ElseIf Item > 0 And lst���.Selected(Item) Then
        lst���.Selected(0) = False
    End If
End Sub

Private Sub lst���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call lst���_Validate(False)
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub lst���_LostFocus()
    lst���.Visible = False
End Sub

Private Sub lst���_Validate(Cancel As Boolean)
    Dim objGrid As Object, i As Integer
    
    Set objGrid = Bill(bill_���ʱ���)
    
    With objGrid
        .TextMatrix(.Row, .Col) = Get���ѡ��
        If .TextMatrix(.Row, .Col) = "�������" Then
            For i = 3 To 5
                If i <> .Col Then .TextMatrix(.Row, i) = " "
            Next
        End If
    End With
    mblnChange = True
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     If Index = lvw_���� Then
        If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
            lvw(lvw_����).SortOrder = IIF(lvw(lvw_����).SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            mintColumn = ColumnHeader.Index - 1
            lvw(lvw_����).SortKey = mintColumn
            lvw(lvw_����).SortOrder = lvwAscending
        End If
     End If
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Index = lvw_���� Then
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
    Dim lngԭֵ As Long
    
    If Index = lvw_Ʊ�� Then
        lngԭֵ = Val(Item.SubItems(1))
        ud(ud_���볤��).Max = 20
        
        '�������ֵʱ�������Ѿ��������б��е�ֵ
        ud(ud_���볤��).Value = lngԭֵ
        chk(chk_Ʊ�ſ���).Value = IIF(Item.SubItems(2) = "��", 1, 0)
    ElseIf Index = lvw_һ��ͨ Then
        cmdOneCard(1).Enabled = Item.Text <> ""
        cmdOneCard(2).Enabled = cmdOneCard(1).Enabled
    End If
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_Ʊ�� Then
        If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    ElseIf Index = lvw_���� Then
        If KeyAscii = vbKeyReturn Then cmdOperate_Click (1)
    End If
End Sub

Private Sub lvwCheckMed_DblClick()
    If Not Me.lvwCheckMed.SelectedItem Is Nothing Then
        lvwCheckMed.SelectedItem.SubItems(2) = Switch(lvwCheckMed.SelectedItem.SubItems(2) = "0-�����", "1-��飬��������", lvwCheckMed.SelectedItem.SubItems(2) = "1-��飬��������", "2-��飬�����ֹ", lvwCheckMed.SelectedItem.SubItems(2) = "2-��飬�����ֹ", "0-�����")
    End If
End Sub

Private Sub lvwCheckMed_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = "C" Then
        Call lvwCheckMed_DblClick
    End If
End Sub

Private Sub lvwNo_DblClick()
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    Call �ı����
End Sub

Private Sub lvwNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If lvwNo.SelectedItem Is Nothing Then Exit Sub
        Call �ı����
    End If
End Sub

Private Sub lvw����_DblClick()
    If Not lvw����.SelectedItem Is Nothing Then
        If lvw����.SelectedItem.SubItems(4) <> "" Then
            lvw����.SelectedItem.SubItems(4) = ""
        Else
            lvw����.SelectedItem.SubItems(4) = "��"
        End If
        lvw����.Tag = "1"
        Call lvw����_ItemClick(lvw����.SelectedItem)
    End If
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmd��������.Enabled = Item.SubItems(4) <> ""
End Sub

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call lvw����_DblClick
    End If
End Sub

Private Sub msf�ⷿ��λ_DblClick()
    Dim i As Long
    
    If msf�ⷿ��λ.Col > 1 And msf�ⷿ��λ.Row > 0 And Trim(msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, 0)) <> "" Then
        msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, 2) = ""
        msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, 3) = ""
        msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, 4) = ""
        msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, 5) = ""
        msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Row, msf�ⷿ��λ.Col) = "��"
    End If
End Sub

Private Sub msf�ⷿ��λ_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn Or KeyAscii = Asc(" ")) Then
        msf�ⷿ��λ_DblClick
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
                    MsgBox "��λ����Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "��������Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "��", "")
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
                    Bill(bill_�Զ�����).SetFocus
                Else
                    .Row = .Row + 1
                    .Col = 1
                    If .Row - .TopRow > 8 Then .TopRow = .Row - 8
                End If
            End If
        ElseIf KeyAscii = Asc(" ") Then
            If .Row > 0 And (.Col = 1 Or .Col = 3) And .RowData(.Row) <> 0 Then
                If .Col = 1 And Not mblnJRaiseByDate Then
                    MsgBox "��λ����Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If .Col = 3 And Not mblnHRaiseByDate Then
                    MsgBox "��������Ŀ���������Ŀ�ļ۸�������ǰ�����ִ�еģ����顣", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                .Text = IIF(.Text = "", "��", "")
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
'ͨ����ťѡ���շ�ϸĿ
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim rsTmp As New ADODB.Recordset
    
    If Index = bill_���ʱ��� Then
        With Bill(Index)
            Call Set���ѡ��(.TextMatrix(.Row, .Col))
            
            lst���.Left = .Left + .MsfObj.CellLeft
            If .Top + .MsfObj.CellTop + .MsfObj.CellHeight + lst���.Height <= .Container.Height Then
                lst���.Top = .Top + .MsfObj.CellTop + .MsfObj.CellHeight
            Else
                lst���.Top = .Top + .MsfObj.CellTop - lst���.Height - 30
            End If
            lst���.Width = .MsfObj.CellWidth
            lst���.ZOrder
            lst���.Visible = True
            lst���.SetFocus
        End With
    End If
    
    If Index = bill_�Զ����� Then
        With Bill(bill_�Զ�����)
            If .TextMatrix(.Row, 3) <> "2-����һ��" Then
                blnRe = frmChargeListSel.ShowTree(strID, str����, False)
            Else
                blnRe = frmChargeListSel.ShowTree(strID, str����, True)
            End If
            If blnRe And strID <> "" Then
                If .TextMatrix(.Row, 3) <> "2-����һ��" Then
                    If Not IsRaiseByDate(strID) Then
                        MsgBox "��Ŀ[" & str���� & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .SetFocus
                .TextMatrix(.Row, 1) = strID
                .TextMatrix(.Row, 2) = str����
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-��������"
                mblnChange = True
            End If
        End With
    End If
    
    If Index = bill_ҩƷ�������� Then
        gstrSQL = "Select Distinct Id,����,����,���� From ���ű� a,��������˵�� b " & _
                  "Where a.id = b.����id And b.�������� In('��ҩ����') " & _
                  "    and (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) " & _
                  "order by ���� "
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "��ҩ����")
        
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> 1 Then Exit Sub
        If rsTmp.EOF = True Then Exit Sub
        
        With Bill(bill_ҩƷ��������)
            .TextMatrix(.Row, 0) = rsTmp("����") & "-" & rsTmp("����")
            .RowData(.Row) = rsTmp("ID")
        End With
        
    End If
    
End Sub
Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If Index = bill_ҩƷ���� Then
        With Bill(bill_ҩƷ����)
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
            
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
        End With
    End If
    
    If Index = bill_ҩƷ�������� Then
        With Bill(bill_ҩƷ��������)
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
            If Index = bill_ҩƷ���� And .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                If Index <> bill_ҩƷ�������� Then
                    .RowData(.Row) = .ItemData(.ListIndex)
                End If
            End If
            If Index = bill_���ʱ��� Then
                If .TextMatrix(.Row, 1) = "" Then .TextMatrix(.Row, 1) = "1-�ۼƷ���"
            ElseIf Index = bill_ҩƷ���� Then
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
            End If
            If .Index = bill_ҩƷ�������� Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            End If
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'�������һ�еı仯
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_�Զ����� Then
        If .MouseCol <> 3 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "0"
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "�����������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "1-������"
            Case "1"
                .TextMatrix(.Row, .Col) = "2-����һ��"
            Case Else
                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                        MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "�����������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(.Row, .Col) = "0-��������"
        End Select
    ElseIf Index = bill_ҩƷ���� Then
        If .MouseCol <> .Cols - 1 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "1"
                .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
            Case "2"
                .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
            Case Else
                .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
        End Select
    ElseIf Index = bill_���ʱ��� Then
        If .MouseCol <> .Cols - 1 And .MouseCol <> 1 Then Exit Sub
        If .Col = 1 Then
            .TextMatrix(.Row, 1) = IIF(Left(.TextMatrix(.Row, 1), 1) = "1", "2-ÿ�շ���", "1-�ۼƷ���")
            If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                .TextMatrix(.Row, 4) = ""  'ÿ�շ����ޱ�����ʽ2
                
                'Ϊ��ÿ�շ��á�ʱ�ж�һ�½���Ϊ����
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
        If Index = bill_�Զ����� Then
            If .Col = 4 And KeyCode = vbKeyReturn Then
                If .Text <> "" And Not IsDate(.Text) Then
                    If Not IsDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)) Then
                        .Text = ""
                        MsgBox "��������ȷ�����ڸ�ʽ(yyyy-mm-dd����yyyymmdd)��", vbInformation, gstrSysName
                    Else
                        .Text = Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2)
                    End If
                    .TextMatrix(.Row, .Col) = .Text
                End If
            End If
                
            If .Col = 2 Then
                '�շ�ϸĿ��ֻ����س���
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '����һ���ؼ�
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    'ѡ���շ�ϸĿ
                    If IsRecord(.Text) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = .TextMatrix(.Row, 2)
                    If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "0-��������"
                    mblnChange = True
                End If
            End If
        End If
        
        If Index = bill_���ʱ��� Then
            If .Col = 2 Then
                '����ֵ��ֻ����س���
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TxtVisible = False Then
                    If .TextMatrix(.Row, 2) = "" Then
                        '����һ���ؼ�
                        zlCommFun.PressKey vbKeyTab
                    End If
                Else
                    '�ж�����ĺϷ���
                    .Text = Format(.Text, "##########0.00;-##########0.00;0.00;0,00")
                    mblnChange = True
                End If
            ElseIf .Col = 3 Then
                '��ֹ���뱨�����
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            ElseIf .Col = 6 Or .Col = 6 Then
                .Text = Format(.Text, "###0.00;-###.00;0.00;0,00")
                mblnChange = True
            End If
        End If
  
        
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
                              " And b.�������� In ('��ҩ����') and (a.���� Like '" & UCase(strTmp) & "%' or a.���� like '" & UCase(strTmp) & "%' or a.���� like '" & UCase(strTmp) & "%')"
                    
                    lmX = Me.Left + Me.tabMain.Left + Me.fraMain(9).Left + Me.Bill(bill_ҩƷ��������).Left
                    lmY = Me.Top + Me.tabMain.Top + Me.fraMain(9).Top + Me.Bill(bill_ҩƷ��������).Top + Me.Bill(bill_ҩƷ��������).RowHeight(.Row) + 350
                    Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "��ҩ����", , , , , , True, lmX, lmY, 300, , , True)
                    
                    If rsTmp Is Nothing Then Cancel = True: Exit Sub
                    If rsTmp.State <> 1 Then Cancel = True: Exit Sub
                    If rsTmp.EOF = True Then Cancel = True: Exit Sub
        
                    With Bill(bill_ҩƷ��������)
                        .TextMatrix(.Row, 0) = rsTmp("����") & "-" & rsTmp("����")
                        .Text = rsTmp("����") & "-" & rsTmp("����")
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
        If Index = bill_�Զ����� Then
            If .Col = 4 Then
                .TxtCheck = True
                .TextMask = "0123456789-"
            Else
                .TxtCheck = False
            End If
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "0"
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "1-������"
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-����һ��"
                            Case Else
                                If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                                    If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                        MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                    If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                        MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                        Exit Sub
                                    End If
                                End If
                                .TextMatrix(.Row, .Col) = "0-��������"
                        End Select
                        mblnChange = True
                    Case vbKey0
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "ҩƷ�����������Զ����㷽ʽ���ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "0-��������"
                        mblnChange = True
                    Case vbKey1
                        If Len(Trim(.TextMatrix(.Row, 1))) > 0 Then
                            If IsDrugOrStuff(.TextMatrix(.Row, 1)) Then
                                MsgBox "ҩƷ�����������Զ��������Ͳ��ܸı䡣", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If Not IsRaiseByDate(.TextMatrix(.Row, 1)) Then
                                MsgBox "��Ŀ[" & .TextMatrix(.Row, 2) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۺ���ѡ�������Զ����㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        .TextMatrix(.Row, .Col) = "1-������"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-����һ��"
                        mblnChange = True
                End Select
            End If
        ElseIf Index = bill_ҩƷ���� Then
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
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                        mblnChange = True
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                        mblnChange = True
                End Select
            End If
        ElseIf Index = bill_���ʱ��� Then
            .TxtCheck = False
            If .Col = 1 Then
                
                '�л���������
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                        End Select
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-�ۼƷ���"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-ÿ�շ���"
                        mblnChange = True
                End Select
                If InStr(.TextMatrix(.Row, 1), "ÿ�շ���") > 0 Then
                    .TextMatrix(.Row, 4) = ""  'ÿ�շ����ޱ�����ʽ2
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
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub mshBillEdit_EnterCell(Row As Long, Col As Long)
    With mshBillEdit
        Select Case .Col
            Case mGrdCol.����
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
            Case mGrdCol.����
'                If CheckNumberRule_Drug = True Then
'                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
'                        MsgBox "����������룡", vbOKOnly + vbInformation, gstrSysName
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
            Case mGrdCol.����
        End Select
    End With
End Sub

Private Sub mshBillEdit_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
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
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub mshBillEditStuff_EnterCell(Row As Long, Col As Long)
    With mshBillEditStuff
        Select Case .Col
            Case mGrdCol.����
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
            Case mGrdCol.����
'                If CheckNumberRule_Stuff = True Then
'                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
'                        MsgBox "����������룡", vbOKOnly + vbInformation, gstrSysName
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
            Case mGrdCol.����
        End Select
    End With
End Sub

Private Sub mshBillEditStuff_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
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

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub sstSign_Click(PreviousTab As Integer)
    mlngFindItem = 1
End Sub

Private Sub tab����_Click()
    Dim lngRow As Long
    
    mrsWarn.Filter = "���ò���='" & tab����.SelectedItem.Caption & "'"
    
    With Bill(bill_���ʱ���)
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
                .RowData(lngRow) = Nvl(mrsWarn!����id, 0)
                .TextMatrix(lngRow, 0) = IIF(IsNull(mrsWarn!����id), "*����*", mrsWarn!������ & "-" & mrsWarn!������)
                .TextMatrix(lngRow, 1) = IIF(mrsWarn!�������� = 1, "1-�ۼƷ���", "2-ÿ�շ���")
                .TextMatrix(lngRow, 2) = Format(mrsWarn!����ֵ, "##########0.00;-##########0.00;0.00;0.00")
                
                .TextMatrix(lngRow, 3) = Get������ƴ�(Nvl(mrsWarn!������־1), mrs���)
                .TextMatrix(lngRow, 4) = Get������ƴ�(Nvl(mrsWarn!������־2), mrs���)
                .TextMatrix(lngRow, 5) = Get������ƴ�(Nvl(mrsWarn!������־3), mrs���)
                .TextMatrix(lngRow, 6) = Format(mrsWarn!�߿�����, "###0.00;-###0.00;0.00;0.00")
                .TextMatrix(lngRow, 7) = Format(mrsWarn!�߿��׼, "###0.00;-###0.00;0.00;0.00")
                
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
                    MsgBox "��������ȷ�����ڸ�ʽ(yyyy-mm-dd����yyyymmdd)��", vbInformation, gstrSysName
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
        MsgBox "��¼��0-9999����ֵ��Χ��", vbInformation, gstrSysName
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
        If Index = ud_�����¿�ҽ����� Then
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
        MsgBox "��¼��0-999����ֵ��Χ��", vbInformation, gstrSysName
        Cancel = True
    Else
        txtUNExecLimit.Text = Val(txtUNExecLimit.Text)
    End If
End Sub

Private Sub ud_Change(Index As Integer)
    mblnChange = True
    '��̬�ı�Ʊ�ų���
    If Index = ud_���볤�� Then
        lvw(lvw_Ʊ��).SelectedItem.SubItems(1) = ud(ud_���볤��).Value
    End If
    If Index = ud_�����¿�ҽ����� Then
        txtUD(ud_�����¿�ҽ�����).Text = ud(ud_�����¿�ҽ�����).Value
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
        Case 1 '����
            cmb(cmb_סԺ�Ź���).SetFocus
        Case 2 '�ٴ�Ӧ��
            cmb(cmb_���Ʊ���ģʽ).SetFocus
        Case 4 'Ʊ�ݹ���
            If lvw(lvw_һ��ͨ).Enabled Then lvw(lvw_һ��ͨ).SetFocus
        Case 5 '�Զ�����
            mblnJRaiseByDate = IsRaiseByDate("J")
            mblnHRaiseByDate = IsRaiseByDate("H")
            msh(0).SetFocus
        Case 6 '���ʱ���
            tab����.SetFocus
        Case 7 'Ȩ��
            If chk(chk_���ƿ�����).Enabled Then chk(chk_���ƿ�����).SetFocus
        Case 8 '���ݲ���
            lvw(lvw_����).SetFocus
        Case 9 'ҩƷ����
            Bill(bill_ҩƷ����).SetFocus
        Case 10  '�ⷿ���
            Me.lvwCheckMed.SetFocus
        Case 11  'ҩƷ�ⷿ��λ
        Case 12 'ҩƷ��������
            Bill(bill_ҩƷ��������).SetFocus
        Case 13 '���ݱ������
        Case 14 '���ұ��
        Case 15 'ҩ����ҩ����
    End Select
End Sub

Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub

Private Function IsRecord(ByVal strFind As String) As Boolean
'����:�������������Ƿ�����Ч�����ݿ��б�ļ�¼
'����:strFind SQL��������
'����ֵ:��Ч����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    
    rsTemp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strFind, "'") > 0 Then
        MsgBox "�����˷Ƿ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    gstrSQL = "select distinct A.����,A.����,A.���,A.���㵥λ ,A.id from �շ�ϸĿ A,�շѱ��� B,�շ���� C " & _
         " where A.ID=B.�շ�ϸĿID and A.�Ƿ��� <> 1 and A.ĩ��=1 and  A.���=C.���� and  (A.���� like [1] or B.���� like [2] " & _
         " or  upper(B.����) like [2]) and " & Where����ʱ��("A")
          
    With Bill(bill_�Զ�����)
        If .TextMatrix(.Row, 3) <> "2-����һ��" Then
            gstrSQL = gstrSQL & " and C.���� Not In('4','5','6','7') "
        End If
    End With
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind & "%", "%" & UCase(strFind) & "%")
    
    If rsTemp.RecordCount < 1 Then Exit Function
    If rsTemp.RecordCount > 1 Then
        gstrSQL = ""
        gstrSQL = frmSelCurr.ShowCurrSel(Me, rsTemp, "����,1000,0,2;����,1800,0,1;���,2300,0,2;���㵥λ,1000,0,2;id,0,0,2", -1, "ѡ���շ�ϸĿ")
        If gstrSQL = "" Then
            Exit Function
        End If
        If Bill(bill_�Զ�����).TextMatrix(Bill(bill_�Զ�����).Row, 3) <> "2-����һ��" Then
            If Not IsRaiseByDate(Val(Split(gstrSQL, ";")(4))) Then
                MsgBox "��Ŀ[" & Split(gstrSQL, ";")(1) & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_�Զ�����)
            .TextMatrix(.Row, 1) = Split(gstrSQL, ";")(4) ' rsTemp("ID")
            .TextMatrix(.Row, 2) = Split(gstrSQL, ";")(1) 'rsTemp("����")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-��������"
            End If
        End With
    Else
        rsTemp.MoveFirst
        If Bill(bill_�Զ�����).TextMatrix(Bill(bill_�Զ�����).Row, 3) <> "2-����һ��" Then
            If Not IsRaiseByDate(Val(rsTemp!ID)) Then
                MsgBox "��Ŀ[" & rsTemp!���� & "]" & "���������Ŀ�ļ۸�������ǰ�����ִ�еģ������µ��ۡ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        With Bill(bill_�Զ�����)
            .TextMatrix(.Row, 1) = rsTemp("ID")
            .TextMatrix(.Row, 2) = rsTemp("����")
            If .TextMatrix(.Row, 3) = "" Then
                .TextMatrix(.Row, 3) = "0-��������"
            End If
        End With
    End If
    IsRecord = True
End Function

Private Function NumIsValid(ByVal lngIndex As Long, ByVal strNumber As String) As Boolean
'����:�������������Ƿ���һ����Ч������
'����:strNumber  ��������
'����ֵ:��Ч����True,����ΪFalse
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "������һ����ֵ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "�����̫���ˡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�����޸ı��:2500
    'ֻ����ҽ�����˺ͷ�ҽ���˵ı��
    If (lngIndex = 1 Or lngIndex = 2) And Left(Bill(lngIndex).TextMatrix(Bill(lngIndex).Row, 1), 1) = "1" Then
        If Val(strNumber) < -9999999999.999 Then
            MsgBox "�����̫С�ˡ�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If Val(strNumber) < 0 Then
            MsgBox "����Ϊ������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    NumIsValid = True
End Function

Private Sub Set���ѡ��(str��� As String)
'���ܣ���������"���,����..."�Ĵ������б��ѡ�����
    Dim i As Integer, j As Integer
    Dim arr���() As String
    
    For i = 0 To lst���.ListCount - 1
        lst���.Selected(i) = False
    Next
    
    If Trim(str���) = "" Then
        Exit Sub
    ElseIf str��� = "�������" Then
        For i = 0 To lst���.ListCount - 1
            lst���.Selected(i) = (i = 0)
        Next
    Else
        lst���.Selected(0) = False
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    lst���.Selected(j) = True: Exit For
                End If
            Next
        Next
    End If
    
    For i = 0 To lst���.ListCount - 1
        If lst���.Selected(i) Then
            lst���.TopIndex = i: Exit For
        End If
    Next
End Sub

Private Function Get���ѡ��() As String
'���ܣ��������ѡ���ѡ��������������"���,����..."�Ĵ�
    Dim i As Integer, strTmp As String
    
    If lst���.Selected(0) Then
        Get���ѡ�� = "�������"
    Else
        For i = 1 To lst���.ListCount - 1
            If lst���.Selected(i) Then
                strTmp = strTmp & "," & lst���.List(i)
            End If
        Next
        Get���ѡ�� = Mid(strTmp, 2)
        If Get���ѡ�� = "" Then Get���ѡ�� = " " 'Ϊ���ܻس�������
    End If
End Function

Private Function Get������ƴ�(str��� As String, rs��� As ADODB.Recordset) As String
'���ܣ�������"CDEFG"�����ת��Ϊ����"���,����..."��
    Dim i As Integer, strTmp As String
    
    If str��� = "" Then
        Get������ƴ� = " " 'Ϊ���ܰ��س�������
        Exit Function
    End If
    
    If str��� = "-" Then
        Get������ƴ� = "�������"
        Exit Function
    End If
    
    For i = 1 To Len(str���)
        rs���.Filter = "����='" & Mid(str���, i, 1) & "'"
        If Not rs���.EOF Then strTmp = strTmp & "," & rs���!���
    Next
    Get������ƴ� = Mid(strTmp, 2)
End Function

Private Function Get�����봮(str��� As String) As String
'���ܣ���������"���,����"�Ĵ���������"CDEFG"�Ĵ�
    Dim i As Integer, j As Integer
    Dim arr���() As String, strTmp As String
    
    If Trim(str���) = "" Then Exit Function
    
    If str��� = "�������" Then
        Get�����봮 = "-"
    Else
        arr��� = Split(str���, ",")
        For i = 0 To UBound(arr���)
            For j = 1 To lst���.ListCount - 1
                If lst���.List(j) = arr���(i) Then
                    strTmp = strTmp & Chr(lst���.ItemData(j))
                    Exit For
                End If
            Next
        Next
        Get�����봮 = strTmp
    End If
End Function

Sub LoadҩƷ��������()
    '''''''''''''''''''''''''''''''''''''''''
    '����           ����ҩƷ���ò���
    '''''''''''''''''''''''''''''''''''''''''
    
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_ҩƷ��������)
        'װ�������������
        gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') " & _
                   " and  b.����ID=a.ID and " & Where����ʱ��("A") & " order by ����"
        Call OpenRecordset(rsTemp, Me.Caption)
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
        Call OpenRecordset(rsTemp, Me.Caption)
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
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveҩƷ��������()
    Dim strTemp As String
    Dim lngRow As Long
    Dim bln���� As Boolean
    
    On Error GoTo ErrHand
    With Bill(bill_ҩƷ��������)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                If LenB(StrConv(strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ",", vbFromUnicode)) >= 4000 Then
                    If bln���� = True Then
                        gstrSQL = "zl_ҩƷ�����������_Modify('" & strTemp & "'," & 1 & ")"
                    Else
                        gstrSQL = "zl_ҩƷ�����������_Modify('" & strTemp & "'," & 0 & ")"
                    End If
                    bln���� = True
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    strTemp = .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                Else
                    strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & ","
                End If
            End If
        Next
    End With
    
    If bln���� = True Then
        gstrSQL = "zl_ҩƷ�����������_Modify('" & strTemp & "'," & 1 & ")"
    Else
        gstrSQL = "zl_ҩƷ�����������_Modify('" & strTemp & "'," & 0 & ")"
    End If
    bln���� = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    Call SaveErrLog
    End If
End Sub

Private Function Load���ݱ������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lst As ListItem
    
    gstrSQL = "" & _
        "   Select ��Ŀ���,��Ŀ����,��Ź���,decode(��Ź���,2,'2-��ִ�п��ҷ��±��',0,'0-����˳����',1,'1-����˳����','0-����˳����') as ��Ź���˵�� " & _
        "   From ������Ʊ� " & _
        "   where ��Ŀ��� in ( 11,12,13,14,15,16,21,22,23,24,25,26,27,28,29,32,62,68,69,70,71,72,73,74,75,76,77) order by ��Ŀ��� "
    
    Err = 0: On Error GoTo ErrHand:
    Load���ݱ������ = False
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    With rsTmp
        lvwNo.ListItems.Clear
        Do While Not rsTmp.EOF
            Set lst = lvwNo.ListItems.Add(, "K" & Nvl(!��Ŀ���, 0), Nvl(!��Ŀ����))
            lst.SubItems(1) = Nvl(!��Ź���˵��)
            If Nvl(!��Ŀ���) >= 1 And Nvl(!��Ŀ���) <= 16 Then
                lst.ForeColor = &HC85422
                lvwNo.ListItems("K" & Nvl(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &HC85422
            End If
            If Nvl(!��Ŀ���) >= 21 And Nvl(!��Ŀ���) <= 62 Then
                lst.ForeColor = &H68588
                lvwNo.ListItems("K" & Nvl(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &H68588
            End If
            If Nvl(!��Ŀ���) >= 68 And Nvl(!��Ŀ���) <= 77 Then
                lst.ForeColor = &H856701
                lvwNo.ListItems("K" & Nvl(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &H856701
            End If
            lst.Tag = Nvl(!��Ź���, 0)
            If lvwNo.SelectedItem Is Nothing Then
                lst.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '2-סԺ�ţ�3-�����
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select ��Ŀ���,��Ź��� as ����ֵ From ������Ʊ� Where ��Ŀ��� in (2,3)"
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    rsTmp.Filter = "��Ŀ���=2"
    If rsTmp.RecordCount > 0 Then cmb(cmb_סԺ�Ź���).ListIndex = Val("" & rsTmp!����ֵ)
    rsTmp.Filter = "��Ŀ���=3"
    If rsTmp.RecordCount > 0 Then cmb(cmb_����Ź���).ListIndex = Val("" & rsTmp!����ֵ)
    
    Load���ݱ������ = True
    Call SetEdit
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SetEdit
    
End Function
Private Function SetEdit()
    '����:���ñ༭����
    Dim blnEdit As Boolean
    Dim blnData As Boolean
End Function
Private Sub �ı����()
    '�ı�������
    Dim strNo As String
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    strNo = lvwNo.SelectedItem.SubItems(1) & "-"
    Select Case Split(strNo, "-")(0)
        Case 0
            If Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16 Then
                strNo = "1-����˳����"
                lvwNo.SelectedItem.Tag = "1"
            Else
                strNo = "2-��ִ�п��ҷ��±��"
                lvwNo.SelectedItem.Tag = "2"
            End If
        Case 1
            strNo = "0-����˳����"
            lvwNo.SelectedItem.Tag = "0"
        Case 2
            strNo = "0-����˳����"
            lvwNo.SelectedItem.Tag = "0"
    End Select
    lvwNo.SelectedItem.SubItems(1) = strNo
        
'    If Split(StrNo, "-")(0) = 2 Then
'        StrNo = "1-����˳����"
'        lvwNo.SelectedItem.Tag = "0"
'    Else
'        StrNo = "2-��ִ�п��ҷ��±��"
'        lvwNo.SelectedItem.Tag = "2"
'    End If
'    lvwNo.SelectedItem.SubItems(1) = StrNo
End Sub
Sub Save���ݱ������()
    Dim lst As ListItem
    
    On Error GoTo ErrHandle
    For Each lst In lvwNo.ListItems
        gstrSQL = "ZL_������Ʊ�_Rule(" & Mid(lst.Key, 2) & "," & Val(lst.Tag) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Next
    
    '2-סԺ��,3-�����
    gstrSQL = "ZL_������Ʊ�_Rule(2," & cmb(cmb_סԺ�Ź���).ListIndex & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    gstrSQL = "ZL_������Ʊ�_Rule(3," & cmb(cmb_����Ź���).ListIndex & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub InitFace()
    '��ʼ���ؼ�
    With mshBillEdit
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.ѡ��) = "ѡ��"
        .TextMatrix(0, mGrdCol.����) = "����"
        .TextMatrix(0, mGrdCol.����) = "����"

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mGrdCol.ѡ��) = 5
        .ColData(mGrdCol.����) = 5
        .ColData(mGrdCol.����) = 4


        .ColWidth(mGrdCol.ѡ��) = 0
        .ColWidth(mGrdCol.����) = 2000
        .ColWidth(mGrdCol.����) = 1600
        
        .ColAlignment(mGrdCol.����) = 1
        
        .PrimaryCol = mGrdCol.����
        .LocateCol = mGrdCol.����
        .AllowAddRow = False
        .Active = True
    End With
    
    With mshBillEditStuff
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.ѡ��) = "ѡ��"
        .TextMatrix(0, mGrdCol.����) = "����"
        .TextMatrix(0, mGrdCol.����) = "����"

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mGrdCol.ѡ��) = 5
        .ColData(mGrdCol.����) = 5
        .ColData(mGrdCol.����) = 4


        .ColWidth(mGrdCol.ѡ��) = 0
        .ColWidth(mGrdCol.����) = 2000
        .ColWidth(mGrdCol.����) = 1600
        
        .ColAlignment(mGrdCol.����) = 1
        
        .PrimaryCol = mGrdCol.����
        .LocateCol = mGrdCol.����
        .AllowAddRow = False
        .Active = True
    End With
End Sub
Sub Save����()
    '������ұ��
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With mshBillEdit
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then  'And Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then
                '����ID_IN   IN ���ұ��.����ID%TYPE,
                '���_IN     IN ���ұ��.���%TYPE
                gstrSQL = "ZL_���Һ����_UPDATE("
                gstrSQL = gstrSQL & .RowData(i) & ","
                gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.����)) & "',1)"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With

    With mshBillEditStuff
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then 'And Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then
                '����ID_IN   IN ���ұ��.����ID%TYPE,
                '���_IN     IN ���ұ��.���%TYPE
                gstrSQL = "ZL_���Һ����_UPDATE("
                gstrSQL = gstrSQL & .RowData(i) & ","
                gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.����)) & "',2)"
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
    '����       ��鵥�ݱ�������Ƿ���"2"��
    '����       ��=True ��=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If .ListItems(i).SubItems(1) = "2-��ִ�п��ҷ��±��" Then
                CheckNumberRule = True
                Exit For
            End If
        Next
    End With
    'Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16
End Function

Function CheckNumberRule_Drug() As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '����       ��鵥�ݱ�������Ƿ���"2"��
    '����       ��=True ��=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 21 And Mid(.ListItems(i).Key, 2) <= 62 Then
                If .ListItems(i).SubItems(1) = "2-��ִ�п��ҷ��±��" Then
                    CheckNumberRule_Drug = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Function CheckNumberRule_Stuff() As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '����       ��鵥�ݱ�������Ƿ���"2"��
    '����       ��=True ��=False
    '''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 68 And Mid(.ListItems(i).Key, 2) <= 77 Then
                If .ListItems(i).SubItems(1) = "2-��ִ�п��ҷ��±��" Then
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
    If Col <> col_ѡ�� Then Cancel = True
End Sub

Private Sub vsDept_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    If Col = col_ѡ�� Then
        Order = 0
        With vsDept(Index)
            If .MouseCol = col_ѡ�� And .MouseRow = .FixedRows - 1 Then
                If sstSign.Enabled = False Then Exit Sub
                If .ColData(col_ѡ��) = "Check" Then
                    .Cell(flexcpPicture, 0, col_ѡ��) = ils16.ListImages("UnCheck").Picture
                    .ColData(col_ѡ��) = ""
                Else
                    .Cell(flexcpPicture, 0, col_ѡ��) = ils16.ListImages("AllCheck").Picture
                    .ColData(col_ѡ��) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .ColData(col_ѡ��) = "Check" Then
                        .Cell(flexcpChecked, i, col_ѡ��) = 1
                    Else
                        .Cell(flexcpChecked, i, col_ѡ��) = 2
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
            vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_ѡ��) = IIF(vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_ѡ��) = 1, 2, 1)
        End If
    End If
End Sub

Private Sub vsfControlItem_DblClick()
    With vsfControlItem
        If .Row < 1 Then Exit Sub
        If .Col < 2 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "��" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            '�˲�ʱ�����޸�"��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���"
            If .TextMatrix(.Row, 1) = "�˲�" And InStr(1, "��Ʊ��,��Ʊ����,��Ʊ����,��Ʊ���", .TextMatrix(0, .Col)) > 0 Then Exit Sub
            
            '�����⹺�����ѡ��
            If .TextMatrix(.Row, 0) = "�����⹺" And .TextMatrix(0, .Col) = "���" Then Exit Sub
            
            .TextMatrix(.Row, .Col) = "��"

        End If
    End With
End Sub

Private Sub Init�����˵��(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    vsUnWriteDept.Clear
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,���� from ���ű� where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnWriteDept
        lngRow = (rsTmp.RecordCount + 3) \ 4
        If lngRow > 5 Then .Rows = lngRow
        
        For i = 1 To rsTmp.RecordCount
            Call mcol����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
                lngRow = (i - 1) \ 4
                lngCol = (i - 1) Mod 4
                .TextMatrix(lngRow, lngCol) = rsTmp!����
                .Cell(flexcpData, lngRow, lngCol) = rsTmp!���� & ""
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

Private Sub Initת�Ƴ�Ժ�������Ŀ(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,���� from ������ĿĿ¼ where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnCheckItem
        .Row = 0: .Col = 0
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(.Row, .Col) = rsTmp!���� & ""
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


Private Function Get��д��������() As String
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
    Get��д�������� = Mid(strIDs, 2)
End Function

Private Function Getת�Ƴ�Ժ�������Ŀ() As String
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
    Getת�Ƴ�Ժ�������Ŀ = Mid(strIDs, 2)
End Function

Private Sub Set��д��������(ByVal blnEdit As Boolean)
'���ܣ��ɲ�¼�볬��ԭ��Ŀ��ң���񣩿�����
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
    strSQL = "select A.ID,A.����,A.���� from ������ĿĿ¼ A Where A.��� not in('4','5','6','7') and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order By ����"
    With vsUnCheckItem
        vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "������Ŀ", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetItemInput(Row, Col, rsTmp)
            Call vsUnCheckItem_AfterRowColChange(-1, -1, Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "û�п��õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
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
            '���ֱ�����뺺�ֵ�����
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
                .ComboList = "" 'ʹ��ť״̬��������״̬
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
        strSQL = "select DISTINCT A.ID,A.����,A.���� from ������ĿĿ¼ A, ������Ŀ���� B where " & _
            " a.Id = b.������Ŀid And B.����=1 And B.����=1 And A.��� not in('4','5','6','7') and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(B.����) Like [2])" & _
            " Order by A.����"
        With vsUnCheckItem
            vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, _
                vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                Call SetItemInput(Row, Col, rsTmp)
                .EditText = .TextMatrix(Row, Col)
                mblnChange = True
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnCheckItem_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
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
    strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
        " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order by A.����"
    With vsUnWriteDept
        vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "�ٴ�����", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsUnWriteDept_AfterRowColChange(-1, -1, Row, Col)
        Else
            If Not blnCancel Then
                MsgBox "û�п��õĿ������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
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
            strSQL = "select Distinct a.id,a.����,a.����,a.���� from ���ű� a,��������˵�� b where a.id=b.����id" & _
                " and b.��������='�ٴ�' And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
                " Order by A.����"
            With vsUnWriteDept
                vPoint = GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�ٴ�����", False, "", "", False, False, True, _
                    vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
                If Not rsTmp Is Nothing Then
                    Call SetDeptInput(Row, Col, rsTmp)
                    .EditText = .TextMatrix(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                End If
            End With
            Call vsUnWriteDept_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        End If
    End With
End Sub

Private Sub vsUnWriteDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
    If vsUnWriteDept.TextMatrix(Row, Col + 4) = "" Then vsUnWriteDept.TextMatrix(Row, Col) = ""
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset)
    '�ȼ���±�����Ƿ����
    Dim strTmp As String
    With vsUnWriteDept
        On Error Resume Next
        strTmp = mcol����("_" & rsTmp!ID)
        If Err.Number = 0 Then
            MsgBox "�ÿ����Ѿ����ڣ����������롣", vbInformation, gstrSysName
            .TextMatrix(lngRow, lngCol) = CStr(.Cell(flexcpData, lngRow, lngCol))
            Exit Sub
        Else
            Err.Clear
        End If
        On Error GoTo 0
        
        If .TextMatrix(lngRow, lngCol + 4) <> "" Then
            Call mcol����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
        End If
        
        .TextMatrix(lngRow, lngCol) = rsTmp!���� & ""
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!���� & ""
        .TextMatrix(lngRow, lngCol + 4) = rsTmp!ID & ""
        Call mcol����.Add(rsTmp!ID & "", "_" & rsTmp!ID)
    End With
End Sub

Private Function SetItemInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset)
    '�ȼ���±�����Ƿ����
    Dim i As Long, j As Long
    
    With vsUnCheckItem
        For i = .FixedCols To .Cols - 1
            For j = .FixedRows To .Rows - 1
                If .Cell(flexcpData, j, i) = rsTmp!ID & "" Then
                    MsgBox "��������Ŀ�Ѿ������б��У���鿴��", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
        
        .TextMatrix(lngRow, lngCol) = rsTmp!���� & ""
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
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsUnWriteDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    With vsUnWriteDept
        If KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsUnWriteDept_KeyPress(KeyCode)
        ElseIf KeyCode = vbKeyDelete Then
            Call mcol����.Remove("_" & Val(.TextMatrix(.Row, .Col + 4)))
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
'���ܣ����λ����һ��
    With vsobj
        If .Col + 1 > .Cols - 1 Then
            If .Row + 1 > .Rows - 1 Then .AddItem ""
            .Row = .Row + 1: .Col = .FixedCols
        Else
            .Col = .Col + 1
        End If
        '�������������ݹ��ٶ�λ����һ��λ��
        If .ColHidden(.Col) = True Then Call EnterNextCell(vsobj)
        .ShowCell .Row, .Col
    End With
End Sub
