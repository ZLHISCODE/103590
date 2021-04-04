VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISManage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "电子病历访问授权"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18780
   Icon            =   "frmCISManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   18780
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   5160
      ScaleHeight     =   3495
      ScaleWidth      =   3255
      TabIndex        =   8
      Top             =   3480
      Width           =   3255
      Begin VB.Frame fraFillter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "查询过滤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   17055
         Begin VB.CommandButton cmdFind 
            Caption         =   "查询(&F)"
            Height          =   375
            Left            =   12120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "待审批"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   7320
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   13
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已审批"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   8565
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   14
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已作废"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9795
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已拒绝"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   11040
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   277
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   11
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   5010
            TabIndex        =   26
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   2
            Left            =   225
            Picture         =   "frmCISManage.frx":6852
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "申请时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   32
            Top             =   330
            Width           =   855
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   7020
            Picture         =   "frmCISManage.frx":6DDC
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   8250
            Picture         =   "frmCISManage.frx":D62E
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   9495
            Picture         =   "frmCISManage.frx":13E80
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   10725
            Picture         =   "frmCISManage.frx":1A6D2
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.Frame picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "授权信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   12840
         TabIndex        =   25
         Top             =   840
         Width           =   4095
         Begin VSFlex8Ctl.VSFlexGrid vsInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4750
            _cx             =   8378
            _cy             =   12197
            Appearance      =   0
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
            MouseIcon       =   "frmCISManage.frx":20F24
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISManage.frx":217FE
            ScrollTrack     =   -1  'True
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
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   7275
         Left            =   0
         TabIndex        =   18
         Top             =   840
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":21827
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":22101
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   33
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   1800
      ScaleHeight     =   3375
      ScaleWidth      =   3255
      TabIndex        =   27
      Top             =   3480
      Width           =   3255
      Begin VB.Frame fraLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "查询过滤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   17055
         Begin VB.ComboBox cboLogTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   277
            Width           =   1365
         End
         Begin VB.CommandButton cmdLogFind 
            Caption         =   "查询(&F)"
            Height          =   375
            Left            =   7200
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpLogTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   21
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpLogTime 
            Height          =   300
            Index           =   1
            Left            =   5040
            TabIndex        =   22
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   0
            Left            =   240
            Picture         =   "frmCISManage.frx":2219C
            Top             =   300
            Width           =   240
         End
         Begin VB.Line LineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            Index           =   1
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "访问时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   36
            Top             =   330
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLog 
         Height          =   7275
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":22726
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":23000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   37
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picManage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   9120
      ScaleHeight     =   4215
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   3600
      Width           =   4455
      Begin VB.Frame fraManageInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "授权信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   13080
         TabIndex        =   30
         Top             =   960
         Width           =   4095
         Begin VSFlex8Ctl.VSFlexGrid vsManageInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4750
            _cx             =   8378
            _cy             =   12197
            Appearance      =   0
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
            MouseIcon       =   "frmCISManage.frx":2309B
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISManage.frx":23975
            ScrollTrack     =   -1  'True
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
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.Frame fraManageFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "查询过滤"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   17055
         Begin VB.CheckBox chk已作废 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "显示已作废的授权"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   7515
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdManageFind 
            Caption         =   "查询(&F)"
            Height          =   375
            Left            =   9480
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboManageTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   277
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpManageTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   1
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpManageTime 
            Height          =   300
            Index           =   1
            Left            =   5040
            TabIndex        =   2
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   1
            Left            =   200
            Picture         =   "frmCISManage.frx":2399E
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   7200
            Picture         =   "frmCISManage.frx":23F28
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "授权时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   29
            Top             =   330
            Width           =   855
         End
         Begin VB.Line LineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            Index           =   0
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsManage 
         Height          =   7275
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":2A77A
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":2B054
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   34
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5580
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   8130
      _Version        =   589884
      _ExtentX        =   14340
      _ExtentY        =   9842
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   10560
      Width           =   18780
      _ExtentX        =   33126
      _ExtentY        =   635
      SimpleText      =   $"frmCISManage.frx":2B0EF
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCISManage.frx":2B136
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   28046
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":2B9CA
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3222C
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":38A8E
            Key             =   "访问时限"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":39028
            Key             =   "访问内容"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":395C2
            Key             =   "访问医生"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":39B5C
            Key             =   "访问病人"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3A0F6
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3A250
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   600
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCISManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum colList
    COL_申请ID = 1
    COL_访问内容 = 2
    COL_内容时限 = 3
    COL_撤消时间 = 4
    COL_撤消人 = 5

    COL_申请时间 = 6
    COL_申请人 = 7
    COL_申请访问病人 = 8
    COL_访问开始时间 = 9
    COL_访问结束时间 = 10
    COL_申请原因 = 11
    COL_审批状态 = 12
End Enum

Private Enum colManage
    '隐藏列
    COLM_ID = 1
    COLM_访问内容 = 2
    COLM_内容时限 = 3 '针对住院病历：0-无限制，1-未归档病历，2-已归档病历；门诊病历无限制'
    COLM_授权类型 = 4 '0-审批申请,1-主动授权
    COLM_访问病人 = 5 '0-全院病人，1-本科病人，2-指定科室病人，3-指定病人，4-诊断为指定疾病的病人，5-指定手术的病人
    COLM_病人范围详情 = 6
    '显示列
    COLM_方案名 = 7
    COLM_备注 = 8
    COLM_访问开始时间 = 9
    COLM_访问结束时间 = 10
    COLM_授权人 = 11
    COLM_授权时间 = 12
    COLM_作废人 = 13
    COLM_作废时间 = 14
    COLM_访问者 = 15
End Enum


Private Enum colLog
    '隐藏列
    COLG_ID = 1
    COLG_病人ID = 2
    COLG_就诊ID = 3 '门诊为挂号ID，住院为主页ID';
    COLG_病人来源 = 4 '1-门诊病人，2-住院病人
    COLG_内容ID = 5  '内容ID中记录对应的业务文件标识ID
    
    '显示列
    COLG_访问时间 = 6
    COLG_访问人 = 7
    COLG_病人姓名 = 8
    COLG_病人性别 = 9
    COLG_病人年龄 = 10
    COLG_病人标识号 = 11
    COLG_病人科室 = 12
    COLG_病人类型 = 13
    COLG_访问内容 = 14
End Enum



Private Enum RowInfo
    Row_访问病人标题 = 0
    Row_访问病人 = 1
    Row_内容时限标题 = 3
    Row_内容时限 = 4
    Row_访问内容标题 = 6
    Row_访问内容 = 7
End Enum

Private Enum RowMInfo
    RowM_访问者标题 = 0
    RowM_访问者 = 1
    RowM_访问病人标题 = 3
    RowM_访问病人 = 4
    RowM_内容时限标题 = 6
    RowM_内容时限 = 7
    RowM_访问内容标题 = 9
    RowM_访问内容 = 10
End Enum



Private Sub cboManageTime_Click()
    Dim curDate As Date
    
    dtpManageTime(0).Enabled = cboManageTime.ListIndex = cboManageTime.ListCount - 1
    dtpManageTime(1).Enabled = cboManageTime.ListIndex = cboManageTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpManageTime(0).MaxDate = curDate + 1
    dtpManageTime(1).MaxDate = curDate + 1

    
    Select Case cboManageTime.ListIndex
    Case 0 '今日
        dtpManageTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '最近二天
        dtpManageTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '最近三天
        dtpManageTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '最近一周
        dtpManageTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '最近一月
        dtpManageTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 '指  定
        If Me.Visible Then
            dtpManageTime(0).SetFocus
        End If
    End Select
End Sub

Private Sub cboLogTime_Click()
    Dim curDate As Date
    
    dtpLogTime(0).Enabled = cboLogTime.ListIndex = cboLogTime.ListCount - 1
    dtpLogTime(1).Enabled = cboLogTime.ListIndex = cboLogTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpLogTime(0).MaxDate = curDate + 1
    dtpLogTime(1).MaxDate = curDate + 1

    
    Select Case cboLogTime.ListIndex
    Case 0 '今日
        dtpLogTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '最近二天
        dtpLogTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '最近三天
        dtpLogTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '最近一周
        dtpLogTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '最近一月
        dtpLogTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 '指  定
        If Me.Visible Then
            dtpLogTime(0).SetFocus
        End If
    End Select
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpTime(0).MaxDate = curDate + 1
    dtpTime(1).MaxDate = curDate + 1

    
    Select Case cboTime.ListIndex
    Case 0 '今日
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '最近二天
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '最近三天
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '最近一周
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '最近一月
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 '指  定
        If Me.Visible Then
            dtpTime(0).SetFocus
        End If
    End Select
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngApplyID As Long
    Select Case Control.ID
        Case conMenu_Edit_ApplyAdd
            If frmCISManageEdit.ShowEdit(Me, 0, lngApplyID) Then
                   Call LoadManage(lngApplyID)
            End If
        Case conMenu_Edit_ApplyEdit
            If Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) = 0 Then Exit Sub
            lngApplyID = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID))
            If frmCISManageEdit.ShowEdit(Me, 1, lngApplyID) Then
                Call LoadManage(lngApplyID)
            End If
        Case conMenu_Edit_Delete
            If tbcSub.Selected.Tag = "授权记录" Then
                If Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) = 0 Or vsManage.TextMatrix(vsManage.Row, COLM_作废时间) <> "" Then Exit Sub
                lngApplyID = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID))
                If ManageDelete(lngApplyID) Then
                    Call LoadManage(lngApplyID)
                End If
            ElseIf tbcSub.Selected.Tag = "审批记录" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_审批状态) <> "已审批" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
                If ApplyUpdate(lngApplyID, 2) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Manage_Complete
            If tbcSub.Selected.Tag = "审批记录" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_审批状态) <> "待审批" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
                If ApplyUpdate(lngApplyID, 1) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Manage_Undone
            If tbcSub.Selected.Tag = "审批记录" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_审批状态) <> "待审批" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
                If ApplyUpdate(lngApplyID, 3) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Edit_Untread
            If tbcSub.Selected.Tag = "审批记录" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_申请ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_审批状态) <> "已拒绝" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_申请ID))
                If ApplyUpdate(lngApplyID, 5) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_View_Refresh
            If tbcSub.Selected.Tag = "授权记录" Then
                Call LoadManage
            ElseIf tbcSub.Selected.Tag = "审批记录" Then
                Call LoadList
            End If
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Function ApplyUpdate(lngApplyID As Long, ByVal lngType As Long) As Boolean
    'lngType '1-审批，2-作废，3-拒绝'，5-取消拒绝'
    Dim strSQL As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    If MsgBox("确定要" & Decode(lngType, 1, "审批", 2, "作废", 3, "拒绝", 5, "对") & "选中的授权申请记录" & IIf(lngType = 5, "取消拒绝", "") & "吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zldatabase.Currentdate
    strSQL = "Zl_电子病历访问申请_审批状态(" & lngApplyID & "," & lngType & ",'" & UserInfo.姓名 & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ApplyUpdate = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ManageDelete(ByVal lng授权ID As Long) As Boolean
    '授权作废
    Dim strSQL As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("确定要作废选中的授权记录吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zldatabase.Currentdate
    strSQL = "Zl_电子病历访问授权_作废(" & lng授权ID & ",'" & UserInfo.姓名 & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ManageDelete = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function




Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '设置可见
    Select Case Control.ID
    Case conMenu_Edit_ApplyAdd
        Control.Visible = tbcSub.Selected.Tag = "授权记录"
    Case conMenu_Edit_ApplyEdit
        If tbcSub.Selected.Tag = "授权记录" Then
            Control.Visible = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) <> 0 And vsManage.TextMatrix(vsManage.Row, COLM_作废人) = ""
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_Complete
        If tbcSub.Selected.Tag = "审批记录" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_审批状态) = "待审批"
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_Undone
        If tbcSub.Selected.Tag = "审批记录" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_审批状态) = "待审批"
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Delete
        If tbcSub.Selected.Tag = "授权记录" Then
            Control.Visible = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) <> 0
        ElseIf tbcSub.Selected.Tag = "审批记录" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_审批状态) = "已审批"
        Else
            Control.Visible = False
        End If
    Case conMenu_File_Excel
        Control.Visible = tbcSub.Selected.Tag = "访问日志"
    Case conMenu_Edit_Untread
        If tbcSub.Selected.Tag = "审批记录" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_审批状态) = "已拒绝"
        Else
            Control.Visible = False
        End If
    End Select
End Sub


Private Sub chkFilter_Click(Index As Integer)
    Dim i As Long
    Dim blnCheck As Boolean
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then
        MsgBox "请至少选择一种分类用于过滤。", vbInformation, gstrSysName
        chkFilter(Index).Value = 1
        Exit Sub
    End If
End Sub


Private Sub cmdFind_Click()
    Call LoadList
End Sub


Private Sub LoadManage(Optional lng授权ID As Long)
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date

    On Error GoTo errH
    If cboManageTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "Select a.Id, a.授权类型, a.申请id, a.方案名, a.访问病人, a.访问开始时间, a.访问结束时间, a.内容时限, a.授权人, a.授权时间, a.作废人, a.作废时间,a.备注" & vbNewLine & _
                "From 电子病历访问授权 A Where A.授权类型 = 1  And a.授权时间 Between [1] And [2]" & IIf(chk已作废.Value = 0, " And A.作废时间 is null", "") & vbNewLine & _
                "Order by A.id"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpManageTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboManageTime.ListIndex <> 5, dtpManageTime(1).Value + 1, dtpManageTime(1).Value), "yyyy-MM-dd hh:mm")))
    With vsManage
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '隐藏列
                .TextMatrix(i, COLM_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COLM_内容时限) = Val(rsTmp!内容时限 & "")
                .TextMatrix(i, COLM_授权类型) = Val(rsTmp!授权类型 & "")
                .TextMatrix(i, COLM_访问病人) = Val(rsTmp!访问病人 & "")
                
                    
                 '显示列
                .TextMatrix(i, COLM_方案名) = rsTmp!方案名 & ""
                .TextMatrix(i, COLM_备注) = rsTmp!备注 & ""
                .TextMatrix(i, COLM_授权人) = rsTmp!授权人 & ""
                .TextMatrix(i, COLM_授权时间) = Format(rsTmp!授权时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_访问开始时间) = Format(rsTmp!访问开始时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_访问结束时间) = Format(rsTmp!访问结束时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_作废人) = rsTmp!作废人 & ""
                .TextMatrix(i, COLM_作废时间) = Format(rsTmp!作废时间 & "", "yyyy-mm-dd hh:mm")
                
                If rsTmp!作废时间 & "" <> "" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(4).Picture
                Else
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(1).Picture
                End If
                
                

                If Val(rsTmp!ID & "") = lng授权ID Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .ColHidden(COLM_作废人) = chk已作废.Value = 0
             .ColHidden(COLM_作废时间) = chk已作废.Value = 0
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "当前过滤查找到 " & rsTmp.RecordCount & " 份授权信息"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "当前过滤没有查找到授权信息"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1

        .WordWrap = True
        '自动调整行高
        .AutoSize COLM_备注, COLM_方案名
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Public Function GetRs病人姓名(rsTmp As ADODB.Recordset) As Boolean
    Dim str病人IDs As String
    Dim arrTmp As Variant
    Dim colPati As Collection
    Dim i As Long, j As Long
    Dim str姓名 As String, colValue As Collection
    
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.EOF Then Exit Function
    
    
    '加载病人信息
    str病人IDs = ""
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
             If rsTmp!病人ids & "" <> "" Then
                arrTmp = Split(rsTmp!病人ids & "", ",")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    If InStr("," & str病人IDs & ",", "," & Val(arrTmp(j)) & ",") = 0 Then
                       str病人IDs = str病人IDs & "," & Val(arrTmp(j))
                    End If
                Next
             End If
             rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    End If

    
    If str病人IDs <> "" Then
        str病人IDs = Mid(str病人IDs, 2)
        Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, "", "", "", "", str病人IDs)
        
        If Not colPati Is Nothing Then
            Set rsTmp = zldatabase.CopyNewRec(rsTmp)
            Do While Not rsTmp.EOF
               If rsTmp!病人ids & "" <> "" Then
                    arrTmp = Split(rsTmp!病人ids & "", ",")
                    str姓名 = ""
                    For j = LBound(arrTmp) To UBound(arrTmp)
                        If Val(arrTmp(j)) <> 0 Then
                            Set colValue = GetColObj(colPati, "_" & arrTmp(j))
                            If Not colValue Is Nothing Then
                                If GetColVal(colValue, "_pati_name") <> "" Then
                                    str姓名 = str姓名 & "," & GetColVal(colValue, "_pati_name")
                                End If
                            End If
                        End If
                    Next
                End If
                
                If str姓名 <> "" Then
                    str姓名 = Mid(str姓名, 2)
                    rsTmp!病人姓名 = str姓名
                End If
                
                rsTmp.MoveNext
            Loop
            rsTmp.MoveFirst
        End If
    End If
End Function


Private Sub LoadList(Optional lng申请id As Long)
    Dim strSQL As String
    Dim strFilter As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then strFilter = strFilter & "," & i
    Next
    strFilter = Mid(strFilter, 2)
    
    On Error GoTo errH
    If cboTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "Select a.Id, a.访问开始时间, a.访问结束时间, a.内容时限, a.申请原因, a.审批状态, a.申请人, a.申请时间,A.撤消时间,A.撤消人," & vbNewLine & _
                "       f_List2str(Cast(Collect(b.病人id || '') As t_Strlist)) As 病人ids,null as 病人姓名" & vbNewLine & _
                "From 电子病历访问申请 A, 电子病历申请访问病人 B" & vbNewLine & _
                "Where a.Id = b.申请id And a.申请时间 Between [1] And [2]" & vbNewLine & _
                " And Instr([3], a.审批状态) > 0 and A.撤消时间 is null" & vbNewLine & _
                "Group By a.Id, a.访问开始时间, a.访问结束时间, a.内容时限, a.申请原因, a.审批状态, a.申请人, a.申请时间,A.撤消时间,A.撤消人" & vbNewLine & _
                "Order by a.审批状态,A.id"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboTime.ListIndex <> 5, dtpTime(1).Value + 1, dtpTime(1).Value), "yyyy-MM-dd hh:mm")), strFilter)
    Call GetRs病人姓名(rsTmp)
    With vsList
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '隐藏列
                .TextMatrix(i, COL_申请ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_内容时限) = Val(rsTmp!内容时限 & "")
                .TextMatrix(i, COL_撤消时间) = Format(rsTmp!撤消时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_撤消人) = rsTmp!撤消人 & ""
                '显示列
                .TextMatrix(i, COL_申请人) = rsTmp!申请人 & ""
                .TextMatrix(i, COL_申请时间) = Format(rsTmp!申请时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_申请访问病人) = rsTmp!病人姓名 & ""
                .TextMatrix(i, COL_访问开始时间) = Format(rsTmp!访问开始时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_访问结束时间) = Format(rsTmp!访问结束时间 & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_申请原因) = rsTmp!申请原因 & ""
                
                .TextMatrix(i, COL_审批状态) = Decode(Val(rsTmp!审批状态 & ""), 0, "待审批", 1, "已审批", 2, "已作废", 3, "已拒绝")
                Set .Cell(flexcpPicture, i, 0) = imgFilter(Val(rsTmp!审批状态 & "")).Picture

                If Val(rsTmp!ID & "") = lng申请id Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "当前过滤查找到 " & rsTmp.RecordCount & " 份申请信息"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "当前过滤没有查找到申请信息"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1
        .WordWrap = True
        '自动调整行高
        .AutoSize COL_申请访问病人, COL_申请原因
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get访问者(lngRow As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
    
    If vsManage.TextMatrix(lngRow, COLM_访问者) <> "" Then Get访问者 = vsManage.TextMatrix(lngRow, COLM_访问者): Exit Function
    
    strSQL = "Select f_List2str(Cast(Collect(b.姓名 || '') As t_Strlist)) As 授权人员" & vbNewLine & _
                "From 电子病历授权访问人员 A, 人员表 B" & vbNewLine & _
                "Where a.人员id = b.Id And a.授权id =[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            Get访问者 = Replace(rsTmp!授权人员 & "", ",", "、")
            vsManage.TextMatrix(lngRow, COLM_访问者) = Replace(rsTmp!授权人员 & "", ",", "、")
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get访问范围(lngRow As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strOut As String
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
    If vsManage.TextMatrix(lngRow, COLM_病人范围详情) <> "" Then Get访问范围 = vsManage.TextMatrix(lngRow, COLM_病人范围详情): Exit Function
    
    Select Case Val(vsManage.TextMatrix(lngRow, COLM_访问病人))
        Case 0 '全院病人
            strOut = "操作人员拥有查看全院病人电子病历的权限"
        Case 1 '本科病人
            strOut = "操作人员拥有查看所在科室的病人电子病历权限"
        Case 2 '指定科室病人
            strSQL = "Select f_List2str(Cast(Collect(b.名称 || '') As t_Strlist)) As 访问科室" & vbNewLine & _
                        "From 电子病历授权访问病人 A, 部门表 B" & vbNewLine & _
                        "Where a.授权内容 = b.Id And a.授权id = [1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    strOut = Replace(rsTmp!访问科室 & "", ",", "、")
                End If
            End If
            strOut = "访问科室：" & strOut
        Case 3 '指定病人
            strSQL = "Select f_List2str(Cast(Collect(a.授权内容 || '') As t_Strlist)) As 病人ids,null as 病人姓名" & vbNewLine & _
                        "From 电子病历授权访问病人 A" & vbNewLine & _
                        "Where a.授权id = [1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
            Call GetRs病人姓名(rsTmp)
            
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    strOut = Replace(rsTmp!病人姓名 & "", ",", "、")
                End If
            End If
            strOut = "访问病人：" & strOut
    End Select
    
    vsManage.TextMatrix(lngRow, COLM_病人范围详情) = strOut
    Get访问范围 = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdLogFind_Click()
    Call LoadLog
End Sub

Private Sub cmdManageFind_Click()
    Call LoadManage
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    stbThis.Panels(2).Text = "查看" & tbcSub.Selected.Tag
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsList.Col >= vsList.FixedCols Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsInfo
        If Val(vsList.TextMatrix(NewRow, COL_申请ID)) <> 0 Then
            '访问病人
            .TextMatrix(Row_访问病人, 0) = vsList.TextMatrix(NewRow, COL_申请访问病人) & ""
            
            '内容时限
            .TextMatrix(Row_内容时限, 0) = "于 " & Format(vsList.TextMatrix(NewRow, COL_访问开始时间), "yyyy-mm-dd hh:mm") & vbCrLf & "至 " & _
                                        Format(vsList.TextMatrix(NewRow, COL_访问结束时间), "yyyy-mm-dd hh:mm") & "期间" & vbCrLf & "访问病人" & Decode(Val(vsList.TextMatrix(NewRow, COL_内容时限)), 0, "所有病历内容", 1, "未归档的病历", "已归档的病历")
                             
            '访问内容
            .TextMatrix(Row_访问内容, 0) = GetXmlInfo(1, NewRow)
        Else
            .TextMatrix(Row_访问病人, 0) = ""
            .TextMatrix(Row_内容时限, 0) = ""
            .TextMatrix(Row_访问内容, 0) = ""
        End If
        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With
End Sub

Private Function GetXmlInfo(ByVal intType As Integer, ByVal lngRow As Long) As String
    '获取申请内容的Xml并解析
    'lngType =0 授权详情 =1 审批详情
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String
    Dim strOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function

    If intType = 0 Then
        If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
        '读取缓存
        If vsManage.TextMatrix(lngRow, COLM_访问内容) <> "" Then GetXmlInfo = vsManage.TextMatrix(lngRow, COLM_访问内容): Exit Function

        strXML = Sys.ReadXML("电子病历访问授权", "访问内容", "ID=[1]", strErr, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
    Else
        If Val(vsList.TextMatrix(lngRow, COL_申请ID)) = 0 Then Exit Function
        '读取缓存
        If vsList.TextMatrix(lngRow, COL_访问内容) <> "" Then GetXmlInfo = vsList.TextMatrix(lngRow, COL_访问内容): Exit Function

        strXML = Sys.ReadXML("电子病历访问申请", "访问内容", "ID=[1]", strErr, Val(vsList.TextMatrix(lngRow, COL_申请ID)))
    End If

    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function

    '所有内容
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    If Val(strValue) = 1 Then
        strOut = "无限制访问所有内容"
    Else
        '病案首页、医嘱、临床路径
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then strOut = "病案首页、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "病人医嘱、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "临床路径、" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "体检报告、" & vbCrLf & vbCrLf
        
        '护理记录
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber)
            If Val(strValue) = 1 Then
                strOut = strOut & "护理记录(所有护理记录)" & vbCrLf & vbCrLf
            Else
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then strTmp = "体温单、"
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber)
                If Val(strValue) = 1 Then
                    Call GetXmlString(objXML, "nursing_info/file_name", strValue)
                    strValue = Replace(strValue, ",", "、")
                    strTmp = strTmp & strValue
                Else
                    strTmp = Replace(strTmp, "、", "")
                End If
                strOut = strOut & "护理记录" & vbCrLf & "(记录范围：" & strTmp & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '检查报告
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0所有检查报告 =1指定类型的检查报告
            If Val(strValue) = 0 Then
                strOut = strOut & "检查报告(所有检查报告)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "检查报告" & vbCrLf & "(类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '检验报告
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 所有检验报告 =1指定类型的检验报告
            If Val(strValue) = 0 Then
                strOut = strOut & "检验报告(所有检验报告)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "检验报告" & vbCrLf & "(类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '电子病历
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 所有电子病历  =1指定类型的电子病历  =1指定种类的电子病历
            If Val(strValue) = 0 Then
                strOut = strOut & "电子病历(所有电子病历)" & vbCrLf & vbCrLf
            ElseIf Val(strValue) = 1 Then
                Call GetXmlString(objXML, "emr_info/standard_class/class_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "电子病历" & vbCrLf & "(病历类型范围：" & strValue & ")" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue)
                strValue = Replace(strValue, ",", "、")
                strOut = strOut & "电子病历" & vbCrLf & "(病历范围：" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
    End If
    
    If Right(strOut, 5) = "、" & vbCrLf & vbCrLf Then strOut = Left(strOut, Len(strOut) - 5)
    '缓存内容数据
    If intType = 0 Then
        vsManage.TextMatrix(lngRow, COLM_访问内容) = strOut
    Else
        vsList.TextMatrix(lngRow, COL_访问内容) = strOut
    End If
    GetXmlInfo = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "授权记录", picManage.hwnd, 0).Tag = "授权记录"
        .InsertItem(1, "审批记录", picApply.hwnd, 0).Tag = "审批记录"
        .InsertItem(2, "访问日志", picLog.hwnd, 0).Tag = "访问日志"
        
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call InitListTable
    Call InitManageTable
    Call InitLogTable
    
    '初始化详情表格
    With vsInfo
        '访问病人
        .TextMatrix(Row_访问病人标题, 0) = "访问病人："
        .Cell(flexcpForeColor, Row_访问病人标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_访问病人标题, 0) = img16.ListImages("访问病人").Picture
        .Cell(flexcpFontBold, Row_访问病人标题, 0) = True
        
        '内容时限
        .TextMatrix(Row_内容时限标题, 0) = "内容时限："
        .Cell(flexcpForeColor, Row_内容时限标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_内容时限标题, 0) = img16.ListImages("访问时限").Picture
        .Cell(flexcpFontBold, Row_内容时限标题, 0) = True

        '访问内容
        .TextMatrix(Row_访问内容标题, 0) = "访问内容："
        .Cell(flexcpForeColor, Row_访问内容标题, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_访问内容标题, 0) = img16.ListImages("访问内容").Picture
        .Cell(flexcpFontBold, Row_访问内容标题, 0) = True

        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With

    '初始化详情表格
    With vsManageInfo
        '访问者
        .TextMatrix(RowM_访问者标题, 0) = "访问者："
        .Cell(flexcpForeColor, RowM_访问者标题, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_访问者标题, 0) = img16.ListImages("访问医生").Picture
        .Cell(flexcpFontBold, RowM_访问者标题, 0) = True
    
        '访问病人
        .TextMatrix(RowM_访问病人标题, 0) = "访问范围："
        .Cell(flexcpForeColor, RowM_访问病人标题, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_访问病人标题, 0) = img16.ListImages("访问病人").Picture
        .Cell(flexcpFontBold, RowM_访问病人标题, 0) = True
        
        '内容时限
        .TextMatrix(RowM_内容时限标题, 0) = "内容时限："
        .Cell(flexcpForeColor, RowM_内容时限标题, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_内容时限标题, 0) = img16.ListImages("访问时限").Picture
        .Cell(flexcpFontBold, RowM_内容时限标题, 0) = True

        '访问内容
        .TextMatrix(RowM_访问内容标题, 0) = "访问内容："
        .Cell(flexcpForeColor, RowM_访问内容标题, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_访问内容标题, 0) = img16.ListImages("访问内容").Picture
        .Cell(flexcpFontBold, RowM_访问内容标题, 0) = True

        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With
    
    '---cboTime
    cboTime.AddItem "今    日"
    cboTime.AddItem "最近二天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指  定]"
    cboTime.ListIndex = 3
    
    '---cboManageTime
    cboManageTime.AddItem "今    日"
    cboManageTime.AddItem "最近二天"
    cboManageTime.AddItem "最近三天"
    cboManageTime.AddItem "最近一周"
    cboManageTime.AddItem "最近一月"
    cboManageTime.AddItem "[指  定]"
    cboManageTime.ListIndex = 3
    
    '---cboLogTime
    cboLogTime.AddItem "今    日"
    cboLogTime.AddItem "最近二天"
    cboLogTime.AddItem "最近三天"
    cboLogTime.AddItem "最近一周"
    cboLogTime.AddItem "最近一月"
    cboLogTime.AddItem "[指  定]"
    cboLogTime.ListIndex = 3
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Call LoadList
    Call LoadManage
    Me.Tag = "1"
End Sub
'


Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "新增授权(&A)")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "调整授权(&E)")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "审批申请(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "拒绝申请(&N)")
            objControl.IconId = 4114
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消拒绝(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "作废授权(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "导出到&Excel")
            objControl.IconId = 30134
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "新增授权")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "调整授权")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "审批申请")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "拒绝申请")
            objControl.IconId = 4114
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消拒绝")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "作废授权")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "导出到&Excel")
            objControl.IconId = 3013
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call cbsMain_Resize
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = Me.Height - stbThis.Height - 1500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub picApply_Resize()
    On Error Resume Next
    '固定详细信息4000长度
    picInfo.Width = 5000

    fraFillter.Top = 100: fraFillter.Left = 30
    fraFillter.Width = picApply.Width - 60
    
    vsList.Top = fraFillter.Top + fraFillter.Height + 150: vsList.Height = picApply.Height - fraFillter.Height - 260

    
    vsList.Left = fraFillter.Left
    vsList.Width = fraFillter.Width - 5000 - 30
    
    picInfo.Top = vsList.Top - 70: picInfo.Left = vsList.Left + vsList.Width + 50
    picInfo.Height = vsList.Height + 70
    vsInfo.Height = picInfo.Height - 300
End Sub


Private Sub picManage_Resize()
    On Error Resume Next
    '固定详细信息4000长度
    fraManageInfo.Width = 5000

    fraManageFilter.Top = 100: fraManageFilter.Left = 30
    fraManageFilter.Width = picManage.Width - 60
    
    vsManage.Top = fraManageFilter.Top + fraManageFilter.Height + 150: vsManage.Height = picManage.Height - fraManageFilter.Height - 260

    
    vsManage.Left = fraManageFilter.Left
    vsManage.Width = fraManageFilter.Width - 5000 - 30
    
    fraManageInfo.Top = vsManage.Top - 70: fraManageInfo.Left = vsManage.Left + vsManage.Width + 50
    fraManageInfo.Height = vsManage.Height + 70
    vsManageInfo.Height = fraManageInfo.Height - 300
End Sub


Private Sub picLog_Resize()
    On Error Resume Next
    fraLog.Top = 100: fraLog.Left = 30
    fraLog.Width = picLog.Width - 60
    
    vsLog.Top = fraLog.Top + fraLog.Height + 150: vsLog.Height = picLog.Height - fraLog.Height - 260
    vsLog.Left = fraLog.Left
    vsLog.Width = fraLog.Width
End Sub


Private Sub InitListTable()
'功能：初始化列表清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "申请id;访问内容;内容时限;撤消时间;撤消人;" & _
                "申请时间,2000,1;申请人,800,4;申请访问病人,3200,1;访问开始时间,2000,1;访问结束时间,2000,1;申请原因,3800,1;审批状态,1050,4"
    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
    End With
End Sub


Private Sub InitManageTable()
'功能：初始化授权列表清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;访问内容;内容时限;授权类型;访问病人;病人范围详情;" & _
                "方案名,2500,1;备注,4000,1;访问开始时间,1700,1;访问结束时间,1700,1;授权人,1050,4;授权时间,1700,1;作废人,1050,4;作废时间,1700,1;访问者"
    arrHead = Split(strHead, ";")
    With vsManage
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsManage_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsManage.Col >= vsManage.FixedCols Then
        vsManage.ForeColorSel = vsManage.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsManageInfo
        If Val(vsManage.TextMatrix(NewRow, COLM_ID)) <> 0 Then
            .TextMatrix(RowM_访问者, 0) = Get访问者(NewRow)
        
            '访问病人
            .TextMatrix(RowM_访问病人, 0) = Get访问范围(NewRow)
            
            '内容时限
            .TextMatrix(RowM_内容时限, 0) = "于 " & Format(vsManage.TextMatrix(NewRow, COLM_访问开始时间), "yyyy-mm-dd hh:mm") & vbCrLf & "至 " & _
                                        Format(vsManage.TextMatrix(NewRow, COLM_访问结束时间), "yyyy-mm-dd hh:mm") & "期间" & vbCrLf & "访问病人" & Decode(Val(vsManage.TextMatrix(NewRow, COLM_内容时限)), 0, "所有病历内容", 1, "未归档的病历", "已归档的病历")
                                      
            '访问内容
            .TextMatrix(RowM_访问内容, 0) = GetXmlInfo(0, NewRow)
        Else
            .TextMatrix(RowM_访问者, 0) = ""
            .TextMatrix(RowM_访问病人, 0) = ""
            .TextMatrix(RowM_内容时限, 0) = ""
            .TextMatrix(RowM_访问内容, 0) = ""
        End If
        .WordWrap = True
        '自动调整行高
        .AutoSize 0
    End With
End Sub


Private Sub InitLogTable()
'功能：初始化列表清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;病人ID;就诊ID;病人来源;内容ID;" & _
                "访问时间,2000,1;访问者,1500,4;访问病人,1400,4;性别,700,4;年龄,700,4;标识号,950,4;科室,1700,4;病人类型,4000,1;访问内容,5000,1"
    arrHead = Split(strHead, ";")
    With vsLog
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
    End With
End Sub


Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Val(vsLog.TextMatrix(1, COLG_ID)) = 0 Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsLog
    
    objPrint.Title.Text = "电子病历访问记录"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub LoadLog()
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date

    On Error GoTo errH
    If cboLogTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If

    strSQL = "Select g.*, f.名称 As 科室名称" & vbNewLine & _
                "From (Select b.Id, b.病人id, b.就诊id, b.病人来源, b.内容id, b.访问时间, b.访问人, a.姓名, a.性别, a.年龄, a.门诊号 As 标识号, a.执行部门id As 科室," & vbNewLine & _
                "              a.发生时间 As 开始时间, Null As 结束时间, b.访问内容, -1 As 病人性质" & vbNewLine & _
                "       From 病人挂号记录 A, 电子病历访问日志 B" & vbNewLine & _
                "       Where a.病人id = b.病人id And a.Id = b.就诊id And b.病人来源 = 1 And b.访问时间 Between [1] And [2]" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select d.Id, d.病人id, d.就诊id, d.病人来源, d.内容id, d.访问时间, d.访问人, c.姓名, c.性别, c.年龄, c.住院号 As 标识号, c.出院科室id As 科室," & vbNewLine & _
                "              c.入院日期 As 开始时间, c.出院日期 As 结束时间, d.访问内容, Nvl(病人性质, 0) As 病人性质" & vbNewLine & _
                "       From 病案主页 C, 电子病历访问日志 D" & vbNewLine & _
                "       Where c.病人id = d.病人id And c.主页id = d.就诊id And d.病人来源 = 2 And d.访问时间 Between [1] And [2]) G, 部门表 F" & vbNewLine & _
                "Where g.科室 = f.Id" & vbNewLine & _
                "Order By 访问时间 Desc"

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpLogTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboLogTime.ListIndex <> 5, dtpLogTime(1).Value + 1, dtpLogTime(1).Value), "yyyy-MM-dd hh:mm")))
    With vsLog
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '隐藏列
                .TextMatrix(i, COLG_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COLG_病人ID) = Val(rsTmp!病人ID & "")
                .TextMatrix(i, COLG_就诊ID) = Val(rsTmp!就诊id & "")
                .TextMatrix(i, COLG_病人来源) = Val(rsTmp!病人来源 & "")
                .TextMatrix(i, COLG_内容ID) = rsTmp!内容ID & ""
                 '显示列

                .TextMatrix(i, COLG_访问时间) = Format(rsTmp!访问时间 & "", "yyyy-mm-dd hh:mm")
                Set .Cell(flexcpPicture, i, COLG_访问时间) = img16.ListImages("访问时限").Picture
                .TextMatrix(i, COLG_访问人) = rsTmp!访问人 & ""
                .TextMatrix(i, COLG_病人姓名) = rsTmp!姓名 & ""
                Set .Cell(flexcpPicture, i, COLG_病人姓名) = img16.ListImages(IIf(rsTmp!性别 & "" = "女", "girl", "boy")).Picture
                .TextMatrix(i, COLG_病人性别) = rsTmp!性别 & ""
                .TextMatrix(i, COLG_病人年龄) = rsTmp!年龄 & ""
                .TextMatrix(i, COLG_病人标识号) = rsTmp!标识号 & ""
                .TextMatrix(i, COLG_病人科室) = rsTmp!科室名称 & ""
                .TextMatrix(i, COLG_访问内容) = rsTmp!访问内容 & ""
                .TextMatrix(i, COLG_病人类型) = IIf(Val(rsTmp!病人来源) = 2, "第" & rsTmp!就诊id & "次" & IIf(rsTmp!病人性质 = 1, "门诊留观", IIf(rsTmp!病人性质 = 2, "住院留观", "住院")), "门诊就诊") & " " & Format(rsTmp!开始时间, "yyyy-MM-dd HH:mm") & _
                    IIf(Not IsNull(rsTmp!结束时间), "～" & Format(rsTmp!结束时间, "yyyy-MM-dd HH:mm"), "")
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "当前过滤查找到 " & rsTmp.RecordCount & " 条访问信息"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "当前过滤没有查找访问信息"
        End If

        If .Row <= 0 Then .Row = .Rows - 1

        .WordWrap = True
        '自动调整行高
        .AutoSize COLG_访问内容, COLG_病人类型
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


