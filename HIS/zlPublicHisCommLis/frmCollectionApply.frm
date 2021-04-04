VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCollectionApply 
   Caption         =   "检验手工申请单"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   Icon            =   "frmCollectionApply.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13080
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picLeftTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   60
      ScaleHeight     =   1335
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   390
      Width           =   10695
      Begin VB.TextBox txt年龄1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   5
         Top             =   112
         Width           =   555
      End
      Begin VB.ComboBox cbo医生 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   10
         Top             =   487
         Width           =   1605
      End
      Begin VB.ComboBox cbo开单科室 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08CA
         Left            =   3510
         List            =   "frmCollectionApply.frx":08CC
         TabIndex        =   9
         Top             =   487
         Width           =   1635
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8610
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   127
         Width           =   1635
      End
      Begin VB.TextBox txtPatientDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   8
         Top             =   495
         Width           =   1785
      End
      Begin VB.TextBox txtBed 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   6
         Top             =   127
         Width           =   1035
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08CE
         Left            =   4740
         List            =   "frmCollectionApply.frx":08DE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   112
         Width           =   750
      End
      Begin VB.TextBox txt年龄 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4290
         MaxLength       =   5
         TabIndex        =   3
         Top             =   112
         Width           =   435
      End
      Begin VB.ComboBox cbo性别 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08F4
         Left            =   3165
         List            =   "frmCollectionApply.frx":08F6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   112
         Width           =   675
      End
      Begin VB.TextBox txt姓名 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
         Top             =   112
         Width           =   1785
      End
      Begin VB.ComboBox cbo执行科室 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   11
         Text            =   "cbo执行科室"
         Top             =   487
         Width           =   1635
      End
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   12
         Top             =   900
         Width           =   9375
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2760
         TabIndex        =   25
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓       名"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标  识 号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   7890
         TabIndex        =   19
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   7845
         TabIndex        =   24
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单        位"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请医生"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   5370
         TabIndex        =   26
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6300
         TabIndex        =   23
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所在科室"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   20
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3885
         TabIndex        =   21
         Top             =   165
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   10725
      TabIndex        =   28
      Top             =   2100
      Width           =   10755
      Begin VB.Frame fraWE 
         BorderStyle     =   0  'None
         Height          =   4785
         Left            =   2430
         MousePointer    =   9  'Size W E
         TabIndex        =   35
         Top             =   780
         Width           =   60
      End
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   3300
         ScaleHeight     =   4125
         ScaleWidth      =   4215
         TabIndex        =   32
         Top             =   750
         Width           =   4245
         Begin VB.PictureBox picFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   30
            ScaleHeight     =   435
            ScaleWidth      =   4215
            TabIndex        =   33
            Top             =   0
            Width           =   4215
            Begin VB.TextBox txtFind 
               ForeColor       =   &H80000011&
               Height          =   315
               Left            =   480
               TabIndex        =   14
               Top             =   45
               Width           =   2355
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "查找"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   60
               TabIndex        =   34
               Top             =   60
               Width           =   360
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfItem 
            Height          =   3285
            Left            =   420
            TabIndex        =   17
            Top             =   1530
            Width           =   3225
            _cx             =   5689
            _cy             =   5794
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
      End
      Begin VB.PictureBox picGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4935
         Left            =   180
         ScaleHeight     =   4905
         ScaleWidth      =   2085
         TabIndex        =   31
         Top             =   720
         Width           =   2115
         Begin VSFlex8Ctl.VSFlexGrid vsfGroup 
            Height          =   3195
            Left            =   180
            TabIndex        =   16
            Top             =   420
            Width           =   1485
            _cx             =   2619
            _cy             =   5636
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
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
            SelectionMode   =   1
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
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         ScaleHeight     =   345
         ScaleWidth      =   10155
         TabIndex        =   29
         Top             =   120
         Width           =   10185
         Begin VB.ComboBox cboSampleType 
            Height          =   300
            Left            =   1140
            TabIndex        =   13
            Text            =   "cboSampleType"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CheckBox chkConcatenation 
            BackColor       =   &H80000005&
            Caption         =   "保存当前项目连续输入"
            Height          =   225
            Left            =   2940
            TabIndex        =   15
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标本类型"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   240
            TabIndex        =   30
            Top             =   60
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   8205
      Left            =   10800
      ScaleHeight     =   8145
      ScaleWidth      =   2145
      TabIndex        =   36
      Top             =   360
      Width           =   2205
      Begin VSFlex8Ctl.VSFlexGrid VSFSeled 
         Height          =   7725
         Left            =   30
         TabIndex        =   38
         Top             =   330
         Width           =   2085
         _cx             =   3678
         _cy             =   13626
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16706793
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "已选择(双击取消选择)"
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
         Left            =   30
         TabIndex        =   37
         Top             =   60
         Width           =   2100
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCollectionApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mrsRelativeAdvice As ADODB.Recordset                             '登记的相关医嘱
Private PatientType As Integer, mlng病人ID As Long, mstrNO As String    '门诊收费单据号
Private mlngCapID As Long                                               '采集项目ID
Private mlngReqDept As Long, mstrReqDoctor As String                    '默认的登记科室和医生
Private mblnSaveAdvice As Boolean                                       '是否需要保存医嘱，用于修改在院病人标本信息
Private mstrKeys As String                                              '当前核收的申请医嘱ID
Private mblnBarCode As Boolean                                          '条码
Private miInputType As Integer

Private mlngDeptID As Long                                              '科室ID
Private mrsItem As ADODB.Recordset              '组合项目
Private mstrItemSel As String                   '选择的组合项目
Private mblnFindEOF As Boolean                  '查找时，是否已经到达记录集末尾
Private mblnEdit As Boolean                     '是否编辑了数据
Private mstrSQLPro() As String                  '提交数据用的sql
Private mblnLoad As Boolean                     '是否首次加载

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const mConst_病人信息_列名 As String = "a.病人id,a.门诊号,a.住院号,a.就诊卡号,a.卡验证码,a.费别,a.医疗付款方式,a.姓名,a.性别,a.年龄,a.出生日期," & _
                                              "a.出生地点,a.身份证号,a.身份,a.职业,a.民族,a.国籍,a.区域,a.学历,a.婚姻状况,a.家庭地址,a.家庭电话,a.家庭地址邮编," & _
                                              "a.联系人关系,a.联系人地址,a.联系人电话,a.合同单位ID,a.工作单位,a.单位电话,a.单位邮编,a.单位开户行,a.单位帐号," & _
                                              "a.就诊时间,a.就诊状态,a.就诊诊室,a.住院次数,a.当前科室ID,a.当前病区ID,a.入院时间,a.出院时间," & _
                                              "a.IC卡号,a.健康号,a.险类,a.登记时间,a.停用时间,a.当前床号,a.医保号,a.查询密码,a.在院,a.其他证件,a.监护人,a.锁定,a.主页id"


Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cboSampleType_Click()
    Call vsfGroup_RowColChange
End Sub

Private Sub cboSampleType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Click()
    If cbo开单科室.ListIndex > -1 Then InitDoctors cbo开单科室.ItemData(cbo开单科室.ListIndex)
End Sub

Private Sub cbo开单科室_GotFocus()
    Call TxtSelAll(cbo开单科室)
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String, intIdx As Long
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean
              
1         On Error GoTo cbo开单科室_Validate_Error

2         If cbo开单科室.ListIndex <> -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex): Exit Sub '已选中
3         If cbo开单科室.Text = "" Then '无输入
4             Exit Sub
5         End If
          
6         strInput = UCase(NeedName(cbo开单科室.Text))
          '全院临床科室
7         strSQL = _
              " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
              " From 部门表 A,部门性质说明 B " & _
              " Where B.部门ID = A.ID " & _
              " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
              " And (B.工作性质 IN('临床','体检'))" & _
              " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
              " Order by A.编码"
          
8         vRect = GetControlRect(cbo开单科室.hWnd)
9         Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱科室", False, "", "", False, False, _
              True, vRect.Left, vRect.Top, cbo开单科室.Height, blnCancel, False, True, strInput & "%", strInput & "%")
10        If Not rsTmp Is Nothing Then
11            If Not CboLocate(cbo开单科室, rsTmp!名称) Then
12                cbo开单科室.Text = ""
13            End If
14        Else
15            If Not blnCancel Then
16                MsgBox "未找到对应的科室。", vbInformation, Me.Caption
17            End If
18            Cancel = True: Exit Sub
19        End If
20        If Me.cbo开单科室.ListIndex > -1 Then mlngReqDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)


21        Exit Sub
cbo开单科室_Validate_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(cbo开单科室_Validate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/27
'功    能:根据不同性别查询不同的标本类型（防止出现类似女性有精液标本这种笑话）
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub cbo性别_Click()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo cbo性别_Click_Error
              
2         strSQL = "select 名称 from 检验标本类型 where 适用性别=[1] or nvl(适用性别,0)=0"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验标本类型", IIf(Me.cbo性别.Text = "男", 1, 2))
4         With cboSampleType
5             .Clear
6             .AddItem "所有标本"
7             Do While Not rsTmp.EOF
8                 .AddItem rsTmp("名称") & ""
9                 rsTmp.MoveNext
10            Loop
11            If .ListCount > 0 Then .ListIndex = 0
12        End With
          

13        Exit Sub
cbo性别_Click_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(cbo性别_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear
End Sub


Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cbo医生_Click()
    Call TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo医生_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String, intIdx As Long
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean
              
1         On Error GoTo cbo医生_Validate_Error

2         If cbo医生.ListIndex <> -1 Then mstrReqDoctor = Me.cbo医生.Text: Exit Sub '已选中
3         If cbo医生.Text = "" Then '无输入
4             Exit Sub
5         End If
          
6         strInput = UCase(NeedName(cbo医生.Text))
          '全院医生
7         strSQL = "Select Distinct 部门ID From 部门性质说明 Where 服务对象 IN(1,2,3)"
8         strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
              " From 人员表 A,部门人员 B,人员性质说明 C" & _
              " Where A.ID=B.人员ID And A.ID=C.人员ID And C.人员性质='医生'" & _
              " And B.部门ID IN(" & strSQL & ")" & _
              " And (Upper(A.编号) Like [1] Or Upper(A.姓名) Like [2] Or Upper(A.简码) Like [2])" & _
              " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
              " Order by A.简码"
          
9         vRect = GetControlRect(cbo医生.hWnd)
10        Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱医生", False, "", "", False, False, _
              True, vRect.Left, vRect.Top, cbo医生.Height, blnCancel, False, True, strInput & "%", strInput & "%")
11        If Not rsTmp Is Nothing Then
12            cbo医生.Text = rsTmp!姓名
13        Else
14            If Not blnCancel Then
15                MsgBox "未找到对应的医生。", vbInformation, Me.Caption
16            End If
17            Cancel = True: Exit Sub
18        End If
19        If Len(Trim(Me.cbo医生.Text)) > 0 Then mstrReqDoctor = Me.cbo医生.Text


20        Exit Sub
cbo医生_Validate_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(cbo医生_Validate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
22        Err.Clear

End Sub

Private Sub cbo执行科室_Click()
    mlngDeptID = cbo执行科室.ItemData(cbo执行科室.ListIndex)
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo执行科室_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean

1         On Error GoTo cbo执行科室_Validate_Error

2         If cbo执行科室.ListIndex <> -1 Then mlngReqDept = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex): Exit Sub    '已选中
3         If cbo执行科室.Text = "" Then    '无输入
4             Exit Sub
5         End If

6         strInput = UCase(NeedName(cbo执行科室.Text))
          '全院临床科室
7         strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                 " From 部门表 A,部门性质说明 B " & _
                 " Where B.部门ID = A.ID " & _
                 " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
                 " And (B.工作性质 IN('检验'))" & _
                 " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
                 " Order by A.编码"


8         vRect = GetControlRect(cbo执行科室.hWnd)
9         Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "开嘱科室", False, "", "", False, False, _
                                               True, vRect.Left, vRect.Top, cbo执行科室.Height, blnCancel, False, True, strInput & "%", strInput & "%")
10        If Not rsTmp Is Nothing Then
11            If Not CboLocate(cbo执行科室, rsTmp!名称) Then
12                cbo执行科室.Text = ""
13            End If
14        Else
15            If Not blnCancel Then
16                MsgBox "未找到对应的科室。", vbInformation, Me.Caption
17            End If
18            Cancel = True: Exit Sub
19        End If
20        If Me.cbo执行科室.ListIndex > -1 Then mlngReqDept = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex)


21        Exit Sub
cbo执行科室_Validate_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(cbo执行科室_Validate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear

End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Save        '保存
            Call getSelItems
        Case ConMenu_Browse_Cancel      '取消
            Call cancelEdit
        Case ConMenu_Appfro_Exit        '退出
            Unload Me
    End Select
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picRight
        .Left = Right - .Width
        .Top = Top
        .Height = Bottom - Top
    End With
    
    With Me.picLeftTop
        .Left = Left
        .Top = Top
        .Width = Me.picRight.Left - Left
    End With
    
    With Me.picMain
        .Left = Left
        .Top = Me.picLeftTop.Top + Me.picLeftTop.Height + 50
        .Width = Me.picLeftTop.Width
        .Height = Bottom - .Top
    End With
End Sub


Private Function SaveData(ByRef strNewAdvice As String) As Boolean
    
    SaveData = SaveAdviceData(strNewAdvice)

End Function

Private Sub cancelEdit()
    Me.VSFSeled.Rows = 0
    Call CheckSelItem
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/3
'功    能:获取选择的项目
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub getSelItems()
          Dim rsTmp As New ADODB.Recordset
          Dim strItemSel As String
          Dim strOldNames As String
          Dim lngRow As Long
          Dim strSampleType As String
          Dim lngLoop As Long
          Dim blnTrs As Boolean
          Dim strCodeBefor As String  '试管编码
          Dim strCode As String   '试管编码
          Dim strNewAdvice As String
          Dim strErr As String

1         On Error GoTo getSelItems_Error

2         ReDim mstrSQLPro(0)

          '检查输入数据的合法性
3         If Not ValidAdvice Then Exit Sub

          '不同标本类型需要分批提交
4         With Me.VSFSeled
5             For lngRow = 0 To .Rows - 1
6                 strCodeBefor = GetSampleCode(Val(.TextMatrix(lngRow, .ColIndex("oldid"))))
7                 If (strSampleType = .TextMatrix(lngRow, .ColIndex("标本")) Or strSampleType = "") And (strCode = strCodeBefor Or strCode = "") Then
8                     strItemSel = strItemSel & "," & .TextMatrix(lngRow, .ColIndex("oldid"))
9                     strOldNames = strOldNames & "," & .TextMatrix(lngRow, .ColIndex("oldName"))
10                Else
11                    If strItemSel <> "" Then strItemSel = Mid(strItemSel, 2) & ";" & strSampleType
12                    If strOldNames <> "" Then strOldNames = Mid(strOldNames, 2) & ";" & strSampleType
13                    mstrItemSel = strOldNames
14                    If mstrItemSel <> "" Then
                          '获取采集方式
15                        Set rsTmp = SelectCap(Split(Split(strItemSel, ";")(0), ",")(0))
16                        If rsTmp Is Nothing Then
17                            MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, Me.Caption
18                            Exit Sub
19                        End If
20                        mlngCapID = rsTmp("ID")
21                        Call AdviceSet检查手术(3, strItemSel)
22                    End If
23                    If Not SaveData(strNewAdvice) Then Exit Sub

24                    strItemSel = ""
25                    strItemSel = strItemSel & "," & .TextMatrix(lngRow, .ColIndex("oldid"))
26                End If
27                strSampleType = .TextMatrix(lngRow, .ColIndex("标本"))
28                strCode = strCodeBefor
29            Next
30        End With

          '处理最后一批
31        If strItemSel <> "" Then
32            strItemSel = Mid(strItemSel, 2) & ";" & strSampleType
33            strOldNames = Mid(strOldNames, 2) & ";" & strSampleType
34            mstrItemSel = strOldNames
35            If mstrItemSel <> "" Then
                  '获取采集方式
36                Set rsTmp = SelectCap(Split(Split(strItemSel, ";")(0), ",")(0))
37                If rsTmp Is Nothing Then
38                    MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, Me.Caption
39                    Exit Sub
40                End If
41                mlngCapID = rsTmp("ID")
42                Call AdviceSet检查手术(3, strItemSel)
43            End If
44            If Not SaveData(strNewAdvice) Then Exit Sub
45        End If

          '提交数据
46        gcnHisOracle.BeginTrans
47        blnTrs = True
48        For lngLoop = 1 To UBound(mstrSQLPro)
49            Call ComExecuteProc(Sel_His_DB, mstrSQLPro(lngLoop), Me.Caption)
50        Next

          '生成新版医嘱
51        If SampleBarcodeUpdate(strNewAdvice, "", "", strErr, 0) = False Then
52            gcnHisOracle.RollbackTrans
53            blnTrs = False
54            If strErr <> "" Then
55                MsgBox strErr, vbInformation, gSysInfo.AppName
56            End If
57            Exit Sub
58        End If

59        gcnHisOracle.CommitTrans
60        blnTrs = False

61        If Me.chkConcatenation.value = 1 Then
62            Me.txt姓名.SetFocus
63        Else
64            Me.txt姓名 = ""
65            Me.txt年龄 = "": Me.txt年龄1 = "": Me.cboAge.ListIndex = 0
66            Me.txtBed = "": Me.txtID = ""
67            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
68            Me.cbo开单科室.ListIndex = -1
69            Me.cbo医生.ListIndex = -1
70            txtUnit.Text = ""

71            Me.VSFSeled.Rows = 0
72            Call CheckSelItem

73            Me.txt姓名.SetFocus
74        End If

75        MsgBox "登记成功，请刷新病人列表查看！", vbInformation, Me.Caption

76        Exit Sub
getSelItems_Error:
77        If blnTrs Then gcnHisOracle.RollbackTrans
78        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(getSelItems)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
79        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/9/20
'功    能:获取试管编码
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Private Function GetSampleCode(ByVal lngOldID As Long) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetSampleCode_Error

2         strSQL = "Select 试管编码 From 诊疗项目目录 Where ID = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", lngOldID)
4         If Not rsTmp.EOF Then
5             GetSampleCode = rsTmp("试管编码") & ""
6         End If


7         Exit Function
GetSampleCode_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(GetSampleCode)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
9         Err.Clear
End Function

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If VSFSeled.Rows > 0 And Trim(txt姓名.Text) <> "" Then
        mblnEdit = True
    Else
        mblnEdit = False
    End If
    Select Case Control.ID
        Case ConMenu_Browse_Save        '保存
            Control.Enabled = mblnEdit
        Case ConMenu_Browse_Cancel      '取消
            Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_Activate()
    If mblnLoad Then
        txt姓名.SetFocus
        mblnLoad = False
    End If
End Sub

Private Sub Form_Load()
    
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Save, "保存(Crl+S)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Cancel, "取消(Crl+U)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "退出(Crl+Q)")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, vbKeyS, ConMenu_Browse_Save
        .Add FCONTROL, vbKeyU, ConMenu_Browse_Cancel
        .Add FCONTROL, vbKeyQ, ConMenu_Appfro_Exit
    End With
         
    With VSFSeled
        .ExplorerBar = flexExSortShow
        .Rows = 0
        .Cols = 6
        .ColKey(0) = "ID": .ColHidden(0) = True
        .ColKey(1) = "名称"
        .ColKey(2) = "诊疗编码": .ColHidden(2) = True
        .ColKey(3) = "oldID": .ColHidden(3) = True
        .ColKey(4) = "标本": .ColHidden(4) = True
        .ColKey(5) = "oldName": .ColHidden(5) = True
    End With
    
    Call InitDepts                      '取得科室和性别
    Call intData                        '加载项目
    
    '设置文本框提示字
    Call setTxtTip(txtFind, "输入编码、简码或名称敲击回车查找")
    
    mblnLoad = True
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/26
'功    能:加载数据
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub intData()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsGroup As ADODB.Recordset
          Dim strFenLei As String
          Dim blnHasGroup As Boolean  '是否有执行小组


1         On Error GoTo intData_Error


          '查询组合项目
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
3         strSQL = "Select 0 选择, a.Id, a.编码, a.名称, a.简码, b.申请单id, Nvl(b.分组id, 0) 分组id, a.诊疗编码, a.检验标本 标本" & vbCrLf & _
                 " From 检验组合项目 A, 检验申请单明细 B" & vbCrLf & _
                 " Where a.Id = b.组合id And a.停用日期 Is Null And a.诊疗编码 Is Not Null And nvl(a.是否耐受项目, 0) = 0"
4         Else
5             strSQL = "Select 0 选择, a.Id, a.编码, a.名称, a.简码, b.申请单id, Nvl(b.分组id, 0) 分组id, a.诊疗编码, a.检验标本 标本" & vbCrLf & _
                 " From 检验组合项目 A, 检验申请单明细 B" & vbCrLf & _
                 " Where a.Id = b.组合id And a.停用日期 Is Null And a.诊疗编码 Is Not Null"
6         End If
7         If gUserInfo.NodeNo <> "-" Then
8             strSQL = strSQL & " And (a.站点 = [1] Or Nvl(a.站点, 0) = 0)"
9         End If
10        strSQL = strSQL & " Order By b.申请单id,b.分组id,b.排列顺序"
11        Set mrsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "检验组合项目", gUserInfo.NodeNo)

          '查询申请单分类
12        If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
13            strSQL = "Select ID,申请单ID, 分类, 名称 分组,执行小组" & vbCrLf & _
                     " From (Select Distinct a.名称 分类, b.Id,a.ID 申请单ID, b.名称,a.执行小组" & vbCrLf & _
                     "        From 检验申请单 A, 检验申请单分组 B, 检验申请单明细 C" & vbCrLf & _
                     "        Where a.Id = b.申请单id(+) And a.Id = c.申请单id And c.分组id Is Not Null And" & vbCrLf & _
                     "        nvl(a.是否耐受申请单, 0) = 0 And (a.科室id = [1] Or Nvl(a.科室id, 0) = 0)" & vbCrLf & _
                     "        Union all" & vbCrLf & _
                     "        Select Distinct a.名称, 0 id,a.ID 申请单ID, '未分组' 名称,a.执行小组" & vbCrLf & _
                     "        From 检验申请单 A, 检验申请单分组 B, 检验申请单明细 C" & vbCrLf & _
                     "        Where a.Id = b.申请单id(+) And a.Id = c.申请单id And c.分组id Is Null And" & vbCrLf & _
                     "        nvl(a.是否耐受申请单, 0) = 0 And (a.科室id = [1] Or Nvl(a.科室id, 0) = 0))" & vbCrLf & _
                     " Order By 分类,申请单ID,ID"
14        Else
15            strSQL = "Select ID,申请单ID, 分类, 名称 分组,执行小组" & vbCrLf & _
                     " From (Select Distinct a.名称 分类, b.Id,a.ID 申请单ID, b.名称,a.执行小组" & vbCrLf & _
                     "        From 检验申请单 A, 检验申请单分组 B, 检验申请单明细 C" & vbCrLf & _
                     "        Where a.Id = b.申请单id(+) And a.Id = c.申请单id And c.分组id Is Not Null And" & vbCrLf & _
                     "              (a.科室id = [1] Or Nvl(a.科室id, 0) = 0)" & vbCrLf & _
                     "        Union all" & vbCrLf & _
                     "        Select Distinct a.名称, 0 id,a.ID 申请单ID, '未分组' 名称,a.执行小组" & vbCrLf & _
                     "        From 检验申请单 A, 检验申请单分组 B, 检验申请单明细 C" & vbCrLf & _
                     "        Where a.Id = b.申请单id(+) And a.Id = c.申请单id And c.分组id Is Null And" & vbCrLf & _
                     "              (a.科室id = [1] Or Nvl(a.科室id, 0) = 0))" & vbCrLf & _
                     " Order By 分类,申请单ID,ID"
16        End If
17        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请单", mlngDeptID)

          '查询站点对应的执行小组
18        If gUserInfo.NodeNo <> "-" Then
19            strSQL = "Select Distinct 编码 From 检验小组记录 Where 站点 = [1] or 站点 is null"
20        Else
21            strSQL = "Select Distinct 编码 From 检验小组记录"
22        End If
23        Set rsGroup = ComOpenSQL(Sel_Lis_DB, strSQL, "检验小组记录", gUserInfo.NodeNo)


24        With Me.vsfGroup
25            .ExplorerBar = flexExSort
26            .Rows = 1
27            .Cols = 4
28            .FixedRows = 1
29            .OutlineBar = flexOutlineBarComplete
30            .OutlineCol = 1
              '        .SubtotalPosition = flexSTAbove
31            .ExtendLastCol = True

              '1.边框
32            .Appearance = flex3DLight
33            .BorderStyle = flexBorderFlat
34            .GridLines = flexGridNone
35            .GridColorFixed = flexGridNone

              '2.颜色
36            .BackColor = vbWindowBackground    '窗口背景
37            .BackColorAlternate = vbWindowBackground
38            .BackColorBkg = vbWindowBackground
39            .BackColorFixed = vbButtonFace    '按钮表面
40            .BackColorFrozen = &H0&         '黑
41            .FloodColor = &HC0&             '红
42            .BackColorSel = &HFFEBD7        '浅绿
43            .ForeColor = vbWindowText       '窗口文本
44            .ForeColorFixed = vbButtonText  '按钮文本
45            .ForeColorFrozen = &H0&         '黑
46            .ForeColorSel = vbWindowText

47            .GridColor = vbApplicationWorkspace    '应用程序工作区
48            .GridColorFixed = vbApplicationWorkspace
49            .SheetBorder = vbWindowBackground
50            .TreeColor = vbButtonShadow         '按钮阴影


51            .ColKey(0) = "id": .ColWidth(.ColIndex("id")) = 0: .ColHidden(.ColIndex("id")) = True
52            .ColKey(1) = "申请单ID": .ColWidth(.ColIndex("申请单ID")) = 250
53            .ColKey(2) = "分类": .ColAlignment(.ColIndex("分类")) = flexAlignCenterCenter: .ColHidden(.ColIndex("分类")) = True
54            .ColKey(3) = "分组": .ColWidth(.ColIndex("分组")) = 250: .ColAlignment(.ColIndex("分组")) = flexAlignCenterCenter: .TextMatrix(0, .ColIndex("分组")) = "分组"

55            Do While Not rsTmp.EOF
56                blnHasGroup = False
                  '判断申请单是否有执行小组
57                If Not rsGroup Is Nothing Then
58                    If rsGroup.RecordCount > 0 Then
59                        rsGroup.MoveFirst
60                        Do While Not rsGroup.EOF
61                            If InStr("," & rsTmp("执行小组") & ",", "," & rsGroup("编码") & ",") > 0 Then
62                                blnHasGroup = True
63                            End If
64                            rsGroup.MoveNext
65                        Loop
66                    End If
67                End If

68                If blnHasGroup Then
69                    If InStr(";" & strFenLei & ";", ";" & rsTmp("分类") & ";") <= 0 Then
70                        .Rows = .Rows + 2

71                        .TextMatrix(.Rows - 2, .ColIndex("ID")) = rsTmp("分类") & ""
72                        .TextMatrix(.Rows - 2, .ColIndex("申请单ID")) = rsTmp("分类") & ""
73                        .TextMatrix(.Rows - 2, .ColIndex("分类")) = rsTmp("分类") & ""
74                        .TextMatrix(.Rows - 2, .ColIndex("分组")) = rsTmp("分类") & "": .Cell(flexcpAlignment, .Rows - 2, .ColIndex("分类")) = flexAlignLeftCenter

                          '加粗
75                        .Cell(flexcpFontBold, .Rows - 2, 0, .Rows - 2, .Cols - 1) = True

                          '合并
76                        .MergeRow(.Rows - 2) = True
77                        .MergeCellsFixed = flexMergeRestrictRows

                          '缩进
78                        .IsSubtotal(.Rows - 2) = True
79                        .RowOutlineLevel(.Rows - 2) = 1

80                        .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
81                        .TextMatrix(.Rows - 1, .ColIndex("申请单ID")) = rsTmp("申请单ID") & ""
82                        .TextMatrix(.Rows - 1, .ColIndex("分组")) = rsTmp("分组") & ""
83                        strFenLei = strFenLei & ";" & rsTmp("分类") & ""
84                    Else
85                        .Rows = .Rows + 1
86                        .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
87                        .TextMatrix(.Rows - 1, .ColIndex("申请单ID")) = rsTmp("申请单ID") & ""
88                        .TextMatrix(.Rows - 1, .ColIndex("分组")) = rsTmp("分组") & ""
89                    End If
90                End If

91                rsTmp.MoveNext
92            Loop

              '默认选中第一个分组
93            If .Rows > 2 Then .Row = 2
94        End With



95        Exit Sub
intData_Error:
96        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(intData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
97        Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetPara Sel_His_DB, "采集工作站登记", chkConcatenation.value, 100, 1211
    Set mrsItem = Nothing
    Set mrsRelativeAdvice = Nothing
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button <> 1 Then Exit Sub
    With Me.fraWE
        If .Left + X < 2000 Or picMain.Width - (.Left + X) < 2000 Then Exit Sub
        .Left = .Left + X
        .Tag = .Left
    End With
    With Me.picGroup
        .Width = Me.fraWE.Left
    End With
    
    With Me.picItem
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Width = Me.picMain.Width - .Left
    End With
    
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    With Me.txtFind
        .Width = Me.picFind.Width - .Left - 200
    End With
End Sub

Private Sub picGroup_Resize()
    On Error Resume Next
    With Me.vsfGroup
        .Left = 0
        .Top = 0
        .Width = Me.picGroup.Width
        .Height = Me.picGroup.Height
    End With
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    With Me.picFind
        .Left = 0
        .Top = 0
        .Width = Me.picItem.Width
    End With
    With Me.vsfItem
        .Left = 0
        .Top = picFind.Height
        .Width = Me.picItem.Width
        .Height = Me.picItem.Height - .Top
    End With
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.picFilter
        .Left = 0
        .Top = 0
        .Width = Me.picMain.Width
        .BorderStyle = 0
    End With
    With Me.fraWE
        .Top = picFilter.Height
        .Height = picMain.Height - .Top
    End With
    With Me.picGroup
        .Left = 0
        .Top = picFilter.Height
        .Width = Me.fraWE.Left
        .Height = Me.picMain.Height - .Top
    End With
    With Me.picItem
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Top = picGroup.Top
        .Width = Me.picMain.Width - .Left
        .Height = picGroup.Height
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With VSFSeled
        .Height = Me.picRight.Height - .Top
        .Width = Me.picRight.Width - 100
    End With
End Sub

Private Sub txtBed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtFind_GotFocus()
    Call selAllText(txtFind)
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/2
'功    能:查找项目
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub txtFind_KeyPress(KeyAscii As Integer)
          Dim strFind As String
          Dim strFilter As String
          Dim strTmp() As String
          Dim strSub() As String
          Dim lngRow As Long
          Dim i As Integer
          
          
1         On Error GoTo txtFind_KeyPress_Error

2         If KeyAscii <> 13 Then
3             vsfItem.Tag = ""
4             mblnFindEOF = False
5             Exit Sub
6         End If
7         If mrsItem Is Nothing Then Exit Sub
          
8         strFind = UCase(Trim(Me.txtFind.Text))
9         If strFind = "" Then Exit Sub
          '通过数据的内容去过滤记录集
10        mrsItem.Filter = "编码 like '%" & strFind & "%'"
11        If mrsItem.RecordCount <= 0 Then
12            mrsItem.Filter = "名称 like '%" & strFind & "%'"
13            If mrsItem.RecordCount <= 0 Then
14                mrsItem.Filter = "简码 like '%" & strFind & "%'"
15            End If
16        End If

          
          '由于不同分组下可能存在相同的项目，所以查找时可能会存在多行记录，需要查找每一行记录
17        If mblnFindEOF = False Then  '只有当上次过滤的记录集中的内容都已经查找了一遍之后才进行新的记录
18            strFilter = ""
19            Do While Not mrsItem.EOF
20                strFilter = strFilter & ";" & mrsItem("申请单ID") & "," & mrsItem("分组ID") & "," & mrsItem("ID")
21                mrsItem.MoveNext
22            Loop
23            If strFilter <> "" Then Me.txtFind.Tag = Mid(strFilter, 2)
24        End If
              
25        If Trim(Me.txtFind.Tag) = "" Then Exit Sub
26        strTmp = Split(Trim(Me.txtFind.Tag), ";")
          
          '开始遍历查找
27        For i = 0 To UBound(strTmp)
28            strSub = Split(strTmp(i), ",")
29            mblnFindEOF = True
              '先选中项目对应的分组
30            With vsfGroup
31                For lngRow = 1 To .Rows - 1
32                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = Val(strSub(1)) And Val(.TextMatrix(lngRow, .ColIndex("申请单ID"))) = Val(strSub(0)) Then
33                        .Row = lngRow
34                        .ShowCell .Row, 0
35                        vsfItem.Tag = ""
36                        Exit For
                      
37                    End If
38                Next
39            End With
              
              '再选中分组下面的项目
40            With Me.vsfItem
41                For lngRow = Val(.Tag) + 1 To .Rows - 1
42                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = Val(strSub(2)) Then
43                        .Row = lngRow
44                        .ShowCell .Row, 0
45                        If lngRow >= .Rows - 1 Then
46                            .Tag = 0
47                        Else
48                            .Tag = lngRow
49                        End If
50                        Exit For
51                    End If
52                Next
53            End With
              
              '清除已经查找过的内容
54            Me.txtFind.Tag = Replace(Me.txtFind.Tag, strSub(0) & "," & strSub(1) & "," & strSub(2), "")
              '清楚两端的分号
55            If Mid(Me.txtFind.Tag, 1, 1) = ";" Then
56                Me.txtFind.Tag = Mid(Me.txtFind.Tag, 2)
57            End If
58            If Me.txtFind.Tag <> "" Then
59                If Mid(Me.txtFind.Tag, Len(Me.txtFind.Tag) - 1, 1) = ";" Then
60                    Me.txtFind.Tag = Mid(Me.txtFind.Tag, 1, Len(Me.txtFind.Tag) - 1)
61                End If
62            End If
              
63            Exit For
64        Next
          

65        mrsItem.Filter = ""
66        If Trim(Me.txtFind.Tag) = "" Then
67            vsfItem.Tag = ""
68            mblnFindEOF = False
69        End If


70        Exit Sub
txtFind_KeyPress_Error:
71        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(txtFind_KeyPress)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
72        Err.Clear
End Sub

Private Sub txtFind_LostFocus()
    Call setTxtTip(txtFind, "输入编码、简码或名称敲击回车查找")
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtPatientDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txt年龄_GotFocus()
    TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt年龄1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txt姓名_GotFocus()
    TxtSelAll txt姓名
End Sub

Private Sub txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    End If
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txt姓名.Text)) = 0 Then Exit Sub
        Call txt姓名_Validate(False)
        cbo性别.SetFocus
    End If
End Sub

Private Sub txt姓名_Validate(Cancel As Boolean)
          Dim rsTmp As New ADODB.Recordset, i As Integer
          Dim strField As String
          Dim strBarCode As String
          Dim rsDept As ADODB.Recordset, strSQL As String
          Dim strAge As String
          Dim aAge() As String

1         On Error GoTo txt姓名_Validate_Error

2         If Len(Trim(txt姓名)) = 0 Then Exit Sub
3         If txt姓名 = txt姓名.Tag Then Exit Sub

4         Call AdjustEditState(True)

5         mblnSaveAdvice = True
6         Cancel = Not StrIsValid(txt姓名.Text, txt姓名.MaxLength)

          '初始病人信息
7         Set rsTmp = GetPatient(txt姓名)
8         strBarCode = txt姓名
9         If rsTmp.EOF Then
              '登记新病人
10            mlng病人ID = 0
11            mstrKeys = ""
12            Me.txtBed = "": Me.txtID = ""
13            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0

              '如果想输入院内病人，则不允许继续
14            If InStr("+-*./", Left(Me.txt姓名.Text, 1)) > 0 Or mblnBarCode Then
15                Me.txt姓名.Text = "": Cancel = True
16                Exit Sub
17            End If
18            PatientType = 1
19        Else
20            Me.txt姓名.Text = NVL(rsTmp("姓名"))
21            Me.txt年龄.Text = "": Me.txt年龄1.Text = ""
22            strAge = IIf(IsNull(rsTmp("年龄")), "", rsTmp("年龄")): If Me.txt年龄 = "0" Then Me.txt年龄 = ""

23            strAge = Replace(strAge, "小时", "时")
24            strAge = Replace(strAge, "分钟", "分")

25            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "岁", ""), "月", ""), "天", ""), "时", ""), "分", "")) <> "" Then
26                If InStr(strAge, "成人") > 0 Or InStr(strAge, "婴儿") > 0 Then
27                    Me.txt年龄.Text = ""
28                    Me.cboAge.Text = Trim(strAge)
29                Else
30                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "岁", "岁;"), "月", "月;"), "天", "天;"), "时", "时;"), "分", "分;")
31                    aAge = Split(strAge, ";")
32                    If UBound(aAge) = 1 Then
33                        Me.txt年龄.Text = Val(aAge(0))
34                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
35                    Else
36                        Me.txt年龄.Text = Val(aAge(0))
37                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "分", "分钟"), "时", "小时")
38                        Me.txt年龄1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "分", "分钟"), "时", "小时")
39                    End If
40                End If
41            Else
42                Me.txt年龄.Text = ""
43                Me.cboAge.ListIndex = 0
44            End If

45            If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
46            Me.cbo性别 = NVL(rsTmp("性别"))    ' CombIndex(cbo性别, Nvl(rsTmp("性别")))

47            mlng病人ID = NVL(rsTmp("病人ID"), 0): PatientType = NVL(rsTmp("PatientType"), 1)

              '设置默认开单科室、医生
48            cbo开单科室.ListIndex = FindComboItem(cbo开单科室, NVL(rsTmp("病人科室"), 0))

              '病人单位
49            txtUnit.Text = NVL(rsTmp("工作单位"))

50            strField = ""
51            strField = rsTmp.Fields("医生").Name
52            If strField = "医生" Then
53                Me.cbo医生.Text = NVL(rsTmp("医生"))
54                For i = 0 To Me.cbo医生.ListCount - 1
55                    If Me.cbo医生.List(i) Like NVL(rsTmp("医生")) Then
56                        Me.cbo医生.ListIndex = i
57                        Exit For
58                    End If
59                Next
60            End If

              '显示病人科室
61            strSQL = "Select 名称 From 部门表 Where ID=[1]"
62            Set rsDept = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, CLng(NVL(rsTmp("病人科室"), 0)))
63            If rsDept.EOF Then
64                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
65            Else
66                Me.txtPatientDept.Text = rsDept("名称") & "": Me.txtPatientDept.Tag = NVL(rsTmp("病人科室"), 0)
67            End If
68            Me.txtID = rsTmp("住院号") & "": If Len(Me.txtID) = 0 Then Me.txtID = rsTmp("门诊号") & ""
69            Me.txtBed = NVL(rsTmp("当前床号"))

              '处理登记的默认科室、医生
70            If Me.cbo开单科室.ListIndex = -1 And mlngReqDept > 0 Then
71                cbo开单科室.ListIndex = FindComboItem(cbo开单科室, mlngReqDept)
72            End If
73        End If

74        txt姓名.Tag = txt姓名.Text


75        Exit Sub
txt姓名_Validate_Error:
76        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(txt姓名_Validate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
77        Err.Clear
End Sub

Private Sub InitDoctors(ByVal lng科室ID As Long)
      '功能：读取当前开单科室中包含的所有人员
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          
1         On Error GoTo InitDoctors_Error

2         Me.cbo医生.Clear
          
          '科室医生或护士
3         strSQL = _
              "Select Distinct A.ID,B.部门ID,A.编号,A.姓名,Upper(A.简码) as 简码," & _
              " C.人员性质,Nvl(A.聘任技术职务,0) as 职务" & _
              " From 人员表 A,部门人员 B,人员性质说明 C" & _
              " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
              " And C.人员性质 IN('医生') And B.部门ID=[1] " & _
              " And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
              
4         strSQL = strSQL & " Order by 简码,人员性质 Desc"
          
5         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng科室ID)
          
6         If Not rsTmp.EOF Then
7             For i = 1 To rsTmp.RecordCount
8                 cbo医生.AddItem rsTmp!姓名
9                 cbo医生.ItemData(cbo医生.ListCount - 1) = rsTmp!部门ID
                  
10                If rsTmp!ID = gUserInfo.ID And cbo医生.ListIndex = -1 Then
11                    cbo医生.ListIndex = cbo医生.NewIndex
12                ElseIf cbo医生.ListCount > 0 Then
13                    cbo医生.ListIndex = 0
14                End If
15                rsTmp.MoveNext
16            Next
              
17            If cbo医生.ListCount = 1 And cbo医生.ListIndex = -1 Then cbo医生.ListIndex = 0
18        End If


19        Exit Sub
InitDoctors_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(InitDoctors)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear
End Sub
Public Sub ShowMe(objFrm As Object)
    Me.Show vbModal, objFrm
End Sub

Private Function InitDepts() As Boolean
      '功能：初始化住院临床科室
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strOldText As String
          Dim intloop As Integer
              
1         On Error GoTo InitDepts_Error

2         strSQL = _
              " Select Distinct A.ID,A.编码,A.名称" & _
              " From 部门表 A,部门性质说明 B " & _
              " Where B.部门ID = A.ID " & _
              " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
              " And (B.工作性质 IN('检验'))" & _
              " Order by A.编码"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
          
4         With Me.cbo执行科室
5             Do While Not rsTmp.EOF
6                 .AddItem NVL(rsTmp("名称"))
7                 .ItemData(.NewIndex) = rsTmp("ID")
8                 rsTmp.MoveNext
9             Loop
10            If .ListCount > 0 Then
11                .ListIndex = 0
12            End If
13        End With
          
          
14        strOldText = Me.cbo开单科室.Text
15        Me.cbo开单科室.Clear
          
16        strSQL = _
              " Select Distinct A.ID,A.编码,A.名称" & _
              " From 部门表 A,部门性质说明 B " & _
              " Where B.部门ID = A.ID " & _
              " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
              " And (B.工作性质 IN('临床','体检'))" & _
              " Order by A.编码"
17        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
          
18        For i = 1 To rsTmp.RecordCount
19            cbo开单科室.AddItem rsTmp!名称
20            cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!ID
              
21            rsTmp.MoveNext
22        Next
          
23        On Error Resume Next
24        Me.cbo开单科室.Text = strOldText
          
           '性别
26        Set rsTmp = Nothing
27        Set rsTmp = GetDictData("性别")
28        cbo性别.Clear
29        If Not rsTmp Is Nothing Then
30            For intloop = 1 To rsTmp.RecordCount
31                cbo性别.AddItem rsTmp!名称
32                If rsTmp!缺省 = 1 Then
33                    cbo性别.ItemData(cbo性别.NewIndex) = 1
34                    cbo性别.ListIndex = cbo性别.NewIndex
35                End If
36                rsTmp.MoveNext
37            Next
38        End If
          
39        chkConcatenation.value = GetPara(Sel_His_DB, "采集工作站登记", 100, 1211, 0)

40        InitDepts = True


41        Exit Function
InitDepts_Error:
42        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(InitDepts)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
43        Err.Clear

End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '功能:              调整编辑状态
    'Me.txt姓名.Enabled = blEnable
    cbo性别.Enabled = blEnable
    txt年龄.Enabled = blEnable
    txt年龄1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo开单科室.Enabled = blEnable
    cbo医生.Enabled = blEnable
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
      '功能：读取病人信息，并显示该病人存在的医嘱时间
          Dim strSQL As String
          Dim strNO As String, str姓名 As String
          Dim strSeek As String
          
          
1         On Error GoTo GetPatient_Error

2         If BlnIsNumber(strCode) Then
          '预置条码单独处理
3             mblnBarCode = True
4             strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,B.主页ID,B.病人科室id As 病人科室,B.开嘱医生 As 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A,病人医嘱记录 B,病人医嘱发送 C Where A.病人ID=B.病人ID+0 And B.ID=C.医嘱ID+0" & _
                  " And C.样本条码=[1]"
5             Set GetPatient = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, strCode)
6             Exit Function
7         End If
8         mblnBarCode = False
          
9         strSeek = strCode
          '判断当前输入模式
10        If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then  '刷卡
11            miInputType = 0
12        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
13            miInputType = 1
14            strSeek = Mid(strCode, 2)
15        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号
16            miInputType = 2
17            strSeek = Mid(strCode, 2)
18        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号
19            miInputType = 3
20            strSeek = Mid(strCode, 2)
21        ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '挂号单
22            miInputType = 4
23            strSeek = Mid(strCode, 2)
24        ElseIf Left(strCode, 1) = "/" Then '收费单据号
25            miInputType = 5
26            strSeek = Mid(strCode, 2)
27        ElseIf Not IsNumeric(Mid(strCode, 2)) Then '当作姓名
28            miInputType = 6
29            strSeek = Replace(strCode, "(婴儿)", "")
30        End If
          
31        If miInputType = 0 Then '刷卡
32            strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A,病人挂号记录 B Where A.就诊卡号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and (b.病人ID is null or (b.记录性质 =1 and b.记录状态 =1)) "
      '            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
33        ElseIf miInputType = 1 Then '病人ID
34            strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,0) As 病人科室,'' 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A Where A.病人ID=[2]"
35        ElseIf miInputType = 2 Then '住院号
36            strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.入院科室ID,0),A.当前科室id) As 病人科室,B.住院医师 As 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A,病案主页 B Where A.住院号=[2] And A.病人ID=B.病人ID" ' And A.当前科室id IS NOT NULL And B.出院日期 Is NULL"
37        ElseIf miInputType = 3 Then '门诊号
38            strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Decode(A.当前科室id,Null,Nvl(B.执行部门ID,0),A.当前科室id) As 病人科室,B.执行人 As 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A,病人挂号记录 B Where A.门诊号=[1] And A.病人ID=B.病人ID(+) And A.门诊号=B.门诊号(+) and (b.病人ID is null or (b.记录性质 =1 and b.记录状态 =1)) "
      '            " And (A.当前科室id IS NOT NULL Or NVL(B.执行状态,1) IN (0,2))"
39        ElseIf miInputType = 4 Then '挂号单
40            strNO = GetFullNO(strSeek, 12)
41            strSQL = "Select 1 As PatientType,0 As 主页ID,Nvl(B.执行部门ID,0) As 病人科室,B.执行人 As 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A,门诊费用记录 B " & _
                  " Where B.记录性质=4 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID"
42        ElseIf miInputType = 5 Then '收费单据号
43            strNO = GetFullNO(strSeek, 13): mstrNO = strNO
              
44            strSQL = "Select 1 As PatientType,0 As 主页ID,B.开单部门ID As 病人科室,B.开单人 As 医生,B.姓名,B.性别,B.年龄,a.住院号,a.当前床号," & _
                  "A.病人ID,A.单位电话,A.工作单位,A.单位邮编,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.门诊号,A.身份证号,A.费别,A.医疗付款方式," & _
                  "A.国籍,A.婚姻状况,A.民族,A.职业 From 病人信息 A,门诊费用记录 B" & _
                  " Where Mod(B.记录性质,10)=1 And B.记录状态 IN(1,3) And B.NO=[3] And B.病人ID=A.病人ID(+) Order By B.病人ID" ' And B.医嘱序号 Is Null"
45        Else '当作姓名
46            strSQL = "Select Decode(A.当前科室id,Null,1,2) As PatientType,A.主页ID,Nvl(A.当前科室id,0) As 病人科室,'' 医生," & mConst_病人信息_列名 & _
                  " From 病人信息 A Where A.姓名=[1] and 1 = 2 " '所有输入姓名的病人当新病人处理
47        End If
          
48        Set GetPatient = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, strSeek, Val(strSeek), strNO)


49        Exit Function
GetPatient_Error:
50        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(GetPatient)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
51        Err.Clear

End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'功能：从指定的字典中读取数据
'参数：strDict=字典对应的表名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
        
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From " & strDict & " Order by 编码"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp

End Function
Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
      '功能：由用户输入的部份单号，返回全部的单号。
      '参数：intNum=项目序号,为0时固定按年产生
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, intType As Integer
          Dim curDate As Date
          
1         On Error GoTo GetFullNO_Error

2         If Len(strNO) >= 8 Then
3             GetFullNO = Right(strNO, 8)
4             Exit Function
5         ElseIf Len(strNO) = 7 Then
6             GetFullNO = PreFixNO & strNO
7             Exit Function
8         ElseIf intNum = 0 Then
9             GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
10            Exit Function
11        End If
12        GetFullNO = strNO
          
13        strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
15        If Not rsTmp.EOF Then
16            intType = NVL(rsTmp!编号规则, 0)
17            curDate = rsTmp!日期
18        End If

19        If intType = 1 Then
              '按日编号
20            strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
21            GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
22        Else
              '按年编号
23            GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
24        End If


25        Exit Function
GetFullNO_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(GetFullNO)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear

End Function

Private Function SelectCap(Optional ByVal lngItemid As Long = 0) As ADODB.Recordset
      '获取采集方式
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset



1         On Error GoTo SelectCap_Error

2         strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                   "From 诊疗项目目录 A,诊疗用法用量 D Where A.ID=D.用法ID" + _
                 " And A.类别='E' And A.操作类型='6'" & _
                 " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                 " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
                   IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
                 " And Nvl(A.执行频率,0) IN(0,1)" + _
                 " And D.项目ID=" & lngItemid
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
4         If rsTmp.EOF Then
5             strSQL = "Select Distinct A.ID,A.编码,A.名称 " + _
                       "From 诊疗项目目录 A Where " + _
                     " A.类别='E' And A.操作类型='6'" & _
                     " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " + _
                     " And A.服务对象 IN(" & PatientType & ",3) And Nvl(A.适用性别,0) IN (" + _
                       IIf(Me.cbo性别.Text Like "*男*", "1,0)", "2,0)") + _
                     " And Nvl(A.执行频率,0) IN(0,1)"
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
7         End If

8         If Not rsTmp.EOF Then Set SelectCap = rsTmp


9         Exit Function
SelectCap_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(SelectCap)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear

End Function

Private Sub AdviceSet检查手术(ByVal int类型 As Integer, ByVal strDataIDs As String)
      '功能：1.重新设置指定检查组合项目的部位行,用于新输入检查组合项目或修改部位
      '      2.重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
      '参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
      '      strDataIDs=检查:包含检查部位信息,手术:包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
          Dim strSQL As String

                  
          '处理检验项目
1         On Error GoTo AdviceSet检查手术_Error

2         strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
          
3         If strDataIDs <> "" Then
4             If Not mrsRelativeAdvice Is Nothing Then
5                 mrsRelativeAdvice.Close
6             Else
7                 Set mrsRelativeAdvice = New ADODB.Recordset
8             End If
9             strSQL = "Select ID,编码,名称,nvl(标本部位,' ') As 标本部位," + _
              "类别,nvl(计价性质,0) As 计价性质,nvl(执行科室,0) As 执行科室,操作类型 From 诊疗项目目录 Where ID IN(" & strDataIDs & ")"
10            Set mrsRelativeAdvice = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
11        Else
12            If Not mrsRelativeAdvice Is Nothing Then mrsRelativeAdvice.Close: Set mrsRelativeAdvice = Nothing
13        End If


14        Exit Sub
AdviceSet检查手术_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(AdviceSet检查手术)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
16        Err.Clear

End Sub

Private Function Get检查手术名称(ByVal int类型 As Integer, ByVal txtMainAdvice As String) As String
      '功能：重新生成检查手术内容的医嘱内容
      '参数：int类型=1=处理检查部位项目,2=处理附加手术及麻醉项目
          Dim lngBegin As Long
          Dim strTmp As String

1         On Error GoTo Get检查手术名称_Error

2         If mrsRelativeAdvice Is Nothing Or int类型 = 1 Then Get检查手术名称 = txtMainAdvice: Exit Function
              
3         mrsRelativeAdvice.MoveFirst
4         Do While Not mrsRelativeAdvice.EOF
5             If Len(Trim(mrsRelativeAdvice("名称"))) > 0 Then
6                 strTmp = strTmp & "," & mrsRelativeAdvice("名称")
7             End If
              
8             mrsRelativeAdvice.MoveNext
9         Loop
          
10        If strTmp <> "" Then
11            Get检查手术名称 = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " 及 ") & Mid(strTmp, 2)
12        Else
13            Get检查手术名称 = txtMainAdvice
14        End If


15        Exit Function
Get检查手术名称_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(Get检查手术名称)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear
End Function

'检查医嘱内容的合法性
Private Function ValidAdvice() As Boolean
1         On Error GoTo ValidAdvice_Error

2         ValidAdvice = True
          
3         On Error Resume Next
4         If txt姓名.Text = "" Then
5             ValidAdvice = False
6             MsgBox "请输入病人的姓名！", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.姓名
7             txt姓名.SetFocus: Exit Function
8         End If
          
      '    If Len(Trim(Me.txt医嘱内容)) = 0 Then
      '        ValidAdvice = False
      '        MsgBox "必须输入申请项目！", vbInformation, Me.Caption: DoEvents
      ''        mintFocusItem = FocusItem.医嘱内容
      '        Me.txt医嘱内容.SetFocus: Exit Function
      '    End If
9         If Me.cbo开单科室.ListIndex = -1 Then
10            ValidAdvice = False
11            MsgBox "请指定开单科室！", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.开单科室
12            Me.cbo开单科室.SetFocus: Exit Function
13        End If
14        If Me.cbo执行科室.ListIndex = -1 Then
15            ValidAdvice = False
16            MsgBox "请指定执行科室!", vbInformation, Me.Caption: DoEvents
17            Me.cbo执行科室.SetFocus: Exit Function
18        End If
19        If Len(Trim(Me.cbo医生.Text)) = 0 Then
20            ValidAdvice = False
21            MsgBox "请指定开单医生！", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.医生
22            Me.cbo医生.SetFocus: Exit Function
23        End If


24        Exit Function
ValidAdvice_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(ValidAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
26        Err.Clear
End Function

Private Function SaveAdviceData(ByRef strNewAdvice As String) As Boolean
          Dim strSQL As String, strDate As String, strNO As String
          Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
          Dim iMaxSeq As Integer, iSendSeq As Integer
          Dim rsTmp As New ADODB.Recordset
          Dim lng开嘱科室ID As Long, strDoctor As String, i As Integer
          Dim str执行科室id As String, str执行科室ID1 As String
          Dim tmpstr类别 As String, tmplngClinicID As Long, tmpint计价特性 As Integer
          Dim lngJ As Long, strCostType As String

          Dim strAge As String
          Dim strInfo As String
          Dim lngTmp As Long

1         On Error GoTo SaveAdviceData_Error

          '保存病人信息
2         strDate = "To_Date('" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
3         If PatientType = 1 Then    '门诊病人
4             If mlng病人ID > 0 Then    '已有的病人
                  '            strSQL = _
                               "zl_挂号病人病案_INSERT(3," & mlng病人ID & ",Null," & _
                               "'',''," & _
                               "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
                               "'自费','自费'," & _
                               "'','',''," & _
                               "'','','',0,'','','','',''," & strDate & ",NULL)"
5             Else    '新病人
6                 If txt年龄.Locked = False Then
7                     strAge = txt年龄.Text
8                     If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt年龄1.Text
9                     strInfo = CheckAge(strAge)
10                    If InStr(1, strInfo, "|") > 0 Then
11                        lngTmp = Val(Split(strInfo, "|")(0))    '1禁止,0提示
12                        strInfo = Split(strInfo, "|")(1)
13                        If lngTmp = 1 Then
14                            MsgBox strInfo, vbInformation, Me.Caption
15                            If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus: Exit Function
16                        End If
17                    End If
18                End If
                  '添加获取默认费别
19                strSQL = "select 名称,缺省标志 from 费别 order by 编码"
20                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlLisWork")
21                Do While Not rsTmp.EOF
22                    lngJ = lngJ + 1
23                    If lngJ = 1 Then
24                        strCostType = rsTmp("名称")
25                    End If
26                    If rsTmp("缺省标志") = 1 Then
27                        strCostType = rsTmp("名称")
28                        Exit Do
29                    End If
30                    rsTmp.MoveNext
31                Loop
32                If strCostType = "" Then strCostType = "自费"

33                mlng病人ID = GetNextNo(Sel_His_DB, 1)
34                ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
35                mstrSQLPro(UBound(mstrSQLPro)) = "zl_挂号病人病案_INSERT(1," & mlng病人ID & ",Null," & _
                                                   "'',''," & _
                                                   "'" & txt姓名.Text & "','" & NeedName(cbo性别.Text) & "','" & txt年龄.Text & Me.cboAge.Text & Me.txt年龄1.Text & "'," & _
                                                   "'" & strCostType & "','" & strCostType & "'," & _
                                                   "'','',''," & _
                                                   "'','','" & Me.txtUnit.Text & "',0,'','','','',''," & strDate & ",NULL)"
36            End If
37        End If
          '保存医嘱并发送
38        lngAdviceID = GetNextId("病人医嘱记录")
39        iMaxSeq = 0

40        lng开嘱科室ID = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
41        strDoctor = NeedName(Me.cbo医生.Text)

42        If mrsRelativeAdvice.RecordCount = 0 Then
43            str执行科室id = mlngDeptID
44        Else
              'PatientType
45            If mlng病人ID > 0 Then
46                strSQL = "select  执行科室ID from  诊疗执行科室 where 病人来源 = [1] and 诊疗项目ID = [2] "
47            End If
48            mrsRelativeAdvice.MoveFirst
49            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, PatientType, CLng(mrsRelativeAdvice("Id")))
50            If Not rsTmp.EOF Then str执行科室id = Val(NVL(rsTmp("执行科室ID")))
51        End If

          '选择了执行科室按执行科室进行
52        If Me.cbo执行科室.Text <> "" Then
53            str执行科室id = Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex)
54        End If

55        iSendSeq = 1
          '检验项目将采集方式作为主医嘱
56        tmplngClinicID = mlngCapID
          '取采集方式的执行部门
57        str执行科室ID1 = gUserInfo.DeptID

58        lngSendNO = GetNextNo(Sel_His_DB, 10)
59        strNO = GetNextNo(Sel_His_DB, IIf(PatientType = 2, 14, 13))

          '保存相关医嘱
60        If Not mrsRelativeAdvice Is Nothing Then
61            i = 2
62            mrsRelativeAdvice.MoveFirst
63            Do While Not mrsRelativeAdvice.EOF
64                lngTmpID = GetNextId("病人医嘱记录")
65                With mrsRelativeAdvice
66                    strNewAdvice = strNewAdvice & "," & lngTmpID
67                    ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
68                    mstrSQLPro(UBound(mstrSQLPro)) = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                                                       (iMaxSeq + i) & ",3," & mlng病人ID & ",NULL," & _
                                                       "0,1," & _
                                                       "1,'" & .Fields("类别") & "'," & _
                                                       .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                                                       "'" & Replace(.Fields("名称"), "'", "''") & "',''," & _
                                                       "'" & .Fields("标本部位") & "','一次性',NULL,NULL,'',NULL," & _
                                                       .Fields("计价性质") & "," & _
                                                       str执行科室id & "," & _
                                                       .Fields("执行科室") & ",0," & strDate & ",NULL," & _
                                                       IIf(Val(Me.txtPatientDept.Tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.Tag)) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                                                       "Sysdate,'',Null)"
69                    iSendSeq = iSendSeq + 1

70                    ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
71                    mstrSQLPro(UBound(mstrSQLPro)) = "ZL_病人医嘱发送_Insert(" & _
                                                       lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                                                       iSendSeq & ",NULL,NULL,NULL," & _
                                                       "Sysdate+1/(24*3600)," & _
                                                       "0," & str执行科室id & ",0,0)"
72                    i = i + 1
73                    .MoveNext
74                End With
75            Loop
76        End If
          '检验申请的采集方式放到最后
77        iMaxSeq = iMaxSeq + 1
78        strNewAdvice = strNewAdvice & "," & lngAdviceID
79        ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
80        mstrSQLPro(UBound(mstrSQLPro)) = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & _
                                           iMaxSeq & ",3," & mlng病人ID & ",NULL," & _
                                           "0,1," & _
                                           "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
                                           "'" & Replace(mstrItemSel, "'", "''") & "',''," & _
                                           "'','一次性',NULL,NULL,'',NULL,2," & _
                                           str执行科室ID1 & ",3,0," & strDate & ",NULL," & _
                                           IIf(Val(Me.txtPatientDept.Tag) = 0, lng开嘱科室ID, Val(Me.txtPatientDept.Tag)) & "," & lng开嘱科室ID & ",'" & strDoctor & "'," & _
                                           "Sysdate,'',Null)"
81        iSendSeq = iSendSeq + 1
          '发送主医嘱
82        ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
83        mstrSQLPro(UBound(mstrSQLPro)) = "ZL_病人医嘱发送_Insert(" & _
                                           lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                                           iSendSeq & ",NULL,NULL,NULL," & _
                                           "Sysdate+1/(24*3600)," & _
                                           "0," & str执行科室id & ",0,1)"


84        If strNewAdvice <> "" Then strNewAdvice = Mid(strNewAdvice, 2)
85        SaveAdviceData = True

86        Exit Function
SaveAdviceData_Error:
87        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(SaveAdviceData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
88        Err.Clear

End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Sub TxtSelAll(ByVal objTxt As Object)
    With objTxt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'功能：获取指定控件在屏幕中的位置(Twip/Pixel)
'返回：blnTwip=True-返回Twip单位，False-返回像素单位
    Dim vRect As RECT
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Private Function CboLocate(ByVal cboobj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'建议弃用，使用Cbo.SeekIndex代替
'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboobj.ListCount - 1
        If blnItem Then
            If cboobj.ItemData(lngLocate) = Val(strValue) Then
                cboobj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboobj.List(lngLocate), InStr(1, cboobj.List(lngLocate), "-") + 1) = strValue Then
                cboobj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789<>", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.-<>+Ee", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, Me.Caption
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Function FindComboItem(objCombox As Object, ByVal lngFind As Long) As Integer
    Dim i As Integer
    
    For i = 0 To objCombox.ListCount - 1
        If objCombox.ItemData(i) = lngFind Then Exit For
    Next
    If i > objCombox.ListCount - 1 Then i = -1
    
    FindComboItem = i
End Function

Private Function BlnIsNumber(ByVal strCode As String) As Boolean
    '数字，及条码判断
     If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        BlnIsNumber = True
     Else
        BlnIsNumber = False
     End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "") As String
    '功能:年龄合法性检查
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay))
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = NVL(rsTemp.Fields(0).value)

End Function

Private Sub vsfGroup_RowColChange()
          Dim lngRow As Long
          Dim lngCol As Long
          Dim lngGroupId As Long      '分组ID
          Dim lngAppID As Long        '申请单ID
          Dim strSampleType As String '标本类型
          Dim strErr As String
          
1         On Error GoTo vsfGroup_RowColChange_Error
          
          '获取选择的标本
2         strSampleType = IIf(Trim(Me.cboSampleType.Text) = "所有标本", "", Trim(Me.cboSampleType.Text))

3         With Me.vsfGroup
4             lngRow = .Row
5             lngCol = .Col
6             If lngRow <= 1 Or lngCol < 0 Then Exit Sub
7             If .Cell(flexcpFontBold, lngRow, lngCol) = True Then Exit Sub
8             If mrsItem Is Nothing Then Exit Sub
9             If mrsItem.RecordCount <= 0 Then Exit Sub
10            mrsItem.MoveFirst
11            lngAppID = Val(.TextMatrix(lngRow, .ColIndex("申请单ID")))      '申请单ID
12            lngGroupId = Val(.TextMatrix(lngRow, .ColIndex("ID")))          '分组ID
13            mrsItem.Filter = "申请单ID=" & lngAppID & " and 分组ID=" & lngGroupId & IIf(strSampleType = "", "", " and 标本='" & strSampleType & "'")
14            If vfgLoadFromRecord(vsfItem, mrsItem, strErr) = False Then
15                If strErr <> "" Then
16                    MsgBox strErr, vbInformation, Me.Caption
17                    mrsItem.Filter = ""
18                    Exit Sub
19                End If
20            End If
              
21        End With
22        With Me.vsfItem
23            .ExtendLastCol = True
24            .ColDataType(.ColIndex("选择")) = flexDTBoolean
25            .ExplorerBar = flexExSortShow
              
              
26            .ColHidden(.ColIndex("ID")) = True
27            .ColHidden(.ColIndex("申请单ID")) = True
28            .ColHidden(.ColIndex("分组ID")) = True
29            .ColHidden(.ColIndex("诊疗编码")) = True
30            .ColHidden(.ColIndex("标本")) = True
              
31            .ColWidth(.ColIndex("选择")) = 500
32            .ColWidth(.ColIndex("编码")) = 1000
33            .ColWidth(.ColIndex("名称")) = 4000
34            .ColWidth(.ColIndex("简码")) = 1500
              
35            Call CheckSelItem
              
36        End With
              
37        mrsItem.Filter = ""


38        Exit Sub
vsfGroup_RowColChange_Error:
39        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(vsfGroup_RowColChange)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
40        Err.Clear

End Sub

Private Sub VSFItem_Click()
          Dim lngRow As Long
          Dim lngCol As Long
          
1         On Error GoTo VSFItem_Click_Error

2         With Me.vsfItem
3             lngRow = .MouseRow
4             lngCol = .MouseCol
5             .Editable = flexEDNone
              
6             If lngRow > 0 And lngCol = .ColIndex("选择") Then
7                 If .Cell(flexcpChecked, lngRow, lngCol) = 1 Then
8                     .Cell(flexcpChecked, lngRow, lngCol) = 0
9                     Call selOrDelItem(2, lngRow)
10                Else
11                    .Cell(flexcpChecked, lngRow, lngCol) = 1
12                    Call selOrDelItem(1, lngRow)
13                End If
14            End If
15        End With


16        Exit Sub
VSFItem_Click_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(VSFItem_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear
End Sub

Private Sub VSFSeled_DblClick()
    With Me.VSFSeled
        If .Row < 0 Or .Col < 0 Then Exit Sub
        Call selOrDelItem(3, .Row)
        Call CheckSelItem
    End With
End Sub


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/27
'功    能:选择或者取消选择项目
'入    参:
'           intType     1=点击选择,2=点击取消选择，3=双击VSFSeled取消选择
'           lngSelRow   中选择的行
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub selOrDelItem(ByVal intType As Integer, ByVal lngSelRow As Long)
          Dim lngRow As Long
          
1         On Error GoTo selOrDelItem_Error

2         With VSFSeled
3             If intType = 1 Then
4                 .Rows = .Rows + 1
5                 .TextMatrix(.Rows - 1, .ColIndex("ID")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("ID"))
6                 .TextMatrix(.Rows - 1, .ColIndex("名称")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("名称"))
7                 .TextMatrix(.Rows - 1, .ColIndex("诊疗编码")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("诊疗编码"))
8                 .TextMatrix(.Rows - 1, .ColIndex("oldid")) = GetOldID(.TextMatrix(.Rows - 1, .ColIndex("诊疗编码")))
9                 .TextMatrix(.Rows - 1, .ColIndex("oldName")) = GetOldName(.TextMatrix(.Rows - 1, .ColIndex("诊疗编码")))
10                .TextMatrix(.Rows - 1, .ColIndex("标本")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("标本"))
11            ElseIf intType = 2 Then
12                For lngRow = 0 To .Rows - 1
13                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("ID")) Then
14                        .RemoveItem lngRow
15                        Exit Sub
16                    End If
17                Next
18            ElseIf intType = 3 Then
19                .RemoveItem lngSelRow
20            End If
              '按照标本排序
21            If .Rows > 0 Then .Cell(flexcpSort, .FixedRows, .ColIndex("标本"), .Rows - 1, .ColIndex("标本")) = 2
22        End With


23        Exit Sub
selOrDelItem_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(selOrDelItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/27
'功    能:切换分组时勾选已经选择的项目
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub CheckSelItem()
          Dim lngLoop As Long
          Dim lngRow As Long
          
          '先取消选择
1         On Error GoTo CheckSelItem_Error

2         With vsfItem
3             .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
4         End With
          
          '再选择
5         With Me.VSFSeled
6             For lngLoop = 0 To .Rows - 1
7                 For lngRow = 1 To vsfItem.Rows - 1
8                     If vsfItem.TextMatrix(lngRow, vsfItem.ColIndex("ID")) = .TextMatrix(lngLoop, .ColIndex("ID")) Then
9                         vsfItem.Cell(flexcpChecked, lngRow, vsfItem.ColIndex("选择")) = 1
10                    End If
11                Next
12            Next
13        End With


14        Exit Sub
CheckSelItem_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "执行(CheckSelItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
16        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/27
'功    能:通过诊疗编码获取老版项目ID
'入    参:
'           strCode     诊疗编码
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Function GetOldID(ByVal strCode As String) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetOldID_Error

2         strSQL = "select ID from 诊疗项目目录 where 编码=[1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strCode)
4         If rsTmp.RecordCount > 0 Then GetOldID = Val(rsTmp("ID") & "")


5         Exit Function
GetOldID_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(GetOldID)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/27
'功    能:通过诊疗编码获取老版项目名称
'入    参:
'           strCode     诊疗编码
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Function GetOldName(ByVal strCode As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetOldID_Error

2         strSQL = "select 名称 from 诊疗项目目录 where 编码=[1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strCode)
4         If rsTmp.RecordCount > 0 Then GetOldName = rsTmp("名称") & ""


5         Exit Function
GetOldID_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "执行(GetOldID)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/21
'功    能:设置文本框的提示字
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub setTxtTip(objTxt As TextBox, Optional ByVal strTip As String)
    On Error Resume Next
    With objTxt
        If .Text <> "" Then Exit Sub
        .ToolTipText = strTip
        .Text = strTip
        .ForeColor = &H80000002
        .Tag = "T"
    End With
End Sub



