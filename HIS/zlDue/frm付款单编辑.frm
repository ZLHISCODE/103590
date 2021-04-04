VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm付款编辑 
   Caption         =   "付款"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frm付款单编辑.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10125
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2610
      TabIndex        =   24
      Top             =   660
      Width           =   2010
   End
   Begin MSComctlLib.ImageList ilt24 
      Left            =   870
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm付款单编辑.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   7080
      ScaleHeight     =   2340
      ScaleWidth      =   2430
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3945
      Width           =   2430
      Begin VSFlex8Ctl.VSFlexGrid vsFp 
         Height          =   1800
         Left            =   60
         TabIndex        =   18
         Top             =   390
         Width           =   2385
         _cx             =   4207
         _cy             =   3175
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm付款单编辑.frx":0E64
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
         ExplorerBar     =   5
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
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin XtremeSuiteControls.ShortcutCaption stcFpTittle 
         Height          =   375
         Left            =   -30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         _Version        =   589884
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "临时信息-发票汇总"
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
      End
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   5025
      ScaleHeight     =   2340
      ScaleWidth      =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1860
      Begin VSFlex8Ctl.VSFlexGrid vsTemp 
         Height          =   1920
         Left            =   150
         TabIndex        =   15
         Top             =   375
         Width           =   1845
         _cx             =   3254
         _cy             =   3387
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm付款单编辑.frx":0EF9
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
         ExplorerBar     =   5
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
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTempTittle 
         Height          =   375
         Left            =   -30
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1860
         _Version        =   589884
         _ExtentX        =   3281
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "临时数据-分段汇总"
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
      End
   End
   Begin VB.PictureBox pic预付 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   45
      ScaleHeight     =   2235
      ScaleWidth      =   4905
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   4905
      Begin VSFlex8Ctl.VSFlexGrid vs预付 
         Height          =   1770
         Left            =   0
         TabIndex        =   6
         Top             =   390
         Width           =   4845
         _cx             =   8546
         _cy             =   3122
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm付款单编辑.frx":0F48
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
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "冲预交:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   3615
         TabIndex        =   12
         Top             =   105
         Width           =   630
      End
      Begin VB.Label lbl金额 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交累计:"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   3
         Left            =   1305
         TabIndex        =   11
         Top             =   90
         Width           =   810
      End
      Begin XtremeSuiteControls.ShortcutCaption stc预付 
         Height          =   375
         Left            =   -30
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   4890
         _Version        =   589884
         _ExtentX        =   8625
         _ExtentY        =   661
         _StockProps     =   6
         Caption         =   "预付清单"
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
      End
   End
   Begin VB.PictureBox picPayList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2460
      Left            =   0
      ScaleHeight     =   2460
      ScaleWidth      =   9600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1380
      Width           =   9600
      Begin VSFlex8Ctl.VSFlexGrid vsPayList 
         Height          =   1425
         Left            =   45
         TabIndex        =   3
         Top             =   810
         Width           =   8295
         _cx             =   14631
         _cy             =   2514
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   12632256
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm付款单编辑.frx":0FE8
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
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.PictureBox picCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   -30
         ScaleHeight     =   780
         ScaleWidth      =   9570
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   9570
         Begin VB.CommandButton cmdSelDept 
            Caption         =   "…"
            Height          =   255
            Left            =   4410
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   255
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            Left            =   675
            TabIndex        =   1
            Top             =   60
            Width           =   4005
         End
         Begin VB.Label lbl金额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "本次应付:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   5
            Left            =   8640
            TabIndex        =   23
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lbl金额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "累计应付:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   1
            Left            =   3450
            TabIndex        =   22
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lbl金额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款金额:"
            ForeColor       =   &H00000040&
            Height          =   180
            Index           =   2
            Left            =   6150
            TabIndex        =   21
            Top             =   510
            Width           =   810
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   4875
            TabIndex        =   13
            Top             =   120
            Width           =   90
         End
         Begin VB.Label lbl供应商 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "供应商"
            Height          =   180
            Left            =   75
            TabIndex        =   0
            Top             =   135
            Width           =   540
         End
         Begin XtremeSuiteControls.ShortcutCaption stcPayTittle 
            Height          =   330
            Left            =   90
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   450
            Width           =   9525
            _Version        =   589884
            _ExtentX        =   16801
            _ExtentY        =   582
            _StockProps     =   6
            Caption         =   "未付款清单"
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
         End
         Begin XtremeSuiteControls.ShortcutCaption stcTop 
            Height          =   450
            Left            =   75
            TabIndex        =   20
            Top             =   0
            Width           =   9480
            _Version        =   589884
            _ExtentX        =   16722
            _ExtentY        =   794
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
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   6375
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm付款单编辑.frx":130C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12779
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   435
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frm付款单编辑.frx":1BA0
      Left            =   -15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm付款编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmEdit As frmPayNoEdit
Attribute mfrmEdit.VB_VarHelpID = -1

Private mintStep As Integer
Private mstrFindKey As String       '查找主键
Private mstrNo As String                   '单据号
Private mlng单位ID As Long
Private mblnFirst As Boolean
Private mblnChange As Boolean
Private mblnSave As Boolean
Private mfrmMain  As Object
Private mEditType As gEditType
Private mint记录状态 As RecBillStatus       '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mErrBillStatusInfor As ErrBillStatusInfor       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mblnEdit As Boolean             '编辑状态
Private mblnSuccess As Boolean          '是否有单据保存成功
Private mstrPrivs  As String
Private mlng付款序号 As Long    '付款序号

Private mdbl累计应付 As Double
Private mdbl本次应付 As Double
Private mdbl本次预交 As Double
Private mdbl累计预交 As Double
Private mlng效期 As Long
Private mstrFind As String
Private mcllFilter As Collection
Private mcbrToolBar As CommandBar
Private mcbrMenuBar As CommandBarPopup
Private mcbrControl As CommandBarControl
Private mobjFindKey As CommandBarControl
Private mint统计方式 As Integer     '0-按随货同行单号分类统计,1-按发票号统计
Private mint标记 As Integer
Private mbln付款标志 As Boolean
Private mint显示单位 As Integer     '0：最小单位；  1：最大单位

Private Const mConMenu_Hide = 99
Private Const mConMenu_Hide_TempSave = 9981
Private Const mConMenu_Hide_TempClearAll = 9982
Private Const mConMenu_Popu = 88
Private Const mConMenu_Popu_FP = 8801
Private Const mConMenu_Popu_SH = 8802
Private Const mConMenu_Report = 102

Private Const mlngModule = 1323
Private Enum mPanIndex
    pane_应付列表 = 0
    pane_预交列表 = 1
    pane_付款单 = 2
    pane_临时数据 = 3
    pane_发票合计 = 4
End Enum

Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim cbrCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mConMenu_Hide, "隐藏(&D)", -1, False)
    mcbrMenuBar.Visible = False
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempSave, "保存临时信息(&S)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Save, IIf(mEditType = g预审, "预审(&O", "保存(&O)"))
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "审核(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "冲销(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_FilterView, "重置(&F)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
 
    mint统计方式 = Val(zlDatabase.GetPara("统计方式", glngSys, mlngModule, "1"))
        
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mConMenu_Popu, "统计方式(&T)", -1, False)
    mcbrMenuBar.Visible = False
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Popu_FP, "按发票号进行分类统计(&1)")
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Popu_SH, "按随货单号进行分类统计(&2)")
    End With
        
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FALT, Asc("O"), conMenu_Edit_Save
        .Add FALT, Asc("V"), conMenu_Manage_Audit
        .Add FALT, Asc("S"), conMenu_Edit_ChargeOff
        .Add FALT, Asc("R"), conMenu_View_FilterView
        .Add FCONTROL, Asc("S"), mConMenu_Hide_TempSave
        .Add FCONTROL, Asc("C"), mConMenu_Hide_TempClearAll
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清")
        
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempSave, "保存临时信息"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3200
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Hide_TempClearAll, "清除临时信息"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3702
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Forward, "上一步"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Backward, "下一步")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Save, IIf(mEditType = g预审, "预审", "保存")): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "审核"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "冲销"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_FilterView, "重置"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 254
        Set mcbrControl = .Add(xtpControlButton, mConMenu_Report, "药品付款查询"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
     
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        mcbrControl.Flags = xtpFlagRightAlign
        mstrFindKey = Trim(zlDatabase.GetPara("定位依据", glngSys, mlngModule, "随货单号"))
        
        Set mobjFindKey = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
       ' mobjFindKey.BeginGroup = True
        mobjFindKey.IconId = conMenu_View_Find
        mobjFindKey.Flags = xtpFlagRightAlign
        mobjFindKey.Style = xtpButtonIconAndCaption
        If mstrFindKey = "" Then mstrFindKey = "随货单号"
        
        With mobjFindKey.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&1.随货单号")
            mcbrControl.Parameter = "随货单号"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&2.发票号")
            mcbrControl.Parameter = "发票号"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&3.入库单号")
            mcbrControl.Parameter = "入库单号"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&4.品名")
            mcbrControl.Parameter = "品名"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&5.数量")
            mcbrControl.Parameter = "数量"
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&6.发票金额")
            mcbrControl.Parameter = "发票金额"
            
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "&7.审核人")
            mcbrControl.Parameter = "审核人"
        End With
    
        Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.Handle = txtFind.hwnd
        cbrCustom.Flags = xtpFlagRightAlign
        
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SaveTempData(Optional blnClsAll As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '功能:可存临时数据到网格
    '入参:blnClsAll-是否清除历史存储的信息
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-20 11:36:23
    '-----------------------------------------------------------------------------------------------------------
    Dim dbl金额 As Double, i As Long, lngMax序号 As Long, blnHaveData As Boolean
    With vsTemp
        If blnClsAll Then
            .Clear 1
            .Rows = 2
        Else
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("序号"))) > lngMax序号 Then lngMax序号 = Val(.TextMatrix(i, .ColIndex("序号")))
            Next
            lngMax序号 = lngMax序号 + 1
        End If
    End With
    
    blnHaveData = False
    With vsPayList
        dbl金额 = 0
        If blnClsAll Then
            .Cell(flexcpText, 1, .ColIndex("汇总序号"), .Rows - 1, .ColIndex("汇总序号")) = ""
            Exit Sub
        Else
            For i = 1 To .Rows - 1
                .Redraw = flexRDNone
                If Trim(.TextMatrix(i, .ColIndex("选定"))) = "√" And Val(.TextMatrix(i, .ColIndex("汇总序号"))) = 0 Then
                    dbl金额 = dbl金额 + Val(.Cell(flexcpData, i, .ColIndex("发票金额")))
                    .TextMatrix(i, .ColIndex("汇总序号")) = lngMax序号
                    blnHaveData = True
                End If
                .Redraw = flexRDBuffered
            Next
        End If
    End With
    If blnHaveData Then
        '填临时汇总数
        With vsTemp
            If Val(.TextMatrix(.Rows - 1, .ColIndex("序号"))) <> 0 Then
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = lngMax序号
            .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(dbl金额, gVbFmtString.FM_金额)
            .Cell(flexcpData, .Rows - 1, .ColIndex("金额")) = dbl金额
        End With
    End If
End Sub

Private Sub InitPancel()
    '功能:初始化区域控件:2008-07-14 15:04:29
    Dim panThis As Pane
        
    Set mfrmEdit = New frmPayNoEdit
    Load mfrmEdit
    '问题27930 by lesfeng 2010-03-23
    mfrmEdit.zlInitPara Me, mlngModule, mstrPrivs, mint标记
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_付款单, 250, 580, DockTopOf, Nothing)
    panThis.Title = "付款通知单"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.Close
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_应付列表, 250, 580, DockTopOf, Nothing)
    panThis.Title = "付款清单"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_预交列表, 250, 580, DockBottomOf, panThis)
    panThis.Title = "预付款清单"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_临时数据, 124, 580, DockRightOf, panThis)
    panThis.Title = "应付临时数据"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.MinTrackSize.Width = 124
    panThis.MaxTrackSize.Width = 348
    
    Set panThis = dkpMan.CreatePane(mPanIndex.pane_发票合计, 124, 580, DockRightOf, panThis)
    panThis.Title = "应付发票合计信息"
    panThis.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    panThis.MinTrackSize.Width = 124
    panThis.MaxTrackSize.Width = 348
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.UseSplitterTracker = False '实时拖动
    dkpMan.Options.AlphaDockingContext = True
    Me.dkpMan.Options.HideClient = True
End Sub

Private Sub InitFilter()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化过滤条件
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-18 16:10:40
    '-----------------------------------------------------------------------------------------------------------
    Set mcllFilter = New Collection
    mcllFilter.Add mlng单位ID, "供应商ID"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "审核日期"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "发票日期"
    mcllFilter.Add "", "发票号列表"
    mcllFilter.Add "", "随货单号列表"
    mcllFilter.Add "", "系统标识"
    mcllFilter.Add "", "品名"
    mcllFilter.Add "", "规格"
    mcllFilter.Add "", "产地"
    mcllFilter.Add Array("", ""), "批号"
    mcllFilter.Add Array("", ""), "入库单号"
    mcllFilter.Add Array("", ""), "发票号"
    mcllFilter.Add Array("", ""), "随货单号"
    mcllFilter.Add "", "填制人"
    mcllFilter.Add "", "审核人"
    mcllFilter.Add "", "过滤"
    mcllFilter.Add "", "库房"
    mcllFilter.Add "0", "按所有药品库存数量小于发票数量"
    mstrFind = ""
End Sub
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strSQL As String
    Dim rsTemp As New Recordset
    
    GetDepend = False
    
    '读取结算方式
    Err = 0: On Error GoTo ErrHand:
    
    strSQL = "Select 1 From 结算方式应用 Where 应用场合='付货款' and rownum<=1 Order by 缺省标志 desc"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "结算方式应用信息不全,请在结算方式管理中进行设置！"
        Exit Function
    End If
    GetDepend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub initGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化网格控件的默认属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-20 11:11:42
    '-----------------------------------------------------------------------------------------------------------
    With vsPayList
        .Cell(flexcpPicture, 0, .ColIndex("汇总序号"), 0, .ColIndex("汇总序号")) = ilt24.ListImages(1).Picture
        .Cell(flexcpPictureAlignment, 0, .ColIndex("汇总序号"), 0, .ColIndex("汇总序号")) = 4
        .Cell(flexcpAlignment, 0, .ColIndex("汇总序号"), 0, .ColIndex("汇总序号")) = 4
        If (mEditType = g新增 Or mEditType = g修改) And InStr(1, mstrPrivs, ";修改发票信息;") > 0 Then
            '有此权限时，才允许修改发票信息
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub initCard()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化卡片信息
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 13:39:41
    '-----------------------------------------------------------------------------------------------------------
    Dim intErrInfor As Integer
    Call initGrid
    '初始表格
    If mfrmEdit.zlLoadData(mEditType, mlng单位ID, mstrNo, mint记录状态, intErrInfor) = False Then
        If intErrInfor = 1 Then
            mErrBillStatusInfor = 已经删除
        ElseIf intErrInfor = 2 Then
            mErrBillStatusInfor = 已经审核
        End If
        Exit Sub
    End If
    If mEditType = g新增 Then Exit Sub
    Call LoadPayMoney
    Call 汇总发票信息
    SetEditPro
End Sub

Public Sub ShowCard(FrmMain As Form, _
    ByVal int编辑状态 As gEditType, ByVal strPrivs As String, _
    Optional strNO As String = "", _
    Optional lng单位ID As Long = 0, _
    Optional int记录状态 As RecBillStatus = 1, _
    Optional blnSuccess As Boolean = False, _
    Optional int标记 As Integer = 0)
    
    mstrNo = strNO
    mblnSave = False
    mblnSuccess = False
    mEditType = int编辑状态
    mint记录状态 = int记录状态
    mstrPrivs = strPrivs

    mlng单位ID = lng单位ID
    mint标记 = int标记
    
    mblnChange = False
    mErrBillStatusInfor = 正常情况
    Set mfrmMain = FrmMain
    '问题27930 by lesfeng 2010-03-23
    If int标记 = 1 Then
        Me.Caption = "标记付款"
    End If
    '初始化过滤条件:2008-08-18 17:48:29
    Call InitFilter
    
    '检查数据依赖关系
    If Not GetDepend Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
     
    If mEditType = g新增 Then
        mblnEdit = True
    ElseIf mEditType = g修改 Then
        mblnEdit = True
    ElseIf mEditType = g预审 Then
        mblnEdit = True
    ElseIf mEditType = g审核 Then
        mblnEdit = False
    ElseIf mEditType = g取消 Then
        mblnEdit = False
    ElseIf mEditType = g查看 Then
        mblnEdit = False
    End If
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
End Sub

Private Sub LoadPayMoney()
    '--------------------------------------------------------------
    '功能：填充供选择的应付款数据
    '参数：
    '返回：
    '说明：
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    Dim strSQL As String, strWhere As String, strMultiPay As String, strHead As String, strAmount As String
    Dim blnMaterialSys As Boolean
    
    '标志,发票号,入库单号,品名,规格,单位,数量,发票金额
    Call zlCommFun.ShowFlash("正在搜索付款记录,请稍候 ...", Me)
    
    vsPayList.Redraw = False
    DoEvents
    Screen.MousePointer = vbHourglass
    
    '检查安装系统
    strSQL = "select 编号 from zlSystems where 编号=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查安装系统", 400)
    If Not rsTemp.EOF Then
        blnMaterialSys = True
    End If
    rsTemp.Close
    
    '根据操作类型设定记录读取条件
    If mEditType = g新增 Then
        '新增时读取付款序号为空的应付款供选择
        If Format(CDate(mcllFilter("审核日期")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
            strWhere = " and a.付款序号 Is Null and a.审核日期 between [2] and [3] "
        Else
            strWhere = " and a.付款序号 Is Null "
        End If
    ElseIf mEditType = g修改 Then
        '编辑时读取付款序号为空或当前编辑的付款序号所对应的应付款
        If Format(CDate(mcllFilter("审核日期")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
            strWhere = " And (a.付款序号 Is Null Or a.付款序号=[22]) And a.审核日期 between [2] and [3] "
        Else
            strWhere = " And (a.付款序号 Is Null Or a.付款序号=[22]) "
        End If
    Else
        '查看或审核时仅读取当前编辑的付款单所对应的应付款
        strWhere = " And a.付款序号=[22]"
    End If
    
    If mEditType = g新增 Then
        strMultiPay = "(Select a.ID, 0 未付金额, " & _
                      "   Sum(decode(b.审核人, null, a.计划金额, 0)) 本次付款, " & _
                      "   Sum(decode(b.审核人, null, 0, a.计划金额)) 已付金额  " & _
                      " From 应付记录 A, 付款记录 B " & _
                      " Where a.付款序号 = b.付款序号 And instr('12345',a.系统标识)>0 And a.记录性质 = 2 " & _
                      "   and a.记录状态 = 1 and a.付款序号 is not null " & _
                      " Group by a.ID ) A1, "
    Else
        strMultiPay = "select min(a.id) id,a.no,a.序号,a.发票号, " & _
                      "    sum(case when a.记录性质=2 then 0 else a.发票金额 end) 发票金额, a.记录性质,a.付款序号,a.计划金额," & _
                      "    sum(a.数量) 数量 " & _
                      "from (select distinct a.* from 应付记录 a, 应付记录 b " & _
                      "      where b.付款序号=[22] and a.no=b.no and a.序号=b.序号 and a.发票号=b.发票号 and a.系统标识=b.系统标识 and a.项目id=b.项目id " & _
                      "     ) a " & _
                      "group by a.no,a.序号,a.发票号,a.系统标识,a.项目id,a.记录性质,a.付款序号,a.计划金额 "
        
        strMultiPay = "(Select ID, 发票金额 - nvl(已付金额,0) 未付金额, 已付金额, 发票金额, 数量, " & _
                      "   Case When Nvl(已付金额, 0) = 0 and 本次付款 = 0 then 发票金额 else 本次付款 end 本次付款 " & _
                      " From (" & _
                      "   Select a.ID, Max(a.发票金额) 发票金额, sum(数量) 数量, " & _
                      "     Sum(Case When a.计划金额 Is Not Null And a.记录性质 = 2 And Nvl(a.计划金额, 0) <> a.发票金额 And b.Id Is Null Then a.计划金额 Else 0 End) 已付金额," & _
                      "     Sum(Case When a.计划金额 Is Not Null And Nvl(a.计划金额, 0) <> a.发票金额 And b.Id Is Not Null Then a.计划金额 " & _
                      "              When a.计划金额 Is Null And a.付款序号 Is Null    and b.ID is not null   Then a.发票金额 Else 0 End) 本次付款 " & _
                      "   From (" & strMultiPay & ") A, 付款记录 B " & _
                      "   Where a.付款序号 = b.付款序号(+) And b.付款序号(+) = [22] " & _
                      "     And a.Id in (Select ID From 应付记录 Where 付款序号 = [22])  And a.记录性质 <> -1 " & _
                      "   Group By a.ID ) ) A1,"
                      
    End If
    strSQL = "With T1 as (Select " & IIf(mEditType = g预审, "Max(Decode(a.预审, 1, '√', ''))", "Decode(a.付款序号, Null, '', '√')") & " As 标志, " & _
             "        min(a.ID) ID,max(a.记录状态) 记录状态, " & _
             "        '' 计划日期,a.随货单号,a.发票号,to_char(max(A.发票日期),'yyyy-mm-dd') as 发票日期,a.入库单据号," & _
             "        max(a.审核人) as 审核人, to_char(max(a.审核日期),'yyyy-mm-dd') as 审核日期, " & _
             "        max(a.品名) as 品名,max(a.规格) 规格,max(a.计量单位) 计量单位, " & _
             "        max(b.药价级别) as 药价级别," & _
             "        sum(nvl(a.数量,0)) as 数量, " & _
             "        sum(nvl(a.发票金额,0)) as 发票金额, max(a.系统标识) 系统标识, max(a.库房id) 库房id, a.项目id, a.NO " & _
             "  From (" & _
             "        Select Distinct c.* " & vbCr & _
             "        From 应付记录 A, 应付记录 C " & vbCr & _
             "        Where a.系统标识=c.系统标识 And a.No = c.No And a.记录性质 = c.记录性质 And a.项目id=c.项目id " & _
             "            And a.序号=c.序号 And a.计划序号 = c.计划序号 " & vbCr & _
             "            And a.计划日期 Is Null and a.单位ID = [1] " & vbCr & _
             IIf(mEditType = g新增 Or mEditType = g修改, " And not A.记录性质 in (-1, 2) ", " And a.记录性质 <> -1 ") & _
             IIf(mbln付款标志, " And (A.付款标志=1 and nvl(a.系统标识,0)=1 or nvl(a.系统标识,0)<>1 ) ", "") & _
                    strWhere & Replace(CStr(mcllFilter("过滤")), "[alias]", "a.") & _
             ") A, 药品规格 B" & _
             "  Where decode(A.系统标识,1,a.项目id,0)=b.药品ID(+) " & _
             "  group by a.记录性质,a.NO,a.项目id,a.序号," & IIf(mEditType = g预审, "", "a.付款序号,") & "a.发票号,a.发票日期,a.入库单据号,a.随货单号 " & _
             "  having sum(nvl(a.发票金额,0))<>0 " & _
             ") "
    
    If mEditType = g修改 Then
        strHead = "Select Decode(nvl(a1.本次付款,0), 0, a.标志, '√') 标志,"
    Else
        strHead = "Select a.标志, "
    End If
    
    strSQL = strSQL & _
             strHead & _
             "  a.Id, a.记录状态, a.计划日期, a.随货单号, a.发票号, a.发票日期, a.入库单据号, a.审核人, a.审核日期, " & _
             "  a.品名, a.规格, a.药价级别, a.库房id, a.项目id, a.计量单位, " & _
             IIf(mEditType = g新增, "a.数量,", "decode(nvl(a.数量,0), 0, a1.数量, a.数量) 数量, ") & _
             IIf(mEditType = g新增, "a.发票金额, ", "decode(a.发票金额, null, a1.发票金额, a.发票金额) 发票金额, ") & _
             "  Decode(a.系统标识, 1, a.数量 / b.药库包装, 5, a.数量 / e1.换算系数, " & IIf(blnMaterialSys, "2, a.数量 / e2.换算系数,", "") & " null, a.数量, null) 药库数量, " & _
             "  Decode(a.系统标识, 1, b.全院库存, 0) 全院库存, " & _
             "  Decode(a.系统标识, 1, c.当前库房库存, 0) 当前库房库存, " & _
             "  Decode(a.系统标识, 1, b.药库单位, 5, E1.包装单位, " & IIf(blnMaterialSys, "2, E2.包装单位,", "") & " null, a.计量单位, '') 药库单位, " & _
             "  d.名称 当前库房, a.系统标识, a1.已付金额, a1.本次付款, a1.未付金额 " & _
             "From T1 A, " & strMultiPay & _
             "  (Select a.药品id, " & IIf(mint显示单位 = 1, "Round(a.全院库存 / b.药库包装, 5)", "a.全院库存") & " 全院库存, b.药库单位, b.药库包装 " & _
             "   From (Select a.药品id, Sum(a.实际数量) 全院库存 From 药品库存 a Where Exists(Select 1 From T1 Where 项目id = a.药品id) " & _
             "         Group By a.药品id) A, 药品规格 B Where a.药品id = b.药品id) B," & _
             "  (Select a.库房id, a.药品id, " & IIf(mint显示单位 = 1, "Round(a.当前库房库存 / b.药库包装, 5)", "a.当前库房库存") & " 当前库房库存 " & _
             "   From (Select a.库房id, a.药品id, Sum(a.实际数量) 当前库房库存 From 药品库存 A Where Exists(Select 1 From T1 Where 项目id = a.药品id) " & _
             "         Group By a.库房id, a.药品id) A, 药品规格 B Where a.药品id = b.药品id) C, " & _
             "  部门表 D, 材料特性 E1 " & IIf(blnMaterialSys, ", 物资目录 E2 ", "")
    
    '按所有药品库存数量小于发票数量
    If Val(mcllFilter("按所有药品库存数量小于发票数量")) = 1 Then
        strSQL = strSQL & _
                 ",(Select a.项目id, a.数量, b.全院库存 " & _
                 "  From (Select 项目id, Sum(数量) 数量 " & _
                 "        From 应付记录 Where Nvl(系统标识, 0) = 1 And 单位id = [1] And 付款标志 = 1 And 付款序号 Is Null " & _
                 "        Group By 单位id, 项目id) A, (Select 药品id, Sum(实际数量) 全院库存 From 药品库存 Group By 药品id) B " & _
                 "  Where a.项目id = b.药品id(+) And a.数量 > nvl(b.全院库存,0)) F "
    End If

    strSQL = strSQL & _
             "Where a.ID = a1.ID(+) " & IIf(True, "", " and a.付款序号 = a1.付款序号(+) ") & _
             "  and a.项目id=b.药品id(+) and a.库房id=c.库房id(+) and a.项目id=c.药品id(+) and a.库房id=d.id(+) And a.项目id = E1.材料id(+) " & _
             IIf(blnMaterialSys, " And a.项目id = E2.Id(+) ", "") & _
             IIf(Val(mcllFilter("按所有药品库存数量小于发票数量")) = 1, " And a.项目id = f.项目id ", "") & _
             IIf(mEditType = g新增, "  and nvl(a.发票金额,0) - nvl(a1.已付金额,0) - nvl(a1.本次付款,0) <> 0 ", "") & _
             "Order by a.发票号 "
             
    strSQL = "Select * From (" & strSQL & ")"
             
'    If Val(mcllFilter("按所有药品库存数量小于发票数量")) = 1 Then
'        strSQL = "select * from (" & strSQL & ") where 全院库存 < 药库数量 "
'    End If
    
    '供应商ID: [1]
    '审核日期: [2] [3]
    '发票日期: [4] [5]
    '发票号列表: [6]
    '随货单号列表: [7]
    '系统标识: [8]
    '品名: [9]
    '规格: [10]
    '产地: [11]
    '批号: [12] , [13]
    '入库单据号: [14] [15]
    '发票号: [16] [17]
    '随货单号: [18] [19]
    '填制人: [20]
    '审核人: [21]
    '付款序号:[22]
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, _
     CDate(mcllFilter("审核日期")(0)), CDate(mcllFilter("审核日期")(1)), _
     CDate(mcllFilter("发票日期")(0)), CDate(mcllFilter("发票日期")(1)), _
     CStr(mcllFilter("发票号列表")), CStr(mcllFilter("随货单号列表")), _
     CStr(mcllFilter("系统标识")), CStr(mcllFilter("品名")), _
     CStr(mcllFilter("规格")), CStr(mcllFilter("产地")), _
     CStr(mcllFilter("批号")(0)), CStr(mcllFilter("批号")(1)), _
     CStr(mcllFilter("入库单号")(0)), CStr(mcllFilter("入库单号")(1)), _
     CStr(mcllFilter("发票号")(0)), CStr(mcllFilter("发票号")(1)), _
     CStr(mcllFilter("随货单号")(0)), CStr(mcllFilter("随货单号")(1)), _
     CStr(mcllFilter("填制人")), CStr(mcllFilter("审核人")), _
     mlng付款序号, Val(mcllFilter("库房")))
    
    '初始化并填充数据
    With vsPayList
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Clear 1
        mdbl本次应付 = 0
        mdbl累计应付 = 0
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("选定")) = NVL(rsTemp!标志)
            'ID,记录状态,计划日期
            .Cell(flexcpData, i, .ColIndex("选定")) = NVL(rsTemp!ID) & "," & NVL(rsTemp!记录状态) & "," & NVL(rsTemp!计划日期)
            '.TextMatrix(i, .ColIndex("付款标志")) = Nvl(rsTemp!付款标志)
            '.Cell(flexcpData, i, .ColIndex("付款标志")) = Nvl(rsTemp!系统标识)
            .TextMatrix(i, .ColIndex("随货单号")) = NVL(rsTemp!随货单号)
            .TextMatrix(i, .ColIndex("入库单号")) = NVL(rsTemp!入库单据号)
            .TextMatrix(i, .ColIndex("发票号")) = NVL(rsTemp!发票号)
            .Cell(flexcpData, i, .ColIndex("发票号")) = NVL(rsTemp!发票号)
            .TextMatrix(i, .ColIndex("发票日期")) = NVL(rsTemp!发票日期)
            .Cell(flexcpData, i, .ColIndex("发票日期")) = NVL(rsTemp!发票日期)
            .TextMatrix(i, .ColIndex("药价级别")) = NVL(rsTemp!药价级别)
            .TextMatrix(i, .ColIndex("系统标识")) = NVL(rsTemp!系统标识)
            .TextMatrix(i, .ColIndex("药品ID")) = NVL(rsTemp!项目ID)
            .TextMatrix(i, .ColIndex("库房ID")) = NVL(rsTemp!库房ID)
            .TextMatrix(i, .ColIndex("当前库房")) = NVL(rsTemp!当前库房)
            .TextMatrix(i, .ColIndex("当前库房库存")) = Format(Val(NVL(rsTemp!当前库房库存)), gVbFmtString.FM_数量)
            .TextMatrix(i, .ColIndex("全院库存")) = Format(Val(NVL(rsTemp!全院库存)), gVbFmtString.FM_数量)
'            .TextMatrix(i, .ColIndex("药库单位")) = Nvl(rsTemp!药库单位)
'            .TextMatrix(i, .ColIndex("药库数量")) = Format(Val(Nvl(rsTemp!药库数量)), gVbFmtString.FM_数量)
            .TextMatrix(i, .ColIndex("审核人")) = NVL(rsTemp!审核人)
            .TextMatrix(i, .ColIndex("审核日期")) = NVL(rsTemp!审核日期)
            .TextMatrix(i, .ColIndex("品名")) = NVL(rsTemp!品名)
            .TextMatrix(i, .ColIndex("规格")) = NVL(rsTemp!规格)
            If mint显示单位 = 1 Then
                .TextMatrix(i, .ColIndex("单位")) = NVL(rsTemp!药库单位)
                .TextMatrix(i, .ColIndex("数量")) = Format(Val(NVL(rsTemp!药库数量)), gVbFmtString.FM_数量)
            Else
                .TextMatrix(i, .ColIndex("单位")) = NVL(rsTemp!计量单位)
                .TextMatrix(i, .ColIndex("数量")) = Format(Val(NVL(rsTemp!数量)), gVbFmtString.FM_数量)
            End If
            .TextMatrix(i, .ColIndex("发票金额")) = Format(Val(NVL(rsTemp!发票金额)), gVbFmtString.FM_金额)
            .Cell(flexcpData, i, .ColIndex("发票金额")) = Val(NVL(rsTemp!发票金额))
            
            .TextMatrix(i, .ColIndex("已付金额")) = Format(Val(NVL(rsTemp!已付金额)), gVbFmtString.FM_金额)
            .Cell(flexcpData, i, .ColIndex("已付金额")) = Val(NVL(rsTemp!已付金额))
            '.TextMatrix(i, .ColIndex("未付金额")) = Format(.Cell(flexcpData, i, .ColIndex("发票金额")) - .Cell(flexcpData, i, .ColIndex("已付金额")), gVbFmtString.FM_金额)
            '.Cell(flexcpData, i, .ColIndex("未付金额")) = .Cell(flexcpData, i, .ColIndex("发票金额")) - .Cell(flexcpData, i, .ColIndex("已付金额"))
            If mEditType = g新增 Then
                .TextMatrix(i, .ColIndex("未付金额")) = Format(NVL(rsTemp!发票金额) - Val(NVL(rsTemp!已付金额)), gVbFmtString.FM_金额)
                .Cell(flexcpData, i, .ColIndex("未付金额")) = Val(NVL(rsTemp!发票金额) - Val(NVL(rsTemp!已付金额)))
                If IsNull(rsTemp!本次付款) Then
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(.Cell(flexcpData, i, .ColIndex("未付金额")), gVbFmtString.FM_金额)
                Else
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(.Cell(flexcpData, i, .ColIndex("未付金额")) - NVL(rsTemp!本次付款), gVbFmtString.FM_金额)
                End If
                '限制本次付款金额
                .Cell(flexcpData, i, .ColIndex("本次付款")) = .TextMatrix(i, .ColIndex("本次付款"))
            ElseIf mEditType = g修改 Then
                .TextMatrix(i, .ColIndex("未付金额")) = Format(NVL(rsTemp!发票金额, 0) - Val(NVL(rsTemp!已付金额)), gVbFmtString.FM_金额)
                .Cell(flexcpData, i, .ColIndex("未付金额")) = Val(NVL(rsTemp!发票金额, 0) - Val(NVL(rsTemp!已付金额)))
                If IsNull(rsTemp!本次付款) Then
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(.Cell(flexcpData, i, .ColIndex("未付金额")), gVbFmtString.FM_金额)
                Else
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(NVL(rsTemp!本次付款), gVbFmtString.FM_金额)
                End If
                '限制本次付款金额
                If Val(.TextMatrix(i, .ColIndex("本次付款"))) < .Cell(flexcpData, i, .ColIndex("未付金额")) Then
                    .Cell(flexcpData, i, .ColIndex("本次付款")) = .Cell(flexcpData, i, .ColIndex("未付金额"))
                Else
                    .Cell(flexcpData, i, .ColIndex("本次付款")) = .TextMatrix(i, .ColIndex("本次付款"))
                End If
            Else
                .TextMatrix(i, .ColIndex("未付金额")) = Format(Val(NVL(rsTemp!未付金额)), gVbFmtString.FM_金额)
                .Cell(flexcpData, i, .ColIndex("未付金额")) = Val(NVL(rsTemp!未付金额))
                If IsNull(rsTemp!本次付款) And .Cell(flexcpData, i, .ColIndex("发票金额")) = .Cell(flexcpData, i, .ColIndex("未付金额")) Then
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(.Cell(flexcpData, i, .ColIndex("未付金额")), gVbFmtString.FM_金额)
                Else
                    .TextMatrix(i, .ColIndex("本次付款")) = Format(NVL(rsTemp!本次付款), gVbFmtString.FM_金额)
                End If
                '限制本次付款金额
                .Cell(flexcpData, i, .ColIndex("本次付款")) = .Cell(flexcpData, i, .ColIndex("未付金额"))
            End If
            
            mdbl累计应付 = mdbl累计应付 + .Cell(flexcpData, i, .ColIndex("未付金额"))
            If Trim(.TextMatrix(i, .ColIndex("选定"))) <> "" Then
                'mdbl本次应付 = mdbl本次应付 + .Cell(flexcpData, i, .ColIndex("本次付款"))
                mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(i, .ColIndex("本次付款")))
            End If
            
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        Call zl_vsGrid_Para_Restore(mlngModule, vsPayList, Me.Caption, "一般付款列表")
        
'        .ColHidden(.ColIndex("当前库房")) = Not mbln付款标志
'        .ColHidden(.ColIndex("当前库房库存")) = Not mbln付款标志
'        .ColHidden(.ColIndex("全院库存")) = Not mbln付款标志
'        .ColHidden(.ColIndex("药库单位")) = True
'        .ColHidden(.ColIndex("药库数量")) = True
        
        .ColHidden(.ColIndex("已付金额")) = Not mbln付款标志
        .ColHidden(.ColIndex("未付金额")) = Not mbln付款标志
        .ColHidden(.ColIndex("本次付款")) = Not mbln付款标志
        
        .Redraw = flexRDBuffered
    End With
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    
    Call SetMoneyLbl
    Call Get预付数据             '读取预付款
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub
 
Private Sub Get预付数据()
    '--------------------------------------------------------------
    '功能：读取并填充预付款记录供选择
    '参数：
    '返回：
    '说明：
    '--------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    Dim strSQL As String
    Dim strWhere As String
    Dim lngLoop As Long
    
    '标志,结算方式,结算金额,结算号码
    Call zlCommFun.ShowFlash("正在搜索预付款记录,请稍候 ...", Me)
    Screen.MousePointer = vbHourglass
    
    If mEditType = g新增 Then
        strWhere = " And 付款序号 Is Null"
    ElseIf mEditType = g修改 Then
        strWhere = " and (付款序号 Is Null Or 付款序号=[2])"
    Else
        strWhere = " And 付款序号=[2]"
    End If
    On Error GoTo errHandle
    strSQL = "" & _
        "   Select Decode(付款序号,Null,'','√') As 标志,ID,结算方式,金额,结算号码 " & _
        "   From 付款记录 " & _
        "   Where 审核日期 Is not  Null And (记录状态=1 and 预付款=1)  And 单位ID=[1]" & strWhere & _
        "   Order By ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID, mlng付款序号)
    With vs预付
        .Redraw = flexRDNone
        .Clear 1
        .Tag = 0
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        mdbl累计预交 = 0: mdbl本次预交 = 0
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("标志")) = NVL(rsTemp!标志)
            .Cell(flexcpData, i, .ColIndex("标志")) = NVL(rsTemp!ID)
            .TextMatrix(i, .ColIndex("结算方式")) = NVL(rsTemp!结算方式)
            .TextMatrix(i, .ColIndex("结算金额")) = Format(Val(NVL(rsTemp!金额)), gVbFmtString.FM_金额)
            .Cell(flexcpData, i, .ColIndex("结算金额")) = Val(NVL(rsTemp!金额))
            .TextMatrix(i, .ColIndex("结算号码")) = NVL(rsTemp!结算号码)
            
            mdbl累计预交 = mdbl累计预交 + Val(NVL(rsTemp!金额))
            If Trim(.TextMatrix(i, .ColIndex("标志"))) = "√" Then
                mdbl本次预交 = mdbl本次预交 + Val(NVL(rsTemp!金额))
            End If
            If Val(NVL(rsTemp!金额)) < 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H0&
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Call SetMoneyLbl
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        zlCommFun.StopFlash
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Full预付()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充本次预付
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
    
    With mfrmEdit.vs冲预付
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
         lngRow = 1
        For i = 1 To vs预付.Rows - 1
            If Trim(vs预付.TextMatrix(i, 0)) = "√" Then
                .TextMatrix(lngRow, .ColIndex("ID")) = vs预付.Cell(flexcpData, i, vs预付.ColIndex("标志"))
                .TextMatrix(lngRow, .ColIndex("付款方式")) = vs预付.TextMatrix(i, vs预付.ColIndex("结算方式"))
                .TextMatrix(lngRow, .ColIndex("结算号码")) = vs预付.TextMatrix(i, vs预付.ColIndex("结算号码"))
                .TextMatrix(lngRow, .ColIndex("付款金额")) = vs预付.TextMatrix(i, vs预付.ColIndex("结算金额"))
                If Val(.TextMatrix(lngRow, .ColIndex("付款金额"))) < 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H0&
                End If
                lngRow = lngRow + 1
                .Rows = .Rows + 1
            End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lng序号 As Long
    '------------------------------------
    Select Case Control.ID
    Case mConMenu_Report    '调用报表(药品付款查询)
        PrintReport 0
    Case conMenu_File_Print     '打印通知单
        '打印
        If mEditType = g新增 Then Exit Sub
        printbill
    Case conMenu_Edit_Save:       '保存
        Call zlSaveData
    Case conMenu_Edit_SelAll:       '全选
        Call zlAllSelData
    Case conMenu_Edit_ClsAll:       '全清
        Call zlAllClsData
    Case conMenu_View_Backward:       '上一步
        Call zlBackForward
    Case conMenu_View_Forward:       '下一步
        Call zlBackward
    Case conMenu_Manage_Audit:       '审核
        Call zlSaveCheck
    Case conMenu_Edit_ChargeOff  '冲销
        Call zlSaveStrike
    Case mConMenu_Hide_TempSave '保存临时信息
        Call SaveTempData
    Case mConMenu_Hide_TempClearAll '清除临时信息
        Call SaveTempData(True)
    Case conMenu_View_Location
         If Trim(txtFind.Text) = "" Then Exit Sub
         FindRow Trim(txtFind.Text), IIf(vsPayList.Row + 1 >= vsPayList.Rows - 1, 1, vsPayList.Row + 1)
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsThis.RecalcLayout
    Case conMenu_View_FilterView '重置条件
        Call ShowFilterCon
    Case mConMenu_Popu_FP   '按发票统计
        mint统计方式 = 0
        Call Set统计方式
        Call 汇总发票信息
    Case mConMenu_Popu_SH
        mint统计方式 = 1
        Call Set统计方式
        Call 汇总发票信息
    Case conMenu_View_Refresh   '重新刷新数据
        If mEditType = g新增 Then
            Me.Tag = -1
            Call FillDeptDue
        Else
            Call initCard
        End If
        
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit
        Unload Me: Exit Sub
    Case Else
        If Control.ID > 401 And Control.ID < 499 Then
            '暂无报表处理
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub zlAllClsData()
    Dim lngLoop As Long
    Dim objTemp As Object
    
    If mEditType <> g新增 And mEditType <> g修改 And mEditType <> g预审 Then Exit Sub
    If Not (Me.ActiveControl Is vsPayList) And Not (Me.ActiveControl Is vs预付) Then vsPayList.SetFocus
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        objTemp.TextMatrix(lngLoop, objTemp.ColIndex(IIf(objTemp.Name = "vsPayList", "选定", "标志"))) = ""
    Next
    If objTemp Is vsPayList Then
        mdbl本次应付 = 0
    Else
        mdbl本次预交 = 0
    End If
    vsFp.Rows = 2
    vsFp.Clear 1
    Call SetMoneyLbl
End Sub

Private Sub zlAllSelData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:全选数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 19:30:01
    '-----------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim objTemp As Object
    Dim blnHaveData As Boolean
    
    If mEditType <> g新增 And mEditType <> g修改 And mEditType <> g预审 Then Exit Sub
    If Not (Me.ActiveControl Is vsPayList) And Not (Me.ActiveControl Is vs预付) Then vsPayList.SetFocus
    
    Set objTemp = Me.ActiveControl
    For lngLoop = 1 To objTemp.Rows - 1
        blnHaveData = False
        If Me.ActiveControl Is vsPayList Then
            blnHaveData = Trim(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex(IIf(objTemp.Name = "vsPayList", "选定", "标志")))) <> ""
        Else
            blnHaveData = Trim(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("结算方式"))) <> ""
        End If
        If blnHaveData Then
            If objTemp.Name = "vsPayList" Then
                If Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("发票号"))) = "" Then
                    objTemp.TextMatrix(lngLoop, objTemp.ColIndex("选定")) = ""
                Else
'                    If mbln付款标志 Then
'                        '付款、药品类
'                        If Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("付款标志"))) = "付款" And Val(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("付款标志"))) = 1 _
'                            Or Trim(objTemp.TextMatrix(lngLoop, objTemp.ColIndex("付款标志"))) <> "付款" And Val(objTemp.Cell(flexcpData, lngLoop, objTemp.ColIndex("付款标志"))) <> 1 Then
'                            objTemp.TextMatrix(lngLoop, objTemp.ColIndex("选定")) = "√"
'                        Else
'                            objTemp.TextMatrix(lngLoop, objTemp.ColIndex("选定")) = ""
'                        End If
'                    Else
                        objTemp.TextMatrix(lngLoop, objTemp.ColIndex("选定")) = "√"
'                    End If
                End If
            Else
                objTemp.TextMatrix(lngLoop, objTemp.ColIndex("标志")) = "√"
            End If
        End If
    Next
    mblnChange = True
    If objTemp Is vs预付 Then
        mdbl本次预交 = mdbl累计预交
    Else
        mdbl本次应付 = mdbl累计应付
    End If
    Call 汇总发票信息
    Call SetMoneyLbl
End Sub

Private Sub zlBackward()
    '-----------------------------------------------------------------------------------------------------------
    '功能:上一步
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 19:31:50
    '-----------------------------------------------------------------------------------------------------------
    ChangeMode 1
    zlControl.IsCtrlSetFocus vsPayList
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub zlBackForward()
    Dim dblCount As Double
    Dim lngRow As Long
    Dim i As Long, j As Long
    
    If mEditType = g新增 Or mEditType = g修改 Or mEditType = g预审 Then
        If mdbl本次预交 < 0 Then
            MsgBox "本次冲预付款总额不能小于零", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        '功能:玉溪医院要求允许输入负数，因为可能存在退货的情况（如供应商不存供货，要求退货，但要退款）:2008-08-19 15:45:04
        '        '检查各结算方式的预付款总额的累计是否为负数
        '        Dim str结算方式 As String
        '        Dim dbl金额 As Double
        '        str结算方式 = ","
        '        With vs预付
        '            For i = 1 To .Rows - 1
        '                dbl金额 = 0
        '                If InStr(1, str结算方式, "," & .TextMatrix(i, .ColIndex("结算方式")) & ",") = 0 And Trim(.TextMatrix(i, .ColIndex("标志"))) = "√" Then
        '                    For j = 1 To .Rows - 1
        '                        If .TextMatrix(i, .ColIndex("结算方式")) = .TextMatrix(j, "结算方式") And Trim(.TextMatrix(j, .ColIndex("标志"))) = "√" Then
        '                            dbl金额 = dbl金额 + Val(.Cell(flexcpData, j, .ColIndex("结算金额")))
        '                        End If
        '                    Next
        '                    If dbl金额 < 0 Then
        '                        MsgBox "结算方式为:" & .TextMatrix(i, .ColIndex("结算方式")) & "的总额不能负数!", vbInformation + vbDefaultButton1, gstrSysName
        '                        Exit Sub
        '                    End If
        '                    str结算方式 = str结算方式 & .TextMatrix(i, .ColIndex("结算方式")) & ","
        '                End If
        '            Next
        '        End With
    End If
    mfrmEdit.zldbl本次应付 = mdbl本次应付
    mfrmEdit.zldbl本次预交 = mdbl本次预交
    '落列本次冲预付的数据
    If mEditType <> g预审 Then
        Call Full预付
    End If

    ChangeMode 2
    zlControl.IsCtrlSetFocus mfrmEdit.vsPayEdit
End Sub

Private Sub zlSaveData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 19:24:16
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim blnSuccess As Boolean
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
    If blnSuccess = False Then Exit Sub
    
    mblnChange = False
    If blnSuccess = True Then
        If IIf(Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
            '打印
            '问题27930 by lesfeng 2010-03-23
            If mint标记 = 0 Then
                If InStr(mstrPrivs, ";付款通知单;") <> 0 Then
                    printbill
                End If
            Else
                If InStr(mstrPrivs, ";标记付款单;") <> 0 Then
                    printbill
                End If
            End If
        End If
        mblnSuccess = True
        If mEditType = g修改 Then    '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    Call mfrmEdit.ClearData
    vsFp.Clear 1
    vsFp.Rows = 2
    vsTemp.Clear 1
    vsTemp.Rows = 2
    txtDept.Text = "": txtDept.Tag = "-1": Me.Tag = "-1": mlng单位ID = 0:
    ChangeMode 1
    FillDeptDue
    mblnSave = False
    mblnEdit = True
    mblnChange = False
End Sub

Private Sub zlSaveCheck()
    '-----------------------------------------------------------------------------------------------------------
    '功能:审核处理
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 19:33:35
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim blnSuccess As Boolean
    
    If mEditType = g审核 Then        '审核
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '打印
                '问题27930 by lesfeng 2010-03-23
                If mint标记 = 0 Then
                    If InStr(mstrPrivs, ";付款通知单;") <> 0 Then
                        printbill
                    End If
                Else
                    If InStr(mstrPrivs, ";标记付款单;") <> 0 Then
                        printbill
                    End If
                End If
            End If
            mblnChange = False
            mblnSuccess = True
            Unload Me
        End If
        Exit Sub
    End If
End Sub

Private Sub zlSaveStrike()
    '-----------------------------------------------------------------------------------------------------------
    '功能:费用冲销
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-19 19:37:57
    '-----------------------------------------------------------------------------------------------------------

    Dim strReg As String
    Dim blnSuccess As Boolean
    
   If ValidData = False Then Exit Sub
    
   If mEditType = g取消 Then
        If SaveStrike() = True Then
            mblnChange = False
            mblnSuccess = True
            If IIf(Val(zlDatabase.GetPara("审核打印", glngSys, mlngModule)) = 1, 1, 0) = 1 Then
                '打印
                '问题27930 by lesfeng 2010-03-23
                If mint标记 = 0 Then
                    If InStr(mstrPrivs, ";付款通知单;") <> 0 Then
                        printbill
                    End If
                Else
                    If InStr(mstrPrivs, ";标记付款单;") <> 0 Then
                        printbill
                    End If
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case mConMenu_Report
        Control.Enabled = InStr(mstrPrivs, ";药品付款查询;") > 0
    Case conMenu_File_Preview, conMenu_File_Excel
    Case conMenu_Edit_Save:       '保存
        Control.Enabled = mintStep >= 2 And mblnChange
        Control.Visible = (mEditType = g修改 Or mEditType = g新增 Or mEditType = g预审)
    Case conMenu_Edit_SelAll:       '全选
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g修改 Or mEditType = g新增 Or mEditType = g预审)
    Case conMenu_Edit_ClsAll:       '全清
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g修改 Or mEditType = g新增 Or mEditType = g预审)
    Case conMenu_View_FilterView    '重置
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g修改 Or mEditType = g新增)
    Case conMenu_View_Backward:     '上一步
        Control.Enabled = mintStep < 2
    Case conMenu_View_Forward:      '下一步
        Control.Enabled = mintStep >= 2
    Case conMenu_Manage_Audit:      '审核
        Control.Visible = mEditType = g审核
    Case conMenu_Edit_ChargeOff  '冲销
        Control.Visible = mEditType = g取消
    Case conMenu_File_Print     '打印通知单
        Control.Visible = Not (mEditType = g新增) And InStr(mstrPrivs, ";付款通知单;") > 0
    Case mConMenu_Hide_TempSave
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g修改 Or mEditType = g新增)
   Case mConMenu_Hide_TempClearAll   '保存临时信息
        Control.Enabled = mintStep < 2
        Control.Enabled = Control.Enabled And (mEditType = g修改 Or mEditType = g新增)
    Case conMenu_View_Location
        Control.Enabled = mintStep < 2
    Case conMenu_View_LocationItem
        Control.Enabled = mintStep < 2
    Case mConMenu_Popu_FP
        Control.Checked = mint统计方式 <= 0
    Case mConMenu_Popu_SH
        Control.Checked = mint统计方式 >= 1
    Case conMenu_View_Refresh
        Control.Enabled = mEditType = g修改 Or mEditType = g新增
        Control.Visible = Control.Enabled
'    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
'    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
'    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
'    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
'    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdSelDept_Click()
    Dim strTemp As String
    
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp <> "" Then
        txtDept.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
        mlng单位ID = Val(Left(strTemp, InStr(strTemp, ",") - 1))
        FillDeptDue
    End If
    Unload frm供应商选择
    If vsPayList.Enabled Then vsPayList.SetFocus
End Sub

Private Function ShowFilterCon() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:显示付款条件
    '入参:
    '出参:
    '返回:设置了条件,返回true,否则返回false
    '修改人:刘兴宏
    '修改时间:2007/1/25
    '------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, lng供应商ID As Long
    If frm付款条件.ShowFind(Me, lng供应商ID, mstrPrivs, cllFilter) = False Then: Exit Function
    Set mcllFilter = cllFilter
 
    If Format(CDate(mcllFilter("审核日期")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
        lblDate.Caption = "审核日期:" & Format(CDate(mcllFilter("审核日期")(0)), "yyyy-mm-dd") & " 至 " & Format(CDate(mcllFilter("审核日期")(1)), "yyyy-mm-dd")
    Else
        lblDate.Caption = ""
    End If
    
    If Format(CDate(mcllFilter("发票日期")(0)), "yyyy-mm-dd") <> "1901-01-01" Then
        lblDate.Caption = lblDate.Caption & IIf(lblDate.Caption = "", "", Space(5))
        lblDate.Caption = lblDate.Caption & "发票日期:" & Format(CDate(mcllFilter("发票日期")(0)), "yyyy-mm-dd") & " 至 " & Format(CDate(mcllFilter("发票日期")(1)), "yyyy-mm-dd")
    End If
    Me.Tag = ""
    mlng单位ID = Val(mcllFilter("供应商ID"))
    Call FillDeptDue
    ShowFilterCon = True
End Function
 
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPanIndex.pane_付款单
        Item.Handle = mfrmEdit.hwnd
    Case mPanIndex.pane_应付列表
        Item.Handle = picPayList.hwnd
    Case mPanIndex.pane_预交列表
        Item.Handle = pic预付.hwnd
    Case mPanIndex.pane_临时数据
        Item.Handle = picTemp.hwnd
    Case mPanIndex.pane_发票合计
        Item.Handle = picFp.hwnd
    End Select
End Sub

 
Private Sub Form_Activate()
    If mErrBillStatusInfor = 已经删除 Then
        ShowMsgbox "该单据已经被他人删除,不能继续!"
        Unload Me
        Exit Sub
    End If
    If mErrBillStatusInfor = 已经审核 Then
        ShowMsgbox "该单据已经被他人审核,不能再进行审核!"
        Unload Me
        Exit Sub
    End If
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    ChangeMode 1
    
    SetEditPro
    mblnChange = False
    If mEditType = g新增 Then
        If ShowFilterCon = False Then Unload Me: Exit Sub
        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
    ElseIf mEditType = g修改 Then
        If txtDept.Enabled And txtDept.Visible Then txtDept.SetFocus
    ElseIf mEditType = g查看 Then
    End If
    mblnChange = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mbln付款标志 = Val(zlDatabase.GetPara("外购入库需要经过标记付款后才能进行付款管理", glngSys, 0)) = 1
    mint显示单位 = Val(zlDatabase.GetPara("显示单位选择", glngSys, mlngModule))
     
    mblnFirst = True
    Call InitFilter
    Call InitPancel
    Call InitComandBars
    Call Set统计方式
    Call zl_vsGrid_Para_Restore(mlngModule, vsPayList, Me.Caption, "一般付款列表")
    Call zl_vsGrid_Para_Restore(mlngModule, vs预付, Me.Caption, "一般预付列表")
'    If mEditType <> g新增 Then
'        Call initCard
'    End If
    Call initCard
    Call vsPayList_LostFocus
    Call vs预付_LostFocus
    Call vsTemp_LostFocus
    Call vsFp_LostFocus
    mintStep = 0

End Sub

Private Sub Set统计方式()
    stcFpTittle.Caption = IIf(mint统计方式 = 0, "临时信息-发票汇总", "临时信息-随货单汇总")
    vsFp.TextMatrix(0, vsFp.ColIndex("发票号")) = IIf(mint统计方式 = 0, "发票号", "随货单号")
End Sub

Private Sub ChangeMode(intMode As Integer)
    Dim panThis As Pane
    
    If intMode = mintStep Then Exit Sub
    mintStep = intMode
    If mintStep = 1 Then
        dkpMan.CloseAll
        dkpMan.ShowPane (mPanIndex.pane_应付列表)
        '问题27930 by lesfeng 2010-03-23
        If mint标记 = 0 Then
            dkpMan.ShowPane (mPanIndex.pane_预交列表)
        End If
        If mEditType = g新增 Or mEditType = g修改 Then
            dkpMan.ShowPane (mPanIndex.pane_临时数据)
        End If
        dkpMan.ShowPane (mPanIndex.pane_发票合计)
    ElseIf mintStep = 2 Then
        dkpMan.CloseAll
        dkpMan.ShowPane (mPanIndex.pane_付款单)
        Set panThis = dkpMan.FindPane(mPanIndex.pane_付款单)
        
        panThis.MaxTrackSize.Height = Me.ScaleHeight
        panThis.MaxTrackSize.Width = Me.ScaleWidth
        dkpMan_AttachPane panThis
        
    ElseIf mintStep = 3 Then
    End If
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width < 10245 Then Me.Width = 10245
    If Me.Height < 7140 Then Me.Height = 7140
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnYes As Boolean
    Call zl_vsGrid_Para_Save(mlngModule, vsPayList, Me.Caption, "一般付款列表")
    Call zl_vsGrid_Para_Save(mlngModule, vs预付, Me.Caption, "一般预付列表")
    
    If mblnChange Then
        ShowMsgbox "你已经更改了单据信息,你这样退出的话," & vbCrLf & "所更改的数据将不能保存,真的要退出吗?", True, blnYes
        If blnYes = False Then Cancel = 1: Exit Sub
    End If
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '使用个性化设置
        Call zlDatabase.SetPara("定位依据", mstrFindKey, glngSys, mlngModule)
        Call zlDatabase.SetPara("统计方式", mint统计方式, glngSys, mlngModule)
    End If
    If Not mfrmEdit Is Nothing Then Unload mfrmEdit
    Set mfrmEdit = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub mfrmEdit_InitCard(ByVal lng付款序号 As Long, ByVal lng单位ID As Long, ByVal str单位名称 As String)
    '单位名称:
    txtDept.Text = str单位名称
    txtDept.Tag = lng单位ID
    mlng单位ID = lng单位ID: mlng付款序号 = lng付款序号
End Sub

Private Sub mfrmEdit_zlChangeData(ByVal blnChange As Boolean)
    '数据发生改变时,发生此事件
    mblnChange = blnChange
End Sub

Private Sub picCon_Resize()
    Err = 0: On Error Resume Next
    With picCon
          stcTop.Top = .ScaleTop
          stcTop.Width = .ScaleWidth
          stcTop.Left = .ScaleLeft
          stcPayTittle.Move .ScaleLeft, stcTop.Height + stcTop.Top, .ScaleWidth
    End With
End Sub

Private Sub picPayList_Resize()
    Err = 0: On Error Resume Next
    With picPayList
        picCon.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsPayList.Move .ScaleLeft, picCon.Height, .ScaleWidth, .ScaleHeight - picCon.Height
    End With
End Sub

Private Sub picTemp_Resize()
    Err = 0: On Error Resume Next
    With picTemp
        stcTempTittle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsTemp.Move .ScaleLeft, stcTempTittle.Top + stcTempTittle.Height, .ScaleWidth
        vsTemp.Height = .ScaleHeight - vsTemp.Top
    End With
End Sub

Private Sub picFp_Resize()
    Err = 0: On Error Resume Next
    With picFp
        stcFpTittle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vsFp.Move .ScaleLeft, stcFpTittle.Top + stcFpTittle.Height, .ScaleWidth
        vsFp.Height = .ScaleHeight - vsFp.Top
    End With
End Sub

Private Sub pic预付_Resize()
    Err = 0: On Error Resume Next
    With pic预付
        stc预付.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        vs预付.Move .ScaleLeft, stc预付.Top + stc预付.Height, .ScaleWidth, .ScaleHeight - (stc预付.Height + stc预付.Top)
    End With
End Sub

Private Sub 临时汇总数据处理()
    '-----------------------------------------------------------------------------------------------------------
    '功能:重新算汇总发料数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-20 14:14:58
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCur序号 As Long, blnHaveTempSum As Boolean   '存在汇总序号
    Dim i As Long
    Dim lngRow As Long
    Dim dbl金额 As Double
    
    Err = 0: On Error GoTo ErrHand:
    With vsPayList
        lngCur序号 = Val(.TextMatrix(.Row, .ColIndex("汇总序号")))
        If lngCur序号 <= 0 Then Exit Sub
        .TextMatrix(.Row, .ColIndex("汇总序号")) = ""
        dbl金额 = Val(.Cell(flexcpData, .Row, .ColIndex("发票金额")))
        blnHaveTempSum = .FindRow(lngCur序号, , .ColIndex("汇总序号"), , True) > 0
        
    End With
    
    With vsTemp
        lngRow = .FindRow(lngCur序号, 1, .ColIndex("序号"), , True)
        If lngRow > 0 And lngRow <= .Rows - 1 Then
            If blnHaveTempSum Then
                .Cell(flexcpData, lngRow, .ColIndex("金额")) = Val(.Cell(flexcpData, lngRow, .ColIndex("金额"))) - dbl金额
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("金额"))), gVbFmtString.FM_金额)
            Else
                If lngRow = .Rows - 1 And lngRow = 1 Then
                    .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                    .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                Else
                    .RemoveItem lngRow
                End If
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub 汇总发票信息(Optional ByVal strInvoiceNO As String, Optional ByVal strParamInvoiceDate As String)
    '-----------------------------------------------------------------------------------------------------------
    '功能:汇总发票信息或发货单信息
    '入参:
    '  strInvoiceNO: 选定的应付款明细发票号
    '  strParamInvoiceDate: 选定的应付款明细发票日期
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-20 13:20:44
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, strNO As String, intCol As Integer, cllPro As Collection, strKey As String
    Dim bln发票号 As Boolean
    Dim dbl金额 As Double
    Dim strInvoiceDate As String
    Dim varTmp As Variant
    Dim intCountCol As Integer
    Dim intOldRow As Integer
    Dim blnFind As Boolean
    
    intCountCol = vsPayList.ColIndex("本次付款")
    
    bln发票号 = IIf(mint统计方式 = 0, True, False)
    Set cllPro = New Collection
    
    strNO = ""
    With vsPayList
        mdbl本次应付 = 0

        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("选定")) <> "" And .TextMatrix(i, .ColIndex("选定")) = "√" Then
                strInvoiceDate = IIf(Trim(.TextMatrix(i, .ColIndex("发票日期"))) = "", "", .TextMatrix(i, .ColIndex("发票日期")))
                intCol = IIf(bln发票号, .ColIndex("发票号"), .ColIndex("随货单号"))
                strKey = UCase(Trim(.TextMatrix(i, intCol))) & "_" & strInvoiceDate
                
                'mdbl本次应付 = mdbl本次应付 + Val(.Cell(flexcpData, i, intCountCol))
                mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(i, intCountCol))
                
                On Error Resume Next
                varTmp = cllPro(strKey)
                If Err.Number = 0 Then
                    '存在
                    dbl金额 = varTmp(1) + Val(.TextMatrix(i, intCountCol))
                    cllPro.Remove strKey
                Else
                    '不存在
                    dbl金额 = Val(.TextMatrix(i, intCountCol))
                End If
                Err.Clear: On Error GoTo 0
                cllPro.Add Array(Mid(strKey, 1, InStr(strKey, "_") - 1), dbl金额, strInvoiceDate), strKey
                
            End If
        Next
    
    End With
    
    '填充数据
    With vsFp
        intOldRow = .Row
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
        For i = 1 To cllPro.Count
            .TextMatrix(i, .ColIndex("序号")) = i
            .TextMatrix(i, .ColIndex("发票号")) = cllPro(i)(0)
            .TextMatrix(i, .ColIndex("发票日期")) = Format(cllPro(i)(2), "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("金额")) = Format(Val(cllPro(i)(1)), gVbFmtString.FM_金额)
            If cllPro.Count > i Then
                .Rows = .Rows + 1
            End If
            '光栅定位
            If blnFind = False Then
                If UCase(strInvoiceNO) = UCase(.TextMatrix(i, .ColIndex("发票号"))) And strParamInvoiceDate = .TextMatrix(i, .ColIndex("发票日期")) Then
                    .Row = i
                    blnFind = True
                End If
            End If
        Next
        .TopRow = IIf(blnFind, .Row, 1)
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsFp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsFp, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsFp_AfterSort(ByVal Col As Long, Order As Integer)
    Dim intInvCol As Integer
    With vsFp
        
        If .Rows <= 1 Then Exit Sub
        
        intInvCol = .ColIndex("发票号")
        
        Select Case Col
            Case intInvCol
                .ColSort(Col) = Order
                .ColSort(.ColIndex("发票日期")) = 0
                .Select 1, 0, .Rows - 1, .Cols - 1
                .Sort = flexSortUseColSort
                zl_VsGridAfterSort vsFp, Col, Order
            Case .ColIndex("发票日期")
                .ColSort(Col) = Order
                .ColSort(intInvCol) = 0
                .Select 1, 0, .Rows - 1, .Cols - 1
                .Sort = flexSortUseColSort
                zl_VsGridAfterSort vsFp, Col, Order
        End Select
        
    End With
End Sub

Private Sub vsFp_GotFocus()
    zl_VsGridGotFocus vsFp
End Sub

Private Sub vsFp_LostFocus()
    zl_VsGridLOSTFOCUS vsFp
End Sub

Private Sub vsFp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
   
   If Button <> vbRightButton Then Exit Sub
 
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.ID, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub vsPayList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsPayList
        Select Case Col
        Case .ColIndex("发票号"), .ColIndex("发票日期")
        Case .ColIndex("本次付款")
            If Val(.TextMatrix(Row, Col)) > .Cell(flexcpData, Row, .ColIndex("本次付款")) Then
                MsgBox "“本次付款”大于“可未付金额=" & Format(.Cell(flexcpData, Row, .ColIndex("本次付款")), gVbFmtString.FM_金额) & "”？", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = Format(.Cell(flexcpData, Row, .ColIndex("本次付款")), gVbFmtString.FM_金额)
            End If
        Case Else
            Exit Sub
        End Select
        
        If .Cell(flexcpData, Row, Col) <> .TextMatrix(Row, Col) And .ColIndex("本次付款") <> Col Then
            .Cell(flexcpForeColor, Row, Col) = vbRed
        ElseIf .ColIndex("本次付款") = Col Then
            .Cell(flexcpForeColor, Row, Col) = .ForeColor
        Else
            .Cell(flexcpForeColor, Row, Col) = .ForeColor
            Exit Sub
        End If
        
        If .TextMatrix(Row, .ColIndex("选定")) = "" Then
            Call Set应付选择标志
        Else
            If Col = .ColIndex("本次付款") Then Call Set应付选择标志
        End If
    End With
End Sub

Private Sub vsPayList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Set mcbrControl = Me.cbsThis.FindControl(, mConMenu_Report)
    If Not mcbrControl Is Nothing Then
        mcbrControl.Enabled = Val(vsPayList.TextMatrix(NewRow, vsPayList.ColIndex("系统标识"))) = 1 And InStr(mstrPrivs, ";药品付款查询;") > 0
    End If
    Call zl_VsGridRowChange(vsPayList, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vsPayList_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vsPayList, Col, Order)
End Sub

Private Sub vsPayList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mEditType <> g新增 And mEditType <> g修改 Then Cancel = True: Exit Sub
    
    With vsPayList
        Select Case Col
        Case .ColIndex("发票号"), .ColIndex("发票日期")
            
'            If Trim(.TextMatrix(.Row, .ColIndex("发票号"))) = "" And .Row > 1 Then
'                If Trim(.TextMatrix(.Row - 1, .ColIndex("发票号"))) <> "" Then
'                    .TextMatrix(.Row, .ColIndex("发票号")) = .TextMatrix(.Row - 1, .ColIndex("发票号"))
'                    .TextMatrix(.Row, .ColIndex("发票日期")) = .TextMatrix(.Row - 1, .ColIndex("发票日期"))
'                    If .TextMatrix(.Row, .ColIndex("标志")) = "" Then
'                        Call vsPayList_DblClick
'                    End If
'                End If
'            End If
        Case .ColIndex("本次付款")
            Cancel = .TextMatrix(Row, .ColIndex("选定")) = ""         '未选定，不能录入
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPayList_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsPayList_EnterCell()
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    With vsPayList
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("发票号")
            .EditMaxLength = 200
        Case .ColIndex("发票日期")
            .EditMaxLength = 16
        End Select
    End With
End Sub

Private Sub vsPayList_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPayList
        Select Case Col
        Case .ColIndex("发票号")
        Case .ColIndex("发票日期")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsPayList, 0, .Cols - 1, False)
    End With
End Sub

Private Sub vsPayList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsPayList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vsPayList
        Select Case Col
        Case .ColIndex("发票号")
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("发票日期")
            '主要可能存在退款情况
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m文本式)
        Case .ColIndex("本次付款")
            Call VsFlxGridCheckKeyPress(vsPayList, Row, Col, KeyAscii, m金额式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPayList_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If mEditType <> g修改 And mEditType <> g新增 Then Exit Sub
    If Set发票号及发票日期 Then
        '需要自动选中此应付记录
        Call Set应付选择标志
    End If
End Sub

Private Function Set发票号及发票日期() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置发票号及发票日期信息
    '入参:
    '出参:
    '返回:自动设置了的，返回ture,否则返回False
    '编制:刘兴洪
    '日期:2008-08-21 11:19:39
    '说明:主要是当前行的信息根据上行的信息的来获取
    '-----------------------------------------------------------------------------------------------------------
    If mEditType <> g修改 And mEditType <> g新增 Then Exit Function
    If InStr(1, mstrPrivs, ";修改发票信息;") = 0 Then Exit Function
    
    With vsPayList
        If Trim(.TextMatrix(.Row, .ColIndex("发票号"))) = "" And .Row > 1 Then
            If Trim(.TextMatrix(.Row - 1, .ColIndex("发票号"))) <> "" Then '
                .TextMatrix(.Row, .ColIndex("发票号")) = .TextMatrix(.Row - 1, .ColIndex("发票号"))
                .TextMatrix(.Row, .ColIndex("发票日期")) = .TextMatrix(.Row - 1, .ColIndex("发票日期"))
                If .Cell(flexcpData, .Row, .ColIndex("发票号")) <> .TextMatrix(.Row, .ColIndex("发票号")) Then
                    .Cell(flexcpForeColor, .Row, .ColIndex("发票号")) = vbRed
                Else
                    .Cell(flexcpForeColor, .Row, .ColIndex("发票号")) = .ForeColor
                End If
                If .Cell(flexcpData, .Row, .ColIndex("发票日期")) <> .TextMatrix(.Row, .ColIndex("发票日期")) Then
                    .Cell(flexcpForeColor, .Row, .ColIndex("发票日期")) = vbRed
                Else
                    .Cell(flexcpForeColor, .Row, .ColIndex("发票日期")) = .ForeColor
                End If
                Set发票号及发票日期 = True
            End If
        End If
        Select Case .Col
        Case .ColIndex("发票号")
            .EditText = .TextMatrix(.Row, .ColIndex("发票号"))
        Case .ColIndex("发票日期")
            .EditText = .TextMatrix(.Row, .ColIndex("发票日期"))
        End Select
    End With
End Function

Private Sub vsPayList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    With vsPayList
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("发票日期")
            If strKey = "" Then Exit Sub
            strKey = CheckIsDate(strKey, "发票日期")
            If strKey = "" Then Cancel = True: Exit Sub
            .EditText = strKey
        End Select
    End With
End Sub

Private Sub vsTemp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsTemp, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsTemp_GotFocus()
    zl_VsGridGotFocus vsTemp
End Sub

Private Sub vsTemp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, lng序号 As Long
    With vsTemp
        If KeyCode = vbKeyDelete Then
            
            If (mEditType = g新增 Or mEditType = g修改) And Val(.TextMatrix(.Row, .ColIndex("序号"))) > 0 Then
                lng序号 = Val(.TextMatrix(.Row, .ColIndex("序号")))
                If MsgBox("你真的要删除汇总序号为" & lng序号 & "  的临时汇总数据吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    '先清除付款汇总序号
                    With vsPayList
                        For i = 1 To .Rows - 1
                            If Val(.TextMatrix(.Row, .ColIndex("汇总序号"))) = lng序号 Then
                                .TextMatrix(.Row, .ColIndex("汇总序号")) = ""
                            End If
                        Next
                    End With
                    '移除当前行数据
                    If .Rows - 1 = .Row And .Row = 1 Then
                        .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    Else
                        .RemoveItem .Row
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsTemp_LostFocus()
    zl_VsGridLOSTFOCUS vsTemp

End Sub

Private Sub vs预付_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vs预付, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs预付_AfterSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridAfterSort(vs预付, Col, Order)
End Sub

Private Sub vs预付_DblClick()
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    With vs预付
        If .Col <> .ColIndex("标志") Then Exit Sub
        If Trim(.TextMatrix(.Row, .ColIndex("结算方式"))) <> "" Then
            .TextMatrix(.Row, .ColIndex("标志")) = IIf(Trim(.TextMatrix(.Row, .ColIndex("标志"))) = "", "√", "")
            If Trim(.TextMatrix(.Row, 0)) = "" Then
                mdbl本次预交 = mdbl本次预交 - Val(.TextMatrix(.Row, .ColIndex("结算金额")))
            Else
                mdbl本次预交 = mdbl本次预交 + Val(.TextMatrix(.Row, .ColIndex("结算金额")))
            End If
        End If
    End With
    Call SetMoneyLbl
End Sub

Private Sub SetMoneyLbl()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置标签金额
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    mfrmEdit.zldbl本次应付 = mdbl本次应付
    mfrmEdit.zldbl本次预交 = mdbl本次预交
    lbl金额(1).Caption = "累计应付:" & Format(mdbl累计应付, "###0.00;-###0.00;0.00;0.00") & ""
    '问题27930 by lesfeng 2010-03-23
    If mint标记 = 0 Then
        lbl金额(2).Caption = "付款金额:" & Format(mdbl本次应付, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(3).Caption = "预交累计:" & Format(mdbl累计预交, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(4).Caption = "冲预付:" & Format(mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(5).Caption = "本次应付:" & Format(mdbl本次应付 - mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
    Else
        lbl金额(2).Caption = "标记付款金额:" & Format(mdbl本次应付, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(3).Caption = "预交累计:" & Format(mdbl累计预交, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(4).Caption = "冲预付:" & Format(mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
        lbl金额(5).Caption = "本次标记应付:" & Format(mdbl本次应付 - mdbl本次预交, "###0.00;-###0.00;0.00;0.00") & ""
    End If
End Sub

Private Sub vs预付_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        vs预付_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vs预付_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, blnHaveData As Boolean
    
    If Button <> 2 Then Exit Sub
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    
    blnHaveData = False
    With vs预付
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, vs预付.ColIndex("标志")) <> "" Then blnHaveData = True: Exit For
        Next
    End With
    If blnHaveData = False Then Exit Sub
    If vs预付.Enabled Then vs预付.SetFocus
End Sub

Private Sub vsPayList_Click()
    With vsPayList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
           ' SetColumnSort vsPayList, mintPreCol, mintsort
            Exit Sub
         End If
    End With
End Sub

Private Sub vsPayList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
   
    With vsPayList
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeySpace
                Call vsPayList_DblClick
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
 
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    With vsPayList
        Select Case .Col
        Case .Cols - 1
            .Col = 0: .LeftCol = 0
            Exit Sub
        End Select
        
        Call zlVsMoveGridCell(vsPayList, 0, vsPayList.Cols - 1, False)
    End With
End Sub

Private Sub vsPayList_DblClick()
    Dim intCol As Integer
    With vsPayList
        
        If mEditType <> g新增 And mEditType <> g修改 And mEditType <> g预审 Then
            If Val(.TextMatrix(.Row, .ColIndex("系统标识"))) = 1 And InStr(mstrPrivs, ";药品付款查询;") > 0 Then
                '调用报表
                PrintReport 1
            End If
            Exit Sub
        End If
        
        If Trim(.TextMatrix(.Row, .ColIndex("发票号"))) = "" Then Exit Sub
        
        If .Col <> .ColIndex("选定") And .Col <> .ColIndex("本次付款") Then
            If Val(.TextMatrix(.Row, .ColIndex("系统标识"))) = 1 And InStr(mstrPrivs, ";药品付款查询;") > 0 Then
                '调用报表
                PrintReport 1
            End If
            Exit Sub
        End If
        
        Call Set应付选择标志
        If .TextMatrix(.Row, .ColIndex("选定")) = "√" Then
            Call Set发票号及发票日期
        End If
        mblnChange = True
    End With
End Sub

Private Sub Set应付选择标志()
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置应付选择标志
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-08-21 11:25:11
    '-----------------------------------------------------------------------------------------------------------
    Dim intCol As Integer
    Dim strInvoice As String
    Dim strInvoiceDate As String
    
    If mEditType <> g新增 And mEditType <> g修改 And mEditType <> g预审 Then Exit Sub

    With vsPayList
        If Trim(.Cell(flexcpData, .Row, .ColIndex("选定"))) = "" Then Exit Sub
'        If mbln付款标志 Then
'            If .TextMatrix(.Row, .ColIndex("付款标志")) = "付款" And Val(.Cell(flexcpData, .Row, .ColIndex("付款标志"))) = 1 _
'                Or .TextMatrix(.Row, .ColIndex("付款标志")) <> "付款" And Val(.Cell(flexcpData, .Row, .ColIndex("付款标志"))) <> 1 Then
'                .TextMatrix(.Row, .ColIndex("选定")) = IIf(.TextMatrix(.Row, .ColIndex("选定")) = "", "√", "")
'            Else
'                .TextMatrix(.Row, .ColIndex("选定")) = ""
'                Exit Sub
'            End If
'        Else
            '.TextMatrix(.Row, .ColIndex("选定")) = IIf(.TextMatrix(.Row, .ColIndex("选定")) = "", "√", "")
'        End If
        If vsPayList.Col <> vsPayList.ColIndex("本次付款") Then
            .TextMatrix(.Row, .ColIndex("选定")) = IIf(.TextMatrix(.Row, .ColIndex("选定")) = "", "√", "")
            strInvoice = IIf(mint统计方式 = 0, .TextMatrix(.Row, .ColIndex("发票号")), .TextMatrix(.Row, .ColIndex("随货单号")))
            strInvoiceDate = .TextMatrix(.Row, .ColIndex("发票日期"))
        End If

'        If Trim(.TextMatrix(.Row, .ColIndex("选定"))) = "" Then
'            mdbl本次应付 = mdbl本次应付 - Val(.TextMatrix(.Row, .ColIndex("本次付款")))
'        Else
'            mdbl本次应付 = mdbl本次应付 + Val(.TextMatrix(.Row, .ColIndex("本次付款")))
'        End If
        mblnChange = True
    End With
    Call 临时汇总数据处理
    Call 汇总发票信息(strInvoice, strInvoiceDate)
    Call SetMoneyLbl
End Sub

Private Sub vsPayList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, blnHaveData As Boolean
    
    If Button <> 2 Then Exit Sub
    If mEditType <> g新增 And mEditType <> g修改 Then Exit Sub
    blnHaveData = False
    With vsPayList
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("选定")) <> "" Then blnHaveData = True: Exit Sub
        Next
    End With
    
    If blnHaveData = False Then Exit Sub
    If vsPayList.Enabled Then vsPayList.SetFocus
 End Sub

Private Sub txtDept_Change()
    mlng单位ID = 0
End Sub

Private Sub txtDept_GotFocus()
    SetTxtGotFocus txtDept, True
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtDept.Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Val(txtDept.Tag) <> 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelMltProvide = False Then
        Exit Sub
    End If
End Sub

Private Sub FillDeptDue()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载部门数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select ID,编码,名称,地址,电话,开户银行,税务登记号,信用期 From 供应商 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng单位ID)
    Err.Clear: On Error GoTo 0
    If Not rsTemp.EOF Then
        txtDept.Text = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        mlng单位ID = Val(NVL(rsTemp!ID))
        mlng效期 = NVL(rsTemp!信用期, 0)
    End If
    Call mfrmEdit.zlLoadPrivder(mlng单位ID)
    zlControl.IsCtrlSetFocus vsPayList
    If mlng单位ID <> Val(Me.Tag) Then
        Me.Tag = mlng单位ID
        LoadPayMoney
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtDept_LostFocus()
    ImeLanguage False
End Sub

Private Function ValidData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim strTemp As String
    Dim dblCount As Double
    If mlng单位ID = 0 Then
         ShowMsgbox "供应商选择有误,请重新选择!"
         Call zlBackForward
         If txtDept.Enabled Then txtDept.SetFocus
         Exit Function
    End If
    If mfrmEdit.zlValidData = False Then
        Exit Function
    End If
    If mEditType = g新增 Or mEditType = g修改 Then
        If InStr(1, mstrPrivs, "修改发票信息") > 0 Then
            With vsPayList
                For lngRow = 1 To .Rows - 1
                    If Trim(.Cell(flexcpData, lngRow, .ColIndex("选定"))) <> "" _
                        And Trim(.TextMatrix(lngRow, .ColIndex("选定"))) <> "" Then
                        If Trim(.TextMatrix(.Row, .ColIndex("发票号"))) = "" Then
                            ShowMsgbox "发票号未输入，请检查!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("发票号"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                        strTemp = .TextMatrix(.Row, .ColIndex("发票日期"))
                        If strTemp = "" Then
                            ShowMsgbox "发票日期未输入，请检查!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("发票日期"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                        If IsDate(strTemp) = False Or IsNumeric(strTemp) Then
                            ShowMsgbox "输入的发票日期不是日期类型，请检查!"
                            If mintStep >= 2 Then
                                Call zlBackForward
                            End If
                            .Col = .ColIndex("发票日期"): .Row = lngRow: .TopRow = lngRow
                            zlControl.IsCtrlSetFocus vsPayList
                            Exit Function
                        End If
                    End If
                Next
            End With
        End If
    End If
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-19 15:20:25
    '-----------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection
    Dim strNO_IN As String
    Dim lng付款序号_IN As Long
    Dim str发票号 As String, str发票日期 As String
    Dim lngRow As Long
    Dim varData As Variant

    SaveCard = False
    Set cllPro = New Collection
    
    Err = 0: On Error GoTo errHandle:
    
    If mEditType = g预审 Then
        '预审
        If vsPayList.Rows <= 1 Then Exit Function
        'ID, 记录状态, 计划日期
        varData = Split(vsPayList.Cell(flexcpData, 1, vsPayList.ColIndex("选定")), ",")
        '清理预审标志
        gstrSQL = "Zl_付款管理_CheckClear(" & _
                  varData(0) & "," & _
                  mlng付款序号 & _
                  ")"
        AddArray cllPro, gstrSQL
        With vsPayList
            For lngRow = 1 To .Rows - 1
                If Trim(.TextMatrix(lngRow, .ColIndex("选定"))) <> "" Then
                    varData = Split(.Cell(flexcpData, lngRow, .ColIndex("选定")), ",")
                    gstrSQL = "Zl_付款管理_Check(" & varData(0)     '作预审标志
                    gstrSQL = gstrSQL & "," & mlng付款序号          '付款序号
                    gstrSQL = gstrSQL & ",'" & UserInfo.姓名 & "'"  '预审人
                    gstrSQL = gstrSQL & ",to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd hh24:mi:ss' ) ) "       '预审日期
                    
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
    Else
        '付款单的保存
        If mfrmEdit.zlSaveCard(cllPro, lng付款序号_IN, strNO_IN) = False Then Exit Function
        
        '对应采购清单
        With vsPayList
            For lngRow = 1 To .Rows - 1
                If Trim(.TextMatrix(lngRow, .ColIndex("选定"))) <> "" Then
                    If InStr(1, mstrPrivs, "修改发票信息") > 0 Then
                        str发票号 = Trim(.TextMatrix(lngRow, .ColIndex("发票号")))
                        If Trim(.Cell(flexcpData, lngRow, .ColIndex("发票号"))) = str发票号 Then
                            '如果没发生改变，就不更改发票号了
                            str发票号 = "NULL"
                        Else
                            str发票号 = "'" & str发票号 & "'"
                        End If
                        str发票日期 = Trim(.TextMatrix(lngRow, .ColIndex("发票日期")))
                        If Trim(.Cell(flexcpData, lngRow, .ColIndex("发票日期"))) = str发票日期 Or str发票日期 = "" Then
                            '如果没发生改变，就不更改发票号了
                            str发票日期 = "NULL"
                        Else
                            str发票日期 = "To_date('" & str发票日期 & "','yyyy-mm-dd')"
                        End If
                    Else
                        str发票号 = "NULL": str发票日期 = "NULL"
                    End If
                    
                    ' .Cell(flexcpData, .Row, .ColIndex("标志")) : 'ID,记录状态 ,计划日期
                    varData = Split(.Cell(flexcpData, lngRow, .ColIndex("选定")), ",")
                    '过程参数
                    'Zl_付款序号_Update(
                    gstrSQL = "zl_付款序号_UPDATE("
                    'Id_In       In Varchar2 := Null,
                    gstrSQL = gstrSQL & Val(varData(0)) & ","
                    '计划序号_In In Varchar2 := Null, --以0,1,2,3方式传入
                    gstrSQL = gstrSQL & "NULL,"
                    '付款序号_In In 付款记录.付款序号%Type := Null,
                    gstrSQL = gstrSQL & "" & lng付款序号_IN & ","
                    '预付款_In   In 付款记录.预付款%Type := 0,
                    gstrSQL = gstrSQL & "" & 0 & ","
                    '发票金额_In     In 应付记录.发票金额%Type := 0
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("发票金额"))) & ","
                    '本次付款_In
                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("本次付款"))) & ","
                    '--应付记录:发票号和发票日期为NULL的情况下，将不更改应付记录中的发票号，同时，只能更改是普通付款才处理发票号
                    '发票号_In   In 应付记录.发票号%Type := Null,
                    gstrSQL = gstrSQL & "" & str发票号 & ","
                    '发票日期_In In 应付记录.发票日期%Type := Null
                    gstrSQL = gstrSQL & "" & str发票日期 & ")"
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
        
        '保存预付款
        With vs预付
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, .ColIndex("标志")) <> "" Then
                    'Zl_付款序号_Update
                    gstrSQL = "zl_付款序号_UPDATE("
                    '  Id_In       In Varchar2 := Null,
                    gstrSQL = gstrSQL & Val(.Cell(flexcpData, lngRow, .ColIndex("标志"))) & ","
                    '  计划序号_In In Varchar2 := Null, --以0,1,2,3方式传入
                    gstrSQL = gstrSQL & "NULL,"
                    '  付款序号_In In 付款记录.付款序号%Type := Null,
                    gstrSQL = gstrSQL & "" & lng付款序号_IN & ","
                    '  预付款_In   In 付款记录.预付款%Type := 0,
                    gstrSQL = gstrSQL & "" & 1 & ","
                    '  金额_In     In 应付记录.发票金额%Type := 0
                    gstrSQL = gstrSQL & "" & Val(.Cell(flexcpData, lngRow, .ColIndex("结算金额"))) & ")"
                    '--应付记录:发票号和发票日期为NULL的情况下，将不更改应付记录中的发票号，同时，只能更改是普通付款才处理发票号
                    '发票号_In   In 应付记录.发票号%Type := Null,
                    '发票日期_In In 应付记录.发票日期%Type := Null
                    AddArray cllPro, gstrSQL
                End If
            Next
        End With
    End If
    
    Err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    If mEditType <> g预审 Then
        If Check付款与应付明细(lng付款序号_IN) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    '提交事务
    gcnOracle.CommitTrans
    SaveCard = True
    Me.stbThis.Panels(2).Text = "上张单据号为:" & strNO_IN
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetEditPro()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:设置编辑属性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    cmdSelDept.Enabled = mEditType = g新增
    txtDept.Enabled = mEditType = g新增
End Sub

Private Function SaveCheck() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:审核单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String, cllPro As New Collection
    
    SaveCheck = False
    
    strNO_IN = mfrmEdit.txtNo
    If mfrmEdit.zlCheck(cllPro) = False Then
        ChangeMode 2
        zlControl.IsCtrlSetFocus mfrmEdit.vsPayEdit
        Exit Function
    End If
    '   zl_付款管理_VERIFY(NO_IN);
    gstrSQL = "zl_付款管理_VERIFY('" & _
        strNO_IN & "')"
    AddArray cllPro, gstrSQL
    
    Err = 0: On Error GoTo errHandle:
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveCheck = True
    Exit Function
    
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
 '-----------------------------------------------------------------------------------------------------------
    '--功  能:冲销单据
    '--入参数:
    '--出参数:
    '--返  回:成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strNO_IN As String
    
    SaveStrike = False
    
    strNO_IN = mfrmEdit.txtNo
    On Error GoTo errHandle:
    '   zl_付款管理_VERIFY(NO_IN);
    gstrSQL = "zl_付款管理_strike('" & _
        strNO_IN & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveStrike = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'打印单据
Private Sub printbill()
    '问题27930 by lesfeng 2010-03-23
    If mint标记 = 0 Then
        ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_1", Me, "单据编号=" & mfrmEdit.txtNo.Tag, "记录状态=" & mint记录状态
    Else
        ReportOpen gcnOracle, glngSys, "ZL1_BILL_1323_3", Me, "单据编号=" & mfrmEdit.txtNo.Tag, "记录状态=" & mint记录状态
    End If
End Sub

Private Function SelMltProvide() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取供应商数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String, vRect As RECT, lngH As Long, blnCancel As Boolean
    Dim str权限 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    
    If Trim(txtDept.Text) = "" Then Exit Function
    
    strKey = GetMatchingSting(UCase(txtDept.Text), False)
    SelMltProvide = False
    
    str权限 = " and " & Get分类权限(mstrPrivs)
    
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    
    strSQL = "" & _
        "  Select   ID,编码,名称,简码,许可证号," & _
        "           to_char(许可证效期,'yyyy-mm-dd') as 许可证效期,执照号," & _
        "           to_char(执照效期,'yyyy-mm-dd') as 执照效期,税务登记号,联系人 " & _
        "  From  供应商 " & _
        "   Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & zl_获取站点限制() & "   " & _
        "           and 末级=1 And ( 编码 Like upper([1]) or 名称 like [1] or 简码  like  upper([1]) ) " & str权限
    
    
    vRect = zlControl.GetControlRect(txtDept.hwnd)
    lngH = txtDept.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "供应商选择", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "没有找到满足条件的供应商,请检查!"
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp.State = 0 Then Exit Function
    If txtDept.Enabled Then txtDept.SetFocus
    
    txtDept.Text = NVL(rsTemp!名称)
    mlng单位ID = NVL(rsTemp!ID, 0)
    txtDept.Tag = mlng单位ID
    '填充数据
    FillDeptDue
    zlCommFun.PressKey vbKeyTab
    SelMltProvide = True
End Function
 
Private Sub vsPayList_GotFocus()
    Call zl_VsGridGotFocus(vsPayList)
End Sub

Private Sub vsPayList_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsPayList)
End Sub

Private Sub vs预付_GotFocus()
    Call zl_VsGridGotFocus(vs预付)
End Sub

Private Sub vs预付_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs预付)
End Sub

Private Sub FindRow(ByVal strFind As String, Optional lngRow As Long = 1)
    '功能:查找指列的数据是否满足相关的条件
    '参数:intMachType:0-左匹配,1-完全匹配
    Dim i As Long, lngCol As Long
    Dim blnAll As Boolean
    With vsPayList
        lngCol = .ColIndex(mstrFindKey)
        If lngCol < 0 Then Exit Sub
        Select Case lngCol
        Case .ColIndex("汇总序号")
            blnAll = True
        Case .ColIndex("数量")
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_数量)
        Case .ColIndex("发票金额")
            blnAll = True
            strFind = Format(Val(strFind), gVbFmtString.FM_金额)
        Case Else
            blnAll = False
        End Select
       i = .FindRow(strFind, lngRow, lngCol, False, blnAll)
       If i > 0 Then
            .Row = i: .TopRow = i
       Else
            ShowMsgbox "已经查到末尾,没有发现满足条件的数据,请检查!"
       End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    If mstrFindKey = "姓名" Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNO As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtFind) = "" Then Exit Sub
    FindRow Trim(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtFind, KeyAscii, m文本式
End Sub

Private Sub PrintReport(ByVal bytStyle As Byte) '(ByVal lngDeptID As Long, ByVal lngPurveryID As Long, ByVal lngDrugID As Long)
    Dim lngDeptID As Long, lngDrugID As Long
    Dim strDept As String, strDrug As String, strSupplier As String
    
    strSupplier = Mid(txtDept.Text, InStr(txtDept.Text, "-") + 1)
    With vsPayList
        lngDeptID = Val(.TextMatrix(.Row, .ColIndex("库房ID")))
        lngDrugID = Val(.TextMatrix(.Row, .ColIndex("药品ID")))
        strDept = .TextMatrix(.Row, .ColIndex("当前库房"))
        strDrug = .TextMatrix(.Row, .ColIndex("品名"))
    End With
    
    If bytStyle = 1 Then
        ReportOpen gcnOracle, glngSys, "ZL1_REPORT_1323", Me, _
            "药品名称=" & strDrug & "|" & lngDrugID, _
            "库房=" & strDept & "|" & IIf(lngDeptID = 0, " is not null ", "=" & lngDeptID), _
            "供药单位=" & strSupplier & "|" & mlng单位ID
    Else
        ReportOpen gcnOracle, glngSys, "ZL1_REPORT_1323", Me, _
            "供药单位=" & strSupplier & "|" & mlng单位ID
    End If
End Sub
