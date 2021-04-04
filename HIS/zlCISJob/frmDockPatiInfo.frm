VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.0#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmDockPatiInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "病人信息"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   129
      Top             =   11160
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   635
      SimpleText      =   $"frmDockPatiInfo.frx":0000
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDockPatiInfo.frx":0047
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
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
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8235
      Index           =   2
      Left            =   600
      ScaleHeight     =   8235
      ScaleWidth      =   9915
      TabIndex        =   64
      Top             =   4920
      Width           =   9915
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   24
         Left            =   4065
         MaxLength       =   100
         TabIndex        =   84
         Top             =   1170
         Width           =   3270
      End
      Begin VB.CommandButton cmdE 
         Height          =   220
         Index           =   23
         Left            =   2370
         Picture         =   "frmDockPatiInfo.frx":08DB
         Style           =   1  'Graphical
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1035
         Width           =   240
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   23
         Left            =   810
         MaxLength       =   20
         TabIndex        =   83
         Text            =   "2013-06-20 18:00"
         Top             =   1065
         Width           =   1500
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   675
         TabIndex        =   135
         Top             =   825
         Width           =   750
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   -30
            Width           =   700
         End
      End
      Begin VB.TextBox txtSL 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Left            =   105
         MaxLength       =   3
         TabIndex        =   81
         Top             =   825
         Width           =   300
      End
      Begin VB.CheckBox chkNoAller 
         Appearance      =   0  'Flat
         Caption         =   "无过敏记录"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1515
         TabIndex        =   134
         Top             =   1290
         Width           =   1215
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   2
         Left            =   9360
         Picture         =   "frmDockPatiInfo.frx":09D1
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         ToolTipText     =   "诊断上移"
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   3
         Left            =   9360
         Picture         =   "frmDockPatiInfo.frx":2F79
         Style           =   1  'Graphical
         TabIndex        =   132
         TabStop         =   0   'False
         ToolTipText     =   "诊断下移"
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   0
         Left            =   9480
         Picture         =   "frmDockPatiInfo.frx":56A4
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "诊断上移"
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdDiagMove 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Height          =   375
         Index           =   1
         Left            =   9480
         Picture         =   "frmDockPatiInfo.frx":7C4C
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "诊断下移"
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdSaveZY 
         Height          =   300
         Left            =   8595
         Picture         =   "frmDockPatiInfo.frx":A377
         Style           =   1  'Graphical
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "将当前摘要设置为常用摘要。"
         Top             =   1110
         Width           =   300
      End
      Begin VB.CommandButton cmdShowZY 
         Caption         =   "…"
         Height          =   300
         Left            =   8280
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1110
         Width           =   300
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2280
         Index           =   6
         Left            =   1260
         ScaleHeight     =   2280
         ScaleWidth      =   8145
         TabIndex        =   105
         Top             =   5775
         Width           =   8145
         Begin VB.CommandButton cmdE 
            Caption         =   "…"
            Height          =   220
            Index           =   25
            Left            =   2745
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
            Top             =   1695
            Width           =   240
         End
         Begin VB.TextBox txtE 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   26
            Left            =   4725
            MaxLength       =   100
            TabIndex        =   94
            Top             =   1710
            Width           =   3270
         End
         Begin VB.TextBox txtE 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   25
            Left            =   810
            MaxLength       =   100
            TabIndex        =   93
            Top             =   1695
            Width           =   1965
         End
         Begin VB.Frame fraC 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   22
            Left            =   4605
            TabIndex        =   108
            Top             =   30
            Width           =   3390
            Begin VB.ComboBox cboE 
               BackColor       =   &H00FDFDFD&
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   91
               Top             =   -30
               Width           =   3345
            End
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            Caption         =   "传染病上传"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1995
            TabIndex        =   90
            Top             =   30
            Width           =   1260
         End
         Begin VB.PictureBox picPanel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   5
            Left            =   0
            ScaleHeight     =   360
            ScaleWidth      =   1845
            TabIndex        =   107
            Top             =   45
            Width           =   1845
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "复诊"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   945
               TabIndex        =   106
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton optInfo 
               Appearance      =   0  'Flat
               Caption         =   "初诊"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   0
               TabIndex        =   89
               Top             =   0
               Width           =   690
            End
         End
         Begin zl9CISJob.UCPatiVitalSigns UCPatiVitalSigns 
            Height          =   660
            Left            =   150
            TabIndex        =   95
            Top             =   450
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   1164
            TextBackColor   =   -2147483643
            LblBackColor    =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Style           =   1
            XDis            =   200
            YDis            =   200
            LabToTxt        =   0
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "生命体征"
            Height          =   180
            Index           =   28
            Left            =   0
            TabIndex        =   127
            Top             =   2025
            Width           =   720
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "去向"
            Height          =   180
            Index           =   22
            Left            =   3780
            TabIndex        =   111
            Top             =   30
            Width           =   360
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "其他医学警示"
            Height          =   180
            Index           =   26
            Left            =   3060
            TabIndex        =   110
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            Caption         =   "医学警示"
            Height          =   180
            Index           =   25
            Left            =   0
            TabIndex        =   109
            Top             =   1695
            Width           =   720
         End
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   3825
         ScaleHeight     =   345
         ScaleWidth      =   3510
         TabIndex        =   96
         Top             =   1260
         Width           =   3510
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "药品目录"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "过敏源"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   1425
            TabIndex        =   98
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.PictureBox picPanel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   3825
         ScaleHeight     =   345
         ScaleWidth      =   3330
         TabIndex        =   102
         Top             =   2925
         Width           =   3330
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "诊断编码"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   1395
            TabIndex        =   104
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optInfo 
            Appearance      =   0  'Flat
            Caption         =   "疾病编码"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.TextBox txtE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDFDFD&
         Height          =   630
         Index           =   27
         Left            =   1515
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   450
         Width           =   7335
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   1260
         Left            =   1230
         TabIndex        =   85
         Top             =   1665
         Width           =   7980
         _cx             =   14076
         _cy             =   2222
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockPatiInfo.frx":A901
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   1260
         Left            =   1305
         TabIndex        =   87
         Top             =   3270
         Width           =   7980
         _cx             =   14076
         _cy             =   2222
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockPatiInfo.frx":A9B4
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   960
         Left            =   1215
         TabIndex        =   88
         Top             =   4665
         Width           =   7980
         _cx             =   14076
         _cy             =   1693
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   24
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockPatiInfo.frx":AC70
         ScrollTrack     =   -1  'True
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
      Begin VB.Line linD 
         Index           =   3
         X1              =   1290
         X2              =   1680
         Y1              =   240
         Y2              =   270
      End
      Begin VB.Line linD 
         Index           =   2
         X1              =   180
         X2              =   390
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Line linD 
         Index           =   1
         X1              =   255
         X2              =   570
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line linD 
         Index           =   0
         X1              =   150
         X2              =   930
         Y1              =   255
         Y2              =   270
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "发病地址"
         Height          =   180
         Index           =   24
         Left            =   2760
         TabIndex        =   138
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "发病时间"
         Height          =   180
         Index           =   23
         Left            =   0
         TabIndex        =   137
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "就诊信息"
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
         Index           =   52
         Left            =   4425
         TabIndex        =   120
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lblLinkAdd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "将诊断添加至就诊摘要"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   1515
         MouseIcon       =   "frmDockPatiInfo.frx":AE62
         MousePointer    =   99  'Custom
         TabIndex        =   116
         ToolTipText     =   "将诊断描述添加至就诊摘要中。"
         Top             =   2985
         Width           =   1800
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "诊断记录"
         Height          =   180
         Index           =   54
         Left            =   360
         TabIndex        =   101
         Top             =   2985
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "过敏记录"
         Height          =   180
         Index           =   53
         Left            =   495
         TabIndex        =   100
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "就诊摘要"
         Height          =   180
         Index           =   27
         Left            =   495
         TabIndex        =   99
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.VScrollBar vsc 
      Height          =   7935
      Left            =   10905
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   10845
      TabIndex        =   114
      Top             =   5970
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   172883969
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   0
      Left            =   900
      ScaleHeight     =   4530
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   330
      Width           =   9705
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   55
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   26
         Top             =   3960
         Width           =   2475
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   2
         Left            =   1245
         TabIndex        =   16
         Top             =   2375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Style           =   1
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   1
         Left            =   1245
         TabIndex        =   11
         Top             =   1640
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Style           =   1
      End
      Begin ZlPatiAddress.PatiAddress PatiAddress 
         Height          =   270
         Index           =   0
         Left            =   1245
         TabIndex        =   8
         Top             =   1245
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Items           =   3
         Style           =   1
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   7
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   6
         Left            =   6840
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   1245
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   4
         Left            =   6840
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   870
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   1
         Left            =   6840
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   465
         Width           =   240
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   41
         Top             =   450
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   -30
            TabIndex        =   1
            Text            =   "cboEdit"
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   1
         Left            =   4665
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   2430
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   8130
         TabIndex        =   40
         Top             =   450
         Width           =   1230
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   3
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   4
         Top             =   870
         Width           =   2475
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   4
         Left            =   4665
         MaxLength       =   100
         TabIndex        =   5
         Top             =   870
         Width           =   2430
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   8130
         TabIndex        =   37
         Top             =   840
         Width           =   1230
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   6
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1260
         Width           =   5850
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   20
         Left            =   1230
         TabIndex        =   35
         Top             =   3570
         Width           =   2535
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   -30
            Width           =   2520
         End
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   8
         Left            =   6840
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   1635
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   10
         Left            =   6840
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   1995
         Width           =   240
      End
      Begin VB.CommandButton cmdE 
         Caption         =   "…"
         Height          =   220
         Index           =   12
         Left            =   6840
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   2370
         Width           =   240
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   9
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1650
         Width           =   1170
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   11
         Left            =   8130
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2010
         Width           =   1170
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   13
         Left            =   8130
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2385
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1245
         TabIndex        =   31
         Top             =   2775
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   15
         Left            =   4665
         TabIndex        =   30
         Top             =   2775
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   16
         Left            =   8130
         MaxLength       =   6
         TabIndex        =   20
         Top             =   2805
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   17
         Left            =   1245
         TabIndex        =   29
         Top             =   3180
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   -30
            Width           =   2505
         End
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   18
         Left            =   4665
         TabIndex        =   28
         Top             =   3195
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   19
         Left            =   8130
         MaxLength       =   64
         TabIndex        =   23
         Top             =   3210
         Width           =   1170
      End
      Begin VB.Frame fraC 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   21
         Left            =   4665
         TabIndex        =   27
         Top             =   3585
         Width           =   2520
         Begin VB.ComboBox cboE 
            BackColor       =   &H00FDFDFD&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   -30
            Width           =   2490
         End
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   8
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1650
         Width           =   5850
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   10
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2010
         Width           =   5850
      End
      Begin VB.TextBox txtE 
         BackColor       =   &H00FDFDFD&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   12
         Left            =   1245
         MaxLength       =   100
         TabIndex        =   15
         Top             =   2385
         Width           =   5850
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "监护人身份证号"
         Height          =   180
         Index           =   55
         Left            =   360
         TabIndex        =   139
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "基本信息"
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
         Index           =   50
         Left            =   3795
         TabIndex        =   119
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "展开快捷病历"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   480
         MouseIcon       =   "frmDockPatiInfo.frx":C7E4
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   4275
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "单位邮编"
         Height          =   180
         Index           =   7
         Left            =   7350
         TabIndex        =   92
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblN 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "身份证号"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "家庭地址"
         Height          =   180
         Index           =   12
         Left            =   480
         TabIndex        =   61
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "监护人"
         Height          =   180
         Index           =   19
         Left            =   7530
         TabIndex        =   60
         Top             =   3210
         Width           =   540
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "其他证件"
         Height          =   180
         Index           =   3
         Left            =   480
         TabIndex        =   59
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "出生地点"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   58
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "户口地址"
         Height          =   180
         Index           =   8
         Left            =   480
         TabIndex        =   57
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "Rh"
         Height          =   180
         Index           =   21
         Left            =   4410
         TabIndex        =   56
         Top             =   3615
         Width           =   180
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "文化程度"
         Height          =   180
         Index           =   2
         Left            =   7350
         TabIndex        =   55
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "单位电话"
         Height          =   180
         Index           =   11
         Left            =   7350
         TabIndex        =   54
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "家庭电话"
         Height          =   180
         Index           =   13
         Left            =   7350
         TabIndex        =   53
         Top             =   2415
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "血型"
         Height          =   180
         Index           =   20
         Left            =   810
         TabIndex        =   52
         Top             =   3615
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "生育状况"
         Height          =   180
         Index           =   5
         Left            =   7350
         TabIndex        =   51
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "籍贯"
         Height          =   180
         Index           =   4
         Left            =   4230
         TabIndex        =   50
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "区域"
         Height          =   180
         Index           =   1
         Left            =   4230
         TabIndex        =   49
         Top             =   510
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "户口邮编"
         Height          =   180
         Index           =   9
         Left            =   7350
         TabIndex        =   48
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "婚姻状况"
         Height          =   180
         Index           =   14
         Left            =   465
         TabIndex        =   47
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "单位名称"
         Height          =   180
         Index           =   10
         Left            =   480
         TabIndex        =   46
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "家庭邮编"
         Height          =   180
         Index           =   16
         Left            =   7350
         TabIndex        =   45
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "国籍"
         Height          =   180
         Index           =   15
         Left            =   4230
         TabIndex        =   44
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "民族"
         Height          =   180
         Index           =   18
         Left            =   4230
         TabIndex        =   43
         Top             =   3210
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "职业"
         Height          =   180
         Index           =   17
         Left            =   825
         TabIndex        =   42
         Top             =   3210
         Width           =   360
      End
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   375
      Left            =   90
      TabIndex        =   113
      Top             =   450
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3765
      Index           =   1
      Left            =   810
      ScaleHeight     =   3765
      ScaleWidth      =   9750
      TabIndex        =   63
      Top             =   5280
      Width           =   9750
      Begin VB.PictureBox picOutDoc 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   180
         ScaleHeight     =   3015
         ScaleWidth      =   9195
         TabIndex        =   65
         Top             =   540
         Width           =   9200
         Begin VB.PictureBox picSentence 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   585
            ScaleHeight     =   240
            ScaleWidth      =   1155
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   375
            Visible         =   0   'False
            Width           =   1185
            Begin VB.TextBox txtSentence 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Left            =   0
               TabIndex        =   68
               Top             =   30
               Width           =   930
            End
            Begin VB.Image imgSentence 
               Height          =   210
               Left            =   960
               Picture         =   "frmDockPatiInfo.frx":E166
               ToolTipText     =   "请按 * 号键选择"
               Top             =   15
               Width           =   180
            End
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "取消签名(&Q)"
            Height          =   350
            Left            =   6885
            TabIndex        =   77
            Top             =   2400
            Width           =   1200
         End
         Begin VB.PictureBox picPrompt 
            BackColor       =   &H00FDFDFD&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4545
            Picture         =   "frmDockPatiInfo.frx":E690
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   66
            Top             =   2078
            Width           =   260
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "全文编辑(&U)"
            Height          =   350
            Left            =   5595
            TabIndex        =   78
            Top             =   2385
            Width           =   1200
         End
         Begin VB.CommandButton cmdImportEPRDemo 
            Caption         =   "导入范文(&I)"
            Height          =   350
            Left            =   4290
            TabIndex        =   76
            Top             =   2400
            Width           =   1200
         End
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   3
            Left            =   5580
            TabIndex        =   72
            Top             =   900
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EA91
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   4
            Left            =   345
            TabIndex        =   74
            Top             =   1785
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EB2E
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   0
            Left            =   240
            TabIndex        =   69
            Top             =   225
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EBCB
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   2
            Left            =   435
            TabIndex        =   71
            Top             =   1050
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":EC68
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
         Begin RichTextLib.RichTextBox rtfEdit 
            Height          =   800
            Index           =   1
            Left            =   5250
            TabIndex        =   70
            Top             =   105
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16645629
            ScrollBars      =   2
            MaxLength       =   4000
            Appearance      =   0
            TextRTF         =   $"frmDockPatiInfo.frx":ED05
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
         Begin VB.Label lblEPRname 
            AutoSize        =   -1  'True
            Caption         =   "(门诊病历)"
            Height          =   180
            Left            =   4935
            TabIndex        =   128
            Top             =   1785
            Width           =   900
         End
         Begin VB.Line linDoc 
            BorderColor     =   &H00808080&
            X1              =   795
            X2              =   2040
            Y1              =   2715
            Y2              =   2715
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "查体"
            Height          =   180
            Index           =   4
            Left            =   210
            TabIndex        =   126
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "现病史"
            Height          =   180
            Index           =   1
            Left            =   4605
            TabIndex        =   125
            Top             =   300
            Width           =   540
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "主诉"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   124
            Top             =   15
            Width           =   360
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "过去史"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   123
            Top             =   990
            Width           =   540
         End
         Begin VB.Label lblDoc 
            AutoSize        =   -1  'True
            Caption         =   "家族史"
            Height          =   180
            Index           =   3
            Left            =   4665
            TabIndex        =   122
            Top             =   990
            Width           =   540
         End
         Begin VB.Label lblDoctor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   6705
            TabIndex        =   79
            Top             =   1785
            Width           =   450
         End
         Begin VB.Label lblDoctor 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "刘力红"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   7215
            TabIndex        =   75
            Top             =   1785
            Width           =   540
         End
         Begin VB.Label lblTip 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输入病历时按 ~ 键可提取或选择词句示范."
            ForeColor       =   &H00404040&
            Height          =   180
            Left            =   4920
            TabIndex        =   73
            Top             =   2115
            Width           =   3420
         End
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         Caption         =   "快键病历"
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
         Index           =   51
         Left            =   4980
         TabIndex        =   121
         Top             =   285
         Width           =   780
      End
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   11055
      Picture         =   "frmDockPatiInfo.frx":EDA2
      Top             =   1140
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   11055
      Picture         =   "frmDockPatiInfo.frx":155F4
      Top             =   780
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDockPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event EditFullDoc(ByVal lngEPRFileID As Long, ByVal lngFileID As Long, ByVal strDoctor As String, ByVal strIn As String)    '全文编辑病历
Public Event EPRRefresh() '病历刷新
Public Event UpdatePatiInfo(ByVal strBirthday As String, ByVal strAge As String, ByVal strSex As String, ByVal strTag As String)  '身份证号保存后连动项 生日，年龄，性别，strTag 扩展就诊摘要
Public Event UpdatePatiState(ByVal strInfo As String, ByRef strTag As String) '更新病人生命体征，
    '参数：strInfo 用<split>分割，病人ID<split>挂号ID<split>身高.......... strTag 扩展参数
Public Event UpdateDiagInfo(ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)  'strTag 扩展参数

Private Const DColor = &HEEEEEE, EColor = &HFDFDFD, HColor = &HFFDFDF
Private Const GRD_UNEDITCELL_COLOR = &H8000000B  '未编辑的单元格颜色：灰蓝色
Public Event SetEdit()
Private Enum E_ITEM_INDEX
    I身份证号 = 0
    I区域 = 1
    I文化程度 = 2
    I其它证件 = 3
    I籍贯 = 4
    I生育状况 = 5
    I出生地点 = 6
    I单位邮编 = 7
    I户口地址 = 8
    I户口邮编 = 9
    I单位名称 = 10
    I单位电话 = 11
    I家庭地址 = 12
    I家庭电话 = 13
    I婚姻状况 = 14
    I国籍 = 15
    I家庭邮编 = 16
    I职业 = 17
    I民族 = 18
    I监护人 = 19
    I血型 = 20
    IRH = 21
    I去向 = 22
    I发病时间 = 23
    I发病地址 = 24
    I医学警示 = 25
    I其他医学警示 = 26
    I就诊摘要 = 27
    
    I过敏记录 = 53
    I诊断记录 = 54
    I监护人身份证号 = 55
    
    I主诉 = 0
    I现病史 = 1
    I过去史 = 2
    I家族史 = 3
    I查体 = 4
    
    I日期 = 1

End Enum

Private Enum m_Ctl_ID
    picPanel_基本信息 = 0
    picPanel_快键病历 = 1
    picPanel_就诊信息 = 2
    picPanel_过敏源 = 8
    
    picPanel_诊断 = 3
    picPanel_初复诊 = 5
    picPanel_附加 = 6
    picPanel_过敏输入方式 = 8
    fraLine_基本信息 = 0
    fraLine_快键病历 = 1
    fraLine_就诊信息 = 2
     
    opt药品目录 = 5
    opt过敏源 = 4
    opt疾病 = 1
    opt诊断 = 0
    opt初诊 = 3
    opt复诊 = 2
    
    lbl标题基本 = 50
    lbl标题病历 = 51
    lbl标题就诊 = 52
    lbl生命体征 = 28
End Enum

Private Enum AllerColsIndex
    AI_过敏药物
    AI_过敏反应
    AI_过敏时间
    AI_过敏源编码
    AI_药物ID
    AI_过敏来源
End Enum

Private Enum Change_State
    CS_删除行 = -1
    CS_未改变 = 0
    CS_更新行 = 1
    CS_替换行 = 2
    CS_新增行 = 3
End Enum
 
Private Enum PaddType
    PT_出生地点 = 0
    PT_户口地址 = 1
    PT_家庭地址 = 2
End Enum
 
Private Enum DiagColsIndex
    DI_诊断类型 = 0
    DI_关联 = 1
    DI_诊断编码 = 2
    DI_诊断描述 = 3
    DI_中医证候 = 4
    DI_发病时间 = 5
    DI_备注 = 6
    DI_ICD附码 = 7
    DI_是否疑诊 = 8
    DI_增加 = 9
    DI_Del = 10
    DI_诊断ID = 11
    DI_疾病ID = 12
    DI_证候ID = 13
    DI_医嘱IDs = 14 '与当前诊断关联的医嘱ID组成的字符串，医嘱ID间以逗号分割
    DI_诊断分类 = 15
    DI_固定附码 = 16
    DI_附码ID = 17
    DI_诊断来源 = 18
    DI_疾病编码 = 19
    DI_疾病类别 = 20
    DI_证候编码 = 21
    DI_记录日期 = 22
    DI_记录人员 = 23
End Enum

Private mlng病人ID As Long
Private mstr挂号单 As String
Private mstr门诊号 As String
Private mlng挂号ID As Long
Private mlng科室ID As Long '挂号中的执行科室
Private mstr婚姻状况 As String
Private mlng病历文件id As Long
Private mlng病历ID As Long
Private mstr保存人 As String
Private mbln签名 As Boolean
Private mlng执行状态 As Long
Private mbln急 As Boolean
Private mstr出生日期 As String
Private mstr年龄 As String
Private mstr性别 As String
Private mstr姓名 As String
Private mint险类 As Integer '当前病人险类
Private mlng合同单位ID As Long
Private mblnEdit合同单位 As Boolean '是否有修改合同单位的权限，true 可以修改，false不能修改
Private mclsMipModule As zl9ComLib.clsMipModule
Private mobjKernel As zlPublicAdvice.clsPublicAdvice         '临床核心部件
Private mobjPatient As Object
Private mobjCtl As Object '当前活动控件
Private mblnUseEPR As Boolean '是否用可用的快捷病历

Private mblnMoved As Boolean '门诊医生站传入
Private mblnDocInput As Boolean '门诊医生站传入，是否显示快捷病历
Private mbln中医 As Boolean '跟据 mlng科室ID 来判断
Private mbln录中医诊断 As Boolean '参数：门诊西医科允许录入中医诊断
Private mblnEdit As Boolean '内容是否可以修改 false 可以编辑，true 不能编辑
Private mblnChange As Boolean
Private mblnPatiChange As Boolean
Private mblnReturn As Boolean
Private mblnID加密 As Boolean '身份证号用掩码显示
Private mint诊断输入 As Integer '1-允许自由输入,2-从数据库提取输入,3-仅医保病人从数据库输入
Private mint调用 As Integer '调用场合，0-门诊医生工作站，1-门诊医嘱编辑界面
Private mblnOK As Boolean '是否向后台提交了数据，不包含病历相关数据的提交

Private mint简码 As Integer
Private mintAllerInput As Integer                     'AllerInput:过敏输入来源：0-按药品目录输入，1-按过敏源输入
Private mintDiagInput As Integer                      '0-根据诊断标准输入,1-根据疾病编码输入
Private mblnSizeTmp As Boolean
Private mbytSize As Byte '9－小字体，12－大字体
Private mlngTopVsc As Long

Private mclsZip As zlRichEPR.cZip
Private mclsUnZip As zlRichEPR.cUnzip
Private mrsMainInfo As ADODB.Recordset
Private mrsSecdInfo As ADODB.Recordset
Private mrsPreEditCtl As ADODB.Recordset '上一个控件编辑信息，格式：控件类型|是否为数组控件|控件名|控件下标
Private mblnNoSave As Boolean '不保存数据，true时不保存，false保存
Private mblnCboNoClick As Boolean '触发下拉列表Click事件
 
Private mstrCtlName As String '当前正编辑的控件名字
Private mintCtlIndex As Integer '当前正编辑的控件的索引
Private mstrTagAller As String '记录来源标记
Private mstrTagDiagXY As String
Private mstrTagDiagZY As String
Private mblnFreeInput As Boolean      '诊断是否允许自由调整
Private mdatCurDate As Date
Private mblnChk  As Boolean  '是否执行chk点击事件
Private mblnStructAdress As Boolean, mblnShowTown As Boolean
Private mblnUpdate As Boolean '是否更新结构化地址
Private mbln不更新诊断 As Boolean

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal blnEdit As Boolean, ByVal blnMoved As Boolean, Optional ByRef objMip As Object, Optional ByVal int调用 As Integer) As Boolean
'功能：公共接中，用于刷新
    Dim blnTmp As Boolean
    
        mblnUpdate = True
    Call SavePreItem
    mint调用 = int调用
    mlng病人ID = lng病人ID
    mlng挂号ID = lng挂号id
    mblnMoved = blnMoved
    mblnEdit = blnEdit
    If mlng挂号ID = 0 Then
        mblnEdit = True
    End If
    If Not objMip Is Nothing Then Set mclsMipModule = objMip
    mblnNoSave = True
    Call ClearPatiInfo
    If lng挂号id <> 0 Then
        Call LoadPatiInfo
        Call LoadAllerData
        Call LoadDiagData
        If mint调用 = 1 Then
            mblnUseEPR = False
        Else
            mblnUseEPR = CanUseFastEPR
        End If
        If mblnUseEPR Then
            If mblnDocInput Then
                Call LoadDocData
                lblLink.Caption = "收起快捷病历"
            Else
                lblLink.Caption = "展开快捷病历"
            End If
            lblLink.Visible = True
        Else
            lblLink.Visible = False
            PicPanel(picPanel_快键病历).Visible = False
            mblnDocInput = False
        End If
        If Not mblnEdit Then mdatCurDate = zlDatabase.Currentdate
    End If
    Call SetFaceEditable(mblnEdit)
    mblnNoSave = False
    Set mobjCtl = Nothing
End Function

Private Sub SetAllerEdit(ByVal blnReadOnly As Boolean)
    vsAller.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
    vsAller.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    vsAller.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    vsAller.TextMatrix(1, AI_过敏药物) = IIf(blnReadOnly, "—", "")
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    LockWindowUpdate Me.hwnd
    mbytSize = IIf(bytSize = 0, 9, 12)
    Call zlControl.SetPubFontSize(Me, bytSize)
    Call Grid.SetFontSize(vsAller, mbytSize)
    Call Grid.SetFontSize(vsDiagZY, mbytSize)
    Call Grid.SetFontSize(vsDiagXY, mbytSize)
    Set UCPatiVitalSigns.Font = txtE(I区域).Font
    If mblnStructAdress Then
        PatiAddress(PT_出生地点).Font.Size = mbytSize
        PatiAddress(PT_户口地址).Font.Size = mbytSize
        PatiAddress(PT_家庭地址).Font.Size = mbytSize
    End If
    Call SetRTFEditFontSize
    Call SetCtlPos(0)
    Call SetCtlPos(1)
    Call SetCtlPos(2)
    Call Form_Resize
    LockWindowUpdate 0
    Me.Refresh
End Sub

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'功能：根据当前是否只读，设置界面的可编辑属性
    Dim strObjName As String
    Dim i As Long
    Dim objControl As Object
    Dim intTmp As Integer
    
    On Error GoTo errH
    
    For Each objControl In Me.Controls
        strObjName = TypeName(objControl)
        If InStr("TextBox;ComboBox;CheckBox;VSFlexGrid;OptionButton;PatiAddress", strObjName) > 0 Then '
            If strObjName = "TextBox" Then
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.Locked = blnReadOnly
            ElseIf strObjName = "PatiAddress" Then
                objControl.ControlLock = blnReadOnly
                objControl.Enabled = Not blnReadOnly
            ElseIf strObjName = "ComboBox" Then
                objControl.Enabled = True
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.Locked = blnReadOnly
            ElseIf strObjName = "CheckBox" Or strObjName = "OptionButton" Then
                '没有Locked属性,用Enabled实现
                objControl.Enabled = Not blnReadOnly
            ElseIf strObjName = "VSFlexGrid" Then
                '同时注意要在键盘鼠标事件中进行一些控制
                objControl.Editable = IIf(blnReadOnly, flexEDNone, flexEDKbdMouse)
                objControl.BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
                objControl.BackColorBkg = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
            End If
        End If
    Next
    PicPanel(picPanel_基本信息).Enabled = True
    UCPatiVitalSigns.Enabled = Not blnReadOnly
    picOutDoc.Enabled = Not blnReadOnly
    
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).BackColor = IIf(blnReadOnly, vbButtonFace, vbWindowBackground)
    Next
     
    lblLinkAdd.ForeColor = IIf(blnReadOnly, &HC0C0C0, &HC00000)
    
    If Not blnReadOnly Then
        If chkNoAller.Value = 1 Then
            Call SetAllerEdit(True)
        End If
        Call SetDocEditable
    End If
    
    intTmp = IIf(optInfo(opt疾病).Value, 1, 0)
    Call SetInputRoot(0, gint诊断来源, intTmp, optInfo(opt诊断), optInfo(opt疾病))
    If gint诊断来源 = 1 Then
        If intTmp = 1 Then
            optInfo(opt疾病).Value = True
            optInfo(opt诊断).Value = False
        Else
            optInfo(opt疾病).Value = False
            optInfo(opt诊断).Value = True
        End If
    End If
    
    intTmp = IIf(optInfo(opt过敏源).Value, 2, 1)
    If Not gobjPass Is Nothing Then
        Call SetInputRoot(2, gint过敏输入来源, intTmp, optInfo(opt药品目录), optInfo(opt过敏源))
    Else
        Call SetInputRoot(1, 1, 1, optInfo(opt药品目录), optInfo(opt过敏源))
    End If
    If gint过敏输入来源 = 0 Then
        If intTmp = 2 Then
            optInfo(opt药品目录).Value = False
            optInfo(opt过敏源).Value = True
        Else
            optInfo(opt药品目录).Value = True
            optInfo(opt过敏源).Value = False
        End If
    End If
    
    If blnReadOnly Then
        For i = 0 To 5
            optInfo(i).Enabled = Not blnReadOnly
        Next
    End If
    If Not blnReadOnly Then
        If Not mblnEdit合同单位 And mlng合同单位ID <> 0 Then
            txtE(I单位名称).BackColor = vbButtonFace
            txtE(I单位名称).Locked = True
        End If
    End If
    If mint调用 = 0 Then
        UCPatiVitalSigns.Visible = False
        lblN(lbl生命体征).Visible = False
    End If
    Exit Sub
errH:
    If 2 = 1 Then
        Resume
    End If
    err.Clear
End Sub

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    With vsAller
        If .Col = AI_过敏反应 Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AI_过敏药物
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, Optional rsInput As ADODB.Recordset, Optional ByVal strTYTInput As String)
'功能：处理过敏药物的输入
'参数：strTYTInput=太元通合理用药接口返回的字符串
    Dim strSQL As String, curDate As Date
    Dim arrTmp As Variant
    Dim strAllerOld As String, strAllerNew As String
    
    With vsAller
        
        strAllerOld = .Cell(flexcpData, lngRow, AI_过敏药物) & ";" & .TextMatrix(lngRow, AI_过敏源编码)
        If Not gobjPass Is Nothing Then
            If optInfo(opt过敏源).Value Then
                arrTmp = Split(strTYTInput, ";")
                
                If UBound(arrTmp) < 1 Then Exit Sub
                If strAllerOld <> strTYTInput Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_过敏药物) = arrTmp(1)
                    .TextMatrix(lngRow, AI_过敏源编码) = arrTmp(0)
                    .RowData(lngRow) = 0
                End If
            Else
                If Not rsInput Is Nothing Then
                    .RowData(lngRow) = CLng(rsInput!ID)
                    .TextMatrix(lngRow, AI_过敏药物) = NVL(rsInput!名称)
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_过敏药物) = .EditText
                End If
                
                strAllerNew = .TextMatrix(lngRow, AI_过敏药物) & ";" & .TextMatrix(lngRow, AI_过敏源编码)
                
                If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_过敏源编码) = ""
                End If
            End If
        Else
            If optInfo(opt药品目录).Value Then
                If Not rsInput Is Nothing Then
                    .RowData(lngRow) = CLng(rsInput!ID)
                    .TextMatrix(lngRow, AI_过敏药物) = NVL(rsInput!名称)
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_过敏药物) = .EditText
                End If
                
                strAllerNew = .TextMatrix(lngRow, AI_过敏药物) & ";" & .TextMatrix(lngRow, AI_过敏源编码)
                
                If strAllerOld <> strAllerNew Or Val(.RowData(lngRow) & "") <> 0 Then
                    .TextMatrix(lngRow, AI_过敏源编码) = ""
                End If
            Else
                If Not rsInput Is Nothing Then
                    .TextMatrix(lngRow, AI_过敏药物) = rsInput!名称 & ""
                    .TextMatrix(lngRow, AI_过敏源编码) = rsInput!编码 & ""
                    .RowData(lngRow) = 0
                Else
                    .RowData(lngRow) = 0
                    .TextMatrix(lngRow, AI_过敏药物) = .EditText
                End If
            End If
        End If
        .Cell(flexcpData, lngRow, AI_过敏药物) = .TextMatrix(lngRow, AI_过敏药物)
        .TextMatrix(lngRow, AI_药物ID) = Val(.RowData(lngRow) & "")

        If .TextMatrix(lngRow, AI_过敏时间) = "" Then
            .TextMatrix(lngRow, AI_过敏时间) = Format(mdatCurDate, "YYYY-MM-DD")
        End If
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
    End With
End Sub

Private Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True, Optional ByVal strMintime As String, Optional strMaxtTime As String) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
'参数：blnTime=是否处理时间部份
'参数：strMintime=生成时间的下线
'          strOutTime=生成时间的上限
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = mdatCurDate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) Then
        If strMintime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") < Format(strMintime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMintime
            End If
        End If
        If strMaxtTime <> "" Then
            If Format(strTmp, "yyyy-MM-dd HH:mm") > Format(strMaxtTime, "yyyy-MM-dd HH:mm") Then
                strTmp = strMaxtTime
            End If
        End If
        If Not blnTime Then
            strTmp = Format(strTmp, "yyyy-MM-dd")
        End If
    End If
    GetFullDate = strTmp
End Function

Private Sub InitEditData()
'功能：初始化编辑环境和必要的数据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboE(I民族), cboE(I民族).Height * 16)
    Call zlControl.CboSetHeight(cboE(I国籍), cboE(I国籍).Height * 16)
    Call zlControl.CboSetHeight(cboE(I职业), cboE(I职业).Height * 16)
    
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
 
    strSQL = _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '职业' 表名 From 职业 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '民族' 表名 From 民族 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '国籍' 表名 From 国籍 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '血型' 表名 From 血型 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '学历' 表名 From 学历 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '病人去向' 表名 From 病人去向 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '婚姻状况' 表名 From 婚姻状况 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '身份证未录原因' 表名 From 身份证未录原因"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call SetCboFromRec(rsTmp, Array("职业", "民族", "国籍", "血型", "学历", "病人去向", "婚姻状况", "身份证未录原因"), Array(I职业, I民族, I国籍, I血型, I文化程度, I去向, I婚姻状况, I身份证号))
    Call SetCboFromList(Array("0-未生育", "1-生育1胎", "2-生育2胎及以上", "4-不详"), Array(I生育状况))
    Call SetCboFromList(Array("0-未查", "1-阴", "2-阳", "3-不详"), Array(IRH))
    Call SetCboFromList(Array(" ", "小时前", "日前", "周前", "月前", "年前"), Array(I日期), 2)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCboFromRec(ByVal rsTmp As ADODB.Recordset, ByVal arrTab As Variant, ByVal arrCboIdx As Variant)
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
     
    For i = 0 To UBound(arrTab)
        rsTmp.Filter = "表名='" & arrTab(i) & "'"
        If Not rsTmp.EOF Then
            rsTmp.Sort = "编码,ID"
            Set objCboTmp = cboE(arrCboIdx(i))
                objCboTmp.Clear
                
            For j = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!编码) Then
                    objCboTmp.AddItem rsTmp!名称
                Else
                    objCboTmp.AddItem rsTmp!编码 & "-" & rsTmp!名称
                End If
                objCboTmp.ItemData(objCboTmp.NewIndex) = NVL(rsTmp!ID, 0)
                If Val(rsTmp!缺省 & "") = 1 Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                    objCboTmp.Tag = objCboTmp.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
    Next
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboE(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboE(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboE(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub


Private Sub ClearPatiInfo()
    Dim i As Long
  
    mrsMainInfo.Filter = 0
    mrsMainInfo.MoveFirst
    mblnNoSave = True
    For i = 1 To mrsMainInfo.RecordCount
        If mrsMainInfo!控件名 = "txtE" Then
            txtE(mrsMainInfo!Index).Text = ""
        ElseIf mrsMainInfo!控件名 = "PatiAddress" Then
            PatiAddress(mrsMainInfo!Index).Tag = ""
            PatiAddress(mrsMainInfo!Index).Value = ""
        ElseIf mrsMainInfo!控件名 = "cboE" Then
            cboE(mrsMainInfo!Index).ListIndex = -1
            If cboE(mrsMainInfo!Index).Style = 0 Then
                cboE(mrsMainInfo!Index).Text = ""
            End If
        End If
        mrsMainInfo!信息原值 = Null
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    txtSL.Text = ""
    vsAller.Rows = vsAller.FixedRows
    vsAller.AddItem ""
    vsDiagXY.Rows = vsAller.FixedRows
    vsDiagXY.AddItem ""
    vsDiagZY.Rows = vsAller.FixedRows
    vsDiagZY.AddItem ""
    '1、过敏记录，
    '2、过敏记录录入方式
    '3、诊断记录  诊断区分中医西，控件表格的显示问题
    '4、诊断录入方式
    UCPatiVitalSigns.ClearData
    For i = I主诉 To I查体
        rtfEdit(i).Text = ""
    Next
    mblnNoSave = False
End Sub

Private Sub InitBaseInfo()
    Dim arrMainFileds() As Variant, arrSecdFileds() As Variant

    '首页标准改变，初始化记录集
    '1、主记录结构定义
    Set mrsMainInfo = New ADODB.Recordset
    With mrsMainInfo
        .Fields.Append "序号", adInteger, , adFldKeyColumn              '主键，标识信息
        .Fields.Append "信息名", adVarChar, 100, adFldKeyColumn   '信息名称
        '该记录集仅记录一个信息对应一个控件的情况或多个信息对应一个控件，其他情况不填写
        .Fields.Append "控件名", adVarChar, 100, adFldIsNullable      '展示信息的控件名称
        .Fields.Append "Index", adInteger, , adFldIsNullable                '为空时表示不是控件数组
        .Fields.Append "ExpState", adInteger                                        '信息扩展状态，0-不扩展，1-初始扩展，2-加载扩展
        .Fields.Append "页码", adInteger                                                '信息所在的页码
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable  '信息在首页加载时的值
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable  '信息在首页检查时的值
        .Fields.Append "ErrInfo", adVarChar, 4000, adFldIsNullable  '控件录入信息不合法提示信息，
        .Fields.Append "Edit", adInteger                                                 '0-可编辑,1-不可编辑，只用于展示,2-不可编辑不保存
        .Fields.Append "是否改变", adInteger                                          '信息是否有改变0-未改变，1-改变了
        .Fields.Append "来源", adInteger  '信息来至于那张表0－病人信息，1－病人信息从表,-1额外的表
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    '2、次级信息记录集结构定义
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '本记录集的主键
        .Fields.Append "序号", adInteger                                              '标识信息，引用主记录集
        .Fields.Append "控件名", adVarChar, 100                                       '展示信息的控件名称
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable               '行号或控件数组Index
        .Fields.Append "页码", adInteger                                                     '信息所在的页码
        .Fields.Append "原ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable      '信息在首页加载时的值
        .Fields.Append "主信息原值", adVarChar, 2000, adFldIsNullable    '信息的主要部分，标识一个信息是否被彻底改变，信息在首页加载时的值
        .Fields.Append "现ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable      '信息在首页检查时的值
        .Fields.Append "主信息现值", adVarChar, 2000, adFldIsNullable    '信息在首页检查时的值
        .Fields.Append "中医诊候", adVarChar, 2000, adFldIsNullable    '信息在首页检查时的值 同一中诊断，不同诊候不能超过3条记录
        .Fields.Append "Edit", adInteger                                                      '0-可编辑,1-不可编辑，只用于展示,2-不可编辑不保存
        .Fields.Append "改变状态", adInteger                                              '信息改变程度0-未改变，1-次级信息改变，2-主信息改变,3-新增,-1，删除
        .Fields.Append "ID", adBigInt, , adFldIsNullable                             '信息行在数据库中的ID,一般表格类控件使用
        .Fields.Append "Tag", adVarChar, 2000                                           '存储额外数据
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With mrsMainInfo
        arrMainFileds = Array("信息名", "控件名", "Index", "来源")
        '基本信息页
        .AddNew arrMainFileds, Array("国籍", "cboE", I国籍, 0)
        .AddNew arrMainFileds, Array("民族", "cboE", I民族, 0)
        .AddNew arrMainFileds, Array("婚姻状况", "cboE", I婚姻状况, 0)
        .AddNew arrMainFileds, Array("职业", "cboE", I职业, 0)
        .AddNew arrMainFileds, Array("身份证号", "无", -1, 0)
        .AddNew arrMainFileds, Array("身份证号状态", "cboE", I身份证号, 0)
        .AddNew arrMainFileds, Array("外籍身份证号", "cboE", I身份证号, 0)
        .AddNew arrMainFileds, Array("区域", "txtE", I区域, 0)
        .AddNew arrMainFileds, Array("其他证件", "txtE", I其它证件, 0)
        .AddNew arrMainFileds, Array("籍贯", "txtE", I籍贯, 0)
        .AddNew arrMainFileds, Array("出生地点", "txtE", I出生地点, 0)
        .AddNew arrMainFileds, Array("单位名称", "txtE", I单位名称, 0) '界面显示和数据库字段不一样，保存字段为  工作单位
        .AddNew arrMainFileds, Array("单位电话", "txtE", I单位电话, 0)
        .AddNew arrMainFileds, Array("单位邮编", "txtE", I单位邮编, 0)
        .AddNew arrMainFileds, Array("家庭地址", "txtE", I家庭地址, 0)
        .AddNew arrMainFileds, Array("家庭电话", "txtE", I家庭电话, 0)
        .AddNew arrMainFileds, Array("家庭地址邮编", "txtE", I家庭邮编, 0)
        .AddNew arrMainFileds, Array("户口地址", "txtE", I户口地址, 0)
        .AddNew arrMainFileds, Array("户口地址邮编", "txtE", I户口邮编, 0)
        .AddNew arrMainFileds, Array("监护人", "txtE", I监护人, 0)
        .AddNew arrMainFileds, Array("文化程度", "cboE", I文化程度, 1)
        .AddNew arrMainFileds, Array("生育状况", "cboE", I生育状况, 1)
        .AddNew arrMainFileds, Array("血型", "cboE", I血型, 1)
        .AddNew arrMainFileds, Array("RH", "cboE", IRH, 1)
        .AddNew arrMainFileds, Array("摘要", "txtE", I就诊摘要, 0)
        .AddNew arrMainFileds, Array("传染病上传", "chkInfo", Null, 0)
        .AddNew arrMainFileds, Array("去向", "cboE", I去向, 1)
        .AddNew arrMainFileds, Array("发病地址", "txtE", I发病地址, 0)
        .AddNew arrMainFileds, Array("发病时间", "txtE", I发病时间, 0)
        .AddNew arrMainFileds, Array("医学警示", "txtE", I医学警示, 1)
        .AddNew arrMainFileds, Array("其他医学警示", "txtE", I其他医学警示, 1)
        .AddNew arrMainFileds, Array("无过敏记录", "chkNoAller", Null, 1)
        .AddNew arrMainFileds, Array("监护人身份证号", "txtE", I监护人身份证号, 1)
         
        If mblnStructAdress Then
            .AddNew arrMainFileds, Array("出生地点结构化", "PatiAddress", PT_出生地点, 0)
            .AddNew arrMainFileds, Array("户口地址结构化", "PatiAddress", PT_户口地址, 0)
            .AddNew arrMainFileds, Array("家庭地址结构化", "PatiAddress", PT_家庭地址, 0)
        End If
        
        .AddNew arrMainFileds, Array("身命体征", "UCPatiVitalSigns", Null, -1)
        .AddNew arrMainFileds, Array("主诉", "rtfEdit", I主诉, -1)
        .AddNew arrMainFileds, Array("现病史", "rtfEdit", I现病史, -1)
        .AddNew arrMainFileds, Array("过去史", "rtfEdit", I过去史, -1)
        .AddNew arrMainFileds, Array("家族史", "rtfEdit", I家族史, -1)
        .AddNew arrMainFileds, Array("查体", "rtfEdit", I查体, -1)
        
    End With
    
    Set mrsPreEditCtl = New ADODB.Recordset
    With mrsPreEditCtl
        .Fields.Append "控件类型", adVarChar, 100              '如TextBox,Combox
        .Fields.Append "控件名", adVarChar, 100
        .Fields.Append "Index", adInteger, , adFldIsNullable   '行号或控件数组Index，如果不是控件数组为－1
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
        .AddNew Array("控件类型", "控件名", "Index"), Array("", "", -1)
        .MoveFirst
    End With
End Sub

Private Sub SetCurCtlInfo(ByVal strType As String, ByVal strName As String, Optional ByVal Index As Integer = -1)
'功能：记录当前控件
    '在记录当前控件之前保存上一个控件的值
    Call SavePreItem(1)
    Call mrsPreEditCtl.Update(Array("控件类型", "控件名", "Index"), Array(strType, strName, Index))
    If strName = "txtE" Then
        Set mobjCtl = txtE(Index)
'    ElseIf strName = "PatiAddress" Then
'        Set mobjCtl = PatiAddress(Index)
    ElseIf strName = "cboE" Then
        Set mobjCtl = cboE(Index)
    Else
        Set mobjCtl = Nothing
    End If
End Sub

Private Sub SavePreItem(Optional intTpye As Integer = 0)
'功能：保存数据
    Dim strType As String
    Dim strName As String
    Dim Index As Integer
    Dim blnDo As Boolean
    Dim objCtl As Object
    Dim strValue As String
    Dim str身份证号 As String
    Dim strMsg As String
    
    If mblnNoSave Then Exit Sub
    If mblnEdit Then Exit Sub
    If intTpye = 0 Then mblnUpdate = True
    If Not mrsPreEditCtl Is Nothing Then
        If mrsPreEditCtl.RecordCount > 0 And Not mrsPreEditCtl.EOF Then
            If mrsPreEditCtl!控件名 & "" <> "" Then
                blnDo = True
            End If
        End If
    End If
    
    If blnDo Then
        strType = mrsPreEditCtl!控件类型 & ""
        strName = mrsPreEditCtl!控件名 & ""
        Index = Val(mrsPreEditCtl!Index)
        Select Case strType
        Case "TextBox"
            If strName = "txtE" And Not txtE(Index).Locked Then
                If Index <> I监护人身份证号 Then
                    Call UpDateInfo(txtE(Index).Text, "txtE", Index)
                End If
            End If
        Case "CheckBox"
            If "chkNoAller" = strName Then
                 Call UpDateInfo(chkNoAller.Value, "chkNoAller")
            Else
                Call UpDate挂号信息("传染病上传", chkInfo.Value)
            End If
        Case "PatiAddress"
            Call UpDate结构化地址(Index)
        Case "OptionButton"
            If Index = opt初诊 Or Index = opt复诊 Then
                Call UpDate挂号信息("复诊", IIf(optInfo(opt复诊).Value, 1, 0))
            End If
        Case "VSFlexGrid"
            If strName = "vsAller" Then
                Call UpDateAller
            ElseIf strName = "vsDiagXY" Then
                Call UpDateDiag(vsDiagXY)
            ElseIf strName = "vsDiagZY" Then
                Call UpDateDiag(vsDiagZY)
            End If
        End Select
        If strName = "UCPatiVitalSigns" Then
            Call UCPatiVitalSigns_Validate(False)
        ElseIf strName = "rtfEdit" Then
            Call rtfEdit_Validate(Index, False)
        End If
    End If
End Sub

Private Sub LoadPatiInfo()
'功能：加载病人基本信息
    Dim rsTmp As ADODB.Recordset
    Dim rsOther As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long, bln中医 As Boolean
    Dim lngidx As Long
    Dim strValue As String
    
    On Error GoTo errH
    
    Call ClearPatiInfo
    
    strSQL = "Select b.NO,B.Id As 挂号id,A.门诊号,A.病人id,A.出生地点,A.出生日期,A.身份证号,Null as 身份证号状态,Null as 外籍身份证号,A.其他证件, A.职业, A.民族, A.国籍, A.籍贯, A.区域,A.婚姻状况," & vbNewLine & _
        "       A.家庭地址, A.家庭电话, A.家庭地址邮编, A.监护人, A.户口地址, A.户口地址邮编, A.合同单位id, A.工作单位 as 单位名称, A.单位电话, A.单位邮编, Nvl(A.险类, 0) as 险类," & vbNewLine & _
        "       Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) 年龄, B.发病时间, B.发病地址, B.传染病上传,B.急诊, B.复诊," & vbNewLine & _
        "       Nvl(Nvl(B.续诊科室id, Decode(B.转诊状态, 1, B.转诊科室id, Null)), B.执行部门id) As 科室id, B.摘要, B.社区,b.执行状态,a.合同单位ID" & vbNewLine & _
        "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
        "Where A.病人id = B.病人id And b.id=[1] And B.记录性质=1 And B.记录状态=1"
     If mblnMoved Then
        strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng挂号ID)
    If rsTmp.EOF Then Exit Sub
    
    mstr挂号单 = rsTmp!NO & ""
    mstr门诊号 = rsTmp!门诊号 & ""
    mstr出生日期 = Format(rsTmp!出生日期 & "", "yyyy-MM-dd")
    mstr年龄 = rsTmp!年龄 & ""
    mstr姓名 = rsTmp!姓名 & ""
    mstr性别 = rsTmp!性别 & ""
    mint险类 = Val(rsTmp!险类 & "")
    mlng合同单位ID = Val(rsTmp!合同单位id & "")
    mlng执行状态 = Decode(NVL(rsTmp!执行状态, 0), 0, 0, 2, 1, 1, 2)
    
    strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID=[1] And (就诊ID=[2] Or 就诊ID is Null and instr(',去向,无过敏记录,',','||信息名||',')=0) Order by Nvl(就诊ID,999999999)"
    Set rsOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    Call LoadCache(rsTmp, rsOther)
    
    mrsMainInfo.Filter = 0
    mrsMainInfo.MoveFirst
    mblnNoSave = True

    For i = 1 To mrsMainInfo.RecordCount
        lngidx = Val(mrsMainInfo!Index & "")
        strValue = mrsMainInfo!信息原值 & ""
        
        If mrsMainInfo!控件名 = "txtE" Then
            txtE(lngidx).Text = strValue
            txtE(lngidx).Tag = ""
            txtE(lngidx).BackColor = vbWindowBackground
        ElseIf mrsMainInfo!控件名 = "PatiAddress" Then
            Call zlReadAddrInfo(PatiAddress(Val(mrsMainInfo!Index)), Val(NVL(mlng病人ID)), 0, Decode(Val(mrsMainInfo!Index), PT_出生地点, 1, PT_户口地址, 4, PT_家庭地址, 3), strValue)
            If mrsMainInfo!信息原值 <> strValue Then
                mrsMainInfo!信息原值 = PatiAddress(Val(mrsMainInfo!Index)).Value
                mrsMainInfo.Update
            End If
        ElseIf mrsMainInfo!控件名 = "cboE" Then
            If lngidx <> I身份证号 Then
                If Mid(strValue, 1, 1) = Chr(30) Then
                    strValue = zlStr.TrimEx(strValue, Chr(30))
                End If
                Call GetCboIndex(cboE(lngidx), strValue)
            End If
        End If
        mrsMainInfo.MoveNext
    Next
    
    mblnReturn = True
    '*身份证号单独加载
    lngidx = I身份证号
    mrsMainInfo.Filter = "控件名='无' and Index=-1 and 信息名='身份证号'"
    If Not mrsMainInfo.EOF Then
        strValue = mrsMainInfo!信息原值 & ""
        If Mid(strValue, 1, 1) = Chr(30) Then
            strValue = zlStr.TrimEx(strValue, Chr(30))
        End If
        Call GetCboIndex(cboE(lngidx), strValue)
        If cboE(lngidx).ListIndex = -1 And strValue <> "" Then
            cboE(lngidx).Tag = strValue
            If mblnID加密 Then
                strValue = Mid(strValue, 1, 12) & String(Len(Mid(strValue, 13, 2)), "*") & Mid(strValue, 15)
            End If
            cboE(lngidx).Text = strValue
        End If
    End If
    If cboE(lngidx).Text = "" Then
        mrsMainInfo.Filter = "控件名='cboE' and Index=" & lngidx & " and 信息名='身份证号状态'"
        If Not mrsMainInfo.EOF Then
            strValue = mrsMainInfo!信息原值 & ""
            If Mid(strValue, 1, 1) = Chr(30) Then
                strValue = zlStr.TrimEx(strValue, Chr(30))
            End If
            Call GetCboIndex(cboE(lngidx), strValue)
        End If
        If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) <> "中国" Then
            If cboE(lngidx).Text = "" Then
                mrsMainInfo.Filter = "控件名='cboE' and Index=" & lngidx & " and 信息名='外籍身份证号'"
                If Not mrsMainInfo.EOF Then
                    strValue = mrsMainInfo!信息原值 & ""
                    cboE(lngidx).Tag = strValue
                    cboE(lngidx).Text = strValue
                End If
            End If
        End If
    End If
    mblnReturn = False
    '1、过敏记录，
    '2、过敏记录录入方式
    '3、诊断记录  诊断区分中医西，控件表格的显示问题
    '4、诊断录入方式
    
    optInfo(opt复诊).Value = Val(rsTmp!复诊 & "") = 1
    optInfo(opt初诊).Value = Val(rsTmp!复诊 & "") = 0
    
    chkInfo.Value = Val(rsTmp!传染病上传 & "")
    
    If mint调用 = 1 Then
        Call UCPatiVitalSigns.LoadPatiVitalSigns(mlng病人ID, mlng挂号ID)
        strSQL = UCPatiVitalSigns.GetSaveSQL(mlng病人ID, mlng挂号ID)
        UCPatiVitalSigns.Visible = True
        lblN(lbl生命体征).Visible = True
    Else
        strSQL = ""
        UCPatiVitalSigns.Visible = False
        lblN(lbl生命体征).Visible = False
    End If
    
    mrsMainInfo.Filter = "控件名='UCPatiVitalSigns'"
    mrsMainInfo!信息原值 = strSQL
    mrsMainInfo.Update
    
    mbln急 = Val(rsTmp!急诊 & "") = 1
    mlng科室ID = Val(rsTmp!科室ID & "")
    
    mbln中医 = Sys.DeptHaveProperty(mlng科室ID, "中医科")
    mbln中医 = mbln中医 Or mbln录中医诊断
    mblnNoSave = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCache(ByVal rsMain As ADODB.Recordset, ByVal rsSecond As ADODB.Recordset)
'功能：初始化缓存信息
    Dim i As Long
    Dim strTmp As String
    
    On Error GoTo errH
    
    mrsMainInfo.Filter = "来源=0"
    For i = 1 To mrsMainInfo.RecordCount
        If mrsMainInfo!控件名 <> "PatiAddress" Then
            mrsMainInfo!信息原值 = rsMain(CStr(mrsMainInfo!信息名)).Value
        Else
            mrsMainInfo!信息原值 = rsMain(Replace(CStr(mrsMainInfo!信息名), "结构化", "")).Value
        End If
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    
    mrsMainInfo.Filter = "来源=1"
    For i = 1 To mrsMainInfo.RecordCount
        rsSecond.Filter = "信息名='" & mrsMainInfo!信息名 & "'"
        If Not rsSecond.EOF Then
            mrsMainInfo!信息原值 = NVL(rsSecond!信息值)
            mrsMainInfo.Update
        End If
        mrsMainInfo.MoveNext
    Next
    
    mrsMainInfo.Filter = "控件名='txtE' and Index=" & I发病时间
    If Not mrsMainInfo.EOF Then
        strTmp = rsMain(CStr(mrsMainInfo!信息名)).Value & ""
        If strTmp <> "" Then
            mrsMainInfo!信息原值 = Format(rsMain(CStr(mrsMainInfo!信息名)).Value, "yyyy-MM-dd HH:MM")
            mrsMainInfo.Update
        End If
    End If
    
    mrsMainInfo.Filter = "控件名='cboE' and Index=" & I身份证号 & " and 信息名='身份证号状态' or 信息名='外籍身份证号'"
    If Not mrsMainInfo.EOF Then
        rsSecond.Filter = "信息名='身份证号状态'"
        If Not rsSecond.EOF Then
            mrsMainInfo!信息原值 = NVL(rsSecond!信息值)
            mrsMainInfo.Update
        End If
    End If
     
    mrsMainInfo.Filter = "控件名='cboE' and Index=" & I身份证号 & " and  信息名='外籍身份证号'"
    If Not mrsMainInfo.EOF Then
        rsSecond.Filter = "信息名='外籍身份证号'"
        If Not rsSecond.EOF Then
            mrsMainInfo!信息原值 = NVL(rsSecond!信息值)
            mrsMainInfo.Update
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboE_Change(Index As Integer)
    If Index = I身份证号 Then
        Call cboSpecificInfoChange(Index)
    End If
End Sub

Private Sub chkInfo_Click()
    Call UpDate挂号信息("传染病上传", chkInfo.Value)
End Sub

Private Function SaveRTFData(ByVal lng病历ID As Long, Optional blnSign As Boolean) As Boolean
'功能：保存病人病历格式RTF数据
'参数：
    Dim strZipFile As String, strTempFile As String, i As Long
    Dim bFinded As Boolean, lngStartPos As Long, lngEndPos As Long, arrTmp As Variant
    Dim strContent As String, lngRecID As Long
    
    If mlng病历ID = 0 Then
        lngRecID = lng病历ID
    Else
        lngRecID = mlng病历ID
    End If
    
    If blnSign = False Then
        '替换提纲内容
        edtEditor.Freeze
        edtEditor.ForceEdit = True
        
        For i = 0 To lblDoc.UBound
            bFinded = FindOutLinePosition(edtEditor, CStr(lblDoc(i).Tag), lngStartPos, lngEndPos)
            If bFinded Then
                strContent = rtfEdit(i).Text    '去掉尾部的回车或换行
                Do While Len(strContent) > 2
                    If Mid(strContent, Len(strContent) - 1) = vbLf Or Mid(strContent, Len(strContent) - 1) = vbCr Then
                        strContent = Mid(strContent, Len(strContent) - 1)
                    Else
                        Exit Do
                    End If
                Loop
                edtEditor.Range(lngStartPos, lngEndPos).Text = strContent
            End If
        Next
        
        edtEditor.UnFreeze
        edtEditor.ForceEdit = False
        '要素内容更新
        If mlng病历ID = 0 Then Call ElementsUpdate(lngRecID)
    End If
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    '压缩文件
    strZipFile = zlFileZip(strTempFile)
    '保存格式
    Call Sys.SaveLob(glngSys, 5, lngRecID, strZipFile)
    
    '删除临时文件
    If strTempFile <> "" Then Kill strTempFile
    If strZipFile <> "" Then Kill strZipFile

    SaveRTFData = True
    Exit Function
errH:
    SaveRTFData = False
End Function

Private Function ElementsUpdate(ByVal lng病历ID As Long) As Boolean
'功能：更新Editor控件中的替换要素内容，以便保存为RTF文件
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSQL = "Select 对象标记,ID From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 终止版=0 and 保留对象 =0 And 替换域 =1 order by 对象标记 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(NVL(rsTmp("对象标记"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.Count
        If ThisElements(i).替换域 = 1 Then
            ThisElements(i).内容文本 = GetReplaceEleValue(ThisElements(i).要素名称, mlng病人ID, mlng挂号ID, 1, 0)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).替换域 = 1 And ThisElements(i).自动转文本 Then
            EleToString edtEditor, ThisElements(i)     '自动转化为纯文本（暂时不删除该要素）
        End If
    Next
    Set ThisElements = Nothing
End Function

Private Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    
    bBeteenKeys = FindNextKey(edtThis, 0, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str内容 As String
        str内容 = Ele.内容文本
        lngLen = Len(str内容)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str内容
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
    End If
End Sub

Private Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lng医嘱ID As Long) As String

    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5]) From Dual"
    err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱ID)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

Private Function FindOutLinePosition(ByRef edtThis As Object, ByVal strOName As String, ByRef lngS As Long, lngE As Long) As Boolean
'功能：根据指定的提纲名称，返回提纲内容文本的起止位置
    Dim blnFindedNext As Boolean, lngCur As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim strTmp As String
    
    bFinded = True
    While bFinded
        bFinded = FindNextKey(edtThis, lngCur, "O", 0, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = edtThis.Range(lKEE, lKEE + Len(strOName))
            If strOName = strTmp Then
                lngS = lKEE + Len(strOName)
                blnFindedNext = FindNextAnyKey(edtThis, lngS, strTmp, lKSS, lKSE, lKES, lKEE, 0, bNeeded)
                If blnFindedNext Then
                    lngE = lKSS
                Else
                    lngE = Len(edtThis.Text)
                End If
                Do While lngE > lngS + 1    '去掉尾部的回车或换行
                    If edtThis.Range(lngE - 1, lngE) = vbLf Or edtThis.Range(lngE - 1, lngE) = vbCr Then
                        lngE = lngE - 1
                    Else
                        Exit Do
                    End If
                Loop
                FindOutLinePosition = True
                Exit Function
            Else
                lngCur = lKEE
            End If
        End If
    Wend
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Private Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    Set mclsZip = New zlRichEPR.cZip
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    Set mclsZip = Nothing
End Function

Private Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim objFSO As New Scripting.FileSystemObject    'FSO对象
    
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If objFSO.FileExists(strZipPath & "TMP.RTF") Then objFSO.DeleteFile strZipPath & "TMP.RTF"
    
    Set mclsUnZip = New zlRichEPR.cUnzip
    With mclsUnZip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
    Set mclsUnZip = Nothing
End Function

Private Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Private Function FindNextAnyKey(ByRef edtThis As Object, _
    ByRef lngCurPosition As Long, _
    ByRef strKeyType As String, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef lngKey As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = .TOM.TextDocument.Range(i - 2, i - 1)
                lngKSS = i - 2 '转换为0开始的坐标位置。
                lngKSE = i + 14
                lngKES = j - 2
                lngKEE = j + 14
                lngKey = Val(.TOM.TextDocument.Range(i + 1, i + 9))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 10, i + 11))
                FindNextAnyKey = True
            End If
        End If
    End With
End Function
   
Private Sub GetSQLOutDoc(ByRef arrSQL As Variant, ByVal lng病历ID As Long)
'功能：组织快捷病历的数据保存SQL
'参数：lng病历ID-新增时传入新取的病历ID
    Dim i As Long, k As Long
    Dim strTmp(5) As String
    
    If mlng病历ID = 0 Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then Exit Sub     '新增时，如果没有填内容，则不保存
    End If
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
     
    If mlng病历ID = 0 Then
        If rtfEdit(I主诉).Locked Then Exit Sub
        arrSQL(UBound(arrSQL)) = "Zl_简单门诊病历_Update(1," & mlng病人ID & "," & _
            mlng挂号ID & "," & mlng科室ID & "," & mlng病历文件id & "," & lng病历ID & ",'" & UserInfo.姓名 & "','" & _
            Replace(Trim(rtfEdit(I主诉).Text), "'", "’") & "','" & Replace(Trim(rtfEdit(I家族史).Text), "'", "’") & "','" & Replace(Trim(rtfEdit(I现病史).Text), "'", "’") & "','" & _
            Replace(Trim(rtfEdit(I查体).Text), "'", "’") & "','" & Replace(Trim(rtfEdit(I过去史).Text), "'", "’") & "')"
    Else
        k = 0
        For i = 0 To rtfEdit.UBound
            If rtfEdit(i).Locked = False Then
                strTmp(i) = rtfEdit(i).Tag & "|" & Replace(Trim(rtfEdit(i).Text), "'", "’")
                k = k + 1
            End If
        Next
        If k = 0 Then Exit Sub
        arrSQL(UBound(arrSQL)) = "Zl_简单门诊病历_Update(2," & mlng病人ID & "," & _
            mlng挂号ID & "," & mlng科室ID & ",0," & mlng病历ID & ",'" & UserInfo.姓名 & "','" & _
            strTmp(0) & "','" & strTmp(3) & "','" & strTmp(1) & "','" & strTmp(4) & "','" & strTmp(2) & "')"
    End If
End Sub

Private Function GetEPRDoc() As zlRichEPR.cEPRDocument
'功能：读取病历文件的RTF数据到editor控件中，并返回文档对象
    Dim objDoc As New zlRichEPR.cEPRDocument
   
    objDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, mlng病历ID, cprPF_门诊, mlng病人ID, mlng挂号ID, , mlng科室ID
    
    If objDoc.ReadFileStructure(edtEditor) = True Then
        Set GetEPRDoc = objDoc
    End If
End Function

Private Sub chkInfo_GotFocus()
    Call SetCurCtlInfo(TypeName(chkInfo), "chkInfo")
End Sub

Private Sub chkInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chkNoAller_Click()
    Dim strValue As String
    If mblnChk = False Then
        strValue = chkNoAller.Value
        If vsAller.TextMatrix(vsAller.FixedRows, AI_过敏药物) <> "" And vsAller.TextMatrix(vsAller.FixedRows, AI_过敏药物) <> "—" Then
            If strValue = "1" Then
                MsgBox "已经有过敏药物，不能标记为无。", vbInformation, gstrSysName
                mblnChk = True
                chkNoAller.Value = 0
                Exit Sub
            End If
        End If
        Call SetAllerEdit(strValue = "1")
        Call UpDateInfo(strValue, "chkNoAller")
    End If
    mblnChk = False
End Sub

Private Sub cmdDiagMove_Click(Index As Integer)
'功能：移动诊断行
    Dim strTmp As String
    Dim i As Long, lngRow As Long
    Dim vsDiag As VSFlexGrid                '当前诊断表格
    Dim intStep As Integer                  '移动位置，1-向下移动，-1向上移动

    If Index = 0 Or Index = 1 Then
        Set vsDiag = vsDiagXY               '西医
    Else
        Set vsDiag = vsDiagZY               '中医
    End If
    
    If vsDiag.Editable = flexEDNone Then
        Exit Sub
    End If
    
'    intStep=移动位置，1-向下移动，-1向上移动
    If Index = 0 Or Index = 2 Then
        intStep = -1
    ElseIf Index = 1 Or Index = 3 Then
        intStep = 1
    End If
    
    With vsDiag
        If .Row < 0 Then
            Exit Sub
        Else
            If Not DiagRowCanMove(vsDiag, intStep, .Row) Then Exit Sub
            lngRow = IIf(intStep = 1, .Row, .Row + intStep)
        End If
        
        For i = .FixedCols To .Cols - 1
            '交换界面数据
            strTmp = .TextMatrix(.Row + intStep, i)
            .TextMatrix(.Row + intStep, i) = .TextMatrix(.Row, i)
            .TextMatrix(.Row, i) = strTmp
            '交换隐藏数据
            strTmp = .Cell(flexcpData, .Row + intStep, i)
            .Cell(flexcpData, .Row + intStep, i) = .Cell(flexcpData, .Row, i)
            .Cell(flexcpData, .Row, i) = strTmp
        Next
        
        '交换隐藏数据
        strTmp = .RowData(.Row + intStep)
        .RowData(.Row + intStep) = .RowData(.Row)
        .RowData(.Row) = Val(strTmp)
        Call SetDiagReletedInfo(vsDiag)
        .Row = .Row + intStep
    End With
End Sub

Private Function DiagRowCanMove(ByRef vsDiag As VSFlexGrid, ByVal intStep As Integer, ByVal lngRow As Long) As Boolean
'功能：设置诊断移动控件状态
'参数：intStep=移动位置，1-向下移动，-1向上移动
'   lngRow=需判定的行，一般为当前行
    Dim lngBgn As Long, lngEnd As Long
    Dim i As Long
    
    lngBgn = 0
    lngEnd = 0
    '根据当前行的位置设置移动诊断控件的可用性
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                If lngBgn < .FixedRows Then lngBgn = i
                lngEnd = i
            End If
        Next
    End With
    
    If lngBgn = lngEnd Then '只有一行诊断，则不可移动
        DiagRowCanMove = False
    ElseIf lngRow = lngBgn Then '当前行是本分类第一行，则只能下移
        DiagRowCanMove = intStep = 1
    ElseIf lngRow = lngEnd Then '当前行是本分类最后一行，则只能上
        DiagRowCanMove = intStep = -1
    Else  '当前行是本分类中间某一行，则可以上下移动
        DiagRowCanMove = True
    End If
End Function

Private Sub cmdSaveZY_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    Dim Index As Long
    Dim objTxt As Object
    Set objTxt = txtE(I就诊摘要)
    Index = I就诊摘要
    If objTxt.Locked Then Exit Sub
    If Trim(objTxt.Text) = "" Then
        MsgBox "请输入摘要内容。", vbInformation, gstrSysName
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From 常用就诊摘要 Where 名称=[1] And 人员ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(objTxt.Text), UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该内容已经在常用摘要中。", vbInformation, gstrSysName
        If objTxt.Enabled Then objTxt.SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(objTxt.Text, CByte(mint简码))
    strSQL = "Zl_常用就诊摘要_Update(0,Null,'" & Replace(objTxt.Text, "'", "''") & "','" & strSQL & "'," & UserInfo.ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem objTxt.hwnd, CB_ADDSTRING, 0, objTxt.Text
    MsgBox "已设置为常用摘要。", vbInformation, gstrSysName
    If objTxt.Enabled Then objTxt.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdShowZY_Click()
    If txtE(I就诊摘要).Locked Then Exit Sub
    If AbstractSelect("") Then Exit Sub
End Sub

Private Sub cmdSign_Click()
    Dim i As Long, str对象属性 As String, strSource As String, strSQL As String
    Dim arrSQL As Variant, blnTrans As Boolean
    Dim patiSign As cEPRSign, objEPRDoc As cEPRDocument
    
    If mbln签名 = False Then
        For i = 0 To rtfEdit.UBound
            If Trim(rtfEdit(i).Text) <> "" Then Exit For
        Next
        If i > rtfEdit.UBound Then
            MsgBox "请先输入病历信息后再进行签名。", vbInformation, gstrSysName
            Exit Sub    '新增时，如果没有填内容，则不保存
        End If
                                         
        If edtEditor.Text = "" Then
            If ReadRTFData(mlng病历ID) = False Then Exit Sub
        End If
        
        strSource = edtEditor.Text
        If cmdSign.Visible And cmdSign.Enabled Then cmdSign.SetFocus
        Set patiSign = frmOutDocterSign.ShowMe(Me, strSource, mlng病人ID, mlng挂号ID)
        If patiSign Is Nothing Then Exit Sub
        With patiSign
            .Key = "1"
            str对象属性 = .签名方式 & ";" & .签名规则 & ";" & .证书ID & ";" & IIf(.显示手签, 1, 0) & ";" & _
                    Format(.签名时间, "yyyy-mm-dd hh:mm:ss") & ";" & .显示时间 & ";" & .签名要素
                    
            strSQL = "Zl_简单门诊病历_签名(1," & mlng病历ID & ",'" & str对象属性 & "','" & UserInfo.姓名 & "','" & _
                    .前置文字 & "','" & .时间戳 & "','" & .签名级别 & "','" & .签名信息 & "')"
        End With
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.InsertIntoEditor(edtEditor, Len(edtEditor.Text), , objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
    Else
        If MsgBox("你确定要取消签名吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        Set patiSign = GetSign(mlng病历ID)
        If patiSign Is Nothing Then Exit Sub
        
        Set objEPRDoc = GetEPRDoc()
        If objEPRDoc Is Nothing Then Exit Sub
        Call patiSign.DeleteFromEditor(edtEditor, objEPRDoc)
        Set objEPRDoc = Nothing
        Set patiSign = Nothing
        
        strSQL = "Zl_简单门诊病历_签名(0," & mlng病历ID & ")"
    End If
    
   
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If SaveRTFData(mlng病历ID, True) = False Then GoTo errH
    gcnOracle.CommitTrans: blnTrans = False
    
        
    Call LoadDocData
    
    RaiseEvent EPRRefresh
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objEPRDoc = Nothing: Set patiSign = Nothing
    Call SaveErrLog
End Sub

Private Sub cmdImportEPRDemo_Click()
    Dim objImportEPRDemo As New frmImportEPRDemo
    Dim rsDemo As New Recordset
    
    If mlng病历ID <> 0 Then
        MsgBox "该病人已经产生了病历文件，不能再导入范文。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If objImportEPRDemo.ShowMe(Me, mlng病历文件id, mlng病人ID, mlng挂号ID, rsDemo) > 0 Then
    Call SetDocData(rsDemo, 1)
    End If
End Sub

Private Sub SetDocData(ByVal rsTmp As Recordset, ByVal intType As Integer)
'功能：设置快捷面板的内容
'参数：intType=0病历读取，intType=1范文导入，不清空病历ID
    Dim i As Long, j As Long, arrTmp As Variant
    Dim strContent As String
    
    With rsTmp
        If .RecordCount > 0 Then
            arrTmp = Split("-10,2,3,5,6", ",") '病人主诉,现病史,既往史,家族史,体格检查
            For i = 0 To UBound(arrTmp)
                .Filter = "预制提纲id=" & arrTmp(i)
                rtfEdit(i).Text = ""
                If intType = 1 Then
                    '导入范文后所有都可用。
                    rtfEdit(i).Locked = False
                    rtfEdit(i).BackColor = HColor
                End If
                For j = 1 To .RecordCount
                    If j = 1 Then
                        strContent = "" & !内容文本
                        If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                        rtfEdit(i).Text = strContent
                        If intType = 0 Then rtfEdit(i).Tag = !ID
                    Else
                        rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !内容文本
                        If intType = 0 Then rtfEdit(i).Tag = rtfEdit(i).Tag & "," & !ID
                    End If
                    .MoveNext
                Next
            Next
        End If
    End With
End Sub

Private Function ReadRTFData(ByVal lng病历ID As Long) As Boolean
'功能：读取病历文件的RTF数据到editor控件中
    Dim strZipFile As String, strTempFile As String
    Dim lngRecID As Long
    
    If mlng病历ID = 0 Then
        lngRecID = lng病历ID
    Else
        lngRecID = mlng病历ID
    End If
    
    On Error GoTo errH
        
    '判断本次是不是新增的病历
    If mlng病历ID = 0 Then
        strZipFile = Sys.ReadLob(glngSys, 1, mlng病历文件id)
    Else
        strZipFile = Sys.ReadLob(glngSys, 5, lngRecID)
    End If
    
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
     '删除临时文件
    If strTempFile <> "" Then Kill strTempFile
    If strZipFile <> "" Then Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Private Function GetSign(ByVal lng病历ID As Long) As cEPRSign
'功能：获取当前用户的签名对象
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim OneSign As New cEPRSign, intSign As Integer, strUserName As String
    
    strUserName = UserInfo.姓名
    intSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    If intSign = 1 Then
        strSQL = "Select 签名 From 人员表 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
        If rsTemp.RecordCount > 0 Then
            If Not IsNull(rsTemp!签名) Then strUserName = rsTemp!签名
        End If
    End If
    strSQL = "Select Id,对象标记 From 电子病历内容 Where 文件id= [1] And 对象类型=8 And Instr(';'||内容文本||';',[2])>0 Order By 对象标记"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历ID, ";" & strUserName & ";")
    If rsTemp.RecordCount > 0 Then
        OneSign.Key = NVL(rsTemp!对象标记, 0)
        If OneSign.GetSignFromDB(rsTemp!ID) = True Then Set GetSign = OneSign
    End If
End Function

Private Sub cmdUpdate_Click()
    RaiseEvent EditFullDoc(mlng病历文件id, mlng病历ID, mstr保存人, lblDoctor(1).Tag)
End Sub

Private Sub lblLink_Click()
'功能：设置快捷病历的可见性
    mblnDocInput = Not mblnDocInput
    Call zlDatabase.SetPara("显示病历快捷输入", IIf(mblnDocInput, 1, 0), glngSys, p门诊医生站, InStr(";" & gstrPrivs & ";", ";参数设置;") > 0)
    
    If mblnDocInput Then Call LoadDocData
    PicPanel(picPanel_快键病历).Visible = mblnDocInput
    If mblnDocInput Then
        lblLink.Caption = "收起快捷病历"
    Else
        lblLink.Caption = "展开快捷病历"
    End If
    If Not mblnEdit Then Call SetDocEditable
 
    Call Form_Resize
End Sub

Private Sub lblLinkAdd_Click()
    If lblLinkAdd.ForeColor = &HC0C0C0 Then Exit Sub
    Call MakeLog
End Sub

Private Sub optInfo_Click(Index As Integer)
    Dim blnTmp As Boolean
    Dim strValue As String
    If mblnNoSave Then Exit Sub
    If Index = opt初诊 Or Index = opt复诊 Then
        Call UpDate挂号信息("复诊", IIf(optInfo(opt复诊).Value, 1, 0))
    End If
End Sub

Private Sub optInfo_GotFocus(Index As Integer)
    Call SetCurCtlInfo(TypeName(optInfo(Index)), "optInfo", Index)
End Sub

Private Sub PatiAddress_GotFocus(Index As Integer)
    Call SetCurCtlInfo(TypeName(PatiAddress(Index)), "PatiAddress", Index)
End Sub

Private Sub PatiAddress_SetEdit(Index As Integer, blnEdit As Boolean)
    mblnUpdate = blnEdit
    RaiseEvent SetEdit
End Sub
Private Sub picPanel_Click(Index As Integer)
    Call SavePreItem
End Sub

Private Sub picPanel_Paint(Index As Integer)
    Call DrawLine(picPanel_附加)
    Call DrawLine(picPanel_基本信息)
End Sub

Private Sub SetCtlPos(Index As Integer)
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim lngH As Long
    Dim lngW1 As Long, lngW2 As Long
    Dim lngTmp As Long
    Dim lngDXCboLin As Long
    
    On Error Resume Next
 
    If Index = picPanel_基本信息 Then
        lblN(I身份证号).Top = lblN(lbl标题基本).Top + lblN(lbl标题基本).Height + 200
        lngH = 150
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I身份证号), lngH, lblN(I其它证件), lngH, lblN(I出生地点), lngH, lblN(I户口地址), lngH, _
            lblN(I单位名称), lngH, lblN(I家庭地址), lngH, lblN(I婚姻状况), lngH, lblN(I职业), lngH, lblN(I血型), lngH, lblN(I监护人身份证号), lngH, lblLink)
        
        lngW1 = 40: lngW2 = 350
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I身份证号), lngW1, fraC(I身份证号), lngW2, lblN(I区域), lngW1, txtE(I区域), lngW2, _
            lblN(I文化程度), lngW1, fraC(I文化程度))
            
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I监护人身份证号), lngW1, txtE(I监护人身份证号), lngW2 - 100)
        '--------------------------------------------------------------------------
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I身份证号), lngH, txtE(I其它证件), lngH, txtE(I出生地点), lngH, txtE(I户口地址), lngH, _
            txtE(I单位名称), lngH, txtE(I家庭地址), lngH, fraC(I婚姻状况), lngH, fraC(I职业), lngH, txtE(I监护人身份证号), lngH, fraC(I血型))
            
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I身份证号), lngH, txtE(I其它证件), lngH, txtE(I出生地点), lngH, txtE(I户口地址), lngH, _
            txtE(I单位名称), lngH, txtE(I家庭地址), lngH, fraC(I婚姻状况), lngH, fraC(I职业), lngH, txtE(I监护人身份证号), lngH, fraC(I血型))
        
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I区域), lngH, lblN(I籍贯), lngH, lblN(I国籍), lngH, lblN(I民族), lngH, lblN(IRH))
        
        Call zlControl.SetPubCtrlPos(True, -1, txtE(I区域), lngH, txtE(I籍贯), lngH, fraC(I国籍), lngH, fraC(I民族), lngH, fraC(IRH))
        
        Call zlControl.SetPubCtrlPos(True, 1, lblN(I文化程度), lngH, lblN(I生育状况), lngH, lblN(I单位邮编), lngH, lblN(I户口邮编), lngH, _
            lblN(I单位电话), lngH, lblN(I家庭电话), lngH, lblN(I家庭邮编), lngH, lblN(I监护人))
            
        Call zlControl.SetPubCtrlPos(True, -1, fraC(I文化程度), lngH, fraC(I生育状况), lngH, txtE(I单位邮编), lngH, txtE(I户口邮编), lngH, _
             txtE(I单位电话), lngH, txtE(I家庭电话), lngH, txtE(I家庭邮编), lngH, txtE(I监护人))
             
        '----------------------------------------------------------------------------
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I其它证件), lngW1, txtE(I其它证件))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I籍贯), lngW1, txtE(I籍贯))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I生育状况), lngW1, fraC(I生育状况))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I监护人身份证号), lngW1, txtE(I监护人身份证号))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I出生地点), lngW1, txtE(I出生地点))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I单位邮编), lngW1, txtE(I单位邮编))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I户口地址), lngW1, txtE(I户口地址))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I户口邮编), lngW1, txtE(I户口邮编))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I单位名称), lngW1, txtE(I单位名称))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I单位电话), lngW1, txtE(I单位电话))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I家庭地址), lngW1, txtE(I家庭地址))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I家庭电话), lngW1, txtE(I家庭电话))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I婚姻状况), lngW1, fraC(I婚姻状况))
        lblN(I国籍).Top = lblN(I婚姻状况).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I国籍), lngW1, fraC(I国籍))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I家庭邮编), lngW1, txtE(I家庭邮编))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I职业), lngW1, fraC(I职业))
        lblN(I民族).Top = lblN(I职业).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I民族), lngW1, fraC(I民族))
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I监护人), lngW1, txtE(I监护人))
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I血型), lngW1, fraC(I血型))
        lblN(IRH).Top = lblN(I血型).Top
        Call zlControl.SetPubCtrlPos(False, 0, lblN(IRH), lngW1, fraC(IRH))
        
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I区域), lngW1, cmdE(I区域))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I籍贯), lngW1, cmdE(I籍贯))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I出生地点), lngW1, cmdE(I出生地点))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I单位名称), lngW1, cmdE(I单位名称))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I家庭地址), lngW1, cmdE(I家庭地址))
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I户口地址), lngW1, cmdE(I户口地址))
        
        lngW1 = 60
        
        cboE(I身份证号).Left = -30
        cboE(I身份证号).Top = -30
        fraC(I身份证号).Width = cboE(I身份证号).Width
        fraC(I身份证号).Height = cboE(I身份证号).Height - lngW1
        
        cboE(I婚姻状况).Left = -30
        cboE(I婚姻状况).Top = -30
        fraC(I婚姻状况).Width = cboE(I婚姻状况).Width
        fraC(I婚姻状况).Height = cboE(I婚姻状况).Height - lngW1
        
        cboE(I职业).Left = -30
        cboE(I职业).Top = -30
        fraC(I职业).Width = cboE(I职业).Width
        fraC(I职业).Height = cboE(I职业).Height - lngW1
        
        cboE(I血型).Left = -30
        cboE(I血型).Top = -30
        fraC(I血型).Width = cboE(I血型).Width
        fraC(I血型).Height = cboE(I血型).Height - lngW1
        
        cboE(IRH).Left = -30
        cboE(IRH).Top = -30
        fraC(IRH).Width = cboE(IRH).Width
        fraC(IRH).Height = cboE(IRH).Height - lngW1
        
        cboE(I国籍).Left = -30
        cboE(I国籍).Top = -30
        fraC(I国籍).Width = cboE(I国籍).Width
        fraC(I国籍).Height = cboE(I国籍).Height - lngW1
        
        cboE(I民族).Left = -30
        cboE(I民族).Top = -30
        fraC(I民族).Width = cboE(I民族).Width
        fraC(I民族).Height = cboE(I民族).Height - lngW1
        
        cboE(I文化程度).Left = -30
        cboE(I文化程度).Top = -30
        fraC(I文化程度).Width = cboE(I文化程度).Width
        fraC(I文化程度).Height = cboE(I文化程度).Height - lngW1
        
        cboE(I生育状况).Left = -30
        cboE(I生育状况).Top = -30
        fraC(I生育状况).Width = cboE(I生育状况).Width
        fraC(I生育状况).Height = cboE(I生育状况).Height - lngW1
 
        
        txtE(I籍贯).Width = txtE(I区域).Width
        
        txtE(I出生地点).Width = txtE(I籍贯).Width + txtE(I籍贯).Left - txtE(I出生地点).Left
        txtE(I户口地址).Width = txtE(I出生地点).Width
        txtE(I单位名称).Width = txtE(I出生地点).Width
        txtE(I家庭地址).Width = txtE(I出生地点).Width
        
        txtE(I户口邮编).Width = txtE(I单位邮编).Width
        txtE(I单位电话).Width = txtE(I单位邮编).Width
        txtE(I家庭电话).Width = txtE(I单位邮编).Width
        txtE(I家庭邮编).Width = txtE(I单位邮编).Width
        txtE(I监护人).Width = txtE(I单位邮编).Width
        
        cmdE(I区域).Left = txtE(I区域).Left + txtE(I区域).Width - cmdE(I区域).Width
        cmdE(I籍贯).Left = cmdE(I区域).Left
        cmdE(I出生地点).Left = cmdE(I区域).Left
        cmdE(I户口地址).Left = cmdE(I区域).Left
        cmdE(I单位名称).Left = cmdE(I区域).Left
        cmdE(I家庭地址).Left = cmdE(I区域).Left
        
        lblLink.Left = lblN(I身份证号).Left
        
        If mblnStructAdress Then
            PatiAddress(PT_出生地点).Top = txtE(I出生地点).Top: PatiAddress(PT_出生地点).Left = txtE(I出生地点).Left: PatiAddress(PT_出生地点).Height = txtE(I出生地点).Height: PatiAddress(PT_出生地点).Width = txtE(I出生地点).Width
            PatiAddress(PT_户口地址).Top = txtE(I户口地址).Top: PatiAddress(PT_户口地址).Left = txtE(I户口地址).Left: PatiAddress(PT_户口地址).Height = txtE(I户口地址).Height: PatiAddress(PT_户口地址).Width = txtE(I户口地址).Width
            PatiAddress(PT_家庭地址).Top = txtE(I家庭地址).Top: PatiAddress(PT_家庭地址).Left = txtE(I家庭地址).Left: PatiAddress(PT_家庭地址).Height = txtE(I家庭地址).Height: PatiAddress(PT_家庭地址).Width = txtE(I家庭地址).Width
        End If
        lblN(I监护人身份证号).Left = lblN(I家庭地址).Left
        txtE(I监护人身份证号).Left = lblN(I监护人身份证号).Left + lblN(I监护人身份证号).Width + 30
        Call DrawLine(Index)
    ElseIf picPanel_就诊信息 = Index Then
        lngW1 = 40: lngH = 150
        lngW2 = 350
        
        lblN(I发病时间).Top = lblN(lbl标题就诊).Top + lblN(lbl标题就诊).Height + 200
        
        lblN(I发病时间).Left = lblN(I就诊摘要).Left
     
        lblN(I过敏记录).Top = lblN(I发病时间).Top + lblN(I发病时间).Height + 200
        
        lblN(I过敏记录).Left = lblN(I就诊摘要).Left
        
        chkNoAller.Left = lblN(I过敏记录).Left + lblN(I过敏记录).Width + 200
        chkNoAller.Top = lblN(I过敏记录).Top - 20
            
        PicPanel(picPanel_过敏源).Top = lblN(I过敏记录).Top - 20
        PicPanel(picPanel_过敏源).Height = lblN(I过敏记录).Height + 40
        optInfo(opt过敏源).Left = optInfo(opt药品目录).Width
        
        vsAller.Top = lblN(I过敏记录).Top + lblN(I过敏记录).Height + 20
        vsAller.Left = lblN(I过敏记录).Left
        vsAller.Width = txtE(I监护人).Width + txtE(I监护人).Left - lblN(I过敏记录).Left
        
        lblN(I就诊摘要).Top = vsAller.Top + vsAller.Height + 200
        txtE(I就诊摘要).Top = lblN(I就诊摘要).Top + lblN(I就诊摘要).Height + 20
        txtE(I就诊摘要).Left = lblN(I就诊摘要).Left
        txtE(I就诊摘要).Height = IIf(mbytSize = 0, 600, 700)
        txtE(I就诊摘要).Width = txtE(I监护人).Width + txtE(I监护人).Left - txtE(I就诊摘要).Left
      
        cmdSaveZY.Top = txtE(I就诊摘要).Top + txtE(I就诊摘要).Height + 20
        cmdSaveZY.Left = txtE(I就诊摘要).Width + txtE(I就诊摘要).Left - cmdSaveZY.Width
        
        cmdShowZY.Left = cmdSaveZY.Left - cmdShowZY.Width - 10
        cmdShowZY.Top = cmdSaveZY.Top
        
        '----诊断
        lblN(I诊断记录).Left = lblN(I就诊摘要).Left
        lblN(I诊断记录).Top = cmdShowZY.Top + cmdShowZY.Height + 160
        
        lblLinkAdd.Top = lblN(I诊断记录).Top
        lblLinkAdd.Left = chkNoAller.Left + PicPanel(picPanel_诊断).Width
        
        PicPanel(picPanel_诊断).Left = chkNoAller.Left
        PicPanel(picPanel_诊断).Top = lblN(I诊断记录).Top - 20
        PicPanel(picPanel_诊断).Height = lblN(I诊断记录).Height + 40
        optInfo(opt诊断).Left = optInfo(opt疾病).Width
        
        vsDiagXY.Top = lblN(I诊断记录).Top + lblN(I诊断记录).Height + 20
        vsDiagXY.Left = vsAller.Left
        vsDiagXY.Width = txtE(I就诊摘要).Width
        
        cmdDiagMove(0).Move vsDiagXY.Left + vsDiagXY.Width + 70, vsDiagXY.Top + 150, 375, 375
        cmdDiagMove(1).Move vsDiagXY.Left + vsDiagXY.Width + 70, vsDiagXY.Top + cmdDiagMove(0).Height + 250, 375, 375
        
        If mbln中医 Then
            vsDiagZY.Visible = True
            vsDiagZY.Top = vsDiagXY.Top + vsDiagXY.Height + lngH
            vsDiagZY.Left = vsDiagXY.Left
            vsDiagZY.Width = txtE(I就诊摘要).Width
            lngTmp = vsDiagZY.Height
            
            cmdDiagMove(2).Visible = True
            cmdDiagMove(3).Visible = True
            cmdDiagMove(2).Move vsDiagZY.Left + vsDiagZY.Width + 70, vsDiagZY.Top + 80, 375, 375
            cmdDiagMove(3).Move vsDiagZY.Left + vsDiagZY.Width + 70, vsDiagZY.Top + cmdDiagMove(2).Height + 180, 375, 375
            
        Else
            vsDiagZY.Visible = False
            lngTmp = 0
            cmdDiagMove(2).Visible = False
            cmdDiagMove(3).Visible = False
        End If
        
        PicPanel(picPanel_附加).Left = vsDiagXY.Left
        PicPanel(picPanel_附加).Top = vsDiagXY.Height + vsDiagXY.Top + lngTmp + 2 * lngH
        PicPanel(picPanel_附加).Width = txtE(I就诊摘要).Width + 1200
            optInfo(opt复诊).Left = optInfo(opt初诊).Width

        lblN(I医学警示).Top = lblN(I去向).Top + lblN(I去向).Height + 180
        
        lblN(lbl生命体征).Top = lblN(I医学警示).Top + lblN(I医学警示).Height + 180
        
        UCPatiVitalSigns.Top = lblN(lbl生命体征).Top + lblN(lbl生命体征).Height + 100
        
        fraC(I去向).Width = cboE(I去向).Width
        fraC(I去向).Height = cboE(I去向).Height - 60
        
        Call zlControl.SetPubCtrlPos(False, 0, chkInfo, lngW2, lblN(I去向), lngW1, fraC(I去向))
        
        fraC(I日期).Width = cboE(I日期).Width
        fraC(I日期).Height = cboE(I日期).Height - 60
         
        txtE(I发病地址).Width = fraC(I去向).Width
        
        txtE(I其他医学警示).Width = fraC(I去向).Width
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I发病时间), lngW1, txtSL, 0, fraC(I日期), 30, txtE(I发病时间), lngW1, cmdE(I发病时间), lngW2, lblN(I发病地址), lngW1, txtE(I发病地址))
        
        txtE(I发病地址).Left = vsDiagXY.Left + vsDiagXY.Width - txtE(I发病地址).Width
        lblN(I发病地址).Left = txtE(I发病地址).Left - lblN(I发病地址).Width - 10
        
        
        txtE(I发病时间).Width = txtE(I医学警示).Width
        
        cmdE(I发病时间).Left = txtE(I发病时间).Left + txtE(I发病时间).Width
        
        Call zlControl.SetPubCtrlPos(False, 0, lblN(I医学警示), lngW1, txtE(I医学警示), 900, lblN(I其他医学警示), lngW1, txtE(I其他医学警示))
        
        Call zlControl.SetPubCtrlPos(False, 0, txtE(I医学警示), lngW1, cmdE(I医学警示))
        
        cmdE(I医学警示).Left = txtE(I医学警示).Left + txtE(I医学警示).Width
        
        fraC(I去向).Left = txtE(I就诊摘要).Width - fraC(I去向).Width
        lblN(I去向).Left = fraC(I去向).Left - lngW1 - lblN(I去向).Width
        
        lblN(I其他医学警示).Left = lblN(I去向).Left + lblN(I去向).Width - lblN(I其他医学警示).Width
        txtE(I其他医学警示).Left = lblN(I其他医学警示).Left + lblN(I其他医学警示).Width + lngW1
        
        chkInfo.Left = cmdE(I医学警示).Left + cmdE(I医学警示).Width - chkInfo.Width + 130
        
        Call DrawLine(picPanel_附加)
        
        '下划线位置设定
        lngDXCboLin = 60
        
        x1 = txtSL.Left
        y1 = txtSL.Top + txtSL.Height
        x2 = txtSL.Left + txtSL.Width + txtSL.Width
        y2 = y1
        linD(0).x1 = x1
        linD(0).x2 = x2
        linD(0).y1 = y1
        linD(0).y2 = y2
        
        x1 = fraC(I日期).Left
        y1 = fraC(I日期).Top + fraC(I日期).Height
        x2 = fraC(I日期).Left + fraC(I日期).Width - lngDXCboLin
        y2 = y1
        linD(1).x1 = x1
        linD(1).x2 = x2
        linD(1).y1 = y1
        linD(1).y2 = y2
        
        x1 = txtE(I发病时间).Left
        y1 = txtE(I发病时间).Top + txtE(I发病时间).Height
        x2 = txtE(I发病时间).Left + txtE(I发病时间).Width + cmdE(I发病时间).Width
        y2 = y1

        linD(2).x1 = x1
        linD(2).x2 = x2
        linD(2).y1 = y1
        linD(2).y2 = y2
        

        x1 = txtE(I发病地址).Left
        y1 = txtE(I发病地址).Top + txtE(I发病地址).Height
        x2 = txtE(I发病地址).Left + txtE(I发病地址).Width
        y2 = y1

        linD(3).x1 = x1
        linD(3).x2 = x2
        linD(3).y1 = y1
        linD(3).y2 = y2
        
    ElseIf picPanel_快键病历 = Index Then
        '''''
        cmdSign.Left = txtE(I监护人).Width + txtE(I监护人).Left - cmdSign.Width
       
        cmdUpdate.Left = cmdSign.Left - cmdUpdate.Width - 30
        
        cmdImportEPRDemo.Left = cmdUpdate.Left - cmdImportEPRDemo.Width - 30
        
        lngW1 = txtE(I就诊摘要).Width \ 2
        
        For i = 0 To I查体
            rtfEdit(i).Width = lngW1 - 200
        Next
        
        rtfEdit(I现病史).Left = txtE(I监护人).Width + txtE(I监护人).Left - rtfEdit(I现病史).Width
        
        lblDoc(I主诉).Top = 0
        lblDoc(I现病史).Top = 0
        
        lblDoc(I主诉).Left = lblN(I身份证号).Left
        
        rtfEdit(I主诉).Left = lblDoc(I主诉).Left
        rtfEdit(I主诉).Top = lblDoc(I主诉).Top + lblDoc(I主诉).Height + 20
        
        
        lblDoc(I过去史).Top = rtfEdit(I主诉).Top + rtfEdit(I主诉).Height + 100
        lblDoc(I过去史).Left = lblN(I身份证号).Left
        
        rtfEdit(I过去史).Left = lblDoc(I过去史).Left
        rtfEdit(I过去史).Top = lblDoc(I过去史).Top + lblDoc(I过去史).Height + 20
        
        lblDoc(I查体).Top = rtfEdit(I过去史).Top + rtfEdit(I过去史).Height + 100
        lblDoc(I查体).Left = lblN(I身份证号).Left
        
        rtfEdit(I查体).Left = lblDoc(I查体).Left
        rtfEdit(I查体).Top = lblDoc(I查体).Top + lblDoc(I查体).Height + 20
        
        lblDoc(I现病史).Left = rtfEdit(I现病史).Left
        
        rtfEdit(I现病史).Left = lblDoc(I现病史).Left
        rtfEdit(I现病史).Top = lblDoc(I现病史).Top + lblDoc(I现病史).Height + 20
        
        lblDoc(I家族史).Top = lblDoc(I过去史).Top
        lblDoc(I家族史).Left = lblDoc(I现病史).Left
        
        rtfEdit(I家族史).Left = lblDoc(I家族史).Left
        rtfEdit(I家族史).Top = lblDoc(I家族史).Top + lblDoc(I家族史).Height + 20
        
        picPrompt.Top = rtfEdit(I查体).Top + rtfEdit(I查体).Height / 2 - lblTip.Height
        picPrompt.Left = lblDoc(I家族史).Left
        
        lblTip.Top = picPrompt.Top
        lblTip.Left = picPrompt.Left + picPrompt.Width + 20
        
        lngW1 = rtfEdit(I查体).Top + rtfEdit(I查体).Height + 120
        
        linDoc.y1 = lngW1 - 60
        linDoc.y2 = linDoc.y1
        linDoc.x1 = lblDoc(I查体).Left
        linDoc.x2 = rtfEdit(I家族史).Left + rtfEdit(I家族史).Width
        
        cmdSign.Top = lngW1
        cmdUpdate.Top = lngW1
        cmdImportEPRDemo.Top = lngW1
        
        lblEPRname.Left = lblDoc(I查体).Left
        lblEPRname.Top = lngW1 + 20
        
        lblDoctor(0).Left = lblEPRname.Left + lblEPRname.Width + 300
        lblDoctor(0).Top = lblEPRname.Top
        
        lblDoctor(1).Left = lblDoctor(0).Left + lblDoctor(0).Width + 20
        lblDoctor(1).Top = lblEPRname.Top
        
    End If
End Sub

Private Sub DrawLine(ByVal intIdx As Integer)
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim objPic As Object
    Dim lngDXCboLin As Long
    
    If Not (intIdx = picPanel_附加 Or intIdx = picPanel_基本信息) Then
        Exit Sub
    End If
    
    On Error Resume Next
    lngDXCboLin = 60
    Set objPic = PicPanel(intIdx)
    
    objPic.Cls
    If intIdx = picPanel_附加 Then
        
        x1 = txtE(I医学警示).Left
        y1 = txtE(I医学警示).Top + txtE(I医学警示).Height
        x2 = txtE(I医学警示).Left + txtE(I医学警示).Width + cmdE(I医学警示).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I其他医学警示).Left
        y1 = txtE(I其他医学警示).Top + txtE(I其他医学警示).Height
        x2 = txtE(I其他医学警示).Left + txtE(I其他医学警示).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
    
        x1 = fraC(I去向).Left
        y1 = fraC(I去向).Top + fraC(I去向).Height
        x2 = fraC(I去向).Left + fraC(I去向).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        '--------画生命体征的边框 4 条线-------------
        If mint调用 = 1 Then
            x1 = 0
            y1 = UCPatiVitalSigns.Top - 80
            x2 = txtE(I其他医学警示).Left + txtE(I其他医学警示).Width
            y2 = y1
            objPic.Line (x1, y1)-(x2, y2)
        
            x1 = 0
            y1 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            x2 = txtE(I其他医学警示).Left + txtE(I其他医学警示).Width
            y2 = y1
            objPic.Line (x1, y1)-(x2, y2)
            
            x1 = 0
            y1 = UCPatiVitalSigns.Top - 80
            x2 = 0
            y2 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            objPic.Line (x1, y1)-(x2, y2)
            
            x1 = txtE(I其他医学警示).Left + txtE(I其他医学警示).Width
            y1 = UCPatiVitalSigns.Top - 80
            x2 = x1
            y2 = UCPatiVitalSigns.Top + UCPatiVitalSigns.Height + 20
            objPic.Line (x1, y1)-(x2, y2)
        End If
        '-------------------------
        
    ElseIf intIdx = picPanel_基本信息 Then
        x1 = txtE(I区域).Left
        y1 = txtE(I区域).Top + txtE(I区域).Height
        x2 = txtE(I区域).Left + txtE(I区域).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I其它证件).Left
        y1 = txtE(I其它证件).Top + txtE(I其它证件).Height
        x2 = txtE(I其它证件).Left + txtE(I其它证件).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I监护人身份证号).Left
        y1 = txtE(I监护人身份证号).Top + txtE(I监护人身份证号).Height
        x2 = txtE(I监护人身份证号).Left + txtE(I监护人身份证号).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I籍贯).Left
        y1 = txtE(I籍贯).Top + txtE(I籍贯).Height
        x2 = txtE(I籍贯).Left + txtE(I籍贯).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I单位邮编).Left
        y1 = txtE(I单位邮编).Top + txtE(I单位邮编).Height
        x2 = txtE(I单位邮编).Left + txtE(I单位邮编).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I户口邮编).Left
        y1 = txtE(I户口邮编).Top + txtE(I户口邮编).Height
        x2 = txtE(I户口邮编).Left + txtE(I户口邮编).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I家庭邮编).Left
        y1 = txtE(I家庭邮编).Top + txtE(I家庭邮编).Height
        x2 = txtE(I家庭邮编).Left + txtE(I家庭邮编).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I家庭电话).Left
        y1 = txtE(I家庭电话).Top + txtE(I家庭电话).Height
        x2 = txtE(I家庭电话).Left + txtE(I家庭电话).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I单位电话).Left
        y1 = txtE(I单位电话).Top + txtE(I单位电话).Height
        x2 = txtE(I单位电话).Left + txtE(I单位电话).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I监护人).Left
        y1 = txtE(I监护人).Top + txtE(I监护人).Height
        x2 = txtE(I监护人).Left + txtE(I监护人).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = txtE(I单位名称).Left
        y1 = txtE(I单位名称).Top + txtE(I单位名称).Height
        x2 = txtE(I单位名称).Left + txtE(I单位名称).Width
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        If Not mblnStructAdress Then
           x1 = txtE(I出生地点).Left
           y1 = txtE(I出生地点).Top + txtE(I出生地点).Height
           x2 = txtE(I出生地点).Left + txtE(I出生地点).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
           
           x1 = txtE(I家庭地址).Left
           y1 = txtE(I家庭地址).Top + txtE(I家庭地址).Height
           x2 = txtE(I家庭地址).Left + txtE(I家庭地址).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
           
           x1 = txtE(I户口地址).Left
           y1 = txtE(I户口地址).Top + txtE(I户口地址).Height
           x2 = txtE(I户口地址).Left + txtE(I户口地址).Width
           y2 = y1
           objPic.Line (x1, y1)-(x2, y2)
        End If
        
        x1 = fraC(I身份证号).Left
        y1 = fraC(I身份证号).Top + fraC(I身份证号).Height
        x2 = fraC(I身份证号).Left + fraC(I身份证号).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I文化程度).Left
        y1 = fraC(I文化程度).Top + fraC(I文化程度).Height
        x2 = fraC(I文化程度).Left + fraC(I文化程度).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I文化程度).Left
        y1 = fraC(I文化程度).Top + fraC(I文化程度).Height
        x2 = fraC(I文化程度).Left + fraC(I文化程度).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(IRH).Left
        y1 = fraC(IRH).Top + fraC(IRH).Height
        x2 = fraC(IRH).Left + fraC(IRH).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I血型).Left
        y1 = fraC(I血型).Top + fraC(I血型).Height
        x2 = fraC(I血型).Left + fraC(I血型).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I职业).Left
        y1 = fraC(I职业).Top + fraC(I职业).Height
        x2 = fraC(I职业).Left + fraC(I职业).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I婚姻状况).Left
        y1 = fraC(I婚姻状况).Top + fraC(I婚姻状况).Height
        x2 = fraC(I婚姻状况).Left + fraC(I婚姻状况).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        x1 = fraC(I国籍).Left
        y1 = fraC(I国籍).Top + fraC(I国籍).Height
        x2 = fraC(I国籍).Left + fraC(I国籍).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        
        x1 = fraC(I民族).Left
        y1 = fraC(I民族).Top + fraC(I民族).Height
        x2 = fraC(I民族).Left + fraC(I民族).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
        
        
        x1 = fraC(I生育状况).Left
        y1 = fraC(I生育状况).Top + fraC(I生育状况).Height
        x2 = fraC(I生育状况).Left + fraC(I生育状况).Width - lngDXCboLin
        y2 = y1
        objPic.Line (x1, y1)-(x2, y2)
    End If
End Sub

Private Sub SetRTFEditFontSize()
'功能：设置病人病史信息的输入框字体
    Dim i As Long
    lblEPRname.Visible = True
    mblnSizeTmp = True
    For i = 0 To rtfEdit.UBound
        Call zlControl.RTFSetFontSize(rtfEdit(i), IIf(mbytSize = 0, 9, 12))
    Next
    mblnSizeTmp = False
End Sub

Private Sub lblDoctor_Click(Index As Integer)
    lblEPRname.Visible = Not lblEPRname.Visible
End Sub

Private Sub rtfEdit_Change(Index As Integer)
    If mblnSizeTmp = True Then Exit Sub
    mblnChange = True
End Sub

Private Sub rtfEdit_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True)
    Call SetCurCtlInfo(TypeName(rtfEdit(Index)), "rtfEdit", Index)
End Sub

Private Sub rtfEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not mblnPatiChange Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '连续两次回车光标跳转
            With rtfEdit(Index)
                If Trim(.Text) = "" Then
                    KeyAscii = 0
                    Call zlCommFun.PressKey(vbKeyTab)
                ElseIf .SelStart - 1 > 0 Then
                    If Mid(.Text, .SelStart - 1, 2) = vbCrLf Then
                        KeyAscii = 0
                        Call zlCommFun.PressKey(vbKeyBack)
                        Call zlCommFun.PressKey(vbKeyTab)
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub rtfEdit_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme
End Sub

Private Sub rtfEdit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp As String
    mrsMainInfo.Filter = "控件名='rtfEdit' and Index=" & Index
    strTmp = mrsMainInfo!ErrInfo & ""
    Call zlCommFun.ShowTipInfo(rtfEdit(Index).hwnd, strTmp, True, True)
End Sub

Private Sub rtfEdit_SelChange(Index As Integer)
    With rtfEdit(Index)
        If .SelLength = 0 And .SelStart > 0 And PicPanel(picPanel_快键病历).Tag = "" Then
            If Mid(.Text, .SelStart, 1) = "`" Or Mid(.Text, .SelStart, 1) = "·" Then
                PicPanel(picPanel_快键病历).Tag = "UnChange"
                .SelStart = .SelStart - 1
                .SelLength = 1
                .SelText = ""
                Call ShowWordInput(rtfEdit(Index))
                PicPanel(picPanel_快键病历).Tag = ""
            End If
        End If
    End With
End Sub

Private Sub rtfEdit_Validate(Index As Integer, Cancel As Boolean)
    Call UpDate病历(Index)
End Sub

Private Sub txtE_Change(Index As Integer)
    Dim txtTmp As Object
    Dim lngPos As Long, lngLen As Long
    If Index = I监护人身份证号 Then
        If mblnReturn Then Exit Sub
        Set txtTmp = txtE(Index)
        mblnReturn = True
        '不规则的输入
        If Not zlStr.CheckCharScope(txtTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
            txtTmp.Text = ""
        Else
            If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                If zlCommFun.ActualLen(txtTmp.Text) > 18 Then
                    txtTmp.Text = Mid(txtTmp.Text, 1, 18)
                End If
            End If
        End If
        If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
            lngPos = InStr(txtTmp.Text, "*")
            lngLen = Len(Mid(txtTmp.Text, 13, 2))
            Select Case lngPos
                Case 0
                    txtTmp.Tag = txtTmp.Text
                Case Else
                    txtTmp.Tag = Mid(txtTmp.Text, 1, lngPos - 1)
                    txtTmp.Text = txtTmp.Tag
                    txtTmp.SelStart = Len(txtTmp.Text)
            End Select
        Else
            txtTmp.Tag = txtTmp.Text
        End If
        mblnReturn = False
    End If
End Sub

Private Sub txtE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(txtE(Index).hwnd, txtE(Index).Tag, True, True)
End Sub

Private Sub txtE_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String, str身份证号 As String
    If Not txtE(Index).Locked Then
        If Index = I发病时间 Then
            txtE(Index).Text = GetFullDate(txtE(Index).Text)
        End If
        If Index = I发病时间 Then
            txtSL.Text = ""
            cboE(I日期).ListIndex = -1
        End If
        If Index = I监护人身份证号 Then
            '完整的身份证号是存在不加掩码  cboE(Index).Tag
            strValue = txtE(Index).Tag
            mrsMainInfo.Filter = "信息名='身份证号'"
            str身份证号 = mrsMainInfo!信息原值 & ""
            If strValue <> str身份证号 Then
                If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                    If Not Check身份证号(strValue, txtE(Index)) Then
                        Cancel = True
                        txtE(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If
            Call Update监护人身份证
        Else
            Call UpDateInfo(txtE(Index).Text, "txtE", Index)
        End If
    End If
End Sub

Private Sub txtSentence_GotFocus()
    Call zlCommFun.OpenIme(True)
    Call zlControl.TxtSelAll(txtSentence)
End Sub

Private Sub txtSentence_KeyPress(KeyAscii As Integer)
    Dim strSentence As String, blnCancel As Boolean, strType As String
       
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        Select Case Val(picSentence.Tag)
        Case I主诉
            strType = "病人主诉"
        Case I家族史
            strType = "家族史"
        Case I现病史
            strType = "现病史"
        Case I查体
            strType = "体格一般检查"
        Case I过去史
            strType = "既往史"
        End Select
                
        strSentence = frmSentenceSel.ShowMe(Me, mlng病历文件id, mstr性别, mstr婚姻状况, strType, txtSentence.Text, picSentence.hwnd, blnCancel)
        If strSentence <> "" Then
            rtfEdit(Val(picSentence.Tag)).SelText = strSentence
            Call HideWordInput
        Else
            If Not blnCancel Then
                MsgBox "没有找到匹配的词句。", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtSentence)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call imgSentence_Click
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        Call HideWordInput
    End If
End Sub


Private Sub imgSentence_Click()
    Dim strSentence As String, strType As String
    
    Select Case Val(picSentence.Tag)
    Case I主诉
        strType = "病人主诉"
    Case I家族史
        strType = "家族史"
    Case I现病史
        strType = "现病史"
    Case I查体
        strType = "体格一般检查"
    Case I过去史
        strType = "既往史"
    End Select
    
    strSentence = frmSentenceSel.ShowMe(Me, mlng病历文件id, mstr性别, mstr婚姻状况, strType)
    If strSentence <> "" Then
        rtfEdit(Val(picSentence.Tag)).SelText = strSentence
        Call HideWordInput
    End If
End Sub

Private Sub txtSentence_LostFocus()
    If Not frmSentenceSel.mblnShow Then
        Call HideWordInput   '隐藏词句输入
    End If
End Sub

Private Sub ShowWordInput(ByRef txtThis As RichTextBox)
'功能：显示词句输入
    Dim vPos As POINTAPI
    
    If txtThis.Visible And txtThis.Enabled And Not txtThis.Locked Then
        picSentence.Tag = txtThis.Index '记下以便隐藏返回后定位
        
        If txtThis.Text = "" Then PicPanel(picPanel_快键病历).Tag = "UnChange": txtThis.Text = " " '必须要有一个空字符才能返回其坐标
        vPos = zlControl.GetCaretPos(txtThis.hwnd)
        If txtThis.Text = " " Then PicPanel(picPanel_快键病历).Tag = "UnChange": txtThis.Text = ""
        
        If vPos.X <> -1 And vPos.Y <> -1 Then
            If txtThis.Left + vPos.X + Screen.TwipsPerPixelX * 2 < txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX Then
                picSentence.Left = txtThis.Left + vPos.X + Screen.TwipsPerPixelX * 2
            Else
                picSentence.Left = txtThis.Left + txtThis.Width - picSentence.Width - 2 * Screen.TwipsPerPixelX
            End If
            picSentence.Top = txtThis.Top + vPos.Y + Screen.TwipsPerPixelY
            txtSentence.Text = ""
            picSentence.Visible = True
            txtSentence.SetFocus
        End If
    End If
End Sub

Private Sub HideWordInput()
'功能：隐藏词句输入
    Dim idx As Long
    
    If picSentence.Visible Then
        picSentence.Visible = False
        txtSentence.Text = ""
        
        idx = Val(picSentence.Tag)
        picSentence.Tag = ""
        
        If rtfEdit(idx).Visible And rtfEdit(idx).Enabled And Not rtfEdit(idx).Locked Then
            rtfEdit(idx).SetFocus
        End If
    End If
End Sub

Private Sub LoadDocData()
    Dim rsTmp As ADODB.Recordset, strSQL As String, strContent As String
    Dim i As Long, j As Long, arrTmp  As Variant, blnLoading As Boolean
    
    mblnNoSave = True
    For i = 0 To rtfEdit.UBound
        rtfEdit(i).Text = ""
        rtfEdit(i).Tag = ""
    Next
    lblEPRname.Caption = ""    '仅用于技术人员查原因，例如：选择提纲词句时，没有列出预计的词句，可根据病历文件名称查是否设置了提纲词句对应
    
    '只显示简单病历模式下产生的文件
    strSQL = "Select id,文件id,签名级别,病历名称,保存人 From 电子病历记录 A Where 病人id = [1] And 主页id = [2] And 病历种类 = 1" & vbNewLine & _
        " And Exists(Select 1 From 病历文件列表 B Where A.文件ID = B.ID And B.保留 = '3')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    If rsTmp.RecordCount > 0 Then
        If SetCompendsTag(Val("" & rsTmp!文件ID)) Then
            mlng病历文件id = Val("" & rsTmp!文件ID)
            mbln签名 = IIf(Val("" & rsTmp!签名级别) > 0, True, False)
            mlng病历ID = rsTmp!ID
            mstr保存人 = "" & rsTmp!保存人
            lblEPRname.Caption = "" & rsTmp!病历名称
            '读取提纲下的段落文本,对象属性为-1表示提纲标题文本
            strSQL = "Select A.预制提纲id, B.内容文本, B.ID" & vbNewLine & _
                    "From 电子病历内容 A, 电子病历内容 B" & vbNewLine & _
                    "Where A.文件id = [1] And A.对象类型 = 1 And A.预制提纲id+0 In(-10,5,2,6,3)" & vbNewLine & _
                    "      And B.父id = A.ID And B.对象类型 = 2 Order By A.预制提纲id, B.对象序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病历ID)
            With rsTmp
                If .RecordCount > 0 Then
                    arrTmp = Split("-10,2,3,5,6", ",") '病人主诉,现病史,既往史,家族史,体格检查
                    For i = 0 To UBound(arrTmp)
                        .Filter = "预制提纲id=" & arrTmp(i)
                        For j = 1 To .RecordCount
                            If j = 1 Then
                                strContent = "" & !内容文本
                                If InStr(strContent, lblDoc(i).Tag) = 1 Then strContent = Mid(strContent, Len(lblDoc(i).Tag) + 1)
                                rtfEdit(i).Text = strContent
                                rtfEdit(i).Tag = !ID
                            Else
                                rtfEdit(i).Text = rtfEdit(i).Text & vbCrLf & !内容文本
                                rtfEdit(i).Tag = rtfEdit(i).Tag & "," & !ID
                            End If
                            .MoveNext
                        Next
                    Next
                End If
            End With
        End If
    Else
        If mbln急 Then
            strSQL = " And (R.事件 = '急诊'  OR R.事件 IS NUll)"
        Else
            If optInfo(opt复诊).Value Then
                strSQL = " And (R.事件 = '门诊' Or R.事件 = '复诊'  OR R.事件 IS NUll)"
            Else
                strSQL = " And (R.事件 = '门诊' Or R.事件 = '初诊'  OR R.事件 IS NUll )"
            End If
        End If
        '系统定义了门(急)诊病历且对当前病人适用，具有5个固定预制提纲,才显示病历录入面板.
        strSQL = "Select F.ID, F.名称 as 病历名称" & vbNewLine & _
                "From (Select F.ID, F.通用, A.科室id, F.名称,Decode(R.事件,Null,2,1) 事件" & vbNewLine & _
                "       From 病历文件列表 F, 病历应用科室 A, 病历时限要求 R" & vbNewLine & _
                "       Where F.ID = A.文件id(+) And F.ID = R.文件id(+) And F.种类 = 1 And F.保留= '3'" & strSQL & ") F" & vbNewLine & _
                "Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [2]" & vbNewLine & _
                "Order By F.事件,F.通用 Desc,F.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng挂号ID, mlng科室ID)
        If rsTmp.RecordCount > 0 Then
            mlng病历文件id = rsTmp!ID
            lblEPRname.Caption = "" & rsTmp!病历名称
            If SetCompendsTag(mlng病历文件id) = False Then
                mlng病历文件id = 0: lblEPRname.Caption = ""
            End If
        Else
            mlng病历文件id = 0: lblEPRname.Caption = ""
        End If
        
        mlng病历ID = 0
        mbln签名 = False
    End If
    
    mrsMainInfo.Filter = "控件名='rtfEdit'"
    
    For i = 1 To mrsMainInfo.RecordCount
        mrsMainInfo!信息原值 = rtfEdit(mrsMainInfo!Index).Text
        mrsMainInfo.Update
        mrsMainInfo.MoveNext
    Next
    mblnNoSave = False
    cmdImportEPRDemo.Visible = mlng病历文件id <> 0
    Call SetDocEditable
    Call SetRTFEditFontSize
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCompendsTag(ByVal lng病历文件id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    
    strSQL = "Select Decode(A.预制提纲id, -10, 0, 5, 3, 2, 1, 6, 4, 3, 2) As 序号, B.内容文本" & vbNewLine & _
            "From 病历文件结构 A, 病历文件结构 B" & vbNewLine & _
            "Where A.文件id = [1] And A.预制提纲id+0 In (-10,5,2,6,3) And A.Id = B.父id And B.对象类型 = 2" & vbNewLine & _
            "Order By 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病历文件id)
    If rsTmp.RecordCount > 0 And rsTmp.RecordCount <= 5 Then
        If rsTmp!序号 & "" = "0" Then  '必须包含主诉
            For i = 0 To rsTmp.RecordCount - 1
                lblDoc(Val(rsTmp!序号 & "")).Tag = rsTmp!内容文本       '用于保存Rtf文件替换内容时定位
                rsTmp.MoveNext
            Next
            For i = 1 To lblDoc.Count - 1
                If lblDoc(i).Tag = "" Then
                    lblDoc(i).Visible = False
                    rtfEdit(i).Visible = False
                Else
                    lblDoc(i).Visible = True
                    rtfEdit(i).Visible = True
                End If
            Next
            SetCompendsTag = True
        End If
    End If
End Function

Private Function CanUseFastEPR() As Boolean
'功能：是否有可用的快捷病历文件
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    blnTmp = InStr(GetInsidePrivs(p门诊病历管理), "病历书写") > 0
    
    If Not blnTmp Then Exit Function
    
    If mbln急 Then
        strSQL = " And (R.事件 = '急诊'  OR R.事件 IS NUll)"
    Else
        If optInfo(opt复诊).Value Then
            strSQL = " And (R.事件 = '门诊' Or R.事件 = '复诊'  OR R.事件 IS NUll)"
        Else
            strSQL = " And (R.事件 = '门诊' Or R.事件 = '初诊'  OR R.事件 IS NUll )"
        End If
    End If
    '系统定义了门(急)诊病历且对当前病人适用，具有5个固定预制提纲,才显示病历录入面板.
    strSQL = "Select F.ID, F.名称 as 病历名称" & vbNewLine & _
            "From (Select F.ID, F.通用, A.科室id, F.名称,Decode(R.事件,Null,2,1) 事件" & vbNewLine & _
            "       From 病历文件列表 F, 病历应用科室 A, 病历时限要求 R" & vbNewLine & _
            "       Where F.ID = A.文件id(+) And F.ID = R.文件id(+) And F.种类 = 1 And F.保留= '3'" & strSQL & ") F" & vbNewLine & _
            "Where F.通用 = 1 Or F.通用 = 2 And F.科室id = [2]" & vbNewLine & _
            "Order By F.事件,F.通用 Desc,F.id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng挂号ID, mlng科室ID)
    If rsTmp.RecordCount > 0 Then
        mlng病历文件id = rsTmp!ID
        lblEPRname.Caption = "" & rsTmp!病历名称
        If SetCompendsTag(mlng病历文件id) Then
            CanUseFastEPR = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NeedName(strList As String) As String
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）> - > [] > ()

    '优先判断以回车符分割
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '以[]分割
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '以()分割
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '以-分割
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function

Private Sub GetCboIndex(objCbo As Object, ByVal strFind As String)
'功能：由字符串在ComboBox中查找索引
'参数：blnKeep=如果未匹配，是否保持原索引
    Dim i As Integer

    '先精确查找
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind And strFind <> "" Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '最后模糊查找
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 And strFind <> "" Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
End Sub

Private Sub SetCtlBackColor()
    Dim objCtl As Object
    For Each objCtl In Me.Controls
        If UCase(TypeName(objCtl)) <> "LINE" And InStr(",LINE,VSCROLLBAR,IMAGE,STATUSBAR,", "," & UCase(TypeName(objCtl)) & ",") = 0 Then
        objCtl.BackColor = vbWhite ' &HFFFFFF
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsc.Value
    lngMin = vsc.Min
    lngMax = vsc.Max
    
    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsc.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsc.Value = lngMin
        End If
    End If
    
End Sub

Private Sub Form_Activate()
'鼠标滚轮
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    If Not mobjCtl Is Nothing Then
        mobjCtl.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
'鼠标滚轮
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If Me.ActiveControl.Name = "dtpDate" Then
            dtpDate.Visible = False
        End If
    ElseIf 39 = KeyAscii Then '单引号
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strHead As String
    Dim blnTmp As Boolean
    Dim blnHave As Boolean
    Dim strTmp As String
    Dim intTmp As Integer
    
    Call RestoreWinState(Me, App.ProductName)
    
    Set mobjKernel = New zlPublicAdvice.clsPublicAdvice
    
    mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
    mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
    mblnUpdate = True

    Call InitBaseInfo
    
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    mblnID加密 = Val(zlDatabase.GetPara(247, glngSys, , 0)) = 1
    mblnFreeInput = Val(zlDatabase.GetPara("诊断手术名称自由调整", glngSys, 0)) = 1
    mbln录中医诊断 = Val(zlDatabase.GetPara("门诊西医科允许录入中医诊断", glngSys, p门诊医嘱下达, "0")) = 1

    '诊断输入方式
    strTmp = gstr诊断输入
    If Val(Mid(strTmp, 1, 1)) = 0 Then
        strTmp = "1"
    Else
        strTmp = Mid(strTmp, 1, 1)
    End If
    mint诊断输入 = Val(strTmp)
    
    mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")
    mblnEdit合同单位 = InStr(GetInsidePrivs(p门诊医生站), "合约病人登记") > 0
    '先读参数，菜单定义中需要判断
    blnTmp = InStr(GetInsidePrivs(p门诊病历管理), "病历书写") > 0
    If blnTmp Then
        lblDoctor(1).Tag = IIf(InStr(GetInsidePrivs(p门诊病历管理), "他人病历") > 0, 1, 0)
        mblnDocInput = Val(zlDatabase.GetPara("显示病历快捷输入", glngSys, p门诊医生站, 0)) = 1
        blnHave = IIf(InStr(GetInsidePrivs(1070), "签名权") > 0, True, False)
        lblDoctor(1).Caption = UserInfo.姓名
        lblLink.Visible = True
    Else
        mblnDocInput = False
        blnHave = False
        lblLink.Visible = False
    End If
    
    '初始化地址控件
    PatiAddress(PT_出生地点).Visible = mblnStructAdress
    PatiAddress(PT_户口地址).Visible = mblnStructAdress
    PatiAddress(PT_家庭地址).Visible = mblnStructAdress
    txtE(I出生地点).Visible = Not mblnStructAdress: cmdE(I出生地点).Visible = Not mblnStructAdress
    txtE(I户口地址).Visible = Not mblnStructAdress: cmdE(I户口地址).Visible = Not mblnStructAdress
    txtE(I家庭地址).Visible = Not mblnStructAdress: cmdE(I家庭地址).Visible = Not mblnStructAdress
    If mblnStructAdress Then
        PatiAddress(PT_出生地点).ShowTown = False
        PatiAddress(PT_户口地址).ShowTown = mblnShowTown
        PatiAddress(PT_家庭地址).ShowTown = mblnShowTown
    End If
 
    cmdSign.Visible = blnHave
    lblDoctor(0).Visible = blnHave
    lblDoctor(1).Visible = blnHave
    PicPanel(picPanel_快键病历).Visible = mblnDocInput

    intTmp = Val(zlDatabase.GetPara("门诊诊断输入", glngSys, p门诊医生站, 0, Array(optInfo(opt疾病), optInfo(opt诊断))))
    Call SetInputRoot(0, gint诊断来源, intTmp, optInfo(opt诊断), optInfo(opt疾病))
    
    intTmp = Val(zlDatabase.GetPara("过敏输入来源", glngSys, p门诊医生站, 0, Array(optInfo(opt药品目录), optInfo(opt过敏源))))
    If Not gobjPass Is Nothing Then
        Call SetInputRoot(2, gint过敏输入来源, intTmp, optInfo(opt药品目录), optInfo(opt过敏源))
    Else
        Call SetInputRoot(intTmp, 1, intTmp, optInfo(opt药品目录), optInfo(opt过敏源))
    End If
    
    strHead = "过敏物,3100,1;过敏反应,3800,1;过敏时间,1100,4;过敏源编码;药物ID;过敏来源"
    Call InitVSFlexGrid(vsAller, strHead)
    
    strHead = ",460,4;关联;诊断编码,840,4;诊断描述,4000,1;中医证候;发病时间,1600,1;备注;ICD附码;疑诊,460,4;,270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
    Call InitVSFlexGrid(vsDiagXY, strHead, "0,西医,18,1", 1, 1)
    
    strHead = ",460,4;关联;诊断编码,840,4;诊断描述,2700,1;中医证候,1300,1;发病时间,1600,1;备注;ICD附码;疑诊,460,4;,270,4;,270,4;诊断ID;疾病ID;证候ID;医嘱IDs;诊断分类;固定附码;附码ID;诊断来源;疾病编码;疾病类别;证候编码;记录日期;记录人员"
    Call InitVSFlexGrid(vsDiagZY, strHead, "0,中医,18,11", 1, 0)
    
    Call InitEditData
    
    vsc.Max = 600
    vsc.Min = 0
    vsc.LargeChange = 100
    Call SetCtlBackColor
    
    Call SetCtlPos(0)
    Call SetCtlPos(1)
    Call SetCtlPos(2)
    If mint调用 = 1 Then
        stbThis.Visible = True
    Else
        stbThis.Visible = False
    End If
    Set mobjCtl = Nothing
    mblnCboNoClick = False
    mblnOK = False
End Sub

Private Sub LoadDiagData()
'功能：加载诊断信息
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    Set rsTmp = GetPatiDiagData(False)
    
    '删除之前的缓存
    mrsSecdInfo.Filter = "控件名='vsDiagXY' or 控件名='vsDiagZY'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If
    Call LoadVsDiagData(vsDiagXY, rsTmp, "1")
    Call LoadVsDiagData(vsDiagZY, rsTmp, "11")
End Sub

Private Function FilterDiagByType(ByRef rsInput As ADODB.Recordset, ByVal intDiagType As Integer) As Boolean
'功能：诊断的过滤
'返回：true-首页诊断
    rsInput.Filter = "记录来源=3 And 诊断类型=" & intDiagType
    FilterDiagByType = Not rsInput.EOF
    If rsInput.EOF Then rsInput.Filter = "记录来源=2 And 诊断类型=" & intDiagType
    If rsInput.EOF Then rsInput.Filter = "记录来源=1 And 诊断类型=" & intDiagType
    If rsInput.EOF Then rsInput.Filter = "记录来源=4 And 诊断类型=" & intDiagType
End Function

Private Sub LoadVsDiagData(ByRef vsDiagInput As VSFlexGrid, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal strDiagType As String)
'功能：将诊断加载到表格中并且缓存
'参数：vsDiagInput=需要加载诊断的表格
'      rsInput=读取的诊断记录集
'      strDiagType=诊断类型字符串，各类型以逗号分割
'      blnOnlyCache=是否只缓存数据，True-界面经过检查后缓存，False-界面初始加载缓存
'说明：LoadMedPageData的子函数

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long
    Dim bln分化程度 As Boolean
    Dim bln西医 As Boolean
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim lngTmp As Long
    
    On Error GoTo errH
    With vsDiagInput
        bln西医 = vsDiagInput.Name = "vsDiagXY"
        .Rows = .FixedRows
        
        .Rows = .FixedRows + 1
        .TextMatrix(.Rows - 1, 0) = IIf(bln西医, "西医", "中医")
        .TextMatrix(.Rows - 1, DI_诊断分类) = IIf(bln西医, 1, 11)
        If Not FilterDiagByType(rsInput, Val(strDiagType)) Then
            .Tag = "1"
        Else
            .Tag = ""
        End If
        If bln西医 Then
            mstrTagDiagXY = .Tag
        Else
            mstrTagDiagZY = .Tag
        End If
        .Tag = ""
        
        Do While Not rsInput.EOF
            '确定当前显示行
            lngRow = .FindRow(strDiagType, , DI_诊断分类, , True)
            For j = lngRow To .Rows - 1
                If Val(.TextMatrix(j, DI_诊断分类)) = Val(strDiagType) Then
                    lngRow = j
                    If .TextMatrix(j, DI_诊断描述) = "" Then Exit For
                Else
                    Exit For
                End If
            Next
            '新增行
            If .TextMatrix(lngRow, DI_诊断描述) <> "" Then
                lngRow = lngRow + 1: .AddItem "", lngRow
                .TextMatrix(lngRow, DI_诊断分类) = strDiagType
            End If
            
            strTmp = rsInput!诊断描述 & ""
            If Not (IsNull(rsInput!诊断id) And IsNull(rsInput!疾病id)) Then
                '读取诊断编码，诊断描述为(编码)描述，或(编码)描述(证候) 类型的可以获取诊断描述
                If strTmp Like "(?*)?*" Then
                    lngPos = InStr(1, strTmp, ")")
                    .TextMatrix(lngRow, DI_诊断编码) = Mid(strTmp, 2, lngPos - 2)
                    strTmp = Mid(strTmp, lngPos + 1)
                End If
            End If
            If .TextMatrix(lngRow, DI_诊断编码) = "" And Not (IsNull(rsInput!诊断id) And IsNull(rsInput!疾病id)) Then
                '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                .TextMatrix(lngRow, DI_诊断编码) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病编码 & "", rsInput!诊断编码 & "")
            End If
            '获取中医证候，由于诊断描述可能会增加前后缀，前后缀包含括号，所以反向截取字符串
            If strTmp Like "?*(?*)" And Not bln西医 Then
                strTmp = StrReverse(strTmp)
                lngPos = InStr(1, strTmp, "(")
                .TextMatrix(lngRow, DI_中医证候) = StrReverse(Mid(strTmp, 2, lngPos - 2))
                strTmp = StrReverse(Mid(strTmp, lngPos + 1))
            End If
            '取诊断描述
            .TextMatrix(lngRow, DI_诊断描述) = strTmp
            '诊断描述的备份数据
            If Not (IsNull(rsInput!诊断id) And IsNull(rsInput!疾病id)) Then
                .Cell(flexcpData, lngRow, DI_诊断描述) = IIf(Not IsNull(rsInput!疾病id), rsInput!疾病名称 & "", rsInput!诊断名称 & "")
            Else
                .Cell(flexcpData, lngRow, DI_诊断描述) = .TextMatrix(lngRow, DI_诊断描述)
            End If
             .Cell(flexcpData, lngRow, DI_诊断编码) = .TextMatrix(lngRow, DI_诊断编码)
             .Cell(flexcpData, lngRow, DI_中医证候) = .TextMatrix(lngRow, DI_中医证候)
            '其他列数据加载
            .TextMatrix(lngRow, DI_发病时间) = Format(rsInput!发病时间 & "", "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, DI_备注) = rsInput!备注 & ""
            .TextMatrix(lngRow, DI_ICD附码) = rsInput!附码 & ""
            .TextMatrix(lngRow, DI_是否疑诊) = IIf(Val(rsInput!是否疑诊 & "") = 1, "？", "")
            .TextMatrix(lngRow, DI_诊断ID) = rsInput!诊断id & ""
            .TextMatrix(lngRow, DI_疾病ID) = rsInput!疾病id & ""
            .TextMatrix(lngRow, DI_证候ID) = rsInput!证候id & ""
            .TextMatrix(lngRow, DI_医嘱IDs) = rsInput!医嘱ID & ""
            .TextMatrix(lngRow, DI_固定附码) = IIf(IsNull(rsInput!附码), "", "1")
            .TextMatrix(lngRow, DI_附码ID) = IIf(IsNull(rsInput!分娩), "0", "1")
            .TextMatrix(lngRow, DI_诊断来源) = Val(rsInput!记录来源 & "") '保存记录来源，以便保存时，保存为首页或病案来源
            .TextMatrix(lngRow, DI_疾病编码) = rsInput!疾病编码 & ""
            .TextMatrix(lngRow, DI_疾病类别) = rsInput!疾病类别 & ""
            .TextMatrix(lngRow, DI_证候编码) = rsInput!证候编码 & ""
            .TextMatrix(lngRow, DI_记录日期) = Format(rsInput!记录日期 & "", "YYYY-MM-DD HH:mm")
            .TextMatrix(lngRow, DI_记录人员) = rsInput!记录人 & ""
            .RowData(lngRow) = Val(rsInput!ID & "")
            rsInput.MoveNext
        Loop
   
        '设置诊断相关信息
        .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
        .Cell(flexcpText, .FixedRows, DI_诊断类型, .Rows - 1, DI_诊断类型) = IIf(bln西医, "西医", "中医")
        
        '数据缓存
        lngTmp = 1
        strTmp = ""
        arrMain = Array(DI_诊断编码, DI_诊断分类, DI_诊断ID, DI_疾病ID)
        arrWhole = Array(DI_诊断分类, DI_疾病编码, DI_诊断编码, DI_ICD附码, DI_疾病类别, DI_证候编码, DI_中医证候, DI_是否疑诊, DI_诊断ID, DI_疾病ID, DI_诊断描述, DI_备注, DI_发病时间)
        For i = .FixedRows To .Rows - 1
            blnFreeDiag = Val(.TextMatrix(i, DI_诊断ID)) = 0 And Val(.TextMatrix(i, DI_疾病ID)) = 0 '自由录入诊断
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                If strTmp <> .TextMatrix(i, DI_诊断分类) Then
                    j = 1: strTmp = .TextMatrix(i, DI_诊断分类)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = ""
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                Next

                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_诊断描述) '自由录入诊断加上诊断描述
                mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值", "Tag", "信息现值", "主信息现值"), Array(lngTmp, Val(.RowData(i)), vsDiagInput.Name, strInfo, strMainInfo, IIf(.TextMatrix(i, DI_诊断来源) = "", 3, .TextMatrix(i, DI_诊断来源)), Null, Null)
                lngTmp = lngTmp + 1
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAllerData()
'功能：加载过敏信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngRow As Long, j As Long
    Dim strInfo As String '全信息
    Dim strMainInfo As String '主键信息
    Dim str过敏时间 As String
    Dim lngTmp As Long
    Dim strAll As String
    Dim blnTmp As Boolean
    blnTmp = mblnNoSave
    mblnNoSave = True
    On Error GoTo errH
    chkNoAller.Value = 0
    mrsMainInfo.Filter = "信息名='无过敏记录'"
    If Val(mrsMainInfo!信息原值 & "") = 1 Then
        chkNoAller.Value = 1
        mblnNoSave = blnTmp
        Exit Sub
    End If
    mblnNoSave = blnTmp
 
    strSQL = "Select a.ID,a.记录来源,a.过敏时间,a.药物id,a.药物名,a.过敏反应,a.过敏源编码" & vbNewLine & _
        "From 病人过敏记录 A" & vbNewLine & _
        "Where a.结果 = 1 And a.病人id =[1] And a.主页id =[2] And Not Exists" & vbNewLine & _
        " (Select b.药物id From 病人过敏记录 b" & vbNewLine & _
        "       Where (Nvl(b.药物id, 0) = Nvl(A.药物id, 0) Or Nvl(药物名, 'Null') = Nvl(A.药物名, 'Null')) And Nvl(结果, 0) = 0 And" & vbNewLine & _
        "             b.记录时间 > A.记录时间 And b.病人id =[1] And b.主页id =[2])" & vbNewLine & _
        "Order By Nvl(a.过敏时间,a.记录时间) Desc,a.记录时间 desc, a.药物名"

    If mblnMoved Then
        strSQL = Replace(strSQL, "病人过敏记录", "H病人过敏记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "首页获取过敏信息", mlng病人ID, mlng挂号ID)
    
    mstrTagAller = ""
    rsTmp.Filter = "记录来源=3"
    If rsTmp.EOF Then
        rsTmp.Filter = "记录来源<>3"
        mstrTagAller = "1" '非首页来源
    End If
    
    '删除之前的缓存
    mrsSecdInfo.Filter = "控件名='vsAller'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If
    
    '处理历史数据103674
    If rsTmp.RecordCount > 0 Then
        chkNoAller.Value = 0
        Call SetAllerEdit(False)
        Call UpDateInfo(0, "chkNoAller")
    End If
    
    lngTmp = 1
    With vsAller
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            '其它来源的可能有重复 唯一键：药物ID，药物名，过敏源编码，过敏时间
            str过敏时间 = Format(rsTmp!过敏时间 & "", "yyyy-MM-dd")
            strMainInfo = Val(rsTmp!药物ID & "") & "|" & rsTmp!药物名 & "|" & rsTmp!过敏源编码 & "|" & str过敏时间
            
            If InStr("," & strAll & ",", "," & strMainInfo & ",") = 0 Then
                strAll = strAll & "," & strMainInfo
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .TextMatrix(lngRow, AI_过敏时间) = str过敏时间
                .TextMatrix(lngRow, AI_过敏药物) = NVL(rsTmp!药物名)
                .TextMatrix(lngRow, AI_过敏反应) = NVL(rsTmp!过敏反应)
                .TextMatrix(lngRow, AI_过敏源编码) = NVL(rsTmp!过敏源编码)
                .TextMatrix(lngRow, AI_药物ID) = Val(rsTmp!药物ID & "")
                .TextMatrix(lngRow, AI_过敏来源) = rsTmp!记录来源 & ""
                '数据备份存储
                .Cell(flexcpData, lngRow, AI_过敏药物) = .TextMatrix(lngRow, AI_过敏药物)
                .RowData(lngRow) = Val(rsTmp!ID & "")
                
                strInfo = strMainInfo & "|" & rsTmp!过敏反应
                mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值"), Array(lngTmp, Val(rsTmp!ID & ""), "vsAller", strInfo, strMainInfo)
                lngTmp = lngTmp + 1
            End If
            rsTmp.MoveNext
        Next
        .Rows = .Rows + 1 '增加一行空行
        .Row = 1: .Col = AI_过敏药物
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Function GetAllerSaveSQL(ByRef arrSQL As Variant) As Boolean
'功能：获取过敏药保存的SQL
    Dim i As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim lngRow As Long
    Dim lngTmp As Long
    Dim strAll As String
    Dim strDels As String
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    With vsAller
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, AI_过敏药物) <> "" Then
                strMainInfo = Val(.TextMatrix(i, AI_药物ID)) & "|" & .TextMatrix(i, AI_过敏药物) & "|" & .TextMatrix(i, AI_过敏源编码) & "|" & .TextMatrix(i, AI_过敏时间)
                strInfo = strMainInfo & "|" & .TextMatrix(i, AI_过敏反应)
 
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '相同过每记录
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, AI_过敏反应) = &HC0C0FF
                    Call .ShowCell(i, AI_过敏药物)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '收集所有诊断用于判断是否有重复行
                End If
                
                mrsSecdInfo.Filter = "控件名='vsAller' and 序号=" & lngTmp
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!序号 = lngTmp
                    mrsSecdInfo!控件名 = "vsAller"
                End If
                mrsSecdInfo!现ID = Val(.RowData(i))
                mrsSecdInfo!信息现值 = strInfo
                mrsSecdInfo!主信息现值 = strMainInfo
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
 
                mrsSecdInfo.Filter = 0
            End If
        Next
    
        mrsSecdInfo.Filter = "控件名='vsAller'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
   
        '删除行以及主信息改变行需要调用删除方法
        mrsSecdInfo.Filter = "(改变状态=" & CS_删除行 & " And 控件名='vsAller') OR (改变状态=" & CS_替换行 & " And 控件名='vsAller')"
        Do While Not mrsSecdInfo.EOF
            strDels = strDels & "," & mrsSecdInfo!原ID
            mrsSecdInfo.MoveNext
        Loop
        
        '主信息改变以及新增行需要调用插入过程        '次级信息改变，调用更新过程
        mrsSecdInfo.Filter = "控件名='vsAller' And 改变状态>" & CS_未改变
        
        If mstrTagAller = "1" Then
            '如果修改了过敏记录把其他来源的过敏记录新增一份。
            If (strDels <> "" Or Not mrsSecdInfo.EOF) Then
                For lngRow = .FixedRows To .Rows - 1
                    If .TextMatrix(lngRow, AI_过敏药物) <> "" And .TextMatrix(lngRow, AI_过敏药物) <> "—" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Insert(" & mlng病人ID & "," & mlng挂号ID & "," & _
                            "3," & ZVal(.TextMatrix(lngRow, AI_药物ID)) & ",'" & .TextMatrix(lngRow, AI_过敏药物) & "',1," & _
                            ToDateOracle(.TextMatrix(lngRow, AI_过敏时间), "ymd") & ",SysDate,'" & _
                            .TextMatrix(lngRow, AI_过敏反应) & "','" & .TextMatrix(lngRow, AI_过敏源编码) & "')"
                    End If
                Next
            End If
        Else
            If strDels <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人过敏记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,'" & Mid(strDels, 2) & "')"
            End If
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                If .TextMatrix(lngRow, AI_过敏药物) <> "—" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mrsSecdInfo!改变状态 <> CS_更新行 Then
                        arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Insert(" & mlng病人ID & "," & mlng挂号ID & "," & _
                                "3," & ZVal(.TextMatrix(lngRow, AI_药物ID)) & ",'" & .TextMatrix(lngRow, AI_过敏药物) & "',1," & _
                                ToDateOracle(.TextMatrix(lngRow, AI_过敏时间), "ymd") & ",SysDate,'" & _
                                .TextMatrix(lngRow, AI_过敏反应) & "','" & .TextMatrix(lngRow, AI_过敏源编码) & "')"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_病人过敏记录_Update(" & mrsSecdInfo!原ID & "," & mlng病人ID & "," & mlng挂号ID & "," & _
                                "3," & ZVal(.TextMatrix(lngRow, AI_药物ID)) & ",'" & .TextMatrix(lngRow, AI_过敏药物) & "',1," & _
                                ToDateOracle(.TextMatrix(lngRow, AI_过敏时间), "ymd") & ",'" & _
                                .TextMatrix(lngRow, AI_过敏反应) & "','" & .TextMatrix(lngRow, AI_过敏源编码) & "')"
                    End If
                End If
                mrsSecdInfo.MoveNext
            Loop
        End If
    End With
    If UBound(arrSQL) <> -1 Then GetAllerSaveSQL = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetDiagSaveSQL(ByRef vsDiagInput As VSFlexGrid, ByRef arrSQL As Variant) As Boolean
'功能：获取诊断保存的SQL
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim blnFreeDiag As Boolean
    Dim datCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    Dim strAllDiag As String
    Dim strTag As String
    
    On Error GoTo errH
    arrSQL = Array()
    With vsDiagInput
        .Tag = ""
        strVsName = .Name
        strTmp = ""
        lngTmp = 1
        arrMain = Array(DI_诊断编码, DI_诊断分类, DI_诊断ID, DI_疾病ID)
        arrWhole = Array(DI_诊断分类, DI_疾病编码, DI_诊断编码, DI_ICD附码, DI_疾病类别, DI_证候编码, DI_中医证候, DI_是否疑诊, DI_诊断ID, DI_疾病ID, DI_诊断描述, DI_备注, DI_发病时间)
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                blnFreeDiag = Val(.TextMatrix(i, DI_诊断ID)) = 0 And Val(.TextMatrix(i, DI_疾病ID)) = 0 '自由录入诊断
                If strTmp <> .TextMatrix(i, DI_诊断分类) Then
                    j = 1: strTmp = .TextMatrix(i, DI_诊断分类)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = ""
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                Next
                If blnFreeDiag Then strMainInfo = strMainInfo & "|" & .TextMatrix(i, DI_诊断描述) '自由录入诊断加上诊断描述
                
                If strVsName = "vsDiagZY" Then
                    If InStr("," & strAllDiag & ",", "," & strMainInfo & ",") = 0 Then
                        strAllDiag = strAllDiag & "," & strMainInfo '收集所有诊断用于判断是否有重复行
                    End If
                Else
                    If InStr("," & strAllDiag & ",", "," & strMainInfo & ",") > 0 Then
                        '有相同诊断
                        .Tag = i
                        .Cell(flexcpBackColor, i, .FixedCols, i, DI_是否疑诊) = &HC0C0FF
                        Call .ShowCell(i, DI_诊断描述)
                        Exit Function
                    Else
                        strAllDiag = strAllDiag & "," & strMainInfo '收集所有诊断用于判断是否有重复行
                    End If
                End If
                mrsSecdInfo.Filter = "控件名='" & strVsName & "' and 序号=" & lngTmp
 
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!序号 = lngTmp
                    mrsSecdInfo!控件名 = strVsName
                End If
                mrsSecdInfo!信息现值 = strInfo
                mrsSecdInfo!主信息现值 = strMainInfo
                mrsSecdInfo!中医诊候 = .TextMatrix(i, DI_中医证候)
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
                mrsSecdInfo.Filter = 0
            End If
        Next
        
        If strVsName = "vsDiagZY" Then
            If strAllDiag <> "" Then
                arrOther = Split(strAllDiag, ",")
                For i = 0 To UBound(arrOther)
                    strTmp = arrOther(i)
                    If strTmp <> "" Then
                        mrsSecdInfo.Filter = "控件名='vsDiagZY' and 主信息现值='" & strTmp & "'"
                        If mrsSecdInfo.RecordCount > 2 Then
                            strInfo = "|同一中医诊断超过了2条。"
                            '超过3条了
                            .Tag = mrsSecdInfo!IndexEx & strInfo
                            lngTmp = Val(mrsSecdInfo!IndexEx)
                            .Cell(flexcpBackColor, lngTmp, .FixedCols, lngTmp, DI_是否疑诊) = &HC0C0FF
                            Call .ShowCell(lngTmp, DI_诊断描述)
                            Exit Function
                        ElseIf mrsSecdInfo.RecordCount > 1 Then
                            strAllDiag = "无"
                            For j = 1 To mrsSecdInfo.RecordCount
                                If InStr("," & strAllDiag & ",", "," & mrsSecdInfo!中医诊候 & ",") > 0 Then
                                    strInfo = "|同一中医诊断对应了两个相同的证候。"
                                    '有相同诊断
                                    .Tag = mrsSecdInfo!IndexEx & strInfo
                                    lngTmp = Val(mrsSecdInfo!IndexEx)
                                    .Cell(flexcpBackColor, lngTmp, .FixedCols, lngTmp, DI_是否疑诊) = &HC0C0FF
                                    Call .ShowCell(lngTmp, DI_诊断描述)
                                    Exit Function
                                Else
                                    strAllDiag = strAllDiag & "," & mrsSecdInfo!中医诊候 '收集所有诊断用于判断是否有重复行
                                End If
                                mrsSecdInfo.MoveNext
                            Next
                        End If
                    End If
                Next
            End If
        End If
        mrsSecdInfo.Filter = "控件名='" & strVsName & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
        
        '删除行以及主信息改变行需要调用删除方法
        mrsSecdInfo.Filter = "(改变状态=" & CS_删除行 & " And 控件名='" & strVsName & "') OR (改变状态=" & CS_替换行 & " And 控件名='" & strVsName & "')": strTmp = ""
        Do While Not mrsSecdInfo.EOF
            strTmp = strTmp & "," & mrsSecdInfo!原ID
            mrsSecdInfo.MoveNext
        Loop
        '主信息改变以及新增行需要调用插入过程
        '次级信息改变，调用更新过程
        mrsSecdInfo.Filter = "改变状态>" & CS_未改变 & " And 控件名='" & strVsName & "'"
        
        datCur = mdatCurDate
         
        If strVsName = "vsDiagXY" Then
            strTag = mstrTagDiagXY
        Else
            strTag = mstrTagDiagZY
        End If
        
        If strTag = "1" Then
            If strTmp <> "" Or Not mrsSecdInfo.EOF Then
                j = 1
                For lngRow = .FixedRows To .Rows - 1
                    If .TextMatrix(lngRow, DI_诊断描述) <> "" Then
                        If Trim(.TextMatrix(lngRow, DI_诊断编码)) = "" Then
                            strTmp = .TextMatrix(lngRow, DI_诊断描述) & IIf(.TextMatrix(lngRow, DI_中医证候) <> "", "(" & .TextMatrix(lngRow, DI_中医证候) & ")", "")
                        Else
                            strTmp = "(" & .TextMatrix(lngRow, DI_诊断编码) & ")" & .TextMatrix(lngRow, DI_诊断描述) & IIf(.TextMatrix(lngRow, DI_中医证候) <> "", "(" & .TextMatrix(lngRow, DI_中医证候) & ")", "")
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3,NULL," & .TextMatrix(lngRow, DI_诊断分类) & "," & _
                                ZVal(.TextMatrix(lngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(lngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(lngRow, DI_证候ID)) & ",'" & _
                                strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_是否疑诊) = "", 0, 1) & "," & ToDateOracle(datCur, "ymdhms") & ",'" & .TextMatrix(lngRow, DI_医嘱IDs) & "' ," & j & ",'" & _
                                .TextMatrix(lngRow, DI_备注) & "',Null," & ToDateOracle(.TextMatrix(lngRow, DI_发病时间), "ymdhm") & ")"
                        j = j + 1
                    End If
                Next
            End If
        Else
            If strTmp <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                '病案系统存储过程需加系统编号
                arrSQL(UBound(arrSQL)) = "Zl_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,NULL,NUll,'" & Mid(strTmp, 2) & "')"
            End If
            
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx: j = Val(Mid(mrsSecdInfo!信息现值, 1, InStr(mrsSecdInfo!信息现值, "|") - 1))
                If Trim(.TextMatrix(lngRow, DI_诊断编码)) = "" Then
                    strTmp = .TextMatrix(lngRow, DI_诊断描述) & IIf(.TextMatrix(lngRow, DI_中医证候) <> "", "(" & .TextMatrix(lngRow, DI_中医证候) & ")", "")
                Else
                    strTmp = "(" & .TextMatrix(lngRow, DI_诊断编码) & ")" & .TextMatrix(lngRow, DI_诊断描述) & IIf(.TextMatrix(lngRow, DI_中医证候) <> "", "(" & .TextMatrix(lngRow, DI_中医证候) & ")", "")
                End If
    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If mrsSecdInfo!改变状态 <> CS_更新行 Then
                    lngID = zlDatabase.GetNextId("病人诊断记录")
                    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3,NULL," & .TextMatrix(lngRow, DI_诊断分类) & "," & _
                                        ZVal(.TextMatrix(lngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(lngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(lngRow, DI_证候ID)) & ",'" & _
                                        strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_是否疑诊) = "", 0, 1) & "," & ToDateOracle(datCur, "ymdhms") & ",'" & .TextMatrix(lngRow, DI_医嘱IDs) & "' ," & j & ",'" & .TextMatrix(lngRow, DI_备注) & "'," & _
                                        "null," & ToDateOracle(.TextMatrix(lngRow, DI_发病时间), "ymdhm") & ",Null," & lngID & ")"
                Else
                    arrSQL(UBound(arrSQL)) = "Zl_病人诊断记录_Update(" & mrsSecdInfo!原ID & "," & mlng病人ID & "," & mlng挂号ID & ",3," & .TextMatrix(lngRow, DI_诊断分类) & "," _
                                        & ZVal(.TextMatrix(lngRow, DI_疾病ID)) & "," & ZVal(.TextMatrix(lngRow, DI_诊断ID)) & "," & ZVal(.TextMatrix(lngRow, DI_证候ID)) & ",'" & _
                                        strTmp & "',null,null," & IIf(.TextMatrix(lngRow, DI_是否疑诊) = "", 0, 1) & "," & j & ",'" & .TextMatrix(lngRow, DI_备注) & "',null," & ToDateOracle(.TextMatrix(lngRow, DI_发病时间), "ymdhm") & ")"
                End If
                mrsSecdInfo.MoveNext
            Loop
        End If
    End With
    If UBound(arrSQL) <> -1 Then GetDiagSaveSQL = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function ToDateOracle(ByVal strDate As String, ByVal strType As String) As String
'功能：获取ORACLE Date类型串
'参数：strDate=时间字符串
'      strType=格式字符串类型，ymd-年月日（yyyy-mm-dd)，ymdhm-（yyyy-mm-dd hh:mm),ymdhms-（yyyy-mm-dd hh:mm:ss)
    If Not IsDate(strDate) Then ToDateOracle = "Null": Exit Function
    Select Case UCase(strType)
        Case "YMD"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "YMDHM"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Case "YMDHMS"
           ToDateOracle = "To_Date('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
           ToDateOracle = "Null"
    End Select

End Function

Private Sub UpDateAller()
'功能：保存过敏记录
    Dim i As Long
    Dim blnTrans As Boolean
    Dim arrSQL As Variant
    Dim blnUpdate As Boolean
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    
    If Not GetAllerSaveSQL(arrSQL) Then
        '清除部分缓存
        mrsSecdInfo.Filter = "控件名='vsAller'"
        For i = 1 To mrsSecdInfo.RecordCount
            Call mrsSecdInfo.Update(Array("信息现值", "主信息现值"), Array(Null, Null))
            mrsSecdInfo.MoveNext
        Next
        Exit Sub
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        blnUpdate = True
    Next
    gcnOracle.CommitTrans: blnTrans = False
    If blnUpdate Then
        mblnOK = True
        Call LoadAllerData
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub UpDateDiag(ByRef vsDiagInput As VSFlexGrid)
'功能：保存诊断记录
    Dim i As Long
    Dim blnTrans As Boolean
    Dim arrSQL As Variant
    Dim str疾病ID As String
    Dim str诊断ID As String
    Dim lngRowXY As Long
    Dim lngColXY As Long
    Dim lngRowZY As Long
    Dim lngColZY As Long
    
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    If mbln不更新诊断 Then Exit Sub
    
    If Not GetDiagSaveSQL(vsDiagInput, arrSQL) Then
        '清除部分缓存
        mrsSecdInfo.Filter = "控件名='" & vsDiagInput.Name & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            Call mrsSecdInfo.Update(Array("信息现值", "主信息现值", "中医诊候"), Array(Null, Null, ""))
            mrsSecdInfo.MoveNext
        Next
        Exit Sub
    End If
    Call MsgDis
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    
    lngRowXY = vsDiagXY.Row
    lngColXY = vsDiagXY.Col
    lngRowZY = vsDiagZY.Row
    lngColZY = vsDiagZY.Col
    
    Call LoadDiagData
    
    With vsDiagXY
        If lngRowXY < .Rows And lngRowXY >= .FixedRows Then .Row = lngRowXY
        If lngColXY < .Cols And lngColXY >= .FixedCols Then .Col = lngColXY
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then
                If InStr("," & str诊断ID & ",", "," & Val(.TextMatrix(i, DI_诊断ID)) & ",") = 0 Then
                    str诊断ID = str诊断ID & "," & Val(.TextMatrix(i, DI_诊断ID))
                End If
            End If
            If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                If InStr("," & str疾病ID & ",", "," & Val(.TextMatrix(i, DI_疾病ID)) & ",") = 0 Then
                    str疾病ID = str疾病ID & "," & Val(.TextMatrix(i, DI_疾病ID))
                End If
            End If
        Next
    End With
    
    If mbln中医 Then
        With vsDiagZY
            If lngRowZY < .Rows And lngRowZY >= .FixedRows Then .Row = lngRowZY
            If lngColZY < .Cols And lngColZY >= .FixedCols Then .Col = lngColZY
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then
                    If InStr("," & str诊断ID & ",", "," & Val(.TextMatrix(i, DI_诊断ID)) & ",") = 0 Then
                        str诊断ID = str诊断ID & "," & Val(.TextMatrix(i, DI_诊断ID))
                    End If
                End If
                If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                    If InStr("," & str疾病ID & ",", "," & Val(.TextMatrix(i, DI_疾病ID)) & ",") = 0 Then
                        str疾病ID = str疾病ID & "," & Val(.TextMatrix(i, DI_疾病ID))
                    End If
                End If
            Next
        End With
    End If
    str疾病ID = Mid(str疾病ID, 2): str诊断ID = Mid(str诊断ID, 2)
    If str疾病ID <> "" Or str诊断ID <> "" Then RaiseEvent UpdateDiagInfo(str疾病ID, str诊断ID, "")

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Function GetPatiDiagData(ByVal blnLast As Boolean) As ADODB.Recordset
'功能：获取病人的诊断记录，以记录集形式返回方便加载到表格中
'参数：blnLast 是否是最后一次就诊
    Dim strSQL As String, strSQLTmp As String, strDiagType As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If blnLast Then
    '最后一次的就诊ID
        strSQLTmp = "(Select Max(ID) As 主页id" & vbNewLine & _
            "From 病人挂号记录" & vbNewLine & _
            "Where 病人id = [1] And 记录性质 = 1 And 记录状态 = 1 And" & vbNewLine & _
            "      登记时间 =" & vbNewLine & _
            "      (Select Max(A.登记时间)" & vbNewLine & _
            "       From 病人挂号记录 A" & vbNewLine & _
            "       Where A.病人id = [1] And A.记录性质 = 1 And A.记录状态 = 1 And A.登记时间 < (Select 登记时间 From 病人挂号记录 Where ID = [2])))"
    End If
    
    '设置读取诊断的类别以及诊断来源
    If mbln中医 Then
        strDiagType = " And A.记录来源 IN(1,3) And A.诊断类型 IN(1,11) "
    Else
        strDiagType = " And A.记录来源 IN(1,3) And A.诊断类型=1 "
    End If

    '组装SQL,电子病案查阅不用查询医嘱记录
    strSQL = "Select A.备注, A.Id, A.病人id, A.主页id, A.医嘱id, A.记录来源, A.诊断次序, A.编码序号, A.诊断类型, A.入院病情, A.疾病id, A.诊断id, A.证候id,B.名称 疾病名称,C.名称 诊断名称,D.名称 证候名称," & vbNewLine & _
            "       A.诊断描述, A.出院情况, A.是否未治, A.是否疑诊, A.发病时间, B.编码 As 疾病编码,B.类别 As 疾病类别 , C.编码 As 诊断编码, D.编码 As 证候编码,B.附码," & vbNewLine & _
            "(Select F_List2str(Cast(Collect(C.医嘱id|| '') As T_Strlist)) 医嘱id From 病人诊断医嘱 C Where C.诊断id = A.Id) As 医嘱id," & _
            "B.性别限制, B.疗效限制, B.分娩, B.附码, E.Id As 大类, E.是否病人,A.记录日期,A.记录人 " & vbNewLine & _
            "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C, 疾病编码目录 D,疾病编码分类 E" & vbNewLine & _
            "Where A.疾病id = B.Id(+) And A.诊断id = C.Id(+) And A.证候id = D.Id(+)  And  B.分类id = E.Id(+)" & strDiagType & "And A.取消时间 Is Null And A.诊断描述 Is Not Null And 病人id = [1] And 主页id =[2]" & strSQLTmp & vbNewLine & _
            "Order By A.诊断类型, A.记录来源 Desc, A.诊断次序, A.编码序号, A.Id"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取首页诊断", mlng病人ID, mlng挂号ID)
    Set GetPatiDiagData = rsTmp
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub InitVSFlexGrid(ByRef vsInput As VSFlexGrid, ByVal strHead As String, Optional ByVal strRowContent As String, Optional ByVal intFixedCols As Integer, Optional ByVal intFixedRows As Integer = 1)
'功能：初始化表格内容，用在窗体个性化设置恢复之前
'参数：vsInput=要设置格式的表格
'          strHead=表格的列格式，格式为：列标题1,列宽1,对齐1,数据类型1,格式串1;列标题2,列宽2,对齐2,数据类型2,格式串2.....
'          strRowContent=表格的预定义行内容,格式为：列1,内容1,列2,内容2:行1;列1,内容1,列2,内容2:行2;....(列要从小到大排列）:行1
    Dim i As Integer, lngRow As Long, j As Long
    Dim arrHead As Variant, arrCol As Variant, arrRow As Variant
    Dim arrTmp As Variant
    On Error GoTo errH
    '设置列
    With vsInput
        If strHead <> "" Then
            arrHead = Split(strHead, ";")
            .Clear: .Cols = 0: .Rows = 0
            .Rows = intFixedRows + 1: .Cols = UBound(arrHead) + 1
            .FixedRows = intFixedRows: .FixedCols = intFixedCols
            For i = LBound(arrHead) To UBound(arrHead)
                arrCol = Split(arrHead(i), ",")
                .FixedAlignment(i) = 4
                If intFixedRows <> 0 Then .TextMatrix(0, i) = arrCol(0)
                If UBound(arrCol) > 0 Then
                    .ColWidth(i) = Val(arrCol(1))
                Else
                    .ColHidden(i) = True
                End If
                If UBound(arrCol) > 1 Then .ColAlignment(i) = Val(arrCol(2))
                If UBound(arrCol) > 2 Then .ColDataType(i) = Val(arrCol(3))
                If UBound(arrCol) > 3 Then .ColFormat(i) = arrCol(4)
                If UBound(arrCol) > 4 Then .ColHidden(i) = Val(arrCol(5))
            Next
        End If
        '设置解析行
        If strRowContent <> "" Then
            .Rows = .FixedRows
            lngRow = .FixedRows - 1: arrRow = Split(strRowContent, ";")
            For i = LBound(arrRow) To UBound(arrRow)
                arrTmp = Split(arrRow(i), ";")
                '确定行号
                lngRow = lngRow + 1
                If UBound(arrTmp) > 0 Then lngRow = Val(arrTmp(1))
                .Rows = lngRow + 1 '设置行数
                '设置内容
                arrCol = Split(arrTmp(0), ",")
                For j = LBound(arrCol) To UBound(arrCol) Step 2
                    .TextMatrix(lngRow, Val(arrCol(j))) = arrCol(j + 1)
                Next
            Next
        End If
    End With
    vsAller.ExtendLastCol = False
    Exit Sub
errH:
    Debug.Print err.Source & "-InitVSFlexGrid:" & err.Description
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MovePanel()

    PicPanel(picPanel_基本信息).Top = mlngTopVsc
    
    If mblnDocInput Then
        PicPanel(picPanel_快键病历).Top = PicPanel(picPanel_基本信息).Height + PicPanel(picPanel_基本信息).Top
        PicPanel(picPanel_就诊信息).Top = PicPanel(picPanel_快键病历).Height + PicPanel(picPanel_快键病历).Top
    Else
        PicPanel(picPanel_就诊信息).Top = PicPanel(picPanel_基本信息).Height + PicPanel(picPanel_基本信息).Top
    End If
    
    dtpDate.Top = PicPanel(picPanel_就诊信息).Top - dtpDate.Height + txtE(I发病时间).Top - 20
    
    dtpDate.Left = PicPanel(picPanel_就诊信息).Left + txtE(I发病时间).Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intTmp As Integer
    
    intTmp = IIf(optInfo(opt疾病).Value, 1, 0)
    Call zlDatabase.SetPara("门诊诊断输入", intTmp, glngSys, p门诊医生站, InStr(gstrPrivs, "参数设置") > 0)
    
    intTmp = IIf(optInfo(opt过敏源).Value, 2, 1)
    Call zlDatabase.SetPara("过敏输入来源", intTmp, glngSys, p门诊医生站, gint过敏输入来源 = 0 And gbytPass = 3 And InStr(gstrPrivs, "参数设置") > 0)
    Call SavePreItem
    Call ClearPatiInfo
    Set mobjKernel = Nothing
    Set mobjPatient = Nothing
    Set mobjCtl = Nothing
    Set mclsZip = Nothing
    Set mclsUnZip = Nothing
    Set mrsMainInfo = Nothing
    Set mrsSecdInfo = Nothing
    Set mrsPreEditCtl = Nothing
    mblnCboNoClick = False
    mlng挂号ID = 0
End Sub

Private Sub txtSL_GotFocus()
    Call zlControl.TxtSelAll(txtSL)
End Sub

Private Sub txtSL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtSL_Validate(Cancel As Boolean)
    Dim datCur As Date, datRes As Date
    
    If Trim(txtSL.Text) <> "" Then
        If IsNumeric(txtSL.Text) Then
            If Val(txtSL.Text) <= 0 Then
                MsgBox "发病时间推算值必须为正数。", vbInformation, gstrSysName
                txtSL.SetFocus: Exit Sub
            End If
        Else
            MsgBox "发病时间推算值必须为数字。", vbInformation, gstrSysName
            txtSL.Text = "": txtSL.SetFocus: Exit Sub
        End If
    Else
         Exit Sub
    End If
    If cboE(I日期).ListIndex <= 0 Then Exit Sub
    datCur = Format(mdatCurDate, "yyyy-MM-dd HH:mm")
    Select Case cboE(I日期).ListIndex
    Case 1 '小时
        datRes = DateAdd("n", -1 * Val(txtSL.Text) * 60, CDate(datCur))
    Case 2 '天
        datRes = DateAdd("h", -1 * Val(txtSL.Text) * 24, CDate(datCur))
    Case 3 '周
        datRes = DateAdd("d", -1 * 7 * Val(txtSL.Text), CDate(datCur))
    Case 4 '月
        datRes = DateAdd("M", -1 * Int(Val(txtSL.Text)), CDate(datCur))
        datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 30, datRes)
    Case 5 '年
        If Val(txtSL.Text) < 100 Then
            datRes = DateAdd("yyyy", -1 * Int(Val(txtSL.Text)), CDate(datCur))
            datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 365, datRes)
        Else
            MsgBox "发病时间推算不能超过100年。", vbInformation, gstrSysName
            txtSL.SetFocus: Exit Sub
        End If
    End Select
    txtE(I发病时间).Text = Format(CDate(datRes), "YYYY-MM-DD HH:mm")
    If Not txtE(I发病时间).Locked Then
        Call UpDateInfo(txtE(I发病时间).Text, "txtE", I发病时间)
    End If
End Sub

Private Sub UCPatiVitalSigns_Change(ByVal int序号 As Integer)
    mblnChange = True
End Sub

Private Sub UCPatiVitalSigns_GotFocus()
    Call SetCurCtlInfo(TypeName(UCPatiVitalSigns), "UCPatiVitalSigns")
End Sub

Private Sub UCPatiVitalSigns_Validate(Cancel As Boolean)
    Dim strSQL As String
    If mblnNoSave Then Exit Sub
    If mblnChange Then
        strSQL = UCPatiVitalSigns.GetSaveSQL(mlng病人ID, mlng挂号ID)
        mrsMainInfo.Filter = "控件名='UCPatiVitalSigns'"
        If mrsMainInfo!信息原值 & "" <> strSQL Then
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            mrsMainInfo!信息原值 = strSQL
            mrsMainInfo.Update
            strSQL = ""
            With UCPatiVitalSigns
                strSQL = mlng病人ID & "<split>" & mlng挂号ID & "<split>" & .value身高 & "<split>" & .value体重 & "<split>" & .value体温 & "<split>" & _
                    .value脉搏 & "<split>" & .value呼吸 & "<split>" & .value收缩压 & "<split>" & .value舒张压 & "<split>" & .value血压单位
            End With
            RaiseEvent UpdatePatiState(strSQL, "")
            mblnOK = True
        End If
    End If
    mblnChange = False
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    On Error GoTo errH
 
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    Me.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsAller_GotFocus()
    Call SetCurCtlInfo(TypeName(vsAller), "vsAller")
End Sub

Private Sub vsAller_Validate(Cancel As Boolean)
''''''''''''''
    Dim vsTmp As Object
    Dim i As Long
    Dim j As Long
    
    Set vsTmp = vsAller
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, AI_过敏药物)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, AI_过敏药物)) > 60 Then
                    .Row = i: .Col = AI_过敏药物
                    Call ShowMessage(vsTmp, "过敏药物名太长，只允许60个字符或30个汉字。")
                    Cancel = True
                    Exit Sub
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, AI_过敏反应)) > 100 Then
                    .Row = i: .Col = AI_过敏反应
                    Call ShowMessage(vsTmp, "过敏反应内容太长，只允许100个字符或50个汉字。")
                    Cancel = True
                    Exit Sub
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, AI_过敏药物)) <> "" And Format(.TextMatrix(i, AI_过敏时间), "yyyy-mm-dd") = Format(.TextMatrix(j, AI_过敏时间), "yyyy-mm-dd") Then
                        If .TextMatrix(j, AI_过敏药物) = .TextMatrix(i, AI_过敏药物) Then
                            .Row = i: .Col = AI_过敏药物
                            Call ShowMessage(vsTmp, "发现" & Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。")
                            Cancel = True
                            Exit Sub
                        ElseIf Val(.TextMatrix(i, AI_药物ID)) <> 0 And .TextMatrix(i, AI_药物ID) = .TextMatrix(j, AI_药物ID) Then
                            .Row = i: .Col = AI_过敏药物
                             Call ShowMessage(vsTmp, "发现" & Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。")
                             Cancel = True
                            Exit Sub
                        ElseIf .TextMatrix(i, AI_过敏源编码) <> "" And .TextMatrix(i, AI_过敏源编码) = .TextMatrix(j, AI_过敏源编码) Then
                            .Row = i: .Col = AI_过敏药物
                            Call ShowMessage(vsTmp, "发现" & Format(.TextMatrix(j, AI_过敏时间), "yyyy年mm月dd日") & "内存在相同的过敏药物记录。")
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                Next
            End If
        Next
    End With
    Call UpDateAller
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    mlngTopVsc = -1 * vsc.Value * Screen.TwipsPerPixelY
    Call MovePanel
End Sub

Private Sub cboE_Click(Index As Integer)
    Dim strValue As String
    Dim datCur As Date, datRes As Date
    
    If mblnCboNoClick Then Exit Sub
    strValue = NeedName(cboE(Index).Text)
    Call UpDateInfo(strValue, "cboE", Index)
    
    If Index = I日期 Then
        If cboE(I日期).ListIndex <= 0 Then Exit Sub
        If Trim(txtSL.Text) = "" Then Exit Sub
        datCur = Format(mdatCurDate, "yyyy-MM-dd HH:mm")
        Select Case cboE(I日期).ListIndex
            Case 1 '小时
                datRes = DateAdd("n", -1 * Val(txtSL.Text) * 60, CDate(datCur))
            Case 2 '天
                datRes = DateAdd("h", -1 * Val(txtSL.Text) * 24, CDate(datCur))
            Case 3 '周
                datRes = DateAdd("d", -1 * 7 * Val(txtSL.Text), CDate(datCur))
            Case 4 '月
                datRes = DateAdd("M", -1 * Int(Val(txtSL.Text)), CDate(datCur))
                datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 30, datRes)
            Case 5 '年
                If Val(txtSL.Text) < 100 Then
                    datRes = DateAdd("yyyy", -1 * Int(Val(txtSL.Text)), CDate(datCur))
                    datRes = DateAdd("d", -1 * (Val(txtSL.Text) - Int(Val(txtSL.Text))) * 365, datRes)
                Else
                    MsgBox "发病时间推算不能超过100年。", vbInformation, gstrSysName
                    txtSL.SetFocus: Exit Sub
                End If
        End Select
        txtE(I发病时间).Text = Format(CDate(datRes), "YYYY-MM-DD HH:mm")
        If Not txtE(I发病时间).Locked Then
            Call UpDateInfo(txtE(I发病时间).Text, "txtE", I发病时间)
        End If
    End If
End Sub

Private Sub cboE_GotFocus(Index As Integer)
    If Index = I身份证号 Then
        Call zlControl.TxtSelAll(cboE(Index))
    End If
    Call SetCurCtlInfo(TypeName(cboE(Index)), "cboE", Index)
End Sub

Private Sub cboE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboE(Index).Style = 0 Then
            cboE(Index).ListIndex = -1
            cboE(Index).Text = ""
        Else
            If cboE(Index).ListIndex <> -1 Then
               cboE(Index).ListIndex = -1
            End If
        End If
    End If
End Sub

Private Sub cboE_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    Dim lngidx As Long
    Dim blnCancel As Boolean
    
    If Index = I身份证号 Then
        Call cboSpecificInfoKeyPress(Index, KeyAscii)
    End If
    
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = I身份证号 Then
            Call cboE_Validate(I身份证号, blnCancel)
        End If
        If Not blnCancel Then
            
            If Index = IRH Then
                cboE(I身份证号).SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        If Index = I身份证号 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            Else
                lblN(I身份证号).Tag = "1"
            End If
        Else
            lngidx = Cbo.MatchIndex(cboE(Index).hwnd, KeyAscii)
            If lngidx = -1 And cboE(Index).ListCount > 0 Then lngidx = 0
            cboE(Index).ListIndex = lngidx
        End If
    ElseIf KeyAscii = 8 Then
        If Index = I身份证号 Then
            lblN(I身份证号).Tag = "1"
        End If
    End If
End Sub

Private Sub cboE_Validate(Index As Integer, Cancel As Boolean)
    Dim strValue As String
    Dim str身份证号 As String
    
    If Index = I身份证号 Then
        '完整的身份证号是存在不加掩码  cboE(Index).Tag
        If cboE(Index).ListIndex = -1 Then
            strValue = cboE(Index).Tag
            mrsMainInfo.Filter = "信息名='身份证号'"
            str身份证号 = mrsMainInfo!信息原值 & ""
            If strValue <> str身份证号 Then
                If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                    If Not Check身份证号(strValue, cboE(Index)) Then
                        Cancel = True
                        cboE(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        Call UpDate身份证号
    End If
End Sub

Private Function Check身份证号(ByVal strNO As String, objTmp As Object) As Boolean
'功能：身份证号检查
    Dim strTmp As String
    Dim lngColor As Long
    Dim str出生日期 As String
    Dim lng性别 As Long
    Dim strBirthday As String, strAge As String, strSex As String, strErrIfno As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If mobjPatient Is Nothing Then
        On Error Resume Next
        Set mobjPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        err.Clear: On Error GoTo 0
    End If
    If mobjPatient Is Nothing Then
        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
        Exit Function
    End If

    On Error GoTo errH
    Call mobjPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
    
    strTmp = strNO
    lngColor = objTmp.BackColor
    
    If mobjPatient.CheckPatiIdcard(strTmp, strBirthday, strAge, strSex, strErrIfno) Then '省份证合法则检查是否匹配
        '判断是否已经存了，要禁止
        If gblnPatiByID Then
            strSQL = "select 1 from 病人信息 a where a.身份证号=[1] and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
            If Not rsTmp.EOF Then
                cboE(I身份证号).BackColor = &HC0C0FF
                MsgBox "该身份证号已经建档，同一身份证只能对应一个建档病人!", vbInformation, Me.Caption
                cboE(I身份证号).BackColor = lngColor
                Exit Function
            End If
        End If

        If objTmp.Index = I身份证号 Then
            strTmp = ""
            If Format(strBirthday, "yyyy-MM-dd") <> Format(mstr出生日期, "yyyy-MM-dd") Then
                strTmp = "出生日期"
            End If
            If mstr性别 <> strSex Then
                strTmp = strTmp & IIf(strTmp <> "", "、", "") & "性别"
            End If
            If strAge <> mstr年龄 Then
                strTmp = strTmp & IIf(strTmp <> "", "、", "") & "年龄"
            End If
            
            If strTmp <> "" Then
                If InStr(GetInsidePrivs(p病人信息公共部件), "基本信息调整") = 0 Then
                    cboE(I身份证号).BackColor = &HC0C0FF
                    If MsgBox("身份证号码获取的" & strTmp & "与病人当前的" & strTmp & "不相符，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboE(I身份证号).BackColor = lngColor
                        Exit Function
                    Else
                        cboE(I身份证号).BackColor = lngColor
                    End If
                Else
                    strErrIfno = "身份证号码获取的" & strTmp & "与病人当前的" & strTmp & "不相符，是否继续？继续则将自动跟新界面上的" & strTmp & "。"
                    
                    cboE(I身份证号).BackColor = &HC0C0FF
                    If MsgBox(strErrIfno, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboE(I身份证号).BackColor = lngColor
                        Exit Function
                    Else
                        cboE(I身份证号).BackColor = lngColor
                    End If
                    
                    If mobjPatient.SavePatiBaseInfo(mlng病人ID, mlng挂号ID, mstr姓名, strSex, strAge, strBirthday, "门诊首页", 1, strErrIfno) Then
                        
                    Else
                        cboE(I身份证号).BackColor = &HC0C0FF
                        If MsgBox(strErrIfno & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            cboE(I身份证号).BackColor = lngColor
                            Exit Function
                        Else
                            cboE(I身份证号).BackColor = lngColor
                        End If
                    End If
                    mstr性别 = strSex
                    mstr年龄 = strAge
                    mstr出生日期 = Format(strBirthday, "yyyy-MM-dd")
                    RaiseEvent UpdatePatiInfo(strBirthday, strAge, strSex, "")
                End If
            End If
        End If
    Else '身份证不合法则退出
        objTmp.BackColor = &HC0C0FF
        MsgBox strErrIfno, vbInformation, gstrSysName
        objTmp.BackColor = lngColor
        Exit Function
    End If
         
    Check身份证号 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendMsgDiag(ByVal datCur As Date) As Boolean
'功能：发送诊断消息
    Dim i As Long
    Dim arrTmp As Variant
    Dim strFilter As String
    On Error GoTo errH
    If mclsMipModule Is Nothing Then SendMsgDiag = True: Exit Function
  
    mrsSecdInfo.Filter = "改变状态<>" & CS_未改变 & " And 改变状态<>" & CS_更新行
    Do While Not mrsSecdInfo.EOF
        If mrsSecdInfo!控件名 = "vsDiagXY" Or mrsSecdInfo!控件名 = "vsDiagZY" Then
            arrTmp = Split(mrsSecdInfo!信息原值 & "", "|")
            If mrsSecdInfo!改变状态 <> CS_新增行 Then '删除行与替换行先触发删除诊断消息
'                Call ZLHIS_CIS_011(mclsMipModule, mlng病人ID, mstr姓名, 1, mlng挂号ID, gclsPros.出院科室ID, mrsSecdInfo!ID, arrTmp(DMP_诊断编码), arrTmp(DMP_疾病编码))
            End If
            arrTmp = Split(mrsSecdInfo!信息现值 & "", "|")
            If mrsSecdInfo!改变状态 <> CS_删除行 Then  '新增行与替换行触发下达诊断消息
'                Call ZLHIS_CIS_010(mclsMipModule, mlng病人ID, mstr姓名, 1, mlng挂号ID, gclsPros.出院科室ID, Val(mrsSecdInfo!Tag & ""), arrTmp(DMP_诊断类型), arrTmp(DMP_是否疑诊), arrTmp(DMP_诊断次序), arrTmp(DMP_诊断编码), arrTmp(DMP_疾病编码), arrTmp(DMP_疾病附码), arrTmp(DMP_疾病类别), arrTmp(DMP_证候编码), arrTmp(DMP_证候名称), datCur, UserInfo.姓名)
            End If
        End If
        mrsSecdInfo.MoveNext
    Loop
    
    SendMsgDiag = True
    '病原学诊断不触发消息
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function UpDateInfo(ByVal strValue As String, ByVal strCtlName As String, Optional ByVal intIdx As Integer = -1) As Boolean
    Dim strSQL As String, strSqlTwo As String
    Dim strFilter As String
    Dim str信息名 As String
    Dim strInfo As String
    Dim objTmp As Object
    Dim strMask As String
    Dim strTmp As String
    Dim i As Long
    Dim lngTmp As Long
    Dim strPar As String
    Dim datCur As Date
    Dim blnTrans As Boolean, blnEMPI As Boolean, strMsg As String
    
    If mblnNoSave Then Exit Function
    If mblnEdit Then Exit Function
    If strCtlName = "cboE" And intIdx = I身份证号 Then
        '身份证号特殊处理
        Call UpDate身份证号
    ElseIf strCtlName = "PatiAddress" Then
        '结构化地址特殊处理
        Call UpDate结构化地址(intIdx)
    Else
        strFilter = "控件名='" & strCtlName & "' and Index= " & intIdx
        If strCtlName = "chkNoAller" Then
            '无过敏记录处理
            strFilter = "控件名='" & strCtlName & "'"
        End If
        mrsMainInfo.Filter = strFilter
        If Not mrsMainInfo.EOF Then
            If strCtlName = "txtE" Then
                Set objTmp = txtE(intIdx)
                If (intIdx = I出生地点 Or intIdx = I户口地址 Or intIdx = I家庭地址) And mblnStructAdress Then
                    strInfo = PatiAddress(Decode(intIdx, I出生地点, PT_出生地点, I户口地址, PT_户口地址, I家庭地址, PT_家庭地址)).Value
                Else
                    strInfo = Trim(objTmp.Text)
                End If
                strInfo = Replace(strInfo, "'", "’")
                If InStr(",摘要,发病时间,发病地址,", "," & mrsMainInfo!信息名 & ",") = 0 Then
                    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
                        objTmp.BackColor = &HC0C0FF
                        objTmp.Tag = mrsMainInfo!信息名 & "-内容太长(允许录入" & objTmp.MaxLength & "个字符或" & objTmp.MaxLength \ 2 & "个汉字)。"
                        Exit Function
                    Else
                        objTmp.BackColor = vbWindowBackground
                        If intIdx <> I监护人身份证号 Then
                            objTmp.Tag = ""
                        End If
                    End If
                End If
                
                If strInfo <> "" Then
                    Select Case intIdx
                        Case I家庭电话, I单位电话
                            strMask = "1234567890-()"
                            lngTmp = Len(strInfo)
                            strTmp = strInfo
                            objTmp.BackColor = vbWindowBackground
                            objTmp.Tag = ""
                            For i = 1 To lngTmp
                                If InStr(strMask, Mid(strTmp, i, 1)) = 0 Then
                                    objTmp.BackColor = &HC0C0FF
                                    objTmp.Tag = mrsMainInfo!信息名 & "-内容中包含非法字符(允许录入以下字符：‘" & strMask & "’)。"
                                    Exit Function
                                End If
                            Next
                        Case I家庭邮编, I单位邮编, I户口邮编
                            strMask = "1234567890"
                            If (Not IsNumeric(strInfo)) Or InStr(strInfo, ".") > 0 Then
                                objTmp.BackColor = &HC0C0FF
                                objTmp.Tag = mrsMainInfo!信息名 & "-内容中包含非法字符(允许录入0－9的数字)。"
                                Exit Function
                            Else
                                objTmp.BackColor = vbWindowBackground
                                objTmp.Tag = ""
                            End If
                    End Select
                End If
                objTmp.Text = strInfo
                strValue = strInfo
            End If
            
            If mrsMainInfo!信息原值 & "" <> strValue Then
                strTmp = strValue
                If InStr(",摘要,发病时间,发病地址,", "," & mrsMainInfo!信息名 & ",") > 0 Then
                    Call UpDate挂号信息(mrsMainInfo!信息名 & "", strValue)
                ElseIf mrsMainInfo!信息名 = "监护人" Then
                    If mrsMainInfo!信息原值 & "" <> "" And strValue = "" Then
                        MsgBox "监护人信息只能修改，不能清除。", vbInformation, gstrSysName
                        txtE(I监护人).Text = mrsMainInfo!信息原值 & ""
                        Exit Function
                    Else
                        strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                        strSQL = "Zl_病人信息_更新信息(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & ")"
                        blnEMPI = True
                    End If
                ElseIf InStr(",RH,血型,其他医学警示,医学警示,", "," & mrsMainInfo!信息名 & ",") > 0 Then
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    strSQL = "Zl_病人信息从表_Update(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & "," & mlng挂号ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                ElseIf mrsMainInfo!信息名 = "单位名称" Then
                    If strValue = "" Then
                        mlng合同单位ID = 0
                    Else
                        strValue = "'" & strValue & "'"
                    End If
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    strSQL = "Zl_病人信息_更新信息(" & mlng病人ID & ",'工作单位'," & strValue & ")"
                    strSqlTwo = "Zl_病人信息_更新信息(" & mlng病人ID & ",'合同单位id'," & IIf(mlng合同单位ID = 0, "Null", mlng合同单位ID) & ")"
                    blnEMPI = True
                ElseIf mrsMainInfo!信息名 = "婚姻状况" Then
                    i = 0
                    Set objTmp = cboE(I婚姻状况)
                    If objTmp.Text <> "" And objTmp.ListIndex <> -1 Then
                        If InStr(objTmp.Text, "未婚") = 0 And InStr(objTmp.Text, "其他") = 0 Then
                            If IsDate(mstr出生日期) Then
                                datCur = mdatCurDate
                                If DateDiff("yyyy", CDate(mstr出生日期), datCur) < 15 Then
                                     If MsgBox("该病人年龄小于15岁，婚姻状况应该写为未婚或其他，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                        strTmp = mrsMainInfo!信息原值 & ""
                                        i = 1
                                        '恢复原值
                                        mblnCboNoClick = True
                                        objTmp.ListIndex = -1
                                        Call GetCboIndex(objTmp, strTmp)
                                        mblnCboNoClick = False
                                     End If
                                End If
                            End If
                        End If
                    End If
                    If i = 0 Then
                        strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                        strSQL = "Zl_病人信息_更新信息(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & ")"
                        blnEMPI = True
                    End If
                ElseIf InStr(",监护人身份证号,", "," & mrsMainInfo!信息名 & ",") > 0 Then
                    Call Update监护人身份证
                Else
                    strValue = IIf(strValue = "", "null", "'" & strValue & "'")
                    If mrsMainInfo!来源 = 0 Then
                        strSQL = "Zl_病人信息_更新信息(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & ")"
                        blnEMPI = True
                    Else
                        If mrsMainInfo!信息名 = "无过敏记录" Then
                            strSQL = "Zl_病人信息从表_Update(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & "," & mlng挂号ID & ")"
                        Else
                            strSQL = "Zl_病人信息从表_Update(" & mlng病人ID & ",'" & mrsMainInfo!信息名 & "'," & strValue & "," & mlng挂号ID & ")"
                            If mrsMainInfo!信息名 = "文化程度" Then
                                blnEMPI = True
                            End If
                        End If
                    End If
                    On Error GoTo errH
                    gcnOracle.BeginTrans: blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    If blnEMPI Then
                        If EMPIModifyPatiInfo(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, strMsg) = 0 Then
                            gcnOracle.RollbackTrans
                            blnTrans = False
                            MsgBox strMsg, vbInformation, gstrSysName
                            Exit Function
                        End If
                        blnEMPI = False
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If InStr(",家庭地址,家庭电话,", "," & mrsMainInfo!信息名 & ",") > 0 Then
                        If HaveRIS Then
                            If gobjRis.HISModPati(1, mlng病人ID, mlng挂号ID) <> 1 Then
                                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
                            End If
                        ElseIf gbln启用影像信息系统接口 = True Then
                            MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
                        End If
                    End If
                End If
                
                If blnEMPI Then
                    gcnOracle.BeginTrans: blnTrans = True
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    If strSqlTwo <> "" Then
                        Call zlDatabase.ExecuteProcedure(strSqlTwo, Me.Caption)
                    End If
                    If EMPIModifyPatiInfo(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, strMsg) = 0 Then
                        gcnOracle.RollbackTrans
                        blnTrans = False
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                End If
                
                mrsMainInfo!信息原值 = strTmp
                mrsMainInfo.Update
                mblnOK = True
            End If
        End If
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Update监护人身份证()
'功能：监护人身份证号
    Dim strSQL As String, i As Long, blnTrans As Boolean
    Dim strValue As String
    Dim intIndex As Integer
    Dim strPar As String, strMsg As String
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    mblnReturn = True
    intIndex = cboE(I身份证号).ListIndex
    strValue = Trim(txtE(I监护人身份证号).Text)
    strPar = IIf(strValue = "", "null", "'" & strValue & "'")
    strSQL = "Zl_病人信息从表_Update(" & mlng病人ID & ",'监护人身份证号'," & strPar & ",'')"
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    If EMPIModifyPatiInfo(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, strMsg) = 0 Then
        gcnOracle.RollbackTrans
        blnTrans = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    mblnReturn = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function EMPIModifyPatiInfo(ByVal lngSys As Long, ByVal lngModule As Long, ByVal lngPatiID As Long, ByVal lngClinicID As Long, ByRef strMsg As String) As Integer
    If Not gobjPlugIn Is Nothing Then
        err.Clear: On Error Resume Next
        If gobjPlugIn.EMPI_ModifyPatiInfo(lngSys, lngModule, lngPatiID, 0, lngClinicID, strMsg) = 0 Then
            If err.Number = 0 Then
                strMsg = "当前启用了EMPI系统接口，但EMPI系统接口(EMPI_ModifyPatiInfo)未调用成功：" & strMsg
                EMPIModifyPatiInfo = 0
                Exit Function
            End If
        End If
        If err.Number <> 0 And err.Number <> 438 Then
            strMsg = "zlPlugIn 外挂部件执行 EMPI_ModifyPatiInfo 方法时出错：" & vbCrLf & err.Number & vbCrLf & err.Description
            EMPIModifyPatiInfo = 0
            Exit Function
        End If
        err.Clear: On Error GoTo 0
    End If
    EMPIModifyPatiInfo = 1
End Function

Private Function UpDate挂号信息(ByVal strInfoName As String, ByVal strValue As String)
'功能：更新 复诊，摘要，传染病上传，发病时间，发病地址  几个字段的值
'参数：strInfoName 字段名称,strValue 字段的值
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strTmp As String
    Dim strArr(4) As String
    Dim strInfo As String
    Dim objTmp As Object
    Dim datCur As Date
    
    If mblnNoSave Then Exit Function
    On Error GoTo errH
    '检查
    Set objTmp = txtE(I就诊摘要)
    strInfo = Trim(objTmp.Text)
    strInfo = Replace(strInfo, "'", "’")
    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "就诊摘要-内容太长(允许录入" & objTmp.MaxLength & "个字符或" & objTmp.MaxLength \ 2 & "个汉字)。"
        Exit Function
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    Set objTmp = txtE(I发病时间)
    strInfo = Trim(objTmp.Text)
    If Not IsDate(strInfo) And strInfo <> "" Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "发病时间-日期格式不对。"
        Exit Function
    ElseIf IsDate(strInfo) Then
        datCur = mdatCurDate
        If CDate(strInfo) > datCur Then
            objTmp.BackColor = &HC0C0FF
            objTmp.Tag = "发病时间应当小于当前时间。"
            Exit Function
        End If
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    strInfo = Trim(txtE(I发病地址).Text)
    strInfo = Trim(objTmp.Text)
    strInfo = Replace(strInfo, "'", "’")
    If zlCommFun.ActualLen(strInfo) > objTmp.MaxLength Then
        objTmp.BackColor = &HC0C0FF
        objTmp.Tag = "发病地址-内容太长(允许录入" & objTmp.MaxLength & "个字符或" & objTmp.MaxLength \ 2 & "个汉字)。"
        Exit Function
    Else
        objTmp.BackColor = vbWindowBackground
        objTmp.Tag = ""
    End If
    objTmp.Text = strInfo
    
    strSQL = "select 姓名,性别,年龄,民族,国籍,区域,籍贯,职业,出生日期,出生地点,身份证号,其他证件,婚姻状况,医疗付款方式," & _
        "家庭地址,家庭电话,家庭地址邮编,户口地址,户口地址邮编,合同单位id,工作单位,单位电话,单位邮编,联系人姓名,联系人关系," & _
        "联系人电话,联系人地址,Email,Qq,监护人 from 病人信息 where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTmp.EOF Then Exit Function
    
    strSQL = "Zl_病人信息_首页整理(" & mlng病人ID & ",'" & mstr门诊号 & "',"
    For i = 0 To rsTmp.Fields.Count - 1
        If rsTmp.Fields(i).Name = "出生日期" Then
            strTmp = IIf(IsNull(rsTmp!出生日期), "NULL,", "To_Date('" & Format(rsTmp!出生日期, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),")
        Else
            strTmp = IIf(IsNull(rsTmp.Fields(i).Value), "NULL,", "'" & rsTmp.Fields(i).Value & "',")
        End If
        strSQL = strSQL & strTmp
    Next
    strSQL = strSQL & "'" & mstr挂号单 & "'"
    
    strArr(0) = IIf(optInfo(opt复诊).Value, 1, 0)
    
    strTmp = Trim(txtE(I就诊摘要).Text)
    strArr(1) = IIf(strTmp = "", "NULL", "'" & strTmp & "'")
    
    strArr(2) = chkInfo.Value
    
    strTmp = Trim(txtE(I发病时间).Text)
    If strTmp <> "" Then
        strArr(3) = "To_Date('" & Format(strTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strArr(3) = "NULL"
    End If
    
    strTmp = Trim(txtE(I发病地址).Text)
    strArr(4) = IIf(strTmp = "", "NULL", "'" & strTmp & "'")
    
    Select Case strInfoName
    Case "复诊"
        strArr(0) = strValue
    Case "摘要"
        If strValue <> "" Then
            strArr(1) = "'" & strValue & "'"
        Else
            strArr(1) = "NULL"
        End If
    Case "传染病上传"
        strArr(2) = strValue
    Case "发病时间"
        If strValue <> "" Then
            strArr(3) = "To_Date('" & Format(strValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strArr(3) = "NULL"
        End If
    Case "发病地址"
        strArr(4) = IIf(strValue = "", "NULL", "'" & strValue & "'")
    End Select
    strSQL = strSQL & "," & strArr(0) & "," & strArr(1) & "," & strArr(2) & "," & strArr(3) & "," & strArr(4) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If "摘要" = strInfoName Then
        strSQL = IIf(strValue = "", "NULL", strValue)
        RaiseEvent UpdatePatiInfo("", "", "", strSQL)
    End If
    mblnOK = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub UpDate病历(ByVal intIdx As Integer)
'功能：保存病历
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim lng病历ID As Long
    Dim strValue As String
    Dim strTmp As String
 
    On Error GoTo errH
    If mblnNoSave Then Exit Sub
    If Not mblnChange Then Exit Sub
    If rtfEdit(intIdx).BackColor = DColor Then
        Exit Sub
    End If
    strValue = Trim(rtfEdit(intIdx).Text)
 
    mrsMainInfo.Filter = "控件名='rtfEdit' and Index=" & intIdx
    
    strValue = Replace(strValue, "'", "’")
    If zlCommFun.ActualLen(strValue) > 4000 Then
        rtfEdit(intIdx).BackColor = &HC0C0FF
        strTmp = mrsMainInfo!信息名 & "-内容太长(允许录入4000个字符或2000个汉字)。"
        mrsMainInfo.Update "ErrInfo", strTmp
        Exit Sub
    Else
        rtfEdit(intIdx).BackColor = vbWindowBackground
        mrsMainInfo.Update "ErrInfo", ""
    End If
    
    If mrsMainInfo!信息原值 & "" = strValue Then
        Exit Sub
    End If

    arrSQL = Array()
    If mlng病历ID = 0 Then lng病历ID = zlDatabase.GetNextId("电子病历记录")
    
    Call GetSQLOutDoc(arrSQL, lng病历ID)
    If UBound(arrSQL) = -1 Then Exit Sub
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
  
    If ReadRTFData(lng病历ID) = False Then GoTo errH
    If SaveRTFData(lng病历ID) = False Then GoTo errH
      
    gcnOracle.CommitTrans: blnTrans = False
    
    mrsMainInfo!信息原值 = strValue
    mrsMainInfo.Update
 
    If mlng病历ID = 0 Then Call LoadDocData
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UpDate结构化地址(ByVal intIdx As Integer)
'功能：更新结构化地址
    Dim strSQL As String
    Dim strValue As String
    Dim blnTrans As Boolean
    Dim strSQLTmp As String
On Error GoTo errH
    If mblnNoSave Then Exit Sub
    If Not mblnUpdate Then Exit Sub
    strValue = Trim(PatiAddress(intIdx).Value)
    mrsMainInfo.Filter = "控件名='PatiAddress' and Index=" & intIdx

    If mrsMainInfo!信息原值 & "" = strValue Then
        Exit Sub
    End If

    If mblnStructAdress Then
        If PatiAddress(intIdx).Value <> "" Then
           strSQL = "zl_病人地址信息_update(1," & mlng病人ID & ",NULL," & Decode(intIdx, PT_出生地点, 1, PT_户口地址, 4, PT_家庭地址, 3) & ",'" & PatiAddress(intIdx).value省 & "','" & _
               PatiAddress(intIdx).value市 & "','" & PatiAddress(intIdx).value区县 & "','" & PatiAddress(intIdx).value乡镇 & "','" & _
               PatiAddress(intIdx).value详细地址 & "','" & PatiAddress(intIdx).Code & "')"
        Else
           strSQL = "zl_病人地址信息_update(2," & mlng病人ID & ",NULL," & Decode(intIdx, PT_出生地点, 1, PT_户口地址, 4, PT_家庭地址, 3) & ")"
        End If
    End If
    
    strSQLTmp = "Zl_病人信息_更新信息(" & mlng病人ID & ",'" & Decode(intIdx, PT_出生地点, "出生地点", PT_户口地址, "户口地址", PT_家庭地址, "家庭地址") & "','" & PatiAddress(intIdx).Value & "')"
    
On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call zlDatabase.ExecuteProcedure(strSQLTmp, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    mrsMainInfo!信息原值 = strValue
    mrsMainInfo.Update
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub UpDate身份证号()
'功能：身份证号
'功能：intType 0-身份证号，1－身份证号状态
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean
    Dim strValue As String
    Dim intIndex As Integer
    Dim str身份证号 As String
    Dim str身份证号状态 As String
    Dim blnDo As Boolean
    Dim strPar As String, strMsg As String
    On Error GoTo errH
    
    If mblnNoSave Then Exit Sub
    mblnReturn = True
    intIndex = cboE(I身份证号).ListIndex
    
    arrSQL = Array()
    If intIndex = -1 Then
        strValue = cboE(I身份证号).Tag
    Else
        strValue = cboE(I身份证号).Text
    End If
    
    strPar = IIf(strValue = "", "null", "'" & strValue & "'")
    If intIndex = -1 Then
        If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人信息_更新信息(" & mlng病人ID & ",'身份证号'," & strPar & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'外籍身份证号',Null,Null)"
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'外籍身份证号'," & strPar & ",'')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人信息_更新信息(" & mlng病人ID & ",'身份证号',Null)"
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'身份证号状态',Null,Null)"
    Else
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息_更新信息(" & mlng病人ID & ",'身份证号',Null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'外籍身份证号',Null,Null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人信息从表_Update(" & mlng病人ID & ",'身份证号状态'," & strPar & ",Null)"
    End If
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If EMPIModifyPatiInfo(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, strMsg) = 0 Then
        gcnOracle.RollbackTrans
        blnTrans = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    mblnOK = True
    If intIndex = -1 Then
        If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
            mrsMainInfo.Filter = "信息名='身份证号'"
            mrsMainInfo.Update "信息原值", strValue
        Else
            mrsMainInfo.Filter = "信息名='外籍身份证号'"
            mrsMainInfo.Update "信息原值", strValue
        End If
        mrsMainInfo.Filter = "信息名='身份证号状态'"
        mrsMainInfo.Update "信息原值", ""
        
        cboE(I身份证号).Tag = strValue
        If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
            If mblnID加密 Then
                strValue = Mid(strValue, 1, 12) & String(Len(Mid(strValue, 13, 2)), "*") & Mid(strValue, 15)
            End If
        End If
        cboE(I身份证号).Text = strValue
    Else
        mrsMainInfo.Filter = "信息名='身份证号'"
        mrsMainInfo.Update "信息原值", ""
        mrsMainInfo.Filter = "信息名='外籍身份证号'"
        mrsMainInfo.Update "信息原值", ""
        mrsMainInfo.Filter = "信息名='身份证号状态'"
        mrsMainInfo.Update "信息原值", strValue
    End If
    cboE(I身份证号).ToolTipText = ""
    cboE(I身份证号).BackColor = vbWindowBackground
    mblnReturn = False
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtE_GotFocus(Index As Integer)
    If Index <> I就诊摘要 Then
        Call zlControl.TxtSelAll(txtE(Index))
    ElseIf txtE(Index).SelLength = 0 Then
        Call zlControl.TxtSelAll(txtE(Index))
    End If
    Call SetCurCtlInfo(TypeName(txtE(Index)), "txtE", Index)
End Sub

Private Sub txtE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Index = I医学警示 Then
            txtE(I医学警示) = ""
        End If
    End If
End Sub

Private Sub txtE_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    Dim txtTmp As Object
    Dim strValue As String, str身份证号 As String
    If Index = I监护人身份证号 Then
        If KeyAscii = vbKeyReturn Then
            zlCommFun.PressKey vbKeyTab
        Else
            Set txtTmp = txtE(Index)
            If Not (KeyAscii >= 0 And KeyAscii < 32) Then
                If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                    If zlCommFun.ActualLen(txtTmp.Text) >= 18 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                            KeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(txtTmp.Text) Then
                            txtTmp.Text = "": txtTmp.Tag = ""
                        End If
                        If KeyAscii <> 0 Then
                            Select Case zlCommFun.ActualLen(txtTmp.Text)
                                Case 12
                                    txtTmp.Tag = txtTmp.Text & Chr(KeyAscii)
                                Case 13
                                    txtTmp.Tag = txtTmp.Tag & Chr(KeyAscii)
                            End Select
                        End If
                    End If
                Else
                    If Not (KeyAscii >= 0 And KeyAscii < 32) Then
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                            KeyAscii = 0
                        ElseIf zlCommFun.IsCharChinese(txtTmp.Text) Then
                            txtTmp.Text = "": txtTmp.Tag = ""
                        End If
                        If KeyAscii <> 0 Then
                            txtTmp.Tag = txtTmp.Text & Chr(KeyAscii)
                        End If
                    End If
                End If
            End If
        End If
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = I区域 Or Index = I籍贯) And txtE(Index).Text <> "" Then
            '输入区域数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 区域 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2]) And Nvl(级数, 0) < 3" & _
                " Order by 编码"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(Index = I区域, "区域", "籍贯"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!名称
            End If
            txtE(Index).SetFocus
        ElseIf (Index = I出生地点 Or Index = I家庭地址 Or Index = I户口地址) And txtE(Index).Text <> "" Then
            '输入地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                " Where (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!名称
            End If
            txtE(Index).SetFocus
            
        ElseIf Index = I家庭邮编 Or Index = I户口邮编 Or Index = I单位邮编 Then
            If ((Not IsNumeric(txtE(Index).Text)) Or Len(txtE(Index).Text) > 6 Or InStr(txtE(Index).Text, ".") > 0) And txtE(Index).Text <> "" Then
                If txtE(Index).Text <> "" Then
                    If zlCommFun.IsCharChinese(txtE(Index).Text) Then
                        strSQL = strSQL & " And A.名称 Like [1] "
                    Else
                        strSQL = strSQL & " And A.简码 Like [1] "
                    End If
                End If
                strSQL = "Select Rownum as ID,名称,简码,邮编  From 区域 A " & _
                "Where 邮编 is not null " & strSQL & " Order by 编码"
                vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "区域", False, "", "", False, _
                                        False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                                        UCase(txtE(Index).Text) & "%")
                '可以任意输入,不一定要匹配
                If Not rsTmp Is Nothing Then
                    txtE(Index).Text = rsTmp!邮编 & ""
                End If
                txtE(Index).SetFocus
            End If
        ElseIf Index = I单位名称 And txtE(Index).Text <> "" Then
            '输入工作单位
            strSQL = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (编码 Like [1] Or 简码 Like [2] Or 名称 Like [2])" & _
                " Order by 编码"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "工作单位", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, False, False, _
                UCase(txtE(Index).Text) & "%", gstrLike & UCase(txtE(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtE(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If mblnEdit合同单位 Then
                    mlng合同单位ID = Val(rsTmp!ID)
                Else
                    mlng合同单位ID = 0
                End If
                If txtE(I单位电话).Text = "" Then
                    txtE(I单位电话).Text = NVL(rsTmp!电话)
                End If
            Else
                txtE(Index).Tag = ""
                mlng合同单位ID = 0
            End If
            txtE(Index).SetFocus
        ElseIf Index = I就诊摘要 Then
            Call TxtKeyPress摘要(KeyAscii)
        ElseIf Index = I监护人身份证号 Then
            Call txtE_Validate(I监护人身份证号, blnCancel)
            If blnCancel Then
                mblnNoSave = True
                Exit Sub
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '非控制按键
        If Index = I医学警示 Then
            KeyAscii = 0
        ElseIf Index = I监护人身份证号 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            Else
                lblN(I监护人身份证号).Tag = "1"
            End If
        End If
        If KeyAscii = 39 Then KeyAscii = 0 '单引号蔽屏
        '选择快捷键
        If KeyAscii = Asc("*") Then
            '注意界面上要求CMD和对应TXT的Index相同
            On Error Resume Next
            strSQL = ""
            strSQL = cmdE(Index).Name
            err.Clear: On Error GoTo 0
            If strSQL <> "" Then
                KeyAscii = 0
                Call cmdE_Click(Index)
                Exit Sub
            End If
        End If
        
        '限制输入长度
        If txtE(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtE(Index).Text) > txtE(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '限制输入内容
        Select Case Index
            Case I家庭电话, I单位电话
                strMask = "1234567890-()"
            Case I发病时间
                strMask = "1234567890-: "
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = I医学警示 Then
            KeyAscii = 0
        ElseIf Index = I监护人身份证号 Then
            lblN(I监护人身份证号).Tag = "1"
        End If
    End If
End Sub

Private Sub TxtKeyPress摘要(KeyAscii As Integer)
    Dim objTxt As Object
    Set objTxt = txtE(I就诊摘要)
    If objTxt.Text <> "" Then
        If AbstractSelect(objTxt.Text) Then Exit Sub
    End If
End Sub

Private Sub cmdE_Click(Index As Integer)
'说明：注意界面上要求CMD和对应TXT的Index相同
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    Dim blnNoSave As Boolean
    
    '使用Lock的方式,不采用Enabled的方式
    If Not cmdE(Index).Enabled Or txtE(Index).Locked Then
        If txtE(Index).Enabled Then txtE(Index).SetFocus
        Exit Sub
    End If
    
    Select Case Index
        Case I出生地点, I家庭地址, I户口地址
            '选择地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""地区""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                Call txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!名称
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I单位名称
            '选择单位信息
            strSQL = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "合约单位", , , , , True, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""合约单位""数据，请先到合约单位管理中设置。", vbInformation, gstrSysName
                End If
                txtE(Index).Tag = ""
                If txtE(Index).Enabled Then txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!名称 & IIf(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If mblnEdit合同单位 Then
                    mlng合同单位ID = Val(rsTmp!ID)
                Else
                    mlng合同单位ID = 0
                End If
                If txtE(I单位电话).Text = "" Then
                    txtE(I单位电话).Text = NVL(rsTmp!电话)
                End If
                If txtE(Index).Enabled Then txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I区域, I籍贯
            '选择区域数据
            strSQL = "Select 1  From 区域 Where Nvl(级数,0)<>0 And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.RecordCount > 0 Then blnLevel = True
            
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            If blnLevel Then
                strSQL = _
                        "Select Id, 上级id, Id 编码, 名称, 简码, 末级" & vbNewLine & _
                        "From (Select Rpad(编码, 15, '0') As Id, Rpad(Substr(编码, 1, Decode(Nvl(级数, 0), 0, 0, 1, 2, 4)), 15, '0') As 上级id, 名称, 简码," & vbNewLine & _
                        "              Decode(Nvl(级数, 0), 2, 1, 3, 1, 0) As 末级" & vbNewLine & _
                        "       From 区域" & vbNewLine & _
                        "       Where Nvl(级数, 0) < 3" & vbNewLine & _
                        "       Order By 编码)" & vbNewLine & _
                        "Start With 上级id Is Null" & vbNewLine & _
                        "Connect By Prior Id = 上级id"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "区域", , , , , , , vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            Else
                strSQL = "Select Rownum as ID,编码,名称,简码 From 区域 Order by 编码"
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""区域""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtE(Index).SetFocus
                blnNoSave = True
            Else
                txtE(Index).Text = rsTmp!名称
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I医学警示
            '选择医学警示
            On Error GoTo errH
            vPoint = zlControl.GetCoordPos(txtE(Index).Container.hwnd, txtE(Index).Left, txtE(Index).Top)
            strSQL = "Select Rownum ID,编码,名称,简码 From 医学警示 Order by 编码"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtE(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""医学警示""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtE(Index).SetFocus
                blnNoSave = True
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!名称
                    rsTmp.MoveNext
                Wend
                txtE(Index).Text = Mid(strResult, 2)
                txtE(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case I发病时间
            dtpDate.ZOrder
            If IsDate(txtE(I发病时间).Text) Then
                 dtpDate.Value = CDate(txtE(I发病时间).Text)
             Else
                 dtpDate.Value = mdatCurDate
             End If
             dtpDate.Visible = True
             dtpDate.SetFocus
        Case Else
            blnNoSave = True
    End Select
    
    If Not blnNoSave Then
        Call UpDateInfo(txtE(Index).Text, "txtE", Index)
    End If
    Set mobjCtl = txtE(Index)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    '取值
    If IsDate(txtE(I发病时间).Text) Then
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtE(I发病时间).Text, "yyyy-MM-dd HH:mm"), 12, 5)
    Else
        strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(mdatCurDate, "yyyy-MM-dd HH:mm"), 12, 5)
    End If
    
    txtE(I发病时间).Text = strDate
    dtpDate.Tag = ""
    txtE(I发病时间).SetFocus
    dtpDate.Visible = False
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        txtE(I发病时间).SetFocus
        dtpDate.Tag = ""
        dtpDate.Visible = False
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'''
    Dim strDate As String
    
    With vsAller
        Select Case Col
            Case AI_过敏时间
                strDate = GetFullDate(.TextMatrix(Row, Col), False)
                If IsDate(strDate) Then
                    .TextMatrix(Row, Col) = strDate
                End If
        End Select
        Call vsAller_AfterRowColChange(-1, -1, .Row, .Col)
    End With
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    With vsAller
        If NewCol = AI_过敏药物 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(.TextMatrix(NewRow, AI_过敏药物)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int性别 As Integer
    Dim vPoint As POINTAPI
    
    With vsAller
        If Not gobjPass Is Nothing Then
            If optInfo(opt过敏源).Value Then
                strSQL = gobjPass.zlPassInputAllergy()
                If InStr(strSQL, ";") > 0 Then
                    Call SetAllerInput(Row, , strSQL)
                    Call AllerEnterNextCell
                End If
            Else
                If mstr性别 Like "*男*" Then
                    int性别 = 1
                ElseIf mstr性别 Like "*女*" Then
                    int性别 = 2
                End If
                
                strSQL = _
                    " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                    " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                    " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                    " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                    IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
                    End If
                Else
                    Call SetAllerInput(Row, rsTmp)
                    Call AllerEnterNextCell
                End If
            End If
        Else
            If mstr性别 Like "*男*" Then
                int性别 = 1
            ElseIf mstr性别 Like "*女*" Then
                int性别 = 2
            End If
            If optInfo(opt药品目录).Value Then
                strSQL = _
                    " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
                    " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
                    " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
                    " From 诊疗分类目录 Where 类型 IN (1,2,3) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                    " Union All" & _
                    " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
                    " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
                    IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
            Else
                strSQL = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Order By 编码"
                vPoint = zlControl.GetCoordPos(vsAller.hwnd, vsAller.Left, vsAller.CellTop)
                Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "过敏源", , , , , True, True, vPoint.X, vPoint.Y, vsAller.Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    If optInfo(opt药品目录).Value Then
                        MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
                    Else
                        MsgBox "没有过敏源数据可以选择。", vbInformation, gstrSysName
                    End If
                End If
            Else
                Call SetAllerInput(Row, rsTmp)
                Call AllerEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    If vsAller.Editable = flexEDNone Then Exit Sub
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AI_过敏药物) <> "" Then
                If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call UpDateAller
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    If vsAller.Editable = flexEDNone Then Exit Sub
    If 39 = KeyAscii Then KeyAscii = 0 '单引号
    With vsAller
        If KeyAscii = vbKeySpace Then   'Space
            If .Col = AI_过敏药物 And Not gobjPass Is Nothing And optInfo(opt过敏源).Value Then KeyAscii = 0: Exit Sub
        End If
        If KeyAscii = 13 Then
             KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AI_过敏药物 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If 39 = KeyAscii Then KeyAscii = 0
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    With vsAller
        If Col = AI_过敏时间 Then
            If KeyAscii = 13 Then
                .Col = .Col + 1
                .ShowCell Row, Col
                .Col = .Col - 1
            End If
        ElseIf Col = AI_过敏药物 Then
            If KeyAscii <> 13 Then
                If Not gobjPass Is Nothing And optInfo(opt过敏源).Value Then KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAller
        If Col = AI_过敏药物 Or Col = AI_过敏时间 Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        End If
    End With
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AI_过敏反应 And Trim(vsAller.TextMatrix(Row, AI_过敏药物)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    Dim curDate As Date
    Dim strDate As String
    
    With vsAller
        If Col = AI_过敏药物 Then
            If .EditText = "" Then
                If .Cell(flexcpData, Row, Col) <> "" Then
                    If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem .Row
                        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
                    Else
                        .EditText = .Cell(flexcpData, Row, Col)
                    End If
                End If
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                If LenB(StrConv(.EditText, vbFromUnicode)) > 60 Then
                    MsgBox "药物名称不能超过30个汉字的长度。", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                strInput = UCase(.EditText)
                If mstr性别 Like "*男*" Then
                    int性别 = 1
                ElseIf mstr性别 Like "*女*" Then
                    int性别 = 2
                End If
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                If optInfo(opt药品目录) Then
                    strSQL = _
                        " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                        " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                        " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                        " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                        " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                        IIf(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                        Decode(mint简码, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.编码"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", int性别, mint简码 + 1)
                Else
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSQL = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where 名称 Like [1] Order By 编码"
                    Else
                        If mint简码 = 1 Then
                            strSQL = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where zlWbCode(名称) Like [1] Order By 编码"
                        Else
                            strSQL = "Select Rownum As ID, 编码, 名称, 简码 From 过敏源 Where 简码 Like [1] Order By 编码"
                        End If
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏源", False, "", "", False, _
                        False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        gstrLike & UCase(strInput) & "%")
                End If
                If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AI_过敏时间 Then
            If .EditText <> "" Then
                strDate = GetFullDate(.EditText, False)
                If IsDate(strDate) Then
                    curDate = mdatCurDate
                    If CDate(strDate) > curDate Then
                        MsgBox "您输入的日期不能大于当前时间。当前时间：" & Format(curDate, "yyyy-mm-dd") & "。"
                        Cancel = True
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    .EditText = Format(strDate, "yyyy-MM-dd")
                Else
                    MsgBox "请输入正确的过敏时间，例如：""2012-12-21""或""121221""。"
                    Cancel = True
                End If
            End If
        Else
            If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                MsgBox "过敏反应不能超过50个汉字的长度。", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Function DiagCellEditable(ByRef vsDiagTmp As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    Dim bln西医 As Boolean
    Dim blnJudge As Boolean
    Dim dtTmp As DiagType
    Dim lng出院Row As Long
    
    If lngRow < 0 Then
        Exit Function
    End If
    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        '隐藏列不可编辑
        If .ColHidden(lngCol) Then Exit Function
        '必须先输入诊断描述，才能输入其他列(部分逻辑属于新增）
        If .TextMatrix(lngRow, DI_诊断描述) = "" Then
            If Not (lngCol = DI_Del Or lngCol = DI_诊断描述) Then Exit Function
        Else
            If lngCol <> DI_关联 And lngCol <> DI_Del And lngCol <> DI_增加 Then
                '关联医嘱不可编辑
                If .TextMatrix(lngRow, DI_医嘱IDs) <> "" Then Exit Function
            End If
        End If
        DiagCellEditable = True
    End With
End Function

Private Sub DiagKeyDown(ByRef vsDiag As VSFlexGrid, ByRef intKeyCode As Integer, ByRef intShift As Integer)
'vsDiagXY_KeyDown事件，vsDiagZY_KeyDown事件
    Dim i As Long, j As Long
    Dim dtCurRow As DiagType, lngRow As Long
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        If intKeyCode = vbKeyF4 Then
            If .Col = DI_诊断描述 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf intKeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, DI_诊断描述) <> "" Or .Rows = .FixedRows + 1 Then
                If .TextMatrix(.Row, DI_诊断描述) = "" Then Exit Sub
                If Not DiagCellEditable(vsDiag, .Row, DI_诊断描述) Then Exit Sub
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '删除主/次要诊断后调用外挂接口
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        Call gobjPlugIn.DiagnosisDeleted(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(.TextMatrix(.Row, DI_诊断ID)), .TextMatrix(.Row, DI_诊断描述))
                        Call zlPlugInErrH(err, "DiagnosisDeleted")
                        err.Clear: On Error GoTo 0
                    End If
                    dtCurRow = Val(.TextMatrix(.Row, DI_诊断分类))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, DI_诊断分类) = dtCurRow
                    '下面的同类诊断数据上移
                    If .TextMatrix(.Row, DI_诊断类型) = "" Or .Rows <> .FixedRows + 1 Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            For i = .Row + 1 To .Rows - 1
                                '数据移动到前面一行,若当前行为另一个分类，看前面一行是否是另一个分类的起始行，不是，则删除行
                                If .TextMatrix(i, DI_诊断分类) <> "" Then
                                    If .TextMatrix(i - 1, DI_诊断分类) = "" Then .RemoveItem i - 1
                                    Exit For
                                End If
                                '数据移动到前面一行
                                For j = .FixedCols To .Cols - 1
                                    .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                    .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                Next
                                .RowData(i - 1) = .RowData(i)
                                '最后一行删除
                                If i = .Rows - 1 Then
                                    .RemoveItem i: Exit For
                                End If
                            Next
                        End If
                    End If
                    Call UpDateDiag(vsDiag)
                End If
            ElseIf .TextMatrix(.Row, DI_诊断类型) = "" Or .Rows <> .FixedRows + 1 Then
                .RemoveItem .Row
            End If
            '设置诊断相关信息
            '如果填写了发病时间，则下面的发病时间则不允许填写了
'            If gclsPros.FuncType <> f诊断选择 Then Call SetDiagReletedInfo(vsDiag)
        ElseIf intKeyCode = vbKeyInsert Then '新增行
            lngRow = .Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, DI_诊断分类) = .TextMatrix(lngRow - 1, DI_诊断分类)
            .TextMatrix(lngRow, DI_诊断类型) = .TextMatrix(lngRow - 1, DI_诊断类型)
            .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
            .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
            .Row = lngRow: .Col = DI_诊断编码
            .ShowCell .Row, .Col
        ElseIf intKeyCode > 127 Then
            '解决直接输入汉字的问题
            Call DiagKeyPress(vsDiag, intKeyCode)
        End If
    End With
End Sub

Private Sub DiagKeyPress(ByRef vsDiag As VSFlexGrid, ByRef intKeyAscii As Integer)
'vsDiagXY_KeyPress事件，vsDiagZY_KeyPress事件
    If vsDiag.Editable = flexEDNone Then Exit Sub
    If intKeyAscii = 39 Then intKeyAscii = 0 '单引号蔽屏
    With vsDiag
        If intKeyAscii = vbKeyReturn Then
            intKeyAscii = 0
            Call EnterNextCellDiag(vsDiag)
        Else
            If .Col <> DI_是否疑诊 Then
                If Not DiagCellEditable(vsDiag, .Row, .Col) Then Exit Sub
            End If
            Select Case .Col
                Case DI_是否疑诊
                    If intKeyAscii <> vbKeySpace Then Exit Sub
                    intKeyAscii = 0
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", IIf(.Col = DI_是否疑诊, "？", "√"), "")
                Case DI_诊断编码, DI_诊断描述, DI_中医证候, DI_ICD附码 '西医中医证候隐藏,中医无ICD附码隐藏
                    If intKeyAscii = Asc("*") Then
                        intKeyAscii = 0
                        Call DiagCellButtonClick(vsDiag, .Row, .Col)
                    Else
                        .ComboList = "" '使按钮状态进入输入状态
                    End If
            End Select
        End If
    End With
End Sub

Private Sub EnterNextCellDiag(ByRef vsDiagTmp As VSFlexGrid)
    Dim i As Long, j As Long
    
    With vsDiagTmp
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, DI_诊断描述) To DI_Del
                If Not .ColHidden(j) Then
                    If DiagCellEditable(vsDiagTmp, i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= DI_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > DI_Del And .TextMatrix(.Rows - 1, DI_诊断描述) <> "" Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, DI_诊断分类) = .TextMatrix(.Rows - 2, DI_诊断分类)
            .TextMatrix(.Rows - 1, DI_诊断类型) = .TextMatrix(.Rows - 2, DI_诊断类型)
            .ShowCell i, DI_诊断描述
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub DiagCellButtonClick(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'vsDiagZY_CellButtonClick事件，vsDiagXY_CellButtonClick事件
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngCurRow As Long
    Dim bln西医 As Boolean
    
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        Select Case lngCol
            Case DI_诊断描述, DI_诊断编码
                If optInfo(opt诊断).Value Then
                    '按诊断输入:中医部份，一个诊断可能属于多个分类
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(bln西医, "1", "2"), mlng科室ID, , True, False)
                Else
                    'B-中医疾病编码，7-损伤中毒：Y-损伤中毒的外部原因；6-病理诊断：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(bln西医, "D", "B"), mlng科室ID, mstr性别, True, True, , glngSys)
                End If
                If Not rsTmp Is Nothing Then
                    Call SetDiagInput(vsDiag, lngRow, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                    zlControl.ControlSetFocus vsDiag, True
                End If
            Case DI_中医证候
                If optInfo(opt诊断).Value Then
                    '按诊断输入:先查是否有对应
                    If Set中医证候(lngRow, Val(.TextMatrix(lngRow, DI_诊断ID))) Then
                        zlControl.ControlSetFocus vsDiag, True
                        Exit Sub
                    End If
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, mstr性别, True, , , glngSys)
                Else
                    'Z-中医疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng科室ID, mstr性别, True, , , glngSys)
                End If
                If Not rsTmp Is Nothing Then
                    Call Set中医证候(lngRow, 0, rsTmp)
                    Call EnterNextCellDiag(vsDiag)
                    zlControl.ControlSetFocus vsDiag, True
                End If
            Case DI_增加
                If Not .Cell(flexcpPicture, lngRow, DI_增加) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyInsert, 0)
                End If
            Case DI_Del
                If Not .Cell(flexcpPicture, lngRow, DI_Del) Is Nothing Or Not .CellButtonPicture Is Nothing Then
                    Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
                End If
        End Select
    End With
End Sub

Private Sub SetDiagInput(ByRef vsDiagTmp As VSFlexGrid, ByVal lngRow As Long, rsInput As ADODB.Recordset, Optional bln附码 As Boolean)
'功能：处理诊断项目的输入
'      bln附码=是否是附码输入
    Dim str性别 As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String, bln分化程度 As Boolean
    Dim bln西医 As Boolean, blnRCodeIn As Boolean
    Dim lngTmpRow As Long, lng出院Row As Long
    Dim lng原诊断ID As Long, int诊断次序 As Integer
    
    With vsDiagTmp
        bln西医 = .Name = "vsDiagXY"
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                '不是单独的附码输入
                If Not bln附码 Then
                    If i > 1 Then
                        '最后一种诊断类别（中医：其他诊断，西医：损伤中毒）选择多条时的处理
                        lng原诊断ID = 0
                        If lngRow = .Rows - 1 Then
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, DI_诊断分类) = .TextMatrix(lngRow, DI_诊断分类)
                            .TextMatrix(.Rows - 1, DI_诊断类型) = .TextMatrix(lngRow, DI_诊断类型)
                        End If
                        '确定当前显示行
                        If Val(.TextMatrix(lngRow + 1, DI_诊断分类)) = Val(.TextMatrix(lngRow, DI_诊断分类)) Then
                            For j = lngRow + 1 To .Rows - 1
                                If Val(.TextMatrix(j, DI_诊断分类)) = Val(.TextMatrix(lngRow, DI_诊断分类)) Then
                                    lngRow = j
                                    If .TextMatrix(j, DI_诊断描述) = "" Then Exit For
                                Else
                                    Exit For
                                End If
                            Next
                            If .TextMatrix(lngRow, DI_诊断描述) <> "" Then
                                lngRow = lngRow + 1: .AddItem "", lngRow
                                .TextMatrix(lngRow, DI_诊断分类) = .TextMatrix(lngRow - 1, DI_诊断分类)
                                .TextMatrix(lngRow, DI_诊断类型) = .TextMatrix(lngRow - 1, DI_诊断类型)
                            End If
                        Else
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, DI_诊断分类) = .TextMatrix(lngRow - 1, DI_诊断分类)
                            .TextMatrix(lngRow, DI_诊断类型) = .TextMatrix(lngRow - 1, DI_诊断类型)
                        End If
                    Else
                        lng原诊断ID = Val(.TextMatrix(lngRow, DI_诊断ID))
                    End If
                    
                    .TextMatrix(lngRow, DI_诊断编码) = rsInput!编码 & ""
                    .TextMatrix(lngRow, DI_诊断描述) = rsInput!名称
                    .Cell(flexcpData, lngRow, DI_诊断描述) = rsInput!名称 & ""  '保存原名
                    .Cell(flexcpData, lngRow, DI_诊断编码) = rsInput!编码 & ""
                    .TextMatrix(lngRow, DI_诊断ID) = rsInput!诊断id & ""
                    .TextMatrix(lngRow, DI_疾病ID) = rsInput!疾病id & ""
                    .TextMatrix(lngRow, DI_疾病编码) = rsInput!疾病编码 & ""
                    .TextMatrix(lngRow, DI_疾病类别) = rsInput!疾病类别 & ""
                End If
                If Not bln附码 Then .TextMatrix(lngRow, DI_固定附码) = IIf(Not IsNull(rsInput!附码), "1", "")
                .TextMatrix(lngRow, DI_ICD附码) = IIf(bln附码, rsInput!编码 & "", rsInput!附码 & "")
                .TextMatrix(lngRow, DI_附码ID) = IIf(bln附码, rsInput!项目ID & "", rsInput!附码ID & "")
                
                If Not bln西医 Then
                    '中医根据疾病诊断参考取证候
                    Call Set中医证候(lngRow, Val(.TextMatrix(lngRow, DI_诊断ID)))
                End If
                
                If CreatePlugInOK(p门诊医生站, -1) Then
                    int诊断次序 = 0
                    int诊断次序 = IIf(lngRow = .FixedRows, -1, -2)
                    On Error Resume Next
                    Select Case int诊断次序
                        Case -1
                            Call gobjPlugIn.DiagnosisEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, DI_诊断描述), lng原诊断ID)
                            Call zlPlugInErrH(err, "DiagnosisEnter")
                        Case -2
                            Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p门诊医生站, mlng病人ID, mlng挂号ID, Val(rsInput!项目ID), .TextMatrix(lngRow, DI_诊断描述), lng原诊断ID)
                            Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End Select
                    err.Clear: On Error GoTo errH
                End If
                
                rsInput.MoveNext
            Next
        Else
            If Not bln附码 Then
                .TextMatrix(lngRow, DI_诊断描述) = .EditText
              
                .Cell(flexcpData, lngRow, DI_诊断描述) = .TextMatrix(lngRow, DI_诊断描述)
                .TextMatrix(lngRow, DI_诊断编码) = ""
                 .Cell(flexcpData, lngRow, DI_诊断描述) = ""
                .TextMatrix(lngRow, DI_诊断ID) = ""
                .TextMatrix(lngRow, DI_疾病ID) = ""
                .TextMatrix(lngRow, DI_证候ID) = ""
           
            Else
                .TextMatrix(lngRow, DI_固定附码) = ""
                .TextMatrix(lngRow, DI_ICD附码) = ""
                .TextMatrix(lngRow, DI_附码ID) = ""
            End If
        End If
        .Cell(flexcpForeColor, .FixedRows, DI_是否疑诊, .Rows - 1, DI_是否疑诊) = vbRed
        .Cell(flexcpBackColor, .FixedRows, DI_诊断编码, .Rows - 1, DI_诊断编码) = GRD_UNEDITCELL_COLOR      '灰蓝色
        
        '设置诊断相关信息
        Call SetDiagReletedInfo(vsDiagTmp, lngRow)
        If optInfo(opt复诊).Value = False Then
            If PatiReSeeDoctor Then
                If MsgBox("病人就诊科室、医生、诊断与上次相同，要标记为复诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    optInfo(opt复诊).Value = True
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'cmdMakeLog事件
Private Sub MakeLog()
'功能：将诊断描述加至摘要中
    Dim strLog As String, i As Long
    Dim strTmp As String
 
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, DI_诊断描述) & IIf(.TextMatrix(i, DI_是否疑诊) <> "", "(？)", "")
            End If
        Next
    End With

    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, DI_诊断描述) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, DI_诊断描述) & IIf(.TextMatrix(i, DI_是否疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        With txtE(I就诊摘要)
            If .SelStart = 0 And .SelLength = Len(.Text) Then
                .SelStart = Len(.Text)
            End If
            i = .SelStart
            .SelText = Mid(strLog, 2)
            .SelStart = i
            .SelLength = Len(Mid(strLog, 2))
            .SetFocus
            strTmp = .Text
        End With
        
        Call UpDateInfo(strTmp, "txtE", I就诊摘要)
    End If
End Sub


Private Function PatiReSeeDoctor() As Boolean
'功能：判断病人本次是否复诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL1 As String, strSQL2 As String
    Dim strSQL As String
    Dim vsTmp As VSFlexGrid
    
    On Error GoTo errH
    
    '医生、科室与上次相同：没有转诊、续诊的
    strSQL1 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=[2] And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSQL2 = "Select Max(ID) as ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
            " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
    strSQL2 = "Select 病人ID,执行人 as 医生,执行部门ID as 科室ID From 病人挂号记录 Where ID=(" & strSQL2 & ") And 转诊科室ID Is Null And 续诊科室ID Is Null"
    
    strSQL = "Select 1 From (" & strSQL1 & ") A,(" & strSQL2 & ") B Where A.病人ID=B.病人ID And A.医生=B.医生 And A.科室ID=B.科室ID"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID)
    If rsTmp.EOF Then Exit Function
    
    '主要诊断与上次相同
    With vsDiagXY
        If .TextMatrix(.FixedRows, DI_诊断描述) <> "" Then
            strSQL = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                    " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
            strSQL = "Select 1 From 病人诊断记录" & _
                " Where 病人ID=[1] And 主页ID=(" & strSQL & ")" & _
                " And 诊断类型=1 And 记录来源 IN(1,3) And 诊断次序=1" & _
                " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID, _
                Val(.TextMatrix(.FixedRows, DI_疾病ID)), Val(.TextMatrix(.FixedRows, DI_诊断ID)), .TextMatrix(.FixedRows, DI_诊断描述))
            If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
        End If
    End With
    
    If mbln中医 Then
        With vsDiagZY
            If .TextMatrix(.FixedRows, DI_诊断描述) <> "" Then
                strSQL = "Select Max(ID) as 主页ID From 病人挂号记录 Where 病人ID=[1] And 记录性质=1 And 记录状态=1" & _
                       " And 登记时间 =(Select Max(a.登记时间) From 病人挂号记录 A Where a.病人id=[1] And a.记录性质=1 And a.记录状态=1 And a.登记时间<(Select 登记时间 From 病人挂号记录 Where ID=[2])) "
                strSQL = "Select 1 From 病人诊断记录" & _
                    " Where 病人ID=[1] And 主页ID=(" & strSQL & ")" & _
                    " And 诊断类型=11 And 记录来源 IN(1,3) And 诊断次序=1" & _
                    " And (疾病ID=[3] And 疾病ID<>0 Or 诊断ID=[4] And 诊断ID<>0 Or 诊断描述=[5])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatiReSeeDoctor", mlng病人ID, mlng挂号ID, _
                    Val(.TextMatrix(.FixedRows, DI_疾病ID)), Val(.TextMatrix(.FixedRows, DI_诊断ID)), .TextMatrix(.FixedRows, DI_诊断描述))
                If Not rsTmp.EOF Then PatiReSeeDoctor = True: Exit Function
            End If
        End With
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDiagReletedInfo(ByRef vsDiagTmp As VSFlexGrid, Optional ByVal lngRow As Long = -1)
'‘’‘’‘’‘’‘’‘’‘’
    '如果填写了发病时间，则下面的发病时间则不允许填写了
End Sub

Private Function Set中医证候(ByVal lngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'返回：是否有对应关系
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    On Error GoTo errH
    
    With vsDiagZY
        If blnFreeInput Then
            .TextMatrix(lngRow, DI_证候ID) = ""
            .TextMatrix(lngRow, DI_证候编码) = ""
            .TextMatrix(lngRow, DI_中医证候) = .EditText
        Else
            '去掉已有的证候
            If .TextMatrix(lngRow, DI_诊断描述) Like "?*(?*)" Then
                strTmp = Mid(.TextMatrix(lngRow, DI_诊断描述), 1, InStrRev(.TextMatrix(lngRow, DI_诊断描述), "(") - 1)
            Else
                strTmp = .TextMatrix(lngRow, DI_诊断描述)
            End If
            
            If rsInput Is Nothing Then
                If lng诊断ID = 0 Then Exit Function
                strSQL = "Select Distinct A.证候序号 As ID, A.证候id As 项目id, B.编码, B.附码, A.证候名称 名称," & IIf(mint简码 = 0, "B.简码", "B.五笔码 As 简码") & ", B.说明" & vbNewLine & _
                            "From 疾病诊断参考 A, 疾病编码目录 B" & vbNewLine & _
                            "Where A.证候id = B.Id(+) And A.诊断id = [1] And A.证候名称 Is Not Null" & vbNewLine & _
                            "Order By A.证候序号"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsInput = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                If rsInput Is Nothing Then
                    If Not blnCancel Then Exit Function
                    If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, DI_中医证候)
                    Set中医证候 = True: Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, DI_证候ID) = NVL(rsInput!项目ID)
            .TextMatrix(lngRow, DI_证候编码) = NVL(rsInput!编码)
            If Not IsNull(rsInput!名称) Then
                .TextMatrix(lngRow, DI_诊断描述) = strTmp
                .Cell(flexcpData, lngRow, DI_诊断描述) = .TextMatrix(lngRow, DI_诊断描述)
                .TextMatrix(lngRow, DI_中医证候) = NVL(rsInput!名称)
                .Cell(flexcpData, lngRow, DI_中医证候) = .TextMatrix(lngRow, DI_中医证候)
                If .EditText <> "" Then .EditText = .TextMatrix(lngRow, DI_中医证候)
            End If
        End If
        Set中医证候 = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DiagAfterEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
    'vsDiagXY_AfterEdit事件,vsDiagZY_AfterEdit事件
    Dim bln西医 As Boolean
    
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        If lngCol = DI_诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                Call DiagKeyDown(vsDiag, vbKeyDelete, 0)
            End If
        End If
        Call DiagAfterRowColChange(vsDiag, -1, -1, .Row, .Col)
        zlControl.ControlSetFocus vsDiag, True
    End With
End Sub

Private Sub DiagAfterRowColChange(ByRef vsDiag As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngOldCol As Long, ByVal lngNewRow As Long, ByVal lngNewCol As Long)
'vsDiagZY_AfterRowColChange事件，vsDiagXY_AfterRowColChange事件
    Dim i As Long
    Dim bln西医 As Boolean
    Dim vPoint As POINTAPI
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiag.Editable = flexEDNone Then Exit Sub
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        '清除图片
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, DI_增加) = Nothing
            Set .Cell(flexcpPicture, i, DI_Del) = Nothing
        Next
        
        If Not DiagCellEditable(vsDiag, lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
             
            Select Case lngNewCol
                Case DI_诊断描述
                    .ComboList = "..."
                Case DI_增加, DI_Del
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = IIf(lngNewCol = DI_增加, imgButtonNew.Picture, imgButtonDel.Picture)
                Case DI_中医证候
                    If .TextMatrix(lngNewRow, DI_诊断描述) = "" Then
                        .ComboList = ""
                        .FocusRect = flexFocusLight
                    Else
                        .ComboList = "..."
                    End If
                Case Else
                    .ComboList = ""
            End Select
        End If
        If lngNewRow >= .FixedRows Then
            '显示图片
            If lngNewCol <> DI_增加 And .TextMatrix(lngNewRow, DI_诊断描述) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '下一行诊断为空则不能新增行
                    If Not (.TextMatrix(lngNewRow, DI_诊断分类) = .TextMatrix(lngNewRow + 1, DI_诊断分类) And .TextMatrix(lngNewRow + 1, DI_诊断描述) = "") Then
                         Set .Cell(flexcpPicture, lngNewRow, DI_增加) = imgButtonNew.Picture
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, DI_增加) = imgButtonNew.Picture
                End If
            End If
            '显示图片
            If lngNewCol <> DI_Del Then Set .Cell(flexcpPicture, lngNewRow, DI_Del) = imgButtonDel.Picture
        End If
        zlControl.ControlSetFocus vsDiag, True
    End With
End Sub

Private Sub DiagAfterUserResize(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'vsDiagZY_BeforeUserResize事件，vsDiagXY_BeforeUserResize事件
    If lngCol = DI_诊断描述 Then
        If vsDiagZY.ColWidth(DI_中医证候) < vsDiagXY.ColWidth(lngCol) Then
             vsDiagZY.ColHidden(DI_中医证候) = False
             vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol) - vsDiagZY.ColWidth(DI_中医证候)
        Else
             vsDiagZY.ColHidden(DI_中医证候) = True
             vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol)
        End If
    Else
         vsDiagZY.ColWidth(lngCol) = vsDiagXY.ColWidth(lngCol)
    End If
End Sub

Private Sub DiagBeforeUserResize(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagZY_BeforeUserResize事件，vsDiagXY_BeforeUserResize事件
    If lngCol = DI_增加 Or lngCol = DI_Del Or lngCol < DI_诊断描述 Then blnCancel = True
End Sub

Private Sub DiagClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_Click事件，vsDiagZY_Click事件
    Dim bln西医 As Boolean
    
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        If (.MouseCol = DI_增加 Or .MouseCol = DI_Del) And .MouseRow >= .FixedRows Then
            If .MouseCol = DI_增加 Then
                If .TextMatrix(.MouseRow, DI_诊断描述) = "" Or .TextMatrix(.MouseRow, 0) = IIf(bln西医, "出院诊断", "主要诊断") Then Exit Sub
            End If
            .Select .MouseRow, .MouseCol
            Call DiagCellButtonClick(vsDiag, .MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub DiagDblClick(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_DblClick事件，vsDiagZY_DblClick事件
    Call DiagKeyPress(vsDiag, vbKeySpace)
End Sub

Private Sub DiagGotFocus(ByRef vsDiag As VSFlexGrid)
'vsDiagXY_GotFocus事件，vsDiagZY_GotFocus事件
    Call SetCurCtlInfo(TypeName(vsDiag), vsDiag.Name)

    If vsDiag.Row >= vsDiag.FixedRows And vsDiag.Col >= vsDiag.FixedCols Then
        Call DiagAfterRowColChange(vsDiag, -1, -1, vsDiag.Row, vsDiag.Col)
    End If

    zlControl.ControlSetFocus vsDiag, True
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_GotFocus()
    Call DiagGotFocus(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
End Sub

Private Sub vsAller_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strInfo As String
 
    With vsAller
        If .Tag <> "" Then
            lngRow = Val(.Tag)
            If .MouseRow = lngRow Then
                strInfo = "存在两行相同的过敏记录。"
            End If
        End If
    End With
    Call zlCommFun.ShowTipInfo(vsAller.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strInfo As String
 
    With vsDiagXY
        If .Tag <> "" Then
            lngRow = Val(.Tag)
            If .MouseRow = lngRow Then
                strInfo = "存在两行相同诊断。"
            End If
        End If
    End With
    Call zlCommFun.ShowTipInfo(vsDiagXY.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagZY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Dim lngRow As Long
    Dim strTmp As String
    
    strInfo = vsDiagZY.Tag
    If InStr(strInfo, "|") <> 0 Then
        lngRow = Val(Split(strInfo, "|")(0))
        strInfo = Split(strInfo, "|")(1)
        With vsDiagZY
            If .MouseRow = lngRow Then
                strTmp = strInfo
            End If
        End With
    End If
    strInfo = strTmp
    Call zlCommFun.ShowTipInfo(vsDiagZY.hwnd, strInfo, True, True)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub DiagSetupEditWindow(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByVal lngEditWindow As Long, ByVal blnIsCombo As Boolean)
'vsDiagXY_SetupEditWindow事件，vsDiagZY_SetupEditWindow事件
    With vsDiag
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub DiagStartEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_StartEdit事件，vsDiagZY_StartEdit事件
    
    If Not DiagCellEditable(vsDiag, lngRow, lngCol) Then
        blnCancel = True
    ElseIf lngCol = DI_是否疑诊 Then
        blnCancel = True '不直接编辑
    End If
    
End Sub

Private Sub vsDiagXY_Validate(Cancel As Boolean)
    Call UpDateDiag(vsDiagXY)
End Sub

Private Sub vsDiagZY_Validate(Cancel As Boolean)
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub DiagValidateEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef blnCancel As Boolean)
'vsDiagXY_ValidateEdit事件，vsDiagZY_ValidateEdit事件
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnInputCancel As Boolean
    Dim int诊断输入 As Integer
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim bln西医 As Boolean
    Dim str性别 As String
    
    With vsDiag
        bln西医 = .Name = "vsDiagXY"
        Select Case lngCol
            Case DI_诊断描述, DI_诊断编码
                If bln西医 Then
                    strDiagType = "'D'"
                Else
                    strDiagType = IIf(optInfo(opt诊断).Value, "", "B")
                End If
                
                If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    .EditText = ""
                ElseIf .EditText = .Cell(flexcpData, lngRow, lngCol) Then
                    If mblnReturn Then Call EnterNextCellDiag(vsDiag)
                ElseIf .TextMatrix(lngRow, DI_诊断编码) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And .EditText Like "*" & .Cell(flexcpData, lngRow, lngCol) & "*" Then
                    '判断加了前缀后的名称是否存在其他的诊断编码
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别)
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, strDiagType, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID)
                    If rsTmp.RecordCount = 1 Then
                        Call SetDiagInput(vsDiag, lngRow, rsTmp)
                        .EditText = .Text
                    Else
                        '允许在标准的名称前后输入附加信息
                        '不处理.Cell(flexcpData, lngRow, lngCol)，以便修改内容时再次使用like判断
                        .TextMatrix(lngRow, DI_诊断描述) = .EditText
                    End If
                ElseIf .TextMatrix(lngRow, DI_诊断编码) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And mblnFreeInput Then
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别)
                    On Error GoTo errH
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInfo(opt疾病).Value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", strDiagType, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(lngRow, DI_诊断描述) = .EditText
                        Else
                             Call SetDiagInput(vsDiag, lngRow, rsTmp): .EditText = .Text
                        End If
                    End If
                Else
                    int诊断输入 = mint诊断输入
                    strInput = UCase(.EditText)
                    strSQL = GetMedInputSQL(IIf(bln西医, 0, 1), strInput, str性别)
                    '损伤中毒码：Y-损伤中毒的外部原因；病理诊断允许：M-肿瘤形态学编码；其它诊断：D-ICD-10疾病编码
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIf(optInfo(opt疾病).Value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", gstrLike & strInput & "%", strDiagType, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                        ElseIf Not (rsTmp Is Nothing) Then
                            Call SetDiagInput(vsDiag, lngRow, rsTmp)
                            .EditText = .Text
                        Else
                            '没有匹配成功再次当成自由录入
                            If int诊断输入 = 1 Or (int诊断输入 = 3 And (rsTmp Is Nothing) And mint险类 = 0) Then
                                Call SetDiagInput(vsDiag, lngRow, Nothing)
                                .EditText = .Text
                            Else
                                blnCancel = True
                            End If
                        End If
                    End If
                End If
      
                mblnReturn = False
            Case DI_中医证候
                If .EditText = "" And .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    .EditText = ""
                    '中医证候则清除备份数据
                    .Cell(flexcpData, lngRow, lngCol) = ""
                ElseIf .EditText = .Cell(flexcpData, lngRow, lngCol) Then
                    If mblnReturn Then Call EnterNextCellDiag(vsDiag)
                
                ElseIf .TextMatrix(lngRow, DI_诊断编码) <> "" And .Cell(flexcpData, lngRow, lngCol) <> "" And mblnFreeInput Then
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSQL = GetMedInputSQL(1, strInput, str性别, strDiagType)
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDiagType, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        If rsTmp Is Nothing Then
                            .TextMatrix(lngRow, DI_中医证候) = .EditText
                        Else
                            Call Set中医证候(lngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                Else
                    int诊断输入 = mint诊断输入
                    strInput = UCase(.EditText)
                    strDiagType = "Z"
                    strSQL = GetMedInputSQL(1, strInput, str性别, strDiagType)
                    If optInfo(opt疾病).Value Then
                        '按诊断输入:先查是否有对应
                        If Set中医证候(lngRow, Val(.TextMatrix(lngRow, DI_诊断ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = zlControl.GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, strInput & "%", gstrLike & strInput & "%", strDiagType, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnInputCancel Then '无匹配输入时,按任意输入处理,取消不同
                        blnCancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            blnCancel = True
                        Else
                            Call Set中医证候(lngRow, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                End If
                mblnReturn = False
            Case DI_发病时间
                If .EditText <> "" Then
                    strInput = GetFullDate(.EditText)
                    If IsDate(strInput) Then
                        .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    Else
                        MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                        blnCancel = True
                    End If
                End If
                If lngRow = .FixedRows Then
'                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病时间), IsDate(.EditText), True)
'                    Call SetCtrlLocked(gclsPros.CurrentForm.mskDateInfo(DC_发病日期), IsDate(.EditText), True)
                End If
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMedInputSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str性别 As String, Optional ByVal strOtherInfo As String) As String
'功能：获得查询首页输入查询的SQL
'参数：intType:获取的SQL类型,0-西医诊断，1-中医诊断，2-手术操作
'    strInput-查询条件，str性别--病人的性别
'    strOtherInfo:中医诊断-疾病编码种类
'返回：strsql--查询诊断的SQL

    Dim strSQL As String

    If mstr性别 Like "*男*" Then
        str性别 = "男"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "女"
    End If

    Select Case intType
        Case 0, 1 '西医诊断,中医诊断
            If intType = 0 And optInfo(opt诊断).Value Or intType = 1 And optInfo(opt诊断).Value And strOtherInfo <> "Z" Then
            '按诊断输入:一个诊断可能属于多个分类
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                End If
                strSQL = "Select A.Id, A.Id 项目ID, A.编码, Null 序号, Null 附码, Null 附码id, Null 附码名称, A.名称, A.说明, A.编者, B.简码, 0 疗效限制, 0 分娩," & vbNewLine & _
                                "              0 是否病人, Max(D.疾病id) 疾病id, A.Id 诊断id" & vbNewLine & _
                                "       From 疾病诊断目录 A, 疾病诊断别名 B, 疾病诊断对照 D" & vbNewLine & _
                                " Where A.ID=B.诊断ID And A.ID=D.诊断ID(+) And A.类别=" & IIf(intType = 0, 1, 2) & vbNewLine & _
                                " And B.码类=[5] And (" & strSQL & ")" & vbNewLine & _
                                " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                "Group By A.Id, A.编码, A.名称, A.说明, A.编者,B.简码"
                '读取诊断对应疾病编码附码
                strSQL = "Select distinct A.ID,A.项目ID, A.编码, B.序号, B.附码, Null 附码id, Null 附码名称, A.名称, A.说明, Null 编者,A.简码, A.疗效限制, A.分娩, A.是否病人," & vbNewLine & _
                                "       B.编码 疾病编码, B.Id 疾病id, B.类别 疾病类别, A.诊断id," & vbNewLine & _
                                "      Decode(a.名称, [6], 1, Decode(A.简码,[6],1,decode(A.编码,[6],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                                "      Decode(Substr(A.名称, 1, Length([6])), [6], 1, Decode(Substr(A.简码, 1, Length([6])),[6],1,decode(Substr(a.编码, 1, Length([6])),[6],1,NULL))) As 排序3ID" & _
                                " From (" & strSQL & ") A, 疾病编码目录 B, 疾病诊断科室 C, 疾病诊断科室 D" & vbNewLine & _
                                " Where A.疾病id = B.Id(+)" & vbNewLine & _
                                " And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And c.科室id(+)=[8]  And d.人员id(+) = [7]" & _
                                " Order By 排序1ID, 排序2ID, 排序3ID, A.编码"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "A.名称 Like [2]" '输入汉字时只匹配名称
                Else
                    strSQL = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]"
                End If
           
                strSQL = _
                    "Select A.Id,A.Id 项目ID, A.编码, A.序号, A.附码,Null 附码ID, Null 附码名称, A.名称, A.说明, Null 编者, A.分类id, " & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " as 简码,  A.疗效限制, A.分娩, C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别," & vbNewLine & _
                    "       Max(B.诊断id) 诊断id" & vbNewLine & _
                    "From 疾病编码目录 A, 疾病诊断对照 B, 疾病编码分类 C " & vbNewLine & _
                    "Where A.Id = B.疾病id(+) And A.分类id = C.Id(+)  And" & vbNewLine & _
                    " Instr([3],A.类别)>0 And (" & strSQL & ")" & _
                    IIf(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
                    " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    "Group By A.Id, A.编码, A.序号, A.附码, A.名称, A.说明, A.分类id," & IIf(mint简码 = 0, "A.简码", "A.五笔码") & ", A.疗效限制, A.分娩, A.类别,C.是否病人"
                 
                strSQL = "Select distinct A.Id,A.项目ID, A.编码, A.序号, A.附码,A.附码ID, A.附码名称, A.名称, A.说明, A.编者, A.分类id, A.简码,  A.疗效限制, A.分娩, A.是否病人,A.疾病编码, A.疾病id,A.疾病类别,A.诊断id, " & _
                        " Decode(a.名称, [6], 1, Decode(A.简码,[6],1,decode(a.编码,[6],1,NULL))) As 排序1ID," & vbNewLine & _
                        "    Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                        "   Decode(Substr(a.名称, 1, Length([6])), [6], 1, Decode(Substr(A.简码, 1, Length([6])),[6],1,decode(Substr(a.编码, 1, Length([6])),[6],1,NULL))) As 排序3ID" & vbNewLine & _
                        " From (" & strSQL & ") A, 疾病编码科室 C, 疾病编码科室 D " & _
                        " Where  c.疾病id(+) = a.Id And d.疾病id(+) = a.Id And c.科室id(+)=[8]  And d.人员id(+) = [7] " & _
                        " Order By 排序1ID, 排序2ID, 排序3ID, A.编码"
            End If
    End Select
    GetMedInputSQL = strSQL
End Function

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterUserResize(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mbln不更新诊断 = True
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
    mbln不更新诊断 = False
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_GotFocus()
    Call DiagGotFocus(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub DiagKeyPressEdit(ByRef vsDiag As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long, ByRef intKeyAscii As Integer)
    If intKeyAscii = 39 Then intKeyAscii = 0 '单引号蔽屏
    If intKeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mbln不更新诊断 = True
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
    mbln不更新诊断 = False
    Call UpDateDiag(vsDiagZY)
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngTop As Long, lngHeight As Long
    Dim lngH As Long
    Dim lngFrmWidth As Long
    Dim bln小字体 As Boolean
 
    On Error Resume Next
    If mint调用 = 1 Then
        stbThis.Visible = True
    Else
        stbThis.Visible = False
    End If
    vsc.Top = 0
    vsc.Height = Me.ScaleHeight - IIf(mint调用 = 1, stbThis.Height, 0)
    vsc.Left = Me.ScaleWidth - vsc.Width
    vsc.Visible = True
    
    Me.BackColor = &H808080  ' &HE0E0E0
    
    PicPanel(picPanel_基本信息).Top = mlngTopVsc
    
    bln小字体 = mbytSize = 9
 
    lngFrmWidth = vsc.Left - 200
    
    If bln小字体 Then
        PicPanel(picPanel_基本信息).Width = 9665
        If lngFrmWidth < 9665 Then
            PicPanel(picPanel_基本信息).Left = 0
        Else
            PicPanel(picPanel_基本信息).Left = 100 + (lngFrmWidth - 9665) / 2
        End If
        PicPanel(picPanel_基本信息).Height = 4600
        PicPanel(picPanel_快键病历).Height = 4500
        If mbln中医 Then
            PicPanel(picPanel_就诊信息).Height = 9300
        Else
            PicPanel(picPanel_就诊信息).Height = 8340
        End If
        PicPanel(picPanel_附加).Height = 5000
    Else
        PicPanel(picPanel_基本信息).Width = 12440
        If lngFrmWidth < 12440 Then
            PicPanel(picPanel_基本信息).Left = 0
        Else
            PicPanel(picPanel_基本信息).Left = 100 + (lngFrmWidth - 12440) / 2
        End If
        PicPanel(picPanel_基本信息).Height = 5300
        PicPanel(picPanel_快键病历).Height = 5400
        If mbln中医 Then
            PicPanel(picPanel_就诊信息).Height = 10500
        Else
            PicPanel(picPanel_就诊信息).Height = 9540
        End If
        PicPanel(picPanel_附加).Height = 5000
    End If
    
    lngWidth = PicPanel(picPanel_基本信息).Width
    
    lblN(lbl标题基本).Left = (lngWidth - lblN(lbl标题基本).Width) / 2 - 50
    lblN(lbl标题基本).Caption = "基本信息"
    lblN(lbl标题基本).Top = 500
    
    lblN(lbl标题就诊).Left = lblN(lbl标题基本).Left
    lblN(lbl标题就诊).Top = 100
    
    lblN(lbl标题病历).Left = lblN(lbl标题基本).Left
    lblN(lbl标题病历).Top = 100
    
    PicPanel(picPanel_快键病历).Top = PicPanel(picPanel_基本信息).Height + PicPanel(picPanel_基本信息).Top
    PicPanel(picPanel_快键病历).Width = lngWidth
    PicPanel(picPanel_快键病历).Left = PicPanel(picPanel_基本信息).Left

    picOutDoc.Width = lngWidth
    picOutDoc.Top = lblN(lbl标题就诊).Top + lblN(lbl标题就诊).Height + 200
    picOutDoc.Left = 0
    picOutDoc.Height = PicPanel(picPanel_快键病历).Height - picOutDoc.Top
    
    PicPanel(picPanel_就诊信息).Top = IIf(mblnDocInput, PicPanel(picPanel_快键病历).Height, 0) + PicPanel(picPanel_快键病历).Top
    PicPanel(picPanel_就诊信息).Left = PicPanel(picPanel_基本信息).Left
    PicPanel(picPanel_就诊信息).Width = lngWidth
    
    lngH = PicPanel(picPanel_基本信息).Height + PicPanel(picPanel_就诊信息).Height + IIf(mblnDocInput, PicPanel(picPanel_快键病历).Height, 0)
    lngH = lngH - Me.ScaleHeight
    
    vsc.Max = lngH \ (Screen.TwipsPerPixelY)
    
End Sub

Private Function AbstractSelect(ByVal strFind As String) As Boolean
'常用摘要选择器
    Dim blnCancle As Boolean
    Dim strRetrun As String
    Dim lngLeft As Long, lngTop As Long
    Dim strName As String
    
    Dim objTxt As Object
    Set objTxt = txtE(I就诊摘要)
    
    lngLeft = objTxt.Left + objTxt.Container.Left + 5800
    lngTop = objTxt.Top + objTxt.Container.Top - 210
    
    strRetrun = mobjKernel.ShowCommItem(Me, strFind, blnCancle, lngLeft, lngTop, 4)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "没有找到可用的就诊摘要。", vbInformation, Me.Caption
                Exit Function
            End If
            objTxt.Text = strFind
        Else
            objTxt.Text = strRetrun
        End If
        Call UpDateInfo(objTxt.Text, "txtE", I就诊摘要)
    End If
    AbstractSelect = blnCancle
End Function

Private Sub SetDocEditable()
'功能：快捷病历的可编辑性
    Dim blnDoc As Boolean
    Dim k As Long, i As Long
    
    If mblnDocInput Then
        blnDoc = mlng科室ID <> 0 And (mlng病历ID = 0 And mlng病历文件id <> 0 Or mlng病历ID <> 0 And mbln签名 = False) And (mlng执行状态 = 1 Or mlng执行状态 = 5)
        If blnDoc And mlng病历ID <> 0 And lblDoctor(1).Tag = "0" Then    '没有修改他人病历的权限
            blnDoc = mstr保存人 = UserInfo.姓名
        End If
        k = 0
        For i = 0 To rtfEdit.Count - 1
            rtfEdit(i).Locked = Not blnDoc Or InStr(rtfEdit(i).Tag, ",") > 0   '存在多行内容时(进行了全文编辑保存)，不允许再修改
            If rtfEdit(i).Locked = False Then
                rtfEdit(i).BackColor = vbWindowBackground
                k = k + 1
            Else
                rtfEdit(i).BackColor = DColor
            End If
        Next
        If mlng病历ID = 0 Or mlng病历ID <> 0 And mbln签名 = False Then
            cmdSign.Caption = "签名(&S)"
        Else
            cmdSign.Caption = "取消签名(&S)"
        End If
        cmdSign.Enabled = mlng科室ID <> 0 And (mlng病历ID = 0 And mlng病历文件id <> 0 Or mlng病历ID <> 0) And (mlng执行状态 = 1 Or mlng执行状态 = 5)
        
        If cmdSign.Enabled And mlng病历ID <> 0 And lblDoctor(1).Tag = "0" Then   '没有修改他人病历的权限
            cmdSign.Enabled = mstr保存人 = UserInfo.姓名
        End If
        cmdUpdate.Enabled = cmdSign.Enabled
        cmdImportEPRDemo.Enabled = cmdSign.Enabled
    End If
End Sub

Private Sub cboSpecificInfoKeyPress(ByRef intIndex As Integer, ByRef intKeyAscii As Integer)
'cboSpecificInfo_KeyPress事件
    Dim lngidx As Long
    Dim cboTmp As ComboBox
    If intKeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        Set cboTmp = cboE(intIndex)
        If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
            If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                If zlCommFun.ActualLen(cboTmp.Text) >= 18 Then
                    intKeyAscii = 0
                Else
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If intKeyAscii <> 0 Then
                        Select Case zlCommFun.ActualLen(cboTmp.Text)
                            Case 12
                                cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                            Case 13
                                cboTmp.Tag = cboTmp.Tag & Chr(intKeyAscii)
                        End Select
                    End If
                End If
            Else
                If Not (intKeyAscii >= 0 And intKeyAscii < 32) Then
                    intKeyAscii = Asc(UCase(Chr(intKeyAscii)))
                    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(intKeyAscii)) = 0 Then
                        intKeyAscii = 0
                    ElseIf zlCommFun.IsCharChinese(cboTmp.Text) Then
                        cboTmp.Text = "": cboTmp.Tag = ""
                    End If
                    If intKeyAscii <> 0 Then
                        cboTmp.Tag = cboTmp.Text & Chr(intKeyAscii)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboSpecificInfoChange(ByRef intIndex As Integer)
'cboSpecificInfo_Change事件
    Dim cboTmp As ComboBox
    Dim lngPos As Long, lngLen As Long

    If mblnReturn Then Exit Sub
    Select Case intIndex
        Case I身份证号
            Set cboTmp = cboE(intIndex)
            mblnReturn = True
            If Cbo.FindIndex(cboTmp, cboTmp.Text, True) = -1 Then
                '不规则的输入
                If Not zlStr.CheckCharScope(cboTmp.Text, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ*") Then
                    cboTmp.Text = ""
                Else
                    If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                        If zlCommFun.ActualLen(cboTmp.Text) > 18 Then
                            cboTmp.Text = Mid(cboTmp.Text, 1, 18)
                        End If
                    End If
                End If
            End If
            If Trim(zlCommFun.GetNeedName(cboE(I国籍).Text)) = "中国" Then
                lngPos = InStr(cboTmp.Text, "*")
                lngLen = Len(Mid(cboTmp.Text, 13, 2))
                Select Case lngPos
                    Case 0
                        cboTmp.Tag = cboTmp.Text
                    Case Else
                        cboTmp.Tag = Mid(cboTmp.Text, 1, lngPos - 1)
                        cboTmp.Text = cboTmp.Tag
                        cboTmp.SelStart = Len(cboTmp.Text)
                End Select
            Else
                cboTmp.Tag = cboTmp.Text
            End If
            mblnReturn = False
    End Select
End Sub

Private Function SetInputRoot(ByVal intType As Integer, ByVal intSysPara As Integer, ByRef intModPara As Integer, ParamArray arrControls() As Variant) As Boolean
'说明：该函数用于系统参数与模块参数共同控制一组单选按钮，系统参数值一般为A(0或1),A+1,A+2....,模块参数为B,B+1,....系统参数为A时，模块参数起作用,且有以下条件
'           模块参数=B(系统参数=A)产生的业务效果与系统参数=A+1相同
'           模块参数=B+1(系统参数=A)产生的业务效果与系统参数=A+2相同
'功能：设置来源，可以设置模块变量西医诊断来源，中医诊断来源，过敏输入来源
'参数：intType=0-西医诊断来源设置，1-中医诊断来源，2-过敏诊断来源
'      intSysPara=系统参数，参数值为A(0或1),A+1,A+2，..，值为A时模块参数起作用
'      intModPara=模块参数
'返回：是否成功
'      intModPara=实际参数值。如系统参数为，0，1，2，模块为0，1 ，系统为0时模块起作用，此时模块参数实际值=模块参数值，当系统参数<>0，如1，模块参数实际值=系统参数-1

    Dim blnVisual As Boolean, blnEnable As Boolean
    Dim i As Long
 
    On Error GoTo errH
    '过敏输入来源，当不启用太元通时控件不可见,其余情况可见
    blnVisual = intType = 2 And gbytPass = 3 Or intType <> 2
    blnEnable = intSysPara = IIf(intType <> 2, 1, 0)
    If Not blnVisual Then intModPara = 0
    If Not blnEnable Then intModPara = intSysPara - IIf(intType <> 2, 2, 1)
    '设置控件的值以及可用性
    For i = LBound(arrControls) To UBound(arrControls)
        arrControls(i).Visible = blnVisual
        If blnVisual Then
            arrControls(i).Enabled = blnEnable And arrControls(i).Enabled
            '实际模块参数值与控件数组下标起始值一样，顺序一样
            If i = intModPara Then
                arrControls(i).Value = 1
            Else
                arrControls(i).Value = 0
            End If
        End If
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateLastItem() As String
'功能：保存上次编辑的控件信息
    Call SavePreItem
End Function

Public Function IsDataSaved() As Boolean
    IsDataSaved = mblnOK
End Function


Private Sub MsgDis()
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim str疾病ID As String
    Dim str诊断ID As String
    On Error GoTo ErrHand
    '判断当前病人是否填写传染病报告卡
    strSQL = "Select 文件ID From 电子病历记录 Where 病人ID=[1] And 主页ID=[2] And 病历种类=5  and 创建人=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MsgDis", mlng病人ID, mlng挂号ID, UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        '判断用户是否修改或删除诊断
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then
                    If InStr("," & str诊断ID & ",", "," & Val(.TextMatrix(i, DI_诊断ID)) & ",") = 0 Then
                        str诊断ID = str诊断ID & "," & Val(.TextMatrix(i, DI_诊断ID))
                    End If
                End If
                If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                    If InStr("," & str疾病ID & ",", "," & Val(.TextMatrix(i, DI_疾病ID)) & ",") = 0 Then
                        str疾病ID = str疾病ID & "," & Val(.TextMatrix(i, DI_疾病ID))
                    End If
                End If
            Next
        End With
        
        If mbln中医 Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, DI_诊断ID)) <> 0 Then
                        If InStr("," & str诊断ID & ",", "," & Val(.TextMatrix(i, DI_诊断ID)) & ",") = 0 Then
                            str诊断ID = str诊断ID & "," & Val(.TextMatrix(i, DI_诊断ID))
                        End If
                    End If
                    If Val(.TextMatrix(i, DI_疾病ID)) <> 0 Then
                        If InStr("," & str疾病ID & ",", "," & Val(.TextMatrix(i, DI_疾病ID)) & ",") = 0 Then
                            str疾病ID = str疾病ID & "," & Val(.TextMatrix(i, DI_疾病ID))
                        End If
                    End If
                Next
            End With
        End If
        str疾病ID = Mid(str疾病ID, 2): str诊断ID = Mid(str诊断ID, 2)
        strSQL = ""
        If str疾病ID <> "" Then
            strSQL = " Union Select 疾病id,诊断id From 疾病报告前提 Where 疾病ID IN (Select Column_Value From Table(f_Num2list([3])))"
        End If
        If str诊断ID <> "" Then
            strSQL = strSQL & " Union Select 疾病id,诊断id From 疾病报告前提 Where 诊断ID IN (Select Column_Value From Table(f_Num2list([4])))"
        End If
        strSQL = "Select a.疾病id, a.诊断id From 病人诊断记录 A, 疾病报告前提 B Where a.病人id = [1] And a.主页id = [2] And a.编码序号 = 1 And (a.疾病id = b.疾病id Or a.诊断id = b.诊断id) " & IIf(strSQL = "", "", "Minus (" & Mid(strSQL, 8) & ") ")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "MsgDis", mlng病人ID, mlng挂号ID, str疾病ID, str诊断ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "当前病人传染病诊断数据发生了改变,请修改传染病报告卡！", vbInformation, gstrSysName
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

