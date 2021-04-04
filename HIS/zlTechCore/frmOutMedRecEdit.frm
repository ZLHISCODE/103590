VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmOutMedRecEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊首页"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmOutMedRecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   0
      Left            =   195
      TabIndex        =   34
      Top             =   465
      Width           =   6480
      Begin VB.CommandButton cmdEdit 
         Caption         =   "…"
         Height          =   255
         Index           =   5
         Left            =   6060
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   2265
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   900
         MaxLength       =   30
         TabIndex        =   14
         Top             =   2235
         Width           =   5475
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   2
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   495
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "…"
         Height          =   255
         Index           =   6
         Left            =   6060
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   2625
         Width           =   285
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "…"
         Height          =   255
         Index           =   9
         Left            =   6060
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "选择(*)"
         Top             =   3345
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   4905
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3675
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   900
         MaxLength       =   20
         TabIndex        =   22
         Top             =   3675
         Width           =   3090
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   900
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3315
         Width           =   5475
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4905
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2955
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   900
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2955
         Width           =   3090
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   900
         MaxLength       =   100
         TabIndex        =   16
         Top             =   2595
         Width           =   5475
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   900
         MaxLength       =   18
         TabIndex        =   13
         Top             =   1875
         Width           =   5475
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   6
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1365
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   5
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1365
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   4
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1005
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   3
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1005
         Width           =   2355
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   495
         Width           =   615
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   135
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3030
         MaxLength       =   5
         TabIndex        =   6
         Top             =   495
         Width           =   675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   20
         TabIndex        =   1
         Top             =   135
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         BackColor       =   &H8000000F&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txt出生时间 
         Height          =   300
         Left            =   1950
         TabIndex        =   5
         Top             =   495
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt出生日期 
         Height          =   300
         Left            =   900
         TabIndex        =   4
         Top             =   495
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   51
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   180
         Index           =   5
         Left            =   4380
         TabIndex        =   50
         Top             =   555
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -105
         X2              =   7245
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   -150
         X2              =   7200
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   -15
         X2              =   7335
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7290
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭邮编"
         Height          =   180
         Index           =   17
         Left            =   4095
         TabIndex        =   49
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   47
         Top             =   3375
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Index           =   14
         Left            =   4095
         TabIndex        =   46
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   45
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   44
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   43
         Top             =   1935
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   42
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Index           =   7
         Left            =   3555
         TabIndex        =   41
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Index           =   9
         Left            =   3555
         TabIndex        =   40
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Index           =   2
         Left            =   2610
         TabIndex        =   38
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   37
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   35
         Top             =   195
         Width           =   540
      End
   End
   Begin VB.Frame fraInfo 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   1
      Left            =   195
      TabIndex        =   52
      Top             =   465
      Width           =   6480
      Begin VB.CheckBox chkEdit 
         Caption         =   "复诊"
         Height          =   195
         Index           =   0
         Left            =   5505
         TabIndex        =   30
         Top             =   3195
         Width           =   750
      End
      Begin VB.CommandButton cmdMakeLog 
         Height          =   255
         Left            =   1260
         Picture         =   "frmOutMedRecEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "根据诊断生成就诊摘要(F12)"
         Top             =   3135
         Width           =   345
      End
      Begin VB.TextBox txtEdit 
         Height          =   660
         Index           =   12
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   3405
         Width           =   6405
      End
      Begin VB.OptionButton optInput 
         Caption         =   "根据诊断标准输入(&1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   0
         Left            =   2295
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1230
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optInput 
         Caption         =   "根据疾病编码输入(&2)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   4365
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   960
         Left            =   30
         TabIndex        =   27
         Top             =   1440
         Width           =   6405
         _cx             =   11298
         _cy             =   1693
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
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":0102
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   115
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
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   960
         Left            =   30
         TabIndex        =   24
         Top             =   225
         Width           =   6405
         _cx             =   11298
         _cy             =   1693
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":019E
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
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   735
         Left            =   30
         TabIndex        =   28
         Top             =   2400
         Width           =   6405
         _cx             =   11298
         _cy             =   1296
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
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmOutMedRecEdit.frx":01EF
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   115
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " 就诊摘要 "
         Height          =   180
         Index           =   20
         Left            =   285
         TabIndex        =   55
         Top             =   3195
         Width           =   900
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   5300
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   5300
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " 诊断记录 "
         Height          =   180
         Index           =   19
         Left            =   285
         TabIndex        =   54
         Top             =   1230
         Width           =   900
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   6400
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   6400
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   " 过敏记录 "
         Height          =   180
         Index           =   18
         Left            =   285
         TabIndex        =   53
         Top             =   15
         Width           =   900
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   75
         X2              =   6400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   75
         X2              =   6400
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   32
      ToolTipText     =   "热键：F2"
      Top             =   4740
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5280
      TabIndex        =   33
      Top             =   4740
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tbsInfo 
      Height          =   4515
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   7964
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "就诊信息"
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
End
Attribute VB_Name = "frmOutMedRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnDiagnose As Boolean
Private mlng病人ID As Long
Private mstr挂号单 As String
Private mlng挂号ID As Long
Private mint险类 As Integer
Private mbln中医 As Boolean
Private mstrLike As String
Private mint简码 As Integer
Private mblnChange As Boolean
Private mblnOK As Boolean

Private Enum TXT_ENUM
    txt姓名 = 0
    txt门诊号 = 1
    'txt出生日期 = 2
    txt年龄 = 3
    txt身份证号 = 4
    txt出生地点 = 5
    txt工作单位 = 6
    txt单位电话 = 7
    txt单位邮编 = 8
    txt家庭地址 = 9
    txt家庭电话 = 10
    txt家庭邮编 = 11
    txt就诊摘要 = 12
End Enum
Private Enum CBO_ENUM
    cbo性别 = 0
    cbo年龄 = 1
    cbo付款 = 2
    cbo国籍 = 3
    cbo民族 = 4
    cbo婚姻 = 5
    cbo职业 = 6
End Enum
Private Enum CHK_ENUM
    chk复诊 = 0
End Enum
Private Enum COL_ENUM
    col类型 = 0
    col诊断 = 1
    col疑诊 = 2
    col诊断ID = 3
    col疾病ID = 4
    col证候ID = 5
End Enum

Public Function ShowMe(frmParent As Object, ByVal str挂号单 As String, Optional blnDiagnose As Boolean) As Boolean
'参数：blnDiagnose=是否调用用于填写诊断
'返回：blnDiagnose=是否填写了病人的诊断
    mblnDiagnose = blnDiagnose
    mstr挂号单 = str挂号单
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    blnDiagnose = mblnDiagnose
    ShowMe = mblnOK
End Function

Private Function InitMedData() As Boolean
'功能：初始化编辑环境和必要的数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    Call zlControl.CboSetHeight(cboEdit(cbo民族), cboEdit(cbo民族).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo国籍), cboEdit(cbo国籍).Height * 16)
    Call zlControl.CboSetHeight(cboEdit(cbo职业), cboEdit(cbo职业).Height * 16)
    vsDiagXY.MergeCol(0) = True
    vsDiagZY.MergeCol(0) = True
    
    Call SetCboFromList(Array("岁", "月", "天"), Array(cbo年龄), 0)
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 性别 Order by 编码", Array(cbo性别))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 医疗付款方式 Order by 编码", Array(cbo付款))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 民族 Order by 编码", Array(cbo民族))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 国籍 Order by 编码", Array(cbo国籍))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 婚姻状况 Order by 编码", Array(cbo婚姻))
    Call SetCboFromSQL("Select 0 as ID,编码 as 简码,名称,缺省标志 From 职业 Order by 编码", Array(cbo职业))
    
    optInput(0).TabStop = False: optInput(1).TabStop = False '要强行代码执行一次
    
    InitMedData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadMedRec() As Boolean
'功能：读取门诊首页的各种信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long
    
    On Error GoTo errH
    
    '基本信息
    strSQL = "Select A.病人ID,B.ID as 挂号ID,B.执行部门ID as 科室ID,B.摘要,B.复诊," & _
        " A.险类,A.门诊号,A.姓名,A.性别,A.年龄,A.出生日期,A.医疗付款方式," & _
        " A.国籍,A.民族,A.婚姻状况,A.职业,A.身份证号,A.出生地点," & _
        " A.家庭地址,A.家庭电话,A.户口邮编,A.工作单位,A.单位电话,A.单位邮编" & _
        " From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And B.NO=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    If rsTmp.EOF Then Exit Function
    
    mlng病人ID = rsTmp!病人ID
    mlng挂号ID = rsTmp!挂号ID
    mint险类 = Nvl(rsTmp!险类, 0)
    mbln中医 = Have部门性质(rsTmp!科室ID, "中医科")
        
    txtEdit(txt姓名).Text = rsTmp!姓名
    Call GetCboIndex(cboEdit(cbo性别), Nvl(rsTmp!性别))
    txtEdit(txt门诊号).Text = Nvl(rsTmp!门诊号)
    
    If Not IsNull(rsTmp!出生日期) Then
        txt出生日期.Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
        If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
            txt出生时间.Text = Format(rsTmp!出生日期, "HH:mm")
        End If
    End If
    
    Call LoadOldData(Nvl(rsTmp!年龄))
    
    Call txt出生日期_Validate(False)
    
    If IsNumeric(txtEdit(txt年龄).Text) Then
         If Val(txtEdit(txt年龄).Text) <> CLng(txtEdit(txt年龄).Text) Then
            cboEdit(cbo年龄).ListIndex = 2                    '以天为单位
            txtEdit(txt年龄).Text = CLng(Val(txtEdit(txt年龄).Text) * 365)
        End If
    End If
    
    Call GetCboIndex(cboEdit(cbo付款), Nvl(rsTmp!医疗付款方式))
    Call GetCboIndex(cboEdit(cbo国籍), Nvl(rsTmp!国籍))
    Call GetCboIndex(cboEdit(cbo民族), Nvl(rsTmp!民族))
    Call GetCboIndex(cboEdit(cbo婚姻), Nvl(rsTmp!婚姻状况))
    Call GetCboIndex(cboEdit(cbo职业), Nvl(rsTmp!职业))
    txtEdit(txt身份证号).Text = Nvl(rsTmp!身份证号)
    txtEdit(txt出生地点).Text = Nvl(rsTmp!出生地点)
    txtEdit(txt工作单位).Text = Nvl(rsTmp!工作单位)
    txtEdit(txt单位电话).Text = Nvl(rsTmp!单位电话)
    txtEdit(txt单位邮编).Text = Nvl(rsTmp!单位邮编)
    txtEdit(txt家庭地址).Text = Nvl(rsTmp!家庭地址)
    txtEdit(txt家庭电话).Text = Nvl(rsTmp!家庭电话)
    txtEdit(txt家庭邮编).Text = Nvl(rsTmp!户口邮编)
    txtEdit(txt就诊摘要).Text = Nvl(rsTmp!摘要)
    chkEdit(chk复诊).Value = Nvl(rsTmp!复诊, 0)
    
    '过敏信息:本次挂号的,过敏的
    strSQL = "Select 记录来源,记录时间,药物ID,药物名 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>=A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by 记录时间,药物名"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '固定行+新行
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!药物ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!记录时间, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, 0) = Format(rsTmp!记录时间, "yyyy-MM-dd HH:mm:ss") '用于保存
                    .TextMatrix(i, 1) = Nvl(rsTmp!药物名)
                    .Cell(flexcpData, i, 1) = .TextMatrix(i, 1) '用于输入恢复
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = 1
    
    '诊断信息:本次挂号的
    strSQL = "Select 记录来源,诊断类型,疾病ID,诊断ID,证候ID,诊断描述,是否疑诊 From 病人诊断记录" & _
        " Where 记录来源 IN(1,3) And 诊断类型 IN(1,11)" & _
        " And 取消时间 is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by 诊断类型,诊断次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    If Not rsTmp.EOF Then
        '西医诊断
        rsTmp.Filter = "诊断类型=1 And 记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "诊断类型=1 And 记录来源<>3" '其它来源的作为缺省显示
        With vsDiagXY
            .Rows = rsTmp.RecordCount + 2
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col诊断) = Nvl(rsTmp!诊断描述)
                .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                .TextMatrix(i, col疑诊) = IIF(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断ID, 0)
                .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                rsTmp.MoveNext
            Next
            .Cell(flexcpText, .FixedRows, col类型, .Rows - 1, col类型) = "西医"
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        End With
        '中医诊断
        If mbln中医 Then
            rsTmp.Filter = "诊断类型=11 And 记录来源=3"
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=11 And 记录来源<>3"
            With vsDiagZY
                .Rows = rsTmp.RecordCount + 1
                For i = 0 To rsTmp.RecordCount - 1
                    .TextMatrix(i, col诊断) = Nvl(rsTmp!诊断描述)
                    .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                    .TextMatrix(i, col疑诊) = IIF(Nvl(rsTmp!是否疑诊, 0) = 1, "？", "")
                    .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断ID, 0)
                    .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                    .TextMatrix(i, col证候ID) = Nvl(rsTmp!证候ID, 0)
                    rsTmp.MoveNext
                Next
                .Cell(flexcpText, .FixedRows, col类型, .Rows - 1, col类型) = "中医"
                .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
            End With
        End If
    End If
    vsDiagXY.Row = vsDiagXY.FixedRows: vsDiagXY.Col = 0: vsDiagXY.Col = col诊断
    vsDiagZY.Row = vsDiagZY.FixedRows: vsDiagZY.Col = 0: vsDiagZY.Col = col诊断
        
    If Not mbln中医 Then
        vsDiagZY.Visible = False
        vsDiagXY.Height = vsDiagZY.Top + vsDiagZY.Height - vsDiagXY.Top
        vsDiagXY.ColHidden(0) = True
        vsDiagXY.ColWidth(1) = vsDiagXY.ColWidth(1) + vsDiagXY.ColWidth(0)
    End If
    
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CalcHowOld() As Long
'功能：根据当前的出生日期和年龄单位，计算年龄值
'返回：-1表示未计算
    Dim datBase As Date, lngTmp As Long
        
    CalcHowOld = -1
    If Not (IsDate(txt出生日期.Text) And IsNumeric(txtEdit(txt年龄).Text)) Then Exit Function
    
    datBase = zlDatabase.Currentdate
    lngTmp = DateDiff("yyyy", CDate(txt出生日期.Text), datBase)
    
    '只对岁这种情况进行检查
    If cboEdit(cbo年龄).ListIndex = 0 Then
        If Format(datBase, "MMdd") < Format(txt出生日期.Text, "MMdd") Then
            lngTmp = lngTmp - 1
        End If
        CalcHowOld = lngTmp
    End If
End Function

Private Function CheckMedRec(Optional blnDiagnose As Boolean) As Boolean
'功能：检查首页输入数据合法性
'返回：blnDiagnose=是否填写了诊断
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str身份证 As String, str出生日期 As String, lng性别 As Long
    Dim i As Long, j As Long
    
    blnDiagnose = False
    curDate = zlDatabase.Currentdate
    
    '必须要输入的内容检查
    '-----------------------------------------------------------------------------------------
    arrInfo = Array(txt姓名, txt年龄)
    arrName = Array("姓名", "年龄")
    For i = 0 To UBound(arrInfo)
        If txtEdit(arrInfo(i)).Enabled And Not txtEdit(arrInfo(i)).Locked And txtEdit(arrInfo(i)).Text = "" Then
            Call ShowMessage(txtEdit(arrInfo(i)), "必须输入病人的" & arrName(i) & "。")
            Exit Function
        End If
    Next
    
    Select Case cboEdit(cbo年龄).Text
        Case "岁"
            If Val(txtEdit(txt年龄).Text) > 200 Then
                MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                txtEdit(txt年龄).SetFocus: Exit Function
            End If
        Case "月"
            If Val(txtEdit(txt年龄).Text) > 2400 Then
                MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                txtEdit(txt年龄).SetFocus: Exit Function
            End If
        Case "天"
            If Val(txtEdit(txt年龄).Text) > 73000 Then
                MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                txtEdit(txt年龄).SetFocus: Exit Function
            End If
        Case Else
            Exit Function
    End Select
    If Not IsDate(txt出生日期.Text) Then
        Call ShowMessage(txt出生日期, "必须输入病人的出生日期。")
        Exit Function
    ElseIf txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        Call ShowMessage(txt出生时间, "请输入正确的病人出生时间。")
        Exit Function
    End If
    
    i = CalcHowOld
    If i <> -1 And i <> Val(txtEdit(txt年龄).Text) Then
        If ShowMessage(txt出生日期, "年龄和出生日期不一致，" & txt出生日期.Text & "出生现在应该是" & i & cboEdit(cbo年龄).Text & "。" & _
            vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", True) = vbNo Then
            Exit Function
        End If
    End If
    
    arrInfo = Array(cbo付款, cbo性别)
    arrName = Array("付款方式", "性别")
    For i = 0 To UBound(arrInfo)
        If cboEdit(arrInfo(i)).Enabled And Not cboEdit(arrInfo(i)).Locked And cboEdit(arrInfo(i)).ListIndex = -1 Then
            Call ShowMessage(cboEdit(arrInfo(i)), "必须输入病人的" & arrName(i) & "。")
            Exit Function
        End If
    Next
    
    '项目输入的长度检查
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtEdit
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "输入内容过长，请检查。(该项目最多允许 " & objTmp.MaxLength & " 个字符或 " & objTmp.MaxLength \ 2 & " 个汉字)")
                Exit Function
            End If
        End If
    Next
    
    '输入内容的有效性检查
    '-----------------------------------------------------------------------------------------
    '出生日期必须早于当前时间
    If Format(txt出生日期.Text, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
        Call ShowMessage(txt出生日期, "出生日期不应该比当前日期还晚。")
        Exit Function
    End If

    '15岁以下应为未婚
    If Not (cboEdit(cbo婚姻).Text = "" Or cboEdit(cbo婚姻).ListIndex = -1) Then
        If DateDiff("yyyy", CDate(txt出生日期.Text), curDate) < 15 Then
            If InStr(cboEdit(cbo婚姻).Text, "已婚") > 0 _
                Or InStr(cboEdit(cbo婚姻).Text, "丧偶") > 0 Or InStr(cboEdit(cbo婚姻).Text, "离婚") > 0 Then
                Call ShowMessage(cboEdit(cbo婚姻), "婚姻状况信息填写不对。")
                Exit Function
            End If
        End If
    End If
            
    '身份证号码检查
    '对身份证号进行验证
    str身份证 = txtEdit(txt身份证号).Text
    If str身份证 <> "" Then
        If Len(str身份证) <> 15 And Len(str身份证) <> 18 Then
            Call ShowMessage(txtEdit(txt身份证号), "身份证号码的长度不正确，应为15位或18位。")
            Exit Function
        End If

        If Len(str身份证) = 15 Then
            str出生日期 = Mid(str身份证, 7, 6)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Right(str身份证, 1))
        Else
            str出生日期 = Mid(str身份证, 7, 8)
            str出生日期 = Format(GetFullDate(str出生日期), "yyyy-MM-dd")
            lng性别 = Val(Mid(str身份证, 17, 1))
        End If
        If Not IsDate(str出生日期) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息不正确，是否继续？", True) = vbNo Then Exit Function
        Else
            If Format(str出生日期, "yyyy-MM-dd") <> Format(txt出生日期.Text, "yyyy-MM-dd") Then
                If ShowMessage(txtEdit(txt身份证号), "身份证号码中的出生日期信息与病人的出生日期不符，是否继续？", True) = vbNo Then Exit Function
            End If
        End If
        If (lng性别 Mod 2 = 1 And InStr(cboEdit(cbo性别).Text, "女") > 0) Or (lng性别 Mod 2 = 0 And InStr(cboEdit(cbo性别).Text, "男") > 0) Then
            If ShowMessage(txtEdit(txt身份证号), "身份证号码中的性别信息与病人的性别不符，是否继续？", True) = vbNo Then Exit Function
        End If
    End If
    
    '诊断表格的检查
    '-----------------------------------------------------------------------------------------
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 150 Then
                    .Row = i: .Col = col诊断
                    Call ShowMessage(vsDiagXY, "诊断内容太长，只允许150个字符或75个汉字。")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col诊断)) <> "" Then
                        If .TextMatrix(j, col诊断) = .TextMatrix(i, col诊断) Then
                            .Row = i: .Col = col诊断
                            Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                            Exit Function
                        ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                            If Val(.TextMatrix(j, col诊断ID)) = Val(.TextMatrix(i, col诊断ID)) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                            If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagXY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            End If
                        End If
                    End If
                Next
                blnDiagnose = True
            End If
        Next
    End With
        
    If mbln中医 Then
        With vsDiagZY
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 150 Then
                        .Row = i: .Col = col诊断
                        Call ShowMessage(vsDiagZY, "诊断内容太长，只允许150个字符或75个汉字。")
                        Exit Function
                    End If
                    For j = i + 1 To .Rows - 1
                        If Trim(.TextMatrix(j, col诊断)) <> "" Then
                            If .TextMatrix(j, col诊断) = .TextMatrix(i, col诊断) Then
                                .Row = i: .Col = col诊断
                                Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
                                Exit Function
                            ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 Then
                                '因中医诊断带证候,可能无对应证候ID,诊断ID又相同
'                                If Val(.TextMatrix(j, col诊断ID)) & "," & Val(.TextMatrix(j, col证候ID)) _
'                                    = Val(.TextMatrix(i, col诊断ID)) & "," & Val(.TextMatrix(i, col证候ID)) Then
'                                    .Row = i: .Col = col诊断
'                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
'                                    Exit Function
'                                End If
                            ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                                If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                    .Row = i: .Col = col诊断
                                    Call ShowMessage(vsDiagZY, "发现存在两行相同的诊断信息。")
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                    blnDiagnose = True
                End If
            Next
        End With
    End If
    
    '过敏药物表格检查
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                If zlCommFun.ActualLen(.TextMatrix(i, 1)) > 60 Then
                    .Row = i: .Col = 1
                    Call ShowMessage(vsAller, "过敏药物名太长，只允许60个字符或30个汉字。")
                    Exit Function
                End If
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, 1)) <> "" Then
                        If .TextMatrix(j, 1) = .TextMatrix(i, 1) Then
                            .Row = i: .Col = 1
                            Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                            Exit Function
                        ElseIf .RowData(i) <> 0 Then
                            If .RowData(j) = .RowData(i) Then
                                .Row = i: .Col = 1
                                Call ShowMessage(vsAller, "发现存在两行相同的过敏药物。")
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End With
    
    CheckMedRec = True
End Function

Private Function SaveMedRec() As Boolean
'功能：保存门诊首页的各种信息
    Dim arrSQL As Variant, i As Integer
    Dim curDate As Date, intIdx As Integer
    Dim str生日 As String
    
    arrSQL = Array()
    curDate = zlDatabase.Currentdate
    
    If IsDate(txt出生时间.Text) Then
        str生日 = "To_Date('" & Format(txt出生日期.Text & " " & txt出生时间.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
    Else
        str生日 = "To_Date('" & Format(txt出生日期.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    End If
    
    '病人信息
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人信息_首页整理(" & _
        mlng病人ID & "," & Val(txtEdit(txt门诊号).Text) & ",'" & txtEdit(txt姓名).Text & "'," & _
        "'" & NeedName(cboEdit(cbo性别).Text) & "','" & txtEdit(txt年龄).Text & cboEdit(cbo年龄).Text & "'," & _
        str生日 & ",'" & txtEdit(txt出生地点).Text & "','" & txtEdit(txt身份证号).Text & "'," & _
        "'" & NeedName(cboEdit(cbo民族).Text) & "','" & NeedName(cboEdit(cbo国籍).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo婚姻).Text) & "','" & NeedName(cboEdit(cbo职业).Text) & "'," & _
        "'" & NeedName(cboEdit(cbo付款).Text) & "','" & txtEdit(txt家庭地址).Text & "'," & _
        "'" & txtEdit(txt家庭电话).Text & "','" & txtEdit(txt家庭邮编).Text & "'," & _
        "'" & txtEdit(txt工作单位).Text & "','" & txtEdit(txt单位电话).Text & "'," & _
        "'" & txtEdit(txt单位邮编).Text & "',Null,Null,Null,Null,'" & mstr挂号单 & "'," & _
        chkEdit(chk复诊).Value & ",'" & txtEdit(txt就诊摘要).Text & "')"
    
    '过敏药物
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_病人过敏记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3)"
    With vsAller
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "zl_病人过敏记录_Insert(" & mlng病人ID & "," & mlng挂号ID & "," & _
                    "3,Null," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, 1) & "',1," & _
                    "To_Date('" & .Cell(flexcpData, i, 0) & "','YYYY-MM-DD HH24:MI:SS'))"
            End If
        Next
    End With
    
    '诊断记录
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,Null,'1')"
    With vsDiagXY
        intIdx = 0
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                    " Null,1," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & ",Null," & _
                    "'" & .TextMatrix(i, col诊断) & "',Null,Null," & IIF(.TextMatrix(i, col疑诊) = "", 0, 1) & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & intIdx & ")"
            End If
        Next
    End With
    
    If mbln中医 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,Null,'11')"
        With vsDiagZY
            intIdx = 0
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                    arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                        "Null,11," & ZVal(.TextMatrix(i, col疾病ID)) & "," & ZVal(.TextMatrix(i, col诊断ID)) & "," & _
                        ZVal(.TextMatrix(i, col证候ID)) & ",'" & .TextMatrix(i, col诊断) & "',Null,Null," & _
                        IIF(.TextMatrix(i, col疑诊) = "", 0, 1) & ",To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null," & intIdx & ")"
                End If
            Next
        End With
    End If
    
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    mblnChange = False
    SaveMedRec = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadOldData(strOld As String)
'功能:将数据库中保存的年龄按估计的格式加载到界面
    Dim strTmp As Long
    
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            txtEdit(txt年龄).Text = strTmp
            If cboEdit(cbo年龄).ListCount > 0 Then cboEdit(cbo年龄).ListIndex = 0
        Else
            txtEdit(txt年龄).Text = strOld
            cboEdit(cbo年龄).ListIndex = -1
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            txtEdit(txt年龄).Text = strTmp
            If cboEdit(cbo年龄).ListCount > 1 Then cboEdit(cbo年龄).ListIndex = 1
        Else
            txtEdit(txt年龄).Text = strOld
            cboEdit(cbo年龄).ListIndex = -1
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            txtEdit(txt年龄).Text = strTmp
            If cboEdit(cbo年龄).ListCount > 2 Then cboEdit(cbo年龄).ListIndex = 2
        Else
            txtEdit(txt年龄).Text = strOld
            cboEdit(cbo年龄).ListIndex = -1
        End If
    ElseIf IsNumeric(strOld) Then
        txtEdit(txt年龄).Text = strOld
        If cboEdit(cbo年龄).ListCount > 0 Then cboEdit(cbo年龄).ListIndex = 0
    Else
        txtEdit(txt年龄).Text = strOld
        cboEdit(cbo年龄).ListIndex = -1
    End If
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    tbsInfo.Tabs(objTmp.Container.Index + 1).Selected = True
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
    objTmp.SetFocus
    Me.Refresh
End Function

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboEdit(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboEdit(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub

Private Sub SetCboFromSQL(ByVal strSQL As String, ByVal arrCboIdx As Variant)
'功能：将指定数据源中的数据装入指定索引的一个或多个ComboBox
'参数：strSQL=包含"ID,简码,名称,缺省标志"字段
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, j As Long
    
    '清除原有数据
    For i = 0 To UBound(arrCboIdx)
        cboEdit(arrCboIdx(i)).Clear
    Next
    
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '装入数据
    For i = 1 To rsTmp.RecordCount
        For j = 0 To UBound(arrCboIdx)
            If IsNull(rsTmp!简码) Then
                cboEdit(arrCboIdx(j)).AddItem rsTmp!名称
            Else
                cboEdit(arrCboIdx(j)).AddItem rsTmp!简码 & "-" & Chr(13) & rsTmp!名称
            End If
            cboEdit(arrCboIdx(j)).ItemData(cboEdit(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
            If Nvl(rsTmp!缺省标志, 0) = 1 Then
                Call zlControl.CboSetIndex(cboEdit(arrCboIdx(j)).Hwnd, cboEdit(arrCboIdx(j)).NewIndex)
            End If
        Next
        rsTmp.MoveNext
    Next
    '无缺省时,为未选中
End Sub

Private Sub cboEdit_Click(Index As Integer)
    Dim strTmp As String
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If Index = cbo年龄 Then
        '根据出生日期重算年龄
'        If Not mblnChange Then Exit Sub
'        If IsDate(txt出生日期.Text) And cboEdit(cbo年龄).ListIndex <> -1 Then
'            strTmp = cboEdit(cbo年龄).Text
'            strTmp = Switch(strTmp = "岁", "yyyy", strTmp = "月", "m", strTmp = "天", "d")
'
'            txtEdit(txt年龄).Text = DateDiff(strTmp, txt出生日期.Text, zlDatabase.Currentdate)
'            If strTmp = "d" And txtEdit(txt年龄).Text = "0" Then txtEdit(txt年龄).Text = "1"
'        End If
    End If
End Sub

Private Sub cboEdit_GotFocus(Index As Integer)
    If cboEdit(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboEdit(Index))
    End If
End Sub

Private Sub cboEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cboEdit(Index).Hwnd, KeyAscii)
        If lngIdx = -1 And cboEdit(Index).ListCount > 0 Then lngIdx = 0
        cboEdit(Index).ListIndex = lngIdx
    End If
End Sub

Private Sub cboEdit_LostFocus(Index As Integer)
    Dim strTmp As String, lngTmp As Long
    Dim datTemp As Date, datBase As Date
    
    On Local Error Resume Next
    
    If Index = cbo年龄 Then
        If IsNumeric(txtEdit(txt年龄).Text) Then
            'And Between(Val(txtEdit(txt年龄).Text), 0, 200) Then
            'txt出生日期.Text = Year(zlDatabase.Currentdate) - Int(txtEdit(txt年龄).Text) & "-01-01"
            
            If Len(txtEdit(txt年龄).Text) > txtEdit(txt年龄).MaxLength Then Exit Sub  '以前输入的超长的不管
    
            Select Case cboEdit(cbo年龄).Text
                Case "岁"
                    If Val(txtEdit(txt年龄).Text) > 200 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtEdit(txt年龄).SetFocus: Exit Sub
                    End If
                Case "月"
                    If Val(txtEdit(txt年龄).Text) > 2400 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtEdit(txt年龄).SetFocus: Exit Sub
                    End If
                Case "天"
                    If Val(txtEdit(txt年龄).Text) > 73000 Then
                        MsgBox "年龄值过大，请检查输入是否正确。", vbInformation, gstrSysName
                        txtEdit(txt年龄).SetFocus: Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select
            
            If txtEdit(txt年龄).Text = "0" And cboEdit(cbo年龄).Text = "天" Then  '不足一天按一天算
                txtEdit(txt年龄).Text = 1
            End If
            
            If Not IsDate(txt出生日期.Text) Then
                '如果出生日期是年,月,并且与根据年龄倒算的是相同的,则不改变出生日期(避免改变输入的出生日)
                datBase = zlDatabase.Currentdate
                
                If IsDate(txt出生日期.Text) Then
                    If strTmp = "岁" Then
                        datTemp = DateAdd("yyyy", txtEdit(txt年龄).Text * -1, datBase)
                        If Year(txt出生日期.Text) = Year(datTemp) Then Exit Sub
                    ElseIf strTmp = "月" Then
                        datTemp = DateAdd("m", txtEdit(txt年龄).Text * -1, datBase)
                        If Year(txt出生日期.Text) = Year(datTemp) And Month(txt出生日期.Text) = Month(datTemp) Then Exit Sub
                    End If
                End If
                
                If Val(txtEdit(txt年龄).Text) < 1 Then
                    strTmp = "d"
                    datTemp = DateAdd(strTmp, txtEdit(txt年龄).Text * 365 * -1, datBase)
                Else
                    strTmp = Switch(strTmp = "岁", "yyyy", strTmp = "月", "m", strTmp = "天", "d")
                    datTemp = DateAdd(strTmp, txtEdit(txt年龄).Text * -1, datBase)
                End If
                txt出生日期.Text = Format(datTemp, "yyyy-MM-dd")    '不保留以前输过的月和日
            Else
                lngTmp = CalcHowOld
                If lngTmp <> -1 And lngTmp <> Val(txtEdit(txt年龄).Text) Then
                    If MsgBox("年龄和出生日期不一致，" & txt出生日期.Text & "出生现在应该是" & lngTmp & cboEdit(cbo年龄).Text & "。" & _
                        vbCrLf & vbCrLf & "请检查年龄或出生日期的正确性，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        txtEdit(txt年龄).SetFocus: Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub chkEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click(Index As Integer)
'说明：注意界面上要求CMD和对应TXT的Index相同
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    '使用Lock的方式,不采用Enabled的方式
    If Not cmdEdit(Index).Enabled Or txtEdit(Index).Locked Then
        txtEdit(Index).SetFocus: Exit Sub
    End If
    
    Select Case Index
        Case txt出生地点, txt家庭地址
            '选择地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, , , , , , , True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""地区""数据，请先到字典管理工具中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!名称
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt工作单位
            '选择单位信息
            strSQL = "Select ID,上级ID,末级,编码,名称,简码,地址,电话,开户银行,帐号,联系人" & _
                " From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "合约单位", , , , , , True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "没有设置""合约单位""数据，请先到合约单位管理中设置。", vbInformation, gstrSysName
                End If
                txtEdit(Index).SetFocus
            Else
                txtEdit(Index).Text = rsTmp!名称 & IIF(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
                txtEdit(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
End Sub

Private Sub cmdMakeLog_Click()
    Dim strLog As String, i As Long
    
    With vsDiagXY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col诊断) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, col诊断) & IIF(.TextMatrix(i, col疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    With vsDiagZY
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col诊断) <> "" Then
                strLog = strLog & "　" & .TextMatrix(i, col诊断) & IIF(.TextMatrix(i, col疑诊) <> "", "(？)", "")
            End If
        Next
    End With
    If strLog <> "" Then
        txtEdit(txt就诊摘要).Text = Mid(strLog, 2)
    End If
    txtEdit(txt就诊摘要).SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckMedRec(blnDiagnose) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("病人的诊断信息还没有输入，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If Not SaveMedRec Then Exit Sub
        
    mblnDiagnose = blnDiagnose
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnDiagnose Then
        On Error Resume Next
        vsDiagXY.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call cmdMakeLog_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔
    optInput(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "门诊诊断输入", 0))).Value = True
    
    '诊断输入来源
    If gint诊断来源 > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        If gint诊断来源 = 2 Then
            optInput(0).Value = True
        ElseIf gint诊断来源 = 3 Then
            optInput(1).Value = True
        End If
    End If
    
    If Not InitMedData Then Unload Me: Exit Sub
    If Not LoadMedRec Then Unload Me: Exit Sub
    
    tbsInfo.Tabs(2).Selected = True
    Call tbsInfo_Click
    
    mblnChange = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("如果关闭窗体，你所作的更改将不会保存。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "门诊诊断输入", IIF(optInput(0).Value, 0, 1)
End Sub

Private Sub optInput_LostFocus(Index As Integer)
    optInput(0).TabStop = False: optInput(1).TabStop = False '要强行代码执行一次
End Sub

Private Sub tbsInfo_Click()
    Dim i As Integer
    
    For i = 0 To fraInfo.UBound
        If i = tbsInfo.SelectedItem.Index - 1 Then
            fraInfo(i).Visible = True
            fraInfo(i).ZOrder
        Else
            fraInfo(i).Visible = False
        End If
    Next
End Sub

Private Sub tbsInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Visible Then mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt出生地点 Or Index = txt家庭地址) And txtEdit(Index).Text <> "" Then
            '输入地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 " & _
                " Where (Upper(编码) Like [1] Or Upper(简码) Like [2] Or Upper(名称) Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "地区", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", mstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt工作单位 And txtEdit(Index).Text <> "" Then
            '输入工作单位
            strSQL = "Select ID,编码,名称,简码,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                " Where (撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL)" & _
                " And (Upper(编码) Like [1] Or Upper(简码) Like [2] Or Upper(名称) Like [2])" & _
                " Order by 编码"
            vPoint = GetCoordPos(txtEdit(Index).Container.Hwnd, txtEdit(Index).Left, txtEdit(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "工作单位", False, "", "", False, _
                False, True, vPoint.x, vPoint.y, txtEdit(Index).Height, blnCancel, False, False, _
                UCase(txtEdit(Index).Text) & "%", mstrLike & UCase(txtEdit(Index).Text) & "%")
            '可以任意输入,不一定要匹配
            If Not rsTmp Is Nothing Then
                txtEdit(Index).Text = rsTmp!名称 & IIF(Not IsNull(rsTmp!地址), "(" & rsTmp!地址 & ")", "")
                If txtEdit(txt单位电话).Text = "" Then
                    txtEdit(txt单位电话).Text = Nvl(rsTmp!电话)
                End If
            End If
            txtEdit(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("*") Then
        '注意界面上要求CMD和对应TXT的Index相同
        KeyAscii = 0
        Call cmdEdit_Click(Index)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '非控制按键
        
        '限制输入长度
        If txtEdit(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(Index).Text) > txtEdit(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '限制输入内容
        Select Case Index
            Case txt年龄
                strMask = "1234567890"
            'Case txt出生日期
                'strMask = "1234567890-"
            Case txt身份证号
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                strMask = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Case txt家庭电话, txt单位电话
                strMask = "1234567890-()"
            Case txt家庭邮编, txt单位邮编
                strMask = "1234567890"
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txt出生日期_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt出生日期_GotFocus()
    Call zlControl.TxtSelAll(txt出生日期)
End Sub

Private Sub txt出生日期_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt出生日期_Validate(Cancel As Boolean)
    Dim datBase As Date, lngTmp As Long
    
    On Local Error Resume Next
    
    If IsDate(txt出生日期.Text) Then
        'If txtEdit(txt年龄).Text = "" Then
            'strTmp = Get年龄值(CDate(txt出生日期.Text))
            'If strTmp <> "" Then txtEdit(txt年龄).Text = strTmp
            
            datBase = zlDatabase.Currentdate
            lngTmp = Val(Format(datBase, "yyyy")) - Val(Format(CDate(txt出生日期.Text), "yyyy"))
            
            If lngTmp > 1 Then '2岁以上
                '未过生日
                If Format(datBase, "MMdd") < Format(txt出生日期.Text, "MMdd") Then
                    lngTmp = lngTmp - 1
                End If
                txtEdit(txt年龄).Text = lngTmp
                cboEdit(cbo年龄).ListIndex = 0
            Else
                '2岁以下按月计
                lngTmp = Val(Format(datBase, "MM")) - Val(Format(CDate(txt出生日期.Text), "MM")) + IIF(lngTmp = 1, 12, 0)
                
                If lngTmp > 1 Then '月
                   txtEdit(txt年龄).Text = lngTmp
                   cboEdit(cbo年龄).ListIndex = 1
                Else
                    '2月以下按天计
                    lngTmp = Val(datBase - CDate(txt出生日期.Text))
                    txtEdit(txt年龄).Text = IIF(lngTmp = 0, 1, lngTmp)   '不足一天算一天
                    cboEdit(cbo年龄).ListIndex = 2
                End If
            End If
        'End If
    Else
        txt出生日期.Text = "____-__-__"
        txt出生时间.Text = "__:__"
        Cancel = True
    End If
End Sub

Private Sub txt出生时间_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt出生时间_GotFocus()
    Call zlControl.TxtSelAll(txt出生时间)
End Sub

Private Sub txt出生时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not IsDate(txt出生日期.Text) Then
        KeyAscii = 0: txt出生时间.Text = "__:__"
    End If
End Sub

Private Sub txt出生时间_Validate(Cancel As Boolean)
    If txt出生时间.Text <> "__:__" And Not IsDate(txt出生时间.Text) Then
        txt出生时间.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsAller
        If Col = 1 Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsAller_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = 1 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int性别 As Integer
    
    With vsAller
        If cboEdit(cbo性别).Text Like "*男*" Then
            int性别 = 1
        ElseIf cboEdit(cbo性别).Text Like "*女*" Then
            int性别 = 2
        End If
        
        strSQL = _
            " Select -1 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'西成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select -2 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中成药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select -3 as ID,-NULL as 上级ID,0 as 末级,NULL as 编码,'中草药' as 名称,NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试 From Dual Union ALL" & _
            " Select ID,Nvl(上级ID,-类型) as 上级ID,0 as 末级,NULL as 编码,名称," & _
            " NULL as 单位,NULL as 剂型,NULL as 毒理分类,NULL as 皮试" & _
            " From 诊疗分类目录 Where 类型 IN (1,2,3)" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
            " Union All" & _
            " Select Distinct A.ID,A.分类ID as 上级ID,1 as 末级,A.编码,A.名称," & _
            " A.计算单位 as 单位,B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
            " From 诊疗项目目录 A,药品特性 B" & _
            " Where A.类别 IN('5','6','7') And A.ID=B.药名ID" & _
            IIF(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[1])", "") & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "过敏药物", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有药品数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetAllerInput(Row, rsTmp)
            Call AllerEnterNextCell
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 1) <> "" Then
                If MsgBox("确实要清除该行过敏药物吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    With vsAller
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = 1 Then
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
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim StrInput As String, vPoint As POINTAPI
    Dim int性别  As Integer
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsAller
            If Col = 1 And .EditText <> "" Then
                StrInput = UCase(.EditText)
                If cboEdit(cbo性别).Text Like "*男*" Then
                    int性别 = 1
                ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                    int性别 = 2
                End If
                strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位," & _
                    " B.药品剂型 as 剂型,B.毒理分类,Decode(B.是否皮试,1,'√','') as 皮试" & _
                    " From 诊疗项目目录 A,药品特性 B,诊疗项目别名 C" & _
                    " Where A.类别 IN('5','6','7') And A.ID=B.药名ID And A.ID=C.诊疗项目ID" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or C.名称 Like [2] Or C.简码 Like [2])" & _
                    IIF(int性别 <> 0, " And Nvl(A.适用性别,0) IN(0,[3])", "") & _
                    Decode(mint简码, 0, " And C.码类=[4]", 1, " And C.码类=[4]", "") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " Order by A.编码"
                
                vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "过敏药物", False, "", "", False, _
                    False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                    StrInput & "%", mstrLike & StrInput & "%", int性别, mint简码 + 1)
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    Call vsAller_AfterRowColChange(Row, Col, Row, Col)
                    .SetFocus: Exit Sub
                Else
                    Call SetAllerInput(Row, rsTmp)
                    .EditText = .TextMatrix(Row, Col)
                End If
                Call AllerEnterNextCell
            End If
        End With
    End If
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAller.EditSelStart = 0
    vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col诊断 Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsDiagXY_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagXY
        If Not DiagCellEditable(vsDiagXY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col诊断 Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagZY.ColWidth(Col) = vsDiagXY.ColWidth(Col)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str性别 As String
    
    With vsDiagXY
        If optInput(0).Value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            strSQL = _
                " Select 0 As 末级,NULL||ID As ID,上级ID," & _
                " -NULL as 项目ID,编码,名称,Null As 说明,Null As 编者" & _
                " From 疾病诊断分类 Where 类别=1" & _
                " Start With 上级ID Is Null Connect By Prior ID=上级ID" & _
                " Union All" & _
                " Select 1 As 末级,A.ID||'0'||B.分类ID as ID,B.分类ID As 上级ID," & _
                " A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                " From 疾病诊断目录 A,疾病诊断属类 B" & _
                " Where A.ID=B.诊断ID And A.类别=1"
        Else
            If cboEdit(cbo性别).Text Like "*男*" Then
                str性别 = "男"
            ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                str性别 = "女"
            End If
            'D-ICD-10疾病编码
            strSQL = _
                " Select 0 as 末级,ID,上级ID,-NULL as 项目ID,类别||LPAD(序号,3,'0') as 编码," & _
                " NULL as 附码,名称,简码,NULL as 说明 From 疾病编码分类" & _
                " Where 类别='D' Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union ALL " & _
                " Select 1 as 末级,ID,分类ID as 上级ID,ID as 项目ID,编码,附码,名称,简码,说明" & _
                " From 疾病编码目录 Where 类别='D'" & _
                IIF(str性别 <> "", " And (性别限制=[1] Or 性别限制 is NULL)", "")
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, IIF(optInput(0).Value, "疾病诊断", "疾病编码"), _
            False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, str性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有" & IIF(optInput(0).Value, "疾病诊断", "疾病编码") & "数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call XYSetDiagInput(Row, rsTmp)
            Call DiagEnterNextCell(vsDiagXY)
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断) <> "" Then
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagXY)
        ElseIf KeyAscii = 32 And (.Col = col疑诊) Then
            If DiagCellEditable(vsDiagXY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col疑诊 Then
                    .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "", "？", "")
                End If
            End If
        Else
            If .Col = col诊断 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str性别 As String, StrInput As String
    Dim vPoint As POINTAPI
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsDiagXY
            If Col = col诊断 Then
                If .EditText <> "" Then
                    StrInput = UCase(.EditText)
                    If optInput(0).Value Then
                        '按诊断输入:西医部份，一个诊断可能属于多个分类
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "B.名称 Like [2]" '输入汉字时,只匹配名称
                        Else
                            strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                        End If
                        strSQL = _
                            " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                            " From 疾病诊断目录 A,疾病诊断别名 B" & _
                            " Where A.ID=B.诊断ID And A.类别=1" & _
                            Decode(mint简码, 0, " And B.码类=[4]", 1, " And B.码类=[4]", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by A.编码"
                    Else
                        If cboEdit(cbo性别).Text Like "*男*" Then
                            str性别 = "男"
                        ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                            str性别 = "女"
                        End If
                        'D-ICD-10疾病编码
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "名称 Like [2]" '输入汉字时,只匹配名称
                        Else
                            strSQL = "编码 Like [1] Or 名称 Like [2] Or 简码 Like [2]"
                        End If
                        strSQL = _
                            " Select ID,ID as 项目ID,编码,附码,名称,简码,说明" & _
                            " From 疾病编码目录 Where 类别='D'" & _
                            IIF(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by 编码"
                    End If
                    vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(optInput(0).Value, "疾病诊断", "疾病编码"), _
                        False, "", "", False, False, True, vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, _
                        StrInput & "%", mstrLike & StrInput & "%", str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        Call vsDiagXY_AfterRowColChange(Row, Col, Row, Col)
                        .SetFocus: Exit Sub
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (gint诊断输入 = 2 Or gint诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                            Call vsDiagXY_AfterRowColChange(Row, Col, Row, Col)
                            .SetFocus: Exit Sub
                        End If
                    
                        Call XYSetDiagInput(Row, rsTmp)
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    Call DiagEnterNextCell(vsDiagXY)
                End If
            End If
        End With
    End If
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagXY, Row, Col) Then
        Cancel = True
    ElseIf Col = col疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col诊断 Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            End If
            Call vsDiagZY_AfterRowColChange(-1, -1, Row, Col)
        End If
    End With
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDiagZY
        If Not DiagCellEditable(vsDiagZY, NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col诊断 Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsDiagXY.ColWidth(Col) = vsDiagZY.ColWidth(Col)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim str性别 As String
    
    With vsDiagZY
        If optInput(0).Value Then
            '按诊断输入:中医部份，一个诊断可能属于多个分类
            strSQL = _
                " Select 0 As 末级,NULL||ID As ID,上级ID," & _
                " -NULL as 项目ID,编码,名称,Null As 说明,Null As 编者" & _
                " From 疾病诊断分类 Where 类别=2" & _
                " Start With 上级ID Is Null Connect By Prior ID=上级ID" & _
                " Union All" & _
                " Select 1 As 末级,A.ID||'0'||B.分类ID as ID,B.分类ID As 上级ID," & _
                " A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                " From 疾病诊断目录 A,疾病诊断属类 B" & _
                " Where A.ID=B.诊断ID And A.类别=2"
        Else
            If cboEdit(cbo性别).Text Like "*男*" Then
                str性别 = "男"
            ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                str性别 = "女"
            End If
            'B-中医疾病编码
            strSQL = _
                " Select 0 as 末级,ID,上级ID,-NULL as 项目ID,类别||LPAD(序号,3,'0') as 编码," & _
                " NULL as 附码,名称,简码,NULL as 说明 From 疾病编码分类" & _
                " Where 类别='B'" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
                " Union ALL " & _
                " Select 1 as 末级,ID,分类ID as 上级ID,ID as 项目ID,编码,附码,名称,简码,说明" & _
                " From 疾病编码目录 Where 类别='B'" & _
                IIF(str性别 <> "", " And (性别限制=[1] Or 性别限制 is NULL)", "")
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, IIF(optInput(0).Value, "疾病诊断", "疾病编码"), _
            False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, str性别)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有" & IIF(optInput(0).Value, "疾病诊断", "疾病编码") & "数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call ZYSetDiagInput(Row, rsTmp)
            Call DiagEnterNextCell(vsDiagZY)
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col诊断) <> "" Then
                If MsgBox("确实要清除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell(vsDiagZY)
        ElseIf KeyAscii = 32 And (.Col = col疑诊) Then
            If DiagCellEditable(vsDiagZY, .Row, .Col) Then
                KeyAscii = 0
                If .Col = col疑诊 Then
                    .TextMatrix(.Row, .Col) = IIF(.TextMatrix(.Row, .Col) = "", "？", "")
                End If
            End If
        Else
            If .Col = col诊断 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim StrInput As String, vPoint As POINTAPI
    Dim str性别 As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsDiagZY
            If Col = col诊断 Then
                If .EditText <> "" Then
                    StrInput = UCase(.EditText)
                    If optInput(0).Value Then
                        '按诊断输入:中医部份，一个诊断可能属于多个分类
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
                        Else
                            strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
                        End If
                        strSQL = _
                            " Select Distinct A.ID,A.ID as 项目ID,A.编码,A.名称,A.说明,A.编者" & _
                            " From 疾病诊断目录 A,疾病诊断别名 B" & _
                            " Where A.ID=B.诊断ID And A.类别=2" & _
                            Decode(mint简码, 0, " And B.码类=[4]", 1, " And B.码类=[4]", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by A.编码"
                    Else
                        If cboEdit(cbo性别).Text Like "*男*" Then
                            str性别 = "男"
                        ElseIf cboEdit(cbo性别).Text Like "*女*" Then
                            str性别 = "女"
                        End If
                        'B-中医疾病编码
                        If zlCommFun.IsCharChinese(StrInput) Then
                            strSQL = "名称 Like [2]" '输入汉字时只匹配名称
                        Else
                            strSQL = "编码 Like [1] Or 名称 Like [2] Or 简码 Like [2]"
                        End If
                        strSQL = _
                            " Select ID,ID as 项目ID,编码,附码,名称,简码,说明" & _
                            " From 疾病编码目录" & _
                            " Where 类别='B'" & _
                            IIF(str性别 <> "", " And (性别限制=[3] Or 性别限制 is NULL)", "") & _
                            " And (" & strSQL & ")" & _
                            " Order by 编码"
                    End If
                    vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(optInput(0).Value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                        vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%", str性别, mint简码 + 1)
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        Call vsDiagZY_AfterRowColChange(Row, Col, Row, Col)
                        .SetFocus: Exit Sub
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (gint诊断输入 = 2 Or gint诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                            Call vsDiagZY_AfterRowColChange(Row, Col, Row, Col)
                            .SetFocus: Exit Sub
                        End If
                    
                        Call ZYSetDiagInput(Row, rsTmp)
                        .EditText = .TextMatrix(Row, Col)
                    End If
                    Call DiagEnterNextCell(vsDiagZY)
                End If
            End If
        End With
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(vsDiagZY, Row, Col) Then
        Cancel = True
    ElseIf Col = col疑诊 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Function DiagCellEditable(objGrid As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With objGrid
        '必须先输入诊断
        If .TextMatrix(lngRow, col诊断) = "" Then
            If lngCol = col疑诊 Then
                Exit Function
            End If
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub DiagEnterNextCell(objGrid As VSFlexGrid)
    Dim i As Long, j As Long
    
    With objGrid
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col诊断) To col疑诊
                If DiagCellEditable(objGrid, i, j) Then Exit For
            Next
            If j <= col疑诊 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理过敏药物的输入
    Dim strSQL As String, curDate As Date
    
    With vsAller
        If Not rsInput Is Nothing Then
            .RowData(lngRow) = CLng(rsInput!ID)
            .TextMatrix(lngRow, 1) = Nvl(rsInput!名称)
        Else
            .RowData(lngRow) = 0
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        
        curDate = zlDatabase.Currentdate
        .TextMatrix(lngRow, 0) = Format(curDate, "yyyy-MM-dd HH:mm")
        .Cell(flexcpData, lngRow, 0) = Format(curDate, "yyyy-MM-dd HH:mm:ss")
        
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        mblnChange = True
    End With
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理西医诊断项目的输入
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            .TextMatrix(lngRow, col诊断) = IIF(Not IsNull(rsInput!编码), "(" & rsInput!编码 & ")", "") & Nvl(rsInput!名称)
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            '根据诊断确定疾病,或根据疾病确定诊断
            If optInput(0).Value Then
                .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                .TextMatrix(lngRow, col疾病ID) = ""
                strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
            Else
                .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                .TextMatrix(lngRow, col诊断ID) = ""
                strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
            If Not rsTmp.EOF Then
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                Else
                    .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                End If
            End If
        Else
            .TextMatrix(lngRow, col诊断) = .EditText
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
        End If
        
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col类型) = "西医"
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        End If
        mblnChange = True
    End With
End Sub

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理中医诊断项目的输入
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, str编码 As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            If Not IsNull(rsInput!编码) Then
                str编码 = "(" & rsInput!编码 & ")"
            End If
            .TextMatrix(lngRow, col诊断) = Nvl(rsInput!名称)
            
            '根据诊断确定疾病,或根据疾病确定诊断
            If optInput(0).Value Then
                .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                .TextMatrix(lngRow, col疾病ID) = ""
                strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
            Else
                .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                .TextMatrix(lngRow, col诊断ID) = ""
                strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
            End If
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
            If Not rsTmp.EOF Then
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                Else
                    .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                End If
            End If
            
            '中医根据疾病诊断参考取证候
            If Val(.TextMatrix(lngRow, col诊断ID)) <> 0 Then
                strSQL = "Select Distinct 证候序号 as ID,证候ID,证候名称" & _
                    " From 疾病诊断参考" & _
                    " Where 诊断ID=[1] And 证候名称 is Not NULL" & _
                    " Order by 证候序号"
                vPoint = GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                    vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, Val(Val(.TextMatrix(lngRow, col诊断ID))))
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, col证候ID) = Nvl(rsTmp!证候ID)
                    .TextMatrix(lngRow, col诊断) = Nvl(rsTmp!证候名称) & .TextMatrix(lngRow, col诊断)
                End If
            End If
            .TextMatrix(lngRow, col诊断) = str编码 & .TextMatrix(lngRow, col诊断)
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
        Else
            .TextMatrix(lngRow, col诊断) = .EditText
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
            .TextMatrix(lngRow, col证候ID) = ""
        End If
        
        '如果是出院诊断,始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, col类型) = "中医"
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
        End If
        mblnChange = True
    End With
End Sub
