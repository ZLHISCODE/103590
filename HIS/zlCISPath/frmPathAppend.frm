VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmPathAppend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加路径外项目"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   Icon            =   "frmPathAppend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraERPType 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8040
      TabIndex        =   39
      Top             =   3930
      Width           =   2535
      Begin VB.OptionButton optEPRType 
         Caption         =   "老版"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   41
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optEPRType 
         Caption         =   "新版"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   40
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblEPR 
         Caption         =   "病历版本"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.ComboBox cboItems 
      Height          =   300
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   37
      ToolTipText     =   "默认插入选中位置之前"
      Top             =   960
      Width           =   2175
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11400
      MouseIcon       =   "frmPathAppend.frx":1708A
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3885
      Width           =   330
   End
   Begin VB.Frame fraVariation 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2760
      Left            =   6000
      TabIndex        =   32
      Top             =   960
      Width           =   5895
      Begin VB.TextBox txtVariation 
         Height          =   300
         Left            =   3885
         MaxLength       =   1000
         TabIndex        =   4
         Top             =   0
         Width           =   1890
      End
      Begin VSFlex8Ctl.VSFlexGrid vsVariation 
         Height          =   2380
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   5775
         _cx             =   10186
         _cy             =   4198
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   8000
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":171DC
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblSearch 
         Caption         =   "查找(&F)"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblVariation 
         Caption         =   "变异原因"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Frame fraItemKind 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   3930
      Width           =   3615
      Begin VB.OptionButton optType 
         Caption         =   "病历类"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   0
         Width           =   840
      End
      Begin VB.OptionButton optType 
         Caption         =   "医嘱类"
         Height          =   180
         Index           =   0
         Left            =   810
         TabIndex        =   6
         Top             =   0
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton optType 
         Caption         =   "其他类"
         Height          =   180
         Index           =   2
         Left            =   2550
         TabIndex        =   8
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目类型"
         Height          =   180
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame fraExecute 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1420
      Left            =   0
      TabIndex        =   26
      Top             =   6650
      Width           =   11895
      Begin VB.Frame fraExecutor 
         Appearance      =   0  'Flat
         Caption         =   "执行者"
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   3480
         TabIndex        =   27
         Top             =   90
         Width           =   1305
         Begin VB.OptionButton optExecutor 
            Caption         =   "护士"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   660
         End
         Begin VB.OptionButton optExecutor 
            Caption         =   "医生"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   660
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsResult 
         Height          =   1290
         Left            =   5880
         TabIndex        =   11
         Top             =   90
         Width           =   5895
         _cx             =   10398
         _cy             =   2275
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":17241
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   5
         X1              =   0
         X2              =   11880
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   4
         X1              =   0
         X2              =   11880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmPathAppend.frx":172BF
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblExePrompt 
         Caption         =   $"frmPathAppend.frx":17ABD
         Height          =   1335
         Left            =   960
         TabIndex        =   31
         Top             =   90
         Width           =   2295
      End
      Begin VB.Label lblResult 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行结果"
         Height          =   180
         Left            =   5040
         TabIndex        =   28
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.TextBox txtReason 
      Height          =   2415
      Left            =   1080
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.ComboBox cboItemType 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1050
      MaxLength       =   1000
      TabIndex        =   1
      Top             =   3900
      Width           =   3255
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11925
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8055
      Width           =   11925
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9435
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   10665
         TabIndex        =   15
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11880
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11925
      TabIndex        =   16
      Top             =   0
      Width           =   11925
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   11880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   11880
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathAppend.frx":17B5F
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "    标准路径中定义的项目不能满足病人实际需求，但又不至于由于这些因素的影响而退出路径，可以临时添加路径外项目。"
         Height          =   360
         Left            =   1095
         TabIndex        =   18
         Top             =   360
         Width           =   10605
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径外项目"
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
         Left            =   1095
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin MSComctlLib.ImageList imgNature 
      Left            =   0
      Top             =   2760
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
            Picture         =   "frmPathAppend.frx":196A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":19C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1A1D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1A76F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1AD09
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1B2A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1B83D
            Key             =   "Selected"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathAppend.frx":1BBD7
            Key             =   "UnSelected"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAdvice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2380
      Left            =   120
      TabIndex        =   21
      Top             =   4250
      Width           =   11655
      Begin zlCISPath.UCAdviceList UCAdvice 
         Height          =   1935
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   11655
         _extentx        =   20558
         _extenty        =   3413
      End
      Begin VB.CommandButton cmdAdvice 
         Caption         =   "医嘱编辑(&E)"
         Height          =   350
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.Frame fraEPR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2380
      Left            =   120
      TabIndex        =   20
      Top             =   4250
      Width           =   11775
      Begin VSFlex8Ctl.VSFlexGrid vsEPR 
         Height          =   2375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11640
         _cx             =   20532
         _cy             =   4189
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathAppend.frx":1BF71
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   10000
         Y1              =   1335
         Y2              =   1335
      End
   End
   Begin VB.Label lblItemPosition 
      AutoSize        =   -1  'True
      Caption         =   "插入位置"
      Height          =   180
      Left            =   2880
      TabIndex        =   38
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目图标"
      Height          =   180
      Left            =   10680
      TabIndex        =   36
      Top             =   3960
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   8
      X1              =   120
      X2              =   11780
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   6
      X1              =   120
      X2              =   11780
      Y1              =   3825
      Y2              =   3825
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "变异说明"
      Height          =   180
      Left            =   210
      TabIndex        =   25
      Top             =   1320
      Width           =   720
   End
   Begin XtremeCommandBars.CommandBars cbsIcon 
      Left            =   0
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl分类 
      AutoSize        =   -1  'True
      Caption         =   "项目分类"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称"
      Height          =   180
      Left            =   210
      TabIndex        =   22
      Top             =   3960
      Width           =   720
   End
End
Attribute VB_Name = "frmPathAppend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'入口参数
Private mlngFun As Long '0-直接添加,1-医嘱新开时添加,2-修改路径外项目(不提供医嘱和病历文件的内容修改)
Private mfrmParent As Object
Private mPP As TYPE_PATH_Pati
Private mPati As TYPE_Pati
Private mstrItemType As String
Private mint场合 As Integer
Private mstr医嘱IDs As String   'mlngFun=1时传入
Private mlng执行ID As Long       'mlngFun=1或2时传入
Private mclsMipModule As zl9ComLib.clsMipModule ' 消息平台对象

'程序变量
Private mrsResult As ADODB.Recordset '执行结果集
Private mrsNature As ADODB.Recordset '结果性质集

Private mblnUseExecute As Boolean

Private mvItem As TYPE_PATH_ITEM
Private mblnReturn As Boolean
Private mblnOK As Boolean
Private marrSQL As Variant          '医嘱校对的插入过程SQL
Private mdatPathOut As Date     'mlngFun=0 或 mlngFun=1 时,住院医嘱编辑界面返回
Private mlng阶段ID As Long      'mlngFun=0时,路径外项目生成的阶段ID
Private mlng天数 As Long        'mlngFun=0时,路径外项目生成的天数

Private Enum CONST_COL_执行结果
    col执行图标 = 0
    col执行结果 = 1
    col结果性质 = 2
    col缺省结果 = 3
End Enum

Private Enum CONST_COL_变异原因
    col变异分类 = 0
    col变异原因 = 1
    col变异选择 = 2
End Enum


Public Function ShowMe(frmParent As Object, int场合 As Integer, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal strItemType As String, ByVal bytUseType As Byte, ByVal str医嘱IDs As String, ByVal lng执行ID As Long, Optional ByRef objMip As Object, _
    Optional ByVal datDate As Date = CDate(0)) As Boolean
    
    Set mfrmParent = frmParent
    mint场合 = int场合
    mPati = t_pati
    mPP = t_pp
    mstrItemType = strItemType
    mlngFun = bytUseType
    mstr医嘱IDs = str医嘱IDs
    mlng执行ID = lng执行ID
    mblnOK = False
    mdatPathOut = datDate
    mlng阶段ID = 0
    mlng天数 = 0
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub LoadItemType()
'功能：加载路径分类
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long

    strSql = "Select 名称 From 临床路径分类 Where 路径ID = [1] And 版本号 = [2] And NVL(分支ID,0)=[3] Order by 序号"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.路径ID, mPP.版本号, mPP.当前阶段分支ID)
    cboItemType.Clear
    For i = 1 To rsTmp.RecordCount
        cboItemType.AddItem rsTmp!名称
        If rsTmp!名称 = mstrItemType Then cboItemType.ListIndex = cboItemType.NewIndex
        rsTmp.MoveNext
    Next
    If cboItemType.ListIndex = -1 And cboItemType.ListCount > 0 Then cboItemType.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboItemType_Click()
  Call LoadCurrentItems
End Sub

Private Sub cbsIcon_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = -1 Then
        picIcon.Cls
        mvItem.图标ID = 0
    Else
        mvItem.图标ID = Control.ID
        Call DrawPicture(GetPathIcon(mvItem.图标ID))
    End If
End Sub

Private Sub DrawPicture(objPic As StdPicture)
    Dim X As Long, Y As Long, W As Long, H As Long
    
    W = picIcon.ScaleX(objPic.Width, vbHimetric, vbTwips)
    H = picIcon.ScaleY(objPic.Height, vbHimetric, vbTwips)
    
    X = (picIcon.ScaleWidth - W) / 2
    Y = (picIcon.ScaleHeight - H) / 2
    
    picIcon.PaintPicture objPic, X, Y, W, H
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
End Sub

Private Sub Form_Load()
    Dim vItem As TYPE_PATH_ITEM

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsIcon.VisualTheme = xtpThemeOffice2003
    cbsIcon.ActiveMenuBar.Visible = False
    With Me.cbsIcon.Options
        .ToolBarAccelTips = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    fraVariation.BackColor = Me.BackColor
    fraAdvice.BackColor = Me.BackColor
    fraEPR.BackColor = Me.BackColor
    fraExecute.BackColor = Me.BackColor

    mblnUseExecute = Val(zlDatabase.GetPara("是否启用路径执行环节", glngSys, p临床路径应用, 1))
    If mblnUseExecute = False Then
        fraExecute.Visible = False
        fraAdvice.Height = fraAdvice.Height + fraExecute.Height + 60
        fraEPR.Height = fraAdvice.Height
        vsEPR.Height = fraEPR.Height
    Else
        Call Init基本执行结果
    End If

    vsEPR.Row = 0: vsEPR.Row = 1: vsEPR.Col = 0
    vsResult.Row = 0: vsResult.Row = 1: vsResult.Col = 0
    
    Call LoadItemType
    
    mvItem = vItem  '清空之前可能保留的信息
    Call InitVariation
    
    If mlngFun = 1 Or mlngFun = 2 Then
        fraItemKind.Enabled = False
        optType(0).Enabled = False: optType(1).Enabled = False: optType(2).Enabled = False
        cboItemType.Enabled = True

        cmdAdvice.Visible = False
        UCAdvice.Top = 0
        UCAdvice.Height = fraAdvice.Height
    ElseIf mlngFun = 0 Then
        If mint场合 = 1 Then '护士场合,只允许添加（其他类）文本
            optType(0).Enabled = False: optType(1).Enabled = False: optType(2).Enabled = True
            optType(2).Value = True
            optExecutor(1).Value = True
        ElseIf mint场合 = 0 Then
            optExecutor(0).Value = True
        End If
        cboItemType.Enabled = True
    End If
    
    If mlngFun = 1 Then
        optType(0).Value = True
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False

        mvItem.医嘱IDs = mstr医嘱IDs
        Call ShowAdvice
    ElseIf mlngFun = 2 Then
        Me.Caption = "修改路径外项目"
        Call LoadData
        Call SetFormFace
        fraEPR.Enabled = False: vsEPR.Enabled = False
    Else
        UCAdvice.Height = fraAdvice.Height - cmdAdvice.Height - 60
    End If

    Call Cbo.SetListHeight(cboItems, 500)
    Call Cbo.SetListWidth(cboItems.Hwnd, 3500)
    Call SetFormFace
End Sub

Private Sub LoadData()
'功能：读取路径外项目的相关信息
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, str医嘱IDs As String, lngRow As Long
    Dim j As Long, str项目结果 As String, str缺省结果 As String
    Dim arrtmp As Variant, arrtmp2 As Variant
 
    strSql = "Select a.分类,a.项目内容,a.执行者,a.项目结果,a.添加原因,a.图标ID,a.变异原因,b.病人医嘱ID " & vbNewLine & _
            " From 病人路径执行 A,病人路径医嘱 B Where a.ID = [1] And a.id = b.路径执行id(+)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng执行ID)
    For i = 1 To rsTmp.RecordCount
        If i = 1 Then
            Call Cbo.Locate(cboItemType, rsTmp!分类)
            txtItem.Text = "" & rsTmp!项目内容
            txtReason.Text = "" & rsTmp!添加原因
            
            If Not IsNull(rsTmp!变异原因) Then
                lngRow = vsVariation.FindRow(CStr(rsTmp!变异原因)) '按编码查找rowdata
                If lngRow > 0 Then
                    vsVariation.Row = lngRow
                    vsVariation.TopRow = lngRow
                    Call vsVariation_DblClick
                End If
            End If
            
            mvItem.图标ID = Val("" & rsTmp!图标ID)
            If mvItem.图标ID <> 0 Then Call DrawPicture(GetPathIcon(mvItem.图标ID))
            
            If mblnUseExecute Then
                If Not IsNull(rsTmp!执行者) Then
                    If rsTmp!执行者 = 1 Then
                        optExecutor(0).Value = True
                    Else
                        optExecutor(1).Value = True
                    End If
                End If
                
                If Not IsNull(rsTmp!项目结果) Then
                    arrtmp = Split(rsTmp!项目结果, vbTab)
                    str项目结果 = arrtmp(0)
                    If UBound(arrtmp) > 0 Then str缺省结果 = arrtmp(1)
                    
                    arrtmp = Split(str项目结果, ",")
                    With vsResult
                        .Rows = .FixedRows + UBound(arrtmp) + 2 '最后增加一个空行
                        For j = 0 To UBound(arrtmp)
                            arrtmp2 = Split(arrtmp(j), "|")
                            lngRow = .FixedRows + j
                            .TextMatrix(lngRow, col执行结果) = arrtmp2(0)
                            If arrtmp2(0) = str缺省结果 Then
                                .TextMatrix(lngRow, col缺省结果) = 1
                                Call vsResult_AfterEdit(lngRow, col缺省结果)
                            End If
                            
                            If UBound(arrtmp2) > 0 Then
                                mrsNature.Filter = "编码='" & arrtmp2(1) & "'"
                                Set .Cell(flexcpPicture, lngRow, col执行图标) = imgNature.ListImages(Val(arrtmp2(1))).Picture
                                .TextMatrix(lngRow, col结果性质) = mrsNature!名称
                            End If
                        Next
                    End With
                End If
            End If
        End If
        
        If Not IsNull(rsTmp!病人医嘱id) Then
            str医嘱IDs = str医嘱IDs & "," & rsTmp!病人医嘱id
        End If
        rsTmp.MoveNext
    Next
    
    If str医嘱IDs <> "" Then
        optType(0).Value = True
        mvItem.医嘱IDs = Mid(str医嘱IDs, 2)
        Call ShowAdvice
    Else
        strSql = "Select a.病历名称 From 电子病历记录 a Where a.路径执行id = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng执行ID)
        If rsTmp.RecordCount = 0 Then
            optType(2).Value = True
        Else
            optType(1).Value = True
            vsEPR.Rows = vsEPR.FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                vsEPR.TextMatrix(vsEPR.FixedRows - 1 + i, 0) = rsTmp!病历名称
                rsTmp.MoveNext
            Next
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init基本执行结果()
'功能：初始并加载基本执行结果
    Dim strSql As String, rsTmp As ADODB.Recordset
          
    '读取结果性质集
    On Error GoTo errH
    strSql = "Select 编码,名称 From 路径结果性质 Order by 编码"
    Set mrsNature = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsNature, strSql, Me.Caption)
    strSql = ""
    Do While Not mrsNature.EOF
        strSql = strSql & "|" & mrsNature!编码 & "-" & mrsNature!名称
        mrsNature.MoveNext
    Loop
    vsResult.ColData(col结果性质) = Mid(strSql, 2)
    
    '读取可用结果集
    strSql = "Select A.编码,A.名称,Nvl(基本,0) as 基本,B.名称 as 性质" & _
        " From 路径常见结果 A,路径结果性质 B" & _
        " Where A.末级=1 And Nvl(A.性质,0)=B.编码(+)" & _
        " Order by A.编码"
    Set mrsResult = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrsResult, strSql, Me.Caption)
    
    If mlngFun <> 2 Then
        mrsResult.Filter = "基本=1"
        If Not mrsResult.EOF Then
            vsResult.Rows = vsResult.FixedRows + 1
            Call SetResultInput(vsResult.FixedRows, mrsResult)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 1 Then
        If InStr(GetInsidePrivs(p住院病历管理), ";病历书写;") = 0 Then
            MsgBox "你没有病历书写的权限，不能生成包含病历的路径项目。", vbInformation + vbOKOnly, gstrSysName
            optType(2).Value = True
            Exit Sub
        End If
    ElseIf Index = 0 Then
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0 Then
            MsgBox "你没有医嘱下达的权限，不能生成医嘱类的路径项目。", vbInformation + vbOKOnly, gstrSysName
            optType(2).Value = True
            Exit Sub
        End If
    End If
    
    Call SetFormFace
    
    If Visible Then
        If Index = 0 Then
            If cmdAdvice.Enabled Then cmdAdvice.SetFocus
        ElseIf Index = 1 Then
            vsEPR.SetFocus
        End If
    End If
End Sub


Private Sub picBottom_GotFocus()
    If cmdOK.Enabled And cmdOK.Visible Then
        cmdOK.SetFocus
    ElseIf cmdCancel.Enabled And cmdCancel.Visible Then
        cmdCancel.SetFocus
    End If
End Sub

Private Sub SetFormFace()
'功能：根据内容属性设置界面的可见内容和尺寸
    fraAdvice.Enabled = optType(0).Value: fraAdvice.Visible = fraAdvice.Enabled
    fraEPR.Enabled = optType(1).Value: fraEPR.Visible = fraEPR.Enabled
    If optType(1).Value Then
        If gobjEmr Is Nothing Then
            fraERPType.Visible = False
            optEPRType(0).Value = True
        Else
            fraERPType.Visible = True
            optEPRType(1).Value = True
        End If
    Else
        fraERPType.Visible = False
        fraItemKind.Move txtItem.Left + txtItem.Width + 120
    End If
    If fraERPType.Visible Then
        txtItem.Width = txtReason.Width - 1600
        fraItemKind.Left = txtItem.Left + txtItem.Width + 120
    Else
        txtItem.Width = txtReason.Width
        fraItemKind.Left = txtItem.Left + txtItem.Width + 120
    End If
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vPoint As POINTAPI
    
    On Error GoTo errH
    
    If img16.ListImages.count = 0 Then
        strSql = "Select ID,Nvl(性质,0) as 性质 From 临床路径图标 Order by 性质 Desc,ID"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            img16.ListImages.Add , "_" & IIf(rsTmp!性质 = 1, 1, -1) * rsTmp!ID, GetPathIcon(rsTmp!ID)
            img16.ListImages(img16.ListImages.count).Tag = CStr(rsTmp!ID) '要CStr
            rsTmp.MoveNext
        Loop
        cbsIcon.AddImageList img16
    End If
    
    Set objPopup = cbsIcon.Add("Popup", xtpBarPopup)
    objPopup.SetPopupToolBar True
    objPopup.Width = 260
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, -1, "清除项目图标")
        objControl.Flags = xtpFlagControlStretched
        For i = 1 To img16.ListImages.count
            Set objControl = .Add(xtpControlButton, img16.ListImages(i).Tag, "")
            If i = 1 Then
                objControl.BeginGroup = True
            ElseIf Val(Mid(img16.ListImages(i).Key, 2)) < 0 Then
                If Val(Mid(img16.ListImages(i - 1).Key, 2)) > 0 Then
                    objControl.BeginGroup = True
                End If
            End If
        Next
    End With
    
    vPoint.X = (picIcon.Left + picIcon.Width) / Screen.TwipsPerPixelX + 1
    vPoint.Y = picIcon.Top / Screen.TwipsPerPixelY
    ClientToScreen Me.Hwnd, vPoint
    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_GotFocus()
    Call zlControl.TxtSelAll(txtItem)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If TypeName(ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    Dim str变异原因 As String
    Dim strTmp As String
    
    '数据检查
    If Trim(txtItem.Text) = "" Then
        MsgBox "请输入路径项目的内容。", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem.Text) > txtItem.MaxLength Then
        MsgBox "项目内容中最多允许 " & txtItem.MaxLength \ 2 & " 个汉字或者 " & txtItem.MaxLength & "个字符。", vbInformation, gstrSysName
        txtItem.SetFocus: Exit Sub
    End If

    '如果有数据，则必须选择一个变异原因，变异说明可以不输
    If vsVariation.Rows > vsVariation.FixedRows Then
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, col变异选择) = 1 Then
                    mvItem.变异原因 = .RowData(i)
                    str变异原因 = Mid(.TextMatrix(i, col变异原因), InStr(.TextMatrix(i, col变异原因), "-") + 1)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                MsgBox "请选择一种变异原因。", vbInformation, gstrSysName
                If vsVariation.Enabled Then vsVariation.SetFocus
                Exit Sub
            End If
        End With
    End If

    '如果变异原因是其他则要求必须填写变异说明
    If str变异原因 = "其他" Or str变异原因 = "其它" Then
        If Trim(txtReason.Text) = "" Then
            MsgBox "变异原因为其他的，必须填写变异说明。", vbInformation, gstrSysName
            If txtReason.Enabled Then txtReason.SetFocus
            Exit Sub
        End If
    End If

    If zlCommFun.ActualLen(txtReason.Text) > txtReason.MaxLength Then
        MsgBox "添加原因中最多允许 " & txtReason.MaxLength \ 2 & " 个汉字或者 " & txtReason.MaxLength & "个字符。", vbInformation, gstrSysName
        txtReason.SetFocus: Exit Sub
    End If


    '检查医嘱
    If mlngFun = 0 Then
        If optType(0).Value Then
            If mvItem.医嘱IDs = "" Then
                MsgBox "没有定义当前项目所对应的医嘱内容。", vbInformation, gstrSysName
                If cmdAdvice.Enabled Then cmdAdvice.SetFocus
                Exit Sub
            End If
        Else
            If mvItem.医嘱IDs <> "" Then
                gcnOracle.RollbackTrans
                mvItem.医嘱IDs = ""
                mdatPathOut = "0:00:00": mlng阶段ID = 0: mlng天数 = 0
                If cboItems.Tag <> mPP.当前阶段ID & "_" & mPP.当前天数 Then Call LoadCurrentItems
            End If
        End If

        '检查病历
        If optType(1).Value Then
            With vsEPR
                strFilter = ""
                mvItem.病历IDs = "": mvItem.新版病历IDs = ""
                For i = .FixedRows To .Rows - 1
                    If .RowData(i) <> 0 Then
                         strTmp = .RowData(i) '格式：(NEW/OLD)|ID/ID:编辑方式
                        If Split(strTmp, "|")(0) = "OLD" Then
                            mvItem.病历IDs = mvItem.病历IDs & "," & Split(strTmp, "|")(1) '其中存的ID:编辑方式
                        Else
                            mvItem.新版病历IDs = mvItem.新版病历IDs & "," & Split(strTmp, "|")(1)
                        End If
                        If InStr(strFilter & ",", "," & .TextMatrix(i, 0) & ",") = 0 Then
                            strFilter = strFilter & "," & .TextMatrix(i, 0)
                        Else
                            MsgBox "指定了重复的病历文件""" & .TextMatrix(i, 0) & """。", vbInformation, gstrSysName
                            .Row = i: Call .ShowCell(.Row, .Col)
                            .SetFocus: Exit Sub
                        End If
                    End If
                Next
                mvItem.病历IDs = Mid(mvItem.病历IDs, 2)
                mvItem.新版病历IDs = Mid(mvItem.新版病历IDs, 2)
                If mvItem.病历IDs = "" And mvItem.新版病历IDs = "" Then
                    MsgBox "请指定项目所对应的病历文件。", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End With
        Else
            mvItem.病历IDs = "": mvItem.新版病历IDs = ""
        End If
        
    End If

    '执行结果
    If fraExecute.Visible Then
        With vsResult
            strFilter = ""
            mvItem.项目结果 = ""
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col执行结果)) <> "" Then
                    If InStr(strFilter & ",", "," & .TextMatrix(i, col执行结果) & ",") = 0 _
                       And InStr(strFilter, "," & .TextMatrix(i, col执行结果) & "|") = 0 Then
                        strFilter = strFilter & "," & .TextMatrix(i, col执行结果)
                        If .TextMatrix(i, col结果性质) <> "" Then
                            mrsNature.Filter = "名称='" & .TextMatrix(i, col结果性质) & "'"
                            strFilter = strFilter & "|" & mrsNature!编码
                        End If
                    Else
                        MsgBox "指定了重复的执行结果""" & .TextMatrix(i, col执行结果) & """。", vbInformation, gstrSysName
                        .Row = i: Call .ShowCell(.Row, .Col)
                        .SetFocus: Exit Sub
                    End If

                    '缺省结果
                    If Val(.TextMatrix(i, col缺省结果)) <> 0 Then
                        mvItem.项目结果 = .TextMatrix(i, col执行结果)
                    End If
                End If
            Next
            strFilter = Mid(strFilter, 2)
            If strFilter = "" Then
                MsgBox "请指定项目所对应的执行结果。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If mvItem.项目结果 = "" Then
                MsgBox "请指定项目所对应的缺省结果。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            mvItem.项目结果 = strFilter & vbTab & mvItem.项目结果
        End With
    End If


    '其他数据收集
    mvItem.项目内容 = txtItem.Text
    If fraExecute.Visible Then
        mvItem.执行者 = IIf(optExecutor(0).Value, 1, 2)
    Else
        mvItem.执行者 = 0
    End If

    Call SaveData
      
    mblnOK = True
    Unload Me
End Sub

Private Sub SaveData()
'参数:bln评估结果 -是否修改评估结果
    Dim i As Long, blnTrans As Boolean
    Dim arrSQL As Variant, arrFile As Variant, arrItem As Variant
    Dim strAddDate As String, strDate As String
    Dim lng病历ID As Long, str病人病历IDs As String, str病历IDs As String
    Dim lngPosition As Long '插入位置id
    Dim str任务IDs As String
    Dim strTmp As String, strSql As String
    
    Dim colNewDoc As New Collection  '存放新版病历任务ID
    Dim rsTmp As ADODB.Recordset
    Dim lng阶段ID As Long
    Dim lng天数 As Long
    Dim lng标记 As Long, lng自动执行 As Long
    
    lngPosition = cboItems.ItemData(cboItems.ListIndex)
    
    On Error GoTo errH
    If mlngFun = 1 Or mlngFun = 2 Then
        If lngPosition = mlng执行ID Then lngPosition = 0    '插入位置不发生变动
        strSql = "Zl_病人路径生成_Update(" & mlng执行ID & ",'" & cboItemType.Text & "','" & mvItem.项目内容 & _
                 "'," & mvItem.执行者 & ",'" & mvItem.项目结果 & "'," & ZVal(mvItem.图标ID) & ",'" & txtReason.Text & "','" & mvItem.变异原因 & _
                 "'," & lngPosition & "," & mPP.病人路径ID & "," & mPP.当前阶段ID & "," & mPP.当前天数 & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        If mlngFun = 1 And mdatPathOut <> "0:00:00" And mdatPathOut < mPP.当前日期 Then
            strSql = "Select 阶段id, 日期,天数" & vbNewLine & _
                    "From 病人路径执行 A" & vbNewLine & _
                    "Where ID = [1] And Exists" & vbNewLine & _
                    " (Select 1 From 病人路径评估 B Where b.路径记录id = a.路径记录id And b.阶段id = a.阶段id And b.日期 = a.日期)"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng执行ID)
            If rsTmp.RecordCount = 1 Then
                strDate = "To_Date('" & Format(rsTmp!日期, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                strSql = "Zl_病人路径评估_Insert(" & 2 & "," & mPP.病人路径ID & "," & rsTmp!阶段id & _
                "," & strDate & "," & rsTmp!天数 & ",NULL," & 1 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,0,NULL,1)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If
    Else
        strAddDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        arrFile = Array()
        If mvItem.病历IDs <> "" Then
            arrFile = Split(mvItem.病历IDs, ",")
            For i = 0 To UBound(arrFile)
                str病历IDs = str病历IDs & "," & Split(arrFile(i), ":")(0)

                lng病历ID = zlDatabase.GetNextID("电子病历记录")
                str病人病历IDs = str病人病历IDs & "," & lng病历ID
            Next
            str病历IDs = Mid(str病历IDs, 2)
            str病人病历IDs = Mid(str病人病历IDs, 2)
        End If
        
        If mvItem.新版病历IDs <> "" And Not gobjEmr Is Nothing Then '新版
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                strTmp = "": str任务IDs = ""
                For i = 0 To UBound(Split(mvItem.新版病历IDs, ","))
                    strTmp = "<parameter><antetypeid>" & Split(mvItem.新版病历IDs, ",")(i) & "</antetypeid><patient>" & mPati.病人ID & "</patient></parameter>"
                    '记录集包含字段：原型ID,任务ID,生成时间,起始时间,终止时间；
                    On Error Resume Next
                    Set rsTmp = gobjEmr.MakeBeforTask(strTmp)
                    Err.Clear: On Error GoTo 0
                    If rsTmp.State <> adStateClosed Then
                        If rsTmp.RecordCount = 1 Then
                            str任务IDs = str任务IDs & "," & rsTmp!任务ID
                        End If
                    End If
                Next
                str任务IDs = Mid(str任务IDs, 2)
                colNewDoc.Add str任务IDs, "C" & (colNewDoc.count + 1) '记录返回的任务ID,避免事务提交失败,删除生成成功的新版病历
            End If
        End If
        
        If mvItem.医嘱IDs <> "" Then
            If mdatPathOut < mPP.当前日期 Then
                '获取径外医嘱项目应当插入的阶段ID,天数,日期(补录)
                strDate = "To_Date('" & Format(mdatPathOut, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                lng阶段ID = mlng阶段ID
                lng天数 = mlng天数
                lng标记 = 1 '补录标记
                lng自动执行 = 1 '自动执行
            ElseIf mdatPathOut > mPP.当前日期 Then
                lng标记 = 2 '暂存标记
            End If
        End If
        
        If lng标记 = 0 Or lng标记 = 2 Then '医嘱开始日期=当前日期\>当前日期
            strDate = "To_Date('" & Format(mPP.当前日期, "yyyy-MM-dd") & "','YYYY-MM-DD')"
            lng阶段ID = mPP.当前阶段ID
            lng天数 = mPP.当前天数
        End If
        
        arrSQL = Array()
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人路径生成_Insert(1," & mPati.病人ID & "," & mPati.主页ID & ",'0'," & mPati.科室ID & "," & _
                 mPP.病人路径ID & "," & lng阶段ID & _
                 "," & strDate & "," & lng天数 & _
                 ",'" & cboItemType.Text & "',Null,'" & mvItem.医嘱IDs & "','" & str病历IDs & "','" & str病人病历IDs & "'" & _
                 ",'" & UserInfo.姓名 & "'," & strAddDate & ",'" & mvItem.项目内容 & _
                 "'," & mvItem.执行者 & ",'" & mvItem.项目结果 & "'," & ZVal(mvItem.图标ID) & ",'" & txtReason.Text & "','" & mvItem.变异原因 & _
                 "'," & IIf(lng自动执行 = 0, "NULL", "1") & ",Null,Null,Null,Null," & lngPosition & "," & IIf(mint场合 = 0, 1, 2) & ",'" & str任务IDs & IIf(lng标记 = 0, "')", "'," & lng标记 & ")")
       
        If mdatPathOut <> "0:00:00" And mdatPathOut < mPP.当前日期 Then
         '修改评估结果和变异原因
            If CheckPathIsEvaluated(mPP.病人路径ID, lng阶段ID, mdatPathOut) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人路径评估_Insert(" & 2 & "," & mPP.病人路径ID & "," & lng阶段ID & _
                "," & strDate & "," & lng天数 & ",NULL," & 1 & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,0,NULL,1)"
            End If
        End If
        
        If mvItem.医嘱IDs = "" Then
            gcnOracle.BeginTrans: blnTrans = True
        Else
            '医嘱已插入并开始了事务
            '1.医嘱校对的插入过程SQL
            blnTrans = True
            For i = 0 To UBound(marrSQL)
                zlDatabase.ExecuteProcedure CStr(marrSQL(i)), Me.Caption
            Next
        End If

        '2.产生病人路径数据
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next

        '3.产生病历文件RTF数据
        If mvItem.病历IDs <> "" Then
            For i = 0 To UBound(arrFile)
                arrItem = Split(arrFile(i), ":")
                If arrItem(1) = 0 Or arrItem(1) = 1 Then     '全文编辑方式的病历
                    lng病历ID = Split(str病人病历IDs, ",")(i)
                    Call ReadRTFData(lng病历ID, edtEditor)
                    Call SaveRTFData(lng病历ID, mPati.病人ID, mPati.主页ID, 0, edtEditor)   '暂不支持选择指定婴儿的病历
                End If
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        '发送医嘱新开的消息
        Call ZLHIS_CIS_001(mclsMipModule, mPati.病人ID, mPati.主页ID, mPati.病区ID, mPati.科室ID)
    End If


    '清除医嘱ID,以便退出时判断是否回退事务
    mvItem.医嘱IDs = ""
    Exit Sub
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        '回滚新版病历
        If Not gobjEmr Is Nothing Then
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then Set gobjEmr = Nothing
            If Not gobjEmr Is Nothing Then
                For i = 1 To colNewDoc.count
                    strTmp = "<parameter><taskid>" & colNewDoc("C" & i) & "</taskid></parameter>"
                    On Error Resume Next
                    Call gobjEmr.DeleteTask(strTmp)
                    Err.Clear: On Error GoTo 0
                Next
            End If
        End If
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = False And mlngFun = 1 Then
        MsgBox "请选择变异原因后点确定按钮。", vbInformation, gstrSysName
        If vsVariation.Enabled And vsVariation.Visible Then vsVariation.SetFocus
        Cancel = 1: Exit Sub
    End If

    If Not mblnOK And txtItem.Text <> "" Then
        If mlngFun = 2 Then
            If MsgBox("确实要放弃路径外项目的修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        Else
            If MsgBox("确实要放弃路径外项目的添加吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If

    End If

    If Not mrsResult Is Nothing Then
        If mrsResult.State = 1 Then mrsResult.Close
        Set mrsResult = Nothing
    End If
    If Not mrsNature Is Nothing Then
        If mrsNature.State = 1 Then mrsNature.Close
        Set mrsNature = Nothing
    End If
    Set marrSQL = Nothing
    Set mclsMipModule = Nothing
    If mvItem.医嘱IDs <> "" And mlngFun = 0 Then gcnOracle.RollbackTrans
End Sub

Private Sub cmdAdvice_Click()
'功能：编辑项目所对应的医嘱
    Dim strItem As String, i As Long
    Dim strAdviceOfItem As String
    Dim arrSQL As Variant
    Dim datPathOut As Date
    Dim rsTmp As ADODB.Recordset
    
    Call gobjKernel.ShowAdviceEdit(mfrmParent, mint场合, 2, mPati.病人ID, mPati.主页ID, mvItem.医嘱IDs, mPP.当前日期, arrSQL, strAdviceOfItem, , , , mclsMipModule, mdatPathOut)
        
    If mvItem.医嘱IDs <> strAdviceOfItem Then
        mvItem.医嘱IDs = strAdviceOfItem
        marrSQL = arrSQL
    End If
    
    If mdatPathOut <> CDate(0) And mdatPathOut < mPP.当前日期 Then
        Set rsTmp = GetPatiPathAppend(mPP.病人路径ID, mdatPathOut)   '获取要插入的阶段
        If rsTmp.RecordCount > 0 Then
            mlng阶段ID = Val(rsTmp!阶段id & "")
            mlng天数 = Val(rsTmp!天数 & "")
        End If
        If cboItems.Tag <> mlng阶段ID & "_" & mlng天数 Then Call LoadCurrentItems
    Else
        mlng阶段ID = 0: mlng天数 = 0
        If cboItems.Tag <> mPP.当前阶段ID & "_" & mPP.当前天数 Then Call LoadCurrentItems
    End If
    '刷新显示
    Call ShowAdvice
End Sub

Private Function ShowAdvice() As Boolean
'功能：显示路径项目对应的医嘱内容

    If mvItem.医嘱IDs = "" Then
        Call UCAdvice.ShowAdvice(2, "", 0, "")
        ShowAdvice = True: Exit Function
    Else
        Call UCAdvice.ShowAdvice(2, "", 0, mvItem.医嘱IDs)
    End If
    
    If mlngFun <> 2 Then
        '缺省项目内容
        If txtItem.Text = "" Then
            txtItem.Text = UCAdvice.GetAdviceTitle
        End If
    End If
End Function


Private Sub txtReason_GotFocus()
    Call zlControl.TxtSelAll(txtReason)
End Sub

Private Sub txtVariation_GotFocus()
    Call zlControl.TxtSelAll(txtVariation)
End Sub

Private Sub txtVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim i As Long, strtxt As String
        strtxt = "*" & UCase(Trim(txtVariation.Text)) & "*"
        With vsVariation
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col变异原因) Like strtxt Or .RowData(i) Like strtxt Or .Cell(flexcpData, i, col变异原因) Like strtxt Then
                    .SetFocus
                    .Row = i
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
End Sub

Private Sub vsEPR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsEPR_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsEPR_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsEPR
        If NewCol = 0 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsEPR_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsEPR
        If optEPRType(0).Value Then  '表示新版电子病历未安装或安装不正常（走老版流程）
            strSql = "Select A.ID,Decode(A.种类,2,'住院病历',4,'护理病历',5,'疾病证明报告',6,'知情文件') as 种类," & _
                " A.编号,A.名称,A.说明,A.保留 as 编辑方式 From 病历文件列表 A" & _
                " Where A.种类 IN(2,4,5,6) And Nvl(A.保留,0) IN(0,1,2) And A.通用 IN(1,2)" & _
                " Order by A.种类,A.编号"
        Else
            '新版流程
            If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                Set gobjEmr = Nothing
                MsgBox "电子病历服务器不在线或导航台登录时未能成功连接电子病历服务器!", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            Else
                'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                '记录集包含字段：分类编号，分类名称，分组名称，ID，编号，名称，说明
                On Error Resume Next
                Set rsTmp = gobjEmr.GetAntetypeList("")
                Err.Clear: On Error GoTo 0
                strSql = Rec.ToSQL(rsTmp)
            End If
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "病历文件", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有病历文件数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetEPRInput(Row, rsTmp)
            Call EPREnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsEPR_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsEPR
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("确实要清除该行病历文件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsEPR_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsEPR_KeyPress(KeyAscii As Integer)
    With vsEPR
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EPREnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsEPR_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsEPR_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsEPR_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsEPR.EditSelStart = 0
    vsEPR.EditSelLength = zlCommFun.ActualLen(vsEPR.EditText)
End Sub

Private Sub vsEPR_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsEPR
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call EPREnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call EPREnterNextCell
            Else
                strInput = UCase(.EditText)
                If optEPRType(0).Value Then
                    strSql = "Select A.ID,Decode(A.种类,2,'住院病历',4,'护理病历',5,'疾病证明报告',6,'知情文件') as 种类," & _
                        " A.编号,A.名称,A.说明,A.保留 as 编辑方式 From 病历文件列表 A" & _
                        " Where A.种类 IN(2,4,5,6) And Nvl(A.保留,0) IN(0,1,2)" & _
                        " And A.通用 IN(1,2) And (A.编号 Like [1] Or A.名称 Like [2] Or zlSpellCode(A.名称) Like [2])" & _
                        " Order by A.种类,A.编号"
                    
                Else
                '新版流程
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                        MsgBox "电子病历服务器不在线或导航台登录时未能成功连接电子病历服务器!", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        'gobjEmr.GetAntetypeList(byref strParameter as string) as Adodb.RecordSet
                        '记录集包含字段：分类编号，分类名称，分组名称，ID，编号，名称，说明
                        On Error Resume Next
                        Set rsTmp = gobjEmr.GetAntetypeList("")
                        Err.Clear: On Error GoTo 0
                        strSql = Rec.ToSQL(rsTmp)
                        strSql = "select A.* from (" & strSql & ") A where A.编号 Like [1] Or A.名称 Like [2] Or zlSpellCode(A.名称) Like [2]  order by 分类编号,编号"
                    End If
                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "病历文件", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的病历文件。", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetEPRInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call EPREnterNextCell(True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub SetEPRInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理病历文件的输入
    Dim strItem As String, i As Long
    
    With vsEPR
        For i = 1 To rsInput.RecordCount
            If i > 1 Then
                .AddItem "", lngRow + 1
                lngRow = lngRow + 1
            End If
            If optEPRType(0).Value Then
               '旧版
                .RowData(lngRow) = "OLD|" & Val(rsInput!ID) & ":" & rsInput!编辑方式
            Else
               .RowData(lngRow) = "NEW|" & Val(rsInput!ID)
            End If
            .TextMatrix(lngRow, 0) = rsInput!名称
            .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)
            
            strItem = strItem & "、" & rsInput!名称
            
            rsInput.MoveNext
        Next
        
        '缺省项目内容
        If txtItem.Text = "" Then txtItem.Text = "书写" & Mid(strItem, 2)
        
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        
    End With
End Sub

Private Sub EPREnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long
    
    With vsEPR
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub


Private Sub vsResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsResult
        If Col = col结果性质 Then
            .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.TextMatrix(Row, Col))
            If .TextMatrix(Row, Col) <> "" Then
                mrsNature.Filter = "名称='" & .TextMatrix(Row, Col) & "'"
                Set .Cell(flexcpPicture, .Row, col执行图标) = imgNature.ListImages(Val(mrsNature!编码)).Picture
            Else
                Set .Cell(flexcpPicture, .Row, col执行图标) = Nothing
            End If
        ElseIf Col = col缺省结果 Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                For i = .FixedRows To .Rows - 1
                    If i <> Row Then .TextMatrix(i, col缺省结果) = 0
                Next
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsResult
        If Not ResultCellEditable(NewRow, NewCol) Then
            .FocusRect = flexFocusLight
            .ComboList = ""
        ElseIf NewCol = col执行结果 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        ElseIf NewCol = col结果性质 Then
            .ComboList = .ColData(NewCol)
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsResult_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsResult
        '如果用子查询，则数据树形顺序不对，需要特别排序
        strSql = _
            " Select A.编码 As ID,A.上级 As 上级id,A.编码,A.名称,A.简码,Nvl(A.末级,0) As 末级,B.名称 As 性质" & _
            " From 路径常见结果 A,路径结果性质 B Where Nvl(A.性质,0)=B.编码(+)" & _
            " Start With A.上级 Is Null Connect By Prior A.编码=A.上级"
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "常见结果", True, "", "", False, True, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有常见结果数据可以选择。", vbInformation, gstrSysName
            End If
        Else
            Call SetResultInput(Row, rsTmp)
            Call ResultEnterNextCell
        End If
    End With
End Sub

Private Sub vsResult_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If Col = col结果性质 Then
        With vsResult
            If .TextMatrix(Row, Col) <> "" Then
                For i = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = i: Exit For
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub vsResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
        
    With vsResult
        If KeyCode = vbKeyDelete Then
            If .Col = col结果性质 Then
                .TextMatrix(.Row, .Col) = ""
                Set .Cell(flexcpPicture, .Row, col执行图标) = Nothing
            ElseIf .TextMatrix(.Row, col执行结果) <> "" Then
                If MsgBox("确实要清除该行结果吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                End If
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsResult_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsResult_KeyPress(KeyAscii As Integer)
    With vsResult
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ResultEnterNextCell
        ElseIf KeyAscii = Asc(",") Then
            KeyAscii = 0
        ElseIf .Col = col执行结果 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsResult_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" '使按钮状态进入输入状态
            End If
        End If
    End With
End Sub

Private Sub vsResult_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If KeyAscii = Asc(",") Then KeyAscii = 0
End Sub

Private Sub vsResult_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsResult.EditSelStart = 0
    vsResult.EditSelLength = zlCommFun.ActualLen(vsResult.EditText)
End Sub

Private Sub vsResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ResultCellEditable(Row, Col) Then Cancel = True
End Sub

Private Sub vsResult_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsResult
        If Col = col执行结果 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call ResultEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ResultEnterNextCell
            Else
                strInput = UCase(.EditText)
                strSql = "Select A.编码 as ID,A.编码,A.名称,A.简码,B.名称 as 性质" & _
                    " From 路径常见结果 A,路径结果性质 B" & _
                    " Where Nvl(A.性质,0)=B.编码(+) And A.末级=1" & _
                    " And (A.编码 Like [1] Or A.名称 Like [2] Or A.简码 Like [2])" & _
                    " Order by A.编码"
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "常见结果", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%")
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    Call SetResultInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call ResultEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col结果性质 Then
            If mblnReturn Then Call ResultEnterNextCell
            mblnReturn = False
        End If
    End With
End Sub


Private Sub SetResultInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理项目结果的输入
    Dim i As Long
    
    With vsResult
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If
                
                .TextMatrix(lngRow, col执行结果) = rsInput!名称
                
                '处理结果性质
                If Not IsNull(rsInput!性质) Then
                    mrsNature.Filter = "名称='" & rsInput!性质 & "'"
                    Set .Cell(flexcpPicture, lngRow, col执行图标) = imgNature.ListImages(Val(mrsNature!编码)).Picture
                    .TextMatrix(lngRow, col结果性质) = rsInput!性质
                End If
                
                If i = 1 And lngRow = .FixedRows Then
                    .TextMatrix(lngRow, col缺省结果) = 1
                    Call vsResult_AfterEdit(lngRow, col缺省结果)
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col执行结果) = .EditText
        End If
        .Cell(flexcpData, lngRow, col执行结果) = .TextMatrix(lngRow, col执行结果)
                
        '始终保持一空行
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
    End With
End Sub

Private Sub ResultEnterNextCell()
    With vsResult
        If .Col + 1 <= .Cols - 1 Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = col执行结果
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function ResultCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    ResultCellEditable = True
    
    With vsResult
        If lngCol = col执行图标 Then
            ResultCellEditable = False
        ElseIf lngCol > col执行结果 And .TextMatrix(lngRow, col执行结果) = "" Then
            ResultCellEditable = False
        ElseIf lngCol = col结果性质 And .TextMatrix(lngRow, col执行结果) <> "" Then
            '字典中的结果性质不允许更改,手工输入的才允许
            If .TextMatrix(lngRow, col结果性质) <> "" Then
                mrsResult.Filter = "名称='" & .TextMatrix(lngRow, col执行结果) & "'"
                If Not mrsResult.EOF Then
                    If NVL(mrsResult!性质) = .TextMatrix(lngRow, col结果性质) Then
                        ResultCellEditable = False
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub InitVariation()
'功能：读取并加载变异原因数据
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    
    strSql = "Select b.名称 As 分类, a.编码, a.名称, a.简码" & vbNewLine & _
            "From 变异常见原因 A, 变异常见原因 B" & vbNewLine & _
            "Where a.末级 = 1 And a.上级 = b.编码 and a.性质=1" & vbNewLine & _
            "order by b.名称"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
    With vsVariation
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If rsTmp.RecordCount > 0 Then
            .MergeCol(col变异分类) = True
            .Rows = .FixedRows + rsTmp.RecordCount
            '缺省不选择
            Set .Cell(flexcpPicture, .FixedRows, col变异选择, .Rows - 1, col变异选择) = imgNature.ListImages("UnSelected").Picture
            .Cell(flexcpPictureAlignment, .FixedRows, col变异选择, .Rows - 1, col变异选择) = flexPicAlignCenterCenter
            For i = .FixedRows To rsTmp.RecordCount
                .Cell(flexcpData, i, col变异选择) = 0
                
                .RowData(i) = CStr(rsTmp!编码)    '主键
                .TextMatrix(i, col变异分类) = rsTmp!分类
                .TextMatrix(i, col变异原因) = rsTmp!编码 & "-" & rsTmp!名称
                .Cell(flexcpData, i, col变异原因) = "" & rsTmp!简码
                rsTmp.MoveNext
            Next
        End If
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    vsVariation.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsVariation_DblClick()
    With vsVariation
        If .Row >= .FixedRows Then
            Dim i As Long
            .Redraw = flexRDNone
            If .Cell(flexcpData, .Row, col变异选择) = 0 Then
                Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("Selected").Picture
                .Cell(flexcpData, .Row, col变异选择) = 1
                For i = .FixedRows To .Rows - 1
                    If i <> .Row Then
                        If .Cell(flexcpData, i, col变异选择) = 1 Then
                            Set .Cell(flexcpPicture, i, col变异选择) = imgNature.ListImages("UnSelected").Picture
                            .Cell(flexcpData, i, col变异选择) = 0
                        End If
                    End If
                Next
            Else
                Set .Cell(flexcpPicture, .Row, col变异选择) = imgNature.ListImages("UnSelected").Picture
                .Cell(flexcpData, .Row, col变异选择) = 0
            End If
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub vsVariation_GotFocus()
    If vsVariation.Row < vsVariation.FixedRows And vsVariation.Rows > vsVariation.FixedRows Then vsVariation.Row = vsVariation.FixedRows
End Sub

Private Sub vsVariation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call vsVariation_DblClick
    End If
End Sub

Private Sub LoadCurrentItems()
'功能:加载路径项目
    Dim strSql      As String
    Dim rsData      As Recordset
    Dim strTmp      As String
    Dim lng阶段ID   As Long
    Dim lng天数     As Long
    
    cboItems.Clear
    strSql = "Select Decode(a.项目内容, Null, b.项目内容, a.项目内容) As 项目内容, a.Id, a.项目序号" & vbNewLine & _
             "From 病人路径执行 A, 临床路径项目 B" & vbNewLine & _
             "Where a.路径记录id = [1] And a.阶段id = [2] And a.项目id = b.Id(+) And a.天数 = [3] And a.分类 = [4]" & vbNewLine & _
             "Order By a.项目序号"
    strTmp = cboItemType.List(cboItemType.ListIndex)
    On Error GoTo errH
    If mlng阶段ID <> 0 And mlng天数 <> 0 Then '指定某个阶段
        lng阶段ID = mlng阶段ID
        lng天数 = mlng天数
        cboItems.Tag = lng阶段ID & "_" & lng天数   '
    Else
        lng阶段ID = mPP.当前阶段ID
        lng天数 = mPP.当前天数
        cboItems.Tag = lng阶段ID & "_" & lng天数
    End If
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.病人路径ID, lng阶段ID, lng天数, strTmp)
    
    Call Cbo.AddData(cboItems, rsData)

    If mlngFun = 0 Or mlngFun = 1 Then  '0-直接添加,1-医嘱新开时添加,2-修改路径外项目
        rsData.Filter = "项目内容='路径外项目'"
        If rsData.RecordCount < 1 Then
            cboItems.AddItem "路径外项目"
            cboItems.ItemData(cboItems.ListCount - 1) = 0
        End If
        If cboItems.ListCount > 0 Then Call Cbo.Locate(cboItems, "路径外项目", False)
    ElseIf mlngFun = 2 Then
        If Not Cbo.Locate(cboItems, mlng执行ID, True) Then
            If cboItems.ListCount = 0 Then
                cboItems.AddItem ""
                cboItems.ItemData(cboItems.ListCount - 1) = 0
            End If
            If cboItems.ListCount > 0 Then cboItems.ListIndex = cboItems.ListCount - 1
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPatiPathAppend(ByVal lng记录ID As Long, ByVal dat日期 As Date) As ADODB.Recordset
'功能:根据日期获取有效阶段和天数
    Dim strSql      As String
    
    strSql = "Select 阶段id,天数" & vbNewLine & _
            "From (Select a.阶段id, a.天数, a.登记时间" & vbNewLine & _
            "       From 病人路径执行 A" & vbNewLine & _
            "       Where a.路径记录id = [1] And a.日期 = [2]" & vbNewLine & _
            "       Order By a.登记时间 Desc)" & vbNewLine & _
            "Where Rownum < 2"
        
    On Error GoTo errH
    Set GetPatiPathAppend = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng记录ID, dat日期)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPathIsEvaluated(ByVal lng路径记录ID As Long, ByVal lng阶段ID As Long, ByVal dat日期 As Date) As Boolean
'-------------------------------------------------------------------------------------------------------------------------------------
'功能：检查当前病人当前路径表单下,某一天,某一阶段是否评估
'参数：
'   lng路径记录Id-病人路径记录ID
'   lng阶段ID  -当前阶段ID
'   dat日期-当前日期
'返回:True-评估;False-未评估
'-------------------------------------------------------------------------------------------------------------------------------------
     Dim strSql As String, rsTmp As Recordset
     
     strSql = "Select 1 From 病人路径评估 Where 路径记录ID=[1] And 阶段ID =[2] And 日期 =[3] "
     
     On Error GoTo errH
     Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径记录ID, lng阶段ID, dat日期)
     
     CheckPathIsEvaluated = rsTmp.RecordCount > 0
     Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
