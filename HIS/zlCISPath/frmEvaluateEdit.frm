VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEvaluateEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XXXX评估设置"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6420
   Icon            =   "frmEvaluateEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4065
      Index           =   1
      Left            =   105
      TabIndex        =   28
      Top             =   1245
      Visible         =   0   'False
      Width           =   6165
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   300
         Index           =   1
         Left            =   5115
         TabIndex        =   18
         Top             =   2145
         Width           =   990
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "更新(&U)"
         Height          =   300
         Index           =   1
         Left            =   4125
         TabIndex        =   17
         Top             =   2145
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   300
         Index           =   1
         Left            =   3135
         TabIndex        =   16
         Top             =   2145
         Width           =   990
      End
      Begin VB.OptionButton optCombine 
         Caption         =   "满足任一条件"
         Height          =   180
         Index           =   1
         Left            =   1500
         TabIndex        =   20
         Top             =   2280
         Width           =   1380
      End
      Begin VB.OptionButton optCombine 
         Caption         =   "满足所有条件"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   2280
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.ComboBox cboResult 
         Height          =   300
         Left            =   4080
         TabIndex        =   15
         Top             =   1770
         Width           =   2025
      End
      Begin VB.ComboBox cboCond 
         Height          =   300
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1770
         Width           =   735
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   435
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1770
         Width           =   2865
      End
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   1710
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   6165
         _cx             =   10874
         _cy             =   3016
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
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluateEdit.frx":058A
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
      Begin VSFlex8Ctl.VSFlexGrid vsCond 
         Height          =   1560
         Left            =   0
         TabIndex        =   21
         Top             =   2490
         Width           =   6165
         _cx             =   10874
         _cy             =   2752
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
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluateEdit.frx":05D6
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
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "项目"
         Height          =   180
         Left            =   75
         TabIndex        =   12
         Top             =   1830
         Width           =   360
      End
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4065
      Index           =   0
      Left            =   105
      TabIndex        =   27
      Top             =   1245
      Width           =   6165
      Begin VB.CommandButton cmdName 
         Caption         =   "…"
         Height          =   240
         Left            =   4530
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "选择变异常见原因(*键)"
         Top             =   45
         Width           =   255
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Index           =   0
         Left            =   4950
         TabIndex        =   9
         Top             =   1305
         Width           =   1100
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "更新(&U)"
         Height          =   350
         Index           =   0
         Left            =   4950
         TabIndex        =   8
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   350
         Index           =   0
         Left            =   4950
         TabIndex        =   7
         Top             =   615
         Width           =   1100
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmEvaluateEdit.frx":0676
         Left            =   5310
         List            =   "frmEvaluateEdit.frx":0680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   15
         Width           =   795
      End
      Begin VB.TextBox txtName 
         Height          =   555
         Left            =   840
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   15
         Width           =   3975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMark 
         Height          =   2310
         Left            =   0
         TabIndex        =   10
         Top             =   1725
         Width           =   6165
         _cx             =   10874
         _cy             =   4075
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
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluateEdit.frx":0690
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         ExplorerBar     =   8
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
      Begin VSFlex8Ctl.VSFlexGrid vsResult 
         Height          =   1050
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   3975
         _cx             =   7011
         _cy             =   1852
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
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
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmEvaluateEdit.frx":06F4
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "指标结果"
         Height          =   180
         Left            =   75
         TabIndex        =   5
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型"
         Height          =   180
         Left            =   4890
         TabIndex        =   3
         Top             =   75
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "指标名称"
         Height          =   180
         Left            =   75
         TabIndex        =   0
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4035
      TabIndex        =   22
      Top             =   5475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   23
      Top             =   5475
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6420
      TabIndex        =   24
      Top             =   0
      Width           =   6420
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXXX评估"
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
         Left            =   1065
         TabIndex        =   26
         Top             =   120
         Width           =   810
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "  设置用于XXXX评估的指标，包括指标名称，指标值，以及指标之间的计算关系，用于辅助计算评估结果。"
         Height          =   360
         Left            =   1065
         TabIndex        =   25
         Top             =   360
         Width           =   5175
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   105
         Picture         =   "frmEvaluateEdit.frx":073F
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin MSComctlLib.TabStrip tbsFunc 
      Height          =   4485
      Left            =   45
      TabIndex        =   29
      Top             =   885
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   7911
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "评估指标"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "计算条件"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   795
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluateEdit.frx":2281
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvaluateEdit.frx":281B
            Key             =   "Num"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEvaluateEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CheckDataValid(EvalInfo As TYPE_PATH_EVAL, EvalType As Integer, Cancel As Boolean)
Private mvEval As TYPE_PATH_EVAL
Private mColItems As Collection
Private mintType As Integer '1-导入评估,2-阶段评估

Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean

Private Enum ENUM_COND_COL
    col显示 = 0
    col指标ID = 1
    col项目ID = 2
    col关系式 = 3
    col条件值 = 4
End Enum

Public Function ShowEdit(frmParent As Object, intType As Integer, vEval As TYPE_PATH_EVAL, Optional colItems As Collection) As Boolean
'参数：colItems=如果是阶段的评估定义，则传入该阶段已定义的项目集合
    mintType = intType
    mvEval = vEval
    Set mColItems = colItems
    
    Me.Show 1, frmParent
    
    If mblnOK Then vEval = mvEval
    
    Set mvEval.指标集 = Nothing
    Set mvEval.条件集 = Nothing
    Set mColItems = Nothing
    
    ShowEdit = mblnOK
End Function

Private Sub cboCond_Click()
    cboResult.Text = ""
End Sub

Private Sub cboResult_GotFocus()
    Call zlControl.TxtSelAll(cboResult)
End Sub

Private Sub cboResult_KeyPress(KeyAscii As Integer)
    If cboCond.Text <> "Like" Then KeyAscii = 0
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    If Index = 0 Then
        If Not SetMarkInput(False) Then Exit Sub
    ElseIf Index = 1 Then
        If Not SetCondInput(False) Then Exit Sub
    End If
    
    Call SetFuncEnabled
    mblnChange = True
End Sub

Private Sub cmdAdd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Button = 0 Then
        If Index = 0 Then
            strTip = "将当前输入的指标信息增加到下面表格中"
        ElseIf Index = 1 Then
            strTip = "将当前设置的计算条件增加到下面表格中"
        End If
    End If
    ZLCommFun.ShowTipInfo cmdAdd(Index).Hwnd, strTip
End Sub

Private Sub cmdDelete_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Button = 0 Then
        If Index = 0 Then
            strTip = "删除下面表格中当前选择的指标行"
        ElseIf Index = 1 Then
            strTip = "删除下面表格中当前选择的计算条件"
        End If
    End If
    ZLCommFun.ShowTipInfo cmdDelete(Index).Hwnd, strTip
End Sub

Private Sub cmdName_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    strSql = "Select 编码 as ID,上级 as 上级ID,编码,名称,简码,Nvl(末级,0) as 末级" & _
        " From 变异常见原因 Where 性质 = 1 Start With 上级 Is NULL Connect by Prior 编码=上级"
    vPoint = zlControl.GetCoordPos(txtName.Hwnd, -15, 15)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "变异常见原因", True, "", "", True, True, True, _
        vPoint.X, vPoint.Y, txtName.Height, blnCancel, False, True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有变异常见原因可以选择。", vbInformation, gstrSysName
        End If
    Else
        txtName.Text = rsTmp!名称
    End If
    
    txtName.SetFocus: Call txtName_GotFocus
End Sub

Private Sub cmdUpdate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Button = 0 Then
        If Index = 0 Then
            strTip = "将当前输入的指标信息更新到下面表格当前指标行中"
        ElseIf Index = 1 Then
            strTip = "将当前设置的计算条件更新到下面表格当前条件行中"
        End If
    End If
    ZLCommFun.ShowTipInfo cmdUpdate(Index).Hwnd, strTip
End Sub

Private Sub cmdUpdate_Click(Index As Integer)
    If Index = 0 Then
        If Not SetMarkInput(True) Then Exit Sub
    ElseIf Index = 1 Then
        If Not SetCondInput(True) Then Exit Sub
    End If
    
    Call SetFuncEnabled
    mblnChange = True
End Sub

Private Sub cmdDelete_Click(Index As Integer)
    Dim lngRow As Long, i As Long
    
    If Index = 0 Then
        With vsMark
            If MsgBox("确实要删除该指标吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            lngRow = .Row
            
            '删除计算条件中跟该指标相关的
            For i = vsCond.Rows - 1 To vsCond.FixedRows Step -1
                If Val(vsCond.TextMatrix(i, col指标ID)) = .RowData(lngRow) Then
                    vsCond.RemoveItem i
                    If i <= vsCond.Rows - 1 Then
                        vsCond.Row = i
                    ElseIf vsCond.Rows > vsCond.FixedRows Then
                        vsCond.Row = vsCond.Rows - 1
                    End If
                End If
            Next
            
            .RemoveItem lngRow
            If lngRow <= .Rows - 1 Then
                .Row = lngRow
            ElseIf .Rows > .FixedRows Then
                .Row = .Rows - 1
            End If
            .ShowCell .Row, .Col
        End With
    ElseIf Index = 1 Then
        With vsCond
            If MsgBox("确实要删除该计算条件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            lngRow = .Row
            .RemoveItem lngRow
            
            If lngRow <= .Rows - 1 Then
                .Row = lngRow
            ElseIf .Rows > .FixedRows Then
                .Row = .Rows - 1
            End If
            .ShowCell .Row, .Col
        End With
    End If
    
    Call SetFuncEnabled
    mblnChange = True
End Sub

Private Function SetCondInput(ByVal blnUpdate As Boolean) As Boolean
    Dim lngRow As Long, strShow As String, i As Long
    
    If cboCond.ListIndex = -1 Then
        MsgBox "请指定计算关系。", vbInformation, gstrSysName
        cboCond.SetFocus: Exit Function
    End If
    If cboResult.Text = "" Then
        MsgBox "请指定计算条件值。", vbInformation, gstrSysName
        cboResult.SetFocus: Exit Function
    End If
    If ZLCommFun.ActualLen(cboResult.Text) > 50 Then
        MsgBox "计算条件值的内容太长，最多允许 25 个汉字或者 50 个字符。", vbInformation, gstrSysName
        cboResult.SetFocus: Exit Function
    End If
    If vsItem.Cell(flexcpData, vsItem.Row, 0) = 2 Then
        If Not IsNumeric(cboResult.Text) Then
            MsgBox "指定的计算条件值不是数值类型。", vbInformation, gstrSysName
            cboResult.SetFocus: Exit Function
        End If
    End If
    
    With vsCond
        strShow = "[" & txtItem.Text & "] " & cboCond.Text & " [" & cboResult.Text & "]"
        For i = .FixedRows To .Rows - 1
            If Not blnUpdate Or i <> .Row Then
                If strShow = .TextMatrix(i, col显示) Then
                    MsgBox "该计算条件已经存在。", vbInformation, gstrSysName
                    cboCond.SetFocus: Exit Function
                End If
            End If
        Next
        
        If blnUpdate Then
            lngRow = .Row
        Else
            .AddItem "": lngRow = .Rows - 1
        End If
        
        .TextMatrix(lngRow, col显示) = strShow
        If vsItem.TextMatrix(vsItem.Row, 1) = "评估指标" Then
            .TextMatrix(lngRow, col指标ID) = vsItem.RowData(vsItem.Row)
            .TextMatrix(lngRow, col项目ID) = 0
        Else
            .TextMatrix(lngRow, col指标ID) = 0
            .TextMatrix(lngRow, col项目ID) = vsItem.RowData(vsItem.Row)
        End If
        .TextMatrix(lngRow, col关系式) = cboCond.Text
        .TextMatrix(lngRow, col条件值) = cboResult.Text
        
        .Row = lngRow: .Col = 0
        .ShowCell .Row, .Col
    End With
    
    SetCondInput = True
End Function

Private Function SetMarkInput(ByVal blnUpdate As Boolean) As Boolean
    Dim strResult As String, lngRow As Long
    Dim i As Long, j As Long
    
    If txtName.Text = "" Then
        MsgBox "请输入指标名称。", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    If ZLCommFun.ActualLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "指标名称太长，最多允许 " & txtName.MaxLength \ 2 & " 个汉字或者 " & txtName.MaxLength & " 个字符。", vbInformation, gstrSysName
        txtName.SetFocus: Exit Function
    End If
    
    With vsResult
        strResult = ""
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                If cboType.ListIndex = 1 Then
                    If Not IsNumeric(.TextMatrix(i, 0)) Then
                        .Row = i: .Col = 0: .ShowCell .Row, .Col
                        MsgBox "输入的指标结果不是数值类型。", vbInformation, gstrSysName
                        .SetFocus: Exit Function
                    End If
                End If
                For j = .FixedRows To .Rows - 1
                    If j <> i And .TextMatrix(j, 0) = .TextMatrix(i, 0) Then
                        .Row = i: .Col = 0: .ShowCell .Row, .Col
                        MsgBox "输入了重复的指标结果。", vbInformation, gstrSysName
                        .SetFocus: Exit Function
                    End If
                Next
                
                If Val(.TextMatrix(i, 1)) <> 0 Then
                    strResult = strResult & vbCrLf & "●" & .TextMatrix(i, 0)
                Else
                    strResult = strResult & vbCrLf & "○" & .TextMatrix(i, 0)
                End If
            End If
        Next
        strResult = Mid(strResult, 3)
        
        If strResult = "" Then
            MsgBox "请依次输入该指标的结果。", vbInformation, gstrSysName
            .SetFocus: Exit Function
        End If
        If InStr(strResult, "●") = 0 Then
            MsgBox "请指定指标的缺省结果。", vbInformation, gstrSysName
            .SetFocus: Exit Function
        End If
        If ZLCommFun.ActualLen(Replace(Replace(Replace(strResult, vbCrLf, ","), "●", ""), "○", "")) > 500 Then
            MsgBox "指标的结果太多，请适当进行调整。", vbInformation, gstrSysName
            .SetFocus: Exit Function
        End If
    End With
    
    With vsMark
        For i = .FixedRows To .Rows - 1
            If Not blnUpdate Or i <> .Row Then
                If .TextMatrix(i, 1) = txtName.Text Then
                    MsgBox "该指标已经存在。", vbInformation, gstrSysName
                    txtName.SetFocus: Exit Function
                End If
            End If
        Next
        
        If blnUpdate Then
            lngRow = .Row
        Else
            .AddItem ""
            .RowData(.Rows - 1) = zlDatabase.GetNextID("路径评估指标")
            lngRow = .Rows - 1
        End If
        
        .TextMatrix(lngRow, 1) = txtName.Text
        
        If cboType.ListIndex = 0 Then
            .Cell(flexcpData, lngRow, 1) = 1
            Set .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Text").Picture
        ElseIf cboType.ListIndex = 1 Then
            .Cell(flexcpData, lngRow, 1) = 2
            Set .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Num").Picture
        End If
        .Cell(flexcpPictureAlignment, lngRow, 1) = 1
        
        .TextMatrix(lngRow, 2) = strResult
        .AutoSize 2
        
        .Row = lngRow: .Col = .FixedCols
        .ShowCell .Row, .Col
    End With
    
    '清除指标输入
    txtName.Text = ""
    vsResult.Rows = vsResult.FixedRows
    vsResult.Rows = vsResult.FixedRows + 1
    vsResult.Row = vsResult.FixedRows
    txtName.SetFocus
    
    SetMarkInput = True
End Function

Private Sub SetFuncEnabled()
    cmdUpdate(0).Enabled = vsMark.Row >= vsMark.FixedRows
    cmdDelete(0).Enabled = vsMark.Row >= vsMark.FixedRows
    
    cmdAdd(1).Enabled = txtItem.Text <> ""
    cmdUpdate(1).Enabled = vsCond.Row >= vsCond.FixedRows
    cmdDelete(1).Enabled = vsCond.Row >= vsCond.FixedRows
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim blnCancel As Boolean, i As Long
    Dim arrResult As Variant, j As Long
    Dim strResult As String, strDefault As String
    
    If vsMark.Rows = vsMark.FixedRows Then
        If vsCond.Rows > vsCond.FixedRows Then
            If MsgBox("还没有定义用于评估的指标，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                tbsFunc.Tabs(1).Selected = True
                txtName.SetFocus: Exit Sub
            End If
        Else
            MsgBox "还没有定义用于评估的指标。", vbInformation, gstrSysName
            tbsFunc.Tabs(1).Selected = True
            txtName.SetFocus: Exit Sub
        End If
    End If
    If vsCond.Rows = vsCond.FixedRows Then
        MsgBox "还没有定义用于评估的计算条件。", vbInformation, gstrSysName
        tbsFunc.Tabs(2).Selected = True
        cboCond.SetFocus: Exit Sub
    End If
    
    '收集数据
    Set mvEval.指标集 = New Collection
    With vsMark
        For i = .FixedRows To .Rows - 1
            vEvalMark.ID = .RowData(i)
            vEvalMark.序号 = i - .FixedRows + 1
            vEvalMark.评估指标 = .TextMatrix(i, 1)
            vEvalMark.指标类型 = .Cell(flexcpData, i, 1)
            
            arrResult = Split(.TextMatrix(i, 2), vbCrLf)
            strResult = ""
            For j = 0 To UBound(arrResult)
                strResult = strResult & "," & Mid(arrResult(j), 2)
                If Left(arrResult(j), 1) = "●" Then
                    strDefault = Mid(arrResult(j), 2)
                End If
            Next
            vEvalMark.指标结果 = Mid(strResult, 2) & vbTab & strDefault
            
            mvEval.指标集.Add vEvalMark
        Next
    End With
    
    Set mvEval.条件集 = New Collection
    With vsCond
        For i = .FixedRows To .Rows - 1
            vEvalCond.指标ID = Val(.TextMatrix(i, col指标ID))
            vEvalCond.项目ID = Val(.TextMatrix(i, col项目ID))
            vEvalCond.关系式 = .TextMatrix(i, col关系式)
            vEvalCond.条件值 = .TextMatrix(i, col条件值)
            vEvalCond.条件组合 = IIf(optCombine(0).Value, 1, 2)
            
            mvEval.条件集.Add vEvalCond
        Next
    End With
    
    RaiseEvent CheckDataValid(mvEval, mintType, blnCancel)
    If blnCancel Then Exit Sub
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        If TypeName(Me.ActiveControl) <> "VSFlexGrid" And Not Me.ActiveControl Is txtName Then
            KeyAscii = 0
            ZLCommFun.PressKey vbKeyTab
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim arrResult As Variant, strTemp As String
    Dim i As Long, j As Long
    
    mblnOK = False
    
    Me.Caption = Replace(Me.Caption, "XXXX", IIf(mintType = 1, "导入", "阶段"))
    lblInfo.Caption = Replace(lblInfo.Caption, "XXXX", IIf(mintType = 1, "导入", "阶段"))
    lblNote.Caption = Replace(lblNote.Caption, "XXXX", IIf(mintType = 1, "导入", "阶段"))
    cboType.ListIndex = 0
    vsResult.Rows = vsResult.FixedRows + 1
    vsResult.Col = 1: vsResult.Col = 0
    
    '显示评估指标
    With vsMark
        .Rows = .FixedRows
        If Not mvEval.指标集 Is Nothing Then
            For i = 1 To mvEval.指标集.count
                vEvalMark = mvEval.指标集(i)
                
                .AddItem ""
                
                .RowData(i) = vEvalMark.ID
                .TextMatrix(.Rows - 1, 1) = vEvalMark.评估指标
                .Cell(flexcpData, .Rows - 1, 1) = vEvalMark.指标类型
                
                If vEvalMark.指标类型 = 1 Then
                    Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("Text").Picture
                ElseIf vEvalMark.指标类型 = 2 Then
                    Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("Num").Picture
                End If
                .Cell(flexcpPictureAlignment, .Rows - 1, 1) = 1
                
                strTemp = ""
                arrResult = Split(Split(vEvalMark.指标结果, vbTab)(0), ",")
                For j = 0 To UBound(arrResult)
                    If arrResult(j) = Split(vEvalMark.指标结果, vbTab)(1) Then
                        strTemp = strTemp & vbCrLf & "●" & arrResult(j)
                    Else
                        strTemp = strTemp & vbCrLf & "○" & arrResult(j)
                    End If
                Next
                .TextMatrix(.Rows - 1, 2) = Mid(strTemp, 3)
            Next
            If .Rows > .FixedRows Then
                .Row = .FixedRows: .Col = .FixedCols
            End If
            .AutoSize 2
        End If
    End With
    
    '显示计算条件
    Call ShowMarkList
    With vsCond
        .Rows = .FixedRows
        If Not mvEval.条件集 Is Nothing Then
            For i = 1 To mvEval.条件集.count
                vEvalCond = mvEval.条件集(i)
                
                .AddItem ""
                .TextMatrix(.Rows - 1, col显示) = "[" & GetItemName(vEvalCond.指标ID, vEvalCond.项目ID) & "] " & vEvalCond.关系式 & " [" & vEvalCond.条件值 & "]"
                .TextMatrix(.Rows - 1, col指标ID) = vEvalCond.指标ID
                .TextMatrix(.Rows - 1, col项目ID) = vEvalCond.项目ID
                .TextMatrix(.Rows - 1, col关系式) = vEvalCond.关系式
                .TextMatrix(.Rows - 1, col条件值) = vEvalCond.条件值
                
                optCombine(vEvalCond.条件组合 - 1).Value = True
            Next
            If .Rows > .FixedRows Then
                .Row = .FixedRows: .Col = .FixedCols
            End If
        End If
    End With
    
    tbsFunc.Tabs(1).Selected = True
    Call SetFuncEnabled
    mblnChange = False
End Sub

Private Function GetItemName(ByVal lng指标ID As Long, ByVal lng项目ID As Long) As String
    Dim i As Long
    
    With vsItem
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 1) = "评估指标" Then
                If .RowData(i) = lng指标ID Then
                    GetItemName = .TextMatrix(i, 0)
                    Exit Function
                End If
            Else
                If .RowData(i) = lng项目ID Then
                    GetItemName = .TextMatrix(i, 0)
                    Exit Function
                End If
            End If
        Next
    End With
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK And mblnChange Then
        If MsgBox("数据已被更改，确实要放弃退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If

End Sub

Private Sub tbsFunc_Click()
    If tbsFunc.SelectedItem.Index = 1 Then
        fraFunc(0).Visible = True
        fraFunc(1).Visible = False
    ElseIf tbsFunc.SelectedItem.Index = 2 Then
        fraFunc(0).Visible = False
        fraFunc(1).Visible = True
        
        Call ShowMarkList
    End If
    
    Call SetFuncEnabled
End Sub

Private Sub txtItem_GotFocus()
    Call zlControl.TxtSelAll(txtItem)
End Sub

Private Sub txtItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "条件运算关系说明：" & vbCrLf & _
        "  =     等于条件值" & vbCrLf & _
        "  <>    不等于条件值" & vbCrLf & _
        "  >     大于条件值" & vbCrLf & _
        "  >=    大于或者等于条件值" & vbCrLf & _
        "  <     小于条件值" & vbCrLf & _
        "  <=    小于或者等于条件值" & vbCrLf & _
        "  Like  条件值中的*号表示匹配任意内容"
    ZLCommFun.ShowTipInfo txtItem.Hwnd, strTip, True
End Sub

Private Sub txtName_GotFocus()
    Call zlControl.TxtSelAll(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdName.Enabled And cmdName.Visible Then
            Call cmdName_Click
        End If
    ElseIf KeyAscii = 13 Then
        If txtName.Text = "" Then
            KeyAscii = 0
        Else
            KeyAscii = 0
            
            strInput = UCase(txtName.Text)
            strSql = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 变异常见原因 a,变异常见原因 b" & _
                " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 And (a.编码 Like [1] Or a.名称 Like [2] Or a.简码 Like [2])" & _
                " Order by 分类,a.编码"
            vPoint = zlControl.GetCoordPos(txtName.Hwnd, -15, 15)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "变异常见原因", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, txtName.Height, blnCancel, False, True, _
                strInput & "%", gstrLike & strInput & "%")
            If Not blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                If Not rsTmp Is Nothing Then
                    txtName.Text = rsTmp!名称
                End If
                Call ZLCommFun.PressKey(vbKeyTab)
            End If
        End If
    End If
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim arrResult As Variant, i As Long
    
    With vsItem
        If NewRow <> OldRow And NewRow >= .FixedRows Then
            .ForeColorSel = .CellForeColor
        
            txtItem.Text = .TextMatrix(NewRow, 0)
            
            cboCond.Clear
            If .Cell(flexcpData, NewRow, 0) = 1 Then
                cboCond.AddItem "="
                cboCond.AddItem "<>"
                cboCond.AddItem "Like"
            ElseIf .Cell(flexcpData, NewRow, 0) = 2 Then
                cboCond.AddItem "="
                cboCond.AddItem "<>"
                cboCond.AddItem ">"
                cboCond.AddItem ">="
                cboCond.AddItem "<"
                cboCond.AddItem "<="
            End If
            
            cboResult.Clear
            arrResult = Split(.Cell(flexcpData, NewRow, 1), ",")
            For i = 0 To UBound(arrResult)
                cboResult.AddItem arrResult(i)
            Next
            
            Call SetFuncEnabled
        End If
    End With
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vsMark_DblClick()
    If vsMark.MouseRow >= vsMark.FixedRows Then
        Call vsMark_KeyPress(13)
    End If
End Sub

Private Sub vsMark_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim arrResult As Variant, i As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        With vsMark
            lngRow = .Row
            If lngRow >= .FixedRows Then
                txtName.Text = .TextMatrix(lngRow, 1)
                cboType.ListIndex = Decode(.Cell(flexcpData, lngRow, 1), 1, 0, 2, 1)
                
                arrResult = Split(.TextMatrix(lngRow, 2), vbCrLf)
                With vsResult
                    .Rows = .FixedRows
                    .Rows = .FixedRows + (UBound(arrResult) + 1) + 1
                    For i = 0 To UBound(arrResult)
                        .TextMatrix(.FixedRows + i, 0) = Mid(arrResult(i), 2)
                        If Left(arrResult(i), 1) = "●" Then
                            .TextMatrix(.FixedRows + i, 1) = 1
                        End If
                    Next
                    .Row = .FixedRows: .Col = 0
                End With
            End If
        End With
    End If
End Sub

Private Sub vsMark_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    If Button = 0 Then
        strTip = "操作提示：" & vbCrLf & "1.双击或者回车提取当前行指标进行编辑" & vbCrLf & "2.拖动指标行头部可以改变行的顺序"
    End If
    ZLCommFun.ShowTipInfo vsMark.Hwnd, strTip, True
End Sub

Private Sub vsResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsResult
        If Col = 1 Then
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                For i = .FixedRows To .Rows - 1
                    If i <> Row Then .TextMatrix(i, 1) = 0
                Next
            End If
        End If
    End With
End Sub

Private Sub vsResult_AfterMoveRow(ByVal Row As Long, Position As Long)
    mblnChange = True
End Sub

Private Sub vsResult_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsResult
        If NewCol = 0 Then
            .ComboList = "..."
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
        strSql = "Select 编码 as ID,上级 as 上级ID,编码,名称,简码,Nvl(末级,0) as 末级" & _
            " From 路径常见结果 Start With 上级 Is NULL Connect by Prior 编码=上级"
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "常见结果", True, "", "", False, False, True, _
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

Private Sub vsResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsResult
        If KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then .RemoveItem .Row
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
        ElseIf .Col = 0 Then
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
    vsResult.EditSelLength = ZLCommFun.ActualLen(vsResult.EditText)
End Sub

Private Sub vsResult_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 And vsResult.TextMatrix(Row, 0) = "" Then Cancel = True
End Sub

Private Sub vsResult_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsResult
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call ResultEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ResultEnterNextCell
            Else
                strInput = UCase(.EditText)
                strSql = "Select 编码 as ID,编码,名称,简码 From 路径常见结果" & _
                    " Where 末级=1 And (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2])" & _
                    " Order by 编码"
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
        End If
    End With
End Sub

Private Sub SetResultInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'功能：处理指标结果的输入
    Dim i As Long
    
    With vsResult
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If
            
                .TextMatrix(lngRow, 0) = rsInput!名称
                If i = 1 And lngRow = .FixedRows Then
                    .TextMatrix(lngRow, 1) = 1
                    Call vsResult_AfterEdit(lngRow, 1)
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, 0) = .EditText
        End If
        .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)
                
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
            .Row = .Row + 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub ShowMarkList()
'功能：将定义的指标和路径项目显示在条件待选列表中
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long
    
    With vsItem
        .Rows = .FixedRows
        .Rows = vsMark.Rows
        
        '指标部分
        For i = vsMark.FixedRows To .Rows - 1
            .RowData(i) = vsMark.RowData(i)
            .TextMatrix(i, 0) = vsMark.TextMatrix(i, 1)
            .TextMatrix(i, 1) = "评估指标"
            
            .Cell(flexcpData, i, 0) = vsMark.Cell(flexcpData, i, 1) '类型
            .Cell(flexcpData, i, 1) = Replace(Replace(Replace(vsMark.TextMatrix(i, 2), vbCrLf, ","), "●", ""), "○", "") '结果
            
            If vsMark.Cell(flexcpData, i, 1) = 1 Then
                Set .Cell(flexcpPicture, i, 0) = img16.ListImages("Text").Picture
            ElseIf vsMark.Cell(flexcpData, i, 1) = 2 Then
                Set .Cell(flexcpPicture, i, 0) = img16.ListImages("Num").Picture
            End If
            .Cell(flexcpPictureAlignment, i, 0) = 1
        Next
        '项目部份：阶段评估
        If mintType = 2 Then
            For i = 1 To mColItems.count
                vItem = mColItems(i)
                If vItem.项目结果 <> "" Then
                    .AddItem ""
                    .RowData(.Rows - 1) = vItem.ID
                    .TextMatrix(.Rows - 1, 0) = vItem.项目内容
                    .TextMatrix(.Rows - 1, 1) = "路径项目"
                    
                    .Cell(flexcpData, .Rows - 1, 0) = 1 '固定为文本类型
                    .Cell(flexcpData, .Rows - 1, 1) = CStr(Split(vItem.项目结果, vbTab)(0)) '不包含缺省结果标识
                    
                    Set .Cell(flexcpPicture, .Rows - 1, 0) = img16.ListImages("Text").Picture
                    .Cell(flexcpPictureAlignment, .Rows - 1, 0) = 1
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue '区别显示
                End If
            Next
        End If
        
        .Row = .FixedRows - 1
        txtItem.Text = ""
        cboCond.Clear
        cboResult.Clear
    End With
End Sub
