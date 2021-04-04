VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAutoFill 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "自动填充设置"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4860
   Icon            =   "frmAutoFill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   3360
      Width           =   4785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSetCol 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _cx             =   8070
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  2.以“自动拆分”那行作数据源填充空白的各列。"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   4140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  1.只支持字符类型的列，并且非“自动拆分”的列；"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   540
   End
End
Attribute VB_Name = "frmAutoFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_COLS As String = "列,,3,3000|自动填充,,3,1000,B|列序号,,0,0"

Private mobjSetCol As clsVSFlexGridEx
Public mstrColFillInfo As String
Public mstrResult As String

Private Sub cmdCancel_Click()
    mstrResult = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intRow As Integer
    
    mstrResult = ""
    For intRow = 1 To vsfSetCol.Rows - 1
        With vsfSetCol
            mstrResult = mstrResult & mdlPublic.FormatString("|[1],[2],[3]" _
                            , .TextMatrix(intRow, .ColIndex("列序号")) _
                            , .TextMatrix(intRow, .ColIndex("列")) _
                            , "" & Abs(Val(.TextMatrix(intRow, .ColIndex("自动填充")))))
        End With
    Next
    If mstrResult <> "" Then mstrResult = Mid$(mstrResult, 2)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitVSF
    Call FillData
End Sub

Private Sub InitVSF()
    Set mobjSetCol = New clsVSFlexGridEx
    
    With mobjSetCol
        .AppTemplate EM_Display, vsfSetCol, MSTR_COLS, "", True
        .Init
        .ColsReadonly = "列"
        .Binding.Editable = flexEDKbdMouse
    End With
End Sub

Private Sub FillData()
    Dim arrInfo As Variant, arrItem As Variant
    Dim i As Integer
    
    arrInfo = Split(mstrColFillInfo, "|")
    With vsfSetCol
        .Redraw = False
        For i = LBound(arrInfo) To UBound(arrInfo)
            .Rows = .Rows + 1
            arrItem = Split(arrInfo(i), ",")
            .TextMatrix(.Rows - 1, .ColIndex("列序号")) = arrItem(0)
            .TextMatrix(.Rows - 1, .ColIndex("列")) = arrItem(1)
            .TextMatrix(.Rows - 1, .ColIndex("自动填充")) = arrItem(2)
        Next
        .Redraw = True
    End With
End Sub
