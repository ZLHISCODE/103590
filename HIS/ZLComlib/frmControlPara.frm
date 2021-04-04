VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmControlPara 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsControl 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7950
      _cx             =   14023
      _cy             =   7329
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
      BackColor       =   -2147483634
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483634
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483643
      FloodColor      =   192
      SheetBorder     =   -2147483637
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmControlPara.frx":0000
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
Attribute VB_Name = "frmControlPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrControl As String

Private Enum ColIndex
    col输入项目 = 0
    col是否禁用
    col是否必输项
    col光标是否进入
End Enum
Private Sub Form_Resize()
    vsControl.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function zlOpenControlFrm(ByVal strControl As String) As Boolean
    Call InitPara(strControl)
End Function

Private Function InitPara(ByVal strControl As String) As Boolean
    'strControl:是否禁用;光标是否跳过,是否必输项|....
    Dim Arr() As String, subArr() As String, intRow As Integer

    Arr() = Split(strControl & "|", "|")
    With vsControl
        .Clear 1
        .Rows = UBound(Arr) + 1
        For intRow = 1 To .Rows - 1
            subArr() = Split(Arr(intRow - 1) & ",", ",")
            .TextMatrix(intRow, .ColIndex("输入项目")) = subArr(0)
            If UBound(subArr) = 4 Then
                .TextMatrix(intRow, .ColIndex("是否禁用")) = IIf(Val(subArr(1)) = 1, "√", "")
                .TextMatrix(intRow, .ColIndex("是否必输项")) = IIf(Val(subArr(2)) = 1, "√", "")
                .TextMatrix(intRow, .ColIndex("光标是否进入")) = IIf(Val(subArr(3)) = 1, "√", "")
            End If
        Next
    End With
End Function

Public Function GetInputItemControlSet() As String
    Dim i As Integer, strTmp As String
    With vsControl
        For i = 1 To .Rows - 1
            strTmp = strTmp & "|" & .TextMatrix(i, .ColIndex("输入项目"))
            strTmp = strTmp & "," & IIf(.TextMatrix(i, .ColIndex("是否禁用")) = "√", 1, 0)
            strTmp = strTmp & "," & IIf(.TextMatrix(i, .ColIndex("是否必输项")) = "√", 1, 0)
            strTmp = strTmp & "," & IIf(.TextMatrix(i, .ColIndex("光标是否进入")) = "√", 1, 0)
        Next
    End With
    GetInputItemControlSet = Mid(strTmp, 2)
End Function

Private Function ChangePara() As Boolean
    With vsControl
        If .Col = col是否禁用 Then
            .TextMatrix(.Row, .ColIndex("是否禁用")) = IIf(.TextMatrix(.Row, .ColIndex("是否禁用")) = "", "√", "")
            If .TextMatrix(.Row, .ColIndex("是否禁用")) = "√" Then
                .TextMatrix(.Row, .ColIndex("光标是否进入")) = ""
                .TextMatrix(.Row, .ColIndex("是否必输项")) = ""
            End If
        Else
            If .TextMatrix(.Row, .ColIndex("是否禁用")) = "√" Then Exit Function
            If .Col = col光标是否进入 And .TextMatrix(.Row, .ColIndex("是否必输项")) = "√" Then Exit Function
            .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "√", "")
            If .Col = col是否必输项 And .TextMatrix(.Row, .Col) = "√" Then .TextMatrix(.Row, .ColIndex("光标是否进入")) = "√"
        End If
    End With
End Function

Private Sub vsControl_DblClick()
    Call ChangePara
End Sub

Private Sub vsControl_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then KeyAscii = 0
    Call ChangePara
End Sub
