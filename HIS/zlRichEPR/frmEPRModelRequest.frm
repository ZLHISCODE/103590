VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRModelRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应用条件"
   ClientHeight    =   4920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6615
   Icon            =   "frmEPRModelRequest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSave 
      Caption         =   "恢复(&R)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   5415
      TabIndex        =   12
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4380
      TabIndex        =   11
      Top             =   1770
      Width           =   1035
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -45
      TabIndex        =   9
      Top             =   345
      Width           =   6975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   5350
      TabIndex        =   7
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "将条件应用于当前文件的所有示范(&T)…"
      Height          =   350
      Left            =   150
      TabIndex        =   6
      Top             =   4485
      Width           =   3555
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "添加(&A)"
      Height          =   350
      Index           =   0
      Left            =   2055
      TabIndex        =   5
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "删除(&D)"
      Height          =   350
      Index           =   1
      Left            =   3105
      TabIndex        =   4
      Top             =   1770
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgItems 
      Height          =   3765
      Left            =   150
      TabIndex        =   0
      Top             =   645
      Width           =   1875
      _cx             =   3307
      _cy             =   6641
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vfgVal 
      Height          =   1110
      Left            =   2055
      TabIndex        =   1
      Top             =   645
      Width           =   4380
      _cx             =   7726
      _cy             =   1958
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vfgSel 
      Height          =   1980
      Left            =   2055
      TabIndex        =   2
      Top             =   2430
      Width           =   4395
      _cx             =   7752
      _cy             =   3492
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备选条件值: (双击或添加需要的备选值为条件值)"
      Height          =   180
      Left            =   2070
      TabIndex        =   13
      Top             =   2220
      Width           =   3960
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "条件项目:"
      Height          =   180
      Left            =   165
      TabIndex        =   10
      Top             =   435
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmEPRModelRequest.frx":000C
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "根据当前文件定义的种类，可以基于以下项目设置特定应用条件。"
      Height          =   180
      Left            =   525
      TabIndex        =   8
      Top             =   90
      Width           =   5220
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已选条件值: (当项目为以下值时可使用该示范)"
      Height          =   180
      Left            =   2055
      TabIndex        =   3
      Top             =   435
      Width           =   3780
   End
End
Attribute VB_Name = "frmEPRModelRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    条件项 = 0: 条件值
End Enum

Private mlngDemoId As Long      '当前示范ID
Private mintPower As Integer    '示范管理权范围
Private mblnOK As Boolean       '是否确认


'-----------------------------------------------------
'以下为外部公共程序
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, ByVal lngDemoId As Long, ByVal intPower As Integer) As Boolean
    '功能：显示本编辑窗体
    '参数： frmParent-父窗体
    '       lngDemoId-词句示范ID
    Dim rsTemp As New ADODB.Recordset
    mlngDemoId = lngDemoId: mintPower = intPower
    
    '装入可选分类数据
    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select 文件id, 性质 From 病历范文目录 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDemoId)
    If rsTemp.RecordCount <= 0 Then MsgBox "当前示范不存在(可能被其他用户删除)！", vbInformation, gstrSysName: Exit Function
    Me.cmdApply.Tag = rsTemp!文件ID
    If Val("" & rsTemp!性质) = 0 Then
        Me.Caption = "范文应用条件"
        Me.cmdApply.Caption = "将条件应用于当前文件的所有范文(&T)…"
    Else
        Me.Caption = "片段应用条件"
        Me.cmdApply.Caption = "将条件应用于当前文件的所有片段(&T)…"
    End If
    
    If RefList = False Then MsgBox "没有合适的条件项目！", vbInformation, gstrSysName: Exit Function
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

'-----------------------------------------------------
'以下为内部共用程序
'-----------------------------------------------------
Private Function RefList(Optional strItem As String) As Boolean
    '功能：刷新装入项目列表，并定位到指定项目
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select 名称 As 条件项, 简码 As 条件值 From Table(Cast(f_Segment_条件项([1]) As " & gstrDbOwner & ".t_Dic_Rowset))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDemoId)
    With Me.vfgItems
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mCol.条件值) = 0: .ColHidden(mCol.条件值) = True
        For lngCount = .FixedRows To .Rows - 1
            .Cell(flexcpFontBold, lngCount, mCol.条件项) = (.TextMatrix(lngCount, mCol.条件值) <> "")
            If .TextMatrix(lngCount, mCol.条件项) = strItem Then .Row = lngCount
        Next
        If .Row < .FixedRows Then .Row = .FixedRows
        Call vfgItems_AfterRowColChange(.Row, .Col, .Row, .Col)
    End With
    RefList = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cmdApply_Click()
    Err = 0: On Error GoTo ErrHand
    If MsgBox("真的将该条件应用于当前文件的所有示范吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Me.vfgItems.SetFocus: Exit Sub
    End If
    gstrSQL = "Zl_病历范文条件_Apply(" & mlngDemoId & "," & mintPower & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "范文条件"
    Me.vfgItems.SetFocus: Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim strItem As String
    If Index = 0 Then
        If Me.vfgSel.Rows < 1 Then MsgBox "没有可添加的值！", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        If Me.vfgSel.Row < 0 Then MsgBox "没有可添加的值！", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        Select Case Me.vfgItems.TextMatrix(Me.vfgItems.Row, mCol.条件项)
        Case "诊疗类别", "检查类型"
            If Me.vfgVal.Rows >= 1 Then MsgBox "该项目只能设置一个条件值！", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        End Select
        Me.vfgVal.AddItem Me.vfgSel.TextMatrix(Me.vfgSel.Row, 0)
        Me.vfgSel.RemoveItem Me.vfgSel.Row
        Me.vfgVal.SetFocus
    Else
        If Me.vfgVal.Rows < 1 Then MsgBox "没有可移去的值！", vbInformation, gstrSysName: Me.vfgSel.SetFocus: Exit Sub
        If Me.vfgVal.Row < 0 Then MsgBox "没有可移去的值！", vbInformation, gstrSysName: Me.vfgSel.SetFocus: Exit Sub
        Me.vfgSel.AddItem Me.vfgVal.TextMatrix(Me.vfgVal.Row, 0)
        Me.vfgVal.RemoveItem Me.vfgVal.Row
        Me.vfgSel.SetFocus
    End If
    If Me.vfgVal.Rows > 0 And Me.vfgVal.Row < 0 Then Me.vfgVal.Row = 0
    If Me.vfgSel.Rows > 0 And Me.vfgSel.Row < 0 Then Me.vfgSel.Row = 0
    
    Me.cmdEdit(0).Enabled = (Me.vfgSel.Rows > 0): Me.cmdEdit(1).Enabled = (Me.vfgVal.Rows > 0)
    Me.cmdSave(0).Enabled = True: Me.cmdSave(1).Enabled = True
End Sub

Private Sub cmdSave_Click(Index As Integer)
Dim strItem As String, strTerm As String
Dim lngCount As Long
    Err = 0: On Error GoTo ErrHand
    strItem = Me.vfgItems.TextMatrix(Me.vfgItems.Row, mCol.条件项)
    If Index = 0 Then
        strTerm = ""
        With Me.vfgVal
            For lngCount = .FixedRows To .Rows - 1
                strTerm = strTerm & vbTab & .TextMatrix(lngCount, 0)
            Next
        End With
        If strTerm <> "" Then strTerm = Mid(strTerm, 2)
        gstrSQL = "Zl_病历范文条件_Edit(" & mlngDemoId & ",'" & strItem & "','" & strTerm & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "范文条件"
        mblnOK = True
    Else
        If MsgBox("真的放弃当前条件的修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Me.cmdSave(0).Enabled = False: Me.cmdSave(1).Enabled = False
    Call RefList(strItem)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim rsTemp As New ADODB.Recordset
Dim strItem As String, aryVal() As String
Dim lngCount As Long
    
    Me.cmdEdit(0).Enabled = False: Me.cmdEdit(1).Enabled = False
    Me.cmdSave(0).Enabled = False: Me.cmdSave(1).Enabled = False
    Me.vfgVal.Clear: Me.vfgVal.Rows = 0
    Me.vfgSel.Clear: Me.vfgSel.Rows = 0
    If NewRow < 0 Then Exit Sub
    strItem = Me.vfgItems.TextMatrix(NewRow, mCol.条件项)
    aryVal = Split(Me.vfgItems.TextMatrix(NewRow, mCol.条件值), vbTab)
    With Me.vfgVal
        For lngCount = 0 To UBound(aryVal)
            .AddItem aryVal(lngCount)
        Next
        If .Rows > 0 Then .Row = 0
    End With
    
    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select 名称 As 可选值 From Table(Cast(f_Segment_可选值([1], [2]) As " & gstrDbOwner & ".t_Dic_Rowset))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDemoId, strItem)
    Set Me.vfgSel.DataSource = rsTemp
    Me.cmdEdit(0).Enabled = (Me.vfgSel.Rows > 0)
    Me.cmdEdit(1).Enabled = (Me.vfgVal.Rows > 0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgItems_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strItem As String
    If Me.cmdSave(0).Enabled = False Then Exit Sub
    strItem = Me.vfgItems.TextMatrix(OldRow, mCol.条件项)
    If MsgBox("已经更改了'" & strItem & "'的条件值，要放弃吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Exit Sub
    Cancel = True
End Sub

Private Sub vfgSel_DblClick()
    If Me.vfgSel.Rows < 1 Then Exit Sub
    If Me.vfgSel.Row < 0 Then Exit Sub
    Call cmdEdit_Click(0)
End Sub

Private Sub vfgVal_DblClick()
    If Me.vfgVal.Rows < 1 Then Exit Sub
    If Me.vfgVal.Row < 0 Then Exit Sub
    Call cmdEdit_Click(1)
End Sub
