VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemCost 
   BorderStyle     =   0  'None
   Caption         =   "项目耗用试剂"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7710
      _cx             =   13600
      _cy             =   3334
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   7710
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2295
      Width           =   7710
      Begin VB.CheckBox chkHand 
         Caption         =   "加入手工列(&M)"
         Height          =   195
         Left            =   6075
         TabIndex        =   11
         Top             =   135
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.OptionButton opt内容 
         Caption         =   "选择仪器(&1)"
         Height          =   180
         Index           =   1
         Left            =   6090
         TabIndex        =   10
         Top             =   1995
         Width           =   1560
      End
      Begin VB.OptionButton opt内容 
         Caption         =   "选择试剂(&0)"
         Height          =   180
         Index           =   0
         Left            =   6090
         TabIndex        =   9
         Top             =   1680
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨ 从试剂耗用列表中删除"
         Height          =   350
         Index           =   1
         Left            =   2595
         TabIndex        =   6
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧ 添加到试剂耗用列表中"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找…    "
         Height          =   350
         Left            =   6075
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   6075
         TabIndex        =   3
         Top             =   720
         Width           =   1605
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2280
         Left            =   0
         TabIndex        =   2
         Top             =   450
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查找内容:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6075
         TabIndex        =   7
         Top             =   495
         Width           =   810
      End
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 本表说明耗用单位用量的以下各种试剂可完成本项目检验多少人次。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   345
      TabIndex        =   8
      Top             =   120
      Width           =   5490
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   120
      Picture         =   "frmLabItemCost.frx":0000
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "frmLabItemCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的检验项目的诊疗项目id
Private mlngLabID As Long          '当前显示的检验项目的诊治项目id

Private Enum mCol
    ID = 0: 编码: 名称: 单位: 手工
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = mCol.手工 + 1
            .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.编码) = "编码"
            .TextMatrix(0, mCol.名称) = "名称": .TextMatrix(0, mCol.单位) = "单位"
            .TextMatrix(0, mCol.手工) = "手工"
        End If
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.编码) = 1000
        .ColWidth(mCol.名称) = 2500: .ColWidth(mCol.单位) = 500
        If Me.chkHand.Value = 0 Then
            .ColWidth(mCol.手工) = 0
        Else
            .ColWidth(mCol.手工) = 500
        End If
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .AutoSize mCol.手工, .Cols - 1
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '功能：根据仪器id刷新当前显示内容
    '参数：当前项目id
    Dim rsTemp As New ADODB.Recordset
    Dim rsApt As New ADODB.Recordset, strColSql As String
    
    mlngItemID = lngItemID
    mlngLabID = 0
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
    
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    On Error GoTo ErrHand
    gstrSql = "Select R.报告项目id From 检验报告项目 R Where R.诊疗项目id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    If rsTemp.RecordCount > 0 Then
        mlngLabID = Nvl(rsTemp!报告项目ID, 0)
'    Else
'        MsgBox "该检验项目信息部分丢失！", vbInformation, gstrSysName
    End If
    
    gstrSql = "Select Distinct A.ID, A.编码, A.名称" & vbNewLine & _
            "From 检验试剂关系 L, 检验仪器 A" & vbNewLine & _
            "Where L.仪器id = A.ID And L.项目id = [1]" & vbNewLine & _
            "Order By A.编码"
    Set rsApt = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngLabID)
    With rsApt
        strColSql = ""
        Do While Not .EOF
            strColSql = strColSql & ",Max(Decode(L.仪器id," & !ID & ",L.数量)) As C" & !ID
            .MoveNext
        Loop
    End With
    
    gstrSql = "Select L.材料id As ID, I.编码, I.名称, I.计算单位 As 单位, Max(Decode(L.仪器id, Null, L.数量)) As 手工" & strColSql & vbNewLine & _
            "From 诊疗项目目录 I, 材料特性 T, 检验试剂关系 L" & vbNewLine & _
            "Where L.材料id = T.材料id And T.诊疗id = I.ID And L.项目id = [1]" & vbNewLine & _
            "Group By L.材料id, I.编码, I.名称, I.计算单位"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngLabID)
    
    Me.vfgList.FixedCols = 0
    Set Me.vfgList.DataSource = rsTemp
    With rsApt
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Me.vfgList.ColData(.AbsolutePosition + mCol.手工) = CLng(!ID)
            Me.vfgList.Cell(flexcpData, 0, .AbsolutePosition + mCol.手工) = CStr(!编码)
            Me.vfgList.TextMatrix(0, .AbsolutePosition + mCol.手工) = "" & !名称
            .MoveNext
        Loop
    End With
    Me.vfgList.FixedCols = mCol.手工
    
    Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    '参数： lngItemId-指定编辑的项目
    Me.Tag = "编辑": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strLists As String, lngRow As Long, lngCol As Long
    
    strLists = ""
    With Me.vfgList
        For lngRow = .FixedRows To .Rows - 1
            If Me.chkHand.Value = vbChecked Then
                If Val(.TextMatrix(lngRow, mCol.手工)) > 99999999 Then
                    MsgBox "“" & .TextMatrix(lngRow, mCol.名称) & "”(" & lngRow & "行)，“手工”列人次数量太大！", vbInformation, gstrSysName
                    .Row = lngRow: .Col = mCol.手工: .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Val(.TextMatrix(lngRow, mCol.手工)) <> 0 Then
                    strLists = strLists & "|" & .TextMatrix(lngRow, mCol.ID) & ";;" & Val(.TextMatrix(lngRow, mCol.手工))
                End If
            End If
            For lngCol = mCol.手工 + 1 To .Cols - 1
                If Val(.TextMatrix(lngRow, lngCol)) > 99999999 Then
                    MsgBox "“" & .TextMatrix(lngRow, mCol.名称) & "”(" & lngRow & "行)，“" & .TextMatrix(0, lngCol) & "”(" & lngCol & "列)人次数量太大！", vbInformation, gstrSysName
                    .Row = lngRow: .Col = lngCol: .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then
                    strLists = strLists & "|" & .TextMatrix(lngRow, mCol.ID) & ";" & .ColData(lngCol) & ";" & Val(.TextMatrix(lngRow, lngCol))
                End If
            Next
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '数据保存
    gstrSql = "Zl_检验试剂关系_Edit(" & mlngLabID & ",'" & strLists & "')"
    If LenB(gstrSql) > 4000 Then
        MsgBox "设置了太多的试剂消耗品，不能保存！", vbInformation, gstrSysName
        Me.vfgList.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0
End Function

Private Sub chkKind_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkUpper_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkHand_Click()
    With Me.vfgList
        If Me.chkHand.Value = 0 Then
            .ColWidth(mCol.手工) = 0: .ColHidden(mCol.手工) = True
        Else
            .ColWidth(mCol.手工) = 500: .ColHidden(mCol.手工) = False
        End If
    End With
End Sub

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCol As Long
    With Me.vfgList
        If Me.opt内容(0).Value Then
            Select Case Index
            Case 0         '添加
                If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
                Set objItem = Me.lvwItem.SelectedItem
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, mCol.ID) = Mid(objItem.Key, 2)
                .TextMatrix(.Rows - 1, mCol.编码) = objItem.Text
                .TextMatrix(.Rows - 1, mCol.名称) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1)
                .TextMatrix(.Rows - 1, mCol.单位) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_单位").Index - 1)
                If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
                Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
            Case 1          '删除
                If .Row < .FixedRows Then MsgBox "所有试剂行已经删除！", vbInformation, gstrSysName: Exit Sub
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.编码))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = .TextMatrix(.Row, mCol.名称)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_单位").Index - 1) = .TextMatrix(.Row, mCol.单位)
                objItem.Selected = True
                .RemoveItem .Row
            End Select
        Else
            Select Case Index
            Case 0         '添加
                If .Cols >= mCol.手工 + 6 Then MsgBox "最多只能设置6种仪器！", vbInformation, gstrSysName: Exit Sub
                If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
                Set objItem = Me.lvwItem.SelectedItem
                .Cols = .Cols + 1
                .ColData(.Cols - 1) = Val(Mid(objItem.Key, 2))
                .Cell(flexcpData, 0, .Cols - 1) = objItem.Text
                .TextMatrix(0, .Cols - 1) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1)
                .Col = .Cols - 1
                If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
                Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
            Case 1          '删除
                If .Col <= mCol.手工 Then MsgBox "请选中想删除的仪器列！", vbInformation, gstrSysName: Exit Sub
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .ColData(.Col), .Cell(flexcpData, 0, .Col))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = .TextMatrix(0, .Col)
                objItem.Selected = True
                If .Col < .Cols - 1 Then
                    For lngCol = .Col To .Cols - 2
                        .ColData(lngCol) = .ColData(lngCol + 1)
                        .Cell(flexcpData, 0, lngCol) = .Cell(flexcpData, 0, lngCol + 1)
                        .TextMatrix(0, lngCol) = .TextMatrix(0, lngCol + 1)
                    Next
                End If
                .Cols = .Cols - 1
            End Select
            .AutoSize mCol.手工, .Cols - 1
        End If
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String, strColdId As String
    
    strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
    If Me.opt内容(0).Value Then
        gstrSql = "Select Distinct M.材料id As ID, I.编码, I.名称, I.计算单位 As 单位" & vbNewLine & _
                "From 诊疗项目目录 I, 诊疗项目别名 N, 材料特性 M" & vbNewLine & _
                "Where I.ID = N.诊疗项目id And I.ID = M.诊疗id And" & vbNewLine & _
                "      (I.编码 Like '" & strFind & "%' Or N.名称 Like '" & gstrMatch & strFind & "%' Or N.简码 Like '" & gstrMatch & strFind & "%')"
    Else
        gstrSql = "Select I.ID, I.编码, I.名称, '' As 单位" & vbNewLine & _
                "From 检验仪器 I" & vbNewLine & _
                "Where I.编码 Like '" & strFind & "%' Or I.名称 Like '" & gstrMatch & strFind & "%' Or I.简码 Like '" & gstrMatch & strFind & "%'"
    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = "" & !名称
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_单位").Index - 1) = "" & !单位
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgList
        If Me.opt内容(0).Value Then
            For lngCount = .FixedRows To .Rows - 1
                Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mCol.ID)
            Next
        Else
            For lngCount = mCol.手工 + 1 To .Cols - 1
                strColdId = .ColData(lngCount)
                Me.lvwItem.ListItems.Remove "_" & strColdId
            Next
        End If
    End With
    
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "没有匹配的项目！", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgList.SetFocus
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 1000
        .Add , "_名称", "名称", 3500
        .Add , "_单位", "单位", 1000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 105
    If Me.Tag = "编辑" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 105
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub opt内容_Click(Index As Integer)
    Me.lvwItem.ListItems.Clear
    Me.txtFind.Text = "": Me.txtFind.SetFocus
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.lvwItem.Height = Me.picEdit.ScaleHeight - Me.lvwItem.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "编辑" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col < mCol.手工 Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < mCol.手工 Then Cancel = True: Exit Sub
    If Row < Me.vfgList.FixedRows Then Cancel = True
End Sub


