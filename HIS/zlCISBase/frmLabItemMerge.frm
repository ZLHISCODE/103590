VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemMerge 
   BorderStyle     =   0  'None
   Caption         =   "检验组合项目构成"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   3960
      ScaleHeight     =   3675
      ScaleWidth      =   3855
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkUpper 
         Caption         =   "区分大小写(&U)"
         Height          =   210
         Left            =   2055
         TabIndex        =   10
         Top             =   0
         Width           =   2040
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨"
         Height          =   350
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Top             =   2235
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   ">"
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   1260
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "<"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   900
         Width           =   390
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   3495
         Picture         =   "frmLabItemMerge.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   225
         Width           =   360
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   435
         TabIndex        =   2
         Top             =   225
         Width           =   3045
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧"
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   1860
         Width           =   390
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   3120
         Left            =   435
         TabIndex        =   4
         Top             =   555
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   5503
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
         Caption         =   "查找项目:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   0
         Width           =   810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3825
      _cx             =   6747
      _cy             =   6482
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
      Cols            =   5
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
End
Attribute VB_Name = "frmLabItemMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id
Private mbln组合 As Boolean         '当前项目是否组合项目
Private mstr类型 As String          '当前项目的检验类型

Private Enum mCol
    ID = 0: 序号: 编码: 中文名: 英文名
End Enum

Dim ObjItem As ListItem
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
            .Rows = 1: .FixedRows = 1: .Cols = 5: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.序号) = "序号": .TextMatrix(0, mCol.编码) = "编码"
        .TextMatrix(0, mCol.中文名) = "中文名": .TextMatrix(0, mCol.英文名) = "英文名"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.序号) = 500: .ColWidth(mCol.编码) = 1000
        .ColWidth(mCol.中文名) = 2000: .ColWidth(mCol.英文名) = 2000
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColAlignment(mCol.序号) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    '参数：当前项目id
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
        
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    mbln组合 = False: mstr类型 = ""
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 组合项目, 操作类型, 单独应用 From 诊疗项目目录 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        If .RecordCount > 0 Then
            mbln组合 = (Val("" & !组合项目) = 1)
            mstr类型 = "" & !操作类型
'            If Val("" & !单独应用) <> 1 Then
'                zlRefresh = False: Exit Function
'            End If
        Else
            zlRefresh = False: Exit Function
        End If
    End With
    
    gstrSql = "Select Distinct R.诊疗项目id As ID, R.排列序号 As 序号, V.编码, V.中文名, V.英文名" & vbNewLine & _
            " From 检验报告项目 R, 诊治所见项目 V, 检验合并规则 H , 诊疗项目目录 I " & vbNewLine & _
            " Where R.报告项目id = V.ID And R.诊疗项目id =H.合并项目ID And H.主项目ID=[1] And 细菌id Is Null" & vbNewLine & _
            " and r.诊疗项目ID  = i.id and i.组合项目 <> 1 and 单独应用 = 1"
    gstrSql = gstrSql & " Union ALL " & _
            " Select Distinct  R.诊疗项目id As ID, 0 As 序号, i.编码, i.名称, '' As 英文名" & vbNewLine & _
            " From 检验报告项目 R, 诊治所见项目 V, 检验合并规则 H, 诊疗项目目录 I" & vbNewLine & _
            " Where R.报告项目id = V.ID And R.诊疗项目id =H.合并项目ID And H.主项目ID=[1] And 细菌id Is Null" & vbNewLine & _
            " and r.诊疗项目ID  = i.id and i.组合项目 = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    '参数： lngItemId-指定编辑的项目
        
    Me.Tag = "编辑": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strLists As String
    
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '数据保存
    gstrSql = "Zl_检验合并规则_Edit(" & mlngItemID & ",'" & strLists & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long
    With Me.vfgList
        Select Case Index
        Case 0         '添加
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set ObjItem = Me.lvwItem.SelectedItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mCol.ID) = Mid(ObjItem.Key, 2)
            .TextMatrix(.Rows - 1, mCol.编码) = ObjItem.Text
            .TextMatrix(.Rows - 1, mCol.中文名) = ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1)
            .TextMatrix(.Rows - 1, mCol.英文名) = ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1)
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwItem.ListItems.Remove ObjItem.Key: Me.lvwItem.SetFocus
        Case 1          '删除
            If .Row < .FixedRows Then Exit Sub
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.编码))
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) = .TextMatrix(.Row, mCol.中文名)
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1) = .TextMatrix(.Row, mCol.英文名)
            ObjItem.Selected = True
            .RemoveItem .Row
        Case 2          '上移
            If .Row <= .FixedRows Then Exit Sub
            lngCurRow = .Row
            .AddItem "", lngCurRow - 1
            .TextMatrix(lngCurRow - 1, mCol.ID) = .TextMatrix(lngCurRow + 1, mCol.ID)
            .TextMatrix(lngCurRow - 1, mCol.编码) = .TextMatrix(lngCurRow + 1, mCol.编码)
            .TextMatrix(lngCurRow - 1, mCol.中文名) = .TextMatrix(lngCurRow + 1, mCol.中文名)
            .TextMatrix(lngCurRow - 1, mCol.英文名) = .TextMatrix(lngCurRow + 1, mCol.英文名)
            .RemoveItem lngCurRow + 1
            .Row = lngCurRow - 1
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
            
        Case 3          '下移
            If .Row >= .Rows - 1 Then Exit Sub
            lngCurRow = .Row
            .AddItem "", lngCurRow
            .TextMatrix(lngCurRow, mCol.ID) = .TextMatrix(lngCurRow + 2, mCol.ID)
            .TextMatrix(lngCurRow, mCol.编码) = .TextMatrix(lngCurRow + 2, mCol.编码)
            .TextMatrix(lngCurRow, mCol.中文名) = .TextMatrix(lngCurRow + 2, mCol.中文名)
            .TextMatrix(lngCurRow, mCol.英文名) = .TextMatrix(lngCurRow + 2, mCol.英文名)
            .RemoveItem lngCurRow + 2
            .Row = lngCurRow + 1
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        End Select
        
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String
    
    If Me.chkUpper.Value = 0 Then
        strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
        gstrSql = "Select distinct I.ID, I.编码, I.名称 As 中文名, L.缩写 As 英文名, Nvl(H.主项目ID,0) as 使用" & vbNewLine & _
                "From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L, 检验合并规则 H " & vbNewLine & _
                "Where I.ID=H.合并项目ID(+) And I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 <> 1 and i.单独应用 = 1 And I.操作类型 = '" & mstr类型 & "' And" & vbNewLine & _
                "      (I.编码 Like '" & strFind & "%' Or Upper(I.名称) Like '" & gstrMatch & strFind & "%' Or Upper(L.缩写) Like '" & gstrMatch & strFind & "%')"
        gstrSql = gstrSql & " Union ALL " & _
                " Select distinct I.ID, I.编码, I.名称 As 中文名, '' As 英文名, Nvl(H.主项目ID,0) as 使用" & vbNewLine & _
                " From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L, 检验合并规则 H " & vbNewLine & _
                " Where I.ID=H.合并项目ID(+) And I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 = 1 And I.操作类型 = '" & mstr类型 & "' And" & vbNewLine & _
                "      (I.编码 Like '" & strFind & "%' Or Upper(I.名称) Like '" & gstrMatch & strFind & "%' Or Upper(L.缩写) Like '" & gstrMatch & strFind & "%')"
    Else
        strFind = DelInvalidChar(Trim(Me.txtFind.Text))
        gstrSql = "Select distinct I.ID, I.编码, I.名称 As 中文名, L.缩写 As 英文名, Nvl(H.主项目ID,0) as 使用" & vbNewLine & _
                "From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L, 检验合并规则 H" & vbNewLine & _
                "Where I.ID=H.合并项目ID(+) And I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 <> 1 and i.单独应用 =1 And I.操作类型 = '" & mstr类型 & "' And" & vbNewLine & _
                "      (I.编码 Like '" & strFind & "%' Or I.名称 Like '" & gstrMatch & strFind & "%' Or L.缩写 Like '" & gstrMatch & strFind & "%')"
        gstrSql = gstrSql & " Union ALL " & _
                " Select distinct I.ID, I.编码, I.名称 As 中文名, '' As 英文名, Nvl(H.主项目ID,0) as 使用" & vbNewLine & _
                " From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L, 检验合并规则 H" & vbNewLine & _
                " Where I.ID=H.合并项目ID(+) And I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 = 1 And I.操作类型 = '" & mstr类型 & "' And" & vbNewLine & _
                "      (I.编码 Like '" & strFind & "%' Or I.名称 Like '" & gstrMatch & strFind & "%' Or L.缩写 Like '" & gstrMatch & strFind & "%')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        .Filter = " 使用 = 0 "
        Do While Not .EOF
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) = "" & !中文名
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1) = "" & !英文名
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    
    '不能选自身
    Me.lvwItem.ListItems.Remove "_" & mlngItemID
    
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
        .Add , "_中文名", "中文名", 2300
        .Add , "_英文名", "英文名", 600
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 180
    Me.picEdit.Height = Me.ScaleHeight - Me.picEdit.Top - 180
    If Me.Tag = "编辑" Then
        Me.vfgList.Width = Me.picEdit.Left - Me.vfgList.Left
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
    Else
        Me.vfgList.Width = Me.picEdit.Left + Me.picEdit.Width - Me.vfgList.Left
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
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
