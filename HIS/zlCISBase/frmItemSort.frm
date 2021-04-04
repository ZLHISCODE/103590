VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检验项目排序"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10965
   Icon            =   "frmItemSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   5475
      ScaleHeight     =   5895
      ScaleWidth      =   5460
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   5460
      Begin VB.CommandButton cmdEdit 
         Height          =   350
         Index           =   5
         Left            =   0
         Picture         =   "frmItemSort.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   350
         Index           =   4
         Left            =   0
         Picture         =   "frmItemSort.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "保存"
         Top             =   2820
         Width           =   390
      End
      Begin VB.ComboBox cbo类型 
         Height          =   300
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
         Width           =   3855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧"
         Height          =   350
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   1860
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "<"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   900
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   ">"
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   1260
         Width           =   390
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨"
         Height          =   350
         Index           =   3
         Left            =   0
         TabIndex        =   1
         Top             =   2235
         Width           =   390
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   4800
         Left            =   435
         TabIndex        =   5
         Top             =   390
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   8467
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
      Begin VB.Label lbl操作类型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检验类型(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   465
         TabIndex        =   10
         Top             =   75
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   5490
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   5300
      _cx             =   9349
      _cy             =   9684
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
Attribute VB_Name = "frmItemSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnEdit As Boolean '是否修改过
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
        .ColWidth(mCol.中文名) = 2600: .ColWidth(mCol.英文名) = 1000
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

Public Function zlRefresh() As Boolean
    '功能：根据项目类别刷新当前显示内容
    '
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    Dim str类型 As String
    Me.lvwItem.ListItems.Clear
    
    On Error GoTo ErrHand
    If cbo类型.ListIndex < 0 Then Call setListFormat: zlRefresh = True: Exit Function

    '获取指定项目的信息
    str类型 = cbo类型.List(cbo类型.ListIndex)
    strSQL = "Select /*+ Rule */" & vbNewLine & _
            " a.诊治项目id, Decode(a.排列序号,Null,Null,a.排列序号-1000) as 序号, c.编码, c.名称 as 中文名, a.缩写 as 英文名" & vbNewLine & _
            "From 检验项目 a, 检验报告项目 b, 诊疗项目目录 c" & vbNewLine & _
            "Where a.诊治项目id = b.报告项目id And b.诊疗项目id = c.Id And A.排列序号 is not null And c.组合项目 = 0 And c.操作类型 = [1] And" & vbNewLine & _
            "           (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By a.排列序号, c.编码"
    str类型 = Mid(str类型, InStr(str类型, " ") + 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类型)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows

    strSQL = "Select /*+ Rule */" & vbNewLine & _
            " a.诊治项目id, Decode(a.排列序号,Null,Null,a.排列序号-1000) as 序号, c.编码, c.名称 as 中文名, a.缩写 as 英文名" & vbNewLine & _
            "From 检验项目 a, 检验报告项目 b, 诊疗项目目录 c" & vbNewLine & _
            "Where a.诊治项目id = b.报告项目id And b.诊疗项目id = c.Id And A.排列序号 is Null And c.组合项目 = 0 And c.操作类型 = [1] And" & vbNewLine & _
            "           (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By a.排列序号, c.编码"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类型)

    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & !诊治项目id, !编码)
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
    Me.vfgList.SetFocus

    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strLists As String
    Dim strSQL As String
    
    If Not mblnEdit Then Exit Function
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '数据保存
    strSQL = "ZL_检验项目_SORT('" & strLists & "')"

    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    mblnEdit = False: Call Form_Resize
    zlEditSave = 1: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub cbo类型_Click()
    Call zlRefresh
End Sub

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
            mblnEdit = True
        Case 1          '删除
            If .Row < .FixedRows Then Exit Sub
            Set ObjItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.编码))
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) = .TextMatrix(.Row, mCol.中文名)
            ObjItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1) = .TextMatrix(.Row, mCol.英文名)
            ObjItem.Selected = True
            .RemoveItem .Row
            mblnEdit = True
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
            mblnEdit = True
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
            mblnEdit = True
        Case 4          '保存
            Call zlEditSave
        Case 5          '退出
            If mblnEdit Then
                If MsgBox("您刚才所做的修改还未保存，是否退出？", vbInformation + vbYesNo, Me.Caption) = vbYes Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
        End Select
        cmdEdit(4).Enabled = mblnEdit
        
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
        If Index <> 5 Then .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 1000
        .Add , "_中文名", "中文名", 2300
        .Add , "_英文名", "英文名", 900
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With

    cbo类型.Clear
    strSQL = "Select 编码,名称 From 诊疗检验类型 Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do Until rsTmp.EOF
        cbo类型.AddItem rsTmp!编码 & " " & rsTmp!名称
        rsTmp.MoveNext
    Loop
    If cbo类型.ListCount > 0 Then cbo类型.ListIndex = 0
    Me.vfgList.ZOrder 0
    Me.Tag = "编辑"
    cmdEdit(4).Enabled = mblnEdit
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
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

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "编辑" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub


