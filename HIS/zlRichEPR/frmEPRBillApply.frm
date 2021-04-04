VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRBillApply 
   BorderStyle     =   0  'None
   Caption         =   "单据适用项目"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picEdit 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3285
      Left            =   0
      ScaleHeight     =   3285
      ScaleWidth      =   6285
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3810
      Width           =   6285
      Begin VB.Frame fraLine 
         Height          =   15
         Left            =   -45
         TabIndex        =   12
         Top             =   495
         Width           =   6690
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "按项目门诊是否可用选择(&P)"
         Height          =   350
         Index           =   2
         Left            =   3690
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   2475
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "门诊全清除(&R)"
         Height          =   350
         Index           =   1
         Left            =   1860
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   1815
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "门诊全选择(&L)"
         Height          =   350
         Index           =   0
         Left            =   105
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   45
         Width           =   1740
      End
      Begin VB.ComboBox cboKind 
         Height          =   300
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1650
         Width           =   1155
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   5010
         TabIndex        =   6
         Top             =   2235
         Width           =   1155
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找…   "
         Height          =   350
         Left            =   5010
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   2565
         Width           =   1155
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧ 添加到适用项目中(&A)"
         Height          =   350
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   585
         Width           =   2400
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨ 从适用项目中删除(&D)"
         Height          =   350
         Index           =   1
         Left            =   2520
         TabIndex        =   2
         Top             =   585
         Width           =   2400
      End
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2205
         Left            =   90
         TabIndex        =   3
         Top             =   975
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   3889
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
         Left            =   5010
         TabIndex        =   5
         Top             =   2040
         Width           =   810
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgItem 
      Height          =   3630
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   6045
      _cx             =   10663
      _cy             =   6403
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      Rows            =   7
      Cols            =   4
      FixedRows       =   1
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
End
Attribute VB_Name = "frmEPRBillApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ID = 0: 类别: 编码: 名称: 对象: 门诊: 住院: 体检
End Enum

Private mlngBillID As Long          '当前显示的单据id


'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngBillId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    
    mlngBillID = lngBillId
    
    '获取指定项目的信息
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select I.ID, K.名称 As 类别, I.编码, I.名称, I.服务对象 As 对象, A.门诊, A.住院, A.体检 " & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗项目类别 K," & vbNewLine & _
            "     (Select 诊疗项目id, Max(Decode(应用场合, 1, 1)) As 门诊, Max(Decode(应用场合, 2, 1)) As 住院" & vbNewLine & _
            "       , Max(Decode(应用场合, 4, 1)) As 体检" & _
            "       From 病历单据应用" & vbNewLine & _
            "       Where 病历文件id = [1]" & vbNewLine & _
            "       Group By 诊疗项目id) A" & vbNewLine & _
            "Where I.ID = A.诊疗项目id And I.类别 = K.编码" & vbNewLine & _
            "Order By I.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
    With Me.vfgItem
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        Call .AutoSize(mCol.类别, mCol.名称)
        .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        .ColWidth(mCol.对象) = 0: .ColHidden(mCol.对象) = True
        .ColWidth(mCol.门诊) = 450
        .ColWidth(mCol.住院) = 450
        .ColWidth(mCol.体检) = 450
        .ColWidth(mCol.名称) = .Width - .ColWidth(mCol.类别) - .ColWidth(mCol.编码) - .ColWidth(mCol.门诊) - .ColWidth(mCol.住院) - .ColWidth(mCol.体检)
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        For lngCount = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, lngCount, mCol.门诊) = IIf(.TextMatrix(lngCount, mCol.门诊) = "", flexUnchecked, flexChecked)
            .TextMatrix(lngCount, mCol.门诊) = ""
            .Cell(flexcpChecked, lngCount, mCol.住院) = IIf(.TextMatrix(lngCount, mCol.住院) = "", flexUnchecked, flexChecked)
            .TextMatrix(lngCount, mCol.住院) = ""
            .Cell(flexcpChecked, lngCount, mCol.体检) = IIf(.TextMatrix(lngCount, mCol.体检) = "", flexUnchecked, flexChecked)
            .TextMatrix(lngCount, mCol.体检) = ""
        Next
        .Redraw = flexRDDirect
    End With
    zlRefresh = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    Dim rsTemp As New ADODB.Recordset
    
    Me.lvwItem.ListItems.Clear
    Me.picEdit.Enabled = True: Call Form_Resize
    zlEditStart = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = ""
    Me.picEdit.Enabled = False: Call Form_Resize
    Call Me.zlRefresh(mlngBillID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
Dim strSQL() As String, strLists As String, blnTran As Boolean
Dim lngCount As Long
    
    '一般特性检查
    ReDim Preserve strSQL(0)
    strSQL(0) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",1,'')"
    ReDim Preserve strSQL(1)
    strSQL(1) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",2,'')"
    ReDim Preserve strSQL(2)
    strSQL(2) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",4,'')"
    With Me.vfgItem
        strLists = ""
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.门诊) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
                If Len(strLists) > 1900 Then
                    ReDim Preserve strSQL(UBound(strSQL) + 1)
                    strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",1,'" & Mid(strLists, 2) & "')"
                    strLists = ""
                End If
            End If
        Next
        If strLists <> "" Then
            ReDim Preserve strSQL(UBound(strSQL) + 1)
            strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",1,'" & Mid(strLists, 2) & "')"
        End If
        
        strLists = ""
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.住院) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
                If Len(strLists) > 1900 Then
                    ReDim Preserve strSQL(UBound(strSQL) + 1)
                    strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",2,'" & Mid(strLists, 2) & "')"
                    strLists = ""
                End If
            End If
        Next
        If strLists <> "" Then
            ReDim Preserve strSQL(UBound(strSQL) + 1)
            strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",2,'" & Mid(strLists, 2) & "')"
        End If
    
        strLists = ""
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.体检) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mCol.ID)
                If Len(strLists) > 1900 Then
                    ReDim Preserve strSQL(UBound(strSQL) + 1)
                    strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",4,'" & Mid(strLists, 2) & "')"
                    strLists = ""
                End If
            End If
        Next
        If strLists <> "" Then
            ReDim Preserve strSQL(UBound(strSQL) + 1)
            strSQL(UBound(strSQL)) = "Zl_诊疗单据目录_Apply(" & mlngBillID & ",4,'" & Mid(strLists, 2) & "')"
        End If
    End With
    
    
    '数据保存语句组织
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    blnTran = True
    For lngCount = 0 To UBound(strSQL)
        Call zlDatabase.ExecuteProcedure(strSQL(lngCount), "保存适用项目")
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    Call zlRefresh(mlngBillID)
    Me.picEdit.Enabled = False: Call Form_Resize
    zlEditSave = mlngBillID: Exit Function
    
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function


'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub cmdEdit_Click(Index As Integer)
Dim lngCol As Long
Dim objItem As ListItem
    With Me.vfgItem
        Select Case Index
        Case 0         '添加
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwItem.SelectedItem
            .Rows = .Rows + 1: .Row = .Rows - 1
            .TextMatrix(.Row, mCol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, mCol.编码) = objItem.Text
            .TextMatrix(.Row, mCol.名称) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1)
            .TextMatrix(.Row, mCol.类别) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_类别").Index - 1)
            .TextMatrix(.Row, mCol.对象) = objItem.Tag
            Select Case Val(objItem.Tag)
            Case 4
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexChecked
            Case 3
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexUnchecked
            Case 2
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexUnchecked
            Case 1
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexUnchecked
            Case Else
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexUnchecked
            End Select
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
            Me.lvwItem.ListItems.Remove objItem.Key
        Case 1          '删除
            If .Row < .FixedRows Then MsgBox "所有行已经删除！", vbInformation, gstrSysName: Exit Sub
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.编码))
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = .TextMatrix(.Row, mCol.名称)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_类别").Index - 1) = .TextMatrix(.Row, mCol.类别)
            objItem.Tag = .TextMatrix(.Row, mCol.对象)
            objItem.Selected = True
            objItem.EnsureVisible
            .RemoveItem .Row
        End Select
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
Dim rsTemp As New ADODB.Recordset
Dim strFind As String, strColdId As String
Dim objItem As ListItem
Dim lngCount As Long
    
    If Me.cboKind.ListIndex = 0 And Trim(Me.txtFind.Text) = "" Then
        MsgBox "在所有类别时，必须输入查找内容！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strFind = Trim(UCase(Me.txtFind.Text))
    gstrSQL = "Select Distinct I.ID, I.编码, I.名称, K.名称 As 类别, I.服务对象" & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗项目别名 N, 诊疗项目类别 K" & vbNewLine & _
            "Where I.ID = N.诊疗项目id And I.类别 = K.编码 And (I.撤档时间>sysdate or I.撤档时间 is null) And (0 = [3] Or K.编码 = [4]) And" & vbNewLine & _
            "      (I.编码 Like [1] || '%' Or N.名称 Like [2] || '%' Or N.简码 Like [2] || '%')"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFind, gstrMatch & strFind, Me.cboKind.ListIndex, Left(Me.cboKind.Text, 1))
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_名称").Index - 1) = "" & !名称
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_类别").Index - 1) = "" & !类别
            objItem.Tag = "" & !服务对象
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgItem
        For lngCount = .FixedRows To .Rows - 1
            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "没有匹配的项目！", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgItem.SetFocus
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click(Index As Integer)
Dim lngCount As Long
    With Me.vfgItem
        For lngCount = .FixedRows To .Rows - 1
            Select Case Index
            '全选
            Case 0
                '体检
                If mCol.体检 = .Col Then
                    .Cell(flexcpChecked, lngCount, .Col) = flexChecked
                    .Cell(flexcpChecked, lngCount, mCol.门诊) = flexUnchecked
                    .Cell(flexcpChecked, lngCount, mCol.住院) = flexUnchecked
                Else
                    .Cell(flexcpChecked, lngCount, .Col) = flexChecked
                    .Cell(flexcpChecked, lngCount, mCol.体检) = flexUnchecked
                End If
            '全清
            Case 1: .Cell(flexcpChecked, lngCount, .Col) = flexUnchecked
            '有条件选
            Case 2
                Select Case Val(.TextMatrix(lngCount, mCol.对象))
                Case 4  '体检
                    .Cell(flexcpChecked, lngCount, .Col) = flexChecked
                    .Cell(flexcpChecked, lngCount, mCol.门诊) = flexUnchecked
                    .Cell(flexcpChecked, lngCount, mCol.住院) = flexUnchecked
                Case 3:
                    .Cell(flexcpChecked, lngCount, mCol.门诊) = flexChecked
                    .Cell(flexcpChecked, lngCount, mCol.住院) = flexChecked
                    .Cell(flexcpChecked, lngCount, mCol.体检) = flexUnchecked
                Case 2
                    If .Col = mCol.住院 Then
                        .Cell(flexcpChecked, lngCount, .Col) = flexChecked
                    Else
                        .Cell(flexcpChecked, lngCount, .Col) = flexUnchecked
                    End If
                    .Cell(flexcpChecked, lngCount, mCol.体检) = flexUnchecked
                Case 1
                    If .Col = mCol.门诊 Then
                        .Cell(flexcpChecked, lngCount, .Col) = flexChecked
                    Else
                        .Cell(flexcpChecked, lngCount, .Col) = flexUnchecked
                    End If
                    .Cell(flexcpChecked, lngCount, mCol.体检) = flexUnchecked
                Case Else: .Cell(flexcpChecked, lngCount, .Col) = flexUnchecked
                End Select
            End Select
        Next
    End With
    Me.vfgItem.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mlngBillID = 0
    Me.picEdit.BackColor = Me.BackColor
    
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 1000
        .Add , "_名称", "名称", 2800
        .Add , "_类别", "类别", 900
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgItem.ZOrder 0

    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 编码, 名称 From 诊疗项目类别 Where 编码 Not In ('8', '9', 'M') Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.cboKind.Clear
        Me.cboKind.AddItem "所有项目"
        Do While Not .EOF
            Me.cboKind.AddItem !编码 & "-" & !名称
            .MoveNext
        Loop
        Me.cboKind.ListIndex = 0
    End With
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.picEdit.Enabled Then
        Me.picEdit.Visible = True
        Me.vfgItem.FocusRect = flexFocusHeavy
        Me.vfgItem.Height = Me.ScaleHeight - Me.picEdit.Height - Me.vfgItem.Top
    Else
        Me.picEdit.Visible = False
        Me.vfgItem.FocusRect = flexFocusNone
        Me.vfgItem.Height = Me.ScaleHeight - Me.vfgItem.Top - 120
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

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
End Sub

Private Sub vfgItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case NewCol
    Case mCol.门诊
        Me.cmdSel(0).Enabled = True: Me.cmdSel(0).Caption = "门诊全选择(&L)"
        Me.cmdSel(1).Enabled = True: Me.cmdSel(1).Caption = "门诊全清除(&R)"
        Me.cmdSel(2).Enabled = True: Me.cmdSel(2).Caption = "按项目门诊是否可用选择(&P)"
    Case mCol.住院
        Me.cmdSel(0).Enabled = True: Me.cmdSel(0).Caption = "住院全选择(&L)"
        Me.cmdSel(1).Enabled = True: Me.cmdSel(1).Caption = "住院全清除(&R)"
        Me.cmdSel(2).Enabled = True: Me.cmdSel(2).Caption = "按项目住院是否可用选择(&P)"
    Case mCol.体检
        Me.cmdSel(0).Enabled = True: Me.cmdSel(0).Caption = "体检全选择(&L)"
        Me.cmdSel(1).Enabled = True: Me.cmdSel(1).Caption = "体检全清除(&R)"
        Me.cmdSel(2).Enabled = True: Me.cmdSel(2).Caption = "按项目体检是否可用选择(&P)"
    Case Else
        Me.cmdSel(0).Enabled = False
        Me.cmdSel(1).Enabled = False
        Me.cmdSel(2).Enabled = False
    End Select
End Sub

Private Sub vfgItem_DblClick()
    If Me.picEdit.Enabled = False Then Exit Sub
    With Me.vfgItem
        If .MouseRow < .FixedRows Then Exit Sub
        Select Case .Col
        Case mCol.门诊, mCol.住院
            If .Cell(flexcpChecked, .Row, .Col) = flexChecked Then
                .Cell(flexcpChecked, .Row, .Col) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, .Col) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.体检) = flexUnchecked
            End If
        Case mCol.体检
            If .Cell(flexcpChecked, .Row, .Col) = flexChecked Then
                .Cell(flexcpChecked, .Row, .Col) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, .Col) = flexChecked
                .Cell(flexcpChecked, .Row, mCol.门诊) = flexUnchecked
                .Cell(flexcpChecked, .Row, mCol.住院) = flexUnchecked
            End If
        Case Else
            Call cmdEdit_Click(1)
        End Select
    End With
End Sub

