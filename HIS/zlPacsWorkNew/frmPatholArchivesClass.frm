VERSION 5.00
Begin VB.Form frmPatholArchivesClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "档案分类设置"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   Icon            =   "frmPatholArchivesClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSave 
      Caption         =   "保 存(&S)"
      Height          =   400
      Left            =   10080
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删 除(&D)"
      Height          =   400
      Left            =   8760
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9975
      GridRows        =   501
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
   End
End
Attribute VB_Name = "frmPatholArchivesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function IsAllowDelArchivesClass(ByVal lngClassID As Long) As Boolean
'是否允许删除档案类别数据
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select 分类ID from 病理档案信息 where 分类ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngClassID)
    
    
    IsAllowDelArchivesClass = IIf(rsData.RecordCount > 0, False, True)
End Function



Private Sub cmdDel_Click()
'清除档案分类数据，已使用的类别不能进行删除
On Error GoTo ErrHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要删除的档案类别。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要删除的档案类别。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断是否允许删除该类别
    If Not IsAllowDelArchivesClass(ufgData.KeyValue(ufgData.SelectionRow)) Then
        Call MsgBoxD(Me, "该档案类别已被使用，不能进行删除。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "确认要删除选择的档案类别吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    '删除行
    Call ufgData.DelCurRow
    
    '保存删除的档案类别数据
    Call SaveArchivesClassData(True)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    Dim blnValid As Boolean
    
    '档案类别保存
    If ufgData.ShowingDataRowCount <= 0 Then
        Call MsgBoxD(Me, "没有找到需要保存的档案类别信息，请录入档案类别数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到档案类别列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        
        ufgData.SetFocus
        
        Exit Sub
    End If
    
    Call SaveArchivesClassData
    
    Call MsgBoxD(Me, "数据已成功保存。", vbOKOnly, Me.Caption)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub SaveArchivesClassData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'------------------------------------------------------------------------------
'blnIsSaveOnlyDel:是否仅仅保存删除的数据
'------------------------------------------------------------------------------
'档案类别保存


    Dim i As Long
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim dtSerivcesTime As Date
    
    Dim strNewId As String
    

    For i = 1 To ufgData.GridRows - 1
        If ufgData.RowState(i) = TDataRowState.Add And Not blnIsSaveOnlyDel Then
            
            dtSerivcesTime = zlDatabase.Currentdate
            strSQL = "select Zl_病理档案_新增分类([1],[2],[3],[4],[5],[6])  as 返回值 from dual"
            
            Set rsResult = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                    ufgData.Text(i, gstrArchivesClass_分类名称), _
                                                    Val(ufgData.Text(i, gstrArchivesClass_材料类型)), _
                                                    ufgData.Text(i, gstrArchivesClass_报表名称), _
                                                    UserInfo.姓名, _
                                                    CDate(dtSerivcesTime), _
                                                    ufgData.Text(i, gstrArchivesClass_备注) _
                                                    )
                                                            
            
            If rsResult.RecordCount <= 0 Then
                Call err.Raise(0, "SaveArchivesData", "未成功获取新增后的档案分类ID,处理失败。")
                Exit Sub
            End If
            
            '更新档案分类列表
            ufgData.Text(i, gstrArchivesClass_ID) = Nvl(rsResult!返回值)
            ufgData.Text(i, gstrArchivesClass_创建人) = UserInfo.姓名
            ufgData.Text(i, gstrArchivesClass_创建时间) = dtSerivcesTime
            
        ElseIf ufgData.RowState(i) = TDataRowState.Update And Not blnIsSaveOnlyDel Then
            
            strSQL = "Zl_病理档案_更新分类('" & ufgData.KeyValue(i) & "','" & _
                                                ufgData.Text(i, gstrArchivesClass_分类名称) & "'," & _
                                                Val(ufgData.Text(i, gstrArchivesClass_材料类型)) & ",'" & _
                                                ufgData.Text(i, gstrArchivesClass_报表名称) & "','" & _
                                                ufgData.Text(i, gstrArchivesClass_备注) & "')"
            
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
        ElseIf ufgData.RowState(i) = TDataRowState.Del Then
            '删除档案分类记录
            If Trim(ufgData.KeyValue(i)) <> "" Then
                strSQL = "Zl_病理档案_删除分类('" & ufgData.KeyValue(i) & "')"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
        
        
        '更新行状态
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
    
End Sub



Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitArchivesClassList
    
    Call LoadArchivesClassData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitArchivesClassList()
'配置档案分类显示列表

    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
    ufgData.DefaultColNames = gstrArchivesClassCols
    ufgData.ColNames = gstrArchivesClassCols
    ufgData.ColConvertFormat = gstrArchivesClassConvertFormat
End Sub



Private Sub LoadArchivesClassData()
'载入档案分类数据
    Dim strSQL As String
    
    strSQL = "select ID, 分类名称,材料类型,报表名称, 备注,创建人,创建时间 from 病理档案分类 order by 创建时间"
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim strNewArchivesClassName
    Dim iCol As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        Exit Sub
    End If
    
    
    If Col = ufgData.GetColIndex(gstrArchivesClass_分类名称) Then
    '检查档案类别是否重复
    
        strNewArchivesClassName = ufgData.CheckEquateValue(Row, Col)
        If strNewArchivesClassName <> "" Then
            Call MsgBoxD(Me, "档案类别 [" & ufgData.Text(Row, gstrArchivesClass_分类名称) & "]已经存在。", vbOKOnly, Me.Caption)
            
            ufgData.Text(Row, gstrArchivesClass_分类名称) = strNewArchivesClassName
        End If
    End If
    
    
    '如果未录入类别名称，则显示淡红色
    iCol = ufgData.GetColIndex(gstrArchivesClass_分类名称)
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrArchivesClass_分类名称) = "", ufgData.ErrCellColor, ufgData.BackColor)
          
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
