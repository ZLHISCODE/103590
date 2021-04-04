VERSION 5.00
Begin VB.Form frmPatholReportDelay 
   Caption         =   "报告延迟"
   ClientHeight    =   7095
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10875
   Icon            =   "frmPatholReportDelay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10875
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   10335
      TabIndex        =   5
      Top             =   6000
      Width           =   10335
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打 印(&P)"
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "数据保存(&S)"
         Height          =   400
         Left            =   9000
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "删除记录(&C)"
         Height          =   400
         Left            =   7680
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame framReportDelay 
      Caption         =   "报告延迟记录"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10335
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         _ExtentX        =   17171
         _ExtentY        =   4471
         GridRows        =   21
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholReportDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngCurAdviceId As Long
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngCurDepartmentId As Long

Private mrecStudyInf As TStudyStateInf

Dim WithEvents zlReport As zl9Report.clsReport
Attribute zlReport.VB_VarHelpID = -1


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigReportDelayFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId


    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)

    If Trim(mrecStudyInf.strPatholNumber) = "" Then
        Call ConfigReportDelayFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")
        
        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
        End If
        
        Exit Sub
    Else
        Call ConfigReportDelayFace(True)
    End If

    Call LoadReportDelayData(mrecStudyInf.lngPatholAdviceId)
    
    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub



Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowDelay As Boolean
    
    blnIsAllowDelay = CheckPopedom(mstrPrivs, "报告延迟")
    
    cmdCancel.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    cmdSave.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    
    cmdPrint.Enabled = blnIsAllowDelay
    
    ufgData.ReadOnly = blnIsReadOnly
End Sub


Private Sub ConfigReportDelayFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置特检界面
    cmdSave.Enabled = blnIsValid
    cmdCancel.Enabled = blnIsValid
    cmdPrint.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
    End If
End Sub


Private Sub AdjustFace()
'调整界面布局
    framReportDelay.Left = 120
    framReportDelay.Top = 120
    framReportDelay.Width = Me.Width - 360
    framReportDelay.Height = Me.Height - picControl.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framReportDelay.Width - 240
    ufgData.Height = framReportDelay.Height - 360
    
    
    picControl.Left = 120
    picControl.Top = Me.Height - picControl.Height - 620
    picControl.Width = Me.Width - 240
    
    
    cmdPrint.Left = 0
    cmdPrint.Top = 0
    
    cmdSave.Left = picControl.Width - cmdSave.Width - 120
    cmdSave.Top = 0
    
    cmdCancel.Left = cmdSave.Left - cmdCancel.Width - 120
    cmdCancel.Top = 0
End Sub



Private Sub InitReportDelayList()
'初始化报告延迟显示列表
    Dim strTemp As String
    
    

     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("报告延迟列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrReportDelayCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrReportDelayCols
    Else
        ufgData.ColNames = strTemp
    End If
    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrReportDelayConvertFormat
End Sub


Private Sub ufgData_OnColFormartChange()
 '关闭窗口时保存列表配置
    zlDatabase.SetPara "报告延迟列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadReportDelayData(ByVal lngPatholAdviceId As Long)
'载入报告延迟数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID,延迟原因,延迟天数,临时诊断,转达人,登记人,登记时间,当前状态 from 病理报告延迟 where 病理医嘱ID=[1] order by 登记时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub SaveReportDelayData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'blnIsSaveOnlyDel:是否保存仅删除的数据

'保存报告延迟数据
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    For i = 1 To ufgData.GridRows - 1
        Select Case ufgData.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                dtServicesTime = zlDatabase.Currentdate
                
                strSql = "select Zl_病理报告延迟_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                        mrecStudyInf.lngPatholAdviceId, _
                                                        ufgData.Text(i, gstrReportDelay_延迟原因), _
                                                        Val(ufgData.Text(i, gstrReportDelay_延迟天数)), _
                                                        ufgData.Text(i, gstrReportDelay_临时诊断), _
                                                        ufgData.Text(i, gstrReportDelay_转达人), _
                                                        UserInfo.姓名, _
                                                        CDate(dtServicesTime))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveReportDelayData", "未成功获取新增后的报告延迟ID,处理失败。")
                    Exit Sub
                End If
                
                
                ufgData.Text(i, gstrReportDelay_ID) = rsData!返回值
                ufgData.Text(i, gstrReportDelay_登记人) = UserInfo.姓名
                ufgData.Text(i, gstrReportDelay_当前状态) = "未打印"
                ufgData.Text(i, gstrReportDelay_登记时间) = dtServicesTime
                                                        
            Case TDataRowState.Del
                strSql = "Zl_病理报告延迟_删除(" & Val(ufgData.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                strSql = "Zl_病理报告延迟_更新(" & Val(ufgData.KeyValue(i)) & ",'" & _
                                                ufgData.Text(i, gstrReportDelay_延迟原因) & "'," & _
                                                Val(ufgData.Text(i, gstrReportDelay_延迟天数)) & ",'" & _
                                                ufgData.Text(i, gstrReportDelay_临时诊断) & "','" & _
                                                ufgData.Text(i, gstrReportDelay_转达人) & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

        End Select
        
        '更新行状态
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
End Sub


Private Sub cmdCancel_Click()
'删除套餐
On Error GoTo errHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行删除的报告延迟记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除该报告延迟数据吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgData.DelCurRow
    
    '保存删除的数据
    Call SaveReportDelayData(True)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub PrintReportDelay(ByVal lngReportDelayId As Long)
'打印报告延迟单
    Call zlReport.ReportOpen(gcnOracle, 100, "ZL1_Inside_1294_07", Me, "报告延迟ID=" & lngReportDelayId, Decode((Val(zlDatabase.GetPara("是否直接打印", glngSys, glngModul, 0)) = 1), 0, 0, 2))
End Sub


Private Sub cmdPrint_Click()
'报告延迟单打印
On Error GoTo errHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行打印的报告延迟记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要进行打印的报告延迟记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        If MsgBoxD(Me, "延迟报告尚未保存，不能进行打印，需要自动保存吗？", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        Else
            '保存报告延迟信息
            Call SaveReportDelayData
        End If
    End If
    
    Call PrintReportDelay(ufgData.KeyValue(ufgData.SelectionRow))
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub CmdRefresh_Click()
''将数据恢复到初始状态
'On Error GoTo errHandle
'    Call mvfgReportDelay.RestoreList
'
'    Call mvfgReportDelay.RefreshReadColColor
'
'    Call RefreshRecordInf
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到报告延迟列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '保存套餐信息
    Call SaveReportDelayData
    
    Call MsgBoxD(Me, "数据已保存成功。", vbOKOnly, Me.Caption)
'    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Initialize()
    Set zlReport = New zl9Report.clsReport
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitReportDelayList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set zlReport = Nothing
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
        
    
    '如果未录入标本名称，则显示淡红色
    iCol = ufgData.GetColIndex(gstrReportDelay_延迟原因)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrReportDelay_延迟原因) = "", ufgData.ErrCellColor, ufgData.BackColor)
           
    
    
    '如果未录入主取医师，则显示淡红色
    iCol = ufgData.GetColIndex(gstrReportDelay_延迟天数)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrReportDelay_延迟天数)) <= 0, ufgData.ErrCellColor, ufgData.BackColor)
           
End Sub



Private Sub ShowReasonWindow(ByVal Row As Long, ByVal Col As Long)
    Dim strReason As String
    
    Dim frmReason As New frmPatholReportDelay_Select
    On Error GoTo errFree
    Call frmReason.ShowReasonWindow(ufgData.Text(Row, gstrReportDelay_延迟原因), Me)
    
    strReason = ""
    
    If frmReason.IsOk Then
        With frmReason
            If .chkJF.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需缴费", "需缴费")
            If .chkTG.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需脱钙", "需脱钙")
            If .chkBQC.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需补取材", "需补取材")
            If .chkSQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需深切", "需深切")
            If .chkCQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需重切", "需重切")
            If .chkLQ.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需连切", "需连切")
            If .chkMYZH.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需免疫组化", "需免疫组化")
            If .chkFZBL.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需分子病理", "需分子病理")
            If .chkTSRS.value <> 0 Then strReason = strReason & IIf(strReason <> "", "、需特殊染色", "需特殊染色")
            
            If Trim(.txtOther.Text) <> "" Then strReason = strReason & IIf(strReason <> "", "、", "") & .txtOther.Text
        End With
        
        ufgData.Text(Row, gstrReportDelay_延迟原因) = strReason
    End If
errFree:
    Call Unload(frmReason)
    Set frmReason = Nothing
End Sub



Private Sub ufgData_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errHandle
    Call ShowReasonWindow(Row, Col)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub





Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle
    Call LoadReportDelayData(mrecStudyInf.lngPatholAdviceId)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = ufgData.GetColIndex(gstrReportDelay_延迟天数) Then
        If Val(ufgData.Text(Row, gstrReportDelay_延迟天数)) <= 0 Then ufgData.Text(Row, gstrReportDelay_延迟天数) = "1"
        Exit Sub
    End If
End Sub

Private Sub zlReport_AfterPrint(ByVal ReportNum As String)
'报告单已打印
On Error GoTo errHandle
    Dim strSql As String
    
    strSql = "Zl_病理报告延迟_打印(" & ufgData.KeyValue(ufgData.SelectionRow) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '修改界面列表的打印状态
    ufgData.Text(ufgData.SelectionRow, gstrReportDelay_当前状态) = "已打印"
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

