VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub DeliveryInvoice(ByVal cnVal As ADODB.Connection)
'------------------------------
'功能: 调用送货发票导入操作界面
'------------------------------
    If cnVal Is Nothing Then
        MsgBox "创建ZLHIS的数据库连接失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "创建zl9ComLib部件失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal

    Call gobjComLib.InitCommon(cnVal)              '定义内部Connection
    
    frmDeliveryInvoice.Show vbModal
End Sub

Public Sub PurchasePlan(ByVal cnVal As ADODB.Connection)
'------------------------------
'功能: 调用采购计划导出操作界面
'------------------------------
    If cnVal Is Nothing Then
        MsgBox "创建ZLHIS的数据库连接失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "创建zl9ComLib部件失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal

    Call gobjComLib.InitCommon(cnVal)              '定义内部Connection
    
    frmPurchasePlan.Show vbModal
End Sub

Public Sub ClearDrugInfo(ByVal lngDrugID As Long, ByVal lngPurchaseID As Long)
'------------------------------
'功能: 删除药品信息
'参数:
'  lngDrugID: 药品ID
'  lngPurchaseID: 中标单位ID
'------------------------------
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strClear As String, strResultMESS As String
    Dim rsClear As New ADODB.Recordset
    
'    Set gcnOracle = cnVal
'    Call initCommon(cnVal)
    
    '连接参数
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    
    '默认MSSQL方式连接
    If MSSQLServerOpen(strServer, strDB, strUser, strPWD) = False Then
        If MsgBox("是否立即设置连接？", vbQuestion + vbYesNo + vbDefaultButton1, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        Else
            frmOutsideLinkSet.Show vbModal
            If gblnSetupFinish = False Then Exit Sub
        End If
    End If
    
    strClear = "declare @msg varchar(2000) " & vbNewLine _
             & "execute p_ins_waredict '000006'" _
             & "  ,'" & IIf(lngPurchaseID = 0, "", lngPurchaseID) & "'" _
             & "  ,''" _
             & "  ,'" & IIf(lngDrugID = 0, "", lngDrugID) & "'" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,''" _
             & "  ,0" _
             & "  ,''" _
             & "  ,''" _
             & "  ,1 ,@msg output " & vbNewLine _
             & "select @msg result_msg "
    On Error Resume Next
    rsClear.Open strClear, gcnOutside
    strResultMESS = rsClear!result_msg
    rsClear.Close
    On Error GoTo 0
    
    If InStr(strResultMESS, "成功") = 0 Or strResultMESS = "" Then
        If strResultMESS = "" Then
            MsgBox "同步物流平台药品信息失败！", vbInformation, GSTR_MESSAGE
        Else
            MsgBox "同步物流平台药品信息：" & vbNewLine & strResultMESS, vbInformation, GSTR_MESSAGE
        End If
    End If
    
End Sub

Public Sub UploadDrugInfo(ByVal frmOwner As Object, ByVal cnVal As ADODB.Connection, Optional ByVal lngDrugID As Long)
'------------------------------
'功能: 上传药品信息
'参数:
'  frmOwner:  宿主窗体对象
'  lngDrugID: 药品ID    为0时就上传全部药品信息
'------------------------------
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strSQL As String, strUpload As String, strResultMESS As String
    Dim rsTemp As ADODB.Recordset, rsUpload As New ADODB.Recordset
    Dim lngFailedQTY As Long
    
    If cnVal Is Nothing Then
        MsgBox "创建ZLHIS的数据库连接失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    If gobjComLib Is Nothing Then
        MsgBox "创建zl9ComLib部件失败，终止进入！", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    Set gcnOracle = cnVal
    Call gobjComLib.InitCommon(cnVal)
    
    '连接参数
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    
    '默认MSSQL方式连接
    If MSSQLServerOpen(strServer, strDB, strUser, strPWD) = False Then
        If MsgBox("是否立即设置连接？", vbQuestion + vbYesNo + vbDefaultButton1, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        Else
            frmOutsideLinkSet.Show vbModal
            If gblnSetupFinish = False Then Exit Sub
        End If
    End If
    
    strSQL = "Select a.Id, a.编码, a.规格, a.产地, b.药库单位, Round(b.成本价 * b.药库包装, 4) 药库单价, b.批准文号, c.名称 通用名, " _
           & "  c.拼音码, c.化学名, d.单位id 供应商代码, e.名称 供应商 " _
           & "From 收费项目目录 A, 药品规格 B," _
           & "  (Select 收费细目id, Min(Decode(性质, 1, 名称)) 名称, Min(Decode(码类, 1, 简码)) 拼音码, Min(Decode(码类, 0, 名称)) 化学名 " _
           & "   From 收费项目别名 " _
           & "   Where 码类 In (0, 1) And 性质 In (1, 2) " _
           & "   Group By 收费细目id) C, " _
           & "  (Select 药品id, 单位id From 药品中标单位 Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) D, " _
           & "  (Select ID, 名称 From 供应商 Where 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') And 末级 = 1 And Substr(类型, 1, 1) = '1') E " _
           & "Where a.Id = b.药品id And a.Id = c.收费细目id(+) And b.药品id = d.药品id(+) And d.单位id = e.Id(+) " _
           & "  And a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') And a.类别 In ('5', '6') "
    If lngDrugID > 0 Then
        strSQL = strSQL & " and a.ID=[1] "
    End If
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "上传药品信息至物流平台", lngDrugID)
    
    If Not rsTemp.EOF And lngDrugID = 0 Then
        With frmMessage
            .Show , frmOwner
            .Caption = "批量上传药品信息"
            .lblMess.Caption = ""
            .cmdClose.Enabled = False
            .pbrMess.Min = 0: .pbrMess.Max = rsTemp.RecordCount: .pbrMess.Value = 0
        End With
        rsTemp.MoveFirst
        '清空药品信息
        strUpload = "declare @msg varchar(2000) " & vbNewLine _
                  & "execute p_ins_waredict '000006'" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,0 " _
                  & "  ,''" _
                  & "  ,''" _
                  & "  ,-1 ,@msg output " & vbNewLine _
                  & "select @msg result_msg "
        On Error Resume Next
        rsUpload.Open strUpload, gcnOutside
        rsUpload.Close
        On Error GoTo 0
    End If
    
    Do While Not rsTemp.EOF
        strResultMESS = ""
        strUpload = "declare @msg varchar(2000) " & vbNewLine _
                  & "execute p_ins_waredict '000006'" _
                  & "  ,'" & rsTemp!供应商代码 & "'" _
                  & "  ,'" & rsTemp!供应商 & "'" _
                  & "  ,'" & rsTemp!Id & "'" _
                  & "  ,'" & rsTemp!化学名 & "'" _
                  & "  ,'" & rsTemp!通用名 & "'" _
                  & "  ,'" & rsTemp!拼音码 & "'" _
                  & "  ,'" & rsTemp!规格 & "'" _
                  & "  ,'" & rsTemp!药库单位 & "'" _
                  & "  ," & rsTemp!药库单价 _
                  & "  ,'" & rsTemp!产地 & "'" _
                  & "  ,'" & rsTemp!批准文号 & "'" _
                  & "  ,0 ,@msg output " & vbNewLine _
                  & "select @msg result_msg "
        On Error Resume Next
        rsUpload.Open strUpload, gcnOutside
        DoEvents
        strResultMESS = rsUpload!result_msg
        rsUpload.Close
        On Error GoTo 0
        '更新显示信息
        If InStr(strResultMESS, "成功") = 0 Or strResultMESS = "" Then
            If lngDrugID = 0 Then
                '批量上传
                lngFailedQTY = lngFailedQTY + 1
                With frmMessage
                    .txtContents.Text = "上传药品编码：" & rsTemp!编码 & "； 通用名：" & rsTemp!通用名 & " 。" & vbNewLine _
                                      & "    " & IIf(strResultMESS = "", "上传失败！", strResultMESS) & vbNewLine _
                                      & .txtContents.Text
                End With
            Else
                '单个上传
                MsgBox "上传失败！", vbInformation, GSTR_MESSAGE
            End If
            Err.Clear
        End If
        If lngDrugID = 0 Then
            With frmMessage
                .pbrMess.Value = rsTemp.AbsolutePosition
                .lblMess = "总共 " & .pbrMess.Max & " 条记录，未上传 " & lngFailedQTY + .pbrMess.Max - .pbrMess.Value & " 条记录。"
            End With
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If lngDrugID = 0 Then
        frmMessage.txtContents.Text = "批量上传完成！" & vbNewLine & vbNewLine & frmMessage.txtContents.Text
        frmMessage.cmdClose.Enabled = True
    End If
End Sub

Private Sub Class_Initialize()
    '创建zl9ComLib部件
    On Error Resume Next
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
End Sub

Private Sub Class_Terminate()
    Set gobjComLib = Nothing
End Sub
