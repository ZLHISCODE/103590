VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_DIH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口（蝶和）类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjSOAP As Object

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mobjSOAP = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'说明：开启接口为活动状态

    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mobjSOAP Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "连接蝶和服务器失败", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "开始连接蝶和服务器", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        Call mobjSOAP.MSSoapInit(mobjBase.URL)
        mobjBase.mobjLog.Add "成功连接蝶和服务器", 1
    Else
        Set mobjSOAP = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Add "连接蝶和服务器异常", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的Filter方法
    
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim strSQL As String, strInfo As String
    Dim objDB As Object
    
    On Error GoTo hErr
    
    strInfo = Split(strData, "|")(1)

    Select Case intType
    Case Val("22-开始发药"), Val("23-完成发药")
        If TypeName(mobjBase.mobjComLib) = "clsComLib" Then
            'ZLHIS、自助机、药品自动化设备管理工具
            Set objDB = mobjBase.mobjComLib.zlDatabase
        ElseIf TypeName(mobjBase.mobjComLib) = "clsPublic" Then
            '支付宝
            Set objDB = mobjBase.mobjComLib
        Else
            mobjBase.mobjLog.Add "对象类型不正确，终止获取标准数据", 1
            mobjBase.mobjLog.Add "" & intType, 1, 1
            mobjBase.mobjLog.Save
            Exit Function
        End If
        
        If Val(strData) = 1 Then
            strSQL = _
                "Select count(1) REC " & vbNewLine & _
                "From All_Tab_Cols " & vbNewLine & _
                "Where Table_Name = '未发药品记录' And (Lower(Column_Name) = 'groupno' Or Lower(Column_Name) = 'ordertype')"
            mobjBase.mobjLog.Add strSQL, 1, 1
            
            Set rsResult = objDB.OpenSQLRecord(strSQL, "获取DIH回写的GroupNO字段是否存在")

            If rsResult!Rec >= 2 Then
                strSQL = _
                    "Select Distinct a.单据, a.No 处方号, a.库房id, b.病人id, g.编码 窗口编码, f.GroupNO, f.OrderType," & _
                    "    Zl_Fun_Drug_Machine(a.库房id, d.药品剂型) 接口编号 " & vbNewLine & _
                    "From 药品收发记录 A, 门诊费用记录 B, 药品规格 C, 药品特性 D, 药品收发门诊标志 E, 未发药品记录 F, 发药窗口 G " & vbNewLine & _
                    "   , Table(f_Str2list2([1], ';', ',')) I " & vbNewLine & _
                    "Where a.费用id = b.Id And a.药品id = c.药品id And c.药名id = d.药名id(+) And a.单据 = i.C1 " & vbNewLine & _
                    "    And a.NO = e.处方号(+) And a.单据 = e.单据(+) And a.库房id = e.库房id(+) " & vbNewLine & _
                    "    And a.NO = f.NO And a.单据 = f.单据 And a.库房id = f.库房id And f.发药窗口 = g.名称(+) " & vbNewLine & _
                    "    And e.业务分类(+) = [2] And a.No = i.C2 And (Mod(a.记录状态, 3) = 0 Or a.记录状态 = 1) " & vbNewLine & _
                    "Order By 接口编号, a.NO, a.单据, a.库房id "
            Else
                mobjBase.mobjLog.Add "请为“未发药品记录”增加“GROUPNO和ORDERTYPE”两个字段", 1, 1
                strSQL = _
                    "Select Distinct a.单据, a.No 处方号, a.库房id, b.病人id, g.编码 窗口编码, '' GroupNO, '' OrderType, " & _
                    "    Zl_Fun_Drug_Machine(a.库房id, d.药品剂型) 接口编号 " & vbNewLine & _
                    "From 药品收发记录 A, 门诊费用记录 B, 药品规格 C, 药品特性 D, 药品收发门诊标志 E, 未发药品记录 F, 发药窗口 G " & vbNewLine & _
                    "   , Table(f_Str2list2([1], ';', ',')) I " & vbNewLine & _
                    "Where a.费用id = b.Id And a.药品id = c.药品id And c.药名id = d.药名id(+) And a.单据 = i.C1 " & vbNewLine & _
                    "    And a.NO = e.处方号(+) And a.单据 = e.单据(+) And a.库房id = e.库房id(+) " & vbNewLine & _
                    "    And a.NO = f.NO And a.单据 = f.单据 And a.库房id = f.库房id And f.发药窗口 = g.名称(+) " & vbNewLine & _
                    "    And e.业务分类(+) = [2] And a.No = i.C2 And (Mod(a.记录状态, 3) = 0 Or a.记录状态 = 1) " & vbNewLine & _
                    "Order By 接口编号, a.NO, a.单据, a.库房id "
            End If
            rsResult.Close

            mobjBase.mobjLog.Add strSQL, 1, 1
            Set rsResult = objDB.OpenSQLRecord(strSQL, "获取开始发药数据", strInfo, intType - 20)
            mobjBase.mobjLog.Add "获取开始发药数据完成", 1
        Else
            Set rsResult = Nothing
            mobjBase.mobjLog.Add "住院不支持", 1
        End If
        
        Set clsINF_GetData = rsResult
    
    Case Else
        '使用基类的数据
        Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    End Select
    
    Exit Function
    
hErr:
    blnFinish = False
    mobjBase.mobjLog.Add "获取标准数据失败", 1
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的SendData方法

'    Const STR_TAKE_END As String = "</outpOrderTake>"
'    Const STR_COMP_END As String = "</outpOrderCompletion>"
    
    Dim strInput As String
    Dim strSingle As String, strKey As String, strTmp As String
    Dim strReturnCode As String, strWindows As String
    Dim blnFinish As Boolean
    Dim objDB As Object
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "开始传送数据", 1
    
    Select Case intType
    Case 5, 21 To 23
        rsData.Filter = ""
    Case Else
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
makExit:
        mobjBase.mobjLog.Add "无数据可传送", 1
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    '标准数据传送开始
    mobjBase.mobjLog.Add "标准数据传送开始", 1
    
    '传送数据
    If intType = Val("22-开始发药") Or intType = Val("23-完成发药") Then
        '不同于其它功能的特殊处理
        
        rsData.MoveFirst
        Do While rsData.EOF = False
            blnFinish = True
        
            '格式化数据（单个记录）
            mobjBase.mobjLog.Add "格式化数据", 1
            strInput = FormatData(intType, intIO, strCode, rsData, Val("1-单记录") = 1)
            If strInput = "" Then
                strMessage = "格式化数据失败"
                'Exit Function
            End If
            
            If SendDataSingle(intType, strInput, strReturnCode, strWindows, strMessage) Then
                '标准数据传送完成
                mobjBase.mobjLog.Add "标准数据传送完成", 1
            Else
                '标准数据传送失败
                blnFinish = False
                strMessage = strWindows
                mobjBase.mobjLog.Add strMessage, 1
                mobjBase.mobjLog.Add "蝶和接口异常", 1
                
                '准备记录集对象
                With rsTemp
                    mobjBase.mobjLog.Add "准备记录集对象", 2
                    If .Fields.Count <= 0 Then
                        .LockType = adLockOptimistic
                        .CursorType = adOpenStatic
                        .CursorLocation = adUseClient
                        
                        .Fields.Append "库房ID", adBigInt, , adFldIsNullable
                        .Fields.Append "单据", adInteger, , adFldIsNullable
                        .Fields.Append "处方号", adVarChar, 10, adFldIsNullable
                        .Open
                    End If
                    
                    mobjBase.mobjLog.Add "准备失败标记的数据开始", 2
                    mobjBase.mobjLog.Add "库房ID:" & mdlDrugMachine.NVL(rsData!库房id), 3, 1
                    mobjBase.mobjLog.Add "单据:" & mdlDrugMachine.NVL(rsData!单据), 3, 1
                    mobjBase.mobjLog.Add "处方号:" & mdlDrugMachine.NVL(rsData!处方号), 3, 1
                    
                    .AddNew
                    .Fields!库房id = mdlDrugMachine.NVL(rsData!库房id, 0)
                    .Fields!单据 = mdlDrugMachine.NVL(rsData!单据)
                    .Fields!处方号 = mdlDrugMachine.NVL(rsData!处方号)
                    .Update
                    
                    mobjBase.mobjLog.Add "准备失败标记的数据完成", 2
                End With
                                
            End If
            
            '标记成功（药品收发记录），之所以失败也标记，是为了让接口管理工具的定时传送可补充传送。
            If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsTemp, mobjBase, blnFinish) Then
                mobjBase.mobjLog.Add "“传送”标记成功", 1
            Else
                mobjBase.mobjLog.Add "“传送”标记失败", 1
            End If
            If rsTemp.State = adStateOpen Then rsTemp.Close
            
            rsData.MoveNext
        Loop
        
    Else
    
        '格式化数据
        mobjBase.mobjLog.Add "格式化数据", 1
        strInput = FormatData(intType, intIO, strCode, rsData)
        If strInput = "" Then
            strMessage = "格式化数据失败"
            Exit Function
        End If
   
        If SendDataSingle(intType, strInput, strReturnCode, strWindows, strMessage) Then
            If intType = 21 Then
                '回写发药窗口
                If intAppType = Val("3-支付宝") Then
                    Set objDB = mobjBase.mobjComLib
                Else
                    Set objDB = mobjBase.mobjComLib.zlDatabase
                End If
                If mdlDrugMachine.UpdateDispenseWindow(rsData, strWindows, objDB, mobjBase.mobjLog) Then
                    mobjBase.mobjLog.Add "回写发药窗口成功", 1
                Else
                    mobjBase.mobjLog.Add "回写发药窗口失败", 1
                End If
            End If
        Else
            strMessage = strWindows
            mobjBase.mobjLog.Add strMessage, 1
            mobjBase.mobjLog.Add "蝶和接口异常", 1
        End If
        
        '标记成功（药品收发记录），之所以失败也标记，是为了让接口管理工具的定时传送可补充传送。
        If intType > 20 Then
            If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
                mobjBase.mobjLog.Add "“传送”标记成功", 1
            Else
                mobjBase.mobjLog.Add "“传送”标记失败", 1
            End If
        End If
        
        '标准数据传送完成
        mobjBase.mobjLog.Add "标准数据传送完成", 1
        mobjBase.mobjLog.Save
    End If
    
    clsINF_SendData = True
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "标准数据传送失败", 1
    mobjBase.mobjLog.Save
    Exit Function
End Function

Private Function SendDataSingle(ByVal intType As Integer, ByVal strInput As String, _
    ByRef strReturnCode As String, ByRef strWindows As String, _
    ByRef strMessage As String) As Boolean
'功能：传送数据至设备接口
'参数：
'返回：True成功；False失败
    
    Dim strOutput As String
    Dim objXML As New clsXML
    
    On Error GoTo hErr
    
    strReturnCode = ""
    
    Select Case intType
    Case Val("5-发药窗口")
        strOutput = mobjSOAP.outpWindowNotify(strInput)
        GoSub hResult
        strReturnCode = Val(strOutput)
        
    Case Val("21-配药")
        '再发送配药
        strOutput = mobjSOAP.outpOrderDispense(strInput)
        GoSub hResult
        
        '解析返回信息
        Call objXML.OpenXMLDocument(strOutput)
        Call objXML.GetNodePropertyValue("status", "code", strReturnCode)
        Call objXML.GetNodePropertyValue("status", "message", strMessage)
        Call objXML.GetSingleNodeValue("windowNo", strWindows)
        objXML.CloseXMLDocument
        
    Case Val("22-开始发药")
        strOutput = mobjSOAP.outpOrderTakeNotify(strInput)
        GoSub hResult
        
        '解析返回信息
        Call objXML.OpenXMLDocument(strOutput)
        Call objXML.GetNodePropertyValue("status", "code", strReturnCode)
        Call objXML.GetNodePropertyValue("status", "message", strMessage)
        objXML.CloseXMLDocument
        
    Case Val("23-完成发药")
        strOutput = mobjSOAP.outpOrderCompletionNotify(strInput)
        GoSub hResult
        strReturnCode = Val(strOutput)
    
    End Select
    
    SendDataSingle = Val(strReturnCode) = 0 _
                     Or Val(strReturnCode) = 5 And intType = Val("21-配药")
    
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "标准数据传送失败", 1
    mobjBase.mobjLog.Save
    Exit Function
    
hResult:
    mobjBase.mobjLog.Add "接口返回信息：", 2
    mobjBase.mobjLog.Add strOutput, 2
    Return
End Function

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, _
    ByVal rsData As ADODB.Recordset, _
    Optional ByVal blnSingle As Boolean = False) As String
'功能：将记录集对象的记录转换成信息平台的入参格式
'参数：
'  intType：业务类型
'  intIO：门诊与住院
'  strCode：接口编号
'  rsData：记录集对象
'  blnSingle：True单记录格式化；False所有记录格式化
'返回：标准服务的入参
    
    Dim objXML As New clsXML
    Dim lngPatientID As Long
    Dim strNO As String
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    lngPatientID = -9999        '记录集的“病人ID”字段值可能为Null。防止组织XML不正确
    
    With rsData
        If blnSingle = False Then
            .MoveFirst
        End If
        Do
    
            Select Case intType
            Case Val("5-发药窗口")
                '注意：蝶和要求“windowNo”为数字，请确保“发药窗口”表的“编码”也为纯数字。
            
                objXML.AppendNode "outpWindow"
                
                objXML.AppendData "windowNo", mdlDrugMachine.NVL(!编码)
                objXML.AppendData "windowStatus", IIf(Val(mdlDrugMachine.NVL(!上班)) = 1, "open", "close")
                objXML.AppendData "windowType", "indirect"
                
                objXML.AppendNode "outpWindow", True
                
            Case Val("21-配药")
                '病人
                If lngPatientID <> Val(mdlDrugMachine.NVL(!病人id)) Then
                    objXML.AppendNode "outpOrder"
                    
                    objXML.AppendNode "patient"
                    objXML.AppendData "windowNo", ""
                    objXML.AppendData "patientID", mdlDrugMachine.NVL(!病人id)
                    objXML.AppendData "patientName", mdlDrugMachine.NVL(!姓名)
                    objXML.AppendData "patientGender", mdlDrugMachine.NVL(!性别)
                    objXML.AppendData "patientAge", mdlDrugMachine.NVL(!年龄)
                    objXML.AppendData "identity", mdlDrugMachine.NVL(!身份)
                    objXML.AppendData "insuranceType", mdlDrugMachine.NVL(!医保类别)
                    objXML.AppendData "chargeType", mdlDrugMachine.NVL(!费别)
                    objXML.AppendNode "patient", True
                
                    objXML.AppendNode "prescriptions"
                    
                    GoTo makRecipe
                    
                Else
makRecipe:
                    '处方
                    If strNO <> mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                        objXML.AppendNode "prescription no=""" & mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) & """ " & _
                                          "type=""" & IIf(Val(mdlDrugMachine.NVL(!是否急诊)) = 1, "J", "M") & """ " & _
                                          "paymentDT=""" & Format(mdlDrugMachine.NVL(!缴费时间), "yyyy-MM-DD hh:mm:ss") & """"
                        
                        objXML.AppendData "outpNo", ""
                        objXML.AppendData "visitNo", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
                        objXML.AppendData "deptCode", mdlDrugMachine.NVL(!开嘱科室id)
                        objXML.AppendData "deptName", mdlDrugMachine.NVL(!开嘱科室)
                        objXML.AppendData "doctCode", ""
                        objXML.AppendData "doctName", mdlDrugMachine.NVL(!开嘱医生)
                        objXML.AppendData "diagnosis", mdlDrugMachine.NVL(!诊断描述)
                        objXML.AppendData "paymentDT", mdlDrugMachine.NVL(!缴费时间), xsDateTime
                        objXML.AppendNode "drugList"
                        
                        GoTo makDetail
                        
                    Else
makDetail:
                        objXML.AppendNode "drug"
                        
                        objXML.AppendData "drugCode", mdlDrugMachine.NVL(!药品id)
                        objXML.AppendData "drugName", mdlDrugMachine.NVL(!通用名)
                        objXML.AppendData "drugSpec", mdlDrugMachine.NVL(!规格)
                        objXML.AppendData "firmName", mdlDrugMachine.NVL(!生产商)
                        objXML.AppendData "amount", mdlDrugMachine.NVL(!门诊单位数量)
                        objXML.AppendData "takeUnit", mdlDrugMachine.NVL(!门诊单位)
                        objXML.AppendData "takeMethod", mdlDrugMachine.NVL(!用法)
                        If Val(mdlDrugMachine.NVL(!单次用量)) = 0 Then
                            objXML.AppendData "takeDosage", ""
                        Else
                            objXML.AppendData "takeDosage", mdlDrugMachine.NVL(!单次用量) & _
                                                            mdlDrugMachine.NVL(!剂量单位) & "；" & _
                                                            mdlDrugMachine.NVL(!执行频次)
                        End If
                        objXML.AppendData "takeType", ""
                        objXML.AppendData "takeNote", mdlDrugMachine.NVL(!医生嘱托)
                        objXML.AppendData "pharmacyCode", mdlDrugMachine.NVL(!库房id)
                        objXML.AppendData "sortNo", mdlDrugMachine.NVL(!序号)
                        
                        objXML.AppendNode "drug", True
                    End If
                        
                End If
            
                lngPatientID = Val(mdlDrugMachine.NVL(!病人id))
                strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
            
            Case Val("22-开始发药")
                
                objXML.AppendNode "outpOrderTake"
                
                objXML.AppendData "windowNo", mdlDrugMachine.NVL(!窗口编码)
                objXML.AppendData "patientID", mdlDrugMachine.NVL(!病人id)
                objXML.AppendData "groupNo", mdlDrugMachine.NVL(!groupNo)
                objXML.AppendData "orderType", mdlDrugMachine.NVL(!OrderType)
                
                objXML.AppendNode "outpOrderTake", True
                
            Case Val("23-完成发药")
            
                objXML.AppendNode "outpOrderCompletion"
                
                objXML.AppendData "patientID", mdlDrugMachine.NVL(!病人id)
                objXML.AppendData "groupNo", mdlDrugMachine.NVL(!groupNo)
                
                objXML.AppendNode "outpOrderCompletion", True
                
            End Select
        
            
            If blnSingle = False Then
                .MoveNext
            End If
            
            '结点收尾
            Select Case intType
            Case 21
                If .EOF Then
                    objXML.AppendNode "drugList", True
                    objXML.AppendNode "prescription", True
                    objXML.AppendNode "prescriptions", True
                    objXML.AppendNode "outpOrder", True
                Else
                    If strNO <> mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                        objXML.AppendNode "drugList", True
                        objXML.AppendNode "prescription", True
                    End If
                    If lngPatientID <> Val(mdlDrugMachine.NVL(!病人id)) Then
                        objXML.AppendNode "prescriptions", True
                        objXML.AppendNode "outpOrder", True
                    End If
                End If
            Case 22, 23
                If blnSingle Then
                    '单个记录格式化
                    Exit Do
                End If
            End Select
            
        Loop While .EOF = False
        
    End With
    
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "格式化数据完成", 1
    mobjBase.mobjLog.Save
    FormatData = objXML.XmlText
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "格式化数据失败", 1
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
'
End Property

Private Property Get clsINF_Token() As String
'
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
        
    Call CreateSOAP(mobjSOAP, mobjBase)
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
    
    'Call mdlDrugMachine.GetUserInfo(mobjBase.UserName, mobjBase.mobjComLib, mobjBase.mobjLog, UserInfo)
End Property
