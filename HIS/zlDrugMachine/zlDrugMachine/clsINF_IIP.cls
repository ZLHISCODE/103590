VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_IIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口（中联信息交互平台）类
'IIP：Information Interchange Platform
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjHTTP As Object

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'说明：开启接口为活动状态

    mobjBase.Action = RHS
    
    If RHS Then
        If mobjHTTP Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "创建WinHTTP部件失败", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "开始连接信息交互平台", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        
        On Error Resume Next
        Call mobjHTTP.Open("POST", mobjBase.URL)
        If Err.Number = -2147012891 Then
            mobjBase.mobjLog.Add mobjHTTP.ResponseText, 1, 1
            mobjBase.mobjLog.Add "连接信息交互平台异常", 1
        Else
            mobjBase.mobjLog.Add "成功连接信息交互平台", 1
        End If
        On Error GoTo 0
    Else
        Set mobjHTTP = Nothing
    End If
    
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Public Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的Filter方法

    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)

End Function

Public Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, ByRef blnFinish As Boolean) As ADODB.Recordset
    Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '如果不使用基类的记录集对象，请将前面代码屏蔽，再以下面编写新的代码
    'TODO：
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

Public Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的SendData方法

    Dim arrFilter As Variant
    Dim lngPatientID As Long, lngStockID As Long
    Dim intBill As Integer
    Dim strNO As String
    Dim i As Long
    Dim rsNew As New ADODB.Recordset
    
    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "开始传送数据", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "停用=0"
    Case 21 To 25
        rsData.Filter = "是否传送=1"
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

    If intType = 21 And intIO = 1 Then
        '需要返回窗口信息，按“病人、库房”分别传送
        
        arrFilter = Array()
        
        '取病人与库房
        Do While rsData.EOF = False
            If Not (lngPatientID = rsData!病人id And lngStockID = rsData!库房id) Then
                ReDim Preserve arrFilter(UBound(arrFilter) + 1)
                arrFilter(UBound(arrFilter)) = "病人id=" & rsData!病人id & " And 库房id=" & rsData!库房id & " And 是否传送=1"
            End If
            lngPatientID = rsData!病人id
            lngStockID = rsData!库房id
            
            rsData.MoveNext
        Loop
        
        '传送数据
        For i = LBound(arrFilter) To UBound(arrFilter)
            rsData.Filter = arrFilter(i)
            Call SendDataEx(intAppType, intType, intIO, strCode, rsData, strMessage)
        Next
    
    ElseIf intType >= 22 And intType <= 24 And intIO = 1 Then
        '22-开始发药；23-完成发药；24-处方退药；按“单据、处方号”分别传送
        
        '取单据与处方号
        Do While rsData.EOF = False
            If Not (strNO = Trim(rsData!处方号) And intBill = rsData!单据) Then
                ReDim Preserve arrFilter(UBound(arrFilter) + 1)
                arrFilter(UBound(arrFilter)) = "处方号='" & Trim(rsData!处方号) & "' And 单据=" & intBill & " And 是否传送=1"
            End If
            strNO = Trim(rsData!处方号)
            intBill = rsData!单据
            
            rsData.MoveNext
        Loop
        
        '传送数据
        For i = LBound(arrFilter) To UBound(arrFilter)
            rsData.Filter = arrFilter(i)
            Call SendDataEx(intAppType, intType, intIO, strCode, rsData, strMessage)
        Next
        
    Else
        '按记录行传送
        
        Set rsNew = mdlDrugMachine.CopyStructure(rsData.Fields)
        If rsNew.State = adStateOpen Then
            i = 1
            Do While rsData.EOF = False
                '复制记录
                strMessage = CopyRecord(rsData.Fields, rsNew)
                If strMessage = "" Then
                    '按50行传送1次
                    If i Mod 50 = 0 Then
                        '传送数据
                        Call SendDataEx(intAppType, intType, intIO, strCode, rsNew, strMessage)
                        '清除数据
                        If rsNew.RecordCount > 0 Then Call ClearRecord(rsNew)
                    End If
                    i = i + 1
                Else
                    mobjBase.mobjLog.Add "创建记录失败。" & strMessage, 1
                End If
                
                rsData.MoveNext
            Loop
            If rsNew.State = adStateOpen Then
                If rsNew.RecordCount > 0 Then
                    Call SendDataEx(intAppType, intType, intIO, strCode, rsNew, strMessage)
                    Call ClearRecord(rsNew)
                End If
                rsNew.Close
            End If
        Else
            mobjBase.mobjLog.Add "创建记录集对象失败", 1
        End If
    End If
    
    mobjBase.mobjLog.Add "传送数据完成", 1
    mobjBase.mobjLog.Save
    
    '执行基类的代码
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
hErr:
    '标记失败（药品收发记录）
    strMessage = Err.Number & "：" & Err.Description
    mobjBase.mobjLog.Add strMessage, 1
    mobjBase.mobjLog.Add "传送数据失败", 1
    mobjBase.mobjLog.Save
End Function

Private Sub SendDataEx(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, ByRef strMessage As String)
    
    Const WinHttpRequestOption_SslErrorIgnoreFlags As Long = 4
    
    Dim strInput As String, strOutput As String, strVal As String
    Dim strCData As String, strWindows As String, strTmp As String
    Dim strErrCode As String, strErrMsg As String
    Dim strFuncURL As String, strService As String
    Dim blnFinish As Boolean
    Dim objXML As New clsXML, objData As New clsXML
    Dim objDB As Object, objEncrypt As Object
    
    On Error GoTo hErr
    
    '格式化数据
    strTmp = "格式化数据"
    mobjBase.mobjLog.Add "开始" & strTmp, 1
    strInput = FormatData(intType, intIO, strCode, rsData)
    If strInput = "" Then
        strMessage = "格式化数据失败"
        Exit Sub
    End If
    
    '标准数据传送开始
    strTmp = "标准数据传送"
    mobjBase.mobjLog.Add "开始" & strTmp, 1
    
    '平台接口
    strService = "Drug/Machine"
    If Right(mobjBase.URL, 1) = "/" Then
        strFuncURL = mobjBase.URL & strService
    Else
        strFuncURL = mobjBase.URL & "/" & strService
    End If
    
    '传送数据
    On Error Resume Next
    Call mobjHTTP.Open("POST", strFuncURL, False)
'    mobjHTTP.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300     '忽略错误
    If Err.Number <> -2147012891 Then
        If Err.Number = 0 Then
            On Error GoTo hErr
        
            mobjHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            'mobjHTTP.SetRequestHeader "Content-Length", LenB(StrConv(strInput, vbFromUnicode))
            mobjHTTP.Send "redata=" & strInput
            strOutput = mobjHTTP.ResponseText
            mobjBase.mobjLog.Add strOutput, 1, 1
            mobjBase.mobjLog.Add "传送数据到信息交互平台完成", 1
        Else
            mobjBase.mobjLog.Add Err.Description, 1, 1
            mobjBase.mobjLog.Add "WinHttpRequest.Open失败", 1
        End If
    Else
        mobjBase.mobjLog.Add "WinHTTP连接失败，传送数据到信息交互平台不成功", 1
    End If
    On Error GoTo hErr
    
'<ROOT>
'   <STATE><![CDATA[状态]]></STATE>
'   <DATAPARAM><![CDATA[业务数据]]></DATAPARAM>
'   <ERROR>
'       <ERRCODE><![CDATA[错误码]]></ERRCODE>
'       <MSG><![CDATA[错误信息]]></MSG>
'   </ERROR>
'</ROOT>
    
    '解析返回信息
    Call objXML.OpenXMLDocument(strOutput)
    Call objXML.GetSingleNodeValue("STATE", strVal)
    
    If UCase(strVal) = "T" Then
        '成功
        Call objXML.GetSingleNodeValue("DATAPARAM", strCData)
        If intType = Val("21-配药") And intIO = Val("1-门诊") Then
            On Error Resume Next
            Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
            If Err.Number <> 0 Then
                mobjBase.mobjLog.Add "zlEncryptPub部件未注册，拒绝格式化数据", 1
                Err.Clear
                On Error GoTo hErr
            Else
                On Error GoTo hErr
                strCData = objEncrypt.AES.DecryptStr(strCData, objEncrypt.Base64Decode(mobjBase.KEY))
                Call objData.OpenXMLDocument(strCData)
                Call objData.GetSingleNodeValue("FYCK", strWindows)       '发药窗口
                objData.CloseXMLDocument
            End If
        End If
        blnFinish = True
        
    ElseIf UCase(strVal) = "F" Then
        '失败
        Call objXML.GetSingleNodeValue("ERRCODE", strErrCode)
        Call objXML.GetSingleNodeValue("MSG", strErrMsg)
        
        strMessage = "[" & strErrCode & "]" & strErrMsg
        
    Else
        '返回信息异常
        strMessage = "返回信息异常"
    End If
    
    objXML.CloseXMLDocument
    
    If blnFinish Then
        '成功
        If intType = Val("21-配药") And intIO = Val("1-门诊") Then
            '回写发药窗口
            If strWindows <> "" Then
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
        End If
    ElseIf Val(strVal) = Val("2-信息交互平台异常，第三方无异常") Then
        mobjBase.mobjLog.Add strMessage, 1
        mobjBase.mobjLog.Add "信息交互平台异常", 1
    Else
        mobjBase.mobjLog.Add strMessage, 1
        mobjBase.mobjLog.Add "第三方接口异常", 1
    End If
    
    '标准数据传送完成
    mobjBase.mobjLog.Add "完成" & strTmp, 1
    
    '标记成功（药品收发记录），之所以失败也标记，是为了让接口管理工具的定时传送可补充传送。
    If intType > 20 Then
        If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
            mobjBase.mobjLog.Add "“传送”标记成功", 1
        Else
            mobjBase.mobjLog.Add "“传送”标记失败", 1
        End If
    End If
    mobjBase.mobjLog.Save
    
    Exit Sub
    
hErr:
    strMessage = Err.Number & ":" & Err.Description
    mobjBase.mobjLog.Add strMessage, 1
    mobjBase.mobjLog.Add strTmp & "失败", 1
    mobjBase.mobjLog.Save
End Sub

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset) As String
'功能：将记录集对象的记录转换成信息平台的入参格式
'参数：
'  intType：业务类型
'  intIO：门诊与住院
'  strCode：接口编号
'  rsData：记录集对象
'返回：标准服务的入参

    Dim strTmp As String, strNO As String, strService As String
    Dim l As Long, lngDeptID As Long, lngPatientID As Long, lngDeptTakeID As Long
    Dim arrTmp As Variant
    Dim objEncrypt As Object
    Dim objXML As New clsXML, objData As New clsXML

    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error Resume Next
    Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        mobjBase.mobjLog.Add "zlEncryptPub部件未注册，拒绝格式化数据", 1
        mobjBase.mobjLog.Save
        Exit Function
    End If
    On Error GoTo hErr
    
    lngPatientID = -9999        '记录集的“病人ID”字段值可能为Null。防止组织XML不正确
    
    '标准服务名
    strService = "NotSupport"

    Select Case intType
    Case 1      '部门
        strService = "Drug.Machine.Base.Dept"
    Case 2      '人员
        strService = "Drug.Machine.Base.Worker"
    Case 3      '药品
        strService = "Drug.Machine.Base.Drug"
    Case 4      '库存
        strService = "Drug.Machine.Base.Stock"
    Case 5      '窗口
        strService = "Drug.Machine.Base.Window"
    Case 21     '配药
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.Dosage"
        Else
            strService = "Drug.Machine.Inpatient.Dosage"
        End If
    Case 22     '开始发药
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.SendStart"
        End If
    Case 23     '完成发药
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.SendComplete"
        End If
    Case 24     '退药
        If intIO = 1 Then
            strService = "Drug.Machine.Outpatient.Back"
        End If
    End Select
    
'<ROOT>
'    <TOKEN><![CDATA[令牌]]></TOKEN>
'    <SERVICE><![CDATA[服务名称]]></SERVICE>
'    <DATAPARAM>
'       <![CDATA[业务数据]]>
'    </DATAPARAM>
'</ROOT>
    
    objXML.AppendNode "ROOT"
    strTmp = "<![CDATA[" & objEncrypt.AES.EncryptStr(objEncrypt.Base64Decode(mobjBase.Token), objEncrypt.Base64Decode(mobjBase.KEY)) & "]]>"
    objXML.AppendData "TOKEN", strTmp, xsString, True
    objXML.AppendData "SERVICE", "<![CDATA[" & strService & "]]>", xsString, True
    
    objData.AppendData "JKBH", strCode
    objData.AppendNode "ITEMS"
    
    With rsData
        
        If .RecordCount > 0 Then .MoveFirst
        Do While .EOF = False
            
            Select Case intType
            Case 1 To 5
                objData.AppendNode "ITEM"
            End Select
            
            Select Case intType
            Case 1      '部门信息
                
                objData.AppendData "ID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "BM", mdlDrugMachine.NVL(!编码)
                objData.AppendData "MC", mdlDrugMachine.NVL(!名称)
                objData.AppendData "FWDX", mdlDrugMachine.NVL(!服务对象)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!停用)
                
                arrTmp = Split(mdlDrugMachine.NVL(!工作性质), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "BMXZ"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "XZ", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "BMXZ", True
                    
            Case 2      '人员信息
            
                objData.AppendData "ID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "BH", mdlDrugMachine.NVL(!编号)
                objData.AppendData "XM", mdlDrugMachine.NVL(!姓名)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!停用)
                
                '人员性质
                arrTmp = Split(mdlDrugMachine.NVL(!人员性质), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "RYXZ"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "XZ", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "RYXZ", True
                
                '所属部门
                arrTmp = Split(mdlDrugMachine.NVL(!部门ID), ",")
                
                If UBound(arrTmp) >= 0 Then objData.AppendNode "SSBM"
                For l = LBound(arrTmp) To UBound(arrTmp)
                    objData.AppendData "BMID", arrTmp(l)
                Next
                If UBound(arrTmp) >= 0 Then objData.AppendNode "SSBM", True
            
            Case 3      '药品目录
            
                objData.AppendData "YPID", mdlDrugMachine.NVL(!ID)
                objData.AppendData "YPBM", mdlDrugMachine.NVL(!编码)
                objData.AppendData "YPGG", mdlDrugMachine.NVL(!规格)
                objData.AppendData "YPMC", mdlDrugMachine.NVL(!通用名)
                objData.AppendData "SPMC", mdlDrugMachine.NVL(!商品名)
                
                objData.AppendData "YWMC", mdlDrugMachine.NVL(!英文名)
                objData.AppendData "PYJM", mdlDrugMachine.NVL(!通用名简码)
                objData.AppendData "YPLB", mdlDrugMachine.NVL(!药品类别)
                objData.AppendData "YPJX", mdlDrugMachine.NVL(!药品剂型)
                objData.AppendData "YKDW", mdlDrugMachine.NVL(!药库单位)
                
                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!住院单位)
                objData.AppendData "MZDW", mdlDrugMachine.NVL(!门诊单位)
                objData.AppendData "ZXDW", mdlDrugMachine.NVL(!计算单位)
                objData.AppendData "JLXS", mdlDrugMachine.NVL(!剂量系数)
                objData.AppendData "JLDW", mdlDrugMachine.NVL(!剂量单位)
                
                objData.AppendData "YKBZ", mdlDrugMachine.NVL(!药库包装)
                objData.AppendData "ZYBZ", mdlDrugMachine.NVL(!住院包装)
                objData.AppendData "MZBZ", mdlDrugMachine.NVL(!门诊包装)
                objData.AppendData "DLFL", mdlDrugMachine.NVL(!毒理分类)
                objData.AppendData "PZWH", mdlDrugMachine.NVL(!批准文号)
                
                objData.AppendData "SCSDM", mdlDrugMachine.NVL(!生产商编码)
                objData.AppendData "SCS", mdlDrugMachine.NVL(!生产商)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!停用)
            
            Case 4      '药品库存与库位
            
                objData.AppendData "KFID", mdlDrugMachine.NVL(!库房id)
                objData.AppendData "YPID", mdlDrugMachine.NVL(!药品id)
                objData.AppendData "YPBM", mdlDrugMachine.NVL(!药品编码)
                objData.AppendData "YPWZ", mdlDrugMachine.NVL(!库房货位)
                objData.AppendData "PC", mdlDrugMachine.NVL(!批次)
                
                objData.AppendData "PH", mdlDrugMachine.NVL(!上次批号)
                objData.AppendData "SCRQ", mdlDrugMachine.NVL(!上次生产日期)
                objData.AppendData "SXRQ", mdlDrugMachine.NVL(!效期)
                objData.AppendData "YKDW", mdlDrugMachine.NVL(!药库单位)
                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!住院单位)
                
                objData.AppendData "MZDW", mdlDrugMachine.NVL(!门诊单位)
                objData.AppendData "SJDW", mdlDrugMachine.NVL(!计算单位)
                objData.AppendData "YKBZ", mdlDrugMachine.NVL(!药库包装)
                objData.AppendData "ZYBZ", mdlDrugMachine.NVL(!住院包装)
                objData.AppendData "MZBZ", mdlDrugMachine.NVL(!门诊包装)
                
                objData.AppendData "SJDWSL", mdlDrugMachine.NVL(!实际数量)
                objData.AppendData "SCSBM", mdlDrugMachine.NVL(!生产商编码)
                objData.AppendData "SCS", mdlDrugMachine.NVL(!生产商)
                objData.AppendData "GYSBM", mdlDrugMachine.NVL(!供应商编码)
                objData.AppendData "GYS", mdlDrugMachine.NVL(!供应商)
                
            Case 5      '发药窗口
            
                objData.AppendData "KFID", mdlDrugMachine.NVL(!药房id)
                objData.AppendData "CKBM", mdlDrugMachine.NVL(!编码)
                objData.AppendData "CKMC", mdlDrugMachine.NVL(!名称)
                objData.AppendData "SFTY", mdlDrugMachine.NVL(!上班)
            
            Case 21     '配药
            
                If intIO = 1 Then
                    '门诊配（摆）药
                    
                    If lngPatientID <> Val(mdlDrugMachine.NVL(!病人id)) Then
                        '病人
                        objData.AppendNode "BRXX"
                        
                        objData.AppendData "ID", mdlDrugMachine.NVL(!病人id)
                        objData.AppendData "XM", mdlDrugMachine.NVL(!姓名)
                        objData.AppendData "XB", mdlDrugMachine.NVL(!性别)
                        objData.AppendData "NL", mdlDrugMachine.NVL(!年龄)
                        objData.AppendData "JZKH", mdlDrugMachine.NVL(!就诊卡号)
                        
                        objData.AppendData "YBLB", mdlDrugMachine.NVL(!医保类别)
                        objData.AppendData "FB", mdlDrugMachine.NVL(!费别)
                        objData.AppendData "YXJ", mdlDrugMachine.NVL(!优先级)
                        
                        objData.AppendNode "BRXX", True
                        
                        GoTo makBill
                    Else
makBill:
                        '处方
                        If strNO <> mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                            objData.AppendNode "DJ"
                            
                            objData.AppendData "CFSJ", mdlDrugMachine.NVL(!处方时间), xsDateTime
                            objData.AppendData "CFH", mdlDrugMachine.NVL(!处方号)
                            objData.AppendData "DJLX", mdlDrugMachine.NVL(!单据)
                            objData.AppendData "FYYFID", Val(mdlDrugMachine.NVL(!库房id))
                            objData.AppendData "FYYF", mdlDrugMachine.NVL(!发药药房)
                            
                            objData.AppendData "CFZD", mdlDrugMachine.NVL(!诊断描述)
                            objData.AppendData "CFBZ", ""
                            objData.AppendData "JFSJ", mdlDrugMachine.NVL(!缴费时间), xsDateTime
                            objData.AppendData "KFKSID", Val(mdlDrugMachine.NVL(!开嘱科室id))
                            objData.AppendData "KFKS", mdlDrugMachine.NVL(!开嘱科室)
                            
                            objData.AppendData "KFYS", mdlDrugMachine.NVL(!开嘱医生)
                            
                            GoTo makDetail
                        Else
makDetail:
                            '处方明细
                            objData.AppendNode "MX"
                            
                            objData.AppendData "XH", Val(mdlDrugMachine.NVL(!序号))
                            objData.AppendData "YZID", Val(mdlDrugMachine.NVL(!医嘱id))
                            objData.AppendData "YPID", Val(mdlDrugMachine.NVL(!药品id))
                            objData.AppendData "YPBM", mdlDrugMachine.NVL(!编码)
                            objData.AppendData "YPGG", mdlDrugMachine.NVL(!规格)
                            
                            objData.AppendData "YPMC", mdlDrugMachine.NVL(!通用名)
                            objData.AppendData "SPMC", mdlDrugMachine.NVL(!商品名)
                            objData.AppendData "SCS", mdlDrugMachine.NVL(!生产商)
                            objData.AppendData "MZDWSL", Val(mdlDrugMachine.NVL(!门诊单位数量))
                            objData.AppendData "MZDW", mdlDrugMachine.NVL(!门诊单位)
                            
                            objData.AppendData "YF", mdlDrugMachine.NVL(!用法)
                            objData.AppendData "DL", mdlDrugMachine.NVL(!单次用量)
                            objData.AppendData "DLDW", mdlDrugMachine.NVL(!剂量单位)
                            objData.AppendData "ZXPC", mdlDrugMachine.NVL(!执行频次)
                            objData.AppendData "YPJX", mdlDrugMachine.NVL(!药品剂型)
                            
                            objData.AppendData "DLFL", mdlDrugMachine.NVL(!毒理分类)
                            objData.AppendData "YSZT", mdlDrugMachine.NVL(!医生嘱托)
                            
                            objData.AppendNode "MX", True
                            
                        End If
                        strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
                            
                    End If
                    lngPatientID = Val(mdlDrugMachine.NVL(!病人id))
                
                ElseIf intIO = 2 Then
                    '住院配（摆）药
                    
                    '发药药房
                    If lngDeptID <> Val(mdlDrugMachine.NVL(!库房id)) Then
                        objData.AppendNode "FYYF"
                        
                        objData.AppendData "FYYFID", mdlDrugMachine.NVL(!发药药房id)
                        objData.AppendData "FYYF", mdlDrugMachine.NVL(!发药药房)
                        objData.AppendData "FYH", mdlDrugMachine.NVL(!发药号)
                        objData.AppendData "CSSJ", mdlDrugMachine.NVL(!传送时间), xsDateTime
                        
                        GoTo makDeptTake
                    Else
makDeptTake:
                        '领药科室
                        If lngDeptTakeID <> Val(mdlDrugMachine.NVL(!领药科室ID)) Then
                            objData.AppendNode "LYKS"
                            
                            objData.AppendData "LYKSID", mdlDrugMachine.NVL(!领药科室ID)
                            objData.AppendData "LYKSMC", mdlDrugMachine.NVL(!领药科室)
                            
                            '病人
                            If lngPatientID <> Val(mdlDrugMachine.NVL(!病人id)) Then
                                objData.AppendNode "BR"
                                
                                objData.AppendData "BRID", mdlDrugMachine.NVL(!病人id)
                                objData.AppendData "BRXM", mdlDrugMachine.NVL(!病人姓名)
                                objData.AppendData "BRXB", mdlDrugMachine.NVL(!病人性别)
                                objData.AppendData "BRNL", mdlDrugMachine.NVL(!病人年龄)
                                objData.AppendData "JZKH", mdlDrugMachine.NVL(!就诊卡号)
                                
                                objData.AppendData "BQID", mdlDrugMachine.NVL(!病区id)
                                objData.AppendData "BQ", mdlDrugMachine.NVL(!病区)
                                objData.AppendData "CH", mdlDrugMachine.NVL(!床号)
                                objData.AppendData "YXJ", mdlDrugMachine.NVL(!优先级)
                                
                                GoTo makMedical
                            Else
makMedical:
                                '医嘱
                                objData.AppendNode "MX"
                                
                                objData.AppendData "KZYS", mdlDrugMachine.NVL(!开嘱医生)
                                objData.AppendData "KZSJ", mdlDrugMachine.NVL(!开嘱时间), xsDateTime
                                objData.AppendData "YZID", mdlDrugMachine.NVL(!开嘱id)
                                objData.AppendData "YPJX", mdlDrugMachine.NVL(!药品剂型)
                                objData.AppendData "YPID", mdlDrugMachine.NVL(!药品id)
                                
                                objData.AppendData "YPMC", mdlDrugMachine.NVL(!药品名称)
                                objData.AppendData "SPMC", mdlDrugMachine.NVL(!商品名称)
                                objData.AppendData "YF", mdlDrugMachine.NVL(!用法)
                                objData.AppendData "DL", mdlDrugMachine.NVL(!单量)
                                objData.AppendData "DLDW", mdlDrugMachine.NVL(!单量单位)
                                
                                objData.AppendData "ZYDW", mdlDrugMachine.NVL(!住院单位)
                                objData.AppendData "ZYDWSL", mdlDrugMachine.NVL(!住院单位数量)
                                objData.AppendData "YZQX", mdlDrugMachine.NVL(!医嘱期效)
                                objData.AppendData "ZXSJ", mdlDrugMachine.NVL(!执行时间), xsDateTime
                                objData.AppendData "YSZT", mdlDrugMachine.NVL(!医生嘱托)
                                
                                objData.AppendData "DLFL", mdlDrugMachine.NVL(!毒理分类)
                                objData.AppendData "HSJDSJ", mdlDrugMachine.NVL(!护士校对时间), xsDateTime
                                objData.AppendData "SFTZR", mdlDrugMachine.NVL(!收发填制人)
                                objData.AppendData "SFSHR", mdlDrugMachine.NVL(!收发审核人)
                                
                                objData.AppendNode "MX", True
                                
                            End If
                            lngPatientID = Val(mdlDrugMachine.NVL(!病人id))
                            
                        End If
                        lngDeptTakeID = Val(mdlDrugMachine.NVL(!领药科室ID))
                        
                    End If
                    lngDeptID = Val(mdlDrugMachine.NVL(!库房id))
                    
                End If
            
            Case 22, 23     '22-开始发药；23-完成发药
                
                If intIO = 1 Then
                    objData.AppendNode "DJ"
                    
                    objData.AppendData "CFSJ", mdlDrugMachine.NVL(!处方时间), xsDateTime
                    objData.AppendData "CFH", mdlDrugMachine.NVL(!处方号)
                    objData.AppendData "DJLX", mdlDrugMachine.NVL(!单据)
                    objData.AppendData "FYYFID", Val(mdlDrugMachine.NVL(!库房id))
                    objData.AppendData "BRID", Val(mdlDrugMachine.NVL(!病人id))
                    objData.AppendData "YPJX", mdlDrugMachine.NVL(!药品剂型)
                    
                    objData.AppendNode "DJ", True
                End If
            
            Case 24     '退药
            
                If intIO = 1 Then
                    
                    If strNO <> mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                        objData.AppendNode "DJ"
                        
                        objData.AppendData "CFSJ", mdlDrugMachine.NVL(!处方时间), xsDateTime
                        objData.AppendData "CFH", mdlDrugMachine.NVL(!处方号)
                        objData.AppendData "DJLX", mdlDrugMachine.NVL(!单据)
                        objData.AppendData "FYYFID", mdlDrugMachine.NVL(!库房id)
                        objData.AppendData "BRID", Val(mdlDrugMachine.NVL(!病人id))
                        
                        GoTo makReturnDetail
                    Else
makReturnDetail:
                        objData.AppendNode "YPMX"
                        
                        objData.AppendData "YPJX", mdlDrugMachine.NVL(!药品剂型)
                        objData.AppendData "YZID", Val(mdlDrugMachine.NVL(!医嘱id))
                        objData.AppendData "YPID", Val(mdlDrugMachine.NVL(!药品id))
                        objData.AppendData "YPBM", mdlDrugMachine.NVL(!药品编码)
                        objData.AppendData "MZDWSL", mdlDrugMachine.NVL(!门诊单位数量)
                        
                        objData.AppendData "MZDW", mdlDrugMachine.NVL(!门诊单位)
                        objData.AppendData "DLFL", mdlDrugMachine.NVL(!毒理分类)
                        
                        objData.AppendNode "YPMX", True
                    End If
                    
                    strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
                    
                End If
            
            End Select
            
            Select Case intType
            Case 1 To 5
                objData.AppendNode "ITEM", True
            End Select
            
            .MoveNext
            
            'XML收尾

            Select Case intType
            Case 21, 24    '21-配药；24-退药
                If intIO = 1 Then
                    '门诊
                    If .EOF Then
                        objData.AppendNode "DJ", True
                    Else
                        If strNO <> mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                            objData.AppendNode "DJ", True
                        End If
                    End If
                    
                ElseIf intIO = 2 Then
                    '住院
                    If .EOF Then
                        objData.AppendNode "BR", True
                        objData.AppendNode "LYKS", True
                        objData.AppendNode "FYYF", True
                    Else
                        If lngPatientID <> Val(!病人id) Then objData.AppendNode "BR", True
                        If lngDeptTakeID <> Val(!领药科室ID) Then objData.AppendNode "LYKS", True
                        If lngDeptID <> Val(!库房id) Then objData.AppendNode "FYYF", True
                    End If
                End If
            End Select
            
        Loop
    End With
    
    objData.AppendNode "ITEMS", True
    
    '加密
    strTmp = objData.XmlText
    mobjBase.mobjLog.Add Replace(strTmp, vbCrLf, ""), 1, 1
    objData.ClearXmlText
    
    If Not objEncrypt Is Nothing Then
        strTmp = objEncrypt.AES.EncryptStr(strTmp, objEncrypt.Base64Decode(mobjBase.KEY))
        strTmp = "<![CDATA[" & strTmp & "]]>"
        mobjBase.mobjLog.Add "加密成功", 1
    Else
        strTmp = "Encrypt.Exception"
        mobjBase.mobjLog.Add "加密失败（zlEncryptPub部件未注册）", 1
    End If
    
    objXML.AppendData "DATAPARAM", strTmp, xsString, True
    objXML.AppendNode "ROOT", True
    
    strTmp = objXML.XmlText
    objXML.ClearXmlText
    
    mobjBase.mobjLog.Add "格式化数据完成", 1
    mobjBase.mobjLog.Save
    FormatData = strTmp
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objData.XmlText, 1, 1
    mobjBase.mobjLog.Add "格式化数据失败", 1
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    mobjBase.Token = RHS
End Property

Private Property Get clsINF_Token() As String
    clsINF_Token = mobjBase.Token
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
        
    Call CreateHTTP(mobjHTTP, mobjBase)
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property

Private Function GetXML_CDATA(ByVal strVar As String) As String
    Const STR_CDATA_S As String = "<![CDATA["
    Const STR_CDATA_E As String = "]]>"
    
    If InStr(UCase(strVar), STR_CDATA_S) > 0 Then
        strVar = Mid(strVar, InStr(UCase(strVar), STR_CDATA_S) + Len(STR_CDATA_S))
        If InStr(strVar, STR_CDATA_E) > 0 Then
            strVar = Left(strVar, InStr(strVar, STR_CDATA_E) - Len(STR_CDATA_E))
        Else
            strVar = ""
        End If
    Else
        strVar = ""
    End If
End Function
