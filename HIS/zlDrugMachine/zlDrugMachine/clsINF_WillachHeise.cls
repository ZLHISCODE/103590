VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_WillachHeise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口（韦乐海茨）类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mobjSOAP As Object
Private UserInfo As TYPE_USER_INFO

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
            mobjBase.mobjLog.Add "连接韦乐海茨服务器失败", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "开始连接韦乐海茨服务器", 1
        mobjBase.mobjLog.Add "URL: " & mobjBase.URL, 2
        Call mobjSOAP.MSSoapInit(mobjBase.URL)
        mobjBase.mobjLog.Add "成功连接韦乐海茨服务器", 1
    Else
        Set mobjSOAP = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Add "连接韦乐海茨服务器异常", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Public Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的Filter方法
    
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)

End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
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

    Dim strInput As String, strOutput As String
    Dim strVal As String, strReturn As String, strWindows As String
    Dim objXML As New clsXML
    Dim blnFinish As Boolean
    Dim objDB As Object
    
    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "开始传送数据", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "停用=0"
    Case 4 To 5
        rsData.Filter = ""
    Case 21 To 24
        rsData.Filter = "是否传送=1"
    Case 25
        rsData.Filter = "是否传送=1 And 单据=8"
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
    
    '格式化数据
    mobjBase.mobjLog.Add "格式化数据", 1
    strInput = FormatData(intType, intIO, strCode, rsData)
    If strInput = "" Then
        strMessage = "无数据可格式化"
        Exit Function
    End If
    
    '标准数据传送开始
    mobjBase.mobjLog.Add "标准数据传送开始(HisTransData)", 1
    
    '传送数据
    On Error GoTo makFinish
    strOutput = mobjSOAP.HisTransData(strInput)
    On Error GoTo hErr
    mobjBase.mobjLog.Add "接口返回信息：", 2
    mobjBase.mobjLog.Add strOutput, 2
    
    '解析返回信息
    Call objXML.OpenXMLDocument(strOutput)
    Call objXML.GetSingleNodeValue("RETVAL", strVal)                '设备处理当天的第几张处方
    Call objXML.GetSingleNodeValue("RETCODE", strReturn)            '1-执行成功；其他值-执行失败
    Call objXML.GetSingleNodeValue("RETMSG", strWindows)            '当前处方发药窗口编码或错误信息
    objXML.CloseXMLDocument
    blnFinish = Val(strReturn) = 1
    
    If blnFinish Then
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
        mobjBase.mobjLog.Add "韦乐海茨接口异常", 1
    End If
    
makFinish:
    If Err.Number <> 0 Then
        blnFinish = False
        Err.Clear: On Error GoTo hErr
        mobjBase.mobjLog.Add "韦乐海茨接口异常或者网络异常", 1
    End If
    
    '标记成功（药品收发记录），之所以失败也标记，是为了让接口管理工具的定时传送可补充传送。
    If InStr(";21;22;23;25;", ";" & intType & ";") > 0 Then
        If mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, blnFinish) Then
            mobjBase.mobjLog.Add "“传送”标记成功", 1
        Else
            mobjBase.mobjLog.Add "“传送”标记失败", 1
        End If
    End If
    
    '标准数据传送完成
    mobjBase.mobjLog.Add "标准数据传送完成", 1
    
    mobjBase.mobjLog.Save
    
    '执行基类的代码
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add "标准数据传送失败", 1
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Save
End Function

Private Function FormatData(ByVal intType As Integer, ByVal intIO As Integer, ByVal strCode As String, ByVal rsData As ADODB.Recordset) As String
'功能：将记录集对象的记录转换成信息平台的入参格式
'参数：
'  intType：业务类型
'  intIO：门诊与住院
'  strCode：接口编号
'  rsData：记录集对象
'返回：标准服务的入参

    Dim objXML As New clsXML
    Dim strIP As String, strPCode As String, strPName As String, strTmp As String
    Dim strItem As String, strNO As String
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    With rsData
    
        strIP = mdlDrugMachine.IP
        strPCode = UserInfo.编号
        strPName = UserInfo.姓名
        
        Select Case intType
        Case 1      '部门
            strTmp = "104"
        Case 2      '人员
            strTmp = "107"
        Case 3      '药品
            strTmp = "101"
        Case 4      '库存与库位
            strTmp = "102"
        Case 5      '发药窗口
            strTmp = "106"
        Case 21     '配药
            strTmp = "201"
        Case 22     '开始发药
            strTmp = "202"
        Case 23     '完成发药
            strTmp = "203"
        Case 24     '退药
            strTmp = "204"
        Case 25     '整单退药
            strTmp = "205"
        Case Else
            Exit Function
        End Select
    
        .MoveFirst
        
        objXML.AppendNode "ROOT"
        objXML.AppendData "OPSYSTEM", "HIS"
        Select Case intType
        Case 22, 23
            objXML.AppendData "OPWINID", mdlDrugMachine.NVL(!窗口编码)
        Case Else
            objXML.AppendData "OPWINID", ""
        End Select
        
        objXML.AppendData "OPTYPE", strTmp
        objXML.AppendData "OPIP", strIP
        objXML.AppendData "OPMANNO", SpecialChar(strPCode)
        objXML.AppendData "OPMANNAME", SpecialChar(strPName)
        
        Do
            Select Case intType
            Case 1      '部门
                strItem = "CONSIS_BASIC_DEPTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DEPTCODE", mdlDrugMachine.NVL(!ID)
                objXML.AppendData "DEPTNAME", SpecialChar(mdlDrugMachine.NVL(!名称))
                objXML.AppendData "OUTP_OR_INP", mdlDrugMachine.NVL(!服务对象)
                
                objXML.AppendNode strItem, True
                
            Case 2      '人员
                strItem = "CONSIS_BASIC_STAFFERVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "STAFFERID", mdlDrugMachine.NVL(!ID)
                objXML.AppendData "STAFFERNO", SpecialChar(mdlDrugMachine.NVL(!编号))
                objXML.AppendData "DEPTCODE", mdlDrugMachine.NVL(!部门ID)
                objXML.AppendData "STAFFERNAME", SpecialChar(mdlDrugMachine.NVL(!姓名))
'                objXML.AppendData "JOB", ""
'                objXML.AppendData "TITLE", ""
                
                objXML.AppendNode strItem, True
            
            Case 3      '药品
                strItem = "CONSIS_BASIC_DRUGSVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!编码))
                objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!通用名))
                objXML.AppendData "TRADE_NAME", SpecialChar(mdlDrugMachine.NVL(!商品名))
                objXML.AppendData "ENGLISHNAME", SpecialChar(mdlDrugMachine.NVL(!英文名))
                objXML.AppendData "PINYIN", mdlDrugMachine.NVL(!通用名简码)
                
                objXML.AppendData "SORTTYPE1", mdlDrugMachine.NVL(!药品类别)
                objXML.AppendData "SORTTYPE2", SpecialChar(mdlDrugMachine.NVL(!药品剂型))
                objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!规格))
                objXML.AppendData "MINSPECS", ""
                
                objXML.AppendData "UNIT", SpecialChar(mdlDrugMachine.NVL(!门诊单位))
                objXML.AppendData "MAXUNIT", SpecialChar(mdlDrugMachine.NVL(!药库单位))
                objXML.AppendData "MINUNIT", SpecialChar(mdlDrugMachine.NVL(!门诊单位))
                objXML.AppendData "DOSAGE", mdlDrugMachine.NVL(!剂量系数)
                objXML.AppendData "DOSAGEUNIT", SpecialChar(mdlDrugMachine.NVL(!剂量单位))
                
                objXML.AppendData "PRICE1", "0"
                objXML.AppendData "CONVERTION1", Round(mdlDrugMachine.NVL(!药库包装, 0) / mdlDrugMachine.NVL(!门诊包装), 2)
                objXML.AppendData "CONVERTION2", mdlDrugMachine.NVL(!门诊包装)
                objXML.AppendData "FIRM_ID", SpecialChar(mdlDrugMachine.NVL(!生产商编码))
                objXML.AppendData "FIRM_NAME", SpecialChar(mdlDrugMachine.NVL(!生产商))
                
                objXML.AppendData "PASSNO", SpecialChar(mdlDrugMachine.NVL(!批准文号))
                objXML.AppendData "BARCODE", "0"
                objXML.AppendData "STORAGECONDITION", ""
                objXML.AppendData "STORAGETYPE", ""
                objXML.AppendData "ALLOWIND", IIf(mdlDrugMachine.NVL(!停用) = "1", "Y", "N")
                
                objXML.AppendNode strItem, True
            
            Case 4      '库存与库位
                strItem = "CONSIS_BASIC_LOCATIONVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!库房id)
                objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!药品编码))
                objXML.AppendData "LOCATIONINFO", SpecialChar(mdlDrugMachine.NVL(!库房货位))
                objXML.AppendData "BATCHID", mdlDrugMachine.NVL(!批次)
                objXML.AppendData "BATCHNO", mdlDrugMachine.NVL(!上次批号)
                
                objXML.AppendData "PRODUCEDATE", mdlDrugMachine.NVL(!上次生产日期), xsDateTime
                objXML.AppendData "DISABLEDDATE", mdlDrugMachine.NVL(!效期), xsDateTime
                objXML.AppendData "DRUG_QUANTITY", mdlDrugMachine.NVL(!实际数量)
                
                objXML.AppendNode strItem, True
                
            Case 5      '发药窗口
                strItem = "CONSIS_BASIC_OPWINIDVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "OPWINID_CODE", SpecialChar(mdlDrugMachine.NVL(!编码))
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!药房id)
                objXML.AppendData "OPWINID_NAME", SpecialChar(mdlDrugMachine.NVL(!名称))
                                
                objXML.AppendNode strItem, True
            
            Case 21     '配药
                
                If Not strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW"  '处方
                    
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!处方时间), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                    objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!库房id)
                    objXML.AppendData "PATIENT_ID", SpecialChar(mdlDrugMachine.NVL(!就诊卡号))
                    objXML.AppendData "PATIENT_NAME", SpecialChar(mdlDrugMachine.NVL(!姓名))
                    
                    objXML.AppendData "INVOICE_NO", ""
                    objXML.AppendData "PATIENT_TYPE", IIf(Val(mdlDrugMachine.NVL(!优先级)) = 1, "01", "00")
                    objXML.AppendData "DATE_OF_BIRTH", mdlDrugMachine.NVL(!出生日期)
                    objXML.AppendData "SEX", SpecialChar(mdlDrugMachine.NVL(!性别))
                    objXML.AppendData "PRESC_IDENTITY", SpecialChar(mdlDrugMachine.NVL(!身份))
                    
                    objXML.AppendData "CHARGE_TYPE", mdlDrugMachine.NVL(!医保类别)
                    objXML.AppendData "PRESC_ATTR", ""
                    objXML.AppendData "PRESC_INFO", ""
                    objXML.AppendData "RCPT_INFO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                    objXML.AppendData "RCPT_REMARK", ""
                    
                    objXML.AppendData "REPETITION", "1"
                    objXML.AppendData "COSTS", "0"
                    objXML.AppendData "PAYMENTS", "0"
                    objXML.AppendData "ORDERED_BY", mdlDrugMachine.NVL(!开嘱科室id)
                    objXML.AppendData "ORDERED_BY_NAME", SpecialChar(mdlDrugMachine.NVL(!开嘱科室))
                    
                    objXML.AppendData "PRESCRIBED_BY", SpecialChar(mdlDrugMachine.NVL(!开嘱医生))
                    objXML.AppendData "ENTERED_BY", SpecialChar(mdlDrugMachine.NVL(!开嘱医生))
                    objXML.AppendData "DISPENSE_PRI", IIf(Val(mdlDrugMachine.NVL(!优先级)) = 1, "1", "2")
                
                    GoTo makDetail
                Else
makDetail:
                    
                    strItem = "CONSIS_PRESC_DTLVW"
                    objXML.AppendNode strItem           '处方明细
                
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                    objXML.AppendData "ITEM_NO", mdlDrugMachine.NVL(!序号)
                    objXML.AppendData "ADVICE_CODE", mdlDrugMachine.NVL(!医嘱id)
                    objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!编码))
                    objXML.AppendData "DRUG_SPEC", SpecialChar(mdlDrugMachine.NVL(!规格))
                    
                    objXML.AppendData "DRUG_NAME", SpecialChar(mdlDrugMachine.NVL(!通用名))
                    objXML.AppendData "FIRM_ID", ""
                    objXML.AppendData "FIRM_NAME", SpecialChar(mdlDrugMachine.NVL(!生产商))
                    objXML.AppendData "PACKAGE_SPEC", SpecialChar(mdlDrugMachine.NVL(!规格))
                    objXML.AppendData "PACKAGE_UNITS", SpecialChar(mdlDrugMachine.NVL(!门诊单位))
                    
                    objXML.AppendData "QUANTITY", mdlDrugMachine.NVL(!门诊单位数量)
                    objXML.AppendData "UNIT", SpecialChar(mdlDrugMachine.NVL(!门诊单位))
                    objXML.AppendData "COSTS", "0"
                    objXML.AppendData "PAYMENTS", "0"
                    objXML.AppendData "DOSAGE", SpecialChar(mdlDrugMachine.NVL(!单次用量))
                    
                    objXML.AppendData "DOSAGE_UNITS", SpecialChar(mdlDrugMachine.NVL(!剂量单位))
                    objXML.AppendData "ADMINISTRATION", SpecialChar(mdlDrugMachine.NVL(!用法))
                    objXML.AppendData "FREQUENCY", SpecialChar(mdlDrugMachine.NVL(!执行频次))
                    objXML.AppendData "ADDITIONUSAGE", ""
                    objXML.AppendData "RCPT_REMARK", ""
                    
                    objXML.AppendNode strItem, True
                End If
                
                strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
                
            Case 22, 23    '22开始发药；23-完成发药
                strItem = "CONSIS_PRESC_MSTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!处方时间), xsDateTime
                objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                objXML.AppendData "INVOICE_NO", ""
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!库房id)
                
                objXML.AppendNode strItem, True
            
            Case 24     '退药
                If Not strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW"  '处方
                    
                    objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!处方时间), xsDateTime
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                    objXML.AppendData "INVOICE_NO", ""
                    objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!库房id)
                
                    GoTo makReturn
                Else
makReturn:
                    strItem = "CONSIS_PRESC_DTLVW"
                    objXML.AppendNode strItem           '处方明细
                    
                    objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                    objXML.AppendData "ADVICE_CODE", mdlDrugMachine.NVL(!医嘱id)
                    objXML.AppendData "DRUG_CODE", SpecialChar(mdlDrugMachine.NVL(!药品编码))
                    objXML.AppendData "QUANTITY", mdlDrugMachine.NVL(!门诊单位数量)
                    
                    objXML.AppendNode strItem, True
                End If
                
                strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id)
                
            Case 25     '整单退药
                strItem = "CONSIS_PRESC_MSTVW"
                objXML.AppendNode strItem
                
                objXML.AppendData "PRESC_DATE", mdlDrugMachine.NVL(!处方时间), xsDateTime
                objXML.AppendData "PRESC_NO", mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据)
                objXML.AppendData "INVOICE_NO", ""
                objXML.AppendData "DISPENSARY", mdlDrugMachine.NVL(!库房id)
                
                objXML.AppendNode strItem, True
            End Select
            
            'XML收尾
            .MoveNext
            
            Select Case intType
            Case 21, 24    '21-配药；24-退药
                If .EOF Then
                    objXML.AppendNode "CONSIS_PRESC_MSTVW", True
                Else
                    If Not strNO = mdlDrugMachine.NVL(!处方号) & "_" & mdlDrugMachine.NVL(!单据) & "_" & mdlDrugMachine.NVL(!库房id) Then
                        objXML.AppendNode "CONSIS_PRESC_MSTVW", True
                    End If
                End If
            End Select
            
        Loop While .EOF = False
    End With
    
    objXML.AppendNode "ROOT", True
    
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "格式化数据完成", 1
    mobjBase.mobjLog.Save
    FormatData = objXML.XmlText
    Exit Function
    
hErr:
    mobjBase.mobjLog.Add objXML.XmlText, 1, 1
    mobjBase.mobjLog.Add "格式化数据失败", 1
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description
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
    Call mdlDrugMachine.GetUserInfo(mobjBase.UserName, mobjBase.mobjComLib, mobjBase.mobjLog, UserInfo)
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property



