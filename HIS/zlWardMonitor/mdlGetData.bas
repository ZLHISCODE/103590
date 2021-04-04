Attribute VB_Name = "mdlGetData"
Option Explicit
Private Declare Function CEC_DevNo2His Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strInPatient As String) As Boolean
'lngType:1监护仪床号, 2HIS床号, 3病历编号

Private Declare Function CEC_UpdateDataBase Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal strResult As String) As Boolean
Private Declare Function CEC_HisSetDataToCec Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal strResult As String) As Boolean
Private Declare Function CEC_GetMonitorData Lib "CecDeviceToHis.dll" (ByVal lngDevice As Long, ByVal lngType As Long, ByVal strResult As String) As Boolean

Public gcnOracle As New ADODB.Connection    '公共数据库连接
Private mobjRichEPR As Object           '病历核心部件
Const SP = "[|]"
Const SPN = "[^]"


Public Function RequestData(ByVal lngDevice As Long, ByVal lngCmd As Long, ByVal obj As Object) As Boolean
'功能：根据监护仪上的操作指令返回所请求的数据
'参数：lngDevice-设备号，lngCmd-请求的命令
    Dim strCmd As String, strdata As String, strResult As String * 20
    Dim strBedNO As String, strInPatient As String, strTmp As String
        
    On Error Resume Next
    strCmd = Hex(lngCmd)
        
    Select Case strCmd
        '请求病人信息
        Case "F0001"
            Call CEC_GetMonitorData(lngDevice, 6, strResult)
            strTmp = Replace(Replace(Trim(strResult), "{", ""), "}", "")
            strBedNO = Split(strTmp, "|")(0)
            strInPatient = Split(strTmp, "|")(1)
                
            strdata = GetPatientInfor(strInPatient)
            strdata = strBedNO & "|" & strdata
            If strdata <> strBedNO & "|" Then Call CEC_UpdateDataBase(lngDevice, 1, strdata)
            
        '请求费用信息
        Case "A0001", "A0002", "A0003", "A0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetFee(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
            
        '请求医嘱信息
        Case "B0001", "B0002", "B0003", "B0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
        
            strdata = GetAdvice(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
    
        '请求病历信息
        Case "C0001", "C0002", "C0003", "C0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetCase(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
        
        '请求报告信息
        Case "D0001", "D0002", "D0003", "D0004"
            Call CEC_DevNo2His(lngDevice, 3, strResult)
            strInPatient = Trim(strResult)
            
            strdata = GetReport(strInPatient, strCmd)
            If strdata <> "" Then Call CEC_HisSetDataToCec(lngDevice, lngCmd, strdata)
        
    End Select

    RequestData = True
End Function

Public Function GetPatientInfor(ByVal strInPatient As String) As String
'功能：获取病人信息
'参数：strInPatient-住院号
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLimit As Long
    
    strSQL = "select Nvl(Zl_Getsysparameter(147), 12) par from dual"    '儿童年龄界定上限
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", strSQL)
    lngLimit = rsTmp!par
     
    strSQL = "Select Nvl(b.出院病床,' ')||'|'||Nvl(d.名称,' ')||'|'||b.住院号||'|'||b.主页ID||'|'||a.姓名||'|'||a.年龄||'|'||a.性别||'| | |'||" & vbNewLine & _
        "to_char(b.入院日期,'yyyy-mm-dd')||'|'||to_char(a.出生日期,'yyyy-mm-dd')||'|'||" & vbNewLine & _
        "Decode(sign(zl_to_number(Substr(a.年龄, 1, Instr(a.年龄, '岁') - 1))-" & lngLimit & "),1,0,1)||'|'||Nvl(b.血型,' ')||'|'||Nvl(c.信息值,' ')||'|'||Nvl(a.身份证号,' ')||'|'||Nvl(b.家庭电话,' ')||'|'||Nvl(b.家庭地址,' ') as Data" & vbNewLine & _
        "From 病人信息 a,病案主页 b,病案主页从表 c,部门表 d" & vbNewLine & _
        "Where a.病人id=b.病人id And a.住院次数=b.主页id And b.病人id=c.病人id(+) And b.主页id=c.主页id(+) And c.信息名(+) = '主治医师' And a.当前科室id=d.id(+) And a.住院号 = [1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
    
    '"监护仪床号|HIS床号|科室|病历号[病案号]|住院次数|病人姓名|年龄|性别|身高|体重|住院日期|出生日期|类型|血型|主治医生|身份证号|电话|住址"
    
    If rsTmp.RecordCount > 0 Then GetPatientInfor = rsTmp!Data
    
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetFee(ByVal strInPatient As String, ByVal strCmd As String) As String
'功能：获取病人费用信息
'参数：strInPatient-住院号
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    Dim strHead As String, curSum As Currency
    
    strSQL = "Select 姓名,性别,年龄,当前床号 From 病人信息 Where 住院号=[1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "病人:" & !姓名 & SP & "性别:" & !性别 & SP & "年龄:" & Replace(!年龄, "岁", "") & SP & _
                 "住院号:" & Val(strInPatient) & SP & "床号:" & !当前床号 & SPN
    End With
    strHead = strHead & "发生时间" & SP & "科室" & SP & "费用项目" & SP & "数次" & SP & "单价" & SP & "实收金额" & SPN
    
    Select Case strCmd
        Case "A0001"
            strIF = "And A.发生时间 Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60"
        Case "A0002"
            strIF = "And A.发生时间 Between Trunc(Sysdate - 1) And Trunc(Sysdate ) - 1 / 24 / 60 / 60"
        Case "A0003"
            strIF = "And A.发生时间 Between Trunc(Sysdate-180, 'mm') And Sysdate"
        Case "A0004"
            strIF = ""
    End Select
    
    strSQL = "Select To_Char(A.发生时间, 'yyyy/mm/dd') 发生时间, C.名称 开单科室, D.名称 收费项目," & vbNewLine & _
            "       Decode(Nvl(A.付数, 1), 1, '', 0, '', A.付数 || ' 付 × ') || A.数次 || ' ' || A.计算单位 As 数量, Ltrim(To_Char(A.标准单价,'9999999990.00000')) as 标准单价, Ltrim(To_Char(Nvl(Sum(A.实收金额),0),'9999999990.00')) as  实收金额" & vbNewLine & _
            "From 病人费用记录 A, 病人信息 B, 部门表 C, 收费项目目录 D" & vbNewLine & _
            "Where A.病人id = B.病人id And A.开单部门id = C.ID And A.收费细目id = D.ID And A.记录状态 > 0 And B.住院号 = [1]" & vbNewLine & _
            "      " & strIF & vbNewLine & _
            "Group By A.NO, Mod(A.记录性质, 10), Nvl(A.价格父号, A.序号), A.记录状态, To_Char(A.发生时间, 'yyyy/mm/dd'), C.名称, D.名称, A.付数, A.数次, A.计算单位, A.标准单价" & vbNewLine & _
            "Order By 发生时间, A.NO"
    
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
    
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            For i = 1 To .RecordCount
                strValue = strValue & vbNewLine & !发生时间 & SP & !开单科室 & SP & !收费项目 & SP & !数量 & SP & !标准单价 & SP & !实收金额 & SPN
                curSum = curSum + !实收金额
                .MoveNext
            Next
        End With
        
        GetFee = strHead & strValue & "合计" & SP & curSum
    Else
        GetFee = strHead
    End If

    Exit Function
errH:
    Debug.Print Err.Description
    Call WriteLog(Err.Description)
End Function

Private Function GetAdvice(ByVal strInPatient As String, ByVal strCmd As String) As String
'功能：获取病人医嘱信息
'参数：strInPatient-住院号
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    Dim strHead As String, strDoctor As String
    
     strSQL = "Select A.姓名, A.性别, A.年龄, A.当前床号, B.信息值 As 主治医师" & vbNewLine & _
            "From 病人信息 A, 病案主页从表 B" & vbNewLine & _
            "Where A.住院号 = [1] And A.病人id = B.病人id(+) And A.住院次数 = B.主页id(+) And B.信息名(+) = '主治医师'"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "病人:" & !姓名 & SP & "性别:" & !性别 & SP & "年龄:" & Replace(!年龄, "岁", "") & SP & _
                 "住院号:" & Val(strInPatient) & SP & "床号:" & !当前床号 & SPN
        '最后两行，要求只有两列
        strDoctor = " [|] [^]主治医生[|]" & IIf(IsNull(!主治医师), "无", !主治医师)
    End With
    
    strHead = strHead & "期效" & SP & "开始时间" & SP & "医嘱内容" & SP & "用法" & SP & "频率" & SPN
    Select Case strCmd
        Case "B0001"
            strIF = " And A.开嘱时间 + 0 Between Trunc(Sysdate) And Trunc(Sysdate + 1) - 1 / 24 / 60 / 60"
        Case "B0002"
            strIF = " And A.医嘱期效 = 0"
        Case "B0003"
            strIF = " And A.医嘱期效 = 1"
    End Select
    
    strSQL = "Select Decode(A.医嘱期效, 0, '长嘱', '临嘱') As 期效, To_Char(A.开始执行时间, 'MM-DD HH24:MI') As 开始时间," & vbNewLine & _
            "       Decode(D.诊疗类别, '5', D.医嘱内容, '6', D.医嘱内容, A.医嘱内容) 医嘱内容," & vbNewLine & _
            "       Decode(A.诊疗类别, 'E', Decode(Instr('2468', Nvl(E.操作类型, '0')), 0, Null, E.名称), Null) As 用法, A.执行频次 As 频率" & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱记录 D, 病人信息 B, 诊疗项目类别 C, 诊疗项目目录 E" & vbNewLine & _
            "Where A.病人id = B.病人id And A.主页id = B.住院次数 And A.诊疗类别 = C.编码(+) And A.医嘱状态 <> -1 And B.住院号 = [1] And A.ID = D.相关id(+) And" & vbNewLine & _
            "      A.相关id Is Null And Instr('5,6', D.诊疗类别(+)) > 0 And A.诊疗项目id = E.ID(+) And A.医嘱状态 Not In (4, 8, 9)" & vbNewLine & strIF

    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
    If rsTmp.RecordCount > 0 Then
        With rsTmp
            For i = 1 To .RecordCount
                strValue = strValue & !期效 & SP & !开始时间 & SP & !医嘱内容 & SP & !用法 & SP & !频率 & SPN
                .MoveNext
            Next
        End With
        GetAdvice = strHead & strValue & strDoctor
    Else
        GetAdvice = strHead & strDoctor
    End If

    Debug.Print strSQL
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetCase(ByVal strInPatient As String, ByVal strCmd As String) As String
'功能：获取病人病历信息
'参数：strInPatient-住院号
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strHead As String, strText As String, i As Long
    Dim objTmp As Object
    
    strSQL = "Select 姓名,性别,年龄,当前床号 From 病人信息 Where 住院号=[1]"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
        
    If rsTmp.RecordCount = 0 Then Exit Function
    With rsTmp
        strHead = "病人:" & !姓名 & SP & "性别:" & !性别 & SP & "年龄:" & Replace(!年龄, "岁", "") & SP & _
                 "住院号:" & Val(strInPatient) & SP & "床号:" & !当前床号 & SPN
    End With
    
    Select Case strCmd
        Case "C0001"
            strIF = " And A.病历名称 = '入院记录'"
        Case "C0002"
            strIF = " And A.病历名称 = '首次病程记录'"
        Case "C0003"
            strIF = " And A.病历名称 = '手术记录'"
        Case "C0004"
            strIF = " And (A.病历名称 = '会诊记录' or A.病历名称 = '请会诊记录')"
    End Select
    
    strSQL = "Select A.ID" & vbNewLine & _
        "From 电子病历记录 A, 病人信息 B, 病历文件目录 C" & vbNewLine & _
        "Where A.病人id = B.病人id And B.住院号 = [1] And A.文件id = C.ID And C.种类 = 2" & strIF & vbNewLine & _
        "Order by 创建时间 Desc"
    
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))
    '有多个病历文件时，只取最近一个
    If rsTmp.RecordCount > 0 Then
        If mobjRichEPR Is Nothing Then
            '不能写在类的初始化中，因为消息触发时，是一个新线程或对象，无法访问类中的对象。
            If mobjRichEPR Is Nothing Then Set mobjRichEPR = CreateObject("zlRichEPR.cRichEPR")
            Call mobjRichEPR.InitRichEPR(gcnOracle, objTmp, 100, True)
        End If
                
        strText = "内容" & SP & mobjRichEPR.GetDocumentText(Val(rsTmp!ID))
        GetCase = strHead & strText
    Else
        GetCase = strHead & "内容" & SP & "无"
    End If
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function


Private Function GetReport(ByVal strInPatient As String, ByVal strCmd As String) As String
'功能：获取病人报告信息
'参数：strInPatient-住院号
    Dim rsTmp As ADODB.Recordset, strSQL As String, strIF As String, strValue As String, i As Long
    
    strSQL = ""
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSQL, "读取病人数据", Val(strInPatient))

    GetReport = ""
    Exit Function
errH:
    Call WriteLog(Err.Description)
End Function




Public Sub WriteLog(ByVal strInfo As String)
    '将调试信息写入文件中
    Dim objFile As Object
    Dim objText As Object
    Dim strFile As String
    
    On Error Resume Next
    Set objFile = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\zlWardMonitor.Log"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, 8) '8-ForAppending
    objText.WriteLine Now()
    objText.WriteLine strInfo
    objText.Close
End Sub


Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax = 0 Or intMax < 200 Then intMax = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax = 0 Or intMax < 200 Then intMax = 200
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If
    'Debug.Print strLog
    cmdData.CommandText = strSQL
    Set OpenSQLRecord = cmdData.Execute
End Function
