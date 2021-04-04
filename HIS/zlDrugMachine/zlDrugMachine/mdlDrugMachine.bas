Attribute VB_Name = "mdlDrugMachine"
Option Explicit

'------------------------------------------------------------------------------
'说明：药品自动化设备接口模块
'编制：余智勇
'------------------------------------------------------------------------------

Public Sub ReadParams(ByRef typVar As TYPE_PARAMS)
'功能：读取参数，并保存到变量
    
    Dim objXML As New clsXML
    Dim strFile As String

    '调试
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\" & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & GSTR_CONFIG_FILE
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, "\" & App.EXEName, "") & "\" & App.EXEName & "\zlDrugMachineManage\zlDrugMachine.cfg"
    Else
        Exit Sub
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        With typVar
            .输出日志 = True
            .详细日志 = False
            .保存日志天数 = 7
        End With
        Exit Sub
    End If

    With typVar
        .输出日志 = Val(GetParameter(objXML, "output", "0")) = 1
        .详细日志 = Val(GetParameter(objXML, "detailed", "0")) = 1
        .保存日志天数 = Val(GetParameter(objXML, "savedays", "7"))
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
End Sub

Public Function VerifyConfigFile(ByVal strFile As String) As Boolean
'功能：检查配置文档是否存在，不存在就自动创建
'参数：
'返回：True检查成功；False检查失败

    Dim fsoFile As New FileSystemObject
    Dim tsmFile As TextStream
    
    On Error GoTo hErr
    
    If fsoFile.FileExists(strFile) = False Then
        '创建配置文档
        Set tsmFile = fsoFile.CreateTextFile(strFile)
        
        '默认生成文档内容
        With tsmFile
            .WriteLine "<root>"
            .WriteLine "    <log>"
            .WriteLine "        <output>0</output>"
            .WriteLine "        <detailed>0</detailed>"
            .WriteLine "        <savedays>7</savedays>"
            .WriteLine "    </log>"
            .WriteLine "    <timer>"
            .WriteLine "        <enabled>0</enabled>"
            .WriteLine "        <businessdata></businessdata>"
            .WriteLine "        <cycle>5</cycle>"
            .WriteLine "        <validdays>2</validdays>"
            .WriteLine "        <viewlines>200</viewlines>"
            .WriteLine "    </timer>"
            .WriteLine "</root>"
        End With
        tsmFile.Close
    End If
    
    VerifyConfigFile = True
    Exit Function
    
hErr:
End Function

Private Function GetParameter(ByVal objXML As clsXML, ByVal strName As String, Optional ByVal strDefaultVal As String) As String
'功能：从zlDrugMachine.cfg文件中获取指定参数的值
'参数：
'  objXML：cfg文件的内容加载后的XML对象
'  strName：参数名称，即：XML结点名称
'返回：参数值

    Dim strValue As String

    If objXML Is Nothing Then
        GetParameter = strDefaultVal
        Exit Function
    End If
    
    strName = LCase(strName)
    
    If objXML.GetSingleNodeValue(strName, strValue) Then
        GetParameter = strValue
    Else
        GetParameter = strDefaultVal
    End If

End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub CreateSOAP(ByRef objSOAP As Object, ByVal objBase As Object)
'创建SOAP部件
        
    On Error Resume Next
    
    Set objSOAP = CreateObject("MSSOAP.SoapClient30")
    If Err.Number <> 0 Then
        Err.Clear
        objBase.mobjLog.Add "创建“SoapClient30”部件失败！", 1
        
        Set objSOAP = CreateObject("MSSOAP.SoapClient")
        If Err.Number <> 0 Then
            Err.Clear
            objBase.mobjLog.Add "程序尝试创建“SoapClient20”部件失败！", 1
        Else
            objBase.mobjLog.Add "创建“SoapClient20”部件完成！", 1
        End If
    Else
        objBase.mobjLog.Add "创建“SoapClient30”部件完成！", 1
    End If
    
    On Error GoTo 0
End Sub

Public Sub CreateHTTP(ByRef objHTTP As Object, ByVal objBase As Object)
    On Error Resume Next
    Set objHTTP = Nothing
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number <> 0 Then
        Err.Clear
        objBase.mobjLog.Add "创建“WinHttp”部件失败，请联系技术人员", 1
    End If
    On Error GoTo 0
End Sub

Public Function TransmitFlag(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, ByVal rsData As ADODB.Recordset, _
    ByVal objPub As Object, ByVal blnFinish As Boolean) As Boolean

'功能：传送后对ZLHIS数据作标志
'参数：
'  intType：业务类型
'  intIO：门诊与住院
'  rsData：记录集对象
'  objPub：公共对象
'  blnFinish：True完成标志；False失败标志
'返回：True成功；Fase失败
    
    Dim strSQL As String, strInfo As String
    Dim lngStockID As Long
    Dim objDB As Object
    
    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    If intAppType = Val("3-支付宝") Then
        Set objDB = objPub.mobjComLib
    Else
        Set objDB = objPub.mobjComLib.zlDatabase
    End If
    
    With rsData
        .MoveFirst
        Do
            If intIO = 1 Then
                strInfo = strInfo & ";" & !单据 & "," & !处方号
            Else
                strInfo = strInfo & ";" & !收发id
            End If
            lngStockID = !库房id
            
            .MoveNext
            
            If .EOF = False Then
                If lngStockID <> !库房id Then
                    GoTo makProc
                End If
            Else
makProc:
                If Left(strInfo, 1) = ";" Then strInfo = Mid(strInfo, 2)
                If intIO = 1 Then
                    strSQL = "ZL_药品收发门诊标志_FLAG(" & _
                        IIf(intType >= 20, intType - 20, intType) & "," & _
                        lngStockID & ",'" & strInfo & "'," & IIf(blnFinish, 1, 0) & ")"
                    objPub.mobjLog.Add strSQL, 2, 1
                    Call objDB.ExecuteProcedure(strSQL, "药品收发门诊标志")
                Else
                    strSQL = "ZL_药品收发住院标志_FLAG(" & _
                        IIf(intType >= 20, intType - 20, intType) & "," & _
                        "'" & strInfo & "'," & IIf(blnFinish, 1, 0) & ")"
                    objPub.mobjLog.Add strSQL, 2, 1
                    Call objDB.ExecuteProcedure(strSQL, "药品收发住院标志")
                    strInfo = ""
                End If
            End If
        Loop While .EOF = False
    End With
    
    objPub.mobjLog.Save
    TransmitFlag = True
    
    Exit Function
    
hErr:
    objPub.mobjLog.Add Err.Number & ":" & Err.Description, 2
    objPub.mobjLog.Save
End Function

Private Function GetWinName(ByVal lngDeptID As Long, ByVal lngWinCode As Long, ByVal objDB As Object, ByVal objLog As Object) As String
'功能：将窗口编码转成窗口名称
'参数：
'  lngDeptID：库房ID
'  lngWinCode：窗口编码
'返回：窗口名称
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = "Select 名称 From 发药窗口 Where 药房id = [1] And 编码 = [2] "
    Set rsTemp = objDB.OpenSQLRecord(strSQL, "将窗口编码转成窗口名称", lngDeptID, CStr(lngWinCode))
    If rsTemp.EOF = False Then
        GetWinName = NVL(rsTemp!名称)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    objLog.Add "将窗口编码转成窗口名称失败", 2
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Save
End Function

Public Function UpdateDispenseWindow(ByVal rsData As ADODB.Recordset, ByVal strWin As String, ByVal objDB As Object, ByVal objLog As Object) As Boolean
'功能：更新数据库的窗口信息
'参数：
'  rsData：数据源
'  strWin：窗口
'返回：True成功；False失败

    Dim lngStockID As Long
    Dim strNO As String, strSQL As String
    Dim intBill As Integer
    Dim strWinName As String

    If rsData.State <> adStateOpen Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    On Error GoTo hErr
    
    With rsData
        .MoveFirst
        Do
            '同库房、同单据、同处方号只能更新一个窗口
            lngStockID = !库房id
            strNO = Trim(!处方号)
            intBill = !单据
            
            '窗口编码转窗口名称
            strWinName = GetWinName(lngStockID, Val(strWin), objDB, objLog)
            
            .MoveNext
            
            If .EOF = False Then
                If Not (lngStockID = !库房id And strNO = Trim(!处方号) And intBill = !单据) Then
                    GoTo makProc
                End If
            Else
makProc:
                strSQL = "Zl_未发药品记录_分配发药窗口(" & _
                         "'" & strNO & "'," & _
                         intBill & "," & _
                         lngStockID & "," & _
                         IIf(Trim(strWinName) = "", "Null", "'" & strWinName & "'") & ")"
                objLog.Add strSQL, 2, 1
                Call objDB.ExecuteProcedure(strSQL, "更新发药窗口")
            End If
        Loop While .EOF = False
    End With
    
    objLog.Save
    UpdateDispenseWindow = True
    Exit Function
      
hErr:
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Save
End Function

Public Function CopyStructure(ByVal fdsSource As ADODB.Fields) As ADODB.Recordset
'功能：
'参数：
'返回：

    Dim i As Integer

    On Error GoTo hErr

    Set CopyStructure = New ADODB.Recordset
    
    With CopyStructure
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        
        '结构复制
        For i = 0 To fdsSource.Count - 1
            .Fields.Append fdsSource(i).Name, IIf(fdsSource(i).Type = adNumeric, adDouble, fdsSource(i).Type), fdsSource(i).DefinedSize, adFldIsNullable
        Next
        
        .Open
    End With
    
    Exit Function

hErr:
    Set CopyStructure = Nothing
End Function

Public Function CopyRecord(ByVal fdsSource As ADODB.Fields, ByRef rsTarget As ADODB.Recordset) As String
'功能：
'参数：
'返回：

    Dim i As Integer
    
    On Error GoTo hErr
    
    rsTarget.AddNew
    For i = 0 To fdsSource.Count - 1
        rsTarget.Fields(i).Value = fdsSource(i).Value
    Next
    rsTarget.Update
    
    Exit Function
    
hErr:
    CopyRecord = Err.Number & ":" & Err.Description
End Function

Public Sub ClearRecord(ByRef rsSource As ADODB.Recordset)
    With rsSource
        .MoveLast
        Do While .BOF = False
            .Delete
            .MovePrevious
        Loop
    End With
End Sub

Public Function IP(Optional ByRef strErr As String) As String
    '功能：通过API获取临时IP
    
    Dim ret As Long, Tel As Long
    Dim bBytes() As Byte
    Dim TempList() As String
    Dim TempIP As String
    Dim Tempi As Long
    Dim Listing As MIB_IPADDRTABLE
    Dim L3 As String
    Dim strTmpErr As String, strALLErr As String
    
    strErr = ""
    On Error GoTo Errhand
    GetIpAddrTable ByVal 0&, ret, True
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
    'retrieve the data
    GetIpAddrTable bBytes(0), ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr, strTmpErr)
        If strTmpErr <> "" Then strALLErr = strALLErr & IIf(strALLErr = "", "", "|") & strTmpErr
    Next Tel
    'Sort Out The IP For WAN
        TempIP = TempList(0)
        For Tempi = 0 To Listing.dEntrys - 1
            L3 = Left(TempList(Tempi), 3)
            If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
                TempIP = TempList(Tempi)
            End If
        Next Tempi
        IP = TempIP 'Return The TempIP
    Exit Function
    strErr = strALLErr
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    strErr = strALLErr & IIf(strALLErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Private Function ConvertAddressToString(longAddr As Long, Optional ByRef strErr As String) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    
    strErr = ""
    On Error GoTo errH
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errH:
    strErr = Err.Description
    Err.Clear
End Function

Public Function GetUserInfo(ByVal strDBUser As String, ByVal objComLib As Object, ByVal objLog As Object, ByRef typUserInfo As TYPE_USER_INFO) As Boolean
'功能：获取当前用户的基本信息
'返回：返回Ado记录集
    Dim strSQL As String, strDefault As String
    Dim rsTemp As ADODB.Recordset
    
    If objComLib Is Nothing Then
        objLog.Add "GetUserInfo的objComLib参数为Nothing，已终止过程执行", 1
        objLog.Save
        Exit Function
    End If
    
    On Error GoTo hErr
    
    objLog.Add "获取用户信息", 1
    objLog.Add "用户名：" & UCase(strDBUser), 2
    
    strDefault = " And C.缺省 = 1"
    strSQL = "Select User,A.Id, A.编号, A.简码, A.姓名, A.专业技术职务,B.用户名, C.部门id, D.编码 As 部门码, D.名称 As 部门名 " & vbNewLine & _
             "From 人员表 A, 上机人员表 B, 部门人员 C, 部门表 D " & vbNewLine & _
             "Where A.Id = B.人员id And A.Id = C.人员id And C.部门id = D.Id And B.用户名 = [1] "
    If TypeName(objComLib) = "clsPublic" Then
        Set rsTemp = objComLib.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    Else
        Set rsTemp = objComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
    End If
    objLog.Add strSQL & strDefault, 2
    
    If rsTemp.RecordCount = 0 Then
        strDefault = " And Rownum < 2"
        Set rsTemp = objComLib.zlDatabase.OpenSQLRecord(strSQL & strDefault, "GetUserInfo", UCase(strDBUser))
        objLog.Add strSQL & strDefault, 2
    End If
    
    If rsTemp.RecordCount > 0 Then
        typUserInfo.ID = rsTemp!ID
        typUserInfo.编号 = rsTemp!编号
        typUserInfo.部门ID = mdlDrugMachine.NVL(rsTemp!部门ID, 0)
        typUserInfo.简码 = mdlDrugMachine.NVL(rsTemp!简码)
        typUserInfo.姓名 = mdlDrugMachine.NVL(rsTemp!姓名)
        typUserInfo.用户名 = rsTemp!用户名
        GetUserInfo = True
        objLog.Add "获取用户信息成功", 1
    Else
        typUserInfo.ID = 0
        typUserInfo.编号 = ""
        typUserInfo.部门ID = 0
        typUserInfo.简码 = ""
        typUserInfo.姓名 = ""
        typUserInfo.用户名 = ""
        objLog.Add "获取用户信息失败", 1
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    objLog.Add "获取用户信息失败", 1
    objLog.Add Err.Number & ":" & Err.Description, 1
    objLog.Save
End Function

Public Function SpecialChar(ByVal strVal As Variant) As String
'功能：特殊字符转换
'说明：
' < 转 &lt;
' > 转 &gt;
' & 转 &amp;
' ' 转 &apos;
' " 转 &quot;
    Dim strReturn As String
    
    If IsNull(strVal) Then
        strVal = ""
        GoTo errHandle
    End If
    If strVal = "" Then
        GoTo errHandle
    End If
    On Error GoTo errHandle
    strReturn = strVal
    strReturn = Replace(strReturn, "<", "&lt;")
    strReturn = Replace(strReturn, ">", "&gt;")
    strReturn = Replace(strReturn, "&", "&amp;")
    strReturn = Replace(strReturn, "'", "&apos;")
    strReturn = Replace(strReturn, """", "&quot;")
    SpecialChar = strReturn
    Exit Function
    
errHandle:
    SpecialChar = strVal
End Function

Public Function GetInterfaceLink(ByVal objPublic As Object, ByVal strCode As String) As String
'功能：获取接口连接信息
'参数：
'  objPublic：公共对象
'  strCode：接口编号
'返回：解密后的连接串

    Dim strSQL As String, strLink As String
    Dim rsTemp As ADODB.Recordset
    Dim objEncrypt As Object
    
    On Error Resume Next
    Set objEncrypt = CreateObject("zlEncryptPub.clsEncrypt")
    If Err.Number <> 0 Then
        objPublic.mobjLog.Add "zlEncryptPub部件未注册，影响接口连接信息的解密", 1
    End If
    Err.Clear
    
    On Error GoTo hErr
    
    strSQL = "Select 连接信息 From 药品设备接口 Where 编号 = [1] And 停用日期 Is Null And 启用日期 Is Not Null "
    If LCase(TypeName(objPublic.mobjComLib)) = "clspublic" Then
        Set rsTemp = objPublic.mobjComLib.OpenSQLRecord(strSQL, "获取接口连接信息", strCode)
    Else
        Set rsTemp = objPublic.mobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取接口连接信息", strCode)
    End If
    If rsTemp.EOF = False Then
        If Not IsNull(rsTemp!连接信息) Then
            strLink = objEncrypt.Base64Decode(rsTemp!连接信息)
        End If
    End If
    rsTemp.Close
    
    objPublic.mobjLog.Add "获取接口连接信息完成", 1
    objPublic.mobjLog.Save
    
    GetInterfaceLink = strLink
    Exit Function
    
hErr:
    objPublic.mobjLog.Add Err.Number & ":" & Err.Description, 1
    objPublic.mobjLog.Save
End Function

Public Sub ExecuteProcedureBeach(ByVal cllProcs As Variant, ByVal strCaption As String, ByVal cnThird As ADODB.Connection, _
    ByRef objLog As Object, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True)
'---------------------------------------------------------------------------------------------
'功能:执行相关的Oracle过程集
'参数:cllProcs-oracle过程集，可以为数组，也可以为集合，不能为其他类型
'     strCaption -执行过程的父窗口标题
'     blnTrans-是否存在事务
'     blnCommit-执行完过程后,提交数据(前题:blnTrans=true)
'---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo hErr
    
    If blnTrans Then cnThird.BeginTrans
    
    If TypeName(cllProcs) = "Collection" Then '集合形式
        For i = 1 To cllProcs.Count
            strSQL = cllProcs(i)
            objLog.Add strSQL, 2, 1
            Call ExecuteProcedure(cnThird, strSQL, strCaption)
        Next
    ElseIf Not IsObject(cllProcs) Then
        If VarType(cllProcs) = vbArray + vbVariant Or VarType(cllProcs) = vbArray + vbString Then  '数组形式
            For i = LBound(cllProcs) To UBound(cllProcs)
                strSQL = cllProcs(i)
                objLog.Add strSQL, 2, 1
                Call ExecuteProcedure(cnThird, strSQL, strCaption)
            Next
        End If
    End If
    
    If blnCommit And blnTrans Then
        cnThird.CommitTrans
    End If
    objLog.Save
    Exit Sub
    
hErr:
    If blnCommit And blnTrans Then
        cnThird.RollbackTrans
    End If
    objLog.Add strSQL, 2, 1
    objLog.Add Err.Number & ":" & Err.Description, 2
    objLog.Add "ExecuteProcedureBeach()", 2
    objLog.Save
End Sub

Public Sub ExecuteProcedure(ByVal cnThird As ADODB.Connection, ByRef strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = Now()
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        Set cmdData.ActiveConnection = cnThird      '这句比较慢
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
        Call cmdData.Execute
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
    
NoneVarLine:
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    cnThird.Execute strSQL, , adCmdText
End Sub

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

