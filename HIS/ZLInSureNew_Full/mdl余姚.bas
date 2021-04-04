Attribute VB_Name = "mdl余姚"
Option Explicit
    
'Public Declare Function f_Init Lib "dhpDLL.DLL" () As Integer
'Public Declare Function f_Close Lib "dhpDLL.DLL" () As Integer
'Public Declare Function f_Apply Lib "dhpDLL.DLL" (ByVal lngTradeTypeID As Long, _
'    ByVal dblTradeID As Double, ByVal strData As String, ByRef strMessage As String) As Integer
'cd 50223
'VVVV
Private Declare Function f_UserBargaingInit Lib "BargaingApply.DLL" (ByVal strData As String, ByVal strMessage As String) As Integer
Private Declare Function f_UserBargaingClose Lib "BargaingApply.DLL" (ByVal strData As String, ByVal strMessage As String) As Integer
Private Declare Function f_UserBargaingApply Lib "BargaingApply.DLL" (ByVal lngTradeTypeID As Integer, ByVal dblTradeID As Double, ByVal strData As String, ByVal strMessage As String) As Integer

Public gstrOutput余姚 As String, gstrInput余姚 As String, gcn余姚 As New ADODB.Connection, gstrIC明文 As String
Private mstrBillNo As String, mcur非医保 As Currency, mstr流水号 As String
Private mblnInit As Boolean

Public Function makeBillNO(lng病人ID As Long) As String
    Dim datCurr As Date
    datCurr = zlDatabase.Currentdate
    makeBillNO = toHex(CDbl(Format(datCurr, "yyyyMMddHHmmss") & lng病人ID), 36)
End Function

Public Function makeICInfo(lng病人ID As Long) As String
    Dim rsTemp As New ADODB.Recordset
    '生成IC明文
    gstrSQL = "Select A.卡号,B.姓名,B.性别,A.单位编码 From 保险帐户 A,病人信息 B Where A.病人ID=[1]" & _
        " And A.险类=[2] And A.病人ID=B.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_余姚)
    If rsTemp.EOF Then
        MsgBox "没有找到该病人的身份信息", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    makeICInfo = Space(18 - LenB(StrConv(rsTemp(0), vbFromUnicode))) & rsTemp(0) & _
                 String(18, "0") & _
                 Space(20 - LenB(StrConv(rsTemp(1), vbFromUnicode))) & rsTemp(1) & _
                 Space(2 - LenB(StrConv(rsTemp(2), vbFromUnicode))) & rsTemp(2) & _
                 String(56, "0") & _
                 Space(10 - LenB(StrConv(rsTemp(3), vbFromUnicode))) & rsTemp(3) & _
                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
'    makeICInfo = Right(Space(18) & rsTemp(0), 18) & _
'                 String(18, "0") & _
'                 RightB(Space(20) & rsTemp(1), 20) & _
'                 RightB(Space(2) & rsTemp(2), 2) & _
'                 String(56, "0") & _
'                 Right(Space(10) & rsTemp(3), 10) & _
'                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
               'cd 050223  Right(Space(20) & rsTemp(1), 20) & _
               'cd 050223  Right(Space(2) & rsTemp(2), 2) & _

End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = IIf(dblMod <> 0, dblMod, "") & strTemp
    End If
    toHex = strTemp
End Function

Public Function CheckReturn_余姚() As Boolean
    If glngReturn < 0 Then
        If Split(gstrOutput余姚, "$$")(1) = "" Then
            MsgBox "进行医保调用时发生错误", vbInformation, gstrSysName
        Else
            MsgBox "医保操作返回以下错误：" & vbCrLf & "    " & Split(gstrOutput余姚, "$$")(1), vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CheckReturn_余姚 = True
End Function

Public Function 申请交易流水_余姚(str交易类型 As String) As String
    Dim strTemp As String
    申请交易流水_余姚 = ""
    strTemp = "$$" & str交易类型 & "$$"
    gstrOutput余姚 = Space(4000) 'cd 050223
    glngReturn = f_UserBargaingApply(23, 0, strTemp, gstrOutput余姚)
    If CheckReturn_余姚() = False Then Exit Function
    申请交易流水_余姚 = CDbl(Split(gstrOutput余姚, "$$")(2))
End Function

Public Function openConn余姚() As Boolean
    Dim rsTemp As New ADODB.Recordset, strServer As String, strUser As String, strPass As String, _
        strTemp As String, strDatabase As String
    On Error GoTo ErrH
    If gcn余姚.State <> adStateOpen Then
        '首先读出参数，打开连接
        gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚)
        Do Until rsTemp.EOF
            strTemp = Nvl(rsTemp("参数值"), "")
            Select Case rsTemp("参数名")
                Case "余姚服务器"
                    strServer = strTemp
                Case "余姚用户名"
                    strUser = strTemp
                Case "余姚用户密码"
                    strPass = strTemp
                Case "余姚数据库"
                    strDatabase = strTemp
            End Select
            rsTemp.MoveNext
        Loop
    
        On Error Resume Next
        gcn余姚.ConnectionString = "Provider=SQLOLEDB.1;Initial Catalog=" & strDatabase & ";Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
        gcn余姚.CursorLocation = adUseClient
        gcn余姚.Open
        If Err.Number <> 0 Then
            MsgBox "医保前置服务器连接失败。", vbInformation, gstrSysName
            openConn余姚 = False
            Exit Function
        End If
        On Error GoTo ErrH
    End If
    openConn余姚 = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    openConn余姚 = False
End Function

Public Function 医保初始化_余姚() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    If openConn余姚() = False Then
        医保初始化_余姚 = False
        Exit Function
    End If
    
    If mblnInit = False Then
        gstrInput余姚 = "$$$$"
        gstrOutput余姚 = Space(4000)
        
        glngReturn = f_UserBargaingInit(gstrInput余姚, gstrOutput余姚)
        
        医保初始化_余姚 = CheckReturn_余姚()
        If 医保初始化_余姚 Then mblnInit = True
    Else
        医保初始化_余姚 = True
    End If
'cd 050223
'    gstrInput余姚 = "$$$$": gstrOutput余姚 = "$$$$$$"
''    glngReturn = f_UserBargaingInit(gstrInput余姚, gstrOutput余姚)
'    医保初始化_余姚 = CheckReturn_余姚()
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    医保初始化_余姚 = False
End Function

Public Function 医保终止_余姚() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    On Error GoTo errHandle
    Set gcn余姚 = Nothing
    'cd 050223 VVVVVV
    gstrInput余姚 = "$$$$"
    gstrOutput余姚 = Space(4000)

    glngReturn = f_UserBargaingClose(gstrInput余姚, gstrOutput余姚)
    '^^^^^^^^^^^^^^^
'    glngReturn = f_Close()
    医保终止_余姚 = CheckReturn_余姚()
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    医保终止_余姚 = False
End Function

Public Function 医保设置_余姚() As Boolean
    医保设置_余姚 = frmSet余姚.参数设置()
End Function

Public Function 门诊虚拟结算_余姚(rs费用明细 As Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str医保号 As String, lng病人ID As Long, datCurr As Date, rsTemp As New ADODB.Recordset, str项目类型 As String, _
        str病种ID As String, str病种 As String, strSQL As String, strTemp As String, iLoop As Long, str处方明细流水号 As String, _
        str医保类型 As String, str明细编码 As String, str项目名称 As String, str规格 As String, str自付比例 As String
    WriteInfo vbCrLf & "门诊预结算"
    On Error GoTo errHandle
    If rs费用明细.RecordCount = 0 Then
        MsgBox "没有病人费用明细，不能进行医保操作", vbInformation, gstrSysName
        Exit Function
    End If
    rs费用明细.MoveFirst
    lng病人ID = rs费用明细!病人ID
    datCurr = zlDatabase.Currentdate
    
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_余姚

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        str病种ID = rsTemp!编码
        str病种 = rsTemp!名称
    Else
        门诊虚拟结算_余姚 = False
        Exit Function
    End If
    
    mstrBillNo = makeBillNO(lng病人ID)
    gstrSQL = "Select * From 保险帐户 Where 险类=" & TYPE_余姚 & " And 病人ID=" & lng病人ID
'    Call OpenRecordset(rsTemp, gstrSysName)
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    
    If rsTemp.EOF Then
        MsgBox "没有找到该病人的医保信息", vbInformation, gstrSysName
        Exit Function
    End If
    str医保号 = rsTemp!卡号
    mstr流水号 = 申请交易流水_余姚(29)
    If mstr流水号 = "" Then Exit Function
    '写处方表
    strSQL = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr流水号 & "','" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & _
        "',1," & Trim(gstr医院编码) & ",'" & str医保号 & "','" & lng病人ID & "',Null,Null,'" & rs费用明细!开单人 & _
        "','" & str病种 & "','" & str病种ID & "',Null,'" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & _
        "','" & UserInfo.姓名 & "')"
    WriteInfo "写前置机处方数据:" & strSQL
    gcn余姚.Execute strSQL
    mcur非医保 = 0
    iLoop = 1
    strSQL = "select max(SerialNum) as SerialNum From "
    strSQL = strSQL & "(Select SerialNum From hi_upClinicPrescription  "
    strSQL = strSQL & " union all"
    strSQL = strSQL & " select SerialNum From hi_ClinicPrescription ) A"
    Set rsTemp = gcn余姚.Execute(strSQL)
    If rsTemp.EOF Then
        str处方明细流水号 = 0
    Else
        str处方明细流水号 = Nvl(rsTemp(0), 0)
    End If
    str处方明细流水号 = AddNum(str处方明细流水号)
    
    While Not rs费用明细.EOF
        '取收费明细
        gstrSQL = "Select 编码,名称,类别,nvl(规格,'') as 规格 From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs费用明细!收费细目ID))
        str明细编码 = rsTemp!编码: str项目名称 = rsTemp!名称
        str规格 = Left(Left(rsTemp!规格 & " |", InStr(rsTemp!规格 & " |", "|") - 1), InStr(rsTemp!规格 & " |", " ") - 1)
        '判断项目类型
        str项目类型 = IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6", "药品", IIf(rsTemp!类别 = "7", "中药", "诊疗"))
        ''''cd 2005 0301
        '卫材在药品中，但类型为诊疗，所以不能直接按rstemp!类别来判断
        gstrSQL = "select * from 药品目录 where 药品ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs费用明细!收费细目ID))
        
        If Not rsTemp.EOF Then '如果在药品目录中有此项目，再根据药品目录的附注来判断是否为卫材
           If Nvl(rsTemp!标识码, "药品") = "材料" Then str项目类型 = "诊疗"
        End If
        
        '^^^^^^^^^^^^^^^

        
        '从保险支付项目中查找是否有该医保项目
        gstrSQL = "Select 项目编码,项目名称 From 保险支付项目 Where 是否医保=1 And 险类=[1] And 收费细目ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, CLng(rs费用明细!收费细目ID))
        If rsTemp.EOF Then      '没有项目处理
           
            If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                str医保类型 = "丙类": str自付比例 = "1"
            Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                str医保类型 = "甲类": str自付比例 = "0"
            End If
        Else            '有该项目时处理
            str明细编码 = rsTemp!项目编码
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str明细编码 & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str明细编码 & "'"
            
            Set rsTemp = gcn余姚.Execute(gstrSQL)
            
            If rsTemp.EOF Then      '如果医保中心目录中未找到该项目
                If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                    str医保类型 = "丙类": str自付比例 = "1"
                Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                    str医保类型 = "甲类": str自付比例 = "0"
                End If
            Else        '如果医保中心目录中有该药品
                str医保类型 = IIf(rsTemp!zfbl = 0, "甲类", IIf(rsTemp!zfbl = 1, "丙类", "乙类"))
                str自付比例 = rsTemp!zfbl
            End If
        End If
        strSQL = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
            "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & str处方明细流水号 & "," & Trim(gstr医院编码) & ",'" & mstr流水号 & _
            "','" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "','" & mstr流水号 & "'," & IIf(str项目类型 = "诊疗", 2, 1) & ",'" & _
            str明细编码 & "','" & str项目名称 & "','" & str规格 & "'," & Format(rs费用明细!实收金额 / rs费用明细!数量, "#.###") & "," & _
            rs费用明细!数量 & "," & str自付比例 & ",'" & UserInfo.姓名 & "')"
        WriteInfo "传递明细(写处方明细):" & strSQL
        gcn余姚.Execute strSQL
        str处方明细流水号 = AddNum(str处方明细流水号)
        
        'gstrSQL = "ZL_保险支付项目_Modify(" & rs费用明细!收费细目ID & "," & TYPE_余姚 & ",NULL,'" & str明细编码 & "','" & _
            str项目名称 & "','" & str医保类型 & "',1)"
'        WriteInfo "修改保险支付项目:" & gstrSQL
'        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        

'
        If str自付比例 = 1 Then mcur非医保 = mcur非医保 + rs费用明细!实收金额
        rs费用明细.MoveNext
    Wend
    WriteInfo " "
    gstrInput余姚 = "$$" & mcur非医保 & "~1~" & mstr流水号 & "~" & gstrIC明文 & "~0000$$"
    gstrOutput余姚 = Space(4000)
    WriteInfo "预结算调用:f_UserBargaingApply(29, " & mstr流水号 & ", """ & Replace(gstrInput余姚, String(1053, "0"), "") & """, "" "")"
    glngReturn = f_UserBargaingApply(29, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "预结算返回:" & gstrOutput余姚
    门诊虚拟结算_余姚 = CheckReturn_余姚()
    
    WriteInfo "完成预结算"
    门诊虚拟结算_余姚 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    门诊虚拟结算_余姚 = False
End Function

Public Function 门诊结算_余姚(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rsTemp As New ADODB.Recordset, datCurr As Date, cur费用 As Currency, rs记录 As New ADODB.Recordset
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, _
        cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, strTemp As String, _
        str项目类型 As String, cur自付比例 As Currency, str明细编码 As String
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From 门诊费用记录 Where 门诊标志=1 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng病人ID = rsTemp!病人ID
    While Not rsTemp.EOF
        cur费用 = cur费用 + rsTemp!实收金额
        'cd 050301 同时更新进入统筹金额
        
        str项目类型 = IIf(rsTemp!收费类别 = "5" Or rsTemp!收费类别 = "6", "药品", IIf(rsTemp!收费类别 = "7", "中药", "诊疗"))
        ''''cd 2005 0301
        '卫材在药品中，但类型为诊疗，所以不能直接按rstemp!类别来判断
        gstrSQL = "select * from 药品目录 where 药品ID=[1]"
        Set rs记录 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!收费细目ID))
        
        If Not rs记录.EOF Then '如果在药品目录中有此项目，再根据药品目录的附注来判断是否为卫材
           If Nvl(rs记录!标识码, "药品") = "材料" Then str项目类型 = "诊疗"
        End If

        
        '从保险支付项目中查找是否有该医保项目
        gstrSQL = "Select 项目编码,项目名称 From 保险支付项目 Where 是否医保=1 And 险类=[1] And 收费细目ID=[2]"
        Set rs记录 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, CLng(rsTemp!收费细目ID))
        If rs记录.EOF Then      '没有项目处理
           
            If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                cur自付比例 = "1"
            Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                cur自付比例 = "0"
            End If
        Else            '有该项目时处理
            str明细编码 = rs记录!项目编码
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str明细编码 & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str明细编码 & "'"
            
            Set rs记录 = gcn余姚.Execute(gstrSQL)
            
            If rs记录.EOF Then      '如果医保中心目录中未找到该项目
                If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                    cur自付比例 = "1"
                Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                    cur自付比例 = "0"
                End If
            Else        '如果医保中心目录中有该药品
                'str医保类型 = IIf(rs记录!zfbl = 0, "甲类", IIf(rs记录!zfbl = 1, "丙类", "乙类"))
                cur自付比例 = rs记录!zfbl
            End If
        End If
                '在费用记录中记录进入统筹金额
                '项目编码中保存项目类型（药品，诊疗）,摘要中保存自付比例,可根据比例得到甲类，乙类
        gstrSQL = "ZL_病人费用记录_更新医保(" & rsTemp!ID & ","
        gstrSQL = gstrSQL & rsTemp!实收金额 - rsTemp!实收金额 * cur自付比例 & ",NULL,1,'" & str项目类型 & "',NULL,'" & cur自付比例 & "')"
        WriteInfo "修改明细项目的进入统筹金额:" & gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '^^^^^^^^^^
        rsTemp.MoveNext
    Wend
    
    ''cd 050225 预结算时未填写科室，在此补填VVVVV
    gstrSQL = "select 编码,名称 from 部门表 where ID=(Select 开单部门ID From 门诊费用记录 Where 门诊标志=1 and 结帐ID=[1] And Rownum < 2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If Not rsTemp.EOF Then
        gstrSQL = "Update HI_ClinicRx Set Department='" & rsTemp!名称 & "',DepartmentID='" & rsTemp!编码 & "' Where BillID=" & mstr流水号
        WriteInfo "更新科室名称：" & gstrSQL
        gcn余姚.Execute gstrSQL
    End If
    
    
    ''^^^^^^^^^^^^^^^^^^
'    gstrOutput余姚 = Space(4000)
'
'    gstrInput余姚 = "$$1~" & cur费用 & "~1~" & mstr流水号 & "~" & gstrIC明文 & "$$"
'    WriteInfo vbCrLf & "结算调用:f_UserBargaingApply(30, " & mstr流水号 & ", """ & Replace(gstrInput余姚, String(1053, "0"), "") & """, "" "")"
'    glngReturn = f_UserBargaingApply(30, mstr流水号, gstrInput余姚, gstrOutput余姚)
'    WriteInfo "结算返回:" & gstrOutput余姚
'    门诊结算_余姚 = CheckReturn_余姚()
'    If 门诊结算_余姚 = False Then
'        Exit Function
'    End If
'    strTemp = Split(gstrOutput余姚, "$$")(2)
'    cur费用 = CCur(Split(strTemp, "~")(0))
'
'    gcn余姚.Execute "Delete hi_ClinicRx Where BillID='" & mstr流水号 & "'"
'    gcn余姚.Execute "Delete hi_ClinicPrescription Where BillID='" & mstr流水号 & "'"
    
    Call Get帐户信息(TYPE_余姚, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_余姚 & "," & _
            lng病人ID & "," & Year(datCurr) & ",0,0,0,0," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur费用 & "," & cur全自付 & ",0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & mstr流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    '---------------------------------------------------------------------------------------------
    门诊结算_余姚 = True
    WriteInfo "结算完成"
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 门诊结算冲销_余姚(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, str就诊编号 As String, rs记录 As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, lng冲销ID As Long, strTemp As String
    Dim datCurr As Date, strSQL As String
    Dim cur退费费用 As Currency
    Dim cur退费自费 As Currency
    cur退费费用 = 0
    cur退费自费 = 0
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrIC明文 = makeICInfo(lng病人ID)
    If gstrIC明文 = "" Then Exit Function
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    
    '取原单据交易流水号
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    If IsNull(rsTemp!备注) Then
        Err.Raise 9000, gstrSysName, "该单据的交易流水号丢失，不能作废。"
        Exit Function
    End If
    strTemp = rsTemp!发生费用金额
    str就诊编号 = rsTemp!备注
'    strSql = "Insert Into hi_ClinicRx (BillID,DateDiagnose,ChargeType,HospitalID,PIN,ClinicSerial,Department,DepartmentID," & _
'        "Doctor,Disease,DiseaseID,Description,DateOccur,Operator) values ('" & mstr流水号 & "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "',1," & Trim(gstr医院编码) & ",'" & str医保号 & "','" & lng病人id & "',Null,Null,'" & rs费用明细!开单人 & _
'        "','" & str病种 & "','" & str病种ID & "',Null,'" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & _
'        "','" & UserInfo.姓名 & "')"
'    strSql = "Insert Into hi_ClinicPrescription (SerialNum,HospitalID,BillID,DateDiagnose,RecipeSerial,Class,ItemID,ItemName," & _
'        "Specification,Price,Dosage,ScaleSelf,Operator) Values (" & iLoop + lng流水 & "," & Trim(gstr医院编码) & ",'" & mstrBillNo & _
'        "','" & Format(datCurr, "yyyy-mm-dd HH:MM:SS") & "','" & mstr流水号 & "'," & IIf(str项目类型 = "诊疗", 2, 1) & ",'" & _
'        str明细编码 & "','" & str项目名称 & "','" & str规格 & "'," & Format(rs费用明细!实收金额 / rs费用明细!数量, "#.###") & "," & _
'        rs费用明细!数量 & "," & str自付比例 & ",'" & UserInfo.姓名 & "')"
'    strSql = "Select * From hi_ClinicRx Where BillID='" & str就诊编号 & "'"
    strSQL = "Select * From hi_upClinicRx Where BillID='" & str就诊编号 & "'"
    Set rsTemp = gcn余姚.Execute(strSQL)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "前置机中未找到该单据数据，已上传的数据不能退费"
        门诊结算冲销_余姚 = False
        Exit Function
    End If
    
'    strSql = "Select * From hi_ClinicPrescription Where RecipeSerial='" & str就诊编号 & "'"
'    Set rsTemp = gcn余姚.Execute(strSql)
'    If rsTemp.EOF Then
'        Err.Raise 9000,gstrSysName, "前置机中未找到该单据数据，已上传的数据不能退费"
'        门诊结算冲销_余姚 = False
'        Exit Function
'    End If
'    gcn余姚.Execute "Delete hi_ClinicRx Where BillID='" & str就诊编号 & "'"
'    gcn余姚.Execute "Delete hi_ClinicPrescription Where RecipeSerial='" & str就诊编号 & "'"
    
    mstr流水号 = 申请交易流水_余姚(31)

    '调用接口数冲销
    gstrInput余姚 = "$$" & str就诊编号 & "~" & gstrIC明文 & "$$"
    gstrOutput余姚 = Space(4000)
    WriteInfo "退费调用：f_UserBargaingApply(31, CDbl(" & mstr流水号 & "), " & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    glngReturn = f_UserBargaingApply(31, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "退费返回：" & gstrOutput余姚
    门诊结算冲销_余姚 = CheckReturn_余姚()
    If 门诊结算冲销_余姚 = False Then
        Exit Function
    End If
    
    strTemp = Split(gstrOutput余姚, "$$")(2)
    cur退费费用 = CCur(Split(strTemp, "~")(0))
    cur退费自费 = CCur(Split(strTemp, "~")(1))
    '帐户年度信息
    Call Get帐户信息(TYPE_余姚, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)

    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_余姚 & "," & lng病人ID & "," & _
        Year(datCurr) & ",0,0,0,0," & int住院次数累计 & ",0,0,0," & cur退费费用 & "," & cur退费自费 & ",0,0," & _
        "0,0,0,0,NULL,NULL,NULL,'" & mstr流水号 & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    门诊结算冲销_余姚 = True
    WriteInfo "退费完成"

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
    WriteInfo "错误信息：" & Err.Description & vbCrLf & "接口返回：" & gstrOutput余姚
End Function

Public Function 入院登记_余姚(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date, strTemp As String
    Dim lng病种ID As Long
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
'    If rsTmp.BOF Then 入院登记_余姚 = False: Exit Function
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_余姚
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
        str病种 = rsTemp!名称
        str病种编码 = rsTemp!ID
    Else
        入院登记_余姚 = False
        Exit Function
    End If
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码,D.编码 As 科室编码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [2] And A.病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    
    mstr流水号 = 申请交易流水_余姚(32)
'    gstrIC明文 = makeICInfo(lng病人ID)
'                                             '????? 住院号，用流水号代替行不行
'    gstrInput余姚 = "$$" & gstrIC明文 & "~" & mstr流水号 & "~" & _
'        Format(Nvl(rsTemp(0), datCurr), "YYYY.M.D") & "~" & Nvl(rsTemp(4), "医生") & "~" & strInNote & "~" & _
'        str病种编码 & "~" & Nvl(rsTemp!住院科室, " ") & "~" & Nvl(rsTemp!科室编码, "0") & "~" & Nvl(rsTemp!入院病床, "1") & "$$"
'    gstrOutput余姚 = Space(4000)
'
'    WriteInfo "入院登记：f_UserBargaingApply(32, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
'
'    glngReturn = f_UserBargaingApply(32, mstr流水号, gstrInput余姚, gstrOutput余姚)
'    WriteInfo "反回：" & gstrOutput余姚
'    入院登记_余姚 = CheckReturn_余姚()
'    If 入院登记_余姚 = False Then
'        Exit Function
'    End If
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_余姚 & ",'病种ID'," & lng病种ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_余姚 & ",'顺序号'," & mstr流水号 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_余姚 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    入院登记_余姚 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_余姚 = False
End Function

'VVVVVV
Public Function 入院登记撤销_余姚(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim str住院号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '取病人的住院号
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = rsTemp!顺序号 ' Format(Val(rsTemp!顺序号), "0" & String(16, "#")) ' Val(rsTemp!顺序号)

    '入口参数 (Data)
    '交易体： 住院登记交易号+~+要注销的住院号+~+IC暗文信息

    mstr流水号 = 申请交易流水_余姚(40)
    gstrIC明文 = makeICInfo(lng病人ID)
    
    '调用接口
    gstrInput余姚 = "$$" & str住院号 & "~" & str住院号 & "~" & gstrIC明文 & "$$"
    gstrOutput余姚 = Space(4000)
    
    WriteInfo "入院登记：f_UserBargaingApply(40, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    glngReturn = f_UserBargaingApply(40, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "反回：" & gstrOutput余姚
    
    入院登记撤销_余姚 = CheckReturn_余姚()
    If 入院登记撤销_余姚 = False Then
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_余姚 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")

    入院登记撤销_余姚 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'^^^^^^

Public Function 记帐传输_余姚(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, lng主页ID As Long, iLoop As Long, strSQL As String, lng流水 As String, _
        rs明细 As New ADODB.Recordset, strTemp As String, str住院号 As String, str明细编码 As String, str项目名称 As String, _
        str规格 As String, str项目类型 As String, str医保类型 As String, str自付比例 As String
    On Error GoTo errHandle
    '取病人最大主页ID
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    lng主页ID = rsTemp(0)
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = rsTemp!顺序号 ' Format(Val(rsTemp!顺序号), "0" & String(16, "#")) ' Val(rsTemp!顺序号)
    
    '取病人费用记录
    If str单据号 <> "" Then
        gstrSQL = "Select * From 住院费用记录 Where 实收金额<>0 And 实收金额 Is Not Null And 记录状态<>0" & _
            "And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 记录性质=[1] and NO=[2] order by 主页ID,序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", int性质, str单据号)
    Else
        gstrSQL = "Select * From 住院费用记录 Where 实收金额<>0 And 实收金额 Is Not Null And 记录状态<>0 " & _
            " And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 病人id=[1] And 主页id=[2] order by 主页ID,序号"
        Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "", lng病人ID, lng主页ID)
    End If
    
    mstr流水号 = 申请交易流水_余姚(33)
    iLoop = 1
    strSQL = "select max(SerialNum) as SerialNum From "
    strSQL = strSQL & "(Select SerialNum From hi_upInpatientPrescription  "
    strSQL = strSQL & " union all"
    strSQL = strSQL & " select SerialNum From hi_InpatientPrescription ) A"
    
    Set rsTemp = gcn余姚.Execute(strSQL)
    If rsTemp.EOF Then
        lng流水 = 0
    Else
        lng流水 = Nvl(rsTemp(0), 0)
    End If
    While Not rs明细.EOF
        gstrSQL = "Select 编码,名称,类别,nvl(规格,'') as 规格 From 收费细目 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        str明细编码 = rsTemp!编码: str项目名称 = rsTemp!名称
        str规格 = Left(Left(rsTemp!规格 & " |", InStr(rsTemp!规格 & " |", "|") - 1), InStr(rsTemp!规格 & " |", " ") - 1)
        '判断项目类型
        str项目类型 = IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6", "药品", IIf(rsTemp!类别 = "7", "中药", "诊疗"))
        ''''cd 2005 0301
        '卫材在药品中，但类型为诊疗，所以不能直接按rstemp!类别来判断
        gstrSQL = "select * from 药品目录 where 药品ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs明细!收费细目ID))
        
        If Not rsTemp.EOF Then '如果在药品目录中有此项目，再根据药品目录的附注来判断是否为卫材
           If Nvl(rsTemp!标识码, "药品") = "材料" Then str项目类型 = "诊疗"
        End If
        
        '^^^^^^^^^^^^^^^
        '从保险支付项目中查找是否有该医保项目
        gstrSQL = "Select 项目编码,项目名称 From 保险支付项目 Where 是否医保=1 And 险类=[1] And 收费细目ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, CLng(rs明细!收费细目ID))
        If rsTemp.EOF Then      '没有项目处理
            If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                str医保类型 = "丙类": str自付比例 = "1"
            Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                str医保类型 = "甲类": str自付比例 = "0"
            End If
        Else            '有该项目时处理
            str明细编码 = rsTemp!项目编码
            gstrSQL = "Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & str明细编码 & "'"
            gstrSQL = gstrSQL & " union all Select MedicineID,zfbl From hi_Medicine Where MedicineID='" & str明细编码 & "'"
            
            Set rsTemp = gcn余姚.Execute(gstrSQL)
            
            If rsTemp.EOF Then      '如果医保中心目录中未找到该项目
                If str项目类型 = "药品" Then    '类型为药品时，医保类型为“丙类”
                    str医保类型 = "丙类": str自付比例 = "1"
                Else        '项目类型为诊疗或中药时，医保类型为“甲类”
                    str医保类型 = "甲类": str自付比例 = "0"
                End If
            Else        '如果医保中心目录中有该药品
                str医保类型 = IIf(rsTemp!zfbl = 0, "甲类", IIf(rsTemp!zfbl = 1, "丙类", "乙类"))
                str自付比例 = rsTemp!zfbl
            End If
        End If
''''''''''        strSQL = "Insert Into hi_InpatientPrescription (SerialNum,InpatientID,HospitalID,FeeType,RecipeSerial,DateDiagnose," & _
''''''''''            "Class,ItemID,ItemName,Specification,Price,Dosage,ScaleSelf,Operator) Values (" & lng流水 & ",'" & _
''''''''''            str住院号 & "'," & Trim(gstr医院编码) & ",1,Null,'" & Format(rs明细!发生时间, "yyyy-MM-dd HH:mm:ss") & _
''''''''''            "'," & IIf(str项目类型 = "诊疗", 1, 2) & ",'" & str明细编码 & _
''''''''''            "','" & str项目名称 & "','" & str规格 & "'," & Format(Nvl(rs明细!实收金额, 0) / (rs明细!付数 * rs明细!数次), _
''''''''''            "0.000") & "," & rs明细!付数 * rs明细!数次 & "," & str自付比例 & ",'" & UserInfo.姓名 & "')"
''''''''''
''''''''''        WriteInfo "传递明细(写处方明细):" & strSQL
''''''''''        gcn余姚.Execute strSQL
''''''''''        lng流水 = AddNum(lng流水)
''''''''''
''''''''''        'gstrSQL = "ZL_保险支付项目_Modify(" & rs明细!收费细目ID & "," & TYPE_余姚 & ",NULL,'" & str明细编码 & "','" & _
''''''''''        '    str项目名称 & "','" & str医保类型 & "',1)"
'''''''''''        WriteInfo "修改保险支付项目:" & gstrSQL
''''''''''        'Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
''''''''''
''''''''''        gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
''''''''''        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        
                '在费用记录中记录进入统筹金额
                '项目编码中保存项目类型（药品，诊疗）,摘要中保存自付比例,可根据比例得到甲类，乙类
        gstrSQL = "ZL_病人费用记录_更新医保(" & rs明细!ID & ","
        gstrSQL = gstrSQL & rs明细!实收金额 - rs明细!实收金额 * str自付比例 & ",NULL,1,'" & str项目类型 & "',NULL,'" & str自付比例 & "')"
        WriteInfo "修改明细项目的进入统筹金额:" & gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
        '^^^^^^^^^^
        rs明细.MoveNext
    Wend
    '调用接口
''''''''''    gstrIC明文 = makeICInfo(lng病人ID)
''''''''''    gstrInput余姚 = "$$" & str住院号 & "~" & gstrIC明文 & "~0000$$"
''''''''''    gstrOutput余姚 = Space(4000)
''''''''''
''''''''''    WriteInfo "记帐：f_UserBargaingApply(33, " & mstr流水号 & ", " & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
''''''''''
''''''''''    glngReturn = f_UserBargaingApply(33, mstr流水号, gstrInput余姚, gstrOutput余姚)
''''''''''    WriteInfo "反回：" & gstrOutput余姚
''''''''''
''''''''''    记帐传输_余姚 = CheckReturn_余姚()
''''''''''    If 记帐传输_余姚 = False Then
''''''''''        Exit Function
''''''''''    End If

    记帐传输_余姚 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    记帐传输_余姚 = False
End Function

Public Function 住院虚拟结算_余姚(rs费用明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim datCurr As Date, str住院号 As String
    
    On Error GoTo errHandle
    Set rs明细 = rs费用明细.Clone
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    '需要先上传费用明细
    If 记帐传输_余姚("", 0, "", lng病人ID) = False Then Exit Function
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = rsTemp!顺序号 'Format(Val(rsTemp!顺序号), "0" & String(16, "#"))
    
    '计算非医保金额
    mcur非医保 = 0
    While Not rs明细.EOF
'        gstrSQL = "Select B.类别,A.项目编码,A.项目名称,Nvl(B.规格,'') As 规格 From 保险支付项目 A,收费细目 B " & _
'            "Where A.收费细目ID=B.ID And A.是否医保=1 And A.险类=" & TYPE_余姚 & " And A.收费细目ID=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If Not rsTemp.EOF Then
'            '判断医保前置机上是否有该项目
'               gstrSQL = "Select MedicineID,zfbl From hi_Medicine  Where MedicineID='" & rsTemp(1) & "'"
'               gstrSQL = gstrSQL & " union all Select DiagnoseID,zfbl From hi_Diagnose Where DiagnoseID='" & rsTemp(1) & "'"
'            Set rsTemp = gcn余姚.Execute(gstrSQL)
'            If rsTemp.EOF Then mcur非医保 = mcur非医保 + rs明细!金额
'        Else
'            mcur非医保 = mcur非医保 + rs明细!金额
'        End If
        '
        If Nvl(rs明细!摘要, 0) = 1 Then mcur非医保 = mcur非医保 + rs明细!金额
        rs明细.MoveNext
    Wend
    mstr流水号 = 申请交易流水_余姚(34)
    '调用接口
    '该接口是全自费，虚拟结算无意义，为减少出错可能，暂不调用接口
    gstrIC明文 = makeICInfo(lng病人ID)
    gstrInput余姚 = "$$" & str住院号 & "~" & mcur非医保 & "~" & gstrIC明文 & "~0000$$"
    gstrOutput余姚 = Space(4000)
    WriteInfo "虚拟结算：f_UserBargaingApply(34, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    glngReturn = f_UserBargaingApply(34, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "反回：" & gstrOutput余姚
    'if CheckReturn_余姚 then
    
    住院虚拟结算_余姚 = "医保基金;0;0"
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    住院虚拟结算_余姚 = ""
End Function

Public Function 出院登记_余姚(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim rsTemp As New ADODB.Recordset, datCurr As Date, bln零费用出院 As Boolean, str住院号 As String, _
        strInNote As String, str病种编码 As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额 from 住院费用记录 where nvl(附加标志,0)<>9 and 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, lng主页ID)
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
    gstrSQL = "Select * From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID)
    str住院号 = rsTemp!顺序号 ' Format(Val(rsTemp!顺序号), "0" & String(16, "#")) ' Val()
    
    If bln零费用出院 = True Then
        '调用入院登记撤消
        mstr流水号 = 申请交易流水_余姚(40)
        gstrIC明文 = makeICInfo(lng病人ID)
        
        '调用接口
        gstrInput余姚 = "$$" & str住院号 & "~" & str住院号 & "~" & gstrIC明文 & "$$"
        gstrOutput余姚 = Space(4000)
        WriteInfo "出院撤消：f_UserBargaingApply(40, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
        glngReturn = f_UserBargaingApply(40, mstr流水号, gstrInput余姚, gstrOutput余姚)
        WriteInfo "反回：" & gstrOutput余姚

        出院登记_余姚 = CheckReturn_余姚()
        Exit Function
    End If
    
    '获取出院诊断
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID, False, True, True)
    
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_余姚
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "确诊疾病")
    If rsTemp.State = 1 Then
        str病种编码 = rsTemp!ID
    Else
        出院登记_余姚 = False
        Exit Function
    End If
    '获取住院医师
    gstrSQL = "select 住院医师 from 病案主页 Where 主页ID = [2] And 病人ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        MsgBox "不能取得病人的入院登记信息", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr流水号 = 申请交易流水_余姚(35)
    gstrIC明文 = makeICInfo(lng病人ID)
    
    '调用接口
    'gstrInput余姚 = "$$" & str住院号 & "~" & Nvl(rsTemp(0), " ") & "~" & strInNote & "~" & _
        str病种编码 & "~" & Format(datCurr, "yyyy-MM-dd") & "$$"
    'gstrOutput余姚 = Space(4000)
    'WriteInfo "出院登记：f_UserBargaingApply(35, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    'glngReturn = f_UserBargaingApply(35, mstr流水号, gstrInput余姚, gstrOutput余姚)
    'WriteInfo "反回：" & gstrOutput余姚
    '出院登记_余姚 = CheckReturn_余姚()
    'If 出院登记_余姚 = False Then Exit Function
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_余姚 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    出院登记_余姚 = True
  
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记_余姚 = False
End Function

Public Function 住院结算_余姚(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String, strTemp As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur发生费用金额 As Currency, cur全自付金额 As Currency
    
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From 住院费用记录 Where 记录状态<>0 And nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rs明细.EOF Then
        Err.Raise 9000, gstrSysName, "没有费用明细，不能进行出院结算"
        Exit Function
    End If
    lng病人ID = rs明细!病人ID
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng病人ID, TYPE_余姚)
    str就诊编号 = rsTemp!顺序号 ' Format(Val(rsTemp!顺序号), "0" & String(16, "#")) ' Nvl(rsTemp!顺序号)
    datCurr = zlDatabase.Currentdate
   
    '计算非医保项目金额
    mcur非医保 = 0
    cur发生费用金额 = 0
    cur全自付金额 = 0
    While Not rs明细.EOF
        If Nvl(rs明细!摘要, 0) = 1 Then mcur非医保 = mcur非医保 + Nvl(rs明细!实收金额, 0)
        cur发生费用金额 = cur发生费用金额 + Nvl(rs明细!实收金额, 0)
        rs明细.MoveNext
    Wend
    'mstr流水号 = 申请交易流水_余姚(36)
    'gstrIC明文 = makeICInfo(lng病人ID)
    'gstrInput余姚 = "$$1~" & str就诊编号 & "~" & mcur非医保 & "~" & gstrIC明文 & "$$"
    'gstrOutput余姚 = Space(4000)
    
    'WriteInfo "住院结算：f_UserBargaingApply(36, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    'glngReturn = f_UserBargaingApply(36, mstr流水号, gstrInput余姚, gstrOutput余姚)
    'WriteInfo "返回：" & gstrOutput余姚
    
    '住院结算_余姚 = CheckReturn_余姚()
    'If 住院结算_余姚 = False Then Exit Function
    'strTemp = Split(gstrOutput余姚, "$$")(2)
    '余姚返回结果不对，所以暂时不用
    '    cur发生费用金额 = CCur(Split(strTemp, "~")(0))
    
    'cur全自付金额 = CCur(Split(strTemp, "~")(1))
    '余姚地区只返回总额和全自负部分金额,。
    
    '帐户年度信息
    Call Get帐户信息(TYPE_余姚, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)

    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_余姚 & "," & _
            lng病人ID & "," & Year(datCurr) & ",0," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",0,NULL,NULL," & cur发生费用金额 & _
            "," & mcur非医保 & ",0,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,'" & _
            str就诊编号 & "~" & mstr流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    住院结算_余姚 = True
    '---------------------------------------------------------------------------------------------

    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 住院结算冲销_余姚(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, lng冲销ID As Long, str流水号 As String, str就诊编号 As String, _
        lng病人ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, _
        cur统筹报销累计 As Currency, int住院次数累计 As Integer, cur票据总金额 As Currency, cur全自付金额 As Currency
    Dim strTemp As String
    Dim datCurr As Date
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 住院费用记录 Where 记录状态<>0 And nvl(附加标志,0)<>9 and 结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "没找到病人的费用明细记录，不能退费"
        Exit Function
    End If
    lng病人ID = rsTemp("病人ID")
    Do Until rsTemp.EOF
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费

    gstrSQL = "select distinct A.ID as 结帐ID from 病人结帐记录 A,病人结帐记录 B " & _
              "  where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng结帐ID)
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=[1] and 记录ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, lng结帐ID)
    
    If rsTemp.EOF = True Then
        Err.Raise 9000, gstrSysName, "原单据的医保记录不存在，不能作废。"
        Exit Function
    End If
    If IsNull(rsTemp!备注) Then
        Err.Raise 9000, gstrSysName, "该单据的就诊编号丢失，不能作废。"
        Exit Function
    End If
    str就诊编号 = Split(rsTemp!备注, "~")(0)
    str流水号 = Split(rsTemp!备注, "~")(1)
    
    '调用接口数冲销
    mstr流水号 = 申请交易流水_余姚(37)
    gstrIC明文 = makeICInfo(lng病人ID)

    '调用接口
    gstrInput余姚 = "$$" & str流水号 & "~" & gstrIC明文 & "$$"
    gstrOutput余姚 = Space(4000)
    WriteInfo "住院结算：f_UserBargaingApply(37, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    glngReturn = f_UserBargaingApply(37, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "返回：" & gstrOutput余姚
    
    住院结算冲销_余姚 = CheckReturn_余姚()
    If 住院结算冲销_余姚 = False Then Exit Function
    strTemp = Split(gstrOutput余姚, "$$")(2)
    'cur票据总金额 = CCur(Split(strTemp, "~")(0))
    cur全自付金额 = CCur(Split(strTemp, "~")(1))

'^^^^^^^^^^^^^^
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_余姚 & ",'顺序号'," & str流水号 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    
    '帐户年度信息
    Call Get帐户信息(TYPE_余姚, lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_余姚 & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & "," & cur全自付金额 & ",0,0,0,0,0,0," & _
        "NULL,NULL,NULL,'" & str就诊编号 & "~" & str流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)

    住院结算冲销_余姚 = True
    Exit Function
errHandle:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 身份标识_余姚(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify余姚
    Dim strPatiInfo As String, cur余额 As Currency, str就诊编号 As String
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    'MODIFIED BY ZYB 宁海医保接口开发
    strPatiInfo = frmIDentified.GetPatient(bytType, lng病人ID)
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
        lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID, TYPE_余姚)
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        Unload frmIDentified
    Else
        身份标识_余姚 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_余姚 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_余姚 = ""
End Function

Public Function 个人余额_余姚(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    '余姚不能提取个人帐户余额
    个人余额_余姚 = 0
End Function

Public Function 转科转床_余姚(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSQL As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date, strTemp As String
    Dim lng病种ID As Long
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
    If rsTmp.BOF Then 转科转床_余姚 = False: Exit Function
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_余姚
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
        str病种 = rsTemp!名称
        str病种编码 = rsTemp!ID
    Else
        转科转床_余姚 = False
        Exit Function
    End If
    
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码,D.编码 As 科室编码,C.顺序号 As 住院流水 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = [1] And A.病人ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, lng主页ID, lng病人ID)
    
    mstr流水号 = 申请交易流水_余姚(38)
    gstrIC明文 = makeICInfo(lng病人ID)
    
    gstrInput余姚 = "$$" & rsTemp!住院流水 & "~" & Format(datCurr, "yyyy-MM-dd") & "~" & _
        rsTemp(3) & "~" & Nvl(rsTemp(4), " ") & "~" & strInNote & "~" & _
        str病种编码 & "~" & Nvl(rsTemp!住院科室, " ") & "~" & Nvl(rsTemp!科室编码, "0") & "$$"
    gstrOutput余姚 = Space(4000)
    
    WriteInfo "住院结算：f_UserBargaingApply(38, " & mstr流水号 & "," & Replace(gstrInput余姚, String(1053, "0"), "") & ", gstrOutput余姚)"
    glngReturn = f_UserBargaingApply(38, mstr流水号, gstrInput余姚, gstrOutput余姚)
    WriteInfo "返回：" & gstrOutput余姚
    
    转科转床_余姚 = CheckReturn_余姚()
    If 转科转床_余姚 = False Then
        Exit Function
    End If
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_余姚 & ",'病种ID'," & lng病种ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, gstrSysName)
    转科转床_余姚 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    转科转床_余姚 = False
End Function

Private Function AddNum(ByVal strSequence As String) As String
    Dim intAdd As Integer
    Dim intDO As Integer, intLen As Integer
    '完成加法运算
    intLen = Len(strSequence)
    intAdd = Val(Right(strSequence, 1))
    intAdd = intAdd + 1
'    If intAdd > 9 Then
        '进位
        For intDO = intLen To 1 Step -1
            intAdd = Val(Mid(strSequence, intDO, 1))
            intAdd = intAdd + 1
            If intAdd <= 9 Then
                '当前位及以前的位不动，后面的位全置为零
                AddNum = Mid(strSequence, 1, intDO - 1) & CStr(intAdd) & String(intLen - intDO, "0")
                Exit For
            End If
        Next
        '发生进位到最高位，增加最高位1，其余位为零
        If intDO = 0 Then AddNum = "1" & String(intLen, "0")
'    Else
'        AddNum = Mid(strSequence, 1, intLen - 1) & CStr(intAdd)
'    End If
End Function
