Attribute VB_Name = "mdl慈溪农医"
Option Explicit
Public Declare Sub CXNY_SetRemoteServerAddr Lib "JFNetLib.dll" Alias "SetRemoteServerAddr" _
    (ByVal lngPort As Long, ByVal strIP As String)
Public Declare Function CXNY_SendRequestPack Lib "JFNetLib.dll" Alias "SendRequestPack" _
    (ByVal strSend As String, ByVal lngSend As Long, ByVal strReceive As String, lngReceive As Long, ByVal lngWaitSecs As Long) As Long

'业务主体功能函数
Public Const gstrFunc慈溪农医_GetServerTime As String = "H888"     '取服务器时间
Public Const gstrFunc慈溪农医_GetHospitalInfo As String = "H301"   '取医院名称及编码
Public Const gstrFunc慈溪农医_GetPersonalInfo As String = "H302"   '取个人信息
Public Const gstrFunc慈溪农医_OutRegist As String = "H101"         '门诊登记
Public Const gstrFunc慈溪农医_InRegist As String = "H201"          '入院登记
Public Const gstrFunc慈溪农医_InRegistCancel As String = "H401"    '入院登记取消
Public Const gstrFunc慈溪农医_UploadDetail As String = "H303"      '上传处方明细
Public Const gstrFunc慈溪农医_InBalance As String = "H202"         '住院结算
Public Const gstrFunc慈溪农医_OutBalance As String = "H102"        '门诊结算
Public Const gstrFunc慈溪农医_BalanceCancel As String = "H103"     '结算作废
Public Const gstrFunc慈溪农医_ModifyInfo As String = "H606"        '修改住院信息
Public Const gstrFunc慈溪农医_DelAllDetail As String = "H607"      '删除住院部分所有已上传处方明细
Public Const gstrFunc慈溪农医_DelAllDetail1 As String = "H608"      '删除门诊所有已上传处方明细
'用于对帐的功能函数
Public Const gstrFunc慈溪农医_OutQuery As String = "H601"          '核对门诊
Public Const gstrFunc慈溪农医_InQuery As String = "H602"           '核对住院
Public Const gstrFunc慈溪农医_AllowQuery As String = "H603"        '备案查询
Public Const gstrFunc慈溪农医_Exception_Out As String = "H604"     '例外时的门诊业务取消
Public Const gstrFunc慈溪农医_Exception_In As String = "H605"      '例外时的住院业务取消

Private Const mstrAmountFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrPriceFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrMoneyFormat As String = "#0.0000;-#0.0000;0;"
Private Const mstrDateFormat As String = "yyyy-MM-dd HH:mm:ss"
Private Const mstrSplit As String = "&"
Private mblnInit As Boolean                                         '是否正常初始化

Private Type ComInfo_慈溪农医
    医院编码 As String
    医院名称 As String
    业务类型 As String
    医疗证号 As String
    个人编号 As String
    就诊流水号 As String
    结算流水号 As String
    总费用 As Currency                      'HIS
    总费用_中心 As Currency                 '中心的费用总额
    结算串 As String
    结算串反 As String
    结算类型 As String
    疾病编码 As String
    
End Type
Public gComInfo_慈溪农医 As ComInfo_慈溪农医

'06-03-25应军杰修改
Public g门诊农保报销 As Currency    '门诊虚拟农保可以报销的金额
Public g门诊流水号 As String
Public g门诊临时标志 As String      '是否门诊
'06-03-25



Private mstrFunc As String              '功能号
Private mstrInput As String             '输入串
Private mlngInput As Long               '输入串长度
Private mstrOutput As String            '输出串
Private mlngOutput As Long              '输出串长度
Private mlngReturn As Long              '等待秒数>=30
Public gstrOutput_慈溪农医 As String

'测试用代码----------------------------------------------
'Private gobjCXXNY As New clsT_CXXNY

Public Function 医保设置_慈溪农医() As Boolean
    医保设置_慈溪农医 = frmSet慈溪农医.ShowME
End Function

Public Sub 调用接口_准备_慈溪农医(ByVal strFunc As String, ByVal StrInput As String)
    mstrFunc = strFunc
    mstrOutput = String(2000, " ")
    mlngOutput = 2000
    mstrInput = "exchcode=" & mstrFunc & mstrSplit & StrInput
    mlngInput = LenB(StrConv(mstrInput, vbFromUnicode))
End Sub

Public Function 调用接口_慈溪农医() As Boolean
    Dim strMsg As String
    Dim arrReturn
    Dim blnSuccess As Boolean
    
'    mlngReturn = gobjCXXNY.CXNY_SendRequestPack(mstrInput, mlngInput, mstrOutput, mlngOutput, 30)
    mlngReturn = CXNY_SendRequestPack(mstrInput, mlngInput, mstrOutput, mlngOutput, 30)
    Select Case mlngReturn
    Case 0
        blnSuccess = True
    Case 1
        strMsg = "连接医院前置机服务器失败"
    Case 2
        strMsg = "向远程服务器发送数据失败"
    Case 3
        strMsg = "接收返回值结果失败"
    Case 4
        strMsg = "动态链接库不存在"
    Case Else
        strMsg = "发生未知错误"
    End Select
    
    arrReturn = Split(mstrOutput, "&")
    If blnSuccess Then
        blnSuccess = (Val(Split(arrReturn(0), "=")(1)) = 0)
        If blnSuccess = False Then strMsg = Split(arrReturn(1), "=")(1)
    End If
    
    If blnSuccess = False Then
        MsgBox strMsg & vbCrLf & "功能：" & mstrFunc & "|错误号：" & mlngReturn, vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrOutput_慈溪农医 = mstrOutput
    调用接口_慈溪农医 = True
End Function

Public Function 身份标识_慈溪农医(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim arrReturn
    Dim StrInput As String
    Dim intVerify As Integer
    Dim strDiseaseCode As String            '疾病编码
    Dim strIdentify As String
    Dim strRegistCode As String             '挂号单号
    Dim strRegisterOffice As String         '就诊科室
    Dim strRegisterDoctor As String         '医生
    Dim rsTemp As New ADODB.Recordset
    Dim STR设置时间 As String
    STR设置时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If STR设置时间 > "2007-04-01" Then
       Exit Function
    End If
       
    On Error GoTo errHand
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    
      
    strIdentify = frmIdentify慈溪农医.GetPatient(bytType, lng病人ID)
    If strIdentify = "" Then Exit Function
    If Not (bytType = 1 Or bytType = 0 Or bytType = 3) Then Exit Function
    
    '进行门诊登记
    gComInfo_慈溪农医.结算串 = ""
    If bytType = 0 Then
        '检查是否已经过备案
        'Balx    varchar(2)  就医类型(=1 急诊备案 =2 转院备案 =3 特殊病种备案)
        'Kzhm    varchar(30) 卡证号码
        '接收数据包参数
        'Returncode  Long    如果返回0表示成功
        'Returninfo  varchar(50) 对应的错误提示
        'Spbz    varchar(1)  审批标志=0未审批 =1 审批通过 =2审批未通过
       ' StrInput = "Balx=3" & mstrSplit & "Kzhm=" & gComInfo_慈溪农医.个人编号
        
       ' Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_AllowQuery, StrInput)
        'If Not 调用接口_慈溪农医() Then Exit Function
        '检查是否通过备案
        'arrReturn = Split(gstrOutput_慈溪农医, mstrSplit)
        'intVerify = Val(Split(arrReturn(2), "=")(1))
        'Select Case intVerify
        'Case 0
         '   MsgBox "农医办还未审批，不允许办理门诊业务！", vbInformation, gstrSysName
            'Exit Function
        'Case 2
         '   MsgBox "农医办没有通过审批，不允许办理门诊业务！", vbInformation, gstrSysName
            'Exit Function
        'End Select
        
        '入参：合作医疗号码│合作医疗病人在医院就诊的挂号号码│就诊的医疗类别│医院就诊的科室│就诊的医生│" & _
        医院的诊断│医院就诊登记的日期│并发症│就诊机构的机构编码│就诊机构的机构名称│经办单位│经办人
        '取当天挂号的科室与医生
        gstrSQL = " Select B.名称 AS 挂号科室,执行人 AS 医生 " & _
                  " From 门诊费用记录 A,部门表 B " & _
                  " Where A.记录性质=4 And A.记录状态=1 And A.病人ID=" & lng病人ID & _
                  " And A.执行部门ID=B.ID And A.登记时间 Between to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 00:00:00','yyyy-MM-dd hh24:mi:ss')" & _
                  " And to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss') And Rownum<2"
      '  Call OpenRecordset(rsTemp, "取当天挂号的科室与医生")
      '  If rsTemp.RecordCount = 0 Then
      '      MsgBox "今天没有有效的挂号记录,无法进行门诊就诊登记！", vbInformation, gstrSysName
       '     Exit Function
      '  End If
        'strRegisterOffice = rsTemp!挂号科室
         'strRegisterDoctor = rsTemp!医生
        strRegisterOffice = "001"
        strRegisterDoctor = "001"
        '取疾病编码
        gstrSQL = "Select 编码 From 疾病编码目录 Where ID=(Select nvl(病种ID,0) From 保险帐户 Where 险类=[1] And 病人ID=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取疾病编码", TYPE_慈溪农医, lng病人ID)
        If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!编码
        
        '获取挂号单号，十位，唯一标识
        strRegistCode = CStr(zlDatabase.GetNextID("部门表"))
         
        gComInfo_慈溪农医.就诊流水号 = strRegistCode
        '门诊就诊登记入参
'        Mzdjh   varchar(20) 门诊登记号，在医院数据库中的唯一号码
'        Kzhm    varchar(30) 卡证号码
'        Jzks    varchar(20) 就诊科室
'        Ysxm    varchar(10) 医生姓名
'        Bzdm    varchar(20) 病种代码
'        Mzrq    Datetime    就诊日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
'        Czy     Varchar(10) 操作员姓名
      
        StrInput = "Mzdjh=" & strRegistCode & mstrSplit & "Kzhm=" & gComInfo_慈溪农医.个人编号 & mstrSplit & _
            "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
            "Bzdm=" & strDiseaseCode & mstrSplit & "Mzrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
         
        If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
           Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_OutRegist, StrInput)
        
           If Not 调用接口_慈溪农医() Then Exit Function
        End If
        
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_慈溪农医 & ",'业务类型','''" & gComInfo_慈溪农医.业务类型 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    End If
    
    If bytType = 1 Then
        '更新保险帐户相关信息（业务类型）
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_慈溪农医 & ",'业务类型','''" & gComInfo_慈溪农医.业务类型 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    End If
    
    '返回病人信息串
    身份标识_慈溪农医 = strIdentify
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保初始化_慈溪农医(Optional ByVal blnTest As Boolean = False) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim lngPort As Long
    Dim strIP As String
    Dim rsTemp As New ADODB.Recordset
    Dim cnTest As New ADODB.Connection

    On Error Resume Next
    
    If mblnInit = False Then
        '取医院编码
        gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取医院编码", TYPE_慈溪农医)
        gComInfo_慈溪农医.医院编码 = Nvl(rsTemp!医院编码)
        
        '取保险参数
        gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取保险参数", TYPE_慈溪农医)
        Do While Not rsTemp.EOF
            Select Case rsTemp!参数名
            Case "IP地址"
                strIP = Nvl(rsTemp!参数值, "127.0.0.1")
            Case "端口号"
                lngPort = Nvl(rsTemp!参数值, 8801)
            End Select
            rsTemp.MoveNext
        Loop
        
        '测试是否连接的通
'        Call gobjCXXNY.CXNY_SetRemoteServerAddr(lngPort, strIP)
'yjj1row
        Call CXNY_SetRemoteServerAddr(lngPort, strIP)
        mblnInit = True
    End If
    
    医保初始化_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保终止_慈溪农医() As Boolean
    On Error Resume Next
    
    医保终止_慈溪农医 = True
End Function

Public Function 门诊虚拟结算_慈溪农医(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim StrInput As String, strinput1 As String
    Dim lng病人ID As Long, 退条数 As Long
    Dim str开方日期 As String, str处方号 As String, str医保编码 As String, str项目名称 As String, str规格 As String, str单位 As String
    Dim dbl帐户支付 As Double, dbl现金 As Double, dbl统筹基金 As Double
    Dim fpze As Double, ylzl As Double, blzf As Double
    Dim qfbz As Double, sbje As Double, Zhzf As Double
    Dim Grzf As Double, Zhye As Double, dnbxlj As Double
    Dim cfdkbje As Double
    Dim jsdh As Double
    Dim tbkbje As Double
    Dim tbsbje As Double
    Dim strRegisterOffice  As String
    Dim strRegisterDoctor As String
    Dim strDiseaseCode As String
    Dim strRegistCode As String
    Dim lngrowcount As String
    
    
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    
    '处方明细作废入参（因门诊没有该函数，对方答应考虑一下是否提供）
   ' If gComInfo_慈溪农医.结算串 <> "" Then
        
    '    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_DelAllDetail1, gComInfo_慈溪农医.结算串反)
     '   If Not 调用接口_慈溪农医() Then Exit Function
    'End If
    
    lng病人ID = rs明细!病人ID
    
    ''''''''''''''''''
    
      gComInfo_慈溪农医.结算串 = ""
    
        
        '入参：合作医疗号码│合作医疗病人在医院就诊的挂号号码│就诊的医疗类别│医院就诊的科室│就诊的医生│" & _
        医院的诊断│医院就诊登记的日期│并发症│就诊机构的机构编码│就诊机构的机构名称│经办单位│经办人
       ' 取当天挂号的科室与医生
       ' strRegisterOffice = rs明细!开单部门ID
          strRegisterOffice = "00001"
         strRegisterDoctor = rs明细!开单人
         
       'gstrSQL = " Select 名称 AS 挂号科室 from 部门表 " & _
      '            " Where ID=" & strRegisterOffice
      '           WriteInfo (gstrSQL)
      '  Call OpenRecordset(rsTemp, "取当天挂号的科室")
  '    If rsTemp.RecordCount = 0 Then
   '         MsgBox "今天没有有效的挂号记录,无法进行门诊就诊登记！", vbInformation, gstrSysName
   '    Exit Function
    '   End If
     '   strRegisterOffice = rsTemp!挂号科室
      '
        '取疾病编码
        gstrSQL = "Select 编码 From 疾病编码目录 Where ID=(Select nvl(病种ID,0) From 保险帐户 Where 险类=[1] And 病人ID=[2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取疾病编码", TYPE_慈溪农医, lng病人ID)
        If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!编码
        gComInfo_慈溪农医.疾病编码 = strDiseaseCode
        '获取挂号单号，十位，唯一标识
        strRegistCode = CStr(zlDatabase.GetNextID("部门表"))
         
        gComInfo_慈溪农医.就诊流水号 = strRegistCode
        '门诊就诊登记入参
'        Mzdjh   varchar(20) 门诊登记号，在医院数据库中的唯一号码
'        Kzhm    varchar(30) 卡证号码
'        Jzks    varchar(20) 就诊科室
'        Ysxm    varchar(10) 医生姓名
'        Bzdm    varchar(20) 病种代码
'        Mzrq    Datetime    就诊日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
'        Czy     Varchar(10) 操作员姓名
        strinput1 = "Mzdjh=" & strRegistCode & mstrSplit & "Kzhm=" & gComInfo_慈溪农医.个人编号 & mstrSplit & _
            "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
            "Bzdm=" & strDiseaseCode & mstrSplit & "Mzrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
        
        If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
           Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_OutRegist, strinput1)
           If Not 调用接口_慈溪农医() Then Exit Function
        End If
     
    ''''''''''''''''''''
       str开方日期 = Format(zlDatabase.Currentdate, mstrDateFormat)
        
    '得到本次结算的总费用
    With rs明细
        '求费用总额
        gComInfo_慈溪农医.总费用 = 0
        Do While Not .EOF
            gComInfo_慈溪农医.总费用 = gComInfo_慈溪农医.总费用 + !实收金额
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
          lngrowcount = 0
        Do While Not .EOF
            '提取收费细目的相关信息
            If Nvl(!实收金额, 0) <> 0 Then
            lngrowcount = lngrowcount + 1
             gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,A.计算单位 AS 单位,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                      " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
            Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_慈溪农医, CLng(!收费细目ID))
           
            
            str项目名称 = Nvl(rsItem!名称, "无")
             If IsNull(rsItem!项目编码) = True Or rsItem!项目编码 = "" Then
                MsgBox "医院项目" & str项目名称 & "没有对码，请先对码！"
                Exit Function
            End If
            str医保编码 = Nvl(rsItem!项目编码)
            str规格 = Nvl(rsItem!规格, "无")
            str单位 = Nvl(rsItem!单位, "无")
            If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
            str规格 = Nvl(str规格, "无")
            '处方明细上传入参
        '    Jylx    varchar(1)  就医类型(=0 门诊；=1 住院)
        '    Jyhm    varchar(20) 就医号码(入院登记号或门诊登记号)
        '    Recordcount numeric(10) 记录条数(＜=100)
        '    xmdm[i] varchar(10) 项目代码(新农医项目代码)
        '    zxxmmc[i]   varchar(80) 中心项目名称(医院项目名称)
        '    xmdw[i] varchar(10) 项目单位
        '    xmdj[i] numeric(10,4)   项目单价
        '    xmsl[i] numeric(10,4)   项目数量
        '    xmje[i] numeric(10,4)   项目金额
        '    xmgg[i] varchar(50) 项目规格
        '    yzrq[i] Datetime    医嘱日期或门诊日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
            
            StrInput = StrInput & mstrSplit & "xmdm[" & lngrowcount & "]=" & str医保编码 & mstrSplit & "zxxmmc[" & lngrowcount & "]=" & ToVarchar(str项目名称, 80) & mstrSplit & _
                "xmdw[" & lngrowcount & "]=" & ToVarchar(str单位, 10) & mstrSplit & "xmdj[" & lngrowcount & "]=" & Format(!单价, mstrPriceFormat) & mstrSplit & _
                "xmsl[" & lngrowcount & "]=" & Format(!数量, mstrAmountFormat) & mstrSplit & "xmje[" & lngrowcount & "]=" & Format(!实收金额, mstrMoneyFormat) & mstrSplit & _
                "xmgg[" & lngrowcount & "]=" & ToVarchar(str规格, 50) & mstrSplit & "yzrq[" & lngrowcount & "]=" & str开方日期
             End If
     
            .MoveNext
            
          
        Loop
         If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
            StrInput = "Jylx=0" & mstrSplit & "Jyhm=" & gComInfo_慈溪农医.就诊流水号 & mstrSplit & "Recordcount=" & lngrowcount & StrInput
             Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
             If Not 调用接口_慈溪农医() Then Exit Function
         End If
    End With
    'gComInfo_慈溪农医.结算串反 = strinput1
    '预结算的入参
'    Jsbj    Varchar(1)  结算标记(=0 预算 =1 结算)
'    Mzdjh   varchar(20) 门诊登记号
'    Kzhm    varchar(30) 卡证号码
'    Jsrq    Datetime    结算日期(门诊日期) (固定格式19位：yyyy-mm-dd hh:mm:ss)
'    Recordcount Long    明细记录数
'    Fpze    numeric(10,2)   发票总额
'    Czy Varchar(10) 操作员姓名
    StrInput = "Jsbj=0" & mstrSplit & "Mzdjh=" & gComInfo_慈溪农医.就诊流水号 & mstrSplit & "Kzhm=" & gComInfo_慈溪农医.个人编号 & mstrSplit & _
        "Jsrq=" & str开方日期 & mstrSplit & "Recordcount=" & lngrowcount & mstrSplit & "Fpze=" & Format(gComInfo_慈溪农医.总费用, "#0.00") & mstrSplit & "Czy=" & gstrUserName
    
    gComInfo_慈溪农医.结算串 = StrInput
    gComInfo_慈溪农医.结算串反 = "Mzdjh=" & gComInfo_慈溪农医.就诊流水号 & mstrSplit & "Kzhm=" & gComInfo_慈溪农医.个人编号
    
    If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
       Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_OutBalance, StrInput)
       If Not 调用接口_慈溪农医() Then Exit Function
    End If
    '出参：
'    Returncode  Long    如果返回0表示成功
'    Returninfo  varchar(50) 对应的错误提示
'    Fpze    numeric(10,2)   发票总额
'    Ylzl    numeric(10,2)   乙类自理
'    Blzf    numeric(10,2)   丙类自费
'    Qfbz    numeric(10,2)   起付标准
'    Sbje    numeric(10,2)   实报金额
'    Zhzf    numeric(10,2)   帐户支付
'    Grzf    numeric(10,2)   个人自付
'    Zhye    numeric(10,2)   帐户余额
'    Dnbxlj  Numeric(10,2)   当年统筹报销累计
'    Cfdkbje numeric(10,2)   超封顶有效金额(用于报销大病)
    If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
        gComInfo_慈溪农医.总费用_中心 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(2), "=")(1), "#0.00"))
        If Format(gComInfo_慈溪农医.总费用_中心, "#0.00") <> Format(gComInfo_慈溪农医.总费用, "#0.00") Then
            Err.Raise 9000, gstrSysName, "医院的总费用与医保中心的总费用不一致！" & vbCrLf & _
            "医院：" & Format(gComInfo_慈溪农医.总费用, "#0.00") & Space(10) & "医保中心：" & Format(gComInfo_慈溪农医.总费用_中心, "#0.00"), vbInformation, gstrSysName
        End If
    
        dbl统筹基金 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(6), "=")(1), "#0.00"))
        dbl帐户支付 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(7), "=")(1), "#0.00"))
        fpze = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(2), "=")(1), "#0.00"))
        ylzl = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(3), "=")(1), "#0.00"))
        blzf = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(4), "=")(1), "#0.00"))
        qfbz = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(5), "=")(1), "#0.00"))
        Grzf = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(8), "=")(1), "#0.00"))
        Zhye = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(9), "=")(1), "#0.00"))
        dnbxlj = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(10), "=")(1), "#0.00"))
        cfdkbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(11), "=")(1), "#0.00"))
        jsdh = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(12), "=")(1), "#0.00"))
        tbkbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(13), "=")(1), "#0.00"))
        tbsbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(14), "=")(1), "#0.00"))
        dbl现金 = gComInfo_慈溪农医.总费用 - dbl统筹基金 - dbl帐户支付
    Else
        dbl帐户支付 = 0
        dbl统筹基金 = 0
    End If
    str结算方式 = "个人帐户;" & dbl帐户支付 & ";0|统筹基金;" & dbl统筹基金 & ";0"
      
    门诊虚拟结算_慈溪农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算_慈溪农医(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    Dim lng病人ID As Long
    Dim StrInput As String, str结算单号 As String
    
    Dim dbl起付线 As Double, dbl统筹基金 As Double, dbl现金 As Double
        Dim fpze As Double, ylzl As Double, blzf As Double
    Dim qfbz As Double, sbje As Double, Zhzf As Double
    Dim Grzf As Double, Zhye As Double, dnbxlj As Double
    Dim cfdkbje As Double
    Dim jsdh As Double
    Dim tbkbje As Double
    Dim tbsbje As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    StrInput = gComInfo_慈溪农医.结算串
    StrInput = Replace(StrInput, "Jsbj=0", "Jsbj=1")
     If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
        Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_OutBalance, StrInput)
        If Not 调用接口_慈溪农医() Then Exit Function
     End If
    '出参：
'    Returncode  Long    如果返回0表示成功
'    Returninfo  varchar(50) 对应的错误提示
'    Fpze    numeric(10,2)   发票总额
'    Ylzl    numeric(10,2)   乙类自理
'    Blzf    numeric(10,2)   丙类自费
'    Qfbz    numeric(10,2)   起付标准
'    Sbje    numeric(10,2)   实报金额
'    Zhzf    numeric(10,2)   帐户支付
'    Grzf    numeric(10,2)   个人自付
'    Zhye    numeric(10,2)   帐户余额
'    Dnbxlj  Numeric(10,2)   当年统筹报销累计
'    Cfdkbje numeric(10,2)   超封顶有效金额(用于报销大病)
'   Jsdh    numeric(15) 结算单号
'   Tbkbje  numeric(10,2)   特病可报金额
'   Tbsbje  numeric(10,2)   特病实报金额

      '取病人ID
    gstrSQL = "Select 病人ID From 门诊费用记录 Where 结帐ID=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的ID", lng结帐ID)
    lng病人ID = rsTemp!病人ID
    If gComInfo_慈溪农医.结算类型 <> "事后补报" Then
        dbl起付线 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(5), "=")(1), "#0.00"))
        dbl统筹基金 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(6), "=")(1), "#0.00"))
        cur个人帐户 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(7), "=")(1), "#0.00"))
        dbl现金 = gComInfo_慈溪农医.总费用_中心 - dbl统筹基金 - cur个人帐户
        str结算单号 = str(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(12), "=")(1), "#0"))
         fpze = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(2), "=")(1), "#0.00"))
        ylzl = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(3), "=")(1), "#0.00"))
        blzf = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(4), "=")(1), "#0.00"))
        qfbz = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(5), "=")(1), "#0.00"))
        Grzf = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(8), "=")(1), "#0.00"))
        Zhye = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(9), "=")(1), "#0.00"))
        dnbxlj = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(10), "=")(1), "#0.00"))
        cfdkbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(11), "=")(1), "#0.00"))
        jsdh = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(12), "=")(1), "#0.00"))
        tbkbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(13), "=")(1), "#0.00"))
        tbsbje = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(14), "=")(1), "#0.00"))
 
    
    Else
       dbl起付线 = 0
       dbl统筹基金 = 0
       dbl现金 = 0
       str结算单号 = "事后补报"
        gstrSQL = "insert into 病人费用记录_事后补报慈系 Select * From 门诊费用记录 Where 结帐ID=" & lng结帐ID
        gcnOracle.Execute gstrSQL
        gstrSQL = "delete from 病人费用记录_事后补报慈系  Where 是否上传=1 and 结帐ID=" & lng结帐ID
       gcnOracle.Execute gstrSQL
        
       gstrSQL = "update 病人费用记录_事后补报慈系 set 执行人='" & gComInfo_慈溪农医.疾病编码 & "' Where 结帐ID=" & lng结帐ID
       gcnOracle.Execute gstrSQL
        
    End If
         
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_慈溪农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
         dnbxlj & "," & "NULL" & "," & qfbz & "," & 0 & "," & 0 & "," & _
        gComInfo_慈溪农医.总费用 & "," & blzf & "," & ylzl & "," & fpze - blzf - ylzl & "," & dbl统筹基金 & "," & tbkbje & " ," & tbsbje & "," & _
        cur个人帐户 & ",'" & gComInfo_慈溪农医.就诊流水号 & "',null,null,'" & str结算单号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    gComInfo_慈溪农医.结算串 = ""
    门诊结算_慈溪农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_慈溪农医(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str卡证号码 As String
    Dim StrInput As String
    Dim str结算单号
    Dim rsTemp As New ADODB.Recordset, rsTemp1 As New ADODB.Recordset
    On Error GoTo errHand
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    '只能退最后一笔
    '取冲销记录的结帐ID，单据号
    
    '取卡证号码
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读卡证号码", TYPE_慈溪农医, lng病人ID)
    str卡证号码 = Nvl(rsTemp!卡号)
    
    gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 And 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID)
    If rsTemp.RecordCount = 0 Then
        Err.Raise 9000, gstrSysName, "没有找到原始结算记录，无法进行门诊结算冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    gComInfo_慈溪农医.就诊流水号 = Nvl(rsTemp!支付顺序号)
    str结算单号 = Nvl(rsTemp!备注)
    
    '调用结算冲销
'    Jylx    varchar(1)  就医类型(=0 门诊作废 = 1 取消出院)
'    Jyhm    varchar(20) 就医号码
'    Kzhm    varchar(30) 卡证号码
'    Tfrq    Datetime    就诊日期(退费日期) (固定格式19位：yyyy-mm-dd hh:mm:ss)
'    Czy     Varchar(10) 操作员姓名
   If str结算单号 <> "事后补报" Then
        StrInput = "Jylx=0" & mstrSplit & "Jyhm=" & gComInfo_慈溪农医.就诊流水号 & mstrSplit & _
        "Kzhm=" & str卡证号码 & mstrSplit & "Tfrq=" & Format(zlDatabase.Currentdate(), mstrDateFormat) & mstrSplit & "Czy=" & gstrUserName
        Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_BalanceCancel, StrInput)
        If Not 调用接口_慈溪农医() Then Exit Function
    Else '事后补报的如果要退费要先判断是否已经补报，如果已经补报则需要先在补报处先退费，然后在这里退费
       gstrSQL = "Select * From 病人费用记录_事后补报慈系 Where 是否上传='1' and 结帐ID=[1]"
       Set rsTemp1 = zlDatabase.OpenSQLRecord(gstrSQL, "取事后补报记录的状态", lng结帐ID)
       If rsTemp1.RecordCount > 0 Then '说明已经结报，不允许退费
           MsgBox "该单据是事后结报单据，并且已经结算，不能退费！"
           Exit Function
       End If
       gstrSQL = "update 病人费用记录_事后补报慈系 set 是否上传='8' Where 结帐ID=" & lng结帐ID
       gcnOracle.Execute gstrSQL
    End If
    '保存本次结算情况
       
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_慈溪农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & "," & -1 * Nvl(rsTemp!首先自付金额, 0) & "," & -1 * Nvl(rsTemp!进入统筹金额, 0) & "," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",0," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & _
        -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & rsTemp!支付顺序号 & "',null,null,'" & rsTemp!备注 & "')"
        
    Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊结算冲销")
    
    门诊结算冲销_慈溪农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 入院登记_慈溪农医(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim StrInput As String
    Dim strKey As String                    '入院登记号(基数=21)
    Dim strRegister As String               '结算类型
    Dim strCardNO As String, lngDisease As Long '病人的医疗卡号及病种ID
    Dim strRegistCode As String             '挂号单号
    Dim strInHospitalDate As String         '入院日期
    Dim strRegisterOffice As String         '就诊科室
    Dim strDiseaseCode As String            '病种代码
    Dim strDiagnose As String               '入院诊断
    Dim strRegisterDoctor As String         '医生
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strKey = GetKey(lng病人ID, lng主页ID)
    '取病人的医疗卡号
    gstrSQL = "Select 卡号,Nvl(病种ID,0) 病种ID,Nvl(业务类型,'21') AS 结算类型 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的医疗卡号", TYPE_慈溪农医, lng病人ID)
    strCardNO = rsTemp!卡号
    lngDisease = rsTemp!病种ID
    strRegister = Val(rsTemp!结算类型) - 21
    
    '取科室与医生
    gstrSQL = " Select A.入院日期,B.名称 科室,A.门诊医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A.入院科室ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取科室与医生", lng病人ID, lng主页ID)
    strInHospitalDate = Format(rsTemp!入院日期, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!医生)
    strRegisterOffice = Nvl(rsTemp!科室)
    '取病种代码
    gstrSQL = "Select 编码 From 疾病编码目录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病种代码", lngDisease)
    If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!编码
    '取入院诊断
    strDiagnose = 获取入出院诊断(lng病人ID, lng主页ID, True, True, False)
    
    '入参：
'    Rydjh   varchar(20) 入院登记号，在医院数据库中的唯一号码
'    Kzhm    varchar(30) 卡证号码
'    Jzks    varchar(20) 就诊科室
'    Jslx    Varchar(1)  结算类型(=0 普通 =1 交通事故 =2 大病救助 =3 难产 =4 其他)
'    Ysxm    varchar(10) 医生姓名
'    Bzdm    Varchar(20) 病种代码
'    Ryzdsm  varchar(254)    入院诊断说明
'    Ryrq    Datetime    入院日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
'    Czy Varchar(10) 操作员姓名
    '返回参数:
'    Returncode  Long    如果返回0表示成功
'    Returninfo  varchar(50) 对应的错误提示
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & _
        "Jzks=" & strRegisterOffice & mstrSplit & "Jslx=" & strRegister & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
        "Bzdm=" & strDiseaseCode & mstrSplit & "Ryzdsm=" & strDiagnose & mstrSplit & _
        "Ryrq=" & strInHospitalDate & mstrSplit & "Czy=" & UserInfo.姓名
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_InRegist, StrInput)
    If Not 调用接口_慈溪农医() Then Exit Function

    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_慈溪农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")

    入院登记_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_慈溪农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim StrInput As String
    Dim strKey As String, strCardNO As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '如果主页ID小于一位，则前面加一个零
    strKey = GetKey(lng病人ID, lng主页ID)
    
    '取病人的医疗卡号
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的医疗卡号", TYPE_慈溪农医, lng病人ID)
    strCardNO = rsTemp!卡号

    '入参：
'    Rydjh   varchar(20) 入院登记号，在医院数据库中的唯一号码
'    Kzhm    varchar(30) 卡证号码
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_InRegistCancel, StrInput)
    If Not 调用接口_慈溪农医 Then Exit Function

    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_慈溪农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    gstrSQL = "zl_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    
    入院登记撤销_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_慈溪农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHand
    
    Call 更新在院信息_慈溪农医(lng病人ID, lng主页ID, True)
    '允许操作员修改就诊类型，同时保存
    Call frm就诊类型修改.ShowME(lng病人ID)
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_慈溪农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    
    出院登记_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_慈溪农医(lng病人ID As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHand
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_慈溪农医 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 个人余额_慈溪农医(strSelfNo As String) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '功能: 提取参保病人个人帐户余额
    '参数: strSelfNO-病人个人编号
    '返回: 返回个人帐户余额的金额
    '如果是门诊，返回家庭帐户余额；住院返回个人帐户余额
    gstrSQL = "Select Nvl(帐户余额,0) AS 个人帐户 From 保险帐户 Where 医保号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人ID", strSelfNo)
    个人余额_慈溪农医 = rsTemp!个人帐户
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 处方上传_慈溪农医(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal strNO As String) As Boolean
    Dim intCOUNT As Integer
    Dim lng主页ID As Long, lng病人ID As Long
    Dim StrInput As String
    Dim blnInsure As Boolean, blnTrans As Boolean
    Dim str项目名称 As String, str医保编码 As String, str规格 As String, str单位 As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo errHand
    '上传处方明细（记帐完成后上传，建一个触发器，不允许录入未对码项目，然后一一打上传标记）
    '打开本次待上传的处方明细
    gstrSQL = " Select A.ID,A.记录性质,A.记录状态,A.NO,A.序号,A.收费类别,A.病人ID,A.主页ID,A.收费细目ID,A.登记时间,A.实收金额," & _
              " Nvl(A.付数,1)*A.数次 AS 数量,A.实收金额/(Nvl(A.付数,1)*A.数次) AS 价格" & _
              " From 住院费用记录 A,保险帐户 B" & _
              " Where A.记录性质=" & int性质 & " ANd A.记录状态=" & int状态 & " And A.NO='" & strNO & "' And Nvl(A.是否上传,0)=0 And Nvl(A.实收金额,0)<>0 " & _
              " And A.病人ID=B.病人ID And B.险类=" & TYPE_慈溪农医 & _
              " Order by A.病人ID"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次待上传的处方明细")

    '处方完成后上传
    gcnOracle.BeginTrans
    blnTrans = True
    With rsDetail
        '上传处方
        lng病人ID = 0
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                If lng病人ID <> 0 Then
                    '说明处方明细已组装好了，准备上传
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(!病人ID, !主页ID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
                    If 调用接口_慈溪农医() Then
                        gcnOracle.CommitTrans
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                End If
                
                intCOUNT = 0
                StrInput = ""
                lng病人ID = !病人ID
                lng主页ID = !主页ID
                blnInsure = IsYBPatient(lng病人ID)
            End If

            If blnInsure Then
                '提取收费细目的相关信息
                gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,A.计算单位 AS 单位,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                          " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_慈溪农医, CLng(!收费细目ID))
                str项目名称 = Nvl(rsItem!名称)
                str医保编码 = Nvl(rsItem!项目编码)
                str医保编码 = get小儿用药编码(lng病人ID, !收费细目ID, str医保编码)
                str规格 = Nvl(rsItem!规格)
                str单位 = Nvl(rsItem!单位)
                If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
        
                '处方明细上传入参
            '    Jylx    varchar(1)  就医类型(=0 门诊；=1 住院)
            '    Jyhm    varchar(20) 就医号码(入院登记号或门诊登记号)
            '    Recordcount numeric(10) 记录条数(＜=100)
            '    xmdm[i] varchar(10) 项目代码(新农医项目代码)
            '    zxxmmc[i]   varchar(80) 中心项目名称(医院项目名称)
            '    xmdw[i] varchar(10) 项目单位
            '    xmdj[i] numeric(10,4)   项目单价
            '    xmsl[i] numeric(10,4)   项目数量
            '    xmje[i] numeric(10,4)   项目金额
            '    xmgg[i] varchar(50) 项目规格
            '    yzrq[i] Datetime    医嘱日期或门诊日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
                
                intCOUNT = intCOUNT + 1
                StrInput = StrInput & mstrSplit & "xmdm[" & intCOUNT & "]=" & str医保编码 & mstrSplit & "zxxmmc[" & intCOUNT & "]=" & ToVarchar(str项目名称, 80) & mstrSplit & _
                    "xmdw[" & intCOUNT & "]=" & ToVarchar(str单位, 10) & mstrSplit & "xmdj[" & intCOUNT & "]=" & Format(!价格, mstrPriceFormat) & mstrSplit & _
                    "xmsl[" & intCOUNT & "]=" & Format(!数量, mstrAmountFormat) & mstrSplit & "xmje[" & intCOUNT & "]=" & Format(!实收金额, mstrMoneyFormat) & mstrSplit & _
                    "xmgg[" & intCOUNT & "]=" & ToVarchar(str规格, 50) & mstrSplit & "yzrq[" & intCOUNT & "]=" & Format(!登记时间, mstrDateFormat)
                
                gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                If intCOUNT = 20 Then
                    '说明处方明细已组装好了，准备上传
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(lng病人ID, lng主页ID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
                    If 调用接口_慈溪农医() Then
                        gcnOracle.CommitTrans
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                    intCOUNT = 0
                    StrInput = ""
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If intCOUNT <> 0 Then
        '说明处方明细已组装好了，准备上传
        StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & GetKey(lng病人ID, lng主页ID) & mstrSplit & "Recordcount=" & intCOUNT & StrInput
        Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
        If 调用接口_慈溪农医() Then
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    blnTrans = False

    处方上传_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function 住院虚拟结算_慈溪农医(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim StrInput As String
    Dim strKey As String
    Dim blnTrans As Boolean
    Dim strCardNO As String, str出院日期 As String, str出院诊断 As String
    Dim intBalance As Integer, intRecords As Integer, intCOUNT As Integer
    Dim lng主页ID As Long, dbl承担比例 As Double
    Dim dbl帐户支付 As Double, dbl现金 As Double, dbl医保基金 As Double, dbl起付线 As Double
    Dim str项目名称 As String, str医保编码 As String, str规格 As String, str单位 As String
    Dim rsItem As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim rs明细1 As New ADODB.Recordset
    On Error GoTo errHand
    
    '取结算类型
    gstrSQL = "Select 卡号,业务类型,Nvl(退休证号,0) 承担比例 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取结算类型", TYPE_慈溪农医, lng病人ID)
    intBalance = Val(Right(rsItem!业务类型, 1)) - 1
    dbl承担比例 = Val(rsItem!承担比例)
    strCardNO = rsItem!卡号

    '取主页ID
    gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID")
    lng主页ID = rsItem!主页ID
    strKey = GetKey(lng病人ID, lng主页ID)
    
    '取出院日期
    gstrSQL = "Select 出院日期 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "取出院日期", lng病人ID, lng主页ID)
    str出院日期 = Format(rsItem!出院日期, mstrDateFormat)
    If str出院日期 = "" Then str出院日期 = Format(zlDatabase.Currentdate, mstrDateFormat)
    str出院诊断 = 获取入出院诊断(lng病人ID, lng主页ID, False, True, False)

    '提取本次费用明细
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价,A.开单人 AS 医生,A.登记时间 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND Nvl(A.实收金额,0)<>0 " & _
              "        And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "  Order by A.病人ID,A.发生时间"
    Set rs明细 = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", lng病人ID, lng主页ID, TYPE_慈溪农医)

 '提取本次费用明细1
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.收费类别,A.记录性质,A.记录状态,A.序号,A.收费细目ID,C.项目编码 AS 医保项目编码,B.编码,B.名称,A.实收金额 AS 金额" & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 单价,A.开单人 AS 医生,A.登记时间 " & _
              "  From 住院费用记录 A,收费细目 B,保险支付项目 C " & _
              "  where A.病人ID=[1] and A.主页ID=[2] and A.记帐费用=1 And A.操作员姓名 is not null AND Nvl(A.实收金额,0)<>0 " & _
              "        And Nvl(A.记录状态,0)<>0 and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类= [3]" & _
              "  Order by A.病人ID,A.发生时间"
    Set rs明细1 = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次费用明细", lng病人ID, lng主页ID, TYPE_慈溪农医)

    With rs明细1
        '求费用总额
        gComInfo_慈溪农医.总费用 = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!金额, 0) <> 0 Then
                intRecords = intRecords + 1
                gComInfo_慈溪农医.总费用 = gComInfo_慈溪农医.总费用 + !金额
            End If
            .MoveNext
        Loop
        
    End With
  
    gcnOracle.BeginTrans
    blnTrans = True
    With rs明细
        Do While Not .EOF
            If Nvl(!金额, 0) <> 0 Then
               '提取收费细目的相关信息
                gstrSQL = " Select A.类别 AS 收费类别,A.名称,A.规格,A.计算单位 AS 单位,B.项目编码 From 收费细目 A,保险支付项目 B" & _
                          " Where A.ID=B.收费细目ID(+) And B.险类(+)=[1] And A.ID=[2]"
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSQL, "提取项目信息", TYPE_慈溪农医, CLng(!收费细目ID))
                str项目名称 = Nvl(rsItem!名称)
                str医保编码 = Nvl(rsItem!项目编码)
                str医保编码 = get小儿用药编码(lng病人ID, !收费细目ID, str医保编码)
                str规格 = Nvl(rsItem!规格)
                str单位 = Nvl(rsItem!单位)
                If InStr(1, str规格, "|") <> 0 Then str规格 = Mid(str规格, 1, InStr(1, str规格, "|") - 1)
        
                '处方明细上传入参
            '    Jylx    varchar(1)  就医类型(=0 门诊；=1 住院)
            '    Jyhm    varchar(20) 就医号码(入院登记号或门诊登记号)
            '    Recordcount numeric(10) 记录条数(＜=100)
            '    xmdm[i] varchar(10) 项目代码(新农医项目代码)
            '    zxxmmc[i]   varchar(80) 中心项目名称(医院项目名称)
            '    xmdw[i] varchar(10) 项目单位
            '    xmdj[i] numeric(10,4)   项目单价
            '    xmsl[i] numeric(10,4)   项目数量
            '    xmje[i] numeric(10,4)   项目金额
            '    xmgg[i] varchar(50) 项目规格
            '    yzrq[i] Datetime    医嘱日期或门诊日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
                
                intCOUNT = intCOUNT + 1
                StrInput = StrInput & mstrSplit & "xmdm[" & intCOUNT & "]=" & str医保编码 & mstrSplit & "zxxmmc[" & intCOUNT & "]=" & ToVarchar(str项目名称, 80) & mstrSplit & _
                    "xmdw[" & intCOUNT & "]=" & ToVarchar(str单位, 10) & mstrSplit & "xmdj[" & intCOUNT & "]=" & Format(!单价, mstrPriceFormat) & mstrSplit & _
                    "xmsl[" & intCOUNT & "]=" & Format(!数量, mstrAmountFormat) & mstrSplit & "xmje[" & intCOUNT & "]=" & Format(!金额, mstrMoneyFormat) & mstrSplit & _
                    "xmgg[" & intCOUNT & "]=" & ToVarchar(str规格, 50) & mstrSplit & "yzrq[" & intCOUNT & "]=" & Format(!登记时间, mstrDateFormat)
                    
                gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                
                If intCOUNT = 20 Then
                    '说明处方明细已组装好了，准备上传
                    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & strKey & mstrSplit & "Recordcount=" & intCOUNT & StrInput
                    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
                    If Not 调用接口_慈溪农医() Then
                        gcnOracle.RollbackTrans
                        Exit Function
                    End If
                    gcnOracle.CommitTrans
                    gcnOracle.BeginTrans
                    intCOUNT = 0
                    StrInput = ""
                End If
            End If
            .MoveNext
        Loop
    
        If intCOUNT <> 0 Then
            '说明处方明细已组装好了，准备上传
            StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & strKey & mstrSplit & "Recordcount=" & intCOUNT & StrInput
            Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_UploadDetail, StrInput)
            If Not 调用接口_慈溪农医() Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
            gcnOracle.CommitTrans
        Else
            gcnOracle.RollbackTrans
        End If
    End With
    blnTrans = False
    
    '入参：
    'Jsbj    varchar(1)  结算标记(=0 预算 =1 结算)
    'Jslx    Varchar(1)  结算类型(=0 普通 =1 交通事故 =2 大病救助 =3 难产 =4 其他)
    'Cdbl    Numeric(10,2)   承担比例(交通事故)
    'Rydjh   varchar(20) 入院登记号，在医院数据库唯一
    'Kzhm    varchar(30) 卡证号码
    'Cyrq    Datetime    出院日期(固定格式19位：yyyy-mm-dd hh:mm:ss)
    'Cyzdsm  varchar(254)    出院诊断说明
    'Recordcount Long    明细记录数
    'Fpze    numeric(10,2)   发票总额
    'Czy Varchar(10) 操作员姓名
    '出参：
    'Returncode  Long    如果返回0表示成功
    'Returninfo  varchar(50) 对应的错误提示
    'Fpze    numeric(10,2)   发票总额
    'Ylzl    numeric(10,2)   乙类自理
    'Blzf    numeric(10,2)   丙类自费
    'Qfbz    numeric(10,2)   起付标准
    'Sbje    numeric(10,2)   报销金额
    'Zhzf    numeric(10,2)   帐户支付
    'Grzf    numeric(10,2)   个人自付
    'Zhye    numeric(10,2)   帐户余额
    'Dnbxlj  numeric(10,2)   当年统筹报销累计
    'Cfdkbje Numeric(10,2)   超封顶有效金额
    'Jsdh    Number(15)  结算单号
    'Kbje    Number(10,2)    本次可报金额
    'grzyljkb    Number(10,2)    当年累计可报金额(包含本次)
    'dc  Char(10)    补偿档次
    'Ld_bcfd[1]  Number(10,2)    分段1
    'Ld_bckbje[1]    Number(10,2)    本段可报金额
    'Ld_fdbxje[1]    Number(10,2)    本段实报金额
    'Ld_bcfd[2]  Number(10,2)    分段2
    'Ld_bckbje[2]    Number(10,2)    本段可报金额
    'Ld_fdbxje[2]    Number(10,2)    本段实报金额
    'Ld_bcfd[3]  Number(10,2)    分段3
    'Ld_bckbje[3]    Number(10,2)    本段可报金额
    'Ld_fdbxje[3]    Number(10,2)    本段实报金额
    'Ld_bcfd[4]  Number(10,2)    分段4
    'Ld_bckbje[4]    Number(10,2)    本段可报金额
    'Ld_fdbxje[4]    Number(10,2)    本段实报金额
    'Ld_bcfd[5]  Number(10,2)    分段5
    'Ld_bckbje[5]    Number(10,2)    本段可报金额
    'Ld_fdbxje[5]    Number(10,2)    本段实报金额

    '注：(1)发票总额＝报销金额＋个人自付
    '现金金额 = 个人自付
    '(2)住院结算前，所有的明细必须已经保存
    '(3)门诊结算交易和住院交易是同样的处理方法和返回值
    StrInput = "Jsbj=0" & mstrSplit & "Jslx=" & intBalance & mstrSplit & "Cdbl=" & IIf(intBalance = 1, dbl承担比例 / 100, 0) & mstrSplit & _
        "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & "Cyrq=" & str出院日期 & mstrSplit & _
        "Cyzdsm=" & str出院诊断 & mstrSplit & "recordcount=" & intRecords & mstrSplit & "Fpze=" & gComInfo_慈溪农医.总费用 & mstrSplit & _
        "Czy=" & UserInfo.姓名
    gComInfo_慈溪农医.结算串 = StrInput
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_InBalance, StrInput)
    If Not 调用接口_慈溪农医() Then Exit Function
    
    dbl医保基金 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(6), "=")(1))
    dbl帐户支付 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(7), "=")(1))
    dbl现金 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(8), "=")(1))

    住院虚拟结算_慈溪农医 = "个人帐户;" & dbl帐户支付 & ";0|医保基金;" & dbl医保基金 & ";0"
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
End Function

Public Function 住院结算_慈溪农医(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    Dim StrInput As String
    Dim str结算流水号 As String
    Dim lng主页ID As Long
    Dim dbl现金 As Double, dbl医保基金 As Double, dbl起付线 As Double, dbl帐户支付 As Double
    Dim dbl本次可报 As Double, dbl本年累计 As Double, str补偿档次 As String
    Dim dbl分段1 As Double, dbl分段1可报 As Double, dbl分段1实报 As Double
    Dim dbl分段2 As Double, dbl分段2可报 As Double, dbl分段2实报 As Double
    Dim dbl分段3 As Double, dbl分段3可报 As Double, dbl分段3实报 As Double
    Dim dbl分段4 As Double, dbl分段4可报 As Double, dbl分段4实报 As Double
    Dim dbl分段5 As Double, dbl分段5可报 As Double, dbl分段5实报 As Double
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    '必须先出院，才能进行结算
    If Not 医保病人已经出院(lng病人ID) Then
        Err.Raise 9000, gstrSysName, "必须先出院，才能进行结算！", vbInformation, gstrSysName
        Exit Function
    End If

    '取主页ID
    gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    StrInput = gComInfo_慈溪农医.结算串
    StrInput = Replace(StrInput, "Jsbj=0", "Jsbj=1")
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_InBalance, StrInput)
    If Not 调用接口_慈溪农医() Then Exit Function
    '出参：
    'Returncode  Long    如果返回0表示成功
    'Returninfo  varchar(50) 对应的错误提示
    'Fpze    numeric(10,2)   发票总额
    'Ylzl    numeric(10,2)   乙类自理
    'Blzf    numeric(10,2)   丙类自费
    'Qfbz    numeric(10,2)   起付标准
    'Sbje    numeric(10,2)   报销金额
    'Zhzf    numeric(10,2)   帐户支付
    'Grzf    numeric(10,2)   个人自付
    'Zhye    numeric(10,2)   帐户余额
    'Dnbxlj  numeric(10,2)   当年统筹报销累计
    'Cfdkbje Numeric(10,2)   超封顶有效金额
    'Jsdh    Number(15)  结算单号
    'Kbje    Number(10,2)    本次可报金额
    'grzyljkb    Number(10,2)    当年累计可报金额(包含本次)
    'dc  Char(10)    补偿档次
    'Ld_bcfd[1]  Number(10,2)    分段1
    'Ld_bckbje[1]    Number(10,2)    本段可报金额
    'Ld_fdbxje[1]    Number(10,2)    本段实报金额
    'Ld_bcfd[2]  Number(10,2)    分段2
    'Ld_bckbje[2]    Number(10,2)    本段可报金额
    'Ld_fdbxje[2]    Number(10,2)    本段实报金额
    'Ld_bcfd[3]  Number(10,2)    分段3
    'Ld_bckbje[3]    Number(10,2)    本段可报金额
    'Ld_fdbxje[3]    Number(10,2)    本段实报金额
    'Ld_bcfd[4]  Number(10,2)    分段4
    'Ld_bckbje[4]    Number(10,2)    本段可报金额
    'Ld_fdbxje[4]    Number(10,2)    本段实报金额
    'Ld_bcfd[5]  Number(10,2)    分段5
    'Ld_bckbje[5]    Number(10,2)    本段可报金额
    'Ld_fdbxje[5]    Number(10,2)    本段实报金额
    
    dbl起付线 = Val(Format(Split(Split(gstrOutput_慈溪农医, mstrSplit)(5), "=")(1), "#0.00"))
    dbl医保基金 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(6), "=")(1))
    dbl帐户支付 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(7), "=")(1))
    dbl现金 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(8), "=")(1))
    str结算流水号 = Split(Split(gstrOutput_慈溪农医, mstrSplit)(12), "=")(1)
    
    dbl本次可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(13), "=")(1))
    dbl本年累计 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(14), "=")(1))
    str补偿档次 = Split(Split(gstrOutput_慈溪农医, mstrSplit)(15), "=")(1)
    dbl分段1 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(16), "=")(1))
    dbl分段1可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(17), "=")(1))
    dbl分段1实报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(18), "=")(1))
    dbl分段2 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(19), "=")(1))
    dbl分段2可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(20), "=")(1))
    dbl分段2实报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(21), "=")(1))
    dbl分段3 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(22), "=")(1))
    dbl分段3可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(23), "=")(1))
    dbl分段3实报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(24), "=")(1))
    dbl分段4 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(25), "=")(1))
    dbl分段4可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(26), "=")(1))
    dbl分段4实报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(27), "=")(1))
    dbl分段5 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(28), "=")(1))
    dbl分段5可报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(29), "=")(1))
    dbl分段5实报 = Val(Split(Split(gstrOutput_慈溪农医, mstrSplit)(30), "=")(1))
    
    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_慈溪农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        gComInfo_慈溪农医.总费用 & "," & dbl现金 & ",0," & dbl起付线 & "," & dbl医保基金 & ",0,0," & _
        dbl帐户支付 & ",'" & str结算流水号 & "'," & lng主页ID & ",null,'" & GetKey(lng病人ID, lng主页ID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    gstrSQL = "ZL_结算附加信息_INSERT(" & lng结帐ID & ",'" & str结算流水号 & "'," & dbl本次可报 & "," & dbl本年累计 & "," & _
        "'" & str补偿档次 & "'," & dbl分段1 & "," & dbl分段1可报 & "," & dbl分段1实报 & "," & _
        dbl分段2 & "," & dbl分段2可报 & "," & dbl分段2实报 & "," & _
        dbl分段3 & "," & dbl分段3可报 & "," & dbl分段3实报 & "," & _
        dbl分段4 & "," & dbl分段4可报 & "," & dbl分段4实报 & "," & _
        dbl分段5 & "," & dbl分段5可报 & "," & dbl分段5实报 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算附加信息")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "将结帐记录打上上传标志")

    住院结算_慈溪农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_慈溪农医(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    Dim StrInput As String
    Dim strCardNO As String
    Dim lng冲销ID As Long
    Dim lng病人ID As Long, lng主页ID As Long, lng主页ID_当前 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsBalance As New ADODB.Recordset
    On Error GoTo errHand

    '取冲销ID
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!ID

    '取结算流水号
    gstrSQL = "Select * From 保险结算记录 Where 性质=2 And 记录ID=[1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(gstrSQL, "取结算流水号", lng结帐ID)
    If rsBalance.RecordCount = 0 Then
        MsgBox "没有找到原始结算记录，无法进行住院结算冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rsBalance!病人ID
    lng主页ID = rsBalance!主页ID
    
    '取当前主页ID
    gstrSQL = "Select Nvl(住院次数,0) AS 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    lng主页ID_当前 = rsTemp!主页ID
    
    If lng主页ID <> lng主页ID_当前 Then
        MsgBox "不能冲销上次住院期间的结算单，请先撤销本次入院登记！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取结算类型
    gstrSQL = "Select 卡号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取结算类型", TYPE_慈溪农医, lng病人ID)
    strCardNO = rsTemp!卡号

    '调用结算冲销
'    Jylx    varchar(1)  就医类型(=0 门诊作废 = 1 取消出院)
'    Jyhm    varchar(20) 就医号码
'    Kzhm    varchar(30) 卡证号码
'    Tfrq    Datetime    就诊日期(退费日期) (固定格式19位：yyyy-mm-dd hh:mm:ss)
'    Czy Varchar(10) 操作员姓名
    StrInput = "Jylx=1" & mstrSplit & "Jyhm=" & rsBalance!备注 & mstrSplit & _
        "Kzhm=" & strCardNO & mstrSplit & "Tfrq=" & Format(zlDatabase.Currentdate, mstrDateFormat) & mstrSplit & _
       "Czy=" & UserInfo.姓名
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_BalanceCancel, StrInput)
    If Not 调用接口_慈溪农医() Then Exit Function

    '保存本次结算情况
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & TYPE_慈溪农医 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsBalance!发生费用金额, 0) & "," & -1 * Nvl(rsBalance!全自付金额, 0) & "," & -1 * Nvl(rsBalance!首先自付金额, 0) & "," & -1 * Nvl(rsBalance!进入统筹金额, 0) & "," & -1 * Nvl(rsBalance!统筹报销金额, 0) & ",0,0," & _
        -1 * Nvl(rsBalance!个人帐户支付, 0) & ",null,null,null,'" & Nvl(rsBalance!备注) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "门诊结算冲销")
    
    gstrSQL = "Select * From 结算附加信息 Where 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取结算附加信息", lng结帐ID)
    If rsTemp.RecordCount <> 0 Then
        gstrSQL = "ZL_结算附加信息_INSERT(" & lng冲销ID & ",'" & Nvl(rsTemp!结算单号) & "'," & -1 * Nvl(rsTemp!本次可报金额, 0) & "," & -1 * Nvl(rsTemp!当年累计可报金额, 0) & "," & _
            "'" & Nvl(rsTemp!补偿档次) & "'," & -1 * Nvl(rsTemp!分段1, 0) & "," & -1 * Nvl(rsTemp!分段1可报, 0) & "," & -1 * Nvl(rsTemp!分段1实报, 0) & "," & _
            -1 * Nvl(rsTemp!分段2, 0) & "," & -1 * Nvl(rsTemp!分段2可报, 0) & "," & -1 * Nvl(rsTemp!分段2实报, 0) & "," & _
            -1 * Nvl(rsTemp!分段3, 0) & "," & -1 * Nvl(rsTemp!分段3可报, 0) & "," & -1 * Nvl(rsTemp!分段3实报, 0) & "," & _
            -1 * Nvl(rsTemp!分段4, 0) & "," & -1 * Nvl(rsTemp!分段4可报, 0) & "," & -1 * Nvl(rsTemp!分段4实报, 0) & "," & _
            -1 * Nvl(rsTemp!分段5, 0) & "," & -1 * Nvl(rsTemp!分段5可报, 0) & "," & -1 * Nvl(rsTemp!分段5实报, 0) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存结算附加信息")
    End If

    住院结算冲销_慈溪农医 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function IsYBPatient(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '判断指定病人本次是否以医保身份就诊
    gstrSQL = " Select 1 From 病案主页 Where 险类=" & TYPE_慈溪农医 & " And (病人ID,主页ID) IN " & _
              "     (Select 病人ID,住院次数 From 病人信息 Where 病人ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断指定病人本次是否以医保身份就诊", lng病人ID)
    IsYBPatient = (rsTemp.RecordCount <> 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetKey(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim strKey As String
    '如果主页ID小于一位，则前面加一个零
    strKey = lng主页ID
    If Len(strKey) = 1 Then strKey = "0" & strKey
    GetKey = lng病人ID & strKey
End Function

Private Sub UploadDetail(ByVal int记录性质 As Integer, ByVal int记录状态 As Integer, ByVal strNO As String, ByVal lng病人ID As Long)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select NO,序号,记录性质,记录状态 From 住院费用记录" & _
              " Where 记录性质=[1] And 记录状态=[2] And NO=[3] And 病人ID=[4] And Nvl(是否上传,0)=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该病人的处方明细", int记录性质, int记录状态, strNO, lng病人ID)
    
    '将指定病人的处方明细打上传标记
    With rsTemp
        Do While Not .EOF
            gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "打上传标志")
            .MoveNext
        Loop
    End With
End Sub

Public Function 更新在院信息_慈溪农医(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln出院 As Boolean = False) As Boolean
    Dim StrInput As String
    Dim strKey As String                    '入院登记号
    Dim strCardNO As String, lngDisease As Long '病人的医疗卡号及病种ID
    Dim strRegistCode As String             '挂号单号
    Dim strInHospitalDate As String         '入院日期
    Dim strRegisterOffice As String         '就诊科室
    Dim strDiseaseCode As String            '病种代码
    Dim strDiagnose As String               '入院诊断
    Dim strRegisterDoctor As String         '医生
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strKey = GetKey(lng病人ID, lng主页ID)
    '取病人的医疗卡号
    gstrSQL = "Select 卡号,Nvl(病种ID,0) 病种ID From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的医疗卡号", TYPE_慈溪农医, lng病人ID)
    strCardNO = rsTemp!卡号
    lngDisease = rsTemp!病种ID
    
    '取科室与医生
    gstrSQL = " Select A.入院日期,B.名称 科室,A.住院医师 医生 From 病案主页 A,部门表 B " & _
              " Where A.病人ID=[1] And A.主页ID=[2] And A." & IIf(bln出院 = False, "入院科室ID", "出院科室ID") & "=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取科室与医生", lng病人ID, lng主页ID)
    strInHospitalDate = Format(rsTemp!入院日期, mstrDateFormat)
    strRegisterDoctor = Nvl(rsTemp!医生)
    strRegisterOffice = Nvl(rsTemp!科室)
    '取病种代码
    gstrSQL = "Select 编码 From 疾病编码目录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病种代码", lngDisease)
    If rsTemp.RecordCount = 1 Then strDiseaseCode = rsTemp!编码
    '取入院诊断
    strDiagnose = 获取入出院诊断(lng病人ID, lng主页ID, True, True, False)
    
    '入参：
'    Rydjh   varchar(20) 入院登记号，在医院数据库中的唯一号码
'    Kzhm    varchar(30) 卡证号码
'    Jzks    varchar(20) 就诊科室
'    Ysxm    varchar(10) 医生姓名
'    Bzdm    Varchar(20) 病种代码
'    Ryzdsm  varchar(254)    入院诊断说明
'    Ryrq    Datetime    入院日期(固定格式19位：yyyy-mm-dd hh:mm:ss)，以下相同
'    Czy Varchar(10) 操作员姓名
    '返回参数:
'    Returncode  Long    如果返回0表示成功
'    Returninfo  varchar(50) 对应的错误提示
    StrInput = "Rydjh=" & strKey & mstrSplit & "Kzhm=" & strCardNO & mstrSplit & _
        "Jzks=" & strRegisterOffice & mstrSplit & "Ysxm=" & strRegisterDoctor & mstrSplit & _
        "Bzdm=" & strDiseaseCode & mstrSplit & "Ryzdsm=" & strDiagnose & mstrSplit & _
        "Ryrq=" & strInHospitalDate & mstrSplit & "Czy=" & UserInfo.姓名
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_ModifyInfo, StrInput)
    If Not 调用接口_慈溪农医() Then Exit Function
    
    更新在院信息_慈溪农医 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub 更新病种_慈溪农医(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    Dim lng疾病ID As Long
    lng疾病ID = frm病种选择_慈溪农医.ChooseDisease(lng病人ID)
    
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_慈溪农医 & ",'病种ID','''" & lng疾病ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种ID")
    
    Call 更新在院信息_慈溪农医(lng病人ID, lng主页ID)
End Sub

Public Function get小儿用药编码(ByVal lng病人ID As Long, ByVal lng收费细目ID As Long, ByVal str编码 As String) As String
    Dim rsTemp As New ADODB.Recordset
    '提取小儿用药编码
    get小儿用药编码 = str编码
    
    '判断本次入院是否是小儿用药的方式入院
    gstrSQL = "Select Nvl(小儿用药,0) AS 小儿用药 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断本次入院是否是小儿用药的方式入院", lng病人ID, TYPE_慈溪农医)
    If rsTemp.RecordCount = 0 Then Exit Function
    If rsTemp!小儿用药 = 0 Then Exit Function
    
    '提取小儿用药编码并返回
    gstrSQL = "Select 附注 From 保险支付项目 Where 险类=[1] ANd 收费细目ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取小儿用药编码", TYPE_慈溪农医, lng收费细目ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If InStr(1, Nvl(rsTemp!附注), "|") = 0 Then Exit Function
    get小儿用药编码 = Split(rsTemp!附注, "|")(0)
    If get小儿用药编码 = "" Then get小儿用药编码 = str编码
End Function
