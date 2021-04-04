Attribute VB_Name = "mdl重大校园卡"
Option Explicit
#Const gverControl = 99
Public gintComport_重大校园卡 As Long
Private Type 病人身份
    串口设备            As Long
    卡类型              As Long
    卡流水号            As Long
    卡号                As String
    卡消费分组          As Long
    卡有效期            As String
    姓名                As String
    注册日期            As String
    身份证号            As String
    电子钱包1余额       As Long   '以分为单位
    电子钱包2余额       As Long   '以分为单位
    卡固有序号          As Long
    上次交易流水号      As Integer
    上次交易金额        As Long
    上次交易时间        As String
    日交易累计金额      As Long
    上次交易终端号      As Integer
    卡等待时间          As Integer
    性别                As String
    年龄                As Long
    出生日期            As String
End Type
Public g病人身份_重大校园卡 As 病人身份
Public gdbl交易限额_重大校园卡 As Double
Const mbln测试 As Boolean = False
'--校园卡声明
Public Declare Function CloseComm Lib "cqcardtsgl.dll" (ByVal icdev As Long) As Integer
Public Declare Function OpenComm Lib "cqcardtsgl.dll" (ByVal CommPort As Long) As Long

Public Declare Function Query_Pos_UserCard Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByRef CardType As Long, ByRef CardSerno As Long, _
             ByVal Cardno As String, ByRef CardGroup As Long, ByVal CardDate As String, ByVal Name As String, _
                 ByVal RegDate As String, ByVal Passport As String, ByRef Account0 As Long, ByRef Account1 As Long, _
                 ByRef Serno As Long, ByRef LastSerno As Integer, ByRef LastAccount As Long, ByVal LastTime As String, _
                 ByRef DayAccount As Long, ByRef LastTermno As Integer, ByVal WAITTIME As Integer) As Integer
                 
Public Declare Function extSys_IsInBlackList Lib "cqcardtsgl.dll" (ByVal ulCardSerno As Long) As Integer
Public Declare Function extSys_WithDraw Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByVal ulCardSerno As Long, ByVal ulPurseNo As Long, ByVal p_TransCode As String, ByVal p_Amount As Long) As Integer
Public Declare Function rf_beep Lib "cqcardtsgl.dll" (ByVal icdev As Long, ByVal WAITTIME As Integer) As Integer

Public Const G_WAIT_TIME = 30
Public Const G_WAIT_TIME1 = 10
Private mblnInit As Boolean '是否初始化

Public Function 医保初始化_重大校园卡() As Boolean
    '功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
    '返回：初始化成功，返回true；否则，返回false
  
    Dim strReg As String
    Dim lngHandle As Long
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    If mblnInit = True Then
        医保初始化_重大校园卡 = True
        Exit Function
    End If
    
    '设置端口号
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)


    If Val(strReg) = 0 Then
        gintComport_重大校园卡 = 0
    Else
        gintComport_重大校园卡 = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    gstrSQL = "Select * From 保险参数 where 参数名 ='交易限额' and 险类=" & TYPE_重大校园卡
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取参数"
    If Not rsTemp.EOF Then
            gdbl交易限额_重大校园卡 = Val(Nvl(rsTemp!参数值))
    Else
        
    End If
    gdbl交易限额_重大校园卡 = IIf(gdbl交易限额_重大校园卡 <= 0, 2000, gdbl交易限额_重大校园卡)
    If mbln测试 Then
        医保初始化_重大校园卡 = True
        Exit Function
    End If
    If g病人身份_重大校园卡.串口设备 <> 0 Then
       lngHandle = CloseComm(g病人身份_重大校园卡.串口设备)
    End If
    lngHandle = OpenComm(gintComport_重大校园卡)
    If lngHandle < 0 Then
        ShowMsgbox ("串口打开失败,请检查设备是否连接正常!")
        Exit Function
    End If
    g病人身份_重大校园卡.串口设备 = lngHandle
    mblnInit = True
    
    医保初始化_重大校园卡 = True
End Function

Public Function 医保终止_重大校园卡() As Boolean
    Dim intReutn As Integer
    mblnInit = False
    '关闭
    If mbln测试 Then
        医保终止_重大校园卡 = True
        Exit Function
    End If
    intReutn = CloseComm(g病人身份_重大校园卡.串口设备)
    If intReutn <> 0 Then
     '  ShowMsgbox GetErrInfo(CStr(intReutn))
       Exit Function
    End If
  医保终止_重大校园卡 = True
End Function
Public Function 医保设置_重大校园卡(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    医保设置_重大校园卡 = frmSet重大校园卡.ShowME(lng险类, lng医保中心)
End Function
Public Function 身份标识_重大校园卡2(ByVal strCard As String, ByVal strPass As String, Optional lng病人ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    身份标识_重大校园卡2 = frmIdentify重大校园卡.GetPatient(3, lng病人ID, True)
End Function
Public Function 身份标识_重大校园卡(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim str备注 As String, RSPATIENT As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    '/**?
    身份标识_重大校园卡 = frmIdentify重大校园卡.GetPatient(bytType, lng病人ID)
End Function


Public Function 个人余额_重大校园卡(ByVal lng病人ID As Long, ByVal bytplance As Byte) As Currency
    '功能: 根据病人id取出余额
    '参数: 病人id
    '返回: 返回个人帐户余额
    '读卡失败则退出
    'bytplance=10门诊
    Dim rsTmp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    
    #If gverControl >= 5 Then
        gstrSQL = "select  费用余额 from 病人余额 where 病人id= " & lng病人ID & " And 性质=1 And 类型=" & IIf(bytplance = 10, "1", "2")
    #Else
        gstrSQL = "select  费用余额 from 病人余额 where 病人id= " & lng病人ID & " And 性质=1 "
    #End If
    zlDatabase.OpenRecordset rsTmp, gstrSQL, "获取余额"
    If Not rsTmp.EOF Then
        个人余额_重大校园卡 = Nvl(rsTmp!费用余额, 0)
    End If
    
    If 个人余额_重大校园卡 < (g病人身份_重大校园卡.电子钱包1余额 + g病人身份_重大校园卡.电子钱包2余额) / 100 Then
        个人余额_重大校园卡 = (g病人身份_重大校园卡.电子钱包1余额 + g病人身份_重大校园卡.电子钱包2余额) / 100
    End If
    If bytplance = 10 Then
        If 个人余额_重大校园卡 < gdbl交易限额_重大校园卡 Then
            个人余额_重大校园卡 = gdbl交易限额_重大校园卡
        End If
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
        个人余额_重大校园卡 = 0
End Function

Public Function 门诊虚拟结算_重大校园卡(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim curTotal As Currency
    Dim dbl余额 As Double
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '明细字段
    '   病人ID,收费类别,收据费目,计算单位,开单人,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,摘要,是否急诊
    With rs明细
        '取出本次发生费用的金额合计
        Do While Not .EOF
            '先判断是否都设置了医保对应项目编码
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
    
    dbl余额 = g病人身份_重大校园卡.电子钱包1余额 + g病人身份_重大校园卡.电子钱包2余额
    
    If curTotal * 100 > dbl余额 Then
        dbl余额 = dbl余额 / 100
    Else
        dbl余额 = curTotal
    End If
    
    str结算方式 = "个人帐户;" & Format(dbl余额, "###0.00;-###0.00;0;0") & ";1"  '不充许修改
    门诊虚拟结算_重大校园卡 = True
End Function
Public Function 门诊结算_重大校园卡(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim lng病人ID As Long
    门诊结算_重大校园卡 = Set门诊结算或冲销(False, lng结帐ID, cur个人帐户, lng病人ID, strSelfNo)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Private Function Set门诊结算或冲销(ByVal bln冲销 As Boolean, lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long, strSelfNo As String) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim curTotal As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim strTmp As String
    Dim int业务 As Integer
    Dim lng冲销ID As Long
    Dim strNO As String
    Dim lng记录性质 As Long
    Dim lngTmp As Long

    int业务 = IIf(bln冲销, 1, 0)
     Set门诊结算或冲销 = False

    If bln冲销 Then
        '重新读卡
'
'        If GetUserCardInfor = False Then
'            Exit Function
'        End If
        
        '验证是否为该病人的IC卡
        gstrSQL = "Select * From  保险帐户 where 病人id=" & lng病人ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的卡流水号"
        If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "该病人在保险帐户中无记录!"
            Exit Function
        End If
        
        '确定退费记录
        '退费
          gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                    " where A.NO=B.NO and A.记录性质=B.记录性质  and A.记录状态=2 and B.结帐ID=[1]"
          Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊退费", lng结帐ID)
          If rsTemp.EOF Then
            Err.Raise 9000, gstrSysName, "不存在病人费用冲销记录!"
            Exit Function
          Else
            lng冲销ID = rsTemp("结帐ID")
          End If
          
    End If
    
    '打开本次结算明细记录
    gstrSQL = " " & _
        "  Select A.收费类别,a.病人ID,sum(nvl(A.结帐金额,0)) as 实收金额" & _
        "  From 门诊费用记录  A" & _
        "  Where A.记录状态<>0 and A.结帐ID=" & IIf(bln冲销, lng冲销ID, lng结帐ID) & " and  Nvl(A.附加标志,0)<>9 " & _
        "  Group by A.收费类别,a.病人id" & _
        "  Order by A.收费类别"
        
    zlDatabase.OpenRecordset rs明细, gstrSQL, "读取本次结帐费用明细"
    With rs明细
        Do While Not .EOF
            '计算总额,待用
            '扣款.
            lng病人ID = Nvl(!病人ID, 0)
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
    curTotal = IIf(bln冲销, -1, 1) * cur个人帐户
    If bln冲销 Then
    Else
        If Not 扣款_重大校园卡("0", Val(Format(curTotal, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
            
            '遇到中断无法回滚，建议使用总额进行扣款.
             Set门诊结算或冲销 = False
            Exit Function
        End If
    End If
    '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(1," & IIf(bln冲销, lng冲销ID, lng结帐ID) & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "保存门诊收费数据"
    Set门诊结算或冲销 = True
End Function
Public Function 门诊结算冲销_重大校园卡(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Err = 0
    On Error GoTo errHand:
    门诊结算冲销_重大校园卡 = Set门诊结算或冲销(True, lng结帐ID, cur个人帐户, lng病人ID, "")
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function

Public Function 个人帐户转预交_重大校园卡(lng预交ID As Long, curMoney As Currency, rs预交记录 As ADODB.Recordset) As Boolean
    '功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
    '参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    
    Dim cur金额 As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    gstrSQL = "select 病人id,nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "校园卡转预交", TYPE_重大校园卡, g病人身份_重大校园卡.卡流水号)
    If rsTemp.RecordCount > 0 Then
        If rsTemp("状态") <> 1 Then
            MsgBox "该医保病人尚未入院,不能执行个人帐户转预交交易！", vbInformation, gstrSysName
            Exit Function
        End If
        lng病人ID = Nvl(rsTemp!病人ID, 0)
    End If
    Err = 0
    On Error GoTo errHand:
    If curMoney <> 0 Then
        '扣款
        If 扣款_重大校园卡(" ", curMoney * 100, False) = False Then
            Exit Function
        End If
    End If
   '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curMoney & "," & curMoney & ",0,0,0,0," & _
        curMoney & ",NULL,NULL," & curMoney & "," & curMoney & ",Null," & 0 & "," & _
        curMoney & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "个人帐户转预交款"
    
    个人帐户转预交_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    个人帐户转预交_重大校园卡 = False
End Function


Public Function 个人帐户转预交冲销_重大校园卡(lng预交ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
    '参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    Dim curMoney As Currency
    curMoney = cur个人帐户
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curMoney & "," & curMoney & ",0,0,0,0," & _
        curMoney & ",NULL,NULL," & curMoney & "," & curMoney & ",Null," & 0 & "," & _
        curMoney & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "个人帐户转预交款"

    个人帐户转预交冲销_重大校园卡 = True
End Function

Public Function 住院虚拟结算_重大校园卡(rsExse As Recordset, ByVal lng病人ID As Long) As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '      字段:记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID, _
    '           收费类别,收费细目ID,收费名称,开单部门,规格,产地,数量,价格,金额,医生,登记时间, _
    '           是否上传,是否急诊,保险项目否,摘要
    
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
  
   Dim str结算方式  As String
    Dim dbl金额 As String
    dbl金额 = 0
    If GetUserCardInfor = False Then
        住院虚拟结算_重大校园卡 = ""
         Exit Function
    End If
    Do While Not rsExse.EOF
        dbl金额 = dbl金额 + Nvl(rsExse!金额, 0)
        rsExse.MoveNext
    Loop
    
    Dim dbl余额 As Double
    dbl余额 = g病人身份_重大校园卡.电子钱包1余额 '+ g病人身份_重大校园卡.电子钱包2余额
    
    If dbl金额 * 100 > dbl余额 Then
        dbl余额 = dbl余额 / 100
    Else
        dbl余额 = dbl金额
    End If
    
    str结算方式 = "个人帐户;" & Format(dbl余额, "###0.00;-###0.00;0;0") & ";1" '本次基本个人帐户支付,不充许修改

  
   住院虚拟结算_重大校园卡 = str结算方式
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 住院结算_重大校园卡(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean

    Dim lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
     '重新读卡
    If GetUserCardInfor() = False Then
        Exit Function
    End If
     
    On Error GoTo errHand
   
    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份 " & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=[1] And A.主页ID=B.住院次数 And A.病人ID=B.病人ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人入院时间", lng病人ID)
    lng主页ID = rsTemp!主页ID
   住院结算_重大校园卡 = 住院结算及冲帐_重大校园卡(False, lng病人ID, lng结帐ID, lng结帐ID, lng主页ID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function


Private Function 住院结算及冲帐_重大校园卡(ByVal bln冲销 As Boolean, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal 原结帐id As Long, ByVal lng主页ID As Long) As Boolean

    Dim rs明细 As New ADODB.Recordset
    Dim curTotal As Double
    Dim int业务 As Integer
    Dim cur帐户支付 As Double
    Dim dblTmp As Double
    Dim rsTemp As New ADODB.Recordset
    int业务 = IIf(bln冲销, 1, 0)
    
  
    '读取帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And not( a.记录性质  in(11,1)) and  B.险类=[1]" & _
        " And A.结算方式='个人帐户' And A.结帐ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户支付额", TYPE_重大校园卡, lng结帐ID)
    cur帐户支付 = 0
    If Not rsTemp.EOF Then
        cur帐户支付 = Nvl(rsTemp!个人帐户, 0)
    End If
    
    Err = 0
    On Error GoTo errHand:
    
    '住院应用保险支付大类中的住院比额
    gstrSQL = " " & _
        "  Select A.收费类别,sum(nvl(A.结帐金额,0)) as 实收金额" & _
        "  From 住院费用记录  A" & _
        "  Where a.记录状态<>0 and A.结帐ID=" & lng结帐ID & " and  Nvl(A.附加标志,0)<>9 " & _
        "  Group by A.收费类别" & _
        "  Order by A.收费类别"
        
    zlDatabase.OpenRecordset rs明细, gstrSQL, "提取住院结帐明细"
    With rs明细
        dblTmp = 0
        Do While Not .EOF
            '计算总额,待用
            '扣款
'
'            dblTmp = dblTmp + NVL(!实收金额)
'            If dblTmp > cur帐户支付 Then
'                '确定是否已经超过个人帐户
'                If cur帐户支付 - curTotal > 0 Then
'                    If Not 扣款_重大校园卡(NVL(!收费类别, "0"), (cur帐户支付 - curTotal) * 100, True) Then
'                        '
'                    End If
'                End If
'                curTotal = cur帐户支付
'                Exit Do
'            Else
                curTotal = curTotal + Round(Nvl(!实收金额, 0), 2)
''            End If
            .MoveNext
        Loop
    End With
    curTotal = cur帐户支付
    If bln冲销 = False Then
            If Not 扣款_重大校园卡("0", Val(Format(cur帐户支付, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
                '
                住院结算及冲帐_重大校园卡 = False
                Exit Function
            End If
    End If
    '保存结算记录
     '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & "," & lng主页ID & ",0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        cur帐户支付 & ",0,NULL,null,null" & _
         " )"
           
        zlDatabase.ExecuteProcedure gstrSQL, "保存住院结帐收费数据"
        住院结算及冲帐_重大校园卡 = True
        Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_重大校园卡(lng结帐ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo errHand
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "大连医保", lng结帐ID)
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    '重新读卡
    If GetUserCardInfor() = False Then
        Exit Function
    End If

    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "大连医保", lng结帐ID)
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "在保险结算记录中无该结算记录!"
        Exit Function
    End If
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    lng主页ID = Nvl(rsTemp!主页ID, 0)
        
        
    
    '验证是否为该病人的IC卡
    gstrSQL = "Select * From  保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
    If rsTemp.EOF Then
        Err.Raise 9000, gstrSysName, "该病人在保险帐户中无记录!"
        Exit Function
    End If
    
    If g病人身份_重大校园卡.卡号 <> Nvl(rsTemp!卡号) Then
        Err.Raise 9000, gstrSysName, "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
        Exit Function
    End If
    
    '调用撤销结算接口
    住院结算冲销_重大校园卡 = 住院结算及冲帐_重大校园卡(True, lng病人ID, lng冲销ID, lng结帐ID, lng主页ID)
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    Exit Function
End Function
Public Function 入院登记_重大校园卡(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHand
    
    '读取病人的相关保险信息

    gstrSQL = "select * From 保险帐户 where  险类=" & TYPE_重大校园卡 & "  and 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重大校园卡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    入院登记_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 入院登记撤销_重大校园卡(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
                
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    gstrSQL = " Select Count(*) Records From 住院费用记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "撤销入院检查", lng病人ID, lng主页ID)
    
    If rsTemp!Records <> 0 Then
        MsgBox "已经存在费用记录，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If

    
    On Error GoTo errHand
    
    '读取病人的相关保险信息

    gstrSQL = "select * From 保险帐户 where  险类=" & TYPE_重大校园卡 & "  and 病人id=" & lng病人ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "撤消入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重大校园卡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销入院登记")
    入院登记撤销_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 出院登记_重大校园卡(lng病人ID As Long, lng主页ID As Long) As Boolean
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_重大校园卡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "出院登记")
    出院登记_重大校园卡 = True
End Function
Public Function 出院登记撤销_重大校园卡(lng病人ID As Long, lng主页ID As Long) As Boolean
    '存在未结费用的病人才允许撤销HIS出院；否则认为已办理医保出院，不允许再办理HIS出院
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        MsgBox "医保已出院的病人不允许撤销出院！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_重大校园卡 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理撤销出院登记")
    出院登记撤销_重大校园卡 = True
End Function


Public Function 挂号结算_重大校园卡(ByVal lng结帐ID As Long) As Boolean
    Dim curTotal As Currency '上传费用总额
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    
    '先调门诊预结算,主要是取个人帐户支付额,再调门诊结算
    '保险参数中保存的是允许帐户支付的收入项目ID，挂号结算时判断，如果未设置，表明全自付，不需上传，否则仅上传那一笔明细
    
    On Error GoTo errHand
    
  gstrSQL = " " & _
        "  Select A.收费类别,A.病人id,sum(nvl(A.结帐金额,0)) as 实收金额" & _
        "  From 门诊费用记录  A" & _
        "  Where A.记录状态<>0 and  A.结帐ID=" & lng结帐ID & " and  Nvl(A.附加标志,0)<>9 " & _
        "  Group by A.收费类别,A.病人id" & _
        "  Order by A.收费类别"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取本次结帐费用明细"
    
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgbox "无任何费用记录发生!"
            Exit Function
        End If
        Do While Not .EOF
            '计算总额,待用
            '扣款.
            lng病人ID = Nvl(!病人ID, 0)
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
    If Not 扣款_重大校园卡("0", Val(Format(curTotal, "####0.00;-####0.00;0.00;0.00")) * 100, True) Then
        '
        挂号结算_重大校园卡 = False
        Exit Function
    End If
 
   '保存结算记录
     '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
    zlDatabase.ExecuteProcedure gstrSQL, "插入保险记录"
     挂号结算_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 挂号冲销_重大校园卡(ByVal lng结帐ID As Long) As Boolean
   Dim curTotal As Currency '上传费用总额
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long
    
    '先调门诊预结算,主要是取个人帐户支付额,再调门诊结算
    '保险参数中保存的是允许帐户支付的收入项目ID，挂号结算时判断，如果未设置，表明全自付，不需上传，否则仅上传那一笔明细
    
    On Error GoTo errHand
    
  gstrSQL = " " & _
        "  Select A.收费类别,A.病人id,sum(nvl(A.结帐金额,0)) as 实收金额" & _
        "  From 门诊费用记录  A" & _
        "  Where A.记录状态<>0 and  A.结帐ID=" & lng结帐ID & " and  Nvl(A.附加标志,0)<>9 " & _
        "  Group by A.收费类别,A.病人id" & _
        "  Order by A.收费类别"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取本次结帐费用明细"
    
    With rsTemp
        If .RecordCount = 0 Then
            ShowMsgbox "无任何费用记录发生!"
            Exit Function
        End If
        Do While Not .EOF
            '计算总额,待用
            '扣款.
            lng病人ID = Nvl(!病人ID, 0)
            curTotal = curTotal + Round(Nvl(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
  
 
   '保存结算记录
     '性质_IN,记录ID_IN,险类_IN,病人ID_IN,年度_IN,帐户累计增加_IN,帐户累计支出_IN,
    '累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '发生费用金额_IN,全自付金额_IN,首先自付金额_IN,进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN
    '个人帐户支付_IN,支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_重大校园卡 & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
        0 & "," & 0 & "," & _
        curTotal & "," & curTotal & ",0,0,0,0," & _
        curTotal & ",NULL,NULL," & curTotal & "," & curTotal & ",Null," & 0 & "," & _
        curTotal & ",0,NULL,null,null" & _
         " )"
      zlDatabase.ExecuteProcedure gstrSQL, "插入保险记录"
         
    挂号冲销_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function GetUserCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------
    '功能:读取卡信息(用户卡查询)
    '返回:获取成功,返回True,否则返回False
    '---------------------------------------------------------------------------------------------------
    Dim intReturn As Integer
    GetUserCardInfor = False
    '打开

    With g病人身份_重大校园卡
        .卡号 = Space(17)
        .注册日期 = Space(7)
        .姓名 = Space(21)
        .卡有效期 = Space(7)
        .身份证号 = Space(19)
        .上次交易时间 = Space(13)
        
        '读卡
        If mbln测试 Then
            intReturn = testdATA
        Else
             intReturn = Query_Pos_UserCard(.串口设备, .卡类型, .卡流水号, .卡号, .卡消费分组, .卡有效期, _
                    .姓名, .注册日期, .身份证号, .电子钱包1余额, .电子钱包2余额, .卡固有序号, .上次交易流水号, _
                    .上次交易金额, .上次交易时间, .日交易累计金额, .上次交易终端号, .卡等待时间)
                    
        End If
        If intReturn <> 0 Then
            '发生错误
            Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME)
            ShowMsgbox GetErrInfo(CStr(intReturn), TYPE_重大校园卡)
            Exit Function
        End If
        
        .卡号 = Trim(Replace(.卡号, Chr(0), "", 1))
        .注册日期 = Trim(Replace(.注册日期, Chr(0), "", 1))
        .姓名 = Trim(Replace(.姓名, Chr(0), "", 1))
        .卡有效期 = Trim(Replace(.卡有效期, Chr(0), "", 1))
        .身份证号 = Trim(Replace(.身份证号, Chr(0), "", 1))
        .上次交易时间 = Trim(Replace(.上次交易时间, Chr(0), "", 1))
        
        '计算出相关信息
        Dim int性别 As Integer
        int性别 = Val(IIf(Len(Trim(.身份证号)) = 18, Mid(Trim(.身份证号), 17, 1), Right(Trim(.身份证号), 1))) Mod 2
        '根据身份证取出相应的性别
        .性别 = IIf(int性别 = 0, "女", "男")
        .出生日期 = zlCommFun.GetIDCardDate(Trim(.身份证号))
        '计算年龄
        If IsDate(.出生日期) And .出生日期 <> "" Then
            .年龄 = Abs(Int((zlDatabase.Currentdate - CDate(.出生日期)) / 365))
        Else
            .年龄 = 0
        End If
        '判断该卡是否过期
        If "20" & Trim(.卡有效期) < Format(zlDatabase.Currentdate, "yyyymmdd") Then
            Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME)
            ShowMsgbox "该卡已经过期,不能再用(有效期为:20" & .卡有效期 & ")!"
            Exit Function
        End If
        If Not mbln测试 Then
            '判断该卡是否在默名单中
            intReturn = extSys_IsInBlackList(.卡流水号)
            If intReturn <> 0 Then
                Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME)
               ShowMsgbox "该卡为黑名单卡,不能使用该卡!"
               Exit Function
            End If
            Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME1)
        End If
    End With
    GetUserCardInfor = True
End Function

Public Function 扣款_重大校园卡(ByVal str费用代码 As String, ByVal lng金额 As Long, Optional bln门诊 As Boolean) As Boolean
    '功能:对费用进行扣款
    Dim intRetun As Integer
    Dim lngTmp As Long
    扣款_重大校园卡 = False
    Err = 0
    On Error GoTo errHand:
    '门诊收费用要先扣1钱包,否则扣除0
    lngTmp = lng金额
    If mbln测试 Then
        扣款_重大校园卡 = True
        Exit Function
    End If
    If gdbl交易限额_重大校园卡 < lng金额 / 100 Then
        ShowMsgbox "超过交易限额,不能结算!"
        扣款_重大校园卡 = False
        Exit Function
    End If
    If bln门诊 Then
        '先扣电钱包2的钱
        If g病人身份_重大校园卡.电子钱包2余额 <> 0 Then
            If lng金额 < g病人身份_重大校园卡.电子钱包2余额 Then
                intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 1, "200", lng金额)
                lngTmp = 0
            Else
                intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 1, "200", g病人身份_重大校园卡.电子钱包2余额)
                lngTmp = lng金额 - g病人身份_重大校园卡.电子钱包2余额
            End If
        End If
        If g病人身份_重大校园卡.电子钱包1余额 <> 0 And lngTmp > 0 Then
            '再扣电子钱包1的钱
            intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 0, "200", lngTmp)
        End If
        
'        If lng金额 <= g病人身份_重大校园卡.电子钱包2余额 And g病人身份_重大校园卡.电子钱包2余额 <> 0 Then
'            intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 1, "200", lngTmp)
'        Else
'            intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 0, "200", lngTmp)
'        End If
    Else
        intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 0, "200", lngTmp)
    End If
    
    If intRetun <> 0 Then
        Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME)
        If intRetun = -14 Then
            ShowMsgbox "电子钱包余额(" & g病人身份_重大校园卡.电子钱包1余额 / 100 & ")不足,不能进行结算!"
        Else
            MsgBox GetErrInfo(CStr(intRetun), TYPE_重大校园卡)
        End If
        Exit Function
    Else
        Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME1)
    End If
    扣款_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 扣款_余额2_重大校园卡() As Boolean
    Dim intRetun As Integer
    '清除电子钱包2的钱.
    扣款_余额2_重大校园卡 = False
    If GetUserCardInfor = False Then Exit Function
    intRetun = extSys_WithDraw(g病人身份_重大校园卡.串口设备, g病人身份_重大校园卡.卡流水号, 1, "200", g病人身份_重大校园卡.电子钱包2余额)
    If intRetun <> 0 Then
        Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME)
        MsgBox GetErrInfo(CStr(intRetun), TYPE_重大校园卡)
        Exit Function
    Else
        Call rf_beep(g病人身份_重大校园卡.串口设备, G_WAIT_TIME1)
    End If
    扣款_余额2_重大校园卡 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function testdATA() As Long
    '测试数据
    With g病人身份_重大校园卡
        .串口设备 = 1234
        .卡类型 = 3
        .卡流水号 = 1
        .卡号 = "13424"
        .卡消费分组 = 1
        .卡有效期 = "041236"
        .姓名 = "王磊"
        .注册日期 = "20020303"
        .身份证号 = "510221197404282859"
        .电子钱包1余额 = 3000
        .电子钱包2余额 = 4000
        .卡固有序号 = 1
        .上次交易流水号 = 2
        .上次交易金额 = 1000
        .上次交易时间 = "20040402"
        .日交易累计金额 = 1090
        .上次交易终端号 = 1
        .卡等待时间 = 200
    End With
    testdATA = 0
End Function






