Attribute VB_Name = "mdl毕节"
Option Explicit
'IC卡相关参数需要提供参数设置
'帐户余额表[险种名称]：基金医疗、大病、公务员、住院统筹，住院与前三种排斥，对结算无影响，直接按政策表进行计算
'门诊结算: 药品与诊疗各汇总为一笔传到中心
'医疗费用支出明细表: 结算汇总表
'当年住院费用: 本年分档计算的统筹累计金额
'支持结算方式种类: 个人帐户、统筹基金
'每天运行程序前，检查是否存在未上传或未下载，存在则不允许运行
'下载数据时，条件为：基本险种=医疗保险
'结算算法
'   门诊：全由帐户支付或卡支付
'   住院：先按项目表计算进入统筹金额，再按病种计算实际进入统筹金额，最后按分档计算得出实际统筹基金支付金额
'日期格式: yyyy.MM.dd，时间格式:yyyy.MM.dd HH:mm:ss
'如果帐户停用，则只能当作普通病人处理
'未对码项目，如果启用药品诊疗自付，则做为全自付；否则全进入统筹金额，按分档计算
'不支持中途结算,结算时必须出院
'实时还是脱机使用由中心端控制，需实时上传所有明细，结算后需清除中心端的未结算明细，将本次结算明细上传到中心，再将结算信息上传到中心
'需要完成结算部件（门诊、住院结算）、上传下载（费用和结算数据的上传、政策数据的下载）、IC卡读写及连接中心的部件
'结算如果使用了个人帐户的，需要更新中心的个人帐户余额表，本地的个人帐户余额表，及下卡
'要求每天零点必须退出程序，如要继续使用，必须重新启动一次
'如果数据已写入中间库，则费用明细的是否上传填为1；如果上传到中心，则中间库中的是否上传填为1

#Const gverControl = 99
Private Enum ic
    shbzh = 0
    xm
    dwdm
    xb
    Csrq
    cjqzrq
    jyqkdm
    yxkh
    grjbdm
    ye
    zhjzrq
    yydm
    pass
End Enum

Private Type IC_Struct
    社会保障号 As String
    姓名 As String
    单位代码 As String
    性别 As String
    出生日期 As String
    参加工作日期 As String
    就业情况代码 As String
    有效卡号 As Integer
    个人级别代码 As String
    个人帐户余额 As Double
    最后就诊日期 As String
    最后就诊医院代码 As String
    个人IC卡密码    As String
End Type
Public IC_Data_毕节 As IC_Struct

Private Type gCominfo
    strHospitalCode As String       '医院代码
    strHospitalName As String       '医院名称
    strConnectPass As String        '连接密码
    blnOnLine As Boolean            '实时联机还是脱机
    blnICPassVerify As Boolean      '是否使用IC卡密码
    blnDiseaseCash As Boolean       '是否启用病种自付
    blnPhysicCash As Boolean        '是否启用药品诊疗自付
    blnYearBase As Boolean          '是否按年度起付
    '以下住院使用
    str险种名称 As String           '记录第一个适应的险种名称，并保存到中心
    str就诊流水号 As String         '就诊流水号
    dbl费用总额 As Double
    dbl年度统筹 As Double           '年度统筹累计（指进入分档计算时的统筹金额累计）
    dbl统筹金额 As Double           '同上，只是本次进入分档计算时的统筹金额
    dbl年度报销 As Double           '按年度报销金额累计，需要再次计算
    dbl统筹报销 As Double           '本次统筹报销金额，如果是年度起付，需减去dbl年度报销得到本次实际统筹支付金额
    str社会保障号 As String         '医保号
    str有效卡号 As String           '卡号
    用户名 As String          '中间库用户名
End Type
Public gCominfo_毕节 As gCominfo

Public gcnCenter As New ADODB.Connection
Public gobjCenter As Object

Private mblnInit As Boolean
Private mstrFirstStart As String        '记录第一次登录日期，如果不同则禁止使用

Private Const gstr药品代码 As String = "000"
Private Const gstr药品大类 As String = "西药类"
Private Const gstr诊疗代码 As String = "01"
Private Const gstr诊疗大类 As String = "检查费"

Public Function gcnGYBJYB() As ADODB.Connection
    Dim strUser As String
    Dim strServer As String
    Dim strPass As String
    Dim rsTemp As ADODB.Recordset
On Error GoTo ErrH
    If GetSetting("ZLSOFT", "公共模块\医保\" & TYPE_毕节, "NoLink", 0) = "1" Then  '不需要连接中心
        If gcnCenter Is Nothing Then
            If MsgBox("本工作站未设置连接到中心，现在是否连接？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo Link
            End If
        ElseIf gcnCenter.State <> 1 Then
            If MsgBox("本工作站未设置连接到中心，现在是否连接？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo Link
            End If
        Else
            Set gcnGYBJYB = gcnCenter
        End If
    Else '需要连接到中心
        If gcnCenter Is Nothing Then
            If MsgBox("中心连接已断开，现在是否重新连接？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo Link
            End If
        ElseIf gcnCenter.State <> 1 Then
            If MsgBox("中心连接已断开，现在是否重新连接？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                GoTo Link
            End If
        Else
            Set gcnGYBJYB = gcnCenter
        End If
    End If
    Exit Function
Link:
    '对 gCnCenter 初始化
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", gintInsure)
    Do While Not rsTemp.EOF
        Select Case rsTemp!参数名
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "连接密码"
                gCominfo_毕节.strConnectPass = Nvl(rsTemp!参数值)
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(gcnCenter, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接中间库，请查看医保参数是否设置真确", vbInformation, gstrSysName
        Exit Function
    End If
    Set gcnGYBJYB = gcnCenter
    
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Resume Next
    Exit Function
End Function

Public Function GetAge(ByVal strServer As String, ByVal strTest As String) As Long
    Dim strServerYear As String, strTestYear As String
    Dim strServerMonth As String, strTestMonth As String
    Dim strServerDay As String, strTestDay As String
    Dim lngAge As Long
    Dim intDef As Integer
    
    '计算年龄或工龄,例:未满31岁,按30岁算
    If Not IsDate(strServer) Then
        MsgBox "传入的第一个参数不是日期型！[GetAge]", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsDate(strTest) Then
        MsgBox "传入的第二个参数不是日期型！[GetAge]", vbInformation, gstrSysName
        Exit Function
    End If
    
    '分解年、月、日
    strServerYear = Mid(strServer, 1, 4)
    strServerMonth = Mid(strServer, 6, 2)
    strServerDay = Mid(strServer, 9, 2)
    strTestYear = Mid(strTest, 1, 4)
    strTestMonth = Mid(strTest, 6, 2)
    strTestDay = Mid(strTest, 9, 2)
    
    '先按年算，得出大概的年龄
    lngAge = Val(strServerYear) - Val(strTestYear)
    '如果服务器月份大于出生月份，不修正直接返回；如果小于，则年龄减1；如果相同，则继续判断
    intDef = Val(strServerMonth) - Val(strTestMonth)
    If intDef > 0 Then
        GetAge = lngAge
        Exit Function
    ElseIf intDef < 0 Then
        GetAge = (lngAge - 1)
        Exit Function
    Else
        intDef = Val(strServerDay) - Val(strTestDay)
        If intDef >= 0 Then
            GetAge = lngAge
            Exit Function
        Else
            GetAge = (lngAge - 1)
            Exit Function
        End If
    End If
End Function

Public Function 身份标识_毕节(Optional bytType As Byte, Optional lng病人ID As Long) As String
    身份标识_毕节 = frmIdentify毕节.GetIdentify(bytType, lng病人ID)
End Function

Public Function 医保初始化_毕节(Optional ByVal blnTest As Boolean = False) As Boolean
    '功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
    '返回：初始化成功，返回true；否则，返回false
    Dim bln禁止登录 As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strServer As String, strUser As String, strPass As String
    
On Error Resume Next

    If gobjCenter Is Nothing Then
        Err = 0
        Set gobjCenter = CreateObject("Interface.clsInterface")
        If Err <> 0 Then
            MsgBox "无法创建连接中心部件！请与医保中心或开发商联系！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
On Error GoTo errHand
    
    '取保险类别中的医院编码
    gCominfo_毕节.strHospitalCode = ""
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医院编码", TYPE_毕节)
    If rsTemp.RecordCount <> 0 Then
        gCominfo_毕节.strHospitalCode = Nvl(rsTemp!医院编码)
    End If
    If gCominfo_毕节.strHospitalCode = "" Then
        MsgBox "还未初始化，请在保险类别管理里设置本医疗机构的单位代码！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '连接服务器
    If GetSetting("ZLSOFT", "公共模块\医保\" & App.EXEName, "NoLink", 0) <> "1" Then '不需要连接中心
        If Not mblnInit Then
            '读出连接医保服务器的配置
            gCominfo_毕节.strConnectPass = ""
            gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医院编码", TYPE_毕节)
            Do Until rsTemp.EOF
                Select Case rsTemp("参数名")
                    Case "医保用户名"
                        strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保服务器"
                        strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "医保用户密码"
                        strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                    Case "连接密码"
                        gCominfo_毕节.strConnectPass = Nvl(rsTemp!参数值)
                End Select
                rsTemp.MoveNext
            Loop
            gCominfo_毕节.用户名 = strUser
            If Not OraDataOpen(gcnCenter, strServer, strUser, strPass, False) Then
                MsgBox "无法连接中间库，请查看医保参数是否设置真确", vbInformation, gstrSysName
                Exit Function
            End If
            '取连接方式
            If Not 获取连接方式() Then Exit Function
            
            '检查是否存在未上传的费用明细或结算数据，如果存在则不允许使用，提示用户使用上传程序
            If Not 检查是否上传明细 Then Exit Function
            
            '检查今天是否进行过下载，如果没有，也禁止使用（同时提取出单位名称）
            If Not 检查是否下载 Then Exit Function
            
            '检查是否连的通中心
            If Not 获取中心连接 Then Exit Function
            If Not 检查连接密码 Then Exit Function
            Call 关闭中心连接
            
            If mstrFirstStart = "" Then mstrFirstStart = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            mblnInit = True
        End If
    End If
    医保初始化_毕节 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 医保设置_毕节() As Boolean
    医保设置_毕节 = frmSet毕节.参数设置()
End Function

Public Function 医保终止_毕节() As Boolean
    On Error Resume Next
    mblnInit = False
    If gCominfo_毕节.blnOnLine Then
        If Not gobjCenter Is Nothing Then
            Call gobjCenter.CloseConnector
            Set gobjCenter = Nothing
        End If
    End If
End Function

Public Function 门诊虚拟结算_毕节(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim dbl金额 As Double
    Dim cur帐户支付 As Double
On Error GoTo ErrH
    With rs明细
        '检查是否存在负数记帐
        Do While Not .EOF
            If Nvl(!实收金额, 0) < 0 Then
                Err.Raise 9000, gstrSysName, "医保政策规定，不允许负数记帐！"
                Exit Function
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        
        '取费用总额
        Do While Not .EOF
            dbl金额 = dbl金额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    '如果个人帐户余额大于本次结算金额，本次帐户支付额等于结算金额，否则等于个人帐户余额
    cur帐户支付 = IIf(IC_Data_毕节.个人帐户余额 >= dbl金额, dbl金额, IC_Data_毕节.个人帐户余额)
    
    str结算方式 = "个人帐户;" & Format(cur帐户支付, "#####0.00") & ";1"
    门诊虚拟结算_毕节 = True
    Exit Function
ErrH:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算_毕节(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    Dim strCard As String
    Dim blnTrans As Boolean
    Dim lng病人ID As Long
    Dim bln药品 As Boolean
    Dim str就诊登记号 As String
    Dim cur帐户支付 As Currency
    Dim cur帐户支付额 As Currency           '记录药品或诊疗中，帐户实际支付额
    Dim dbl费用总额 As Double
    Dim rsDetail As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHand
    
    cur帐户支付 = cur个人帐户
    '提取本次结算费用明细
    gstrSQL = " Select  A.病人ID,A.收费类别,A.收费细目ID,round(A.实收金额,2) 实收金额,B.项目编码,Nvl(B.项目名称,C.名称) AS 项目名称,B.附注" & _
              " From 门诊费用记录 A,保险支付项目 B,收费细目 C" & _
              " Where A.结帐ID=[2] And A.收费细目ID=B.收费细目ID(+) And B.险类(+)=[1]" & _
              " And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0" & _
              " And A.收费细目ID=C.ID"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本次结算费用明细", TYPE_毕节, lng结帐ID)
    lng病人ID = rsDetail!病人ID
    
    '判断卡是不是当前病人的
    gstrSQL = "Select 帐户余额,卡号,医保号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "判断卡是不是当前病人的", lng病人ID, TYPE_毕节)
    'IC_Data_毕节.个人帐户余额 = Nvl(rsCheck!帐户余额, 0)
    
    '读卡
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call 数据转换_毕节(strCard, True)
    If Not (IC_Data_毕节.有效卡号 = Nvl(rsCheck!卡号, 0) And IC_Data_毕节.社会保障号 = rsCheck!医保号) Then
        Err.Raise 9000, "医保中心提示", "当前IC卡不是该病人的卡或该卡已失效，请与医保中心联系！"
        Call IC_End(True)
        Exit Function
    End If
    
    IC_Data_毕节.个人帐户余额 = Nvl(rsCheck!帐户余额, 0)
    str就诊登记号 = Get流水号_毕节
    
    blnTrans = True
    If Not 事务_开始 Then
        Call IC_End(True)
        Exit Function
    End If
    
    With rsDetail
        Do While Not .EOF
            '填写中间库数据(填写药品费用明细表、诊疗费用明细表及医疗费用支出明细表)
            bln药品 = (InStr(1, "5,6,7", !收费类别) <> 0)
            cur帐户支付额 = IIf(cur帐户支付 >= !实收金额, !实收金额, cur帐户支付)
            If bln药品 Then
                gstrSQL = "" & _
                    " INSERT INTO 药品费用明细表" & _
                    " (ID,社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                    " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                    " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                    " VALUES" & _
                    " (药品费用明细表_ID.Nextval,'" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
                    "'" & Nvl(!项目编码, gstr药品代码) & "','" & !项目名称 & "','" & Nvl(!附注, gstr药品大类) & "'," & _
                    "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & !实收金额 & "," & _
                    "0," & cur帐户支付额 & "," & !实收金额 - cur帐户支付额 & "," & _
                    "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','是','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                gcnGYBJYB.Execute gstrSQL
            Else
                gstrSQL = "" & _
                    " INSERT INTO 诊疗费用明细表" & _
                    " (ID,社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                    " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                    " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                    " VALUES" & _
                    " (诊疗费用明细表_ID.Nextval,'" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
                    "'" & Nvl(!项目编码, gstr诊疗代码) & "','" & !项目名称 & "','" & Nvl(!附注, gstr诊疗大类) & "'," & _
                    "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & !实收金额 & "," & _
                    "0," & cur帐户支付额 & "," & !实收金额 - cur帐户支付额 & "," & _
                    "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','是','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                gcnGYBJYB.Execute gstrSQL
            End If
            
            '填写中心库数据，同上
            If gCominfo_毕节.blnOnLine Then
                If bln药品 Then
                    gstrSQL = "" & _
                        " INSERT INTO 药品费用明细表" & _
                        " (社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员)" & _
                        " VALUES" & _
                        " ('" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
                        "'" & Nvl(!项目编码, gstr药品代码) & "','" & !项目名称 & "','" & Nvl(!附注, gstr药品大类) & "'," & _
                        "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & !实收金额 & "," & _
                        "0," & cur帐户支付额 & "," & !实收金额 - cur帐户支付额 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "')"
                    If Not ExecuteSQL(gstrSQL) Then
                        Call IC_End(True)
                        Exit Function
                    End If
                Else
                    gstrSQL = "" & _
                        " INSERT INTO 诊疗费用明细表" & _
                        " (社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员)" & _
                        " VALUES" & _
                        " ('" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
                        "'" & Nvl(!项目编码, gstr诊疗代码) & "','" & !项目名称 & "','" & Nvl(!附注, gstr诊疗大类) & "'," & _
                        "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & !实收金额 & "," & _
                        "0," & cur帐户支付额 & "," & !实收金额 - cur帐户支付额 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "')"
                    If Not ExecuteSQL(gstrSQL) Then
                        Call IC_End(True)
                        Exit Function
                    End If
                End If
            End If
            
            cur帐户支付 = cur帐户支付 - cur帐户支付额
            dbl费用总额 = dbl费用总额 + !实收金额
            .MoveNext
        Loop
    End With
    
    '为所有费用明细打上上传标记
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '填写保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_毕节 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        dbl费用总额 & "," & dbl费用总额 - cur个人帐户 & ",0,0,0,0," & _
        0 & "," & cur个人帐户 & ",'" & str就诊登记号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    gstrSQL = "" & _
        " INSERT INTO 医疗费用支出明细表" & _
        " (ID,社会保障号,姓名,门诊住院号,发生时间,发生类别,总费用,统筹基金支付," & _
        " 个人帐户支付,个人自付金额,医疗机构代码,医疗机构名称,操作员,是否上传)" & _
        " VALUES" & _
        " (医疗费用支出明细表_ID.Nextval,'" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
        "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & dbl费用总额 & "," & _
        "0," & cur个人帐户 & "," & dbl费用总额 - cur个人帐户 & "," & _
        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = "" & _
            " INSERT INTO 医疗费用支出明细表" & _
            " (社会保障号,姓名,门诊住院号,发生时间,发生类别,总费用,统筹基金支付," & _
            " 个人帐户支付,个人自付金额,医疗机构代码,医疗机构名称,操作员)" & _
            " VALUES" & _
            " ('" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
            "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','门诊'," & dbl费用总额 & "," & _
            "0," & cur个人帐户 & "," & dbl费用总额 - cur个人帐户 & "," & _
            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "')"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    '下帐户（中心库的个人帐户余额表要更新，中间库的个人帐户余额表要更新）
    '住院还需要更新当前住院费用字段（统筹支付、住院次数）
    IC_Data_毕节.个人帐户余额 = IC_Data_毕节.个人帐户余额 - cur个人帐户
    IC_Data_毕节.最后就诊日期 = Format(zlDatabase.Currentdate, "yyyy.MM.dd")
    IC_Data_毕节.最后就诊医院代码 = gCominfo_毕节.strHospitalCode
    Call 数据转换_毕节(strCard, False)
    
    gstrSQL = " Update 个人帐户余额表 " & _
              " Set 余额=Nvl(余额,0)-" & Val(cur个人帐户) & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then 'SqlServer非空函数是IsNULL();而Oracle是Nvl()
        gstrSQL = " Update 个人帐户余额表 " & _
                  " Set 本年支出=IsNull(本年支出,0)+" & Val(cur个人帐户) & "," & _
                  "     累计支出=IsNull(累计支出,0)+" & Val(cur个人帐户) & "," & _
                  "     余额=IsNull(余额,0)-" & Val(cur个人帐户) & _
                  " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call 事务_回滚
        Call IC_End(True)
        Exit Function
    End If
    
    If 事务_提交 Then
        门诊结算_毕节 = True
    Else
        Call 事务_回滚
    End If
    blnTrans = False
    
    Call IC_End
    
    '如果使用个人帐户支付，显示消费提示框
    If cur个人帐户 <> 0 And 门诊结算_毕节 Then
        Call Frm交易提示框.ShowME(IC_Data_毕节.姓名, IC_Data_毕节.个人帐户余额 + cur个人帐户, _
            IC_Data_毕节.个人帐户余额, dbl费用总额, cur个人帐户, dbl费用总额 - cur个人帐户)
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 事务_回滚
    Call IC_End(True)
End Function

Public Function 门诊结算冲销_毕节(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    '只有门诊允许结算作废，且只能退最后一笔，退完后卡上的最后就诊医疗机构编码填为H000，删除中间库与中心的费用明细与支出明细
    Dim lng冲销ID As Long
    Dim strCard As String
    Dim str就诊登记号 As String, str退单流水号 As String
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHand
    
    '判断卡是不是当前病人的
    gstrSQL = "Select 卡号,医保号 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "判断卡是不是当前病人的", lng病人ID, TYPE_毕节)
    
    '读卡
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call 数据转换_毕节(strCard, True)
    If Not (IC_Data_毕节.有效卡号 = Nvl(rsCheck!卡号, 0) And IC_Data_毕节.社会保障号 = rsCheck!医保号) Then
        Call IC_End(True)
        Err.Raise 9000, "医保中心提示", "当前IC卡不是该病人的卡或该卡已失效，请与医保中心联系！"
        Exit Function
    End If
    
    '如果最后就诊医疗机构代码都不相同，则直接退出
    If IC_Data_毕节.最后就诊医院代码 <> gCominfo_毕节.strHospitalCode Then
        Call IC_End(True)
        Err.Raise 9000, "医保中心提示", "你已在其他医疗机构发生就诊行为，不能退单！"
        Exit Function
    End If
    
    '取本次结帐ID
    gstrSQL = "select distinct A.结帐ID,A.NO from 门诊费用记录 A,门诊费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读新产生的结帐ID", lng结帐ID)
    lng冲销ID = rsTemp!结帐ID
    
    '提取原始的保险结算记录
    gstrSQL = "Select * From 保险结算记录 Where 性质=1 AND 记录ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取原始的保险结算记录", lng结帐ID)
    str就诊登记号 = Nvl(rsTemp!支付顺序号)
    
    str退单流水号 = InputBox("请输入原始单据的就诊登记号：", "就诊登记号")
    '判断卡内记录的最后就诊医院机构代码是否与中心的一致（便于处理脱机，只从中间库取）
    gstrSQL = " Select Max(门诊住院号) 就诊登记号 From 医疗费用支出明细表" & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "' And 医疗机构代码='" & gCominfo_毕节.strHospitalCode & "'"
    With rsCheck
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open gstrSQL, gcnGYBJYB
        If IsNull(!就诊登记号) Then
            Call IC_End(True)
            Err.Raise 9000, gstrSysName, "没有找到最后一次的就诊登记号，无法退单！"
            Exit Function
        End If
        If Nvl(!就诊登记号) <> str就诊登记号 Then
            Call IC_End(True)
            Err.Raise 9000, gstrSysName, "只能退该病人在本院就诊的最后一笔门诊单据！"
            Exit Function
        End If
        If str就诊登记号 <> str退单流水号 Then
            Call IC_End(True)
            Err.Raise 9000, gstrSysName, "输入的就诊登记号与原单据的就诊登记号不符，无法退单！"
            Exit Function
        End If
    End With
    
    '为所有费用明细打上上传标记
    gstrSQL = "zl_病人结帐记录_上传(" & lng冲销ID & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '保存保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & TYPE_毕节 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",0,0,0,0,0," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",NULL)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    blnTrans = True
    If Not 事务_开始 Then
        Call IC_End(True)
        Exit Function
    End If
    
    '删除中间库、中心库中的费用明细与支出明细记录
    gstrSQL = " Delete 药品费用明细表 " & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "' And 住院号='" & str就诊登记号 & "' And 医疗机构代码='" & gCominfo_毕节.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    gstrSQL = " Delete 诊疗费用明细表 " & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "' And 住院号='" & str就诊登记号 & "' And 医疗机构代码='" & gCominfo_毕节.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    gstrSQL = " Delete 医疗费用支出明细表 " & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "' And 门诊住院号='" & str就诊登记号 & "' And 医疗机构代码='" & gCominfo_毕节.strHospitalCode & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    '写卡，修改余额与帐户支出累计
    '下帐户（中心库的个人帐户余额表要更新，中间库的个人帐户余额表要更新）
    '不管本次交易是否使用个人帐户，都要写卡的目的是更新最后就诊医疗机构代码
    IC_Data_毕节.个人帐户余额 = IC_Data_毕节.个人帐户余额 + cur个人帐户
    IC_Data_毕节.最后就诊医院代码 = "H000"
    Call 数据转换_毕节(strCard, False)
    
    gstrSQL = " Update 个人帐户余额表 " & _
              " Set 余额=Nvl(余额,0)+" & Val(cur个人帐户) & _
              " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Update 个人帐户余额表 " & _
                  " Set 本年支出=IsNull(本年支出,0)-" & Val(cur个人帐户) & "," & _
                  "     累计支出=IsNull(累计支出,0)-" & Val(cur个人帐户) & "," & _
                  "     余额=IsNull(余额,0)+" & Val(cur个人帐户) & _
                  " Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
        If Not ExecuteSQL(gstrSQL) Then
            Call IC_End(True)
            Exit Function
        End If
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call IC_End(True)
        Call 事务_回滚
        Exit Function
    End If
    
    If 事务_提交 Then
        门诊结算冲销_毕节 = True
    Else
        Call 事务_回滚
    End If
    
    Call IC_End
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 事务_回滚
    Call IC_End(True)
End Function

Public Function 入院登记_毕节(lng病人ID As Long, lng主页ID As Long) As Boolean    '因入院登记前必须要进行身份验证,因此,病人相关信息可直接从卡数据中获取
    Dim blnTrans As Boolean
    Dim str就诊登记号 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    str就诊登记号 = Get流水号_毕节
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_毕节 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理入院登记")
    
    '检查该病人是否已在中心在院（可能在其他医院办理了入院）
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Select 1 From 住院登记表 Where rtrim(社会保障号)='" & Trim(IC_Data_毕节.社会保障号) & "' And IsNULL(出院时间,'')=''"
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call gobjCenter.CloseConnector
            Exit Function
        End If
        If rsTemp.RecordCount <> 0 Then
            Err.Raise 9000, gstrSysName, "该病人已在中心登记为在院，请核实是否已在其它医院办理了医保住院！"
            Call gobjCenter.CloseConnector
            Exit Function
        End If
        Call gobjCenter.CloseConnector
    End If
    
    '获取本次入院相关信息
    gstrSQL = " Select B.入院日期,D.名称 As 入院科室,B.入院病床,B.登记人 As 操作员,E.病种名称,Sum(Nvl(F.金额,0)) As 预交款" & _
              " From 保险帐户 A,病案主页 B,病人信息 C,部门表 D," & gCominfo_毕节.用户名 & ".病种目录表 E,病人预交记录 F" & _
              " Where A.险类=[1] And A.病人ID=[2] And B.主页ID=" & lng主页ID & " ANd A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.主页ID=C.住院次数" & _
              " And A.病种ID=E.ID(+) And B.入院科室ID=D.ID(+) And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And F.记录性质(+)=1 " & _
              " Group by B.入院日期,D.名称,B.入院病床,B.登记人,E.病种名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次入院相关信息", TYPE_毕节, lng病人ID)
    
    If Not 事务_开始 Then Exit Function
    
    '产生入院登记记录(中间库与中心库)
    gstrSQL = " Insert Into 入院登记表" & _
              " (ID,社会保障号,姓名,住院号,入院时间,出院时间,病种名称," & _
              " 科室,床位号,预付款,操作员,是否结算,是否上传)" & _
              " Values" & _
              " (入院登记表_ID.Nextval,'" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
              "'" & Format(rsTemp!入院日期, "yyyy.MM.dd") & "',NULL,'" & rsTemp!病种名称 & "','" & Nvl(rsTemp!入院科室) & "'," & _
              "'" & ToVarchar(Nvl(rsTemp!入院病床), 4) & "'," & Val(Nvl(rsTemp!预交款, 0)) & ",'" & Nvl(rsTemp!操作员, "ZLHIS") & "','否'," & IIf(gCominfo_毕节.blnOnLine, "'10'", "'00'") & ")"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Insert Into 住院登记表" & _
                  " (社会保障号,姓名,住院号,入院时间,出院时间,病种名称," & _
                  " 科室,床位号,预付款,操作员,是否结算,医疗机构代码,医疗机构名称)" & _
                  " Values" & _
                  " ('" & IC_Data_毕节.社会保障号 & "','" & IC_Data_毕节.姓名 & "','" & str就诊登记号 & "'," & _
                  "'" & Format(rsTemp!入院日期, "yyyy.MM.dd") & "','','" & rsTemp!病种名称 & "','" & Nvl(rsTemp!入院科室) & "'," & _
                  "'" & ToVarchar(Nvl(rsTemp!入院病床), 4) & "'," & Val(Nvl(rsTemp!预交款, 0)) & ",'" & Nvl(rsTemp!操作员, "ZLHIS") & "','否'," & _
                  "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "')"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    If 事务_提交 Then
        入院登记_毕节 = True
    Else
        Call 事务_回滚
    End If
    
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 事务_回滚
End Function

Public Function 入院登记撤销_毕节(lng病人ID As Long, lng主页ID As Long, Optional ByVal bln无费出院 As Boolean = False) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
    '删除入院登记表即可，如果本次登记了费用，则不允许撤销入院
    Dim str住院号 As String, str社会保障号 As String, str入院日期 As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 医保号 From 保险帐户 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取社会保障号", lng病人ID)
    str社会保障号 = rsTemp!医保号
    
    gstrSQL = " Select 入院日期 From 病案主页 Where 病人ID=[1] And 主页ID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期", lng病人ID, lng主页ID)
    str入院日期 = Format(rsTemp!入院日期, "yyyy.MM.dd")
    
    gstrSQL = " Select 住院号 From " & gCominfo_毕节.用户名 & ".入院登记表 " & _
              " Where 社会保障号=[1] And 入院时间=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出本次住院号", str社会保障号, str入院日期)
    str住院号 = rsTemp!住院号
    
    If Not bln无费出院 Then
        '再取本次住院是否发生费用
        gstrSQL = " Select Count(*) Records From 药品费用明细表 Where 住院号='" & str住院号 & "'" & _
                " Union ALL" & _
                " Select Count(*) Records From 诊疗费用明细表 Where 住院号='" & str住院号 & "'"
        gstrSQL = "Select SUM(Records) AS Records From (" & gstrSQL & ")"
        With rsTemp
            If .State = 1 Then .Close
            .Open gstrSQL, gcnGYBJYB
            If !Records > 0 Then
                MsgBox "已经发生费用，不能撤销入院！", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    Else
        '存在结算过的明细则不允许转普通病人
        gstrSQL = " Select Count(*) Records From 药品费用明细表 Where 住院号='" & str住院号 & "' And Nvl(是否结算,'否')='是'" & _
                " Union ALL" & _
                " Select Count(*) Records From 诊疗费用明细表 Where 住院号='" & str住院号 & "' And Nvl(是否结算,'否')='是'"
        gstrSQL = "Select SUM(Records) AS Records From (" & gstrSQL & ")"
        With rsTemp
            If .State = 1 Then .Close
            .Open gstrSQL, gcnGYBJYB
            If !Records > 0 Then
                MsgBox "本次住院已经有部分明细进行了结算，不能转为普通病人！", vbInformation, gstrSysName
                Exit Function
            End If
        End With
    End If
    
    If Not 事务_开始 Then Exit Function
    
    '删除中间表中的明细
    gstrSQL = " Delete 药品费用明细表 " & _
              " Where 社会保障号='" & str社会保障号 & "' And 住院号='" & str住院号 & "'"
    gcnGYBJYB.Execute gstrSQL
    gstrSQL = " Delete 诊疗费用明细表 " & _
              " Where 社会保障号='" & str社会保障号 & "' And 住院号='" & str住院号 & "'"
    gcnGYBJYB.Execute gstrSQL
    '删除所有未结算明细
    gstrSQL = " Delete 住院未结算药品费用日记帐表" & _
              " Where 社会保障号='" & str社会保障号 & "' And 住院号='" & str住院号 & "'"
    If Not ExecuteSQL(gstrSQL) Then Exit Function
    gstrSQL = " Delete 住院未结算诊疗费用日记帐表" & _
              " Where 社会保障号='" & str社会保障号 & "' And 住院号='" & str住院号 & "'"
    If Not ExecuteSQL(gstrSQL) Then Exit Function
    
    '删除住院登记表
    gstrSQL = "Delete 入院登记表 Where 住院号='" & str住院号 & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = "Delete 住院登记表 Where 住院号='" & str住院号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_毕节 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "办理出院登记")
    
    If 事务_提交 Then
        入院登记撤销_毕节 = True
    Else
        Call 事务_回滚
    End If
End Function

Public Function 出院登记_毕节(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim blnTrans As Boolean
    Dim str社会保障号 As String
        Dim bln零费用出院  As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '必须全部结清才能出院
    If 存在未结费用(lng病人ID, lng主页ID) Then
        Err.Raise 9000, gstrSysName, "只有费用结清后才允许出院！"
        Exit Function
    End If

        '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额  from 住院费用记录 where 病人ID=[1] and 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人出院", lng病人ID, lng主页ID)
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
        If bln零费用出院 Then
            出院登记_毕节 = 入院登记撤销_毕节(lng病人ID, lng主页ID, True)
        Else
            '提取参保人的社会保障号
            gstrSQL = "Select 医保号 From 保险帐户 Where 病人ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参保人的社会保障号", lng病人ID)
            str社会保障号 = Nvl(rsTemp!医保号)
            
            If Not 事务_开始 Then Exit Function
            blnTrans = True
            If Not 病人变动记录上传_毕节(lng病人ID, lng主页ID, False) Then
                Call 事务_回滚
                Exit Function
            End If
            
            gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_毕节 & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "办理出院登记")
            
            '增加住院次数
            gstrSQL = " Update 个人帐户余额表 " & _
                        " Set 住院次数=Nvl(住院次数,0)+1" & _
                        " Where 社会保障号='" & str社会保障号 & "'"
            gcnGYBJYB.Execute gstrSQL
            If gCominfo_毕节.blnOnLine Then 'SqlServer非空函数是IsNULL();而Oracle是Nvl()
                gstrSQL = " Update 个人帐户余额表 " & _
                            " Set 住院次数=IsNull(住院次数,0)+1" & _
                            " Where 社会保障号='" & str社会保障号 & "'"
                If Not ExecuteSQL(gstrSQL) Then Exit Function
            End If

            If 事务_提交 Then
                出院登记_毕节 = True
            Else
                Call 事务_回滚
            End If
            blnTrans = False
        End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 事务_回滚
End Function

Public Function 出院登记撤销_毕节(lng病人ID As Long, lng主页ID As Long) As Boolean
    '朱科长说不允许撤销出院
    MsgBox "医保政策不允许撤销出院，请与医保中心联系！", vbInformation, gstrSysName
    出院登记撤销_毕节 = False
End Function

Public Function 个人余额_毕节(ByVal lng病人ID As Long) As Currency
    '功能: 提取参保病人个人帐户余额
    '因为门诊每次就诊时都会更新，而住院只能结算一次，所以可以直接以库中的余额为准
    Dim rsTemp As New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Open "Select Nvl(帐户余额,0) 余额 From 保险帐户 Where 险类=" & TYPE_毕节 & " And 病人ID=" & lng病人ID, gcnOracle
        个人余额_毕节 = !余额
    End With
End Function

Public Function 住院虚拟结算_毕节(rsExse As Recordset, ByVal lng病人ID As Long) As String
    Dim lng主页ID As Long
    Dim dbl进入统筹 As Double       '单笔明细的进入统筹金额
    Dim dbl个人自付 As Double
    Dim dbl个人帐户 As Double
    Dim dbl帐户余额 As Double
    Dim bln药品 As Boolean
    Dim blnTrans As Boolean
    Dim bln自费病人 As Boolean
    Dim str项目编码 As String, str项目名称 As String, str大类 As String, STR姓名 As String
    Dim dbl金额 As Double, dbl数量 As Double, dbl单价 As Double
    
    Dim rsTemp As New ADODB.Recordset
    Dim cnOracle As New ADODB.Connection
    On Error GoTo errHand
    
    With gCominfo_毕节
        .dbl费用总额 = 0
        .dbl年度报销 = 0
        .dbl年度统筹 = 0
        .dbl统筹报销 = 0
        .dbl统筹金额 = 0
    End With
    
    Set cnOracle = GetNewConnection
    
    
    '取该病人的社会保障号
    gstrSQL = " Select B.姓名,A.卡号,A.医保号,A.帐户余额 From 保险帐户 A,病人信息 B" & _
              " Where A.险类=[2] And A.病人ID=B.病人ID And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的社会保障号", lng病人ID, TYPE_毕节)
    gCominfo_毕节.str社会保障号 = rsTemp!医保号
    gCominfo_毕节.str有效卡号 = Nvl(rsTemp!卡号, 0)
    STR姓名 = rsTemp!姓名
    dbl帐户余额 = Nvl(rsTemp!帐户余额, 0)
    
    '取该病人的住院流水号
    gstrSQL = "Select 住院号 From 入院登记表 Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "' And 出院时间 Is Null"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "没有找到该病人有效的在院记录，无法进行结算！", vbInformation, gstrSysName
            Exit Function
        End If
        gCominfo_毕节.str就诊流水号 = Nvl(!住院号)
    End With
    
    '读卡
'    Dim strCard As String
'    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
'    Call 数据转换_毕节(strCard, True)
'    If Not (IC_Data_毕节.有效卡号 = gCominfo_毕节.str有效卡号 And IC_Data_毕节.社会保障号 = gCominfo_毕节.str社会保障号) Then
'        Call IC_End(True)
'        MsgBox "当前IC卡不是该病人的卡或该卡已失效，请与医保中心联系！", vbInformation, gstrSysName
'        Exit Function
'    End If
    '因住院期间只能结算一次，所以从保险帐户中取，保险帐户中的帐户余额肯定是正确的，在身份验证时已校正
'    IC_Data_毕节.个人帐户余额 = dbl帐户余额
    
    '检查该病人的卡状态
    '以下数据需要从中间库中获取
    gstrSQL = "Select 住院次数,帐户冻结,冻结原因,有效卡号,当年住院费用,冻结时间,冻结说明 " & _
        " From 个人帐户余额表 Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'"
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call IC_End(True)
            Call gobjCenter.CloseConnector
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTemp
        If .RecordCount = 0 Then
'            Call IC_End(True)
            MsgBox "没发现该病人的有效记录，请与中心联系！", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(!帐户冻结, "否") = "是" Then
'            Call IC_End(True)
            MsgBox "该病人的帐户已经被冻结，只能以现金结算！" & vbCrLf & "冻结原因：" & Nvl(!冻结原因) & vbCrLf & "冻结说明：" & Nvl(!冻结说明) & vbCrLf & "冻结时间：" & Nvl(!冻结时间), vbInformation, gstrSysName
            bln自费病人 = True
            Exit Function
        End If
        '将中心的有效卡号回写入公共变量，用于结算的判断，虚拟结算不再判断有效卡号
        gCominfo_毕节.str有效卡号 = Val(Nvl(!有效卡号, 0))
'        If Val(Nvl(gCominfo_毕节.str有效卡号, 0)) <> Val(Nvl(!有效卡号, 0)) Then
''            Call IC_End(True)
'            MsgBox "当前的IC卡片是一张无效的卡！", vbInformation, gstrSysName
'            Exit Function
'        End If
    End With
    If gCominfo_毕节.blnOnLine Then gobjCenter.CloseConnector
    
    lng主页ID = rsExse!主页ID
    '将病人在院相关状态上传到中心
    If Not 病人变动记录上传_毕节(lng病人ID, lng主页ID) Then
'        Call IC_End(True)
        Exit Function
    End If
    
    If Not 事务_开始() Then
'        Call IC_End(True)
        Exit Function
    End If
    cnOracle.BeginTrans         '上传了明细就打上标记，单独提交
    blnTrans = True
    
    '根据传入记录集计算进入统筹金额（只计算未计算部分明细，那当然，未上传的肯定就没有计算）
    With rsExse
        Do While Not .EOF
            dbl金额 = Nvl(!金额, 0)
            If Nvl(!是否上传, 0) = 0 Then
                '计算统筹金额
                str项目编码 = "": str项目名称 = "": str大类 = ""
                bln药品 = (InStr(1, "5,6,7", !收费类别) <> 0)
                dbl数量 = !数量     '数量不可能为零或空
                dbl金额 = Nvl(!金额, 0)
                dbl单价 = Nvl(!金额, 0) / dbl数量
                
                '取该医保项目相关信息
                gstrSQL = " Select A.项目编码,A.项目名称,B.名称,A.附注 As 大类 From 保险支付项目 A,收费细目 B " & _
                          " Where B.ID=[1] And B.ID=A.收费细目ID(+) And A.险类(+)=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该医保项目相关信息", CLng(!收费细目ID), TYPE_毕节)
                If rsTemp.RecordCount <> 0 Then
                    str项目编码 = Nvl(rsTemp!项目编码)
                    str项目名称 = Nvl(rsTemp!项目名称)
                    str大类 = Nvl(rsTemp!大类)
                End If
                
                dbl进入统筹 = Calc统筹金额_明细(bln药品, str项目编码, str项目名称, dbl数量, dbl单价)
                
                '写入中间库
                If bln药品 Then
                    gstrSQL = "" & _
                        " INSERT INTO 药品费用明细表" & _
                        " (ID,社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                        " VALUES" & _
                        " (药品费用明细表_ID.Nextval,'" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
                        "'" & IIf(str项目编码 = "", gstr药品代码, str项目编码) & "','" & Nvl(rsTemp!项目名称, rsTemp!名称) & "','" & IIf(str大类 = "", gstr药品大类, str大类) & "'," & _
                        "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!金额, 0) & "," & _
                        "" & dbl进入统筹 & ",0," & Nvl(!金额, 0) - dbl进入统筹 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & !医生 & "','否','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                    gcnGYBJYB.Execute gstrSQL
                Else
                    gstrSQL = "" & _
                        " INSERT INTO 诊疗费用明细表" & _
                        " (ID,社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                        " VALUES" & _
                        " (诊疗费用明细表_ID.Nextval,'" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
                        "'" & IIf(str项目编码 = "", gstr诊疗代码, str项目编码) & "','" & Nvl(rsTemp!项目名称, rsTemp!名称) & "','" & IIf(str大类 = "", gstr诊疗大类, str大类) & "'," & _
                        "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!金额, 0) & "," & _
                        "" & dbl进入统筹 & ",0," & Nvl(!金额, 0) - dbl进入统筹 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & !医生 & "','否','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                    gcnGYBJYB.Execute gstrSQL
                End If
                
                '写医保中心库（由于住院登记表中已有病种名称，科室，床位号，入院时间与出院时间，因此，这两张表中相同内容不再填写）
                If gCominfo_毕节.blnOnLine Then
                    If bln药品 Then
                        gstrSQL = "" & _
                            " INSERT INTO 住院未结算药品费用日记帐表" & _
                            " (社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                            " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                            " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                            " VALUES" & _
                            " ('" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
                            "'" & IIf(str项目编码 = "", gstr药品代码, str项目编码) & "','" & Nvl(rsTemp!项目名称, rsTemp!名称) & "','" & IIf(str大类 = "", gstr药品大类, str大类) & "'," & _
                            "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!金额, 0) & "," & _
                            "" & dbl进入统筹 & ",0," & Nvl(!金额, 0) - dbl进入统筹 & "," & _
                            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & Nvl(!医生, "ZLHIS") & "','否','是')"
                        If Not ExecuteSQL(gstrSQL) Then
'                            Call IC_End(True)
                            Call 事务_回滚
                            cnOracle.RollbackTrans
                            Exit Function
                        End If
                    Else
                        gstrSQL = "" & _
                            " INSERT INTO 住院未结算诊疗费用日记帐表" & _
                            " (社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                            " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                            " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                            " VALUES" & _
                            " ('" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
                            "'" & IIf(str项目编码 = "", gstr诊疗代码, str项目编码) & "','" & Nvl(rsTemp!项目名称, rsTemp!名称) & "','" & IIf(str大类 = "", gstr诊疗大类, str大类) & "'," & _
                            "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!金额, 0) & "," & _
                            "" & dbl进入统筹 & ",0," & Nvl(!金额, 0) - dbl进入统筹 & "," & _
                            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & Nvl(!医生, "ZLHIS") & "','否','是')"
                        If Not ExecuteSQL(gstrSQL) Then
                            'Call IC_End(True)
                            Call 事务_回滚
                            cnOracle.RollbackTrans
                            Exit Function
                        End If
                    End If
                End If
                    
                '打上传标记
                gstrSQL = "zl_病人费用记录_上传('" & !NO & "'," & !序号 & "," & !记录性质 & "," & !记录状态 & ")"
                cnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            gCominfo_毕节.dbl费用总额 = gCominfo_毕节.dbl费用总额 + dbl金额
            .MoveNext
        Loop
    End With
    
    If 事务_提交 Then
        cnOracle.CommitTrans
    Else
        Call 事务_回滚
        cnOracle.RollbackTrans
'        Call IC_End(True)
        Exit Function
    End If
    blnTrans = False
    
    '从中间库中取所有未结算的费用明细的统筹金额
    gstrSQL = " Select Sum(Nvl(统筹基金金额,0)) 统筹金额" & _
              " From 药品费用明细表 A,入院登记表 B" & _
              " Where A.社会保障号=B.社会保障号 And Nvl(A.是否结算,'否')='否' And A.发生类别='住院' And A.住院号=B.住院号" & _
              " And A.社会保障号='" & gCominfo_毕节.str社会保障号 & "' And A.住院号='" & gCominfo_毕节.str就诊流水号 & "'"
    gstrSQL = gstrSQL & " Union All" & _
              " Select Sum(Nvl(统筹基金金额,0)) 统筹金额" & _
              " From 诊疗费用明细表 A,入院登记表 B" & _
              " Where A.社会保障号=B.社会保障号 And Nvl(A.是否结算,'否')='否' And A.发生类别='住院' And A.住院号=B.住院号" & _
              " And A.社会保障号='" & gCominfo_毕节.str社会保障号 & "' And A.住院号='" & gCominfo_毕节.str就诊流水号 & "'"
    gstrSQL = " Select Sum(统筹金额) 统筹金额 From (" & gstrSQL & ")"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        gCominfo_毕节.dbl统筹金额 = !统筹金额
    End With
    
    '根据选择的病种再次计算进入统筹金额
    gCominfo_毕节.dbl统筹金额 = Calc统筹金额_病种(gCominfo_毕节.dbl统筹金额, lng病人ID)
    
    '取年度累计
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Select IsNull(余额,0) 余额,IsNull(当年住院费用,0) 年度统筹累计 " & _
                  " From 个人帐户余额表 " & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'"
    Else
        gstrSQL = " Select Nvl(余额,0) 余额,nvl(当年住院费用,0) 年度统筹累计 " & _
                  " From 个人帐户余额表 " & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'"
    End If
    
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call gobjCenter.CloseConnector
'            Call IC_End(True)
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    With rsTemp
        If .RecordCount <> 0 Then
            gCominfo_毕节.dbl年度统筹 = !年度统筹累计
            dbl个人帐户 = !余额
        End If
    End With
    
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.CloseConnector
    End If
    
    '再根据病人的参保险种，就业情况等，分档计算得出最终的统筹报销金额
    '如果是按年度起付，则本年度只扣一次起付线，否则每次结算都要扣起付线，而累计是始终要用的
    'If gCominfo_毕节.blnYearBase Then gCominfo_毕节.dbl年度报销 = Calc统筹金额_分档(gCominfo_毕节.dbl年度统筹, gCominfo_毕节.str社会保障号)
    'gCominfo_毕节.dbl统筹报销 = Calc统筹金额_分档(gCominfo_毕节.dbl统筹金额 + IIf(gCominfo_毕节.blnYearBase, gCominfo_毕节.dbl年度统筹, 0), gCominfo_毕节.str社会保障号)
    gCominfo_毕节.dbl统筹报销 = Calc统筹金额_分档(gCominfo_毕节.dbl统筹金额, gCominfo_毕节.str社会保障号)
    
    '实际本次结算的统筹报销金额
    'gCominfo_毕节.dbl统筹报销 = gCominfo_毕节.dbl统筹报销 - gCominfo_毕节.dbl年度报销
    gCominfo_毕节.dbl统筹报销 = gCominfo_毕节.dbl统筹报销
    
    '计算个人帐户支付额
    dbl个人自付 = gCominfo_毕节.dbl费用总额 - gCominfo_毕节.dbl统筹报销
    dbl个人帐户 = IIf(dbl个人帐户 >= dbl个人自付, dbl个人自付, dbl个人帐户)
    
    If bln自费病人 Then
        dbl个人自付 = gCominfo_毕节.dbl费用总额
        dbl个人帐户 = 0
        gCominfo_毕节.dbl统筹报销 = 0
        gCominfo_毕节.dbl统筹金额 = 0
    End If
    
    '住院虚拟结算_毕节 = "个人帐户;" & dbl个人帐户 & ";1"
    住院虚拟结算_毕节 = "个人帐户;0;1"
    住院虚拟结算_毕节 = 住院虚拟结算_毕节 & "|医保基金;" & gCominfo_毕节.dbl统筹报销 & ";0"
    
'    Call IC_End(True)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        Call 事务_回滚
        cnOracle.RollbackTrans
    End If
'    Call IC_End(True)
End Function

Public Function 住院结算_毕节(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
    On Error GoTo errHand
    Dim STR姓名 As String
    Dim strCard As String
    Dim blnTrans As Boolean
    Dim dbl个人帐户 As Double
    Dim lng主页ID As Long
    Dim int住院次数 As Integer
    Dim rsTemp As New ADODB.Recordset
    '医保要求出院结算后必须自动出院，但HIS中有参数控制，需要实施人员注意
    
    '读卡
    If Not gobjCenter.IC_ReadCard(strCard) Then Exit Function
    Call 数据转换_毕节(strCard, True)
    If Not (IC_Data_毕节.有效卡号 = gCominfo_毕节.str有效卡号 And IC_Data_毕节.社会保障号 = gCominfo_毕节.str社会保障号) Then
        Call IC_End(True)
        Err.Raise 9000, gstrSysName, "当前IC卡不是该病人的卡或该卡已失效，请与医保中心联系！", vbInformation, gstrSysName
        Exit Function
    End If
    STR姓名 = IC_Data_毕节.姓名
    
    '取病人的主页ID
    gstrSQL = "Select nvl(住院次数,0) AS 主页ID From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的主页ID", lng病人ID)
    lng主页ID = rsTemp!主页ID
    
    '取病人的住院次数
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = "Select IsNull(住院次数,0) 住院次数 From 个人帐户余额表 Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
    Else
        gstrSQL = "Select Nvl(住院次数,0) 住院次数 From 个人帐户余额表 Where 社会保障号='" & IC_Data_毕节.社会保障号 & "'"
    End If
    
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then
            Call gobjCenter.CloseConnector
            Call IC_End(True)
            Exit Function
        End If
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    int住院次数 = rsTemp!住院次数
    
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.CloseConnector
    End If
    
    '取本次结算实际个人帐户支付额
    gstrSQL = "Select Nvl(A.冲预交,0) 个人帐户 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=[2]" & _
        " And A.结算方式 in ('个人帐户') And A.记录性质=2 And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次个人帐户支付额", lng结帐ID, TYPE_毕节)
    dbl个人帐户 = 0
    If Not rsTemp.EOF Then
        dbl个人帐户 = rsTemp!个人帐户
    End If
    
    If Not 事务_开始 Then
        Call IC_End(True)
        Exit Function
    End If
    blnTrans = True
    
    '先填写保险结算记录
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_毕节 & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & int住院次数 & "," & 0 & "," & 0 & "," & 0 & "," & _
        gCominfo_毕节.dbl费用总额 & "," & gCominfo_毕节.dbl费用总额 - gCominfo_毕节.dbl统筹报销 - dbl个人帐户 & ",0," & _
        gCominfo_毕节.dbl统筹金额 & "," & gCominfo_毕节.dbl统筹报销 & ",0,0," & dbl个人帐户 & ",'" & gCominfo_毕节.str就诊流水号 & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存保险结算记录")
    
    '将中间库中该病人所有费用明细的结算标志填上
    gstrSQL = " Update 药品费用明细表 " & _
              " Set 是否结算='是' " & _
              " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
              " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
    gcnGYBJYB.Execute gstrSQL
    gstrSQL = " Update 诊疗费用明细表 " & _
              " Set 是否结算='是' " & _
              " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
              " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
    gcnGYBJYB.Execute gstrSQL
    
    gstrSQL = "" & _
        " INSERT INTO 医疗费用支出明细表" & _
        " (ID,社会保障号,姓名,门诊住院号,发生时间,发生类别,总费用,统筹基金支付," & _
        " 个人帐户支付,个人自付金额,医疗机构代码,医疗机构名称,操作员,是否上传)" & _
        " VALUES" & _
        " (医疗费用支出明细表_ID.Nextval,'" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
        "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','住院'," & gCominfo_毕节.dbl费用总额 & "," & _
        gCominfo_毕节.dbl统筹报销 & "," & dbl个人帐户 & "," & gCominfo_毕节.dbl费用总额 - gCominfo_毕节.dbl统筹报销 - dbl个人帐户 & "," & _
        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_毕节.blnOnLine Then
        '将中心的未结算明细转入结算明细表中
        gstrSQL = "" & _
            " INSERT INTO 药品费用明细表" & _
            "     (社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
            "     发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
            "     个人自付金额,医疗机构代码,医疗机构名称,操作员)" & _
            " Select 社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
            "     发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
            "     个人自付金额,医疗机构代码,医疗机构名称,操作员" & _
            " From 住院未结算药品费用日记帐表" & _
            " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
            " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        gstrSQL = "" & _
            " INSERT INTO 诊疗费用明细表" & _
            "     (社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
            "     发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
            "     个人自付金额,医疗机构代码,医疗机构名称,操作员)" & _
            " Select 社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
            "     发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
            "     个人自付金额,医疗机构代码,医疗机构名称,操作员" & _
            " From 住院未结算诊疗费用日记帐表" & _
            " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
            " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        '删除所有未结算明细
        gstrSQL = " Delete 住院未结算药品费用日记帐表" & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
                  " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        gstrSQL = " Delete 住院未结算诊疗费用日记帐表" & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'" & _
                  " And 住院号='" & gCominfo_毕节.str就诊流水号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
        
        '结算时将本次结算明细写入中心库
        gstrSQL = "" & _
            " INSERT INTO 医疗费用支出明细表" & _
            " (社会保障号,姓名,门诊住院号,发生时间,发生类别,总费用,统筹基金支付," & _
            " 个人帐户支付,个人自付金额,医疗机构代码,医疗机构名称,操作员,险种名称)" & _
            " VALUES" & _
            " ('" & gCominfo_毕节.str社会保障号 & "','" & STR姓名 & "','" & gCominfo_毕节.str就诊流水号 & "'," & _
            "'" & Format(zlDatabase.Currentdate, "yyyy.MM.dd HH:mm:ss") & "','住院'," & gCominfo_毕节.dbl费用总额 & "," & _
            gCominfo_毕节.dbl统筹报销 & "," & dbl个人帐户 & "," & gCominfo_毕节.dbl费用总额 - gCominfo_毕节.dbl统筹报销 - dbl个人帐户 & "," & _
            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','" & gCominfo_毕节.str险种名称 & "')"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    '更新入院登记表
    gstrSQL = " Update 入院登记表" & _
              " Set 是否结算='是'" & _
              " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "' And 出院时间 Is NULL"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Update 住院登记表" & _
                  " Set 是否结算='是'" & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "' And IsNull(出院时间,'')=''"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    '下帐户（中心库的个人帐户余额表要更新，中间库的个人帐户余额表要更新）
    '住院还需要更新当前住院费用字段（统筹支付、住院次数）
    IC_Data_毕节.个人帐户余额 = IC_Data_毕节.个人帐户余额 - dbl个人帐户
    IC_Data_毕节.最后就诊日期 = Format(zlDatabase.Currentdate, "yyyy.MM.dd")
    IC_Data_毕节.最后就诊医院代码 = gCominfo_毕节.strHospitalCode
    Call 数据转换_毕节(strCard, False)
    
    '在出院时，自动增加住院次数 Set 住院次数=Nvl(住院次数,0)+1
    gstrSQL = " Update 个人帐户余额表 " & _
              " Set 余额=Nvl(余额,0)-" & Val(dbl个人帐户) & "," & _
              "     当年住院费用=Nvl(当年住院费用,0)+" & gCominfo_毕节.dbl统筹金额 & _
              " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'"
    gcnGYBJYB.Execute gstrSQL
    If gCominfo_毕节.blnOnLine Then 'SqlServer非空函数是IsNULL();而Oracle是Nvl()
        gstrSQL = " Update 个人帐户余额表 " & _
                  " Set 本年支出=IsNull(本年支出,0)+" & Val(dbl个人帐户) & "," & _
                  "     累计支出=IsNull(累计支出,0)+" & Val(dbl个人帐户) & "," & _
                  "     余额=IsNull(余额,0)-" & Val(dbl个人帐户) & "," & _
                  "     当年住院费用=IsNull(当年住院费用,0)+" & gCominfo_毕节.dbl统筹金额 & _
                  " Where 社会保障号='" & gCominfo_毕节.str社会保障号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Call IC_End(True): Exit Function
    End If
    
    If Not gobjCenter.IC_WriteCard(strCard) Then
        Call 事务_回滚
        Call IC_End(True)
        Exit Function
    End If
    
    If 事务_提交 Then
        住院结算_毕节 = True
    Else
        Call 事务_回滚
    End If
    blnTrans = False
    
    Call IC_End
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
    If blnTrans Then Call 事务_回滚
    Call IC_End(True)
End Function

Public Function 住院结算冲销_毕节(lng结帐ID As Long) As Boolean
    '只有门诊允许结算作废，住院不允许
    ErrMsgBox "医保接口不支持作废住院结算单！", vbInformation, gstrSysName
    住院结算冲销_毕节 = False
End Function

Public Function 处方上传_毕节(ByVal int性质 As Integer, ByVal int状态 As Integer, ByVal str单据号 As String) As Boolean
    Dim blnTrans As Boolean                 '当前是否处于事务中
    Dim blnInsure As Boolean                '本次是否做为医保病人的身份进行就诊
    Dim bln药品 As Boolean
    Dim int序号 As Integer
    Dim lng病人ID As Long
    Dim dbl统筹金额 As Double
    Dim str就诊登记号 As String, str社会保障号 As String, STR姓名 As String
    Dim rsDetail As New ADODB.Recordset
    Dim rsInsure As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '将本次产生的处方全部产生到中心库的未结日记帐表及中间库的明细表中
    If Not 事务_开始 Then Exit Function
    gcnOracle.BeginTrans
    blnTrans = True
    
    gstrSQL = " Select A.病人ID,A.收费类别,A.收费细目ID,A.序号,A.发生时间," & _
              " round(A.实收金额,2) 实收金额,A.实收金额/(A.付数*A.数次) As 单价,(A.付数*A.数次) AS 数量," & _
              " B.项目编码,B.项目名称,B.附注 As 大类,C.名称" & _
              " From 住院费用记录 A,保险支付项目 B,收费细目 C" & _
              " Where A.收费细目ID=C.ID And A.收费细目ID=B.收费细目ID(+) And B.险类(+)=[1]" & _
              " And A.记录性质=[2] And A.记录状态=[3] And A.NO=[4]" & _
              " And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0 And Nvl(A.是否上传,0)=0"
    Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "提取本张处方明细", TYPE_毕节, int性质, int状态, str单据号)
    
    With rsDetail
        Do While Not .EOF
            If lng病人ID <> !病人ID Then
                '检查本次是否以医保身份入院
                gstrSQL = "Select Count(*) Records From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.病人ID=[1] And A.主页ID=B.住院次数 And A.险类=[2]"
                Set rsInsure = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否医保病人", CLng(!病人ID), TYPE_毕节)
                blnInsure = (rsInsure!Records = 1)
                If blnInsure Then
                    lng病人ID = !病人ID
                    '提取病人的社会保障号
                    gstrSQL = "Select 医保号 As 社会保障号 From 保险帐户 Where 病人ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人的社会保障号", lng病人ID)
                    str社会保障号 = Nvl(rsTemp!社会保障号)
                    '取本次入院的就诊登记号(从中间库的入院登记表中取)
                    gstrSQL = " Select B.医保号 As 社会保障号,C.姓名,A.住院号" & _
                              " From " & gCominfo_毕节.用户名 & ".入院登记表 A,保险帐户 B,病人信息 C" & _
                              " Where A.社会保障号=B.医保号 And A.出院时间 Is NULL And B.病人ID=C.病人ID" & _
                              " And B.险类=" & TYPE_毕节 & " And B.病人ID=" & lng病人ID
                    With rsTemp
                        If .State = 1 Then .Close
                        .Open gstrSQL, gcnOracle
                        If .RecordCount = 0 Then
                            MsgBox "没有找到该病人[社会保障号：" & str社会保障号 & "]的有效在院记录,可能该病人已经出院！", vbInformation, gstrSysName
                            Call 事务_回滚
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                        str就诊登记号 = Nvl(!住院号)
                        str社会保障号 = Nvl(!社会保障号)
                        STR姓名 = Nvl(!姓名)
                    End With
                    
                    '判断卡的有效性，如果无效，判断预交余额是否大于所有费用总额，如果小于提示并不准保存单据
                    Dim str冻结原因 As String, str冻结说明 As String, str冻结时间 As String
                    If Not CheckCard(str社会保障号, str冻结原因, str冻结说明, str冻结时间) Then
                        If Not BalanceLack(lng病人ID) Then
                            MsgBox "该病人[社会保障号：" & str社会保障号 & "]的卡已被冻结，必须全部缴现金，而预交款不足，请缴款！" & vbCrLf & _
                            "冻结原因：" & str冻结原因 & vbCrLf & "冻结说明：" & str冻结说明 & vbCrLf & "冻结时间：" & str冻结时间, vbInformation, gstrSysName
                            Call 事务_回滚
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            If blnInsure Then
                int序号 = !序号
                bln药品 = (InStr(1, "5,6,7", !收费类别) <> 0)
                '计算当笔明细的进入统筹金额
                dbl统筹金额 = Calc统筹金额_明细(bln药品, Nvl(!项目编码), Nvl(!项目名称), Nvl(!单价, 0), Nvl(!数量, 0))
                
                '写中间库
                If bln药品 Then
                    gstrSQL = "" & _
                        " INSERT INTO 药品费用明细表" & _
                        " (ID,社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                        " VALUES" & _
                        " (药品费用明细表_ID.Nextval,'" & str社会保障号 & "','" & STR姓名 & "','" & str就诊登记号 & "'," & _
                        "'" & Nvl(!项目编码, gstr药品代码) & "','" & Nvl(!项目名称, !名称) & "','" & Nvl(!大类, gstr药品大类) & "'," & _
                        "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!实收金额, 0) & "," & _
                        "" & dbl统筹金额 & ",0," & Nvl(!实收金额, 0) - dbl统筹金额 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','否','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                    gcnGYBJYB.Execute gstrSQL
                Else
                    gstrSQL = "" & _
                        " INSERT INTO 诊疗费用明细表" & _
                        " (ID,社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                        " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                        " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                        " VALUES" & _
                        " (诊疗费用明细表_ID.Nextval,'" & str社会保障号 & "','" & STR姓名 & "','" & str就诊登记号 & "'," & _
                        "'" & Nvl(!项目编码, gstr诊疗代码) & "','" & Nvl(!项目名称, !名称) & "','" & Nvl(!大类, gstr诊疗大类) & "'," & _
                        "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!实收金额, 0) & "," & _
                        "" & dbl统筹金额 & ",0," & Nvl(!实收金额, 0) - dbl统筹金额 & "," & _
                        "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','否','" & IIf(gCominfo_毕节.blnOnLine, "是", "否") & "')"
                    gcnGYBJYB.Execute gstrSQL
                End If
                
                '写医保中心库（由于住院登记表中已有病种名称，科室，床位号，入院时间与出院时间，因此，这两张表中相同内容不再填写）
                If gCominfo_毕节.blnOnLine Then
                    If bln药品 Then
                        gstrSQL = "" & _
                            " INSERT INTO 住院未结算药品费用日记帐表" & _
                            " (社会保障号,姓名,住院号,药品代码,药品名称,药品大类," & _
                            " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                            " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                            " VALUES" & _
                            " ('" & str社会保障号 & "','" & STR姓名 & "','" & str就诊登记号 & "'," & _
                            "'" & Nvl(!项目编码, gstr药品代码) & "','" & Nvl(!项目名称, !名称) & "','" & Nvl(!大类, gstr药品大类) & "'," & _
                            "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!实收金额, 0) & "," & _
                            "" & dbl统筹金额 & ",0," & Nvl(!实收金额, 0) - dbl统筹金额 & "," & _
                            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','否','是')"
                        If Not ExecuteSQL(gstrSQL) Then
                            Call 事务_回滚
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                    Else
                        gstrSQL = "" & _
                            " INSERT INTO 住院未结算诊疗费用日记帐表" & _
                            " (社会保障号,姓名,住院号,诊疗项目代码,诊疗项目名称,费用类别," & _
                            " 发生时间,发生类别,总费用,统筹基金金额,个人帐户金额," & _
                            " 个人自付金额,医疗机构代码,医疗机构名称,操作员,是否结算,是否上传)" & _
                            " VALUES" & _
                            " ('" & str社会保障号 & "','" & STR姓名 & "','" & str就诊登记号 & "'," & _
                            "'" & Nvl(!项目编码, gstr诊疗代码) & "','" & Nvl(!项目名称, !名称) & "','" & Nvl(!大类, gstr诊疗大类) & "'," & _
                            "'" & Format(!发生时间, "yyyy.MM.dd HH:mm:ss") & "','住院'," & Nvl(!实收金额, 0) & "," & _
                            "" & dbl统筹金额 & ",0," & Nvl(!实收金额, 0) - dbl统筹金额 & "," & _
                            "'" & gCominfo_毕节.strHospitalCode & "','" & gCominfo_毕节.strHospitalName & "','" & UserInfo.姓名 & "','否','是')"
                        If Not ExecuteSQL(gstrSQL) Then
                            Call 事务_回滚
                            gcnOracle.RollbackTrans
                            Exit Function
                        End If
                    End If
                End If
                
                '打上传标记
                gstrSQL = "zl_病人费用记录_上传('" & str单据号 & "'," & int序号 & "," & int性质 & "," & int状态 & ")"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
            .MoveNext
        Loop
    End With
    
    If 事务_提交 Then
        gcnOracle.CommitTrans
        处方上传_毕节 = True
    Else
        Call 事务_回滚
        gcnOracle.RollbackTrans
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then
        Call 事务_回滚
        gcnOracle.RollbackTrans
    End If
End Function

Public Function 更新病种_毕节(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim lng病种ID As Long
    Dim blnTrans As Boolean
    Dim str社会保障号 As String
    Dim str病种名称 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '更新病种
    If Not frm病种选择_毕节.病种选择(lng病人ID, lng病种ID, str病种名称) Then Exit Function
    
    If Not 事务_开始 Then Exit Function
    gcnOracle.BeginTrans
    blnTrans = True
    
    '更新保险帐户
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_毕节 & ",'病种ID','" & lng病种ID & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种信息")
    
    '更新中间库与中心库
    gstrSQL = " Update 入院登记表 " & _
              " Set 病种名称='" & str病种名称 & "'" & _
              " Where 社会保障号='" & str社会保障号 & "'"
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = " Update 住院登记表 " & _
                  " Set 病种名称='" & str病种名称 & "'" & _
                  " Where 社会保障号='" & str社会保障号 & "'"
        If Not ExecuteSQL(gstrSQL) Then Exit Function
    End If
    
    If 事务_提交 Then
        gcnOracle.CommitTrans
        更新病种_毕节 = True
    Else
        gcnOracle.RollbackTrans
        Call 事务_回滚
    End If
    
    blnTrans = False
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call 事务_回滚
    End If
End Function

Public Function 病人变动记录上传_毕节(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln开始事务 As Boolean = True) As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    '当病人科室，床位、主治医生，护理等级发生变化时触发该事件（本医保只关心科室、床位、预交款与病种名称）
    gstrSQL = " Select A.医保号 As 社会保障号,to_Char(B.出院日期,'yyyy.MM.dd') AS 出院日期," & _
              " D.名称 As 当前科室,C.当前床号,E.病种名称,Sum(Nvl(F.金额,0)) As 预交款" & _
              " From 保险帐户 A,病案主页 B,病人信息 C,部门表 D," & gCominfo_毕节.用户名 & ".病种目录表 E,病人预交记录 F" & _
              " Where A.险类=[3] And A.病人ID=[1] And B.主页ID=[2] ANd A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.主页ID=C.住院次数" & _
              " And A.病种ID=E.ID(+) And C.当前科室ID=D.ID(+) And B.病人ID=F.病人ID(+) And B.主页ID=F.主页ID(+) And F.记录性质(+)=1 " & _
              " Group by A.医保号,B.出院日期,D.名称,C.当前床号,B.登记人,E.病种名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次入院相关信息", lng病人ID, lng主页ID, TYPE_毕节)
    
    If bln开始事务 Then
        blnTrans = True
        If Not 事务_开始 Then Exit Function
    End If
    
    '修改入院登记记录(中间库与中心库)
    '由于出院时已无当前科室与床位号，因此，只要出院时间不为空，则不再更新科室与床位号
    If Nvl(rsTemp!出院日期) = "" Then
        gstrSQL = " Update 入院登记表" & _
                  " Set 病种名称='" & Nvl(rsTemp!病种名称) & "'," & _
                  "     科室='" & Nvl(rsTemp!当前科室) & "'," & _
                  "     床位号='" & ToVarchar(Nvl(rsTemp!当前床号), 4) & "'," & _
                  "     预付款=" & Nvl(rsTemp!预交款, 0) & _
                  " Where 社会保障号='" & rsTemp!社会保障号 & "' And 出院时间 Is NULL"
    Else
        gstrSQL = " Update 入院登记表" & _
                  " Set 病种名称='" & Nvl(rsTemp!病种名称) & "'," & _
                  "     出院时间='" & Nvl(rsTemp!出院日期) & "'" & _
                  " Where 社会保障号='" & rsTemp!社会保障号 & "' And 出院时间 Is NULL"
    End If
    gcnGYBJYB.Execute gstrSQL
    
    If gCominfo_毕节.blnOnLine Then
        If Nvl(rsTemp!出院日期) = "" Then
            gstrSQL = " Update 住院登记表" & _
                      " Set 病种名称='" & Nvl(rsTemp!病种名称) & "'," & _
                      "     科室='" & Nvl(rsTemp!当前科室) & "'," & _
                      "     床位号='" & ToVarchar(Nvl(rsTemp!当前床号), 4) & "'," & _
                      "     预付款=" & Nvl(rsTemp!预交款, 0) & _
                      " Where 社会保障号='" & rsTemp!社会保障号 & "' And IsNull(出院时间,'')=''"
        Else
            gstrSQL = " Update 住院登记表" & _
                      " Set 病种名称='" & Nvl(rsTemp!病种名称) & "'," & _
                      "     出院时间='" & Nvl(rsTemp!出院日期) & "'" & _
                      " Where 社会保障号='" & rsTemp!社会保障号 & "' And IsNull(出院时间,'')=''"
        End If
        If Not ExecuteSQL(gstrSQL, bln开始事务) Then Exit Function
    End If
    
    If bln开始事务 Then
        If 事务_提交 Then
            病人变动记录上传_毕节 = True
        Else
            Call 事务_回滚
        End If
    Else
        病人变动记录上传_毕节 = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then Call 事务_回滚
End Function

Private Function Calc统筹金额_明细(ByVal bln药品 As Boolean, ByVal str项目编码 As String, ByVal str项目名称 As String, _
    ByVal dbl单价 As Double, ByVal dbl数量 As Double) As Double
    Dim dbl统筹金额 As Double, dbl自付比例 As Double, dbl起付金额 As Double, dbl限价 As Double
    Dim dbl金额 As Double
    Dim bln冲销 As Boolean
    Dim str医院级别 As String
    Dim rsCalc As New ADODB.Recordset
    '计算单笔明细的进入统筹金额
    
    '先取医院级别
    If Not bln药品 Then
        gstrSQL = "Select 级别 From 医疗机构基本情况表 Where 单位代码='" & gCominfo_毕节.strHospitalCode & "'"
        With rsCalc
            If .State = 1 Then .Close
            .Open gstrSQL, gcnGYBJYB
            str医院级别 = !级别
        End With
    End If
    
    '再取项目的基本信息
    dbl金额 = dbl单价 * dbl数量
    bln冲销 = (dbl金额 < 0)
    gstrSQL = " Select nvl(个人自付比例,0) As 自付比例,Nvl(个人起付金额,0) As 起付金额" & _
              "" & IIf(bln药品, "", ",nvl(一级医院定额,0) 一级,Nvl(二级医院定额,0) 二级,Nvl(三级医院定额,0) 三级") & _
              " From " & IIf(bln药品, "药品目录表", "诊疗项目表") & _
              " Where " & IIf(bln药品, "药品代码", "诊疗项目代码") & "='" & str项目编码 & "'" & _
              " And " & IIf(bln药品, "中文名称", "诊疗项目名称") & "='" & str项目名称 & "'"
    With rsCalc
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            '如果是未对码的项目，如果系统参数“住院启用药品诊疗自付”为真，则该明细为全自付，否则全部进入统筹
            dbl自付比例 = IIf(gCominfo_毕节.blnPhysicCash, 100, 0)
            dbl起付金额 = 0
        Else
            dbl自付比例 = IIf(gCominfo_毕节.blnPhysicCash, !自付比例, 0)
            dbl起付金额 = IIf(gCominfo_毕节.blnPhysicCash, !起付金额, 0)
        End If
        
        '先判断实际单价，如果超过限价，以限价为准
        If Not bln药品 And .RecordCount > 0 Then
            dbl限价 = IIf(str医院级别 = "一级医院", !一级, IIf(str医院级别 = "二级医院", !二级, !三级))
        End If
        
        '重新计算本笔明细的实际金额
        dbl金额 = Abs(IIf(dbl单价 >= dbl限价 And dbl限价 <> 0, dbl限价, dbl单价) * dbl数量) * IIf(bln冲销, -1, 1)
        
        '先扣起付金额，如果结果小于等于零，则直接退出
        dbl金额 = (Abs(dbl金额) - dbl起付金额) * IIf(bln冲销, -1, 1)
        If dbl金额 <= 0 And Not bln冲销 Then Exit Function
        Calc统筹金额_明细 = Round(dbl金额 * (100 - dbl自付比例) / 100, 2)
    End With
End Function

Private Function Calc统筹金额_病种(ByVal dbl进入统筹 As Double, ByVal lng病人ID As Long) As Double
    Dim dbl自付比例 As Double, dbl起付金额 As Double
    Dim lng病种ID As Long
    Dim rsCalc As New ADODB.Recordset
    '根据选择的病种再次计算进入统筹金额
    
    '先取该病人当前的病种ID
    gstrSQL = "Select Nvl(病种ID,0) AS 病种ID From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsCalc = zlDatabase.OpenSQLRecord(gstrSQL, "先取该病人当前的病种ID", TYPE_毕节, lng病人ID)
    lng病种ID = rsCalc!病种ID
    
    '再取该病种的属性
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = "Select IsNull(个人自付比例,0) AS 自付比例,IsNull(个人起付金额,0) AS 起付金额 From 病种目录表 Where ID=" & lng病种ID
    Else
        gstrSQL = "Select Nvl(个人自付比例,0) AS 自付比例,Nvl(个人起付金额,0) AS 起付金额 From 病种目录表 Where ID=" & lng病种ID
    End If
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsCalc) Then
            Call gobjCenter.CloseConnector
            Exit Function
        End If
    Else
        If rsCalc.State = 1 Then rsCalc.Close
        rsCalc.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsCalc
        If .RecordCount = 0 Then
            MsgBox "没有找到病种记录，无法完成住院结算，请与中心联系！", vbInformation, gstrSysName
            Exit Function
'            dbl自付比例 = IIf(gCominfo_毕节.blnDiseaseCash, 100, 0)
'            dbl起付金额 = 0
        Else
            dbl自付比例 = IIf(gCominfo_毕节.blnDiseaseCash, !自付比例, 0)
            dbl起付金额 = IIf(gCominfo_毕节.blnDiseaseCash, !起付金额, 0)
        End If
    End With
    
    '先扣去起付金额，再按比例计算统筹金额
    dbl进入统筹 = dbl进入统筹 - dbl起付金额
    If dbl进入统筹 <= 0 Then Exit Function
    dbl进入统筹 = dbl进入统筹 * (100 - dbl自付比例) / 100
    Calc统筹金额_病种 = dbl进入统筹
End Function

Private Function Calc统筹金额_分档(ByVal dbl统筹金额 As Double, ByVal str社会保障号 As String) As Double
    Dim intDO As Integer, intLoops As Integer   '用来循环计算的，如果是住院统筹，则只循环一次
    Dim blnMatch As Boolean
    
    Dim bln第一段 As Boolean
    Dim dbl实际起付线 As Double
    
    Dim str参加险种 As String
    Dim str险种名称 As String
    Dim str就业情况 As String
    Dim str医院级别 As String
    Dim str性别 As String
    Dim lng年龄 As Long
    Dim lng工龄 As Long
    
    Dim dbl年度统筹累计 As Double
    Dim dbl报销累计 As Double
    Dim dbl进入金额 As Double
    Dim dbl进入报销 As Double
    Dim dbl上限 As Double
    Dim dbl下限 As Double
    Dim dbl起付金额 As Double
    Dim dbl报销比例 As Double
    Dim rsBase As New ADODB.Recordset       '参保人基本信息
    Dim rsDisease As New ADODB.Recordset    '病种基本信息
    Dim rsRule As New ADODB.Recordset       '报销政策
    '按传入的统筹金额计算统筹报销金额
    
    gCominfo_毕节.str险种名称 = ""
    dbl年度统筹累计 = gCominfo_毕节.dbl年度统筹
    '先取出该病人的基本信息
    If gCominfo_毕节.blnOnLine Then
        gstrSQL = "" & _
            " Select C.性别,C.出生年月,C.参加工作时间,C.就业情况,C.参保险种,B.级别 AS 医院级别" & _
            " From 医疗机构基本情况表 B,参保职工基本情况表 C" & _
            " Where C.社会保障号='" & str社会保障号 & "' And B.单位代码='" & Trim(gstr医院编码) & "'"
    Else
        gstrSQL = "" & _
            " Select A.性别,A.出生年月,A.参加工作时间,A.就业情况,A.参保险种,B.级别 AS 医院级别" & _
            " From 个人帐户余额表 A,医疗机构基本情况表 B" & _
            " Where A.社会保障号='" & str社会保障号 & "' And B.单位代码='" & Trim(gstr医院编码) & "'"
    End If
    If gCominfo_毕节.blnOnLine Then
        Call gobjCenter.InitConnect("")
        If Not gobjCenter.GetRecordset(gstrSQL, rsBase) Then
            Call gobjCenter.CloseConnector
            Exit Function
        End If
    Else
        If rsBase.State = 1 Then rsBase.Close
        rsBase.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsBase
        If .RecordCount = 0 Then
            MsgBox "没有找到该病人的基本信息，无法进行结算！[提取病人基本信息]", vbInformation, gstrSysName
            Exit Function
        Else
            str参加险种 = Nvl(!参保险种, String(8, "0"))
            str参加险种 = Mid(str参加险种, 1, 4)        '只有前四位与计算相关
            str就业情况 = Trim(!就业情况)
            str医院级别 = Trim(!医院级别)
            str性别 = Trim(!性别)
            lng年龄 = GetAge(Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Replace(!出生年月, ".", "-"))
            lng工龄 = GetAge(Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Replace(!参加工作时间, ".", "-"))
        End If
    End With
    
    '因参保险种共八位，只有前四位有用，且前三位与第四位相排斥
    '提取所有险种（不排序，以原始顺序按位对应str参加险种）
    gstrSQL = "Select Rownum 序号,险种名称 From 参保险种表 Where 基本险种='医疗保险'" '下载时也只下了基本险种等于医疗保险的
    With rsDisease
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "参保险种不全，无法进行结算，请与中心联系！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    '准备进行分档计算
    intLoops = IIf(Right(str参加险种, 1) = "1", 4, 2)
    intDO = IIf(Right(str参加险种, 1) = "1", 4, 1)
    bln第一段 = True
    For intDO = intDO To intLoops
        If Mid(str参加险种, intDO, 1) = "1" Then
            rsDisease.Filter = "序号=" & intDO
            If rsDisease.RecordCount = 0 Then
                rsDisease.Filter = 0
                MsgBox "参保险种不全，无法进行结算，请与中心联系！", vbInformation, gstrSysName
                Exit Function
            End If
            str险种名称 = rsDisease!险种名称
            
            '提取医疗保险结算政策（取出来的可能多条，但符合所有条件的只可能有一条记录）
            gstrSQL = " Select 年龄段,工龄段,就业情况,费用段,起付金额,报销比例 " & _
                      " From 医疗费支付政策表" & _
                      " Where 险种名称='" & str险种名称 & "' And Nvl(性别,'" & str性别 & "')='" & str性别 & "' And Nvl(医院级别,'" & str医院级别 & "')='" & str医院级别 & "'" & _
                      " Order By 费用段"
            With rsRule
                If .State = 1 Then .Close
                .Open gstrSQL, gcnGYBJYB
                If .RecordCount > 0 Then
                    blnMatch = False
                    Do While Not .EOF
                        blnMatch = CheckMatch(Nvl(!年龄段, "00-99"), lng年龄, "-")
                        If blnMatch Then blnMatch = CheckMatch(Nvl(!工龄段, "00-99"), lng工龄, "-")
                        If blnMatch Then blnMatch = (Nvl(!就业情况, str就业情况) = str就业情况)
                        If blnMatch Then
                            '费用段的比较需要单独处理
                            dbl下限 = Split(!费用段, "-")(0)
                            dbl上限 = Split(!费用段, "-")(1)
                            dbl报销比例 = Nvl(!报销比例, 0)
                            dbl起付金额 = Nvl(!起付金额, 0)
                            blnMatch = ((dbl统筹金额 + dbl年度统筹累计) >= dbl下限 And dbl上限 > dbl年度统筹累计)
                            '不管第一段匹配否，都把起付线取出来
                            If bln第一段 Then
                                '如果是年度起付，且已支付起付线，本次起付为零，否则等于起付线
                                If gCominfo_毕节.blnYearBase Then
                                    If dbl年度统筹累计 >= dbl起付金额 Then
                                        dbl实际起付线 = 0
                                    Else
                                        dbl实际起付线 = dbl起付金额 - dbl年度统筹累计
                                        If dbl实际起付线 < 0 Then dbl实际起付线 = 0
                                    End If
                                Else
                                    dbl实际起付线 = dbl起付金额
                                End If
                                bln第一段 = False
                            End If
                        End If
                        
                        '找到对应的，就按规则进行计算（只有第一段的起付线有用）
                        If blnMatch Then
                            '记录第一次适应的险种名称
                            If gCominfo_毕节.str险种名称 = "" Then gCominfo_毕节.str险种名称 = str险种名称
                            '得出进入此段的金额
                            If (dbl统筹金额 + dbl年度统筹累计) <= dbl上限 Then
                                dbl上限 = (dbl统筹金额 + dbl年度统筹累计)
                            End If
                            If dbl年度统筹累计 >= dbl下限 Then
                                dbl进入金额 = dbl上限 - dbl年度统筹累计
                            Else
                                dbl进入金额 = dbl上限 - dbl下限
                            End If
                            If dbl进入金额 >= dbl实际起付线 Then
                                dbl进入金额 = dbl进入金额 - dbl实际起付线
                                dbl实际起付线 = 0
                            Else
                                dbl实际起付线 = dbl实际起付线 - dbl进入金额
                                dbl进入金额 = 0
                            End If
                            dbl进入报销 = dbl进入金额 * dbl报销比例
                            dbl报销累计 = dbl报销累计 + dbl进入报销
                        End If
                        .MoveNext
                    Loop
                End If
            End With
        End If
    Next
    rsDisease.Filter = 0
    
    Calc统筹金额_分档 = dbl报销累计
End Function

Private Function CheckMatch(ByVal str范围 As String, ByVal strValue As String, ByVal str分隔 As String) As Boolean
    'str缺省：str范围为空时的缺省值
    Dim arrData
    arrData = Split(str范围, str分隔)
    CheckMatch = (strValue >= Val(arrData(0)) And strValue <= Val(arrData(1)))
End Function

Private Function 获取连接方式() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '从中间库中提取连接方式
On Error GoTo ErrH
    gstrSQL = "Select NVL(连接方式,'脱机') 连接方式 From 医疗机构基本情况表 Where 单位代码='" & gCominfo_毕节.strHospitalCode & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 0 Then
            MsgBox "今天还未从中心下载，请先执行下载程序！[获取连接方式]", vbInformation, gstrSysName
            Exit Function
        Else
            gCominfo_毕节.blnOnLine = (!连接方式 <> "脱机")
        End If
    End With
    获取连接方式 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Private Function 检查连接密码() As Boolean
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrH
    '检查连接密码是否与中心一致,不同则禁止登录（因为每天都要下载，所以只从中间库中进行比较）
    gstrSQL = " Select 连接密码" & _
              " From 医疗机构基本情况表 " & _
              " Where 单位代码='" & gCominfo_毕节.strHospitalCode & "'"
    If gCominfo_毕节.blnOnLine Then
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then Exit Function
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "没有找到本医疗机构的基本信息，请与中心联系！[检查连接密码1]", vbInformation, gstrSysName
            Exit Function
        Else
            If gCominfo_毕节.strConnectPass <> Nvl(!连接密码) Then
                MsgBox "连接密码错误，可能中心已禁止本医疗机构使用，请与中心联系！[检查连接密码2]", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End With
    
    gstrSQL = " Select 单位名称,是否使用IC卡密码,住院启用病种自付,住院启用药品诊疗自付,住院按年度起付" & _
              " From 医疗机构基本情况表 Where 单位代码='" & gCominfo_毕节.strHospitalCode & "'"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnGYBJYB
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "没有找到本医疗机构的基本信息，请与中心联系！[检查连接密码3]", vbInformation, gstrSysName
            Exit Function
        End If
        
        '将系统参数赋予全局变量
        With gCominfo_毕节
            .strHospitalName = Nvl(rsTemp!单位名称)
            .blnICPassVerify = (InStr(1, "否,NO", UCase(Nvl(rsTemp!是否使用IC卡密码, "否"))) = 0)
            .blnDiseaseCash = (InStr(1, "否,NO", UCase(Nvl(rsTemp!住院启用病种自付, "否"))) = 0)
            .blnPhysicCash = (InStr(1, "否,NO", UCase(Nvl(rsTemp!住院启用药品诊疗自付, "否"))) = 0)
            .blnYearBase = (InStr(1, "否,NO", UCase(Nvl(rsTemp!住院按年度起付, "否"))) = 0)
        End With
    End With
    If gCominfo_毕节.strHospitalName = "" Then
        MsgBox "获取医疗机构名称失败！", vbInformation, gstrSysName
        Exit Function
    End If
'
'    '检查医院编码与名称是否与中心一致，必须出现不一致从而产生错误数据
'    If Not gobjCenter.CheckValid(gCominfo_毕节.strHospitalCode, gCominfo_毕节.strHospitalName) Then
'        MsgBox "医院编码错误，请检查医院编码！", vbInformation, gstrSysName
'        Exit Function
'    End If
    检查连接密码 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Private Function 获取中心连接() As Boolean
    Dim strCurDate As String
On Error GoTo ErrH
    If gCominfo_毕节.blnOnLine = False Then
        获取中心连接 = True
        Exit Function
    End If
    
    '先检查当前日期是否与第一次使用的日期相同，不同则不允许使用
    If mstrFirstStart <> "" Then
        strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        If mstrFirstStart <> strCurDate Then
            MsgBox "请重新启动程序，才能正常进行医保交易！[获取中心连接]", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '检查是否连的通中心
    If Not gobjCenter.InitConnect("") Then Exit Function
    获取中心连接 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Private Sub 关闭中心连接()
    If gCominfo_毕节.blnOnLine = False Then Exit Sub
    Call gobjCenter.CloseConnector
End Sub

Private Function 检查是否上传明细() As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If gCominfo_毕节.blnOnLine Then
        检查是否上传明细 = True
        Exit Function
    End If
    
    '检查是否存在未上传的明细
    gstrSQL = " Select 1 " & _
        " From 保险帐户 A,住院费用记录 B,病人信息 C" & _
        " Where A.险类=[1] And A.病人ID=B.病人ID And C.病人ID=B.病人ID And B.主页ID=C.住院次数" & _
        " And B.记录性质=3 And Nvl(B.是否上传,0)=0 And B.发生时间 Between Sysdate-3 and Sysdate-1 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未上传的费用明细", TYPE_毕节)
    检查是否上传明细 = (rsTemp.RecordCount = 0)
    If 检查是否上传明细 = False Then MsgBox "存在未上传的费用明细，请先运行上传下载程序将明细上传至医保中心！", vbInformation, gstrSysName
End Function

Private Function 检查是否下载() As Boolean
    Dim strCurDate As String, strDownDate As String
    Dim rsTemp As New ADODB.Recordset
On Error GoTo ErrH
    If gCominfo_毕节.blnOnLine Then
        检查是否下载 = True
        Exit Function
    End If
    
    '检查当天是否下载
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = " Select 下载日期 From 医疗机构基本情况表 Where 单位代码='" & gCominfo_毕节.strHospitalCode & "'"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount = 1 Then
            strDownDate = Nvl(!下载日期)
        End If
    End With
    If strCurDate <> strDownDate Then
        MsgBox "今天还未从中心下载，请先执行下载程序！[检查是否下载]", vbInformation, gstrSysName
        Exit Function
    End If
    
    检查是否下载 = True
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Private Function 事务_开始() As Boolean
    On Error GoTo errHand
    If gCominfo_毕节.blnOnLine Then
        If Not 获取中心连接 Then Exit Function
        Call gobjCenter.BeginTrans
    End If
    gcnGYBJYB.BeginTrans
    
    事务_开始 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 事务_提交() As Boolean
    On Error GoTo errHand
    If gCominfo_毕节.blnOnLine Then
        事务_提交 = gobjCenter.CommitTrans
        If 事务_提交 Then gcnGYBJYB.CommitTrans
        Call 关闭中心连接
    Else
        gcnGYBJYB.CommitTrans
        事务_提交 = True
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 事务_回滚() As Boolean
    On Error GoTo errHand
    gcnGYBJYB.RollbackTrans
    If gCominfo_毕节.blnOnLine Then Call gobjCenter.RollbackTrans
    
    事务_回滚 = True
    If gCominfo_毕节.blnOnLine Then 关闭中心连接
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ExecuteSQL(ByVal strSQL As String, Optional ByVal bln回滚 As Boolean = True) As Boolean
    '如果bln回滚=TRUE，则说明处于事务控制中，该过程自动回滚事务，减少外层函数的繁杂处理
    If Not gobjCenter.ExecuteSQL(strSQL) Then
        If bln回滚 Then Call 事务_回滚
        Exit Function
    End If
    ExecuteSQL = True
End Function

Public Sub 数据转换_毕节(CardData As String, Optional blnRead As Boolean = True)
    'IC卡内数据构成
    'shbzh           社会保障号         32      18      string
    'xm              姓名               50      10      string
    'dwdm            单位代码           60      15      string
    'xb              性别               75      2       string
    'csrq            出生日期           77      10      string
    'cjgzrq          参加工作日期       87      10      string
    'jyqkdm          就业情况代码       97      1       string
    'yxkh            有效卡号           99      2       int
    'grjbdm          个人级别代码       101     2       string
    'ye              个人帐户余额       120     10      decimal
    'zhjzrq          最后就诊日期       151     10      string
    'yydm            最后就诊医院代码   161     4       string
    'pass            个人ic卡密码       168     8       string
    Dim arrData
    
    If Not blnRead Then
        CardData = IC_Data_毕节.社会保障号 & "||" & IC_Data_毕节.姓名 & "||" & IC_Data_毕节.单位代码 & "||" & _
            IC_Data_毕节.性别 & "||" & IC_Data_毕节.出生日期 & "||" & IC_Data_毕节.参加工作日期 & "||" & _
            IC_Data_毕节.就业情况代码 & "||" & IC_Data_毕节.有效卡号 & "||" & IC_Data_毕节.个人级别代码 & "||" & _
            IC_Data_毕节.个人帐户余额 & "||" & IC_Data_毕节.最后就诊日期 & "||" & IC_Data_毕节.最后就诊医院代码 & "||" & _
            IC_Data_毕节.个人IC卡密码
    Else
        arrData = Split(CardData, "||")
        IC_Data_毕节.社会保障号 = arrData(ic.shbzh)
        IC_Data_毕节.姓名 = arrData(ic.xm)
        IC_Data_毕节.单位代码 = arrData(ic.dwdm)
        IC_Data_毕节.性别 = arrData(ic.xb)
        IC_Data_毕节.出生日期 = arrData(ic.Csrq)
        IC_Data_毕节.参加工作日期 = arrData(ic.cjqzrq)
        IC_Data_毕节.就业情况代码 = arrData(ic.jyqkdm)
        IC_Data_毕节.有效卡号 = arrData(ic.yxkh)
        IC_Data_毕节.个人级别代码 = arrData(ic.grjbdm)
        IC_Data_毕节.个人帐户余额 = arrData(ic.ye)
        IC_Data_毕节.最后就诊日期 = arrData(ic.zhjzrq)
        IC_Data_毕节.最后就诊医院代码 = arrData(ic.yydm)
        IC_Data_毕节.个人IC卡密码 = arrData(ic.pass)
    End If
End Sub

Private Function Get流水号_毕节() As String
    Dim lng就诊登记号 As Long
    Dim rsCheck As New ADODB.Recordset

    '获取当前就诊登记号
    gstrSQL = "Select 就诊登记号_ID.Nextval from dual"
    With rsCheck
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        lng就诊登记号 = .Fields(0).Value
    End With
    Get流水号_毕节 = gCominfo_毕节.strHospitalCode & Format(zlDatabase.Currentdate, "yyyyMMdd") & String(6 - Len(CStr(lng就诊登记号)), "0") & lng就诊登记号
End Function

Private Function CheckCard(ByVal str社会保障号 As String, str冻结原因 As String, str冻结说明 As String, str冻结时间 As String) As Boolean
    '检查该病人的卡状态，如果被冻结，则返回假
    Dim rsTemp As New ADODB.Recordset
    
    '以下数据需要从中间库中获取
    gstrSQL = "Select 住院次数,帐户冻结,冻结原因,有效卡号,当年住院费用,冻结时间,冻结说明 " & _
        " From 个人帐户余额表 Where 社会保障号='" & str社会保障号 & "'"
    If gCominfo_毕节.blnOnLine Then
        If Not gobjCenter.GetRecordset(gstrSQL, rsTemp) Then Exit Function
    Else
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnGYBJYB
    End If
    
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        str冻结原因 = Nvl(!冻结原因)
        str冻结说明 = Nvl(!冻结说明)
        str冻结时间 = Nvl(!冻结时间)
        If Nvl(!帐户冻结, "否") = "是" Then Exit Function
    End With
    CheckCard = True
End Function

Private Function BalanceLack(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查病人预交余额是否足够，不足则提示缴款，单据不保存
    #If gverControl >= 5 Then
        gstrSQL = "Select 预交余额,费用余额 From 病人余额 Where 性质=1 And 类型=2 And 病人ID=[1]"
    #Else
        gstrSQL = "Select 预交余额,费用余额 From 病人余额 Where 性质=1 And 病人ID=[1]"
    #End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取预交余额及费用余额", lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If Nvl(rsTemp!预交余额, 0) < Nvl(rsTemp!费用余额, 0) Then Exit Function
    BalanceLack = True
End Function

Private Sub IC_End(Optional ByVal blnPull As Boolean = False)
    On Error Resume Next
    '在打开IC设备后，如果出错，是否仅仅弹卡，否则就在弹卡后关闭端口
    Call gobjCenter.IC_PullCard
    If blnPull Then Exit Sub
    
    Call gobjCenter.IC_CloseDevice
End Sub

Public Sub 取消就诊_毕节()
    '因未进行门诊结算前，未做任何处理，所以取消时需完成的内容只有弹卡：将病人的卡片弹出
    Call IC_End(True)
End Sub
