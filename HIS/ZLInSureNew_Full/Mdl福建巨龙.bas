Attribute VB_Name = "Mdl福建巨龙"
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;99-所有交易增加附加参数(最新版)

'Modified by 朱玉宝 20031218 地区：福州 省市医保要分开，因此将所有TYPE_福建巨龙改为intInsure，注意提交程序

'----------------------------------------------字段常量----------------------------------------------
'以下是私有变量申明
Private glng病人ID As Long                                  '补充入院登记时使用
Private bln补充入院 As Boolean
Private mstrFields As String
Private mstrValues As String
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Private Const mintStyle As Integer = 1                       '1-输出到文本;2-输出到屏幕

Public mgstrPatientInfo As String                           '病人信息串
Public Const mstrPath_福建巨龙 As String = "C:\HIS"
Public Const mstrSearch_福建巨龙 As String = "打印.avi"
Public Const mstrReply_福建巨龙 As String = "Reply.txt"
Public Const mstrRequest_福建巨龙 As String = "Request.txt"
Public Const mstrTemp_福建巨龙 As String = "Temp.txt"

Enum 数据类型
    数值型 = 1
    字符型 = 2
    布尔型 = 3
    日期型 = 4
    时间型 = 5
End Enum
Enum 操作方式
    登录 = 1
    入院 = 2
    挂号 = 3
    收费 = 4
    结帐 = 5
    出院 = 6
    验证 = 7
End Enum
Enum 请求目的
    申请 = 1
    冲销 = 2
    刷卡 = 3
    明细 = 4
    查询 = 5    '专用于登录查询
    预结算 = 6
End Enum

Public mrsIniItems As New ADODB.Recordset                   'ini交换文件中的项目
Private mrsIniSection As New ADODB.Recordset                 'ini交换文件中的节
Private mrsDetail As New ADODB.Recordset
Private curTotalMoney As Currency                           '保存预结算时上传的金额总额

'------------------------------------医保常规函数------------------------------------
Public Function 医保初始化_福建巨龙() As Boolean
    Static gbln初始化 As Boolean
    Dim rsTemp As New ADODB.Recordset
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    
    医保初始化_福建巨龙 = False
    If gbln初始化 Then
        医保初始化_福建巨龙 = True
        Exit Function
    End If
    
    If Not InitStruc Then Exit Function
    
    医保初始化_福建巨龙 = True 'frm等待响应.ShowME(操作方式.登录, 请求目的.申请)
    If 医保初始化_福建巨龙 Then gbln初始化 = True
End Function

'Modified by 朱玉宝 20031218 地区：福州
Public Function 医保设置_福建巨龙(ByVal lng险类 As Long) As Boolean
'功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
'返回：接口配置成功，返回true；否则，返回false
    Dim strConn As String
    
    医保设置_福建巨龙 = FrmSet巨龙.ShowSet(lng险类)
End Function

Public Function 医保终止_福建巨龙() As Boolean
    '医保取消登录
    'If Not frm等待响应.ShowME(操作方式.登录, 请求目的.冲销) Then Exit Function
    
    '清除内部记录集
    Set mrsIniItems = Nothing
    Set mrsIniSection = Nothing
    
    医保终止_福建巨龙 = True
End Function

Public Function 身份标识_福建巨龙(ByVal intinsure As Integer, ByVal 操作方式_IN As Integer, Optional ByVal lng病人ID As Long = 0) As String
    Dim str顺序号 As String, str科室 As String, arrReturn
    身份标识_福建巨龙 = ""
    
    '如果是单纯的在帐户管理中进行验证，验证完即退出（提供给用户修改保险病种的一种途径）
    If 操作方式_IN = 操作方式.验证 Then
        If Not frm等待响应.ShowME(intinsure, 操作方式_IN, 请求目的.刷卡) Then Exit Function
        身份标识_福建巨龙 = "小宝"  '用来强迫刷新
        Exit Function
    End If
    
    '如果传入的病人ID不为空，则表示进行的操作是补充入院登记
    bln补充入院 = (lng病人ID <> 0)
    If bln补充入院 Then glng病人ID = lng病人ID
    If Not frm等待响应.ShowME(intinsure, 操作方式_IN, 请求目的.刷卡) Then Exit Function
    身份标识_福建巨龙 = mgstrPatientInfo
    
    '   因同一病人可能存在挂多个科室的号别，所以，门诊收费刷卡时，会返回一段时间内（参数：一天或三天），
    '这个病人挂号的多个流水号及多个科室名称，以分号隔开，需要操作员选择其中一个，作为本次使用的流水号
    Call Record_Locate(mrsIniItems, "名称,Mzlsh0")
    str顺序号 = Nvl(mrsIniItems!值, "")
    If InStr(1, str顺序号, ";") <> 0 Then
        '存在多个挂号科室及挂号流水号
        Call Record_Locate(mrsIniItems, "名称,Ghksmc")
        str科室 = Nvl(mrsIniItems!值, "")
        arrReturn = Split(frmShowList.ShowME(str顺序号 & "||" & str科室), ";")
        str顺序号 = arrReturn(0)
        str科室 = arrReturn(1)
        Call UpdateData("Mzlsh0", str顺序号)
        Call UpdateData("Ghksmc", str科室)
    Else
        Call Record_Locate(mrsIniItems, "名称,Ghksmc")
        str科室 = Nvl(mrsIniItems!值, "")
    End If
    身份标识_福建巨龙 = 身份标识_福建巨龙 & ";" & str科室
End Function

Public Function 个人余额_福建巨龙(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset

    
    gstrSQL = "select A.帐户余额 from 保险帐户 A where A.病人ID=[1] and A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取个人帐户余额", lng病人ID, intinsure)
    
    If rsTemp.EOF Then
        个人余额_福建巨龙 = 0
    Else
        个人余额_福建巨龙 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If

End Function

Public Function 入院登记_福建巨龙(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    Dim strObj As String, str住院号 As String, blnExist As Boolean
    Dim strLine As TextStream, FileSys As New FileSystemObject
    Dim str顺序号 As String
    '先发出刷卡请求，再发出入院请求，得到应答文件后，允许入院则继续
    
    On Error GoTo errHand
    入院登记_福建巨龙 = False
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.入院, 请求目的.申请, lng病人ID) Then Exit Function
    
    '入院登记
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "入院登记")
    '更新入院流水号
    Call Record_Locate(mrsIniItems, "名称,Zylsh0")
    str顺序号 = Nvl(mrsIniItems!值, "")
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'顺序号','''" & str顺序号 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新入院流水号")
    
    入院登记_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_福建巨龙(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln撤销入院 As Boolean = False) As Boolean
    '先结算后出院:
    '   如果发生费用--医保出院、HIS出院
    '   未发生费用  --医保撤销入院、HIS出院
    '先出院后结算:
    '   如果发生费用--HIS出院
    '   未发生费用  --调医保撤销入院（无费用发生，根本不调用医保的结算接口，所以直接出院）
    
    'bln撤销入院=TRUE:表示由入出院管理处以撤销入院的方式在调用，而本医保不支持HIS方面撤销入院，
    '无论何种方式，在HIS方面都反映为出院，而医保端反映为撤销入院或出院
    On Error GoTo errHand
    出院登记_福建巨龙 = False
    
    If bln撤销入院 Then
        MsgBox "不支持该功能，请为病人办理出院手续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If 存在费用记录(lng病人ID, lng主页ID) Then
        If 发生费用(lng病人ID, lng主页ID) Then
            If 操作模式(intinsure) = 0 Then
                '先结算后出院，说明是办理医保出院手续
                If 存在未结费用(lng病人ID, lng主页ID) Then
                    MsgBox "该病人还存在未结费用，不能出院！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.刷卡) Then Exit Function
                If lng病人ID <> 获取病人ID(intinsure) Then
                    MsgBox "病人信息不符！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.申请, lng病人ID) Then Exit Function
                MsgBox "该病人在医保中心成功办理出院手续！", vbInformation, gstrSysName
            Else
                '先出院后结算，如果存在未结费用，则仅HIS出院；否则先调医保出院，再调HIS出院
                If Not 存在未结费用(lng病人ID, lng主页ID) Then
                    If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.刷卡) Then Exit Function
                    If lng病人ID <> 获取病人ID(intinsure) Then
                        MsgBox "病人信息不符！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.申请, lng病人ID) Then Exit Function
                    MsgBox "该病人在医保中心成功办理出院手续！", vbInformation, gstrSysName
                End If
            End If
        Else
            If Not frm等待响应.ShowME(intinsure, 操作方式.入院, 请求目的.冲销, lng病人ID) Then Exit Function
        End If
    Else
        If Not frm等待响应.ShowME(intinsure, 操作方式.入院, 请求目的.冲销, lng病人ID) Then Exit Function
    End If
    
    '出院登记
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & intinsure & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "福建巨龙")
    
    出院登记_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记撤销_福建巨龙(ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    '先结算后出院:
    '   如果发生费用--医保撤销出院
    '   未发生费用  --调医保入院
    '先出院后结算:
    '   存在未结费用--HIS入院
    '   不存在未结费用
    '       如果发生费用--医保撤销出院、HIS入院
    '       未发生费用  --医保入院
    '   结算时：医保结算、医保出院（无费用时，医保不允许结算）
    On Error GoTo errHand
    Dim lng主页ID As Long
    Dim rs住院次数 As New ADODB.Recordset
    出院登记撤销_福建巨龙 = False
    
    '取得主页ID
    gstrSQL = "Select Nvl(住院次数,0) 主页ID From 病人信息 Where 病人ID=[1]"
    Set rs住院次数 = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    lng主页ID = rs住院次数!主页ID
    
    If 存在费用记录(lng病人ID, lng主页ID) Then
        If 发生费用(lng病人ID, lng主页ID) Then
            If 操作模式(intinsure) = 0 Then    '先结算后出院
                If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.冲销, lng病人ID) Then Exit Function
            Else
                If Not 存在未结费用(lng病人ID, lng主页ID) Then
                    If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.冲销, lng病人ID) Then
                        '由于分辨不出是否该调用医保的撤销出院接口，调用未成功时，提示操作员是否仅办理HIS的撤销出院
                        If MsgBox("调用医保的撤销出院接口时发生错误，可能由于该病人未在医保中心办理出院手续！" & vbCrLf & _
                                "你是否忽略该错误，继续办理HIS系统中的撤销出院手续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
            End If
        
            gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & intinsure & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "入院登记")
        Else
            If Not 入院登记_福建巨龙(lng病人ID, intinsure) Then Exit Function
        End If
    Else
        If Not 入院登记_福建巨龙(lng病人ID, intinsure) Then Exit Function
    End If
    
    出院登记撤销_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号结算_福建巨龙(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng病人ID As Long
    Dim cur个人帐户 As Currency, cur现金 As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    挂号结算_福建巨龙 = False
    
    lng病人ID = 获取病人ID(intinsure)
    If Not frm等待响应.ShowME(intinsure, 操作方式.挂号, 请求目的.申请, lng病人ID, lng结帐ID) Then Exit Function
'    If intInsure = TYPE_南平市 Then
'        gstrSQL = "Select A.结算方式,Nvl(A.冲预交,0) 金额 " & _
'                    " From 病人预交记录 A,保险帐户 B " & _
'                    " Where A.病人ID=B.病人ID And B.险类=" & intInsure & " And A.结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
'        '肯定只有现金支付
'        Call OpenRecordset(rsTemp, "读取现金支付额")
'        cur现金 = Nvl(rsTemp!金额, 0)
'
'        cur个人帐户 = 个人余额_福建巨龙(lng病人ID, intInsure)
'        If cur个人帐户 <> 0 Then
'            If cur个人帐户 > cur现金 Then cur个人帐户 = cur现金
'            gstrSQL = " insert into 病人预交记录(ID,记录性质,NO,记录状态,病人ID,主页ID,科室ID,缴款单位," & _
'                     " 单位开户行,单位帐号,摘要,金额,结算方式,结算号码,收款时间,操作员编号,操作员姓名,冲预交,结帐ID) " & _
'                     " select 病人预交记录_ID.nextval ID,记录性质,NO,记录状态,病人ID,主页ID,科室ID, " & _
'                     " 缴款单位,单位开户行,单位帐号,摘要,金额,'个人帐户',结算号码,收款时间,操作员编号, " & _
'                     " 操作员姓名," & cur个人帐户 & ",结帐ID " & _
'                     " from 病人预交记录" & _
'                     " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
'            gcnOracle.Execute gstrSQL
'            '修正现金支付额
'            cur现金 = Val(Format(cur现金 - cur个人帐户, "#####0.00"))
'            If cur现金 <> 0 Then
'                '修改现金支付额
'                gstrSQL = " Update 病人预交记录 Set 冲预交= " & cur现金 & _
'                          " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
'            Else
'                '无现金支付额，删除该预交记录
'                gstrSQL = " Delete 病人预交记录 " & _
'                          " Where 结帐ID=" & lng结帐ID & " And 结算方式='现金' And not (记录性质=1 or 记录性质=11)"
'            End If
'            gcnOracle.Execute gstrSQL
'        End If
'    End If
    If Not 保存挂号结算记录(intinsure, lng病人ID, lng结帐ID) Then Exit Function
'    If intInsure = TYPE_南平市 Then frm结算信息.ShowME (lng结帐ID)
    
    挂号结算_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号结算冲销_福建巨龙(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng病人ID As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    挂号结算冲销_福建巨龙 = False
    
    '此处不使用病人预交记录，是因为考虑到零费用的时候，预交记录中无数据
    gstrSQL = "Select Distinct B.病人ID,B.卡号,B.医保号,B.密码,B.顺序号" & _
        " From 门诊费用记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=[1]" & _
        " And A.结帐ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "医保接口", intinsure, lng结帐ID)
    lng病人ID = rsTmp!病人ID
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.挂号, 请求目的.冲销, lng病人ID, lng结帐ID) Then Exit Function
    If Not 保存挂号结算记录(intinsure, lng病人ID, lng结帐ID, True) Then Exit Function
    
    挂号结算冲销_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 保存挂号结算记录(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal bln冲销 As Boolean = False) As Boolean
    Dim curGhzfy As Currency, strGhlsh As String                            '挂号总费用,挂号流水号
    Dim curZhzfe As Currency, curJjzfe As Currency '帐户支付额，基金支付额
    Dim rsTemp As New ADODB.Recordset
    
    Dim strAdvance As String
    On Error GoTo errHand
    保存挂号结算记录 = False
    
    '获取挂号总费用，挂号流水号
    Call Record_Locate(mrsIniItems, "名称,Ghfy00")
    curGhzfy = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Ghlsh0")
    strGhlsh = Nvl(mrsIniItems!值, 0)
    Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
    curZhzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
    curJjzfe = Val(Nvl(mrsIniItems!值, 0))
    
    '取冲销ID
    If bln冲销 Then
        gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                  " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID", lng结帐ID)
        lng结帐ID = rsTemp("结帐ID")
    Else
        '校正结果
        strAdvance = "个人帐户|" & curZhzfe & "||医保基金|" & curJjzfe
        gstrSQL = " zl_病人结算记录_Update(" & lng结帐ID & ",'" & strAdvance & "')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        curGhzfy & "," & curGhzfy - curZhzfe - curJjzfe & "," & 0 & "," & curJjzfe & "," & curJjzfe & ",0," & _
        0 & "," & curZhzfe & ",'" & strGhlsh & "')"
'        .年度 & "," & .帐户累计增加 & "," & .帐户累计支出 & "," & .累计进入统筹 & "," & _
'        .累计统筹报销 & "," & IIf(.主页ID = 0, "NULL", .主页ID) & "," & .起付线 & "," & .封顶线 & "," & .实际起付线 & "," & _
'        .存在费用记录金额 & "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & _
'        .超限自付金额 & "," & cur个人帐户 & ",'')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存挂号数据")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
    
    保存挂号结算记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_福建巨龙(ByVal rs明细 As ADODB.Recordset, str结算方式 As String, ByVal intinsure As Integer) As Boolean
    Dim curMoney As Currency
    '发出预结算请求
    On Error GoTo errHand
    门诊虚拟结算_福建巨龙 = False
    
    Set mrsDetail = rs明细.Clone
    If Not frm等待响应.ShowME(intinsure, 操作方式.收费, 请求目的.预结算, 获取病人ID(intinsure)) Then Exit Function
    
    Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
    curMoney = Val(Nvl(mrsIniItems!值, 0))
    str结算方式 = "个人帐户;" & curMoney & ";0"
    Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
    curMoney = Val(Nvl(mrsIniItems!值, 0))
    str结算方式 = str结算方式 & "|医保基金;" & curMoney & ";0"
    
    门诊虚拟结算_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_福建巨龙(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String) As Boolean
    Dim curMoney As Currency
    Dim str结算方式 As String
    Dim rsTemp As New ADODB.Recordset
    '发出结算请求
    On Error GoTo errHand
    门诊结算_福建巨龙 = False
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.收费, 请求目的.申请, 获取病人ID(intinsure), lng结帐ID) Then Exit Function
    
    'Modified by 朱玉宝 20031218 地区：福州
    '如果是省市医保，由于无虚拟结算，需要将结算信息保存
    If intinsure <> TYPE_福建巨龙 Then
        Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
        curMoney = Val(Nvl(mrsIniItems!值, 0))
        If curMoney <> 0 Then
            str结算方式 = str结算方式 & "||个人帐户|" & curMoney
        End If
        Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
        curMoney = Val(Nvl(mrsIniItems!值, 0))
        If curMoney <> 0 Then
            str结算方式 = str结算方式 & "||医保基金|" & curMoney
        End If
        
        '如果存在
        If str结算方式 <> "" Then
            str结算方式 = Mid(str结算方式, 3)
            #If gverControl < 2 Then
                gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
            #Else
                strAdvance = str结算方式
                gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
        End If
   End If
    
    If Not 保存门诊收费结算记录(intinsure, lng病人ID, lng结帐ID) Then Exit Function
    
    'Modified by 朱玉宝 20031218 地区：福州
    '如果是省市医保，由于无虚拟结算，需要将结算信息显示出来
    If intinsure <> TYPE_福建巨龙 Then
        #If gverControl < 2 Then
            frm结算信息.ShowME (lng结帐ID)
        #End If
    End If
    
    门诊结算_福建巨龙 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 门诊结算冲销_福建巨龙(ByVal lng结帐ID As Long, ByVal lng病人ID As Long, ByVal intinsure As Integer) As Boolean
    On Error GoTo errHand
    门诊结算冲销_福建巨龙 = False
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.收费, 请求目的.冲销, lng病人ID, lng结帐ID) Then Exit Function
    If Not 保存门诊收费结算记录(intinsure, lng病人ID, lng结帐ID, True) Then Exit Function
    
    门诊结算冲销_福建巨龙 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 保存门诊收费结算记录(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal bln冲销 As Boolean = False) As Boolean
    Dim curMzzfy As Currency, curMzzhzfe As Currency, curMzjjzfe As Currency, curMzgrzfe As Currency, curDbgrzf As Currency, strMzlsh As String     '门诊总费用,帐户支付,基金支付,个人自付,大病个人自付,单据流水号
    Dim rsTemp As New ADODB.Recordset
    Dim blnOld As Boolean
    
    On Error GoTo errHand
    保存门诊收费结算记录 = False
    #If gverControl < 2 Then
        blnOld = True
    #End If
    
    '获取挂号总费用，挂号流水号
    Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
    curMzzhzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Grzfe0")
    curMzgrzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
    curMzjjzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Bcbxf0")
    curMzzfy = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Dbgrzf")
    curDbgrzf = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Djlsh0")
    strMzlsh = Nvl(mrsIniItems!值, 0)
    
    '取冲销ID
    If bln冲销 Then
        gstrSQL = "select distinct A.结帐ID from 门诊费用记录 A,门诊费用记录 B " & _
                  " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID", lng结帐ID)
        lng结帐ID = rsTemp("结帐ID")
    End If
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & "NULL" & "," & 0 & "," & 0 & "," & 0 & "," & _
        curMzzfy & "," & curMzgrzfe + curDbgrzf & "," & curDbgrzf & "," & curMzjjzfe & "," & curMzjjzfe & ",0," & _
        0 & "," & curMzzhzfe & ",'" & strMzlsh & "',NULL,NULL,NULL" & IIf(blnOld, "", IIf(intinsure <> TYPE_福建巨龙, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊收费数据")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
    
    保存门诊收费结算记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_福建巨龙(ByVal lng病人ID As Long, ByVal intinsure As Integer) As String
    Dim curMoney As Currency, str结算方式 As String, lngPatient As Long
    '发出结算请求
    On Error GoTo errHand
    住院虚拟结算_福建巨龙 = ""
    
    'Modified by 朱玉宝 20031218 地区：福州 省市医保不支持预结算
    If Not (intinsure = TYPE_福建巨龙) Then
        住院虚拟结算_福建巨龙 = "个人帐户;0;0"
        Exit Function
    End If
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.结帐, 请求目的.刷卡) Then Exit Function
    lngPatient = 获取病人ID(intinsure)
    If lngPatient <> lng病人ID Then
        MsgBox "病人信息不符！", vbInformation, gstrSysName
        Exit Function
    End If

    '先出院后结算（给操作员的感觉就像是刷卡后提示）
    If 操作模式(intinsure) = 1 Then
        If Not 医保病人已经出院(lngPatient) Then
            MsgBox "由于该医保病人在院，本次结算将做为中途结算！", vbInformation, gstrSysName
        Else
            MsgBox "本次结算将做为出院结算（结算完后自动调医保的出院接口）！", vbInformation, gstrSysName
        End If
    End If
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.结帐, 请求目的.预结算, lng病人ID) Then Exit Function
    
    Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
    curMoney = Nvl(mrsIniItems!值, 0)
    str结算方式 = "个人帐户;" & curMoney & ";0"
    Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
    curMoney = Nvl(mrsIniItems!值, 0)
    str结算方式 = str结算方式 & "|医保基金;" & curMoney & ";0"
    
    住院虚拟结算_福建巨龙 = str结算方式
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_福建巨龙(ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal intinsure As Integer, Optional ByRef strAdvance As String) As Boolean
    Dim curMoney As Currency
    Dim lng预交ID As Long
    Dim lngPatient As Long
    Dim bln出院 As Boolean '记录是否调用出院接口
    Dim bln出院成功 As Boolean '记录调用出院接口是否成功
    Dim str结算方式 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    住院结算_福建巨龙 = False
    bln出院 = False: bln出院成功 = False
    
    'Modified by 朱玉宝 20031218 地区：福州 无预结算接口，此处需调用结帐刷卡，为正式结算做准备
    If intinsure <> TYPE_福建巨龙 Then
        If Not frm等待响应.ShowME(intinsure, 操作方式.结帐, 请求目的.刷卡) Then Exit Function
        lngPatient = 获取病人ID(intinsure)
        If lngPatient <> lng病人ID Then
            Err.Raise 9000, gstrSysName, "病人信息不符！"
            Exit Function
        End If
    
        '先出院后结算（给操作员的感觉就像是刷卡后提示）
        If 操作模式(intinsure) = 1 Then
            If Not 医保病人已经出院(lngPatient) Then
                Err.Raise 9000, gstrSysName, "由于该医保病人在院，本次结算将做为中途结算！"
            Else
                Err.Raise 9000, gstrSysName, "本次结算将做为出院结算（结算完后自动调医保的出院接口）！"
            End If
        End If
    End If
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.结帐, 请求目的.申请, lng病人ID, lng结帐ID) Then Exit Function
    
    'Modified by 朱玉宝 20031218 地区：福州
    '如果是省市医保，由于无虚拟结算，需要将结算信息保存
    If intinsure <> TYPE_福建巨龙 Then
        Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
        curMoney = Val(Nvl(mrsIniItems!值, 0))
        If curMoney <> 0 Then
            str结算方式 = str结算方式 & "||个人帐户|" & curMoney
        End If
        Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
        curMoney = Val(Nvl(mrsIniItems!值, 0))
        If curMoney <> 0 Then
            str结算方式 = str结算方式 & "||医保基金|" & curMoney
        End If
        
        '如果存在
        If str结算方式 <> "" Then
            str结算方式 = Mid(str结算方式, 3)
            #If gverControl < 2 Then
                gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
            #Else
                strAdvance = str结算方式
                gstrSQL = "zl_医保核对表_Insert(" & lng结帐ID & ",'" & str结算方式 & "')"
            #End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
        End If
    End If
    
    If Not 保存住院结算记录(intinsure, lng病人ID, lng结帐ID) Then Exit Function
    住院结算_福建巨龙 = True
    
    'Modified by 朱玉宝 20031218 地区：福州
    '如果是省市医保，由于无虚拟结算，需要将结算信息显示出来
    If intinsure <> TYPE_福建巨龙 Then
        #If gverControl < 2 Then
            frm结算信息.ShowME (lng结帐ID)
        #End If
    End If
    
    '先出院后结算
    If 操作模式(intinsure) = 1 Then
        If 医保病人已经出院(lng病人ID) Then
            If Not frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.刷卡) Then Exit Function
            If lng病人ID <> 获取病人ID(intinsure) Then
                Err.Raise 9000, gstrSysName, "病人信息不符！"
                Exit Function
            End If
            
            bln出院 = True '调用出院接口
            bln出院成功 = frm等待响应.ShowME(intinsure, 操作方式.出院, 请求目的.申请, lng病人ID)
        End If
        
        '显示调用接口给操作员，以便了解当前操作的医保病人是否正常办理出院手续
        If bln出院 Then
            If Not bln出院成功 Then
                Err.Raise 9000, gstrSysName, "出院结算调用失败，请到保险帐户中补办出院手续"
            Else
                Err.Raise 9000, gstrSysName, "该病人在医保中心成功办理出院手续！"
            End If
        End If
    End If
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Public Function 住院结算冲销_福建巨龙(ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    Dim lng病人ID As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    住院结算冲销_福建巨龙 = False
    
    gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码,B.顺序号" & _
        " From 病人结帐记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=[2]" & _
        " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "医保接口", lng结帐ID, intinsure)
    lng病人ID = rsTmp!病人ID
    
    If Not frm等待响应.ShowME(intinsure, 操作方式.结帐, 请求目的.冲销, lng病人ID, lng结帐ID) Then Exit Function
    If Not 保存住院结算记录(intinsure, lng病人ID, lng结帐ID, True) Then Exit Function
    
    住院结算冲销_福建巨龙 = True
    Exit Function
errHand:
    ErrMsgBox Err.Description, IIf(Err.Number > 9000, Err.Number - 9000, vbInformation), Err.Source
    Err.Clear
End Function

Private Function 保存住院结算记录(ByVal intinsure As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, Optional ByVal bln冲销 As Boolean = False) As Boolean
    Dim lng主页ID As Long
    Dim curJszfy As Currency, curJszhzfe As Currency, curJsjjzfe As Currency, curJsgrzfe As Currency, curDbgrzf As Currency, strJslsh As String     '门诊总费用,帐户支付,基金支付,个人自付,大病个人自付,单据流水号
    Dim rsTemp As New ADODB.Recordset
    Dim blnOld As Boolean
    
    On Error GoTo errHand
    保存住院结算记录 = False
    #If gverControl < 2 Then
        blnOld = True
    #End If
    
    '获取挂号总费用，挂号流水号
    Call Record_Locate(mrsIniItems, "名称,Zhzfe0")
    curJszhzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Grzfe0")
    curJsgrzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Jjzfe0")
    curJsjjzfe = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Bcbxf0")
    curJszfy = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Dbgrzf")
    curDbgrzf = Val(Nvl(mrsIniItems!值, 0))
    Call Record_Locate(mrsIniItems, "名称,Djlsh0")
    strJslsh = Nvl(mrsIniItems!值, 0)
    
    '取冲销ID
    If bln冲销 Then
        gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
                  " where A.NO=B.NO and  A.记录状态=2 and B.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取冲销ID", lng结帐ID)
        lng结帐ID = rsTemp("ID") '冲销单据的ID
    End If
    
    gstrSQL = "Select 住院次数 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
    lng主页ID = rsTemp!住院次数
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & intinsure & "," & lng病人ID & "," & _
        Format(zlDatabase.Currentdate, "YYYY") & "," & 0 & "," & 0 & "," & 0 & "," & _
        0 & "," & lng主页ID & "," & 0 & "," & 0 & "," & 0 & "," & _
        curJszfy & "," & curJsgrzfe + curDbgrzf & "," & curDbgrzf & "," & curJsjjzfe & "," & curJsjjzfe & ",0," & _
        0 & "," & curJszhzfe & ",'" & strJslsh & "',NULL,NULL,NULL" & IIf(blnOld, "", IIf(intinsure <> TYPE_福建巨龙, ",1", "")) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存住院结算数据")
    
    gstrSQL = "zl_病人结帐记录_上传(" & lng结帐ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新上传标志")
    
    保存住院结算记录 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 存在费用记录(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rs费用 As New ADODB.Recordset
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(count(病人ID),0) as 金额 " & _
             " From 住院费用记录 " & _
             " Where 病人ID=[1] and 主页ID=[2]" & _
             " And Nvl(记录状态,0)<>0"
    Set rs费用 = zlDatabase.OpenSQLRecord(gstrSQL, "是否住院期间记过帐", lng病人ID, lng主页ID)
    If rs费用.EOF = True Then
        存在费用记录 = False
    Else
        存在费用记录 = (rs费用("金额") <> 0)
    End If
End Function
'
'Public Function 存在未结费用(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'    Dim rs费用 As New ADODB.Recordset
'    '检查该次住院是否还有费用未结算
'    gstrSQL = "Select nvl(金额,0) as 金额  from 病人未结费用 where 病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID
'    Call OpenRecordset(rs费用, "是否存在未结费用")
'    If rs费用.EOF = True Then
'        存在未结费用 = False
'    Else
'        存在未结费用 = (rs费用("金额") <> 0)
'    End If
'End Function

Private Function 发生费用(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    Dim rs费用 As New ADODB.Recordset
    '检查该次住院是否还有费用未结算
    gstrSQL = "Select Sum(nvl(应收金额,0)) as 金额 " & _
              "From 住院费用记录 " & _
              " Where 病人ID=[1] and 主页ID=[2] And Nvl(记录状态,0)<>0"
    Set rs费用 = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在未结费用", lng病人ID, lng主页ID)
    If rs费用.EOF = True Then
        发生费用 = False
    Else
        发生费用 = (rs费用("金额") <> 0)
    End If
End Function

Private Sub 更新相关信息(ByVal lng病人ID As Long, ByVal 操作方式_IN As Integer, ByVal 请求目的_IN As Integer, ByVal intinsure As Integer)
    Dim strSQL  As String, strValue As String
    Dim rsInfo As New ADODB.Recordset
    Dim cur帐户余额 As Currency, int住院次数 As Integer, str工作状态 As String, cur年度医保费用累计 As Currency
    On Error GoTo errHand
    
    '取出各项原来的值
    gstrSQL = " Select 单位编码 As 工作状态,人员身份 As 住院次数,退休证号 As 年度医保费用累计,帐户余额 From 保险帐户" & _
              " Where 险类=[1] And 病人ID=[2]"
    Set rsInfo = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人相关信息", intinsure, lng病人ID)
    cur帐户余额 = IIf(IsNull(rsInfo.Fields("帐户余额").Value), 0, rsInfo.Fields("帐户余额").Value)
    int住院次数 = IIf(IsNull(rsInfo.Fields("住院次数").Value), 0, rsInfo.Fields("住院次数").Value)
    str工作状态 = IIf(IsNull(rsInfo.Fields("工作状态").Value), "", rsInfo.Fields("工作状态").Value)
    cur年度医保费用累计 = IIf(IsNull(rsInfo.Fields("年度医保费用累计").Value), 0, rsInfo.Fields("年度医保费用累计").Value)
    strSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure
    
    '个人帐户余额
    Call Record_Locate(mrsIniItems, "名称,Grzhye")
    strValue = Val(Nvl(mrsIniItems!值, 0))
    cur帐户余额 = Val(strValue)
    If cur帐户余额 < 0 Then cur帐户余额 = 0
    gstrSQL = strSQL & ",'帐户余额'," & cur帐户余额 & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '住院次数
    Call Record_Locate(mrsIniItems, "名称,Bckbcs")
    strValue = Val(Nvl(mrsIniItems!值, 0))
    int住院次数 = IIf(int住院次数 < Val(strValue), Val(strValue), int住院次数)
    gstrSQL = strSQL & ",'人员身份'," & int住院次数 & ")"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '工作状态
    Call Record_Locate(mrsIniItems, "名称,Gzztmc")
    strValue = Nvl(mrsIniItems!值, "")
    str工作状态 = IIf(strValue = "", str工作状态, strValue)
    gstrSQL = strSQL & ",'单位编码','''" & str工作状态 & "''')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '年度医保费用累计
    If 操作方式_IN = 操作方式.入院 And 请求目的_IN = 请求目的.申请 Then
        Call Record_Locate(mrsIniItems, "名称,Ndfylj")
        strValue = Val(Nvl(mrsIniItems!值, 0))
        cur年度医保费用累计 = Val(strValue)
        strSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure
        gstrSQL = strSQL & ",'退休证号'," & cur年度医保费用累计 & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



'------------------------------------其它过程与函数------------------------------------
Public Function SendRequest(ByVal 操作方式_IN As Integer, ByVal 请求目的_IN As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal intinsure As Integer) As Boolean
    Dim objFileSys As New FileSystemObject, objStream As TextStream
    Dim bln明细 As Boolean, curMoney As Currency, cur结帐金额 As Currency
    Dim str收费细目 As String, bln收费 As Boolean
    Dim str收据费目 As String
    Dim rsSecond As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    '发送请求，产生请求文件
    SendRequest = False
    
    '如果请求文件已存在，则先删除后发出请求文件
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrRequest_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrRequest_福建巨龙, True)
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙, True)
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙, True)
    Set objStream = objFileSys.CreateTextFile(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙)
    '先写节头
    With mrsIniSection
        .MoveFirst
        .Find "类型='" & 操作方式_IN & 请求目的_IN & "'"
        If .EOF Then
            MsgBox "该操作未定义，或接口规则发生变化，请与软件提供商联系！", vbInformation, gstrSysName
            GoTo ClearFiles
        End If
        Call OutputData(objStream, Pack(!名称))
    End With
    '填写请求标志
    Call OutputData(objStream, "Request=TRUE")
    If 请求目的_IN = 请求目的.刷卡 Then
        Call Record_Clear(mrsIniItems, True)
        GoTo ReturnCall
    End If
    
    bln明细 = False
    Select Case 请求目的_IN
    Case 请求目的.申请, 请求目的.预结算
        '填写公有信息
        If 操作方式_IN = 操作方式.登录 Then
            Call 获取登录信息(intinsure)
            Call Record_Locate(mrsIniItems, "名称,UserID")
            Call OutputData(objStream, "UserID=" & Nvl(mrsIniItems!值, "supervious"))
            Call Record_Locate(mrsIniItems, "名称,Password")
            Call OutputData(objStream, "Password=" & Nvl(mrsIniItems!值, "yb"))
        Else
            Call OutputData(objStream, "Success=")
            Call OutputData(objStream, "Error=")
            Call Record_Locate(mrsIniItems, "名称,Cardno")
            Call OutputData(objStream, "Cardno=" & Nvl(mrsIniItems!值, ""))
        End If
        
        Select Case 操作方式_IN
        Case 操作方式.入院
            Call 获取入院信息(lng病人ID)
            Call Record_Locate(mrsIniItems, "名称,Ryrq00")
            Call OutputData(objStream, "Ryrq00=" & Nvl(mrsIniItems!值, ""))
            Call Record_Locate(mrsIniItems, "名称,Rysj00")
            Call OutputData(objStream, "Rysj00=" & Nvl(mrsIniItems!值, ""))
            Call Record_Locate(mrsIniItems, "名称,Ryksmc")
            Call OutputData(objStream, "Ryksmc=" & Nvl(mrsIniItems!值, ""))
            Call Record_Locate(mrsIniItems, "名称,Rylb00")
            Call OutputData(objStream, "Rylb00=" & Nvl(mrsIniItems!值, ""))
        Case 操作方式.挂号
            Call 获取挂号信息(lng病人ID, lng结帐ID)
            Call Record_Locate(mrsIniItems, "名称,Ghksmc")
            Call OutputData(objStream, "Ghksmc=" & Nvl(mrsIniItems!值, ""))
            Call Record_Locate(mrsIniItems, "名称,Ghfy00")
            Call OutputData(objStream, "Ghfy00=" & Nvl(mrsIniItems!值, ""))
        Case 操作方式.收费
            Call 获取病种信息(lng病人ID, intinsure)
            Call OutputData(objStream, "Mzlsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
            Call Record_Locate(mrsIniItems, "名称,Bqbm00")
            Call OutputData(objStream, "Bqbm00=" & Nvl(mrsIniItems!值, ""))
            bln明细 = True
        Case 操作方式.结帐
            '取得主页ID
            gstrSQL = "Select Nvl(住院次数,0) 主页ID From 病人信息 Where 病人ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取主页ID", lng病人ID)
            g结算数据.主页ID = rsTmp!主页ID
            Call 获取病种信息(lng病人ID, intinsure)
            Call OutputData(objStream, "Zylsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
            Call Record_Locate(mrsIniItems, "名称,Bqbm00")
            Call OutputData(objStream, "Bqbm00=" & Nvl(mrsIniItems!值, ""))
            bln明细 = True
        Case 操作方式.出院
            Dim rsTemp As New ADODB.Recordset
            Call 获取出院信息(lng病人ID)
            Call OutputData(objStream, "Zylsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
            Call Record_Locate(mrsIniItems, "名称,Cyrq00")
            Call OutputData(objStream, "Cyrq00=" & Nvl(mrsIniItems!值, ""))
            Call Record_Locate(mrsIniItems, "名称,Cysj00")
            Call OutputData(objStream, "Cysj00=" & Nvl(mrsIniItems!值, ""))
            
            'Modified by 朱玉宝 20031218 地区：福州
            '如果是省市医保，需要传递出院状态（治愈，死亡）
            If intinsure = TYPE_福建省 Or intinsure = TYPE_福州市 Then
                gstrSQL = "Select A.出院方式 From 病案主页 A,病人信息 B Where A.病人ID=B.病人ID And A.主页ID=B.住院次数 And A.病人ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取出院状态", lng病人ID)
                Call OutputData(objStream, "cyztlx=" & IIf(rsTemp!出院方式 = "正常", "治愈", rsTemp!出院方式))
            ElseIf intinsure = TYPE_南平市 Then
                gstrSQL = "Select 住院次数 主页ID From 病人信息 Where 病人ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取出院状态", lng病人ID)
                Call OutputData(objStream, "Cyzd00=" & 获取入出院诊断(lng病人ID, rsTemp!主页ID, False, False))
            End If
        End Select
    Case 请求目的.冲销
        '填写公有信息
        If 操作方式_IN <> 操作方式.登录 Then
            Call 获取卡号信息(lng病人ID)
            Call Record_Locate(mrsIniItems, "名称,Cardno")
            Call OutputData(objStream, "Cardno=" & Nvl(mrsIniItems!值, ""))
        End If
        
        Select Case 操作方式_IN
        Case 操作方式.入院
            '要冲销的住院流水号
            Call OutputData(objStream, "Cxlsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
        Case 操作方式.挂号
            '要冲销的挂号流水号
            Call OutputData(objStream, "Ghlsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
        Case 操作方式.收费, 操作方式.结帐
            'Modified by 朱玉宝 20031218 地区：福州 只有福州市医保存在更正收费
'            If 操作方式_IN = 操作方式.收费 And intInsure = TYPE_福州市 Then
'                Call 获取病种信息(lng病人ID)
'                Call Record_Locate(mrsIniItems, "名称,Bqbm00")
'                If Trim(NVL(mrsIniItems!值)) = "" Then
'                    Call OutputData(objStream, "Cxdjh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID))
'                Else
'                    Call OutputData(objStream, "gzdjh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID))
'                End If
'            Else
                Call OutputData(objStream, "Cxdjh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
'            End If
        Case 操作方式.出院
            '要取消出院的住院流水号
            Call OutputData(objStream, "Zylsh0=" & Get流水号(操作方式_IN, 请求目的_IN, lng病人ID, lng结帐ID, intinsure))
        End Select
    End Select
    
    '如果bln明细为真，则提取数据
    If bln明细 Then
        If 操作方式_IN = 操作方式.结帐 Then
            bln收费 = False
            If 请求目的_IN = 请求目的.预结算 Then
                gstrSQL = "Select A.病人ID,A.主页ID,A.婴儿费,C.项目编码 as 医保项目编号,  " & _
                         "  A.保险大类ID,A.收费类别,A.收费细目ID,B.名称 as 医保项目名称,substrb(B.规格,1,20) 规格, " & _
                         "  A.计算单位 单位, sum(A.数量) 数量,sum(A.金额) 金额,'HIS' 医生, " & _
                         " decode(C.是否医保,1,'Y','N') 是否医保,A.收据费目 发票项目  " & _
                         "  From (  " & _
                         "       Select Mod(A.记录性质,10) as 记录性质,A.记录状态,A.NO,Nvl(A.价格父号,序号) as 序号,A.病人ID, " & _
                         "       A.主页ID,Nvl(A.婴儿费,0) as 婴儿费, A.开单人 as 医生,A.开单部门ID,A.收费类别,A.收费细目ID, " & _
                         "       Nvl(A.保险大类ID,0) as 保险大类ID,Avg(Nvl(A.付数,1)*A.数次) as 数量, A.标准单价, " & _
                         "       A.计算单位,A.收据费目,Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 金额,A.发生时间, " & _
                         "       Nvl(A.是否上传,0) as 是否上传,Nvl(A.是否急诊,0) as 是否急诊,A.摘要  " & _
                         "       From 住院费用记录 A,收入项目 B  " & _
                         "       Where A.记帐费用 = 1 And Nvl(A.记录状态,0)<>0 And A.收入项目ID = B.ID And A.病人ID =" & lng病人ID & " And A.主页ID=" & g结算数据.主页ID & _
                         "       Group by Mod(A.记录性质,10),A.记录状态,A.NO,Nvl(A.价格父号,序号),A.病人ID,A.主页ID, " & _
                         "       Nvl(A.婴儿费,0),A.开单人,A.计算单位,A.标准单价,A.收据费目, A.开单部门ID , A.收费类别, A.收费细目ID,  " & _
                         "       NVL(A.保险大类ID, 0), A.发生时间, NVL(A.是否上传, 0), NVL(A.是否急诊, 0), A.摘要  " & _
                         "       Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0 " & _
                         "       ) A, 收费细目 B,收费类别 D,部门表 X, " & _
                         "       (Select * From 保险支付项目 Where 险类=" & intinsure & ") C  " & _
                         "  Where A.收费细目ID = B.ID And B.ID = C.收费细目ID(+) And A.开单部门ID = x.ID And D.编码 = A.收费类别 " & _
                         " Group By A.病人ID,A.主页ID,A.婴儿费,C.项目编码 ,  " & _
                         "  A.保险大类ID,A.收费类别,A.收费细目ID,B.名称 ,substrb(B.规格,1,20) , " & _
                         "  A.计算单位,decode(C.是否医保,1,'Y','N'),A.收据费目" & _
                         " Having Sum(A.数量)<>0"
            Else
                gstrSQL = "Select 'HIS' 医生,A.收费细目ID,C.项目编码 医保项目编号,decode(C.是否医保,1,'Y','N') 是否医保,A.收据费目 发票项目,E.名称 医保项目名称,  " & _
                    " substrb(E.规格,1,20) 规格,A.计算单位 单位,Sum(A.结帐金额) 金额,Sum(A.数次*Nvl(A.付数,1)) 数量  " & _
                    " From 住院费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=" & intinsure & ") C,收费类别 D,收费细目 E  " & _
                    " Where A.执行部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.病人ID=" & lng病人ID & " ANd A.结帐ID=" & lng结帐ID & _
                    " And A.收费类别=D.编码 And A.收费细目ID=E.ID And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And Nvl(A.附加标志,0)<>9 " & _
                    " Group By A.收费细目ID,C.项目编码,C.是否医保,A.收据费目,E.名称,substrb(E.规格,1,20),A.计算单位  " & _
                    " Having Sum(A.数次*Nvl(A.付数,1))<>0"
            End If
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "上传结帐费用")
        Else
            '取收费细目ID
            str收费细目 = ",,"
            bln收费 = True
            
            If 请求目的_IN = 请求目的.申请 Then
                '门诊结算需要重新提取记录集
                gstrSQL = "Select A.开单人,A.收费细目ID,C.项目编码 医保项目编号,decode(C.是否医保,1,'Y','N') 是否医保,A.收据费目,E.名称 医保项目名称,  " & _
                    " substrb(E.规格,1,20) 规格,A.计算单位,A.实收金额,A.标准单价 单价,A.数次*Nvl(A.付数,1) 数量  " & _
                    " From 门诊费用记录 A,部门表 B,(Select * From 保险支付项目 Where 险类=" & intinsure & ") C,收费类别 D,收费细目 E  " & _
                    " Where A.执行部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And A.病人ID=[2] ANd A.结帐ID=[1]" & _
                    " And A.收费类别=D.编码 And A.收费细目ID=E.ID And Nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 And Nvl(A.附加标志,0)<>9"
                Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "获取费用明细", lng结帐ID, lng病人ID)
            End If
            
            With mrsDetail
                Do While Not .EOF
                    If InStr(1, str收费细目, "," & !收费细目ID & ",") = 0 Then
                        str收费细目 = str收费细目 & !收费细目ID & ","
                    End If
                    .MoveNext
                Loop
                str收费细目 = Mid(str收费细目, 3)
                str收费细目 = Mid(str收费细目, 1, Len(str收费细目) - 1)
                If .RecordCount <> 0 Then .MoveFirst
            End With
            
            gstrSQL = "Select D.ID 收费细目ID,D.名称 医保项目名称,substrb(D.规格,1,20) 规格,C.项目编码 医保项目编号" & _
                    " From 收费细目 D,(Select * From 保险支付项目 Where 险类=[1]) C" & _
                    " Where D.ID=C.收费细目ID(+) And D.ID IN ([2])"
            Set rsSecond = zlDatabase.OpenSQLRecord(gstrSQL, "获取费用明细", intinsure, str收费细目)
        End If
        
        With mrsDetail
            'Modified by 朱玉宝 20031218 地区：福州
            '填写处方明细记录数
            If 请求目的_IN = 请求目的.预结算 Or (intinsure <> TYPE_福建巨龙) Then curTotalMoney = 0
            Call OutputData(objStream, "Cfxms0=" & .RecordCount)
            '填写票据项目（如果Cfxms0=0，则不会理会本处内容；如果填写的话，则说明本次是简单收费）
            
            '填写费用明细（一条记录十行）
            Do While Not .EOF
                If .AbsolutePosition = 1 Then
                    With mrsIniSection
                        .MoveFirst
                        .Find "类型='" & 操作方式_IN & 请求目的.明细 & "'"
                        If .EOF Then
                            MsgBox "该操作未定义，或接口规则发生变化，请与软件提供商联系！", vbInformation, gstrSysName
                            GoTo ClearFiles
                        End If
                        '写节头
                        Call OutputData(objStream, Pack(!名称))
                    End With
                End If
                
                'Modified by 朱玉宝 20031218 地区：福州 只有铁路医保需检查项目
                '判断收费细目是否已经审核，如果没有，则提示并退出
'                If intInsure = TYPE_福建巨龙 Then If 检查项目("", !收费细目ID) = False Then GoTo ClearFiles
                
                '以下为十行数据
                If Not bln收费 Then
                    Call OutputData(objStream, Nvl(!医保项目编号, ""))
                    Call OutputData(objStream, !是否医保)
                    str收据费目 = Nvl(!发票项目, "")
                    If intinsure <> TYPE_福建巨龙 Then     '省市医保都是诊察费
                        If str收据费目 = "诊疗费" Then str收据费目 = "诊察费"
                    End If
                    Call OutputData(objStream, str收据费目)
                    Call OutputData(objStream, Nvl(!医保项目名称, ""))
                    Call OutputData(objStream, Nvl(!规格, "无"))
                    Call OutputData(objStream, Nvl(!单位, "无"))
                    curMoney = Nvl(!金额, 0)
                    Call OutputData(objStream, Format(curMoney / !数量, "#####0.0000;-#####0.0000;0;0"))
                    Call OutputData(objStream, Format(!数量, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Format(curMoney, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Nvl(!医生, ""))
                    'Modified by 朱玉宝 20031218 地区：福州
                    If 请求目的_IN = 请求目的.预结算 Or (intinsure <> TYPE_福建巨龙) Then curTotalMoney = curTotalMoney + Nvl(!金额, 0)
                Else
                    rsSecond.MoveFirst
                    rsSecond.Find "收费细目ID=" & !收费细目ID
                    Call OutputData(objStream, Nvl(rsSecond!医保项目编号, ""))
                    If 请求目的_IN = 请求目的.预结算 Then
                        Call OutputData(objStream, IIf(!是否医保 = 1, "Y", "N"))
                    Else
                        Call OutputData(objStream, !是否医保)
                    End If
                    str收据费目 = Nvl(!收据费目, "")
                    If intinsure <> TYPE_福建巨龙 Then
                        If str收据费目 = "诊疗费" Then str收据费目 = "诊察费"
                    End If
                    Call OutputData(objStream, str收据费目)
                    Call OutputData(objStream, Nvl(rsSecond!医保项目名称, ""))
                    Call OutputData(objStream, Nvl(rsSecond!规格, "无"))
                    Call OutputData(objStream, Nvl(!计算单位, "无"))
                    curMoney = Nvl(!实收金额, 0)
                    Call OutputData(objStream, Format(curMoney / !数量, "#####0.0000;-#####0.0000;0;0"))
                    Call OutputData(objStream, Format(!数量, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Format(curMoney, "#####0.00;-#####0.00;0;0"))
                    Call OutputData(objStream, Nvl(!开单人, ""))
                    'Modified by 朱玉宝 20031218 地区：福州
                    If 请求目的_IN = 请求目的.预结算 Or (intinsure <> TYPE_福建巨龙) Then curTotalMoney = curTotalMoney + Nvl(!实收金额, 0)
                End If
                If curMoney < 0 Then
                    MsgBox "医保不支持负数上传，请检查！", vbInformation, gstrSysName
                    GoTo ClearFiles
                End If
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
            
            '如果上传金额与结帐金额不等，则禁止结帐（由于负数冲帐造成的）
            Dim rs结帐 As New ADODB.Recordset
            If lng结帐ID <> 0 Then
                gstrSQL = " Select Sum(结帐金额) 结帐金额 From " & IIf(操作方式_IN = 操作方式.结帐, "住院费用记录", "门诊费用记录") & _
                          " Where 结帐ID=[1] And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
                Set rs结帐 = zlDatabase.OpenSQLRecord(gstrSQL, "本次结帐金额", lng结帐ID)
                cur结帐金额 = Nvl(rs结帐!结帐金额, 0)
                
                If Format(cur结帐金额, "#####.00;-#####.00;0;") <> Format(curTotalMoney, "#####.00;-#####.00;0;") Then
                    MsgBox "上传金额与结帐金额不符，可能是由于负数冲销的单据造成的，请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
ReturnCall:
    objStream.Close
    objFileSys.GetFile(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙).Name = mstrRequest_福建巨龙
    SendRequest = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
ClearFiles:
    On Error Resume Next
    objStream.Close
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrTemp_福建巨龙, True)
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrRequest_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrRequest_福建巨龙, True)
    If objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙) Then Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙, True)
End Function

Private Function Get流水号(ByVal 操作方式_IN As Integer, ByVal 请求目的_IN As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long, ByVal intinsure As Integer) As String
    Dim RSPATIENT As New ADODB.Recordset
    Dim int性质 As Integer
    '挂号请求：无流水号，由应答文件返回挂号流水号
    '挂号冲销：传入挂号流水号，应答文件返回冲销流水号
    '门诊刷卡：返回门诊流水号
    '门诊请求：传入门诊流水号，应答文件返回单据流水号
    '门诊冲销：传入冲销单据号，应答文件返回单据流水号
    '入院请求：无流水号，应答文件返回住院流水号
    '入院冲销：传入住院流水号，应答文件返回冲销流水号
    '结帐请求：传入住院流水号，应答文件返回单据流水号
    '结帐冲销：传入冲销单据号，应答文件返回单据流水号
    '出院请求：传入住院流水号
    '出院冲销：传入住院流水号
    '返回为-1，表示出错
    '保险结算记录的性质含义：1-门诊;2-住院
    
    Get流水号 = ""
    
    '区别对待门诊还是住院
    If 操作方式_IN = 操作方式.结帐 Then
        int性质 = 2
    Else
        int性质 = 1
    End If
    
    '根据操作方式与请求目的，取所需要的流水号
    Select Case 操作方式_IN
    Case 操作方式.入院
        Select Case 请求目的_IN
        'Case 请求目的.申请
        Case 请求目的.冲销
            gstrSQL = " Select 顺序号 as 流水号 From 保险帐户" & _
                     " Where 病人ID=" & lng病人ID & " And 险类=" & intinsure
        End Select
    Case 操作方式.挂号
        Select Case 请求目的_IN
        'Case 请求目的.申请
        Case 请求目的.冲销
            '取原始结帐记录的流水号
            gstrSQL = " Select 支付顺序号 as 流水号 From 保险结算记录 " & _
                      " Where 记录ID=" & lng结帐ID & " And 性质=" & int性质
        End Select
    Case 操作方式.收费
        Select Case 请求目的_IN
        Case 请求目的.申请, 请求目的.预结算
            Call Record_Locate(mrsIniItems, "名称,Mzlsh0")
            Get流水号 = Nvl(mrsIniItems!值, "")
            Exit Function
        Case 请求目的.冲销
            '取原始结帐记录的流水号
            gstrSQL = " Select 支付顺序号 as 流水号 From 保险结算记录 " & _
                      " Where 记录ID=" & lng结帐ID & " And 性质=" & int性质
        'Case 请求目的.刷卡
        End Select
    Case 操作方式.结帐
        Select Case 请求目的_IN
        Case 请求目的.申请, 请求目的.预结算
            gstrSQL = " Select 顺序号 as 流水号 From 保险帐户" & _
                     " Where 病人ID=" & lng病人ID & " And 险类=" & intinsure
        Case 请求目的.冲销
            '取原始结帐记录的流水号
            gstrSQL = " Select 支付顺序号 as 流水号 From 保险结算记录 " & _
                      " Where 记录ID=" & lng结帐ID & " And 性质=" & int性质
        End Select
    Case 操作方式.出院
        gstrSQL = " Select 顺序号 as 流水号 From 保险帐户" & _
                 " Where 病人ID=" & lng病人ID & _
                 " And 险类=" & intinsure
    End Select
    Set RSPATIENT = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人的流水号")
    Get流水号 = RSPATIENT!流水号
End Function

Public Function AnalyseReply(ByVal 操作方式_IN As Integer, ByVal 请求目的_IN As Integer, ByVal intinsure As Integer) As Integer
    Dim objFileSys As New FileSystemObject, objStream As TextStream
    Dim strCompare As String, strLine As String, strSection As String
    Dim strField As String, strValue As String
    Dim strError As String, strIdentify As String, strAddition As String, str顺序号 As String
    Dim lng病人ID As Long, int住院次数 As Integer, lng病种ID As Long
    Dim str卡号 As String, str医保号 As String
    Dim rsTmp As New ADODB.Recordset
    '分析响应文件
    '以下注释的格式为：等号后是该字段对应的中文名称
'    接口应答文件返回时如有参保人信息，都有参保人的各种信息如：姓名、性别、年龄、单位、ic卡状态、工作状态、个人账户余额、地区、分中心等；下面的接口说明中均以"<<参保人其他信息>>"字样代表：
'            xming0=姓名
'            xbie00=性别
'            brnl00=年龄
'            dwmc00=单位名称
'            icztmc=IC卡状态
'            gzztmc=工作状态
'            grzhye=个人帐户余额
'            dqmc00=投保人所属地区名称
'            fzxmc0=投保人所属分中心名称
    '接口应答文件返回时如有处方明细信息，都有收费项目的各种信息如：名称、规格等；下面的接口说明中均以"<<处方明细信息>>"字样代表：
'            医院收费项目在医保中心的编号=医保项目编号
'            是否医保项目=是否医保项目
'            医院收费项目在医保中心的发票项目名称=发票项目
'            医院收费项目在医保中心的名称=医保项目名称
'            医院收费项目在医院的规格=规格
'            医院收费项目在医院的单位=单位
'            医院收费项目在医院的单价=单价
'            医院收费项目的数量=数量
'            医院收费项目的金额=金额
'            医院收费项目的医生姓名=医生姓名
'        此外，接口返回的收费文件的<<处方明细信息>>除有以上信息外，
'        还增加一行信息，为医院收费项目在医保中心的个人自付比例（0 到1）。
'            自付比例=自付比例
    '返回文件中的发票项目均分解到[yb0000]和[fyb000]两个小节中，
    '分别代表按政策医保项目费用和按政策规定个人自付项目费用。

    On Error Resume Next
    AnalyseReply = 0
    
    '清除错误标志，金额等数据
    Call Record_Clear(mrsIniItems, False)
    
    '遇到节名[fpxmbm]、[mzsfmx]、[yb0000]、[fyb000]、[zysfmx]则退出
    strCompare = UCase("[fpxmbm]、[mzsfmx]、[yb0000]、[fybfy0]、[zysfmx]、[tsxm00]、[ybgr00]")
    If Not objFileSys.FileExists(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙) Then Exit Function
    Err = 0
    Set objStream = objFileSys.OpenTextFile(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙, ForReading, False, TristateMixed)
    If Err = 70 Then Err = 0: Exit Function '拒绝的权限，说明医保程序正在对它进行访问
    
    On Error GoTo errHand
    Err = 0
    With objStream
        Do While Not .AtEndOfStream
            strLine = UCase(.ReadLine)
            If InStr(1, strCompare, strLine) <> 0 Then
                Exit Do
            Else
                '如果不是节名，则更新记录集
                If Mid(strLine, 1, 1) = "[" Then
                    '判断是否和操作一致
                    With mrsIniSection
                        .MoveFirst
                        .Find "类型='" & 操作方式_IN & 请求目的_IN & "'"
                        If .EOF Then
                            MsgBox "该操作未定义，或接口规则发生变化，请与软件提供商联系！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        strSection = Pack(!名称)
                    End With
                    If strSection <> strLine Then
                        If .Line = 2 Then
                            'MsgBox "错误的应答文件（例：需要门诊挂号的应答文件，传回来的却是住院登记的应答文件）", vbInformation, gstrSysName
                            Exit Function
                        Else
                            If MsgBox("该节名" & strLine & "未定义，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                Exit Do
                            Else
                                AnalyseReply = 2
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    On Error Resume Next
                    Err = 0
                    strField = UCase(Split(strLine, "=")(0))
                    strValue = Trim(UCase(Split(strLine, "=")(1)))
                    If Err = 0 Then
                        If strField = "REPLY" Then
                            If strValue <> "TRUE" Then Exit Function
                        End If
                        If Not UpdateData(strField, strValue) Then Exit Function
                    End If
                End If
            End If
        Loop
        .Close
    End With
    
    '判断是否出错
    Call Record_Locate(mrsIniItems, "名称,Success")
    If Nvl(mrsIniItems!值, "FALSE") = "FALSE" Then
        Call Record_Locate(mrsIniItems, "名称,Error")
        strError = Nvl(mrsIniItems!值, "")
        MsgBox strError, vbInformation, gstrSysName
        AnalyseReply = 2
        Exit Function
    End If
    If (操作方式_IN = 操作方式.挂号 Or 操作方式_IN = 操作方式.入院) And 请求目的_IN = 请求目的.刷卡 Then
        Call Record_Locate(mrsIniItems, "名称,Valid0")
        If Nvl(mrsIniItems!值, "FALSE") = "FALSE" Then
            Call Record_Locate(mrsIniItems, "名称,Bnghyy")
            If Nvl(mrsIniItems!值, "") <> "" Then strError = Nvl(mrsIniItems!值, "")
            Call Record_Locate(mrsIniItems, "名称,Bndjyy")
            If Nvl(mrsIniItems!值, "") <> "" Then strError = Nvl(mrsIniItems!值, "")
            MsgBox strError, vbInformation, gstrSysName
            AnalyseReply = 2
            Exit Function
        End If
    End If

    
    '如果没有该医保病人的信息，则创建；否则更新病人信息
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    If 请求目的_IN = 请求目的.刷卡 And (操作方式_IN = 操作方式.挂号 Or 操作方式_IN = 操作方式.收费 Or 操作方式_IN = 操作方式.入院) Then
        If lng病人ID <> 0 And lng病人ID <> 获取病人ID(intinsure) Then
            MsgBox "病人信息不符，不能办理！", vbInformation, gstrSysName
            AnalyseReply = 2
            Exit Function
        End If
        
        Call Record_Locate(mrsIniItems, "名称,Cardno")
        str卡号 = Nvl(mrsIniItems!值, "")
        strIdentify = str卡号                              '0卡号
        Call Record_Locate(mrsIniItems, "名称,ID0000")
        str医保号 = Nvl(mrsIniItems!值, "")
        strIdentify = strIdentify & ";" & str医保号          '1医保号
        strIdentify = strIdentify & ";"                                    '2密码
        Call Record_Locate(mrsIniItems, "名称,Xming0")
        strIdentify = strIdentify & ";" & Nvl(mrsIniItems!值, "")     '3姓名
        Call Record_Locate(mrsIniItems, "名称,Xbie00")
        strValue = Nvl(mrsIniItems!值, "")
        'Modified by 朱玉宝 20031218 地区：福州
        If intinsure = TYPE_福建巨龙 Then
            strValue = IIf(strValue = "1", "男", IIf(strValue = "2", "女", ""))
        Else
            strValue = IIf(strValue = "0", "男", IIf(strValue = "1", "女", ""))
        End If
        strIdentify = strIdentify & ";" & strValue '4性别
        Call Record_Locate(mrsIniItems, "名称,Brnl00")
        If Len(str医保号) = 18 Then
            strIdentify = strIdentify & ";" & Mid(str医保号, 7, 4) & "-" & Mid(str医保号, 11, 2) & "-" & Mid(str医保号, 13, 2)  '5出生日期
        Else
            strIdentify = strIdentify & ";19" & Mid(str医保号, 7, 2) & "-" & Mid(str医保号, 9, 2) & "-" & Mid(str医保号, 11, 2) '5出生日期
        End If
        'strIdentify = strIdentify & ";" & DateAdd("YYYY", -1 * Nvl(mrsIniItems!值, 0), zlDatabase.Currentdate)  '5出生日期
        strIdentify = strIdentify & ";"   '6身份证
        Call Record_Locate(mrsIniItems, "名称,Dwmc00")
        strIdentify = strIdentify & ";" & Nvl(mrsIniItems!值, "")  '7.单位名称(编码)
        strAddition = ";"                                  '8.中心代码
        strAddition = strAddition & ";"                             '9.顺序号
        '由于住院次数在门诊时不返回，因此，从数据库中取出住院次数，如果小于等于则依数据库中的数据为准
        int住院次数 = 0: lng病种ID = 0
        If lng病人ID = 0 Then lng病人ID = 获取病人ID(intinsure)
        If lng病人ID <> 0 Then
            gstrSQL = "Select Nvl(人员身份,0) 住院次数 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取住院次数", lng病人ID, intinsure)
            int住院次数 = rsTmp!住院次数
            
            gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 病人ID=[1] And 险类=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病种ID", lng病人ID, intinsure)
            lng病种ID = rsTmp!病种ID
        End If
        Call Record_Locate(mrsIniItems, "名称,Bckbcs")
        If int住院次数 <= Val(Nvl(mrsIniItems!值, 0)) Then int住院次数 = Val(Nvl(mrsIniItems!值, 0))
        strAddition = strAddition & ";" & int住院次数        '10人员身份
        Call Record_Locate(mrsIniItems, "名称,Grzhye")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!值, 0)   '11帐户余额
        strAddition = strAddition & ";0"                            '12当前状态
        Call Record_Locate(mrsIniItems, "名称,Bqbm00")
        strAddition = strAddition & ";" & IIf(lng病种ID = 0, "'NULL'", lng病种ID) '13病种ID
        strAddition = strAddition & ";" & 1 '14在职(1,2,3)
        strAddition = strAddition & ";"     '15退休证号
        Call Record_Locate(mrsIniItems, "名称,Brnl00")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!值, 0) '16年龄段
        strAddition = strAddition & ";"                             '17灰度级
        Call Record_Locate(mrsIniItems, "名称,Grzhye")
        strAddition = strAddition & ";" & Nvl(mrsIniItems!值, 0)      '18帐户增加累计
        strAddition = strAddition & ";0"       '19帐户支出累计
        strAddition = strAddition & ";0;0"       '20进入统筹累计,21统筹报销累计
        strAddition = strAddition & ";" & int住院次数 & ";"       '22住院次数累计
        
        '进行补充入院时，如果保险帐户中存在该病人的信息，且帐户中的病人ID与传入的病人ID不符，提示操作员先合并后再办理补充入院登记
        If bln补充入院 Then
            gstrSQL = "Select 病人ID From 保险帐户 Where 卡号=[1] And 医保号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "读取该病人的保险帐户中的病人ID", str卡号, str医保号)
            If Not rsTmp.EOF Then
                If rsTmp!病人ID <> glng病人ID Then
                    MsgBox "请先将病人身份合并后，再办理补充入院登记！" & vbCrLf & _
                    "当前病人ID[" & glng病人ID & "]；以前的病人ID[" & rsTmp!病人ID & "]", vbInformation, gstrSysName
                    AnalyseReply = 2
                    Exit Function
                End If
            End If
            If lng病人ID = 0 Then lng病人ID = glng病人ID
        Else
            '如果该病人已经存在保险帐户，则不产生病人信息
            gstrSQL = "Select 病人ID From 保险帐户 Where 卡号=[1] And 医保号=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在保险帐户", str卡号, str医保号)
            If Not rsTmp.EOF Then
                lng病人ID = rsTmp!病人ID
            End If
        End If
        lng病人ID = BuildPatiInfo(1, strIdentify & strAddition, lng病人ID, intinsure)
        '返回格式:中间插入病人ID
        If lng病人ID > 0 Then
            mgstrPatientInfo = strIdentify & ";" & lng病人ID & strAddition
        End If
    End If
    
    '更新医保病人帐户余额(挂号结算、门诊结算、住院结算、出院登记)
    Call 更新相关信息(获取病人ID(intinsure), 操作方式_IN, 请求目的_IN, intinsure)
    
    '如果是住院收费刷卡，则更新帐户中的入院顺序号（避免由于其它原因引起的顺序号为空的问题）
    If 请求目的_IN = 请求目的.刷卡 And 操作方式_IN = 操作方式.结帐 Then
        Call Record_Locate(mrsIniItems, "名称,Zylsh0")
        str顺序号 = Nvl(mrsIniItems!值, "")
        gstrSQL = "zl_保险帐户_更新信息(" & 获取病人ID(intinsure) & "," & intinsure & ",'顺序号','''" & str顺序号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新入院流水号")
    End If
    
    '读取完后，删除应答文件
'    If objFileSys.FileExists(mstrPath_福建巨龙 & intInsure & "\" & mstrRequest_福建巨龙) Then
'        Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intInsure & "\" & mstrRequest_福建巨龙, True)
'    End If
    Call objFileSys.DeleteFile(mstrPath_福建巨龙 & intinsure & "\" & mstrReply_福建巨龙, True)
    AnalyseReply = 1
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitStruc() As Boolean
    '初始化记录集
    mstrFields = "名称" & "," & adLongVarChar & "," & "200" & "|" & _
                 "中文名称" & "," & adLongVarChar & "," & "200" & "|" & _
                 "定义" & "," & adLongVarChar & "," & "500" & "|" & _
                 "说明" & "," & adLongVarChar & "," & "2000" & "|" & _
                 "值" & "," & adLongVarChar & "," & "500" & "|" & _
                 "类型" & "," & adDouble & "," & "2" & "|" & _
                 "固定项" & "," & adDouble & "," & "2"
    Call Record_Init(mrsIniItems, mstrFields)
    mstrFields = "名称" & "," & adLongVarChar & "," & "20" & "|" & _
                 "中文名称" & "," & adLongVarChar & "," & "200" & "|" & _
                 "类型" & "," & adLongVarChar & "," & "2"
    Call Record_Init(mrsIniSection, mstrFields)
    '装入初始数据
    InitStruc = Record_Prepare
End Function

Private Function Pack(ByVal strSection As String) As String
    '为节名加上包装 如：[Section]
    Pack = UCase("[" & strSection & "]")
End Function

Private Function UnPack(ByVal strSection As String) As String
    '还原为原始节名
    UnPack = UCase(Mid(strSection, 2, Len(strSection) - 2))
End Function

Private Sub 获取登录信息(ByVal intinsure As Integer)
    Dim rsInfo As New ADODB.Recordset
    
    gstrSQL = "Select * From 保险参数 Where 险类=[1] Order by 序号"
    Set rsInfo = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保参数", intinsure)
    
    Call UpdateData("UserID", Nvl(rsInfo!参数值, "supervisor"))
    rsInfo.MoveNext
    Call UpdateData("Password", Nvl(rsInfo!参数值, "yb"))
    rsInfo.Close
End Sub

Private Sub 获取入院信息(ByVal lng病人ID As Long)
    Dim strValue As String
    Dim rs入院 As New ADODB.Recordset
    
    gstrSQL = "Select to_char(A.入院时间,'yyyy-MM-dd hh24:mi:ss') 入院时间,B.名称 科室 " & _
            " From 病人信息 A,部门表 B " & _
            " Where A.当前科室ID=B.ID And A.病人ID=[1]"
    Set rs入院 = zlDatabase.OpenSQLRecord(gstrSQL, "获取入院信息", lng病人ID)
    
    With rs入院
        strValue = Format(!入院时间, "yyyyMMdd")
        Call UpdateData("Ryrq00", strValue)
        strValue = Format(!入院时间, "HHmm")
        Call UpdateData("Rysj00", strValue)
        Call UpdateData("Ryksmc", IIf(IsNull(!科室), "", !科室))
        Call UpdateData("Rylb00", "普通")
    End With
End Sub

Private Sub 获取挂号信息(ByVal lng病人ID As Long, ByVal lng结帐ID As Long)
    Dim strValue As String
    Dim rs挂号 As New ADODB.Recordset
    
    gstrSQL = "Select Sum(A.实收金额) 费用,B.名称 科室 " & _
            " From 门诊费用记录 A,部门表 B " & _
            " Where A.执行部门ID=B.ID ANd A.病人ID=[1] And 记录性质=4 And 结帐ID=[2]" & _
            " Group by B.名称"
    Set rs挂号 = zlDatabase.OpenSQLRecord(gstrSQL, "获取挂号信息", lng病人ID, lng结帐ID)
    
    With rs挂号
        strValue = Nvl(!科室, "")
        Call UpdateData("Ghksmc", strValue)
        strValue = Format(!费用, "#####0.00;-#####0.00;0;")
        Call UpdateData("Ghfy00", strValue)
    End With
End Sub

Private Function 获取病种信息(ByVal lng病人ID As Long, ByVal intinsure As Integer) As ADODB.Recordset
    Dim strValue As String
    Dim rs收费 As New ADODB.Recordset
    
    gstrSQL = "Select substr(B.名称,1,instr(B.名称,'@@')-1) 病种 From 保险帐户 A,保险病种 B " & _
                    " Where A.险类=B.险类(+) And A.病种ID=B.ID(+) And A.病人ID=[1] And A.险类=[2]"
    Set rs收费 = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人病种信息", lng病人ID, intinsure)
    
    With rs收费
        strValue = Nvl(!病种, "")
        Call UpdateData("Bqbm00", strValue)
    End With
End Function

Private Sub 获取出院信息(ByVal lng病人ID As Long)
    Dim strValue As String
    Dim rs出院 As New ADODB.Recordset
    
    gstrSQL = "Select to_char(出院时间,'yyyy-MM-dd hh24:mi:ss') 出院时间 " & _
            " From 病人信息 " & _
            " Where 病人ID=[1]"
    Set rs出院 = zlDatabase.OpenSQLRecord(gstrSQL, "获取出院信息", lng病人ID)
    
    With rs出院
        strValue = Format(!出院时间, "yyyyMMdd")
        Call UpdateData("Cyrq00", strValue)
        strValue = Format(!出院时间, "HHmm")
        Call UpdateData("Cysj00", strValue)
    End With
End Sub

Private Sub 获取卡号信息(ByVal lng病人ID As Long)
    Dim strValue As String
    Dim rs卡号 As New ADODB.Recordset
    
    gstrSQL = "Select 卡号 " & _
            " From 保险帐户 " & _
            " Where 病人ID=[1]"
    Set rs卡号 = zlDatabase.OpenSQLRecord(gstrSQL, "获取医保病人的卡号", lng病人ID)
    
    With rs卡号
        strValue = Nvl(!卡号, "")
        Call UpdateData("Cardno", strValue)
    End With
End Sub

Public Function 获取病人ID(ByVal intinsure As Integer) As Long
    Dim rsTmp As New ADODB.Recordset
    
    获取病人ID = 0
    Call Record_Locate(mrsIniItems, "名称,ID0000")
    gstrSQL = " Select 病人ID From 保险帐户 Where 医保号=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取该医保病人的基本信息", CStr(mrsIniItems!值), intinsure)
    If Not rsTmp.EOF Then
        获取病人ID = Nvl(rsTmp!病人ID, 0)
    End If
End Function

Private Function 操作模式(ByVal intinsure As Integer) As Long
    Dim intValue As Integer
    Dim rsTmp As New ADODB.Recordset
    
    '获取参数值(0-先结算,后出院;1-先出院,后结算)
    intValue = 0
    gstrSQL = "Select Nvl(参数值,0) Value From 保险参数 Where 险类=[1] And 参数名='操作模式'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取参数值", intinsure)
    
    If Not rsTmp.EOF Then
        intValue = rsTmp!Value
    End If
    操作模式 = intValue
End Function

Private Function 医保病人已经出院(ByVal lng病人ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "Select Nvl(当前状态,0) 状态 From 保险帐户 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "判断医保病人是否出院", lng病人ID)
    
    医保病人已经出院 = (rsTmp!状态 = 0)
End Function

Private Function 检查项目(ByVal strNO As String, ByVal lng收费细目ID As Long) As Boolean
    Dim strCode As String, intVerify As Integer
    Dim rsCheck As New ADODB.Recordset
    检查项目 = False
    
    intVerify = 0
    strCode = ""
    
    '取网点编号
    gstrSQL = "Select Nvl(参数值,'') Value From 保险参数 Where 参数名='服务网点编号'"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取服务网点编号")
    With rsCheck
        If Not .EOF Then
            If Not IsNull(!Value) Then
                strCode = !Value
            End If
        End If
    End With
    If Trim(strCode) = "" Then
        MsgBox "请先设置服务网点编号！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取审核标志
    gstrSQL = "Select nvl(Sfsh00,0) As Value From yydy.Yy_Yydyb0 Where FWWDBH=[1] And YYXMBH=[2]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取审核标志", strCode, lng收费细目ID)
    With rsCheck
        If Not .EOF Then
            If Not IsNull(!Value) Then
                intVerify = !Value
            End If
        End If
    End With
    If intVerify = 0 Then
        gstrSQL = "Select '['||编码||']'||名称 项目 From 收费细目 Where ID=[1]"
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取项目名称", lng收费细目ID)
        If strNO <> "" Then
            MsgBox "项目" & rsCheck!项目 & "还未通过审核（NO：" & strNO & "），不能进行结算操作！", vbInformation, gstrSysName
        Else
            MsgBox "项目" & rsCheck!项目 & "还未通过审核，不能进行结算操作！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    
    检查项目 = True
End Function

Private Function UpdateData(ByVal strColumn As String, ByVal strValue As String) As Boolean
    On Error Resume Next
    
    UpdateData = False
    Call Record_Locate(mrsIniItems, "名称," & strColumn)
    With mrsIniItems
        !值 = strValue
        .Update
    End With
    
    '20030417     刘敏建议取消
'    If Err <> 0 Then
'        MsgBox "应答文件中出现未知的接口项目！", vbInformation, gstrSysName
'        Exit Function
'    End If
    UpdateData = True
End Function







'------------------------------------以下是关于记录集的过程与函数------------------------------------
Private Function Record_Prepare() As Boolean
    '初始化所有内部映射记录集
    On Error Resume Next
    Record_Prepare = False
    
    '-----------------------------mrsIniItems-----------------------------
    mstrFields = "名称|中文名称|定义|说明|值|类型|固定项"
    mstrValues = "UserID|用户名||登录医保数据库的用户名||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Password|密码||用户的密码||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "connected|连接状态||连接状态||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "fwwdmc|医保中心||医保中心||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "czyuan|操作员||操作员||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Request|请求|TRUE or FALSE|各种业务接口请求文件的开始请求标志；TRUE时表示请求文件可以开始被读取||" & 数据类型.布尔型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Reply|响应|TRUE or FALSE|各种业务接口返回文件的回答标志；TRUE时表示应答文件可以开始读取||" & 数据类型.布尔型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Success|成功|TRUE or FALSE|操作成功否||" & 数据类型.布尔型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Error|错误|C400|操作失败原因||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cardno|卡号|C12|医保IC卡号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "ID0000|ID|C19|医疗保险号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Xming0|姓名|C8|姓名||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Xbie00|性别|C1 1男;2女 可能为空|性别||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Brnl00|年龄|N3|年龄||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dwmc00|单位名称|C30|单位名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Icztmc|IC卡状态|C20|IC卡状态名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Gzztmc|工作状态|C30|工作状态名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dqmc00|所属地区|C20|投保人所属地区名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Fzxmc0|所属分中心|C20|投保人所属分中心名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghksmc|挂号科室|C10|挂号科室名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghfy00|挂号费用|N(5,2)|挂号费用||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghlsh0|挂号流水号|C16|挂号流水号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghrq00|挂号日期|C8|挂号日期||" & 数据类型.日期型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ghsj00|挂号时间|C4|挂号时间||" & 数据类型.时间型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cxlsh0|冲销流水号|C16|冲销流水号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Grzhye|个人帐户余额|N(8,2)|个人帐户余额||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bqbm00|病种编码|C20|病种编码||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cfxms0|收费项目数|N(3)|收费项目数||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Djlsh0|单据流水号|C16|单据流水号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Mzlsh0|门诊流水号|C16|门诊流水号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bckbcs|住院次数|N(3)|本次看病次数(同住院次数)||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sftsmz|特殊门诊|C1 Y是;N否|是否特殊门诊||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sftsbz|特殊病种|C1 Y是;N否|是否特殊病种||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zhzfe0|帐户支付额|N(8,2)|帐户支付额||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Grzfe0|个人支付额|N(8,2)|个人现金支付额||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Dbgrzf|大病个人自付|N(8,2)|个人现金支付额||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Jjzfe0|基金支付额|N(8,2)|基金支付额||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bcbxf0|总费用|N(8,2)|总费用||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfrq00|收费日期|C8|收费日期||" & 数据类型.日期型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfsj00|收费时间|C4|收费时间||" & 数据类型.时间型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Sfrxm0|收费操作员|C8|收费人姓名||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cxdjh0|冲销单号|C16|冲销单据号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ryrq00|入院日期|C8|入院日期||" & 数据类型.日期型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rysj00|入院时间|C4|入院时间||" & 数据类型.时间型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ryksmc|入院科室|C10|入院科室名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rylb00|住院类别|C8 普通或家庭病床|住院类别||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ptbcts|普通病床天数|N10 普通病床天数|普通病床天数||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Crbcts|传染病床天数|N10 传染病床天数|传染病床天数||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zylsh0|住院流水号|C16|入院登记流水号||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Rydjr0|入院操作员|C8|入院登记人||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Zyksmc|住院科室|C10|住院科室名称||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cydjr0|出院登记人|C8|出院登记人||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cyrq00|出院日期|C8|出院日期||" & 数据类型.日期型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Cysj00|出院时间|C4|出院时间||" & 数据类型.时间型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Ndfylj|年度医保费用累计|N8,2|年度医保费用累计||" & 数据类型.数值型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Valid0|是否允许操作|True Or False|是否要以入院登记或是否可以挂号||" & 数据类型.布尔型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bnghyy|不能挂号的原因|C400|病人不能挂号的原因||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)
    mstrValues = "Bndjyy|不能入院的原因|C400|病人不能入院登记的原因||" & 数据类型.字符型 & "|1"
    Call Record_Add(mrsIniItems, mstrFields, mstrValues)

    '-----------------------------mrsIniItems-----------------------------
    mstrFields = "名称|中文名称|类型"
    mstrValues = "cydj|出院登记|" & "61"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "cydjcx|出院登记冲销|" & "62"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "cydjsk|出院登记刷卡|" & "63"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "fpxmbm|所有票据项目|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "fybfy0|非医保费用|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "login|登录|" & "11"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "logout|退出|" & "12"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzgh|门诊挂号|" & "31"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzghcx|门诊挂号冲销|" & "32"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzghsk|门诊挂号刷卡|" & "33"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsf|门诊收费|" & "41"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfcx|门诊收费冲销|" & "42"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfmx|门诊收费明细|" & "44"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfsk|门诊收费刷卡|" & "43"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "mzsfyjs|门诊收费预收算|" & "46"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "query|登录查询|" & "15"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydj|住院登记|" & "21"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydjcx|住院登记冲销|" & "22"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "rydjsk|住院登记刷卡|" & "23"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "yb0000|医保费用|" & "70"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysf|住院收费|" & "51"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfcx|住院收费冲销|" & "52"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfmx|住院收费明细|" & "54"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfsk|住院收费刷卡|" & "53"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    mstrValues = "zysfyjs|住院收费预结算|" & "56"
    Call Record_Add(mrsIniSection, mstrFields, mstrValues)
    
    If Err <> 0 Then
        MsgBox "初始化内部数据结构时，发生未知错误！", vbInformation, gstrSysName
        Exit Function
    End If
    Record_Prepare = True
End Function

Private Sub Record_Clear(ByRef rsObj As ADODB.Recordset, Optional ByVal blnAll As Boolean = False)
    '清除记录集的值为空
    
    With rsObj
        If .RecordCount = 0 Then Exit Sub
        Do While Not .EOF
            If blnAll Then
                !值 = ""
                .Update
            Else
                If Nvl(!固定项, 0) = 0 Then
                    !值 = ""
                    .Update
                End If
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, ",")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Private Function Record_Count(ByRef rsObj As ADODB.Recordset, Optional ByVal blnDelete As Boolean = False) As Long
    '返回总记录数
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Count = 0
    
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        If blnDelete = False Then Record_Count = .RecordCount: Exit Function
        .Filter = "删除=0"
        If .RecordCount = 0 Then Exit Function
        Record_Count = .RecordCount
    End With
End Function



'------------------------------------调试用到的过程与函数------------------------------------
'非运行状态调试时，需将相关函数或过程申明为private；调试完成后请还原各函数和过程的申明
Private Sub 调试用主函数(ByVal 操作方式_IN As Integer, ByVal 请求目的_IN As Integer, ByVal intinsure As Integer)
    Call InitStruc
    Call SendRequest(操作方式_IN, 请求目的_IN, 0, 0, intinsure)
End Sub

Private Sub 显示记录集数据()
    Dim intField As Integer, intFields As Integer, strMsg As String
    With mrsIniItems
        intFields = .Fields.Count - 1
        
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            strMsg = ""
            For intField = 0 To intFields
                strMsg = strMsg & "[" & .Fields(intField).Name & "]" & IIf(IsNull(.Fields(intField).Value), "", .Fields(intField).Value)
            Next
            Debug.Print strMsg
            .MoveNext
        Loop
    End With
End Sub

Private Sub OutputData(Optional ByVal objStream As TextStream, Optional ByVal strData As String)
    '输入文本内容
    If mintStyle = 1 Then
        objStream.WriteLine strData
    Else
        Debug.Print strData
    End If
End Sub


