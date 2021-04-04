Attribute VB_Name = "mdlEInvoiceAutoCreater"
Option Explicit '本模块用于存放涉及数据库访问的公共函数
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例

Public glngSys As Long
Public glngModul As Long
Public gstrDBUser As String                 '当前数据库用户

Public gstrSysName As String                '系统名称
Public gstrProductName As String            'OEM产品名称

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gblnExecuting As Boolean
Public gfrmMain As Object
Public glngSplitTime As Long '间隔时间，秒
Private mobjPubEInvoice As Object 'zlPublicExpense.clsPubEInvoice

Private Function GetUserInfo() As Boolean
    '功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.编号 = rsTmp!编号
            UserInfo.部门ID = zlCommFun.NVL(rsTmp!部门ID, 0)
            UserInfo.简码 = zlCommFun.NVL(rsTmp!简码)
            UserInfo.姓名 = zlCommFun.NVL(rsTmp!姓名)
            GetUserInfo = True
        End If
    End If
End Function

Public Sub Main()
    Dim objRelogin As Object, strPrivs As String
    
    On Error GoTo ErrHandler
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "中联软件")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    On Error Resume Next
    Set objRelogin = CreateObject("ZLLogin.clsLogin")
    If objRelogin Is Nothing Then
        MsgBox "创建 ZLLogin.clsLogin 对象失败。请检查是否正确注册  ZLLogin 部件！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set gcnOracle = objRelogin.Login(1, CStr(Command()))
    If gcnOracle Is Nothing Then
        Set objRelogin = Nothing
        Exit Sub
    End If
    
    glngSys = 100
    glngModul = 1145
    gstrDBUser = objRelogin.DBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息，请与系统管理员联系，先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strPrivs = GetPrivFunc(glngSys, glngModul)
    If zlStr.IsHavePrivs(strPrivs, "开具电子票据") = False Then MsgBox "你不具备开具电子票据的权限！", vbExclamation, gstrSysName: Exit Sub
    
    frmEInvoiceManager.ShowMe strPrivs
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, gstrSysName
End Sub

Private Function GetPubEInvoiceObject(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long, _
    objPubEInvoice As Object, Optional ByVal byt场合 As Byte = 1, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取电子票据公共接口部件
    '入参:
    '   frmMain：调用的主窗体
    '   lngModule：当前调用模块号
    '   byt场合：1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '   blnDeviceSet：设备设置调用的初始化
    '出参:
    '返回:初始化成功返回true,否则返回False
    '说明:
    '   1.使用本部件前,必须先调用本接口进行初始化
    '   2.初始化接口,在HIS进入模块时调用(例如：进入收费管理界面)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExtend As String
    
    If objPubEInvoice Is Nothing Then
        On Error Resume Next
        Set objPubEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
        If Err <> 0 Then
            strErrMsg_Out = "不存在可用的电子票据接口部件(zlPublicExpense.clsPubEInvoice)，请与系统管理员联系。详细的错误信息为:" & vbCrLf & Err.Description
            Exit Function
        End If
    End If
    If objPubEInvoice Is Nothing Then Exit Function
    
    GetPubEInvoiceObject = objPubEInvoice.zlInitialize(frmMain, byt场合, gcnOracle, lngSys, lngModule, False, strExtend)
End Function

Private Function GetSwapCollectFromBalanceID(ByVal byt场合 As Byte, ByVal lng原结算ID As Long, _
    ByRef cllSwapData_Out As Collection, Optional ByVal bln补结算 As Boolean, _
    Optional ByVal lng冲销ID As Long, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算ID获取交易结算信息
    '入参:
    '    byt场合-1-收费, 2-预交, 3-结帐, 4-挂号;5-就诊卡
    '   lng原结算ID byt场合=2，病人预交记录.ID；其它，结帐ID
    '出参:
    '   cllSwapData_Out-返回结算信息
    '      |-PatiInfo   Key="_PatiInfo"
    '        |-(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类),key(_节点名称)
    '      |-BalanceInfo Key="_BalanceInfo"
    '        |-(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPati As Collection, cllBalanceInfo As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strInsureSql As String
    
    On Error GoTo ErrHandler
    Select Case byt场合
    Case 1, 4
        If bln补结算 Then
            strWhere = " And b.结帐id In(Select 收费结帐ID From 费用补充记录 Where 结算ID=[1])"
        Else
            strWhere = " And b.结帐id = [1]"
        End If
    
        strSQL = _
            " Select Max(a.病人id) As 病人ID, Max(a.主页id) As 主页ID, Max(a.姓名) As 姓名, Max(a.性别) As 性别, Max(a.年龄) As 年龄," & _
            "        f_List2Str(Cast(Collect(a.No) As t_StrList)) As NO, Sum(a.结帐金额) As 结帐金额, Max(a.登记时间) As 收费时间" & _
            " From (Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号, Sum(a.结帐金额) As 结帐金额, Max(b.登记时间) As 登记时间" & _
            "        From 门诊费用记录 A, 门诊费用记录 B" & _
            "        Where Mod(a.记录性质, 10) = Mod(b.记录性质, 10) And a.No = b.No And a.序号 = b.序号" & strWhere & _
            "        Group By a.病人id, a.主页id, a.姓名, a.性别, a.年龄, a.No, a.序号" & _
            "        Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0) A"
        
        strInsureSql = "Select Max(险类) As 险类 From 保险结算记录 Where 性质 = 1 And 记录id = [1]"
        
        strSQL = _
            " Select a.病人id, a.主页id, a.姓名, a.性别, a.年龄, m.门诊号, Nvl(n.住院号, m.住院号) As 住院号," & _
            "           a.No, a.结帐金额, a.收费时间, b.险类" & _
            " From (" & strSQL & ") A, (" & strInsureSql & ") B, 病人信息 M, 病案主页 N" & _
            " Where a.病人id = m.病人id(+) And a.病人id = n.病人id(+) And a.主页id = n.主页id(+) And a.No Is Not Null"
    Case 2
        strSQL = _
            "   Select a.Id, a.No, a.病人id, a.主页id, Sum(A.金额) As 结帐金额, Max(A.预交电子票据) As 是否电子票据, " & _
            "          Max(Nvl(d.姓名, c.姓名)) As 姓名, " & _
            "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
            "          max(M.险类) as 险类,to_char(max(A.收款时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(a.预交类别) as 预交类别" & _
            "   From  病人预交记录 A, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=3  and 记录ID=[1] ) M" & _
            "   Where a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
            "   Group By a.Id, a.No, a.病人id, a.主页id"
    Case 3
        strSQL = _
            "   Select a.Id, a.No, a.病人id, a.主页id, Sum(b.冲预交) As 结帐金额, Max(b.是否电子票据) As 是否电子票据, " & _
            "          Max(decode(nvl(A.病人ID,0),0,A.原因,Nvl(d.姓名, c.姓名))) As 姓名, " & _
            "          Max(Nvl(d.性别, c.性别)) As 性别, Max(Nvl(d.年龄, c.年龄)) As 年龄, Max(Nvl(d.住院号, c.住院号)) As 住院号, Max(c.门诊号) As 门诊号, " & _
            "          max(M.险类) as 险类,to_char(max(A.收费时间),'yyyy-mm-dd hh24:mi:ss') as 收费时间,max(A.结帐类型) as 结帐类型" & _
            "   From 病人结帐记录 A, 病人预交记录 B, 病人信息 C, 病案主页 D,(Select 记录ID, 险类 From 保险结算记录 where 性质=2  and 记录ID=[1] ) M" & _
            "   Where a.id=b.结帐ID and  a.病人id = c.病人id(+) And a.病人id = d.病人id(+) And a.主页id = d.主页id(+) And a.Id=[1]  And A.ID=M.记录ID(+)" & _
            "   Group By a.Id, a.No, a.病人id, a.主页id"
    Case 5
        strSQL = _
            "   Select a.结帐id As ID, b.No, a.病人id, a.主页id, Sum(a.冲预交) As 结帐金额, Max(a.是否电子票据) As 是否电子票据, Max(c.姓名) As 姓名, Max(c.性别) As 性别, " & _
            "          Max(c.年龄) As 年龄, Max(c.住院号) As 住院号, Max(c.门诊号) As 门诊号, 0 As 险类, " & _
            "          To_Char(Max(a.收款时间), 'yyyy-mm-dd hh24:mi:ss') As 收费时间 " & _
            "   From 病人预交记录 A, (Select  结帐id,No From 住院费用记录 Where 结帐id = [1]) B, 病人信息 C  " & _
            "   Where a.结帐id = b.结帐id And a.病人id = c.病人id(+)  And a.结帐id = [1] " & _
            "   Group By a.结帐id, b.No, a.病人id, a.主页id"
    Case Else
        strErrMsg_Out = "传入场合【" & byt场合 & "】参数无效。": Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据结帐ID构建电子票据信息", IIf(byt场合 = 2 And lng冲销ID <> 0, lng冲销ID, lng原结算ID))
    If rsTemp.EOF Then strErrMsg_Out = "无剩余未退费用数据。": Exit Function
    
    '1.创建病人信息(病人ID,主页ID,姓名,性别,年龄,门诊号,住院号,险类)
    Set cllPati = New Collection
    cllPati.Add Val(NVL(rsTemp!病人ID)), "_病人ID"
    cllPati.Add Val(NVL(rsTemp!主页id)), "_主页ID"
    cllPati.Add NVL(rsTemp!姓名), "_姓名"
    cllPati.Add NVL(rsTemp!性别), "_性别"
    cllPati.Add NVL(rsTemp!年龄), "_年龄"
    cllPati.Add NVL(rsTemp!门诊号), "_门诊号"
    cllPati.Add NVL(rsTemp!住院号), "_住院号"
    cllPati.Add Val(NVL(rsTemp!险类)), "_险类"

    '2.创建结算信息:(发票号,结算ID,冲销ID,单据号(多个用逗号),登记时间(yyyy-mm-dd hh24:mi:ss),是否补结算,是否部分退款,操作员编号,操作员姓名,结算金额,领用ID)
    Set cllBalanceInfo = New Collection
    cllBalanceInfo.Add "", "_发票号"
    cllBalanceInfo.Add lng原结算ID, "_结算ID"
    cllBalanceInfo.Add lng冲销ID, "_冲销ID"
    cllBalanceInfo.Add NVL(rsTemp!No), "_单据号"
    cllBalanceInfo.Add Format(NVL(rsTemp!收费时间), "yyyy-mm-dd HH:MM:SS"), "_登记时间"
    If byt场合 = 1 Or byt场合 = 4 Then
        cllBalanceInfo.Add IIf(bln补结算, 1, 0), "_是否补结算"
    Else
        cllBalanceInfo.Add 0, "_是否补结算"
    End If
    cllBalanceInfo.Add 0, "_是否部分退款"
    cllBalanceInfo.Add UserInfo.编号, "_操作员编号"
    cllBalanceInfo.Add UserInfo.姓名, "_操作员姓名"
    cllBalanceInfo.Add Val(NVL(rsTemp!结帐金额)), "_结算金额"
    cllBalanceInfo.Add 0, "_领用ID"
    Select Case byt场合
    Case 2
        cllBalanceInfo.Add Decode(Val(NVL(rsTemp!预交类别)) = 0, 3, Val(NVL(rsTemp!预交类别))), "_结算类型" '预交类别:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add 0, "_合约单位结帐"
    Case 3
        cllBalanceInfo.Add Decode(Val(NVL(rsTemp!结帐类型)) = 0, 3, Val(NVL(rsTemp!结帐类型))), "_结算类型"  '结帐类型:1-门诊;2-住院 ;3-门诊和住院;
        cllBalanceInfo.Add IIf(Val(NVL(rsTemp!病人ID)) = 0, 1, 0), "_合约单位结帐"
    Case Else
        cllBalanceInfo.Add 1, "_结算类型"
        cllBalanceInfo.Add 0, "_合约单位结帐"
    End Select
    
    Set cllSwapData_Out = New Collection
    cllSwapData_Out.Add cllPati, "_PatiInfo"
    cllSwapData_Out.Add cllBalanceInfo, "_BalanceInfo"
    
    GetSwapCollectFromBalanceID = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Private Function GetExseData(ByVal byt业务场合 As Byte, ByVal str收费员 As String, _
    ByVal dt开始时间 As Date, ByVal dt结束时间 As Date, ByRef rsExse As ADODB.Recordset, ByRef strErrMsg_Out As String) As Boolean
    '获取电子票据费用数据
    '入参：
    '   byt业务场合 0-所有，1-收费，2-预交，3-结帐，4-挂号，5-就诊卡
    Dim strSQL As String, strWhere As String, strSqlSub As String
    
    On Error GoTo ErrHandler
    strWhere = " And a.收款时间 Between [1] And [2]"
    If Trim(str收费员) <> "" Then strWhere = strWhere & " And a.操作员姓名=[3]"
    
    '1)预交款
    If byt业务场合 = 0 Or byt业务场合 = 2 Then
        strSQL = _
            " Select 2 As 业务类型, a.Id As 结算ID, a.No, a.金额, a.操作员姓名, a.收款时间," & _
            "           a.病人id, a.主页id, Null As 姓名, Null As 性别, Null As 年龄, a.预交类别, Null As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 病人预交记录 A" & _
            " Where a.记录性质 = 1 And a.记录状态 = 1 And a.预交电子票据 = 1" & strWhere & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.Id And 票种 = 2 And 记录状态 = 1)"
        '余额退款
        strSQL = strSQL & " Union All" & _
            " Select 12 As 业务类型, b.ID As 结算ID, a.No, a.金额, a.操作员姓名, a.收款时间," & _
            "           a.病人id, a.主页id, Null As 姓名, Null As 性别, Null As 年龄, a.预交类别,a.Id As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 病人预交记录 A,病人预交记录 B" & _
            " Where a.记录性质 = 11 And a.记录状态 = 1 And a.预交电子票据 = 1" & strWhere & _
            "       And Exists(Select 1 From 病人预交记录 Where 记录性质 = 1 And 附加标志 = 1 And 结帐id = a.结帐id)" & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.Id And 票种 = 2 And 记录状态 = 1)" & _
            "       And a.No=b.No And b.记录性质=1 And b.记录状态 In(1,3)"
    End If
    '2)就诊卡
    If byt业务场合 = 0 Or byt业务场合 = 5 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select 5 As 业务类型, b.结帐id As 结算ID, b.No, Sum(a.结帐金额) As 金额, b.操作员姓名, b.收款时间," & _
            "           a.病人id, a.主页id, a.姓名, a.性别, a.年龄, Null As 预交类别, Null As 冲销ID, Null As 结帐类型, Null As 补结算" & _
            " From 住院费用记录 A, 住院费用记录 A1," & _
            "      (Select Distinct a.结帐id, a.操作员姓名, a.收款时间, b.No" & _
            "       From 病人预交记录 A, 住院费用记录 B" & _
            "       Where a.结帐id = b.结帐ID And b.记录性质 = 5 And b.记录状态 In(1,3) And a.是否电子票据 = 1" & strWhere & _
            "             And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐Id And 票种 = 5 And 记录状态 = 1)) B" & _
            " Where a.No = a1.No And a.序号 = a1.序号 And a1.结帐id = b.结帐ID And a.记录性质 = 5" & _
            " Group By b.结帐id, b.No, b.操作员姓名, b.收款时间, a.病人id, a.主页id, a.姓名, a.性别, a.年龄" & _
            " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
    End If
    '3)结帐
    If byt业务场合 = 0 Or byt业务场合 = 3 Then
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct 3 As 业务类型, a.结帐id As 结算ID, b.No, b.结帐金额 As 金额, a.操作员姓名, a.收款时间," & _
            "           b.病人id, b.主页id, Null As 姓名, Null As 性别, Null As 年龄, Null As 预交类别, Null As 冲销ID, b.结帐类型, Null As 补结算" & _
            " From 病人预交记录 A, 病人结帐记录 B" & _
            " Where a.结帐id = b.ID And b.记录状态 = 1 And a.是否电子票据 = 1" & strWhere & _
            "       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐Id And 票种 = 3 And 记录状态 = 1)"
    End If
    '4)挂号、收费
    If byt业务场合 = 0 Or byt业务场合 = 1 Or byt业务场合 = 4 Then
        strSqlSub = _
            " Select a.结帐id, a.操作员姓名, a.收款时间, Mid(b.No) As No, 0 As 补结算, Null As 结算ID" & _
            " From 病人预交记录 A, 门诊费用记录 B" & _
            " Where a.结帐id = b.结帐ID And b.记录性质 = [记录性质] And b.记录状态 In(1,3) And a.是否电子票据 = 1" & strWhere & _
            "             And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐id And 票种 = [票种] And 记录状态 = 1)" & _
            "             And Not Exists(Select 1 From 费用补充记录 Where 收费结帐id = a.结帐id And 记录性质 = 1 And Nvl(附加标志,0) = [附加标志])" & _
            " Group By a.结帐id, a.操作员姓名, a.收款时间"
        '补充结算
        strSqlSub = strSqlSub & " Union All " & _
            " Select 结帐id, 操作员姓名, 收款时间, No, 1 As 补结算, 结算ID" & _
            " From (Select Distinct b.收费结帐ID As 结帐id, a.操作员姓名, a.收款时间,b.No As No, b.结算ID," & _
            "                    Row_Number() Over(Partition By b.记录性质, b.No Order By b.登记时间) As 组号" & _
            "            From 病人预交记录 A, 费用补充记录 B" & _
            "            Where a.结帐ID=b.结算ID And b.记录性质 = 1 And Nvl(b.附加标志,0) = [附加标志] And b.记录状态 In(1, 3) And a.是否电子票据 = 1" & strWhere & _
            "                       And Not Exists(Select 1 From 电子票据使用记录 Where 结算id = a.结帐id And 票种 = [票种] And 记录状态 = 1))" & _
            " Where 组号 = 1"
            
        strSqlSub = _
            " Select [业务类型] As 业务类型, Nvl(b.结算ID, b.结帐id) As 结算ID, Min(b.No) As No, Sum(a.结帐金额) As 金额, b.操作员姓名, b.收款时间," & _
            "        a.病人id, a.主页id, a.姓名, a.性别, a.年龄, Null As 预交类别, Null As 冲销ID, Null As 结帐类型, b.补结算" & _
            " From 门诊费用记录 A, 门诊费用记录 A1,(" & strSqlSub & ") B" & _
            " Where a.No = a1.No And a.序号 = a1.序号 And a1.结帐id = b.结帐ID And Mod(a.记录性质,10)=[记录性质]" & _
            " Group By Nvl(b.结算ID, b.结帐id), b.操作员姓名, b.收款时间, a.病人id, a.主页id, b.补结算, a.姓名, a.性别, a.年龄" & _
            " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
        
        If byt业务场合 = 0 Or byt业务场合 = 1 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[业务类型]", 1), "[记录性质]", 1), "[票种]", 1), "[附加标志]", 0)
        End If
        
        If byt业务场合 = 0 Or byt业务场合 = 4 Then
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                Replace(Replace(Replace(Replace(strSqlSub, "[业务类型]", 4), "[记录性质]", 4), "[票种]", 4), "[附加标志]", 1)
        End If
    End If
    
    strSQL = _
        " Select Nvl(n.姓名,Nvl(m.姓名,a.姓名)) As 姓名,Nvl(n.性别,Nvl(m.性别,a.性别)) As 性别,Nvl(n.年龄,Nvl(m.年龄,a.年龄)) As 年龄," & _
        "           m.门诊号 As 门诊号, Nvl(n.住院号,m.住院号) As 住院号, a.业务类型, a.结算id, a.No, a.金额, a.操作员姓名, a.收款时间, " & _
        "           a.病人id, a.主页id, a.预交类别, a.冲销id, a.结帐类型, a.补结算" & _
        " From (" & strSQL & ") A, 病人信息 M, 病案主页 N" & _
        " Where a.病人ID=m.病人ID(+) And a.病人ID=n.病人ID(+) And a.主页ID=n.主页ID(+)" & _
        " Order By 收款时间"
    Set rsExse = zlDatabase.OpenSQLRecord(strSQL, "获取电子票据数据", dt开始时间, dt结束时间, str收费员)
    GetExseData = True
    Exit Function
ErrHandler:
    strErrMsg_Out = Err.Description
End Function

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    '定时器回调函数
    If gblnExecuting Then Exit Sub
    gblnExecuting = True
    Call AutoCreateEInvoice
    gblnExecuting = False
End Sub

Private Function AutoCreateEInvoice() As Boolean
    '自动开具电子票据
    Dim dtBegin  As Date, dtEnd As Date
    Dim rsExse As ADODB.Recordset
    Dim strErrMsg As String, byt场合 As Byte, bytPre场合 As Byte, blnInit As Boolean
    Dim lng结算ID As Long, bln补结算 As Boolean, lng冲销ID As Long
    Dim cllSwapData As Collection
    
    On Error GoTo ErrHandler
    dtEnd = zlDatabase.Currentdate
    dtBegin = DateAdd("n", -1 * glngSplitTime, dtEnd)
    
    zlWritLog glngModul, "获取自动开具电子票据费用数据", "AutoCreateEInvoice", "收费时间：" & Format(dtBegin, "yyyy-MM-dd HH:mm:ss") & "～" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss")
    If GetExseData(0, "", dtBegin, dtEnd, rsExse, strErrMsg) = False Then
        zlWritLog glngModul, "获取自动开具电子票据费用数据失败", "AutoCreateEInvoice", "[出错了]" & strErrMsg
        Exit Function
    End If
    
    If rsExse.EOF Then
        zlWritLog glngModul, "获取自动开具电子票据费用数据完成", "AutoCreateEInvoice", "不存在需要开具电子票据的费用数据。"
        Exit Function
    End If
    
    rsExse.Sort = "业务类型,收款时间"
    Do While Not rsExse.EOF
        byt场合 = Val(NVL(rsExse!业务类型)) 'Array("1-收费", "2-预交", "3-结帐", "4-挂号", "5-就诊卡")
        lng结算ID = Val(NVL(rsExse!结算ID))
        If byt场合 = 1 Or byt场合 = 4 Then
            bln补结算 = Val(NVL(rsExse!补结算)) = 1
        ElseIf byt场合 = 12 Then '余额退款
            lng冲销ID = Val(NVL(rsExse!冲销ID))
        End If
    
        byt场合 = byt场合 Mod 10
        If byt场合 <> bytPre场合 Then
            bytPre场合 = byt场合
            blnInit = True
            zlWritLog glngModul, "创建费用公共部件", "AutoCreateEInvoice", "业务场合=" & byt场合
            If GetPubEInvoiceObject(gfrmMain, glngSys, glngModul, mobjPubEInvoice, byt场合, strErrMsg) = False Then
                blnInit = False
                zlWritLog glngModul, "创建费用公共部件失败", "AutoCreateEInvoice", strErrMsg
            End If
        End If
        
        If blnInit Then
            zlWritLog glngModul, "根据结算ID获取交易结算信息", "AutoCreateEInvoice", "结算ID=" & lng结算ID
            If GetSwapCollectFromBalanceID(byt场合, lng结算ID, cllSwapData, bln补结算, lng冲销ID, strErrMsg) = False Then
                zlWritLog glngModul, "根据结算ID获取交易结算信息失败", "AutoCreateEInvoice", strErrMsg
            Else
                zlWritLog glngModul, "开具电子票据", "AutoCreateEInvoice", "结算ID=" & lng结算ID
                If mobjPubEInvoice.zlOnlyCreateEinvoice(gfrmMain, byt场合, cllSwapData, Nothing, False, strErrMsg) = False Then
                    zlWritLog glngModul, "开具电子票据失败", "AutoCreateEInvoice", strErrMsg
                End If
            End If
        End If
        
        rsExse.MoveNext
    Loop
    AutoCreateEInvoice = True
    Exit Function
ErrHandler:
    zlWritLog glngModul, "自动开具电子票据", "AutoCreateEInvoice", "[出错了]" & Err.Description
End Function

