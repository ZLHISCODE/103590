Attribute VB_Name = "mdlFeeCommon"
Option Explicit
Private mrs宁波 As ADODB.Recordset
Public Type Ty_FactProperty
    lngShareUseID As Long   '共享领用批次ID
    strUseType As String ' 使用类别
    intInvoiceFormat As Integer '打印的发票格式,发票格式序号
    intInvoicePrint As Integer     '打印方式:0-不打印;1-自动打印;2-提示打印
End Type
Public grs医疗付款方式 As ADODB.Recordset
Public Type TY_PatiMaxLenInfor
    intPatiName As Integer  '姓名最大长度
    intPatiAge  As Integer   '年龄最大长度
    intPatiSex As Integer   '性别最大长度
    intPatiMzNo As Integer   '门诊号最大长度
End Type
Public grsOneCard As ADODB.Recordset

Private gPatiMaxLen As TY_PatiMaxLenInfor

 Public Function zlGetPatiInforMaxLen() As TY_PatiMaxLenInfor
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息的最大长度
    '返回TY_PatiMaxLenInfor
    '编制:刘兴洪
    '日期:2013-11-11 11:44:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If gPatiMaxLen.intPatiName <> 0 Then
        zlGetPatiInforMaxLen = gPatiMaxLen: Exit Function
    End If
    With gPatiMaxLen
        .intPatiName = 100
        .intPatiMzNo = 18
        .intPatiAge = 20
        .intPatiSex = 4
    End With
    '重数据库中读取
    
    strSQL = "" & _
    "   Select /*+ rule */  A.Column_Name ,Nvl(A.Data_Precision, A.Data_Length) as PatiMaxLen " & _
    "   From All_Tab_Columns A,Table(f_str2list([2])) J " & _
    "   Where A.Table_Name = [1] And A.Column_Name=J.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, "病人信息", "姓名,门诊号,性别,年龄")
    With rsTemp
        Do While Not .EOF
            Select Case nvl(!Column_Name)
            Case "姓名"
                gPatiMaxLen.intPatiName = Val(nvl(rsTemp!PatiMaxLen))
            Case "门诊号"
                gPatiMaxLen.intPatiMzNo = Val(nvl(rsTemp!PatiMaxLen))
            Case "性别"
                gPatiMaxLen.intPatiSex = Val(nvl(rsTemp!PatiMaxLen))
            Case "年龄"
                gPatiMaxLen.intPatiAge = Val(nvl(rsTemp!PatiMaxLen))
            End Select
            .MoveNext
        Loop
    End With
    rsTemp.Close: Set rsTemp = Nothing
    zlGetPatiInforMaxLen = gPatiMaxLen: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetColumnLength(strTable As String, strColumn As String) As Long
    GetColumnLength = Sys.FieldsLength(strTable, strColumn)
End Function
Public Function zlExcuteUploadSwap(ByVal lng病人ID As Long, ByRef strOutPut As String, Optional objExcuteObject As Object = Nothing) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用UploadSwap接口
    '入参:strCardNo
    '     objExcuteObject-调用的对象
    '出参:
    '返回:调用成功,返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-07-24 10:32:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnNothing As Boolean, rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select 编号 From 一卡通目录 where nvl(启用,0)=2 and rownum<=1"
    If mrs宁波 Is Nothing Then
        Set mrs宁波 = zlDatabase.OpenSQLRecord(strSQL, "检查一卡通")
    ElseIf mrs宁波.State <> 1 Then
        Set mrs宁波 = zlDatabase.OpenSQLRecord(strSQL, "检查一卡通")
    End If
    If mrs宁波.EOF Then zlExcuteUploadSwap = True: Exit Function
    
    If objExcuteObject Is Nothing Then
        Set objExcuteObject = CreateObject("zlICCard.clsICCard")
        Set objExcuteObject.gcnOracle = gcnOracle
        blnNothing = True
    End If
    If objExcuteObject Is Nothing Then Exit Function
    'UploadSwap(ByVal strCardNO As String, ByVal lng病人ID As Long, ByRef strOut As String) As Boolean'目前只调,没有什么返回值
    Call objExcuteObject.UploadSwap(lng病人ID, strOutPut)
    If blnNothing Then Set objExcuteObject = Nothing
    
    zlExcuteUploadSwap = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
 

Public Function zl_Check特准项目(ByVal objclsInsure As Object, ByVal intInsure As Integer, ByVal lng病人ID As Long, Optional ByVal bln门诊 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否医保病人是否需要检查特准项目
    '入参:objInsure-创建的医保部件
    '     intInsure-险类
    '     lng病人ID-病人ID
    '     bln门诊-是否门诊
    '出参:
    '返回:如果需要检查特准项目,则返回true,否则返回False
    '编制:刘兴洪
    '问题:24862
    '日期:2009-08-12 10:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zl_Check特准项目 = False
     If bln门诊 Then
        If objclsInsure.GetCapability(support门诊病人不受特准项目限制, lng病人ID, intInsure) = False Then zl_Check特准项目 = True
        Exit Function
     End If
    If objclsInsure.GetCapability(support住院病人不受特准项目限制, lng病人ID, intInsure) = False Then zl_Check特准项目 = True
End Function

Public Sub SetNOInputLimit(ByRef objThis As Object, ByRef KeyAscii As Integer, Optional BytType As Byte)
'功能:处理单据号或票据号输入控件的可输入值,目前单据号允许第一位输字母,后面的输数字,票据号允许前两位是字母或数字,后面的输数字
'参数:objThis:可以是txtbox或可输值的combox
'     bytType:0-单据号,1-票据号
    Dim strAbc As String, str123 As String
    Dim str1 As String, str2 As String
    
    strAbc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    str123 = "0123456789"
    str1 = Mid(objThis.Text, 1, 1): str1 = IIf(str1 = "", "空", str1)
    str2 = Mid(objThis.Text, 2, 1): str2 = IIf(str2 = "", "空", str2)
        
    If BytType = 0 Then
        Call zlControl.TxtCheckKeyPress(objThis, KeyAscii, m文本式)
    Else
        If objThis.Text = "" Or objThis.SelLength = Len(objThis.Text) Or _
            objThis.SelStart = 0 And (objThis.SelLength > 0 Or InStr(strAbc, str1) = 0 Or InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0) Or _
            objThis.SelStart = 1 And (objThis.SelLength > 0 Or InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0) Then
            
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            
            '已输入两个字母并且后面有数字,选中第一个字母时,只能输字母
            '已输一个字母,位置在第一个字母之前时,只能输字母
            If objThis.SelStart = 0 And objThis.SelLength = 1 And InStr(strAbc, str2) > 0 And objThis.SelLength <> Len(objThis.Text) Or _
               objThis.SelStart = 0 And objThis.SelLength = 0 And InStr(strAbc, str1) > 0 And InStr(strAbc, str2) = 0 Then
                If InStr(strAbc & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr(str123 & strAbc & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else
            '已输两个字母,位置在第一个字母之前或两个字母之间时,不允许输入
            If (objThis.SelStart = 0 Or objThis.SelStart = 1) And objThis.SelLength = 0 And InStr(strAbc, str1) > 0 And InStr(strAbc, str2) > 0 Then
                If objThis.SelStart = 1 Then    '允许删除第一个字母
                    If Chr(8) <> Chr(KeyAscii) Then KeyAscii = 0: Beep: Exit Sub
                Else
                    KeyAscii = 0: Beep: Exit Sub
                End If
            Else
                If InStr(str123 & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        End If
    End If
End Sub

Public Function ActualMoney(str费别 As String, ByVal lng收入项目ID As Long, ByVal cur应收金额 As Currency, _
    Optional ByVal lng收费细目ID As Long, Optional ByVal lng库房ID As Long, Optional ByVal dbl数量 As Double, Optional ByVal dbl加班加价率 As Double) As Currency
'功能：根据收费细目ID或收入项目ID(前者优先),应收金额,按费别设置的分段比例打折规则计算实收金额；
'       或对药品按成本加收比例规则计算实收金额
'参数：str费别=病人费别；如果是按动态费别,传入格式为"病人费别,动态费别1,动态费别2,..."
'      lng库房ID,dbl数量,对药品类项目按成本价加收打折时才需要传入
'      dbl数量=包含付数在内的售价数量
'      dbl加班加价率=小数比率,传入的应收金额已按加班加价计算时需要，用于还原及重算
'返回：按打折规则和比例计算的实收金额,如果是动态费别,则"str费别"返回最优惠费别(注意如果未打折计算,可能原样返回,也可能返回第一个)
'说明：
'按成本价加收比例打折的两种计算方法(实际是一种)：
'1.打折金额 = 成本金额 * (1 + 加收比例)
'2.打折金额 = 成本价 * (1 + 加收比例) * 零售数量
'相关的计算公式：
'      成本价 = 药品售价 * (1 - 差价率)
'      成本金额 = 售价金额 * (1 - 差价率) = 成本价 * 零售数量
'      有库存金额时:差价率 = 库存差价 / 库存金额,否则:差价率 = 指导差价率
'      对于分批药品，应每个出库批次分别计算成本价和成本金额
'        对于时价分批，"药品售价=实际金额/实际数量"；分批或时价药品库存不足时，不予打折计算。
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Zl_Actualmoney([1],[2],[3],[4],[5],[6]) as Actualmoney From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str费别, lng收费细目ID, lng收入项目ID, cur应收金额 / (1 + dbl加班加价率), dbl数量, lng库房ID)
        
    str费别 = Split(rsTmp!ActualMoney, ":")(0)
    ActualMoney = Format(Split(rsTmp!ActualMoney, ":")(1) * (1 + dbl加班加价率), gstrDec)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetActualMoney(ByVal str费别 As String, ByVal lng收入ID As Long, ByVal cur应收 As Currency, ByVal lng收费细目ID As Long) As Currency
'功能：根据指定的费别和收入项目或收费项目,计算指定金额的实际收款金额
'参数：
'   str费别   ：费别
'   lng收入ID  ：收入项目ID
'   cur应收：应收金额值
'返回：实际应收的金额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select 实收比率" & vbNewLine & _
            "From 费别明细" & vbNewLine & _
            "Where 费别 = [1] And 收费细目id = [3] And Abs([4]) Between 应收段首值 And 应收段尾值" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select 实收比率" & vbNewLine & _
            "From 费别明细 A" & vbNewLine & _
            "Where 费别 = [1] And 收入项目id = [2] And Abs([4]) Between 应收段首值 And 应收段尾值 And Not Exists" & vbNewLine & _
            " (Select 1 From 费别明细 C Where C.费别 = A.费别 And C.收费细目id = [3])"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str费别, lng收入ID, lng收费细目ID, cur应收)
    If rsTmp.EOF Then
        GetActualMoney = cur应收
    Else
        GetActualMoney = cur应收 * rsTmp!实收比率 / 100
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function ReturnMovedExes(ByVal strNO As String, ByVal BytType As Byte, Optional ByVal strFormCaption As String) As Boolean
'功能:根据用户选择抽选后备数据表中的数据到当前数据表中
'参数:bytType表示单据类型,值::1-收费,2-记帐,3-自动记帐,4-挂号,5-就诊卡,6-预交,7-结帐；
'返回:用户选择取消操作,或者抽选数据转出失败,则返回False

    MsgBox "当前操作的单据" & strNO & "在后备数据表中!" & vbCrLf _
        & "请与系统管理员联系,转入到在线数据表再操作!", vbInformation, gstrSysName
    ReturnMovedExes = False

'以下是抽选返回数据的过程，暂存，便于将来透明访问时重用
    
'    If MsgBox("当前操作单据" & strNO & "在后备数据表中,系统需要先把与此单据相关的数据转入到在线数据表才能继续!" & vbCrLf & _
'                             "确定要进行此操作吗?", vbInformation + vbYesNo, gstrSysName) = vbNo Then
'        ReturnMovedExes = False     '此句可省
'        Exit Function
'    End If
'
'    If zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption) Then
'        ReturnMovedExes = True
'    Else
'        '详细错误在之前的执行过程出错时给出
'        MsgBox "因系统错误,与该单据相关的数据未能转入到在线数据表." & vbCrLf & "操作未成功,请与系统管理员联系!", vbInformation, gstrSysName
'        ReturnMovedExes = False
'    End If
End Function

Public Function OverTime(Curdate As Date) As Boolean
'功能：判断当前是否处于加班时间范围内
'返回：真-当前处于加班时间内,假-不处于
    Dim curTime As Date, DateBegin As Date, DateEnd As Date
    Dim str上午 As String, str下午 As String
    
    curTime = CDate(Format(Curdate, "HH:MM:SS"))
    
    str上午 = zlDatabase.GetPara(1, glngSys)
    If str上午 <> "" Then
        DateBegin = CDate(Trim(Split(UCase(str上午), "AND")(0)))
        DateEnd = CDate(Trim(Split(UCase(str上午), "AND")(1)))
    End If
    
    If Not (curTime >= DateBegin And curTime <= DateEnd) Then
        str下午 = zlDatabase.GetPara(2, glngSys)
        If str下午 <> "" Then
            DateBegin = CDate(Trim(Split(UCase(str下午), "AND")(0)))
            DateEnd = CDate(Trim(Split(UCase(str下午), "AND")(1)))
        End If
        
        If Not (curTime >= DateBegin And curTime <= DateEnd) Then OverTime = True
    End If
End Function

Public Function GetInsureName(intInsure As Integer) As String
'功能：根据保险类别序号获取保险类别名称
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 保险类别 Where 序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intInsure)  '一般情况，参数SQL无意义，当有同时连接多个医保时，有点点作用
    If Not rsTmp.EOF Then
        GetInsureName = "" & rsTmp!名称
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockCheck(ByVal BytType As Byte) As Collection
'功能：获取药品或卫材出库检查的集合
'参数：bytType:0-药品，1-卫材
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim colStock As Collection, i As Long
    
    Set colStock = New Collection
    colStock.Add 0, "_0" '避免出错
    
    strSQL = _
        " Select Distinct A.ID,C.检查方式" & _
        " From 部门表 A,部门性质说明 B," & IIf(BytType = 0, "药品出库检查", "材料出库检查") & " C" & _
        " Where B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 " & IIf(BytType = 0, "IN('中药房','西药房','成药房')", "='发料部门'") & _
        " And C.库房ID(+)=A.ID"
        '26046:站点取消.
        '"   And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlFeeCommon")
    For i = 1 To rsTmp.RecordCount
        colStock.Add nvl(rsTmp!检查方式, 0), "_" & rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Set GetStockCheck = colStock
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set GetStockCheck = colStock
End Function

Public Function Get结算方式(str场合 As String, Optional str性质 As String) As ADODB.Recordset
    Dim strSQL As String, strIF As String
    
    On Error GoTo errH
    
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strIF = "And Instr(','||[2]||',',','||B.性质||',')>0 "
        Else
            strIF = "And B.性质 = [2]"
        End If
    End If
    strSQL = _
        " Select B.编码,B.名称,Nvl(Nvl(A.缺省标志,B.缺省标志),0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 " & _
        " And (B.性质<>7 Or B.性质=7 And Exists(Select 1 From 一卡通目录 C Where C.结算方式=B.名称 And C.启用=1))   " & strIF
    If InStr(1, str性质, ",9") > 0 Then
        strSQL = strSQL & " Union " & _
                 " Select 编码,名称,Nvl(缺省标志,0) As 缺省,Nvl(性质,1) as 性质,Nvl(应付款,0) as 应付款 " & _
                 " From 结算方式 " & _
                 " Where 性质=9 " & _
                 " Order by 性质,编码"
    Else
        strSQL = strSQL & " Order by 性质,lpad(编码,3,' ')"
    End If
    Set Get结算方式 = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str场合, str性质)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get病人医疗付款方式(lng病人ID As Long, Optional lng主页ID As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If lng主页ID = 0 Then
        strSQL = "Select 医疗付款方式 From 病人信息 Where 病人ID=[1]"
    Else
        strSQL = "Select 医疗付款方式 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then Get病人医疗付款方式 = "" & rsTmp!医疗付款方式
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMedPayMode(ByVal strName As String, ByRef rsMedPayMode As ADODB.Recordset) As Byte
'功能：根据医疗付款方式名称返回其编码
    Dim strSQL As String
    
    On Error GoTo errH
    
    If rsMedPayMode Is Nothing Then
        strSQL = "Select 编码,名称,缺省标志 From 医疗付款方式"
        Set rsMedPayMode = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    End If
    rsMedPayMode.Filter = "名称='" & strName & "'"
    If rsMedPayMode.RecordCount > 0 Then GetMedPayMode = Val(rsMedPayMode!编码)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMedPayModeName(ByVal strCode As String) As String
'功能：根据医疗付款方式编码返回其名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 名称 From 医疗付款方式 Where 编码 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strCode)
        
    If rsTmp.RecordCount > 0 Then GetMedPayModeName = rsTmp!名称
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiWarnRange(ByVal lngPatient As Long, ByVal lngPage As Long) As String
'功能：获取病人报警适用范围,用于记帐表中报警
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select Zl_Patiwarnscheme([1], [2]) As 适用病人 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngPatient, lngPage)
        
    If rsTmp.RecordCount > 0 Then GetPatiWarnRange = rsTmp!适用病人
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitWarn(Optional ByVal str适用病人 As String, Optional ByVal str病区ID As String) As ADODB.Recordset
'功能：返回病区记帐报警记录集
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(病区ID,0) 病区ID,适用病人,Nvl(报警方法,1) as 报警方法," & _
            " 报警值,报警标志1,报警标志2,报警标志3" & _
            " From 记帐报警线 Where 1=1" & _
            IIf(str适用病人 = "", "", " And 适用病人 = [1]") & _
            IIf(str病区ID = "", "", " And Nvl(病区ID,0) = [2]")
    Set GetUnitWarn = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str适用病人, str病区ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.部门名称 = "" & rsTmp!部门名
        UserInfo.简码 = "" & rsTmp!简码
        UserInfo.姓名 = "" & rsTmp!姓名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPersonnel(str性质 As String, Optional blnBaseInfo As Boolean) As ADODB.Recordset
'功能：读取指定性质的人员列表
    Dim strSQL As String
    On Error GoTo errH
    
    If str性质 <> "" Then
        If blnBaseInfo Then
            strSQL = "Select a.id,a.编号,a.简码,a.姓名 From 人员表 a,人员性质说明 b" & _
            " Where a.ID = b.人员ID And b.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by a.简码"
        Else
            strSQL = "Select a.Id, a.编号, a.姓名, a.简码, a.身份证号, a.出生日期, a.性别, a.民族, a.工作日期, a.办公室电话, a.电子邮件, a.执业类别, a.执业范围, " & _
                    "a.管理职务, a.专业技术职务, a.聘任技术职务, a.学历, a.所学专业, a.留学时间, a.留学渠道, a.接受培训, a.科研课题, a.个人简介, a.建档时间, " & _
                    "a.撤档时间, a.撤档原因, a.别名, a.站点 From 人员表 a,人员性质说明 b" & _
            " Where a.ID = b.人员ID And b.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by a.简码"
        End If
    Else
        If blnBaseInfo Then
            strSQL = "Select id,编号,简码,姓名 From 人员表 A" & _
            " Where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by 简码"
        Else
            strSQL = zlGetFullFieldsTable("人员表", 0, "", False) & _
            " Where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by 简码"
        End If
    End If
    Set GetPersonnel = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str性质)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPersonnelID(str姓名 As String, Optional ByRef rs人员 As ADODB.Recordset) As Long
'功能：根据人员姓名返回ID
'说明：查看收费单时，开单人(医生)可能现在已不是医生了，在mrs开单人中不存在
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If str姓名 = "" Then Exit Function
    
    If Not rs人员 Is Nothing Then
        rs人员.Filter = "姓名='" & str姓名 & "'"
        If rs人员.RecordCount > 0 Then GetPersonnelID = rs人员!ID: Exit Function
    End If
    
    On Error GoTo errH
    strSQL = "Select ID from 人员表 Where 姓名=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str姓名)
    If Not rsTmp.EOF Then GetPersonnelID = rsTmp!ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDepartments(ByVal str性质 As String, _
    ByVal str服务对象 As String, _
    Optional ByVal bln仅操作员部门 As Boolean = False, _
    Optional ByVal blnCheck站点 As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定性质的部门列表
    '入参:str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     bln仅操作员部门-操作员的所属部门
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSQL = " And B.工作性质 = [1]"
        End If
    End If
    If bln仅操作员部门 Then strSQL = strSQL & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & IIf(bln仅操作员部门, ",部门人员 C", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And Instr(',' || [2]|| ',',',' || B.服务对象 || ',')>0 " & strSQL & _
         IIf(blnCheck站点, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
        " Order by A.编码"
    Set GetDepartments = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str性质, str服务对象, UserInfo.ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'该方法与 GetDepartments 方法类似，该方法取消了站点限制
Public Function GetDepts(ByVal str性质 As String, ByVal str服务对象 As String, Optional ByVal bln仅操作员部门 As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定性质的部门列表
    '入参:str性质='临床','护理','中药房',...,允许为空
    '     str服务对象:以,分离:如1,3
    '     bln仅操作员部门-操作员的所属部门
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-10-12 09:44:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    str性质 = Replace(str性质, "'", "")
    If str性质 <> "" Then
        If InStr(1, str性质, ",") > 0 Then
            strSQL = " And Instr(','||[1]||',',','||B.工作性质||',')>0"
        Else
            strSQL = " And B.工作性质 = [1]"
        End If
    End If
    If bln仅操作员部门 Then strSQL = strSQL & "  And A.id=C.部门ID and C.人员id =[3]"
    
    strSQL = _
        " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & IIf(bln仅操作员部门, ",部门人员 C", "") & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And Instr(',' || [2]|| ',',',' || B.服务对象 || ',')>0 " & strSQL & _
        " Order by A.编码"
    Set GetDepts = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str性质, str服务对象, UserInfo.ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitDept() As ADODB.Recordset
'功能：获取病区科室对应关系
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 病区ID, 科室ID From 病区科室对应"
    Set GetUnitDept = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Public Function GetDeptOrUnit(ByVal BytType As Byte, lngDept As Long, ByVal strServiceRange As String) As ADODB.Recordset
'功能：获取指定病区的科室,或指定科室的病区
'参数：bytType=0-指定病区的科室,1-指定科室的病区
'      strServiceRange=服务对象，1-门诊，2-住院，3－门诊和住院
'       lngDept=科室ID或病区ID
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,病区科室对应 B,部门性质说明 C " & _
            " Where " & IIf(BytType = 0, "B.科室ID=A.ID And B.病区ID", "B.病区ID=A.ID And B.科室ID") & "=[1] " & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And C.部门ID=A.ID And Instr(',' || [2]|| ',',',' || C.服务对象 || ',')>0 " & _
            " And C.工作性质=" & IIf(BytType = 0, "'临床'", "'护理'") & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Set GetDeptOrUnit = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", lngDept, strServiceRange)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function StringDelItem(ByVal strAll As String, ByVal strItem As String, Optional strSplit As String = ",") As String
'功能：从指定的字符串列表中删除一项(如果有多个匹配的,只移除第一个)
    Dim i As Long, arrTmp As Variant
    
    arrTmp = Split(strAll, strSplit)
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) = strItem Then
            strItem = ""
        Else
            StringDelItem = StringDelItem & "," & arrTmp(i)
        End If
    Next
    StringDelItem = Mid(StringDelItem, 2)
End Function

Public Function GetOneCardBalance(ByVal lng结帐ID As Long) As ADODB.Recordset
'功能：获取一卡通结算记录
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select A.单位帐号, A.结算号码, B.医院编码, A.冲预交 as 金额" & vbNewLine & _
            "From 病人预交记录 A, 一卡通目录 B" & vbNewLine & _
            "Where A.结帐id = [1] And A.结算方式 = B.结算方式"

    Set GetOneCardBalance = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng结帐ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetOneCard() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一卡通设置记录集
    '返回:返回一卡通设置记录集
    '编制:刘兴洪
    '日期:2014-07-04 10:17:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    If Not grsOneCard Is Nothing Then
        If grsOneCard.State = 1 Then
            Set GetOneCard = grsOneCard
            Exit Function
        End If
    End If
    strSQL = "Select 编号,名称,医院编码,结算方式 From 一卡通目录 Where 启用=1"
    Set grsOneCard = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Set GetOneCard = grsOneCard
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病区ID(ByVal lng科室ID As Long) As Long
'功能：根据科室ID获取对应的病区ID,
'       如果有多个病区,取ID最小的一个,没有找到时返回0
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Min(病区ID) 病区ID From 病区科室对应 Where 科室ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng科室ID)
    
    If Not rsTmp.EOF Then Get病区ID = Val("" & rsTmp!病区ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnit(ByVal blnLimitUnit As Boolean, _
    ByVal strServiceRange As String, ByVal strType As String, _
    Optional bln简码 As Boolean = False, _
    Optional blnNotNode As Boolean = False, _
    Optional blnShowNodeCode As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有病区或科室列表
    '入参:blnLimitUnit=是否有所有病区权限，没有时，只获取操作员所属的科室或病区
    '       blnNotNode-是否区分站点:true,不区分站点,区分站点
    '       blnShowNodeCode:显示站点编号
    '出参:
    '返回:病区或科室数据集
    '编制:刘兴洪
    '日期:2011-02-28 17:21:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strUnitIDs As String
    Dim strWhere As String
    
    On Error GoTo errH
    If blnLimitUnit Then strUnitIDs = GetUserUnits
    strWhere = ""
    If blnNotNode = False Then strWhere = " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) "
    strSQL = _
         " Select A.ID,A.编码,A.名称 " & IIf(bln简码, ",A.简码", "") & IIf(blnShowNodeCode, ",A.站点", "") & _
         " From 部门表 A,部门性质说明 B" & _
         " Where B.部门ID = A.ID And B.服务对象 IN(" & strServiceRange & ") And B.工作性质 = [2]" & _
         " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            strWhere & vbNewLine & _
         IIf(blnLimitUnit, " And Instr(','||[1]||',',','||A.ID||',')>0", "") & _
         " Order by A.编码"
    Set GetUnit = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strUnitIDs, strType)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get科室IDs(lngUnit As Long) As String
'功能：根据病区返回其对应的科室ID串
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    strSQL = "Select Distinct 科室ID From 病区科室对应 Where 病区ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lngUnit)
    
    strSQL = "0"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If Not IsNull(rsTmp!科室ID) Then
                strSQL = strSQL & "," & rsTmp!科室ID
            End If
            rsTmp.MoveNext
        Next
    End If
    Get科室IDs = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserUnits(Optional ByVal blnDept As Boolean) As String
'功能：获取当前用户的所有部门ID或病区ID
'      如果操作员属于科室,则返回科室ID及科室所属病区ID
'      blnDept:True表示获取操作员所属科室,以及所属病区下的所有科室,否则返回操作员所属病区,以及所在科室所属的病区
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    On Error GoTo errH
    
    'union会去掉重复的
    If blnDept Then
        strSQL = "Select A.科室ID 部门ID From 病区科室对应 A,部门人员 B Where A.病区ID=B.部门ID And B.人员ID=[1]" & _
            " Union Select 部门ID as 部门ID From 部门人员 Where 人员ID=[1]"
    Else
        strSQL = "Select A.病区ID 部门ID From 病区科室对应 A,部门人员 B Where A.科室ID=B.部门ID And B.人员ID=[1]" & _
            " Union Select 部门ID as 部门ID From 部门人员 Where 人员ID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        GetUserUnits = GetUserUnits & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    
    If GetUserUnits = "" Then
        GetUserUnits = "0"
    Else
        GetUserUnits = Mid(GetUserUnits, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人科室ID(lng病人ID) As Long
'功能：获取在院病人当前病人科室ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.出院科室ID From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then Get病人科室ID = rsTmp!当前科室id
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GET部门名称(lngDeptID As Long, Optional ByRef rsDept As ADODB.Recordset) As String
'功能：获取部门名称
'参数：lngDeptID=部门ID
'返回：部门名称
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If rsDept Is Nothing Then
        strSQL = "Select 名称 from 部门表 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
    Else
        Set rsTmp = rsDept
        rsTmp.Filter = "ID=" & lngDeptID
        If rsTmp.RecordCount = 0 Then
            strSQL = "Select 名称 from 部门表 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngDeptID)
        End If
    End If
    
    If Not rsTmp.EOF Then GET部门名称 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub GetPersonnelIDCode(ByVal strName As String, Optional ByRef strID As String, Optional ByRef strCode As String)
'功能:根据人员姓名获取其ID和编码
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID,简码 From 人员表 Where 姓名=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", strName)
    
    If Not rsTmp.EOF Then
        strID = rsTmp!ID
        strCode = rsTmp!简码
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function GetDoctorOrNurse(ByVal BytType As Byte, Optional ByVal strUnits As String) As ADODB.Recordset
'功能：获取医生或护士列表.
'参数：bytType=0-医生，1-护士
'       strUnits=科室或病区ID串,如:18,26,31
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If strUnits <> "" Then
        If InStr(1, strUnits, ",") > 0 Then
            strSQL = " And Instr(','|| [2] || ',',',' || C.部门ID || ',')>0"
        Else
            strSQL = " And C.部门ID=[2]"
        End If
    End If
    
    strSQL = _
        "Select Distinct A.ID,A.编号,A.简码,A.姓名" & _
        " From 人员表 A,人员性质说明 B,部门人员 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And C.部门ID=D.部门ID" & _
        " And B.人员性质=[1] And D.服务对象 IN(1,2,3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & strSQL & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        " Order by 简码" '
    Set GetDoctorOrNurse = zlDatabase.OpenSQLRecord(strSQL, "mdlInPatient", IIf(BytType = 0, "医生", "护士"), strUnits)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function is产科(lng科室ID As Long, ByRef rs开单科室 As ADODB.Recordset) As Boolean
'功能：判断科室是否是产科性质
'参数：lng科室ID=指定科室ID
    is产科 = Sys.DeptHaveProperty(lng科室ID, "产科")
End Function

Public Function isMediRoom(lngID As Long) As Boolean
'功能：判断部门是否药房
'参数：lngID=部门ID
     isMediRoom = Sys.DeptHaveProperty(lngID, "中药房") Or Sys.DeptHaveProperty(lngID, "西药房") Or Sys.DeptHaveProperty(lngID, "成药房")
End Function

Public Function isCliniOrNurse(ByVal lngDept As Long) As Boolean
'功能:根据部门ID判断是否是临床或护理部门
    isCliniOrNurse = Sys.DeptHaveProperty(lngDept, "临床") Or Sys.DeptHaveProperty(lngDept, "护理")
End Function

Public Function GetNORule(ByVal intNo As Integer) As Integer
'功能:获取指定NO的编号规则
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 编号规则 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intNo)
    If Not rsTmp.EOF Then GetNORule = Val("" & rsTmp!编号规则)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetShareInvoiceGroupID(ByVal bytKind As Byte) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定票种的共用票据批次
    '编制:刘兴洪
    '日期:2011-04-29 10:24:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    If bytKind = 1 Or bytKind = 3 Then  '收费和结帐
        strSQL = "" & _
        "   Select A.ID,nvl(M.编码,' ') as 使用类别编码,A.使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B,票据使用类别 M" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And A.使用类别=M.名称(+) " & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别编码,剩余数量 Desc"
    ElseIf bytKind = 5 Then
        '就诊卡
        strSQL = "" & _
        "   Select A.ID,nvl(M.编码,' ') as 使用类别编码,M.ID as 使用类别ID,M.名称 as 使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B,医疗卡类别 M" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And to_number(nvl(A.使用类别,'0'))=M.ID(+) " & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别编码,剩余数量 Desc"
    ElseIf bytKind = 2 Then  '预交
        strSQL = "" & _
        "   Select A.ID,to_number(nvl(A.使用类别,'0')) as 使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别,剩余数量 Desc"
    Else
        strSQL = "" & _
        "   Select A.ID,A.使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
        "   Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
        "   Order by 使用类别,剩余数量 Desc"
    End If
    Set GetShareInvoiceGroupID = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, bytKind)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlCheckInvoiceOverplusEnough(ByVal bytKind As Byte, _
    ByVal intNum As Integer, Optional lng剩余数量 As Long, _
    Optional lng领用ID As Long = 0, Optional str使用类别 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查票据的剩余数量是否充足
    '入参:bytKind-1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '     intNum-当前对比的数量(-1代表不提醒)
    '     lng领用ID-只检查当前的领用票据(32455)
    '     str使用类别-使用类别
    '出参:lng剩余数量-返回当前剩余数量
    '返回:充足返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-28 17:16:16
    '问题:26948
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '-1代表不提醒
    If intNum = -1 Then zlCheckInvoiceOverplusEnough = True: Exit Function
    Err = 0: On Error GoTo Errhand:
    
    lng剩余数量 = 0
    
    strSQL = "" & _
        "   Select Sum(nvl(剩余数量,0)) as 剩余数量 " & vbNewLine & _
        "   From 票据领用记录" & vbNewLine & _
        "   Where 票种 = [1]  " & _
        "               And (nvl(使用类别,'LXH')=[4] or nvl(使用类别,'LXH')='LXH')  " & _
        "               And 领用人 = [2] And 使用方式 = 1 and nvl(剩余数量,0)>0" & vbNewLine & _
                    IIf(lng领用ID = 0, "", "             and ID=[3]") & _
        "   Union ALL " & _
        "   Select Sum(nvl(剩余数量,0)) as 剩余数量  " & _
        "   From 票据领用记录 A,人员表 B" & vbNewLine & _
          " Where A.票种=[1] And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
        "             And (nvl(A.使用类别,'LXH')=[4] or nvl(A.使用类别,'LXH')='LXH')  " & _
          "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                       IIf(lng领用ID = 0, "", "             and A.ID=[3]") & _
          "  "
    strSQL = "Select sum(剩余数量) as 剩余数量 From (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, UserInfo.姓名, lng领用ID, str使用类别)
    lng剩余数量 = Val(nvl(rsTemp!剩余数量))
    zlCheckInvoiceOverplusEnough = lng剩余数量 > intNum
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetInvoiceGroupID(ByVal bytKind As Byte, ByVal intNum As Integer, _
    Optional ByVal lngLastUseID As Long, Optional ByVal lngShareUseID As Long, _
    Optional ByVal strBill As String, Optional strUseType As String = "") As Long
'功能：获取张数够用并且指定票据在其可用范围内的领用ID
'参数：bytKind      =   票种
'      intNum       =   要打印的票据张数
'      lngLastUseID =   上次使用的领用ID
'      lngShareUseID=   本地参数指定的共用ID
'      strBill      =   当前票据号，用于检查领用批次的票据范围
'      strUseType-使用类别
'返回：
'      >0   =   成功，可用的领用ID
'      =0   =   失败
'      -1   =   没有自用(用完或不够，或未领用),未设置共用
'      -2   =   没有自用(用完或不够，或未领用),设置的共用已用完或不够
'      -3   =   指定票据号不在当前所有可用领用批次的有效票据号范围内
'      -4   =   指定批次的票据不够用
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
    Dim blnTmp As Boolean, i As Integer, lngReturn As Long
    
    On Error GoTo errH
    '1.上次的领用批次是否可用并够用
    If lngLastUseID > 0 Then
        strSQL = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录 " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3]  " & _
        "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, lngLastUseID, IIf(Trim(strUseType) = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then    '目前的票据号可能和上次不同，所以需要检查范围
                If strBill = "" Then GetInvoiceGroupID = lngLastUseID: Exit Function '可能没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngLastUseID: Exit Function
                
            ElseIf intNum > 1 Then  '不是确定领用批次调用时,当前票据号所在批次不够用
                GetInvoiceGroupID = -4: Exit Function
            End If
        End With
    End If
    
    '2.上次的领用批次不够用或不可用时,取自已领的并且自用的
    '  有多批，最近使用的优先,少的先用,先领先用
    strSQL = "" & _
    "   Select ID, 前缀文本, 开始号码, 终止号码" & vbNewLine & _
    "   From 票据领用记录" & vbNewLine & _
    "   Where 票种 = [1] And 剩余数量 >= [2] And 领用人 = [3]  " & _
    "           And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL ) " & _
    "           And 使用方式 = 1" & vbNewLine & _
    "   Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 desc, 开始号码"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, UserInfo.姓名, IIf(strUseType = "", "LXH", strUseType))
    With rsTmp
        For i = 1 To .RecordCount
            If strBill = "" Then GetInvoiceGroupID = !ID: Exit Function '第一次使用时没有当前票据号
            blnTmp = False
            strPre = "" & !前缀文本
            If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                blnTmp = True
            End If
            If Not blnTmp Then GetInvoiceGroupID = !ID: Exit Function
            .MoveNext
        Next
        lngReturn = IIf(.RecordCount > 0, -3, -1)
    End With
        
    '3.没有自用的,使用本地参数指定的共用批次
    If lngShareUseID > 0 Then
        strSQL = "" & _
        "   Select 前缀文本,开始号码,终止号码" & vbNewLine & _
        "   From 票据领用记录  " & _
        "   Where 票种=[1] And 剩余数量>=[2] And ID=[3] " & _
        "   And (Nvl(使用类别,'LXH')=[4] Or  使用类别 Is NULL) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, intNum, lngShareUseID, IIf(strUseType = "", "LXH", strUseType))
        With rsTmp
            If .RecordCount > 0 Then
                If strBill = "" Then GetInvoiceGroupID = lngShareUseID: Exit Function '第一次使用时没有当前票据号
                blnTmp = False
                strPre = "" & !前缀文本
                If UCase(Left(strBill, Len(strPre))) <> UCase(strPre) Then
                    blnTmp = True
                ElseIf Not (UCase(strBill) >= UCase(!开始号码) And UCase(strBill) <= UCase(!终止号码) And Len(strBill) = Len(!开始号码)) Then
                    blnTmp = True
                End If
                If Not blnTmp Then GetInvoiceGroupID = lngShareUseID: Exit Function
            End If
            lngReturn = IIf(.RecordCount > 0, -3, -2)
        End With
    End If
    GetInvoiceGroupID = lngReturn   '返回未找到的原因代码
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function CheckUsedBill(bytKind As Byte, ByVal lng领用ID As Long, _
    Optional ByVal strBill As String, _
     Optional ByVal strUseType As String = "") As Long
    '功能：检查当前操作员是否有可用票据领用(自用或共用),并返回可用的领用ID
    '参数：bytKind=票种
    '      lng领用ID=第一次检查时为本地设置的共用领用ID,以后为上次使用的领用ID
    '      strBill=要检查范围的票据号
    '说明：
    '    1.在检查范围时,如果病人有多批自用票据,则只要在其中一批之中就行了
    '    2.在检查范围时,长度也在检查范围之内。
    '    3.当有多批自用时,缺省按少的先用,先领先用,"最近使用的优先"原则
    '返回：
    '      正常：票据领用ID>0
    '      0=失败
    '      -1:没有自用(用完或未领用)、也没有共用(未设置)
    '      -2:设置的共用已用完
    '      -3:指定票据号不在当前可用范围内(包含多批自用票据的情况)

    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    
    '操作员有剩余的自用票据集
    strSQL = _
        "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
        "From 票据领用记录" & vbNewLine & _
        "Where 票种 = [1] And 使用方式 = 1 And 剩余数量 > 0 And 领用人 = [2] And (Nvl(使用类别,'LXH')=[3] or  使用类别 is NULL)" & vbNewLine & _
        "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 Desc, 开始号码"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, UserInfo.姓名, IIf(strUseType = "", "LXH", strUseType))
    If lng领用ID = 0 Then
        '程序中第一次检查,且没有设置本地共用
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '也没有自用票据
        '有自用票据,按优先原则返回
        lngReturn = rsSelf!ID
    Else
        '上次使用的领用ID或第一次检查的共用ID,先判断性质
        strSQL = "Select ID,使用方式,剩余数量,前缀文本,开始号码,终止号码 From 票据领用记录 Where 票种=[1]  And (Nvl(使用类别,'LXH')=[3] or  使用类别 is NULL) And ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", bytKind, lng领用ID, IIf(strUseType = "", "LXH", strUseType))
        '问题26352 by 张险华 2009-11-20
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If rsTmp!使用方式 = 2 Then '共用,要先看有没有自用
            If Not rsSelf.EOF Then
                '有自用的，优先
                lngReturn = rsSelf!ID
            Else
                '没有自用取共用
                If rsTmp!剩余数量 = 0 Then CheckUsedBill = -2: Exit Function '共用已经用完
                lngReturn = rsTmp!ID
                blnTmp = True
            End If
        Else
            '自用票据
            If rsTmp!剩余数量 > 0 Then
                '有剩余
                lngReturn = rsTmp!ID
            Else
                '其它有剩余的自用
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '其它自用也没有剩余
                lngReturn = rsSelf!ID
            End If
        End If
    End If
    
    '检查票号范围是否正确
    If strBill <> "" Then
        If blnTmp Then
            '在共用范围内范围判断
            If UCase(Left(strBill, Len(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)))) <> UCase(IIf(IsNull(rsTmp!前缀文本), "", rsTmp!前缀文本)) Then
                lngReturn = -3
            ElseIf Not (UCase(strBill) >= UCase(rsTmp!开始号码) And UCase(strBill) <= UCase(rsTmp!终止号码) And Len(strBill) = Len(rsTmp!开始号码)) Then
                lngReturn = -3
            End If
        Else
            '在可用自用范围内判断
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                blnTmp = True
            ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                blnTmp = True
            End If
            If blnTmp Then
                '该批不满足,则在其它自用中检查
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If UCase(Left(strBill, Len(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)))) <> UCase(IIf(IsNull(rsSelf!前缀文本), "", rsSelf!前缀文本)) Then
                        blnTmp = True
                    ElseIf Not (UCase(strBill) >= UCase(rsSelf!开始号码) And UCase(strBill) <= UCase(rsSelf!终止号码) And Len(strBill) = Len(rsSelf!开始号码)) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!ID: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function


Public Function GetNextBill(lng领用ID As Long) As String
'功能：根据领用批次ID,获取下一个实际票据号
'说明：1.当取不到范围内的有效票据时,返回空由用户输入
'      2.排开已报损的号码
    Dim rsMain As ADODB.Recordset
    Dim rsDelete As ADODB.Recordset
    Dim strSQL As String, strBill As String
    
    On Error GoTo errH
    
    strSQL = "Select 前缀文本,开始号码,终止号码,当前号码" & _
        " From 票据领用记录 Where 剩余数量>0 And ID=[1]"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, "取一下票据号", lng领用ID)
    If rsMain.EOF Then Exit Function
    
    If IsNull(rsMain!当前号码) Then
        strBill = UCase(rsMain!开始号码)
    Else
        strBill = UCase(zlCommFun.IncStr(rsMain!当前号码))
    End If
    
     '问题号:25448
     '刘兴洪:取消了;性质=1 And 原因=5 And 语句:原因是可能存在已经使用了的票据,使用了的,则排除
     '票种: 1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
     '性质:1-发出(原因中1、3、5属该性质)；2-收回(原因中2、4属该性质)
     '原因:1-正常发出票据；2-作废收回废票；3-重打发出票据；4-重打收回票据；5-毁损弃置票据
     
    strSQL = "Select Upper(号码) as 号码 From 票据使用明细" & _
        " Where 号码||''>=[1] And 领用ID=[2]" & _
        " Order by 号码"
        
    Set rsDelete = zlDatabase.OpenSQLRecord(strSQL, "取一下票据号", strBill, lng领用ID)
    Do While True
        '检查范围
        If Left(strBill, Len("" & rsMain!前缀文本)) <> UCase("" & rsMain!前缀文本) Then
            Exit Function
        ElseIf Not (strBill >= UCase(rsMain!开始号码) And strBill <= UCase(rsMain!终止号码)) Then
            Exit Function
        End If
                
        '排开报损号
        rsDelete.Filter = "号码='" & UCase(strBill) & "'"
        If rsDelete.EOF Then Exit Do
        strBill = zlCommFun.IncStr(strBill)
    Loop
   
    GetNextBill = strBill
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub UpdateShareID(ByVal lngModule As Long, ByVal strShareIDs As String, _
    Optional bytKind As Byte = 5, Optional strParName As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新共享批次的磁卡ID
    '入参:strShareIDs:共享的领用批次
    '        bytKind=1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡
    '        strParName-参数名(空时,以常用的名称为准)
    '编制:刘兴洪
    '日期:2011-07-26 17:09:17
    '目前暂对就诊调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strShare As String, varData As Variant, varTemp As Variant, strSQL As String
    Dim i As Long, strIDs As String, lngID As Long, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareIDs, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        If Val(varTemp(0)) <> 0 Then
            strIDs = strIDs & "," & Val(varTemp(0))
        End If
    Next
    If strShare <> "" Then
        strShare = Mid(strShare, 2)
            strSQL = "" & _
            "   Select  /*+ rule */ ID From 票据领用记录 A,Table(f_num2list([1])) J  " & _
            "   Where A.ID=J.Column_value  And A.票种=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查领用ID", lngID, bytKind)
        strShare = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",", ",")
            lngID = Val(varTemp(0))
            If lngID <> 0 Then
                rsTemp.Filter = "ID=" & lngID
                If rsTemp.RecordCount <> 0 Then
                     strShare = strShare & "|" & lngID & "," & varTemp(1)
                End If
                rsTemp.Filter = 0
            End If
            If Val(varTemp(0)) <> 0 Then
                strIDs = strIDs & "," & Val(varTemp(0))
            End If
        Next
    End If
    If strShare <> "" Then strShare = Mid(strShare, 2)
    Select Case bytKind
    Case 1  '收费收据
    Case 2  ' 预交收据
    Case 3   ' 结帐收据
    Case 4   ' 挂号收据
    Case 5   '就诊卡
        If strParName <> "" Then
            zlDatabase.SetPara strParName, strShare, glngSys, lngModule
        Else
            zlDatabase.SetPara "共用医疗卡批次", strShare, glngSys, lngModule
        End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'功能：判断是否存在指定的票据领用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo errH
    
    strSQL = "Select ID From 票据领用记录 Where ID=[1] And 票种=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查领用ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function




Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intTYPE As Integer
    Dim dtCurdate As Date, strMaxNO As String
    Dim strYearStr As String
    
    Err = 0: On Error GoTo errH:
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = zlStr.PrefixNO & strNO
        Exit Function
    End If
'    ElseIf intNum = 0 Then
'        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
'        Exit Function
'    End If
    GetFullNO = strNO
    
    strSQL = "Select 编号规则,Sysdate as 日期,最大号码 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, intNum)
    dtCurdate = Date
    If Not rsTmp.EOF Then
        intTYPE = Val("" & rsTmp!编号规则)
        dtCurdate = rsTmp!日期
        strMaxNO = nvl(rsTmp!最大号码)
    End If
    strYearStr = zlStr.PrefixNO
    If strMaxNO = "" Then strMaxNO = strYearStr & "000001"
    If intTYPE = 1 Then
        '按日编号
        strSQL = Format(CDate(Format(dtCurdate, "YYYY-MM-dd")) - CDate(Format(dtCurdate, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = zlStr.PrefixNO & strSQL & Format(Right(strNO, 4), "0000")
        Exit Function
    End If
    '按年编号
    If Len(strNO) = 6 Then
        GetFullNO = Left(strMaxNO, 2) & strNO: Exit Function
    End If
    GetFullNO = Left(strMaxNO, 2) & zlLeftPad(Right(strNO, 6), 6, "0")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillOperCheck(bytNO As Byte, strOperator As String, Datadd As Date, Optional strMessage As String = "操作", _
    Optional ByVal strNO As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal bytFlag As Byte = 2, Optional ByVal blnOnlyCheckLimit As Boolean, Optional ByVal blnCheckOperator As Boolean = True, _
    Optional ByVal blnCheckCur As Boolean = True) As Boolean
'功能：判断当前人员对单据是否有操作权限
'参数：
'   bytNO：1-挂号单据,2-收费单,3-划价单,4-门诊记帐,5-住院记帐,6-预交款,7-结帐单据,8-就诊卡
'   strOperator：单据实际的操作员
'   DatAdd：单据的登记时间
'   strNO   ：检查金额上限时用来确定单据
'   lngPatientID：检查金额上限时，对于记帐表用来确定单据中的病人
'   bytFlag：1-收费单,2-记帐单,3-记帐单
'   blnOnlyCheckLimit：只检查金额上限
'   blnCheckOperator：要检查是否允许操作他人单据
'   blnCheckCur：是否检查金额上限
'返回：是否有操作权限
'说明：具体的提示在本函数中。

    Dim strSQL As String, strBill As String
    Dim rsTmp As ADODB.Recordset
    Dim curTmp As Currency
    Dim int来源 As Integer
    
    If bytNO = 1 Or bytNO = 2 Or (bytNO = 3 And bytFlag = 1) Or bytNO = 4 Then
        int来源 = 1
    Else
        int来源 = 2
    End If
    
    If glngSys Like "8??" Then
        strBill = Switch(bytNO = 1, "挂号单据", bytNO = 2, "收费单据", bytNO = 3, _
            "划价单据", bytNO = 4, "记帐单据", bytNO = 5, "记帐单据", _
            bytNO = 6, "预交款单据", bytNO = 7, "结帐单据", bytNO = 8, "会员卡")
        
        
    Else
        strBill = Switch(bytNO = 1, "挂号单据", bytNO = 2, "收费单据", bytNO = 3, _
            "划价单据", bytNO = 4, "记帐单据", bytNO = 5, "记帐单据", _
            bytNO = 6, "预交款单据", bytNO = 7, "结帐单据", bytNO = 8, "就诊卡")
    End If
        
    On Error GoTo errH
    
    strSQL = "Select Nvl(时间限制,0) as 时间限制,Nvl(他人单据,0) as 他人单据,Nvl(金额上限,0) as 金额上限 From 单据操作控制 Where 人员ID=[1] And 单据=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.ID, bytNO)
    If rsTmp.EOF Then
        BillOperCheck = True
        Exit Function
    Else
        If Not blnOnlyCheckLimit Then
            If rsTmp!他人单据 = 0 And blnCheckOperator Then
                If strOperator <> UserInfo.姓名 Then
                    MsgBox "你没有权限对" & strOperator & "处理的" & strBill & "进行" & strMessage & "！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            If rsTmp!时间限制 > 0 Then
                If Int(zlDatabase.Currentdate) - Int(CDate(Datadd)) + 1 > rsTmp!时间限制 Then
                    MsgBox "你只能对 " & rsTmp!时间限制 & " 天内处理的" & strBill & "进行" & strMessage & "！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If rsTmp!金额上限 > 0 And blnCheckCur Then
            If strNO <> "" Then
                curTmp = GetBillMoney(int来源, strNO, lngPatientID, bytFlag)
                If curTmp >= rsTmp!金额上限 Then
                    MsgBox "你只能对 " & rsTmp!金额上限 & " 元以下的" & strBill & "进行" & strMessage & "！" & _
                    vbCrLf & "单据[" & strNO & "]的实收金额合计为:" & curTmp, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        BillOperCheck = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetBillMoney(ByVal int来源 As Integer, strNO As String, Optional lng病人ID As Long, Optional ByVal bytFlag As Byte = 2) As Currency
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张单据的实收金额合计,或一张记帐表中指定病人的实收金额合计
    '入参：int来源-1-门诊,2-住院
    '      bytFlag-1-收费单,2-记帐单,3-记帐单(自动记帐单)
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-02 14:26:50
    '说明：int来源增加了参数
    '------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    
    On Error GoTo errH
    
    If lng病人ID = 0 Then
        strSQL = "Select Sum(实收金额) as 金额 From  " & IIf(int来源 = 1, "门诊费用记录", " 住院费用记录") & " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1)"
    Else
        strSQL = "Select Sum(实收金额) as 金额 From " & IIf(int来源 = 1, "门诊费用记录", " 住院费用记录") & " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1) And 病人ID=[3]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, lng病人ID)
    
    If Not rsTmp.EOF Then GetBillMoney = Val("" & rsTmp!金额)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadBillInfo(ByVal int来源 As Integer, ByVal strNO As String, _
    ByVal intFlag As Integer, ByRef strOperator As String, ByRef Datadd As Date, _
    Optional ByRef lng病人ID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：读取一张单据的操作员和登记时间
    '入参：int来源-1-门诊,2-住院
    '      intFlag:-1=结帐,-2=预交,-3=补结算，其它同 住院费用记录或门诊费用记录.记录性质(1-收费记录；2(12)-记帐记录；3(13)-自动记帐记录;4-挂号记录；5(15)-就诊卡记录)
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-02 16:03:22
    '说明：本函数仅配合函数BillOperCheck使用。
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    
    If intFlag = -1 Then
        strSQL = "Select 操作员姓名,收费时间 as 登记时间, 病人ID From 病人结帐记录 Where NO=[1] And 记录状态 IN(1,3)"
    ElseIf intFlag = -2 Then
        strSQL = "Select 操作员姓名,收款时间 as 登记时间,病人ID  From 病人预交记录 Where NO=[1] And 记录状态 IN(1,3)"
    ElseIf intFlag = -3 Then
        strSQL = "Select 操作员姓名, 登记时间, 病人id From 费用补充记录 Where NO = [1] And 记录状态 In (1, 3) And Rownum < 2"
    Else
        strSQL = "Select Nvl(操作员姓名,划价人) as 操作员姓名,登记时间,病人ID  From " & IIf(int来源 = 1, "门诊费用记录", " 住院费用记录") & " Where NO=[1] And 记录性质=[2] And 记录状态 IN(0,1,3) And RowNum=1"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, intFlag)
    If Not rsTmp.EOF Then
        strOperator = rsTmp!操作员姓名
        Datadd = rsTmp!登记时间
        lng病人ID = Val(nvl(rsTmp!病人ID))
        ReadBillInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng病人ID As Long) As Currency
'功能:获取指定病人的划价单金额合计
'入参:IntType:0-门诊,1-住院,2-所有
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '记帐报警包含所有住院划价费用
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara("记帐报警包含所有住院划价费用", glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(主页ID,0) = (Select Nvl(主页ID,0) From 病人信息 Where 病人ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
        "   From 住院费用记录 " & _
        "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志=2" & strWhere
    Else
        If intTYPE = 2 Then
            strSQL = "Select Nvl(Sum(实收金额),0) As 划价费用合计 From 门诊费用记录 Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]"
            strSQL = strSQL & " union ALL  Select Nvl(Sum(实收金额),0) As 划价费用合计 From 住院费用记录 Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]"
            strSQL = "Select Sum(nvl(划价费用合计,0)) as 划价费用合计 From ( " & strSQL & ")"
        Else
            strSQL = "" & _
            "   Select Nvl(Sum(实收金额),0) As 划价费用合计 " & _
            "   From 门诊费用记录  " & _
            "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]  and 门诊标志<>2" & _
            "   Union ALL   " & _
            "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
            "   From 住院费用记录 " & _
            "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] and 门诊标志<>2 "
            strSQL = "" & _
            "   Select Sum(nvl(划价费用合计,0)) as 划价费用合计  " & _
            "   From ( " & strSQL & ")"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取指定病人的划价总额", lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!划价费用合计
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
'功能：获取指定病人当天发生的费用总额
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function HavedInCost(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '零费用返回False,否则返回true
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT SUM(实收金额) 实收金额 FROM 住院费用记录 where 病人ID=[1] AND 主页ID=[2] and 记录状态<>0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否有费用", lng病人ID, lng主页ID)
    If Not rsTemp Is Nothing Then
        If Not rsTemp.EOF Then
            If nvl(rsTemp!实收金额, 0) <> 0 Then HavedInCost = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HavedDirections(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能:检查病人本次住院是否已经产生医嘱
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "SELECT 1 FROM 病人医嘱记录 Where 病人ID = [1] And 主页id = [2] And RowNum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是存在医嘱", lng病人ID, lng主页ID)
    If Not rsTemp Is Nothing Then
        HavedDirections = rsTemp.EOF = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMoneyInfo(lng病人ID As Long, Optional dblModiMoney As Double, _
    Optional blnInsure As Boolean, _
    Optional int类型 As Integer = -1, _
    Optional bln按类型统计 As Boolean = False, _
    Optional bytModiMoneyType As Byte = 0, _
    Optional ByVal blnFamilyMoney As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的剩余额
    '入参:blnInsure=是否排开医保病人的预结费用
    '       curModiMoney=修改时,原单据的当前病人的费用合计
    '       int类型:类型(0-门诊和住院共用;1-门诊;2-住院),-1表示所有
    '       bytModiMoneyType-修改费用的类别(在按类别统计时有效)
    '       blnFamilyMoney-是否读取家属余额
    '出参:
    '返回:病人剩余额
    '编制:刘兴洪
    '日期:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, bln医保 As Boolean, lng主页ID As Long
    Dim strSQL As String
    On Error GoTo errH
    If blnInsure Then
        strSQL = "Select A.险类,A.主页ID From 病案主页 A,病人信息 B" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
                " And B.病人ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
        If Not rsTmp.EOF Then
            bln医保 = Not IsNull(rsTmp!险类)
            lng主页ID = rsTmp!主页ID
        End If
    End If
    strSQL = "Select " & IIf(bln按类型统计, "类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
            "       Nvl(费用余额,0) As 费用余额,Nvl(预交余额,0) As 预交余额" & _
            " From 病人余额" & _
            " Where 性质=1 And 病人ID=[1] " & IIf(int类型 = -1, "", " And 类型=[4]")
    '79868,读取病人家属余额
    If blnFamilyMoney Then
        strSQL = strSQL & " Union All " & _
                " Select " & IIf(bln按类型统计, "a.类型,", "") & IIf(blnFamilyMoney, "1 As 家属,", "") & _
                "       Nvl(a.费用余额, 0) As 费用余额, Nvl(a.预交余额, 0) As 预交余额" & _
                " From 病人余额 A, 病人家属 B" & _
                " Where a.病人id = b.家属id And b.病人id = [1] And a.性质 = 1 " & _
                "       And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) " & _
                IIf(int类型 = -1, "", " And 类型=[4]")
    End If
  
    If dblModiMoney <> 0 Then   '必须要用Union方式,如果直接去减,在病人余额无记录时,不会返回记录
        strSQL = strSQL & " Union All " & _
                " Select " & IIf(bln按类型统计, "[4] as 类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
                "       -1*[3] as 费用余额,0 as 预交余额 From Dual"
    End If
    
    '如果为医保住院病人，则在费用余额中排开预结中的费用(用于报警)
    If blnInsure And bln医保 Then
        strSQL = strSQL & " Union All " & _
        " Select  " & IIf(bln按类型统计, "Decode(主页ID,NULL,1,0,1,2) as 类型,", "") & IIf(blnFamilyMoney, "0 As 家属,", "") & _
        "       -1*Nvl(金额,0) as 费用余额,0 as 预交余额" & _
        " From 保险模拟结算" & _
        " Where 病人ID=[1] And 主页ID=[2] "
    End If
    strSQL = "Select " & IIf(bln按类型统计, "类型,", "") & IIf(blnFamilyMoney, "家属,", "") & _
            "       nvl(Sum(费用余额),0) as 费用余额,nvl(Sum(预交余额),0) as 预交余额 " & _
            " From (" & strSQL & ")" & vbCrLf & _
            IIf(bln按类型统计 And blnFamilyMoney, " Group by 类型,家属", _
                IIf(bln按类型统计, " Group by 类型", IIf(blnFamilyMoney, " Group by 家属", "")))
    
    Set GetMoneyInfo = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID, dblModiMoney, int类型)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function isYBPati(lng病人ID As Long, Optional blnIn As Boolean, Optional int险类 As Integer) As Boolean
'功能：判断一个住院病人是否医保病人
'参数：blnIN=是否必须在院
'      int险类=返回险类
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select A.险类 From 病案主页 A,病人信息 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        " And B.病人ID=[1] " & IIf(blnIn, " And A.出院日期 is NULL", "")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        isYBPati = Not IsNull(rsTmp!险类)
        int险类 = nvl(rsTmp!险类, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetStockSet(ByVal lng药房ID As Long, ByVal lng药品ID As Long) As Recordset
    Dim strSQL As String
    
    If Val(zlDatabase.GetPara(150, glngSys)) = 0 Then '分批药品出库方式：0-按批次先进先出，1-按效期最近先出,效期相同，则再按批次先进先出
        strSQL = "Nvl(批次,0)"
    Else
        strSQL = "效期,Nvl(批次,0)" '效期为空则排在最后
    End If
    
    '药房不分批药品不管效期(这里的库房一定是药房)
    strSQL = "Select Nvl(批次,0) as 批次,Nvl(可用数量,0) as 库存," & _
        " Nvl(零售价,Nvl(Decode(Nvl(实际数量,0),0,0,实际金额/实际数量),0)) as 时价," & _
        " Nvl(实际差价,0) as 实际差价,Nvl(实际金额,0) as 实际金额" & _
        " From 药品库存" & _
        " Where 库房ID=[1] And 药品ID=[2] And Nvl(可用数量,0)>0" & _
        " And 性质=1 And (Nvl(批次,0)=0 Or 效期 is NULL Or 效期>Trunc(Sysdate))" & _
        " Order by " & strSQL
        
    On Error GoTo errH
    Set GetStockSet = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng药房ID, lng药品ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get时价药品应收金额(ByVal lng药房ID As Long, ByVal lng药品ID As Long, ByRef dblAllTime As Double, ByVal strDec As String, ByRef dblPriceSingle As Double) As Currency
'功能：获取分批时价药品的应收金额（根据不同的出库方式按批次合计）
'参数：
'      strDec-费用金额保留位数
'      dblAllTime-传入为出库总数量(售价数量)，传出如果为0则表示库存足够，否则表示库存不足
'      dblPriceSingle-只有一个批次时返回该批次的单价，避免根据数量先乘再除后因四舍五入引起不同的数量的单价不同

    Dim rsPrice As ADODB.Recordset
    Dim dblPrice As Double, dblCurTime As Double, i As Long
    
    Set rsPrice = GetStockSet(lng药房ID, lng药品ID)
    '时价=总金额/总数量
    dblPrice = 0 '本笔总应收金额
    For i = 1 To rsPrice.RecordCount
        If dblAllTime = 0 Then Exit For
        '取小者
        If dblAllTime <= rsPrice!库存 Then
            dblCurTime = dblAllTime
        Else
            dblCurTime = rsPrice!库存
        End If
        If i = 1 Then
            dblPriceSingle = Format(rsPrice!时价, gstrFeePrecisionFmt)
        Else
            dblPriceSingle = 0
        End If
        dblPrice = dblPrice + Format(dblCurTime * Format(rsPrice!时价, gstrFeePrecisionFmt), strDec)
        dblAllTime = dblAllTime - dblCurTime
        rsPrice.MoveNext
    Next
    
    Get时价药品应收金额 = dblPrice
End Function

Public Function GetAuditRecord(lng病人ID As Long, lng主页ID As Long, Optional lng项目id As Long) As ADODB.Recordset
'功能：获取指定病人的费用审批项目,当未设置使用限量或已用数量为空时,可用数量为空
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 项目Id,使用限量,已用数量,使用限量-Nvl(已用数量,0) 可用数量 From 病人审批项目 " & _
            "Where 病人ID=[1] And 主页ID=[2]" & IIf(lng项目id <> 0, " And 项目ID=[3]", "")
    Set GetAuditRecord = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID, lng项目id)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function BillExistInsure(ByVal strNO As String, Optional ByVal strTime As String, _
    Optional ByVal blnAuditing As Boolean, Optional ByVal bytFlag As Byte = 2) As Integer
'功能：判断指定的住院记帐单据是否对医保病人记的帐
'参数：strNO=记帐单据号
'      blnAuditing=是否用于记帐审核,只检查未审核的部份内容
'      bytFlag=2-人工记帐单,3-自动记帐单
'返回：如果是则返回病人险类
'说明：1.只管住院医保病人,不管门诊病人的医技记帐
'      2.记帐表只返回第一个病人的险类
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.险类 From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=[2] And A.记录状态" & IIf(blnAuditing, "=0", " IN(0,1,3)") & " And B.险类 is Not NULL" & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
        IIf(strTime <> "", " And A.登记时间=[3]", "")
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag)
    End If

    If Not rsTmp.EOF Then BillExistInsure = rsTmp!险类
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AdjustCpt(lngID As Long)
    '功能：药品调价
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "zl_药品收发记录_Adjust(" & lngID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, App.ProductName)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function Get医保大类(ByVal lng收费细目ID As Long, ByVal int险类 As Integer) As String
'功能：获取指定收费项目的保险大类名称
'参数：
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select N.名称" & _
        " From 保险支付项目 M,保险支付大类 N " & _
        " Where M.收费细目ID=[1] And M.险类=[2] And M.大类ID=N.ID"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng收费细目ID, int险类)
    If rsTmp.RecordCount > 0 Then Get医保大类 = rsTmp!名称
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get处方限量(lng药品ID As Long) As Double
'功能：获取指定药品的处方限量,以零售单位返回。
'参数：lng药品ID=药品ID
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Nvl(A.处方限量,0) as 处方限量 From 药品特性 A,药品规格 B Where A.药名ID=B.药名ID And B.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng药品ID)
    If Not rsTmp.EOF Then Get处方限量 = rsTmp!处方限量
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get处方职务(lng药品ID As Long) As String
'功能：根据药品ID获取其处方职务
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Get处方职务 = "00"
    strSQL = "Select Nvl(B.处方职务,'00') as 处方职务 From 药品规格 A,药品特性 B Where A.药名ID=B.药名ID And A.药品ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng药品ID)
    If Not rsTmp.EOF Then Get处方职务 = rsTmp!处方职务
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetDiseaseCode(ByRef frmParent As Object, ByRef blnCancel As Boolean, _
    ByVal strInput As String, ByVal strSex As String, ByVal strKind As String, _
    ByVal X As Long, ByVal Y As Long, ByVal txtHeight As Long, Optional ByVal bytSize As Byte) As ADODB.Recordset
'功能:根据输入的字符返回对应的疾病编码记录集
'参数:strCode-输入值,strSex-性别限制,strKind-疾病编码类别
'     x,y弹出的选择器在屏幕中显示的坐标位置,txtHeight-输入框的高度,blnCnacel是否取消选择
'     ："bytSize=?"表示设置字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体。
    Dim strSQL As String, strCode As String
    Dim strLike As String, strWhere As String, lngCodeKind As Long
    
    If Trim(strInput) = "" Then Exit Function
    
    strLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    strCode = strLike & UCase(Trim(strInput)) & "%"
    '简码匹配方式：0-拼音,1-五笔,2-两者
    lngCodeKind = Val(zlDatabase.GetPara("简码方式"))
    
    
    If zlCommFun.IsCharAlpha(strInput) Then
        If lngCodeKind = 0 Then
            strWhere = "(A.编码 Like [1] Or A.简码 Like [1])"
        ElseIf lngCodeKind = 1 Then
            strWhere = "(A.编码 Like [1] Or A.五笔码 Like [1])"
        Else
            strWhere = "(A.编码 Like [1] Or A.简码 Like [1] Or A.五笔码 Like [1])"
        End If
    ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
        strWhere = "A.编码 Like [1]"
    ElseIf zlCommFun.IsCharChinese(strInput) Then
        strWhere = "A.名称 Like [1]"
    Else
        If lngCodeKind = 0 Then
            strWhere = "(A.名称 Like [1] Or A.编码 Like [1] Or A.简码 Like [1])"
        ElseIf lngCodeKind = 1 Then
            strWhere = "(A.名称 Like [1] Or A.编码 Like [1] Or A.五笔码 Like [1])"
        Else
            strWhere = "(A.名称 Like [1] Or A.编码 Like [1] Or A.简码 Like [1] Or A.五笔码 Like [1])"
        End If
    End If
    If strSex <> "" Then strWhere = strWhere & " And (A.性别限制='" & strSex & "' Or A.性别限制 is NULL)"
       
       
'    If strKind <> "" Then
'        strSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.五笔码,A.说明,A.性别限制,B.类别" & _
'            " From 疾病编码目录 A,疾病编码类别 B" & _
'            " Where A.类别=B.编码 And A.类别=[2] And Rownum<=100 And " & strWhere & _
'            " Order by A.类别,A.编码"
'    Else
    '90044取消限制返回行数
        strSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.五笔码,A.说明,A.性别限制" & _
            " From 疾病编码目录 A" & _
            " Where A.类别=[2] And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And " & strWhere & _
            " Order by A.编码"
            
'    End If
    
    Set GetDiseaseCode = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "疾病编码输入", 1, "", "请选择", False, False, True, X, Y, txtHeight, blnCancel, False, True, strCode, strKind, "bytSize=" & bytSize)
End Function

Public Function GetDiseaseCodeNew(ByRef frmParent As Object, ByRef blnCancel As Boolean, _
    ByVal strInput As String, ByVal strSex As String, ByVal strKind As String, _
    ByVal X As Long, ByVal Y As Long, ByVal txtHeight As Long, Optional ByVal bytSize As Byte) As ADODB.Recordset
'功能:根据输入的字符返回对应的疾病编码记录集,最多返回100条记录
'参数:strCode-输入值,strSex-性别限制,strKind-疾病编码类别
'     x,y弹出的选择器在屏幕中显示的坐标位置,txtHeight-输入框的高度,blnCnacel是否取消选择
'     ："bytSize=?"表示设置字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体。
    Dim strSQL As String, strCode As String, strRight As String
    Dim strLike As String, strWhere As String, lngCodeKind As Long
    
    If Trim(strInput) = "" Then Exit Function
    
    strLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    strCode = strLike & UCase(Trim(strInput)) & "%"
    strRight = UCase(Trim(strInput)) & "%"
    '简码匹配方式：0-拼音,1-五笔,2-两者
    lngCodeKind = Val(zlDatabase.GetPara("简码方式"))

    If zlCommFun.IsCharChinese(strInput) Then
        strSQL = "名称 Like [2] or '('||编码||')'||名称 Like [2]" '输入汉字时只匹配名称
    Else
        strSQL = "编码 Like [1] Or 名称 Like [2] Or " & IIf(lngCodeKind = 0, "简码", "五笔码") & " Like [2]"
    End If
    
'    If strSex <> "" Then strWhere = strWhere & " And (A.性别限制='" & strSex & "' Or A.性别限制 is NULL)"
       
    strSQL = _
                " Select ID,ID as 项目ID,编码,附码,名称," & IIf(lngCodeKind = 0, "简码", "五笔码 as 简码") & ",说明" & _
                " From 疾病编码目录 Where Instr([3],类别)>0 And (" & strSQL & ")" & _
                IIf(strSex <> "", " And (性别限制=[4] Or 性别限制 is NULL)", "") & _
                " And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by 编码"

'    strSql = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.五笔码,A.说明,A.性别限制" & _
'             " From  疾病编码目录 A" & _
'             " Where Rownum<=100 And A.类别=[3] And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And " & strWhere & _
'             " Order by A.编码"

    
    Set GetDiseaseCodeNew = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "疾病编码输入", 1, "", "请选择", False, False, True, X, Y, txtHeight, blnCancel, False, True, strRight, strCode, strKind, strSex, "bytSize=" & bytSize)
End Function

Public Function HaveOut(lng病人ID As Long) As Boolean
'功能：判断病人当前是否已经出院
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 出院日期 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And A.主页ID=B.主页ID and A.病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If rsTmp.EOF Then HaveOut = True: Exit Function '未入院病人当作出院病人
    If Not IsNull(rsTmp!出院日期) Then
        If rsTmp!出院日期 <= zlDatabase.Currentdate Then HaveOut = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function HaveExecute(ByVal int来源 As Integer, ByVal strNO As String, _
    ByVal int记录性质 As Integer, Optional blnAll As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：判断费用单据是否包含完全执行或部分执行的内容
    '入参：int来源-1-门诊;2-住院
    '      strNO=费用单据号,
    '      int记录性质=记录性质(1-收费,2-记帐)
    '      blnALL=判别单据中是否全部为完全执行或部分执行的内容
    '返回：存在执行的，返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-03-02 16:23:05
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    On Error GoTo errH
    If int记录性质 = 1 Then
        strWhere = " And mod(记录性质,10)=[2]"
    Else
        strWhere = " And 记录性质=[2]"
    End If
    strWhere = strWhere & " And " & IIf(blnAll, " Not", "") & " 执行状态 IN(1,2)"
    
    strSQL = "" & _
    " Select Nvl(Count(ID),0) as 数目" & _
    " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
    " Where NO=[1] And 记录状态 IN(0,1,3)  " & strWhere

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, int记录性质)
    
    If blnAll Then
        HaveExecute = (rsTemp!数目 = 0)
    Else
        HaveExecute = (rsTemp!数目 > 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function HaveBilling(ByVal int来源 As Integer, ByVal strNO As String, Optional ByVal blnAll As Boolean = True, _
    Optional ByVal strTime As String, Optional ByVal bytFlag As Byte = 2) As Integer
    '------------------------------------------------------------------------------------------------------------------------
    '功能：判断一张记帐单/表是否已经结帐
    '入参：int来源-1-门诊;2-住院
    '      strNO=记帐单据号,不分门诊及住院
    '      blnALL=是否对整张单据内容进行判断,否则只对未销帐部分进行判断
    '出参：
    '返回：0-未结帐,1=已全部结帐,2-已部分结帐
    '编制：刘兴洪
    '日期：2010-03-02 16:37:22
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngTmp As Long
    
    On Error GoTo errH
        
    '求未作废的费用行
    strSQL = _
        " Select 序号 From (" & _
            " Select 记录状态,执行状态,Nvl(价格父号,序号) as 序号, Avg(Nvl(付数, 1) * 数次) As 数量" & _
            " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
            " Where NO=[1] And 记录性质=[2]" & _
            " Group by 记录状态,执行状态,Nvl(价格父号,序号))" & _
        " Group by 序号 Having Sum(数量)<>0"
    
    '求每行的结帐情况
    strSQL = _
        "Select Nvl(价格父号,序号) as 序号,Sum(Nvl(结帐金额,0)) as 结帐金额" & _
        " From " & IIf(int来源 = 1, "门诊费用记录", "住院费用记录") & _
        " Where NO=[1] And mod(记录性质,10)= [2]" & _
        IIf(Not blnAll, " And Nvl(价格父号,序号) IN(" & strSQL & ")", "") & _
        IIf(strTime <> "", " And 登记时间=[3]", "") & _
        " Group by Nvl(价格父号,序号)"
    
    If strTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag, CDate(strTime))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO, bytFlag)
    End If
    
    If Not rsTmp.EOF Then
        lngTmp = rsTmp.RecordCount '单据行数
        rsTmp.Filter = "结帐金额<>0"
        If rsTmp.EOF Then
            HaveBilling = 0 '无结帐行
        ElseIf rsTmp.RecordCount = lngTmp Then
            HaveBilling = 1 '全部行已结帐
        ElseIf rsTmp.RecordCount > 0 Then
            HaveBilling = 2 '部分行已结帐
        End If
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check发生时间(ByVal varDate As Date, ByVal varPaitOrNO As Variant) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：检查发生时间是否合法
    '参数：varDate=发生时间
    '      varPaitOrNO=病人ID或记帐单号(可能是多病人单)
    '返回：错误提示
    '编制:刘兴洪
    '日期:2015-07-10 15:47:18
    '说明：1.检查发生时间不能早于病人的入院时间
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    If TypeName(varPaitOrNO) = "String" Then
        strSQL = "Select Distinct 姓名,病人ID,主页ID  From 住院费用记录 Where 记录性质=2 And NO=[1]"
            
        strSQL = "Select A.姓名,B.主页ID,B.入院日期" & _
            " From (" & strSQL & ") A,病案主页 B" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    Else
        strSQL = "" & _
        "Select nvl(B.姓名,A.姓名) as 姓名,B.主页ID,B.入院日期 From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[2]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, CStr(varPaitOrNO), Val(varPaitOrNO))
    For i = 1 To rsTmp.RecordCount
        If Format(varDate, "yyyy-MM-dd HH:mm:ss") < Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm:ss") Then
            Check发生时间 = "费用的发生时间不能小于病人""" & rsTmp!姓名 & """的入院时间:" & Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm:ss") & "！"
            Exit Function
        End If
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnAuditReFee(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查病人是否存在未批准的退费申请
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSQL = "Select 1" & vbNewLine & _
            "From Dual" & vbNewLine & _
            "Where Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From 住院费用记录 A" & vbNewLine & _
            "       Where A.病人id = [1] And A.主页id = [2] And Exists (Select 1 From 病人费用销帐 B Where B.费用id = A.ID And B.状态 = 0))"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, lng主页ID)
    GetUnAuditReFee = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get保险特准项目(lng病人ID As Long, strField As String, Optional intInsure As Integer) As String
'功能：根据医保病人的病种获取保险特准项目的条件,该条件用于收费细目表
'参数：strField=别名字段,如"C.收费细目ID"
'说明：判断病种后，可以直接返回SQL语句，但效率不高
'    IN (
'        Select 收费细目ID From 保险支付项目
'        Where 险类 = XXXX
'            And 收费细目ID IN (Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=1 And 病种ID=XXXX)
'        ) Or 0=(Select Count(*) From 保险特准项目 Where Nvl(大类,0)=0 And 性质=1 And 病种ID=XXXX)
'
'    Not IN (
'        Select 收费细目ID From 保险支付项目
'        Where 险类 = XXXX
'            And 收费细目ID IN (Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=2 And 病种ID=XXXX)
'        ) Or 0=(Select Count(*) From 保险特准项目 Where Nvl(大类,0)=0 And 性质=2 And 病种ID=XXXX)
'
'    IN (
'        Select 收费细目ID From 保险支付项目
'        Where 险类 = XXXX
'            And Nvl(大类ID,0) IN (Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=1 And 病种ID=XXXX)
'        ) Or 0=(Select Count(*) From 保险特准项目 Where Nvl(大类,0)=1 And 性质=1 And 病种ID=XXXX)
'
'    Not IN (
'        Select 收费细目ID From 保险支付项目
'        Where 险类 = XXXX
'            And Nvl(大类ID,0) IN (Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=2 And 病种ID=XXXX)
'        ) Or 0=(Select Count(*) From 保险特准项目 Where Nvl(大类,0)=1 And 性质=2 And 病种ID=XXXX)

    Dim rsTmp As ADODB.Recordset
    Dim lng病种ID As Long, int险类 As Integer, strSQL As String
    Dim strA1 As String, strA2 As String, strB1 As String, strB2 As String
    
    On Error GoTo errH
    
    '先取病人病种,是该病种是否有各类特准项目设置
    strSQL = _
        " Select A.险类,A.病种ID,Nvl(B.大类,0) as 大类,B.性质,Count(*)" & _
        " From 保险帐户 A,保险特准项目 B" & _
        " Where Nvl(A.病种ID,0)=B.病种ID And Nvl(A.病种ID,0)<>0" & _
        " And B.性质 IN(1,2) And A.病人ID=[1]" & _
        " Group by A.险类,A.病种ID,Nvl(B.大类,0),B.性质"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If rsTmp.EOF Then Exit Function
    
    lng病种ID = rsTmp!病种ID
    int险类 = rsTmp!险类
    
    '允许的收费细目
    rsTmp.Filter = "大类=0 And 性质=1"
    If Not rsTmp.EOF Then
        strA1 = strField & _
            " IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 收费细目ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=1 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '允许的保险大类
    rsTmp.Filter = "大类=1 And 性质=1"
    If Not rsTmp.EOF Then
        strA2 = strField & _
            " IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 大类ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=1 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '禁止的收费细目
    rsTmp.Filter = "大类=0 And 性质=2"
    If Not rsTmp.EOF Then
        strB1 = strField & _
            " Not IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 收费细目ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=0 And 性质=2 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '禁止的保险大类
    rsTmp.Filter = "大类=1 And 性质=2"
    If Not rsTmp.EOF Then
        strB2 = strField & _
            " Not IN (" & _
            "   Select 收费细目ID From 保险支付项目" & _
            "   Where 险类 = " & int险类 & _
            "   And 大类ID IN (" & _
            "       Select 收费细目ID From 保险特准项目 Where Nvl(大类,0)=1 And 性质=2 And 病种ID=" & lng病种ID & ")" & _
            ")"
    End If
    
    '组合SQL(允许部分要用Or)
    strSQL = ""
    If strA1 <> "" And strA2 <> "" Then
        strSQL = " And (" & strA1 & " Or " & strA2 & ")"
    Else
        If strA1 <> "" Then strSQL = " And " & strA1
        If strA2 <> "" Then strSQL = " And " & strA2
    End If
    If strB1 <> "" Then strSQL = strSQL & " And " & strB1
    If strB2 <> "" Then strSQL = strSQL & " And " & strB2
        
    Get保险特准项目 = strSQL
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function is代收款(ByVal strNO As String) As Boolean
'功能:判断一张预交款单据是否是收取的代收款
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select NO From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "Where A.NO = [1] And A.记录性质 = 1 And A.结算方式 = B.名称 And B.性质 = 5"
            
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO)
    is代收款 = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function CheckFeeItemAvailable(ByVal lngFeeItemID As Long, ByVal bytFlag As Byte) As Boolean
'功能:检查收费项目是否未停用,并且服务于病人
'参数:bytFlag:服务对象:1-门诊,2-住院
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select 1 From 收费项目目录 Where ID = [1] And (撤档时间 is Null Or 撤档时间 > Sysdate) And 服务对象 In (" & bytFlag & ",3)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngFeeItemID)
    CheckFeeItemAvailable = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckTextLength(strName As String, txtObj As TextBox) As Boolean
'功能:检查并提示文本框输入长度是否超限
    CheckTextLength = zlControl.TxtCheckInput(txtObj, strName, , True)
End Function

Public Function ReCalcOld(ByVal DateBir As Date, ByRef cbo年龄单位 As ComboBox, Optional ByVal lng病人ID As Long, Optional ByVal blnSetControl As Boolean = True, _
    Optional ByVal datCalc As Date) As String
'功能:根据出生日期重新计算病人的年龄,重设年龄单位
'参数:blnSetControl是否设置年龄单位控件
'     datCalc-指定计算日期,未指定时按系统时间计算
'返回:年龄,年龄单位
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    Else
        strSQL = "Select Zl_Age_Calc([1],[2],[3]) old From Dual"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, DateBir, datCalc)
    If blnSetControl = False Then
        ReCalcOld = Trim(nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*岁" Or rsTmp!old Like "*月" Or rsTmp!old Like "*天" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call cbo.Locate(cbo年龄单位, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo年龄单位.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo年龄单位.ListIndex = 0
            Else
                cbo年龄单位.ListIndex = -1
            End If
        End If
    End If
    If cbo年龄单位.ListIndex = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReCalcBirth(ByVal strOld As String, ByVal str年龄单位 As String) As String
'功能:根据年龄和年龄单位估算病人的出生日期,年龄单位为岁时,出年月日假定为1月1号,年龄单位为月时,出生日期假定为1号
'返回:出生日期
    Dim strTmp As String, strFormat As String, lngDays As Long
    
    strTmp = "____-__-__"
    If str年龄单位 = "" Then
        strFormat = "YYYY-MM-DD"
        If strOld Like "*岁*月" Or strOld Like "*岁*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "岁") + 1))
        ElseIf strOld Like "*月*天" Or strOld Like "*个月*天" Then
            lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "月") + 1))
        ElseIf strOld Like "*岁" Or IsNumeric(strOld) Then
            strFormat = "YYYY-01-01"
            lngDays = 365 * Val(strOld)
        ElseIf strOld Like "*月" Or strOld Like "*个月" Then
            strFormat = "YYYY-MM-01"
            lngDays = 30 * Val(strOld)
        ElseIf strOld Like "*天" Then
            lngDays = Val(strOld)
        End If
        If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, zlDatabase.Currentdate), strFormat)
    ElseIf strOld <> "" Then
        Select Case str年龄单位
            Case "岁"
                If Val(strOld) > 200 Then lngDays = -1
            Case "月"
                If Val(strOld) > 2400 Then lngDays = -1
            Case "天"
                If Val(strOld) > 73000 Then lngDays = -1
        End Select
        
        If lngDays = 0 Then
            strTmp = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
            strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, zlDatabase.Currentdate), "YYYY-MM-DD")
            
            If str年龄单位 = "岁" Then
                strTmp = Format(strTmp, "YYYY-01-01")
            ElseIf str年龄单位 = "月" Then
                strTmp = Format(strTmp, "YYYY-MM-01")
            End If
        End If
    End If
    ReCalcBirth = strTmp
End Function

Public Function CheckOldData(ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox) As Boolean
'功能：检查年龄输入值的有效性
'返回：
    If Not IsNumeric(txt年龄.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo年龄单位.Text
        Case "岁"
            If Val(txt年龄.Text) > 200 Then
                MsgBox "年龄不能大于200岁!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "月"
            If Val(txt年龄.Text) > 2400 Then
                MsgBox "年龄不能大于2400月!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "天"
            If Val(txt年龄.Text) > 73000 Then
                MsgBox "年龄不能大于73000天!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Public Function GetOldAcademic(ByVal DateBir As Date, ByVal str年龄单位 As String) As Long
'功能：根据当前的出生日期和年龄单位，计算理论上的年龄值
'返回：年龄
    Dim DatCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" 岁月天", str年龄单位) < 2 Then Exit Function
    
    DatCur = zlDatabase.Currentdate
    
    strInterval = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
    lngOld = DateDiff(strInterval, DateBir, DatCur)
    If DateAdd(strInterval, lngOld, DateBir) > DatCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function

Public Sub LoadOldData(ByVal strOld As String, ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Call zlControl.LoadOldData(strOld, txt年龄, cbo年龄单位)
End Sub
Public Function zlGetFeeFields(Optional strTableName As String = "门诊费用记录", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定表的值
    '入参：strTableName:如:门诊费用记录;住院费用记录;....
    '      blnReadDatabase-从数据库中读取
    '出参：
    '返回：字段集
    '编制：刘兴洪
    '日期：2010-03-10 10:41:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo Errhand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "门诊费用记录"
        zlGetFeeFields = "" & _
        "Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, " & _
        "姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, " & _
        "加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, " & _
        "发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, " & _
        "保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
        Exit Function
    Case "住院费用记录"
        zlGetFeeFields = "" & _
         " Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, " & _
         " 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, " & _
         " 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, " & _
         " 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, " & _
         " 结帐id , 结帐金额, 保险大类ID, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
         Exit Function
    Case "病人结帐记录"
        zlGetFeeFields = "Id, No, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期, 备注"
        Exit Function
    Case "病人预交记录"
        zlGetFeeFields = "" & _
        " Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, " & _
        " 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补,预交类别,卡类别ID,结算卡序号,卡号,交易流水号,交易说明,合作单位,结算序号,校对标志"
        Exit Function
    Case "人员表"
        zlGetFeeFields = "" & _
        "Id, 编号, 姓名, 简码, 身份证号, 出生日期, 性别, 民族, 工作日期, 办公室电话, 电子邮件, 执业类别, 执业范围, " & _
        "管理职务, 专业技术职务, 聘任技术职务, 学历, 所学专业, 留学时间, 留学渠道, 接受培训, 科研课题, 个人简介, 建档时间, " & _
        "撤档时间, 撤档原因, 别名, 站点"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取列信息", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & nvl(!Column_Name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
Errhand:
    zlGetFeeFields = "*"
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "门诊费用记录", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张数据表中的字段.类似于Select Id,....
    '入参：bytHistory-0-不包含历史数据,1-仅包含历史数据,2-两都都包含( select * from tablename Union select * from Htablename)
    '      strWhere-条件
    '      blnSubTable-是否子表
    '      strAliasName-别名
    '出参：
    '返回：select ID ... From tableName Union ALL
    '编制：刘兴洪
    '日期：2010-03-10 11:19:11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '无
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '仅历史
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '两者都包含
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function
Public Function GetServiceDept(str收费细目IDs As String) As ADODB.Recordset
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    If InStr(1, str收费细目IDs, ",") = 0 Then
        strSQL = "" & _
        "   Select  /*+ rule */ Distinct   收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A " & _
        "   Where   A.收费细目ID  =[2] "
    Else
        strSQL = "" & _
        "   Select  /*+ rule */ Distinct   收费细目ID,Nvl(开单科室ID,0) as 开单科室ID,执行科室id " & _
        "   From 收费执行科室 A," & _
        "          (Select Column_Value From Table(Cast(f_num2list([1]) As Zltools.t_Numlist ))) J " & _
        "   Where   A.收费细目ID  = j.Column_Value"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取执行科室信息", Replace(str收费细目IDs, "'", ""), Val(str收费细目IDs))
    If Not rsTmp.EOF Then Set GetServiceDept = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlComboxLoadFromSQL(ByVal strSQL As String, cboControl As Variant, Optional ByVal blnID As Boolean = False) As Boolean
'本函数的功能是从数据库中读出列表值并装到下拉框中
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cmbArray As Variant
    
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取Cbo数据")
    '下拉框数组
    If IsArray(cboControl) Then
        cmbArray = cboControl
    Else
        '强行组成一个数组
        cmbArray = Array(cboControl)
    End If
    
    For intCount = LBound(cmbArray) To UBound(cmbArray)
        cmbArray(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("编码")) Then
                cmbArray(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
            Else
                cmbArray(intCount).AddItem rsTemp("编码") & "." & rsTemp("名称")
            End If
            If blnID = True Then cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("缺省标志") = 1 Then
                cmbArray(intCount).ListIndex = cmbArray(intCount).NewIndex
                cmbArray(intCount).ItemData(cmbArray(intCount).NewIndex) = 1
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True Then cmbArray(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlComboxLoadFromSQL = False
End Function

Public Function zlAddComboItem(cboControl As Control, strItem As String, Optional ByVal cboType As Integer = 1, Optional ByVal cboItemData As Long) As Boolean
    '参数cboType  = 1时表示下拉框由数字打头，
    '             = 2时表示全是文字
    Dim varTemp As Variant
    Dim strTemp As String
    
    '该项在列表框中
    If IsNull(strItem) Or Trim(strItem) = "" Then Exit Function
    For varTemp = 0 To cboControl.ListCount - 1
        If cboType = 1 Then
            strTemp = Mid(cboControl.List(varTemp), InStr(cboControl.List(varTemp), ".") + 1)
            If strItem = strTemp Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        ElseIf cboType = 2 Then
            If strItem = cboControl.List(varTemp) Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        Else
            If cboItemData = cboControl.ItemData(varTemp) Then
                cboControl.ListIndex = varTemp
                Exit Function
            End If
        End If
    Next
    
    If cboType = 1 Then
        cboControl.AddItem strItem
        cboControl.ListIndex = cboControl.NewIndex
    ElseIf cboType = 2 Then
        cboControl.AddItem strItem
        cboControl.ListIndex = cboControl.NewIndex
    End If
End Function
Public Function zlCboFindItem(ByVal cboObj As Object, ByVal lngFindID As Long, _
    Optional strItem As String = "", Optional blnOnlyFind As Boolean = True, Optional blnFindLocal As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：对Combox的ItemData数据进行定位
    '入参：cboObj-Combox对象
    '         lngFindID-需要查找的ID
    '         strItem-需要查找的或增加的子项(当blnOnlyFind=false)时
    '         blnOnlyFind-是否查找.
    '        blnFindLocal-找到后,定位上
    '出参：
    '返回：找到,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-04-06 17:28:17
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngLocate As Long
    zlCboFindItem = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If cboObj.ItemData(lngLocate) = lngFindID Then
            If blnFindLocal Then cboObj.ListIndex = lngLocate
            zlCboFindItem = True
            Exit Function
        End If
    Next
    If blnOnlyFind Then Exit Function
    cboObj.AddItem strItem
    cboObj.ItemData(cboObj.NewIndex) = lngFindID
    If blnFindLocal Then cboObj.ListIndex = cboObj.NewIndex
    zlCboFindItem = True
End Function
Public Function zlPatiCardCheck(ByVal byt调用场合 As Byte, lng病人ID As Long, str卡号 As String, byt刷卡方式 As Byte) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查病人刷卡方式
    '入参：byt调用场合: 1-挂号;2-收费
    '         lng病人ID:病人ID(未建档的,传入零)
    '         str卡号;未刷卡时,为空
    '         byt刷卡方式: 1-普能刷卡;2-医保刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-04-27 16:09:08
    '说明：一汽集团的离休病人，使用的医保卡同时也是就诊卡；医院要求必须以医保方式进行
    '          身份验证挂号、收费，而不能以自费方式直接刷卡进行；因此要求在挂号、收费时，离休病人刷卡后如果不是以医保身份验证方式刷的卡，
    '          而是直接刷的卡，就提示并不允许继续。
    '问题:29283
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = " Select Zl_Paticardcheck([1],[2],[3],[4]) as 提示信息 From Dual "
    ' Zl_Paticardcheck
    '  调用场合_IN NUMBER ,
    '  病人id_In Number,
    '  卡号_In   Varchar2,
    '  刷卡方式_In Number:=1
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病人刷卡方式是否合法", byt调用场合, lng病人ID, str卡号, byt刷卡方式)
    strSQL = nvl(rsTemp!提示信息)
    If strSQL <> "" Then
        MsgBox strSQL, vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    zlPatiCardCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlPatiIS病案已编目(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional blnMsgbox As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张单据的实收金额合计,或一张记帐表中指定病人的实收金额合计
    '返回：已编目,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-12 11:26:28
    '说明：28725
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    strSQL = "Select nvl(B.姓名,A.姓名) As 姓名 From 病案主页 A,病人信息 B where a.病人id=b.病人id and  A.病人id=[1] and a.主页id=[2] and 编目日期 IS NOT NULL"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查病案是否已经编码", lng病人ID, lng主页ID)
    If rsTemp.EOF Then
        zlPatiIS病案已编目 = False
    Else
        zlPatiIS病案已编目 = True
        If blnMsgbox Then
                MsgBox "病人『" & nvl(rsTemp!姓名) & " 』已经编目,不允许进行记帐或销帐操作!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
        End If
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlCheckIsMzToZY(ByVal strNos As String, ByVal int性质 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否门诊转住院费用是否已经审核
    '入参:strNos-单据号(用逗号分离)
    '        int性质-收费单;2-记帐单
    '返回:如果存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-03-02 16:18:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNO As String
    On Error GoTo errHandle
    strNO = Replace(strNos, "'", "")
     strSQL = "" & _
     "  Select /*+ rule */   1 From 门诊费用记录 A,费用审核记录 B,Table(f_Str2list([1])) J" & _
     "  Where  A.NO=J.Column_Value and A.记录性质=[2] and A.ID=B.费用ID  " & _
     "                  And  B.性质  =1 and Rownum=1"
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否门诊费用转住院费用", strNO, int性质)
    zlCheckIsMzToZY = Not rsTemp.EOF
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zl_GetInvoicePreperty(ByVal lngModule As Long, _
    ByVal int票据 As Integer, Optional str使用类别 As String) As Ty_FactProperty
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票格式
    '入参:int票据:1 - 收费收据, 2 - 预交收据, 3 - 结帐收据, 4 - 挂号收据, 5 - 就诊卡, 12 - 预交红票
    '返回:发票的相关数据
    '编制:刘兴洪
    '日期:2011-07-19 16:43:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Ty_Fact As Ty_FactProperty, strFactType As String, varData As Variant, varTemp As Variant
    Dim strShareTypeUseID As String, lng共用票据 As Long, lng使用票据 As Long
    Dim strFactTypeFormat As String, strFacePrintMode As String
    Dim intPrintMode As Long, intPrintMode1 As Long, lng领用ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, lngFormat As Long, lngFormat1 As Long
    
    strFactType = Switch(int票据 = 1, "共用收费票据批次", int票据 = 2, "共用预交票据批次", int票据 = 12, "共用预交票据批次", int票据 = 3, "共用结帐票据批次", int票据 = 4, "共用挂号票据批次", int票据 = 5, "共用医疗卡批次", True, "")
    strFactTypeFormat = Switch(int票据 = 1, "收费发票格式", int票据 = 2, "预交发票格式", int票据 = 12, "退款发票格式", int票据 = 3, "结帐发票格式", int票据 = 4, "挂号发票格式", int票据 = 5, "医疗卡发票格式", True, "")
    strFacePrintMode = Switch(int票据 = 1, "收费发票打印方式", int票据 = 2, "预交发票打印方式", int票据 = 12, "预交退款打印方式", int票据 = 3, "病人结帐打印", int票据 = 4, "挂号发票打印方式", int票据 = 5, "医疗卡发票打印方式", True, "")
    
    If strFactType = "" Then Exit Function
    '78751:李南春,2014/10/20,增加预交票据打印格式
    Ty_Fact.strUseType = str使用类别
    '初始发票格式
'    If int票据 = 2 Then
'        '预交暂无格
'        Ty_Fact.intInvoiceFormat = 0
'    Else
        strFactTypeFormat = Trim(zlDatabase.GetPara(strFactTypeFormat, glngSys, lngModule, ""))
        '格式:使用类别1,格式1|使用类别2,格式2...
        varData = Split(strFactTypeFormat, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",", ",")
            lngFormat = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then lngFormat1 = lngFormat
            If Trim(varTemp(0)) = str使用类别 And lngFormat <> 0 Then
                Ty_Fact.intInvoiceFormat = lngFormat: Exit For
            End If
        Next
        If Ty_Fact.intInvoiceFormat = 0 And lngFormat1 <> 0 Then Ty_Fact.intInvoiceFormat = lngFormat
'    End If
    
    '打印方式(0-不打印;1-自动打印;2-提示打印)
    '问题50656
'    If int票据 = 2 Then
'        '预交暂为自动打印
'        Ty_Fact.intInvoicePrint = 1
'    Else
        '因为Getpara就缓存了的,所以不用先用变量进行记录
        strFacePrintMode = Trim(zlDatabase.GetPara(strFacePrintMode, glngSys, lngModule, ""))
        Ty_Fact.intInvoicePrint = -1
        '格式:使用类别1,打印方式1|使用类别2,打印方式2...
        varData = Split(strFacePrintMode, "|")
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,", ",")
            intPrintMode = Val(varTemp(1))
            If Trim(varTemp(0)) = "" Then intPrintMode1 = intPrintMode
            If Trim(varTemp(0)) = str使用类别 Then
                Ty_Fact.intInvoicePrint = intPrintMode: Exit For
            End If
        Next
        If Ty_Fact.intInvoicePrint < 0 Then Ty_Fact.intInvoicePrint = intPrintMode1
'    End If
    '共享批次
    
    '格式:领用ID1,使用类别1|....
    strShareTypeUseID = Trim(zlDatabase.GetPara(strFactType, glngSys, lngModule, "0"))
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng领用ID = Val(varTemp(0))
        If int票据 = 2 Or int票据 = 12 Or int票据 = 5 Then
            If Val(varTemp(1)) = 0 Then lng共用票据 = lng领用ID    '共用的.
            If Val(varTemp(1)) = Val(str使用类别) And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        Else
            If Trim(varTemp(1)) = "" Then lng共用票据 = lng领用ID    '共用的.
            If Trim(varTemp(1)) = str使用类别 And lng领用ID <> 0 Then
                lng使用票据 = lng领用ID
            End If
        End If
    Next
    
    On Error GoTo errHandle
    '优先顺序
    '1.先使用
    '2.使用类别不区分的
    '3.具体使用类别的
    strSQL = _
    "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
    "From 票据领用记录" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And 剩余数量 > 0   " & vbNewLine & _
    "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 Desc, 开始号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", lng共用票据, lng使用票据)
    If rsTemp.EOF = False Then
        Ty_Fact.lngShareUseID = Val(nvl(rsTemp!ID)) '共用的领用ID
    End If
    zl_GetInvoicePreperty = Ty_Fact
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  Zl_Billclass([1],[2],[3]) as 使用类别 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取票据使用类别", lng病人ID, lng主页ID, intInsure)
    zl_GetInvoiceUserType = nvl(rsTemp!使用类别)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zl_GetInvoiceShareID(ByVal lngModule As Long, Optional str使用类别 As String = "") As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的共享票据ID
    '返回:共享的领用ID
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, strShareTypeUseID As String
    Dim lng领用ID As Long '共享的领用ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng共用票据 As Long, lng使用票据 As Long
    
    '因为Getpara就缓存了的,所以不用先用变量进行记录
    If lngModule = 1137 Then
        strShareTypeUseID = Trim(zlDatabase.GetPara("共用结帐票据批次", glngSys, lngModule, "0"))
        '格式:领用ID1,使用类别1|....
    Else
        strShareTypeUseID = Trim(zlDatabase.GetPara("共用收费票据批次", glngSys, lngModule, "0"))
        '格式:领用ID1,使用类别1|....
    End If
    
    varData = Split(strShareTypeUseID, "|")
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ",", ",")
        lng领用ID = Val(varTemp(0))
        If Trim(varTemp(1)) = "" Then lng共用票据 = lng领用ID    '共用的.
        If Trim(varTemp(1)) = str使用类别 And lng领用ID <> 0 Then
            lng使用票据 = lng领用ID
        End If
    Next
    On Error GoTo errHandle
    '优先顺序
    '1.先使用
    '2.使用类别不区分的
    '3.具体使用类别的
    strSQL = _
    "Select ID, 前缀文本, 开始号码, 终止号码, 剩余数量, 登记时间, 使用时间" & vbNewLine & _
    "From 票据领用记录" & vbNewLine & _
    "Where (ID =[1] or ID =[2]) And 剩余数量 > 0   " & vbNewLine & _
    "Order By Nvl(使用时间, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc,使用类别 Desc, 开始号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "可用票据批次", lng共用票据, lng使用票据)
    If rsTemp.EOF = False Then
        zl_GetInvoiceShareID = Val(nvl(rsTemp!ID))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查票据是否启用了使用类别的", int票种)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetSpecialItemFee(strClass As String, Optional ByVal strPriceGrade As String, Optional ByVal lng收费细目ID As Long) As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:产生工本费、就诊卡等填写住院费用记录时的必须信息(收费类别,收费细目ID,计算单位,收入项目ID,收入项目,收据费目,原价,现价,是否变价,科室标志)
    '入参:
    '   strClass=工本费、就诊卡、病历费
    '   strPriceGrade 普通价格等级
    '返回:指定的数据类别的费用集
    '编制:刘兴洪
    '日期:2011-07-07 02:17:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strWherePriceGrade As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (b.价格等级 = [2]" & vbNewLine & _
            "          Or (b.价格等级 Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From 收费价目" & vbNewLine & _
            "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    
    If lng收费细目ID = 0 Then
        strSQL = _
            "Select a.类别 As 收费类别, a.Id As 收费细目id, a.计算单位, c.Id As 收入项目id, Nvl(a.屏蔽费别, 0) As 屏蔽费别, c.名称 As 收入项目, c.收据费目, b.原价, b.现价," & vbNewLine & _
            "       Nvl(b.缺省价格, 0) 缺省价格, Nvl(a.是否变价, 0) As 是否变价, Nvl(a.执行科室, 0) As 科室标志" & vbNewLine & _
            "From 收费项目目录 A, 收费价目 B, 收入项目 C, 收费特定项目 D" & vbNewLine & _
            "Where b.收费细目id = a.Id And b.收入项目id = c.Id And d.收费细目id = a.Id And d.特定项目 = [1]" & vbNewLine & _
            "      And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    Else
        strSQL = _
            "Select a.类别 As 收费类别, a.Id As 收费细目id, a.计算单位, c.Id As 收入项目id, Nvl(a.屏蔽费别, 0) As 屏蔽费别, c.名称 As 收入项目, c.收据费目, b.原价, b.现价," & vbNewLine & _
            "       Nvl(b.缺省价格, 0) 缺省价格, Nvl(a.是否变价, 0) As 是否变价, Nvl(a.执行科室, 0) As 科室标志" & vbNewLine & _
            "From 收费项目目录 A, 收费价目 B, 收入项目 C " & vbNewLine & _
            "Where b.收费细目id = a.Id And b.收入项目id = c.Id And A.ID = [3]" & vbNewLine & _
            "      And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "      And Sysdate Between b.执行日期 And Nvl(b.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            strWherePriceGrade
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取特定项目的费用集", strClass, strPriceGrade, lng收费细目ID)
    If Not rsTmp.EOF Then Set zlGetSpecialItemFee = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            zlGetUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSQL = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                zlGetUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                zlGetUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            zlGetUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetSaveCardFeeSQL(ByVal lngCardTypeID As Long, bytStyle As Byte, strNO As String, lng病人ID As Long, lng主页ID As Long, _
        lng病人病区ID As Long, lng病人科室ID As Long, lng标识号 As Long, str费别 As String, _
        str原卡号 As String, str姓名 As String, str性别 As String, str年龄 As String, str卡号 As String, str密码 As String, _
        str变动原因 As String, cur应收金额 As Double, cur实收金额 As Double, str结算方式 As String, dt发卡时间 As Date, lng领用ID As Long, rsMoney As ADODB.Recordset, _
        ByVal strICCard As String, _
        Optional lng刷卡类别ID As Long, Optional bln消费卡 As Boolean, Optional str刷卡卡号 As String, Optional lng结帐ID As Long, Optional str摘要 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:产生一条医疗卡费用记录SQL语句
    '入参:bytStyle=0-发卡,1-补卡,2-换卡
    '       cur金额=就诊卡金额
    '       str结算方式=如果为空,表示记帐,不收现金
    '       rsMoney:包括就诊卡收费信息的记录集
    '       str原卡号=仅换卡时用
    '       lng领用ID=当前可用的就诊卡领用ID
    '       str密码-必须带用oracle的单引号或为空
    '       strICCard=IC卡号,通过读IC卡方式发卡时,同时填写病人信息的IC卡字段
    '返回:医疗卡费用记录SQL语句
    '编制:刘兴洪
    '日期:2011-07-08 01:08:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngUnitID As Long, strSQL As String
    
    '0-不明确,1-病人科室,2-病人病区,3-操作员科室,4-指定科室,5-院外执行(预留,程序暂未用),6-开单人科室
    Select Case rsMoney!科室标志
        Case 4 '指定科室
            lngUnitID = zlGetUnitID(rsMoney!科室标志, rsMoney!收费细目ID)
        Case 1, 2 '病人科室
            If lng病人科室ID <> 0 Then
                lngUnitID = lng病人科室ID
            Else
                lngUnitID = UserInfo.部门ID
            End If
        Case 0, 3, 5, 6
            lngUnitID = UserInfo.部门ID
    End Select
  'Zl_医疗卡记录_Insert
    strSQL = "Zl_医疗卡记录_Insert("
    '  --参数：发卡类型=0-发卡,1-补卡,2-换卡(相当于重打)
    '  --      换卡时,单据号_IN传入的是原发/补卡的单据号。
    '  --      补卡/换卡后,再换卡时是以最后一次卡号为准。
    '  发卡类型_In   Number,
    strSQL = strSQL & "" & bytStyle & ","
    '  单据号_In     住院费用记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  病人id_In     住院费用记录.病人id%Type,
    strSQL = strSQL & "'" & lng病人ID & "',"
    '  主页id_In     住院费用记录.主页id%Type,
    strSQL = strSQL & "" & IIf(lng主页ID = 0, "NULL", lng主页ID) & ","
    '  标识号_In     住院费用记录.标识号%Type,
    strSQL = strSQL & "" & IIf(lng标识号 = 0, "NULL", lng标识号) & ","
    '  费别_In       住院费用记录.费别%Type,
    strSQL = strSQL & "'" & str费别 & "',"
    '  卡类别id_In   医疗卡类别.ID%Type,
    strSQL = strSQL & "" & lngCardTypeID & ","
    '  原卡号_In     病人医疗卡信息.卡号%Type,
    strSQL = strSQL & IIf(str原卡号 = "", "NULL", "'" & str原卡号 & "'") & ","
    '  医疗卡号_In   病人医疗卡信息.卡号%Type,
    strSQL = strSQL & IIf(str卡号 = "", "NULL", "'" & str卡号 & "'") & ","
    '  变动原因_In   病人医疗卡变动.变动原因%Type,
    strSQL = strSQL & IIf(str变动原因 = "", "NULL", "'" & str变动原因 & "'") & ","
    '  密码_In       病人信息.卡验证码%Type,
    strSQL = strSQL & IIf(str密码 = "", "NULL", "'" & str密码 & "'") & ","
    '  姓名_In       住院费用记录.姓名%Type,
    strSQL = strSQL & IIf(str姓名 = "", "NULL", "'" & str姓名 & "'") & ","
    '  性别_In       住院费用记录.性别%Type,
    strSQL = strSQL & IIf(str性别 = "", "NULL", "'" & str性别 & "'") & ","
    '  年龄_In       住院费用记录.年龄%Type,
    strSQL = strSQL & IIf(str年龄 = "", "NULL", "'" & str年龄 & "'") & ","
    '  病人病区id_In 住院费用记录.病人病区id%Type,
    strSQL = strSQL & "" & lng病人病区ID & ","
    '  病人科室id_In 住院费用记录.病人科室id%Type,
    strSQL = strSQL & "" & lng病人科室ID & ","
    '  收费细目id_In 住院费用记录.收费细目id%Type,
    strSQL = strSQL & "" & rsMoney!收费细目ID & ","
    '  收费类别_In   住院费用记录.收费类别%Type,
    strSQL = strSQL & "'" & rsMoney!收费类别 & "',"
    '  计算单位_In   住院费用记录.计算单位%Type,
    strSQL = strSQL & "'" & nvl(rsMoney!计算单位) & "',"
    '  收入项目id_In 住院费用记录.收入项目id%Type,
    strSQL = strSQL & "" & rsMoney!收入项目ID & ","
    '  收据费目_In   住院费用记录.收据费目%Type,
    strSQL = strSQL & "'" & nvl(rsMoney!收据费目) & "',"
    '  标准单价_In   住院费用记录.标准单价%Type,
    strSQL = strSQL & "" & cur应收金额 & ","
    '  执行部门id_In 住院费用记录.执行部门id%Type,
    strSQL = strSQL & "" & lngUnitID & ","
    '  开单部门id_In 住院费用记录.开单部门id%Type,
    strSQL = strSQL & "" & UserInfo.部门ID & ","
    '  操作员编号_In 住院费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  加班标志_In   住院费用记录.加班标志%Type,
    strSQL = strSQL & "" & IIf(OverTime(dt发卡时间), "1", "0") & ","
    '  发卡时间_In   住院费用记录.登记时间%Type,
    strSQL = strSQL & "To_Date('" & Format(dt发卡时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "" & IIf(lng领用ID = 0, "NULL", lng领用ID) & ","
    '  Ic卡号_In     病人信息.Ic卡号%Type := Null,
    strSQL = strSQL & "'" & strICCard & "',"
    '  应收金额_In   住院费用记录.应收金额%Type,
    strSQL = strSQL & "" & cur应收金额 & ","
    '  实收金额_In   住院费用记录.实收金额%Type,
    strSQL = strSQL & "" & cur实收金额 & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "" & IIf(str结算方式 = "", "NULL", "'" & str结算方式 & "'") & ","
    '  刷卡类别id_In 病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & IIf(lng刷卡类别ID = 0, "NULL", lng刷卡类别ID) & ","
    '  消费卡_In     Integer := 0,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  刷卡卡号_In   病人医疗卡信息.卡号%Type
    strSQL = strSQL & "'" & str刷卡卡号 & "',"
    '  结帐ID_IN
    strSQL = strSQL & "" & IIf(lng结帐ID = 0, "NULL", lng结帐ID) & ","
    '  交易流水号_In
    strSQL = strSQL & "NULL,"
    '  交易说明_In
    strSQL = strSQL & "NULL,"
    '  合作单位_In
    strSQL = strSQL & "NULL,"
    '  摘要_In   住院费用记录.摘要%Type,
    strSQL = strSQL & "" & IIf(str摘要 = "", "NULL", "'" & str摘要 & "'") & ")"
    
    zlGetSaveCardFeeSQL = strSQL
End Function
Public Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSQL = strSQL & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSQL = strSQL & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSQL = strSQL & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSQL = strSQL & "" & IIf(bln预交, 1, 0) & ","
    '退费标志 :1-退费;0-付费
    strSQL = strSQL & "0,"
    '校对标志
    strSQL = strSQL & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlAddArray cllPro, strSQL
End Function

Public Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSQL As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSQL = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSQL = strSQL & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSQL = strSQL & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSQL = strSQL & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSQL = strSQL & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
                    zlAddArray cllPro, strSQL
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSQL = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSQL = strSQL & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSQL = strSQL & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSQL = strSQL & IIf(bln预交款, "1", "0") & ")"
        zlAddArray cllPro, strSQL
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsNotSucceedPrintBill(ByVal BytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否已经正常打印
    '入参:bytType-1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '       strNos-本次打印票据的单据,用逗号分离
    '出参:strOutValidNos-打印失败的单据号
    '返回:存在不存功票据的打印,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-16 18:06:01
    '问题:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSQL As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    '应取最后一次打印的最大号码
    strSQL = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From 票据使用明细 A,票据打印内容 B,Table( f_Str2list([2])) J" & _
        " Where A.打印ID =b.ID And B.数据性质=[1] And B.No=J.Column_value "
        'And A.票种=b.数据性质:有可能使用的是其他票据:比如挂号使用门诊收费票据
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查票据是否打印", BytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlIsCheckMedicinePayMode(ByVal str医疗付款名称 As String, _
    Optional ByRef bln医保 As Boolean, Optional ByRef bln公费 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查医疗付款方式是否公费或医保
    '入参:str医疗付款名称-医疗付款名称
    '出参:bln医保-true,表示医保
    '        bln公费-true,表示是公费
    '返回:是医保或公费医疗,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-17 16:25:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "": bln医保 = False: bln公费 = False
    On Error GoTo errHandle
    If grs医疗付款方式 Is Nothing Then
        strSQL = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    ElseIf grs医疗付款方式.State <> 1 Then
        strSQL = "Select 编码,名称,简码,缺省标志,是否医保,是否公费 From 医疗付款方式"
    End If
    If strSQL <> "" Then
        Set grs医疗付款方式 = zlDatabase.OpenSQLRecord(strSQL, "获取医疗付款方式")
    End If
    grs医疗付款方式.Find "名称='" & str医疗付款名称 & "'", , adSearchForward, 1
    If grs医疗付款方式.EOF Then Exit Function
    bln医保 = Val(nvl(grs医疗付款方式!是否医保)) = 1
    bln公费 = Val(nvl(grs医疗付款方式!是否公费)) = 1
    zlIsCheckMedicinePayMode = bln医保 Or bln公费
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlLeftPad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按指定长度填制空格
    '返回:返回字串
    '编制:刘兴洪
    '日期:2012-02-22 17:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlLeftPad = zlStr.LPAD(strCode, lngLen, strChar, True)
End Function

Private Function zlSubstr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定字串的值,字串中可以包含汉字
    '入参:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '返回:子串
    '编制:刘兴洪
    '日期:2012-02-22 18:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlSubstr = zlStr.SubB(strInfor, lngStart, lngLen)
End Function
Public Function zlCheckIsExistsApplied(ByVal strNO As String, ByVal str序号 As String, _
    ByRef str费用IDs As String, Optional ByRef str申请人s As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查销帐的数据中是否存在销帐申请
    '入参:strNo-单据号
    '       str序号-本次销帐的序号(为空为所有)
    '出参:str费用IDs-申请的费用ID
    '返回:存在销帐申请,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-20 15:51:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct A.ID,B.申请人 " & _
    "   From 住院费用记录 A,病人费用销帐 B  " & _
    "   Where A.ID=B.费用ID and A.NO=[1] and A.记录性质=2 And nvl(B.状态,0)=0 " & IIf(str序号 <> "", " and Instr([2],','||序号||',')>0 ", "")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取申请状态", strNO, "," & str序号 & ",")
    If rsTemp.EOF Then
        rsTemp.Close: Set rsTemp = Nothing: Exit Function
    End If
    str申请人s = "": str费用IDs = ""
    With rsTemp
        Do While Not .EOF
            str费用IDs = str费用IDs & "," & Val(nvl(rsTemp!ID))
            If InStr(1, str申请人s & vbCrLf, vbCrLf & nvl(rsTemp!申请人) & vbCrLf) = 0 Then
                str申请人s = str申请人s & vbCrLf & nvl(rsTemp!申请人)
            End If
            .MoveNext
        Loop
    End With
    If str费用IDs <> "" Then str费用IDs = Mid(str费用IDs, 2)
    zlCheckIsExistsApplied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlExecuteChargeRollingCurtain(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行收费轧帐管理
    '入参:frmMain-调用的主窗体
    '编制:刘兴洪
    '日期:2013-10-16 10:15:22
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCargeBill As Object
    Dim strCommon As String
    Dim intAtom As Integer
    Err = 0: On Error Resume Next
    Set objCargeBill = CreateObject("zL9CashBill.clsChargeBill")
    If Err <> 0 Then
        Set objCargeBill = Nothing
        MsgBox "创建轧帐部件(zl9CashBill)失败,估计部件丢失了,请与管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    '6.1.7.1.    InitOracle:初始化连接
    '入参:
      '     cnMain-数据库连接
      '   strDBUser-数据库所有者
      '     lngSys-系统号
      
     '为通讯原子赋值
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    '加入通讯原子
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    If objCargeBill.InitOracle(gcnOracle, gstrDBUser, glngSys) = False Then
        Set objCargeBill = Nothing
        Exit Sub
    End If
    Call GlobalDeleteAtom(intAtom)
    'ChargeRollingCurtain(ByVal frmMain As Object)
    If objCargeBill.ChargeRollingCurtain(frmMain) = False Then
        Set objCargeBill = Nothing
        Exit Sub
    End If
    Set objCargeBill = Nothing
End Sub
 
Public Function zlIsPrintBill(ByVal lng病人ID As Long, _
    ByVal lng诊疗ID As Long, int性质 As Integer, Optional strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否打印了票据的
    '入参:lng病人ID-病人ID-指定的病人ID
    '       lng诊疗ID-挂号ID(0-为所有的)
    '       int性质-1-收费;3-结帐;4-挂号
    '       strNo-性质=1或4时,为挂号单号
    '出参:
    '返回:打印了票据的返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-06 17:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
        
    If int性质 = 4 Then
        strSQL = "" & _
        "   Select  1 " & _
        "   From 病人挂号记录 A,票据打印内容 B" & _
        "   Where a.NO=b.NO and B.数据性质=4 and A.病人ID " & IIf(lng诊疗ID <> 0, "+0", "") & " =[1] " & IIf(lng诊疗ID <> 0, " And A.ID=[2]", "") & _
        "            And  Exists (Select 1 From 票据使用明细 M Where b.Id = m.打印id And 性质 = 1) And Rownum < 2  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否打印票据", lng病人ID, lng诊疗ID)
        zlIsPrintBill = Not rsTemp.EOF
        Exit Function
    End If
    
    If int性质 = 1 Then '收费
        If strNO = "" And lng诊疗ID <> 0 Then
            strSQL = "Select NO From 病人挂号记录 where ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取挂号单号", lng诊疗ID)
            If rsTemp.EOF Then Exit Function
            strNO = nvl(rsTemp!NO)
        End If
        strSQL = "" & _
        "   Select  1 " & _
        "   From 门诊费用记录 A,票据打印内容 B" & _
        "   Where  a.NO=b.NO and B.数据性质=1 and a.记录性质=1 and A.病人ID=[1] " & _
        IIf(strNO = "", "", "      And Exists(Select 1 From 病人医嘱记录 M Where 挂号单=[3]  And M.ID=A.医嘱序号) ") & _
        "      And Exists (Select 1 From 票据使用明细 M Where b.Id = m.打印id And 性质 = 1) And Rownum < 2  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否打印票据", lng病人ID, lng诊疗ID, strNO)
        zlIsPrintBill = Not rsTemp.EOF
        Exit Function
    End If
    
    strSQL = "" & _
    "   Select  1 " & _
    "   From 病人结帐记录 A,票据打印内容 B " & _
    "   Where  a.NO=b.NO and B.数据性质=3  and A.病人ID=[1] " & _
    "      And Exists (Select 1 From 票据使用明细 M Where b.Id = m.打印id And 性质 = 1) And Rownum < 2  "
    If lng诊疗ID <> 0 Then
        strSQL = strSQL & vbCrLf & _
        "        AND exists(SELECT 1 From 住院费用记录 WHERE a.id=结帐ID  AND 病人ID+0=[1] AND 主页ID+0=[2])"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否打印票据", lng病人ID, lng诊疗ID)
    zlIsPrintBill = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlExistOperationData(ByVal lng病人ID As Long, ByVal strNO As String, _
    Optional ByVal lng诊疗ID As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否当前病人发生了业务数据
    '入参:lng病人ID-病人ID-指定的病人ID
    '       strNo-挂号单号
    '       lng诊疗ID-挂号ID
    '出参:
    '返回:存在业务数据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-11-06 17:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If strNO <> "" Then
        strSQL = "" & _
        "   Select 1 From 病人医嘱记录 A Where 病人ID+0=[1] And 挂号单=[2]"
    ElseIf lng诊疗ID <> 0 Then
        strSQL = "" & _
        "   Select 1 From 病人医嘱记录 A,病人挂号记录 B Where  A.挂号单=B.NO And B.ID=[3] "
    Else
        strSQL = "" & _
        "   Select 1 From 病人医嘱记录 A Where 病人ID =[1] AND ROWNUM<2"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断门诊是否发生了业务数据", lng病人ID, strNO, lng诊疗ID)
    zlExistOperationData = rsTemp.EOF = False
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet误差费名称() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取误差费名称
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-05 12:03:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 名称 From 结算方式 where 性质=9 And nvl(是否固定,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取误差费名称")
    If Not rsTemp.EOF Then
        zlGet误差费名称 = nvl(rsTemp!名称)
    Else
        zlGet误差费名称 = "误差费"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGet原结帐ID(ByVal lng冲销ID As Long) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据冲销ID获取原结帐ID
    '入参:lng冲销ID-当前冲销ID
    '出参:
    '返回:返回原结帐ID,0-原结帐ID未获取到
    '编制:刘兴洪
    '日期:2014-06-13 17:26:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = " " & _
    "   Select A.结帐id,A.登记时间 " & _
    "   From 门诊费用记录 A, " & _
    "       (   Select Max(NO) as NO,Max(登记时间) as 登记时间   " & _
    "           From  门诊费用记录 Where 结帐ID=[1] ) B " & _
    "   Where a.No = B.NO And Mod(a.记录性质, 10) = 1 And 记录状态 In (1, 3)  " & _
    "         And  a.登记时间<= B.登记时间 " & _
    "   Order by A.登记时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取原结帐ID", lng冲销ID)
    If Not rsTemp.EOF Then
     zlGet原结帐ID = Val(nvl(rsTemp!结帐ID))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlReadBillFormat(ByVal ReportCode As String) As ADODB.Recordset
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定报表的打印格式
    '入参:ReportCode-报表名称
    '返回:报表打印格式的记录集
    '编制:李南春
    '日期:2014-10-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='" & ReportCode & "'  " & _
    "   Order by  序号"
    Set zlReadBillFormat = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get病人信息从表(ByVal lng病人ID As Long, Optional ByVal str信息串 As String = "") As ADODB.Recordset
'功能：
'    获取病人信息从表项
'参数:
    Dim strSQL As String
    Dim intRet As Integer
    
    intRet = UBound(Split(str信息串, ","))
    If intRet = -1 Then '读取病人所有从表信息
        strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID =[1] And 信息值 is Not Null"
    ElseIf intRet = 0 Then '读取指定某个从表信息
        strSQL = "Select 信息名,信息值 From 病人信息从表 Where 病人ID =[1] And 信息名='" & Split(str信息串, ",")(0) & "'" & " And 信息值 is Not Null "
    ElseIf intRet > 0 Then '读取指定的多个从表信息值
        strSQL = "Select 信息名, 信息值" & vbNewLine & _
            "From 病人信息从表" & vbNewLine & _
            "Where 病人id = [1] And" & vbNewLine & _
            "      信息名 In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And 信息值 is Not Null "
    End If
    
    On Error GoTo errH
    Set Get病人信息从表 = zlDatabase.OpenSQLRecord(strSQL, "读取病人从表", lng病人ID, str信息串)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

