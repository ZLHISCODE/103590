VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlUpdateBillExcutePerson(ByVal byt病人来源 As Byte, _
    ByVal int记录性质 As Integer, strNO As String, _
    Optional ByVal blnExcuteSave As Boolean = True, _
    Optional ByRef strOutExcuteSQL As String, _
    Optional ByVal str执行时间 As String, Optional ByVal str执行人 As String, _
    Optional ByVal str序号 As String = "", _
    Optional ByVal str结论 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新单据的执行人
    '入参:byt病人来源,:1-门诊;2-住院;4-体检
    '    int记录性质-1-收费单;2-记帐单
    '    strNo -单据号
    '    str执行时间 -执行时间(未传入时,填写当前时间),格式:yyyy-mm-dd HH:MM:SS
    '    str执行人 -执行人(未传入时,填写为当前操作员)
    '    blnExcuteSave-是否执行保存操作(缺省为true,表示调用该接口直接执行保存;否则返回保存的SQL语句)
    '    str序号-需要更新的相关序号,多个用逗号分离
    '出参:strOutExcuteSQL-返回执行的SQL语句
    '返回:TRUE:填写成功;False:填写失败或读取失败
    '编制:刘兴洪
    '日期:2014-03-26 16:23:16
    '说明:该过程中不处理异常,主要由调用方处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    If str执行时间 = "" Then
        str执行时间 = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    If str执行人 = "" Then str执行人 = UserInfo.姓名
    
    'Zl_病人费用记录_Execute
    strSql = "Zl_病人费用记录_Execute("
    '  No_In         住院费用记录.No%Type,
    strSql = strSql & "'" & strNO & "',"
    '  记录性质_In   住院费用记录.记录性质%Type,
    strSql = strSql & "" & int记录性质 & ","
    '  序号_In       Varchar2,
    strSql = strSql & IIf(str序号 = "", "NULL", "'" & str序号 & "'") & ","
    '  Int来源       Number := 2,
    '--int来源 :1-门诊;2-住院
    strSql = strSql & IIf(byt病人来源 <> 2, 1, 2) & ","
    '  结论_In       住院费用记录.结论%Type,
    strSql = strSql & IIf(str结论 = "", "NULL", "'" & str结论 & "'") & ","
    '  执行人_In     住院费用记录.执行人%Type,
    strSql = strSql & "'" & str执行人 & "',"
    '  执行时间_In   住院费用记录.执行时间%Type,
    strSql = strSql & "to_date('" & str执行时间 & "','yyyy-mm-dd hh24:mi:ss'),"
    '  仅更新费用_In Number:=0
    ' -- 仅更新费用_In:1表示:只更新病人费用记录.执行人,执行时间,结论;0-根据执行情况填写相关表
    strSql = strSql & "1)"
    If blnExcuteSave Then
        gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "执行登记信息填写"
    End If
    strOutExcuteSQL = strSql
    zlUpdateBillExcutePerson = True
End Function
Public Function zlDelBillExcutePerson(ByVal byt病人来源 As Byte, _
    ByVal int记录性质 As Integer, strNO As String, _
    Optional ByVal blnExcuteSave As Boolean = True, _
    Optional ByRef strOutExcuteSQL As String, _
    Optional ByVal str序号 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将单据的执行人,执行时间等更新为空(即取消执行登记)
    '入参:byt病人来源,:1-门诊;2-住院;4-体检
    '    int记录性质-1-收费单;2-记帐单
    '    strNo -单据号
    '    blnExcuteSave-是否执行删除操作(缺省为true,表示调用该接口直接执行删除;否则返回删除的SQL语句
    '    str序号-需要更新的相关序号,多个用逗号分离
    '出参:strOutExcuteSQL-返回执行的SQL语句
    '返回:TRUE:填写成功;False:填写失败或读取失败
    '编制:刘兴洪
    '日期:2014-03-26 16:23:16
    '说明:该过程中不处理异常,主要由调用方处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
 
    ' Zl_病人费用记录_Unexecute
    strSql = "Zl_病人费用记录_Unexecute("
    '  No_In         住院费用记录.No%Type,
    strSql = strSql & "'" & strNO & "',"
    '  记录性质_In   住院费用记录.记录性质%Type,
    strSql = strSql & "" & int记录性质 & ","
    '  序号_In       Varchar2,
    strSql = strSql & IIf(str序号 = "", "NULL", "'" & str序号 & "'") & ","
    '  来源_In       Number := 2,
    strSql = strSql & IIf(byt病人来源 <> 2, 1, 2) & ","
    '  仅更新费用_In Number:=0
    strSql = strSql & "1)"
    If blnExcuteSave Then
        gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "取消执行登记信息"
    End If
    strOutExcuteSQL = strSql
    zlDelBillExcutePerson = True
End Function
Public Function zlDelBalance(strNO As String, _
    Optional ByVal blnExcuteDel As Boolean = True, _
    Optional ByRef strOutExcuteSQL As String, _
    Optional ByVal strOperatorCode As String, _
    Optional ByVal strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对结帐单进行结帐作废操作
    '入参:int记录性质-1-收费单;2-记帐单
    '    strNo -结帐单据号
    '    strOperatorCode-操作员编号(不传入时,以当前操作员为准)
    '    strOperatorName-操作员姓名(不传入时,以当前操作员为准)
    '    blnExcuteSave-是否执行结帐作废操作(缺省为true,表示调用该接口直接执行作废;否则返回作废的SQL语句)
    '出参:strOutExcuteSQL-返回执行的SQL语句
    '返回:TRUE:填写成功;False:填写失败或读取失败
    '编制:刘兴洪
    '日期:2014-03-26 16:23:16
    '说明:该过程中不处理异常,主要由调用方处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    If strOperatorCode = "" Or strOperatorName = "" Then
        strOperatorCode = UserInfo.编号
        strOperatorName = UserInfo.姓名
    End If
    'Zl_病人结帐记录_Delete
    strSql = "Zl_病人结帐记录_Delete("
    '  No_In           病人结帐记录.NO%Type,
    strSql = strSql & "'" & strNO & "',"
    '  操作员编号_In   病人结帐记录.操作员编号%Type,
    strSql = strSql & "'" & strOperatorCode & "',"
    '  操作员姓名_In   病人结帐记录.操作员姓名%Type,
    strSql = strSql & "'" & strOperatorName & "')"
    '  误差金额_In     病人预交记录.冲预交%Type := 0, --医保或预交退现金产生的误差
    '  误差no_In       住院费用记录.NO%Type := Null,
    '  结帐作废结算_In Varchar2 := Null, --结算方式|结算金额|结算号码||......
    '  预交退现金_In   Number := 0 --当预交款退现金时，结算方式及金额通过参数结帐作废结算_In传入
    If blnExcuteDel Then
        gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "费用结算部件:结帐单作废"
    End If
    strOutExcuteSQL = strSql
    zlDelBalance = True
End Function

Public Function zlRegeventMakeupSave(ByVal lng病人ID As Long, _
    ByVal lng执行部门ID As Long, _
    ByVal str操作员编号 As String, ByVal str操作员姓名 As String, _
    ByVal str医生姓名 As String, _
    ByVal dat登记时间 As Date, _
    ByVal dat发生时间 As Date, _
    ByVal str诊室 As String, _
    ByVal blnExcuteSave As Boolean, _
    Optional ByRef strOutExcuteSQL As String, _
    Optional ByRef strOutNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医嘱是否已经收费
    '入参:lng病人id
    '     lng执行部门id
    '     str医生姓名-医生
    '     blnExcuteSave:是否执行保存操作(缺省为true,表示调用该接口直接执行删除;否则返回删除的SQL语句)
    '出参:strOutExcuteSQL-返回执行的SQL
    '     strOutNo-返回的挂号单号
    '返回:接口返回True或False
    '    TRUE:表示获取或执行成功
    '    False:表示获取获取或执行失败
    '编制:刘兴洪
    '日期:2014-04-01 16:58:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, strNO As String
    
    On Error GoTo errHandle
    strOutExcuteSQL = ""
    strNO = gobjDatabase.GetNextNo(12)
    If strNO = "" Then Exit Function
    strOutNO = strNO
    'Zl_病人挂号记录_补充登记
    strSql = "Zl_病人挂号记录_补充登记("
    '  病人id_In     门诊费用记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '  单据号_In     门诊费用记录.No%Type,
    strSql = strSql & "'" & strNO & "',"
    '  执行部门id_In 门诊费用记录.执行部门id%Type,
    strSql = strSql & "" & lng执行部门ID & ","
    '  操作员编号_In 门诊费用记录.操作员编号%Type,
    strSql = strSql & "'" & str操作员编号 & "',"
    '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
    strSql = strSql & "'" & str操作员姓名 & "',"
    '  发生时间_In   门诊费用记录.发生时间%Type,
    strSql = strSql & "to_Date('" & Format(dat发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  登记时间_In   门诊费用记录.登记时间%Type,
    strSql = strSql & "to_Date('" & Format(dat登记时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  医生姓名_In   挂号安排.医生姓名%Type,
    strSql = strSql & "'" & str医生姓名 & "',"
    '  诊室_In       门诊费用记录.发药窗口%Type
    strSql = strSql & "'" & str诊室 & "')"
    strOutExcuteSQL = strSql
    If blnExcuteSave Then
        gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "挂号补充登记"
    End If
    zlRegeventMakeupSave = True
    Exit Function
errHandle:
End Function

Public Function SaveBillChargePrice(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt病人来源 As Byte, ByVal lng病人ID As Long, _
    ByVal lng主页Id As Long, ByVal rsBillDetail As ADODB.Recordset, _
    Optional blnExcuteSave As Boolean = False, Optional strOutNO As String, _
    Optional ByRef cllOutPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定信息,保存为收费划价单
    '入参:
    '    frmMain -调用主窗体
    '    lngModule -模块号
    '    byt病人来源,:1-门诊;2-住院;4-体检
    '    lng病人id,
    '    lng主页ID
    '    rsBillDetail-单据明细,必须有字段( rsBillDetail-单据明细,必须有字段(序号,从属父号,价格父号,开单部门ID,开单人,医嘱ID,收费类别,收费细目ID,收入项目ID,计算单位,付数,数次,单价,应收金额,实收金额,执行部门ID,发生时间,登记时间,操作员姓名,操作员编号,划价人,加班标志(1-加班,0-正常),附加标志(对手术：0-主要手术，1-附加手术；对就诊卡：0-发卡，1-补卡，2-换卡；自动计算：5-修改标志；对挂号：1-病历费,2-就诊卡费；对发药：0-未配药,1-已配药；其他：8-工本费))
     '    blnExcuteSave-是否执行保存操作(缺省为true,表示调用接口直接执行保存单据操作),否则将保存的数据返回给cllOutPro集合)
    '出参:strOutNo-返回保存的单据号
    '    cllOutPro-保存的Sql集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-11 11:02:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSql As String
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim dtCurDate As Date
    Dim rsInfo As ADODB.Recordset
    Dim lng卫材类别ID As Long, lng药品类别ID As Long
    
    If rsBillDetail Is Nothing Then Exit Function
    If rsBillDetail.State <> 1 Then Exit Function
    If rsBillDetail.RecordCount = 0 Then Exit Function
    If lng病人ID = 0 Then Exit Function
    
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    
    dtCurDate = gobjDatabase.Currentdate
    strSql = "" & _
    "Select" & _
    "   A.病人ID,Nvl(B.主页ID,0) 主页ID,To_Number(Nvl(B.当前病区ID,null)) as 病区ID," & _
    "       B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
    "       A.门诊号,B.住院号,B.出院病床 as 床号,NVL( B.姓名 ,A.姓名) 姓名, " & _
    "       NVL(B.性别 ,A.性别) 性别,Y.编码 as 付款方式编码,NVL( B.年龄 ,A.年龄) 年龄 ,Nvl(B.费别,A.费别) as 费别" & _
    " From 病人信息 A,病案主页 B, 医疗付款方式 Y" & _
    " Where A.病人ID=B.病人ID(+)  And A.病人ID=[1] And B.主页ID(+)=[2]  " & _
    "       And A.医疗付款方式=Y.名称(+)"
        
    Set rsInfo = gobjDatabase.OpenSQLRecord(strSql, "获取病人信息", lng病人ID, lng主页Id, 0)
    If rsInfo.EOF Then Exit Function
    rsBillDetail.Filter = "收费类别='5' or 收费类别='6' or 收费类别='7'"
    If rsBillDetail.EOF = False Then
        lng药品类别ID = ExistIOClass(8)
        If lng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    rsBillDetail.Filter = "收费类别='4'"
    If rsBillDetail.EOF = False Then
        lng卫材类别ID = ExistIOClass(40)
        If lng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    rsBillDetail.Filter = 0
    With rsBillDetail
        Dim strNO As String
        If strOutNO <> "" Then
            strNO = strOutNO
        Else
            strNO = gobjDatabase.GetNextNo(13)
        End If
        rsBillDetail.MoveFirst
        Do While Not .EOF
            'ZL_门诊划价记录_INSERT
            strSql = "ZL_门诊划价记录_INSERT("
            '  No_In         门诊费用记录.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  序号_In       门诊费用记录.序号%Type,
            strSql = strSql & "" & !序号 & ","
            '  病人id_In     门诊费用记录.病人id%Type,
            strSql = strSql & "" & Nvl(rsInfo!病人ID, 0) & ","
            '  主页id_In     住院费用记录.主页id%Type,
            strSql = strSql & "NULL,"
            '  标识号_In     门诊费用记录.标识号%Type,
            strSql = strSql & ZVal(rsInfo!门诊号) & ","
            '  付款方式_In   门诊费用记录.付款方式%Type,
            strSql = strSql & "'" & Nvl(rsInfo!付款方式编码) & "',"
            '  姓名_In       门诊费用记录.姓名%Type,
            strSql = strSql & "'" & Nvl(rsInfo!姓名) & "',"
            '  性别_In       门诊费用记录.性别%Type,
            strSql = strSql & "'" & Nvl(rsInfo!性别) & "',"
            '  年龄_In       门诊费用记录.年龄%Type,
            strSql = strSql & "'" & Nvl(rsInfo!年龄) & "',"
            '  费别_In       门诊费用记录.费别%Type,
            strSql = strSql & "'" & Nvl(rsInfo!费别) & "',"
            '  加班标志_In   门诊费用记录.加班标志%Type,
            strSql = strSql & "" & ZVal(!加班标志) & ","
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            strSql = strSql & "" & ZVal(!病人科室id) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            strSql = strSql & "" & ZVal(!开单部门ID) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            strSql = strSql & "'" & Nvl(!开单人) & "',"
            '  从属父号_In   门诊费用记录.从属父号%Type,
            strSql = strSql & "" & ZVal(Val(Nvl(!从属父号))) & ","
            '  收费细目id_In 门诊费用记录.收费细目id%Type,
            strSql = strSql & "" & ZVal(!收费细目ID) & ","
            '  收费类别_In   门诊费用记录.收费类别%Type,
            strSql = strSql & "'" & Nvl(!收费类别) & "',"
            '  计算单位_In   门诊费用记录.计算单位%Type,
            strSql = strSql & "'" & Nvl(!计算单位) & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type,
            strSql = strSql & "'" & "" & "'," 'Nvl(!发药窗口)
            
            '  付数_In       门诊费用记录.付数%Type,
            strSql = strSql & "" & Val(Nvl(!付数)) & ","
            
            '  数次_In       门诊费用记录.数次%Type,
            strSql = strSql & "" & Val(Nvl(!数次)) & ","
            '  附加标志_In   门诊费用记录.附加标志%Type,
            strSql = strSql & "" & ZVal(!附加标志) & ","
            '  执行部门id_In 门诊费用记录.执行部门id%Type,
            strSql = strSql & "" & !执行部门ID & ","
            '  价格父号_In   门诊费用记录.价格父号%Type,
            strSql = strSql & "" & ZVal(!价格父号) & ","
            '  收入项目id_In 门诊费用记录.收入项目id%Type,
            strSql = strSql & "" & ZVal(!收入项目ID) & ","
            '  收据费目_In   门诊费用记录.收据费目%Type,
            strSql = strSql & "'" & Get收据费目(Val(Nvl(!收入项目ID))) & "',"
            '  标准单价_In   门诊费用记录.标准单价%Type,
            strSql = strSql & "" & Val(Nvl(!单价)) & ","
            '  应收金额_In   门诊费用记录.应收金额%Type,
            strSql = strSql & "" & Val(Nvl(!应收金额)) & ","
            '  实收金额_In   门诊费用记录.实收金额%Type,
            strSql = strSql & "" & Val(Nvl(!实收金额)) & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            If IsNull(!发生时间) Then
                strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            Else
                strSql = strSql & " TO_Date('" & Format(!发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            End If
            '  登记时间_In   门诊费用记录.登记时间%Type,
            If IsNull(!登记时间) Then
                strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            Else
                strSql = strSql & " TO_Date('" & Format(!登记时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
            End If
            '  药品摘要_In   药品收发记录.摘要%Type,
            strSql = strSql & "'" & "" & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & Nvl(!操作员姓名) & "',"
            '  费用摘要_In   门诊费用记录.摘要%Type := Null,
            strSql = strSql & "NULL,"   '需要获取发药窗口
            '  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
            strSql = strSql & "" & ZVal(!医嘱ID) & ","
            '  频次_In       药品收发记录.频次%Type := Null,
            strSql = strSql & "NULL,"
            '  单量_In       药品收发记录.单量%Type := Null,
            strSql = strSql & "NULL,"
            '  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
            strSql = strSql & "NULL,"
            '  期效_In       药品收发记录.扣率%Type := Null,
            strSql = strSql & "NULL,"
            '  计价特性_In   药品收发记录.扣率%Type := Null,
            strSql = strSql & "NULL,"
            '    病人来源_In   Number := 1,
            strSql = strSql & "" & byt病人来源 & ","
            '  保险编码_In   门诊费用记录.保险编码%Type := Null,
            strSql = strSql & "NULL,"
            '  费用类型_In   门诊费用记录.费用类型%Type := Null,
            strSql = strSql & "NULL,"
            '  保险项目否_In 门诊费用记录.保险项目否%Type := Null,
            strSql = strSql & "NULL,"
            '  保险大类id_In 门诊费用记录.保险大类id%Type := Null,
            strSql = strSql & "NULL,"
            '  中药形态_In   门诊费用记录.结论%Type := Null,
            strSql = strSql & "NULL,"
            '  备货材料_In   Number := 0,
            strSql = strSql & "0,"
            '  批次_In       药品收发记录.批次%Type := Null
            strSql = strSql & "NULL,"
            '  执行人_In     门诊费用记录.执行人%Type := Null,
            strSql = strSql & "NULL,"
            '  病人病区id_In 门诊费用记录.病人病区id%Type := Null
            strSql = strSql & "" & ZVal(rsInfo!病区ID) & ")"
            Call zlAddArray(cllOutPro, strSql)
            .MoveNext
        Loop
    End With
    If blnExcuteSave Then
        zlExecuteProcedureArrAy cllOutPro, False, False
    End If
    strOutNO = strNO
    SaveBillChargePrice = True
End Function
Public Function SaveBillBilling(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt病人来源 As Byte, ByVal lng病人ID As Long, _
    ByVal lng主页Id As Long, ByVal blnSavePrice As Boolean, ByVal rsBillDetail As ADODB.Recordset, _
    Optional blnExcuteSave As Boolean = False, Optional strOutNO As String, _
    Optional ByRef cllOutPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定信息,保存为记帐划价单或记帐单
    '入参:
    '    frmMain -调用主窗体
    '    lngModule -模块号
    '    byt病人来源,:1-门诊;2-住院;3-其他;4-体检
    '    blnSavePrice-保存为划价单
    '    lng病人id,
    '    lng主页ID
    '    rsBillDetail-单据明细,必须有字段(序号,从属父号,价格父号,开单部门ID,开单人,医嘱ID,收费类别,收费细目ID,收入项目ID,计算单位,付数,数次,单价,应收金额,实收金额,执行部门ID,病人科室id,开单部门ID,附加标志,加班标志,婴儿费,发生时间,登记时间,操作员姓名,操作员编号,划价人)
    '    blnExcuteSave-是否执行保存操作(缺省为true,表示调用接口直接执行保存单据操作),否则将保存的数据返回给cllOutPro集合)
    '出参:strOutNo-返回保存的单据号
    '    cllOutPro-保存的Sql集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-11 11:02:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    Dim objBill As New ExpenseBill
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim dtCurDate As Date, strSql As String
    Dim rsInfo As ADODB.Recordset
    Dim lng卫材类别ID As Long, lng药品类别ID As Long
    Dim arrSQL As Variant, strTmp As String
    
    If rsBillDetail Is Nothing Then Exit Function
    If rsBillDetail.State <> 1 Then Exit Function
    If rsBillDetail.RecordCount = 0 Then Exit Function
    If lng病人ID = 0 Then Exit Function
    
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    
    arrSQL = Array()
    
    dtCurDate = gobjDatabase.Currentdate
    strSql = "" & _
    "Select" & _
    "   A.病人ID,Nvl(B.主页ID,0) 主页ID,To_Number(Nvl(B.当前病区ID,null)) as 病区ID," & _
    "       B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
    "       A.门诊号,B.住院号,B.出院病床 as 床号,NVL( B.姓名 ,A.姓名) 姓名, " & _
    "       NVL(B.性别 ,A.性别) 性别,Y.编码 as 付款方式编码,NVL( B.年龄 ,A.年龄) 年龄 ,Nvl(B.费别,A.费别) as 费别 " & _
    " From 病人信息 A,病案主页 B, 医疗付款方式 Y" & _
    " Where A.病人ID=B.病人ID(+)  And A.病人ID=[1] And B.主页ID(+)=[2]  " & _
    "       And A.医疗付款方式=Y.名称(+)"
        
    Set rsInfo = gobjDatabase.OpenSQLRecord(strSql, "获取病人信息", lng病人ID, lng主页Id, 0)
    If rsInfo.EOF Then Exit Function
    rsBillDetail.Filter = "收费类别='5' or 收费类别='6' or 收费类别='7'"
    If rsBillDetail.EOF = False Then
        lng药品类别ID = ExistIOClass(9)
        If lng药品类别ID = 0 Then
            MsgBox "不能确定处方单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    rsBillDetail.Filter = "收费类别='4'"
    If rsBillDetail.EOF = False Then
        lng卫材类别ID = ExistIOClass(41)
        If lng卫材类别ID = 0 Then
            MsgBox "不能确定卫材单据的入出类别,请先到入出分类管理中设置！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    rsBillDetail.Filter = 0
    With rsBillDetail
        Dim strNO As String
        If strOutNO <> "" Then
            strNO = strOutNO
        Else
            strNO = gobjDatabase.GetNextNo(14)
        End If
        rsBillDetail.MoveFirst
        Do While Not .EOF
            If byt病人来源 = 2 Then
                '住院
                'Zl_住院记帐记录_Insert
                strSql = "Zl_住院记帐记录_Insert("
                '  No_In             住院费用记录.No%Type,
                strSql = strSql & "'" & strNO & "',"
                '  序号_In           住院费用记录.序号%Type,
                strSql = strSql & "" & !序号 & ","
                '  病人id_In         住院费用记录.病人id%Type,
                strSql = strSql & "" & Nvl(rsInfo!病人ID, 0) & ","
                '  主页id_In         住院费用记录.主页id%Type,
                strSql = strSql & "" & IIf(lng主页Id = 0, ZVal(rsInfo!主页ID), lng主页Id) & ","
                '  标识号_In         住院费用记录.标识号%Type,
                strSql = strSql & "" & ZVal(rsInfo!住院号) & ","
                '  姓名_In           住院费用记录.姓名%Type,
                strSql = strSql & "'" & Nvl(rsInfo!姓名) & "',"
                '  性别_In           住院费用记录.性别%Type,
                strSql = strSql & "'" & Nvl(rsInfo!性别) & "',"
                '  年龄_In           住院费用记录.年龄%Type,
                strSql = strSql & "'" & Nvl(rsInfo!年龄) & "',"
                '  床号_In           住院费用记录.床号%Type,
                strSql = strSql & "'" & Nvl(rsInfo!床号) & "',"
                '  费别_In           住院费用记录.费别%Type,
                strSql = strSql & "'" & Nvl(rsInfo!费别) & "',"
                '  病区id_In         住院费用记录.病人病区id%Type,
                strSql = strSql & "" & ZVal(rsInfo!病区ID) & ","
                '  科室id_In         住院费用记录.病人科室id%Type,
                strSql = strSql & "" & ZVal(!病人科室id) & ","
                '  加班标志_In       住院费用记录.加班标志%Type,
                strSql = strSql & "" & ZVal(!加班标志) & ","
                '  婴儿费_In         住院费用记录.婴儿费%Type,
                strSql = strSql & "" & ZVal(!婴儿费) & ","
                '  开单部门id_In     住院费用记录.开单部门id%Type,
                strSql = strSql & "" & ZVal(!开单部门ID) & ","
                '  开单人_In         住院费用记录.开单人%Type,
                strSql = strSql & "'" & Nvl(!开单人) & "',"
                '  从属父号_In       住院费用记录.从属父号%Type,
                strSql = strSql & "" & ZVal(!从属父号) & ","
                '  收费细目id_In     住院费用记录.收费细目id%Type,
                strSql = strSql & "" & ZVal(!收费细目ID) & ","
                '  收费类别_In       住院费用记录.收费类别%Type,
                strSql = strSql & "'" & Nvl(!收费类别) & "',"
                '  计算单位_In       住院费用记录.计算单位%Type,
                strSql = strSql & "'" & Nvl(!计算单位) & "',"
                '  保险项目否_In     住院费用记录.保险项目否%Type,
                strSql = strSql & "" & "NULL" & ","
                '  保险大类id_In     住院费用记录.保险大类id%Type,
                strSql = strSql & "" & "NULL" & ","
                '  保险编码_In       住院费用记录.保险编码%Type,
                strSql = strSql & "" & "NULL" & ","
                '  付数_In           住院费用记录.付数%Type,
                strSql = strSql & "" & Val(Nvl(!付数)) & ","
                '  数次_In           住院费用记录.数次%Type,
                strSql = strSql & "" & Val(Nvl(!数次)) & ","
                '  附加标志_In       住院费用记录.附加标志%Type,
                strSql = strSql & "" & ZVal(!附加标志) & ","
                '  执行部门id_In     住院费用记录.执行部门id%Type,
                strSql = strSql & "" & !执行部门ID & ","
                '  价格父号_In       住院费用记录.价格父号%Type,
                strSql = strSql & "" & ZVal(!价格父号) & ","
                '  收入项目id_In     住院费用记录.收入项目id%Type,
                strSql = strSql & "" & ZVal(!收入项目ID) & ","
                '  收据费目_In       住院费用记录.收据费目%Type,
                strSql = strSql & "'" & Get收据费目(Val(Nvl(!收入项目ID))) & "',"
                '  标准单价_In       住院费用记录.标准单价%Type,
                strSql = strSql & "" & Val(Nvl(!单价)) & ","
                '  应收金额_In       住院费用记录.应收金额%Type,
                strSql = strSql & "" & Val(Nvl(!应收金额)) & ","
                '  实收金额_In       住院费用记录.实收金额%Type,
                strSql = strSql & "" & Val(Nvl(!实收金额)) & ","
                '  统筹金额_In       住院费用记录.统筹金额%Type,
                strSql = strSql & "" & "NULL" & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                If IsNull(!发生时间) Then
                    strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                Else
                    strSql = strSql & " TO_Date('" & Format(!发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                End If
                '  登记时间_In   门诊费用记录.登记时间%Type,
                If IsNull(!登记时间) Then
                    strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                Else
                    strSql = strSql & " TO_Date('" & Format(!登记时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                End If
                '  药品摘要_In       药品收发记录.摘要%Type,
                strSql = strSql & "'" & "" & "',"
                strSql = strSql & "" & IIf(blnSavePrice, 1, 0) & ","
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                strSql = strSql & "'" & Nvl(!操作员编号) & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                strSql = strSql & "'" & Nvl(!操作员姓名) & "',"
                '  多病人单_In       Number := 0,
                strSql = strSql & "NULL,"
                '  类别id_In         药品单据性质.类别id%Type := Null,
                If InStr(1, ",5,6,7,", "," & Nvl(!收费类别) & ",") > 0 Then
                    strSql = strSql & "" & lng药品类别ID & ","
                ElseIf InStr(1, ",4,", "," & Nvl(!收费类别) & ",") > 0 Then
                    strSql = strSql & "" & lng卫材类别ID & ","
                Else
                    strSql = strSql & "NULL,"   '需要获取发药窗口
                End If
                '  记帐单id_In       住院费用记录.记帐单id%Type := Null,
                strSql = strSql & "NULL,"
                '  费用摘要_In       住院费用记录.摘要%Type := Null,
                strSql = strSql & "NULL,"
                '  是否急诊_In       住院费用记录.是否急诊%Type := 0,
                strSql = strSql & "NULL,"
                '  医嘱序号_In       住院费用记录.医嘱序号%Type := Null,
                strSql = strSql & "" & ZVal(!医嘱ID) & ")"
                '  频次_In           药品收发记录.频次%Type := Null,
                '  单量_In           药品收发记录.单量%Type := Null,
                '  用法_In           药品收发记录.用法%Type := Null, --用法[|煎法]
                '  期效_In           药品收发记录.扣率%Type := Null,
                '  计价特性_In       药品收发记录.扣率%Type := Null,
                '  简单记帐_In       Number := 0,
                '  费用类型_In       住院费用记录.费用类型%Type := Null,
                '  医技补临床费用_In Number := 0,
                '  领药部门id_In     药品收发记录.对方部门id%Type := Null,
                '  中药形态_In       住院费用记录.结论%Type := Null,
                '  医疗小组id_In     住院费用记录.医疗小组id%Type := -1,
                '  备货材料_In       Number := 0,
                '  批次_In           药品收发记录.批次%Type := Null
            Else
                '门诊
                'Zl_门诊记帐记录_Insert
                strSql = "Zl_门诊记帐记录_Insert("
                '  No_In         门诊费用记录.No%Type,
                strSql = strSql & "'" & strNO & "',"
                '  序号_In       门诊费用记录.序号%Type,
                strSql = strSql & "" & !序号 & ","
                '  病人id_In     门诊费用记录.病人id%Type,
                strSql = strSql & "" & Nvl(rsInfo!病人ID, 0) & ","
                '  标识号_In     门诊费用记录.标识号%Type,
                strSql = strSql & IIf(Val(Nvl(rsInfo!门诊号)) = 0, "NULL", Val(Nvl(rsInfo!门诊号))) & ","
                '  姓名_In       门诊费用记录.姓名%Type,
                strSql = strSql & "'" & Nvl(rsInfo!姓名) & "',"
                '  性别_In       门诊费用记录.性别%Type,
                strSql = strSql & "'" & Nvl(rsInfo!性别) & "',"
                '  年龄_In       门诊费用记录.年龄%Type,
                strSql = strSql & "'" & Nvl(rsInfo!年龄) & "',"
                '  费别_In       门诊费用记录.费别%Type,
                strSql = strSql & "'" & Nvl(rsInfo!费别) & "',"
                '  加班标志_In   门诊费用记录.加班标志%Type,
                strSql = strSql & "" & ZVal(!加班标志) & ","
                '   婴儿费_In     门诊费用记录.婴儿费%Type,
                strSql = strSql & "" & ZVal(0) & ","
                
                '  病人科室id_In 门诊费用记录.病人科室id%Type,
                strSql = strSql & "" & ZVal(!病人科室id) & ","
                '  开单部门id_In 门诊费用记录.开单部门id%Type,
                strSql = strSql & "" & ZVal(!开单部门ID) & ","
                '  开单人_In     门诊费用记录.开单人%Type,
                strSql = strSql & "'" & Nvl(!开单人) & "',"
                '  从属父号_In   门诊费用记录.从属父号%Type,
                strSql = strSql & "" & ZVal(Val(Nvl(!从属父号))) & ","
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSql = strSql & "" & ZVal(!收费细目ID) & ","
                '  收费类别_In   门诊费用记录.收费类别%Type,
                strSql = strSql & "'" & Nvl(!收费类别) & "',"
                '  计算单位_In   门诊费用记录.计算单位%Type,
                strSql = strSql & "'" & Nvl(!计算单位) & "',"
                '  付数_In       门诊费用记录.付数%Type,
                strSql = strSql & "" & Val(Nvl(!付数)) & ","
                
                '  数次_In       门诊费用记录.数次%Type,
                strSql = strSql & "" & Val(Nvl(!数次)) & ","
                '  附加标志_In   门诊费用记录.附加标志%Type,
                strSql = strSql & "" & ZVal(!附加标志) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSql = strSql & "" & !执行部门ID & ","
                '  价格父号_In   门诊费用记录.价格父号%Type,
                strSql = strSql & "" & ZVal(!价格父号) & ","
                '  收入项目id_In 门诊费用记录.收入项目id%Type,
                strSql = strSql & "" & ZVal(!收入项目ID) & ","
                '  收据费目_In   门诊费用记录.收据费目%Type,
                strSql = strSql & "'" & Get收据费目(Val(Nvl(!收入项目ID))) & "',"
                '  标准单价_In   门诊费用记录.标准单价%Type,
                strSql = strSql & "" & Val(Nvl(!单价)) & ","
                '  应收金额_In   门诊费用记录.应收金额%Type,
                strSql = strSql & "" & Val(Nvl(!应收金额)) & ","
                '  实收金额_In   门诊费用记录.实收金额%Type,
                strSql = strSql & "" & Val(Nvl(!实收金额)) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                If IsNull(!发生时间) Then
                    strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                Else
                    strSql = strSql & " TO_Date('" & Format(!发生时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                End If
                '  登记时间_In   门诊费用记录.登记时间%Type,
                If IsNull(!登记时间) Then
                    strSql = strSql & " TO_Date('" & Format(dtCurDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                Else
                    strSql = strSql & " TO_Date('" & Format(!登记时间, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
                End If
                '  药品摘要_In   药品收发记录.摘要%Type,
                strSql = strSql & "'" & "" & "',"
                '  划价_In       Number,
                strSql = strSql & "" & IIf(blnSavePrice, 1, 0) & " ,"
                
                '  操作员编号_In 门诊费用记录.操作员编号%Type,
                strSql = strSql & "'" & Nvl(!操作员编号) & "',"
                '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
                strSql = strSql & "'" & Nvl(!操作员姓名) & "',"
                '  记帐单id_In   门诊费用记录.记帐单id%Type := Null,
                strSql = strSql & "NULL,"
                '  费用摘要_In   门诊费用记录.摘要%Type := Null,
                    strSql = strSql & "NULL,"   '需要获取发药窗口
                '  医嘱序号_In   门诊费用记录.医嘱序号%Type := Null,
                strSql = strSql & "" & ZVal(!医嘱ID) & ","
                '   频次_In       药品收发记录.频次%Type := Null,
                strSql = strSql & "NULL,"
                '  单量_In       药品收发记录.单量%Type := Null,
                strSql = strSql & "NULL,"
                '  用法_In       药品收发记录.用法%Type := Null, --用法[|煎法]
                strSql = strSql & "NULL,"
                '  期效_In       药品收发记录.扣率%Type := Null,
                strSql = strSql & "NULL,"
                '  计价特性_In   药品收发记录.扣率%Type := Null,
                strSql = strSql & "NULL,"
                '  门诊标志_In   门诊费用记录.门诊标志%Type := 1,
                strSql = strSql & "" & byt病人来源 & ","
                '  中药形态_In   门诊费用记录.结论%Type := Null,
                strSql = strSql & "NULL,"
                '  备货材料_In   Number := 0,
                strSql = strSql & "0,"
                '  批次_In       药品收发记录.批次%Type := Null
                strSql = strSql & "NULL,"
                '  主页id_In     门诊费用记录.主页id%Type := Null,
                strSql = strSql & "" & IIf(lng主页Id = 0, ZVal(rsInfo!主页ID), lng主页Id) & ","
                '  病人病区id_In 门诊费用记录.病人病区id%Type := Null
                strSql = strSql & "" & ZVal(rsInfo!病区ID) & ")"
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = ZVal(!收费细目ID) & ";" & strSql
            .MoveNext
        Loop
    End With
    
     '对SQL序列按收费细目ID排序
    For i = 0 To UBound(arrSQL) - 1
        For j = i + 1 To UBound(arrSQL)
            If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                strTmp = CStr(arrSQL(j))
                arrSQL(j) = arrSQL(i)
                arrSQL(i) = strTmp
            End If
        Next
    Next
    
    For i = 0 To UBound(arrSQL)
        Call zlAddArray(cllOutPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1))
    Next
    
    If blnExcuteSave Then
        zlExecuteProcedureArrAy cllOutPro, False, False
    End If
    strOutNO = strNO
    SaveBillBilling = True
End Function
Public Function VerfyBillingPriceBill(ByVal byt病人来源 As Byte, ByVal strNos As String, _
   ByVal blnExcuteSave As Boolean, Optional ByRef cllOutPro As Collection, _
   Optional ByVal strVerfyDate As String = "", _
   Optional ByVal strVerfyPersonCode As String, Optional strVerfyPersonName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐划价单审核为记帐单
    '入参:byt病人来源:1-门诊;2-住院
    '    strNos -单据号, 多个用逗号分离
    '    blnExcuteSave-是否执行保存操作(缺省为true,表示调用接口直接执行保存单据操作),否则将保存的数据返回给cllOutPro集合)
    '    strVerfyDate-审核日期(为空时，以当前时间为准):格式：yyyy-mm-dd HH:MM:SS
    '    strVerfyPersonCode-审核人编号(为空时，以当前操作员为准)
    '    strVerfyPersonName-审核人(为空时，以当前操作员为准)
    '出参:cllOutPro-保存的Sql集
    '返回:接口返回True或False：TRUE:执行成功；False:执行失败
    '编制:刘兴洪
    '日期:2014-04-14 09:07:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, i As Long, strSql As String
    
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    
    varNos = Split(strNos, ",")
    If strVerfyDate = "" Or Not IsDate(strVerfyDate) Then
        strVerfyDate = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        strVerfyDate = Format(CDate(strVerfyDate), "yyyy-mm-dd HH:MM:SS")
    End If
    
    If strVerfyPersonCode = "" Or strVerfyPersonName = "" Then
        strVerfyPersonCode = UserInfo.编号
        strVerfyPersonName = UserInfo.姓名
    End If
    For i = 0 To UBound(varNos)
        If byt病人来源 = 1 Then
            'Zl_门诊记帐记录_Verify
            strSql = "zl_门诊记帐记录_Verify("
            '  No_In         门诊费用记录.No%Type,
            strSql = strSql & "'" & varNos(i) & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & strVerfyPersonCode & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & strVerfyPersonName & "',"
            '  序号_In       Varchar2 := Null,
            strSql = strSql & "NULL,"
            '  审核时间_In   门诊费用记录.登记时间%Type := Null
            strSql = strSql & "to_date('" & strVerfyDate & "','yyyy-mm-dd hh24:mi:ss'))"
        Else
            '住院
            'Zl_住院记帐记录_Verify
            strSql = "Zl_住院记帐记录_Verify("
            '  No_In         住院费用记录.No%Type,
            strSql = strSql & "'" & varNos(i) & "',"
            '  操作员编号_In 住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & strVerfyPersonCode & "',"
            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & strVerfyPersonName & "',"
            '  序号_In       Varchar2 := Null,
            strSql = strSql & "NULL,"
            '  病人id_In     住院费用记录.病人id%Type := Null,
            strSql = strSql & "NULL,"
            '  审核时间_In   住院费用记录.登记时间%Type := Null
            strSql = strSql & "to_date('" & strVerfyDate & "','yyyy-mm-dd hh24:mi:ss'))"
        End If
        Call zlAddArray(cllOutPro, strSql)
    Next
    If blnExcuteSave Then
        zlExecuteProcedureArrAy cllOutPro, "执行划价审核", False, False
    End If
    VerfyBillingPriceBill = True
End Function

Public Function DeleteBill(ByVal byt记录性质 As Byte, _
    ByVal strNos As String, ByVal blnExcuteDel As Boolean, _
    Optional ByRef cllOutPro As Collection, _
    Optional ByVal strOperatorCode As String, Optional strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除收费划价单或记帐单
    '入参:byt记录性质-1-收费划价单;2-门诊记帐划价单;3-住院划价记帐单
    '    strNOs -单据号(允许多个，多个用逗号分离)
    '    blnExcuteDel-是否执行删除(缺省为true,表示调用该接口直接执行删除;否则返回删除的SQL语句)
    '    strOperatorName-操作员姓名(不传入,使用当前操作员姓名)
    '    strOperatorCode-操作员编号(不传入,使用当前操作员编号)
    '出参:cllOutPro-删除的Sql集
    '返回:删除或执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-14 09:27:21
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, i As Long, strSql As String
    varNos = Split(strNos, ",")
    
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    If strOperatorCode = "" Or strOperatorName = "" Then
        strOperatorCode = UserInfo.编号
        strOperatorName = UserInfo.姓名
    End If
    
    For i = 0 To UBound(varNos)
        If byt记录性质 = 1 Then
           strSql = "zl_门诊划价记录_DELETE('" & varNos(i) & "')"
        ElseIf byt记录性质 = 2 Then
            'Zl_门诊记帐记录_Delete
            strSql = "Zl_门诊记帐记录_Delete("
            '  No_In         门诊费用记录.No%Type,
            strSql = strSql & "'" & varNos(i) & "',"
            '  序号_In       Varchar2,
            strSql = strSql & "NULL,"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & strOperatorCode & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type
            strSql = strSql & "'" & strOperatorName & "')"
        Else
            'Zl_住院记帐记录_Delete
            strSql = "Zl_住院记帐记录_Delete("
            '  No_In           住院费用记录.No%Type,
            strSql = strSql & "'" & varNos(i) & "',"
            '  序号_In         Varchar2,
            strSql = strSql & "NULL,"
            '  操作员编号_In   住院费用记录.操作员编号%Type,
            strSql = strSql & "'" & strOperatorCode & "',"
            '  操作员姓名_In   住院费用记录.操作员姓名%Type,
            strSql = strSql & "'" & strOperatorName & "')"
            '  记录性质_In     住院费用记录.记录性质%Type := 2,
            '  操作状态_In     Number := 0,
            '  输液配药检查_In Number:=1
       End If
       Call zlAddArray(cllOutPro, strSql)
    Next
    If blnExcuteDel Then
        zlExecuteProcedureArrAy cllOutPro, "执行销帐或删除划价单", False, False
    End If
    DeleteBill = True
End Function
Public Function CreateBillingRecStru(ByRef rsBilling As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建本地记帐的记录集结构
    '出参:rsBilling-返回记帐的本地记录集
    '返回:返回本地记录集成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-15 12:02:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set rsBilling = New ADODB.Recordset
    rsBilling.Fields.Append "序号", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "从属父号", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "价格父号", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsBilling.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "收入项目ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    rsBilling.Fields.Append "付数", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "数次", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "单价", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "应收金额", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsBilling.Fields.Append "附加标志", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "加班标志", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "婴儿费", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "病人科室id", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsBilling.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsBilling.Fields.Append "划价人", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "发生时间", adDate, , adFldIsNullable
    rsBilling.Fields.Append "登记时间", adDate, , adFldIsNullable
    rsBilling.Fields.Append "操作员姓名", adVarChar, 20, adFldIsNullable
    rsBilling.Fields.Append "操作员编号", adVarChar, 10, adFldIsNullable
    rsBilling.CursorLocation = adUseClient
    rsBilling.LockType = adLockOptimistic
    rsBilling.CursorType = adOpenStatic
    rsBilling.Open

    CreateBillingRecStru = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Public Function zlClinicCharge(ByVal lngModule As Long, ByVal strNos As String, byt病人来源 As Byte, _
    ByVal cllBalance As Collection, ByVal dbl缴款 As Double, ByVal dbl找补 As Double, _
    ByRef lng结帐ID As Long, ByRef cllOutPro As Collection, _
    ByRef strErrMsg As String, Optional ByVal blnExcuteCharge As Boolean = True, _
    Optional ByRef strExpendes As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊收费结算
    '入参:strNos-收费划价单号,多个用逗号分离
    '     cllBalance-结算信息集:array(结算类别,卡类别ID,结算方式,结算金额,消费卡ID,卡号,交易流水号,交易说明,结算号码,摘要)
    '                 结算类别:0-非医保和非三方结算;1-三方卡结算;2-消费卡结算;3-冲预交款,4-医保结算;5-误差费
    '                 消费卡ID:结算类别为2时,才传入,为消费卡目录.ID,为零时,根据卡号来查找;其他传入0
    '     byt病人来源-病人来源(1-门诊;2-住院)
    '     dbl缴款-缴款额
    '     dbl找补-找补额
    '     strExpendes-扩展结算信息
    '     blnExcuteCharge-是否执行收费操作
    '     strExpendes-扩展信息,暂无用,以后扩展
    '出参:lng结帐ID-返回结帐ID
    '     cllOutPro-返回需要执行的SQL
    '     strErrMsg-返回False时,返回的出错信息
    '返回:执行成功或获取保存SQL成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 16:23:16
    '说明:该过程中不处理异常,主要由调用方处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strPayWins As String, strSql As String, str发生时间 As String
    Dim rsTemp As ADODB.Recordset, rsInfo As ADODB.Recordset
    Dim cllDept As Collection, varData As Variant, varTemp As Variant, i As Long
    Dim lng病人ID As Long, lng结帐序号 As Long, lngCardTypeID As Long, byt操作类型 As Byte
    Dim dbl结算金额 As Double, dblMoney As Double, dbl冲预交款 As Double, dbl误差费 As Double
    

    Set cllDept = New Collection
    strSql = "" & _
    "   Select  /*+ rule */ distinct 执行部门ID,收费类别,发药窗口" & _
    "   From 门诊费用记录 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) J  " & _
    "  Where  A.No=J.Column_value and 记录状态=0 and A.记录性质=1 and A.收费类别 in ('5','6','7')" & _
    "  "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取本次收费的单据信息", strNos)
    With rsTemp
        Do While Not .EOF
            If InStr(strTemp, "," & Nvl(!收费类别) & "|" & Nvl(!执行部门ID) & ",") = 0 Then
                cllDept.Add Array(Nvl(!收费类别), Val(Nvl(!执行部门ID)), Nvl(!发药窗口))
            End If
            .MoveNext
        Loop
    End With
    strSql = "" & _
    "   Select  /*+ rule */ NO,Max(付款方式) as 付款方式, " & _
    "           Max(病人科室ID) as 病人科室ID,Max(开单部门ID) as 开单部门Id, " & _
    "           Max(发药窗口) as 发药窗口,max(是否急诊) as 是否急诊,Max(A.病人ID) as 病人ID," & _
    "           Sum(实收金额) as 金额" & _
    "   From 门诊费用记录 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) J  " & _
    "  Where  A.No=J.Column_value and 记录状态=0 and A.记录性质=1" & _
    "  Group by NO" & _
    "  Order by NO"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取本次收费的单据信息", strNos)
    With rsTemp
        If .RecordCount = 0 Then
            strErrMsg = "未找到单号为:'" & strNos & "'的收费划价记录": Exit Function
        End If
        .MoveFirst
        strTemp = ""
        lng病人ID = Val(Nvl(!病人ID))
        Do While Not .EOF
            dbl结算金额 = dbl结算金额 + Val(Nvl(rsTemp!金额))
            If lng病人ID <> Val(Nvl(!病人ID)) Then
                strErrMsg = "同一次结算中包含了不同病人的收费信息": Exit Function
            End If
            If InStr(1, "," & strNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                strTemp = strTemp & "," & Nvl(!NO)
            End If
            .MoveNext
        Loop
    End With
    
    dbl结算金额 = FormatEx(dbl结算金额, 7)
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 2)
        strErrMsg = "未找到单号为:'" & strTemp & "'的收费划价记录": Exit Function
    End If
    
    strSql = "" & _
    "   Select Decode(Sign(A.就诊时间-A.登记时间),0,1,0) as 初诊,A.病人ID,A.门诊号,A.住院号, " & _
    "           A.姓名,A.性别,A.年龄 ,A.费别,A.医疗付款方式,M.编码 as 付款方式编码, " & _
    "           nvl(C.预交余额,0)-nvl(C.费用余额,0) as 余额 " & _
    "   From 病人信息 A,医疗付款方式 M,病人余额 C" & _
    "   Where  A.医疗付款方式=M.名称(+)And A.病人ID=[1] " & _
    "           And A.病人ID=C.病人ID(+) and C.性质(+)=1 and C.类型(+)=1"
    Set rsInfo = gobjDatabase.OpenSQLRecord(strSql, "获取病人信息", lng病人ID)
    If rsInfo.EOF Then
        strErrMsg = "未找到指定的病人信息": Exit Function
    End If

    lng结帐ID = gobjDatabase.GetNextId("病人结帐记录")
    lng结帐序号 = -1 * lng结帐ID
    str发生时间 = Format(gobjDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Call zlGetPayDrugWindow(lngModule, lng病人ID, CDate(str发生时间), cllDept, strPayWins)
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '---------------------------------------------------------------
            'Zl_病人划价收费_Insert
            strSql = "Zl_病人划价收费_Insert("
            '  No_In         门诊费用记录.NO%Type,
            strSql = strSql & "'" & Nvl(rsTemp!NO) & "',"
            '  病人id_In     门诊费用记录.病人id%Type,
            strSql = strSql & "" & ZVal(lng病人ID) & ","
            '  病人来源_In   Number,
            strSql = strSql & "" & byt病人来源 & ","
            '  付款方式_In   门诊费用记录.付款方式%Type,
            If Nvl(rsInfo!付款方式编码) <> "" Then
               strSql = strSql & "'" & Nvl(rsInfo!付款方式编码) & "',"
            Else
               strSql = strSql & "'" & Nvl(rsTemp!付款方式) & "',"
            End If
            '  姓名_In       门诊费用记录.姓名%Type,
            strSql = strSql & "'" & Nvl(rsInfo!姓名) & "',"
            '  性别_In       门诊费用记录.性别%Type,
            strSql = strSql & "'" & Nvl(rsInfo!性别) & "',"
            '  年龄_In       门诊费用记录.年龄%Type,
            strSql = strSql & "'" & Nvl(rsInfo!年龄) & "',"
            '  病人科室id_In 门诊费用记录.病人科室id%Type,
            strSql = strSql & "" & IIf(Val(Nvl(rsTemp!病人科室id)) = 0, "NULL", Val(Nvl(rsTemp!病人科室id))) & ","
            '  开单部门id_In 门诊费用记录.开单部门id%Type,
            strSql = strSql & "" & IIf(Val(Nvl(rsTemp!开单部门ID)) = 0, "NULL", Val(Nvl(rsTemp!开单部门ID))) & ","
            '  开单人_In     门诊费用记录.开单人%Type,
            strSql = strSql & "NULL,"    ' 过程内部处理,保持原来的不变
            '  结帐id_In     门诊费用记录.结帐id%Type,
            strSql = strSql & "" & lng结帐ID & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            strSql = strSql & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'),"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSql = strSql & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSql = strSql & "'" & UserInfo.姓名 & "',"
            '  发药窗口_In   门诊费用记录.发药窗口%Type := Null,
              strSql = strSql & "'" & strPayWins & "',"
            '  是否急诊_In   门诊费用记录.是否急诊%Type := 0,
            strSql = strSql & "" & Val(Nvl(rsTemp!是否急诊)) & ","
            '  登记时间_In   门诊费用记录.登记时间%Type := Null,
            strSql = strSql & "to_date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss'))"
            zlAddArray cllOutPro, strSql
            .MoveNext
        Loop
    End With
    
    '保存结算数据
    Dim str医保结算 As String, str收费结算 As String
    dblMoney = 0: dbl误差费 = 0: dbl冲预交款 = 0
    'Zl_门诊收费结算_Modify
    '    操作类型_In   Number,
    '    --操作类型_In:
    '    --   0-普通收费方式:
    '    --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '    --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
    '    --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '    --   1.三方卡结算:
    '    --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
    '    --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '    --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '    --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
    '    --   3-消费卡结算:
    '    --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '    --     ②冲预交_In: 传入零
    '    --     ③退支票额_In:传入零
        
    If Not cllBalance Is Nothing Then
        'array(结算类别,卡类别ID,结算方式,结算金额,消费卡ID,卡号,交易流水号,交易说明,结算号码,摘要)
        For i = 1 To cllBalance.Count
            varTemp = cllBalance(i)
            dblMoney = dblMoney + Val(varTemp(3))
            strTemp = ""
            '结算类别:0-非医保和非三方结算;1-三方卡结算;2-消费卡结算;3-冲预交款
            Select Case Val(varTemp(0))
            Case 0 '非医保结算
            '"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
                str收费结算 = str收费结算 & "||" & Trim(varTemp(2))
                str收费结算 = str收费结算 & "|" & Val(varTemp(3))
                str收费结算 = str收费结算 & "|" & IIf(Trim(varTemp(8)) = "", " ", Trim(varTemp(8)))
                str收费结算 = str收费结算 & "|" & IIf(Trim(varTemp(9)) = "", " ", Trim(varTemp(9)))
            Case 2  '消费卡结算
                '卡类别ID|卡号|消费卡ID|消费金额||..."
                lngCardTypeID = Val(varTemp(1))
                strTemp = Val(varTemp(1))
                strTemp = strTemp & "|" & Trim(varTemp(5))
                strTemp = strTemp & "|" & Val(varTemp(4))
                strTemp = strTemp & "|" & Val(varTemp(3))
                byt操作类型 = 3
            Case 3  '冲预交款
                dbl冲预交款 = dbl冲预交款 + Val(varTemp(3))
            Case 4 '4-医保结算;
                '"结算方式|结算金额||..."
                str医保结算 = str医保结算 & "||" & Trim(varTemp(2))
                str医保结算 = str医保结算 & "|" & Val(varTemp(3))
            Case 5 '误差费
                dbl误差费 = dbl误差费 + Val(varTemp(3))
            Case 1 ' 三方卡结算
                lngCardTypeID = Val(varTemp(1))
                '"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
                strTemp = Trim(varTemp(2))
                strTemp = strTemp & "|" & Val(varTemp(3))
                strTemp = strTemp & "|" & IIf(Trim(varTemp(8)) = "", " ", Trim(varTemp(8)))
                strTemp = strTemp & "|" & IIf(Trim(varTemp(9)) = "", " ", Trim(varTemp(9)))
                
                byt操作类型 = 1
            End Select
            
            If strTemp <> "" Then
                ' Zl_门诊收费结算_Modify
                strSql = "Zl_门诊收费结算_Modify("
                '操作类型_In   Number,
                strSql = strSql & byt操作类型 & ","
                '    病人id_In     门诊费用记录.病人id%Type,
                strSql = strSql & lng病人ID & ","
                '    结帐id_In     病人预交记录.结帐id%Type,
                strSql = strSql & lng结帐ID & ","
                '    结算方式_In   Varchar2,
                strSql = strSql & "'" & strTemp & "',"
                '    冲预交_In     病人预交记录.冲预交%Type := Null,
                strSql = strSql & "" & 0 & ","
                '    退支票额_In   病人预交记录.冲预交%Type := Null,
                strSql = strSql & "NULL,"
                '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
                strSql = strSql & "" & lngCardTypeID & ","
                '    卡号_In       病人预交记录.卡号%Type := Null,
                strSql = strSql & "'" & Trim(varTemp(5)) & "',"
                '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
                strSql = strSql & "'" & Trim(varTemp(6)) & "',"
                '    交易说明_In   病人预交记录.交易说明%Type := Null,
                strSql = strSql & "'" & Trim(varTemp(7)) & "',"
                '    缴款_In       病人预交记录.缴款%Type := Null,
                strSql = strSql & "" & dbl缴款 & ","
                '    找补_In       病人预交记录.找补%Type := Null,
                strSql = strSql & "" & dbl找补 & ","
                '    误差金额_In   门诊费用记录.实收金额%Type := Null,
                '    -- 误差金额_In:存在误差费时,传入
                strSql = strSql & "" & 0 & ","
                '    完成结算_In Number:=0
                '    -- 完成结算_In:1-完成收费;0-未完成收费
                strSql = strSql & "0)"
                zlAddArray cllOutPro, strSql
            End If
        Next
    End If
    If dblMoney <> dbl结算金额 Then
        strErrMsg = "收费的单据金额不等于结算金额!": Exit Function
    End If
    If dbl冲预交款 <> 0 And dbl冲预交款 > Val(Nvl(rsInfo!余额)) Then
        strErrMsg = "预交款余额不足!": Exit Function
    End If
        
    If str医保结算 <> "" Then
        str医保结算 = Mid(str医保结算, 3)
        '消费卡结算
        ' Zl_门诊收费结算_Modify
        strSql = "Zl_门诊收费结算_Modify("
        strSql = strSql & "2" & ","
        '    病人id_In     门诊费用记录.病人id%Type,
        strSql = strSql & lng病人ID & ","
        '    结帐id_In     病人预交记录.结帐id%Type,
        strSql = strSql & lng结帐ID & ","
        '    结算方式_In   Varchar2,
        strSql = strSql & "'" & str医保结算 & "',"
        '    冲预交_In     病人预交记录.冲预交%Type := Null,
        strSql = strSql & "NULL,"
        '    退支票额_In   病人预交记录.冲预交%Type := Null,
        strSql = strSql & "NULL,"
        '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "NULL,"
        '    卡号_In       病人预交记录.卡号%Type := Null,
        strSql = strSql & "NULL,"
        '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "NULL,"
        '    交易说明_In   病人预交记录.交易说明%Type := Null,
        strSql = strSql & "NULL,"
        '    缴款_In       病人预交记录.缴款%Type := Null,
        strSql = strSql & "" & dbl缴款 & ","
        '    找补_In       病人预交记录.找补%Type := Null,
        strSql = strSql & "" & dbl找补 & ","
        '    误差金额_In   门诊费用记录.实收金额%Type := Null,
        '    -- 误差金额_In:存在误差费时,传入
        strSql = strSql & "NULL,"
        '    完成结算_In Number:=0
        '    -- 完成结算_In:1-完成收费;0-未完成收费
        strSql = strSql & "0)"
        zlAddArray cllOutPro, strSql
    End If
    If str收费结算 <> "" Or str收费结算 = "" Then
       If str收费结算 <> "" Then str收费结算 = Mid(str收费结算, 3)
        '消费卡结算
        ' Zl_门诊收费结算_Modify
        strSql = "Zl_门诊收费结算_Modify("
        strSql = strSql & "0" & ","
        '    病人id_In     门诊费用记录.病人id%Type,
        strSql = strSql & lng病人ID & ","
        '    结帐id_In     病人预交记录.结帐id%Type,
        strSql = strSql & lng结帐ID & ","
        '    结算方式_In   Varchar2,
        strSql = strSql & IIf(str收费结算 = "", "NULL", "'" & str收费结算 & "'") & ","
        '    冲预交_In     病人预交记录.冲预交%Type := Null,
        strSql = strSql & IIf(dbl冲预交款 = 0, "NULL", dbl冲预交款) & ","
        '    退支票额_In   病人预交记录.冲预交%Type := Null,
        strSql = strSql & "NULL,"
        '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "NULL,"
        '    卡号_In       病人预交记录.卡号%Type := Null,
        strSql = strSql & "NULL,"
        '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "NULL,"
        '    交易说明_In   病人预交记录.交易说明%Type := Null,
        strSql = strSql & "NULL,"
        '    缴款_In       病人预交记录.缴款%Type := Null,
        strSql = strSql & "" & dbl缴款 & ","
        '    找补_In       病人预交记录.找补%Type := Null,
        strSql = strSql & "" & dbl找补 & ","
        '    误差金额_In   门诊费用记录.实收金额%Type := Null,
        '    -- 误差金额_In:存在误差费时,传入
        strSql = strSql & IIf(dbl误差费 = 0, "NULL", dbl误差费) & ","
        '    完成结算_In Number:=0
        '    -- 完成结算_In:1-完成收费;0-未完成收费
        strSql = strSql & "1)"
        zlAddArray cllOutPro, strSql
    End If
    If blnExcuteCharge Then
        zlExecuteProcedureArrAy cllOutPro, "执行门诊收费", False, False
    End If
    zlClinicCharge = True
End Function

Public Function zlClinicChargeThreeSwapUpdate(ByVal lngModule As Long, ByVal lng结帐ID As String, _
    ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, str交易流水号 As String, str交易说明 As String, ByVal str其他交易信息 As String, _
    ByRef cllOutPro As Collection, _
    ByRef strErrMsg As String, Optional ByVal blnExcuteCharge As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊收费结算
    '入参:lng结帐ID-结帐ID
    '     lng卡类别ID-卡类别ID
    '     str其他交易信息-格式:项目名称|项目内容||...
    '     str交易流水号和str交易说明:都为空时,不更新病人预交记录.交易流水号和交易说明
    '出参:lng结帐ID-返回结帐ID
    '     cllOutPro-返回需要执行的SQL
    '     strErrMsg-返回False时,返回的出错信息
    '返回:执行成功或获取保存SQL成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-26 16:23:16
    '说明:该过程中不处理异常,主要由调用方处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lng卡类别ID = 0 Then Exit Function
    If str交易流水号 <> "" Or str交易说明 <> "" Then
        Call zlAddUpdateSwapSQL(False, lng结帐ID, lng卡类别ID, bln消费卡, str卡号, str交易流水号, str交易说明, cllOutPro)
    End If
    If str其他交易信息 <> "" Then
        Call zlAddThreeSwapSQLToCollection(False, lng结帐ID, lng卡类别ID, bln消费卡, str卡号, str其他交易信息, cllOutPro)
    End If
    If blnExcuteCharge Then
        zlExecuteProcedureArrAy cllOutPro, "执行三方交易信息更正", False, False
    End If
    zlClinicChargeThreeSwapUpdate = True
End Function
Public Function zlGetPayDrugWindow(ByVal lngModule As Long, ByVal lng病人ID As Long, ByVal dt收费时间 As Date, _
    ByVal cllDept As Collection, ByRef strPayWins As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发药窗口
    '入参:lng病人ID-病人ID
    '     dt收费时间-收费时间
    '     cllDept-具体执行部门:array(收费类别,执行部门ID,发药窗口)
    '出参:strPayWins-返回发药窗口集,格式为:执行部门ID|发药窗口;....
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-23 18:18:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发药窗口 As String, strPayDrugWins As String
    Dim str西窗 As String, str中窗 As String, str成窗 As String
    Dim i As Long, varData As Variant
    Err = 0: On Error GoTo Errhand:
    
    strPayDrugWins = ""
    For i = 1 To cllDept.Count
        varData = cllDept(i)
        str发药窗口 = varData(2)
        If str发药窗口 = "" Then
            '判断当前病人是否存在相同执行部门的未发药品，若存在则返回未发药品的发药窗口
            str发药窗口 = Get未发药品发药窗口(lng病人ID, Val(varData(1)))
            If str发药窗口 = "" Then str发药窗口 = zlGetDrugWindow(lngModule, Val(varData(1)), Trim(varData(0)))
            If str发药窗口 = "" Then
                str发药窗口 = zlGet发药窗口(lngModule, dt收费时间, Val(varData(1)), Trim(varData(0)), str西窗, str成窗, str中窗)
            End If
        End If
        If InStr(1, strPayDrugWins & ";", ";" & Val(varData(1)) & "|") = 0 Then
            strPayDrugWins = strPayDrugWins & ";" & Val(varData(1)) & "|" & str发药窗口
        End If
    Next
    strPayWins = IIf(strPayDrugWins <> "", Mid(strPayDrugWins, 2), "")
    zlGetPayDrugWindow = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlCheckRegBillIsExecuted(ByVal strNO As String, ByVal bln不含作废医嘱 As Boolean, _
    ByRef blnExecuted_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断指定的挂号单据是否已经被执行,包括医生接诊下医嘱后作废后,取消接诊,也表示执行过了
    '入参:strNO-挂号单号
    '     bln不含作废医嘱-是否不包含已作废的医嘱
    '出参:
    '返回:True 表示已被执行
    '编制:刘兴洪
    '日期:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    blnExecuted_Out = False
    If bln不含作废医嘱 Then strSql = " And 医嘱状态<>4"
    strSql = _
        " Select count(ID) num From 病人挂号记录 Where NO=[1] And 执行状态>0 and 记录性质=1 and 记录状态 =1 " & _
        " Union All " & _
        " Select count(ID) num From 病人医嘱记录 Where 挂号单=[1] And (病人来源=1 or 病人来源=2)" & strSql
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", strNO)
    Do While Not rsTmp.EOF
        If rsTmp!Num > 0 Then
            blnExecuted_Out = True
        End If
        rsTmp.MoveNext
    Loop
    zlCheckRegBillIsExecuted = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegistLockNo(ByVal lngModule As Long, ByVal byt操作类型 As Byte, _
    ByVal str号码 As String, ByVal dat日期 As Date, ByVal int序号 As Integer, ByRef cllOutPro As Collection, _
    ByRef strErrMsg As String, Optional ByVal blnExcuteLock As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号的锁号/解锁操作
    '入参:lngModule-模块号
    '     byt操作类型-操作类型 0-锁号 1-解锁
    '     dat日期-锁号/解锁的日期,格式:YYYY-MM-DD
    '     str号码 -挂号安排的号码
    '     int序号 -选择锁定 / 解锁的序号
    '     blnExcuteLock-是否执行操作(缺省为true,表示调用该接口直接执行操作;否则返回操作的SQL语句)
    '出参:cllOutPro-返回需要执行的SQL
    '     strErrMsg -返回False时, 返回的出错信息
    '返回:True 表示已被执行
    '编制:刘兴洪
    '日期:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTmpSub As ADODB.Recordset
    Dim strSql As String, lng安排ID As Long, lng计划ID As Long, str日期 As String
    On Error GoTo errH
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    
    If byt操作类型 = 0 Then
        strSql = "Select ID" & vbNewLine & _
                 "      From 挂号安排计划" & vbNewLine & _
                 "      Where 号码 = [1] And [2] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                 "            失效时间 " & vbNewLine & _
                 "      Order By 生效时间 Desc "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码, dat日期)
        If rsTmp.RecordCount = 0 Then
            strSql = "Select ID From 挂号安排 Where 号码 = [1]"
            Set rsTmpSub = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码)
            If rsTmpSub.RecordCount = 0 Then
                strErrMsg = "没有找到号码为" & str号码 & "的挂号安排,请检查！"
                Exit Function
            Else
                lng安排ID = Val(rsTmpSub!ID)
            End If
        Else
            lng计划ID = Val(rsTmp!ID)
        End If
        If lng计划ID <> 0 Then
            strSql = "Select 开始时间" & vbNewLine & _
                     "From 挂号计划时段" & vbNewLine & _
                     "Where 计划id = [1] And 序号 = [2] And 星期 = Decode(To_Char([3], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
                     "                                             '周四', '6', '周五', '7', '周六', Null)"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng计划ID, int序号, dat日期)
            If rsTmp.RecordCount = 0 Then
                strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & "'," & "to_date('" & Format(dat日期, "yyyy-MM-dd") & "','YYYY-MM-DD')," & int序号 & ")"
            Else
                str日期 = Format(dat日期, "yyyy-MM-dd") & " " & Format(rsTmp!开始时间, "hh:mm:ss")
                strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & "'," & "to_date('" & Format(str日期, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & int序号 & ")"
            End If
        Else
            If lng安排ID <> 0 Then
                strSql = "Select 开始时间" & vbNewLine & _
                         "From 挂号安排时段" & vbNewLine & _
                         "Where 安排id = [1] And 序号 = [2] And 星期 = Decode(To_Char([3], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
                         "                                             '周四', '6', '周五', '7', '周六', Null)"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng安排ID, int序号, dat日期)
                If rsTmp.RecordCount = 0 Then
                    strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & "'," & "to_date('" & Format(dat日期, "yyyy-MM-dd") & "','YYYY-MM-DD')," & int序号 & ")"
                Else
                    str日期 = Format(dat日期, "yyyy-MM-dd") & " " & Format(rsTmp!开始时间, "hh:mm:ss")
                    strSql = "Zl_挂号序号状态_Lock(1,'" & UserInfo.姓名 & "','" & str号码 & "'," & "to_date('" & Format(str日期, "yyyy-MM-dd hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & int序号 & ")"
                End If
            End If
        End If
        If blnExcuteLock Then
            gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "挂号锁号"
        Else
            Call zlAddArray(cllOutPro, strSql)
        End If
    Else
        strSql = "Zl_挂号序号状态_Lock(2,'" & UserInfo.姓名 & "','" & str号码 & "'," & "to_date('" & dat日期 & "','yyyy-mm-dd hh24:mi:ss')," & int序号 & ")"
        If blnExcuteLock Then
            gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "挂号解锁"
        Else
            Call zlAddArray(cllOutPro, strSql)
        End If
    End If
    zlRegistLockNo = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegistDel(ByVal lngModule As Long, ByVal strNO As String, _
     ByRef cllOutPro As Collection, ByRef strErrMsg As String, _
    Optional ByVal bln退病历 As Boolean = True, Optional ByVal blnExcuteDel As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:挂号的退号操作
    '入参:lngModule-模块号
    '     strNO -退号记录的单据号
    '     bln退病历 -是否退病历费, 缺省为True
    '     blnExcuteDel-是否执行操作(缺省为true,表示调用该接口直接执行操作;否则返回操作的SQL语句)
    '出参:cllOutPro-返回需要执行的SQL
    '     strErrMsg -返回False时, 返回的出错信息
    '返回:True 表示已被执行
    '编制:刘兴洪
    '日期:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intReuse As Integer
    On Error GoTo errH
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    strSql = "Select 1" & vbNewLine & _
             "From 费用补充记录 A," & vbNewLine & _
             "     (Select Distinct 结帐id" & vbNewLine & _
             "       From 门诊费用记录" & vbNewLine & _
             "       Where NO = [1] And 记录性质 = 4" & vbNewLine & _
             "       Union" & vbNewLine & _
             "       Select Distinct 结帐id From 住院费用记录 Where NO = [1] And 记录性质 = 5) B" & vbNewLine & _
             "Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 1 And Nvl(a.费用状态, 0) <> 2 And Rownum < 2"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", strNO)
    If rsTmp.RecordCount <> 0 Then
         strErrMsg = "传入的挂号单据" & strNO & "已经进行了二次结算,无法退号!"
         Exit Function
    End If
    intReuse = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111, "1"))
    strSql = "Zl_病人挂号记录_Delete('" & strNO & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',Null, Null, Null, " & IIf(bln退病历, 0, 1) & ", Null, " & intReuse & ")"
    If blnExcuteDel Then
        gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & "挂号退号"
    Else
        Call zlAddArray(cllOutPro, strSql)
    End If
    zlRegistDel = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function zlRegistSave(ByVal lngModule As Long, ByVal byt操作类型 As Byte, _
     ByVal lng病人ID As Long, ByVal str号码 As String, ByVal dat发生时间 As Date, ByVal str结算方式 As String, ByVal str金额 As String, _
     ByRef cllOutPro As Collection, ByRef strErrMsg As String, ByRef strOutNO As String, _
     Optional ByVal int序号 As Integer, Optional ByVal str合作单位 As String, Optional ByVal strNO As String, _
     Optional ByVal lng结帐ID As Long, Optional ByVal lng卡类别ID As Long, Optional ByVal bln消费卡 As Boolean, _
      Optional ByVal str卡号 As String, Optional ByVal str交易流水号 As String, Optional ByVal str交易说明 As String, _
     Optional ByVal blnExcuteSave As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存挂号记录
    '入参:lngModule-模块号
    '     byt操作类型-操作类型:0-挂号 1-预约 2-接收
    '     lng病人ID -病人ID
    '     str号码 -挂号安排的号码
    '     dat发生时间-挂号的发生时间,也可只传日期,此时具体时间自动计算
    '     str结算方式 -支付金额的结算方式
    '     str金额-支付金额|预交支付金额|个人账户支付金额
    '     int序号 -选择的序号, 可以不传
    '     str合作单位-可以不传，合作单位预约/挂号时传入
    '     strNO -挂号记录的单据号, 接收时必传, 挂号时可以不传
    '     lng结帐ID -结帐ID, 可以不传
    '     lng卡类别ID -卡类别ID, 可以不传
    '     bln消费卡 -是否消费卡, 可以不传
    '     str卡号 -卡号, 可以不传
    '     str交易流水号 -交易流水号, 可以不传
    '     str交易说明 -交易说明, 可以不传
    '     blnExcuteSave-是否执行保存操作(缺省为true,表示调用该接口直接挂号;否则返回挂号的SQL语句)
    '出参:strOutNO-挂号产生的单据号
    '     cllOutPro-返回需要执行的SQL
    '     strErrMsg -返回False时, 返回的出错信息
    '返回:True 表示已被执行
    '编制:刘兴洪
    '日期:2014-10-10 11:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, lng安排ID As Long, lng计划ID As Long, rsTmpSub As ADODB.Recordset
    Dim strSql As String, intReuse As Integer, bln序号控制 As Boolean, dat当前时间 As Date
    Dim int限号数 As Integer, int限约数 As Integer, strMoney() As String, str保险结算 As String
    Dim int退号重用 As Integer, bln挂号生成队列 As Boolean
    Dim dbl现金 As Double, dbl预交 As Double, dbl个帐 As Double
    On Error GoTo errH
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    dat当前时间 = gobjDatabase.Currentdate
    If byt操作类型 = 2 Then
        bln挂号生成队列 = Val(gobjDatabase.GetPara("排队叫号模式", glngSys, 1113)) <> 0
        If lng结帐ID = 0 Then lng结帐ID = gobjDatabase.GetNextId("病人结帐记录")
        strSql = "Select 门诊号, 姓名, 性别, 年龄, 医疗付款方式, 费别, 险类 From 病人信息 Where 病人ID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng病人ID)
        If rsTmp.RecordCount = 0 Then
            strErrMsg = "未发现传入病人ID为" & lng病人ID & "的病人信息!"
            Exit Function
        End If
        strMoney = Split(str金额 & "|||", "|")
        dbl现金 = Val(strMoney(0))
        dbl预交 = Val(strMoney(1))
        dbl个帐 = Val(strMoney(2))
        '挂号接收
        strSql = "ZL_预约挂号接收_INSERT('" & strNO & "',Null,Null," & _
                 lng结帐ID & ",'" & GetRoom(str号码) & "'," & ZVal(lng病人ID) & "," & Nvl(rsTmp!门诊号, "Null") & ",'" & rsTmp!姓名 & "'," & _
                 "'" & rsTmp!性别 & "','" & rsTmp!年龄 & "','" & rsTmp!医疗付款方式 & "'," & _
                 "'" & rsTmp!费别 & "','" & str结算方式 & "'," & dbl现金 & "," & dbl预交 & "," & dbl个帐 & "," & _
                 dat发生时间 & "," & ZVal(int序号) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(bln挂号生成队列, 1, 0) & "," & _
                 dat当前时间 & ","  '问题号:48350
        '卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSql = strSql & "" & lng卡类别ID & ",Null,"
        '卡号_In       病人预交记录.卡号%Type := Null,
        strSql = strSql & "'" & str卡号 & "',"
        '交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSql = strSql & "'" & str交易流水号 & "',"
        '交易说明_In   病人预交记录.交易说明%Type := Null
        strSql = strSql & "'" & str交易说明 & "',"
        '险类_In       病人挂号记录.险类%Type := Null,
        strSql = strSql & "" & IIf(dbl个帐 = 0, "Null", "'" & rsTmp!险类 & "'") & ")"
        Call zlAddArray(cllOutPro, strSql)
        '处理汇总
        strSql = "" & _
               " Select B.科室id, B.项目id, B.医生id, B.医生姓名,B.号码 " & _
               " From 门诊费用记录 A, 挂号安排 B " & _
               " Where A.记录性质 = 4 And A.记录状态 = 0 And A.NO = [1] And A.序号 = 1 And A.计算单位 = B.号码"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", strNO)
        If rsTmp.EOF = False Then
            strSql = "zl_病人挂号汇总_Update("
            '  医生姓名_In   挂号安排.医生姓名%Type,
            strSql = strSql & "'" & Nvl(rsTmp!医生姓名) & "',"
            '  医生id_In     挂号安排.医生id%Type,
            strSql = strSql & "" & ZVal(Val(Nvl(rsTmp!医生ID))) & ","
            '  收费细目id_In 门诊费用记录.收费细目id%Type,
            strSql = strSql & "" & Val(Nvl(rsTmp!项目ID)) & ","
            '  执行部门id_In 门诊费用记录.执行部门id%Type,
            strSql = strSql & "" & Val(Nvl(rsTmp!科室ID)) & ","
            '  发生时间_In   门诊费用记录.发生时间%Type,
            strSql = strSql & "" & dat发生时间 & ","
            '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收
            strSql = strSql & "2" & ","
            '  号码_In       挂号安排.号码%Type := Null
            strSql = strSql & "'" & Nvl(rsTmp!号码) & "')"
            Call zlAddArray(cllOutPro, strSql)
        End If
        If blnExcuteSave Then
            zlExecuteProcedureArrAy cllOutPro, "预约挂号接收", False, False
        End If
    Else
        int退号重用 = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111, "1"))
        '挂号/预约
        strSql = "" & _
                "Select ID" & vbNewLine & _
                "From 挂号安排计划" & vbNewLine & _
                "Where 号码 = [1] And [2] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                "      失效时间 And Rownum < 2" & vbNewLine & _
                "Order By 生效时间 Desc"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码, dat发生时间)
        If rsTmp.RecordCount = 0 Then
            strSql = "Select ID From 挂号安排 Where 号码 = [1]"
            Set rsTmpSub = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码)
            If rsTmpSub.RecordCount = 0 Then
                lng安排ID = Nvl(rsTmpSub!ID)
            Else
                strErrMsg = "没有发现传入号别" & str号码 & "的挂号安排!"
                Exit Function
            End If
        Else
            lng计划ID = Nvl(rsTmp!ID)
        End If
        '合作单位检查
        If str合作单位 <> "" Then
            If lng计划ID <> 0 Then
                strSql = "" & _
                        "Select 1" & vbNewLine & _
                        "From 合作单位计划控制" & vbNewLine & _
                        "Where 计划id = [1] And " & _
                        "限制项目 = Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And " & _
                        "(序号 = n_号序 Or 序号 = 0) And 合作单位 = [3]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng计划ID, dat发生时间, str合作单位)
                If rsTmp.RecordCount = 0 Then
                    strErrMsg = "没有发现合作单位" & str合作单位 & "的挂号安排!"
                    Exit Function
                End If
            Else
                strSql = "" & _
                        "Select 1" & vbNewLine & _
                        "From 合作单位安排控制" & vbNewLine & _
                        "Where 安排id = [1] And " & _
                        "限制项目 = Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And " & _
                        "(序号 = n_号序 Or 序号 = 0) And 合作单位 = [3]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng安排ID, dat发生时间, str合作单位)
                If rsTmp.RecordCount = 0 Then
                    strErrMsg = "没有发现合作单位" & str合作单位 & "的挂号安排!"
                    Exit Function
                End If
            End If
        End If
        '未传入序号的情况
        If int序号 = 0 Then
            If lng计划ID <> 0 Then
                strSql = "Select Nvl(序号控制,0) As 序号控制 From 挂号安排计划 Where 计划ID = [1]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng计划ID)
            Else
                strSql = "Select Nvl(序号控制,0) As 序号控制 From 挂号安排 Where 安排ID = [1]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng安排ID)
            End If
            bln序号控制 = Val(rsTmp!序号控制) = 1
            If bln序号控制 Then
                '获取序号
                If lng计划ID <> 0 Then
                    strSql = "" & _
                            "Select A.序号, To_Char([2], 'yyyy/mm/dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss') As 开始时间" & vbNewLine & _
                            "From 挂号计划时段 A" & vbNewLine & _
                            "Where a.计划id = [1] And a.星期 = Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6'," & vbNewLine & _
                            "                                      '周五', '7', '周六', Null) And Not Exists" & vbNewLine & _
                            " (Select 1 From 挂号序号状态 Where 号码 = [3] And Trunc(日期) = Trunc([2]) And 序号 = a.序号 And 状态 <> 5) Order By 序号"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng计划ID, dat发生时间, str号码)
                    If rsTmp.RecordCount = 0 Then
                        If dat当前时间 > dat发生时间 Then dat发生时间 = dat当前时间
                        strSql = "Select Nvl(Max(序号), 0) As 序号 From 挂号序号状态 Where 号码 = [1] And Trunc(日期) = Trunc([2]) And 状态 <> 5"
                        Set rsTmpSub = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码, dat发生时间)
                        int序号 = Val(rsTmpSub!序号) + 1
                    Else
                        int序号 = rsTmp!序号
                        dat发生时间 = rsTmp!开始时间
                    End If
                Else
                    strSql = "" & _
                            "Select A.序号, To_Char([2], 'yyyy/mm/dd') || ' ' || To_Char(开始时间, 'hh24:mi:ss') As 开始时间" & vbNewLine & _
                            "From 挂号安排时段 A" & vbNewLine & _
                            "Where a.安排id = [1] And a.星期 = Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6'," & vbNewLine & _
                            "                                      '周五', '7', '周六', Null) And Not Exists" & vbNewLine & _
                            " (Select 1 From 挂号序号状态 Where 号码 = [3] And Trunc(日期) = Trunc([2]) And 序号 = a.序号 And 状态 <> 5) Order By 序号"
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", lng安排ID, dat发生时间, str号码)
                    If rsTmp.RecordCount = 0 Then
                        If dat当前时间 > dat发生时间 Then dat发生时间 = dat当前时间
                        strSql = "Select Nvl(Max(序号), 0) As 序号 From 挂号序号状态 Where 号码 = [1] And Trunc(日期) = Trunc([2]) And 状态 <> 5"
                        Set rsTmpSub = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent", str号码, dat发生时间)
                        int序号 = Val(rsTmpSub!序号) + 1
                    Else
                        int序号 = rsTmp!序号
                        dat发生时间 = rsTmp!开始时间
                    End If
                End If
            End If
        End If
        strMoney = Split(str金额 & "|||", "|")
        dbl现金 = Val(strMoney(0))
        dbl预交 = Val(strMoney(1))
        dbl个帐 = Val(strMoney(2))
        If dbl个帐 <> 0 Then
            strSql = "Select 名称 From 结算方式 Where 性质=3"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlRegEvent")
            If rsTmp.RecordCount = 0 Then
                str保险结算 = "个人帐户|" & dbl个帐
            Else
                str保险结算 = rsTmp!名称 & dbl个帐
            End If
        End If
        If strNO = "" Then strNO = gobjDatabase.GetNextNo(12)
        strSql = "Zl_三方机构挂号_Insert(" & _
                 IIf(byt操作类型 = 0, 1, 2) & "," & lng病人ID & "," & "'" & str号码 & "'," & int序号 & "," & _
                 "'" & strNO & "',Null,'" & str结算方式 & "',Null,to_date('" & dat发生时间 & "','yyyy-mm-dd hh24:mi:ss')," & _
                 "to_date('" & dat当前时间 & "','yyyy-mm-dd hh24:mi:ss'),'" & str合作单位 & "'," & dbl现金 + dbl预交 + dbl个帐 & "," & _
                 "Null,Null,'" & str交易流水号 & "','" & str交易说明 & "',Null,Null," & lng卡类别ID & ",0,0," & lng结帐ID & "," & _
                 "0,'" & str保险结算 & "'," & IIf(dbl预交 = 0, "Null", dbl预交) & ",'" & str卡号 & "'," & int退号重用 & ")"
        If blnExcuteSave Then
            gobjDatabase.ExecuteProcedure strSql, gstrCompentsName & ":" & IIf(byt操作类型 = 0, "挂号", "预约")
        Else
            Call zlAddArray(cllOutPro, strSql)
        End If
        strOutNO = strNO
    End If
    
    zlRegistSave = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function zlBillChargeOff(ByVal strNO As String, Optional ByVal blnExcuteDel As Boolean = True, _
    Optional ByRef cllOutPro As Collection, Optional ByVal strOperatorCode As String, Optional strOperatorName As String, _
    Optional ByVal blnNotBeginTrans As Boolean, Optional ByVal blnNotCommitTrans As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记帐单或记帐表的销帐处理
    '入参: strNO-记帐单
    '    blnExcuteDel-是否执行删除(缺省为true,表示调用该接口直接执行删除;否则返回删除的SQL语句)
    '    strOperatorName-操作员姓名(不传入,使用当前操作员姓名)
    '    strOperatorCode-操作员编号(不传入,使用当前操作员编号)
    '    blnNotBeginTrans-不进行开始事务处理，调用者已经开启事务
    '    blnNotCommitTrans-不提交事务(blnExcuteDel=true时，该参数有效)
    '出参:cllOutPro-删除的Sql集
    '返回:删除或执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-18 13:51:06
    '说明:
    '    1.如果blnExcuteDel=false时，将不进行医保明细上传处理
    '    2.如果blnNotCommit=true且如果blnExcuteDel=true时，医保为"support记帐完成后上传"时，将不调用医保上传接口
    '    3.如果是记帐表销帐时，如果存在一个以上医保上传明细成功时，则销帐成功，否则销帐失败!
    '    4.如果本zlBillChargeOff返回False,则回退了事务处理的。
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInsures As String, arrInsure As Variant, str申请IDs As String, str申请人 As String
    Dim blnTrans As Boolean, strSql As String, strInsure As String
    Dim i As Long
    
    If cllOutPro Is Nothing Then Set cllOutPro = New Collection
    If strOperatorCode = "" Or strOperatorName = "" Then
        strOperatorCode = UserInfo.编号
        strOperatorName = UserInfo.姓名
    End If
        
    Call zlCheckIsExistsApplied(strNO, "", str申请IDs, str申请人)
    If str申请IDs <> "" Then
        strSql = "zl_病人费用销帐_Delete('" & str申请IDs & "')"
        zlAddArray cllOutPro, strSql
    End If
    
    If blnExcuteDel Then
        If GetBillInsures(strInsure, strNO) = False Then Exit Function
    End If
    
    'Zl_住院记帐记录_Delete
    strSql = "Zl_住院记帐记录_Delete("
    '  No_In           住院费用记录.No%Type,
    strSql = strSql & "'" & strNO & "',"
    '  序号_In         Varchar2,
    strSql = strSql & "NULL,"
    '  操作员编号_In   住院费用记录.操作员编号%Type,
    strSql = strSql & "'" & strOperatorCode & "',"
    '  操作员姓名_In   住院费用记录.操作员姓名%Type,
    strSql = strSql & "'" & strOperatorName & "')"
    '  记录性质_In     住院费用记录.记录性质%Type := 2,
    '  操作状态_In     Number := 0,
    '  输液配药检查_In Number:=1
    Call zlAddArray(cllOutPro, strSql)
    If blnExcuteDel Then
    
        blnTrans = True
        zlExecuteProcedureArrAy cllOutPro, "执行住院销帐", Not blnNotBeginTrans, False
        If strInsure <> "" Then
            arrInsure = Split(strInsure, ",")
            For i = 0 To UBound(arrInsure)
                If gclsInsure.GetCapability(support记帐作废上传, , Val(arrInsure(i))) _
                    And Not gclsInsure.GetCapability(support记帐完成后上传, , Val(arrInsure(i))) Then
                    If Not gclsInsure.TranChargeDetail(2, strNO, 2, 2, "", , Val(arrInsure(i))) Then
                        If i = 0 Then gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            Next
        End If
        
        If Not blnNotCommitTrans Then              '提交处理
          gcnOracle.CommitTrans
          If strInsure <> "" Then
                For i = 0 To UBound(arrInsure)
                     If gclsInsure.GetCapability(support记帐作废上传, , Val(arrInsure(i))) _
                        And gclsInsure.GetCapability(support记帐完成后上传, , Val(arrInsure(i))) Then
                         If Not gclsInsure.TranChargeDetail(2, strNO, 2, 2, "", , arrInsure(i)) Then
                             MsgBox "单据中 " & GetInsureName(Val(arrInsure(i))) & " 的费用向医保传送失败，这些费用已销帐。", vbInformation, gstrSysName
                         End If
                     End If
                 Next
          End If
        End If
    End If
    zlBillChargeOff = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    Call gobjComlib.ErrCenter
    Call gobjComlib.SaveErrLog
End Function
    


Public Function zlBillChargeOffIsValied(ByVal strNO As String, Optional ByVal blnNotShowErr As Boolean, Optional blnNotInsureCheck As Boolean, _
    Optional ByVal strDelRows As String, _
    Optional strOutErrMsg As String, Optional ByRef blnYesOrNo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院记帐操作销帐前检查销帐数据的合法性(只支持住院销帐的检查)
    '入参:strNo-销帐的单据信息
    '    strDelRows-(格式为:序号1:数量1;序号2：数量2....):暂时无限制,待以后扩展
    '    blnNotShowErr-不显示错误信息(即发生错误时，不用Msgbox来显示错误信息)
    '    blnNotInsureCheck-在销帐时，不检查对应的医保
    '出参:
    '    strOutErrMsg-返回错误信息
    '    blnYesOrNo-是否返回的错误信息显示确定或取消按钮
    '返回:销帐合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-10-17 14:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln门诊 As Boolean, lng病人ID As Long, lng主页Id As Long, i As Long
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim strPrivsOpt As String, blnNOMoved As Boolean
    Dim strPatis As String, intInsure As Integer, strDate As String
    Dim cllInsure As Collection, strTemp As String, intTemp As Integer
    Dim blnNotInsure As Boolean   '存在非医保病人
    
    On Error GoTo errHandle
    strPrivsOpt = GetInsidePrivs(1150)
    
'    " Select 1 as 标志,a.病人ID,0 as 主页ID,a.序号,0 as 多病人单,0 as 险类,max(B.姓名) as 病人姓名,0 as 病案编目,max(a.记录状态) as 记录状态, " & _
'    "       max(decode(A.记录状态,2,0,1)*Nvl(A.执行状态,0)) as 执行状态, Max(A.登记时间) as 登记时间,max(A.操作员姓名) as 操作员," & _
'    "       sum(decode(a.记录状态,2,0,1)*(nvl(a.付数,0)*nvl(a.数次,0)) as 原始数量 ,sum(nvl(a.付数,0)*nvl(a.数次,0)) as 剩余数量" & _
'    " From 门诊费用记录 A,病人信息 B " & _
'    " Where a.病人ID=B.病人ID(+) and  a.记录性质=2 and A.NO=[1] And A.价格父号 is null   " & _
'    " Group by a.病人ID,a.序号" & _
'    " UNION ALL " & _

    '记录状态=1,3时：0:未执行;1:完全执行;2:部份执行；记录状态=2时：-x:第x次退费
    strSql = "" & _
    " Select a.病人ID,a.主页ID,a.序号,max(a.多病人单) as 多病人单,max(nvl(b.险类,0)) as 险类,max(nvl(B.姓名,C.姓名)) as 病人姓名,max(decode(B.编目日期,NULL,0,1)) as 病案编目,max(a.记录状态) as 记录状态," & _
    "      max(decode(A.记录状态,2,0,1)*Nvl(A.执行状态,0)) as 执行状态, Max(A.登记时间) as 登记时间,max(A.操作员姓名) as 操作员," & _
    "      sum(decode(a.记录状态,2,0,1)*(nvl(a.付数,0)*nvl(a.数次,0))) as 原始数量 ,sum(nvl(a.付数,0)*nvl(a.数次,0)) as 剩余数量,sum(decode(a.记录状态,2,0,1)*(nvl(a.付数,0)*nvl(a.数次,0)))-sum(nvl(a.付数,0)*nvl(a.数次,0)) as 冲销数量  " & _
    " From 住院费用记录 A, 病案主页 B,病人信息 C" & _
    " Where a.记录性质=2 and A.NO=[1] And 价格父号 is null   " & _
    "       And A.病人ID=B.病人ID(+) and a.主页ID=B.主页ID(+) and a.病人ID=C.病人ID(+)" & _
    " Group by a.病人ID,a.主页ID,a.序号"
    
    blnYesOrNo = False
    Set cllInsure = New Collection
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "销帐前检查销帐数据的合法性-获取记帐单信息", strNO)
    If rsTemp.EOF Then
       strOutErrMsg = "没有找到对应的记帐单(" & strNO & "),不能执行销帐处理!"
       If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
       Exit Function
    End If
    If Nvl(rsTemp!记录状态, 0) = 0 Then
        strOutErrMsg = "单据(" & strNO & ")为记帐划价单,不能执行销帐处理!"
        If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    

    blnNOMoved = gobjDatabase.NOMoved("住院费用记录", strNO, , 2, "销帐前检查销帐数据的合法性-检查是否转出到历史数据")
    
    '是否执行
    If blnNOMoved Then
       strOutErrMsg = "当前操作的单据:" & strNO & "在后备数据表中,请与系统管理员联系, " & vbCrLf & "将后备数据转入到在线数据表中后再进行销帐操作!"
       If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
       Exit Function
    End If
    
    strDate = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
        
    '操作员权限判断
    If Not BillOperCheck(5, Nvl(rsTemp!操作员), CDate(strDate), "销帐", strNO, , 2, , , , blnNotShowErr, strOutErrMsg) Then Exit Function

    '项目冲销权限
    If Not CheckDelPriv(strNO, strPrivsOpt, strDate, , 0, blnNotShowErr, strOutErrMsg) Then Exit Function
    '留观病人权限
    strTemp = Check留观病人(strNO, strPrivsOpt, strDate)
    If strTemp <> "" Then
        strOutErrMsg = "单据(" & strNO & ")中包含" & strTemp & ",你没有权限对该单据进行操作！"
        If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If

    '出院病人操作权限判断
    If Not BillCanBeOperate(strNO, strPrivsOpt, "销帐", strDate, , 2, blnNotShowErr, strOutErrMsg) Then Exit Function


    
    '执行状态:记录状态=1,3时：0:未执行;1:完全执行;2:部份执行
    rsTemp.Filter = "执行状态=1 or 执行状态=2 "
    If Not rsTemp.EOF Then
        strOutErrMsg = "当前操作的单据" & strNO & "中的内容已经全部或部分完全执行，不允许按单据销帐!"
        If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    rsTemp.Filter = "剩余数量<>0"
    If rsTemp.EOF Then
       strOutErrMsg = "记帐单(" & strNO & ")没有了可销帐的数量,不能执行销帐处理!"
       If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
       Exit Function
    End If
    
    '暂不检查这一步，原因主要是适用于新网接口。
'    If InStr(1, strPrivsOpt, ";部分销帐;") = 0 Then
'        rsTemp.Filter = "剩余数量<>原始数量"
'        If rsTemp.EOF Then
'           strOutErrMsg = "你没有部分销帐的权限，只能对该单据全部销帐！!"
'           If blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
'           Exit Function
'        End If
'    End If
    
    rsTemp.Filter = 0
    '病案编目检查和病人费用变动等检查
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With rsTemp
        strPatis = ""
        Do While Not .EOF
            lng病人ID = Val(Nvl(!病人ID))
            lng主页Id = Val(Nvl(!主页ID))
            intInsure = Val(Nvl(!险类))
            
            '病人是否病案编目检查
            If InStr(1, strPatis & ",", "," & lng病人ID & ",") = 0 Then
                '病人相关检查
                If Val(Nvl(!病案编目)) = 1 Then
                    strOutErrMsg = "病人『" & Nvl(!病人姓名) & " 』已经病案编目,不允许进行销帐操作!"
                    If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                    Exit Function
                End If
                
                '费用变动检查
                If Not zlIsAllowFeeChange(lng病人ID, lng主页Id, , blnNotShowErr, strOutErrMsg) Then Exit Function
                If intInsure = 0 Then blnNotInsure = True
                If intInsure <> 0 And blnNotInsureCheck = False Then
                    cllInsure.Add Array(lng病人ID, lng主页Id, intInsure, CStr(Nvl(!病人姓名))), "K" & lng病人ID
                End If
                strPatis = strPatis & "," & lng病人ID
            End If
            rsTemp.MoveNext
        Loop
    
    End With
    
    intTemp = 0
    'intTemp:是否已经结帐:0-未结帐,1=已全部结帐,2-已部分结帐
    If Not GetBillingBalanceStatu(2, strNO, True, strDate, 2, intTemp) Then Exit Function
    '医保相关检查
    If cllInsure.Count <> 0 Then
        rsTemp.Filter = "原始数量<0"
        If Not rsTemp.EOF Then  '医保不能针对负数记帐进行销帐
            strOutErrMsg = "该单据存在负数记帐记录,不允许进行医保销帐操作！"
            If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
            Exit Function
        End If
        For i = 1 To cllInsure.Count
            lng病人ID = Val(cllInsure(i)(0))
            intInsure = Val(cllInsure(i)(2))
            
            If gclsInsure.GetCapability(support记帐作废上传, lng病人ID, intInsure) Then
                If gclsInsure.GetCapability(support允许部份冲销单据, lng病人ID, intInsure) = False Then
                   rsTemp.Filter = "冲销数量<>0  and 病人ID=" & lng病人ID
                   If rsTemp.EOF = False Then
                        strOutErrMsg = "因为医保处理需要,该单据中的项目必须全部销帐！"
                        If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                        Exit Function
                   End If
                End If
            End If
            If intTemp <> 0 Then
                If Not gclsInsure.GetCapability(support允许冲销已结帐的记帐单据, lng病人ID, intInsure) Then
                    If intTemp = 1 Then '已全部结帐
                        strOutErrMsg = "该医保记帐单据未销帐部分已经结帐,不能销帐！"
                        If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                        Exit Function
                    End If
                    '按理说因为医保病人必须全部结帐,所以应该不会出现这种情况
                    '可能出现于医保与普通病人混合的记帐表,未精确处理
                    strOutErrMsg = "该医保记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！"
                    If Not blnNotShowErr Then MsgBox strOutErrMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End If
    
    If blnNotInsure Then '存在普通病人 需要检查结帐信息
        Select Case gSysPara.bytBillOpt
        Case 0
        Case 1 '提示
            strOutErrMsg = "该记帐单据包含已经结帐的内容,要销帐吗？"
            blnYesOrNo = True
            If Not blnNotShowErr Then
                If MsgBox(strOutErrMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Case 2  '禁止
            If intTemp = 1 Then '已全部结帐
                strOutErrMsg = "该记帐单据未销帐部分已经结帐,不能销帐！"
            Else
                 strOutErrMsg = "该记帐单据包含已经结帐的内容,只能对未结帐部分进行销帐！"
            End If
            If Not blnNotShowErr Then
                MsgBox strOutErrMsg, vbExclamation, gstrSysName
            End If
            Exit Function
        End Select
    End If
    
    '检查销帐申请
    If zlCheckIsExistsApplied(strNO, "", "", strTemp) Then
        strOutErrMsg = "注意:" & vbCrLf & "    单据" & strNO & "中存在申请销帐的项目,销帐后,将会自动取消" & vbCrLf & "申请人的申请项目,是否继续销帐?" & vbCrLf & "申请人如下: " & strTemp
        blnYesOrNo = True
        If Not blnNotShowErr Then
            If MsgBox(strOutErrMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    zlBillChargeOffIsValied = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlTurnToWard_Fee_Query(ByVal frmMain As Object, ByVal byt操作 As Byte, _
    ByVal lng病人ID As Long, ByVal lng主页Id As Long, _
    Optional ByVal lng变动id As Long, _
    Optional ByVal lng原病区id As Long, Optional lng目标病区id As Long, _
    Optional ByRef cllSql As Collection) As Boolean
     '------------------------------------------------------------------------------------
    '功能:转病区费用查询
    '入参:frmMain-调用的主窗体
    '     byt操作- 0-转病区;1-撤消病区;2-转病区申请;3-历史转出查询
    '     cllSQL - 需要执行的SQL语句，0-转病区/1-撤消病区有效；
    '               0-转病区时，在更改变动记录之后执行；1-撤消病区时，在更改变动记录之前执行
    '   1、int操作=0(转病区)时
    '       lng变动id: 原病区的变动记录的ID
    '       lng原病区id：原病区ID
    '       lng目标病区id:目标病区ID
    '   2、int操作=1(撤消病区)时
    '       lng变动id: 恢复的原病区的变动记录的ID
    '       lng原病区id：被撤消的病区ID
    '       lng目标病区id:恢复的原始病区ID
    '  3. int操作=3(历史转出查询)
    '出参:
    '返回:费用查询时操作员点确认返回true,否则返回False
    '-----------------------------------------------------------------------------------
    Dim frmFeeQuery As New frmTurnToWardFeeQuery
    
    On Error GoTo ErrHandler
    zlTurnToWard_Fee_Query = frmFeeQuery.TurnToWard_Fee_Query(frmMain, byt操作, lng病人ID, lng主页Id, _
        lng变动id, lng原病区id, lng目标病区id, cllSql)
    Exit Function
ErrHandler:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub Class_Initialize()
     glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub
