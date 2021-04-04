VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function zlInitCommon(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
      Optional ByVal strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及数据库连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-当前数据库登录用户名
    '返回:初始化成功,返回true,否则返回False
    '编制:刘鹏飞
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDBUser
    If zlGetComLib = False Then Exit Function
    Call GetUserInfo
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = gobjComlib.GetUnitName
    Call InitLocPar
    
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ModiPatiBaseInfo(frmMain As Object, ByVal str模块 As String, Optional ByVal lng病人ID As Long = 0, _
    Optional ByVal lng就诊ID As Long = 0, Optional ByVal int场合 As Integer = 1, Optional ByVal blnBatch As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用主窗体
    '     str模块-调用模块名称,如：门诊挂号,门诊医生工作站
    '     lng病人ID:病人ID=0通过-病人ID等方式手工过滤病人，病人ID<>0直接提取对应病人信息
    '     lng就诊ID=非0:挂号ID或主页ID(程序将自动定位到要修改的某一次住院或就诊)，0表示用户手工选择是门诊还是住院
    '     int场合 1-门诊;2-住院
    '     blnBatch:True:允许切换病人，调整就诊次数.False:调整具体的某次就诊(不能切换病人和选择就诊次数)
    '返回:TRUE或False
    '编制:刘鹏飞
    '日期:2014-07-16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ModiPatiBaseInfo = frmModiPatiBaseInfo.ShowMe(frmMain, lng病人ID, lng就诊ID, int场合, str模块, blnBatch)
End Function

Public Function PatiImageGatherer(frmMain As Object, ByRef strPictureFilePath As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人图片采集
    '入参:frmMain-调用主窗体
    '出参:strPictureFilePath-返回采集图片的路径(包含文件名及后缀)
    '返回:TRUE或False
    '编制:刘鹏飞
    '日期:2014-07-16
    '调用示例:此接口返回TRue。可直接使用LoadPicture(strPictureFilePath)加载图片
    '---------------------------------------------------------------------------------------------------------------------------------------------
    PatiImageGatherer = frmPatiImageGatherer.ShowMe(frmMain, strPictureFilePath)
End Function

Public Function CheckPatiAge(ByVal strAge As String, Optional ByVal strBirthday As String = "", _
    Optional ByVal bytTag As Byte = 0, Optional ByVal strCalcDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人年龄检查
    '入参:
    '   strAge:病人年龄
    '   strBirthDay:出生日期
    '   bytTag:对于zl_Age_Check函数返回的询问类型的信息，是否要强制终止，还是保持询问.0-保持询问,1-禁止
    '   strCalcDate:计算日期,缺省按照当前时间计算
    '返回：TRUE或FALSE，TRUE:继续,FALSE:终止
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CheckPatiAge = CheckAge(strAge, strBirthday, bytTag, strCalcDate)
End Function

Public Function CheckPatiIdcard(ByVal strIdcard As String, Optional strBirthday As String, Optional strAge As String, Optional strSex As String, Optional strErrInfo As String, Optional datCalc As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：身份证号码合法性校验
    '入参：strIdCard 身份证号码
    '出参：strBirthday  函数返回True为出生日期
    '         strSex 函数返回True为性别
    '         strErrInfo 函数返回False为错误信息
    '         datCalc 计算日期 缺省则按系统时间计算
    '返回：True/False  身份证合法返回True(可从strBirthday，strSex获取出生日期和性别)，否则返回False(可从strErrInfo获取详细错误信息)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    CheckPatiIdcard = CheckIdcard(strIdcard, strBirthday, strAge, strSex, strErrInfo, datCalc)
End Function

Public Function SavePatiBaseInfo(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal strName As String, ByVal strSex As String, _
    ByVal strAge As String, ByVal strBirthday As String, ByVal str模块 As String, Optional ByVal int场合 As Integer = 1, Optional strInfo As String = "", _
    Optional ByVal blnXWHIS As Boolean = True, Optional ByVal blnEMPI As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：调整病人基本信息(含业务数据的同步调整)
    '入参：lng病人ID-病人ID (不能为空/0)
    '         lng就诊ID-挂号ID或主页ID(可为0)
    '         strName-姓名 (不能为空)
    '         strSex-性别 (不能为空)
    '         strAge-年龄 (不能为空)
    '         strBirthDay-出生日期 (不能为空)
    '         str模块-调用该功能的模块描述，如"门诊挂号"，"检查报到"。
    '         int场合 1-门诊;2-住院(lng就诊ID=0,则默认为1;lng就诊ID<>0,1-lng就诊ID为挂号ID,2-lng就诊ID为主页ID)
    '         blnXWHIS-基本信息调整时是否调用RIS的接口 缺省=True调用（该参数用于避免病人信息中重复调用RIS接口）
    '         blnEMPI-T EMPI平台已经建档，F-EMPI平台未建档
    ' 出参：strInfo:更新成功-信息调整导致的变化信息(返回True); 更新失败-信息调整未成功的原因
    ' 返回：TRUE OR False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNote As String
    
    If lng病人ID = 0 Then strInfo = "调整病人基本信息，必须要明确病人。": Exit Function
    If Trim(strName) = "" Then strInfo = "调整病人基本信息，必须要明确病人姓名。": Exit Function
    If Trim(strSex) = "" Then strInfo = "调整病人基本信息，必须要明确病人性别。": Exit Function
    If IsDate(strBirthday) = False Then strInfo = "调整病人基本信息，必须要明确病人出生日期。": Exit Function
    If Trim(strAge) = "" Then strInfo = "调整病人基本信息，必须要明确病人年龄。": Exit Function
    strInfo = "身份证号更新"
    Err.Clear
    On Error GoTo ErrHand
    SavePatiBaseInfo = SaveBaseInfo(lng病人ID, lng就诊ID, strName, strSex, strAge, strBirthday, str模块, int场合, strInfo, blnXWHIS, blnEMPI)
    Exit Function
ErrHand:
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            strInfo = Split(strNote, "[ZLSOFT]")(1)
        Else
            strInfo = strNote
        End If
    Else
        strInfo = Err.Description
    End If
End Function

Public Function RecalcBirthDay(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人年龄获取病人出生日期
    '入参:strAge:病人年龄,如：23岁、1岁2月
    'strCalcDate-传入计算日期
    '返回:传入的病人年龄格式正确则计算返回出生日期,否则返回空
    '  strMsg-返回警告信息
    '正确年龄格式:X岁[X月]、X月[X天]、X天、X小时[X分钟]
    '    X岁:X不能大于200,X月:X不能大于12,X天:X不能大于31,X小时:X不能大于24,X分钟:X不能大于59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    RecalcBirthDay = RecalcBirth(strAge, strDateOfBirth, strCalcDate, strMsg)
    
End Function

Public Function GetPatiState(ByVal lng病人ID As Long, ByRef byt状态 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人某次住院状态(出院、预出院等)
    '入参:lng病人ID：病人唯一标识
    '出参:byt状态：病人状态:1-在院、2-预出院、3-死亡、4-出院
    '返回:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = " Select b.出院日期, b.状态, b.出院方式" & vbNewLine & _
        " From 病人信息 a, 病案主页 b" & vbNewLine & _
        " Where a.病人id = b.病人id And a.主页id = b.主页id And a.病人id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取病人当前状态", lng病人ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    If IsNull(rsTemp!出院日期) Then
        If Val(Nvl(rsTemp!状态, 0)) = 3 Then
            byt状态 = 2
        Else
            byt状态 = 1
        End If
    Else
        If Nvl(rsTemp!出院方式) = "死亡" Then
            byt状态 = 3
        Else
            byt状态 = 4
        End If
    End If
    GetPatiState = True
End Function

Public Function CheckPatiExistBackLists(ByVal byt来源 As Byte, ByVal lng病人ID As Long, ByVal byt提示方式 As Byte, ByRef bln是否继续 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查病人是否是特殊病人，并进行询问或禁止。
    '入参:
    '   byt来源：1-预约;2-先诊疗后结算;3－其他(此参数目前未使用为预留参数，后续需要使用)
    '   lng病人ID: 病人唯一标识
    '   byt提示方式：１-提示;2-禁止
    '出参:
    '   bln是否继续：TRUE继续,False终止.
    '返回:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strBackList As String
    Dim i As Integer
    
    strSQL = "Select 加入原因 From 特殊病人 Where 撤消时间 is Null And 病人ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "特殊病人", lng病人ID)
    If rsTmp.EOF Then
        bln是否继续 = True: CheckPatiExistBackLists = True: Exit Function
    End If
    
    With rsTmp
        i = 1
        strBackList = ""
        Do While Not .EOF
            If i > 5 Then Exit Do
            If i = 4 Then
                strBackList = strBackList & "    " & "……" & vbCrLf
            Else
                strBackList = strBackList & "    " & "□ " & Nvl(!加入原因) & vbCrLf
            End If
            i = i + 1
            .MoveNext
        Loop
    End With
    If byt提示方式 = 1 Then
        If MsgBox("此病人在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & strBackList & vbCrLf & " 请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            bln是否继续 = True
        Else
            bln是否继续 = False
        End If
    Else
        Call MsgBox("此病人在特殊病人名单中。" & vbCrLf & vbCrLf & "原因：" & vbCrLf & vbCrLf & strBackList & vbCrLf & " 不能继续！", vbInformation + vbOKOnly, gstrSysName)
        bln是否继续 = False
    End If
    
    CheckPatiExistBackLists = True
End Function

Public Function GetPatiDiagnose(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal byt功能类型 As Byte, _
    Optional ByVal str诊断类型 As String = "", Optional ByVal str记录来源 As String = "", Optional ByRef str诊断信息 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊、住院病人诊断信息
    '入参:
    '   lng病人ID: 病人唯一标识
    '   lng就诊ID: 主页ID、医嘱ID、挂号ID
    '   byt功能类型：1-提取门诊诊断：２-提取指定类型的诊断信息；３－提取指定医嘱对应的诊断；4-提取最后一次诊断(优先顺序为：出院、入院、门诊)
    '   str诊断类型：诊断类型串,多个诊断类型之间用逗号分割,如:2,3,12,13。(可为空)
    '   str记录来源：诊断来源方式，默认为空。多个来源方式之间用逗号分割。如:3,2,1,4。记录来源值域为1-4。
    ' 说明:
    '   1、byt功能类型：根据主页ID可传１,2,4；根据挂号ID可传１；根据医嘱ID查询只能传３。
    '   2、str诊断类型：byt功能类型=1，则str诊断类型只包含1、11,为空相当于诊断类型="1,11"；byt功能类型=2，str诊断类型不能为空；byt功能类型=3，str诊断类型默认为空；byt功能类型=4，str诊断类型默认为空。
    '   3、str记录来源：只有一个来源，表示提取指定来源方式的诊断；包含多个来源，表示根据传入的来源方式优先级提取诊断(优先级循序为参数顺序)；为空表示来源方式=3,2,1,4
    '出参:
    '   str诊断信息：诊断描述+是否疑诊（疑诊显示为(?)），多条诊断之间用逗号分割。
    '返回:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strReturn As String
    Dim arrCode, i As Integer
    
    If InStr(1, ",1,2,3,4,", "," & byt功能类型 & ",") = 0 Then Exit Function
    
    Select Case byt功能类型
        Case 3 '提取指定医嘱对应的诊断
            strSQL = " Select a.诊断描述, a.是否疑诊" & vbNewLine & _
                " From 病人诊断记录 a, 病人诊断医嘱 b" & vbNewLine & _
                " Where a.Id = b.诊断id And b.医嘱id = [1] And 取消时间 Is Null"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "提取指定医嘱对应的诊断", lng就诊ID)
        Case Else
            strSQL = "Select 记录来源, 诊断类型, 诊断次序, 诊断描述, 是否疑诊, Mod(诊断类型, 10) As 大类" & vbNewLine & _
                " From 病人诊断记录" & vbNewLine & _
                " Where 病人id = [1] And 主页id = [2]  And NVL(编码序号,1) = 1 And Instr([3], ',' || 诊断类型 || ',', 1) <> 0" & vbNewLine & _
                " Order By 记录来源, 诊断类型, 诊断次序"
            If byt功能类型 = 1 Then '-门诊诊断
                If str诊断类型 = "" Then str诊断类型 = ",1,11,"
            ElseIf byt功能类型 = 2 Then '-获取指定类型对应的诊断
                If str诊断类型 = "" Then
                    str诊断信息 = "": Exit Function
                Else
                    str诊断类型 = "," & str诊断类型 & ","
                End If
            Else  '-获取病人最后一次诊断
                str诊断信息 = ",1,2,3,11,12,13,"
            End If
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "提取病人诊断", lng病人ID, lng就诊ID, str诊断类型)
            If str记录来源 = "" Then str记录来源 = "3,2,1,4"
            arrCode = Split(str记录来源, ",")
            For i = 0 To UBound(arrCode)
                rsTmp.Filter = "记录来源=" & Val(arrCode(i))
                If Not rsTmp.EOF Then Exit For
            Next i
            
            If byt功能类型 = 4 And Not rsTmp.EOF Then
                strSQL = rsTmp.Filter
                rsTmp.Filter = strSQL & " And 大类=3"
                If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=2"
                If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=1"
            End If
    End Select
    
    strReturn = ""
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!诊断描述) Then
            strReturn = IIf(strReturn = "", "", strReturn & "，") & rsTmp!诊断描述 & IIf(Nvl(rsTmp!是否疑诊, 0) = 1, "（？）", "")
        End If
        rsTmp.MoveNext
    Loop
    
    str诊断信息 = strReturn
    GetPatiDiagnose = True
End Function

Public Function HosPatiInfoCount(ByVal byt类型 As Byte, ByVal lng部门ID As Long, ByVal dtBegin As Date, ByVal dtEnd As Date, _
    ByVal str状态串 As String, ByRef rsPatiCount As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院病人信息统计
    '入参:
    '   byt类型：0-科室，1-病区
    '   lng部门ID: 科室或病区ID
    '   dtBegin: 统计的开始时间?
    '   dtEnd: 统计的结束时间?
    '   Str状态串：要统计那些状态(如：入院、转入)的病人，格式为：0和1拼接成的字符串(0-不统计、1-统计)，分隔符为分号";"。目前支持类型顺序为：入院、入科、转出、出院、死亡、危、重病人。说明：该参数以后可扩展。
    '出参:
    '   rsPatiCount：字段名称和顺序依次为：入院人数、转入科数、转出人数、出院人数、死亡人数、病危人数、病重人数。
    '返回:TRUE/FALSE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, rsPati As New ADODB.Recordset
    Dim strSQL As String
    With rsPati
        If .State = adStateOpen Then .Close
        .Fields.Append "入院人数", adDouble, 18, adFldIsNullable
        .Fields.Append "入科人数", adDouble, 18, adFldIsNullable
        .Fields.Append "转出人数", adDouble, 18, adFldIsNullable
        .Fields.Append "出院人数", adDouble, 18, adFldIsNullable
        .Fields.Append "死亡人数", adDouble, 18, adFldIsNullable
        .Fields.Append "病危人数", adDouble, 18, adFldIsNullable
        .Fields.Append "病重人数", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
'    '入院人数
'    strSQL = "" & _
'            " SELECT SUM(入院) AS 入院,SUM(入科) AS 入科,SUM(转出) AS 转出,SUM(死亡) AS 死亡,SUM(出院) AS 出院,SUM(危) AS 危,SUM(重) AS 重" & _
'            " FROM (" & _
'            "     SELECT SUM(DECODE(开始原因,2,1,0)) AS 入院,SUM(DECODE(开始原因,3,1,15,1,0)) AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病人变动记录" & _
'            "     Where 病区ID = [1] And NVL(附加床位,0)=0" & _
'            "     AND 开始时间 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     Select SUM(1) as 入院,0 AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病人变动记录 a, 病案主页 b" & _
'            "     Where a.病人id = b.病人id And a.主页id = b.主页id And A.病区ID=[1] And A.开始时间 Between [2] And Sysdate And a.开始原因 = 1 And Nvl(a.附加床位, 0) = 0 And" & _
'            "       Nvl(b.状态, 0) <> 1 And Not Exists" & _
'            "       (Select 1 From 病人变动记录 Where 病人id = a.病人id And 主页id = b.主页id And 开始原因 = 2)"
'    strSQL = strSQL & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,SUM(DECODE(终止原因,3,1,15,1,0)) AS 转出,0 AS 死亡,0 AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病人变动记录" & _
'            "     Where 病区ID = [1] And NVL(附加床位,0)=0" & _
'            "     AND 终止时间 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,SUM(DECODE(出院方式,'死亡',1,0)) AS 死亡,SUM(DECODE(出院方式,'死亡',0,1)) AS 出院,0 AS 危,0 AS 重" & _
'            "     From 病案主页 A,病人信息 B" & _
'            "     Where A.病人ID=B.病人ID  And A.当前病区ID = [1]" & _
'            "     AND 出院日期 BETWEEN [2] AND SYSDATE" & _
'            "     Union" & _
'            "     SELECT 0 AS 入院,0 AS 入科,0 AS 转出,0 AS 死亡,0 AS 出院,SUM(DECODE(当前病况,'危',1,0)) AS 危,SUM(DECODE(当前病况,'重',1,0)) AS 重" & _
'            "     From 病案主页 A,病人信息 B,在院病人 C" & _
'            "     Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And NVL(A.状态,0)<>1 And Nvl(A.病案状态,0)<>5 And A.封存时间 is NULL And B.病人ID=C.病人ID " & _
'            "       And B.当前病区ID=C.病区ID And C.病区ID=[1]" & _
'            ")"
    HosPatiInfoCount = True
End Function

Public Sub MakePatiFamily(frmMain As Object, ByVal lng病人ID As Long, ByVal bytFunc As Byte, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用主窗体
    '     lng病人ID:病人ID
    '     bytFunc:1-查看,2-编辑
    '     lngModul -模块号
    '     rsFamily-病人家属
    '返回:
    '编制:余伟节
    '日期:2015-08-10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPatiFamily As Object
    Set objPatiFamily = New frmPatiFamily
    Call objPatiFamily.ShowMe(frmMain, lng病人ID, bytFunc, lngModule)
End Sub

Public Function GetPatiFamily(ByVal lng病人ID As Long, ByRef strInfo As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:获取病人家属关系
'入参:lng病人ID：病人唯一标识
'出参:病人家属关系 格式:关系1:姓名1,姓名2;关系2:姓名1,姓名2
'返回:TRUE/FALSE
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim strLink As String
    
    strSQL = "Select 关系, b.姓名 From 病人家属 A, 病人信息 B Where a.家属id = b.病人id And a.病人id = [1] Order By 关系"


    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "获取病人家属", lng病人ID)
    
    If rsTemp.RecordCount = 0 Then Exit Function
    For i = 1 To rsTemp.RecordCount
        If strLink <> rsTemp!关系 & "" Then
            strInfo = strInfo & ";" & rsTemp!关系 & ":" & rsTemp!姓名
            strLink = rsTemp!关系
        Else
            strInfo = strInfo & "," & rsTemp!姓名
        End If
        rsTemp.MoveNext
    Next
    
    If strInfo <> "" Then strInfo = Mid(strInfo, 2)
    GetPatiFamily = True
End Function

Public Sub ReadPatiDegreeCard(ByRef frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'---------------------------------------------------------------------------------------------------------------------------------------------
'功能:读取病人详情
'入参:lng病人ID：病人唯一标识
'    lng主页ID
'---------------------------------------------------------------------------------------------------------------------------------------------
    frmDegreeCard.mlng主页ID = lng主页ID
    frmDegreeCard.mlng病人ID = lng病人ID
    frmDegreeCard.Show 1, frmMain
End Sub

Public Function InitLis(Optional ByVal blnMsg As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------
'功能:新版LIS公共部件初始化
'blnMsg      初始化失败是否弹出提示框 T-提示;F-不提示
'返回       True=初始化成功,Flase=初始化失败
'-----------------------------------------------------------------------------------------------------------------------------------------------
    InitLis = InitObjLis(blnMsg)
End Function

Public Function ModifyBabyInfo(ByVal lngBID As Long, ByVal lngZID As Long, ByVal strNO As String, ByVal lngBabyID As Long, ByVal strBabyName As String, ByVal strBabySex As String, Optional strErr As String) As Boolean
    '功能        修改新生儿基本信息
    'lngBID      病人id
    'lngZID      主页id
    'strNO       挂号单
    'lngBabyID   婴儿序号
    'strBabyName 婴儿姓名
    'strBabyAge  婴儿性别
    'StrErr             错误消息
    '返回        True=修改成功,Flase=修改失败
    If InitObjLis(False) Then
        ModifyBabyInfo = gobjLIS.ModifyBabyInfo(lngBID, lngZID, strNO, lngBabyID, strBabyName, strBabySex, strErr)
    End If
End Function

Public Function ShowSelect(rsSelect As ADODB.Recordset, ByVal strKey As String, Optional ByVal strTitle As String, Optional ByVal strNote As String, _
    Optional ByVal blnMutilSelect As Boolean = False, Optional ByVal blnSerach As Boolean = False, Optional ByVal strMshWidth As String = "", Optional ByVal blnHideCancel As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------
'功能:病人信息选择器
'
'返回       True=确定,Flase=取消
'-----------------------------------------------------------------------------------------------------------------------------------------------
   ShowSelect = frmListSel.ShowSelect(rsSelect, strKey, strTitle, strNote, blnMutilSelect, blnSerach, strMshWidth, blnHideCancel)
End Function

Private Sub Class_Terminate()
    If Forms.Count = 0 Then
        Call CloseSquareCardObject '关闭卡结算信息
    End If
End Sub
