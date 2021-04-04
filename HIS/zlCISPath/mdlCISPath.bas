Attribute VB_Name = "mdlCISPath"
Option Explicit
Public gobjFile As New FileSystemObject     '文件操作对象
Public gfrmMain As Object                   '导航台窗体
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gcolPrivs As Collection              '记录内部模块的权限
Public gMainPrivs As String                 '调用主界面所具有的权限,注意非内部模块权限
Public gstrPrivs As String
Public gstrSysName As String                '系统名称
Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String            'OEM产品名称
Public glngSys As Long
Public glngModul As Long

Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjEmr As Object                    '新版智能电子病历对象
Public gcolIcons As Collection              '存放所有临床路径图标集
Public gobjPlugIn As Object                 '外挂功能对象
Public gobjLIS As Object                    'LIS公共部件
Public gbln双审核 As Boolean
Public glngHwnd As Long                     '父窗体句柄
Public gblnGetPath As Boolean

'系统参数
Public gstrLike As String   '如果是双向匹配，则为%
Public gint简码 As Integer  '简码匹配方式：0-拼音,1-五笔

'内部应用模块号定义
Public Enum Enum_Inside_Program
    p临床路径管理 = 1078
    p住院记帐操作 = 1150
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    P临床路径应用 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p电子病案查阅 = 1259
    p门诊医生站 = 1260
    p住院医生站 = 1261
    p住院护士站 = 1262
    p医技工作站 = 1263
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p观片工具管理 = 1289
    P门诊路径应用 = 1248
    P门诊路径管理 = 1083
    P门诊路径跟踪 = 1272
End Enum

Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            GetUserInfo = True
        End If
    End If
    gstrDBUser = UserInfo.用户名
End Function

Public Sub InitSysPar()
'功能：初始化系统参数
    gstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    gint简码 = Val(zlDatabase.GetPara("简码方式"))
End Sub

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
'功能：读取当前登录人员或指定人员的人员性质
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    If str姓名 <> "" Then
        strSql = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", str姓名)
    Else
        strSql = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetAdviceDefineText(ByVal str医嘱IDs As String, Optional rsAdvice As ADODB.Recordset) As String
'功能：获取路径项目对应的医嘱定义内容描述串
'参数：rsAdvice=内存记录集，如果传入则不从数据库读取
    Dim rsTmp As ADODB.Recordset
    Dim strFilter As String, lngPre相关ID As Long
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    
    If Not rsAdvice Is Nothing Then
        '生成动态SQL
        For i = 0 To UBound(Split(str医嘱IDs, ","))
            strFilter = strFilter & " Or ID=" & Split(str医嘱IDs, ",")(i)
        Next
        With rsAdvice
            strSql = ""
            .Filter = Mid(strFilter, 5)
            Do While Not .EOF
                strSql = strSql & " Union ALL Select "
                For i = 0 To .Fields.count - 1
                    If Not IsNull(.Fields(i).Value) Then
                        If Rec.IsType(.Fields(i).Type, adVarChar) Then
                            strSql = strSql & "'" & Replace(Replace(.Fields(i).Value, "[", "("), "]", ")") & "'"
                        Else
                            strSql = strSql & .Fields(i).Value '没有日期型
                        End If
                    Else
                        strSql = strSql & "Null"
                    End If
                    strSql = strSql & " As " & .Fields(i).Name & ","
                Next
                strSql = Left(strSql, Len(strSql) - 1) & " From Dual"
                .MoveNext
            Loop
            .Filter = ""
            strSql = "(" & Mid(strSql, 12) & ")"
        End With
    Else
        strSql = "路径医嘱内容"
    End If
    
    strSql = "Select /*+ Rule*/ A.ID,A.相关ID,Decode(A.期效,1,'临嘱','长嘱') as 期效,B.类别," & _
        " Nvl(A.医嘱内容,B.名称)||Decode(C.规格,NULL,NULL,'('||C.规格||')') as 内容," & _
        " Decode(A.单次用量,NULL,NULL,A.单次用量||'""'||B.计算单位||'""') as 单量," & _
        " Decode(A.总给予量,NULL,NULL,Decode(Instr('56',Nvl(B.类别,'*')),0,A.总给予量||'""'||B.计算单位||'""',A.总给予量/D.住院包装||'""'||D.住院单位||'""')) as 总量," & _
        " A.执行频次,A.医生嘱托" & _
        " From " & strSql & " A,诊疗项目目录 B,收费项目目录 C,药品规格 D" & _
        " Where Nvl(A.诊疗项目ID,0)=B.ID(+) And Nvl(A.收费细目ID,0)=C.ID(+) And Nvl(A.收费细目ID,0)=D.药品ID(+)" & _
        " And A.ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetAdviceDefineText", str医嘱IDs)
    
    strSql = ""
    Do While Not rsTmp.EOF
        If (InStr(",5,6,C,*,", NVL(rsTmp!类别, "*")) > 0 Or IsNull(rsTmp!相关id)) _
            And Not (NVL(rsTmp!类别) = "E" And rsTmp!ID = lngPre相关ID) Then
            strSql = strSql & vbCrLf & "　○" & rsTmp!期效 & "，" & rsTmp!内容 & _
                IIf(Not IsNull(rsTmp!单量), "，每次" & rsTmp!单量, "") & _
                IIf(Not IsNull(rsTmp!总量), "，共" & rsTmp!总量, "") & _
                IIf(Not IsNull(rsTmp!执行频次), "，" & rsTmp!执行频次, "") & _
                IIf(Not IsNull(rsTmp!医生嘱托), "，" & rsTmp!医生嘱托, "")
        End If
        
        lngPre相关ID = NVL(rsTmp!相关id, 0)
        rsTmp.MoveNext
    Loop
    
    GetAdviceDefineText = Mid(strSql, 3)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEPRDefineText(Optional ByVal str病历IDs As String, Optional ByVal lng项目ID As Long) As String
'功能：获取路径项目对应的病历定义内容描述串
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    If lng项目ID <> 0 Then '新版电子病历和老版同时
        strSql = "Select Nvl(a.名称, b.名称) as 名称 From 临床路径病历 A, 病历文件列表 B Where a.项目id = [2] And a.文件id = b.Id(+)" & vbNewLine & _
                "order by a.序号"
    ElseIf str病历IDs <> "" And lng项目ID = 0 Then '老版
        strSql = "Select /*+ Rule*/ 名称 From 病历文件列表" & _
            " Where ID IN(Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " Order by 编号"
    Else '新版
        strSql = "select 名称 from 临床路径病历 t where t.项目id=[2] and t.文件id is null and t.原型id IN (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))) order by 序号"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetEPRDefineText", str病历IDs, lng项目ID)
    
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "、" & rsTmp!名称
        rsTmp.MoveNext
    Loop
    
    GetEPRDefineText = Mid(strSql, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check医嘱项目(ByVal lng执行ID As Long) As Boolean
'功能：检查指定的执行项目是否属于医嘱类
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From 病人路径医嘱 Where 路径执行ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Check医嘱项目", lng执行ID)
    
    Check医嘱项目 = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Check病历项目(ByVal lng执行ID As Long) As Boolean
'功能：检查指定的执行项目是否属于病历类
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 1 From 电子病历记录 Where 路径执行ID = [1] And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Check病历项目", lng执行ID)
    
    Check病历项目 = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckSameDayOfPhase(ByVal lngPhase As Long, ByVal lngDay As Long) As Boolean
'功能：检查当天是否还有适用的其他后续阶段(当前阶段及分支除外)
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    '如果当前是分支阶段，则取其父ID
    strSql = "Select 父ID From 临床路径阶段 Where ID = [1] And 父ID is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取阶段", lngPhase)
    If rsTmp.RecordCount > 0 Then lngPhase = rsTmp!父ID
    
    strSql = "Select 1" & vbNewLine & _
            "From 临床路径阶段 A, 临床路径阶段 B" & vbNewLine & _
            "Where a.Id = [1] And a.路径id = b.路径id And a.版本号 = b.版本号  And nvl(a.分支id,0)=nvl(b.分支ID,0) And b.序号 > a.序号" & vbNewLine & _
            "      And [2] Between b.开始天数 And Nvl(b.结束天数, b.开始天数) And b.父ID is Null And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取阶段", lngPhase, lngDay)
    If rsTmp.RecordCount > 0 Then CheckSameDayOfPhase = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInPath(t_pati As TYPE_Pati, ByVal lng病人路径Id As Long, Optional ByRef lng确诊天数 As Long) As Date
'功能：获取病人的进入路径的开始时间
'参数：返回lng确诊天数=当前可选阶段的的天数
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select a.开始时间,b.确诊天数  From 病人临床路径 a,临床路径目录 b Where a.Id =[1] And a.路径id = b.id"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取入径时间", lng病人路径Id)
    If IsNull(rsTmp!开始时间) Then
        GetPatiInPath = zlDatabase.Currentdate
        'GetPatiInPath = GetPatiInDate(t_pati)
    Else
        GetPatiInPath = rsTmp!开始时间
    End If
    
    lng确诊天数 = Val("" & rsTmp!确诊天数)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInDate(t_pati As TYPE_Pati, Optional lng入院天数 As Long) As Date
'功能：获取病人的入院或入科\转科时间
'返回：lng入院天数：入院或入科天数
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Max(开始时间) 开始时间,To_number(Trunc(Sysdate)-Trunc(Max(开始时间)))+1 as 入院天数" & vbNewLine & _
            "From (Select 入院日期 As 开始时间" & vbNewLine & _
            "       From 病案主页" & vbNewLine & _
            "       Where 病人id = [1] And 主页id = [2]" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 开始时间 From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 In (2, 3) And 科室id = [3])"
           
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取入科时间", t_pati.病人ID, t_pati.主页ID, t_pati.科室ID)
    GetPatiInDate = CDate(rsTmp!开始时间)
    lng入院天数 = rsTmp!入院天数
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiInfo(lng病人ID As Long, lng主页ID As Long) As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select NVL(B.姓名,a.姓名) 姓名,NVL(B.性别,a.性别) 性别 ,NVL(B.年龄,a.年龄) 年龄 , To_Char(a.出生日期, 'yyyy-mm-dd hh24:mi:ss') 出生日期, b.当前病况," & vbNewLine & _
            "       d.名称 As 病例分型,a.门诊号 ,b.住院号, b.出院日期, b.入院日期,e.名称 As 病区" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病案主页从表 C, 临床病例分型 D, 部门表 E" & vbNewLine & _
            "Where a.病人id = b.病人id And b.病人id = [1] And b.主页id = [2] And e.Id = b.当前病区id And" & vbNewLine & _
            "      b.病人id = c.病人id(+) And b.主页id = c.主页id(+) And c.信息名(+) = '病例分型' And c.信息值 = d.编码(+)"
            
    On Error GoTo errH
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSql, "读取病人数据", lng病人ID, lng主页ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetAdvice(strIDs As String) As ADODB.Recordset
'功能：获取路径项目对应的医嘱记录集
    Dim strSql As String
 
    strSql = "Select /*+ rule*/ a.路径项目ID,a.医嘱内容ID,b.期效,Nvl(b.相关ID,b.ID) 相关ID,b.诊疗项目ID" & vbNewLine & _
            "From 临床路径医嘱 A,路径医嘱内容 B,(Select Column_Value As ID From Table(f_Num2list([1]))) C" & vbNewLine & _
            "Where a.医嘱内容id=b.id And a.路径项目id = c.Id" & vbNewLine & _
            "Order by b.序号"
    On Error GoTo errH
    Set GetAdvice = zlDatabase.OpenSQLRecord(strSql, "读取医嘱记录", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFile(strIDs As String, Optional ByVal int场合 As Integer = 2) As ADODB.Recordset
'功能：获取路径项目对应的病历文件记录集
'int场合=1门诊，2-住院
    Dim strSql As String

    strSql = "select A.项目ID as 路径项目ID,A.文件ID,A.原型ID,B.保留  from " & IIf(int场合 = 1, "门诊路径病历", "临床路径病历") & " a,病历文件列表 b  where a.项目id in (Select Column_Value From Table(f_Num2list([1]))) and a.文件ID=b.id(+)"
  
    On Error GoTo errH
    Set GetFile = zlDatabase.OpenSQLRecord(strSql, "读取病历文件", strIDs)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDelPathItem(ByVal lng执行ID As Long, ByVal int场合 As Integer) As Boolean
'功能：检查指定的医嘱类路径项目执行记录是否可以删除或重新生成
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strIDs As String
    Dim i As Long

    '不是当天生成的长嘱，重新生成后自动停止，不管是否发送；
    '是当天生成的长嘱，已校对但未作废，不允许取消(已停止的也不允许)，未校对的，取消时自动删除对应的医嘱。
    strSql = "Select 1" & vbNewLine & _
             "From 病人路径医嘱 A, 病人路径医嘱 B" & vbNewLine & _
             "Where a.路径执行id = [1] And a.病人医嘱id = b.病人医嘱id And b.路径执行id <> a.路径执行id  And rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
    If rsTmp.RecordCount = 0 Then '当天生成
        strSql = "Select 1 From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And c.医嘱状态 <> 4 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "该项目存在已校对但未作废的医嘱，请先作废医嘱后再执行此操作。", vbInformation, gstrSysName
            Exit Function
        End If

        If int场合 = 1 Then
            '对于已经过审核的医嘱，不允许修改删除。
            strSql = "Select 1 From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.开嘱医生 Like '%/%' And rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查医嘱", lng执行ID)
            If rsTmp.RecordCount > 0 Then
                MsgBox "该项目对应的医嘱已经过医生审核，不能执行此操作。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else '非当天生成
        '前期校对后未停用的长嘱，以路径外项目的形式在路径表中展示，如果要删除它,需确定该类长嘱已停止或作废才能删除
        strSql = "Select c.医嘱内容" & vbNewLine & _
                 "From 病人路径医嘱 B, 病人医嘱记录 C,病人路径执行 D " & vbNewLine & _
                 "Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.医嘱状态 > 1 And d.id=b.路径执行Id And d.项目ID is null " & _
                 " And c.医嘱状态 Not In (4,8,9) And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng执行ID)
        If rsTmp.RecordCount > 0 Then
            strIDs = ""
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & vbNewLine & rsTmp!医嘱内容
                rsTmp.MoveNext
            Next
            MsgBox "该路径外项目存在已校对但未作废或停止的长期医嘱：" & strIDs & vbNewLine & "请先作废或停止该医嘱后再执行取消。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckDelPathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPathIcon(ByVal lng图标ID As Long) As StdPicture
'功能：获取指定ID的临床路径图标
'说明：第一次读取时利用集合进行缓存
    Dim rsTmp As ADODB.Recordset
    Dim objIcon As StdPicture
    Dim strFile As String, strSql As String
    Dim blnExist As Boolean
    
    blnExist = True
    If gcolIcons Is Nothing Then
        Set gcolIcons = New Collection
        blnExist = False
    End If
    If blnExist Then
        On Error Resume Next
        Set GetPathIcon = gcolIcons("_" & lng图标ID)
        If Err.Number <> 0 Then
            Err.Clear: blnExist = False
        End If
    End If
    
    On Error GoTo errH
    
    If Not blnExist Then
        Screen.MousePointer = 11
        
        strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\zlTemplate.bmp"
                
        strSql = "Select ID From 临床路径图标 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPathIcon", lng图标ID)
        If Not rsTmp.EOF Then
            If sys.ReadLob(glngSys, 11, lng图标ID, strFile) <> "" Then
                gcolIcons.Add LoadPicture(strFile), "_" & lng图标ID
                gobjFile.DeleteFile strFile
            End If
        End If
        
        Screen.MousePointer = 0
    End If
    
    Set GetPathIcon = gcolIcons("_" & lng图标ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetNextCode(ByVal str分类 As String, Optional ByVal intMode As Integer = 0) As String
'功能：获取指定分类的临床路径的缺省新编码
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intMax As Integer
    Dim strTable As String
    
    On Error GoTo errH
    
    If intMode = 1 Then
        strTable = "门诊路径目录"
    Else
        strTable = "临床路径目录"
    End If
    '取最大长度：按Max，001比0001大
    If str分类 = "" Then
        strSql = "Select Max(Length(编码)) As 长度 From " & strTable & " Where 分类 Is Null"
    Else
        strSql = "Select Max(Length(编码)) As 长度 From " & strTable & " Where 分类=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextCode", str分类)
    If rsTmp.EOF Then
        GetNextCode = "01": Exit Function
    ElseIf IsNull(rsTmp!长度) Then
        GetNextCode = "01": Exit Function
    Else
        intMax = rsTmp!长度
    End If
    
    '按最大长度编码
    If str分类 = "" Then
        strSql = "Select Max(编码) As 编码 From " & strTable & " Where 分类 Is Null And Length(编码)=[2]"
    Else
        strSql = "Select Max(编码) As 编码 From " & strTable & " Where 分类=[1] And Length(编码)=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetNextCode", str分类, intMax)
    GetNextCode = zlCommFun.IncStr(rsTmp!编码)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetNextPhase(ByVal lng阶段ID As Long, ByVal lng当前阶段分支ID As Long) As Long
'功能：获取指定阶段的后续阶段ID
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 父ID From 临床路径阶段 Where id = [1] And 父ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "下一阶段", lng阶段ID)
    If rsTmp.RecordCount > 0 Then lng阶段ID = Val(rsTmp!父ID)
    
    strSql = "Select b.ID From 临床路径阶段 a,临床路径阶段 b " & _
            "Where a.路径ID= b.路径ID And a.版本号= b.版本号 And b.序号>a.序号 And NVL(b.分支ID,0)=[2] And a.ID = [1] And b.父ID Is Null And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "下一阶段", lng阶段ID, lng当前阶段分支ID)
    
    If rsTmp.RecordCount > 0 Then GetNextPhase = Val(rsTmp!ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetMustDay(ByVal lng病人路径Id As Long, ByVal lng当前天数 As Long, Optional ByVal blnIsNotMinus As Boolean, _
        Optional ByVal lng合并路径记录ID As Long) As Long
'功能：获取病人路径执行理论上的当前天数(=当前实际天数-曾经延迟的天数+提前天数(有可能一次提前多天))
'参数：blnIsNotMinus=是否不减去延迟时间（评估时求当前天数）
'      lng阶段ID和lng天数=如果是合并路径，则从导入起点之后算起
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lng延迟天数 As Long
    Dim lng提前天数 As Long
    Dim i As Long
    Dim lng阶段实际天数 As Long
    Dim lng阶段开始天数 As Long
    Dim byt提前进度 As Byte
    
    On Error GoTo errH
    If lng合并路径记录ID <> 0 Then
        strSql = "Select Max(Decode(g.时间进度, 1, 1, 2, 2, 0)) As 阶段是否提前, c.开始天数, Nvl(c.结束天数, c.开始天数) as 结束天数, Sum(Decode(g.时间进度, -1, 1, 0)) As 阶段延后天数," & vbNewLine & _
                "       Count(*) As 阶段实际天数" & vbNewLine & _
                "From 病人合并路径评估 A, 临床路径分支 B, 临床路径阶段 C, 临床路径阶段 D, 临床路径阶段 E, 临床路径阶段 F,病人路径评估 G" & vbNewLine & _
                "Where a.合并路径阶段id = c.Id And c.父id = d.Id(+) And b.Id(+) = c.分支id And b.前一阶段id = e.Id(+) And e.父id = f.Id(+) And a.路径记录id=g.路径记录id And a.阶段id=g.阶段id and a.日期=g.日期" & vbNewLine & _
                "      And a.合并路径记录id = [1]" & vbNewLine & _
                "Group By c.开始天数, Nvl(c.结束天数, c.开始天数), a.阶段id, c.分支id, d.序号, c.序号, f.序号, e.序号" & vbNewLine & _
                "Order By Decode(c.分支id, Null, Nvl(d.序号, c.序号), Nvl(d.序号, c.序号) + Nvl(f.序号, e.序号))"
   
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDay", lng合并路径记录ID)
    Else
       strSql = "Select Max(Decode(a.时间进度, 1, 1, 2, 2, 0)) As 阶段是否提前, c.开始天数, Nvl(c.结束天数, c.开始天数) as 结束天数, Sum(Decode(a.时间进度, -1, 1, 0)) As 阶段延后天数," & vbNewLine & _
                "       Count(1) As 阶段实际天数" & vbNewLine & _
                "From 病人路径评估 A, 临床路径分支 B, 临床路径阶段 C, 临床路径阶段 D, 临床路径阶段 E, 临床路径阶段 F" & vbNewLine & _
                "Where a.阶段id = c.Id And c.父id = d.Id(+) And b.Id(+) = c.分支id And b.前一阶段id = e.Id(+) And e.父id = f.Id(+) And" & vbNewLine & _
                "      a.路径记录id = [1] " & vbNewLine & _
                "Group By c.开始天数, Nvl(c.结束天数, c.开始天数), a.阶段id, c.分支id, d.序号, c.序号, f.序号, e.序号" & vbNewLine & _
                "Order By Decode(c.分支id, Null, Nvl(d.序号, c.序号), Nvl(d.序号, c.序号) + Nvl(f.序号, e.序号))"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetMustDay", lng病人路径Id)
    End If
    For i = 0 To rsTmp.RecordCount - 1
        '延迟天数
        lng延迟天数 = lng延迟天数 + Val(rsTmp!阶段延后天数 & "")
        '提前天数
        If Val(rsTmp!阶段是否提前 & "") = 1 Or Val(rsTmp!阶段是否提前 & "") = 2 Then
            '合并路径的起始阶段可能是提前（首要路径提前评估后再导入合并路径）,第一个阶段就需要计算提前天数(由于已经控制了合并路径必须从第一个阶段开始，所以这个判断暂时无效，如有需求要合并路径从后面开始，则可开启)
'            If i = 0 And rsTmp!开始天数 & "" <> "1" Then
'                lng提前天数 = Val(rsTmp!开始天数 & "") - 1
'                rsTmp.MoveNext
'            Else
                '最后一个阶段是提前的则加1天，因为还不知道后面会选那一个阶段
                If i = rsTmp.RecordCount - 1 Or rsTmp!开始天数 & "" = rsTmp!结束天数 & "" Then
                    If Val(rsTmp!阶段是否提前 & "") = 1 Then
                        lng提前天数 = lng提前天数 + 1
                    ElseIf Val(rsTmp!阶段是否提前 & "") = 2 Then
                        '下一阶段提前至明天,此时不需要像“下一阶段提前到今天”再额外加一天
                    End If
                    rsTmp.MoveNext
                Else
                    '先记录下阶段实际天数和开始天数
                    lng阶段开始天数 = Val(rsTmp!开始天数 & "")
                    lng阶段实际天数 = Val(rsTmp!阶段实际天数 & "")
                    byt提前进度 = Val(rsTmp!阶段是否提前 & "")
                    rsTmp.MoveNext
                    lng提前天数 = lng提前天数 + (Val(rsTmp!开始天数 & "") - lng阶段开始天数 - lng阶段实际天数 + IIf(byt提前进度 = 2, 0, 1))
                End If
'            End If
        Else
            rsTmp.MoveNext
        End If
    Next
    
    GetMustDay = lng当前天数 - IIf(blnIsNotMinus, 0, lng延迟天数) + lng提前天数
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPhaseNO(ByVal lng阶段ID As Long) As Long
'功能：获取指定阶段的序号(如果该阶段是分支，则取父阶段的序号)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select 父ID From 临床路径阶段 Where id = [1] And 父ID is Not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng阶段ID)
    If rsTmp.RecordCount > 0 Then lng阶段ID = Val(rsTmp!父ID)
    
    strSql = "Select 序号 From 临床路径阶段 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng阶段ID)
    If rsTmp.RecordCount > 0 Then GetPhaseNO = Val(rsTmp!序号)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLastPhaseNO(ByVal lng病人路径Id As Long, ByVal lng路径ID As Long)
'功能：获取病人指定路径最近一个阶段的序号
Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Max(Nvl(c.序号, b.序号)) 序号" & vbNewLine & _
            "From 病人路径执行 A, 临床路径阶段 B, 临床路径阶段 C" & vbNewLine & _
            "Where a.路径记录id = [1] And a.阶段id = b.Id And b.路径id = [2] And b.父id = c.Id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "阶段序号", lng病人路径Id, lng路径ID)
    
    GetLastPhaseNO = Val("" & rsTmp!序号)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ExportPathToXML(ByVal lng路径ID As Long, ByVal int版本号 As Integer, ByVal strFile As String) As Boolean
'功能：导出临床路径到XML文件
'参数：strFile=包含路径的文件名
'说明：导出包含路径信息和指定版本的信息
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    Dim xPI As IXMLDOMProcessingInstruction
    
    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsItemAdvice As ADODB.Recordset
    Dim rsItemEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    Set xPath = New DOMDocument
    
    '注释
    xPath.appendChild xPath.createComment(gstrSysName & "  操作员:" & UserInfo.姓名 & ",部门:" & UserInfo.部门名 & ",时间:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
    
    '根结点
    Set xRoot = xPath.createElement("ClinicalPathways")
    Set xPath.documentElement = xRoot
    Call xRoot.setAttribute("ID", lng路径ID)
    Call xRoot.setAttribute("Version", int版本号)

    '临床路径信息
    strSql = "Select A.分类,A.编码,A.名称,A.通用,A.最新版本,A.病例分型," & _
        " A.适用病情,A.适用性别,A.适用年龄,A.说明,B.标准住院日,B.标准费用," & _
        " B.版本说明,B.创建人,B.创建时间,B.审核人,B.审核时间,B.停用人,B.停用时间,A.确诊天数,A.结束路径控制,A.性质" & _
        " From 临床路径目录 A,临床路径版本 B Where A.ID=B.路径ID And A.ID=[1] And B.版本号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    Set xNode = CreateNode(1, xRoot, "PathInfo", NODE_ELEMENT, "")
        CreateNode 2, xNode, "分类", , rsTmp!分类
        CreateNode 2, xNode, "编码", , rsTmp!编码
        CreateNode 2, xNode, "名称", , rsTmp!名称
        CreateNode 2, xNode, "通用", , NVL(rsTmp!通用)
        CreateNode 2, xNode, "最新版本", , NVL(rsTmp!最新版本)
        CreateNode 2, xNode, "病例分型", , NVL(rsTmp!病例分型)
        CreateNode 2, xNode, "适用病情", , NVL(rsTmp!适用病情)
        CreateNode 2, xNode, "适用性别", , NVL(rsTmp!适用性别)
        CreateNode 2, xNode, "适用年龄", , NVL(rsTmp!适用年龄)
        CreateNode 2, xNode, "说明", , NVL(rsTmp!说明)
        CreateNode 2, xNode, "标准住院日", , NVL(rsTmp!标准住院日)
        CreateNode 2, xNode, "标准费用", , NVL(rsTmp!标准费用)
        CreateNode 2, xNode, "版本说明", , NVL(rsTmp!版本说明)
        CreateNode 2, xNode, "创建人", , NVL(rsTmp!创建人)
        CreateNode 2, xNode, "创建时间", , Format(NVL(rsTmp!创建时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "审核人", , NVL(rsTmp!审核人)
        CreateNode 2, xNode, "审核时间", , Format(NVL(rsTmp!审核时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "停用人", , NVL(rsTmp!停用人)
        CreateNode 2, xNode, "停用时间", , Format(NVL(rsTmp!停用时间), "yyyy-MM-dd HH:mm:ss")
        CreateNode 2, xNode, "确诊天数", , NVL(rsTmp!确诊天数)
        CreateNode 2, xNode, "结束路径控制", , NVL(rsTmp!结束路径控制)
        CreateNode 2, xNode, "性质", , NVL(rsTmp!性质, 0)
    
    '临床路径科室
    strSql = "Select B.ID,B.编码,B.名称 From 临床路径科室 A,部门表 B Where A.路径ID=[1] And A.科室ID=B.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDepts", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDept", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "科室ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "编码", , rsTmp!编码
                CreateNode 3, xSubNode1, "名称", , rsTmp!名称
            rsTmp.MoveNext
        Loop
    End If
    
    '临床路径病种
    strSql = "Select A.疾病ID,B.编码 as 疾病码,B.名称 as 疾病名," & _
        " A.诊断ID,C.编码 as 诊断码,C.名称 as 诊断名, a.性质 as 性质" & _
        " From 临床路径病种 A,疾病编码目录 B,疾病诊断目录 C" & _
        " Where Nvl(A.疾病ID,0)=B.ID(+) And Nvl(A.诊断ID,0)=C.ID(+) And A.路径ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathDiseases", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathDisease", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "疾病ID", , NVL(rsTmp!疾病id)
                CreateNode 3, xSubNode1, "疾病码", , NVL(rsTmp!疾病码)
                CreateNode 3, xSubNode1, "疾病名", , NVL(rsTmp!疾病名)
                CreateNode 3, xSubNode1, "诊断ID", , NVL(rsTmp!诊断id)
                CreateNode 3, xSubNode1, "诊断码", , NVL(rsTmp!诊断码)
                CreateNode 3, xSubNode1, "诊断名", , NVL(rsTmp!诊断名)
                CreateNode 3, xSubNode1, "性质", , NVL(rsTmp!性质)
            rsTmp.MoveNext
        Loop
    End If
    
    '临床路径分支
    strSql = "Select ID,路径ID,版本号,名称,说明,前一阶段ID,标准住院日,标准费用,创建人,创建时间 From 临床路径分支 where 路径ID=[1] AND 版本号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathBranchs", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathBranch", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "路径ID", , rsTmp!路径ID
                CreateNode 3, xSubNode1, "版本号", , rsTmp!版本号
                CreateNode 3, xSubNode1, "名称", , NVL(rsTmp!名称)
                CreateNode 3, xSubNode1, "说明", , NVL(rsTmp!说明)
                CreateNode 3, xSubNode1, "前一阶段ID", , NVL(rsTmp!前一阶段ID)
                CreateNode 3, xSubNode1, "标准住院日", , NVL(rsTmp!标准住院日)
                CreateNode 3, xSubNode1, "标准费用", , NVL(rsTmp!标准费用)
                CreateNode 3, xSubNode1, "创建人", , NVL(rsTmp!创建人)
                CreateNode 3, xSubNode1, "创建时间", , Format(NVL(rsTmp!创建时间), "yyyy-MM-dd HH:mm:ss")
            rsTmp.MoveNext
        Loop
    End If
    
    '导入评估
    strSql = "Select B.评估类型,B.阶段ID,A.ID,A.评估指标,A.指标类型,A.指标结果,b.分支ID" & _
        " From 路径评估指标 A,临床路径评估 B" & _
        " Where A.评估ID=B.ID And B.路径ID=[1] And 版本号=[2]" & _
        " Order by B.评估类型,B.阶段ID,A.序号"
    Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    strSql = "Select B.评估类型,B.阶段ID,A.指标ID,A.项目ID,A.关系式,A.条件值,A.条件组合,b.分支ID" & _
        " From 路径评估条件 A,临床路径评估 B" & _
        " Where A.评估ID=B.ID And B.路径ID=[1] And 版本号=[2]" & _
        " Order by B.评估类型,B.阶段ID"
    Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    rsEvalMark.Filter = "评估类型=1"
    rsEvalCond.Filter = "评估类型=1"
    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
        Set xNode = CreateNode(1, xRoot, "ImportEval", NODE_ELEMENT, "")
            If Not rsEvalMark.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Marks", NODE_ELEMENT, "")
                Do While Not rsEvalMark.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Mark", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "ID", , rsEvalMark!ID
                        CreateNode 4, xSubNode2, "评估指标", , rsEvalMark!评估指标
                        CreateNode 4, xSubNode2, "指标类型", , rsEvalMark!指标类型
                        CreateNode 4, xSubNode2, "指标结果", , rsEvalMark!指标结果
                    rsEvalMark.MoveNext
                Loop
            End If
            If Not rsEvalCond.EOF Then
                Set xSubNode1 = CreateNode(2, xNode, "Conditions", NODE_ELEMENT, "")
                Do While Not rsEvalCond.EOF
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode 4, xSubNode2, "指标ID", , rsEvalCond!指标ID
                        CreateNode 4, xSubNode2, "关系式", , rsEvalCond!关系式
                        CreateNode 4, xSubNode2, "条件值", , rsEvalCond!条件值
                        CreateNode 4, xSubNode2, "条件组合", , rsEvalCond!条件组合
                    rsEvalCond.MoveNext
                Loop
            End If
    End If
    
    '路径医嘱内容
    strSql = "Select Distinct A.ID,A.相关ID,A.序号,A.期效,A.诊疗项目ID,D.编码 as 诊疗编码,D.名称 as 诊疗名称," & _
        " A.收费细目ID,E.编码 as 收费编码,E.名称 as 收费名称,A.医嘱内容,A.单次用量,A.总给予量," & _
        " A.标本部位,A.检查方法,A.医生嘱托,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位," & _
        " A.执行性质,A.执行科室ID,F.编码 as 执行科室码,F.名称 as 执行科室名,A.时间方案,A.是否缺省,A.是否备选,A.配方ID,A.组合项目ID" & _
        " From 路径医嘱内容 A,临床路径医嘱 B,临床路径项目 C,诊疗项目目录 D,收费项目目录 E,部门表 F" & _
        " Where A.ID=B.医嘱内容ID And B.路径项目ID=C.ID And C.路径ID=[1] And C.版本号=[2]" & _
        " And Nvl(A.诊疗项目ID,0)=D.ID(+) And Nvl(A.收费细目ID,0)=E.ID(+) And Nvl(A.执行科室ID,0)=F.ID(+)" & _
        " Order by A.序号,A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    If Not rsTmp.EOF Then
        Set xNode = CreateNode(1, xRoot, "PathAdvices", NODE_ELEMENT, "")
        Do While Not rsTmp.EOF
            Set xSubNode1 = CreateNode(2, xNode, "PathAdvice", NODE_ELEMENT, "")
                CreateNode 3, xSubNode1, "ID", , rsTmp!ID
                CreateNode 3, xSubNode1, "相关ID", , NVL(rsTmp!相关id)
                CreateNode 3, xSubNode1, "序号", , rsTmp!序号
                CreateNode 3, xSubNode1, "期效", , rsTmp!期效
                CreateNode 3, xSubNode1, "诊疗项目ID", , NVL(rsTmp!诊疗项目ID)
                CreateNode 3, xSubNode1, "诊疗编码", , NVL(rsTmp!诊疗编码)
                CreateNode 3, xSubNode1, "诊疗名称", , NVL(rsTmp!诊疗名称)
                CreateNode 3, xSubNode1, "收费细目ID", , NVL(rsTmp!收费细目ID)
                CreateNode 3, xSubNode1, "收费编码", , NVL(rsTmp!收费编码)
                CreateNode 3, xSubNode1, "收费名称", , NVL(rsTmp!收费名称)
                CreateNode 3, xSubNode1, "医嘱内容", , NVL(rsTmp!医嘱内容)
                CreateNode 3, xSubNode1, "单次用量", , NVL(rsTmp!单次用量)
                CreateNode 3, xSubNode1, "总给予量", , NVL(rsTmp!总给予量)
                CreateNode 3, xSubNode1, "标本部位", , NVL(rsTmp!标本部位)
                CreateNode 3, xSubNode1, "检查方法", , NVL(rsTmp!检查方法)
                CreateNode 3, xSubNode1, "医生嘱托", , NVL(rsTmp!医生嘱托)
                CreateNode 3, xSubNode1, "执行频次", , NVL(rsTmp!执行频次)
                CreateNode 3, xSubNode1, "频率次数", , NVL(rsTmp!频率次数)
                CreateNode 3, xSubNode1, "频率间隔", , NVL(rsTmp!频率间隔)
                CreateNode 3, xSubNode1, "间隔单位", , NVL(rsTmp!间隔单位)
                CreateNode 3, xSubNode1, "执行性质", , NVL(rsTmp!执行性质)
                CreateNode 3, xSubNode1, "执行科室ID", , NVL(rsTmp!执行科室ID)
                CreateNode 3, xSubNode1, "执行科室码", , NVL(rsTmp!执行科室码)
                CreateNode 3, xSubNode1, "执行科室名", , NVL(rsTmp!执行科室名)
                CreateNode 3, xSubNode1, "时间方案", , NVL(rsTmp!时间方案)
                CreateNode 3, xSubNode1, "是否缺省", , NVL(rsTmp!是否缺省, 0)
                CreateNode 3, xSubNode1, "是否备选", , NVL(rsTmp!是否备选, 0)
                CreateNode 3, xSubNode1, "配方ID", , NVL(rsTmp!配方ID)
                CreateNode 3, xSubNode1, "组合项目ID", , NVL(rsTmp!组合项目ID)
            rsTmp.MoveNext
        Loop
    End If
    
    '临床路径分类
    strSql = "Select 名称,分支ID From 临床路径分类 Where 路径ID=[1] And 版本号=[2] Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    Set xNode = CreateNode(1, xRoot, "PathCategorys", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        Set xSubNode1 = CreateNode(2, xNode, "PathCategory", NODE_ELEMENT, NVL(rsTmp!名称))
        CreateNode 2, xSubNode1, "名称", NODE_ELEMENT, NVL(rsTmp!名称)
        CreateNode 2, xSubNode1, "分支ID", NODE_ELEMENT, NVL(rsTmp!分支ID)
        rsTmp.MoveNext
    Loop
    
    '临床路径阶段/项目
    strSql = "Select ID,Nvl(父ID,0) as 父ID,序号,名称,开始天数,结束天数,标志,分类,说明,分支ID" & _
        " From 临床路径阶段 Where 路径ID=[1] And 版本号=[2] Order by Nvl(父ID,0) Desc,序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    strSql = "Select ID,阶段ID,分类,项目序号,项目内容,执行方式,执行者,生成者,项目结果,图标ID,内容要求,分支ID" & _
        " From 临床路径项目 Where 路径ID=[1] And 版本号=[2] Order by 阶段ID,分类,项目序号"
    Set rsItem = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    strSql = "Select A.路径项目ID,A.医嘱内容ID From 临床路径医嘱 A,临床路径项目 B" & _
        " Where A.路径项目ID=B.ID And B.路径ID=[1] And 版本号=[2]"
    Set rsItemAdvice = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    strSql = "Select A.项目ID,A.文件ID,C.编号,C.名称 From 临床路径病历 A,临床路径项目 B,病历文件列表 C" & _
        " Where A.项目ID=B.ID And A.文件ID=C.ID And B.路径ID=[1] And 版本号=[2]"
    Set rsItemEPR = zlDatabase.OpenSQLRecord(strSql, "ExportPathToXML", lng路径ID, int版本号)
    
    Set rsClone = rsTmp.Clone: rsTmp.Filter = "父ID=0"
    
    Set xNode = CreateNode(1, xRoot, "PathTimeSteps", NODE_ELEMENT, "")
    Do While Not rsTmp.EOF
        '缺省分支
        Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
            CreateNode 3, xSubNode1, "ID", , rsTmp!ID
            CreateNode 3, xSubNode1, "父ID", , ""
            CreateNode 3, xSubNode1, "序号", , rsTmp!序号
            CreateNode 3, xSubNode1, "名称", , rsTmp!名称
            CreateNode 3, xSubNode1, "开始天数", , NVL(rsTmp!开始天数)
            CreateNode 3, xSubNode1, "结束天数", , NVL(rsTmp!结束天数)
            CreateNode 3, xSubNode1, "标志", , NVL(rsTmp!标志)
            CreateNode 3, xSubNode1, "说明", , NVL(rsTmp!说明)
            CreateNode 3, xSubNode1, "分类", , NVL(rsTmp!分类)
            CreateNode 3, xSubNode1, "分支ID", , NVL(rsTmp!分支ID)
            
            '阶段的项目
            rsItem.Filter = "阶段ID=" & rsTmp!ID
            Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
            Do While Not rsItem.EOF
                Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                    CreateNode 5, xSubNode3, "ID", , rsItem!ID
                    CreateNode 5, xSubNode3, "分类", , rsItem!分类
                    CreateNode 5, xSubNode3, "项目序号", , rsItem!项目序号
                    CreateNode 5, xSubNode3, "项目内容", , rsItem!项目内容
                    CreateNode 5, xSubNode3, "执行方式", , NVL(rsItem!执行方式)
                    CreateNode 5, xSubNode3, "执行者", , NVL(rsItem!执行者)
                    CreateNode 5, xSubNode3, "生成者", , NVL(rsItem!生成者, 1)
                    CreateNode 5, xSubNode3, "项目结果", , NVL(rsItem!项目结果)
                    CreateNode 5, xSubNode3, "图标ID", , NVL(rsItem!图标ID)
                    CreateNode 5, xSubNode3, "内容要求", , NVL(rsItem!内容要求, 0)
                    CreateNode 5, xSubNode3, "分支ID", , NVL(rsItem!分支ID)
                    
                    '项目对应的医嘱
                    rsItemAdvice.Filter = "路径项目ID=" & rsItem!ID
                    If Not rsItemAdvice.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                        Do While Not rsItemAdvice.EOF
                            CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!医嘱内容ID
                            rsItemAdvice.MoveNext
                        Loop
                    End If
                    '项目对应的病历
                    rsItemEPR.Filter = "项目ID=" & rsItem!ID
                    If Not rsItemEPR.EOF Then
                        Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                        Do While Not rsItemEPR.EOF
                            Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                CreateNode 7, xSubNode5, "文件ID", , rsItemEPR!文件ID
                                CreateNode 7, xSubNode5, "文件编号", , rsItemEPR!编号
                                CreateNode 7, xSubNode5, "文件名称", , rsItemEPR!名称
                            rsItemEPR.MoveNext
                        Loop
                    End If
                    
                rsItem.MoveNext
            Loop
        
            '阶段的评估
            rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & rsTmp!ID
            rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & rsTmp!ID
            If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                    If Not rsEvalMark.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                        Do While Not rsEvalMark.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                CreateNode 6, xSubNode4, "评估指标", , rsEvalMark!评估指标
                                CreateNode 6, xSubNode4, "指标类型", , rsEvalMark!指标类型
                                CreateNode 6, xSubNode4, "指标结果", , rsEvalMark!指标结果
                                CreateNode 6, xSubNode4, "分支ID", , NVL(rsEvalMark!分支ID)
                            rsEvalMark.MoveNext
                        Loop
                    End If
                    If Not rsEvalCond.EOF Then
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                        Do While Not rsEvalCond.EOF
                            Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                CreateNode 6, xSubNode4, "指标ID", , NVL(rsEvalCond!指标ID)
                                CreateNode 6, xSubNode4, "项目ID", , NVL(rsEvalCond!项目ID)
                                CreateNode 6, xSubNode4, "关系式", , rsEvalCond!关系式
                                CreateNode 6, xSubNode4, "条件值", , rsEvalCond!条件值
                                CreateNode 6, xSubNode4, "条件组合", , rsEvalCond!条件组合
                                CreateNode 6, xSubNode4, "分支ID", , NVL(rsEvalCond!分支ID)
                            rsEvalCond.MoveNext
                        Loop
                    End If
            End If
        
        '备选分支
        rsClone.Filter = "父ID=" & rsTmp!ID
        If Not rsClone.EOF Then
            Do While Not rsClone.EOF
                Set xSubNode1 = CreateNode(2, xNode, "PathTimeStep", NODE_ELEMENT, "")
                    CreateNode 3, xSubNode1, "ID", , rsClone!ID
                    CreateNode 3, xSubNode1, "父ID", , rsClone!父ID
                    CreateNode 3, xSubNode1, "序号", , rsClone!序号
                    CreateNode 3, xSubNode1, "名称", , rsClone!名称
                    CreateNode 3, xSubNode1, "开始天数", , NVL(rsClone!开始天数)
                    CreateNode 3, xSubNode1, "结束天数", , NVL(rsClone!结束天数)
                    CreateNode 3, xSubNode1, "标志", , NVL(rsClone!标志)
                    CreateNode 3, xSubNode1, "说明", , NVL(rsClone!说明)
                    CreateNode 3, xSubNode1, "分支ID", , NVL(rsClone!分支ID)
                
                    '阶段的项目
                    rsItem.Filter = "阶段ID=" & rsClone!ID
                    Set xSubNode2 = CreateNode(3, xSubNode1, "Items", NODE_ELEMENT, "")
                    Do While Not rsItem.EOF
                        Set xSubNode3 = CreateNode(4, xSubNode2, "Item", NODE_ELEMENT, "")
                            CreateNode 5, xSubNode3, "ID", , rsItem!ID
                            CreateNode 5, xSubNode3, "分类", , rsItem!分类
                            CreateNode 5, xSubNode3, "项目序号", , rsItem!项目序号
                            CreateNode 5, xSubNode3, "项目内容", , rsItem!项目内容
                            CreateNode 5, xSubNode3, "执行方式", , NVL(rsItem!执行方式)
                            CreateNode 5, xSubNode3, "执行者", , NVL(rsItem!执行者)
                            CreateNode 5, xSubNode3, "生成者", , NVL(rsItem!生成者, 1)
                            CreateNode 5, xSubNode3, "项目结果", , NVL(rsItem!项目结果)
                            CreateNode 5, xSubNode3, "图标ID", , NVL(rsItem!图标ID)
                            CreateNode 5, xSubNode3, "分支ID", , NVL(rsItem!分支ID)
                            
                            '项目对应的医嘱
                            rsItemAdvice.Filter = "路径项目ID=" & rsItem!ID
                            If Not rsItemAdvice.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "Advices", NODE_ELEMENT, "")
                                Do While Not rsItemAdvice.EOF
                                    CreateNode 6, xSubNode4, "Advice", , rsItemAdvice!医嘱内容ID
                                    rsItemAdvice.MoveNext
                                Loop
                            End If
                            '项目对应的病历
                            rsItemEPR.Filter = "项目ID=" & rsItem!ID
                            If Not rsItemEPR.EOF Then
                                Set xSubNode4 = CreateNode(5, xSubNode3, "EPRFiles", NODE_ELEMENT, "")
                                Do While Not rsItemEPR.EOF
                                    Set xSubNode5 = CreateNode(6, xSubNode4, "EPRFile", NODE_ELEMENT, "")
                                        CreateNode 7, xSubNode5, "文件ID", , rsItemEPR!文件ID
                                        CreateNode 7, xSubNode5, "文件编号", , rsItemEPR!编号
                                        CreateNode 7, xSubNode5, "文件名称", , rsItemEPR!名称
                                    rsItemEPR.MoveNext
                                Loop
                            End If
                            
                        rsItem.MoveNext
                    Loop
                    
                    '阶段的评估
                    rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & rsClone!ID
                    rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & rsClone!ID
                    If Not rsEvalMark.EOF Or Not rsEvalCond.EOF Then
                        Set xSubNode2 = CreateNode(3, xSubNode1, "StepEval", NODE_ELEMENT, "")
                            If Not rsEvalMark.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Marks", NODE_ELEMENT, "")
                                Do While Not rsEvalMark.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Mark", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "ID", , rsEvalMark!ID
                                        CreateNode 6, xSubNode4, "评估指标", , rsEvalMark!评估指标
                                        CreateNode 6, xSubNode4, "指标类型", , rsEvalMark!指标类型
                                        CreateNode 6, xSubNode4, "指标结果", , rsEvalMark!指标结果
                                        CreateNode 6, xSubNode4, "分支ID", , NVL(rsEvalMark!分支ID)
                                    rsEvalMark.MoveNext
                                Loop
                            End If
                            If Not rsEvalCond.EOF Then
                                Set xSubNode3 = CreateNode(4, xSubNode2, "Conditions", NODE_ELEMENT, "")
                                Do While Not rsEvalCond.EOF
                                    Set xSubNode4 = CreateNode(5, xSubNode3, "Condition", NODE_ELEMENT, "")
                                        CreateNode 6, xSubNode4, "指标ID", , NVL(rsEvalCond!指标ID)
                                        CreateNode 6, xSubNode4, "项目ID", , NVL(rsEvalCond!项目ID)
                                        CreateNode 6, xSubNode4, "关系式", , rsEvalCond!关系式
                                        CreateNode 6, xSubNode4, "条件值", , rsEvalCond!条件值
                                        CreateNode 6, xSubNode4, "条件组合", , rsEvalCond!条件组合
                                        CreateNode 6, xSubNode4, "分支ID", , NVL(rsEvalCond!分支ID)
                                    rsEvalCond.MoveNext
                                Loop
                            End If
                    End If
                
                rsClone.MoveNext
            Loop
        End If
        
        rsTmp.MoveNext
    Loop
    
    'XML信息
    Set xPI = xPath.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call xPath.insertBefore(xPI, xPath.childNodes(0))
    
    '保存成文件
    xPath.Save strFile
    Set xPath = Nothing
    
    ExportPathToXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, ByVal SubNodeName As String, Optional ByVal DefaultValue As String) As String
    Dim NodeTMP As IXMLDOMNode
    
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
End Function

Public Function ImportPathFromXML(ByVal strFile As String, _
    Optional ByVal lng路径ID As Long, Optional ByVal int版本号 As Integer, _
    Optional ByVal intLimit As Integer, Optional ByRef blnLimit As Boolean) As Boolean
'功能：导入指定的临床路径XML文件
'参数：lng路径ID,int版本号=如果指定，则只导入版本相关部分信息；如果没有指定，则根据根据XML中的信息进行路径新增或者完全覆盖
'      intLimit=总体限制的最大路径数量,为0表示不限制
'      blnLimit=是否被允许的最大路径数量所限制导入失败
    Dim rsTmp As ADODB.Recordset
    Dim rsIcon As ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim arrSQL As Variant, strSql As String
    Dim colItemID As Collection
    Dim colStepID As Collection
    Dim colMarkID As Collection
    Dim colAdviceID As Collection
    Dim colAdviceOriginalID As Collection
    Dim colBranchID As Collection
    Dim colPreID As Collection
    
    Dim xPath As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim xSubNode1 As IXMLDOMNode
    Dim xSubNode2 As IXMLDOMNode
    Dim xSubNode3 As IXMLDOMNode
    Dim xSubNode4 As IXMLDOMNode
    Dim xSubNode5 As IXMLDOMNode
    
    Dim str编码 As String, lng阶段ID As Long
    Dim strValue As String, strTemp1 As String
    Dim strTemp2 As String, strTemp3 As String
    Dim blnDo As Boolean, blnTran As Boolean
    Dim strValueTurn As String
    Dim strValueTurn1 As String
    Dim i As Long, k As Long, n As Long, m As Long
    Dim strTmp As String
    Dim strPreStep As String
    Dim strtemp4 As String
    Dim lngType As Long
    Dim strImportRef As String
    Dim lng导入结果 As Long '记录同一路径项目医嘱的导入状态0，全部未导入，1，全部导入，2，部分导入
    Dim lngCount As Long, str组IDs As String, arrID As Variant, lng组ID As Long, strFilter As String
    Dim lng项目ID As Long
    
    On Error GoTo errH
    
    rsAdvice.Fields.Append "ID", adBigInt
    rsAdvice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "导入参考", adVarChar, 200, adFldIsNullable
    rsAdvice.Fields.Append "项目ID", adBigInt, , adFldIsNullable
    rsAdvice.Fields.Append "导入状态", adInteger
    
    rsAdvice.CursorLocation = adUseClient
    rsAdvice.LockType = adLockOptimistic
    rsAdvice.CursorType = adOpenStatic
    rsAdvice.Open
    
    blnLimit = False
    
    Set xPath = New DOMDocument
    xPath.Load strFile
    
    '如果不包含任何元素，则退出
    If xPath.documentElement Is Nothing Then
        Set xPath = Nothing
        Screen.MousePointer = 0
        Exit Function
    End If
    
    arrSQL = Array()
    
    '读取XML内容
    Set xRoot = xPath.selectSingleNode("ClinicalPathways")
    Set xNode = xRoot.selectSingleNode("PathInfo")
    If lng路径ID = 0 Then
        '获取应用科室的情况
        strTemp1 = ""
        If Val(GetNodeValue(xNode, "通用")) = 2 Then
            Set xSubNode1 = xRoot.selectSingleNode("PathDepts")
            If Not xSubNode1 Is Nothing Then
                strSql = "Select A.ID,A.编码,A.名称" & _
                    " From 部门表 A,部门性质说明 C" & _
                    " Where A.ID=C.部门ID And C.服务对象 IN(2,3) And C.工作性质='临床'" & _
                    " Order by A.编码"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML")
                
                For Each xSubNode2 In xSubNode1.childNodes
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode2, "编码") & "' And 名称='" & GetNodeValue(xSubNode2, "名称") & "'"
                    If Not rsTmp.EOF Then strTemp1 = strTemp1 & "," & rsTmp!ID
                Next
            
                strTemp1 = Mid(strTemp1, 2)
            End If
        End If
        
        '获取应用疾病的情况
        strValue = ""
        Set xSubNode1 = xRoot.selectSingleNode("PathDiseases")
        If Not xSubNode1 Is Nothing Then
            strTemp2 = "": strTemp3 = ""
            For Each xSubNode2 In xSubNode1.childNodes
                If Val(GetNodeValue(xSubNode2, "性质")) = 0 Then
                    If Val(GetNodeValue(xSubNode2, "疾病ID")) <> 0 Then
                        strSql = "Select ID From 疾病编码目录 Where 编码=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "疾病码"), GetNodeValue(xSubNode2, "疾病名"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & "," & rsTmp!ID
                    ElseIf Val(GetNodeValue(xSubNode2, "诊断ID")) <> 0 Then
                        strSql = "Select ID From 疾病诊断目录 Where 编码=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "诊断码"), GetNodeValue(xSubNode2, "诊断名"))
                        If Not rsTmp.EOF Then strTemp3 = strTemp3 & "," & rsTmp!ID
                    End If
                Else
                    If Val(GetNodeValue(xSubNode2, "疾病ID")) <> 0 Then
                        strSql = "Select ID From 疾病编码目录 Where 编码=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "疾病码"), GetNodeValue(xSubNode2, "疾病名"))
                        If Not rsTmp.EOF Then strValueTurn = strValueTurn & "," & rsTmp!ID
                    ElseIf Val(GetNodeValue(xSubNode2, "诊断ID")) <> 0 Then
                        strSql = "Select ID From 疾病诊断目录 Where 编码=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode2, "诊断码"), GetNodeValue(xSubNode2, "诊断名"))
                        If Not rsTmp.EOF Then strValueTurn1 = strValueTurn1 & "," & rsTmp!ID
                    End If
                End If
            Next
            If strTemp2 <> "" Or strTemp3 <> "" Then
                strValue = Mid(strTemp2, 2) & ";" & Mid(strTemp3, 2)
            End If
            If strValueTurn <> "" Or strValueTurn1 <> "" Then
                strValueTurn = Mid(strValueTurn, 2) & ";" & Mid(strValueTurn1, 2)
            End If
        End If
        
        '产生临床路径信息
        strSql = "Select ID,编码,最新版本 From 临床路径目录 Where 分类=[1] And 名称=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xNode, "分类"), GetNodeValue(xNode, "名称"))
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If Not rsTmp.EOF Then
            '新增版本或者覆盖版本
            lng路径ID = rsTmp!ID
            int版本号 = NVL(rsTmp!最新版本, 0) + 1 '可能覆盖未审核版本
            str编码 = rsTmp!编码
            arrSQL(UBound(arrSQL)) = "zl_临床路径目录_Update(" & _
                lng路径ID & ",'" & GetNodeValue(xNode, "分类") & "','" & str编码 & "'," & _
                "'" & GetNodeValue(xNode, "名称") & "','" & GetNodeValue(xNode, "说明") & "'," & _
                "'" & GetNodeValue(xNode, "病例分型") & "','" & GetNodeValue(xNode, "适用病情") & "'," & _
                Val(GetNodeValue(xNode, "适用性别")) & ",'" & GetNodeValue(xNode, "适用年龄") & "'," & _
                Val(GetNodeValue(xNode, "通用")) & ",'" & strTemp1 & "','" & strValue & "'," & Val(GetNodeValue(xNode, "确诊天数")) & ",'" _
                & strValueTurn & "'," & Val(GetNodeValue(xNode, "结束路径控制")) & "," & Val(GetNodeValue(xNode, "性质")) & ")"
        
        Else
            '检查授权限制
            If intLimit > 0 Then
                strSql = "Select Nvl(Count(*),0) as 数量 From 临床路径目录"
                Set rsTmp = New ADODB.Recordset
                Call zlDatabase.OpenRecordset(rsTmp, strSql, "ImportPathFromXML")
                If rsTmp!数量 >= intLimit Then
                    blnLimit = True
                    Set xPath = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            
            '新增路径
            lng路径ID = zlDatabase.GetNextId("临床路径目录")
            int版本号 = 1
            str编码 = GetNextCode(GetNodeValue(xNode, "分类"))
            arrSQL(UBound(arrSQL)) = "zl_临床路径目录_Insert(" & _
                "'" & GetNodeValue(xNode, "分类") & "','" & str编码 & "'," & _
                "'" & GetNodeValue(xNode, "名称") & "','" & GetNodeValue(xNode, "说明") & "'," & _
                "'" & GetNodeValue(xNode, "病例分型") & "','" & GetNodeValue(xNode, "适用病情") & "'," & _
                Val(GetNodeValue(xNode, "适用性别")) & ",'" & GetNodeValue(xNode, "适用年龄") & "'," & _
                Val(GetNodeValue(xNode, "通用")) & ",'" & strTemp1 & "','" & strValue & "'," & lng路径ID & "," & Val(GetNodeValue(xNode, "确诊天数")) & ",'" & _
                strValueTurn & "'," & Val(GetNodeValue(xNode, "结束路径控制")) & "," & Val(GetNodeValue(xNode, "性质")) & ")"
        End If
    Else
        strSql = "Select 性质 From 临床路径目录 Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", lng路径ID)
        If rsTmp.RecordCount > 0 Then lngType = Val(rsTmp!性质 & "")
    End If
    
    '删除版本相关的内容，重新产生
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Delete(" & lng路径ID & "," & int版本号 & ")"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Update(" & lng路径ID & "," & int版本号 & "," & _
        "'" & GetNodeValue(xNode, "标准住院日") & "','" & GetNodeValue(xNode, "标准费用") & "'," & _
        "'" & GetNodeValue(xNode, "版本说明") & "')"
    
    '导入评估
    Set xNode = xRoot.selectSingleNode("ImportEval")
    If Not xNode Is Nothing Then
        Set xSubNode1 = xNode.selectSingleNode("Marks")
        If Not xSubNode1 Is Nothing Then
            k = 1
            Set colItemID = New Collection
            For Each xSubNode2 In xSubNode1.childNodes
                strValue = zlDatabase.GetNextId("路径评估指标")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode2, "ID")
                            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & lng路径ID & "," & int版本号 & ",NULL,1," & _
                    strValue & "," & k & ",'" & GetNodeValue(xSubNode2, "评估指标") & "'," & _
                    Val(GetNodeValue(xSubNode2, "指标类型")) & ",'" & GetNodeValue(xSubNode2, "指标结果") & "')"
                
                k = k + 1
            Next
        End If
        Set xSubNode1 = xNode.selectSingleNode("Conditions")
        If Not xSubNode1 Is Nothing Then
            For Each xSubNode2 In xSubNode1.childNodes
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & lng路径ID & "," & int版本号 & ",NULL,1," & _
                    colItemID("_" & GetNodeValue(xSubNode2, "指标ID")) & ",NULL,'" & GetNodeValue(xSubNode2, "关系式") & "'," & _
                    "'" & GetNodeValue(xSubNode2, "条件值") & "','" & GetNodeValue(xSubNode2, "条件组合") & "')"
            Next
        End If
    End If
    
    '临床路径分支
    Set xNode = xRoot.selectSingleNode("PathBranchs")
    If Not xNode Is Nothing Then
        Set colBranchID = New Collection
        Set colPreID = New Collection
        For Each xSubNode1 In xNode.childNodes
                strValue = zlDatabase.GetNextId("临床路径分支")
                colBranchID.Add strValue, "_" & GetNodeValue(xSubNode1, "ID")
                If GetNodeValue(xSubNode1, "前一阶段ID") <> "" Then
                    strPreStep = strPreStep & "," & GetNodeValue(xSubNode1, "前一阶段ID")
                    On Error Resume Next
                    If colPreID("_" & GetNodeValue(xSubNode1, "前一阶段ID")) = "" Then
                        Err.Clear
                        colPreID.Add strValue, "_" & GetNodeValue(xSubNode1, "前一阶段ID")
                    Else
                        strTmp = colPreID("_" & GetNodeValue(xSubNode1, "前一阶段ID"))
                        colPreID.Remove "_" & GetNodeValue(xSubNode1, "前一阶段ID")
                        colPreID.Add strTmp & "," & strValue, "_" & GetNodeValue(xSubNode1, "前一阶段ID")
                    End If
                    On Error GoTo 0
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                '前一阶段ID插入Null是因为前面删除了当前版本的阶段，还没有得到新的前一阶段ID，插入旧的话会找不到父项
                arrSQL(UBound(arrSQL)) = "Zl_临床路径分支_Update(" & strValue & "," & lng路径ID & "," & int版本号 & ",'" & GetNodeValue(xSubNode1, "名称") & "',Null,'" & _
                    GetNodeValue(xSubNode1, "标准住院日") & "','" & GetNodeValue(xSubNode1, "标准费用") & "','" & GetNodeValue(xSubNode1, "说明") & "')"
        Next
        strPreStep = Mid(strPreStep, 2)

    End If
    
    '路径医嘱内容
    Set xNode = xRoot.selectSingleNode("PathAdvices")
    If Not xNode Is Nothing Then
        Set colAdviceID = New Collection
        Set colAdviceOriginalID = New Collection
        For Each xSubNode1 In xNode.childNodes
            strValue = zlDatabase.GetNextId("路径医嘱内容")
            strTemp1 = GetNodeValue(xSubNode1, "ID")
            colAdviceID.Add strValue, "_" & strTemp1
            colAdviceOriginalID.Add strTemp1, "_" & strValue
        Next
        k = 1
        For Each xSubNode1 In xNode.childNodes
            blnDo = True: strTemp1 = "": strTemp2 = "": strTemp3 = ""
                
            '验证诊疗项目ID
            If Val(GetNodeValue(xSubNode1, "诊疗项目ID")) <> 0 Then
                strSql = "Select 编码,ID From 诊疗项目目录 Where 名称=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "诊疗名称"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "诊疗编码") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp1 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp1 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '验证收费细目ID
            If blnDo And Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0 Then
                strSql = "Select 编码,ID From 收费项目目录 Where 名称=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "收费名称"))
                If Not rsTmp.EOF Then
                    rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "收费编码") & "'"
                    If rsTmp.RecordCount > 0 Then
                        strTemp2 = rsTmp!ID
                    Else
                        rsTmp.Filter = ""
                        strTemp2 = rsTmp!ID
                    End If
                Else
                    blnDo = False
                End If
            End If
            '获取导入参考
            strImportRef = IIf(Val(GetNodeValue(xSubNode1, "诊疗项目ID")) <> 0, Trim(GetNodeValue(xSubNode1, "诊疗名称")) & _
                IIf(Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0, "(" & Trim(GetNodeValue(xSubNode1, "收费名称")) & ")", ""), "" & _
                IIf(Val(GetNodeValue(xSubNode1, "收费细目ID")) <> 0, Trim(GetNodeValue(xSubNode1, "收费名称")), ""))
            '保存路径医嘱的导入状况进入临时记录集
            rsAdvice.AddNew
            rsAdvice!ID = Val(GetNodeValue(xSubNode1, "ID"))
            rsAdvice!相关id = Val(GetNodeValue(xSubNode1, "相关ID"))
            rsAdvice!导入参考 = strImportRef
            rsAdvice!导入状态 = IIf(blnDo, 1, 0)
            rsAdvice.Update
            
            If blnDo Then
                '验证执行科室ID
                If Val(GetNodeValue(xSubNode1, "执行科室ID")) <> 0 Then
                    strSql = "Select 编码,ID From 部门表 Where 名称=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode1, "执行科室名"))
                    If Not rsTmp.EOF Then
                        rsTmp.Filter = "编码='" & GetNodeValue(xSubNode1, "执行科室码") & "'"
                        If rsTmp.RecordCount > 0 Then
                            strTemp3 = rsTmp!ID
                        Else
                            rsTmp.Filter = ""
                            strTemp3 = rsTmp!ID
                        End If
                    End If
                End If
                
                strValue = GetNodeValue(xSubNode1, "相关ID")
                If strValue <> "" Then strValue = colAdviceID("_" & strValue)
                                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_路径医嘱内容_Insert(" & _
                    colAdviceID("_" & GetNodeValue(xSubNode1, "ID")) & "," & ZVal(strValue) & "," & _
                    k & "," & Val(GetNodeValue(xSubNode1, "期效")) & "," & ZVal(strTemp1) & "," & _
                    "'" & GetNodeValue(xSubNode1, "医嘱内容") & "'," & ZVal(GetNodeValue(xSubNode1, "单次用量")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "总给予量")) & "," & ZVal(strTemp2) & "," & _
                    "'" & GetNodeValue(xSubNode1, "标本部位") & "','" & GetNodeValue(xSubNode1, "检查方法") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "执行频次") & "'," & ZVal(GetNodeValue(xSubNode1, "频率次数")) & "," & _
                    ZVal(GetNodeValue(xSubNode1, "频率间隔")) & ",'" & GetNodeValue(xSubNode1, "间隔单位") & "'," & _
                    "'" & GetNodeValue(xSubNode1, "医生嘱托") & "'," & Val(GetNodeValue(xSubNode1, "执行性质")) & "," & _
                    ZVal(strTemp3) & ",'" & GetNodeValue(xSubNode1, "时间方案") & "',Null,Null," & GetNodeValue(xSubNode1, "是否缺省", 0) & "," & _
                    GetNodeValue(xSubNode1, "是否备选", 0) & ",Null," & ZVal(GetNodeValue(xSubNode1, "配方ID", 0)) & "," & ZVal(GetNodeValue(xSubNode1, "组合项目ID", 0)) & ")"
                k = k + 1
            Else
                '如果有相关ID为该医嘱的，则这些医嘱不应产生
                strValue = GetNodeValue(xSubNode1, "ID")
                For n = 0 To UBound(arrSQL)
                    If arrSQL(n) <> "" Then
                        If Split(arrSQL(n), ",")(1) = colAdviceID("_" & strValue) Then
                            '标明该医嘱不存在
                            strTemp1 = Split(Split(arrSQL(n), ",")(0), "(")(1)
                            colAdviceID.Remove "_" & colAdviceOriginalID("_" & strTemp1)
                            colAdviceID.Add "0", "_" & colAdviceOriginalID("_" & strTemp1)
                            arrSQL(n) = ""
                        End If
                    End If
                Next
                '标明该医嘱不存在
                colAdviceID.Remove "_" & strValue
                colAdviceID.Add "0", "_" & strValue
            End If
        Next
    End If
    
    '临床路径分类
    Set xNode = xRoot.selectSingleNode("PathCategorys")
    k = 1
    For Each xSubNode1 In xNode.childNodes
        strTmp = GetNodeValue(xSubNode1, "分支ID")
        If strTmp = "" Then
            strTmp = "Null"
        Else
            strTmp = colBranchID("_" & strTmp)
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_临床路径分类_Insert(" & lng路径ID & "," & int版本号 & "," & k & ",'" & IIf(GetNodeValue(xSubNode1, "名称") = "", xSubNode1.Text, GetNodeValue(xSubNode1, "名称")) & "',Null," & strTmp & ")"
        k = k + 1
    Next
    
    '临床路径阶段
    Set xNode = xRoot.selectSingleNode("PathTimeSteps")
    k = 1
    Set colStepID = New Collection
    For Each xSubNode1 In xNode.childNodes
        lng阶段ID = zlDatabase.GetNextId("临床路径阶段")
        colStepID.Add lng阶段ID, "_" & GetNodeValue(xSubNode1, "ID")
        
        strTemp1 = GetNodeValue(xSubNode1, "父ID")
        If strTemp1 <> "" Then strTemp1 = colStepID("_" & strTemp1)
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        strTmp = GetNodeValue(xSubNode1, "分支ID")
        If strTmp = "" Then
            strTmp = "Null"
        Else
            strTmp = colBranchID("_" & strTmp)
        End If
        If strPreStep <> "" Then
            If InStr("," & strPreStep & ",", "," & GetNodeValue(xSubNode1, "ID") & ",") > 0 Then
                strtemp4 = colPreID("_" & GetNodeValue(xSubNode1, "ID"))
            End If
        End If
        arrSQL(UBound(arrSQL)) = "Zl_临床路径阶段_Insert(" & _
            lng阶段ID & "," & lng路径ID & "," & int版本号 & "," & ZVal(strTemp1) & "," & _
            IIf(strTemp1 = "", k, GetNodeValue(xSubNode1, "序号")) & ",'" & GetNodeValue(xSubNode1, "名称") & "'," & _
            ZVal(GetNodeValue(xSubNode1, "开始天数")) & "," & ZVal(GetNodeValue(xSubNode1, "结束天数")) & "," & _
            "'" & GetNodeValue(xSubNode1, "标志") & "','" & GetNodeValue(xSubNode1, "说明") & "'," & _
            "'" & GetNodeValue(xSubNode1, "分类") & "'," & strTmp & _
            ",'" & strtemp4 & "')"
        If strTemp1 = "" Then k = k + 1
        strtemp4 = ""
        
        '阶段中的路径项目
        Set xSubNode2 = xSubNode1.selectSingleNode("Items")
        If Not xSubNode2 Is Nothing Then
            Set colItemID = New Collection
            For Each xSubNode3 In xSubNode2.childNodes
                strTemp1 = "": strTemp2 = ""
                '项目关联医嘱
                lng项目ID = Val(GetNodeValue(xSubNode3, "ID"))
                Set xSubNode4 = xSubNode3.selectSingleNode("Advices")
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '在临时结构记录集中设置医嘱与项目的关联
                        rsAdvice.Filter = "ID=" & Val(xSubNode5.Text)
                        If rsAdvice.RecordCount <> 0 Then
                            Call rsAdvice.Update("项目ID", lng项目ID)
                        End If
                        rsAdvice.Filter = ""
                        
                        If Val(colAdviceID("_" & xSubNode5.Text)) <> 0 Then
                            strTemp1 = strTemp1 & "," & colAdviceID("_" & xSubNode5.Text)
                        End If
                    Next
                    strTemp1 = Mid(strTemp1, 2)
                End If
                
                '项目关联病历
                Set xSubNode4 = xSubNode3.selectSingleNode("EPRFiles")
                i = 1
                If Not xSubNode4 Is Nothing Then
                    For Each xSubNode5 In xSubNode4.childNodes
                        '验证病历文件ID
                        strSql = "Select ID From 病历文件列表 Where 编号=[1] And 名称=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML", GetNodeValue(xSubNode5, "文件编号"), GetNodeValue(xSubNode5, "文件名称"))
                        If Not rsTmp.EOF Then strTemp2 = strTemp2 & ";" & rsTmp!ID & ",," & GetNodeValue(xSubNode5, "文件名称") & "," & i + 1
                    Next
                    strTemp2 = Mid(strTemp2, 2)
                End If
                
                '图标的验证：只支持固有图标
                strTemp3 = GetNodeValue(xSubNode3, "图标ID")
                If strTemp3 <> "" Then
                    If rsIcon Is Nothing Then
                        strSql = "Select ID,Nvl(性质,0) as 性质 From 临床路径图标"
                        Set rsIcon = zlDatabase.OpenSQLRecord(strSql, "ImportPathFromXML")
                    End If
                    rsIcon.Filter = "ID=" & strTemp3 & " And 性质=1"
                    If rsIcon.EOF Then strTemp3 = ""
                End If
                
                strValue = zlDatabase.GetNextId("临床路径项目")
                colItemID.Add strValue, "_" & GetNodeValue(xSubNode3, "ID")
                
                strTmp = GetNodeValue(xSubNode3, "分支ID")
                If strTmp = "" Then
                    strTmp = "Null"
                Else
                    strTmp = colBranchID("_" & strTmp)
                End If
                
                rsAdvice.Filter = "项目ID=" & lng项目ID
                
                lngCount = rsAdvice.RecordCount
                strImportRef = ""
                lng导入结果 = 1
                str组IDs = ""
                
                rsAdvice.Filter = rsAdvice.Filter & " And 导入状态=0"
                '获取导入状态
                If rsAdvice.RecordCount <> 0 Then
                    lng导入结果 = IIf(rsAdvice.RecordCount = lngCount, 0, 2)
                    '获取未导入成功医嘱的组ID
                    For n = 1 To rsAdvice.RecordCount
                        lng组ID = rsAdvice!相关id
                        If lng组ID = 0 Then lng组ID = rsAdvice!ID
                        If InStr(str组IDs & ",", "," & lng组ID & ",") = 0 Then
                            str组IDs = str组IDs & "," & lng组ID
                        End If
                        rsAdvice.MoveNext
                    Next
                End If
                If Len(str组IDs) > 0 Then str组IDs = Mid(str组IDs, 2)

                arrID = Split(str组IDs, ",")
                '获取导入参考
                For m = LBound(arrID) To UBound(arrID)
                    '过滤未导入的同一组医嘱
                    strFilter = "(项目ID = " & lng项目ID & " AND 相关ID = " & Val(arrID(m)) & ") OR (项目ID = " & lng项目ID & " AND ID=" & Val(arrID(m)) & ")"
                    rsAdvice.Filter = strFilter
                    rsAdvice.Sort = "相关ID,ID"
                    If rsAdvice.RecordCount <> 0 Then
                        For n = 1 To rsAdvice.RecordCount
                            If n = 1 And strImportRef = "" Then
                                strImportRef = rsAdvice!导入参考
                            ElseIf n = 1 And strImportRef <> "" Then
                                strImportRef = strImportRef & Chr(10) & Chr(13) & rsAdvice!导入参考 '已经有其他组医嘱已经保存在strImportRef
                            Else
                                strImportRef = strImportRef & ";" & rsAdvice!导入参考
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                Next
   
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径项目_Insert(" & _
                    strValue & "," & lng路径ID & "," & int版本号 & "," & lng阶段ID & "," & _
                    "'" & GetNodeValue(xSubNode3, "分类") & "'," & GetNodeValue(xSubNode3, "项目序号") & "," & _
                    "'" & GetNodeValue(xSubNode3, "项目内容") & "'," & Val(GetNodeValue(xSubNode3, "执行方式")) & "," & _
                    ZVal(GetNodeValue(xSubNode3, "执行者")) & ",'" & GetNodeValue(xSubNode3, "项目结果") & "'," & _
                    ZVal(strTemp3) & ",'" & strTemp1 & "','" & strTemp2 & "'," & GetNodeValue(xSubNode3, "内容要求", 0) & _
                    "," & strTmp & ",'" & Trim(strImportRef) & "'," & IIf(Trim(strImportRef) = "" And lng导入结果 = 1, "Null", lng导入结果) & _
                    "," & ZVal(GetNodeValue(xSubNode3, "生成者")) & ")"
            Next
        End If
        
        '阶段评估-如果是合并路径，则不导入阶段评估
        If lngType = 0 Then
        Set xSubNode2 = xSubNode1.selectSingleNode("StepEval")
            If Not xSubNode2 Is Nothing Then
                '评估指标
                Set xSubNode3 = xSubNode2.selectSingleNode("Marks")
                If Not xSubNode3 Is Nothing Then
                    i = 1
                    Set colMarkID = New Collection
                    For Each xSubNode4 In xSubNode3.childNodes
                        strValue = zlDatabase.GetNextId("路径评估指标")
                        colMarkID.Add strValue, "_" & GetNodeValue(xSubNode4, "ID")
                        
                        strTmp = GetNodeValue(xSubNode4, "分支ID")
                        If strTmp = "" Then
                            strTmp = "Null"
                        Else
                            strTmp = colBranchID("_" & strTmp)
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & _
                            lng路径ID & "," & int版本号 & "," & lng阶段ID & ",2," & _
                            strValue & "," & i & ",'" & GetNodeValue(xSubNode4, "评估指标") & "'," & _
                            Val(GetNodeValue(xSubNode4, "指标类型")) & ",'" & GetNodeValue(xSubNode4, "指标结果") & _
                            "," & strTmp & "')"
                        i = i + 1
                    Next
                End If
                '指标条件
                Set xSubNode3 = xSubNode2.selectSingleNode("Conditions")
                If Not xSubNode3 Is Nothing Then
                    For Each xSubNode4 In xSubNode3.childNodes
                        strTemp1 = GetNodeValue(xSubNode4, "指标ID")
                        If strTemp1 <> "" Then strTemp1 = colMarkID("_" & strTemp1)
                        strTemp2 = GetNodeValue(xSubNode4, "项目ID")
                        If strTemp2 <> "" Then strTemp2 = colItemID("_" & strTemp2)
                        
                        strTmp = GetNodeValue(xSubNode4, "分支ID")
                        If strTmp = "" Then
                            strTmp = "Null"
                        Else
                            strTmp = colBranchID("_" & strTmp)
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & _
                            lng路径ID & "," & int版本号 & "," & lng阶段ID & ",2," & _
                            ZVal(strTemp1) & "," & ZVal(strTemp2) & ",'" & GetNodeValue(xSubNode4, "关系式") & "'," & _
                            "'" & GetNodeValue(xSubNode4, "条件值") & "'," & Val(GetNodeValue(xSubNode4, "条件组合")) & _
                            "," & strTmp & ")"
                    Next
                End If
            End If
        End If
    Next
    
    '执行提交数据
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), "ImportPathFromXML"
        End If
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    Set xPath = Nothing
    ImportPathFromXML = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set xPath = Nothing
End Function

'-------------------------------------------------------------
'病历相关处理程序
'-------------------------------------------------------------
Public Function ReadRTFData(ByVal lng病历ID As Long, edtEditor As Editor) As Boolean
'功能：读取病历文件的RTF数据到editor控件中
    Dim strZipFile As String, strTempFile As String
        
    On Error GoTo errH
    strZipFile = ReadLobForPath(glngSys, 5, lng病历ID)
    strTempFile = zlFileUnzip(strZipFile)
    edtEditor.OpenDoc strTempFile
    
     '删除临时文件
    Kill strTempFile
    Kill strZipFile
   
    ReadRTFData = True
    Exit Function
errH:
    ReadRTFData = False
End Function

Public Function ReadLobForPath(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'功能：将指定的LOB字段复制为临时文件
'参数：
'lngSys:系统编号
'Action:操作类型（用以区别是操作哪个表）
'---系统100,Zl_Lob_Append
'0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
'5-电子病历格式;6-电子病历图形;7-病历页面格式(图形)；8-电子病历附件;9-体温重叠标记
'10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
'17-药品规格(使用说明);18-药品规格(图片);23-供应商图片
'---系统2400,Zl24_Lob_Append
'手麻常用图形,无Action
'---系统2100,Zl21_Lob_Append
'1-体质类型调养;2-体检体辨结论(该图片只有读取，没有保存);3-体检申报记录;4-体检任务人员,5-体检任务结果
'---系统2600,Zl26_Lob_Append
'14-导诊控件目录,15-导诊资源目录
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0正常记录,1读取转储后备表记录
'返回：存放内容的文件名，失败则返回零长度""
    Const conChunkSize As Integer = 10240
    
    Dim rsLob As ADODB.Recordset
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim strSql As String
    Dim objFile As New FileSystemObject
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSql = "Select Zl_Lob_ReadForPath([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 24
            strSql = "Select Zl24_Lob_Read([2],[3]) as 片段 From Dual"
        Case 21
            strSql = "Select Zl21_Lob_Read([1],[2],[3]) as 片段 From Dual"
        Case 26
            strSql = "Select Zl26_Lob_Read([1],[2],[3]) as 片段 From Dual"
    End Select
    If strSql = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        Do
            Set rsLob = zlDatabase.OpenSQLRecord(strSql, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).Value) Then Exit Do
            strText = rsLob.Fields(0).Value
            
            ReDim aryChunk(Len(strText) / 2 - 1) As Byte
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLob = zlDatabase.OpenSQLRecord(strSql, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLob.EOF Then Exit Do
            If IsNull(rsLob.Fields(0).Value) Then Exit Do
            strText = rsLob.Fields(0).Value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLobForPath = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        If lngCount = 0 Then
            Kill strFile: ReadLobForPath = ""
        Else
            ReadLobForPath = strFile
        End If
    End If
    Err.Clear
End Function

Public Function SaveRTFData(ByVal lng病历ID As Long, ByVal lng病人ID As Long, lng主页ID As Long, lngBaby As Long, edtEditor As Editor, Optional ByVal intType As Integer) As Boolean
'功能：保存病人病历格式RTF数据
'参数：
    Dim strZipFile As String, strTempFile As String, i As Long
        
    '要素内容更新
    Call ElementsUpdate(lng病历ID, lng病人ID, lng主页ID, lngBaby, edtEditor, intType)
    
    On Error GoTo errH
    strTempFile = App.Path & "\TMP.rtf"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    edtEditor.SaveDoc strTempFile
    '压缩文件
    strZipFile = zlFileZip(strTempFile)
    '保存格式
    sys.SaveLob glngSys, 5, lng病历ID, strZipFile
    
    '删除临时文件
    Kill strTempFile
    Kill strZipFile

    SaveRTFData = True
    Exit Function
errH:
    SaveRTFData = False
End Function

Private Function ElementsUpdate(ByVal lng病历ID As Long, ByVal lng病人ID As Long, lng主页ID As Long, lngBaby As Long, edtEditor As Editor, Optional ByVal intType As Integer) As Boolean
'功能：更新Editor控件中的替换要素内容，以便保存为RTF文件
'    intType=1 门诊
    Dim ThisElements As New zlRichEPR.cEPRElements
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, lngKey As Long
    Dim bFinded As Boolean, bNeeded As Boolean, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long

    strSql = "Select 对象标记,ID From 电子病历内容 Where 文件ID= [1] And 对象类型 = 4 And 终止版=0 and 保留对象 =0 And 替换域 =1 order by 对象标记 "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人病历", lng病历ID)
    For i = 1 To rsTmp.RecordCount
        lngKey = ThisElements.Add(NVL(rsTmp("对象标记"), 0))
        ThisElements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rsTmp("ID"), True
        rsTmp.MoveNext
    Next

     For i = 1 To ThisElements.count
        If ThisElements(i).替换域 = 1 Then
            ThisElements(i).内容文本 = GetReplaceEleValue(ThisElements(i).要素名称, lng病人ID, lng主页ID, IIf(intType = 1, cprPF_门诊, cprPF_住院), 0, lngBaby)
            bFinded = FindNextKey(edtEditor, 0, "E", ThisElements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
            ThisElements(i).Refresh edtEditor
        End If
        If ThisElements(i).替换域 = 1 And ThisElements(i).自动转文本 Then
            EleToString edtEditor, ThisElements(i)     '自动转化为纯文本（暂时不删除该要素）
        End If
    Next
    Set ThisElements = Nothing
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub EleToString(ByRef edtThis As Object, Ele As cEPRElement)
    Dim sKeyType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bNeeded As Boolean, bBeteenKeys As Boolean
    Dim bForce As Boolean, strOldTag As String
    
    bBeteenKeys = FindNextKey(edtThis, 0, "E", Ele.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bBeteenKeys Then
        Dim lngLen As Long, str内容 As String
        str内容 = Ele.内容文本
        lngLen = Len(str内容)
        With edtThis
            .Freeze
            strOldTag = .Tag
            .Tag = "EleToString"
            bForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = str内容
            .Range(lKSS, lKSS + lngLen).Font.Protected = False
            .Range(lKSS, lKSS + lngLen).Font.Hidden = False
            .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
            .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
            .ForceEdit = bForce
            .UnFreeze
            .Tag = strOldTag
        End With
    End If
End Sub

Private Function GetReplaceEleValue(ByVal ElementName As String, _
    ByVal sPatientID As String, _
    ByVal sPageID As String, _
    ByVal iPatientType As PatiFromEnum, _
    ByVal lng医嘱ID As Long, _
    ByVal lngBaby As Long) As String

    Dim rsTmp As ADODB.Recordset, strSql As String
    
    strSql = "Select Zl_Replace_Element_Value([1],[2],[3],[4],[5],[6]) From Dual"
    Err = 0: On Error GoTo DBError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取替换项", ElementName, CLng(sPatientID), _
        CLng(sPageID), CLng(iPatientType), lng医嘱ID, lngBaby)
    If rsTmp.EOF Or rsTmp.BOF Then
        GetReplaceEleValue = ""
    Else
        GetReplaceEleValue = Trim(IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value))
    End If
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    Dim clsZip As zlRichEPR.cZip
    
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    Set clsZip = New zlRichEPR.cZip
    
    With clsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
    Set clsZip = Nothing
End Function

Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim objFSO As New Scripting.FileSystemObject    'FSO对象
    Dim clsUnZip As zlRichEPR.cUnzip
    
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If objFSO.FileExists(strZipPath & "TMP.RTF") Then objFSO.DeleteFile strZipPath & "TMP.RTF"
    
    Set clsUnZip = New zlRichEPR.cUnzip
    With clsUnZip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
    Set clsUnZip = Nothing
End Function

Public Function FindNextKey(ByRef edtThis As Object, _
    ByVal lngCurPosition As Long, _
    ByVal strKeyType As String, _
    ByRef lngKey As Long, _
    ByRef lngKSS As Long, _
    ByRef lngKSE As Long, _
    ByRef lngKES As Long, _
    ByRef lngKEE As Long, _
    ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S("
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = IIf(lngCurPosition = 0, 1, lngCurPosition)
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.Range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E("
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.Range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                lngKey = Val(.Range(i + 2, i + 10))
                blnNeeded = -Val(.TOM.TextDocument.Range(i + 11, i + 12))
                FindNextKey = True
            End If
        End If
    End With
End Function

Public Function Get病种ID(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal lngType As Long, Optional ByVal lng科室ID As Long, Optional ByRef bln中医 As Boolean = False) As ADODB.Recordset
'参数： lngType=1  取病人除首要诊断之外的诊断（（入院非第一诊断或门诊诊断非第一诊断）或住院诊断的其他诊断、并发症诊断）
'             =0 默认 则和以前一样，按次序取入院、门诊、中医入院、中医门诊诊断;如果是中医科时, 优先级：中医入院、入院、中医门诊、门诊
'             =2则是第二次导入有效的首要路径表时，取所有入院、出院诊断
'             =3按次序取入院、门诊、中医入院、中医门诊诊断;如果是中医科时, 优先级：中医入院、入院、中医门诊、门诊（除开首要诊断）同时该诊断对应了首要路径。
'说明:需排除自由录入的诊断
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    gblnGetPath = Val(zlDatabase.GetPara(54, glngSys, 1261)) = 1
    If lngType = 0 Then
        bln中医 = sys.DeptHaveProperty(lng科室ID, "中医科")
        If bln中医 Then
            strSql = "Select 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源" & vbNewLine & _
                    "From 病人诊断记录" & vbNewLine & _
                    "Where 记录来源 In (1, 2, 3) And" & IIf(gblnGetPath, " 诊断类型 In (2,12)", " 诊断类型 In (1, 2, 11, 12)") & " And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 = 1 And" & vbNewLine & _
                    "      Nvl(是否疑诊, 0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                    "Order By Decode(诊断类型, 12, 1, 2, 2, 11, 3, 1, 4), Decode(记录来源, 1, 4, 记录来源) Desc"
        Else
            strSql = "Select 疾病id, 诊断id, 诊断描述,诊断类型,记录来源" & vbNewLine & _
                "From 病人诊断记录" & vbNewLine & _
                "Where 记录来源 In (1, 2, 3) And" & IIf(gblnGetPath, " 诊断类型 In (2,12)", " 诊断类型 In (1, 2, 11, 12)") & " And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 = 1 And" & vbNewLine & _
                "       Nvl(是否疑诊,0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
                "Order By Sign(诊断类型-10),诊断类型 Desc, Decode(记录来源, 1, 4, 记录来源) Desc"
        End If
    ElseIf lngType = 1 Then
        strSql = "Select a.疾病id, a.诊断id, a.诊断描述, a.诊断类型, a.记录来源" & vbNewLine & _
            "From 病人诊断记录 A" & vbNewLine & _
            "Where a.记录来源 In (1, 2, 3) And a.取消时间 Is Null And a.病人id = [1] And a.主页id = [2] And" & vbNewLine & _
            "(" & IIf(gblnGetPath, " 诊断类型 In (2, 3, 12,13)", " a.诊断类型 In (1, 2, 3, 11, 12,13)") & " And a.诊断次序 <> 1 Or" & vbNewLine & _
            "      a.诊断类型 = 10) And Nvl(a.是否疑诊, 0) = 0 And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
            "Order By Sign(a.诊断类型 - 10), a.诊断类型 Desc, Decode(a.记录来源, 1, 4, a.记录来源) Desc"
    ElseIf lngType = 2 Then
        strSql = "Select 疾病id, 诊断id, 诊断描述,诊断类型,记录来源" & vbNewLine & _
            "From 病人诊断记录" & vbNewLine & _
            "Where 记录来源 In (1, 2, 3) And 诊断类型 In ( 2,3,12,13) And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And Nvl(是否疑诊,0) = 0 " & vbNewLine & _
            "       And Not (NVl(疾病ID,0)=0 and NVl(诊断ID,0)=0) " & vbNewLine & _
            "Order By 诊断类型, Decode(记录来源, 1, 4, 记录来源) Desc"
    Else
        bln中医 = sys.DeptHaveProperty(lng科室ID, "中医科")
        If bln中医 Then
            strSql = "Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述, K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
                "From 临床路径目录 A, 临床路径病种 B, 临床路径版本 C," & vbNewLine & _
                "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
                "       From 病人诊断记录" & vbNewLine & _
                "       Where 记录来源 In (1, 2, 3) And" & IIf(gblnGetPath, " 诊断类型 In (2,12)", " 诊断类型 In (1, 2, 11, 12)") & " And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
                "             Nvl(是否疑诊, 0) = 0 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
                "       Order By Decode(诊断类型, 12, 1, 2, 2, 11, 3, 1, 4), Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
                "Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
                "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
                "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
                "Order By k.排序"
        Else
            strSql = "Select Distinct a.Id, k.疾病id, k.诊断id, k.诊断描述,K.诊断类型, K.记录来源,k.排序 " & vbNewLine & _
            "From 临床路径目录 A, 临床路径病种 B, 临床路径版本 C," & vbNewLine & _
            "     (Select Rownum As 排序, 疾病id, 诊断id, 诊断描述, 诊断类型, 记录来源 " & vbNewLine & _
            "       From 病人诊断记录" & vbNewLine & _
            "       Where 记录来源 In (1, 2, 3) And" & IIf(gblnGetPath, " 诊断类型 In (2,12)", " 诊断类型 In (1, 2, 11, 12)") & " And 取消时间 Is Null And 病人id = [1] And 主页id = [2] And 诊断次序 <> 1 And" & vbNewLine & _
            "             Nvl(是否疑诊, 0) = 0 And Not (Nvl(疾病id, 0) = 0 And Nvl(诊断id, 0) = 0)" & vbNewLine & _
            "       Order By Sign(诊断类型 - 10), 诊断类型 Desc, Decode(记录来源, 1, 4, 记录来源) Desc, 诊断次序) K" & vbNewLine & _
            "Where a.Id = b.路径id And a.Id = b.路径id And a.Id = c.路径id And a.最新版本 = c.版本号 And a.性质 = 0 And b.性质 = 0 And" & vbNewLine & _
            "      (b.疾病id = k.疾病id Or b.诊断id = k.诊断id) And" & vbNewLine & _
            "      (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & vbNewLine & _
            "Order By k.排序"
        End If
    End If
    '记录来源:1-病历；2-入院登记；3-首页整理;4-病案
    '诊断类型:1-西医门诊诊断;2-西医入院诊断;11-中医门诊诊断;12-中医入院诊断
    '有多个诊断的情况下，根据诊断次序，只取第一个主要诊断
    '病历里面的诊断优先，主要是为了支持修正诊断。
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病种", lng病人ID, lng主页ID, lng科室ID)
    Set Get病种ID = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetPathTable(ByVal lng疾病ID As Long, ByVal lng诊断ID As Long, ByVal lng科室ID As Long, ByVal lngPathID As Long, Optional ByVal str疾病IDs As String, Optional ByVal lng病人路径Id As Long, _
            Optional ByVal str诊断IDs As String, Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long) As ADODB.Recordset
'参数： str疾病IDs<>"" 合并路径导入时存在多个诊断时
'       lng病人ID<>0 第二次导入有效路径的时候，根据所有入院、出院诊断进行导入，排开已经导入过的路径，=0 导入合并路径
    Dim strSql As String
    
    If str疾病IDs = "" And str诊断IDs = "" Then
        '这里加Distinct是因为，诊断id和疾病id做了绑定对应，所以，查出来会有重复值
        strSql = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, Nvl(a.适用病情,'通用') 适用病情, a.适用性别, a.适用年龄, a.最新版本, c.标准住院日,Nvl(a.病例分型,'无') as 病例分型,Nvl(a.确诊天数,0) as 确诊天数" & vbNewLine & _
                "From 临床路径目录 A, 临床路径病种 B,临床路径版本 C" & vbNewLine & _
                "Where a.Id = b.路径id And (b.疾病id = [1] Or b.诊断id = [2]) And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                "And a.Id = c.路径id And a.性质=0 And b.性质=" & IIf(lngPathID = 0, "0", "1") & " And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & _
                 IIf(lngPathID = 0, "", " And a.id<>[4]")
    Else
        If lng病人ID = 0 Then
            '这里加Distinct是因为，诊断id和疾病id做了绑定对应，所以，查出来会有重复值，排开已经导入了的合并路径
            strSql = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, Nvl(a.适用病情,'通用') 适用病情, a.适用性别, a.适用年龄, a.最新版本, c.标准住院日,Nvl(a.病例分型,'无') as 病例分型,Nvl(a.确诊天数,0) as 确诊天数,b.疾病ID,b.诊断ID" & vbNewLine & _
                    "From 临床路径目录 A, 临床路径病种 B,临床路径版本 C" & vbNewLine & _
                    "Where a.Id = b.路径id And (instr(',' || [5] || ',',',' || b.疾病ID || ',')>0 and [5] is not null Or instr(',' || [7] || ',',',' || b.诊断ID || ',')>0 and [7] is not null)  And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                    "And a.Id = c.路径id And a.性质=1 And b.性质=0 And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & _
                    " And Not Exists(Select 1 From 病人合并路径 D Where a.id=d.路径ID  and d.首要路径记录ID=[6])"
        Else
            strSql = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.说明, Nvl(a.适用病情,'通用') 适用病情, a.适用性别, a.适用年龄, a.最新版本, c.标准住院日,Nvl(a.病例分型,'无') as 病例分型,Nvl(a.确诊天数,0) as 确诊天数" & vbNewLine & _
                    "From 临床路径目录 A, 临床路径病种 B,临床路径版本 C" & vbNewLine & _
                    "Where a.Id = b.路径id And (instr(',' || [5] || ',',',' || b.疾病ID || ',')>0 and [5] is not null Or instr(',' || [7] || ',',',' || b.诊断ID || ',')>0 and [7] is not null)  And a.最新版本 is not null And a.id = b.路径ID And a.最新版本 = c.版本号" & vbNewLine & _
                    "And a.Id = c.路径id And a.性质=0 And b.性质=" & IIf(lngPathID = 0, "0", "1") & " And (a.通用 = 1 Or a.通用 = 2 And Exists (Select 1 From 临床路径科室 D Where a.Id = d.路径id And d.科室id = [3]))" & _
                    " And Not Exists(Select 1 From 病人临床路径 D Where a.ID=d.路径ID And d.病人ID=[8] And D.主页ID=[9])" & _
                    IIf(lngPathID = 0, "", " And a.id<>[4]")
        End If
    End If
    On Error GoTo errH
    Set GetPathTable = zlDatabase.OpenSQLRecord(strSql, "读取路径目录", lng疾病ID, lng诊断ID, lng科室ID, lngPathID, str疾病IDs, lng病人路径Id, str诊断IDs, lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathOutLog() As Boolean
'功能：检查是否存在病人出径登记项目
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 路径报表结构 Where 报表ID = 2 And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取路径报表结构")
    CheckPathOutLog = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathOutDiag(ByVal lng路径ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：如果填写了出院诊断，则判断出院诊断是否和导入路径的诊断相同
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim str疾病IDs As String, str诊断IDs As String, lngDiagType As Long '诊断类型：1/2  西医  ，11.12   中医
 
    strSql = "Select b.疾病ID,b.诊断ID,a.诊断类型 From 病人临床路径 A,临床路径病种 B Where A.路径ID=B.路径ID And A.ID = [1] And NVL(b.性质,0)=0"
    
    CheckPathOutDiag = True
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断出院诊断", lng路径ID)
    If rsTmp.RecordCount > 0 Then
        lngDiagType = Val(rsTmp!诊断类型 & "")
        Do While Not rsTmp.EOF
            If Val(rsTmp!疾病id & "") <> 0 Then
                str疾病IDs = str疾病IDs & "," & Val(rsTmp!疾病id & "")
            End If
            If Val(rsTmp!诊断id & "") <> 0 Then
                str诊断IDs = str诊断IDs & "," & Val(rsTmp!诊断id & "")
            End If
            rsTmp.MoveNext
        Loop
        str疾病IDs = Mid(str疾病IDs, 2)
        str诊断IDs = Mid(str诊断IDs, 2)
        
        strSql = "Select 疾病ID,诊断ID From 病人诊断记录 Where 诊断次序=1 And NVL(编码序号,1) = 1 and 记录来源=3 And 病人ID=[1] And 主页ID=[2]"
        If lngDiagType = 1 Or lngDiagType = 2 Then
            strSql = strSql & " and 诊断类型=3"
        Else
            strSql = strSql & " and 诊断类型=13"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断出院诊断", lng病人ID, lng主页ID)
        '如果没填出院诊断，则不检查
        If rsTmp.RecordCount > 0 Then
            If InStr("," & str疾病IDs & ",", "," & Val(rsTmp!疾病id & "") & ",") = 0 And InStr("," & str诊断IDs & ",", "," & Val(rsTmp!诊断id & "") & ",") = 0 Then
                CheckPathOutDiag = False
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人病案状态(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional lng婴儿科室ID As Long, Optional lng婴儿病区ID As Long) As Long
'功能：获取病人病案提交状态
'      0-未提交;1-等待审查(提交);2-拒绝审查;3-正在审查;4-审查反馈;5-审查归档;6-审查整改;13-正在抽查;14-抽查反馈;16-抽查整改
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 病案状态,婴儿科室ID,婴儿病区ID From 病案主页 Where 病人ID = [1] And 主页ID = [2]"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病案提交状态", lng病人ID, lng主页ID)
    If rsTmp.RecordCount > 0 Then
        Get病人病案状态 = Val("" & rsTmp!病案状态)
        lng婴儿科室ID = Val(rsTmp!婴儿科室ID & "")
        lng婴儿病区ID = Val(rsTmp!婴儿病区ID & "")
    Else
        lng婴儿科室ID = 0
        lng婴儿病区ID = 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPatiPathOutLog(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查是否存在病人出径记录
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select 1 From 病人出径记录 Where 病人ID = [1] And 主页ID = [2] And Rownum=1"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取病人出径记录", lng病人ID, lng主页ID)
    CheckPatiPathOutLog = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get阶段分类(Optional ByVal lng路径记录ID As Long, Optional ByVal lng阶段ID As Long) As String
'功能：获取病人使用过的阶段的分类，只有分支路径才有分类，如果使用了该分类，则病人整个路径期间只能选择该分类，所有只可能有一个分类
'参数：lng阶段ID=指定该参数时，获取指定阶段的分类
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If lng阶段ID <> 0 Then
        strSql = "Select 分类 From 临床路径阶段 Where id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取阶段分类", lng阶段ID)
    Else
        strSql = "Select a.分类" & vbNewLine & _
                "From 临床路径阶段 A, (Select Distinct 阶段id From 病人路径执行 Where 路径记录id = [1]) B" & vbNewLine & _
                "Where a.Id = b.阶段id And a.分类 Is Not Null And rownum<2"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取阶段分类", lng路径记录ID)
    End If
    If rsTmp.RecordCount > 0 Then Get阶段分类 = "" & rsTmp!分类
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDiagnose(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int来源 As Integer) As ADODB.Recordset
'功能：读取病人的首要诊断或次要诊断
'参数：lng就诊ID=挂号ID或主页ID
'      int来源=1-门诊,2-住院
'返回：结果记录集
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select 疾病ID,b.编码,记录来源,Mod(诊断类型,10) as 大类 From 病人诊断记录 a ,疾病编码目录 b" & _
        " Where 病人ID=[1] And 主页ID=[2] And NVL(A.编码序号,1) = 1 And 诊断类型 IN(" & IIf(int来源 = 1, "1,11", "1,2,3,11,12,13") & ") and a.疾病ID=b.ID and 疾病ID is not null and nvl(是否疑诊,0)=0" & _
        " Order by 记录来源,诊断类型,诊断次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "GetPatiDiagnose", lng病人ID, lng就诊ID)
    
    '先按来源优先顺序过滤
    rsTmp.Filter = "记录来源=3" '首页整理
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=2" '入院登记
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=1" '病历
    If rsTmp.EOF Then rsTmp.Filter = "记录来源=4" '病案室录入
    
    '住院再按类型优先顺序过滤
    If Not rsTmp.EOF And int来源 = 2 Then
        strSql = rsTmp.Filter
        rsTmp.Filter = strSql & " And 大类=3"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And 大类=2"
        If rsTmp.EOF Then rsTmp.Filter = strSql & " And 大类=1"
    End If
    
    Set GetPatiDiagnose = zl9ComLib.zlDatabase.CopyNewRec(rsTmp)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckPathSend(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：检查该病人当次住院是否生成过项目
'返回：true=生成过，false=未生成过
    Dim strSql As String, rsPati As Recordset
    
    strSql = "Select Max(状态) as 状态 From 病人临床路径 Where 病人ID=[1] And 主页ID=[2]"
    On Error GoTo errH
    Set rsPati = zlDatabase.OpenSQLRecord(strSql, "CheckPathSend", lng病人ID, lng主页ID)
    If rsPati.RecordCount > 0 Then
        If Val(rsPati!状态 & "") <> 0 Then CheckPathSend = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub AddOutPathItem(ByVal strAdviceIDs As String, ByVal lngMode As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
   Optional ByVal bytType As Byte, Optional ByRef colSQL As Collection)
'功能:将医生没有停止的长嘱，生成为路径外项目
'参数：strAdviceIds - id,id...lngMode=1:传人医嘱ID中包停止和未停止的长期医嘱，
'                             lngMode=2：所有回退的医嘱ID(回退停止长期医嘱,回退作废 长期|临时医嘱)
'       lngMode：1-路径生成 ，
'                2-医生护士站，医嘱状态为：作废或停止的医嘱在回退时调用
'       bytType =当lngMode=2： =4 回退作废;=8 回退停止
'       colSQL =返回可执行SQL

    Dim strSql As String, strStopIds As String, strPathOut As String
    Dim rsTmp       As ADODB.Recordset
    Dim AddDate     As Date
    Dim strDate, strAddDate As String
    Dim i           As Long, j As Long
    Dim str变异原因     As String
    Dim blnTrans    As Boolean
    Dim lng病人路径Id, lng阶段ID, lng天数 As Long
    Dim varTemp As Variant
    Dim strItemType As String
    Set colSQL = New Collection
    On Error GoTo errH
    'a.根据传人的strAdviceIds查询没有停止的长嘱
    If lngMode = 1 Then
        strSql = "Select /*+ rule*/" & vbNewLine & _
                 " Column_Value As 病人医嘱id" & vbNewLine & _
                 "From Table(f_Num2list([1])) A, 病人医嘱记录 B" & vbNewLine & _
                 "Where a.Column_Value = b.Id And b.停嘱时间 Is Null"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strAdviceIDs)

        '获取未停止的长嘱ID
        For i = 1 To rsTmp.RecordCount
            strStopIds = strStopIds & "," & rsTmp!病人医嘱id
            rsTmp.MoveNext
        Next
        strStopIds = Mid(strStopIds, 2)
    ElseIf lngMode = 2 Then    '作废或停止的医嘱回退时，传人医嘱 组ID
        '需要停止的医嘱：排除当天路径中已经生成过的
        strSql = "Select f_List2str(Cast(Collect(t.病人医嘱id || '') As t_Strlist)) As 病人医嘱ids" & vbNewLine & _
                 "From (Select a.Column_Value As 病人医嘱id" & vbNewLine & _
                 "       From Table(f_Num2list([1])) A" & vbNewLine & _
                 "       Minus" & vbNewLine & _
                 "       Select d.病人医嘱id" & vbNewLine & _
                 "       From Table(f_Num2list([1])) C, 病人路径医嘱 D, 病人路径执行 E, 病人临床路径 F" & vbNewLine & _
                 "       Where d.病人医嘱id = c.Column_Value And d.路径执行id = e.Id And f.Id = e.路径记录id And f.当前阶段id = e.阶段id And f.当前天数 = e.天数) T"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strAdviceIDs)
        If rsTmp.RecordCount = 1 Then strStopIds = rsTmp!病人医嘱ids & ""
    End If
    
    If strStopIds <> "" Then
        '过滤掉路径内的医嘱ID(1-存在路径内项目，则自动生成为路径内)返回路径外医嘱ID
        Call CheckStopAdvice(lng病人ID, lng主页ID, strStopIds, colSQL)
        If strAdviceIDs = "" Then Exit Sub
        '获得其他的变异原因编码
        strSql = "Select 编码 From 变异常见原因 Where (名称='其它' Or 名称='其他')  And 性质=1 And rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem")
        If rsTmp.RecordCount > 0 Then
            str变异原因 = rsTmp!编码 & ""
        End If

        '获取当前路径：路径记录ID,当前阶段Id,当前日期，天数
        strSql = "Select a.路径记录id, a.当前阶段id, a.当前天数, b.日期 " & vbNewLine & _
                 "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, Max(b.Id) 执行id" & vbNewLine & _
                 "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
                 "       Where a.病人id = [1] And a.主页id = [2] And a.Id = b.路径记录id And b.阶段id = a.当前阶段id And b.天数 = a.当前天数" & vbNewLine & _
                 "       Group By a.Id, a.当前阶段id, a.当前天数) A, 病人路径执行 B" & vbNewLine & _
                 "Where a.执行id = b.Id"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", lng病人ID, lng主页ID)

        If rsTmp.RecordCount = 1 Then
            lng病人路径Id = Val(rsTmp!路径记录id)
            lng阶段ID = Val(rsTmp!当前阶段ID)
            strDate = "To_Date('" & Format(rsTmp!日期, "yyyy-MM-dd") & "','YYYY-MM-DD')"
            lng天数 = Val(rsTmp!当前天数)
        Else
            Exit Sub
        End If

        strSql = "Select a.Id, a.分类, Decode(a.项目id, Null, a.项目内容, c.项目内容) As 项目内容, Decode(a.项目id, Null, a.执行者, c.执行者) As 执行者," & vbNewLine & _
                 "         Decode(a.项目id, Null, a.项目结果, c.项目结果) As 项目结果, Decode(a.项目id, Null, a.图标id, c.图标id) As 图标id," & vbNewLine & _
                 "         f_List2str(Cast(Collect(b.病人医嘱id || '') As t_Strlist)) As 病人医嘱ids" & vbNewLine & _
                 "  From (Select " & vbNewLine & _
                 "          Row_Number() Over(Partition By d.病人医嘱id Order By d.路径执行id Desc) As Top, d.路径执行id, d.病人医嘱id" & vbNewLine & _
                 "         From 病人路径医嘱 D, Table(f_Num2list([1])) E" & vbNewLine & _
                 "         Where d.病人医嘱id = e.Column_Value" & vbNewLine & _
                 "         Group By d.路径执行id, d.病人医嘱id) B, 病人路径执行 A, 临床路径项目 C" & vbNewLine & _
                 "  Where b.Top = 1 And b.路径执行id = a.Id And a.项目id = c.Id(+)" & vbNewLine & _
                 "  Group By a.Id, a.分类, Decode(a.项目id, Null, a.项目内容, c.项目内容), Decode(a.项目id, Null, a.执行者, c.执行者)," & vbNewLine & _
                 "           Decode(a.项目id, Null, a.项目结果, c.项目结果), Decode(a.项目id, Null, a.图标id, c.图标id)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "AddOutPathItem", strStopIds)

        AddDate = zlDatabase.Currentdate
        strAddDate = "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        '将前一阶段或前一天，未停止的医嘱作为路径外项目，添加到路径中
        For i = 1 To rsTmp.RecordCount
            strSql = "Zl_病人路径生成_Insert(1," & lng病人ID & "," & lng主页ID & ",Null,0," & _
                                     lng病人路径Id & "," & lng阶段ID & "," & strDate & "," & lng天数 & ",'" & NVL(rsTmp!分类) & "',Null" & ",'" & rsTmp!病人医嘱ids & "',Null,Null" & _
                                     ",'" & UserInfo.姓名 & "'," & strAddDate & ",'" & CStr(NVL(rsTmp!项目内容)) & "'" & _
                                     "," & Val(NVL(rsTmp!执行者, 1)) & ",'" & CStr(NVL(rsTmp!项目结果)) & "'," & NVL(rsTmp!图标ID, "Null") & ",'未停用的长嘱','" & str变异原因 & "' ,0)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
            AddDate = AddDate + 1 / 24 / 60 / 60
            '提取未增加成功的医嘱生成路径外项目
            varTemp = Split(rsTmp!病人医嘱ids, ",")
            For j = LBound(varTemp) To UBound(varTemp)
                strStopIds = Replace("," & strStopIds & ",", "," & varTemp(j) & ",", ",")
                If Left(strStopIds, 1) = "," Then strStopIds = Mid(strStopIds, 2)
                If Right(strStopIds, 1) = "," Then strStopIds = Mid(strStopIds, 1, Len(strStopIds) - 1)
            Next
            rsTmp.MoveNext
        Next
        If strStopIds <> "" Then
            Call GetPatiPathInfo(lng病人ID, lng主页ID, strItemType)
            strSql = "Zl_病人路径生成_Insert(1," & lng病人ID & "," & lng主页ID & ",Null,0," & _
                         lng病人路径Id & "," & lng阶段ID & "," & strDate & "," & lng天数 & ",'" & strItemType & "',Null" & ",'" & strStopIds & "',Null,Null" & _
                         ",'" & UserInfo.姓名 & "'," & strAddDate & ",'路径外项目'" & _
                         ",1,'已经执行|1" & vbTab & "已经执行',NULL,NULL,'" & str变异原因 & "' ,1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
        End If
        '提交数据,开启事务
        If lngMode = 1 Then
            gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.count
                Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "AddOutPathItem")
            Next
            gcnOracle.CommitTrans: blnTrans = False
        End If
    End If
    Exit Sub
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CheckStopAdvice(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByRef strUnStopIDs As String, Optional ByRef colSQL As Collection)
'功能:过滤掉能够匹配为路径内医嘱的医嘱ID
'参数:
'   strUnStopIDs-医嘱IDS
'出参:
'   strUnStopIDs-未停止的医嘱ID（一组医嘱的所有ID）返回要添加的路径外项目
'   colSQL-返回可执行SQL

    Dim rsUnStop As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset

    Dim strSql As String
    Dim i As Long, j As Long
    Dim k As Long

    Dim lng病人路径Id  As Long
    Dim lng阶段ID As Long
    Dim lng天数 As Long
    Dim lngPos As Long, lngPathPos As Long
    Dim strDate As String
    Dim strTag As String
    Dim str相关ID As String
    Dim AddDate As Date
    
    Dim blnTrans As Boolean
    Dim str医嘱ID As String
    Dim bln匹配期效 As Boolean
    
    
    On Error GoTo errH
    
    Set colSQL = New Collection
    
    bln匹配期效 = CBool(zlDatabase.GetPara("匹配时期效不同算路径外项目", glngSys, P临床路径应用, "0"))
    '获取当前路径：路径记录ID,当前阶段Id,当前日期，天数
    strSql = "Select a.路径记录id, a.当前阶段id, a.当前天数, b.日期 " & vbNewLine & _
             "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, Max(b.Id) 执行id" & vbNewLine & _
             "       From 病人临床路径 A, 病人路径执行 B" & vbNewLine & _
             "       Where a.病人id = [1] And a.主页id = [2] And a.Id = b.路径记录id And b.阶段id = a.当前阶段id And b.天数 = a.当前天数" & vbNewLine & _
             "       Group By a.Id, a.当前阶段id, a.当前天数) A, 病人路径执行 B" & vbNewLine & _
             "Where a.执行id = b.Id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng病人ID, lng主页ID)

    If rsTmp.RecordCount = 1 Then
        lng病人路径Id = Val(rsTmp!路径记录id)
        lng阶段ID = Val(rsTmp!当前阶段ID)
        strDate = "To_Date('" & Format(rsTmp!日期, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        lng天数 = Val(rsTmp!当前天数)
    Else
        Exit Sub
    End If

    strSql = "select a.ID, a.相关ID, b.类别, a.诊疗项目ID, b.操作类型,a.医嘱期效 " & vbNewLine & _
            "  from 病人医嘱记录 a, 诊疗项目目录 b" & vbNewLine & _
            " where a.诊疗项目ID = b.id" & vbNewLine & _
            "   and a.id in (Select Column_Value As 病人医嘱id" & vbNewLine & _
            "                  From Table(f_Num2list([1]))) Order by a.序号"


    Set rsUnStop = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", strUnStopIDs)

    strSql = "Select c.ID, c.相关ID,c.诊疗项目id,a.id as 路径项目ID,a.分类 as 路径项目分类,c.期效 " & vbNewLine & _
            "From 临床路径项目 a, 临床路径医嘱 b, 路径医嘱内容 c" & vbNewLine & _
            "where a.id = b.路径项目id" & vbNewLine & _
            "   and b.医嘱内容id = c.id" & vbNewLine & _
            "   and a.阶段id = [1]"
            
    Set rsPath = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng阶段ID)

    strTag = ""
    Set rsPathAdvice = Nothing
    For i = 1 To rsUnStop.RecordCount
        lngPos = rsUnStop.AbsolutePosition
        If Val(rsUnStop!相关id & "") = 0 And Not (rsUnStop!类别 & "" = "E" And rsUnStop!操作类型 & "" = "2") Or InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
            '在一并给药中增加一行时，只检查和设置当前行，因为路径外项目可能和路径内项目一并给药
            If InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
                '药品单个药进行匹配 65982
                rsUnStop.Filter = "ID=" & rsUnStop!ID
                str相关ID = Val(rsUnStop!相关id & "")
            Else
                rsUnStop.Filter = "ID=" & rsUnStop!ID & " Or 相关ID=" & rsUnStop!ID
                str相关ID = Val(rsUnStop!ID & "")
            End If
            '药品不含给药途径、用法、煎法，输血不含途径,9-输血采集,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
            If Not (rsUnStop!类别 & "" = "E" And InStr(",2,3,4,6,8,9,", "," & rsUnStop!操作类型 & ",") > 0) _
                And Not (InStr(",G,F,D,", "," & rsUnStop!类别 & ",") > 0 And Val(rsUnStop!相关id & "") <> 0) Then
                If bln匹配期效 Then
                    rsPath.Filter = "诊疗项目id=" & NVL(rsUnStop!诊疗项目ID, 0) & " And 期效 = " & rsUnStop!医嘱期效
                Else
                    rsPath.Filter = "诊疗项目id=" & NVL(rsUnStop!诊疗项目ID, 0)
                    If rsUnStop!医嘱期效 = 0 Then  '优先按期效匹配路径
                        rsPath.Sort = "期效 ASC"
                    Else
                        rsPath.Sort = "期效 DESC"
                    End If
                End If
                If rsPath.RecordCount > 0 Then
                    '路径内项目
                    If InStr("," & strTag & ",", "," & str相关ID & ",") = 0 Then
                        rsUnStop.Filter = "相关ID=" & str相关ID & " OR ID =" & str相关ID
                        If InStr(",5,6,", "," & rsUnStop!类别 & ",") > 0 Then
                            strTag = strTag & "," & rsUnStop!相关id
                        Else
                            strTag = strTag & "," & rsUnStop!ID
                        End If
                        
                        If rsPathAdvice Is Nothing Then Set rsPathAdvice = MakePathAdivceRS
                        For k = 1 To rsUnStop.RecordCount
                            rsPathAdvice.Filter = "路径项目ID = " & rsPath!路径项目ID
                            If rsPathAdvice.RecordCount = 0 Then
                                rsPathAdvice.AddNew
                                rsPathAdvice!路径项目ID = rsPath!路径项目ID & ""
                                rsPathAdvice!路径项目分类 = rsPath!路径项目分类 & ""
                                rsPathAdvice!医嘱IDs = rsUnStop!ID & ""
                            Else
                                rsPathAdvice!医嘱IDs = rsPathAdvice!医嘱IDs & "," & rsUnStop!ID
                            End If
                            rsPathAdvice.Update
                            '从未停止的长嘱中移除
                            strUnStopIDs = Replace("," & strUnStopIDs & ",", "," & rsUnStop!ID & ",", ",")
                            If Left(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 2)
                            If Right(strUnStopIDs, 1) = "," Then strUnStopIDs = Mid(strUnStopIDs, 1, Len(strUnStopIDs) - 1)
                            rsUnStop.MoveNext
                        Next
                    End If
                End If
            End If
        End If
        rsUnStop.Filter = ""
        rsUnStop.AbsolutePosition = lngPos
        rsUnStop.MoveNext
    Next
    
    If rsPathAdvice Is Nothing Then Exit Sub
    rsPathAdvice.Filter = ""
    AddDate = zlDatabase.Currentdate
    For j = 1 To rsPathAdvice.RecordCount
        strSql = "Zl_病人路径生成_Insert(1," & lng病人ID & "," & lng主页ID & ",NULL,0," & lng病人路径Id & "," & lng阶段ID & _
            "," & strDate & "," & lng天数 & ",'" & rsPathAdvice!路径项目分类 & "'," & rsPathAdvice!路径项目ID & ",'" & rsPathAdvice!医嘱IDs & "',Null,Null" & _
            ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
            
        colSQL.Add strSql, "C" & colSQL.count + 1
        '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
        AddDate = AddDate + 1 / 24 / 60 / 60
        rsPathAdvice.MoveNext
    Next
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetPatiPathInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByRef str路径项目分类 As String = "-1") As ADODB.Recordset
'功能：获取路径病人当前路径信息
'返回：str分类=当前天数最后一个路径项目所属的分类
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rsRet As ADODB.Recordset
    Dim blnDo As Boolean
    
    blnDo = str路径项目分类 <> "-1"
    str路径项目分类 = ""
    strSql = "Select a.路径记录id, a.当前阶段id, a.当前天数, a.路径ID, a.版本号, a.开始日期, b.日期, b.分类" & vbNewLine & _
            "From (Select a.Id As 路径记录id, a.当前阶段id, a.当前天数, a.路径ID, a.版本号, Max(b.Id) 执行id, Min(c.日期) As 开始日期" & vbNewLine & _
            "       From 病人临床路径 A, 病人路径执行 B, 病人路径执行 C" & vbNewLine & _
            "       Where a.Id = b.路径记录id And a.Id = c.路径记录id And b.阶段id + 0 = a.当前阶段id And b.天数 = a.当前天数 And a.状态 = 1 And" & vbNewLine & _
            "             a.病人id = [1] And a.主页id = [2]" & vbNewLine & _
            "       Group By a.Id, a.当前阶段id, a.当前天数, a.路径ID, a.版本号) A, 病人路径执行 B" & vbNewLine & _
            "Where a.执行id = b.Id"

    On Error GoTo errH
    Set rsRet = zlDatabase.OpenSQLRecord(strSql, "病人当前路径信息", lng病人ID, lng主页ID)
    If rsRet.RecordCount > 0 And blnDo Then
        str路径项目分类 = "" & rsRet!分类
        
        '如果当天生成了医嘱类项目，则取医嘱类项目的分类
        strSql = "Select 分类" & vbNewLine & _
                "From 病人路径执行" & vbNewLine & _
                "Where ID = (Select Max(ID)" & vbNewLine & _
                "            From 病人路径执行 A" & vbNewLine & _
                "            Where a.路径记录id = [1] And a.阶段id = [2] And a.日期 = [3] And Exists (Select 1 From 病人路径医嘱 B Where a.Id = b.路径执行id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人当前路径信息", Val(rsRet!路径记录id), Val(rsRet!当前阶段ID), CDate(rsRet!日期))
        If rsTmp.RecordCount > 0 Then
            str路径项目分类 = "" & rsTmp!分类
        End If
    End If
    Set GetPatiPathInfo = rsRet
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long, Optional ByVal int场合 As Integer) As Boolean
'功能：外挂创建与检查
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod, int场合)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Function FuncGetEMRInfo(ByVal strInfo As String) As ADODB.Recordset
'功能:将修改后的电子病历信息解析成记录集返回。
'参数: 病历详情：格式：文件ID,原型ID,文件名称,序号;文件ID,原型ID,文件名称,序号...
'说明：修改电子病历后,鼠标移到路径项目上无法显示修改内容,因为新版电子病历内容还没有插入到标准库中
    Dim rsEMR As ADODB.Recordset
    Dim i As Long
    Dim arrtmp As Variant
    Dim arrTmpSub As Variant
    
    Set rsEMR = New ADODB.Recordset
    rsEMR.Fields.Append "文件ID", adBigInt
    rsEMR.Fields.Append "原型ID", adVarChar, 32
    rsEMR.Fields.Append "名称", adVarChar, 100
    rsEMR.Fields.Append "序号", adVarChar, 10
    rsEMR.Fields.Append "版本", adVarChar, 2
    
    rsEMR.CursorLocation = adUseClient
    rsEMR.LockType = adLockOptimistic
    rsEMR.CursorType = adOpenStatic
    rsEMR.Open
    arrtmp = Split(strInfo, ";")
    For i = LBound(arrtmp) To UBound(arrtmp)
        arrTmpSub = Split(arrtmp(i), ",")
        rsEMR.AddNew
        rsEMR!文件ID = Val(arrTmpSub(0))
        rsEMR!原型ID = arrTmpSub(1)
        rsEMR!名称 = arrTmpSub(2)
        rsEMR!序号 = arrTmpSub(3)
        If Val(arrTmpSub(0)) = 0 Then
            rsEMR!版本 = 2
        Else
            rsEMR!版本 = 1
        End If
        rsEMR.Update
    Next
    If rsEMR.RecordCount <> 0 Then rsEMR.MoveFirst
    Set FuncGetEMRInfo = rsEMR
End Function

Public Sub ZLHIS_CIS_001(ByRef objMip As zl9ComLib.clsMipModule, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByVal lng病区ID As Long, ByVal lng科室ID As Long)
'功能：发送医嘱新下达消息  ZLHIS_CIS_001，路径医嘱生成和路径外医嘱添加
    Dim strXML As String
    Dim strSql As String, strTmp As String
    Dim str类别 As String, str就诊科室名称 As String
    Dim bln消息平台 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long
    
    On Error GoTo errH
    
    If Not objMip Is Nothing Then
        If objMip.IsConnect Then bln消息平台 = True
    End If
    
    If bln消息平台 Then
        strSql = "select 名称 from 部门表 where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng科室ID)
        str就诊科室名称 = rsTmp!名称 & ""
    End If
    
    '可以校对的医嘱
    strSql = "select id,紧急标志 from 病人医嘱记录 a where A.医嘱状态=1 and a.病人id=[1] and a.主页id=[2]" & _
        " And Nvl(A.审核状态,0) Not in(1,3,4,5) And Exists ( Select M.姓名 From 人员表 M,执业类别 N" & _
        " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
        " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')) And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng病人ID, lng主页ID)
    If Not rsTmp.EOF Then
        lng医嘱ID = Val(rsTmp!ID & "")
        rsTmp.Filter = "紧急标志=1"
        If Not rsTmp.EOF Then
            lng医嘱ID = Val(rsTmp!ID & "")
        End If
    End If
    
    If lng医嘱ID = 0 Then Exit Sub
    
    '取一条医嘱即可，紧急医嘱优先
    strSql = "select a.id as 医嘱id,a.病人来源,decode(a.紧急标志,1,1,0) as 紧急标志,a.医嘱期效,a.诊疗类别,b.操作类型,a.开嘱医生," & vbNewLine & _
        " to_char(a.开嘱时间,'yyyy-mm-dd hh24:mi:ss') as 开嘱时间,a.开嘱科室id,c.名称" & vbNewLine & _
        " from 病人医嘱记录 a,诊疗项目目录 b,部门表 c where a.id=[1] and a.诊疗项目id=b.id(+) and a.开嘱科室id=c.id"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISPath", lng医嘱ID)
    
    str类别 = rsTmp!诊疗类别 & ""
    If rsTmp!诊疗类别 & "" = "E" Then
        If rsTmp!操作类型 & "" = "2" Then
            str类别 = "5"
        ElseIf rsTmp!操作类型 & "" = "4" Then
            str类别 = "7"
        ElseIf rsTmp!操作类型 & "" = "6" Then
            str类别 = "C"
        End If
    End If
    
    strXML = "": strTmp = ""
    
    strXML = strXML & "<patient_info>"
    strXML = strXML & "   <patient_id>" & lng病人ID & "</patient_id>"
    strXML = strXML & "</patient_info>"
    strXML = strXML & "<patient_clinic>"
    strXML = strXML & "   <patient_source>" & rsTmp!病人来源 & "</patient_source>"
    strXML = strXML & "   <clinic_id>" & lng主页ID & "</clinic_id>"
    strXML = strXML & "   <clinic_area_id>" & lng病区ID & "</clinic_area_id>"
    strXML = strXML & "   <clinic_dept_id>" & lng科室ID & "</clinic_dept_id>"
    strXML = strXML & "   <clinic_dept_title>" & str就诊科室名称 & "</clinic_dept_title>"
    strXML = strXML & "</patient_clinic>"
    strXML = strXML & "<new_order>"
    strXML = strXML & "   <order_id>" & rsTmp!医嘱id & "</order_id>"
    strXML = strXML & "   <order_urgency>" & rsTmp!紧急标志 & "</order_urgency>"
    strXML = strXML & "   <order_expiry>" & rsTmp!医嘱期效 & "</order_expiry>"
    strXML = strXML & "   <order_kind>" & str类别 & "</order_kind>"
    strXML = strXML & "   <operation_kind>" & rsTmp!操作类型 & "</operation_kind>"
    strXML = strXML & "   <create_doctor>" & rsTmp!开嘱医生 & "</create_doctor>"
    strXML = strXML & "   <create_time>" & rsTmp!开嘱时间 & "</create_time>"
    strXML = strXML & "   <create_dept_id>" & rsTmp!开嘱科室id & "</create_dept_id>"
    strXML = strXML & "   <create_dept_title>" & rsTmp!名称 & "</create_dept_title>"
    strXML = strXML & "</new_order>"
    
    If bln消息平台 Then Call objMip.CommitMessage("ZLHIS_CIS_001", strXML, strTmp)
    
    Call zlDatabase.SendMsg("ZLHIS_CIS_001", IIf(strTmp = "", strXML, strTmp))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function MakePathItems() As ADODB.Recordset
    Set MakePathItems = New ADODB.Recordset
    
    MakePathItems.Fields.Append "ID", adBigInt
    MakePathItems.Fields.Append "婴儿", adBigInt
    MakePathItems.Fields.Append "相关id", adBigInt
    MakePathItems.Fields.Append "诊疗项目ID", adBigInt
    MakePathItems.Fields.Append "类别", adVarChar, 10
    MakePathItems.Fields.Append "操作类型", adVarChar, 20
    MakePathItems.Fields.Append "路径项目ID", adBigInt
    MakePathItems.Fields.Append "路径项目分类", adVarChar, 50, adFldIsNullable
    MakePathItems.Fields.Append "期效", adSmallInt
     
    MakePathItems.CursorLocation = adUseClient
    MakePathItems.LockType = adLockOptimistic
    MakePathItems.CursorType = adOpenStatic
    MakePathItems.Open
End Function

Private Function MakePathAdivceRS() As ADODB.Recordset
    Set MakePathAdivceRS = New ADODB.Recordset
    
    MakePathAdivceRS.Fields.Append "行号", adBigInt
    MakePathAdivceRS.Fields.Append "路径项目ID", adBigInt
    MakePathAdivceRS.Fields.Append "原医嘱ID", adBigInt
    
    MakePathAdivceRS.Fields.Append "路径项目分类", adVarChar, 50, adFldIsNullable
    MakePathAdivceRS.Fields.Append "医嘱IDS", adLongVarWChar, 4000, adFldIsNullable
    MakePathAdivceRS.CursorLocation = adUseClient
    MakePathAdivceRS.LockType = adLockOptimistic
    MakePathAdivceRS.CursorType = adOpenStatic
    MakePathAdivceRS.Open
End Function

Private Function MakePathRichEPR() As ADODB.Recordset
    Set MakePathRichEPR = New ADODB.Recordset
    
    MakePathRichEPR.Fields.Append "ID", adBigInt
    MakePathRichEPR.Fields.Append "文件ID", adBigInt
    MakePathRichEPR.Fields.Append "路径项目ID", adBigInt
    MakePathRichEPR.Fields.Append "路径项目分类", adVarChar, 50, adFldIsNullable
    
    MakePathRichEPR.CursorLocation = adUseClient
    MakePathRichEPR.LockType = adLockOptimistic
    MakePathRichEPR.CursorType = adOpenStatic
    MakePathRichEPR.Open
End Function

Private Function Get证候IDs(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
'功能：获取该病人证候IDs，逗号分割
    Dim strSql As String, rsTmp As Recordset
    Dim str证候IDs As String
    
    strSql = "Select 证候ID From 病人诊断记录 Where 病人id = [1] And 主页id = [2] And 证候id Is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get证候IDs", lng病人ID, lng主页ID)
    Do While Not rsTmp.EOF
        str证候IDs = str证候IDs & "," & rsTmp!证候id
        rsTmp.MoveNext
    Loop
    Get证候IDs = Mid(str证候IDs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckPathInItem(ByVal str诊疗项目IDs As String, ByRef str分类 As String, udtPati As TYPE_Pati, Optional ByRef lngDay As Long, Optional ByVal lng阶段ID As Long, _
        Optional ByRef rsStepAdvice As Recordset, Optional ByVal bln中药配方 As Boolean, Optional ByVal byt期效 As Byte) As Long
'功能：检查临床路径病人，当前输入的医嘱（一组诊疗项目）是否是当前阶段的路径内项目，如果是，则返回项目ID
'      必须且仅执行一次的项目，生成时必定已产生，再添加就当成路径外项目。
'参数：
'      str诊疗项目IDs= '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
'      udtPati-病人信息
'      lngDay-当前匹配的是第几天的医嘱
'      rsStepAdvice-当天所有的医嘱
'      rsStepAdvice,如果前一阶段和当前的阶段相同，则为前一阶段和当前阶段医嘱的集合，否则为当前阶段医嘱
'      bln中药配方=中药配方单独处理，根据参数设置的允许修改的中药比例来算
'      byt期效-不同期效的同一诊疗项目，分别在同一阶段的不同路径项目时严格按照期效匹配
'返回：路径项目ID和分类名称
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim k As Long
    Dim strTmp As String
    Dim str证候IDs As String
    Dim lng改动中药味数 As Long, dbl中药味数 As Double
    Dim lng中药味数 As Long
    Dim i As Long
    Dim arrtmp As Variant
    Dim blnTmp As Boolean
    Dim bln匹配期效 As Boolean
    
    str分类 = ""
    If str诊疗项目IDs = "0" Then
    '自由录入的医嘱固定当成路径外项目
        CheckPathInItem = 0
    Else
        '给药途径可能因为收取不同费用的原因，实际使用时不是定义的给药途径，所以只判断药品相同即可
        'bln匹配期效 启用该参数期效不一致的算作路径外项目
        bln匹配期效 = CBool(zlDatabase.GetPara("匹配时期效不同算路径外项目", glngSys, P临床路径应用, "0"))
        lng中药味数 = Val(zlDatabase.GetPara("中药配方允许修改的中药味数上限", glngSys, P临床路径应用, "30"))
        
        If Not bln中药配方 Then
            strSql = "Select 分类, 路径项目id,诊疗项目ids,执行方式,期效 " & vbNewLine & _
                    "From (Select 分类, 路径项目id, 组id, f_List2str(Cast(Collect(To_Char(诊疗项目id)) As t_Strlist)) As 诊疗项目ids,执行方式,期效 " & vbNewLine & _
                    "       From (Select c.路径项目id, b.分类, d.诊疗项目id, Nvl(d.相关id, d.Id) 组id, d.序号,b.执行方式,d.期效" & vbNewLine & _
                    "              From  临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
                    "              Where b.阶段id = [2] And b.Id = c.路径项目id And c.医嘱内容id = d.Id" & vbNewLine & _
                    "                    And Not Exists(Select 1 From 诊疗项目目录 E Where D.诊疗项目ID = E.ID And E.类别 = 'E' And  E.操作类型 In('2','3','4','6'))" & vbNewLine & _
                    "                    And Not Exists(Select 1 From 诊疗项目目录 E Where D.诊疗项目ID = E.ID And E.类别 In('G','F','D') And D.相关ID<>0 )" & vbNewLine & _
                    "              Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
                    "       Group By 分类, 路径项目id, 组id,执行方式,期效)" & vbNewLine & _
                    IIf(InStr(str诊疗项目IDs, ",") > 0, "Where instr(诊疗项目ids,',')>0 ", "Where (诊疗项目ids = [1] or instr(','||诊疗项目ids||',',','||[1]||',')>0)") & IIf(bln匹配期效, " And 期效 =[3]", "")
        
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckPathInItem", str诊疗项目IDs, lng阶段ID, byt期效)
            
            '如果有多个路径项目，则只取第一个
            If rsTmp.RecordCount > 0 Then
                '如果是只能执行一次的，则判断之前是否已经匹配
                If rsTmp!执行方式 & "" = "4" Then
                    Do While Not rsTmp.EOF
                        '排除已经匹配的
                        strTmp = rsTmp!诊疗项目ids & ""
                        rsStepAdvice.Filter = "路径项目ID=" & rsTmp!路径项目ID
                        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.MoveFirst
                        For k = 0 To rsStepAdvice.RecordCount - 1
                            '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
                            If Not (rsStepAdvice!类别 & "" = "E" And InStr(",2,3,4,6,", "," & rsStepAdvice!操作类型 & ",") > 0) _
                                And Not (InStr(",G,F,D,", "," & rsStepAdvice!类别 & ",") > 0 And Val(rsStepAdvice!相关id & "") <> 0) _
                                And Val(rsStepAdvice!诊疗项目ID & "") <> 0 Then
                                strTmp = Replace("," & strTmp & ",", "," & rsStepAdvice!诊疗项目ID & ",", ",")
                                strTmp = Mid(strTmp, 2)
                                If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            End If
                            rsStepAdvice.MoveNext
                        Next
                        If InStr(str诊疗项目IDs, ",") > 0 Then
                            arrtmp = Split(str诊疗项目IDs, ",")
                            blnTmp = True
                            For i = 0 To UBound(arrtmp)
                                If InStr("," & strTmp & ",", "," & arrtmp(i) & ",") = 0 Then
                                    blnTmp = False
                                    Exit For
                                End If
                            Next
                            If blnTmp Then
                                CheckPathInItem = rsTmp!路径项目ID
                                str分类 = rsTmp!分类
                                Exit Function
                            End If
                        Else
                            If strTmp = str诊疗项目IDs Or InStr("," & strTmp & ",", "," & str诊疗项目IDs & ",") > 0 And strTmp <> "" Then
                                CheckPathInItem = rsTmp!路径项目ID
                                str分类 = rsTmp!分类
                                If Not bln匹配期效 Then
                                    '存在多个情况下根据期效进行匹配,优先匹配诊疗项目ID和期效都相同的（为了保证在同一阶段,同一药品,不同项目,期效不一致的情况下,能正常匹配）
                                    For i = 1 To rsTmp.RecordCount
                                        If rsTmp!期效 & "" = byt期效 & "" Then
                                            CheckPathInItem = rsTmp!路径项目ID
                                            str分类 = rsTmp!分类
                                            Exit For
                                        End If
                                        rsTmp.MoveNext
                                    Next
                                End If
                                Exit Do
                            End If
                        End If
                        '如果已经匹配，则继续找下一个匹配的项目
                        rsTmp.MoveNext
                    Loop
                Else
                    If InStr(str诊疗项目IDs, ",") > 0 Then
                        '多个项目判断时，如检验项目，忽略顺序，如果其中有一个是路径外的那么一组就是路径外的
                        arrtmp = Split(str诊疗项目IDs, ",")
                        Do While Not rsTmp.EOF
                            blnTmp = True
                            For i = 0 To UBound(arrtmp)
                                If InStr("," & rsTmp!诊疗项目ids & ",", "," & arrtmp(i) & ",") = 0 Then
                                    blnTmp = False
                                    Exit For
                                End If
                            Next
                            If blnTmp Then
                                CheckPathInItem = rsTmp!路径项目ID
                                str分类 = rsTmp!分类
                                Exit Function
                            End If
                            
                            rsTmp.MoveNext
                        Loop
                    Else
                        CheckPathInItem = rsTmp!路径项目ID
                        str分类 = rsTmp!分类
                        If Not bln匹配期效 Then
                            '存在多个情况下根据期效进行匹配,优先匹配诊疗项目ID和期效都相同的（为了保证在同一阶段,同一药品,不同项目,期效不一致的情况下,能正常匹配）
                            For i = 1 To rsTmp.RecordCount
                                If rsTmp!期效 & "" = byt期效 & "" Then
                                    CheckPathInItem = rsTmp!路径项目ID
                                    str分类 = rsTmp!分类
                                    Exit For
                                End If
                                rsTmp.MoveNext
                            Next
                        End If
                    End If
                End If
            End If
        Else
            '匹配时排开不符合的证候
            str证候IDs = Get证候IDs(udtPati.病人ID, udtPati.主页ID)
            strSql = "Select 分类, 路径项目id,诊疗项目ids,执行方式" & vbNewLine & _
                    "From (Select 分类, 路径项目id, 组id, f_List2str(Cast(Collect(To_Char(诊疗项目id)) As t_Strlist)) As 诊疗项目ids,执行方式" & vbNewLine & _
                    "       From (Select c.路径项目id, b.分类, d.诊疗项目id, Nvl(d.相关id, d.Id) 组id, d.序号,b.执行方式" & vbNewLine & _
                    "              From  临床路径项目 B, 临床路径医嘱 C, 路径医嘱内容 D" & vbNewLine & _
                    "              Where b.阶段id = [1] And b.Id = c.路径项目id And c.医嘱内容id = d.Id" & vbNewLine & _
                    "              And Exists(Select 1 From 诊疗项目目录 E Where d.诊疗项目id = e.Id And e.类别 = '7') " & vbNewLine & _
                    IIf(str证候IDs <> "", " And (Instr(',' || [2] || ',', ',' || d.组合项目id || ',') > 0 Or d.组合项目id Is Null)", "") & vbNewLine & _
                    "              Order By b.分类, b.项目序号, d.序号)" & vbNewLine & _
                    "       Group By 分类, 路径项目id, 组id,执行方式)" & vbNewLine
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckPathInItem", lng阶段ID, str证候IDs)
            Do While Not rsTmp.EOF
                If rsTmp!诊疗项目ids & "" <> "" Then
                    '允许改动的中药
                    dbl中药味数 = (UBound(Split(rsTmp!诊疗项目ids & "", ",")) + 1) * lng中药味数 / 100
                    lng改动中药味数 = 0
                    '先找，配方外的中药
                    For i = 0 To UBound(Split(str诊疗项目IDs, ","))
                        If InStr("," & rsTmp!诊疗项目ids & ",", "," & Split(str诊疗项目IDs, ",")(i) & ",") = 0 Then
                            lng改动中药味数 = lng改动中药味数 + 1
                        End If
                    Next
                    '再找配方中的中药，当且缺少的
                    If rsTmp!诊疗项目ids & "" <> "" Then
                        For i = 0 To UBound(Split(rsTmp!诊疗项目ids & "", ","))
                            If InStr("," & str诊疗项目IDs & ",", "," & Split(rsTmp!诊疗项目ids & "", ",")(i) & ",") = 0 Then
                                lng改动中药味数 = lng改动中药味数 + 1
                            End If
                        Next
                    End If
                    '如果在允许的范围之内，则匹配成功，否则继续匹配
                    If lng改动中药味数 <= dbl中药味数 Then
                        CheckPathInItem = rsTmp!路径项目ID
                        str分类 = rsTmp!分类
                        Exit Do
                    End If
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SetPathRows(ByRef rsAdvice As Recordset, ByRef udtPati As TYPE_Pati, Optional ByVal lngDay As Long, Optional ByRef lng阶段ID As Long, _
            Optional ByRef rsStepAdvice As Recordset) As Boolean
'功能：设置临床路径行及相关行的信息和显示
'参数：lngDay-当前匹配的是第几天的医嘱
'      rsStepAdvice,如果前一阶段和当前的阶段相同，则为前一阶段和当前阶段医嘱的集合，否则为当前阶段医嘱
'返回：lng阶段ID，当前匹配的阶段ID
    Dim k As Long, lngBegin As Long, lngEnd As Long, lngRow As Long
    Dim str诊疗项目IDs As String, lng路径项目ID As Long, str分类 As String
    Dim blnOut As Boolean
    Dim lngRecord As Long
    Dim strSql As String, rsTmp As Recordset
    Dim bln中药配方 As Boolean
    Dim byt期效 As Byte
    
    Do While Not rsAdvice.EOF
        '一组医嘱处理一次,药品单个药进行匹配 65982
        If Val(rsAdvice!相关id & "") = 0 And Not (rsAdvice!类别 & "" = "E" And rsAdvice!操作类型 & "" = "2") Or InStr(",5,6,", "," & rsAdvice!类别 & ",") > 0 Then
            '在一并给药中增加一行时，只检查和设置当前行，因为路径外项目可能和路径内项目一并给药
            lngRecord = rsAdvice.AbsolutePosition
            If InStr(",5,6,", "," & rsAdvice!类别 & ",") > 0 Then
                '药品单个药进行匹配 65982
                rsAdvice.Filter = "ID=" & rsAdvice!ID
            Else
                rsAdvice.Filter = "ID=" & rsAdvice!ID & " Or 相关ID=" & rsAdvice!ID
            End If
            '自由录入的医嘱，不在路径表上体现
            str诊疗项目IDs = ""
            bln中药配方 = False
            If Val(rsAdvice!婴儿 & "") = 0 Then
                For k = 0 To rsAdvice.RecordCount - 1
                    '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
                    If Not (rsAdvice!类别 & "" = "E" And InStr(",2,3,4,6,", "," & rsAdvice!操作类型 & ",") > 0) _
                        And Not (InStr(",G,F,D,", "," & rsAdvice!类别 & ",") > 0 And Val(rsAdvice!相关id & "") <> 0) _
                        And Val(rsAdvice!诊疗项目ID & "") <> 0 Then
                        str诊疗项目IDs = str诊疗项目IDs & "," & rsAdvice!诊疗项目ID
                        If rsAdvice!类别 & "" = "7" Then bln中药配方 = True
                        byt期效 = Val(rsAdvice!期效 & "")
                    End If
                    rsAdvice.MoveNext
                Next
                str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
                If str诊疗项目IDs <> "" Then
                    lng路径项目ID = CheckPathInItem(str诊疗项目IDs, str分类, udtPati, lngDay, lng阶段ID, rsStepAdvice, bln中药配方, byt期效)
                    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
                    For k = 0 To rsAdvice.RecordCount - 1
                        If lng路径项目ID = 0 Then
                            blnOut = True
                        Else
                            rsAdvice!路径项目ID = lng路径项目ID
                            rsAdvice!路径项目分类 = str分类
                            rsStepAdvice.Filter = "ID=" & rsAdvice!ID & " And 路径项目ID = 0"
                            rsStepAdvice!路径项目ID = lng路径项目ID
                            rsStepAdvice!路径项目分类 = str分类
                            rsStepAdvice.Update
                        End If
                        rsAdvice.Update
                        rsAdvice.MoveNext
                    Next
                    rsAdvice.Filter = 0
                    '设置记录集的位置
                    rsAdvice.AbsolutePosition = lngRecord
                    If InStr(",5,6,", "," & rsAdvice!类别 & ",") > 0 Then
                        '西药的给药方式同步修改
                        rsAdvice.Filter = "ID=" & rsAdvice!相关id
                        rsAdvice!路径项目ID = lng路径项目ID
                        rsAdvice!路径项目分类 = str分类
                        rsStepAdvice.Filter = "ID=" & rsAdvice!ID & " And 路径项目ID = 0"
                        If rsStepAdvice.RecordCount > 0 Then
                            rsStepAdvice!路径项目ID = lng路径项目ID
                            rsStepAdvice!路径项目分类 = str分类
                            rsStepAdvice.Update
                        End If
                    End If
                End If
            End If
            rsAdvice.Filter = 0
            '设置记录集的位置
            rsAdvice.AbsolutePosition = lngRecord
        End If
        
        rsAdvice.MoveNext
    Loop
    
    If blnOut Then
        If InStr(GetInsidePrivs(P临床路径应用), ";路径外项目;") = 0 Then
            MsgBox "你没有添加路径外项目的权限，不能自动补齐之前的路径项目。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    SetPathRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetPathRichEPR(ByRef rsRichEPR As Recordset, ByRef rsStepRichEPR As Recordset, ByVal lng阶段ID As Long) As Boolean
'功能：自动匹配已经生成病历
'参数：rsStepRichEPR=一个阶段的病历，rsRichEPR=当前阶段的病历
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim k As Long, strTmp As String
    Dim blnDo As Boolean
    
    On Error GoTo errH
    Do While Not rsRichEPR.EOF
        blnDo = False
        strSql = "Select 分类, 路径项目id,文件ids,执行方式 " & _
                " From (Select 分类, 路径项目id, f_List2str(Cast(Collect(To_Char(文件id)) As t_Strlist)) As 文件ids, 执行方式" & vbNewLine & _
                " From (Select b.Id As 路径项目id, a.文件id, b.分类, b.执行方式" & vbNewLine & _
                "       From 临床路径病历 A, 临床路径项目 B" & vbNewLine & _
                "       Where a.项目id = b.Id And 阶段id = [2])" & vbNewLine & _
                " Group By 分类, 路径项目id, 执行方式)" & _
                " Where 文件ids = [1] or instr(','||文件ids||',',','||[1]||',')>0"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "自动匹配病历", Val(rsRichEPR!文件ID & ""), lng阶段ID)
        Do While Not rsTmp.EOF
            If rsTmp!执行方式 & "" = "4" Then
                '排除已经匹配的
                strTmp = rsTmp!文件ids & ""
                rsStepRichEPR.Filter = "路径项目ID=" & rsTmp!路径项目ID
                If rsStepRichEPR.RecordCount > 0 Then rsStepRichEPR.MoveFirst
                For k = 0 To rsStepRichEPR.RecordCount - 1
                    strTmp = Replace("," & strTmp & ",", "," & rsStepRichEPR!文件ID & ",", ",")
                    strTmp = Mid(strTmp, 2)
                    If strTmp <> "" Then strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                    rsStepRichEPR.MoveNext
                Next
                If strTmp = rsRichEPR!文件ID & "" Or InStr("," & strTmp & ",", "," & rsRichEPR!文件ID & ",") > 0 And strTmp <> "" Then
                    blnDo = True
                    Exit Do
                End If
            Else
                blnDo = True
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
        If blnDo Then
            rsRichEPR!路径项目分类 = rsTmp!分类 & ""
            rsRichEPR!路径项目ID = Val(rsTmp!路径项目ID & "")
            rsStepRichEPR.Filter = "ID=" & rsRichEPR!ID & " And 路径项目ID=0"
            If rsStepRichEPR.RecordCount > 0 Then
                rsStepRichEPR!路径项目分类 = rsTmp!分类 & ""
                rsStepRichEPR!路径项目ID = Val(rsTmp!路径项目ID & "")
                rsStepRichEPR.Update
            End If
            rsRichEPR.Update
        End If
        rsRichEPR.MoveNext
    Loop
    SetPathRichEPR = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFirstType(ByRef udtPP As TYPE_PATH_Pati, Optional ByVal bytFunc As Byte = 0)
'功能：如果一个项目都没有，则从数据库中取第一个分类
'参数:bytFunc=0  -缺省取第一个分类,1-优先匹配含“医嘱”关键字的分类,匹配不上才取缺省的第一个分类
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strFirstName As String
    Dim strType As String
    
    On Error GoTo errH
    If bytFunc = 0 Then
        strSql = "Select 名称 from 临床路径分类 where 路径ID=[1] and 版本号=[2] and NVL(分支ID,0)=0 And 序号=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "取第一个分类", udtPP.路径ID, udtPP.版本号)
        If rsTmp.RecordCount > 0 Then GetFirstType = rsTmp!名称 & ""
    ElseIf bytFunc = 1 Then
        strSql = "Select 名称,序号 from 临床路径分类 where 路径ID=[1] and 版本号=[2] and NVL(分支ID,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "取路径分类", udtPP.路径ID, udtPP.版本号)
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!序号 & "") = 1 Then strFirstName = rsTmp!名称 & ""
            If InStr(rsTmp!名称 & "", "医嘱") > 0 Then
                strType = rsTmp!名称 & ""
                Exit For
            End If
            rsTmp.MoveNext
        Next
        GetFirstType = IIf(strType = "", strFirstName, strType)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreatePathItem(ByVal dateCur As Date, ByVal DateInPath As Date, udtPati As TYPE_Pati, udtPP As TYPE_PATH_Pati, _
    ByVal lng路径记录ID As Long, ByRef colSQL As Collection) As Boolean
'功能:根据已有医嘱或病历生成路径项目
'参数:
'入参:
    
'出参
'   colSQL: 返回可执行的SQL
'返回值:
'   T-调用成功;F调用失败
'
    Dim strSql As String
    Dim strAdivcePathOut As String
    Dim rsTmp As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rsStepAdvice As Recordset    '如果前一个阶段和当前阶段一样，则把医嘱放在一个记录集中，用于判断执行方式=4的，不重复匹配
    Dim rsStepRichEPR As Recordset
    Dim rsRichEPR As Recordset
    Dim lng前一阶段ID As Long
    Dim lng阶段ID As Long
    Dim str变异原因 As String      '变异继续的其他原因。
    Dim strFirstType As String
    Dim strAdviceType As String
    Dim i As Long, j As Long
    Dim AddDate As Date

    AddDate = dateCur
    For i = 1 To Int(dateCur) - Int(DateInPath)
        strSql = "Select a.Id,a.婴儿, a.相关id,a.诊疗项目ID, b.类别, b.操作类型,a.医嘱期效 as 期效 " & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                "Where a.诊疗项目id = b.Id And a.病人id = [1] And a.主页id = [2] And NVL(A.婴儿,0)=0" & vbNewLine & _
                "     And (a.开始执行时间 Between [3] And [4] And a.医嘱期效 = 1 Or a.医嘱期效 = 0 And [3] Between Trunc(a.开始执行时间) And Trunc(Nvl(a.执行终止时间, [4]))) And a.医嘱状态 Not In(-1,4)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "匹配路径项目", udtPati.病人ID, udtPati.主页ID, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1 / 24 / 60 / 60)
        Set rsAdvice = MakePathItems
        Do While Not rsTmp.EOF
            rsAdvice.AddNew
            rsAdvice!ID = rsTmp!ID
            rsAdvice!婴儿 = Val(rsTmp!婴儿 & "")
            rsAdvice!相关id = Val(rsTmp!相关id & "")
            rsAdvice!诊疗项目ID = Val(rsTmp!诊疗项目ID & "")
            rsAdvice!类别 = rsTmp!类别 & ""
            rsAdvice!操作类型 = rsTmp!操作类型 & ""
            rsAdvice!期效 = Val(rsTmp!期效 & "")
            rsTmp.MoveNext
        Loop
        If rsAdvice.RecordCount > 0 Then rsAdvice.Update
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        '匹配项目
        lng前一阶段ID = lng阶段ID
        lng阶段ID = 0
        strAdivcePathOut = ""
        '先取匹配阶段，默认序号最小的
        strSql = "Select ID" & vbNewLine & _
                "      From (Select ID" & vbNewLine & _
                "              From 临床路径阶段" & vbNewLine & _
                "            Where 路径id = [1] And 版本号 = [2] And [3] Between 开始天数 And Nvl(结束天数, 开始天数) And 分支id Is Null And" & vbNewLine & _
                "                  父id Is Null" & vbNewLine & _
                "            Order By 序号)" & vbNewLine & _
                "      Where Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CreatePathItem", udtPP.路径ID, udtPP.版本号, i)
        If Not rsTmp.EOF Then
            lng阶段ID = Val(rsTmp!ID & "")
        Else
            Exit For
        End If
        
        '如果阶段相同，则记录下阶段的医嘱
        If lng阶段ID <> lng前一阶段ID Then
            Set rsStepAdvice = MakePathItems
        End If
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        rsStepAdvice.Filter = 0
        Do While Not rsAdvice.EOF
            rsStepAdvice.AddNew
            rsStepAdvice!ID = rsAdvice!ID
            rsStepAdvice!婴儿 = Val(rsAdvice!婴儿 & "")
            rsStepAdvice!相关id = Val(rsAdvice!相关id & "")
            rsStepAdvice!诊疗项目ID = Val(rsAdvice!诊疗项目ID & "")
            rsStepAdvice!类别 = rsAdvice!类别 & ""
            rsStepAdvice!操作类型 = rsAdvice!操作类型 & ""
            rsStepAdvice!路径项目ID = Val(rsAdvice!路径项目ID & "")
            rsStepAdvice!路径项目分类 = rsAdvice!路径项目分类 & ""
            rsAdvice.MoveNext
        Loop
        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.Update
        If rsStepAdvice.RecordCount > 0 Then rsStepAdvice.MoveFirst
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        
        If Not SetPathRows(rsAdvice, udtPati, i, lng阶段ID, rsStepAdvice) Then Exit Function
        
        Set rsPathAdvice = MakePathAdivceRS
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        Do While Not rsAdvice.EOF
            '路径外项目(婴儿医嘱、自由录入医嘱不当成路径外项目,也不作为路径表上项目)
            If Val(rsAdvice!路径项目ID & "") = 0 Then
                If Val(rsAdvice!婴儿 & "") = 0 And Val(rsAdvice!诊疗项目ID & "") <> 0 Then
                    strAdivcePathOut = strAdivcePathOut & "," & rsAdvice!ID
                End If
            Else
                rsPathAdvice.Filter = "路径项目ID = " & rsAdvice!路径项目ID
                If rsPathAdvice.RecordCount = 0 Then
                    rsPathAdvice.AddNew
                    rsPathAdvice!路径项目ID = rsAdvice!路径项目ID & ""
                    rsPathAdvice!路径项目分类 = rsAdvice!路径项目分类 & ""
                    rsPathAdvice!医嘱IDs = rsAdvice!ID & ""
                Else
                    rsPathAdvice!医嘱IDs = rsPathAdvice!医嘱IDs & "," & rsAdvice!ID
                End If
                rsPathAdvice.Update
            End If
            rsAdvice.MoveNext
        Loop
        If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
        '临床路径数据(要放在临时表提交之后，因为有外键约束)
        rsPathAdvice.Filter = ""
        For j = 1 To rsPathAdvice.RecordCount
            strSql = "Zl_病人路径生成_Insert(1," & udtPati.病人ID & "," & udtPati.主页ID & ",Null,0," & _
                lng路径记录ID & "," & lng阶段ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                ",'" & rsPathAdvice!路径项目分类 & "'," & rsPathAdvice!路径项目ID & ",'" & rsPathAdvice!医嘱IDs & "',Null,Null" & _
                ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1)"
                rsPathAdvice.MoveNext
            colSQL.Add strSql, "C" & colSQL.count + 1
            '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
            AddDate = AddDate + 1 / 24 / 60 / 60
        Next
        '路径外项目
        If strAdivcePathOut <> "" Then
            '获得其他的变异原因编码
            If str变异原因 = "" Then
                strSql = "Select 编码 From 变异常见原因 Where (名称='其它' Or 名称='其他')  And 性质=1 And rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CreatePathItem")
                If rsTmp.RecordCount > 0 Then
                    str变异原因 = rsTmp!编码 & ""
                End If
            End If
            strAdivcePathOut = Mid(strAdivcePathOut, 2)
            If strAdviceType = "" Then strAdviceType = GetFirstType(udtPP, 1)
            strSql = "Zl_病人路径生成_Insert(1," & udtPati.病人ID & "," & udtPati.主页ID & ",Null,0," & _
                lng路径记录ID & "," & lng阶段ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                ",'" & strAdviceType & "',Null" & ",'" & strAdivcePathOut & "',Null,Null" & _
                ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'路径外项目'" & _
                ",1,Null,Null,Null,'" & str变异原因 & "' ,1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
            AddDate = AddDate + 1 / 24 / 60 / 60
        End If
        
        '匹配病历
        strSql = "select ID,文件ID from 电子病历记录 Where 病人id = [1] And 主页id = [2] And 文件ID is not Null And 创建时间 Between [3] And [4] And NVL(婴儿,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "匹配路径项目", udtPati.病人ID, udtPati.主页ID, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1, CDate(Format(DateInPath, "YYYY-MM-DD 00:00:00")) + i - 1 / 24 / 60 / 60)
        Set rsRichEPR = MakePathRichEPR
        If lng阶段ID <> lng前一阶段ID Then
            Set rsStepRichEPR = MakePathRichEPR
        End If
        Do While Not rsTmp.EOF
            rsRichEPR.AddNew
            rsRichEPR!ID = rsTmp!ID
            rsRichEPR!文件ID = rsTmp!文件ID
            rsStepRichEPR.AddNew
            rsStepRichEPR!ID = rsTmp!ID
            rsStepRichEPR!文件ID = rsTmp!文件ID
            rsTmp.MoveNext
        Loop
        If rsRichEPR.RecordCount > 0 Then rsRichEPR.MoveFirst
        If Not SetPathRichEPR(rsRichEPR, rsStepRichEPR, lng阶段ID) Then Exit Function
        If rsRichEPR.RecordCount > 0 Then
            rsRichEPR.Filter = "路径项目ID<>0"
            If rsRichEPR.RecordCount > 0 Then rsRichEPR.MoveFirst
            For j = 1 To rsRichEPR.RecordCount
                strSql = "Zl_病人路径生成_Insert(1," & udtPati.病人ID & "," & udtPati.主页ID & ",Null,0," & _
                    lng路径记录ID & "," & lng阶段ID & _
                    ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                    ",'" & rsRichEPR!路径项目分类 & "'," & rsRichEPR!路径项目ID & ",Null,Null,Null" & _
                    ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & "'',NULL,NULL,NULL,NULL,'',1," & rsRichEPR!ID & ")"
                rsRichEPR.MoveNext
                colSQL.Add strSql, "C" & colSQL.count + 1
                '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
                AddDate = AddDate + 1 / 24 / 60 / 60
            Next
        End If
        
        '如果一个项目都没有，则生成一个特殊项目
        If strAdivcePathOut = "" And rsPathAdvice.RecordCount = 0 And rsRichEPR.RecordCount = 0 Then
            If strFirstType = "" Then strFirstType = GetFirstType(udtPP, 0)
            strSql = "Zl_病人路径生成_Insert(1," & udtPati.病人ID & "," & udtPati.主页ID & ",NULL," & udtPati.科室ID & "," & _
                    lng路径记录ID & "," & lng阶段ID & _
                     ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & _
                    ",'" & strFirstType & "',Null" & ",Null,Null,Null" & _
                    ",'" & UserInfo.姓名 & "'," & "To_Date('" & Format(AddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & ",'未生成任何项目'" & _
                    ",Null,'已经执行|1" & vbTab & "已经执行',Null,Null,'',1)"
            colSQL.Add strSql, "C" & colSQL.count + 1
            '登记时间加一秒，是为了取消生成时取上一个阶段的ID。
            AddDate = AddDate + 1 / 24 / 60 / 60
        End If
        '评估
        strSql = "Zl_病人路径评估_Insert(1," & lng路径记录ID & "," & lng阶段ID & _
                ",To_Date('" & Format(CDate(DateInPath + i - 1), "yyyy-MM-dd") & "','YYYY-MM-DD')," & i & ",'" & _
                UserInfo.姓名 & "'," & IIf(strAdivcePathOut = "", "0", "1") & ",'','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','',0,Null,Null" & ",Null,1" & ")"
                
        colSQL.Add strSql, "C" & colSQL.count + 1
    Next
    
    CreatePathItem = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
