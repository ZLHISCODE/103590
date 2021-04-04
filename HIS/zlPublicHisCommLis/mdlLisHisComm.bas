Attribute VB_Name = "mdlLisHisComm"
Option Explicit
Public gblnInit As Boolean                                         '公共部件是否已初始化

Public grsParas As ADODB.Recordset                                  '系统参数表缓存
Public grsUserParas As ADODB.Recordset                              '系统参数表缓存
Public gcolPrivs As Collection                                      '当前用户具备的所有程序的功能权限
Public gblnAllSite As Boolean                                       '是否能够查看所有站点
  
Public gobjHisComLib As Object
Public gobjHisDatabase As Object
Public gobjHisSystem As Object
Public gobjPlugIn As Object                                         '外挂部件

Public gcnLisOracle As New ADODB.Connection                         'LIS公共数据库连接
Public gcnHisOracle As New ADODB.Connection                         'HIS公共数据库连接

Public gstrDBUser As String

Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"

Public Type TYPE_SYS_INFO   '-----------应用程序信息 及 注册信息
    AppName As String       '系统名称 (产品简称+软件，如中联软件，医业软件)
    ShortName As String     '产品简名
    AppTitle As String      '系统标题，产品全称
    
    Version As String       '系统版本
    AviPath As String       'AVI文件路径
    
    UnitName  As String     '用户单位名称
    Supporter As String     '技术支持商
    Develop As String       '开发商
    SupporterWEB As String  '支持商WEB简名
    SupporterMail As String '支持商邮件
    SupporterURL As String  '支持商网址
    ProductLine  As String  '产品系列，[标准版],[大客户版]
    
    SysNo       As Long     '系统编号
    ModlNo      As Long     '模块号
    VersionHIS As String       'HIS系统版本
    VersionLIS As String       'lis系统版本
End Type

Public Type TYPE_SYS_PARAMETER    '系统参数
    Privs        As String  '模块权限
    MachineCount As Integer '仪器数量
    blnEmerge    As Boolean '是否区分急诊
    BuffDir      As String   '本地缓存记录集的缓存目录
    InvaidWord   As String   '需去掉的非常字符
    intCA        As Integer  'CA中心编号
    strMatch     As String   '输入匹配
End Type

Public Type TYPE_USER_INFO
    ID As Long          '人员ID
    DeptID As Long      '人员对应的部门ID
    DeptName As String  '人员对应的部门名称
    No As String        '人员编号
    Name As String      '人员姓名
    Code As String      '人员简码
    DBUser As String    '人员对应的数据库用户名
    ComputerName As String          '电脑名
    NodeNo As String                '人员登陆站点
End Type

Public gSampleShowColour As SampleValShowColour                    '结果显示颜色
Public Type SampleValShowColour                                    '结果颜色显示
    正常 As Double
    偏低 As Double
    偏高 As Double
    异常 As Double
    警示偏高 As Double
    警示偏低 As Double
    复查偏高 As Double
    复查偏低 As Double
End Type


Public gUserInfo As TYPE_USER_INFO
Public gSysInfo As TYPE_SYS_INFO
Public gSysParameter As TYPE_SYS_PARAMETER
Public gstrSQL  As String

Private Const p医嘱附费管理 As Integer = 1257                       '病人费用模块授权
Private Const p门诊医嘱下达 As Integer = 1252                       '门诊医嘱下达
Private Const p住院医嘱下达 As Integer = 1253                       '住院医嘱下达
Private Const p门诊病历管理 As Integer = 1250                       '门诊病历
Private Const p住院病历管理 As Integer = 1251



Public Declare Function SetParent Lib "user32.dll " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const CB_ADDSTRING = &H143
Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private mstrPara As String

Public Function ComInitComLib(ByRef strErr As String) As Boolean
'初始化公共部件,在程序启动时调用
    Dim strDBUser As String
    On Error GoTo errH
    ComInitComLib = False



    If gblnInit Then
        ComInitComLib = True
        Exit Function
    End If

    If gcnHisOracle.State = 1 Then
        Set gobjHisComLib = CreateObject("zl9ComLib.clsComLib")
        gobjHisComLib.InitCommon gcnHisOracle
        Set gobjHisDatabase = gobjHisComLib.zlDatabase
        If VerCompare(gSysInfo.VersionHIS, "10.35.10") = -1 Then
            strDBUser = GetUserDB(2)
            gobjHisComLib.SetDbUser strDBUser
        End If
        Set gobjHisSystem = CreateObject("zl9ComLib.clsSystem")
    End If


    ComInitComLib = True
    gblnInit = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
End Function

'调用 公共部件ComLib的一些公共函数过程
Public Function ComOpenSQL(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    '功能：通过ComLib对象打开带参数SQL的记录集
    '
    Dim lngCount As Long
    Dim var(30) As Variant
    

    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "不支持超过30个参数的SQL！"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    
    Set ComOpenSQL = OpenSQLRecord(selDB, strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))



End Function

Public Function ComExecuteProc(ByVal selDB As Integer, strSQL As String, ByVal strFormCaption As String) As String
    '功能：执行过程语句,并自动对过程参数进行绑定变量处理
    '返回：无错误返回空串，否则返回错误提示
    

    Call ExecuteProcedure(selDB, strSQL, strFormCaption)
    
End Function

Public Function BeforCreateLisValueStr(ByVal strAdvices As String, Optional ByVal DateE As Date, Optional strErr As String) As Boolean
          '传入医嘱ID和时间,判断传入时间之后是否存在医嘱ID对应的已审核记录
          'stradvices     '医嘱ID,多个医嘱使用","号分隔
          'DateE          '获取在该时间之后是否存在医嘱ID对应的记录
          
          '返回       True=存在记录,False=不存在记录
          
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo BeforCreateLisValueStr_Error

2         BeforCreateLisValueStr = False
3         strErr = ""
          
4         strSQL = "Select /*+cardinality(c,10)*/ B.ID From 检验申请组合 A, 检验报告记录 B,Table(f_Str2list([1])) C" & _
                   " Where a.标本id = b.Id And a.申请id=c.Column_Value and b.审核时间 > [2]"
5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检查是否有数据", strAdvices, DateE)
          
          '有记录时返回True
6         Do While rsTmp.EOF = False
7             If IsNull(rsTmp("ID")) = False Then
8                 BeforCreateLisValueStr = True
9                 Exit Function
10            End If
11            rsTmp.MoveNext
12        Loop


13        Exit Function
BeforCreateLisValueStr_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(BeforCreateLisValueStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear
              
End Function

Public Function CreateLisValueStr(strAdvices As String, Optional lngPatient As Long, Optional strErr As String, Optional intType As Integer) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                  根据传入医嘱ID返回结果
          '参数
          '           strAdvices          申请ID串,用逗号分隔
          '           lngPatient          可选的参数，输入病人ID后只按病人ID查找结果
          '           strType               0-审核，1-取消审核
          '标本组成格式
          '               类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2> 婴儿序号 <split2>
          '                   指标1<split4>检验结果1<split4>单位1<split4>结果标志1<split4>结果参数1<split4>排列序号1<split4>隐私项目1<split4>指标代码1<split4>中文名1<split4>英文名1<split4>参考高值1<split4>参考底值1<split4>小数位数1<split3>
          '                   指标2<split4>检验结果2<split4>单位2<split4>结果标志2<split4>结果参数2<split4>排列序号2<split4>隐私项目2<split4>指标代码2<split4>中文名2<split4>英文名2<split4>参考高值2<split4>参考底值2<split4>小数位数2<split3>
          '                   指标3<split4>检验结果3<split4>单位3<split4>结果标志3<split4>结果参数3<split4>排列序号3<split4>隐私项目3<split4>指标代码3<split4>中文名3<split4>英文名3<split4>参考高值3<split4>参考底值3<split4>小数位数3<split1>
          '
          '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>
          '               细菌名1<split3>描述1<split3>耐药机制1<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
          '               细菌名2<split3>描述2<split3>耐药机制2<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngID As Long
          Dim lngSampleID As Long
          Dim lngSampleGroup As Long
          Dim lngMicroID As Long
          Dim strSampleOne As String
          Dim strSampleTwo As String
          Dim intCount As Integer
          Dim strTemp As String
          Dim stridSQL As String
          Dim i As Long
          Dim str参考高值 As String
          Dim str参考低值 As String
          Dim str结果参考 As String
          Dim str检验结果 As String
          
          '分隔的常量
          Const conSplit1 As String = "<split1>"                        '用于分隔标本,使用“<split1>”分隔，以前使用“|”
          Const conSplit2 As String = "<split2>"                        '用于分隔标本信息,使用“<split2>”分隔，以前使用“;”
          Const conSplit3 As String = "<split3>"                        '用于分隔标本指标信息,使用“<split3>”分隔，以前使用“,”
          Const conSplit4 As String = "<split4>"                        '用于分隔指标内信息,使用“<split4>”分隔，以前使用“^”
          
          
          
          '分别读出普通和微生物项目
          
          '对strAdvices参数字符串长度超过4000做处理
1         On Error GoTo CreateLisValueStr_Error

2         If Len(strAdvices) > 4000 Then
3             For i = 0 To UBound(Split(strAdvices, ","))
4                 strTemp = strTemp & "," & Split(strAdvices, ",")(i)
5                 intCount = intCount + 1
                  
6                 If intCount = 200 Then
7                     stridSQL = stridSQL & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
8                     intCount = 0
9                     strTemp = ""
10                End If
11            Next
12            If strTemp <> "" Then
13                stridSQL = Mid(stridSQL, 12) & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
14            End If
15        Else
16            stridSQL = "Select Column_Value From Table(Cast(F_Num2list([3]) As Zltools.T_Numlist))"
17        End If
          
          '普通
18        strSQL = "Select distinct [关键字] 1 Type,A.申请id,a.病人来源 ,b.报告时间,b.检验人,b.审核人,b.审核时间,e.名称 检验项目名称, " & vbNewLine & _
                  "       D.中文名 || '(' || D.英文名 || ')' 指标,c.检验结果, D.单位," & vbNewLine & _
                  "       Decode(C.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑', '') 结果标志," & vbNewLine & _
                  "       c.结果参考,C.排列序号, 0 隐私项目,b.标本类型,a.婴儿 婴儿序号,d.指标代码,a.标本id,a.组合id,d.中文名,d.英文名,c.参考高值,c.参考低值,nvl(d.小数位数,2) 小数位数,b.病人来源,d.结果类型 " & vbNewLine & _
                  "From 检验申请组合 A, 检验报告记录 B, 检验报告明细 C, 检验指标 D,检验组合项目 e [表格]" & vbNewLine & _
                  "Where A.标本id = B.Id And B.Id = C.标本id And C.项目id = D.Id And Nvl(B.微生物, 0) <> 1 And" & vbNewLine & _
                  "      a.组合id = c.组合id and a.组合id = e.id  [条件1]" & vbNewLine & _
                  "     [条件] " & vbNewLine & _
                  " order by a.申请id,a.标本ID,a.组合id,c.排列序号 "
19        If intType = 1 Then
20           strSQL = Replace(strSQL, "[条件1]", "")
21        Else
22            strSQL = Replace(strSQL, "[条件1]", " and   b.审核人 is not null")
23        End If
24        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) ")
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
25            strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(f,10)*/")
26            strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") f")
27            strSQL = Replace(strSQL, "[条件]", " and a.申请id=f.Column_Value")
28        Else
29            strSQL = Replace(strSQL, "[关键字]", "")
30            strSQL = Replace(strSQL, "[表格]", "")
31            strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
32        End If

33        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
34        lngID = 0
35        Do Until rsTmp.EOF
              '数据格式检查
36            str参考高值 = IIf(IsNumeric(rsTmp("参考高值") & ""), IIf(Mid(rsTmp("参考高值") & "", 1, 1) = ".", "0" & rsTmp("参考高值"), rsTmp("参考高值") & ""), rsTmp("参考高值") & "")
37            str参考低值 = IIf(IsNumeric(rsTmp("参考低值") & ""), IIf(Mid(rsTmp("参考低值") & "", 1, 1) = ".", "0" & rsTmp("参考低值"), rsTmp("参考低值") & ""), rsTmp("参考低值") & "")
              If IsNumeric(rsTmp("检验结果") & "") Then
38               str检验结果 = IIf(Val(rsTmp("结果类型") & "") = 1, Format(rsTmp("检验结果") & "", IIf(Val(rsTmp("小数位数") & "") > 0, "0." & String(Val(rsTmp("小数位数") & ""), "0"), "0")), rsTmp("检验结果") & "")
              Else
                  str检验结果 = rsTmp("检验结果") & ""
              End If
39            If InStr(rsTmp("结果参考") & "", "--") > 0 Then
40                str结果参考 = IIf(IsNumeric(Split(rsTmp("结果参考") & "", "--")(0)), IIf(Mid(Split(rsTmp("结果参考") & "", "--")(0), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "--")(0), Split(rsTmp("结果参考") & "", "--")(0)), Split(rsTmp("结果参考") & "", "--")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("结果参考") & "", "--")(1)), IIf(Mid(Split(rsTmp("结果参考") & "", "--")(1), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "--")(1), Split(rsTmp("结果参考") & "", "--")(1)), Split(rsTmp("结果参考") & "", "--")(1))
41            ElseIf InStr(rsTmp("结果参考") & "", "-") > 0 Then
42                str结果参考 = IIf(IsNumeric(Split(rsTmp("结果参考") & "", "-")(0)), IIf(Mid(Split(rsTmp("结果参考") & "", "-")(0), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "-")(0), Split(rsTmp("结果参考") & "", "-")(0)), Split(rsTmp("结果参考") & "", "-")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("结果参考") & "", "-")(1)), IIf(Mid(Split(rsTmp("结果参考") & "", "-")(1), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "-")(1), Split(rsTmp("结果参考") & "", "-")(1)), Split(rsTmp("结果参考") & "", "-")(1))
43            Else
44                str结果参考 = rsTmp("结果参考") & ""
45            End If
              
46            If lngID <> NVL(rsTmp("申请ID"), 0) Or lngSampleID <> NVL(rsTmp("标本ID"), 0) Or lngSampleGroup <> NVL(rsTmp("组合id"), 0) Then
47                strSampleOne = strSampleOne & conSplit1 & "1" & conSplit2 & rsTmp("申请ID") & conSplit2 & rsTmp("病人来源") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("报告时间")), "", rsTmp("报告时间"))) & conSplit2 & _
                              rsTmp("检验人") & conSplit2 & rsTmp("审核人") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("审核时间")), "", rsTmp("审核时间"))) & conSplit2 & rsTmp("检验项目名称") & conSplit2 & _
                              rsTmp("标本类型") & conSplit2 & NVL(rsTmp("婴儿序号"), "0") & conSplit2 & rsTmp("指标") & conSplit4 & str检验结果 & conSplit4 & rsTmp("单位") & _
                              conSplit4 & rsTmp("结果标志") & conSplit4 & str结果参考 & conSplit4 & rsTmp("排列序号") & conSplit4 & rsTmp("隐私项目") & conSplit4 & rsTmp("指标代码") & _
                              conSplit4 & rsTmp("中文名") & conSplit4 & rsTmp("英文名") & conSplit4 & str参考高值 & conSplit4 & str参考低值 & IIf(rsTmp("病人来源") & "" = "4", conSplit4 & rsTmp("小数位数"), "")
48            Else
49                strSampleOne = strSampleOne & conSplit3 & rsTmp("指标") & conSplit4 & str检验结果 & conSplit4 & rsTmp("单位") & _
                              conSplit4 & rsTmp("结果标志") & conSplit4 & str结果参考 & conSplit4 & rsTmp("排列序号") & conSplit4 & rsTmp("隐私项目") & conSplit4 & rsTmp("指标代码") & _
                              conSplit4 & rsTmp("中文名") & conSplit4 & rsTmp("英文名") & conSplit4 & str参考高值 & conSplit4 & str参考低值 & IIf(rsTmp("病人来源") & "" = "4", conSplit4 & rsTmp("小数位数"), "")
50            End If
51            lngID = NVL(rsTmp("申请ID"), 0)
52            lngSampleID = NVL(rsTmp("标本ID"), 0)
53            lngSampleGroup = NVL(rsTmp("组合id"), 0)
54            rsTmp.MoveNext
55        Loop
          
          
56        lngID = 0
57        lngMicroID = 0
58        strSQL = "Select distinct [关键字] 2 Type,A.申请id,a.病人来源 ,b.报告时间,b.检验人,b.审核人,b.审核时间,g.名称 检验项目名称, " & vbNewLine & _
                  "       E.中文名 || '(' || E.英文名 || ')' 细菌, C.培养描述, C.耐药机制, F.中文名 || '(' || F.英文名 || ')' 抗生素, D.结果 抗生素结果," & vbNewLine & _
                  "       D.结果类型, D.药敏方法, F.用法用量1, F.用法用量2, 血药浓度1, 血药浓度2, 尿药浓度1, 尿药浓度2,c.细菌ID,b.标本类型,a.婴儿 婴儿序号" & vbNewLine & _
                  "From 检验申请组合 A, 检验报告记录 B, 检验报告细菌 C, 检验报告药敏 D, 检验细菌记录 E, 检验药敏 F,检验组合项目 g [表格]" & vbNewLine & _
                  "Where A.标本id = B.Id And B.微生物 = 1 And B.Id = C.标本id And C.Id = D.结果id And C.细菌id = E.Id And D.药敏id = F.Id and a.组合id = g.id " & vbNewLine & _
                  "      [条件] " & vbNewLine & _
                  " order by a.申请id,c.细菌id"
                  
59        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
60            strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(h,10)*/")
61            strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") h")
62            strSQL = Replace(strSQL, "[条件]", " and a.申请id=h.Column_Value")
63        Else
64            strSQL = Replace(strSQL, "[关键字]", "")
65            strSQL = Replace(strSQL, "[表格]", "")
66            strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
67        End If
68        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
          
          
          '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>婴儿序号<split2>
          '               细菌名1<split3>描述1<split3>耐药机制1<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
          '               细菌名2<split3>描述2<split3>耐药机制2<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
          
          
69        If rsTmp.RecordCount <= 0 Then
70            strSQL = "Select  distinct [关键字] 2 Type, a.申请id, a.病人来源, b.报告时间, b.检验人, b.审核人, b.审核时间, c.正常菌 检验项目名称, c.未检出 细菌, c.培养描述, c.耐药机制, '' 抗生素, '' 抗生素结果, '' 结果类型," & vbNewLine & _
                       "          '' 药敏方法, '' 用法用量1, '' 用法用量2, '' 血药浓度1, '' 血药浓度2, '' 尿药浓度1, '' 尿药浓度2, c.细菌id, b.标本类型, a.婴儿 婴儿序号" & vbNewLine & _
                       "   From 检验申请组合 A, 检验报告记录 B, 检验报告细菌 C [表格]" & vbNewLine & _
                       "   Where a.标本id = b.Id And b.微生物 = 1 And b.Id = c.标本id " & vbNewLine & _
                       "      [条件] " & vbNewLine & _
                       "   Order By a.申请id, c.细菌id"
                      
71            If lngPatient = 0 Then
      '            strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
72                strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(d,10)*/")
73                strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") D")
74                strSQL = Replace(strSQL, "[条件]", " and a.申请id=d.Column_Value")
75            Else
76                strSQL = Replace(strSQL, "[关键字]", "")
77                strSQL = Replace(strSQL, "[表格]", "")
78                strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
79            End If
80            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
81        End If
          
82        Do Until rsTmp.EOF
83            If lngID <> NVL(rsTmp("申请ID"), 0) Then
84                strSampleTwo = strSampleTwo & conSplit1 & "2" & conSplit2 & IIf(IsNull(rsTmp("申请ID")), "", rsTmp("申请ID")) & conSplit2 & IIf(IsNull(rsTmp("病人来源")), "", rsTmp("病人来源")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("报告时间")), "", rsTmp("报告时间"))) & conSplit2 & _
                              IIf(IsNull(rsTmp("检验人")), "", rsTmp("检验人")) & conSplit2 & IIf(IsNull(rsTmp("审核人")), "", rsTmp("审核人")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("审核时间")), "", rsTmp("审核时间"))) & conSplit2 & IIf(IsNull(rsTmp("检验项目名称")), "", rsTmp("检验项目名称")) & conSplit2 & _
                              IIf(IsNull(rsTmp("标本类型")), "", rsTmp("标本类型")) & conSplit2 & NVL(rsTmp("婴儿序号"), "0") & conSplit2 & IIf(IsNull(rsTmp("细菌")), "", rsTmp("细菌")) & conSplit3 & IIf(IsNull(rsTmp("培养描述")), "", rsTmp("培养描述")) & conSplit3 & IIf(IsNull(rsTmp("耐药机制")), "", rsTmp("耐药机制")) & _
                              conSplit3 & IIf(IsNull(rsTmp("抗生素")), "", rsTmp("抗生素")) & conSplit4 & IIf(IsNull(rsTmp("抗生素结果")), "", rsTmp("抗生素结果")) & conSplit4 & IIf(IsNull(rsTmp("结果类型")), "", rsTmp("结果类型")) & conSplit4 & IIf(IsNull(rsTmp("药敏方法")), "", rsTmp("药敏方法")) & conSplit4 & IIf(IsNull(rsTmp("用法用量1")), "", rsTmp("用法用量1")) & _
                              conSplit4 & IIf(IsNull(rsTmp("用法用量2")), "", rsTmp("用法用量2")) & conSplit4 & IIf(IsNull(rsTmp("血药浓度1")), "", rsTmp("血药浓度1")) & conSplit4 & IIf(IsNull(rsTmp("血药浓度2")), "", rsTmp("血药浓度2")) & conSplit4 & IIf(IsNull(rsTmp("尿药浓度1")), "", rsTmp("尿药浓度1")) & conSplit4 & IIf(IsNull(rsTmp("尿药浓度2")), "", rsTmp("尿药浓度2"))
85                lngMicroID = NVL(rsTmp("细菌ID"), 0)
86            Else
87                If lngMicroID <> NVL(rsTmp("细菌ID"), 0) Then
88                    strSampleTwo = strSampleTwo & conSplit2 & rsTmp("细菌") & conSplit3 & rsTmp("培养描述") & conSplit3 & rsTmp("耐药机制") & _
                              conSplit3 & rsTmp("抗生素") & conSplit4 & rsTmp("抗生素结果") & conSplit4 & rsTmp("结果类型") & conSplit4 & rsTmp("药敏方法") & conSplit4 & rsTmp("用法用量1") & _
                              conSplit4 & rsTmp("用法用量2") & conSplit4 & rsTmp("血药浓度1") & conSplit4 & rsTmp("血药浓度2") & conSplit4 & rsTmp("尿药浓度1") & conSplit4 & rsTmp("尿药浓度2")
89                Else
90                    strSampleTwo = strSampleTwo & conSplit3 & rsTmp("抗生素") & conSplit4 & rsTmp("抗生素结果") & conSplit4 & rsTmp("结果类型") & conSplit4 & rsTmp("药敏方法") & _
                              conSplit4 & rsTmp("用法用量1") & conSplit4 & rsTmp("用法用量2") & conSplit4 & rsTmp("血药浓度1") & conSplit4 & rsTmp("血药浓度2") & _
                              conSplit4 & rsTmp("尿药浓度1") & conSplit4 & rsTmp("尿药浓度2")
91                End If
92            End If
93            lngID = NVL(rsTmp("申请ID"), 0)
94            lngMicroID = NVL(rsTmp("细菌ID"), 0)
95            rsTmp.MoveNext
96        Loop

97        If strSampleOne <> "" Then
98            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
99        End If
100       If strSampleTwo <> "" Then
101           strSampleTwo = Mid(strSampleTwo, Len(conSplit1) + 1)
102       End If
103       If strSampleTwo <> "" Then
104           If strSampleOne = "" Then
105               CreateLisValueStr = strSampleTwo
106           Else
107               CreateLisValueStr = strSampleOne & conSplit1 & strSampleTwo
108           End If
109       Else
110           CreateLisValueStr = strSampleOne
111       End If


112       Exit Function
CreateLisValueStr_Error:
113       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(CreateLisValueStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
114       Err.Clear
          
End Function

Public Function CreateLisValueStrForTJ(strAdvices As String, Optional lngPatient As Long, Optional strErr As String, Optional intType As Integer) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                  根据传入医嘱ID返回结果
          '参数
          '           strAdvices          申请ID串,用逗号分隔
          '           lngPatient          可选的参数，输入病人ID后只按病人ID查找结果
          '           strType               0-审核，1-取消审核
          '标本组成格式
          '               类型(1=普通)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2> 婴儿序号 <split2>
          '                   指标1<split4>检验结果1<split4>单位1<split4>结果标志1<split4>结果参数1<split4>排列序号1<split4>隐私项目1<split4>指标代码1<split4>中文名1<split4>英文名1<split4>参考高值1<split4>参考底值1<split4>小数位数1<split3>
          '                   指标2<split4>检验结果2<split4>单位2<split4>结果标志2<split4>结果参数2<split4>排列序号2<split4>隐私项目2<split4>指标代码2<split4>中文名2<split4>英文名2<split4>参考高值2<split4>参考底值2<split4>小数位数2<split3>
          '                   指标3<split4>检验结果3<split4>单位3<split4>结果标志3<split4>结果参数3<split4>排列序号3<split4>隐私项目3<split4>指标代码3<split4>中文名3<split4>英文名3<split4>参考高值3<split4>参考底值3<split4>小数位数3<split1>
          '
          '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>
          '               细菌名1<split3>描述1<split3>耐药机制1<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
          '               细菌名2<split3>描述2<split3>耐药机制2<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngID As Long
          Dim lngSampleID As Long
          Dim lngSampleGroup As Long
          Dim lngMicroID As Long
          Dim strSampleOne As String
          Dim strSampleTwo As String
          Dim intCount As Integer
          Dim strTemp As String
          Dim stridSQL As String
          Dim i As Long
          Dim str参考高值 As String
          Dim str参考低值 As String
          Dim str结果参考 As String
          Dim str检验结果 As String
          
          '分隔的常量
          Const conSplit1 As String = "<split1>"                        '用于分隔标本,使用“<split1>”分隔，以前使用“|”
          Const conSplit2 As String = "<split2>"                        '用于分隔标本信息,使用“<split2>”分隔，以前使用“;”
          Const conSplit3 As String = "<split3>"                        '用于分隔标本指标信息,使用“<split3>”分隔，以前使用“,”
          Const conSplit4 As String = "<split4>"                        '用于分隔指标内信息,使用“<split4>”分隔，以前使用“^”
          
          
          
          '分别读出普通和微生物项目
          
          '对strAdvices参数字符串长度超过4000做处理
1         On Error GoTo CreateLisValueStrForTJ_Error

2         If Len(strAdvices) > 4000 Then
3             For i = 0 To UBound(Split(strAdvices, ","))
4                 strTemp = strTemp & "," & Split(strAdvices, ",")(i)
5                 intCount = intCount + 1
                  
6                 If intCount = 200 Then
7                     stridSQL = stridSQL & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
8                     intCount = 0
9                     strTemp = ""
10                End If
11            Next
12            If strTemp <> "" Then
13                stridSQL = Mid(stridSQL, 12) & " Union All " & "Select Column_Value From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))"
14            End If
15        Else
16            stridSQL = "Select Column_Value From Table(Cast(F_Num2list([3]) As Zltools.T_Numlist))"
17        End If
          
          '普通
18        strSQL = "Select distinct [关键字] 1 Type,A.申请id,a.病人来源 ,b.报告时间,b.检验人,b.审核人,b.审核时间,e.名称 检验项目名称, " & vbNewLine & _
                  "       D.中文名 || '(' || D.英文名 || ')' 指标,c.检验结果, D.单位," & vbNewLine & _
                  "       Decode(C.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑',7,'复查下限',8,'复查上限', '') 结果标志," & vbNewLine & _
                  "       c.结果参考,C.排列序号, 0 隐私项目,b.标本类型,a.婴儿 婴儿序号,d.指标代码,a.标本id,a.组合id,d.中文名,d.英文名,c.参考高值,c.参考低值,nvl(d.小数位数,2) 小数位数,b.病人来源,d.结果类型 " & vbNewLine & _
                  "From 检验申请组合 A, 检验报告记录 B, 检验报告明细 C, 检验指标 D,检验组合项目 e [表格]" & vbNewLine & _
                  "Where A.标本id = B.Id And B.Id = C.标本id And C.项目id = D.Id And Nvl(B.微生物, 0) <> 1 And" & vbNewLine & _
                  "      a.组合id = c.组合id and a.组合id = e.id  [条件1]" & vbNewLine & _
                  "     [条件] " & vbNewLine & _
                  " order by a.申请id,a.标本ID,a.组合id,c.排列序号 "
19        If intType = 1 Then
20           strSQL = Replace(strSQL, "[条件1]", "")
21        Else
22            strSQL = Replace(strSQL, "[条件1]", " and   b.审核人 is not null")
23        End If
24        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) ")
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
25            strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(f,10)*/")
26            strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") f")
27            strSQL = Replace(strSQL, "[条件]", " and a.申请id=f.Column_Value")
28        Else
29            strSQL = Replace(strSQL, "[关键字]", "")
30            strSQL = Replace(strSQL, "[表格]", "")
31            strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
32        End If

33        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
34        lngID = 0
35        Do Until rsTmp.EOF
              '数据格式检查
36            str参考高值 = IIf(IsNumeric(rsTmp("参考高值") & ""), IIf(Mid(rsTmp("参考高值") & "", 1, 1) = ".", "0" & rsTmp("参考高值"), rsTmp("参考高值") & ""), rsTmp("参考高值") & "")
37            str参考低值 = IIf(IsNumeric(rsTmp("参考低值") & ""), IIf(Mid(rsTmp("参考低值") & "", 1, 1) = ".", "0" & rsTmp("参考低值"), rsTmp("参考低值") & ""), rsTmp("参考低值") & "")
              If IsNumeric(rsTmp("检验结果") & "") Then
38               str检验结果 = IIf(Val(rsTmp("结果类型") & "") = 1, Format(rsTmp("检验结果") & "", IIf(Val(rsTmp("小数位数") & "") > 0, "0." & String(rsTmp("小数位数") & "", "0"), "0")), rsTmp("检验结果") & "")
              Else
                  str检验结果 = rsTmp("检验结果") & ""
              End If
39            If InStr(rsTmp("结果参考") & "", "--") > 0 Then
40                str结果参考 = IIf(IsNumeric(Split(rsTmp("结果参考") & "", "--")(0)), IIf(Mid(Split(rsTmp("结果参考") & "", "--")(0), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "--")(0), Split(rsTmp("结果参考") & "", "--")(0)), Split(rsTmp("结果参考") & "", "--")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("结果参考") & "", "--")(1)), IIf(Mid(Split(rsTmp("结果参考") & "", "--")(1), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "--")(1), Split(rsTmp("结果参考") & "", "--")(1)), Split(rsTmp("结果参考") & "", "--")(1))
41            ElseIf InStr(rsTmp("结果参考") & "", "-") > 0 Then
42                str结果参考 = IIf(IsNumeric(Split(rsTmp("结果参考") & "", "-")(0)), IIf(Mid(Split(rsTmp("结果参考") & "", "-")(0), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "-")(0), Split(rsTmp("结果参考") & "", "-")(0)), Split(rsTmp("结果参考") & "", "-")(0)) & _
                      "--" & IIf(IsNumeric(Split(rsTmp("结果参考") & "", "-")(1)), IIf(Mid(Split(rsTmp("结果参考") & "", "-")(1), 1, 1) = ".", "0" & Split(rsTmp("结果参考") & "", "-")(1), Split(rsTmp("结果参考") & "", "-")(1)), Split(rsTmp("结果参考") & "", "-")(1))
43            Else
44                str结果参考 = rsTmp("结果参考") & ""
45            End If
              
46            If lngID <> NVL(rsTmp("申请ID"), 0) Or lngSampleID <> NVL(rsTmp("标本ID"), 0) Or lngSampleGroup <> NVL(rsTmp("组合id"), 0) Then
47                strSampleOne = strSampleOne & conSplit1 & "1" & conSplit2 & rsTmp("申请ID") & conSplit2 & rsTmp("病人来源") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("报告时间")), "", rsTmp("报告时间"))) & conSplit2 & _
                              rsTmp("检验人") & conSplit2 & rsTmp("审核人") & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("审核时间")), "", rsTmp("审核时间"))) & conSplit2 & rsTmp("检验项目名称") & conSplit2 & _
                              rsTmp("标本类型") & conSplit2 & NVL(rsTmp("婴儿序号"), "0") & conSplit2 & rsTmp("指标") & conSplit4 & str检验结果 & conSplit4 & rsTmp("单位") & _
                              conSplit4 & rsTmp("结果标志") & conSplit4 & str结果参考 & conSplit4 & rsTmp("排列序号") & conSplit4 & rsTmp("隐私项目") & conSplit4 & rsTmp("指标代码") & _
                              conSplit4 & rsTmp("中文名") & conSplit4 & rsTmp("英文名") & conSplit4 & str参考高值 & conSplit4 & str参考低值 & IIf(rsTmp("病人来源") & "" = "4", conSplit4 & rsTmp("小数位数"), "")
48            Else
49                strSampleOne = strSampleOne & conSplit3 & rsTmp("指标") & conSplit4 & str检验结果 & conSplit4 & rsTmp("单位") & _
                              conSplit4 & rsTmp("结果标志") & conSplit4 & str结果参考 & conSplit4 & rsTmp("排列序号") & conSplit4 & rsTmp("隐私项目") & conSplit4 & rsTmp("指标代码") & _
                              conSplit4 & rsTmp("中文名") & conSplit4 & rsTmp("英文名") & conSplit4 & str参考高值 & conSplit4 & str参考低值 & IIf(rsTmp("病人来源") & "" = "4", conSplit4 & rsTmp("小数位数"), "")
50            End If
51            lngID = NVL(rsTmp("申请ID"), 0)
52            lngSampleID = NVL(rsTmp("标本ID"), 0)
53            lngSampleGroup = NVL(rsTmp("组合id"), 0)
54            rsTmp.MoveNext
55        Loop
          
          
56        lngID = 0
57        lngMicroID = 0
58        strSQL = "Select distinct [关键字] 2 Type,A.申请id,a.病人来源 ,b.报告时间,b.检验人,b.审核人,b.审核时间,g.名称 检验项目名称, " & vbNewLine & _
                  "       E.中文名 || '(' || E.英文名 || ')' 细菌, C.培养描述, C.耐药机制, F.中文名 || '(' || F.英文名 || ')' 抗生素, D.结果 抗生素结果," & vbNewLine & _
                  "       D.结果类型, D.药敏方法, F.用法用量1, F.用法用量2, 血药浓度1, 血药浓度2, 尿药浓度1, 尿药浓度2,c.细菌ID,b.标本类型,a.婴儿 婴儿序号" & vbNewLine & _
                  "From 检验申请组合 A, 检验报告记录 B, 检验报告细菌 C, 检验报告药敏 D, 检验细菌记录 E, 检验药敏 F,检验组合项目 g [表格]" & vbNewLine & _
                  "Where A.标本id = B.Id And B.微生物 = 1 And B.Id = C.标本id And C.Id = D.结果id And C.细菌id = E.Id And D.药敏id = F.Id and a.组合id = g.id " & vbNewLine & _
                  "      [条件] " & vbNewLine & _
                  " order by a.申请id,c.细菌id"
                  
59        If lngPatient = 0 Then
      '        strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
60            strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(h,10)*/")
61            strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") h")
62            strSQL = Replace(strSQL, "[条件]", " and a.申请id=h.Column_Value")
63        Else
64            strSQL = Replace(strSQL, "[关键字]", "")
65            strSQL = Replace(strSQL, "[表格]", "")
66            strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
67        End If
68        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
          
          
          '               类型(2=微生物)<split2>申请ID<split2>病人来源<split2>报告时间<split2>报告人<split2>审核人<split2>审核时间<split2>检项目名称<split2>标本类型<split2>婴儿序号<split2>
          '               细菌名1<split3>描述1<split3>耐药机制1<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split2>
          '               细菌名2<split3>描述2<split3>耐药机制2<split3>
          '                   抗生素1<split4>抗生素结果1<split4>耐药性1<split4>药敏方法1<split4>用法用量11<split4>用法用量21<split4>血药浓度11<split4>血药浓度21<split4>尿药浓度11<split4>尿药浓度21<split3>
          '                   抗生素2<split4>抗生素结果2<split4>耐药性2<split4>药敏方法2<split4>用法用量12<split4>用法用量22<split4>血药浓度12<split4>血药浓度22<split4>尿药浓度12<split4>尿药浓度22<split1>
          
          
69        If rsTmp.RecordCount <= 0 Then
70            strSQL = "Select  distinct [关键字] 2 Type, a.申请id, a.病人来源, b.报告时间, b.检验人, b.审核人, b.审核时间, c.正常菌 检验项目名称, c.未检出 细菌, c.培养描述, c.耐药机制, '' 抗生素, '' 抗生素结果, '' 结果类型," & vbNewLine & _
                       "          '' 药敏方法, '' 用法用量1, '' 用法用量2, '' 血药浓度1, '' 血药浓度2, '' 尿药浓度1, '' 尿药浓度2, c.细菌id, b.标本类型, a.婴儿 婴儿序号" & vbNewLine & _
                       "   From 检验申请组合 A, 检验报告记录 B, 检验报告细菌 C [表格]" & vbNewLine & _
                       "   Where a.标本id = b.Id And b.微生物 = 1 And b.Id = c.标本id " & vbNewLine & _
                       "      [条件] " & vbNewLine & _
                       "   Order By a.申请id, c.细菌id"
                      
71            If lngPatient = 0 Then
      '            strSQL = Replace(strSQL, "[条件]", " and A.申请id In (" & stridSQL & ")")
72                strSQL = Replace(strSQL, "[关键字]", "/*+cardinality(d,10)*/")
73                strSQL = Replace(strSQL, "[表格]", ",(" & stridSQL & ") D")
74                strSQL = Replace(strSQL, "[条件]", " and a.申请id=d.Column_Value")
75            Else
76                strSQL = Replace(strSQL, "[关键字]", "")
77                strSQL = Replace(strSQL, "[表格]", "")
78                strSQL = Replace(strSQL, "[条件]", " and A.HIS病人ID = [1] ")
79            End If
80            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatient, Mid(strTemp, 2), strAdvices)
81        End If
          
82        Do Until rsTmp.EOF
83            If lngID <> NVL(rsTmp("申请ID"), 0) Then
84                strSampleTwo = strSampleTwo & conSplit1 & "2" & conSplit2 & IIf(IsNull(rsTmp("申请ID")), "", rsTmp("申请ID")) & conSplit2 & IIf(IsNull(rsTmp("病人来源")), "", rsTmp("病人来源")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("报告时间")), "", rsTmp("报告时间"))) & conSplit2 & _
                              IIf(IsNull(rsTmp("检验人")), "", rsTmp("检验人")) & conSplit2 & IIf(IsNull(rsTmp("审核人")), "", rsTmp("审核人")) & conSplit2 & ReviseDate(IIf(IsNull(rsTmp("审核时间")), "", rsTmp("审核时间"))) & conSplit2 & IIf(IsNull(rsTmp("检验项目名称")), "", rsTmp("检验项目名称")) & conSplit2 & _
                              IIf(IsNull(rsTmp("标本类型")), "", rsTmp("标本类型")) & conSplit2 & NVL(rsTmp("婴儿序号"), "0") & conSplit2 & IIf(IsNull(rsTmp("细菌")), "", rsTmp("细菌")) & conSplit3 & IIf(IsNull(rsTmp("培养描述")), "", rsTmp("培养描述")) & conSplit3 & IIf(IsNull(rsTmp("耐药机制")), "", rsTmp("耐药机制")) & _
                              conSplit3 & IIf(IsNull(rsTmp("抗生素")), "", rsTmp("抗生素")) & conSplit4 & IIf(IsNull(rsTmp("抗生素结果")), "", rsTmp("抗生素结果")) & conSplit4 & IIf(IsNull(rsTmp("结果类型")), "", rsTmp("结果类型")) & conSplit4 & IIf(IsNull(rsTmp("药敏方法")), "", rsTmp("药敏方法")) & conSplit4 & IIf(IsNull(rsTmp("用法用量1")), "", rsTmp("用法用量1")) & _
                              conSplit4 & IIf(IsNull(rsTmp("用法用量2")), "", rsTmp("用法用量2")) & conSplit4 & IIf(IsNull(rsTmp("血药浓度1")), "", rsTmp("血药浓度1")) & conSplit4 & IIf(IsNull(rsTmp("血药浓度2")), "", rsTmp("血药浓度2")) & conSplit4 & IIf(IsNull(rsTmp("尿药浓度1")), "", rsTmp("尿药浓度1")) & conSplit4 & IIf(IsNull(rsTmp("尿药浓度2")), "", rsTmp("尿药浓度2"))
85                lngMicroID = NVL(rsTmp("细菌ID"), 0)
86            Else
87                If lngMicroID <> NVL(rsTmp("细菌ID"), 0) Then
88                    strSampleTwo = strSampleTwo & conSplit2 & rsTmp("细菌") & conSplit3 & rsTmp("培养描述") & conSplit3 & rsTmp("耐药机制") & _
                              conSplit3 & rsTmp("抗生素") & conSplit4 & rsTmp("抗生素结果") & conSplit4 & rsTmp("结果类型") & conSplit4 & rsTmp("药敏方法") & conSplit4 & rsTmp("用法用量1") & _
                              conSplit4 & rsTmp("用法用量2") & conSplit4 & rsTmp("血药浓度1") & conSplit4 & rsTmp("血药浓度2") & conSplit4 & rsTmp("尿药浓度1") & conSplit4 & rsTmp("尿药浓度2")
89                Else
90                    strSampleTwo = strSampleTwo & conSplit3 & rsTmp("抗生素") & conSplit4 & rsTmp("抗生素结果") & conSplit4 & rsTmp("结果类型") & conSplit4 & rsTmp("药敏方法") & _
                              conSplit4 & rsTmp("用法用量1") & conSplit4 & rsTmp("用法用量2") & conSplit4 & rsTmp("血药浓度1") & conSplit4 & rsTmp("血药浓度2") & _
                              conSplit4 & rsTmp("尿药浓度1") & conSplit4 & rsTmp("尿药浓度2")
91                End If
92            End If
93            lngID = NVL(rsTmp("申请ID"), 0)
94            lngMicroID = NVL(rsTmp("细菌ID"), 0)
95            rsTmp.MoveNext
96        Loop

97        If strSampleOne <> "" Then
98            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
99        End If
100       If strSampleTwo <> "" Then
101           strSampleTwo = Mid(strSampleTwo, Len(conSplit1) + 1)
102       End If
103       If strSampleTwo <> "" Then
104           If strSampleOne = "" Then
105               CreateLisValueStrForTJ = strSampleTwo
106           Else
107               CreateLisValueStrForTJ = strSampleOne & conSplit1 & strSampleTwo
108           End If
109       Else
110           CreateLisValueStrForTJ = strSampleOne
111       End If


112       Exit Function
CreateLisValueStrForTJ_Error:
113       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(CreateLisValueStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
114       Err.Clear
          
End Function

Public Function DelLisApplication(strAdvices As String, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   删除LIS中对应的请单
          '参数                   strAdvices 医嘱内容,格式
          '                       格式：<采集医嘱1,采集医嘱2,.....>
          '                       strErr 有错误信息时返回错误信息
          '返回                   TRUE=发送成功 FALSE=发送失败
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
              
1         On Error GoTo DelLisApplication_Error

2         strSQL = "Zl_检验申请单_Delete('" & strAdvices & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "删除申请单")
4         DelLisApplication = True

5         Exit Function
DelLisApplication_Error:
6         strErr = "删除申请单出错：" & Err.Number & " " & Err.Description
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(DelLisApplication)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
8         Err.Clear
End Function


Public Function SendLisApplication(strAdvices As String, strDiagnose As String, Optional strErr As String) As Boolean
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能                   发送检验申请单到LIS系统中
      '参数                   strAdvices 医嘱内容,格式
      '                       格式：<检验医嘱1,采集医嘱1,执行科室编码1,标本1;检验医嘱2,采集医嘱2,执行科室编码2,标本2;.....>
      '                       strDiagnose 诊断信息
      '返回                   TRUE=发送成功 FALSE=发送失败
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intloop As Integer
          Dim astrList() As String
          Dim astrItem() As String
          Dim strData As String
          Dim astrSQL() As String
          Dim blnRollBack As Boolean

1         On Error GoTo SendLisApplication_Error

2         astrList = Split(strAdvices, ";")
3         For intloop = 0 To UBound(astrList)
4             astrItem = Split(astrList(intloop), ",")
5             strSQL = "Select  distinct a.Id,a.相关id,a.标本部位,a.执行科室id From 病人医嘱记录 a,病人医嘱发送 b  Where 相关id=[1]  and a.id= b.医嘱id  and  b.执行状态=0 "
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查找医嘱记录", astrItem(1))
7             If rsTmp.RecordCount > 0 Then
8                 Do While Not rsTmp.EOF
9                     strData = rsTmp!ID & "," & rsTmp!相关ID & "," & rsTmp!执行科室id & "," & rsTmp!标本部位
10                    blnRollBack = SendLisApplicationAll(strData, strDiagnose, strErr)
11                    If blnRollBack = False Then
12                        SendLisApplication = False
13                        Exit Function   '医嘱发送失败SendLisApplication也会等于true导致临床那边没有提示
14                    Else
15                        SendLisApplication = True
16                    End If
17                    rsTmp.MoveNext
18                Loop
19            End If
20        Next
21        SendLisApplication = True

22        Exit Function
SendLisApplication_Error:
23        SendLisApplication = False
24        strErr = "错误号：" & Err.Number & "    错误描述：" & Err.Description
25        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(SendLisApplication)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
26        Err.Clear
End Function

Public Function SendLisApplicationAll(strAdvices As String, strDiagnose As String, Optional strErr As String) As Boolean
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能                   发送检验申请单到LIS系统中
      '参数                   strAdvices 医嘱内容,格式
      '                       格式：<检验医嘱1,采集医嘱1,执行科室编码1,标本1;检验医嘱2,采集医嘱2,执行科室编码2,标本2;.....>
      '                       strDiagnose 诊断信息
      '返回                   TRUE=发送成功 FALSE=发送失败
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim rsTmp As ADODB.Recordset
          Dim rsGather As ADODB.Recordset
          Dim rsExpenses As ADODB.Recordset
          Dim rsGatherExpenses As ADODB.Recordset       '采集费用记录
          Dim rsBabyName As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim rsExpensesAdvice As ADODB.Recordset
          Dim rsAdviceAddition As ADODB.Recordset        '医嘱附项
          Dim strExpensesAdvice As String
          Dim strAdviceAddition As String                '医嘱附项的所有内容
          Dim strBabyName As String
          Dim strBabySex As String
          Dim strBirthDay As String
          Dim strSQL As String
          Dim intloop As Integer
          Dim astrList() As String
          Dim astrItem() As String
          Dim intItem As Integer
          Dim strData As String
          Dim astrSQL() As String
          Dim blnRollBack As Boolean
          Dim strPayState As String
          Dim strPayStateOne As String
          Dim intTypeFind As Integer
          Dim strExpensesItem As String   '检验项目医嘱
          Dim strRef As String            '参考附项
          Dim strRefComItem As String
          Dim blnHave As Boolean
          Dim lngRefItemID As Long      '参考要素ID
          'Dim strDiagnose As String


1         On Error GoTo SendLisApplicationAll_Error

          Dim strWriteItem As String
2         strWriteItem = "申请id,医嘱id,病人来源,病人id,婴儿,姓名,性别,年龄,诊疗编码,标本类型,申请人,申请时间,申请科室,床号,健康号,病人科室,紧急标志,挂号单,门诊号,住院号,出生日期,主页id," & _
                         "签收人,签收时间,采样人,采样时间,送检人,送检时间,计费状态,病区,病区编码,申请科室编码,病人科室编码,病人类型,路径状态,样本条码"

3         ReDim Preserve astrSQL(0)
4         astrList = Split(strAdvices, ";")
5         For intloop = 0 To UBound(astrList)
6             astrItem = Split(astrList(intloop), ",")
7             strData = ""
              '先查找到医嘱相关的信息
8             strSQL = "Select A.Id 医嘱id, A.相关id 申请id, A.开嘱时间 申请时间, a.病人来源, A.病人id, A.婴儿, C.姓名, decode(C.性别,'男',1,'女',2,'未知',9,0) 性别, a.年龄, A.开嘱医生 申请人, A.开嘱时间 申请时间, D.名称 申请科室," & vbNewLine & _
                       "       C.当前床号 床号, C.健康号, E.名称 病人科室, A.紧急标志, A.挂号单, C.门诊号, C.住院号, C.出生日期, A.主页id, B.接收人 签收人, B.接收时间 签收时间, B.采样人, B.采样时间," & vbNewLine & _
                       "       b.样本条码,decode(a.病人来源,2,s.病人类型,c.病人类型) 病人类型, decode(a.病人来源,2,s.路径状态,null)路径状态,  B.送检人, B.标本送出时间 送检时间,b.计费状态,a.诊疗项目ID,f.编码 诊疗编码,a.标本部位 标本类型," & vbNewLine & _
                       "       b.记录性质,e.编码 病人科室编码,d.编码 申请科室编码,g.编码 病区编码,g.名称 病区,a.皮试结果 耐受方案ID,a.申请序号 申请批号 " & vbNewLine & _
                       "From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 部门表 D, 部门表 E,诊疗项目目录 F,部门表 g,病案主页 s " & vbNewLine & _
                       "Where A.Id = B.医嘱id And A.病人id = C.病人id And A.开嘱科室id = D.Id And A.病人科室id = E.Id and a.诊疗项目id = f.id and a.病人id= s.病人id(+) and a.主页id = s.主页id(+)" & vbNewLine & _
                       "   and c.当前病区ID = g.id(+) and a.id = [1] "

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验医嘱", astrItem(0))

10            If rsTmp.RecordCount > 0 Then
11                If Val(rsTmp("婴儿") & "") > 0 Then
12                    strSQL = "Select b.婴儿姓名, Decode(Substr(b.婴儿性别, Instr(b.婴儿性别, '-')+1),'男',1,'女',2,'未知',9,0) 性别,b.出生时间" & vbNewLine & _
                               "   From 病人医嘱记录 A, 病人新生儿记录 B" & vbNewLine & _
                               "   Where a.病人id = b.病人id And a.主页id = b.主页id And a.婴儿 = b.序号 And" & vbNewLine & _
                               "         a.相关id In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And Rownum = 1"
13                    Set rsBabyName = ComOpenSQL(Sel_His_DB, strSQL, "采集医嘱", astrItem(1))
14                End If
15            End If
16            If Not rsBabyName Is Nothing Then
17                If rsBabyName.RecordCount > 0 Then
18                    strBirthDay = StringFormatDate(rsBabyName("出生时间") & "")
19                Else
20                    strBirthDay = StringFormatDate(rsTmp("出生日期") & "")
21                End If
22            Else
23                strBirthDay = StringFormatDate(rsTmp("出生日期") & "")
24            End If

              '医嘱附项
25            strSQL = "Select b.要素ID, b.项目, b.排列, b.内容" & vbNewLine & _
                       "From 病人医嘱附件 b" & vbNewLine & _
                       "Where b.医嘱id In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) " & vbNewLine & _
                       "Order By 医嘱id, 排列"
26            Set rsAdviceAddition = ComOpenSQL(Sel_His_DB, strSQL, "医嘱附项", astrItem(1))

              '取采集方式和采集科室
27            strSQL = "Select B.名称 采集方式, C.名称 采集科室 " & vbNewLine & _
                       "From 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C " & vbNewLine & _
                       "Where A.诊疗项目id = B.Id And A.执行科室id = C.Id and a.id = [1] "
28            Set rsGather = ComOpenSQL(Sel_His_DB, strSQL, "采集方式", astrItem(1))


29            strSQL = "select ID,诊疗类别 from 病人医嘱记录 where id in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) " & _
                       " union all " & _
                       "select ID,诊疗类别 from 病人医嘱记录 where 相关id in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) "
30            Set rsExpensesAdvice = ComOpenSQL(Sel_His_DB, strSQL, "采集方式", astrItem(1))
31            strExpensesAdvice = ""
32            strExpensesItem = ""
33            Do Until rsExpensesAdvice.EOF
34                If rsExpensesAdvice("诊疗类别") = "E" Then
35                    If strExpensesAdvice = "" Then
36                        strExpensesAdvice = rsExpensesAdvice("ID")
37                    Else
38                        strExpensesAdvice = strExpensesAdvice & "," & rsExpensesAdvice("ID")
39                    End If
40                Else
41                    If strExpensesItem = "" Then
42                        strExpensesItem = rsExpensesAdvice("ID")
43                    Else
44                        strExpensesItem = strExpensesItem & "," & rsExpensesAdvice("ID")
45                    End If
46                End If
47                rsExpensesAdvice.MoveNext
48            Loop
49            intTypeFind = GetAdviceFeeKind(Val(astrItem(1)))
              '剔除采集费用
50            strSQL = "Select 费别 ,Sum(应收金额) As 应收金额, Sum(Decode(记录状态, 1, 实收金额, 0)) 实收金额 From 住院费用记录 Where 执行状态 <> 9   " & vbNewLine & _
                       " and 医嘱序号 in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by 费别 "
51            If intTypeFind = 1 Then
52                strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
53            End If
54            Set rsExpenses = ComOpenSQL(Sel_His_DB, strSQL, "查找费用", strExpensesItem)
              '采集费用
55            strSQL = "Select 费别,Sum(应收金额) As 采集应收金额, Sum(Decode(记录状态, 1, 实收金额, 0)) 采集实收金额 From 住院费用记录 Where 执行状态 <> 9   " & vbNewLine & _
                       " and 医嘱序号 in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by 费别 "
56            If intTypeFind = 1 Then
57                strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
58            End If
59            Set rsGatherExpenses = ComOpenSQL(Sel_His_DB, strSQL, "查找费用", strExpensesAdvice)
60            strPayState = funFindAdvicePay(astrItem(0), rsTmp("病人来源") & "")
61            If strPayState <> "" Then
62                If InStr(strPayState, "|") > 0 Then
63                    strPayState = 0
64                Else
65                    strPayStateOne = Split(strPayState, ",")(1)
66                End If
67            End If
              '写入数据组织
68            astrItem = Split(strWriteItem, ",")

69            For intItem = 0 To UBound(astrItem)
70                If astrItem(intItem) = "年龄" Then
71                    strData = strData & ",'" & rsTmp(astrItem(intItem)) & "','" & GetAgeMid(0, rsTmp(astrItem(intItem)) & "") & "','" & GetAgeMid(1, rsTmp(astrItem(intItem)) & "") & "'"
72                ElseIf astrItem(intItem) Like "*时间*" Or astrItem(intItem) Like "*日期*" Then
73                    If astrItem(intItem) = "出生日期" Then
74                        If strBirthDay = "" Then
75                            strData = strData & ",null"
76                        Else
77                            strData = strData & "," & strBirthDay
78                        End If
79                    Else
80                        If rsTmp(astrItem(intItem)) & "" = "" Then
81                            strData = strData & ",null"
82                        Else
83                            strData = strData & "," & StringFormatDate(rsTmp(astrItem(intItem)))
84                        End If
85                    End If
86                ElseIf astrItem(intItem) = "计费状态" Then
87                    strData = strData & "," & Val(strPayStateOne)
88                ElseIf astrItem(intItem) = "诊疗编码" Then
89                    If gUserInfo.NodeNo <> "-" Then
90                        strSQL = "select id from 检验组合项目 where 诊疗编码 = [1] and (站点=[2] or 站点 is null)"
91                    Else
92                        strSQL = "select id from 检验组合项目 where 诊疗编码 = [1] "
93                    End If
94                    Set rsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "取组合项目", rsTmp(astrItem(intItem)) & "", gUserInfo.NodeNo)
95                    If rsItem.RecordCount > 0 Then
96                        strData = strData & ",'" & rsItem("ID") & "'"
97                    Else
98                        strData = strData & ",null"
99                    End If
100               Else
101                   strData = strData & ",'" & rsTmp(astrItem(intItem)) & "'"
102               End If
103           Next
104           If rsExpenses.RecordCount > 0 Then
105               strData = strData & ",'" & rsGather("采集方式") & "','" & rsGather("采集科室") & "','" & rsExpenses("应收金额") & "','" & rsExpenses("实收金额") & _
                            "','" & rsExpenses("费别") & "'"
106           Else
107               strData = strData & ",'" & rsGather("采集方式") & "','" & rsGather("采集科室") & "','" & "" & "','" & "" & _
                            "','" & "" & "'"
108           End If

              'strDiagnose = GetPatiDiagnose(Val(rsTmp("病人id") & ""), Val(rsTmp("主页id") & ""), Val(rsTmp("病人来源") & ""))

              '诊断
109           strData = strData & ",'" & strDiagnose & "'"

110           If Not rsBabyName Is Nothing Then
111               If rsBabyName.RecordCount > 0 Then
112                   strBabyName = rsBabyName("婴儿姓名") & ""
113                   strBabySex = rsBabyName("性别") & ""
115               End If
116           End If
117           strData = strData & ",'" & strBabyName & "','" & strBabySex & "'"
118           strAdviceAddition = ""

              '从医嘱附项中获取参考附项
119           If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then

120               If rsAdviceAddition.RecordCount > 0 Then
121                   Do Until rsAdviceAddition.EOF
122                       strAdviceAddition = strAdviceAddition & rsAdviceAddition("项目") & ":" & rsAdviceAddition("内容") & vbNewLine
123                       If rsAdviceAddition("要素ID") & "" <> "" And rsAdviceAddition("项目") & "" <> "" Then
124                           strRefComItem = GetRefItem(Val(rsAdviceAddition("要素ID") & ""), rsAdviceAddition("项目") & "", rsAdviceAddition("内容") & "", blnHave, lngRefItemID)
125                           If blnHave Then
126                               strRef = strRef & "<Split D>" & lngRefItemID & "<Split E>" & strRefComItem
127                           End If
128                       End If
129                       rsAdviceAddition.MoveNext
130                   Loop
131               End If
132           Else
133               If rsAdviceAddition.RecordCount > 0 Then
134                   Do Until rsAdviceAddition.EOF
135                       strAdviceAddition = strAdviceAddition & rsAdviceAddition("项目") & ":" & rsAdviceAddition("内容") & vbNewLine
136                       rsAdviceAddition.MoveNext
137                   Loop
138               End If
139           End If
140           strData = strData & ",'" & Replace(strAdviceAddition, vbCrLf, "") & "'"
141           If rsGatherExpenses.RecordCount > 0 Then
142               strData = strData & ",'" & rsGatherExpenses("采集应收金额") & "','" & rsGatherExpenses("采集实收金额") & "'"
143           Else
144               strData = strData & ",'',''"
145           End If

              '耐受试验
146           If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
147               If rsTmp("耐受方案ID") & "" <> "" Then
148                   strData = strData & "," & rsTmp("耐受方案ID") & ",'" & rsTmp("申请批号") & "'"
149               Else
150                   strData = strData & ",null,null"
151               End If
152           End If

              '参考附项
153           If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
154               If strRef <> "" Then
155                   If Mid(strRef, 1, Len("<Split D>")) = "<Split D>" Then strRef = Mid(strRef, Len("<Split D>") + 1)
156                   strData = strData & ",'" & strRef & "'"
157               End If
158           End If

159           strData = Mid(strData, 2)
160           If intloop > 0 Then
161               ReDim Preserve astrSQL(UBound(astrSQL) + 1)
162           End If
163           astrSQL(UBound(astrSQL)) = "Zl_检验申请单_Insert(" & strData & ")"
164       Next
165       blnRollBack = True
          '    gcnHisOracle.BeginTrans
166       For intloop = 0 To UBound(astrSQL)
167           Call ComExecuteProc(Sel_Lis_DB, astrSQL(intloop), "发送医嘱到检验")
168       Next
          '    gcnHisOracle.CommitTrans
169       blnRollBack = False
170       SendLisApplicationAll = True

          '发送刷新科内概况未采样标签申请
171       Call SendMessage("RefreshDeptSurvey0")

172       Exit Function
SendLisApplicationAll_Error:
173       SendLisApplicationAll = False
174       strErr = "错误号：" & Err.Number & "    错误描述：" & Err.Description
175       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(SendLisApplicationAll)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
176       Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-09-20
'功    能:  要素显示条件转计算条件
'入    参:
'           lngID           要素ID
'           strName         要素名
'           strShow         要素显示条件
'出    参:
'           strShow         是否存在要素
'           lngRefItemID    要素ID
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Function GetRefItem(ByVal lngID As Long, ByVal strName As String, ByVal strShow As String, ByRef blnHave As Boolean, ByRef lngRefItemID As Long) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsVal As ADODB.Recordset
          Dim strSQLValue As String
          Dim strArr() As String
          Dim i As Integer

1         On Error GoTo GetRefItem_Error

2         strSQL = "select 值域,值域来源,ID from 检验指标参考要素 where ID=[1] and 要素名=[2]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "参考要素", lngID, strName)
4         If rsTmp.EOF Then
5             strSQL = "select 值域,值域来源,ID from 检验指标参考要素 where 要素名=[1]"
6             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "参考要素", strName)
7         End If
          
8         If Not rsTmp.EOF Then
9             blnHave = True
10            lngRefItemID = Val(rsTmp("ID") & "")
11            If Val(rsTmp("值域来源") & "") > 0 Then
                  '值域为SQL
12                strSQLValue = "select * from (" & rsTmp("值域") & ") where 显示条件=[1]"
13                Set rsVal = ComOpenSQL(Sel_Lis_DB, strSQLValue, "值域", strShow)
14                If Not rsVal.EOF Then
15                    GetRefItem = rsVal("计算条件") & ""
16                End If
17            Else
                  '值域为手动输入
18                strSQLValue = rsTmp("值域") & ""
19                strArr = Split(strSQLValue, "<SP1>")
20                For i = 0 To UBound(strArr)
21                    If Trim(strShow) = Trim(Split(strArr(i), "<SP2>")(0)) Then
22                        GetRefItem = Trim(Split(strArr(i), "<SP2>")(1))
23                        Exit Function
24                    End If
25                Next
26            End If
27        Else
28            blnHave = False
29        End If


30        Exit Function
GetRefItem_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetRefItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
32        Err.Clear
End Function

Private Function GetAgeMid(intType As Integer, strAge As String) As String
    '功能           转换年龄
    '参数           0=取年龄数字 1=取年龄单位
    
    If intType = 0 Then
        GetAgeMid = Val(strAge)
    Else
        'GetAgeMid = Replace(strAge, Val(strAge), "")
        GetAgeMid = Mid(strAge, Len("" & Val(strAge)) + 1)
    End If
End Function

Public Function GetSampleDeptRS(Optional strErr As String) As ADODB.Recordset
          '功能       取得采集科室的数据集
          '返回       找到的采集科室数据集

          Dim strSQL As String
1         On Error GoTo GetSampleDeptRS_Error

2         strSQL = "Select Distinct C.Id, C.编码, C.名称" & vbNewLine & _
                      "From 诊疗项目目录 A, 诊疗执行科室 B, 部门表 C" & vbNewLine & _
                      "Where A.类别 = 'E' And A.操作类型 = '6' And A.Id = B.诊疗项目id And B.执行科室id = C.Id"
3         Set GetSampleDeptRS = ComOpenSQL(Sel_His_DB, strSQL, "采集科室")


4         Exit Function
GetSampleDeptRS_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetSampleDeptRS)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
6         Err.Clear

End Function

Public Function GetSampleTypeRS(Optional strErr As String) As ADODB.Recordset
          '功能       取得采集项目的数据集
          '返回       找到的采集项目数据集

          Dim strSQL As String
1         On Error GoTo GetSampleTypeRS_Error

2         strSQL = "select id,编码,名称 from 诊疗项目目录 where 类别 = 'E' and 操作类型 = '6' "
          
3         Set GetSampleTypeRS = ComOpenSQL(Sel_His_DB, strSQL, "采集科室")


4         Exit Function
GetSampleTypeRS_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetSampleTypeRS)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
6         Err.Clear

End Function

Public Function Get诊疗执行科室(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng项目id As Long, ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, _
    Getlng诊疗执行科室ID As Long, Getstr诊疗执行科室名 As String, _
    Optional ByVal int范围 As Integer = 2, Optional strErr As String) As Boolean
          
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim int执行科室 As Integer
          Dim lng操作员科室ID As Long
          Dim bln上班安排 As Boolean
          Dim lngID As Long
          Dim bytDay As Byte
          
          
1         On Error GoTo Get诊疗执行科室_Error

2         strSQL = "Select A.Id, A.执行科室 From 诊疗项目目录 A, 诊疗用法用量 B Where A.Id = B.用法id And B.项目id = [1] And a.服务对象 IN([2],3)"

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", lng项目id, int范围)
4         If rsTmp.RecordCount > 0 Then
5             lngID = rsTmp("ID")
6             int执行科室 = rsTmp("执行科室")
7         End If
          
8         Select Case int执行科室
              Case 0, 5 '0-无执行的叮嘱,5-院外执行
9                 Get诊疗执行科室 = True: Exit Function
10            Case 1 '1-病人所在科室
11                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([1]) Order by 编码"
12            Case 2 '2-病人所在病区
13                If int范围 = 1 Then
14                    strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([1]) Order by 编码"
15                Else
16                    strSQL = _
                          " Select A.ID,A.编码,A.简码,A.名称" & _
                          " From 部门表 A,病案主页 B" & _
                          " Where A.ID=B.当前病区ID And B.病人ID=[2] And B.主页ID=[3] "
17                End If
18            Case 3 '3-操作员所在科室
19                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([4]) Order by 编码"
20                lng操作员科室ID = Get操作员部门ID(int范围)
21            Case 4 '4-指定科室
22                bln上班安排 = Check上班安排(False)
23                If Not bln上班安排 Then
24                    strSQL = _
                          " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                          " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
                          " Where A.ID=B.执行科室ID And B.诊疗项目ID=[5] And A.ID=C.部门ID" & _
                          " And C.服务对象 IN([6],3) And (B.病人来源 is NULL Or B.病人来源=[6])" & _
                          " And (B.开单科室ID is NULL Or B.开单科室ID=[1])" & _
                          " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)"
25                Else
26                    bytDay = Weekday(Currentdate, vbMonday) Mod 7 '0=周日,1=周一
27                    strSQL = _
                          " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                          " From 诊疗执行科室 A,部门安排 B,部门表 C,部门性质说明 D" & _
                          " Where A.执行科室ID+0=B.部门ID And B.部门ID=C.ID " & _
                          " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                          " And C.ID=D.部门ID And D.服务对象 IN([6],3) And (A.病人来源 is NULL Or A.病人来源=[6])" & _
                          " And (A.开单科室ID is NULL Or A.开单科室ID=[1]) And A.诊疗项目ID=[5]" & _
                          " And (C.站点='" & gUserInfo.NodeNo & "' Or C.站点 is Null)"
28                End If
29            Case 6 '6-开单人所在科室
30                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([11],[6]) Order by 编码"
31        End Select
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng病人科室ID, lng病人ID, lng主页ID, lng操作员科室ID, lngID, int范围)
33        If Not rsTmp.EOF Then
34            Getlng诊疗执行科室ID = rsTmp!ID
35            Getstr诊疗执行科室名 = rsTmp!名称
36            rsTmp.Filter = "ID=" & lng病人科室ID
37            If rsTmp.EOF Then rsTmp.Filter = "ID=" & lng病人科室ID
      '        If rsTmp.EOF And int范围 = 2 Then rsTmp.Filter = "ID=" & Get病区ID(lng病人ID, lng主页Id)
38            If Not rsTmp.EOF Then Getlng诊疗执行科室ID = rsTmp!ID: Getstr诊疗执行科室名 = rsTmp!名称
39        End If
40        rsTmp.Filter = ""
          'Set Get诊疗执行科室 = rsTmp


41        Exit Function
Get诊疗执行科室_Error:
42        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(Get诊疗执行科室)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
43        Err.Clear

End Function

Public Function Get检验执行科室(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng项目id As Long, ByVal lng病人科室ID As Long, ByVal lng开嘱科室ID As Long, _
    Getlng诊疗执行科室ID As Long, Getstr诊疗执行科室名 As String, _
    Optional ByVal int范围 As Integer = 2, Optional strErr As String) As Boolean
          
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim int执行科室 As Integer
          Dim lng操作员科室ID As Long
          Dim bln上班安排 As Boolean
          Dim lngID As Long
          Dim bytDay As Byte
          
1         On Error GoTo Get检验执行科室_Error

2         strSQL = "Select A.Id, A.执行科室 From 诊疗项目目录 A  Where A.Id =   [1]"

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", lng项目id)
4         If rsTmp.RecordCount > 0 Then
5             lngID = rsTmp("ID")
6             int执行科室 = rsTmp("执行科室")
7         End If
          
8         Select Case int执行科室
              Case 0, 5 '0-无执行的叮嘱,5-院外执行
9                 Get检验执行科室 = True: Exit Function
10            Case 1 '1-病人所在科室
11                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([1]) Order by 编码"
12            Case 2 '2-病人所在病区
13                If int范围 = 1 Then
14                    strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([1]) Order by 编码"
15                Else
16                    strSQL = _
                          " Select A.ID,A.编码,A.简码,A.名称" & _
                          " From 部门表 A,病案主页 B" & _
                          " Where A.ID=B.当前病区ID And B.病人ID=[2] And B.主页ID=[3] "
17                End If
18            Case 3 '3-操作员所在科室
19                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([4]) Order by 编码"
20                lng操作员科室ID = Get操作员部门ID(int范围)
21            Case 4 '4-指定科室
22                bln上班安排 = Check上班安排(False)
23                If Not bln上班安排 Then
24                    strSQL = _
                          " Select Distinct A.ID,A.编码,A.简码,A.名称" & _
                          " From 部门表 A,诊疗执行科室 B,部门性质说明 C" & _
                          " Where A.ID=B.执行科室ID And B.诊疗项目ID=[5] And A.ID=C.部门ID" & _
                          " And C.服务对象 IN([6],3) And (B.病人来源 is NULL Or B.病人来源=[6])" & _
                          " And (B.开单科室ID is NULL Or B.开单科室ID=[1])" & _
                          " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)"
25                Else
26                    bytDay = Weekday(Currentdate, vbMonday) Mod 7 '0=周日,1=周一
27                    strSQL = _
                          " Select Distinct C.ID,C.编码,C.简码,C.名称" & _
                          " From 诊疗执行科室 A,部门安排 B,部门表 C,部门性质说明 D" & _
                          " Where A.执行科室ID+0=B.部门ID And B.部门ID=C.ID " & _
                          " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(B.开始时间,'HH24:MI:SS') and To_Char(B.终止时间,'HH24:MI:SS') " & _
                          " And C.ID=D.部门ID And D.服务对象 IN([6],3) And (A.病人来源 is NULL Or A.病人来源=[6])" & _
                          " And (A.开单科室ID is NULL Or A.开单科室ID=[1]) And A.诊疗项目ID=[5]" & _
                          " And (C.站点='" & gUserInfo.NodeNo & "' Or C.站点 is Null)"
28                End If
29            Case 6 '6-开单人所在科室
30                strSQL = "Select ID,编码,简码,名称 From 部门表 Where ID IN([11],[6]) Order by 编码"
31        End Select
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng病人科室ID, lng病人ID, lng主页ID, lng操作员科室ID, lngID, int范围)
33        If Not rsTmp.EOF Then
34            Getlng诊疗执行科室ID = rsTmp!ID
35            Getstr诊疗执行科室名 = rsTmp!名称
      '        rsTmp.Filter = "ID=" & lng病人科室ID
      '        If rsTmp.EOF Then rsTmp.Filter = "ID=" & lng病人科室ID
      '        If rsTmp.EOF And int范围 = 2 Then rsTmp.Filter = "ID=" & Get病区ID(lng病人ID, lng主页Id)
36            If Not rsTmp.EOF Then Getlng诊疗执行科室ID = rsTmp!ID: Getstr诊疗执行科室名 = rsTmp!名称
37        End If
38        rsTmp.Filter = ""


39        Exit Function
Get检验执行科室_Error:
40        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(Get检验执行科室)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
41        Err.Clear

End Function

Public Function Get操作员部门ID(ByVal int服务对象 As Integer, Optional strErr As String) As Long
      '功能：取操作员所属服务对指定对象的部门，缺省部门优先
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, blnNew As Boolean
          
1         On Error GoTo Get操作员部门ID_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          
7         If blnNew Then
8             strSQL = "Select Distinct B.部门ID,Nvl(B.缺省,0) as 缺省,C.服务对象 From 部门人员 B,部门性质说明 C" & _
                  " Where B.人员ID = [1] And B.部门ID=C.部门ID" & _
                  " Order by 缺省 Desc"
9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlLisHisComm", gUserInfo.ID)
10        End If
11        rsTmp.Filter = "服务对象 = 3 or 服务对象 = " & int服务对象
          
12        If Not rsTmp.EOF Then
13            Get操作员部门ID = rsTmp!部门ID
14        Else
15            Get操作员部门ID = gUserInfo.DeptID
16        End If


17        Exit Function
Get操作员部门ID_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(Get操作员部门ID)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
19        Err.Clear

End Function

Public Function Check上班安排(ByVal bln药房 As Boolean, Optional strErr As String) As Boolean
      '功能：检查医院的科室是否使用了上班安排
      '参数：bln药房=是检查药房上班还是其它科室
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String
          Static bln药房Load As Boolean
          Static bln药房Last As Boolean
          Static bln非药Load As Boolean
          Static bln非药Last As Boolean
          
1         On Error GoTo Check上班安排_Error

2         If bln药房 Then '是否有安排只需读取一次
3             If bln药房Load Then Check上班安排 = bln药房Last: Exit Function
4         Else
5             If bln非药Load Then Check上班安排 = bln非药Last: Exit Function
6         End If
              
7         If bln药房 Then
8             strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
                  " Where A.部门ID=B.部门ID  And Rownum<2"
9         Else
10            strSQL = "Select 1 From 部门性质说明 A,部门安排 B" & _
                  " Where A.部门ID=B.部门ID  And Rownum<2"
11        End If
12        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "Check上班安排")
13        Check上班安排 = rsTmp.RecordCount > 0
          
14        If bln药房 Then
15            bln药房Load = True: bln药房Last = Check上班安排
16        Else
17            bln非药Load = True: bln非药Last = Check上班安排
18        End If


19        Exit Function
Check上班安排_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(Check上班安排)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Function

Public Function SampleBarcodeUpdate(strAdvices As String, strBarCode As String, Optional strSamplingName As String, Optional strErr As String, Optional ByVal intContinue As Integer) As Boolean
          '功能                   在采样时完成时写入样本条码到申请记录中，取消写入空的条码信息
          '
          '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
          '                       strBarCode   要写入的条码
          '                       格式:“医嘱1,医嘱2,医嘱3,..
          '                       intContinue 1=让步检验,2=取消检验,3=重采样本,4=生产条码,5=取消采集,6-取消条码,7-完成采集,8-取消采集和条码,9-完成采集和生成条码
          '返回                   如果返回False时可在strErr中显示具体错误
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim intloop As Integer
          Dim strbuff As String
          Dim rsbuff As ADODB.Recordset, rsTmp As ADODB.Recordset
          Dim varAdvices As Variant
          Dim strAdvice As String
          Dim blnGet As Boolean
          Dim strProgram As String
          
1         On Error GoTo SampleBarcodeUpdate_Error
          
2         If Mid(strAdvices, Len(strAdvices), 1) = "," Then
3             strAdvices = Mid(strAdvices, 1, Len(strAdvices) - 1)
4         End If
          
5         varAdvices = Split(strAdvices, ",")
6         For intloop = 0 To UBound(varAdvices)
7             If Val(varAdvices(intloop)) <> 0 Then
8                 strbuff = "Select id from 检验申请组合 where 医嘱id =[1] or 申请id =[1]"
9                 Set rsbuff = ComOpenSQL(Sel_Lis_DB, strbuff, "采集站查询", varAdvices(intloop))
10                If rsbuff.EOF Then
                      '重新生成申请单数据
11                    strbuff = "Select  distinct a.Id,a.相关id,a.标本部位,a.执行科室id From 病人医嘱记录 a,病人医嘱发送 b  Where id=[1]  and a.id= b.医嘱id  and  b.执行状态=0 "
12                    Set rsTmp = ComOpenSQL(Sel_His_DB, strbuff, "查找医嘱记录", varAdvices(intloop))
13                    If rsTmp.RecordCount > 0 Then
14                        If Val(rsTmp!相关ID & "") <> 0 Then
15                            strAdvice = rsTmp!ID & "," & rsTmp!相关ID & "," & rsTmp!执行科室id & "," & rsTmp!标本部位
16                            blnGet = SendLisApplication(strAdvice, "", strErr)
17                        End If
18                    End If
19                End If
20            End If
21        Next
          
22        If VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
23            If strSamplingName = "" Then
24                If intContinue = 2 Then
                      '拒收后取消完成
25                    strSQL = "Zl_检验申请条码_Updatenew(0,'" & strAdvices & "','" & strBarCode & "',Null,2)"
26                ElseIf intContinue = 4 Then
                      '生成条码
27                    strSQL = "Zl_检验申请条码_Updatenew(1,'" & strAdvices & "','" & strBarCode & "')"
28                ElseIf intContinue = 6 Then
                      '取消条码
29                    strSQL = "Zl_检验申请条码_Updatenew(2,'" & strAdvices & "','" & strBarCode & "')"
30                ElseIf intContinue = 5 Then
                      '取消采集
31                    strSQL = "Zl_检验申请条码_Updatenew(4,'" & strAdvices & "','" & strBarCode & "')"
32                ElseIf intContinue = 8 Then
                      '取消采集和条码
33                    strSQL = "Zl_检验申请条码_Updatenew(6,'" & strAdvices & "','" & strBarCode & "')"
34                End If
35            Else
36                If intContinue = 7 Then
                      '完成采集
37                    strSQL = "Zl_检验申请条码_Updatenew(3,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
38                ElseIf intContinue = 9 Then
                      '生成条码和完成采集
39                    strSQL = "Zl_检验申请条码_Updatenew(5,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
40                Else
                      '拒收后让步或重采
41                    strSQL = "Zl_检验申请条码_Updatenew(0,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "'," & intContinue & ")"
42                End If
43            End If
44        Else
              '--0=写入条码 采集人等，1=取消采集人，2-重新写入条码，3-取消条码，4-完成采集
45            If strSamplingName = "" Then
46                If intContinue = 5 Then
47                    strSQL = "Zl_检验申请条码_Update(5,'" & strAdvices & "','" & strBarCode & "')"
48                ElseIf intContinue = 4 Then
49                    strSQL = "Zl_检验申请条码_Update(2,'" & strAdvices & "','" & strBarCode & "')"
50                ElseIf intContinue = 6 Then
51                    strSQL = "Zl_检验申请条码_Update(3,'" & strAdvices & "','" & strBarCode & "')"
52                Else
53                     strSQL = "Zl_检验申请条码_Update(1,'" & strAdvices & "','" & strBarCode & "')"
54                End If
55            Else
56                If intContinue = 4 Then
57                    strSQL = "Zl_检验申请条码_Update(2,'" & strAdvices & "','" & strBarCode & "')"
58                ElseIf intContinue = 7 Then
59                    strSQL = "Zl_检验申请条码_Update(4,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "')"
60                Else
61                    strSQL = "Zl_检验申请条码_Update(0,'" & strAdvices & "','" & strBarCode & "','" & strSamplingName & "'," & intContinue & ")"
62                End If
63            End If
64        End If
          
65        Call ComExecuteProc(Sel_Lis_DB, strSQL, "写入条码")
          
          '拒收业务下的让步检验、取消完成、重采消息
66        If intContinue = 1 Or intContinue = 2 Or intContinue = 3 Then
67            If funWriteInLisNotify(1, strAdvices, intContinue, strErr) = False Then Exit Function
68        End If
          
69        SampleBarcodeUpdate = True
          
          '发送刷新科内概况未送检标签申请
70        Call SendMessage("RefreshDeptSurvey1")
          
71        Exit Function
SampleBarcodeUpdate_Error:
72        SampleBarcodeUpdate = False
73        strErr = "写入申请单条码出错：" & Err.Number & " " & Err.Description
74        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(SampleBarcodeUpdate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
75        Err.Clear
End Function

Public Function funGetItemMoney(strAdvices As String, Optional strErr As String) As Boolean
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能               通过医嘱ID得到当前项目收费的金额
      '参数               strAdvices       多个医嘱ID用逗号分隔
      '                   strErr           如果有错误信息时返回错误信息
      '返回               收费状态,应收金额,实收金额
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsPatient As ADODB.Recordset
          Dim rsGatherExpenses As ADODB.Recordset
          Dim strItem As String
          Dim intPatient As Integer
          Dim intType As Integer
          Dim astrAdvice() As String
          Dim intloop As Integer
          Dim astrSQL() As String
          Dim intTypeFind As Integer
          Dim strFindAdice As String

1         On Error GoTo funGetItemMoney_Error

2         ReDim Preserve astrSQL(0)

3         astrAdvice = Split(strAdvices, ",")

4         For intloop = 0 To UBound(astrAdvice)
5             strSQL = "select a.病人来源,a.id,b.记录性质,a.诊疗类别 from 病人医嘱记录 a,病人医嘱发送 b where a.id = b.医嘱id and a.相关id = [1] "
6             Set rsPatient = ComOpenSQL(Sel_His_DB, strSQL, "查找费用", Val(astrAdvice(intloop)))
7             If rsPatient.RecordCount > 0 Then
8                 Do Until rsPatient.EOF
9                     If rsPatient("记录性质") = "E" Then
10                        strItem = strItem & "," & rsPatient("ID")
11                    Else
12                        strFindAdice = strFindAdice & "," & rsPatient("ID")
13                    End If
14                    intPatient = Val(rsPatient("病人来源") & "")
15                    intType = Val(rsPatient("记录性质") & "")
16                    rsPatient.MoveNext
17                Loop
18                intTypeFind = GetAdviceFeeKind(Val(astrAdvice(intloop)))

19                If strItem <> "" Then
20                    strItem = Val(astrAdvice(intloop)) & strItem
21                Else
22                    strItem = Val(astrAdvice(intloop))
23                End If
                  '剔除采集费用
24                strSQL = "Select /*+ rule */ 记录状态,Sum(应收金额) As 应收金额, Sum(Decode(记录状态, 1, 实收金额, 0)) 实收金额 From 住院费用记录 Where 执行状态 <> 9   " & vbNewLine & _
                         " and 医嘱序号 in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by 记录状态 "
25                If intTypeFind = 1 Then
26                    strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
27                End If
28                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查找费用", strFindAdice)
                  '采集费用
29                strSQL = "Select /*+ rule */ 收费类别,Sum(应收金额) As 采集应收金额, Sum(Decode(记录状态, 1, 实收金额, 0)) 采集实收金额 From 住院费用记录 Where 执行状态 <> 9   " & vbNewLine & _
                         " and 医嘱序号 in (Select Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) group by 收费类别 "
30                If intTypeFind = 1 Then
31                    strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
32                End If
33                Set rsGatherExpenses = ComOpenSQL(Sel_His_DB, strSQL, "查找费用", strItem)

34                If rsTmp.RecordCount > 0 Then
35                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
36                    If rsGatherExpenses.RecordCount > 0 Then
37                        astrSQL(UBound(astrSQL)) = "Zl_申请单金额_Update('" & Val(astrAdvice(intloop)) & "','" & rsTmp("应收金额") & "','" & _
                                                     rsTmp("实收金额") & "','" & rsTmp("记录状态") & "','" & rsGatherExpenses("采集应收金额") & "','" & rsGatherExpenses("采集实收金额") & "')"
38                    Else
39                        astrSQL(UBound(astrSQL)) = "Zl_申请单金额_Update('" & Val(astrAdvice(intloop)) & "','" & rsTmp("应收金额") & "','" & _
                                                     rsTmp("实收金额") & "','" & rsTmp("记录状态") & "')"
40                    End If
41                End If
42                strItem = ""
43               strFindAdice = ""
44            End If
45        Next
46        For intloop = 0 To UBound(astrSQL)
47            If astrSQL(intloop) <> "" Then
48                Call ComExecuteProc(Sel_Lis_DB, astrSQL(intloop), "更新费用信息")
49            End If
50        Next
          
51        funGetItemMoney = True

52        Exit Function
funGetItemMoney_Error:
53        strErr = "更新医嘱费用错误" & Err.Number & " " & Err.Description
54        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetItemMoney)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
55        Err.Clear
End Function

Public Function funSampleCheckinInfoWrite(strAdvices As String, strName As String, strBatchNO As Long, Optional strErr As String, Optional ByVal strSentName As String) As Boolean
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能                   签收时把接收人接收时间写入到LIS中
      '
      '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
      '                       strName      签收人
      '                       strBatchNO   批号
      '返回                   如果返回False时可在strErr中显示具体错误
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strValue As String
          Dim intloop As Integer
          Dim strSQL As String
          Dim strArr() As String
          Dim arrSql() As String
          Dim blnTran As Boolean

1         On Error GoTo funSampleCheckinInfoWrite_Error

2         If strAdvices <> "" Then
3             If Left(strAdvices, 1) = "|" Then strAdvices = Mid(strAdvices, 2)
4             strArr = Str2Array(strAdvices, "|", 4000)
5             ReDim arrSql(UBound(strArr))

6             For intloop = 0 To UBound(strArr)
7                 strValue = Replace(strArr(intloop), "|", ",")
8                 If VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
9                     strSQL = "Zl_检验申请签收_Update('" & strValue & "','" & strName & "'," & IIf(strBatchNO = 0, "null", strBatchNO) & ",'" & strSentName & "',0,1)"
10                Else
11                    strSQL = "Zl_检验申请签收_Update('" & strValue & "','" & strName & "'," & IIf(strBatchNO = 0, "null", strBatchNO) & ",'" & strSentName & "',0)"
12                End If
13                arrSql(intloop) = strSQL
14            Next

15            gcnLisOracle.BeginTrans
16            blnTran = True
17            For intloop = 0 To UBound(arrSql)
18                If arrSql(intloop) <> "" Then
19                    Call ComExecuteProc(Sel_Lis_DB, arrSql(intloop), "签收信息")
20                End If
21            Next
22            gcnLisOracle.CommitTrans
23            blnTran = False

24            funSampleCheckinInfoWrite = True
25        End If

          '发送刷新科内概况未核收标签申请
26        Call SendMessage("RefreshDeptSurvey3")


27        Exit Function
funSampleCheckinInfoWrite_Error:
28        If blnTran Then gcnLisOracle.RollbackTrans
29        strErr = "签收时回写标志错误" & Err.Number & " " & Err.Description
30        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funSampleCheckinInfoWrite)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
31        Err.Clear
End Function

Public Function Str2Array(ByVal strTxt As String, ByVal strDeli As String, ByVal intLength As Integer) As Variant
          '功能: 将超过指定长度字符串,转换成数组
          '参数: strTxt     传入字符串
          '      strDeli    分隔符 - 仅支持单字符分隔符
          '      intLength  指定最大长度
          Dim arrstr() As String
          
1         On Error GoTo Str2Array_Error
          
2         ReDim arrstr(0)
          
3         Do While Len(strTxt) >= intLength And Len(strDeli) = 1 And InStr(1, strTxt, strDeli) > 0
4             arrstr(UBound(arrstr)) = Left(strTxt, InStrRev(strTxt, strDeli, intLength) - 1)
5             strTxt = Mid(strTxt, Len(arrstr(UBound(arrstr))) + Len(strDeli) + 1)
6             If strTxt <> "" Then ReDim Preserve arrstr(UBound(arrstr) + 1)
7         Loop
          
8         If strTxt <> "" Then arrstr(UBound(arrstr)) = strTxt
          
9         Str2Array = arrstr
          
10        Exit Function
Str2Array_Error:
11        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(Str2Array)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
12        Err.Clear
End Function

Public Function funSampleSendInfo(strAdvices As String, intType As Integer, ByVal strUser As String, Optional strErr As String) As Boolean
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   标本发送人发送时间写入LIS中
          '
          '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
          '                       intType      --0为送检 1为取消送检
          '返回                   如果返回False时可在strErr中显示具体错误
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          Dim strSQL As String
      '    SampleCheckinInfoWrite
1         On Error GoTo funSampleSendInfo_Error

2         strAdvices = Replace(strAdvices, "|", ",")
3         If Mid(strAdvices, Len(strAdvices), 1) = "," Then
4             strAdvices = Mid(strAdvices, 1, Len(strAdvices) - 1)
5         End If
6         strSQL = "Zl_检验申请送检_Update('" & strAdvices & "'," & intType & ",'" & strUser & "')"
7         Call ComExecuteProc(Sel_Lis_DB, strSQL, "签收信息")
8         funSampleSendInfo = True
          
          '发送刷新科内概况未登记标签申请
9         Call SendMessage("RefreshDeptSurvey2")
          

10        Exit Function
funSampleSendInfo_Error:
11        strErr = "标本送检回写标志错误" & Err.Number & " " & Err.Description
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funSampleSendInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
13        Err.Clear
End Function

Private Function InPatient() As ADODB.Recordset
    '初始化病人记录集
    Dim rsRetur As New ADODB.Recordset
    
'    If rsRetur.State = adStateOpen Then rsRetur.Close
    rsRetur.Fields.Append "HIS病人ID", adBigInt, 19
    rsRetur.Fields.Append "姓名", adVarChar, 100
    rsRetur.Fields.Append "性别", adVarChar, 4
    rsRetur.Fields.Append "年龄", adVarChar, 20
    rsRetur.Fields.Append "年龄数字", adVarChar, 4
    rsRetur.Fields.Append "年龄单位", adVarChar, 10
    rsRetur.Fields.Append "病人来源", adVarChar, 4
    rsRetur.Fields.Append "床号", adVarChar, 10
    rsRetur.Fields.Append "病历号", adVarChar, 20
    rsRetur.Fields.Append "病人科室", adVarChar, 100
    rsRetur.Fields.Append "门诊号", adVarChar, 19
    rsRetur.Fields.Append "住院号", adVarChar, 19
    rsRetur.Fields.Append "病人科室编码", adVarChar, 10
    rsRetur.Fields.Append "病区", adVarChar, 100
    rsRetur.Fields.Append "病区编码", adVarChar, 100
    rsRetur.Fields.Append "样本条码", adVarChar, 18

    rsRetur.CursorLocation = adUseClient
    rsRetur.LockType = adLockOptimistic
    rsRetur.CursorType = adOpenStatic
    rsRetur.Open

    Set InPatient = rsRetur
End Function

Public Function funGetPatientAndAdivce(strFind As String, lngMachineType As Long, lngMachineID As Long, Optional strErr As String) As ADODB.Recordset
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能               按输入的病人信息查找HIS病人信息并返回对应的记录集
          '                   strFind = 查找的信息
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset, rsbuff As ADODB.Recordset
          Dim intloop As Integer
          Dim strAadvie As String, strbuff As String
          Dim varAdvices As Variant, strAdvice As String
          Dim blnGet As Boolean
          Dim strParentID As String
          Dim strBarCode As String

1         On Error GoTo funGetPatientAndAdivce_Error

2         If strFind = "" Then Exit Function
          
3         strSQL = "Select b.医嘱id, b.样本条码, a.病人id" & vbNewLine & _
                   "   From 病人医嘱发送 B, 病人医嘱记录 A" & vbNewLine & _
                   "   Where a.Id = b.医嘱id And a.相关id Is Not Null And  b.样本条码= [1]"
                  
4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "病人查找", strFind)
5         If rsTmp.RecordCount > 0 Then
6             Do Until rsTmp.EOF
7                 strAadvie = strAadvie & rsTmp("医嘱id") & ","
8                 strBarCode = rsTmp("样本条码")
9                 strParentID = rsTmp("病人id")
10                rsTmp.MoveNext
11            Loop
12            varAdvices = Split(strAadvie, ",")
13            For intloop = 0 To UBound(varAdvices)
14                If varAdvices(intloop) <> "" Then
15                    strbuff = "Select 申请id,样本条码 from 检验申请组合 where 医嘱id =[1] "
16                    Set rsbuff = ComOpenSQL(Sel_Lis_DB, strbuff, "采集站查询", varAdvices(intloop))
17                    If rsbuff.EOF Then
                           '重新生成申请单数据
18                         strbuff = "Select Id,相关id,标本部位,执行科室id From 病人医嘱记录  Where id=[1]"
19                         Set rsTmp = ComOpenSQL(Sel_His_DB, strbuff, "查找医嘱记录", varAdvices(intloop))
20                         If Val(rsTmp!相关ID) <> 0 Then
21                             strAdvice = rsTmp!ID & "," & rsTmp!相关ID & "," & rsTmp!执行科室id & "," & rsTmp!标本部位
22                             blnGet = SendLisApplication(strAdvice, "", strErr)
23                             If blnGet = False Then
'24                                 Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "调用医嘱生成条码 。调用医嘱id：" & varAdvices(intloop) & "，SendLisApplication函数未生成成功", False)
25                             Else
'26                                 Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "调用医嘱生成条码 。调用医嘱id：" & varAdvices(intloop) & "生成成功", False)
27                             End If
28                         End If
29                    Else
30                        If rsbuff("样本条码") & "" = "" Then
31                            strSQL = "Zl_检验申请条码_Update(2,'" & rsbuff("申请id") & "','" & strBarCode & "')"
32                             Call ComExecuteProc(Sel_Lis_DB, strSQL, "写入条码")
33                        End If
34                    End If
                      
35                End If
36            Next
37            Set funGetPatientAndAdivce = GetPatientRecordCode(strBarCode, lngMachineType, lngMachineID, strErr)
38        Else
39            Set funGetPatientAndAdivce = rsTmp
40        End If
          
41        Exit Function
funGetPatientAndAdivce_Error:
42        strErr = "出错函数(funGetPatientAndAdivce),出错信息:" & Err.Number & " " & Err.Description
43        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetPatientAndAdivce)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
44        Err.Clear
End Function

Private Function GetPatientRecordCode(strFind As String, lngMachineType As Long, lngMachineID As Long, Optional strErr As String) As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strWhere As String
          Dim lngSampleID As Long
          Dim blnSelFind As Boolean
          Dim strBarCode As String

1         On Error GoTo GetPatientRecordCode_Error

2         If lngMachineType = 1 Then
              '微生物
3             strSQL = "select distinct a.His病人id,病人ID,nvl(a.婴儿,0) 婴儿 ,decode(a.婴儿,null,a.姓名,0,a.姓名,a.婴儿姓名 ) 姓名,decode(a.婴儿,null,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),decode(a.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别," & vbNewLine & _
                          " decode(a.婴儿,null,a.年龄,0,a.年龄,null) 年龄,decode(a.婴儿,null,a.年龄单位,0,a.年龄单位,'婴') 年龄单位,decode(a.婴儿,null,a.年龄数字,0,a.年龄数字,null) 年龄数字,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                          "from 检验申请组合 a,检验组合项目 b " & vbNewLine & _
                          "where a.组合id = b.id  and nvl(a.申请状态,0) = 0  [条件] " & vbNewLine & _
                          "union all " & vbNewLine & _
                          "select distinct a.His病人id,病人ID,nvl(a.婴儿,0) 婴儿 ,decode(a.婴儿,null,a.姓名,0,a.姓名,a.婴儿姓名 ) 姓名,decode(a.婴儿,null,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),decode(a.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别," & vbNewLine & _
                          "   decode(a.婴儿,null,a.年龄,0,a.年龄,null) 年龄,decode(a.婴儿,null,a.年龄单位,0,a.年龄单位,'婴') 年龄单位,decode(a.婴儿,null,a.年龄数字,0,a.年龄数字,null) 年龄数字,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                          "from 检验申请组合 a,检验组合项目 b " & vbNewLine & _
                          "where a.组合id = b.id  and a.标本id = [3]  [条件] " & vbNewLine
          
4             If IsNumeric(strFind) And InStr("*-+./ABDG", Mid(strFind, 1, 1)) = 0 Then
                  '先按条码查找
5                 strWhere = " and 样本条码 = [1] "
                  
6                 strSQL = Replace(strSQL, "[条件]", strWhere)
7                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "核收", strFind, lngMachineID, lngSampleID)
8                 If rsTmp.RecordCount > 0 Then
9                     blnSelFind = False
10                    strBarCode = strFind
11                Else
12                    blnSelFind = True
13                End If
14            Else
15                blnSelFind = True
16            End If
              
17            If blnSelFind = True Then
18                blnSelFind = False
19                If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then '病人ID
20                    strWhere = " and a.HIS病人ID = [4] "
21                    strFind = Mid(strFind, 2)
22                ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then '住院号
23                    strWhere = " and a.住院号 = [1] "
24                    strFind = Mid(strFind, 2)
25                ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then '门诊号
26                    strWhere = " and a.门诊号 = [1] "
27                    strFind = Mid(strFind, 2)
28                ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then '挂号单
29                    strWhere = " and a.挂号单 = [1] "
30                    strFind = Mid(strFind, 2)
31                ElseIf Left(strFind, 1) = "/" Then '收费单据号
32                    strWhere = " and a.收费单号 = [1] "
33                    strFind = Mid(strFind, 2)
34                Else
                      '没有前缀时在病人id,住院号,门诊号,挂号单,收费单据号中查找
35                    strWhere = ""
                      
36                    strSQL = "Select Distinct B.His病人id,病人ID,nvl(b.婴儿,0) 婴儿 , decode(b.婴儿,null,b.姓名,0,b.姓名,b.婴儿姓名) 姓名,decode(b.婴儿,null,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),decode(b.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别, " & vbNewLine & _
                              "   decode(b.婴儿,null,B.年龄,0,b.年龄,null) 年龄,decode(b.婴儿,null,b.年龄单位,0,b.年龄单位,'婴') 年龄单位,decode(b.婴儿,null,b.年龄数字,0,b.年龄数字,null) 年龄数字 ,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                              "From (Select His病人id" & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where   His病人id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where    住院号 =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where    门诊号 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where    挂号单 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His病人id From 检验申请组合 Where   收费单号 = [1] ) A, 检验申请组合 B, 检验组合项目 C " & vbNewLine & _
                              "Where A.His病人id = B.His病人id And B.组合id = C.Id  " & vbNewLine
                              
37                    strSQL = strSQL & " union all " & vbNewLine & _
                              "Select Distinct B.His病人id,病人ID ,nvl(b.婴儿,0) 婴儿 ,decode(b.婴儿,null,b.姓名,0,b.姓名,b.婴儿姓名) 姓名,decode(b.婴儿,null,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),decode(b.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别, " & vbNewLine & _
                              "   decode(b.婴儿,null,B.年龄,0,b.年龄,null) 年龄,decode(b.婴儿,null,b.年龄单位,0,b.年龄单位,null) 年龄单位,decode(b.婴儿,null,b.年龄数字,0,b.年龄数字,null) 年龄数字,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                              "From (Select His病人id" & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where 标本id = [3] and  His病人id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  住院号 =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  门诊号 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  挂号单 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His病人id From 检验申请组合 Where  标本id = [3] and  收费单号 = [1] ) A, 检验申请组合 B, 检验组合项目 C " & vbNewLine & _
                              "Where A.His病人id = B.His病人id And B.组合id = C.Id " & vbNewLine
38                End If
39                strSQL = Replace(strSQL, "[条件]", strWhere)
40                strSQL = strSQL & strWhere
41                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "核收", strFind, lngMachineID, lngSampleID, Val(strFind))
42            End If
              
43        Else
              '普通
44            strSQL = "select distinct a.His病人id,病人ID,nvl(a.婴儿,0) 婴儿 ,decode(a.婴儿,null,a.姓名,0,a.姓名,a.婴儿姓名 ) 姓名,decode(a.婴儿,null,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(a.性别,1,'男',2,'女',9,'未知','不区分'),decode(a.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别," & vbNewLine & _
                          " decode(a.婴儿,null,a.年龄,0,a.年龄,null) 年龄,decode(a.婴儿,null,a.年龄单位,0,a.年龄单位,'婴') 年龄单位,decode(a.婴儿,null,a.年龄数字,0,a.年龄数字,null) 年龄数字,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                          "from 检验申请组合 a,检验组合项目 b,检验组合指标 c,检验仪器指标 d" & vbNewLine & _
                          "where a.组合id = b.id and b.id = c.组合id and c.项目id = d.项目id and nvl(a.申请状态,0) = 0 and d.仪器id = [2] [条件] " & vbNewLine & _
                          "union all " & vbNewLine & _
                          "select distinct a.His病人id,病人ID,nvl(a.婴儿,0) 婴儿 ,a.姓名,decode(a.性别,1,'男',2,'女',9,'未知','不区分') 性别,a.年龄,a.年龄单位,a.年龄数字,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                          "from 检验申请组合 a,检验组合项目 b,检验组合指标 c,检验仪器指标 d" & vbNewLine & _
                          "where a.组合id = b.id and b.id = c.组合id and c.项目id = d.项目id and a.标本id = [3] and d.仪器id = [2] [条件] " & vbNewLine
          
45            If IsNumeric(strFind) And InStr("*-+./ABDG", Mid(strFind, 1, 1)) = 0 Then
                  '先按条码查找
46                strWhere = " and 样本条码 = [1] "
47                lngSampleID = 0
48                strSQL = Replace(strSQL, "[条件]", strWhere)
49                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "核收", strFind, lngMachineID, lngSampleID)
50                If rsTmp.RecordCount > 0 Then
51                    blnSelFind = False
52                    strBarCode = strFind
53                Else
54                    blnSelFind = True
55                End If
56            Else
57                blnSelFind = True
58            End If
59            If blnSelFind = True Then
60                If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then '病人ID
61                    strWhere = " and a.HIS病人ID = [4] "
62                    strFind = Mid(strFind, 2)
63                ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then '住院号
64                    strWhere = " and a.住院号 = [1] "
65                    strFind = Mid(strFind, 2)
66                ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then '门诊号
67                    strWhere = " and a.门诊号 = [1] "
68                    strFind = Mid(strFind, 2)
69                ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then '挂号单
70                    strWhere = " and a.挂号单 = [1] "
71                    strFind = Mid(strFind, 2)
72                ElseIf Left(strFind, 1) = "/" Then '收费单据号
73                    strWhere = " and a.收费单号 = [1] "
74                    strFind = Mid(strFind, 2)
75                Else
                      '没有前缀时在病人id,住院号,门诊号,挂号单,收费单据号中查找
76                    strWhere = ""
                      
77                    strSQL = "Select Distinct B.His病人id,病人ID,nvl(b.婴儿,0) 婴儿 ,decode(b.婴儿,null,b.姓名,0,b.姓名,b.婴儿姓名) 姓名, decode(b.婴儿,null,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),decode(b.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别,  " & vbNewLine & _
                              "   decode(b.婴儿,null,B.年龄,0,b.年龄,null) 年龄,decode(b.婴儿,null,b.年龄单位,0,b.年龄单位,'婴') 年龄单位,decode(b.婴儿,null,b.年龄数字,0,b.年龄数字,null) 年龄数字,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                              "From (Select His病人id" & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  His病人id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where    住院号 =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where   门诊号 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where    挂号单 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His病人id From 检验申请组合 Where    收费单号 = [1] ) A, 检验申请组合 B, 检验组合项目 C, 检验组合指标 D, 检验仪器指标 E" & vbNewLine & _
                              "Where A.His病人id = B.His病人id And B.组合id = C.Id And C.Id = D.组合id And D.项目id = E.项目id  and e.仪器ID = [2] " & vbNewLine
                              
78                    strSQL = strSQL & " union all " & vbNewLine & _
                              "Select Distinct B.His病人id,病人ID ,nvl(b.婴儿,0) 婴儿 ,decode(b.婴儿,null,b.姓名,0,b.姓名,b.婴儿姓名 ) 姓名, decode(b.婴儿,null,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),0,decode(b.性别,1,'男',2,'女',9,'未知','不区分'),decode(b.婴儿性别,1,'男',2,'女',9,'未知','不区分')) 性别," & vbNewLine & _
                              "   decode(b.婴儿,null,B.年龄,0,b.年龄,null) 年龄,decode(b.婴儿,null,b.年龄单位,0,b.年龄单位,'婴') 年龄单位,decode(b.婴儿,null,b.年龄数字,0,b.年龄数字,null) 年龄数字,decode(b.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源 " & vbNewLine & _
                              "From (Select His病人id" & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where 标本id = [3] and  His病人id = [4] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select  His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  住院号 =[1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select   His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  门诊号 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select    His病人id " & vbNewLine & _
                              "       From 检验申请组合" & vbNewLine & _
                              "       Where  标本id = [3] and  挂号单 = [1] " & vbNewLine & _
                              "       Union All" & vbNewLine & _
                              "       Select His病人id From 检验申请组合 Where  标本id = [3] and  收费单号 = [1] ) A, 检验申请组合 B, 检验组合项目 C, 检验组合指标 D, 检验仪器指标 E" & vbNewLine & _
                              "Where A.His病人id = B.His病人id And B.组合id = C.Id And C.Id = D.组合id And D.项目id = E.项目id  and e.仪器ID = [2] " & vbNewLine
79                End If
80                strSQL = Replace(strSQL, "[条件]", strWhere)
81                strSQL = strSQL & strWhere
82                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "核收", strFind, lngMachineID, lngSampleID, Val(strFind))
83            End If
              
84        End If
              
85        If rsTmp.RecordCount = 0 Then
              '从登记中的病历号中去查找
86            strSQL = "select Distinct HIS病人ID,病人ID,病历号, 姓名, 性别," & vbNewLine & _
                      "  年龄,年龄单位, 年龄数字,decode(病人来源,1,'门诊',2,'住院',3,'院外',4,'体检','') 病人来源,床号,费用类型 费别,decode(病人科室,null,申请科室,病人科室) 病人科室,门诊号,住院号, " & vbNewLine & _
                      " 病人类型,路径状态,病人科室编码,病区,病区编码 " & vbNewLine & _
                      " from 检验报告记录 where 病历号=[1] "
87            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "核收", strFind)
              
88            If rsTmp.RecordCount = 0 Then blnSelFind = True
89        End If
90        Set GetPatientRecordCode = rsTmp

91        Exit Function
GetPatientRecordCode_Error:
92        strErr = "出错函数(GetPatientRecordCode),出错信息:" & Err.Number & " " & Err.Description
93        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetPatientRecordCode)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
94        Err.Clear
End Function


Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer, Optional strErr As String) As String
      '功能：由用户输入的部份单号，返回全部的单号。
      '参数：intNum=项目序号,为0时固定按年产生
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, intType As Integer
          Dim curDate As Date
          
1         On Error GoTo GetFullNO_Error

2         If Len(strNO) >= 8 Then
3             GetFullNO = Right(strNO, 8)
4             Exit Function
5         ElseIf Len(strNO) = 7 Then
6             GetFullNO = PreFixNO & strNO
7             Exit Function
8         ElseIf intNum = 0 Then
9             GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
10            Exit Function
11        End If
12        GetFullNO = strNO
          
13        strSQL = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=" & intNum
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "取号码")
15        If Not rsTmp.EOF Then
16            intType = NVL(rsTmp!编号规则, 0)
17            curDate = rsTmp!日期
18        End If

19        If intType = 1 Then
              '按日编号
20            strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
21            GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
22        Else
              '按年编号
23            GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
24        End If


25        Exit Function
GetFullNO_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetFullNO)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear

End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function


Public Function GetPatiDiagnose(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int来源 As Integer, Optional strErr As String) As String
      '功能：读取病人指定次就诊的门诊诊断
      '参数：lng就诊ID=挂号ID或主页ID
      '      int来源=1-门诊,2-住院
      '返回：用"，"号分隔的多个诊断串
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
              
1         On Error GoTo GetPatiDiagnose_Error

2         strSQL = "Select 记录来源,诊断类型,诊断次序,诊断描述,是否疑诊,Mod(诊断类型,10) as 大类 From 病人诊断记录" & _
              " Where 病人ID=[1] And 主页ID=[2] And 诊断类型 IN(" & IIf(int来源 = 1, "1,11", "1,2,3,11,12,13") & ")" & _
              " Order by 记录来源,诊断类型,诊断次序"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "GetPatiDiagnose", lng病人ID, lng就诊ID)
          
          '先按来源优先顺序过滤
4         rsTmp.Filter = "记录来源=3" '首页整理
5         If rsTmp.EOF Then rsTmp.Filter = "记录来源=2" '入院登记
6         If rsTmp.EOF Then rsTmp.Filter = "记录来源=1" '病历
7         If rsTmp.EOF Then rsTmp.Filter = "记录来源=4" '病案室录入
          
          '住院再按类型优先顺序过滤
8         If Not rsTmp.EOF And int来源 = 2 Then
9             strSQL = rsTmp.Filter
10            rsTmp.Filter = strSQL & " And 大类=3"
11            If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=2"
12            If rsTmp.EOF Then rsTmp.Filter = strSQL & " And 大类=1"
13        End If
          
14        strSQL = ""
15        Do While Not rsTmp.EOF
16            If Not IsNull(rsTmp!诊断描述) Then
17                strSQL = strSQL & "，" & rsTmp!诊断描述 & IIf(NVL(rsTmp!是否疑诊, 0) = 1, "（？）", "")
18            End If
19            rsTmp.MoveNext
20        Loop
          
21        GetPatiDiagnose = Mid(strSQL, 2)


22        Exit Function
GetPatiDiagnose_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetPatiDiagnose)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
24        Err.Clear

End Function

Public Function GetAppendItemValue(ByVal str项目 As String, ByVal lng要素ID As Long, lng病人ID As Long, _
             var就诊ID As Variant, strDiagnosis As String, int婴儿 As Integer, strAdvItem As String) As String
      '功能：获取指定的申请附项值
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, strText As String
          Dim arrItem As Variant, i As Long
              
          '1.如果有对应要素，从要素提取函数读取
1         On Error GoTo GetAppendItemValue_Error

2         If lng要素ID <> 0 Then
3             If TypeName(var就诊ID) = "String" Then
4                 strSQL = "Select Zl_Replace_Element_Value(B.中文名,[1],A.ID,1) as 内容" & _
                      " From 病人挂号记录 A,诊治所见项目 B Where A.NO=[2] And B.ID=[3] And a.记录性质=1 And a.记录状态=1"
5                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS接口", lng病人ID, CStr(var就诊ID), lng要素ID)
6             Else
7                 strSQL = "Select Zl_Replace_Element_Value(中文名,[1],[2],2) as 内容" & _
                      " From 诊治所见项目 Where ID=[3]"
8                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS接口", lng病人ID, CStr(var就诊ID), lng要素ID)
9             End If
10            If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
11        End If
          
          '2.如果诊断，从未保存的已录入诊断中提取
12        If str项目 Like "*诊断" And strText = "" And strDiagnosis <> "" Then
13            strText = strDiagnosis
14        End If

          '3.未取到或未对应要素的，从病人之前未保存的医嘱中提取,以最后填写的为准
15        If strText = "" And strAdvItem <> "" Then
16            arrItem = Split(strAdvItem, "<Split1>")
17            For i = 0 To UBound(arrItem)
18                If Split(arrItem(i), "<Split2>")(0) = str项目 Then
19                    strText = Split(arrItem(i), "<Split2>")(3): Exit For
20                End If
21            Next
22        End If
          
          '4.未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
23        If strText = "" Then
24            strSQL = _
                  " Select 内容 From (" & _
                  " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
                  " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
                  IIf(TypeName(var就诊ID) = "String", " And A.挂号单=[2]", " And A.主页ID=[3]") & _
                  " And B.项目=[5] And B.内容 is Not Null" & _
                  " Order by A.开嘱时间 Desc) Where Rownum=1"
25            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS接口", lng病人ID, CStr(var就诊ID), Val(var就诊ID), int婴儿, str项目)
26            If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
27        End If
          
28        GetAppendItemValue = strText


29        Exit Function
GetAppendItemValue_Error:
30        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetAppendItemValue)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
31        Err.Clear

End Function

Public Function funWriteAdvicesLookState(strAdvices As String, intType As Integer, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   写入医嘱的查阅状态
          '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
          '                       intType      1=已查阅 0=未查阅
          '返回                   True=成功   False=失败
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          
1         On Error GoTo funWriteAdvicesLookState_Error

2         strSQL = "Zl_检验申请查阅_update('" & strAdvices & "','" & intType & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "写入报告查询状态")
          
4         funWriteAdvicesLookState = True


5         Exit Function
funWriteAdvicesLookState_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funWriteAdvicesLookState)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Currency
      '功能：获取指定病人当天发生的费用总额
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String

1         On Error GoTo GetPatiDayMoney_Error

2         strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlCISKernel", lng病人ID)
4         If Not rsTmp.EOF Then GetPatiDayMoney = NVL(rsTmp!金额, 0)


5         Exit Function
GetPatiDayMoney_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetPatiDayMoney)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function

Public Function ReCalcBirth(ByVal strOld As String, ByVal str年龄单位 As String) As String
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能:              根据年龄和年龄单位估算病人的出生日期,年龄单位为岁时,出年月日假定为1月1号,年龄单位为月时,出生日期假定为1号
          '
          '入参:
          '                   strOld               年龄
          '                   str年龄单位          年龄单位
          '
          '返回:              出生日期
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strTmp As String, strFormat As String, lngDays As Long
          
1         On Error GoTo ReCalcBirth_Error

2         strTmp = "____-__-__"
3         If str年龄单位 = "" Then
4             strFormat = "YYYY-MM-DD"
5             If strOld Like "*岁*月" Or strOld Like "*岁*个月" Then
6                 strFormat = "YYYY-MM-01"
7                 lngDays = 365 * Val(strOld) + 30 * Val(Mid(strOld, InStr(1, strOld, "岁") + 1))
8             ElseIf strOld Like "*月*天" Or strOld Like "*个月*天" Then
9                 lngDays = 30 * Val(strOld) + Val(Mid(strOld, InStr(1, strOld, "月") + 1))
10            ElseIf strOld Like "*岁" Or IsNumeric(strOld) Then
11                strFormat = "YYYY-01-01"
12                lngDays = 365 * Val(strOld)
13            ElseIf strOld Like "*月" Or strOld Like "*个月" Then
14                strFormat = "YYYY-MM-01"
15                lngDays = 30 * Val(strOld)
16            ElseIf strOld Like "*天" Then
17                lngDays = Val(strOld)
18            End If
19            If lngDays <> 0 Then strTmp = Format(DateAdd("d", lngDays * -1, Currentdate), strFormat)
20        ElseIf strOld <> "" Then
21            Select Case str年龄单位
                  Case "岁"
22                    If Val(strOld) > 200 Then lngDays = -1
23                Case "月"
24                    If Val(strOld) > 2400 Then lngDays = -1
25                Case "天"
26                    If Val(strOld) > 73000 Then lngDays = -1
27            End Select
              
28            If lngDays = 0 Then
29                strTmp = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
30                strTmp = Format(DateAdd(strTmp, Val(strOld) * -1, Currentdate), "YYYY-MM-DD")
                  
31                If str年龄单位 = "岁" Then
32                    strTmp = Format(strTmp, "YYYY-01-01")
33                ElseIf str年龄单位 = "月" Then
34                    strTmp = Format(strTmp, "YYYY-MM-01")
35                End If
36            End If
37        End If
38        ReCalcBirth = strTmp


39        Exit Function
ReCalcBirth_Error:
40        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(ReCalcBirth)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
41        Err.Clear
End Function

Public Function funModifyApplyItemStateYJ(strAdvices As String, intType As Integer, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   写入申请项目的申请状态
          '参数                   strAdvices   医嘱串,多个医嘱使用","号分隔
          '                       写入
          '返回                   True=成功   False=失败
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
          
1         On Error GoTo funModifyApplyItemStateYJ_Error

2         strSQL = "Zl_检验申请组合_Modify('" & strAdvices & "','" & intType & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "写入报告查询状态")
          
4         funModifyApplyItemStateYJ = True


5         Exit Function
funModifyApplyItemStateYJ_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funModifyApplyItemStateYJ)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function

Public Function funModifyPathState(lngPartentID As Long, lngMainID As Long, lngPathSatae As Long, Optional strErr As String) As Boolean
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   写入临床路径的路径状态
          '参数                   lngPartentID   病人id
          '                       lngMainID      主页id
          '                       lngPathSatae   路径状态
          '                       写入
          '返回                   True=成功   False=失败
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          Dim strSQL As String
          
1         On Error GoTo funModifyPathState_Error

2         strSQL = "Zl_检验路径状态_Modify('" & lngPartentID & "','" & lngMainID & "','" & lngPathSatae & "')"
3         Call ComExecuteProc(Sel_Lis_DB, strSQL, "写入报告查询状态")
          
4         funModifyPathState = True


5         Exit Function
funModifyPathState_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funModifyPathState)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function


'ZL9ComLib中ParseXMLToRecord的相关函数，如果函数改变，需同步改变
Public Function ParseXMLToRecord(ByVal strMsgNo As String, ByVal strXML As String) As ADODB.Recordset
      '功能：解析XML结构的字符串，转换成记录集的形式
          Dim rsMsg As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim objXML As Object
          Dim strTmp1 As String
          Dim strTmp2 As String
          
1         On Error GoTo ParseXMLToRecord_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         If InStr(",ZLHIS_EMR_021,ZLHIS_TRANSFUSION_001,ZLHIS_CHARGE_001,ZLHIS_PACS_005,ZLHIS_LIS_002,ZLHIS_LIS_003,ZLHIS_OPER_001," & _
              "ZLHIS_CIS_001,ZLHIS_CIS_002,ZLHIS_CIS_003,ZLHIS_CIS_004,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgNo & ",") = 0 Then Exit Function
          
          '传入的消息串可能不是完整的XML
4         Call objXML.OpenXMLDocument(IIf(InStr(strXML, "<message>") = 0, "<message>" & strXML & "</message>", strXML))
          
5         Set rsMsg = New ADODB.Recordset
6         rsMsg.Fields.Append "病人ID", adBigInt
7         rsMsg.Fields.Append "就诊ID", adVarChar, 20
8         rsMsg.Fields.Append "就诊科室ID", adBigInt
9         rsMsg.Fields.Append "就诊病区ID", adBigInt
10        rsMsg.Fields.Append "病人来源", adBigInt
11        rsMsg.Fields.Append "消息内容", adVarChar, 4000
12        rsMsg.Fields.Append "提醒场合", adVarChar, 8
13        rsMsg.Fields.Append "类型编码", adVarChar, 60
14        rsMsg.Fields.Append "业务标识", adVarChar, 120
15        rsMsg.Fields.Append "优先程度", adBigInt
16        rsMsg.Fields.Append "是否已阅", adBigInt
17        rsMsg.Fields.Append "登记时间", adVarChar, 60
18        rsMsg.Fields.Append "部门IDs", adVarChar, 4000
19        rsMsg.Fields.Append "提醒人员", adVarChar, 4000
20        rsMsg.CursorLocation = adUseClient
21        rsMsg.LockType = adLockOptimistic
22        rsMsg.CursorType = adOpenStatic
23        rsMsg.Open
          
24        rsMsg.AddNew
25        rsMsg!类型编码 = strMsgNo
26        rsMsg!是否已阅 = 0
27        rsMsg!优先程度 = 1
28        rsMsg!提醒人员 = ""
29        rsMsg!部门IDs = ""
          
30        Call objXML.GetSingleNodeValue("patient_id", strTmp1) '病人id
31        Call objXML.GetSingleNodeValue("clinic_id", strTmp2) '就诊id
32        rsMsg!病人ID = Val(strTmp1): strTmp1 = ""
33        rsMsg!就诊id = strTmp2
          
34        Call objXML.GetSingleNodeValue("send_time", strTmp1)
35        If strTmp1 <> "" Then
36            rsMsg!登记时间 = strTmp1
37        Else
38            rsMsg!登记时间 = Format(Currentdate, "yyyy-MM-dd HH:mm:ss")
39        End If

40        strTmp1 = "": strTmp2 = ""
41        If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_002,ZLHIS_LIS_003,ZLHIS_OPER_001,", "," & strMsgNo & ",") > 0 Then
              
42            Call objXML.GetSingleNodeValue("clinic_dept_id", strTmp1)
43            Call objXML.GetSingleNodeValue("clinic_area_id", strTmp2)
              
             'LIS系统传的时候没有传部门id传的是编码这里再取一次
44            If (strMsgNo = "ZLHIS_LIS_002" Or strMsgNo = "ZLHIS_LIS_003") And Val(strTmp1) = 0 And Val(strTmp2) = 0 Then
45                strTmp1 = "": strTmp2 = ""
46                Call objXML.GetSingleNodeValue("clinic_dept_code", strTmp1)
47                Call objXML.GetSingleNodeValue("clinic_area_code", strTmp2)
48                If strTmp1 <> "" Or strTmp2 <> "" Then
49                    If strTmp1 = strTmp2 Then
50                        strSQL = "select id from 部门表 where 编码=[1]"
51                        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "ParseXMLToRecord", strTmp1)
52                        If Not rsTmp.EOF Then
53                            strTmp1 = rsTmp!ID
54                            strTmp2 = strTmp1
55                        Else
56                            strTmp1 = "": strTmp2 = ""
57                        End If
58                    Else
59                        strSQL = "select id,编码 from 部门表 where 编码 in (Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)))"
60                        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "ParseXMLToRecord", strTmp1 & "," & strTmp2)
61                        If Not rsTmp.EOF Then
62                            rsTmp.Filter = "编码='" & strTmp1 & "'"
63                            strTmp1 = IIf(Not rsTmp.EOF, rsTmp!ID, "")
64                            rsTmp.Filter = "编码='" & strTmp2 & "'"
65                            strTmp2 = IIf(Not rsTmp.EOF, rsTmp!ID, "")
66                        Else
67                            strTmp1 = "": strTmp2 = ""
68                        End If
69                    End If
70                End If
71            End If
              
72            rsMsg!就诊科室id = Val(strTmp1)
73            rsMsg!就诊病区id = Val(strTmp2)
74            If Val(strTmp2) <> Val(strTmp1) Then
75                rsMsg!部门IDs = Val(strTmp1) & "," & Val(strTmp2)
76            Else
77                rsMsg!部门IDs = Val(strTmp1)
78            End If
79            strTmp1 = ""
80            Call objXML.GetSingleNodeValue("patient_source", strTmp1)
81            rsMsg!病人来源 = Val(strTmp1)
82            rsMsg!提醒场合 = "0110"
              
83            strTmp1 = "": strTmp2 = ""
84            If strMsgNo = "ZLHIS_LIS_002" Then
85                rsMsg!消息内容 = "有已阅报告被撤消。"
86                rsMsg!提醒场合 = "0100"
87                Call objXML.GetSingleNodeValue("specimen_id", strTmp1) '标本id
88                rsMsg!业务标识 = Val(strTmp1)
89                rsMsg!优先程度 = 2
90            ElseIf strMsgNo = "ZLHIS_LIS_003" Then
91                Call objXML.GetSingleNodeValue("element_title", strTmp1) '危急值名称
92                Call objXML.GetSingleNodeValue("element_value", strTmp2) '危急值值
93                rsMsg!消息内容 = "危急值：" & strTmp1 & "(" & strTmp2 & ")。"
                  
94                strTmp1 = ""
95                Call objXML.GetSingleNodeValue("order_id", strTmp1) '医嘱id
96                rsMsg!业务标识 = Val(strTmp1)
97                rsMsg!优先程度 = 3
98            ElseIf strMsgNo = "ZLHIS_PACS_005" Then
99                Call objXML.GetSingleNodeValue("check_item_title", strTmp1) '危急值值
100               rsMsg!消息内容 = "危急值：" & strTmp1 & "。"
101               Call objXML.GetSingleNodeValue("order_id", strTmp2) '医嘱id
102               rsMsg!业务标识 = Val(strTmp2)
103               rsMsg!优先程度 = 3
104           ElseIf strMsgNo = "ZLHIS_OPER_001" Then
105               Call objXML.GetSingleNodeValue("operation_item_title", strTmp1) '手术名称
106               Call objXML.GetSingleNodeValue("operation_time", strTmp2) '手术时间
                  
107               strSQL = "select 名称 from 部门表 where id=[1]"
108               Set rsTmp = OpenSQLRecord(strSQL, "ParseXMLToRecord", Val(rsMsg!就诊科室id))
109               rsMsg!消息内容 = rsTmp!名称 & "，" & strTmp1 & "安排到：" & Format(strTmp2, "yyyy-MM-dd HH:mm")
                  
110               strTmp1 = "": strTmp2 = ""
111               Call objXML.GetSingleNodeValue("request_id", strTmp1) '手术医嘱id
112               rsMsg!业务标识 = Val(strTmp1)
113               Call objXML.GetSingleNodeValue("major_doctor", strTmp2) '主刀医师
114               rsMsg!提醒人员 = strTmp2
115           End If
116       End If
117       rsMsg.Update
          
118       If rsMsg.RecordCount > 0 Then
119           rsMsg.MoveFirst
120           Set ParseXMLToRecord = rsMsg
121       End If


122       Exit Function
ParseXMLToRecord_Error:
123       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(ParseXMLToRecord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
124       Err.Clear
          
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
    '动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If Err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "创建电子病历部件"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function funExeDeptID(ByVal lngSampleID As Long) As Long
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能                   获取执行科室id
      '参数
      '                       longSampleid        标本id
      '
      '返回                   执行科室id
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo funExeDeptID_Error

2         strSQL = "Select  d.Id 执行科室id " & vbNewLine & _
                 "   From 检验报告记录 A, 检验仪器记录 B, 检验小组记录 C, 部门表 D" & vbNewLine & _
                 "   Where a.仪器id = b.Id And b.小组id = c.Id And c.His部门编码 = d.编码 and a.id =[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验签名", lngSampleID)

4         If rsTmp.RecordCount > 0 Then
5             strSQL = "select Zl_Fun_Getsignpar(6," & rsTmp("执行科室ID") & ") as tag from dual "
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验签名")
7             If rsTmp.RecordCount > 0 Then
8                 funExeDeptID = rsTmp("tag")
9             End If
10        Else
11            funExeDeptID = 0
12        End If



13        Exit Function
funExeDeptID_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funExeDeptID)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear

End Function

Public Function ReviseDate(ByVal strDate As String) As String
'功能：将时间转化为统一的24小时制时间
    If strDate = "" Then
        ReviseDate = ""
    Else
        ReviseDate = Format(strDate, "yyyy-mm-dd hh:mm:ss")
    End If
End Function

Public Function GetBabyInfor(ByVal lngPatientID As Long, ByVal lngPatientPage As Long, ByVal intBaby As Integer) As Recordset
      '功能：传入母亲的病人ID 主页ID 以及婴儿 返回孩子记录集
          Dim strSQL As String
          
1         On Error GoTo GetBabyInfor_Error

2         strSQL = "Select t.婴儿姓名, t.婴儿性别, t.分娩次数, t.分娩方式, t.胎儿状况,t.出生时间," & vbNewLine & _
                  "Nvl(Round(Nvl(t.死亡时间, Sysdate) - t.出生时间), 0) ||'天' As 年龄,t.序号 婴儿序号" & vbNewLine & _
                  "From 病人新生儿记录 t" & vbNewLine & _
                  "Where t.病人id = [1] And t.主页id = [2] And t.序号 =[3]"
3         Set GetBabyInfor = ComOpenSQL(Sel_His_DB, strSQL, "检验申请", lngPatientID, lngPatientPage, intBaby)


4         Exit Function
GetBabyInfor_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetBabyInfor)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
6         Err.Clear
          
End Function

Public Function funGetLabNewReportList(ByVal lngPatientID As Long, ByVal lngMainID As Long, ByRef strXMLNewLIS As String, Optional lngApplyID As Long) As Boolean
          '功能               LIS的公共部件，调用生成XML格式的病人的检验报告列表
          '参数
          '                   lngPatientID            消息头
          '                   lngPatientID            消息内容
          '                   strXMLNewLIS            返回的字串
          '                   lngApplyID              申请id
          '返回               True=成功   False=失败
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim objXML As Object
          Dim i As Long

1         On Error GoTo funGetLabNewReportList_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "Select b.id 检验报告id,a.申请id ,a.紧急 紧急标志,c.名称 检验项目,b.标本序号,b.微生物 是否微生物,0 报告结果,b.检验人,b.审核人,b.审核时间,a.申请时间 " & vbNewLine & _
                   "   from 检验申请组合 a,检验报告记录 b,检验组合项目 c where a.标本id= b.id and a.组合id = c.id  and  a.病人id = [1] and a.主页id =[2] and a.申请id is not null"
                   
4         If lngApplyID > 0 Then
5             strSQL = strSQL & " and 申请id =[3]"
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人检验报告列表", lngPatientID, lngMainID, lngApplyID)
8         If rsTmp.RecordCount > 0 Then
9             With objXML
10                .ClearXmlText
11                .AppendNode "检验报告列表" ', True '父节点[检验报告列表]
12                For i = 1 To rsTmp.RecordCount
13                    .AppendData "检验报告id", rsTmp!检验报告id '<检验报告id>类型：
14                    .AppendData "申请id", rsTmp!申请id '<申请id>类型：
15                    .AppendData "紧急标志", rsTmp!紧急标志 & "" '<紧急标志>类型：
16                    .AppendData "检验项目", rsTmp!检验项目 & ""  '<检验项目>类型：
17                    .AppendData "标本序号", rsTmp!标本序号 '<标本序号>类型：
18                    .AppendData "是否微生物", rsTmp!是否微生物 & ""  '<微生物标本>类型：
19                    .AppendData "报告结果", Val(rsTmp!报告结果 & "") '<报告结果>类型：
20                    .AppendData "检验人", rsTmp!检验人 & "" '<检验人>类型：
21                    .AppendData "审核人", rsTmp!审核人 & ""  '<审核人>类型：
22                    .AppendData "审核时间", rsTmp!审核时间 & "" '<审核时间>类型：
23                    .AppendData "申请时间", rsTmp!申请时间 & ""  '<申请时间>类型：
24                    rsTmp.MoveNext
25                Next
26                .AppendNode "检验报告列表", True
27                If strXMLNewLIS = "" Then strXMLNewLIS = .XmlText
28            End With
29        End If
30        funGetLabNewReportList = True


31        Exit Function
funGetLabNewReportList_Error:
32        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetLabNewReportList)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
33        Err.Clear

End Function

Public Function funGetLabNewReportResultList(ByVal lngRepottID As Long, ByRef strXMLOldLIS As String) As Boolean
          '功能：                 LIS的公共部件，提取病人的检验报告结果
          '参数
          'lngRepottID            报告id
          'strXMLOldLIS           返回的字串
          '返回                   XML格式的字串
          Dim strSQL As String
          Dim rsNewTmp As ADODB.Recordset
          Dim objXML As Object
          Dim strBH As String
          Dim i As Long

          '查新版
1         On Error GoTo funGetLabNewReportResultList_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "select  id,微生物 from 检验报告记录 where id = [1]"
4         Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人检验结果列表", lngRepottID)
5         If rsNewTmp.RecordCount > 0 Then
6             If Val(rsNewTmp("微生物") & "") = 1 Then
7                 strSQL = "Select distinct a.细菌id, b.中文名 细菌名, a.培养描述 描述, a.耐药机制, e.中文名 抗生素, c.结果 抗生素结果, c.结果类型 耐药性, c.药敏方法, e.用法用量1, e.用法用量2, e.血药浓度1," & vbNewLine & _
                           "          e.血药浓度2 , e.尿药浓度1, e.尿药浓度2" & vbNewLine & _
                           "   From 检验报告细菌 A, 检验细菌记录 B, 检验报告药敏 C, 检验药敏 E" & vbNewLine & _
                           "   Where a.细菌id = b.Id And b.Id = c.细菌id And c.药敏id = e.Id and a.标本id=[1] order by b.中文名"
                  
8                 Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人检验结果列表", lngRepottID)
9                 If rsNewTmp.RecordCount > 0 Then
10                    With objXML
11                        .ClearXmlText
12                        .AppendNode "微生物项目" ', True '父节点[普通项目]
13                        For i = 1 To rsNewTmp.RecordCount
14                            If strBH <> rsNewTmp!细菌名 & "" Then
15                                If strBH <> "" Then
16                                    .AppendNode "抗生素结果列表", True
17                                End If
18                                strBH = rsNewTmp!细菌名 & ""
19                                .AppendData "细菌id", rsNewTmp!细菌id & "" '<细菌id>类型：
20                                .AppendData "细菌名", rsNewTmp!细菌名 & "" '<细菌名>类型：
21                                .AppendData "描述", rsNewTmp!描述 & "" '<描述>类型：
22                                .AppendData "耐药机制", rsNewTmp!耐药机制 & ""  '<耐药机制>类型：
23                                .AppendNode "抗生素结果列表"  ', True '父节点[指标内容]
24                            End If
                          
25                            .AppendData "抗生素", rsNewTmp!抗生素 & "" '<抗生素>类型：
26                            .AppendData "抗生素结果", rsNewTmp!抗生素结果 & "" '<抗生素结果>类型：
27                            .AppendData "耐药性", rsNewTmp!耐药性 & "" '<耐药性>类型：
28                            .AppendData "药敏方法", rsNewTmp!药敏方法 & ""  '<药敏方法>类型：
29                            .AppendData "用法用量1", rsNewTmp!用法用量1 & "" '<用法用量1>类型：
30                            .AppendData "用法用量2", rsNewTmp!用法用量2 & ""  '<用法用量2>类型：
31                            .AppendData "血药浓度1", rsNewTmp!血药浓度1 & "" '< 血药浓度1 > 类型:
32                            .AppendData "血药浓度2", rsNewTmp!血药浓度2 & "" '<血药浓度2>类型：
33                            .AppendData "尿药浓度1", rsNewTmp!尿药浓度1 & ""  '<尿药浓度1>类型：
34                            .AppendData "尿药浓度2", rsNewTmp!尿药浓度2 & ""  '<尿药浓度2>类型：
35                            rsNewTmp.MoveNext
36                        Next
37                        .AppendNode "抗生素结果列表", True
38                        .AppendNode "微生物项目", True
39                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
40                    End With
41                End If
42            Else
43                strSQL = "Select a.项目id 指标id, b.指标代码 指标代码, b.英文名 指标英文名, b.中文名 指标中文名," & vbNewLine & _
                           " a.检验结果, a.结果标志, a.结果参考, a.排列序号, b.隐私项目,a.单位 " & vbNewLine & _
                           "   From 检验报告明细 A, 检验指标 B" & vbNewLine & _
                           "   Where a.项目id = b.Id and a.标本id =[1] "
44                Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人检验结果列表", lngRepottID)
45                If rsNewTmp.RecordCount > 0 Then
46                    With objXML
47                        .ClearXmlText
48                        .AppendNode "普通项目" ', True '父节点[普通项目]
49                        .AppendNode "指标内容" ', True '父节点[指标内容]
50                        For i = 1 To rsNewTmp.RecordCount
51                            .AppendData "指标id", rsNewTmp!指标id & "" '<指标id>类型：
52                            .AppendData "指标代码", rsNewTmp!指标代码 & "" '<指标代码>类型：
53                            .AppendData "指标英文名", rsNewTmp!指标英文名 & "" '<指标英文名>类型：
54                            .AppendData "指标中文名", rsNewTmp!指标中文名 & ""  '<指标中文名>类型：
55                            .AppendData "检验结果", rsNewTmp!检验结果 & ""  '<检验结果>类型：
56                            .AppendData "结果标志", rsNewTmp!结果标志 & ""  '<结果标志>类型：
57                            .AppendData "结果参考", rsNewTmp!结果参考 & "" '< 结果参考 > 类型:
58                            .AppendData "排列序号", rsNewTmp!排列序号 & "" '<排列序号>类型：
59                            .AppendData "隐私项目", rsNewTmp!隐私项目 & ""  '<隐私项目>类型：
60                            .AppendData "单位", rsNewTmp!单位 & ""          '<单位>类型：字符
61                            rsNewTmp.MoveNext
62                        Next
63                        .AppendNode "指标内容", True
64                        .AppendNode "普通项目", True
65                        If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
66                    End With
67                End If
68            End If
69        End If
70        funGetLabNewReportResultList = True


71        Exit Function
funGetLabNewReportResultList_Error:
72        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetLabNewReportResultList)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
73        Err.Clear

End Function

Public Function funGetNewBloodBankItem(ByVal lngApplyID As Long, ByRef strXMLOldLIS As String) As Boolean
          '功能：                 LIS的公共部件，提取病人的检验报告结果
          '参数
          'lngApplyID            相关id ,医嘱id
          '返回                   XML格式的字串
          
          Dim strSQL As String
          Dim rsNewTmp As ADODB.Recordset
          Dim objXML As Object
          Dim i As Long
          
          '查新版
1         On Error GoTo funGetNewBloodBankItem_Error
          
2         Set objXML = CreateObject("zl9ComLib.clsXML")
          
3         strSQL = "Select a.项目id 指标id, b.指标代码 指标代码, b.英文名 指标英文名, b.中文名 指标中文名, a.检验结果, a.结果标志, a.结果参考" & vbNewLine & _
                   "   From 检验报告明细 A, 检验指标 B,检验报告记录 c,检验申请组合 d " & vbNewLine & _
                   "   Where a.项目id = b.Id and a.标本id = c.id  and  c.id = d.标本id  and d.申请id =[1] "
4         Set rsNewTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人检验结果列表", lngApplyID)
5         If rsNewTmp.RecordCount > 0 Then
6             With objXML
7                 .ClearXmlText
8                 .AppendNode "普通项目" ', True '父节点[普通项目]
9                 .AppendNode "指标内容" ', True '父节点[指标内容]
10                For i = 1 To rsNewTmp.RecordCount
11                    .AppendData "指标id", rsNewTmp!指标id & "" '<指标id>类型：
12                    .AppendData "指标代码", rsNewTmp!指标代码 & "" '<指标代码>类型：
13                    .AppendData "指标英文名", rsNewTmp!指标英文名 & "" '<指标英文名>类型：
14                    .AppendData "指标中文名", rsNewTmp!指标中文名 & ""  '<指标中文名>类型：
15                    .AppendData "检验结果", rsNewTmp!检验结果 & ""  '<检验结果>类型：
16                    .AppendData "结果标志", rsNewTmp!结果标志 & ""  '<结果标志>类型：
17                    .AppendData "结果参考", rsNewTmp!结果参考 & "" '<结果参考> 类型:
18                    rsNewTmp.MoveNext
19                Next
20                .AppendNode "指标内容", True
21                .AppendNode "普通项目", True
22                If strXMLOldLIS = "" Then strXMLOldLIS = .XmlText
23            End With
24        End If
25        funGetNewBloodBankItem = True


26        Exit Function
funGetNewBloodBankItem_Error:
27        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetNewBloodBankItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
28        Err.Clear

End Function


Public Function funGetNewTransFusionApplyFor(strItemCodeing As String, lngPatientID As Long, intPatientType As Integer, lngHomePageID As Long, Optional strRegistrationBill As String, _
                                             Optional intBaby As Integer, Optional intType As Integer, Optional ByVal intDay As Integer) As String
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '功能                   LIS的公共部件，根据传入医嘱ID返回结果
      '参数
      '                       strItemCodeing 诊疗项目编码（可传入多个，使用逗号分隔）
      '                       lngPatientID 病人ID
      '                       intPatientType 病人来源 1-门诊，2-住院
      '                       lngHomePageID 主页ID （病人来源=2时查询)
      '                       lngRegistrationBill 挂号单NO（病人来源<>2时查询本次就诊）
      '                       intBaby           是否婴儿
      '                       intType           那种方式，1=再此查7天内的。0 = 不查询 2=指定查询天数（intDay参数）其他= 暂定
      '                       intDay            当intType=2时，此参数才有效，指定要查询多少天的数据
      '标本组成格式
      '                   指标1<split1>诊疗编码1<split1>单位1<split1>隐私项目1<split1>指标代码1<split1>中文名1<split1>英文名1<split1>取值序列1<split1>
      '                       检验结果1<split2>结果标志1<split2>结果参数1<split2>排列序号1<split2>标本类型1<split3>
      '                   指标2<split1>诊疗编码2<split1>单位2<split1>隐私项目2<split1>指标代码2<split1>中文名2<split1>英文名2<split1>取值序列2<split1>
      '                       检验结果2<split2>结果标志2<split2>结果参数2<split2>排列序号2<split2>标本类型2<split3>
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim rsTmpRuest As New ADODB.Recordset
          Dim strItemcodeOne As String
          Dim strSampleOne As String, i As Integer
          Dim strSampleTwo As String
          Dim varItemCodeing As Variant
          Dim strBH As String, strGetSeques As String
          Dim strStartTime As String
          Dim strEndTime As String

1         On Error GoTo funGetNewTransFusionApplyFor_Error

2         strEndTime = Format(Currentdate, "yyyy-mm-dd 23:59:59")
3         If intType = 1 Then
4             strStartTime = Format(Currentdate - 7, "yyyy-mm-dd 00:00:00")
5         ElseIf intType = 2 Then
6             strStartTime = Format(Currentdate - intDay, "yyyy-mm-dd 00:00:00")
7         End If
          '分隔的常量
          Const conSplit1 As String = "<split1>"                        '用于分隔标本,使用“<split1>”分隔
          Const conSplit2 As String = "<split2>"                        '用于分隔标本信息,使用“<split2>”分隔
          Const conSplit3 As String = "<split3>"                        '用于分隔标本指标信息,使用“<split3>”分隔
          Const conSplit4 As String = "<split4>"                        '用于分隔指标内信息,使用“<split4>”分隔

          '--------------------------------------------------------------------------------------------------------------------------------------------------------------
8         varItemCodeing = Split(strItemCodeing, ",")
9         For i = LBound(varItemCodeing) To UBound(varItemCodeing)
10            strItemcodeOne = varItemCodeing(i)
11            If gUserInfo.NodeNo <> "-" Then
12                strSQL = "Select distinct d.id 指标id, d.中文名 || '(' || d.英文名 || ')' 指标, d.单位, 0 隐私项目, d.指标代码, d.中文名, d.英文名, f.取值序列" & vbNewLine & _
                         "   From 检验组合指标 A, 检验指标 D, 检验组合项目 E, 检验仪器指标 F" & vbNewLine & _
                         "   Where a.项目id = d.Id And d.Id = f.项目id And a.组合id = e.Id And e.诊疗编码 = [1] and (e.站点=[2] or e.站点 is null)" & vbNewLine & _
                         "   order by d.id "
13            Else
14                strSQL = "Select distinct d.id 指标id, d.中文名 || '(' || d.英文名 || ')' 指标, d.单位, 0 隐私项目, d.指标代码, d.中文名, d.英文名, f.取值序列" & vbNewLine & _
                         "   From 检验组合指标 A, 检验指标 D, 检验组合项目 E, 检验仪器指标 F" & vbNewLine & _
                         "   Where a.项目id = d.Id And d.Id = f.项目id And a.组合id = e.Id And e.诊疗编码 = [1]" & vbNewLine & _
                         "   order by d.id "
15            End If
16            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", strItemcodeOne, gUserInfo.NodeNo)
17            strBH = "***"
18            Do Until rsTmp.EOF

19                If strBH <> rsTmp("指标id") Then
20                    If strBH <> "***" Then
21                        If strSampleTwo = "空" Then
22                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
23                        Else
24                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
25                        End If
26                        strSampleTwo = ""
27                    End If
                      '                If strBH = "***" Then
28                    strSampleOne = strSampleOne & conSplit3 & rsTmp("指标") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("单位") & _
                                     conSplit1 & rsTmp("隐私项目") & conSplit1 & rsTmp("指标代码") & _
                                     conSplit1 & rsTmp("中文名") & conSplit1 & rsTmp("英文名") & conSplit1
                      '                Else
                      '                     strSampleOne = strSampleOne & strGetSeques & conSplit1 & conSplit3 & rsTmp("指标") & conSplit1 & strItemcodeOne & conSplit1 & rsTmp("单位") & _
                                            '                                    conSplit1 & rsTmp("隐私项目") & conSplit1 & rsTmp("指标代码") & _
                                            '                                    conSplit1 & rsTmp("中文名") & conSplit1 & rsTmp("英文名") & conSplit1
                      '                End If
29                    strSQL = " Select *" & vbNewLine & _
                             "    From (Select b.审核时间, c.检验结果, Decode(c.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑', '') 结果标志, c.结果参考, c.排列序号," & vbNewLine & _
                             "                  B.标本类型" & vbNewLine & _
                             "           From 检验申请组合 A, 检验报告记录 B, 检验报告明细 C, 检验指标 D" & vbNewLine & _
                             "           Where a.标本id = b.Id And b.Id = c.标本id And c.项目id = d.Id And Nvl(b.微生物, 0) <> 1 And a.组合id = c.组合id And" & vbNewLine & _
                             "                  b.审核时间 Is Not Null [条件]  and d.Id =[6] and b.病人来源=[5] order by b.审核时间 desc ) E" & vbNewLine & _
                             "    Where Rownum = 1"
30                    If intPatientType = 2 Then
31                        If intBaby <> 0 Then
32                            strSQL = Replace(strSQL, "[条件]", " and b.HIS病人ID = [1] and b.主页id=[2]   and  a.婴儿 =[7] ")
33                        Else
34                            strSQL = Replace(strSQL, "[条件]", " and b.HIS病人ID = [1] and b.主页id=[2]   and nvl(a.婴儿,0)= 0 ")
35                        End If
36                        Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatientID, lngHomePageID, strRegistrationBill, strItemcodeOne, intPatientType, Val(rsTmp("指标id") & ""), intBaby)
37                        If rsTmpRuest.RecordCount > 0 Then
38                            strSampleTwo = rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型")
39                        Else
40                            If intType = 1 Or intType = 2 Then
41                                strSQL = " Select *" & vbNewLine & _
                                         "    From (Select b.审核时间, c.检验结果, Decode(c.结果标志, 1, '', 2, '↓', 3, '↑', 4, '异常', 5, '↓↓', 6, '↑↑', '') 结果标志, c.结果参考, c.排列序号," & vbNewLine & _
                                         "                  B.标本类型" & vbNewLine & _
                                         "           From 检验申请组合 A, 检验报告记录 B, 检验报告明细 C, 检验指标 D" & vbNewLine & _
                                         "           Where a.标本id = b.Id And b.Id = c.标本id And c.项目id = d.Id And Nvl(b.微生物, 0) <> 1 And a.组合id = c.组合id And" & vbNewLine & _
                                         "                  b.审核时间 Is Not Null [条件]  and d.Id =[5] order by b.审核时间 desc ) E" & vbNewLine & _
                                         "    Where Rownum = 1"
42                                If intBaby <> 0 Then
43                                    strSQL = Replace(strSQL, "[条件]", " and b.HIS病人ID = [1] and b.审核时间 between [2] and [3]  and a.婴儿=[6] ")
44                                Else
45                                    strSQL = Replace(strSQL, "[条件]", " and b.HIS病人ID = [1] and b.审核时间 between [2] and [3]   and nvl(a.婴儿,0)= 0 ")
46                                End If
47                                Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatientID, CDate(strStartTime), CDate(strEndTime), strItemcodeOne, Val(rsTmp("指标id") & ""), intBaby)
48                                If rsTmpRuest.RecordCount > 0 Then
49                                    strSampleTwo = rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型")
50                                Else
51                                    strSampleTwo = "空"
52                                End If
53                            Else
54                                strSampleTwo = "空"
55                            End If
56                        End If
57                    Else
58                        strSQL = Replace(strSQL, "[条件]", " and b.HIS病人ID = [1] and  b.挂号单=[3] ")
59                        Set rsTmpRuest = ComOpenSQL(Sel_Lis_DB, strSQL, "读取结果", lngPatientID, lngHomePageID, strRegistrationBill, strItemcodeOne, intPatientType, Val(rsTmp("指标id") & ""))
60                        If rsTmpRuest.RecordCount > 0 Then
61                            strSampleTwo = rsTmpRuest("检验结果") & conSplit2 & rsTmpRuest("结果标志") & conSplit2 & rsTmpRuest("结果参考") & conSplit2 & rsTmpRuest("排列序号") & conSplit2 & rsTmpRuest("标本类型")
62                        Else
63                            strSampleTwo = "空"
64                        End If
65                    End If
66                    strBH = rsTmp("指标id")
67                    strGetSeques = rsTmp("取值序列") & ""
68                Else
69                    If strGetSeques <> "" Then
70                        strGetSeques = GetSameString(strGetSeques & "," & rsTmp("取值序列"))
71                        If strSampleTwo = "空" Then
72                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
73                        Else
74                            strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
75                        End If
76                        strSampleTwo = ""
77                        strGetSeques = ""
78                    End If
79                End If
80                rsTmp.MoveNext
81            Loop
82            If strSampleTwo = "空" Then
83                strSampleOne = strSampleOne & strGetSeques & conSplit1 & ""
84                strSampleTwo = ""
85                strGetSeques = ""
86            Else
87                strSampleOne = strSampleOne & strGetSeques & conSplit1 & strSampleTwo
88                strSampleTwo = ""
89                strGetSeques = ""
90            End If

91        Next
          '------------------------------------------------------------------------------------------------------------------------
92        If strSampleOne <> "" Then
93            strSampleOne = Mid(strSampleOne, Len(conSplit1) + 1)
94        End If
95        funGetNewTransFusionApplyFor = strSampleOne


96        Exit Function
funGetNewTransFusionApplyFor_Error:
97        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetNewTransFusionApplyFor)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
98        Err.Clear

End Function

Public Function GetSameString(ByVal strGetSeque As String) As String
          '检查重复取值序列
          Dim i As Integer
          Dim varGetSeque As Variant
          Dim strGetSeques As String

1         On Error GoTo GetSameString_Error

2         varGetSeque = Split(strGetSeque, ",")
3         For i = LBound(varGetSeque) To UBound(varGetSeque)
4             If varGetSeque(i) <> "" Then
5                 strGetSeques = "," & varGetSeque(i) & ","
6             Else
7                 strGetSeques = ""
8             End If
9             If Left(strGetSeque, 1) <> "," Then strGetSeque = "," & strGetSeque
10            If Right(strGetSeque, 1) <> "," Then strGetSeque = strGetSeque & ","
11            If InStr(GetSameString, strGetSeques) = 0 Then
12                GetSameString = GetSameString & varGetSeque(i) & ","
13            End If
14        Next
15        If GetSameString <> "" Then GetSameString = Mid(GetSameString, 1, Len(GetSameString) - 1)


16        Exit Function
GetSameString_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetSameString)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear

End Function

Public Function funFindAdvicePay(ByVal strAdvice As String, ByVal intPaentType As Integer, Optional ByVal strErr As String = "") As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim varAdvice As Variant
          Dim intloop As Integer
          Dim strAdvivePay As String
          Dim blnNewPait As Boolean       '病人是否为新门诊病人
          Dim intNewPaitPay As Integer    '新门诊病人是否收费

1         On Error GoTo funFindAdvicePay_Error

2         varAdvice = Split(strAdvice, ",")
3         For intloop = 0 To UBound(varAdvice)

              '检查新门诊病人是否收费
4             If gcnHisOracle.State = 1 Then
5                 If VerCompare(gSysInfo.VersionHIS, "10.35.100") <> -1 Then
6                     blnNewPait = funNewSystemSvr(Val(varAdvice(intloop)))
7                 End If
8             End If
9             If blnNewPait Then  '新门诊病人
                  '不处理
10                Exit Function
11            Else

                  '非新门诊病人
12                If GetAdviceFeeKind(Val(varAdvice(intloop))) = 2 Then
13                    strSQL = "Select 医嘱序号,记录状态,记录性质,实收金额 from 住院费用记录 t where  t.医嘱序号 in  (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)))"
14                Else
15                    strSQL = "Select 医嘱序号,记录状态,记录性质,实收金额 from 门诊费用记录 t where  t.医嘱序号 in  (Select * From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist))) "
16                End If
17                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查询标本ID", varAdvice(intloop))
18                rsTmp.Filter = " 记录性质= 2 and 记录状态=1"
19                If rsTmp.RecordCount > 0 Then
20                    If Val(rsTmp("实收金额")) = 0 Then
21                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",-1|"
22                    Else
23                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",3|"
24                    End If
25                Else
26                    rsTmp.Filter = "记录状态=0"
27                    If rsTmp.RecordCount > 0 Then
28                        strAdvivePay = strAdvivePay & varAdvice(intloop) & ",0|"
29                    Else
30                        rsTmp.Filter = "记录状态=1"
31                        If rsTmp.RecordCount > 0 Then
32                            If Val(rsTmp("实收金额")) = 0 Then
33                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",-1|"
34                            Else
35                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",1|"
36                            End If
37                        Else
38                            rsTmp.Filter = "记录状态=2 or 记录状态=3"
39                            If rsTmp.RecordCount > 0 Then
40                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",2|"
41                            Else
42                                strAdvivePay = strAdvivePay & varAdvice(intloop) & ",0|"
43                            End If
44                        End If
45                    End If
46                End If
47            End If
48        Next
49        If strAdvivePay <> "" Then funFindAdvicePay = Mid(strAdvivePay, 1, Len(strAdvivePay) - 1)

50        Exit Function
funFindAdvicePay_Error:
51        strErr = "函数funFindAdvicePay出错：" & Err.Number & " " & Err.Description
52        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funFindAdvicePay)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
53        Err.Clear
End Function

Private Function GetAdviceFeeKind(lngAdviceID As Long) As Byte
      '功能：根据医嘱ID获取临嘱发送的费用单据的性质，1=门诊费用，2=住院费用
          Dim rsTmp As ADODB.Recordset, strSQL As String

1         On Error GoTo GetAdviceFeeKind_Error


2         GetAdviceFeeKind = 2
3         strSQL = "Select a.记录性质,a.门诊记帐,b.病人来源 From 病人医嘱发送 a,病人医嘱记录 b Where a.医嘱ID = [1] and a.医嘱id= b.id"

4         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查询标本ID", lngAdviceID)
5         If rsTmp.RecordCount > 0 Then
6             If rsTmp!记录性质 = 1 Or rsTmp!记录性质 = 2 And Val("" & rsTmp!门诊记帐) = 1 Then
7                 GetAdviceFeeKind = 1
8             Else
9                 If Val("" & rsTmp!病人来源) = 4 Then
10                    GetAdviceFeeKind = 1
11                End If
12            End If
13        End If

14        Exit Function



15        Exit Function
GetAdviceFeeKind_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetAdviceFeeKind)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear

End Function

Public Function funGetNewDataToXK(ByVal lngPatientID As Long, ByVal strItemCode As String) As String
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                  根据传入病人ID，指标代码 返回当前病人最近一次的检验结果
          '参数
          '                       lngPatientID 病人id
          '                       strItemCode 指标代码串，使用 ，号分隔
          '                      返回格式：   病人id<A>指标代码1<S>检验结果1<B>指标代码2<S>检验结果2<B>指标代码3<S>检验结果3

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTest As ADODB.Recordset
          Dim strErr As String, intloop As Integer
          Dim varItemCode As Variant
          Dim strData As String

1         On Error GoTo funGetNewDataToXK_Error

2         varItemCode = Split(strItemCode, ",")
3         For intloop = 0 To UBound(varItemCode)
4             strSQL = "Select *" & vbNewLine & _
                       "   From (Select d.指标代码, c.检验结果, b.审核时间" & vbNewLine & _
                       "          From 检验报告记录 B, 检验报告明细 C, 检验指标 D" & vbNewLine & _
                       "          Where b.Id = c.标本id And c.项目id = d.Id And b.His病人id = [1]  And d.指标代码 = [2] And b.审核人 Is Not Null " & vbNewLine & _
                       "          Order By b.审核时间 Desc)" & vbNewLine & _
                       "   Where Rownum < 2"
5             Set rsTest = ComOpenSQL(Sel_Lis_DB, strSQL, "查询检验结果", lngPatientID, varItemCode(intloop))
6             If rsTest.RecordCount > 0 Then
7                 strData = strData & "<B>" & rsTest("指标代码") & "<S>" & rsTest("检验结果")
8             End If
              
9         Next
10        If strData <> "" Then strData = Mid(strData, 4)
11        funGetNewDataToXK = lngPatientID & "<A>" & strData

12        Exit Function
funGetNewDataToXK_Error:
13        strErr = "funGetNewData出错：" & Err.Number & " " & Err.Description
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetNewDataToXK)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
15        Err.Clear
End Function

Public Function funGetReadNotify(objFrm As Object, strAdvice As String, ByVal strDicName As String, Optional ByRef strReturn As String) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                   医生工作站反馈危急值处理措施到LIS
          '参数
          '       入参
          '                       objfrm          调用窗体
          '                       strAdvice       医嘱id串
          '                       strDicName      医生姓名
          '       出参
          '                       strReturn       返回医生填写的处理措施
          '返回                   True=成功,False=失败
          '
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

          Dim lngSampleID As Long, strTime As String

1         On Error GoTo funGetReadNotify_Error

2         If strAdvice <> "" Then
3             strSQL = "select b.id,b.姓名,b.标本序号 from 检验申请组合 a,检验报告记录 b " & vbNewLine & _
                          "where a.标本id = b.id  and a.申请id =[1]"

4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "危急值查询", CLng(strAdvice))
5             If rsTmp.EOF = False Then
6                 strTime = Format(Currentdate, "yyyy-mm-dd hh:mm:ss")
7                 lngSampleID = Val(rsTmp("id") & "")
8                 funGetReadNotify = frmAppforCritical.ShowMe(objFrm, strDicName, lngSampleID, strReturn)
9             End If
              
10        End If
          


11        Exit Function
funGetReadNotify_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetReadNotify)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear
          
End Function

Public Function funGetLISPatientRecord(ByVal lngPatientID As Long, ByVal strRecordID As String, ByVal intType As Integer, Optional strErr As String) As ADODB.Recordset
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                       查询LIS病人信息
          '
          '入参
          '                           lngPatientID    病人ID
          '                           lngRecordID     intType=1 挂号单 intType=2 主页ID intType=3 体检
          '                           intType         1=门诊 2=住院 3=体验
          '                           strErr          返回的错误信息
          '返回
          '                           病人信息记录集,记录集中包含以下病人信息
          '                           医嘱ID,相关ID,申请时间,病人来源,婴儿,姓名,性别,年龄,年龄数字,年龄单位,申请人, 申请科室,床号,病人科室,紧急,挂号单,门诊号,住院号,主页ID,标本类型,病区 from 检验申请组合
          '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo funGetLISPatientRecord_Error

2         strSQL = "select 申请ID as 医嘱ID,医嘱ID as 相关ID,申请时间,病人来源,婴儿,姓名,性别,年龄,年龄数字,年龄单位,申请人," & _
                 " 申请科室,床号,病人科室,紧急,挂号单,门诊号,住院号,主页ID,标本类型,病区 from 检验申请组合 where HIS病人ID=[1]"
          
3         Select Case intType
              Case 1
4                 strSQL = strSQL & " and 挂号单=[2] and 病人来源=1"
5             Case 2
6                 strSQL = strSQL & " and 主页ID=[2] and 病人来源=2"
7             Case 3
8                 strSQL = strSQL & " and 挂号单=[2] and 病人来源=4"
9         End Select
          
10        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "查询LIS病人", lngPatientID, strRecordID)
11        Set funGetLISPatientRecord = rsTmp

12        Exit Function
funGetLISPatientRecord_Error:
13        strErr = Err.Number & "   " & Err.Description
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetLISPatientRecord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
15        Err.Clear
End Function

Public Function funModifyPatientBaseIntoLIS(ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal int场合 As Integer, ByVal strName As String, _
                                        ByVal strSex As String, ByVal lngAgeNum As Long, ByVal strAgeUnit As String, ByVal strEditMode As String, _
                                        ByVal strEditUser As String, Optional strErr As String) As Boolean

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能                       同步修改到ZLLIS病人信息中
          '
          '入参
          '                           lng病人ID
          '                           str就诊ID       场合=1 挂号单 场合=2 主页ID 场合=3 体检
          '                           int场合         1=门诊 2=住院 3=体验
          '                           strName         要修改的病人姓名
          '                           strSex          要修改的病人性别
          '                           strAgeNum       要修改的病人年龄数字
          '                           strAgeUnit      要修改的病人年龄单位
          '
          '                           strEditMode     修改源于哪个模块
          '                           strEditUser     修改人
          '                           strErr          返回的错误信息
          '返回
          '                           True=保存成功 False=保存失败
          '
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strAgeAll As String    '年龄
          Dim intBaby As Integer  '婴儿
          
          Dim strAgeUnit1 As String    '第一年龄单位
          Dim strAgeUnit2 As String   '第二年龄单位
          Dim strAge1 As String        '第一年龄
          Dim strAge2 As String       '第二年龄
          Dim strInfo As String

          
1         On Error GoTo funModifyPatientBaseIntoLIS_Error

2         funModifyPatientBaseIntoLIS = False

3         Set rsTmp = funGetLISPatientRecord(lng病人ID, str就诊ID, int场合, strErr)
          
          '没有查找到信息是退出
4         If rsTmp.RecordCount < 1 Then
5             funModifyPatientBaseIntoLIS = True
6             Exit Function
7         Else
8             intBaby = Val(rsTmp("婴儿") & "")
9         End If
          
          '性别
10        Select Case strSex
              Case "男", 1
11                strSex = "1"
12            Case "女", 2
13                strSex = "2"
14            Case "未知", 9
15                strSex = "9"
16            Case Else
17                strSex = "0"
18        End Select
          
          '年龄
19        strAgeAll = lngAgeNum & strAgeUnit
20        If InStr(strAgeAll, "岁") > 0 Then
21            strAgeUnit1 = "岁"
22            strAgeUnit2 = "月"
23        ElseIf InStr(strAgeAll, "月") > 0 Then
24            strAgeUnit1 = "月"
25            strAgeUnit2 = "天"
26        ElseIf InStr(strAgeAll, "天") > 0 Then
27            strAgeUnit1 = "天"
28            strAgeUnit2 = "时"
29        ElseIf InStr(strAgeAll, "时") > 0 Then
30            strAgeUnit1 = "时"
31            strAgeUnit2 = "分"
32        ElseIf InStr(strAgeAll, "成") > 0 Then
33            strAgeUnit1 = "成"
34            strAgeUnit2 = ""
35        ElseIf InStr(strAgeAll, "婴") > 0 Then
36            strAgeUnit1 = "婴"
37            strAgeUnit2 = ""
38        ElseIf InStr(strAgeAll, "分") > 0 Then
39            strAgeUnit1 = "分"
40            strAgeUnit2 = ""
41        End If
42        strAgeAll = Replace(strAgeAll, "个月", "月")
43        strAgeAll = Replace(strAgeAll, "小时", "时")
44        strAgeAll = Replace(strAgeAll, "分钟", "分")
          
45        strAge1 = Mid(strAgeAll, 1, InStr(strAgeAll, strAgeUnit1) - 1)
46        strAge2 = Mid(strAgeAll, InStr(strAgeAll, strAgeUnit1) + 1)
47        strAge2 = Replace(strAge2, strAgeUnit2, "")
48        strInfo = CheckAgeInfo(strAge1, strAgeUnit1, strAge2, strAgeUnit2, strAgeAll)
49        If strInfo <> "" Then
50            strErr = strInfo
51            funModifyPatientBaseIntoLIS = False
52            Exit Function
53        End If
            
          
54        strSQL = "Zl_Lis病人信息_调整(" & lng病人ID & ",'" & str就诊ID & "'," & int场合 & ",'" & strName & "','" & strSex & _
                 "','" & strAgeAll & "'," & lngAgeNum & ",'" & strAgeUnit & "','" & strEditMode & "','" & strEditUser & "'," & intBaby & ")"
55        Call ComExecuteProc(Sel_Lis_DB, strSQL, "修改病人信息")
          
56        strErr = "LIS病人信息已修改,请通知检验科重新审核病人报告"
57        funModifyPatientBaseIntoLIS = True



58        Exit Function
funModifyPatientBaseIntoLIS_Error:
59        strErr = Err.Description
60        If InStr(strErr, "[ZLSOFT]") > 0 Then
61            strErr = Mid(strErr, InStr(strErr, "[ZLSOFT]") + 8, InStrRev(strErr, "[") - InStr(strErr, "[ZLSOFT]") - 8)
62        End If
63        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funModifyPatientBaseIntoLIS)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
64        Err.Clear
End Function

Public Function CheckAgeInfo(strAge As String, strAgeUnit As String, strAge1 As String, strAgeUnit1 As String, Optional ByVal strFullAge As String) As String
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '功能               检验年龄是否合格要求
          '参数
          '                   strAge  = 年龄（2）
          '                   strAgeUnit = 年龄单位 (岁)
          '                   strAge1 = 第二年龄（3）
          '                   strAgeUnit1 = 第二年龄单位(月）
          '                   strfullAge  =完整的年龄字符串
          '返回               正确返加为空字串，错误返回来具体出错提示内容
          '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Dim strTmp As String

          '判断年龄字符串第一个字符是否为数字
1         On Error GoTo CheckAgeInfo_Error

2         If InStr("0123456789", Mid(strFullAge, 1, 1)) <= 0 And strFullAge <> "" Then
3             CheckAgeInfo = "年龄不符合要求,年龄数字不能为非数字！"
4             Exit Function
5         End If
          
          '判断第一年龄
6         If IsNumeric(strAge) = False And Val(Trim(strAge)) <> 0 Then
7             CheckAgeInfo = "年龄不符合要求，不是全数字！"
8             Exit Function
9         End If
          '判断年龄大小
10        If Val(strAge) > 150 And strAgeUnit = "岁" Then
11            CheckAgeInfo = "年龄不能超过150岁！"
12            Exit Function
13        End If

          '判断年龄单位
14        If Trim(strAgeUnit) = "" And Val(Trim(strAge)) <> 0 Then
              '当有年龄数字时，年龄单位不能为空
15            CheckAgeInfo = "当有年龄数字时，年龄单位不能为空！"
16            Exit Function
17        End If

18        If InStr(",岁,月,天,时,成,婴,分,", "," & strAgeUnit & ",") <= 0 And Val(Trim(strAge)) <> 0 Then
19            CheckAgeInfo = "年龄单位不符合要求，请检查！"
20            Exit Function
21        End If

         '判断第二年龄是否为数字
22        If InStr("0123456789", Mid(strFullAge, Len(strAge) + Len(strAgeUnit) + 1, 1)) <= 0 Then
23             CheckAgeInfo = "第二年龄非数字，请检查！"
24             Exit Function
25        End If
          
26        If IsNumeric(strAge1) = False And Val(Trim(strAge1)) <> 0 Then
27            CheckAgeInfo = "年龄不符合要求，不是全数字！"
28            Exit Function
29        End If
         
          
30        If Trim(strAgeUnit) <> "" Then
31            Select Case strAgeUnit
                  Case "岁"
32                    strTmp = "月"
33                Case "月"
34                    strTmp = "天"
35                Case "天"
36                    strTmp = "时"
37                Case "时"
38                    strTmp = "分"
39                Case Else
40                    strTmp = ""
41            End Select
42            If strTmp <> strAgeUnit1 And strAgeUnit1 <> "" Then
43                CheckAgeInfo = "第二年龄单位不符，请检查！"
44                Exit Function
45            End If
46        End If
          
47        CheckAgeInfo = ""


48        Exit Function
CheckAgeInfo_Error:
49        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(CheckAgeInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
50        Err.Clear
End Function



Public Function funModifyBabyInfo(ByVal lngBID As Long, ByVal lngZID As Long, ByVal strNO As String, ByVal lngBabyID As Long, ByVal strBabyName As String, ByVal strBabySex As String, Optional strErr As String) As Boolean
          '功能        修改新生儿基本信息
          'lngBID      病人id
          'lngZID      主页id
          'strNO       挂号单
          'lngBabyID   婴儿序号
          'strBabyName 婴儿姓名
          'strBabyAge  婴儿性别
          '返回        True=修改成功,Flase=修改失败
          
          Dim strSQL As String

1         On Error GoTo funModifyBabyInfo_Error

2         funModifyBabyInfo = False
          
3         strSQL = "Zl_Lis新生儿信息_Update(" & lngBID & "," & lngZID & ",'" & strNO & "'," & lngBabyID & ",'" & strBabyName & "','" & strBabySex & "')"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "修改新生儿信息")
          
5         funModifyBabyInfo = True
          

6         Exit Function
funModifyBabyInfo_Error:
7         strErr = "出错函数(funModifyBabyInfo),出错信息:" & Err.Number & " " & Err.Description
8         funModifyBabyInfo = False
9         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funModifyBabyInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
10        Err.Clear
End Function

'---------------------------------------------------------------------------------------
' 编    码:蔡青松
' 编码日期:2017-3-23
' 功    能:写入检验通知记录
' 入    参:
'           intType                 1=检验科拒收标本，2=临床传染病处理反馈
'           strAdviceID             检验医嘱ID，多个使用","分割
'           [可选][intSampleType    标本状态    采集工作站调用时传入 1=让步检验，2=取消检验，3=样本重采]
' 出    参:
'          strErr                  错误信息
' 返    回:True=成功，False=失败
' 修 改 人:
' 修改日期:
'---------------------------------------------------------------------------------------
Public Function funWriteInLisNotify(ByVal intType As Integer, ByVal strAdviceID As String, _
                                    Optional ByVal intSampleType As Integer, Optional strErr As String) As Boolean
    Dim strSQL As String
    Dim var_tmp As Variant
    Dim strNotify As String
    Dim strBusiness As String
    Dim lngLoop As Long
        
    '如果传入的医嘱第一位为逗号，则截取掉第一位
    On Error GoTo funWriteInLisNotify_Error

    If Mid(strAdviceID, 1, 1) = "," Then
        strAdviceID = Mid(strAdviceID, 2)
    End If
    If strAdviceID = "" Then Exit Function
    
    If intType = 1 Then
        If intSampleType = 1 Then
            strNotify = "需执行让步检验"
        ElseIf intSampleType = 2 Then
            strNotify = "取消检验"
        ElseIf intSampleType = 3 Then
            strNotify = "样本已重采"
        End If
        strBusiness = "标本拒收"
    ElseIf intType = 2 Then
        strNotify = "医生已反馈传染病处理情况"
        strBusiness = "传染病"
    End If
    
    var_tmp = Split(strAdviceID, ",")
    For lngLoop = LBound(var_tmp) To UBound(var_tmp)
        strSQL = "Zl_检验消息记录_Edit(1,3, Null, Null," & var_tmp(lngLoop) & ",Null,Null,'" & strNotify & "','" & strBusiness & "')"
        Call ComExecuteProc(Sel_Lis_DB, strSQL, "检验消息记录")
    Next
    funWriteInLisNotify = True
    
    Exit Function
funWriteInLisNotify_Error:
    strErr = Err.Description
    Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funWriteInLisNotify)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
    Err.Clear
    
End Function

Public Function PrintReportNew(objFrm As Object, lngAdive As Long, lngPaint As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '功能       打印报告
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset
          Dim lngSampleID As Long
          
1         On Error GoTo PrintReportNew_Error

2         strSQL = "select  b.id from 检验申请组合 a ,检验报告记录  b  where  a.标本id = b.id and   a.医嘱id = [1] and a.his病人id= [2]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngAdive, lngPaint)
4         If rsTmp.RecordCount > 0 Then
5             lngSampleID = Val(rsTmp("ID") & "")
6         Else
7             strSQL = "select  b.id from 检验申请组合 a ,检验报告记录  b  where  a.标本id = b.id and   a.申请id = [1] and a.his病人id= [2]"
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngAdive, lngPaint)
9             If rsTmp.RecordCount > 0 Then
10                lngSampleID = Val(rsTmp("ID") & "")
11            Else
12                PrintReportNew = False
13                Exit Function
14            End If
15        End If

16        strSQL = "select b.id 仪器id ,b.名称 仪器名称,b.仪器类别,a.病人来源,a.报告时间,a.阳性报告,a.标本序号 from 检验报告记录 a,检验仪器记录 b where a.仪器id = b.id and a.id = [1]"
17        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngSampleID)

18        If rsTmp.RecordCount = 0 Then Exit Function

19        strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
                      "from 检验仪器记录 where id = [1] "

20        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))


21        rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
22        If Val(rsTmp("仪器类别")) = 1 Then
23            If Val(rsTmp("阳性报告") & "") = 1 Then
                  '阳性
24                intSel = 0
25            Else
                  '阴性
26                intSel = 1
27            End If
28        Else
29            intCount = GetSampleValCount(lngSampleID)
              '没有结果时提示
30            If intCount = 0 Then
31                Exit Function
32            End If
33            If rsReportFormat.RecordCount > 0 Then
34                If Val(rsReportFormat("格式数量") & "") > 0 Then
35                    If intCount > Val(rsReportFormat("格式数量") & "") Then
36                        intSel = 0
37                    Else
38                        intSel = 1
39                    End If
40                End If
41            Else
42                intSel = 0
43            End If

44        End If
45        Select Case Val(rsTmp("病人来源") & "")
              Case 1
46                If intSel = 0 Then
47                    strNO = rsReportFormat("门诊单据号")
48                Else
49                    strNO = rsReportFormat("门诊格式号")
50                End If
51            Case 2
52                If intSel = 0 Then
53                    strNO = rsReportFormat("住院单据号")
54                Else
55                    strNO = rsReportFormat("住院格式号")
56                End If
57            Case 3
58                If intSel = 0 Then
59                    strNO = rsReportFormat("住院单据号")
60                Else
61                    strNO = rsReportFormat("住院格式号")
62                End If
63            Case 4
64                If intSel = 0 Then
65                    strNO = rsReportFormat("院外单据号")
66                Else
67                    strNO = rsReportFormat("院外格式号")
68                End If
69            Case Else
70                If intSel = 0 Then
71                    strNO = rsReportFormat("门诊单据号")
72                Else
73                    strNO = rsReportFormat("门诊格式号")
74                End If
75        End Select
76        If byRunMode = 3 Then
77            If strNO <> "" Then
78                 FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
79            End If
80        Else
             '读图像
81            strTmp = "开始读入图像:" & Now & vbCrLf
82            If ReadSampleImage(lngSampleID, strChart, strErr) = False Then
83                MsgBox strErr: Exit Function
84            End If
85            strTmp = strTmp & "读入图像完成:" & Now & vbCrLf
          
86            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                      "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                      "图形9=" & strChart(8), byRunMode
87            strTmp = strTmp & "打印完成:" & Now & vbCrLf
              
              '对于审核过的标本标识
88            strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ")"
89            Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
90            strTmp = strTmp & "完成打印:" & Now
          
91            SaveDBLog 18, 6, lngSampleID, "打印", "报告打印", 2500, "临床实验室管理"
92        End If

93        PrintReportNew = True


94        Exit Function
PrintReportNew_Error:
95        PrintReportNew = False
96        strErr = "执行(PrintReportNew)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
97        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(PrintReportNew)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
98        Err.Clear
End Function

Public Function GetSampleValCount(lngSampleID As Long, Optional strErr As String) As Integer
          '功能           取当前标本的总数数

          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo GetSampleValCount_Error

2         strSQL = "select count(*) count from 检验报告明细 where 标本id = [1] and 检验结果 is not null "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID)
4         GetSampleValCount = rsTmp("count")

5         Exit Function
GetSampleValCount_Error:
6         strErr = WriteErrLog("zl9LisInsideComm", "mdlLisHisComm", "执行(GetSampleValCount)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear

End Function

Public Function ReadSampleImage(lngSampleID As Long, strChar() As String, Optional strErr As String, Optional intVal As Integer = 25) As Boolean
    '功能   读入标本的图像返回读出的数组
    '读图像
    Dim strReturn As String
    Dim varTmp As Variant, strDir As String
    Dim i As Integer
    Dim gobjFSO As New Scripting.FileSystemObject    'FSO对象
    Dim objImg As Object

    On Error GoTo ReadSampleImage_Error

    strErr = ""
    strDir = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strDir) Then Call gobjFSO.CreateFolder(strDir)

    If objImg Is Nothing Then
        Set objImg = CreateObject("zlLisDev.clsDrawGraph")



        If strErr <> "" Then
            MsgBox strErr
            Exit Function
        End If
    End If
    objImg.GetSampleImgExit strErr
    '标本ID
    '图片保存路径(不存在则自动创建),
    '是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
    '函数返回值为空串时，返回的提示信息
    '返回的图片文件格式，0－cht(默认),1-jgp,2-png
    '是新版LIS还是老版LIS在调用本函件数， 0-老版LIS（默认，从“检验图像结果”中取图形数据），1-新版LIS（从“检验报告图像”中取图形数据）
    If intVal = 25 Then
         Call objImg.GetSampleImgInit(gSysInfo.SysNo, gcnLisOracle, strErr)
        strReturn = objImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 1)
    Else
         Call objImg.GetSampleImgInit(gSysInfo.SysNo, gcnHisOracle, strErr)
        strReturn = objImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 0)
    End If
    If strReturn = "" Then
        If strErr = "无图像数据！" Then
            strErr = ""
            ReadSampleImage = True
        ElseIf strErr <> "" Then
            MsgBox strErr, vbQuestion
        Else
            ReadSampleImage = True
        End If
        Exit Function
    End If

    varTmp = Split(strReturn, ",")

    For i = LBound(varTmp) To UBound(varTmp)
        If i > 8 Then Exit For
        If Trim("" & varTmp(i)) <> "" Then
            If Dir(strDir & "\" & Trim("" & varTmp(i))) <> "" Then strChar(i) = strDir & "\" & Trim("" & varTmp(i))
        End If
    Next

    ReadSampleImage = True

    Exit Function
ReadSampleImage_Error:
    strErr = "出错函数(ReadSampleImage),出错信息:" & Err.Number & " " & Err.Description
    Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(ReadSampleImage)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
    Err.Clear
End Function



'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/5/9
'功    能:向lis消息部件发送刷新质控该概况列表消息
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub SendMessage(ByVal strMessage As String)
1         On Error GoTo SendMessage_Error
          
2         If mstrPara = "" Then mstrPara = ComGetPara(Sel_Lis_DB, "LIS远程通讯参数", 2500, 2500, "")
3         If mstrPara = "" Then Exit Sub
4         If gobjPublicLIS Is Nothing Then
5             Set gobjPublicLIS = CreateObject("zlPublicLIS.clsSampleReprot")
6             If Not gobjPublicLIS Is Nothing Then Call gobjPublicLIS.Init(mstrPara)
7         End If
          
8         If Not gobjPublicLIS Is Nothing Then
9             Call gobjPublicLIS.SendMessage(strMessage, mstrPara)
10        End If


11        Exit Sub
SendMessage_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(SendMessage)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear

End Sub







'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/8/25
'功    能:  查询病人住院期间所有已出报告的申请ID，并且通过标本进行分组
'入    参:
'           lngPatientID    HIS病人ID
'           intPage         主页ID
'出    参:
'           strErr          错误信息
'返    回:  返回按照标本进行分组的医嘱ID串。医嘱之间用","分割，标本之间用";"分割
'---------------------------------------------------------------------------------------
Public Function funGetPatientAdvice(ByVal lngPatientID As Long, ByVal intPage As Integer, Optional ByRef strErr As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strAdvice As String
          
1         On Error GoTo funGetPatientAdvice_Error
          
2         strErr = ""
          
3         strSQL = "select f_List2str(Cast(Collect(a.申请ID || '') As t_Strlist)) 申请ID" & vbCrLf & _
                  " from  检验申请组合 A,检验报告记录 B " & vbCrLf & _
                  " where a.标本id=b.id and a.his病人ID=[1] and a.主页ID=[2] and b.审核人 is not null group by a.标本ID"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngPatientID, intPage)
5         If rsTmp.RecordCount <= 0 Then Exit Function
6         Do While Not rsTmp.EOF
7             strAdvice = strAdvice & ";" & rsTmp("申请ID")
8             rsTmp.MoveNext
9         Loop
10        If Mid(strAdvice, 1, 1) = ";" Then strAdvice = Mid(strAdvice, 2)
11        funGetPatientAdvice = strAdvice

12        Exit Function
funGetPatientAdvice_Error:
13        strErr = Err.Description & "(" & Err.Number & ")"
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetPatientAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
15        Err.Clear
End Function




'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/11/22
'功    能:回写危急值处理措施到LIS系统中
'入    参:
'           lngSampleID     标本ID
'           strUserName     处理人员
'           strNotify       处理措施
'出    参:
'           [strErr         错误信息]
'返    回:  回写成功返回True,否则返回False
'---------------------------------------------------------------------------------------
Public Function funWriteNotifyToLis(ByVal lngSampleID As Long, ByVal strUserName As String, ByVal strNotify As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim strTime As String
          Dim blnTrs As Boolean
          
1         On Error GoTo funWriteNotifyToLis_Error
          
2         strTime = Currentdate
          
3         gcnLisOracle.BeginTrans
4         blnTrs = True
           '检验消息提醒
5         strSQL = "Zl_检验消息记录_Edit(1,2,null," & lngSampleID & ",null,null,null,'医生已查阅危急值','危急值')"
6         Call ComExecuteProc(Sel_Lis_DB, strSQL, "检验消息记录")
          
7         strSQL = "Zl_检验危急值记录_Message(" & lngSampleID & ",'" & strUserName & "',to_date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),'" & strNotify & "')"
8         Call ComExecuteProc(Sel_Lis_DB, strSQL, "危急值记录通知")
9         gcnLisOracle.CommitTrans
10        blnTrs = False
          
11        SaveDBLog 18, 6, Val(lngSampleID), "危急值处理", "确认处理，确认人：" & strUserName & " 确认时间:" & Format(strTime, "yyyy-MM-dd HH:mm:ss") & " 处理措施:" & strNotify, 2500, "临床实验室管理"
          
12        funWriteNotifyToLis = True

13        Exit Function
funWriteNotifyToLis_Error:
14        strErr = Err.Description
15        If blnTrs Then gcnLisOracle.RollbackTrans
16        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funWriteNotifyToLis)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
17        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/14
'功    能:  判断是否允许核收新门诊发送的医嘱
'入    参:
'           strAdvicIDs     医嘱ID，多个医嘱使用“,”分割
'出    参:
'返    回:  True=新门诊病人，False=非新门诊病人
'---------------------------------------------------------------------------------------
Public Function funNewSystemSvr(ByVal strAdvicIDs As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strGHNo As String
          Dim astrtmp() As String
          Dim strJsOut As String          '返回的JSON
          Dim intFinish As Integer
          Dim i As Integer
          Dim strErr As String
          Dim blnTmp As Boolean

          '    json说明
          '    字段       名称        说明
          '    result     执行结果    1-成功；-1-失败 Number(1)   非空
          '    errmsg     错误消息    失败时返回错误消息  Varchar2(200)
          '    kacnt_sign 收费状态    0-未收费，1-已收费；该标志表示执行科室是否可执行：绿色通道返回1，日间病房预交金足够返回1，账单已收费返回1，其它返回0    Number(1)   非空
          '    kacnt_chrg 未收金额    绿色通道返回预交金不足部分金额；其它返回未收费金额  Number(18,2)    非空


          '查询住院记录
1         On Error GoTo funNewSystemSvr_Error

2         strSQL = "select /*+cardinality(c,10)*/  b.附加标志" & vbCrLf & _
                   " from 病人医嘱记录 A,病人挂号记录 B,Table(f_Num2list([1])) C " & vbCrLf & _
                   " where a.挂号单=b.no and a.id=c.Column_Value" & vbCrLf & _
                   "union all " & vbCrLf & _
                   "select /*+cardinality(c,10)*/ b.附加标志" & vbCrLf & _
                   " from 病人医嘱记录 A,病人挂号记录 B,Table(f_Num2list([1])) C " & vbCrLf & _
                   " where a.挂号单=b.no and a.相关id=c.Column_Value"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "病案主页", strAdvicIDs)
4         Do While Not rsTmp.EOF
5             If Val(rsTmp("附加标志") & "") = 3 Then    '3表示新门诊病人

                  '由于新门诊病人在新门诊采集站已经严格控制了收费状态，所以LIS系统中不在检查，新门诊病人一律不处理
6                 funNewSystemSvr = True
7                 Exit Function

8             End If
9             rsTmp.MoveNext
10        Loop

11        Exit Function
funNewSystemSvr_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funNewSystemSvr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear

End Function

Public Function funGetSampleType(ByVal strAdvice As String, Optional strErr As String) As ADODB.Recordset
      '功能       获取标本状态

      '参数
      '           strAdvice       医嘱ID,多个医嘱用","分割
      '           strErr          返回错误信息

      '返回记录集
      '记录集字段: "医嘱ID", adBigInt
      '           "医嘱状态", adVarChar, 20
      '           "操作员", adVarChar, 20
      '           "操作时间", adDate

          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset    '医嘱内容
          Dim strType As String           '标本状态
          Dim strReturn As String         '返回结果
1         Dim strAdviceID As String       '医嘱ID串      格式:老版医嘱ID1,老版医嘱ID2,,|新版医嘱ID1,新版医嘱ID2,,,
          Dim strOldAdvice As String      '老版医嘱ID
          Dim strNewAdvice As String      '新版医嘱ID
          Dim rsReture As ADODB.Recordset    '返回的记录集
          Dim strUser As String           '操作员
          Dim strDate As String           '操作时间
          Dim intType As Integer
          Dim var_tmp As Variant
          Dim intloop As Integer
          Dim strArr As Variant
          Dim i As Integer

          '初试化本地记录集
2         On Error GoTo funGetSampleType_Error

3         Set rsReture = InitRecord

4         strArr = TruncatedExtraLongStr(strAdvice, ",")

5         For i = 0 To UBound(strArr)
6             strOldAdvice = ""
7             strNewAdvice = ""
8             strAdviceID = ""
              
              '进入检验科之前流程
9             strSQL = "select /*+cardinality(c,10)*/ distinct a.医嘱ID 申请ID ,e.执行状态,a.采样人,a.采样时间,a.送检人," & _
                     " a.标本送出时间 送检时间,a.接收人,a.接收时间,b.核收人,b.核收时间,'' 拒收人,'' 拒收时间," & _
                     " b.审核人,b.审核时间 from 病人医嘱发送 A,检验标本记录 B,Table(f_num2list([1])) C,病人医嘱记录 D,病人医嘱发送 E " & _
                     " where a.医嘱id=b.医嘱id(+) and a.医嘱id=d.相关id and d.id=e.医嘱id And a.医嘱ID=c.Column_Value"
10            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "老板医嘱", strArr(i))

              '获取医嘱状态
11            Do While Not rsTmp.EOF
12                strType = ""
13                strUser = ""
14                strDate = ""
15                If IsNull(rsTmp("采样时间")) And Val(rsTmp("执行状态") & "") <> 2 Then  '未采样
16                    strType = "未采样"
17                ElseIf Val(rsTmp("执行状态") & "") = 2 Then     '已拒收
18                    strType = "已拒收"
19                    strUser = rsTmp("拒收人") & ""
20                    strDate = rsTmp("拒收时间") & ""
21                End If

22                If Val(rsTmp("执行状态") & "") <> 2 Then
23                    If Not IsNull(rsTmp("采样时间")) Then    '已采样
24                        strType = "已采样"
25                        strUser = rsTmp("采样人") & ""
26                        strDate = rsTmp("采样时间") & ""
27                    End If
28                    If Not IsNull(rsTmp("送检时间")) Then    '已送检
29                        strType = "已送检"
30                        strUser = rsTmp("送检人") & ""
31                        strDate = rsTmp("送检时间") & ""
32                    End If
33                    If Not IsNull(rsTmp("接收人")) Then   '已接收
34                        strType = "已接收"
35                        strUser = rsTmp("接收人") & ""
36                        strDate = rsTmp("接收时间") & ""
37                    End If
38                End If

                  '添加到本地记录集
39                rsReture.AddNew
40                rsReture("医嘱ID") = CLng(rsTmp("申请ID") & "")
41                rsReture("医嘱状态") = strType
42                If strUser <> "" Then
43                    rsReture("操作员") = strUser
44                End If
45                If strDate <> "" Then
46                    rsReture("操作时间") = CDate(Format(strDate, "yyyy-mm-dd hh:mm:ss"))
47                End If
48                rsTmp.MoveNext
49            Loop


              '进入检验科之后流程
              '查询老版医嘱ID
50            strSQL = "select /*+cardinality(b,10)*/ distinct 医嘱ID from 检验项目分布 A,Table(f_num2list([1])) B where a.医嘱id=b.column_value"
51            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "老板医嘱", strArr(i))
52            Do While Not rsTmp.EOF
53                strOldAdvice = strOldAdvice & "," & rsTmp("医嘱ID")
54                rsTmp.MoveNext
55            Loop
56            If strOldAdvice <> "" Then strOldAdvice = Mid(strOldAdvice, 2)

              '查询新版医嘱ID
57            strSQL = " select /*+cardinality(b,10)*/ distinct 申请ID from 检验申请组合 A,Table(f_num2list([1])) B where a.申请id=b.column_value"
58            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "新版医嘱", strArr(i))
59            Do While Not rsTmp.EOF
60                strNewAdvice = strNewAdvice & "," & rsTmp("申请ID")
61                rsTmp.MoveNext
62            Loop
63            If strNewAdvice <> "" Then strNewAdvice = Mid(strNewAdvice, 2)

64            strAdviceID = strOldAdvice & "|" & strNewAdvice

65            If strAdviceID <> "" Then var_tmp = Split(strAdviceID, "|")
66            For intloop = LBound(var_tmp) To UBound(var_tmp)
67                If intloop = 0 And var_tmp(0) <> "" Then
                      '查询老版医嘱信息
68                    intType = 10
69                    strSQL = "select /*+cardinality(c,10)*/ distinct a.医嘱ID 申请ID ,e.执行状态,a.采样人,a.采样时间,a.送检人," & _
                             " a.标本送出时间 送检时间,a.接收人,a.接收时间,b.核收人,b.核收时间,'' 拒收人,'' 拒收时间," & _
                             " b.审核人,b.审核时间 from 病人医嘱发送 A,检验标本记录 B,Table(f_num2list([1])) C,病人医嘱记录 D,病人医嘱发送 E " & _
                             " where a.医嘱id=b.医嘱id(+) and a.医嘱id=d.相关id and d.id=e.医嘱id And a.医嘱ID=c.Column_Value"

70                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "老板医嘱", var_tmp(intloop))
71                ElseIf intloop = 1 And var_tmp(1) <> "" Then
                      '查询新版医嘱信息
72                    intType = 25
73                    strSQL = "Select /*+cardinality(c,10)*/ distinct a.申请ID,'0' 执行状态,a.标本ID ,a.申请时间, a.采样时间, a.采样人, a.送检时间," & _
                             " a.送检人, a.接收时间, a.接收人, b.核收时间, b.检验人 核收人, b.检验人," & _
                             " b.报告时间, b.审核人, b.审核时间 ,a.拒收人,a.拒收时间,a.门诊号, a.住院号, a.样本条码" & _
                             " from 检验申请组合 A, 检验报告记录 B,Table(f_num2list([1])) C" & _
                               "　Where a.标本id = b.Id(+) And a.申请id=c.Column_Value"

74                    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "新版医嘱", var_tmp(intloop))
75                End If


                  '获取医嘱状态
76                Do While Not rsTmp.EOF
77                    strType = ""
78                    strUser = ""
79                    strDate = ""
80                    If Not IsNull(rsTmp("核收时间")) Then   '已核收
81                        If intType = 25 Then
82                            strType = "已核收"
83                            strUser = rsTmp("核收人") & ""
84                            strDate = rsTmp("核收时间") & ""
85                        ElseIf intType = 10 Then
86                            If Val(rsTmp("执行状态") & "") <> 2 Then
87                                strType = "已核收"
88                                strUser = rsTmp("核收人") & ""
89                                strDate = rsTmp("核收时间") & ""
90                            End If
91                        End If
92                    End If
93                    If Not IsNull(rsTmp("审核时间")) Then   '已审核
94                        If intType = 25 Then
95                            strType = "已审核"
96                            strUser = rsTmp("审核人") & ""
97                            strDate = rsTmp("审核时间") & ""
98                        ElseIf intType = 10 Then
99                            If Val(rsTmp("执行状态") & "") <> 2 Then
100                               strType = "已审核"
101                               strUser = rsTmp("审核人") & ""
102                               strDate = rsTmp("审核时间") & ""
103                           End If
104                       End If
105                   End If


                      '更新当前医嘱对应的检验科内部状态
106                   rsReture.Filter = "医嘱ID=" & CLng(rsTmp("申请ID") & "")
107                   If rsReture.RecordCount > 0 Then
108                       If strType <> "" Then
109                           rsReture("医嘱状态") = strType
110                       End If
111                       If strUser <> "" Then
112                           rsReture("操作员") = strUser
113                       End If
114                       If strDate <> "" Then
115                           rsReture("操作时间") = CDate(Format(strDate, "yyyy-mm-dd hh:mm:ss"))
116                       End If
117                   End If
118                   rsTmp.MoveNext
119               Loop
120           Next
121       Next

122       rsReture.Filter = ""
123       rsReture.MoveFirst
124       Set funGetSampleType = rsReture


125       Exit Function
funGetSampleType_Error:
126       strErr = "执行(BeforCreateLisValueStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
127       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetSampleType)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
128       Err.Clear

End Function

Private Function InitRecord(Optional strErr As String) As ADODB.Recordset
          '初始化本地记录集
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo InitRecord_Error

2         If rsTmp.State = adStateOpen Then rsTmp.Close
3         rsTmp.Fields.Append "医嘱ID", adBigInt
4         rsTmp.Fields.Append "医嘱状态", adVarChar, 20
5         rsTmp.Fields.Append "操作员", adVarChar, 20
6         rsTmp.Fields.Append "操作时间", adDate

7         rsTmp.CursorLocation = adUseClient
8         rsTmp.LockType = adLockOptimistic
9         rsTmp.CursorType = adOpenStatic
10        If rsTmp.State = adStateClosed Then rsTmp.Open

11        Set InitRecord = rsTmp


12        Exit Function
InitRecord_Error:
13        strErr = "执行(BeforCreateLisValueStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
14        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(InitRecord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
15        Err.Clear
          
End Function


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/18
'功    能:检查当前时间是否是业务高峰期，并且业务指定的查询范围是否在许可范围内
'入    参:
'      lngSysNo=系统号
'      lngModuleNo=模块号
'      strFuncName=功能名称
'      datBegin=功能进行查询数据范围的开始时间，当类型为数值类型时
'      datEnd=功能进行查询数据范围的截至时间
'      lngDays=查询的时间范围,当为0时，通过datBegin与datEnd反算，不为0时忽略datBegin与datEnd
'出    参:
'返    回:是否可以进行操作
'---------------------------------------------------------------------------------------
Public Function funCheckRushHours(ByVal lngSysNo As Long, ByVal lngModuleNo As Long, ByVal strFuncName As String, _
                                Optional ByVal datBegin As Date, Optional ByVal datEnd As Date, Optional ByVal lngDays As Long) As Boolean
'    If gzlSystem Is Nothing Then
        funCheckRushHours = True
'        Exit Function
'    End If
'    funCheckRushHours = gzlSystem.CheckRushHours(lngSysNO, lngModuleNo, strFuncName, datBegin, datEnd, lngDays)
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/4/25
'功    能:将拒收信息写入新版LIS
'入    参:
'           strAdviceID         申请ID 多个申请ID使用","分割
'           strUser             拒收人
'           strRefuseInfo       拒收理由
'           strRegName          拒收接收人
'           strRegTime          拒收接收时间

'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function funRefuseSampleInNew(ByVal strAdviceID As String, ByVal strUser As String, ByVal strRefuseInfo As String, _
                                     Optional ByVal strRegName As String, Optional ByVal strRegTime As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strIDs As String

1         On Error GoTo funRefuseSampleInNew_Error

2         strSQL = "select /*+cardinality(b,10)*/ ID from 检验申请组合 where 申请ID  In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B)"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", strAdviceID)
4         Do While Not rsTmp.EOF
5             strIDs = strIDs & "," & rsTmp("ID")
6             rsTmp.MoveNext
7         Loop
8         If strIDs <> "" Then strIDs = Mid(strIDs, 2)

9         If VerCompare(gSysInfo.VersionLIS, "10.35.120") = 0 Or VerCompare(gSysInfo.VersionLIS, "10.35.150") <> -1 Then
10            strSQL = "Zl_检验报告拒收_Edit('" & strIDs & "','" & strUser & "',null,null,'" & strRefuseInfo & "',0,null," & IIf(strRegName = "", "null", "'" & strRegName & "'") & "," & IIf(strRegTime <> "", "to_date('" & strRegTime & "','yyyy-mm-dd hh24:mi:ss')", "null") & ")"
11        Else
12            strSQL = "Zl_检验报告拒收_Edit('" & strIDs & "','" & strUser & "',null,null,'" & strRefuseInfo & "',0,null)"
13        End If
14        Call ComExecuteProc(Sel_Lis_DB, strSQL, "检验申请组合")

15        funRefuseSampleInNew = True


16        Exit Function
funRefuseSampleInNew_Error:
17        strErr = "执行(funRefuseSampleInNew)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
18        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funRefuseSampleInNew)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
19        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/5
'功    能:获取新版LIS中的标本类型
'入    参:
'           strInfo     不查询重复的标本类型
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function funGetSampleTypeNew() As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo funGetSampleTypeNew_Error

2         If intLis_Setup <> 1 Then Exit Function
          
3         strSQL = "select 编码,名称 from 检验标本类型"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "标本类型")
5         Set funGetSampleTypeNew = rsTmp


6         Exit Function
funGetSampleTypeNew_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetSampleTypeNew)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
8         Err.Clear
          
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/6
'功    能:通过医嘱ID打印LIS报告
'入    参:
'           objFrm          调用窗体
'           lngAdvice       医嘱ID（采集医嘱ID)
'           byRunMode       1=打印预览，2=打印，3=打印设置，4=打印PDF
'           BlnlimitPrint   打印新版报告时，是否受到打印次数参数的限制（超出新版LIS中打印次数的报告无法打印）
'           strPDF          需要打印的PDF文件的文件路径
'           strPrinter      指定打印机的名称，若果指定了打印机名称，则默认在指定的打印机上打印
'出    参:
'           strErr          返回错误或打印失败原因
'返    回:  是否打印成功    True=成功，False=失败
'---------------------------------------------------------------------------------------
Public Function funPrintLisReport(ByVal objFrm As Object, ByVal lngAdvice As String, ByVal byRunMode As Byte, _
                                  Optional ByVal BlnLimitPrint As Boolean, Optional ByVal strPDF As String, _
                                  Optional ByVal strPrinter As String, Optional ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset
          Dim blnNewReport As Boolean
          Dim blnOldReport As Boolean
          Dim lngPrintCount As Long
          Dim lngSampleID As Long
          Dim lngPaintID As Long
          Dim intSel As Integer
          Dim intCount As Integer
          Dim strNO As String
          Dim strTmp As String
          Dim strChart(0 To 8) As String
          Dim strReportCode As String
          Dim strReportParaNo As String
          Dim bytReportParaMode As Byte
          Dim lng医嘱ID As Long
          Dim lng发送号 As Long


1         On Error GoTo funPrintLisReport_Error

          '先到新版LIS中去查询检验报告
2         strSQL = "select distinct b.id 标本ID,b.审核人,b.医生站打印,b.病人来源,b.仪器ID,c.仪器类别,b.阳性报告 " & vbCrLf & _
                 " from 检验申请组合 A,检验报告记录 B,检验仪器记录 C " & vbCrLf & _
                 " where a.标本ID=b.ID(+) and b.仪器ID=c.id and 申请ID=[1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngAdvice)
4         If rsTmp.RecordCount > 0 Then blnNewReport = True

          '若新版LIS中没有报告，则再到老版LIS中去查新
5         If Not blnNewReport Then
6             strSQL = "Select Distinct 标本ID, b.审核人, c.发送号, a.医嘱id, b.病人id" & vbCrLf & _
                     "   From 检验项目分布 A, 检验标本记录 B, 病人医嘱发送 C" & vbCrLf & _
                     "   Where a.标本ID = b.id And a.医嘱id = c.医嘱id And a.医嘱id =[1]"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验标本记录", lngAdvice)
8             If rsTmp.RecordCount > 0 Then blnOldReport = True
9         End If
          '判断报告是否已出，未出的报告禁止打印
10        Do While Not rsTmp.EOF
11            If IsNull(rsTmp("标本ID")) Or IsNull(rsTmp("审核人")) Then
12                strErr = "报告未出"
13                Exit Function
14            End If
15            rsTmp.MoveNext
16        Loop

17        If blnNewReport Or blnOldReport Then
18            rsTmp.MoveFirst
19            lngSampleID = Val(rsTmp("标本ID") & "")
20        End If
          '打印新版报告时，检查是否超出打印次数
21        If BlnLimitPrint = True And blnNewReport = True Then    '需要检查，并且是新版LIS报告
22            lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "医生工作站报告打印次数", 2500, 2500, 0))
              '对比打印次数和参数
23            If lngPrintCount > 0 Then
24                If Val(rsTmp("医生站打印") & "") >= lngPrintCount And Val(rsTmp("病人来源") & "") = 2 Then
25                    strErr = "超出打印次数"
26                    Exit Function
27                End If
28            End If
29        End If

          '打印报告
          '新版
30        If blnNewReport Then
31            strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
                     "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
                       "from 检验仪器记录 where id = [1] "

32            Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))
33            rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
34            If Val(rsTmp("仪器类别")) = 1 Then
35                If Val(rsTmp("阳性报告") & "") = 1 Then
                      '阳性
36                    intSel = 0
37                Else
                      '阴性
38                    intSel = 1
39                End If
40            Else
41                intCount = GetSampleValCount(lngSampleID)
                  '没有结果时提示
42                If intCount = 0 Then
43                    Exit Function
44                End If
45                If rsReportFormat.RecordCount > 0 Then
46                    If Val(rsReportFormat("格式数量") & "") > 0 Then
47                        If intCount > Val(rsReportFormat("格式数量") & "") Then
48                            intSel = 0
49                        Else
50                            intSel = 1
51                        End If
52                    End If
53                Else
54                    intSel = 0
55                End If

56            End If
57            Select Case Val(rsTmp("病人来源"))
              Case 1
58                If intSel = 0 Then
59                    strNO = rsReportFormat("门诊单据号")
60                Else
61                    strNO = rsReportFormat("门诊格式号")
62                End If
63            Case 2
64                If intSel = 0 Then
65                    strNO = rsReportFormat("住院单据号")
66                Else
67                    strNO = rsReportFormat("住院格式号")
68                End If
69            Case 3
70                If intSel = 0 Then
71                    strNO = rsReportFormat("住院单据号")
72                Else
73                    strNO = rsReportFormat("住院格式号")
74                End If
75            Case 4
76                If intSel = 0 Then
77                    strNO = rsReportFormat("院外单据号")
78                Else
79                    strNO = rsReportFormat("院外格式号")
80                End If
81            Case Else
82                If intSel = 0 Then
83                    strNO = rsReportFormat("门诊单据号")
84                Else
85                    strNO = rsReportFormat("门诊格式号")
86                End If
87            End Select
88            If byRunMode = 3 Then
89                If strNO <> "" Then
90                    FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
91                End If
92            Else
                  '读图像
93                strTmp = "开始读入图像:" & Now & vbCrLf
94                If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
95                    Exit Function
96                End If
97                strTmp = strTmp & "读入图像完成:" & Now & vbCrLf

98                If strPrinter <> "" Then Call FunSetReportPrintSet(gcnLisOracle, gSysInfo.SysNo, strNO, "printer", strPrinter)    '设置指定打印机
99                FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                                "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                                "图形9=" & strChart(8), "PDF=" & strPDF, byRunMode
100               strTmp = strTmp & "打印完成:" & Now & vbCrLf

                  '对于审核过的标本标识
101               strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ",1)"
102               Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
103               strTmp = strTmp & "完成打印:" & Now

104               SaveDBLog 18, 6, lngSampleID, "打印", "病案检验报告打印", 2500, "临床实验室管理"
105           End If
106       ElseIf blnOldReport Then
107           lng发送号 = Val("" & rsTmp("发送号"))
108           lng医嘱ID = Val("" & rsTmp("医嘱id"))
109           lngPaintID = Val("" & rsTmp("病人ID"))
110           If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, , strErr) Then
111               If byRunMode = 3 Then
112                   FunReportPrintSet gcnHisOracle, 100, strReportCode, objFrm
113               Else
114                   If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
115                       Exit Function
116                   End If

117                   If strPrinter <> "" Then Call FunSetReportPrintSet(gcnHisOracle, 100, strReportCode, "printer", strPrinter)  '设置指定打印机
118                   Call FunReportOpen(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                                         "病人ID=" & lngPaintID, "标本ID=" & lngSampleID, _
                                         "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                                         "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                                         "图形9=" & strChart(8), "PDF=" & strPDF, byRunMode)
119               End If
120           Else
121               Exit Function
122           End If
123       End If

124       If Not blnNewReport And Not blnOldReport Then
125           strErr = "未查询到报告"
126           Exit Function
127       End If


128       funPrintLisReport = True


129       Exit Function
funPrintLisReport_Error:
130       strErr = Err.Description
131       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funPrintLisReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
132       Err.Clear
End Function

Public Function GetReportCode(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False, Optional ByRef strErr As String) As Boolean
          '--------------------------------------------------------------------------------------------------------
          '功能;
          '--------------------------------------------------------------------------------------------------------
          Dim rs As New ADODB.Recordset
          Dim strSQL As String
          
1         On Error GoTo GetReportCode_Error

2         If lng医嘱ID = 0 And lng发送号 = 0 Then Exit Function
          
3         strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                             "A.NO," & _
                             "A.记录性质 " & _
                      "FROM 病人医嘱发送 A,病历文件列表 C,病人医嘱记录 D,病历单据应用 E " & _
                      "Where E.病历文件id = C.ID " & _
                              "AND D.诊疗项目ID=E.诊疗项目ID " & _
                            "AND A.医嘱ID=D.ID AND E.应用场合=Decode(D.病人来源,2,2,4,4,1) " & _
                            " AND D.相关id= [1] "
4         If DataMoved Then
5             strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
6             strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
7         End If

8         Set rs = ComOpenSQL(Sel_His_DB, strSQL, "报告打印", lng医嘱ID, lng发送号)
                            
          
9         If rs.BOF = False Then
10            strCode = NVL(rs("报表编号"))
11            strNO = NVL(rs("NO"))
12            bytMode = NVL(rs("记录性质"), 1)
13        End If
14        GetReportCode = True


15        Exit Function
GetReportCode_Error:
16        strErr = Err.Description
17        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetReportCode)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
18        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-03-13
'功    能:  通过诊疗编码获取当前项目是否为耐受实验项目
'入    参:
'           strItemID       诊疗项目ID
'出    参:
'           strErr          错误或者提示信息
'返    回:  True=当前项目是耐受项目，False=当前项目不是耐受项目
'调整影响:
'---------------------------------------------------------------------------------------
Public Function funIsToleranceItem(ByVal strItemID As Long, ByRef strErr As String) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String       '诊疗编码

1         On Error GoTo funIsToleranceItem_Error
          
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") = -1 Then
3             Exit Function
4         End If
          
          '通过诊疗项目ID获取诊疗项目编码
5         strSQL = "select 编码 from 诊疗项目目录 where ID=[1]"
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strItemID)
7         If rsTmp.EOF Then
      '        strErr = "没有查询到对应的诊疗项目"
8             Exit Function
9         Else
10            strItemCode = Trim(rsTmp("编码") & "")
11        End If

          '判断是否是耐受试验
12        strSQL = "select id,是否耐受项目 from 检验组合项目 where 诊疗编码=[1]"
13        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "耐受项目", strItemCode)
14        If rsTmp.RecordCount > 1 Then
15            strErr = "当前项目对照了多个检验项目，请联系检验科相关人员进行排查"
16            Exit Function
      '    ElseIf rsTmp.RecordCount = 0 Then
      '        strErr = "当前项目未与新版LIS项目进行对码"
17        ElseIf rsTmp.RecordCount = 1 Then
18            If Val(rsTmp("是否耐受项目") & "") = 1 Then
19                funIsToleranceItem = True
20            End If
21        End If


22        Exit Function
funIsToleranceItem_Error:
23        strErr = "执行(funIsToleranceItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl
24        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funIsToleranceItem)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
25        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-15
'功    能:  通过传入医嘱ID串，返回哪些医嘱是属于同一个标本
'入    参:
'           strAdvice       医嘱ID，多个医嘱ID使用英文逗号分割
'出    参:
'返    回:  医嘱ID串，不同标本之间的医嘱ID使用英文分号分割，相同标本的医嘱ID使用逗号分割
'调整影响:
'---------------------------------------------------------------------------------------
Public Function funGetSampleAdvice(ByVal strAdvice As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strOld As String
          Dim strReturn As String
          Dim strArr() As String
          Dim i As Integer

1         On Error GoTo funGetSampleAdvice_Error

2         strOld = "," & strAdvice & ","
3         strArr = TruncatedExtraLongStr(strAdvice, ",")
4         For i = 0 To UBound(strArr)
5             strOld = "," & strArr(i) & ","
              '从新版中去查找
6             strSQL = "Select /*+cardinality(b,10)*/ f_List2str(Cast(Collect(a.申请id || '') As t_Strlist)) 申请id" & vbCrLf & _
                     " From 检验申请组合 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                     " Where A.申请id = B.Column_Value and a.标本ID is not null" & vbCrLf & _
                     " Group By a.标本ID"
7             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", strArr(i))
              '剔除新版的医嘱ID之后剩下的到老版LIS中去查询
8             Do While Not rsTmp.EOF
9                 If Not IsNull(rsTmp("申请id")) Then
10                    strOld = Replace(strOld, rsTmp("申请id") & ",", "")
11                    If InStr(rsTmp("申请id") & "", ",") > 0 Then
12                        strReturn = strReturn & ";" & rsTmp("申请id")
13                    End If
14                End If
15                rsTmp.MoveNext
16            Loop
17            If strOld <> "" Then
18                If Left(strOld, 1) = "," Then strOld = Mid(strOld, 2)
19                If Right(strOld, 1) = "," Then strOld = Mid(strOld, 1, Len(strOld) - 1)
20            End If
21            If strOld <> "" Then
                  '从老版中去查找
22                strSQL = "Select /*+cardinality(b,10)*/ f_List2str(Cast(Collect(a.医嘱id || '') As t_Strlist)) 申请id" & vbCrLf & _
                         " From 检验项目分布 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                         " Where A.医嘱id = B.Column_Value" & vbCrLf & _
                         " Group By a.标本ID"
23                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验申请组合", strOld)
24                Do While Not rsTmp.EOF
25                    If Not IsNull(rsTmp("申请id")) Then
26                        If InStr(rsTmp("申请id") & "", ",") > 0 Then
27                            strReturn = strReturn & ";" & rsTmp("申请id")
28                        End If
29                    End If
30                    rsTmp.MoveNext
31                Loop
32            End If
33        Next
34        If strReturn <> "" Then
35            If Left(strReturn, 1) = ";" Then funGetSampleAdvice = Mid(strReturn, 2)
36        End If

37        Exit Function
funGetSampleAdvice_Error:
38        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetSampleAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
39        Err.Clear
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：外挂部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn 外挂部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, "ZLLIS"
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
        Call gobjPlugIn.Initialize(gcnHisOracle, gSysInfo.SysNo, gSysInfo.ModlNo, int场合)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-19
'功    能:  显示诊疗参考
'入    参:
'           objFrm          父级对象
'           lngSampleID     标本ID
'           lngVer          版本
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub funShowClincHelp(objFrm As Object, ByVal lngSampleID As Long, ByVal lngVer As Long)
          Dim objAdvice As Object
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItemCode As String
          Dim strItemIDs As String
          Dim lngPaitID As Long
          Dim lngPage As Long
          Dim intPaitType As Integer
          Dim strGHNo As String
          Dim lngGHID As Long
          Dim blnContinue As Boolean


1         On Error GoTo funShowClincHelp_Error

2         If lngSampleID <> 0 Then
              '获取诊疗项目ID
3             If lngVer = 25 Then
4                 strSQL = "Select f_List2str(Cast(Collect(b.诊疗编码 || '') As t_Strlist)) 编码" & vbCrLf & _
                           "   From 检验申请组合 A, 检验组合项目 B" & vbCrLf & _
                           "   Where A.组合ID = b.id And a.标本id = [1]"
5                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngSampleID)
6                 If Not rsTmp.EOF Then
7                     strItemCode = rsTmp("编码") & ""
8                 End If

9                 If strItemCode <> "" Then
                      '通过诊疗编码查询诊疗项目ID
10                    strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                               "f_List2str(Cast(Collect(a.ID || '') As t_Strlist)) ID" & vbCrLf & _
                               " From 诊疗项目目录 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                               " Where A.编码 = B.Column_Value"
11                    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strItemCode)
12                    If Not rsTmp.EOF Then strItemIDs = rsTmp("ID") & ""
13                End If

                  '获取病人信息
14                strSQL = "select 病人ID,病人来源,主页ID,挂号单 from 检验报告记录 where ID=[1]"
15                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验报告记录", lngSampleID)
16                If Not rsTmp.EOF Then
17                    lngPaitID = Val(rsTmp("病人ID") & "")
18                    lngPage = Val(rsTmp("主页ID") & "")
19                    intPaitType = Val(rsTmp("病人来源") & "")
20                    strGHNo = rsTmp("挂号单") & ""
21                End If
22            ElseIf lngVer = 10 Then
23                strSQL = " select f_List2str(Cast(Collect(b.诊疗项目ID || '') As t_Strlist)) 诊疗项目ID from 检验标本记录 A, 病人医嘱记录 B where a.医嘱id=b.相关id and a.id=[1]"
24                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目ID", lngSampleID)
25                If Not rsTmp.EOF Then
26                    strItemIDs = rsTmp("诊疗项目ID") & ""
27                End If

                  '获取病人信息
28                strSQL = "select 病人ID,病人来源,主页ID,挂号单 from 检验标本记录 where ID=[1]"
29                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验报告记录", lngSampleID)
30                If Not rsTmp.EOF Then
31                    lngPaitID = Val(rsTmp("病人ID") & "")
32                    lngPage = Val(rsTmp("主页ID") & "")
33                    intPaitType = Val(rsTmp("病人来源") & "")
34                    strGHNo = rsTmp("挂号单") & ""
35                End If

36            End If

              '查询挂号ID
37            If strGHNo <> "" And lngPaitID <> 0 Then
38                strSQL = "SELECT ID FROM 病人挂号记录 where no=[1] AND 病人ID=[2]"
39                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "病人挂号记录", strGHNo, lngPaitID)
40                If Not rsTmp.EOF Then
41                    lngGHID = Val(rsTmp("ID") & "")
42                End If
43            End If
44        End If

          '先调用plugin中的接口，接口调用失败再调用zlPublicAdvice中的接口
45        If VerCompare(gSysInfo.VersionHIS, "10.35.130") <> -1 Then
46            If CreatePlugInOK(2500, 2) Then
47                On Error Resume Next
48                blnContinue = gobjPlugIn.ShowClinicHelp(objFrm.hWnd, 1, intPaitType, lngPaitID, IIf(intPaitType = 2, lngPage, lngGHID), strItemIDs)
49                Call zlPlugInErrH(Err, "ExecuteFunc")
50                Err.Clear: On Error GoTo 0
51            End If
52        End If

          '调用接口
53        If Not blnContinue Then
54            If objAdvice Is Nothing Then
55                Set objAdvice = CreateObject("zlPublicAdvice.clsPublicAdvice")
56                If Not objAdvice Is Nothing Then
57                    On Error Resume Next
58                    Call objAdvice.ShowClincHelp(1, objFrm, 0, False, strItemIDs)
59                    If Err.Number = 438 Then
60                        MsgBox "HIS版本过低", vbInformation, gSysInfo.AppName
61                        Exit Sub
62                    End If
63                End If
64            End If
65        End If



66        Exit Sub
funShowClincHelp_Error:
67        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funShowClincHelp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
68        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-06-27
'功    能:  创建外挂插件按钮
'入    参:
'           objCbr          CommandBar对象
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function CreatePlugInButton(objToolBar As CommandBar) As Boolean
          Dim cbrMenuBar As CommandBarPopup
          Dim cbrControl As CommandBarControl
          Dim cbrToolBar As CommandBar
          Dim strTmp As String
          Dim arrTmp As Variant
          Dim i As Integer

          '-----------------------------------------外接插件-------------------------------------------------
          '插件扩展功能
1         On Error GoTo CreatePlugInButton_Error

2         Call CreatePlugInOK(2500, 1)
3         If Not gobjPlugIn Is Nothing Then
4             With objToolBar.Controls
5                 Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Tool_PlugIn, "扩展功能(&G)")
6                 cbrControl.BeginGroup = True
7                 cbrControl.Style = xtpButtonIconAndCaption
8                 With cbrControl.CommandBar.Controls
9                     If Not gobjPlugIn Is Nothing Then
10                        On Error Resume Next
11                        strTmp = gobjPlugIn.GetFuncNames(2500, 2500, 2)
12                        Call zlPlugInErrH(Err, "GetFuncNames")
13                        Err.Clear: On Error GoTo 0
14                    End If
15                    If strTmp <> "" Then
16                        strTmp = Replace(strTmp, "Auto:", "")
17                        arrTmp = Split(strTmp, ",")
18                        For i = 0 To UBound(arrTmp)
19                            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
20                            If i <= 9 Then cbrControl.Caption = cbrControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
21                            cbrControl.IconId = conMenu_Tool_PlugIn_Item
22                            cbrControl.Parameter = arrTmp(i)
23                        Next
24                    End If
25                End With
26            End With
27        End If
          '-----------------------------------------END-------------------------------------------------


28        Exit Function
CreatePlugInButton_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(CreatePlugInButton)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
30        Err.Clear
End Function

Public Sub ExePlugIn(ByVal strName As String, ByVal lngSampleID As Long)
'功能：执行外挂功能
    Dim lngID As String
    Dim lngPaitID As Long
    Dim lngMainID As Long
    If CreatePlugInOK(2500, 1) Then
        Call gobjPlugIn.ExecuteFunc(2500, 2500, strName, lngPaitID, lngMainID, lngID, lngSampleID, 1)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-06-27
'功    能:  获取标本补充报告
'入    参:
'           lngSampleID     标本ID
'           objVSF          展示数据的VSF
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function GetSupplementReport(ByVal lngSampleID As Long, objVSF As VSFlexGrid) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim intItem As Integer
          Dim strItem As String
          Dim lngRow As Long
          
1         On Error GoTo GetSupplementReport_Error

2         intItem = ComGetPara(Sel_Lis_DB, "检验项目显示", gSysInfo.SysNo, gSysInfo.ModlNo, "1")
          
          
3         Select Case intItem
              Case 1
4                 strItem = "c.中文名  检验项目"

5             Case 2
6                 strItem = "c.英文名  检验项目"

7             Case 3
8                 strItem = "c.中文名 || '(' || c.英文名 || ')'  检验项目"
9             End Select

10            strSQL = "Select b.id, b.补充报告ID, b.项目ID," & strItem & ", b.仪器ID, b.检验结果, b.结果标志, b.结果参考, b.参考高值, b.参考低值, b.单位" & vbCrLf & _
                      " From 检验补充报告记录 A, 检验补充报告明细 B, 检验指标 C" & vbCrLf & _
                      " Where a.ID = b.补充报告id And b.项目id = c.ID And a.标本id = [1] And b.补充结果类型 = 2" & vbCrLf & _
                      " Order By c.排列序号"
11            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "补充报告明细", lngSampleID)
12            If SetDataToVSF(objVSF, rsTmp) = False Then Exit Function
13            With objVSF
14                .SelectionMode = flexSelectionFree
15                .ColHidden(.ColIndex("ID")) = True
16                .ColHidden(.ColIndex("补充报告ID")) = True
17                .ColHidden(.ColIndex("项目ID")) = True
18                .ColHidden(.ColIndex("仪器ID")) = True
19                .ColHidden(.ColIndex("结果标志")) = True
20                .ColHidden(.ColIndex("参考高值")) = True
21                .ColHidden(.ColIndex("参考低值")) = True
                  
22                For lngRow = 1 To .Rows - 1
23                    .Cell(flexcpBackColor, lngRow, .ColIndex("检验结果"), lngRow, .ColIndex("检验结果")) = GetValColour(Val(.TextMatrix(lngRow, .ColIndex("结果标志"))))
24                Next
25            End With
              

26        Exit Function
GetSupplementReport_Error:
27        Call WriteErrLog("ZL9LabWork", "mdlWorkBaseReprot", "执行(GetSupplementReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
28        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-06-27
'功    能:  将结果列表中的作废指标改为删除线
'入    参:
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Sub EditSampleValueList(objVSFSampleValue As VSFlexGrid, objVSFSupplement As VSFlexGrid)
    Dim i As Integer
    Dim J As Integer
    
    With objVSFSampleValue
        For i = 1 To .Rows - 1
            With objVSFSupplement
                For J = 1 To .Rows - 1
                    If Val(objVSFSampleValue.TextMatrix(i, objVSFSampleValue.ColIndex("ID"))) = Val(.TextMatrix(J, .ColIndex("项目ID"))) Then
                        objVSFSampleValue.Cell(flexcpFontStrikethru, i, 0, i, objVSFSampleValue.Cols - 1) = True
                    End If
                Next
            End With
        Next
    End With
End Sub

Public Function GetValColour(intValType As Integer) As Double
    '功能               传入对应的结果类型1-正常、2-偏低、3-偏高、4-阳性(异常)、5-警示下限、6-警示上限、7-复查下限、8-复查上限
    '返回               对应的颜色
    Select Case intValType
        Case 1, 0
            GetValColour = gSampleShowColour.正常
        Case 2
            GetValColour = gSampleShowColour.偏低
        Case 3
            GetValColour = gSampleShowColour.偏高
        Case 4
            GetValColour = gSampleShowColour.异常
        Case 5
            GetValColour = gSampleShowColour.警示偏低
        Case 6
            GetValColour = gSampleShowColour.警示偏高
        Case 7
            GetValColour = gSampleShowColour.复查偏低
        Case 8
            GetValColour = gSampleShowColour.复查偏高
    End Select
End Function



'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-09-26
'功    能:  通过诊疗ID获取组合项目明细
'入    参:
'           strInfo         intType=0,组合项目对应的诊疗项目的ID，多个使用“,”分割;intType=1,诊疗编码，多个使用逗号分割
'出    参:
'返    回:  组成组合项目的指标记录集
'调整影响:
'调用注意:  intType=0时，返回的记录集有4个字段:指标ID，指标名称，耐受方案ID，耐受时间
'           intType=1时，返回的记录集有8个字段:组合名称，组合编码，诊疗编码，标本类型，指标ID，指标名称，耐受方案ID，耐受时间
'---------------------------------------------------------------------------------------
Public Function funGetGroupItemInfo(ByVal strInfo As String, Optional ByVal intType As Integer) As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strCode As String

          '通过诊疗项目ID查询诊疗项目编码
1         On Error GoTo funGetGroupItemInfo_Error

2         If intType = 0 Then
              '通过诊疗项目ID获取组合明细
3             strSQL = "Select /*+cardinality(d,10)*/ f_List2str(Cast(Collect(a.编码) As t_Strlist)) 编码" & vbCrLf & _
                       "   From 诊疗项目目录 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) B" & vbCrLf & _
                       "   Where a.id = b.Column_Value"
4             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "诊疗项目目录", strInfo)
5             If Not rsTmp.EOF Then
6                 strCode = rsTmp("编码") & ""
7             End If

              '通过诊疗项目编码查询检验组合明细
8             If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
9                 strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "   c.ID 指标ID, c.中文名 检验指标, d.id skey, d.耐受时间 sname" & vbCrLf & _
                           "   From 检验组合项目 A, 检验组合指标 B, 检验指标 C, 检验耐受时间方案 D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.组合id And b.项目id = c.id And b.项目id = d.项目id(+) And a.诊疗编码 = e.Column_Value"
10            Else
11                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "   c.ID 指标ID, c.中文名 检验指标, '' skey, '' sname" & vbCrLf & _
                           "   From 检验组合项目 A, 检验组合指标 B, 检验指标 C, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.组合id And b.项目id = c.id And a.诊疗编码 = e.Column_Value"
12            End If
13            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "组合项目明细", strCode)
14        ElseIf intType = 1 Then
              '通过诊疗项目编码查询检验组合明细
15            If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
16                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "    a.名称, a.编码, a.诊疗编码, a.检验标本 标本类型, c.ID 指标ID, c.中文名 检验指标, d.id 耐受方案ID, d.耐受时间" & vbCrLf & _
                           "   From 检验组合项目 A, 检验组合指标 B, 检验指标 C, 检验耐受时间方案 D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where a.id = b.组合id And b.项目id = c.id And b.项目id = d.项目id(+) And a.诊疗编码 = e.Column_Value" & vbCrLf & _
                           "   Order By a.id, c.排列序号"
17            Else
18                strSQL = "Select /*+cardinality(e,10)*/" & vbCrLf & _
                           "    a.名称, a.编码, a.诊疗编码, a.检验标本 标本类型, c.ID 指标ID, c.中文名 检验指标, '' 耐受方案ID, '' 耐受时间" & vbCrLf & _
                           "   From 检验组合项目 A, 检验组合指标 B, 检验指标 C, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E" & vbCrLf & _
                           "   Where A.ID = B.组合id And B.项目id = C.ID And A.诊疗编码 = e.Column_Value" & vbCrLf & _
                           "   Order By a.id, c.排列序号"
19            End If
20            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "组合项目明细", strInfo)
21        End If
22        Set funGetGroupItemInfo = rsTmp


23        Exit Function
funGetGroupItemInfo_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(funGetGroupItemInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-10-29
'功    能:  通过医嘱ID获取XML格式的病人报告（新版LIS和老版LIS）
'入    参:
'           strAdviceID     医嘱ID串，多个使用逗号分割
'出    参:
'           strErr          错误信息或者提示信息
'返    回:  包含传入医嘱ID对应的所有检验指标及结果，记录集内容：中文名,检验结果,单位

'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function funGetPatientReport(ByVal strAdviceID As String, Optional ByRef strErr As String) As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsData As ADODB.Recordset
          Dim strXML As String

          '查询新版LIS报告
1         On Error GoTo funGetPatientReport_Error

2         strSQL = "Select Distinct /*+cardinality(e,10)*/  '[' || f.名称 || ']' || c.中文名 中文名, a.检验结果, c.单位" & vbCrLf & _
                 "   From 检验报告明细 A, 检验申请组合 B, 检验指标 C, 检验报告记录 D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E, 检验组合项目 F" & vbCrLf & _
                 "   Where b.医嘱id = e.Column_Value And a.标本id = b.标本id And a.项目id = c.id And a.标本id = d.ID And a.组合id = f.id And" & vbCrLf & _
                 "         d.审核人 Is Not Null"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "新版结果", strAdviceID)

4         Set rsData = gobjHisDatabase.CopyNewRec(rsTmp)

          '查询老版
5         strSQL = "Select Distinct /*+cardinality(e,10)*/  '[' || f.名称 || ']' || c.中文名 中文名, b.检验结果, c.单位" & vbCrLf & _
                 "   From 检验项目分布 A, 检验普通结果 B, 诊治所见项目 C, 检验标本记录 D, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) E, 诊疗项目目录 F" & vbCrLf & _
                 "   Where a.标本id = b.检验标本id And a.项目id = b.检验项目id And b.检验项目id = c.id And a.标本ID = d.id And b.诊疗项目ID = f.id And" & vbCrLf & _
                 "         a.医嘱Id = e.Column_Value And d.审核人 Is Not Null"
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "老版结果", strAdviceID)

          '将老版和新版报告合并在一起
7         Do While Not rsTmp.EOF
8             rsData.AddNew
9             rsData("中文名") = rsTmp("中文名") & ""
10            rsData("检验结果") = rsTmp("检验结果") & ""
11            rsData("单位") = rsTmp("单位") & ""

12            rsTmp.MoveNext
13        Loop
14        If rsData.RecordCount > 0 Then rsData.MoveFirst

15        Set funGetPatientReport = rsData

16        Exit Function
funGetPatientReport_Error:
17        strErr = WriteErrLog("zl9LisInsideComm", "mdlLisHisComm", "执行(funGetPatientReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
18        Err.Clear
End Function
