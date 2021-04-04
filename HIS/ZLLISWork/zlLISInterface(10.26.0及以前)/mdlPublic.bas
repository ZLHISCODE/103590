Attribute VB_Name = "mdlPublic"
Option Explicit
Public gcnOracle As New ADODB.Connection            '公共数据库连接
Public gstrSql As String                            '公共SQL字串
Public mclsZip As New cZip
Public mclsUnzip As New cUnzip
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public gobjComLib As Object                         '公共函数对象
Private gstrSysName As String

Private gstrDbUser As String                 '当前数据库用户
Private glngUserId As Long                   '当前用户id
Private gstrUserCode As String               '当前用户编码
Private gstrUserName As String               '当前用户姓名
Private gstrUserAbbr As String               '当前用户简码

Private glngDeptId As Long                   '当前用户部门id
Private gstrDeptCode As String               '当前用户部门编码
Private gstrDeptName As String               '当前用户部门名称
Private gstrPrivs As String                  '权限
Private gstr接口权限  As String              '接口本身的授权

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
        '------------------------------------------------
        '功能： 打开指定的数据库
        '参数：
        '   strServerName：主机字符串
        '   strUserName：用户名
        '   strUserPwd：密码
        '返回： 数据库打开成功，返回true；失败，返回false
        '------------------------------------------------
        Dim strSQL As String
        Dim strError As String
    
        On Error Resume Next
100     Err = 0
102     DoEvents
104     With gcnOracle
106         If .State = adStateOpen Then .Close
108         .Provider = "MSDataShape"
110         .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, TranPasswd(strUserPwd)
112         If Err <> 0 Then
                '保存错误信息
114             strError = Err.Description
116             If InStr(strError, "自动化错误") > 0 Then
118                 WriteLog "连接串无法创建，请检查数据访问部件是否正常安装。"
120             ElseIf InStr(strError, "ORA-12154") > 0 Then
122                 WriteLog "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。"
124             ElseIf InStr(strError, "ORA-12541") > 0 Then
126                 WriteLog "无法连接，请检查服务器上的Oracle监听器服务是否启动。"
128             ElseIf InStr(strError, "ORA-01033") > 0 Then
130                 WriteLog "ORACLE正在初始化或在关闭，请稍候再试。"
132             ElseIf InStr(strError, "ORA-01034") > 0 Then
134                 WriteLog "ORACLE不可用，请检查服务或数据库实例是否启动。"
136             ElseIf InStr(strError, "ORA-02391") > 0 Then
138                 WriteLog "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。"
140             ElseIf InStr(strError, "ORA-01017") > 0 Then
142                 WriteLog "由于用户、口令或服务器指定错误，无法登录。"
144             ElseIf InStr(strError, "ORA-28000") > 0 Then
146                 WriteLog "由于用户已经被禁用，无法登录。"
                Else
148                 WriteLog strError
                End If
            
150             OraDataOpen = False
                Exit Function
            End If
        End With
    
152     Err = 0
        On Error GoTo errHand
    
154     gstrDbUser = UCase(strUserName)
    
    
156     Call gobjComLib.InitCommon(gcnOracle)
158     Call GetUserInfo
160     If CheckRegInfo = True Then
162         gstrPrivs = gobjComLib.GetPrivFunc(100, 1208)
164         OraDataOpen = True
        Else
166         gcnOracle.Close
168         Set gcnOracle = Nothing
        End If
        Exit Function
    
errHand:
170     WriteLog "OraOpen " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
172     OraDataOpen = False
174     Err = 0
End Function

Private Function CheckRegInfo() As Boolean

        Dim objFile As Scripting.TextStream, strLine As String, str日期 As String, date日期 As Date
        Dim strUnti As String, strCode As String, strKey As String
        Dim strPrivs As String '存授权功能
        On Error GoTo hErr
100     strKey = "陈东"
    
102     If gobjComLib.RegCheck = False Then
104         WriteLog "未能通HIS系统注册相关检查，请先检查ZLLIS能否正常运行！"
            Exit Function
        End If
106     strUnti = Trim(gobjComLib.zlRegInfo("单位名称", , -1))
108     If gobjFSO.FileExists(App.Path & "\RegFile.ini") Then
110         Set objFile = gobjFSO.OpenTextFile(App.Path & "\RegFile.ini")
        
112         Do Until objFile.AtEndOfLine
114             strLine = objFile.ReadLine
116             If strLine Like "授权截止日期=*" Then
118                 str日期 = Trim(Split(strLine, "=")(1))
120             ElseIf strLine Like "授权码=*" Then
122                 strCode = Trim(Split(strLine, "=")(1))
124             ElseIf strLine Like "授权功能=*" Then
126                 strPrivs = Trim(Split(strLine, "=")(1))
                End If
            Loop
    
128         If IsDate(str日期) Then
130             date日期 = gobjComLib.zlDatabase.Currentdate
132             If date日期 <= CDate(str日期) Then
134                 If strCode <> Md5_String_Calc(strUnti & "|" & str日期 & "|" & strKey & strPrivs) Then
                    
136                     WriteLog "授权码不正确！" & vbNewLine & _
                               "单位：" & strUnti & vbNewLine & _
                               "注册码：" & strCode & vbNewLine & _
                               "日期：" & str日期 & vbNewLine & _
                               "权限：" & strPrivs & vbNewLine & _
                               "注册文件：" & App.Path & "\RegFile.ini"
                    
                    Else
138                     CheckRegInfo = True
140                     If strPrivs <> "" Then gstr接口权限 = strPrivs
                    End If
                Else
142                 WriteLog "已超过试用期限！"
                End If
            Else
144             WriteLog "试用日期错误！"
            End If
        Else
146         WriteLog "部件所在目录缺少授权文件（RegFile.ini）！"
        End If
        Exit Function
hErr:
148     WriteLog "CheckReg " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Public Function GetApplication(strPatientID As String) As String
        '=========================================================================================
        '功能:                              得到病人申请单的记录集
        '参数
        'strPatientID                       数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号
        '=========================================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim lngPatientID As Long
        Dim strData As String, blnBacode As Boolean
    
        '没有查询条件时退出
100     If strPatientID = "" Then Exit Function
102     blnBacode = False
        On Error GoTo errH
    
104     Select Case Mid(strPatientID, 1, 1)
            Case "-"
106             gstrSql = "select 病人ID,姓名,性别,年龄,门诊号,住院号,就诊卡号,身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,健康号,IC卡号,医保号,险类 " & _
                         ",当前床号 from 病人信息 a , 部门表 b where a.当前科室ID = b.ID(+) and 病人id = [1]"
108         Case "+"
110             gstrSql = "select 病人ID,姓名,性别,年龄,门诊号,住院号,就诊卡号,身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,健康号,IC卡号,医保号,险类 " & _
                         ",当前床号 from 病人信息 a , 部门表 b where a.当前科室ID = b.ID(+) and a.住院号 = [1] "
112         Case "*"
114             gstrSql = "select 病人ID,姓名,性别,年龄,门诊号,住院号,就诊卡号,身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,健康号,IC卡号,医保号,险类 " & _
                         ",当前床号 from 病人信息 a , 部门表 b where a.当前科室ID = b.ID(+) and a.门诊号 = [1] "
116         Case "."
118             gstrSql = "select 病人ID,姓名,性别,年龄,门诊号,住院号,就诊卡号,身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,健康号,IC卡号,医保号,险类 " & _
                         ",当前床号 from 病人信息 a , 部门表 b where a.当前科室ID = b.ID(+) and a.挂号单 = [2] "
120         Case "/"
122             gstrSql = "Select Distinct b.病人ID,b.姓名,b.性别,b.年龄,b.门诊号,b.住院号,b.就诊卡号,b.身份证号,c.编码 as 当前科室编码,c.名称 as 当前科室名称,b.健康号,b.IC卡号,b.医保号,b.险类  " & vbNewLine & _
                        "From 病人费用记录 A, 病人信息 B , 部门表 C " & vbNewLine & _
                        "Where A.病人id = B.病人id And A.NO = [2] And A.病人id Is Not Null And A.门诊标志 = 1 and b.当前科室id = c.id(+) " & vbNewLine & _
                        "Order By 病人id Desc"
124         Case "\" '健康号
126             gstrSql = "select a.病人ID,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.就诊卡号,a.身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,a.健康号,a.IC卡号,a.医保号,a.险类 " & _
                         ",当前床号 from 病人信息 a , 部门表 b where a.当前科室ID = b.ID(+) and a.健康号 = [2] "
128         Case Else
130             If Len(strPatientID) >= 12 Then
132                 blnBacode = True
134                 gstrSql = "Select Distinct c.病人ID,c.姓名,c.性别,c.年龄,c.门诊号,c.住院号,c.就诊卡号,c.身份证号,d.编码 as 当前科室编码,d.名称 as 当前科室名称,c.健康号,c.IC卡号,c.医保号,c.险类 " & vbNewLine & _
                                " From 病人医嘱记录 A, 病人医嘱发送 B , 病人信息 C,部门表 d Where A.ID = B.医嘱id and " & vbNewLine & _
                                " a.病人ID = C.病人ID and c.当前科室ID = d.id(+) And B.样本条码 = [2] "
                Else
136                 gstrSql = "Select a.病人ID,a.姓名,a.性别,a.年龄,a.门诊号,a.住院号,a.就诊卡号,a.身份证号,b.编码 as 当前科室编码,b.名称 as 当前科室名称,a.健康号,a.IC卡号,a.医保号,a.险类 " & vbNewLine & _
                                "From 病人信息 a,部门表 b " & vbNewLine & _
                                "Where a.当前科室ID = b.id(+) and  就诊卡号 = [2] "
                End If
        End Select
    
138     If InStr(",-,+,*,.,/,\,", "," & Mid(strPatientID, 1, 1) & ",") > 0 Then
140         strPatientID = Mid(strPatientID, 2)
        End If
        
142     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "提取申清单", Val(strPatientID), CStr(strPatientID))
    
        '没有找到结果时退出
144     If rsTmp.EOF = True Then Exit Function
    
        '当找到多个病人时退出(返回多个病人的ID和信息)
146     If rsTmp.RecordCount > 1 Then
148         Do Until rsTmp.EOF
150             strData = strData & "|" & rsTmp("病人ID") & "^" & rsTmp("姓名") & "^" & rsTmp("性别") & "^" & rsTmp("年龄") & _
                          "^" & rsTmp("门诊号") & "^" & rsTmp("住院号") & "^" & rsTmp("就诊卡号") & "^" & rsTmp("身份证号") & _
                          "^" & rsTmp("当前科室编码") & "^" & rsTmp("当前科室名称") & "^" & rsTmp("健康号") & "^" & rsTmp("险类")
152             rsTmp.MoveNext
            Loop
154         If strData <> "" Then
156             GetApplication = Mid(strData, 2)
            End If
            Exit Function
        End If
    
158     lngPatientID = "" & rsTmp("病人ID")
    
        '提取申请单
    '    gstrSql = "Select A.*, To_Char(B.开嘱时间, 'YYYY-MM-DD HH24:MI') As 申请时间, B.医嘱内容 As 申请项目, B.标本部位 As 标本类型, B.开嘱科室id, B.开嘱医生" & vbNewLine & _
                    "From (Select Decode(Sum(Decode(Z.记录状态, 1, 1, 0)), 0, 0, 1) As 选择, A.相关id As ID," & vbNewLine & _
                    "              C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)') As 姓名, C.门诊号, C.住院号, D.名称 As 申请科室, A.开嘱医生 As 申请人, F.接收人," & vbNewLine & _
                    "              F.接收时间, 'Item' As 图标, Decode(Sum(Decode(Z.记录状态, 1, 1, 0)), 0, '  ', '√') As 收费, Nvl(A.紧急标志, 0) As 紧急, H.操作类型," & vbNewLine & _
                    "              Max(Decode(I.项目类别, 2, 2, 1)) As 项目类别, Max(F.采样人) As 采样人, Max(F.接收时间) As 采样时间" & vbNewLine & _
                    "       From 病人医嘱记录 A, 病人信息 C, 部门表 D, 病人医嘱发送 F, 检验报告项目 G, 诊疗项目目录 H, 检验项目 I, 病人费用记录 Z" & vbNewLine & _
                    "       Where A.诊疗类别 = 'C' And A.病人id = C.病人id And A.开嘱科室id = D.ID And A.相关id Is Not Null And A.医嘱状态 = 8 And A.ID = F.医嘱id And" & vbNewLine & _
                    "             A.诊疗项目id = G.诊疗项目id And G.细菌id Is Null And G.报告项目id = I.诊治项目id And A.诊疗项目id = H.ID And F.执行状态 = 0 And" & vbNewLine & _
                    "             A.病人id = [1] And F.NO = Z.NO(+) And F.记录性质 = Z.记录性质(+) And F.医嘱id = Z.医嘱序号(+) + 0" & vbNewLine & _
                    "       Group By A.相关id, C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)'), C.门诊号, C.住院号, D.名称, A.开嘱医生, 'Item', Nvl(A.紧急标志, 0)," & vbNewLine & _
                    "                H.操作类型, F.接收人, F.接收时间" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sum(Decode(Z.记录状态, 1, 1, 0)), 0, 0, 1) As 选择, A.相关id As ID," & vbNewLine & _
                    "              C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)') As 姓名, C.门诊号, C.住院号, D.名称 As 申请科室, A.开嘱医生 As 申请人, F.接收人," & vbNewLine & _
                    "              F.接收时间, 'Item' As 图标, Decode(Sum(Decode(Z.记录状态, 1, 1, 0)), 0, '  ', '√') As 收费, Nvl(A.紧急标志, 0) As 紧急, H.操作类型," & vbNewLine & _
                    "              Max(Decode(I.项目类别, 2, 2, 1)) As 项目类别, Max(F.采样人) As 采样人, Max(F.接收时间) As 采样时间" & vbNewLine & _
                    "       From 病人医嘱记录 A, 病人信息 C, 部门表 D, 病人医嘱发送 F, 检验报告项目 G, 诊疗项目目录 H, 检验项目 I, 病人费用记录 Z, 检验标本记录 J, 检验项目分布 K" & vbNewLine & _
                    "       Where A.诊疗类别 = 'C' And A.病人id = C.病人id And A.开嘱科室id = D.ID And A.相关id Is Not Null And A.医嘱状态 = 8 And A.ID = F.医嘱id And" & vbNewLine & _
                    "             A.诊疗项目id = G.诊疗项目id And G.细菌id Is Null And G.报告项目id = I.诊治项目id And A.诊疗项目id = H.ID And F.执行状态 = 3 And" & vbNewLine & _
                    "             A.病人id = [1] And F.NO = Z.NO(+) And F.记录性质 = Z.记录性质(+) And F.医嘱id = Z.医嘱序号(+) + 0 And A.相关id = K.医嘱id(+) And" & vbNewLine & _
                    "             J.ID = K.标本id And J.ID = 0" & vbNewLine & _
                    "       Group By A.相关id, C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)'), C.门诊号, C.住院号, D.名称, A.开嘱医生, 'Item', Nvl(A.紧急标志, 0)," & vbNewLine & _
                    "                H.操作类型, F.接收人, F.接收时间) A, 病人医嘱记录 B" & vbNewLine & _
                "Where A.ID = B.ID"
160     gstrSql = "Select Distinct A.相关id As ID, D.姓名, D.性别, D.年龄, A.病人来源, D.门诊号, D.住院号, E.编码 As 申请科室编码, E.名称 As 申请科室名称, A.开嘱医生, A.开嘱时间,D.险类,C.样本条码,C.条码打印 " & vbNewLine & _
                    "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱发送 C, 病人信息 D, 部门表 E" & vbNewLine & _
                    "Where A.诊疗项目id = B.ID And B.类别 = 'C' And A.ID = C.医嘱id And A.相关id Is Not Null And C.执行状态 = 0 And A.病人id = [1] And" & vbNewLine & _
                    "      A.病人id = D.病人id And A.开嘱科室id = E.ID And A.医嘱状态 = 8"
162     If blnBacode Then gstrSql = gstrSql & " And C.样本条码 = [2] "
    
164     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "提取申请单", lngPatientID, strPatientID)
166     Do Until rsTmp.EOF
168         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("姓名") & "^" & rsTmp("性别") & "^" & rsTmp("年龄") & "^" & rsTmp("病人来源") & _
                      "^" & rsTmp("门诊号") & "^" & rsTmp("住院号") & "^" & rsTmp("申请科室编码") & "^" & rsTmp("申请科室名称") & _
                      "^" & rsTmp("开嘱医生") & "^" & rsTmp("开嘱时间") & "^" & rsTmp("险类") & "^" & rsTmp("样本条码") & "^" & rsTmp("条码打印")
170         rsTmp.MoveNext
        Loop
172     If strData <> "" Then
174         GetApplication = Mid(strData, 2)
        End If
    
        Exit Function
errH:
176     WriteLog "GetApplication " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function OraDataClose() As Boolean
    '------------------------------------------------
    '功能： 关闭数据库
    '参数：
    '返回： 关闭数据库，返回True；失败，返回False
    '------------------------------------------------
    Err = 0
    On Error Resume Next
    gcnOracle.Close
    OraDataClose = True
    Err = 0

End Function

Public Function InsertReport(lngID As Long, strReportPath As String, ErrInfo As String, Optional lngDeviceID As Long, Optional strSampleNo As String, Optional strItems As String) As Boolean
        '===================================================================
        '功能                               插入报告到HIS
        '参数
        'lngID                              医嘱ID
        'strReportPath                      报告路径
        '===================================================================
        Dim rsTmp As ADODB.Recordset
        Dim aStrSQL() As String                     '数组SQL字串
        Dim intLoop  As Integer
        Dim strZipFile As String                    '压缩后的文件
        Dim strUnZipFile As String                  '解压后的文件
        Dim strPath As String                       '临时文件路径
    
        On Error GoTo errH
    
100     If Dir(strReportPath) = "" Then Exit Function
102     strPath = IIf(Len(App.Path) <= 3, App.Path & "TMP.RTF", App.Path & "\TMP.RTF")
    
104     If gobjFSO.FileExists(strPath) = True Then gobjFSO.DeleteFile strPath
    
106     Call gobjFSO.CopyFile(strReportPath, strPath)
    
108     If gobjFSO.FileExists(strPath) = False Then Exit Function
    
110     gstrSql = "Zl_检验报告单_Insert(" & lngID & ",0)"
112     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "插入报告"
    
114     gstrSql = "Select Nvl(A.病历id, 0) As 文件id From 病人医嘱报告 A Where A.医嘱id = [1] "
116     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "插入报告", lngID)
118     If rsTmp.EOF = True Then Exit Function
    
    
120     strZipFile = zlFileZip(strPath)
    
122     strUnZipFile = zlFileUnzip(strZipFile)
    
    
124     If zlLisBlobSql(rsTmp("文件ID"), strZipFile, aStrSQL) = False Then Exit Function
    
126     For intLoop = 0 To UBound(aStrSQL)
128         gobjComLib.zlDatabase.ExecuteProcedure Replace(aStrSQL(intLoop), "Call", ""), "插入报告"
    '        Debug.Print aStrSQL(intLoop)
        Next
130     gobjFSO.DeleteFile strZipFile
132     gobjFSO.DeleteFile strPath
134     InsertReport = True
        Exit Function
errH:
136     ErrInfo = CStr(Erl()) & "," & Err.Description
138     WriteLog "InsertReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description

End Function

Private Function zlLisBlobSql(ByVal KeyWord As String, ByVal strFile As String, ByRef arySql() As String) As Boolean
    '生成保存报告文件
    'KeyWord 文件ID
    'strFile 文件路径
    'arySql 生成的SQL存放在此数组中
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '传入数组的最小最大下标
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 500
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
'            If lngCount = 0 Then strText = "100;" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_电子病历格式_Insert(" & KeyWord & ",'" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strZipPath & "TMP.RTF") Then gobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

'################################################################################################################
'## 功能：  将文件压缩为新文件放到相同目录中
'## 参数：  strFile     :原始文件
'## 返回：  压缩文件名，失败则返回零长度""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
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
End Function

Public Function DeleteReport(lngID As Long) As Boolean
        '===================================================================
        '功能                               删除报告
        '参数
        'lngID                              医嘱ID
        '===================================================================
        On Error GoTo errH
100     gstrSql = "Zl_检验报告单_Insert(" & lngID & ",1)"
102     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "删除报告"
104     DeleteReport = True
        Exit Function
errH:
106     WriteLog "DeleteReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description

End Function
    
Public Function GetClinicItem(lngAdivce As Long) As String
        '===================================================================
        '功能                               取得要做的诊疗项目内容
        '参数
        'lngAdivce                          医嘱ID
        '返回                               字串格式:诊疗项目ID^诊疗项目编码^诊疗项目名称^执行科室编码^执行科室名称^单价^金额^是否收费
        '===================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim strData As String, str病人来源 As String
        On Error GoTo errH
    
    '    gstrSql = "Select a.诊疗项目id as ID, b.编码 as 诊疗项目编码, b.名称 as 诊疗项目名称, c.编码 as 执行科室编码, C.名称 As 执行科室名称,E.实收金额,E.标准单价,E.记录状态,'0' as 是否采集" & vbNewLine & _
    '            "From 病人费用记录 E,病人医嘱发送 D,病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
    '            "Where D.记录性质=E.记录性质(+) And D.No=E.No(+) And D.记录序号=E.序号(+) And A.诊疗类别='C' And a.ID=D.医嘱Id And A.诊疗项目id = B.ID And A.执行科室id = C.ID And A.相关id = [1] " & _
    '            "Union all " & _
    '            "Select a.诊疗项目id as ID, b.编码 as 诊疗项目编码, b.名称 as 诊疗项目名称, c.编码 as 执行科室编码, C.名称 As 执行科室名称,E.实收金额,E.标准单价,E.记录状态,'1' as 是否采集" & vbNewLine & _
    '            "From 病人费用记录 E,病人医嘱发送 D,病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
    '            "Where D.记录性质=E.记录性质(+) And D.No=E.No(+) And D.记录序号=E.序号(+) And A.诊疗类别='E' And a.ID=D.医嘱Id And A.诊疗项目id = B.ID And A.执行科室id = C.ID And A.id = [1] "
100     gstrSql = "Select 病人来源 From 病人医嘱记录 Where ID=[1]"
102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "得到病人来源", lngAdivce)
104     Do Until rsTmp.EOF
106         str病人来源 = Trim("" & rsTmp!病人来源)
108         rsTmp.MoveNext
        Loop
110     If str病人来源 = "4" Then
            '体检病人
112         gstrSql = "Select A.诊疗项目id As ID, B.编码 As 诊疗项目编码, B.名称 As 诊疗项目名称, C.编码 As 执行科室编码, C.名称 As 执行科室名称, Sum(E.实收金额) As 实收金额," & vbNewLine & _
                    "       Sum(E.标准单价) As 标准单价, E.记录状态, '0' As 是否采集" & vbNewLine & _
                    "From 病人费用记录 E, 病人医嘱发送 D, 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                    "Where D.记录性质 = E.记录性质(+) And D.No = E.No(+) And D.医嘱id = E.医嘱序号(+) And E.记录状态(+) <> 2 And A.诊疗类别 = 'C' And" & vbNewLine & _
                    "      A.Id = D.医嘱id And A.诊疗项目id = B.Id And A.执行科室id = C.Id And A.相关id = [1]" & vbNewLine & _
                    "Group By A.诊疗项目id, B.编码, B.名称, C.编码, C.名称, E.记录状态" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select A.诊疗项目id As ID, B.编码 As 诊疗项目编码, B.名称 As 诊疗项目名称, C.编码 As 执行科室编码, C.名称 As 执行科室名称, Sum(E.实收金额) As 实收金额," & vbNewLine & _
                    "       Sum(E.标准单价) As 标准单价, Decode(E.记录状态, 1, 1, 3, 1, 0) As 计费状态, '1' As 是否采集" & vbNewLine & _
                    "From 病人费用记录 E, 病人医嘱发送 D, 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                    "Where D.记录性质 = E.记录性质(+) And D.No = E.No(+) And D.医嘱id = E.医嘱序号(+) And E.记录状态(+) <> 2 And A.诊疗类别 = 'E' And" & vbNewLine & _
                    "      A.Id = D.医嘱id And A.诊疗项目id = B.Id And A.执行科室id = C.Id And A.Id = [1]" & vbNewLine & _
                    "Group By A.诊疗项目id, B.编码, B.名称, C.编码, C.名称, E.记录状态"


        Else
114         gstrSql = "Select A.诊疗项目id As ID, B.编码 As 诊疗项目编码, B.名称 As 诊疗项目名称, C.编码 As 执行科室编码, C.名称 As 执行科室名称, Sum(E.数量 * E.单价) As 实收金额," & vbNewLine & _
                "       Sum(E.单价) As 标准单价, D.计费状态, '0' As 是否采集, F.记录状态" & vbNewLine & _
                "From 病人费用记录 F, 病人医嘱计价 E, 病人医嘱发送 D, 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                "Where D.医嘱id = F.医嘱序号(+) And D.No = F.No(+) And D.记录性质 = F.记录性质(+) And D.记录序号 = F.序号(+) And F.记录状态(+) <> 2 And" & vbNewLine & _
                "      A.Id = E.医嘱id And A.诊疗类别 = 'C' And A.Id = D.医嘱id And A.诊疗项目id = B.Id And A.执行科室id = C.Id And A.相关id = [1]" & vbNewLine & _
                "Group By A.诊疗项目id, B.编码, B.名称, C.编码, C.名称, D.计费状态, F.记录状态" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select A.诊疗项目id As ID, B.编码 As 诊疗项目编码, B.名称 As 诊疗项目名称, C.编码 As 执行科室编码, C.名称 As 执行科室名称, Sum(E.数量 * E.单价) As 实收金额," & vbNewLine & _
                "       Sum(E.单价) As 标准单价, D.计费状态, '1' As 是否采集, F.记录状态" & vbNewLine & _
                "From 病人费用记录 F, 病人医嘱计价 E, 病人医嘱发送 D, 病人医嘱记录 A, 诊疗项目目录 B, 部门表 C" & vbNewLine & _
                "Where D.医嘱id = F.医嘱序号(+) And D.No = F.No(+) And D.记录性质 = F.记录性质(+) And D.记录序号 = F.序号(+) And F.记录状态(+) <> 2 And" & vbNewLine & _
                "      A.Id = E.医嘱id And A.诊疗类别 = 'E' And A.Id = D.医嘱id And A.诊疗项目id = B.Id And A.执行科室id = C.Id And A.Id = [1]" & vbNewLine & _
                "Group By A.诊疗项目id, B.编码, B.名称, C.编码, C.名称, D.计费状态, F.记录状态"

        End If
116     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "得到诊疗项目", lngAdivce)
    
118     Do Until rsTmp.EOF
120         strData = strData & "|" & rsTmp("ID") & "^" & rsTmp("诊疗项目编码") & "^" & rsTmp("诊疗项目名称") & "^" & rsTmp("执行科室编码") & _
                        "^" & rsTmp("执行科室名称") & "^" & rsTmp("标准单价") & "^" & rsTmp("实收金额") & "^" & rsTmp("记录状态") & "^" & rsTmp("是否采集")
122         rsTmp.MoveNext
        Loop

124     If strData <> "" Then
126         GetClinicItem = Mid(strData, 2)
        End If
    
        Exit Function
errH:
128     WriteLog "GetClinicItem " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetItemList(lngClinicID As Long) As String
        '===================================================================
        '功能                               取得诊疗项目的指标明细
        '参数
        'lngClinicID                        诊疗项目ID
        '返回
        '===================================================================
        Dim rsTmp As New ADODB.Recordset
        Dim strData As String
        On Error GoTo errH
    
100     gstrSql = "Select B.编码, B.中文名, B.英文名 " & vbNewLine & _
                " From 检验报告项目 A, 诊治所见项目 B " & vbNewLine & _
                " Where A.报告项目id = B.ID And a.诊疗项目ID = [1] "

102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "取得指标明细", lngClinicID)
104     Do Until rsTmp.EOF
106         strData = strData & "|" & rsTmp("编码") & "^" & rsTmp("中文名") & "^" & rsTmp("英文名")
108         rsTmp.MoveNext
        Loop
    
110     If strData <> "" Then
112         GetItemList = Mid(strData, 2)
        End If
    
        Exit Function
errH:
114     WriteLog "GetItemList " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function SetRegister(lngAdivce As Long, intTag As Integer) As Boolean
        '=====================================================================
        '功能                               标本核收或取消核收
        '参数
        'lngAdivce                          医嘱ID
        'intTag                             1=核收 0=取消核收 11-在LIS中核收，10-在LIS中取消核收
        '=====================================================================
        On Error GoTo errH
100     gstrSql = "Zl_检验医嘱标记_Edit(" & lngAdivce & "," & intTag & ")"
102     gobjComLib.zlDatabase.ExecuteProcedure gstrSql, "核收或取消核收"
104     SetRegister = True

        Exit Function
errH:
106     WriteLog "SetRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetAllItem(Optional strFindItem As String) As String()
        '=====================================================================
        '功能                               取得所有的诊疗项目编码和名称
        '参数
        'strItem                            可选，查找编码和名称相同的诊疗项目项目
        '返回                               查找到的诊疗项目数组
        '=====================================================================
        Dim astrItem() As String
        Dim rsTmp As New ADODB.Recordset
        Dim strSQL As String
        Dim strItem As String
        Dim intLoop As Integer
    
100     ReDim Preserve astrItem(0)
102     gstrSql = "select ID,编码,名称,组合项目 from 诊疗项目目录  where 类别 = 'C' "
104     If strFindItem <> "" Then
106         gstrSql = gstrSql & " And (编码 = [1] or 名称 like '%[1]%') "
        End If
    
108     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "提取诊疗项目", CStr(strFindItem))
    
110     Do Until rsTmp.EOF
112         strItem = strItem & ";" & rsTmp("编码") & "," & rsTmp("名称") & "," & rsTmp("组合项目")
114         intLoop = intLoop + 1
116         If intLoop >= 200 Then
118             If astrItem(0) <> "" Then
120                 ReDim Preserve astrItem(UBound(astrItem) + 1)
                End If
122             astrItem(UBound(astrItem)) = Mid(strItem, 2)
124             strItem = ""
126             intLoop = 0
            End If
128         rsTmp.MoveNext
        Loop
130     If intLoop <> 0 Then
132         If astrItem(0) <> "" Then
134             ReDim Preserve astrItem(UBound(astrItem) + 1)
            End If
136         astrItem(UBound(astrItem)) = Mid(strItem, 2)
        End If
    
138     GetAllItem = astrItem
        On Error GoTo errH
        Exit Function
errH:
140     WriteLog "GetAllItem " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function UpdateTestResults(ByVal lngID As Long, ByVal strTestName As String, ByVal strTestTime As String, ByVal strTestResults As String) As String
        '===================================================================
        '功能                               返回检验结果到体检系统
        '参数
        'lngID                              医嘱ID
        'strTestName                        检验人
        'strTestTime                        检验时间，格式 2009-01-01 10:30:01
        'strTestResults                     医嘱ID对应的检验结果，可以对多少个检验指标一起处理，详细格式如下：
        '
        '                                     诊治项目id;检验结果1;单位1;结果参1考;结果标志1|诊治项目id;检验结果2;单位2;结果参考2;结果标志2......
        '
        '                                     其中，结果标志在 “偏低,偏高,异常,空串”中选择一个返回。
        '返回: 空，表示更新成功，非空，表示错误信息。
        '===================================================================
        Dim strSQL As String
        Dim rsTmp As ADODB.Recordset, i As Integer
        Dim varItem As Variant, strItem As String, str体检指标 As String, str诊治项目id As String, strErrInfo As String
        Dim strEditSQL() As String
        On Error GoTo errH
    
100     str体检指标 = ""
102     strErrInfo = ""
104     ReDim strEditSQL(0) As String
    
106     If Not strTestTime Like "####-##-## ##:##:##" Or IsDate(CDate(strTestTime)) = False Then
108         strErrInfo = strErrInfo & "0|检验日期格式不正确，请按yyyy-MM-dd HH24:MI:SS的格式调整！" & vbNewLine
110         UpdateTestResults = strErrInfo
            Exit Function
        End If
            
112     strSQL = "Select /*+Rule */" & vbNewLine & _
                " a.病人id, a.清单id, a.任务id, c.检查人, c.检查时间, c.体检指标id, d.诊治项目id, a.采集医嘱id, f.编码" & vbNewLine & _
                "From 诊疗项目目录 f, 体检指标目录 d, 检验报告项目 e, 体检任务结果 c, 体检任务发送 a" & vbNewLine & _
                "Where a.采集医嘱id = [1] And a.任务id = c.任务id And a.病人id = c.病人id And a.清单id = c.清单id And" & vbNewLine & _
                "           c.体检指标id = d.Id And f.组合项目 = 0 And d.诊治项目id = e.报告项目id And e.诊疗项目id = f.Id"
            
114     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "提取检验指标", lngID)
116     Do Until rsTmp.EOF
118         str体检指标 = str体检指标 & "," & rsTmp!编码
120         rsTmp.MoveNext
        Loop
    
122     If str体检指标 <> "" Then
124         varItem = Split(strTestResults, "|")
126         For i = LBound(varItem) To UBound(varItem)
128             strItem = varItem(i)
130             If InStr(strItem, ";") > 0 Then
132                 If UBound(Split(strItem, ";")) >= 4 Then
134                     If InStr(str体检指标 & ",", "," & Trim(Split(strItem, ";")(0)) & ",") <= 0 Then
136                         strErrInfo = strErrInfo & "0|编码: " & Split(strItem, ";")(0) & "未找到对应申请，请检查!" & vbNewLine
138                     ElseIf InStr(strItem, "'") > 0 Then
140                         strErrInfo = strErrInfo & "0|第" & i & "项检验结果,单引号不能在接口中出现，请调整！" & vbNewLine
142                     ElseIf InStr(strItem, """") > 0 Then
144                         strErrInfo = strErrInfo & "0|第" & i & "项检验结果,双引号不能在接口中出现，请调整！" & vbNewLine
                        Else
146                         rsTmp.MoveFirst
148                         Do Until rsTmp.EOF
        '                        任务id_In     In 体检任务结果.任务id%Type,
        '                        病人id_In     In 体检任务结果.病人id%Type,
        '                        清单id_In     In 体检任务结果.清单id%Type,
        '                        体检指标id_In In 体检任务结果.体检指标id%Type,
        '                        检验人_In     In 体检任务结果.检查人%Type,
        '                        检验时间_In   In 体检任务结果.检查时间%Type,
        '                        结果_In       In 体检任务结果.结果%Type,
        '                        单位_In       In 体检任务结果.单位%Type,
        '                        参考_In       In 体检任务结果.参考%Type,
        '                        报警_In       In 体检任务结果.报警%Type
150                             If Trim("" & rsTmp!编码) = Trim(Split(strItem, ";")(0)) And Trim(Split(strItem, ";")(0)) <> "" Then
152                                 If Trim("" & rsTmp!检查人) = "" Then '仅更新一次
154                                     If strEditSQL(UBound(strEditSQL)) <> "" Then ReDim Preserve strEditSQL(UBound(strEditSQL) + 1)
156                                     strEditSQL(UBound(strEditSQL)) = "Zl_体检指标_Externaledit(" & rsTmp!任务id & "," & rsTmp!病人id & "," & rsTmp!清单id & "," & rsTmp!体检指标id & ",'" & strTestName & "',to_date('" & strTestTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                                                                         "'" & Split(strItem, ";")(1) & "','" & Split(strItem, ";")(2) & "','" & Split(strItem, ";")(3) & "','" & Split(strItem, ";")(4) & "')"
                                    Else
158                                      strErrInfo = strErrInfo & "1|项目" & Val(Split(strItem, ";")(0)) & "已经有结果" & vbNewLine
                                    End If
                                    Exit Do
                                End If
160                             rsTmp.MoveNext
                            Loop
                        End If
                    Else
162                     strErrInfo = strErrInfo & "0|第" & i & "项检验结果,缺少项目，请检查！" & vbNewLine
                    End If
                Else
164                 strErrInfo = strErrInfo & "0|第" & i & "项检验结果,格式不正确，请检查！" & vbNewLine
                End If
            Next
        Else
166         strErrInfo = strErrInfo & "0|未找到医嘱id=" & lngID & "的体检记录!" & vbNewLine
        End If
    
168     For i = LBound(strEditSQL) To UBound(strEditSQL)
170         If strEditSQL(i) <> "" Then gobjComLib.zlDatabase.ExecuteProcedure strEditSQL(i), "保存体检指标"
        Next
172     UpdateTestResults = strErrInfo
    
        Exit Function
errH:
174     UpdateTestResults = strErrInfo & "0|出现错误：" & CStr(Erl()) & "," & Err.Description
176     WriteLog "UpdateTestResults " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function ZipFile(strPath As String) As String
    ZipFile = zlFileZip(strPath)
End Function

Public Function UnZipFile(strPath As String) As String
    UnZipFile = zlFileUnzip(strPath)
End Function

Public Function zlLISRegister(ByVal lngDeviceID As Long, ByVal lngID As Long, ByVal strSampleNo As String, ByRef strErrInfo As String) As Boolean
        '用于核收标本
        Dim strSQL As String, rsTmp As ADODB.Recordset, rs As New ADODB.Recordset
        Dim lngKey As Long, strItemRecords As String
        Dim lngDeptID As Long '当前仪器科室
        Dim rsItem As New ADODB.Recordset
        Dim strItem As String                           '检验项目
        Dim str姓名 As String, str性别 As String, str年龄 As String
        Dim dtSampleDate As Date, dStart As Date, dEnd As Date
    
        On Error GoTo errH
100     If InStr(gstr接口权限, "ZLLIS标本核收") <= 0 Then
102         strErrInfo = "此接口未授权，不能调用！"
            Exit Function
        End If
        '查找仪器科室
104     strSQL = "Select 使用小组id From 检验仪器 Where ID = [1]"
106     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "生成条码标本", lngDeviceID)
108     lngDeptID = 0
110     If Not rsTmp.EOF Then
112         lngDeptID = Val("" & rsTmp("使用小组id"))
        End If
114     If lngDeptID <= 0 Then
            '退出-给出提示
116         strErrInfo = "检验仪器未指定对应的检验小组！"
            Exit Function
        End If
118     dtSampleDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
120     strSQL = "Select ID, 姓名, 性别, 年龄, NO, 项目id, 结果, 标志, 结果参考, 紧急, 采样时间, 采样人, Rownum As 排列序号, 诊疗项目id," & vbNewLine & _
                "       编码,标本部位,开嘱科室ID,开嘱医生,标识号,当前床号,病人科室 " & vbNewLine & _
                "From (Select A.相关id As ID, C.姓名 || Decode(A.婴儿, 0, '', Null, '', '(婴儿)') As 姓名, A.性别, A.年龄, F.NO," & vbNewLine & _
                "              I.诊治项目id As 项目id, Decode(I.结果类型, 3, Nvl(I.默认值, '-'), 2, I.默认值, '') As 结果, '' As 标志," & vbNewLine & _
                "              Trim(Replace(Replace(' ' || Zlgetreference(I.诊治项目id, A.标本部位, Decode(A.性别, '男', 1, '女', 2, 0)," & vbNewLine & _
                "                                                          C.出生日期, Y.仪器id, A.年龄), ' .', '0.'), '～.', '～0.')) As 结果参考," & vbNewLine & _
                "              Nvl(A.紧急标志, 0) As 紧急, F.采样时间, F.采样人, G.排列序号, A.诊疗项目id, M.编码, " & vbNewLine & _
                "              a.标本部位,开嘱科室ID,开嘱医生,decode(a.病人来源,2, decode(nvl(c.住院号,''),'',c.门诊号,c.住院号),c.门诊号) as 标识号,c.当前床号,l.名称 as 病人科室 " & vbNewLine & _
                "       From 病人医嘱记录 A, 病人信息 C, 病人医嘱发送 F, 检验报告项目 G, 检验项目 I, 检验仪器项目 Y, 诊疗项目目录 M ,部门表 L " & vbNewLine & _
                "       Where A.诊疗类别 = 'C' And A.病人id = C.病人id And A.相关id Is Not Null And A.医嘱状态 = 8 And A.ID = F.医嘱id And" & vbNewLine & _
                "             A.诊疗项目id = G.诊疗项目id And G.细菌id Is Null And G.报告项目id = Y.项目id(+) And" & vbNewLine & _
                "             G.报告项目id = I.诊治项目id And A.诊疗项目id = M.ID(+) And a.病人科室ID = l.ID" & vbNewLine & _
                "             and (Y.仪器id + 0 = [1] Or (Y.仪器id Is Null And F.执行部门id = [2])) And nvl(F.执行状态,0) = 0  And A.相关ID = [3]" & vbNewLine & _
                "       Order By M.编码, G.排列序号)"

122     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, lngDeptID, lngID)
124     If rsTmp.EOF Then
126         strErrInfo = "没有找到检验申请！"
            Exit Function
        End If


        
128     If Val(strSampleNo) <= 0 Then
130         strErrInfo = "标本号错误，现只支持大于零的数字！"
            Exit Function
        Else
132         strSampleNo = Val(strSampleNo)
        End If
134     dStart = CDate(Format(dtSampleDate, "yyyy-MM-dd 00:00:00"))
136     dEnd = CDate(Format(dtSampleDate, "yyyy-MM-dd 23:59:59"))
138     strSQL = "Select 核收人,核收时间 from 检验标本记录 where 仪器ID=[1] and 标本序号=[2] And 核收时间 Between [3] and [4]"
140     Set rsItem = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngDeviceID, strSampleNo, dStart, dEnd)
142     If Not rsItem.EOF Then
144         strErrInfo = strSampleNo & "号标本已存在！" & vbNewLine & "核收人：" & rsItem!核收人 & " 核收时间:" & Format(rsItem!核收时间, "yyyy-MM-dd HH:mm:ss")
            Exit Function
        End If
    
146     strSQL = "Select B.病人id, B.主页id, B.序号, B.婴儿姓名, B.婴儿性别" & vbNewLine & _
                        "From 病人医嘱记录 A, 病人新生儿记录 B" & vbNewLine & _
                        "Where A.病人id = B.病人id And A.主页id = B.主页id And A.婴儿 = B.序号 And A.相关id = [1] And Rownum = 1"
148     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "zlLISRegister", lngID)
150     If rs.EOF = False Then
152         str姓名 = Trim("" & rs("婴儿姓名"))
154         str性别 = Trim("" & rs("婴儿性别"))
156         str年龄 = "婴儿"
        Else
158         str姓名 = Trim("" & rsTmp("姓名"))
160         str性别 = Trim("" & rsTmp("性别"))
162         str年龄 = Trim("" & rsTmp("年龄"))
        End If
    
        '读出检验项目
164     strSQL = "select distinct 医嘱内容 from 病人医嘱记录 a , 病人医嘱发送 b, 检验报告项目 c , 检验仪器项目 d " & vbNewLine & _
                  "  where a.id = b.医嘱ID and a.诊疗项目ID = c.诊疗项目ID and " & vbNewLine & _
                  "  c.报告项目ID = d.项目ID(+) and a.相关id=[1] "
166     Set rsItem = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取检验内容", lngID)
168     Do Until rsItem.EOF
170         strItem = strItem & " " & Trim("" & rsItem("医嘱内容"))
172         rsItem.MoveNext
        Loop
174     strItem = Trim(strItem) & "(" & Trim("" & rsTmp("标本部位")) & ")"
        
        '产生标本记录
        '------------10.25
176     lngKey = gobjComLib.zlDatabase.GetNextId("检验标本记录")
     
178     strSQL = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
            rsTmp("ID") & ",'" & _
            strSampleNo & "'," & _
            IIf(IsNull(rsTmp("采样时间")), "Null", "TO_DATE('" & Format(rsTmp("采样时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
            IIf(IsNull(rsTmp("采样人")), "Null", "'" & rsTmp("采样人") & "'") & "," & _
            lngDeviceID & "," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
            "1,'" & _
            gstrUserName & "'," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0,0,0," & _
            rsTmp("紧急") & ",NULL,'" & _
            str姓名 & "','" & str性别 & "','" & str年龄 & "','" & Trim("" & rsTmp("No")) & "','" & _
            Trim("" & rsTmp("标本部位")) & "'," & Trim("" & rsTmp("开嘱科室ID")) & ",'" & Trim("" & rsTmp("开嘱医生")) & "','" & _
            Trim("" & rsTmp("标识号")) & "','" & Trim("" & rsTmp("当前床号")) & "','" & Trim("" & rsTmp("病人科室")) & "','" & _
            strItem & "')"
    
        '---------- 10.26 的SQL
    
    '    gstrSql = "ZL_检验标本记录_标本核收(" & lngKey & "," & _
    '        rsTmp("ID") & ",'" & rsTmp("ID") & "',0,'" & _
    '        strSampleNo & "'," & _
    '        IIf(IsNull(rsTmp("采样时间")), "Null", "TO_DATE('" & Format(rsTmp("采样时间"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
    '        IIf(IsNull(rsTmp("采样人")), "Null", "'" & rsTmp("采样人") & "'") & "," & _
    '        lngDeviceID & "," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
    '        "'" & _
    '        gobjComLib.zlDatabase.GetUserInfo.Fields("姓名").value & "'," & _
    '        "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0," & _
    '        intType & ",NULL,'" & _
    '        str姓名 & "','" & str性别 & "','" & str年龄 & "','" & Trim("" & rsTmp("No")) & "','" & _
    '        Trim("" & rsTmp("标本部位")) & "'," & Trim("" & rsTmp("开嘱科室ID")) & ",'" & Trim("" & rsTmp("开嘱医生")) & "'," & _
    '        Trim("" & rsTmp("标识号")) & ",'" & Trim("" & rsTmp("当前床号")) & "','" & Trim("" & rsTmp("病人科室")) & "','" & _
    '        strItem & "',Null,Null,Null,'" & gstrUserCode & "','" & gstrUserName & "')"
    
        '-------------------------------------------------------------------------------------
    
180     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "生成条码标本"
                                                                
        '填写指标
182     strItemRecords = ""
184     Do While Not rsTmp.EOF
186         strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("项目ID") & "^" & _
                Trim("" & rsTmp("结果")) & "^" & Val("" & rsTmp("标志")) & "^" & Trim("" & rsTmp("结果参考")) & "^" & _
                Trim("" & rsTmp("诊疗项目ID")) & "^" & Trim("" & rsTmp("排列序号"))
            
188         rsTmp.MoveNext
        Loop
    
190     If Len(strItemRecords) > 0 Then
192         strItemRecords = Mid(strItemRecords, 2)
            
194         strSQL = "Zl_检验普通结果_Write(" & lngKey & "," & _
                lngDeviceID & ",'" & strItemRecords & "',0,0)"
196         gobjComLib.zlDatabase.ExecuteProcedure strSQL, "生成条码标本"
        End If
    
198     zlLISRegister = True
        Exit Function
errH:
        'Resume
200     strErrInfo = CStr(Erl()) & "," & Err.Description
202     WriteLog "zlLISRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnRegister(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
        '取消在ZLLIS中已核收的标本
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     If InStr(gstr接口权限, "ZLLIS取消核收") <= 0 Then
102         strErrInfo = "此接口未授权，不能调用！"
            Exit Function
        End If
        '是否可取消核收的操作在存储过程中，所以此处不做检查
104     strSQL = "Zl_检验标本记录_取消核收(" & lngID & ")"
106     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "取消核收"
108     zlLisUnRegister = True
        Exit Function
errH:
110     strErrInfo = CStr(Erl()) & "," & Err.Description
112     WriteLog "zlLisUnRegister " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function ZLLisInsterReport(ByVal lngID As Long, strItems As String, ByRef strErrInfo As String) As Boolean
        Dim str标本 As String, lng仪器ID As Long, str性别 As String, str生日  As String
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim str项目 As String, varItem As Variant
        On Error GoTo errH
100     If InStr(gstr接口权限, "ZLLIS标本审核") <= 0 Then
102         strErrInfo = "此接口未授权，不能调用！"
            Exit Function
        End If
104     If InStr(strItems, "'") > 0 Then
106         strErrInfo = "不允许包含单引号！"
            Exit Function
108     ElseIf InStr(strItems, """") > 0 Then
110         strErrInfo = "不允许包含双引号！"
            Exit Function
112     ElseIf InStr(strItems, "^") < 0 Then
114         strErrInfo = "请至少传入一个结果！"
            Exit Function
        End If
    
116     strSQL = "Select b.Id, b.审核人,b.性别, b.仪器id, b.标本类型, to_char(b.出生日期,'YYYY-MM-DD HH24:MI:SS') as 出生日期, b.微生物标本" & vbNewLine & _
                "From 病人医嘱记录 A, 检验标本记录 B" & vbNewLine & _
                "Where a.Id = b.医嘱id(+) And a.Id = [1]"

118     Set rsSample = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "读取检验内容", lngID)
    
120     If rsSample.EOF Then
122         strErrInfo = "未找到对应医嘱！"
            Exit Function
        End If
    
124     If Trim("" & rsSample!审核人) <> "" Then
126         strErrInfo = "已审核标本，不能修改！"
            Exit Function
        End If
    
128     If InStr(1, gstrPrivs, "审核标本") <= 0 Then
130         strErrInfo = "你没有权限进行审核,请重新登陆具有审核人员进行审核!"
            Exit Function
        End If
    
        '11210 权限“未收费审核”，在审核单个病人时，未生效，
132     If InStr(gstrPrivs, "未收费审核") <= 0 Then
134         strErrInfo = CheckChargeState(lngID, False)
136         If strErrInfo <> "" Then Exit Function
        End If
    
        '21137 已归档报告不能审核
138     gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.病人id = c.病人Id and B.主页id = C.主页ID "
140     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "审核检查", lngID)
142     If rsTmp.EOF = False Then
144         strErrInfo = "病人本次住院的病案已提交审查，不能进行审核！"
            Exit Function
        End If
    
        '检查住院病人是否出院后还有划价单
146     strErrInfo = CheckExesState(lngID)
148     If strErrInfo <> "" Then Exit Function

        '将编码转为项目ID
        Dim i As Integer, strCode As String, strValue As String
150     varItem = Split(strItems, "|")
152     str项目 = ""
154     strErrInfo = ""
156     For i = LBound(varItem) To UBound(varItem)
158         If InStr(varItem(i), "^") > 0 Then
160             strCode = Trim(Split(varItem(i), "^")(0))
162             strValue = Split(varItem(i), "^")(1)
            
164             gstrSql = "Select A.报告项目ID,B.编码, B.中文名, B.英文名 " & vbNewLine & _
                    " From 检验报告项目 A, 诊治所见项目 B " & vbNewLine & _
                    " Where A.报告项目id = B.ID And B.编码 = [1] "
    
166             Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "取得指标ID", strCode)
168             If rsTmp.EOF Then
170                 strErrInfo = strErrInfo & vbNewLine & strCode & " 未找到对应项目!"
                Else
172                 str项目 = str项目 & "|" & rsTmp!报告项目ID & "^" & strValue
                End If
            
            End If
        Next
174     If strErrInfo <> "" Then
            Exit Function
176     ElseIf str项目 = "" Then
178         strErrInfo = "没有要更新的数据！"
            Exit Function
        End If
180     str项目 = Mid(str项目, 2)
        '填结果
182     str性别 = Trim("" & rsSample!性别)
184     If str性别 = "男" Then
186         str性别 = "1"
188     ElseIf str性别 = "女" Then
190         str性别 = "2"
        Else
192         str性别 = "9"
        End If
194     strSQL = "ZL_检验普通结果_BATCHUPDATE(" & rsSample!ID & "," & _
                        rsSample!仪器ID & ",'" & Trim("" & rsSample!标本类型) & "'," & str性别 & "," & _
                        IIf(Trim("" & rsSample!出生日期) = "", "Null", "To_Date('" & Trim("" & rsSample!出生日期) & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                        str项目 & "'," & rsSample!微生物标本 & ")"
196     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "填写结果"

        '审核
198     strSQL = "ZL_检验标本记录_报告审核(" & rsSample!ID & ",'" & gstrUserName & "')"
200     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "审核报告"
202     ZLLisInsterReport = True
        Exit Function
errH:
204     ZLLisInsterReport = False
206     strErrInfo = CStr(Erl()) & "," & Err.Description
208     WriteLog "ZLLisInsterReport " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function zlLisUnAudit(ByVal lngID As Long, strErrInfo As String) As Boolean
        '取消审核
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsSample As ADODB.Recordset
        Dim d审核时间 As Date, dCurr As Date
        On Error GoTo errH
100     If InStr(gstr接口权限, "ZLLIS取消审核") <= 0 Then
102         strErrInfo = "此接口未授权，不能调用！"
            Exit Function
        End If
    
104     strSQL = "Select a.ID,a.打印次数, a.审核时间 From 检验标本记录 A Where 医嘱ID=[1]"
106     Set rsSample = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "取消审核检查", lngID)
    
108     If rsSample.EOF Then
110         strErrInfo = "未找到对应检验记录！"
            Exit Function
        End If
112     If IsNull(rsSample!审核时间) Then
114         strErrInfo = "标本未审核，不用取消审核！"
            Exit Function
        End If
    
116     If InStr(";" & gstrPrivs & ";", ";审核取消;") <= 0 Then
118         d审核时间 = rsTmp!审核时间
120         dCurr = gobjComLib.zlDatabase.Currentdate
122         If DateDiff("h", d审核时间, dCurr) > 24 Then
124             strErrInfo = "只能取消24小时内的审核报告单，请联系上级技师取消审核!"
                Exit Function
            End If
        End If
        '21434
126     If InStr(";" & gstrPrivs & ";", ";已审已打印可回滚;") <= 0 Then
128         If Val("" & rsSample!打印次数) > 0 Then
130             strErrInfo = "只能取消未打印的审核报告单，请联系上级技师取消审核!"
                Exit Function
            End If
        End If
        '21137 已归档报告不能取消
132     gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.医嘱ID=[1] " & vbNewLine & _
                " And b.病人id = c.病人Id and B.主页id = C.主页ID "
134     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "取消审核", lngID)
136     If rsTmp.EOF = False Then
138         strErrInfo = "病人本次住院的病案已提交审查，不能取消审核！"
            Exit Function
        End If
    
140     strSQL = "ZL_检验标本记录_审核取消(" & rsSample!ID & ")"
142     gobjComLib.zlDatabase.ExecuteProcedure strSQL, "取消审核"
144     zlLisUnAudit = True
        Exit Function
errH:
146     strErrInfo = CStr(Erl()) & " " & Err.Description
148     WriteLog "zlLisUnAudit " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Public Function GetAllDevice(ByRef strErrInfo As String) As String
        Dim strSQL As String, rsTmp As ADODB.Recordset
        On Error GoTo errH
100     strSQL = "Select ID,编码,名称 From 检验仪器"
102     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "取检验仪器")
104     GetAllDevice = ""
106     Do Until rsTmp.EOF
108         GetAllDevice = GetAllDevice & "|" & rsTmp!ID & "^" & rsTmp!编码 & "^" & rsTmp!名称
110         rsTmp.MoveNext
        Loop
112     If GetAllDevice <> "" Then GetAllDevice = Mid(GetAllDevice, 2)
114     If GetAllDevice = "" Then
116         strErrInfo = "没有初始化仪器！"
        End If
        Exit Function
errH:
118     strErrInfo = CStr(Erl()) & "," & " " & Err.Description
120     WriteLog "GetAllDevice " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Private Sub GetUserInfo()
    '功能:得到用户的信息

        Dim rsTemp As New ADODB.Recordset
        On Error GoTo errHand
100     glngUserId = 0
102     gstrUserCode = ""
104     gstrUserName = ""
106     gstrUserAbbr = ""
108     glngDeptId = 0
110     gstrDeptCode = ""
112     gstrDeptName = ""
    
114     Set rsTemp = gobjComLib.zlDatabase.GetUserInfo
    
116     Do Until rsTemp.EOF
118         glngUserId = Val("" & rsTemp.Fields("ID").value)               '当前用户id
120         gstrUserCode = "" & rsTemp.Fields("编号").value            '当前用户编码
122         gstrUserName = "" & rsTemp.Fields("姓名").value            '当前用户姓名
124         gstrUserAbbr = "" & rsTemp.Fields("简码").value          '当前用户简码
126         glngDeptId = Val("" & rsTemp.Fields("部门id").value)            '当前用户部门id
128         gstrDeptCode = "" & rsTemp.Fields("部门码").value        '当前用户
130         gstrDeptName = "" & rsTemp.Fields("部门名").value        '当前用户
    
132         rsTemp.MoveNext
        Loop
        Exit Sub
errHand:
134     WriteLog "GetUser " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
136     Err = 0
End Sub


Private Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As String
        '检验收费状态
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        Dim strSQLbak As String
        Dim intPatientType As Integer               '病人来源
        On Error GoTo errH
    
100     CheckChargeState = "单据未收费，不能进行审核！"
    
102     strSQL = "select 病人来源 from 检验标本记录 where id = [1]"
104     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "检验查费用", lngKey)
106     If rs.EOF = True Then Exit Function
108     intPatientType = rs("病人来源")
    
110     If blnOrder Then
112         strSQL = _
                "select NVL(A.记录状态,0) As 记录状态 " & _
                      "from 病人费用记录 A, " & _
                      "( " & _
                           "select No from 病人医嘱发送 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id))  " & _
                           "Union " & _
                           "select No from 病人医嘱附费 where 医嘱id IN (SELECT ID FROM 病人医嘱记录 WHERE [1] In (ID,相关id)) " & _
                      ") B " & _
                    "Where A.NO = B.NO "
    '        If intPatientType <> 2 Then
    '            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    '        End If
        Else
114         strSQL = _
                "select NVL(A.记录状态,0) As 记录状态 " & _
                      "from 病人费用记录 A, " & _
                      "( " & _
                           "select No,记录性质 from 病人医嘱发送 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                           "Union " & _
                           "select No,记录性质 from 病人医嘱附费 where 医嘱id IN (Select ID From 病人医嘱记录 A,(Select 医嘱id From 检验标本记录 Where ID= [1] Union Select 医嘱id From 检验项目分布 Where 标本id= [1]) B where B.医嘱id In (A.ID,A.相关id) and A.诊疗类别 = 'C' ) " & _
                      ") B " & _
                    "Where A.NO = B.NO and a.记录性质 = b.记录性质 "
    '        If intPatientType <> 2 Then
    '            strSQL = Replace(strSQL, "住院费用记录", "门诊费用记录")
    '        End If
        End If
    
116     strSQL = strSQL & " Order by 记录状态 "
    
118     Set rs = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)

120     If rs.BOF Then Exit Function
122     If rs("记录状态").value = 0 Then Exit Function
    
124     CheckChargeState = ""
        Exit Function
errH:
126     CheckChargeState = CStr(Erl()) & "," & Err.Description
128     WriteLog "CheckChargeState " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function

Private Function CheckExesState(lngKey As Long) As String
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '功能:      检查住院病人出院后是否还有划价单需要进行审核
        '参数       标本ID
        '返回       有划价单未审核 = Fasle 没有则 = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim rsTmp As New ADODB.Recordset
        On Error GoTo errH
100     CheckExesState = ""
    
        '81号系统不生效时不检查
        '    执行后自动审核划价单
        '    指当相应的医嘱执行后(包括药品发药)，0-不处理，1-自动审核划价单为记帐单。
102     If gobjComLib.zlDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
        '当前病人是否已出院或预出院
104     gstrSql = "select d.no" & vbNewLine & _
                "from (select distinct d.医嘱id" & vbNewLine & _
                "       from 检验标本记录 a, 病人信息 b, 病案主页 c, 检验项目分布 d" & vbNewLine & _
                "       where a.病人id = b.病人id and a.病人id = c.病人id and a.主页id = c.主页id and" & vbNewLine & _
                "             a.id = [1] and a.病人来源 = 2 and (b.出院时间 is not null or c.状态 = 3) and" & vbNewLine & _
                "             a.id = d.标本id) a, 病人医嘱记录 b, 病人医嘱发送 c, 病人费用记录 d" & vbNewLine & _
                "where a.医嘱id in (b.相关id, b.id) and b.id = c.医嘱id and c.记录性质 = d.记录性质 and" & vbNewLine & _
                "      c.no = d.no and d.记录性质 = 2 and d.记录状态 = 0 "
106     Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSql, "检验技师工作站-费用状态检查", lngKey)
    
108     If rsTmp.EOF Then
110         CheckExesState = ""
        Else
112         CheckExesState = "当前住院病人还有划价单未审核，但已出院或预出院！"
        End If
        Exit Function
errH:
114     CheckExesState = CStr(Erl()) & "," & Err.Description
116     WriteLog "CheckExesState " & CStr(Erl()) & "," & Err.Number & " " & Err.Description
End Function


Private Sub WriteLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）

    strFileName = App.Path & "\zlLisInterface_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not gobjFSO.FileExists(strFileName) Then Call gobjFSO.CreateTextFile(strFileName)
    Set objStream = gobjFSO.OpenTextFile(strFileName, ForAppending)
    
    objStream.WriteLine (strDate & ":" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub
