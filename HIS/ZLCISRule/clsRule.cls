VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitCISRule(cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal lngProgram As Long) As Boolean
          Dim strFilePath As String

1         On Error GoTo ErrH

2         gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
3         gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
4         glngSys = lngSys
5         Set gcnOracle = cnMain
6         Call GetUserInfo
          '如果本地没有日志文件新建日志文件
7         strFilePath = App.Path & "\" & App.EXEName & ".ini"
8         If Dir(strFilePath) = "" Then
              '[DEBUG]
              'OutPutDebugInfo = 0 ;输出调试信息到zlCISRule.log文件，0-不输出,1-输出所有接口参数信息
9             WriteIni "DEBUG", "OutPutDebugInfo", "0 ;输出调试信息到" & App.EXEName & ".log文件,0-不输出,1-输出", strFilePath
               
10        End If
          
11        Call InitObjLis(lngProgram)
          
12        gblnLog = (Val(ReadIni("Debug", "OutPutDebugInfo", strFilePath)) = 1)
          
13        InitCISRule = True

14        Exit Function

ErrH:
15        MsgBox "在zlCISRule.clsRule.InitCISRule的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Public Function SubmitMain(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
    ByVal objMain As Object, ByVal bytScene As Byte, Optional colDiag As Collection) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : SubmitMain
      ' Author    : YWJ
      ' Date      : 2019-09-20 17:16:31
      ' Parameter :
      '             lngPatiID -病人ID
      '             lngVisitId -主页ID
      '             strRegNO -挂号单号
      '             objMain-门诊\住院医嘱下达传入医嘱记录集;门诊\住院医生工作站下达申请单 传入主体集合
      '             bytScene-场合 1-门诊\住院医嘱下达;
      '                           21:检查申请,22-检验申请,23-输血申请,24-手术申请;25-会诊申请;
      '                           3-住院医生首页诊断\门诊病人信息录入诊断
      '                           4-标本采集
      '             colDiag-门诊编辑下达传入
      ' Purpose   : 提交主体信息
      ' Return    : T-禁止操作;F-继续操作
      '---------------------------------------------------------------------------------------
          Dim strJsonIn As String
          Dim strJsonOut As String
          
          Dim strPati As String
          Dim strBusinss As String
          
          Dim strMainInfo As String

          Dim lngRegId As Long    '挂号ID
          Dim bytType As Byte
              
          Dim blnRet As Boolean
          
          Dim objJson As Object
          
          Dim colList As Collection

          '--日志
1         On Error GoTo ErrH
2         WriteLog "函数：SubmitMain 开始" & vbNewLine & _
                   "入参：病人ID=" & lngPatiID & ",主页ID=" & lngVisitId & ",挂号单=" & strRegNo & _
                   ",主体类型=" & TypeName(objMain) & ",场合=" & bytScene & ",诊断=" & (Not colDiag Is Nothing) & vbNewLine
3         bytType = bytScene
4         If InStr(",21,22,23,24,25,", "," & bytScene & ",") Then
5           bytType = 1
6         ElseIf bytScene = 3 Then
7           bytType = 2 '单独传诊断
8         End If
          '获取病人信息
9         strPati = GetPatiInfo(lngPatiID, lngVisitId, strRegNo, bytType, lngRegId)

          '获取医嘱信息
10        If TypeName(objMain) = "Recordset" Then
11            strMainInfo = GetMainInfo(lngPatiID, lngVisitId, strRegNo, lngRegId, objMain, colDiag)
12        ElseIf TypeName(objMain) = "Collection" Then
13            strMainInfo = GetMainJson(objMain)
14        End If
          
15        strJsonIn = "{""接口json_in"":""{\""cdss_in\"":{" & strPati
16        If strMainInfo <> "" Then
17            strJsonIn = strJsonIn & "," & strMainInfo & "}}""}"
18        Else
19            strJsonIn = strJsonIn & "}}""}"
20        End If
          
          '测试YWJ
          'strJsonIn = TestJson(4)
          
21        If SubmitMainInfo(strJsonIn, strJsonOut) Then
              'strJsonOut = TestJson(5)
22            If strJsonOut = "" Then Exit Function
23            Set objJson = mdlJSON.parse(strJsonOut) 'Dictionary
24            If objJson Is Nothing Then
25                MsgBox "【提交主体信息】返回值解析失败！" & vbCrLf & _
                         " 原因:" & mdlJSON.GetParserErrors(), vbExclamation, gstrSysName
26                Exit Function
27            End If
              
28            If TypeName(objJson) = "Dictionary" Then
29                If objJson.Exists("inquiry") Then
                      '存在问诊
30                    Set colList = objJson.Item("inquiry") 'Collection
31                    strJsonOut = ""
32                    If colList.Count > 0 Then
33                        If objJson.Exists("businss") Then
34                            strBusinss = objJson.Item("businss")
35                        End If
36                        If frmInquiryInfo.ShowMe(colList, strJsonOut) Then
37                            strJsonIn = "{""问诊json_in"":""{\""businss\"":\""" & strBusinss & "\"""
38                            If strJsonOut <> "" Then
39                                strJsonIn = strJsonIn & "," & strJsonOut & "}""}"
40                            Else
41                                strJsonIn = strJsonIn & "}""}"
42                            End If
      '                        Call SubmitMainByAdviceID("2187620,2187621")
                              
43                            If SubmitInquiriyInfo(strJsonIn, strJsonOut) Then
44                                If strJsonOut = "" Then Exit Function
45                                Set objJson = mdlJSON.parse(strJsonOut) 'Dictionary
46                                If objJson Is Nothing Then
47                                    MsgBox "【提交问诊信息】返回值解析失败！" & vbCrLf & _
                                             " 原因:" & mdlJSON.GetParserErrors(), vbExclamation, gstrSysName
48                                    Exit Function
49                                End If
50                            End If
51                        End If
52                    End If
53                End If
                  '存在规则干预时给予提示
54                If objJson.Exists("messages") Then
55                    Set colList = objJson.Item("messages")
56                    If colList.Count > 0 Then
57                        If HandleMessage(colList) Then blnRet = True
58                    End If
59                End If
60            Else
61                MsgBox "【提交主体信息】返回值格式有误，无法正常解析！", vbExclamation, gstrSysName
62                Exit Function
63            End If
64        End If
65        SubmitMain = blnRet
66        WriteLog "函数：SubmitMain 结束" & vbNewLine
67        Exit Function
ErrH:
68        MsgBox "在zlCISRule.clsRule.SubmitMain的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function
 
Public Function SubmitMainByAdviceID(ByVal strAdviceIds As String) As Boolean
      '---------------------------------------------------------------------------------------
      ' Procedure : SubmitMainByAdviceID
      ' Author    : YWJ
      ' Date      : 2019-09-20 17:16:31
      ' Parameter :
      '             strAdviceIds-医嘱IDs
      ' Purpose   : 用于标本采集提交主体信息
      ' Return    : T-禁止操作;F-继续操作
      '---------------------------------------------------------------------------------------

          Dim strSQL As String
          Dim rsAdvice As ADODB.Recordset
          
1         On Error GoTo ErrH
2         WriteLog "函数：SubmitMainByAdviceID 开始" & vbNewLine & _
                   "入参：" & strAdviceIds & vbNewLine
3         strSQL = "Select /*+cardinality(b,10)*/" & vbCrLf & _
                   "    a.ID, Nvl(a.相关id, 0) As 相关id, a.病人id, a.主页id, a.挂号单, a.序号, a.诊疗类别, a.医嘱期效, a.诊疗项目id, a.标本部位, a.检查方法, a.医嘱内容,a.开嘱医生, a.总给予量, a.单次用量, a.执行频次" & vbCrLf & _
                   "   From 病人医嘱记录 A, Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b" & vbCrLf & _
                   "   Where a.id = b.Column_Value" & vbCrLf & _
                   "   Order By a.序号"
          
4         Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "SubmitMainByAdviceID", strAdviceIds)
          
5         If rsAdvice.EOF Then Exit Function
          
6         SubmitMainByAdviceID = SubmitMain(CLng(rsAdvice!病人ID), Val(rsAdvice!主页ID & ""), rsAdvice!挂号单 & "", rsAdvice, 4)
7         WriteLog "函数：SubmitMainByAdviceID 结束" & vbNewLine
                   
8         Exit Function

ErrH:
9         MsgBox "在zlCISRule.clsRule.SubmitMainByAdviceID的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

