Attribute VB_Name = "mdlTAT"
Option Explicit

Private mstrItem As String

Public Function getTATTime(ByVal intType As Integer, ByVal strItem As String, _
                           ByVal strDept As String, ByVal strGroup As String, _
                           ByVal strMachine As String, ByVal strSex As String, _
                           intMsg As Integer, Optional strShowBef As String, _
                           Optional lngTATTime As Long, Optional strUser As String) As String
      '功能       检查TAT是否超时

      '入参
      'intType            '1=送检,2=签收,3=核收,4=审核
      'strItem            '申请项目ID  项目ID1,项目名称1,上个时间节点1,急诊1;项目ID2,项目名称2,上个时间节点2,急诊2;
      'strDept            '申请科室ID
      'strGroup           '申请小组ID
      'strMachine         '申请仪器ID
      'strSex             '病人性别
      'strUser            '操作员

      '出参
      'intMsg             '限制       1=只提示,2=提示并限制
      'strShowBef         '提示信息

      'GetTatTime=true表示超时,=false表示未超时

          Dim strSQL As String
          Dim Dtime As Date

          Dim var_tmp As Variant
          Dim var_ItemID As Variant
          Dim rsTAT As ADODB.Recordset
          Dim rsTATMX As ADODB.Recordset
          Dim rsOldItems As ADODB.Recordset
          Dim rsNewItems As ADODB.Recordset
          Dim strOldItemID As String
          Dim strOldItemCode As String
          Dim strNewItemID As String
          Dim strMsgShow As String
          Dim strOldMid As String
          Dim strNewMid As String
          Dim var_MidOld As Variant
          Dim var_MidNew As Variant
          Dim blnFind As Boolean
          Dim lngTATTimeBefor As Long
          Dim i As Integer, J As Integer

1         On Error GoTo getTATTime_Error

2         mstrItem = strItem
3         lngTATTime = 0

4         Select Case intType
          Case 1
5             strMsgShow = "未采样标本不能送检"
6         Case 2
7             strMsgShow = "未送检标本不能登记"
8         Case 3
9             strMsgShow = "未登记标本不能核收"
10        Case 4
11            strMsgShow = "未核收标本不能审核"
12        End Select
13        var_tmp = Split(strItem, ";")

          '项目对码
14        If intType = 1 Or intType = 2 Then
              '送检和签收时需要对项目进行对码
15            strOldItemID = ""
16            For i = LBound(var_tmp) To UBound(var_tmp)
17                strOldItemID = strOldItemID & "," & Split(var_tmp(i), ",")(0)
18            Next
19            If strOldItemID <> "" Then strOldItemID = Mid(strOldItemID, 2)
              '根据老板项目ID查询老板项目编码
20            strSQL = "Select /*+cardinality(b,10)*/ A.id,A.编码 From 诊疗项目目录 A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.Id = b.Column_Value"
21            Set rsOldItems = ComOpenSQL(Sel_His_DB, strSQL, "老版项目编码", strOldItemID)
22            strOldItemCode = ""
23            strOldMid = ""
24            Do While rsOldItems.EOF = False
25                strOldItemCode = strOldItemCode & "," & rsOldItems("编码") & ""
26                strOldMid = strOldMid & ";" & rsOldItems("ID") & "," & rsOldItems("编码") & ""
27                rsOldItems.MoveNext
28            Loop
29            If strOldItemCode <> "" Then strOldItemCode = Mid(strOldItemCode, 2)
30            If strOldMid <> "" Then strOldMid = Mid(strOldMid, 2)
              '根据老板项目编码查询新版项目ID
31            If gUserInfo.NodeNo <> "-" Then
32                strSQL = "Select /*+cardinality(b,10)*/ A.ID,A.诊疗编码 From 检验组合项目 A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.诊疗编码 = b.Column_Value and (a.站点=[2] or a.站点 is null)"
33            Else
34                strSQL = "Select /*+cardinality(b,10)*/ A.ID,A.诊疗编码 From 检验组合项目 A, Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) B Where a.诊疗编码 = b.Column_Value"
35            End If
36            Set rsNewItems = ComOpenSQL(Sel_Lis_DB, strSQL, "新版项目ID", strOldItemCode, gUserInfo.NodeNo)
37            strNewMid = ""
38            Do While rsNewItems.EOF = False
39                strNewItemID = strNewItemID & "," & rsNewItems("ID")
40                strNewMid = strNewMid & ";" & rsNewItems("ID") & "," & rsNewItems("诊疗编码") & ""
41                rsNewItems.MoveNext
42            Loop
43            If strNewItemID <> "" Then
44                strNewItemID = Mid(strNewItemID, 2)
45                If strNewMid <> "" Then strNewMid = Mid(strNewMid, 2)

46                var_ItemID = Split(strNewItemID, ",")
47                var_MidOld = Split(strOldMid, ";")
48                var_MidNew = Split(strNewMid, ";")
49                For J = LBound(var_MidOld) To UBound(var_MidOld)
50                    For i = LBound(var_MidNew) To UBound(var_MidNew)
51                        If Split(var_MidNew(i), ",")(1) = Split(var_MidOld(J), ",")(1) Then
52                            strItem = Replace(strItem, Split(var_MidOld(J), ",")(0) & ",", Split(var_MidNew(i), ",")(0) & ",")
53                        End If
54                    Next
55                Next

56                If strItem <> "" Then
57                    var_tmp = Split(strItem, ";")
58                End If
59            End If
60        End If


61        For i = LBound(var_tmp) To UBound(var_tmp)
62            blnFind = False
              '根据传入项目查询TAT时间
63            strSQL = "Select Distinct a.Id, a.送检限时, a.签收限时, a.核收限时," & _
                     " a.审核限时, a.应用科室,a.应用小组, a.应用仪器, a.性别," & _
                     " a.急诊, a.限制, a.提示信息 From 检验tat时间 A, 检验tat时间明细 B" & _
                     " Where a.Id = b.Tat时间id And a.是否有效 = 1 and b.申请项目id = [1] and a.急诊=[2]"
64            Select Case intType
              Case 1
65                strSQL = strSQL & " and a.送检限时 is  not null"
66            Case 2
67                strSQL = strSQL & " and a.签收限时 is  not null"
68            Case 3
69                strSQL = strSQL & " and a.核收限时 is  not null"
70            Case 4
71                strSQL = strSQL & " and a.审核限时 is  not null"
72            End Select
73            Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "检验TAT时间", Split(var_tmp(i), ",")(0), Split(var_tmp(i), ",")(3))
              '根据查询出来的TAT时间ID来查询相关的TAT时间明细表
74            Do While rsTAT.EOF = False
75                blnFind = True
76                If Not IsNull(rsTAT("ID")) Then
77                    strSQL = "Select f_List2str(Cast(Collect(To_Char(申请科室)) As t_Strlist)) 申请科室," & _
                             " f_List2str(Cast(Collect(To_Char(检验小组id)) As t_Strlist)) 检验小组id," & _
                             " f_List2str(Cast(Collect(To_Char(检验仪器id)) As t_Strlist)) 检验仪器id" & _
                             " From 检验tat时间明细 Where Tat时间id =[1]"
78                    Set rsTATMX = ComOpenSQL(Sel_Lis_DB, strSQL, "检验TAT时间明细", Val(rsTAT("ID")))
                      '记录是提示还是提示并禁止
                      '                intMsg = rsTAT("限制")
79                    strShowBef = rsTAT("提示信息") & ""

                      '急诊
80                    If Split(var_tmp(i), ",")(3) <> 1 And Val(rsTAT("急诊") & "") = 1 Then
81                        blnFind = False
82                    End If
                      '科室
83                    If Val(rsTAT("应用科室") & "") = 2 Then
84                        If InStr("," & rsTATMX("申请科室") & ",", IIf(strDept = "", ",strDept,", "," & strDept & ",")) <= 0 Then
85                            blnFind = False
86                        End If
87                    End If
                      '性别
88                    If rsTAT("性别") & "" <> strSex And rsTAT("性别") & "" <> "所有" And Not IsNull(rsTAT("性别")) Then
89                        blnFind = False
90                    End If

91                    If intType = 3 Or intType = 4 Then
                          '小组
92                        If Val(rsTAT("应用小组") & "") = 2 Then
93                            If rsTATMX("检验小组id") & "" <> "" And InStr("," & rsTATMX("检验小组id") & ",", IIf(strGroup = "", ",strGroup,", "," & strGroup & ",")) <= 0 Then
94                                blnFind = False
95                            End If
96                        End If
                          '仪器
97                        If Val(rsTAT("应用仪器") & "") = 2 Then
98                            If InStr("," & rsTATMX("检验仪器ID") & ",", IIf(strMachine = "", ",strMachine,", "," & strMachine & ",")) <= 0 Then
99                                blnFind = False
100                           End If
101                       End If
102                   End If

103                   If Split(var_tmp(i), ",")(2) = "" Then
104                       Dtime = CDate(Format("2000/01/01 01:01:01", "yyyy/mm/dd hh:mm:ss"))
105                   Else
106                       Dtime = CDate(Split(var_tmp(i), ",")(2))
107                   End If
108                   If blnFind = True Then Exit Do
109               End If
110               rsTAT.MoveNext
111           Loop

112           If blnFind = True Then

113               Call GetMsgItems(Dtime, IIf(intType = 1, Val(rsTAT("送检限时") & ""), _
                                              IIf(intType = 2, Val(rsTAT("签收限时") & ""), _
                                                  IIf(intType = 3, Val(rsTAT("核收限时") & ""), _
                                                      IIf(intType = 4, Val(rsTAT("审核限时") & ""), 0)))), _
                                                      Split(var_tmp(i), ",")(1), strShowBef, CInt(rsTAT("限制")), intType, strUser)

114           End If
              '核收时需要返回审核TAT限时
115           If intType = 3 Then
                  '写入医嘱TAT各阶段的设置时间
116               Call setTATAllTime(Val(Split(var_tmp(i), ",")(0)), Val(Split(var_tmp(i), ",")(5)), Val(Split(var_tmp(i), ",")(4)), strSex, Split(var_tmp(i), ",")(3), strDept, strGroup, strMachine)
                  
                  '核收时获取倒计时
117               lngTATTimeBefor = GetTATAllTimefun(Val(Split(var_tmp(i), ",")(4)), Val(Split(var_tmp(i), ",")(5)))
118               If lngTATTime = 0 And lngTATTimeBefor <> 0 Then
119                   lngTATTime = lngTATTimeBefor
120               ElseIf lngTATTime > lngTATTimeBefor Then
121                   lngTATTime = lngTATTimeBefor
122               End If
123           End If
124       Next

125       getTATTime = mstrItem


126       Exit Function
getTATTime_Error:
127       Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "执行(getTATTime)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
128       Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/12/11
'功    能:获取TAT倒计时
'入    参:
'           lngApplyID      检验申请组合.申请ID
'           lngAdviceID     检验申请组合.医嘱ID
'出    参:
'           lngTATTime      TAT倒计时
'返    回:
'---------------------------------------------------------------------------------------
Public Function GetTATAllTimefun(ByVal lngAdviceID As Long, ByVal lngApplyID As Long) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strTATBegin As String
          Dim intTATBegin As Integer
          
          '读取参数
1         On Error GoTo GetTATAllTimefun_Error
2         If lngAdviceID = 0 Or lngApplyID = 0 Then Exit Function
          
3         strTATBegin = ComGetPara(Sel_Lis_DB, "TAT倒计时起始点", gSysInfo.SysNo, gSysInfo.ModlNo)
4         If strTATBegin = "" Then
5             intTATBegin = 3
6         Else
7             intTATBegin = Val(strTATBegin)
8         End If
          
          '获取倒计时
9         Select Case intTATBegin
              Case 0
10                strSQL = "select 送检限时+签收限时+核收限时+审核限时 TAT剩余时间 from 检验申请组合 where 申请ID=[1] and 医嘱ID=[2]"
11            Case 1
12                strSQL = "select 签收限时+核收限时+审核限时 TAT剩余时间 from 检验申请组合 where 申请ID=[1] and 医嘱ID=[2]"
13            Case 2
14                strSQL = "select 核收限时+审核限时 TAT剩余时间 from 检验申请组合 where 申请ID=[1] and 医嘱ID=[2]"
15            Case 3
16                strSQL = "select 审核限时 TAT剩余时间 from 检验申请组合 where 申请ID=[1] and 医嘱ID=[2]"
17        End Select
18        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngApplyID, lngAdviceID)
19        If rsTmp.RecordCount > 0 Then
20            GetTATAllTimefun = Val(rsTmp("TAT剩余时间") & "")
21        End If


22        Exit Function
GetTATAllTimefun_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "执行(GetTATAllTimefun)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
24        Err.Clear
End Function

Private Function GetMsgItems(ByVal Dtime As Date, ByVal lngTime As Long, _
                            ByVal strItem As String, ByVal strShowBef As String, _
                            ByVal intMsg As Integer, ByVal intType As Integer, ByVal strUser As String) As String
          '返回没有超时的项目
          'Dtime          上一个时间节点
          'lngTime        tat限时
          'strItem        项目
          'strShowBef     提示信息
          '提示控制        0=只写日志,1=提示并写日志,2=写日志并禁止
          'intType        来源 1=送检,2=签收,3=核收,4=审核
          
          '返回字符串格式
              '组合ID,组合名称,上一个时间节点,是否急诊,医嘱ID,相关ID,条码,超时时间,提示信息,提示控制
          Dim dCurrentdate As Date
          Dim var_tmp As Variant
          Dim var_tmp1 As Variant
          Dim strFrom As String
          Dim strSQL As String
          Dim i As Integer
              
1         On Error GoTo GetMsgItems_Error

2         dCurrentdate = Currentdate
          
3         strShowBef = Replace(strShowBef, ",", "")
          
4         If DateDiff("n", Dtime, dCurrentdate) > lngTime Then
          
              '写入检验操作日志
5             Select Case intType
                  Case 1
6                     strFrom = "TAT送检超时"
7                 Case 2
8                     strFrom = "TAT签收超时"
9                 Case 3
10                    strFrom = "TAT核收超时"
11                Case 4
12                    strFrom = "TAT审核超时"
13            End Select
14            strSQL = "Zl_检验操作日志_Insert(19,6,'" & strUser & "',null,'" & strFrom & "','医嘱ID" & Split(mstrItem, ",")(4) & "|" & Replace(Replace(strShowBef, "[项目]", strItem), "[超时]", DateDiff("n", Dtime, dCurrentdate) - lngTime) & "分钟')"
15            ComExecuteProc Sel_Lis_DB, strSQL, "检验操作日志"
              
16            If strItem <> "" Then
17                var_tmp = Split(mstrItem, ";")
18                For i = LBound(var_tmp) To UBound(var_tmp)
19                    var_tmp1 = Split(var_tmp(i), ",")
20                    If InStr(var_tmp(i), strItem) > 0 Then
21                        If UBound(var_tmp1) > 8 Then
22                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp1(0) & "," & var_tmp1(1) & "," & var_tmp1(2) & "," & _
                                          var_tmp1(3) & "," & var_tmp1(4) & "," & var_tmp1(5) & "," & var_tmp1(6) & "," & (DateDiff("n", Dtime, dCurrentdate) - lngTime) & "," & strShowBef & "," & intMsg)
23                        Else
24                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp(i) & "," & DateDiff("n", Dtime, dCurrentdate) - lngTime & "," & strShowBef & "," & intMsg)
25                        End If
26                    Else
27                        If UBound(var_tmp1) < 9 Then
28                            mstrItem = Replace(mstrItem, var_tmp(i), var_tmp(i) & ",0" & "," & strShowBef & "," & intMsg)
29                        End If
30                    End If
31                Next
              
32            End If
33        End If


34        Exit Function
GetMsgItems_Error:
35        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "执行(GetMsgItems)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
36        Err.Clear
          
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/12/11
'功    能:写入医嘱TAT各阶段的设置时间
'入    参:
'           lngItemID           组合项目ID
'           lngApplyID          检验申请组合.申请ID
'           lngAdivceID         检验申请组合.医嘱ID
'           strSex              性别
'           strJiZhen           是否急诊 0=否，1=是
'           strDept             申请科室
'           strGroup            检验小组
'           strMachine          检验仪器
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function setTATAllTime(ByVal lngItemid As Long, ByVal lngApplyID As Long, ByVal lngAdivceID As Long, ByVal strSex As String, ByVal strJiZhen As String, _
                               ByVal strDept As String, ByVal strGroup As String, ByVal strMachine As String)
          Dim strSQL As String
          Dim rsTAT As ADODB.Recordset
          Dim rsTATMX As ADODB.Recordset
          Dim lngSJXS As Long             '送检限时
          Dim lngQSXS As Long             '签收限时
          Dim lngHSXS As Long             '核收限时
          Dim lngSHXS As Long             '审核限时
          
          Dim blnTATTime As Boolean


1         On Error GoTo setTATAllTime_Error

2         If lngItemid = 0 Then Exit Function
3         If lngApplyID = 0 Or lngAdivceID = 0 Then Exit Function
          
'          '检查医嘱是否已经写入TAT各时间，如果已写入，则不再重复写入
'4         strSQL = "Select Count(*) 数量" & vbCrLf & _
'                   " From 检验申请组合 " & vbCrLf & _
'                   " Where 申请id = [1] And 医嘱id = [2] And 送检限时 Is not Null And 签收限时 Is not Null And 核收限时 Is not Null And 审核限时 Is not Null"
'5         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请组合", lngApplyID, lngAdivceID)
'6         If rsTmp.RecordCount > 0 Then
'7             If Val(rsTmp("数量") & "") > 0 Then Exit Function
'8         Else
'9             Exit Function
'10        End If
          
          '查询项目相关的TAT规则
11        strSQL = "Select a.id, Decode(Sign(Length(a.送检限时)), 1, '送检限时',Decode(Sign(Length(a.签收限时)), 1, '签收限时'," & _
                   " Decode(Sign(Length(a.核收限时)), 1, '核收限时', Decode(Sign(Length(a.审核限时)), 1, '审核限时', Null)))) 时间类型," & _
                   " Decode(Sign(Length(a.送检限时)), 1, a.送检限时,Decode(Sign(Length(a.签收限时)), 1, a.签收限时," & _
                   " Decode(Sign(Length(a.核收限时)), 1, a.核收限时, Decode(Sign(Length(a.审核限时)), 1, a.审核限时, Null)))) 时间," & _
                   " A.应用科室 , A.应用小组, A.应用仪器,a.急诊,a.性别 From 检验tat时间 A, 检验tat时间明细 B Where a.Id = b.Tat时间id And a.是否有效 = 1" & _
                   " And b.申请项目id = [1] And (a.性别 = '所有' or a.性别 is null Or a.性别 = [2]) and a.急诊=[3]"

12        Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "计算tat时间", lngItemid, strSex, strJiZhen)
          
          '如果通过性别、急诊没有查询到，则只通过ID去查询
13        If rsTAT.RecordCount < 1 Then
14            strSQL = "Select a.id, Decode(Sign(Length(a.送检限时)), 1, '送检限时',Decode(Sign(Length(a.签收限时)), 1, '签收限时'," & _
                   " Decode(Sign(Length(a.核收限时)), 1, '核收限时', Decode(Sign(Length(a.审核限时)), 1, '审核限时', Null)))) 时间类型," & _
                   " Decode(Sign(Length(a.送检限时)), 1, a.送检限时,Decode(Sign(Length(a.签收限时)), 1, a.签收限时," & _
                   " Decode(Sign(Length(a.核收限时)), 1, a.核收限时, Decode(Sign(Length(a.审核限时)), 1, a.审核限时, Null)))) 时间," & _
                   " A.应用科室 , A.应用小组, A.应用仪器,a.急诊,a.性别 From 检验tat时间 A, 检验tat时间明细 B Where a.Id = b.Tat时间id And a.是否有效 = 1" & _
                   " And b.申请项目id = [1]"
15             Set rsTAT = ComOpenSQL(Sel_Lis_DB, strSQL, "计算tat时间", lngItemid, strSex)
16        End If
               
          '筛选记录，获取正确的TAT限时
17        Do While rsTAT.EOF = False
18            strSQL = "Select f_List2str(Cast(Collect(To_Char(申请科室)) As t_Strlist)) 申请科室," & _
                      " f_List2str(Cast(Collect(To_Char(检验小组id)) As t_Strlist)) 检验小组id," & _
                      " f_List2str(Cast(Collect(To_Char(检验仪器id)) As t_Strlist)) 检验仪器id" & _
                      " From 检验tat时间明细 Where Tat时间id =[1]"
19            Set rsTATMX = ComOpenSQL(Sel_Lis_DB, strSQL, "检验TAT时间明细", Val(rsTAT("ID")))
20            blnTATTime = True
              '急诊
21            If strJiZhen <> 1 And Val(rsTAT("急诊") & "") = 1 Then
22                blnTATTime = False
23            End If
              '科室
24            If Val(rsTAT("应用科室") & "") = 2 Then
25                If InStr("," & rsTATMX("申请科室") & ",", IIf(strDept = "", ",strDept,", "," & strDept & ",")) <= 0 Then
26                    blnTATTime = False
27                End If
28            End If
              '性别
29            If rsTAT("性别") & "" <> strSex And rsTAT("性别") & "" <> "所有" And Not IsNull(rsTAT("性别")) Then
30                blnTATTime = False
31            End If
              
32            If rsTAT("时间类型") & "" = "核收限时" Or rsTAT("时间类型") & "" = "审核限时" Then
                  '小组
33                If Val(rsTAT("应用小组") & "") = 2 Then
34                    If rsTATMX("检验小组id") & "" <> "" And InStr("," & rsTATMX("检验小组id") & ",", IIf(strGroup = "", ",strGroup,", "," & strGroup & ",")) <= 0 Then
35                        blnTATTime = False
36                    End If
37                End If
                  '仪器
38                If Val(rsTAT("应用仪器") & "") = 2 Then
39                    If InStr("," & rsTATMX("检验仪器ID") & ",", IIf(strMachine = "", ",strMachine,", "," & strMachine & ",")) <= 0 Then
40                        blnTATTime = False
41                    End If
42                End If
43            End If
              
              '写入数据库
44            If blnTATTime = True Then
45                If rsTAT("时间类型") & "" = "送检限时" Then
46                    lngSJXS = Val(rsTAT("时间") & "")
47                ElseIf rsTAT("时间类型") & "" = "签收限时" Then
48                    lngQSXS = Val(rsTAT("时间") & "")
49                ElseIf rsTAT("时间类型") & "" = "核收限时" Then
50                    lngHSXS = Val(rsTAT("时间") & "")
51                ElseIf rsTAT("时间类型") & "" = "审核限时" Then
52                    lngSHXS = Val(rsTAT("时间") & "")
53                End If
54            End If
55            rsTAT.MoveNext
56        Loop
          
57        If lngSJXS <> 0 Or lngQSXS <> 0 Or lngHSXS <> 0 Or lngSHXS <> 0 Then
58            strSQL = "Zl_检验申请组合_Tat限时(" & lngApplyID & "," & lngAdivceID & "," & lngSJXS & "," & lngQSXS & "," & lngHSXS & "," & lngSHXS & ")"
59            Call ComExecuteProc(Sel_Lis_DB, strSQL, "TAT限时")
60        End If
          
61        Exit Function
setTATAllTime_Error:
62        Call WriteErrLog("zlPublicHisCommLis", "mdlTAT", "执行(setTATAllTime)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
63        Err.Clear

End Function

