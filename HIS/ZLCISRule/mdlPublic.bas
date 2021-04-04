Attribute VB_Name = "mdlPublic"

Option Explicit
'API
'获得鼠标指针在屏幕坐标上的位置
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'获得窗口在屏幕坐标中的位置
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'判断指定的点是否在指定的矩形内部
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'准备用来使窗体始终在最前面
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'用来移动窗体
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'获取窗体状态
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'HWND hwnd, // 指定分层窗口句柄
'COLORREF crKey, // 指定需要透明的背景颜色值，可用RGB()宏
'BYTE bAlpha, // 设置透明度，0表示完全透明，255表示不透明
'DWORD dwFlags // 透明方式
'       其中，dwFlags参数可取以下值：
'       LWA_ALPHA=&H2时：crKey参数无效，bAlpha参数有效；
'       LWA_COLORKEY=&H1：窗体中的所有颜色为crKey的地方将变为透明，bAlpha参数无效其常量值为1
'       LWA_ALPHA | LWA_COLORKEY：crKey的地方将变为全透明，而其它地方根据bAlpha参数确定透明度
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'自定义数据类型
Public Type TYPE_USER_INFO
    id As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    专业技术编码 As String
    用药级别 As Long
End Type

Public UserInfo As TYPE_USER_INFO

'常量
Public Const GWL_WNDPROC = -4&
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const SWP_NOACTIVATE = &H10 '不激活窗体
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const WS_EX_TOPMOST As Long = &H8
Public Const HWND_TOPMOST As Long = -1
Public Const SW_SHOWMAXIMIZED = 3
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21

Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)

'共有变量
Public glngOldWindowProc As Long '用来保存系统默认的窗口消息处理函数的地址
Public glngSys As Long

Public gobjLIS As Object    '

Public gstrSysName As String                '系统名称
Public gstrProductName As String
Public gstrUnitName As String
Public gblnLog As Boolean         'T-开启日志功能;F-关闭日志功能
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例


Public Function GetXMLResult(ByVal rsRec As ADODB.Recordset)
'功能:构造反向问诊响应XML
    Dim i As Long
    Dim strXML As String
    For i = 1 To rsRec.RecordCount
        strXML = strXML & "    <info name=""" & rsRec!Name & """ type=""" & rsRec!Type & """ index=""" & _
            rsRec!index & """ value=""" & rsRec!Default & """ obsid=""" & rsRec!Obsid & """/>" & vbNewLine
        rsRec.MoveNext
    Next
    GetXMLResult = Replace(strXML, """", "\""")
End Function


'自定义的消息处理函数
Public Function NewWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'功能:捕获滚轮事件进行处理,非滚轮事件调用默认窗口消息处理函数
'参数:vsc-VScrollBar 对象
'     OldWindowProc 默认窗口消息处理函数地址
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        '对鼠标滚轮事件进行处理
        If wParam = -7864320 Then '向下滚动
            If frmInquiryInfo.vsc.Value - 10 < frmInquiryInfo.vsc.Max Then
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Max
            Else
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Value - 10
            End If
        ElseIf wParam = 7864320 Then '向上滚动
            If frmInquiryInfo.vsc.Value + 10 > frmInquiryInfo.vsc.Min Then
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Min
            Else
                frmInquiryInfo.vsc.Value = frmInquiryInfo.vsc.Value + 10
            End If
        End If
    Else
        '调用默认窗口消息处理函数
        NewWindowProc = CallWindowProc(glngOldWindowProc, hWnd, msg, wParam, lParam)
    End If
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = NVL(rsTmp!简码)
            UserInfo.姓名 = NVL(rsTmp!姓名)
            UserInfo.部门ID = NVL(rsTmp!部门ID, 0)
            UserInfo.部门码 = NVL(rsTmp!部门码)
            UserInfo.部门名 = NVL(rsTmp!部门名)
            UserInfo.专业技术职务 = NVL(rsTmp!专业技术职务)
            GetUserInfo = True
        End If
    End If
End Function


Public Function SubmitMainInfo(ByVal strJsonIn As String, ByRef strJsonOut As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : SubmitMainInfo
' Author    : YWJ
' Date      : 2019-09-19 16:40:39
' Parameter : strJsonIn-传入JSON字符串 具体格式见下:
'             strJsonOut-传出JSON字符串
' Purpose   : 提交主体信息
' Return    : T-成功;F-失败
' TEST URL  : http://192.168.32.201:8889/bizdomain/07f7c460-7dd8-49b5-a79b-0a90b9369224
'---------------------------------------------------------------------------------------
    Dim strErr As String
    Dim blnRet As Boolean
    
    '--日志
    WriteLog "标题：提交主体信息" & vbNewLine & _
             "函数：SubmitMainInfo" & vbNewLine & _
             "入参：" & strJsonIn & vbNewLine
    blnRet = Sys.NewSystemSvr("知识库", "提交主体信息", strJsonIn, strJsonOut, strErr)
    WriteLog "标题：提交主体信息" & vbNewLine & _
             "函数：SubmitMainInfo" & vbNewLine & _
             "出参：" & strJsonOut & vbNewLine & _
             "返回值：" & blnRet & vbNewLine & _
             IIf(strErr <> "", "错误信息:" & strErr & vbNewLine, "")
    SubmitMainInfo = blnRet
     
End Function
 


Public Function SubmitInquiriyInfo(ByVal strJsonIn As String, ByRef strJsonOut As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : SubmitInquiriyInfo
' Author    : YWJ
' Date      : 2019-09-19 16:50:43
' Parameter : strJsonIn-传入JSON字符串 具体格式见下:
'             strJsonOut-传出JSON字符串
' Purpose   : 提交问诊信息
' Return    : T-成功;F-失败
' TEST URL  : http://192.168.32.201:8889/bizdomain/c6050afd-135e-454a-b53f-e5a7d7634399
'---------------------------------------------------------------------------------------
'
    Dim strErr As String
    Dim blnRet As Boolean
    
     WriteLog "标题：提交问诊信息" & vbNewLine & _
             "函数：SubmitInquiriyInfo" & vbNewLine & _
             "入参：" & strJsonIn & vbNewLine
    blnRet = Sys.NewSystemSvr("知识库", "提交问诊信息", strJsonIn, strJsonOut, strErr)
    WriteLog "标题：提交问诊信息" & vbNewLine & _
             "函数：SubmitInquiriyInfo" & vbNewLine & _
             "出参：" & strJsonOut & vbNewLine & _
             "返回值：" & blnRet & vbNewLine & _
             IIf(strErr <> "", "错误信息:" & strErr & vbNewLine, "")
    SubmitInquiriyInfo = blnRet
End Function


Public Function GetPatiInfo(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
    ByVal bytScene As Byte, ByRef lngRegId As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : GetPatiInfo
' Author    : YWJ
' Date      : 2019-09-23 13:45:55
' Parameter :
'             lngPatiID -病人ID
'             lngVisitId -主页ID
'             strRegNO -挂号单号
'             bytScene-场合 1-门诊\住院医嘱下达;2-诊断报告;3-检验;4-标本采集
'             lngRegId-出参:挂号ID
' Purpose   : 获取病人信息
'---------------------------------------------------------------------------------------
'            病人基本信息   格式如下:
'            "patient_info":{
'            "pid":"5066404",
'            "visit_id":"1",
'            "visit_no":"314929",
'            "name":"王琳琳",
'            "age":"31岁",
'            "birthday":"1989-10-10 09-10-10",
'            "gender":"女",
'            "marital_status":"已婚",
'            "operator_id":"489b7bba-31cd-4f59-8fef-c12f0570db61",
'            "operator":"郑志鹏",
'            "enc_type":"2","scene":"1"
'            }
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strPati As String
    Dim bytType As Byte
    Dim strVisitId As String
    Dim strVisitNo As String
    On Error GoTo ErrH
    If strRegNo = "" Then
        strSQL = "Select Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.年龄, a.年龄) As 年龄, Nvl(b.性别, a.性别) As 性别, b.婚姻状况,a.出生日期,b.住院号 as 标识号 " & vbNewLine & _
                "From 病人信息 A, 病案主页 B" & vbNewLine & _
                "Where a.病人id = b.病人id And b.病人id = [1] And b.主页id = [2]"
        bytType = 2 '住院
        strVisitId = lngVisitId
         
    Else
                
        strSQL = "Select b.Id As 就诊id, Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.年龄, a.年龄) As 年龄, Nvl(b.性别, a.性别) As 性别, a.婚姻状况, a.出生日期,b.门诊号 as 标识号" & vbNewLine & _
                "From 病人信息 A, 病人挂号记录 B" & vbNewLine & _
                "Where a.病人id = b.病人id And a.病人id = [1] And b.No = [3] And b.记录性质 = 1 And b.记录状态 = 1"
        bytType = 1 '门诊
        strVisitId = strRegNo
    End If

    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiInfo", lngPatiID, lngVisitId, strRegNo)
    If rsTemp.EOF Then Exit Function
    If bytType = 1 Then lngRegId = rsTemp!就诊ID
    strPati = "\""patient_info\"":{\""pid\"":\""" & lngPatiID & "\""," & vbNewLine & _
                                "\""visit_id\"":\""" & strVisitId & "\""," & vbNewLine & _
                                "\""visit_no\"":\""" & rsTemp!标识号 & "\""," & _
                                "\""name\"":\""" & rsTemp!姓名 & "\""," & vbNewLine & _
                                "\""age\"":\""" & rsTemp!年龄 & "\""," & vbNewLine & _
                                "\""birthday\"":\""" & Format(rsTemp!出生日期, "YYYY-MM-DD HH:MM:SS") & "\""," & vbNewLine & _
                                "\""gender\"":\""" & rsTemp!性别 & "\""," & vbNewLine & _
                                "\""marital_status\"":\""" & rsTemp!婚姻状况 & "\""," & vbNewLine & _
                                "\""operator_id\"":\""" & UserInfo.id & "\""," & vbNewLine & _
                                "\""operator\"":\""" & UserInfo.姓名 & "\""," & vbNewLine & _
                                "\""enc_type\"":\""" & bytType & "\""," & vbNewLine & _
                                "\""scene\"":\""" & bytScene & "\""}"
                                
    GetPatiInfo = strPati

    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog

End Function


Public Function GetMainInfo(ByVal lngPatiID As Long, ByVal lngVisitId As Long, ByVal strRegNo As String, _
        ByVal lngRegId As Long, ByVal rsAdvice As ADODB.Recordset, ByVal colDiag As Collection) As String
      '---------------------------------------------------------------------------------------
      ' Procedure : GetMainInfo
      ' Author    : YWJ
      ' Date      : 2019-09-23 14:25:53
      ' Parameter :
      '             lngPatiID-病人ID
      '             lngVisitId-主页ID
      '             strRegNo-挂号单
      '             lngRegId -挂号ID
      '             colDiag -门诊诊断
      ' Purpose   : 获取主体信息
      '---------------------------------------------------------------------------------------
      '
          Dim strInfo As String
          
          Dim strDoctor As String
          Dim strDocPost As String   '医生职务
          Dim strDoctorList As String
          
          Dim strItemIds As String   '记录诊疗项目ID
          Dim strIndex As String
          Dim strCondition As String
          
          Dim strKey As String
          Dim strName As String
          Dim strType As String
          Dim strLisIds As String
          Dim str指标ID As String
          
          Dim i As Long
          Dim lngGroupID As Long
          Dim lngId As Long
          
          Dim blnNext As Boolean
          
          Dim colList As Collection
          Dim colItem As Collection
          Dim colTemp As Collection
          Dim colOther As Collection
          Dim colOtherItem As Collection
          
          Dim arrTemp As Variant
          
          Dim rsDoctor As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim rsDiag As ADODB.Recordset
          

1         On Error GoTo ErrH
2         Set colList = New Collection
          '获取诊断信息
          '门诊编辑界面按界面录入为准
3         If Not colDiag Is Nothing Then
4             For Each colTemp In colDiag
5                 colList.Add colTemp, "K" & colList.Count
6             Next
7         End If
      '    Set rsDiag = Get病人诊断记录(lngPatiID, IIf(strRegNo <> "", lngRegId, lngVisitId), IIf(strRegNo <> "", "1,11", "2,12"))
      '    Do While Not rsDiag.EOF
      '        Set colItem = New Collection
      '        colItem.Add NVL(rsDiag!疾病ID, rsDiag!诊断ID) & "", "key"
      '        colItem.Add rsDiag!名称 & "", "name"
      '        colItem.Add "诊断信息", "type"
      '        colList.Add colItem, "K" & colList.Count
      '        rsDiag.MoveNext
      '    Loop
8         With rsAdvice
              '用药审核或用药研究
9             .Filter = ""
10            For i = 1 To .RecordCount
                  '获取开嘱医生
11                If NVL(!开嘱医生) <> "" Then
12                    strDoctor = NVL(!开嘱医生)
13                    If InStr(strDoctor, "/") > 0 Then strDoctor = Mid(strDoctor, 1, InStr(strDoctor, "/") - 1)
14                    If InStr("," & strDoctorList & ",", "," & strDoctor & ",") = 0 And strDoctor <> "" Then
15                        strDoctorList = strDoctorList & "," & strDoctor
16                    End If
17                End If
                   
18                If InStr("," & strItemIds & ",", "," & !诊疗项目ID & ",") = 0 Then
19                    strItemIds = strItemIds & "," & !诊疗项目ID
20                End If
                   
21                .MoveNext
22            Next
              '取专业技术职务
23            If strDoctorList <> "" Then
24                strDoctorList = Mid(strDoctorList, 2)
25                Set rsDoctor = GetRS("人员表", "编号,姓名,专业技术职务", strDoctorList, "姓名", , 1)
26            End If
              '获取诊疗项目目录
27            If strItemIds <> "" Then
28                strItemIds = Mid(strItemIds, 2)
29                Set rsItem = GetRS("诊疗项目目录", "ID,编码,名称", strItemIds)
30            End If
              
31            .Filter = ""
32            strDoctor = ""
33            lngGroupID = 0

34            Do While Not .EOF
35                blnNext = True
36                If strDoctor <> !开嘱医生 & "" Then strDocPost = GetDoctorPost(rsDoctor, NVL(!开嘱医生)): strDoctor = !开嘱医生 & ""
                  '药品项目
                  '附加项目 给药途径;给药频率;单次用量;总量;开嘱医生职务
37                If InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 Then
38                    lngGroupID = Decode(!相关ID, 0, !id, !相关ID)
39                    strIndex = ""
                      '循环遍历
40                    Do While Not .EOF
41                        If lngGroupID <> Decode(!相关ID, 0, !id, !相关ID) Then Exit Do
42                        Set colItem = New Collection
43                        If InStr(",5,6,7,", "," & !诊疗类别 & ",") > 0 Then
44                            colItem.Add !诊疗项目ID & "", "key"
45                            colItem.Add !标本部位 & "", "name"
46                            colItem.Add "药品项目", "type"
                              '附加项目
47                            Set colOther = New Collection
48                            Set colOtherItem = New Collection
49                            colOtherItem.Add !执行频次 & "", "给药频率"
50                            colOtherItem.Add FormatEx(NVL(!单次用量), 5), "单次用量"
51                            colOtherItem.Add FormatEx(NVL(!总给予量), 5), "总量"
52                            colOtherItem.Add strDocPost, "开嘱医生职务"
53                            colOtherItem.Add "给药途径,给药频率,单次用量,总量,开嘱医生职务", "keys"
                              
54                            colOther.Add colOtherItem
55                            colItem.Add colOther, "other"
                              
56                            If strIndex <> "" Then strIndex = strIndex & ","
57                            strIndex = strIndex & "K" & colList.Count
                          
58                        ElseIf !诊疗类别 & "" = "E" And !id = lngGroupID Then
                              '附加给药途径
59                            arrTemp = Split(strIndex, ",")
60                            For i = LBound(arrTemp) To UBound(arrTemp)
61                                Set colTemp = colList(arrTemp(i))
62                                For Each colOtherItem In colTemp("other")
63                                    colOtherItem.Add !诊疗项目ID & "", "给药途径"
64                                Next
65                            Next
                              
                              '给药途径
                              '附加项目 开嘱医生职务
66                            colItem.Add !诊疗项目ID & "", "key"
67                            colItem.Add !医嘱内容 & "", "name"
68                            colItem.Add "给药途径", "type"
                              
69                            Set colOther = New Collection
70                            Set colOtherItem = New Collection
                              
71                            colOtherItem.Add strDocPost, "开嘱医生职务"
72                            colOtherItem.Add "开嘱医生职务", "keys"
                              
73                            colOther.Add colOtherItem
74                            colItem.Add colOther, "other"
75                        End If
76                        colList.Add colItem, "K" & colList.Count
77                        .MoveNext
78                    Loop
79                    blnNext = False '一组医嘱已经遍历结束或已经到记录集末尾,禁止继续向下MoveNext
80                ElseIf !诊疗类别 & "" = "F" Then
                  '手术项目
                  '附加项目 麻醉方式;开嘱医生职务
                  '麻醉项目
                  '附加项目 开嘱医生职务
81                    lngGroupID = Decode(!相关ID, 0, !id, !相关ID)
82                    strIndex = "": strName = ""
                      '循环遍历
83                    Do While Not .EOF
84                        If lngGroupID <> Decode(!相关ID, 0, !id, !相关ID) Then Exit Do
85                        lngId = Val(rsAdvice!相关ID & "")
86                        Set colItem = New Collection
                                              
87                        If !诊疗类别 & "" = "F" Then
88                            If lngId = 0 Then
                                  '主手术
89                                strName = GetItemInfo(rsItem, CLng(!诊疗项目ID & ""))
90                            Else
91                                strName = !医嘱内容 & ""
92                            End If
                              
93                            colItem.Add !诊疗项目ID & "", "key"
94                            colItem.Add strName, "name"
95                            colItem.Add "手术项目", "type"
              
                              '附加项目
96                            Set colOther = New Collection
97                            Set colOtherItem = New Collection
                              
98                            colOtherItem.Add strDocPost, "开嘱医生职务"
99                            colOtherItem.Add IIf(lngId = 0, 1, 0), "主手术" '1-是主手术;0-附加手术
100                           colOtherItem.Add "麻醉方式,开嘱医生职务,主手术", "keys"
                              
101                           colOther.Add colOtherItem
102                           colItem.Add colOther, "other"
                              
103                           If strIndex <> "" Then strIndex = strIndex & ","
104                           strIndex = strIndex & "K" & colList.Count
                              
105                       ElseIf !诊疗类别 & "" = "G" Then
                              '附加麻醉方式
106                           arrTemp = Split(strIndex, ",")
107                           For i = LBound(arrTemp) To UBound(arrTemp)
108                               Set colTemp = colList(arrTemp(i))
109                               For Each colOtherItem In colTemp("other")
110                                   colOtherItem.Add !诊疗项目ID & "", "麻醉方式"
111                               Next
112                           Next
                              '麻醉项目
113                           colItem.Add !诊疗项目ID & "", "key"
114                           colItem.Add !医嘱内容 & "", "name"
115                           colItem.Add "麻醉项目", "type"
                              
                              '附加项目 开嘱医生职务
116                           Set colOther = New Collection
117                           Set colOtherItem = New Collection
                              
118                           colOtherItem.Add strDocPost, "开嘱医生职务"
119                           colOtherItem.Add "开嘱医生职务", "keys"
                              
120                           colOther.Add colOtherItem
121                           colItem.Add colOther, "other"
122                       End If
                          
123                       colList.Add colItem, "K" & colList.Count
                          
124                       .MoveNext
125                   Loop
                      
126                   blnNext = False '一组医嘱已经遍历结束或已经到记录集末尾,禁止继续向下MoveNext
                      
127               ElseIf !诊疗类别 & "" = "K" Then
                      '输血项目
                      '附加项目 输血方法;输血量;开嘱医生职务
128                   lngGroupID = Decode(!相关ID, 0, !id, !相关ID)
129                   strKey = "": strName = ""
                      
130                   strIndex = ""
                      '循环遍历检查
131                   Do While Not .EOF
132                       If lngGroupID <> Decode(!相关ID, 0, !id, !相关ID) Then Exit Do
133                       Set colItem = New Collection
134                       lngId = Val(rsAdvice!相关ID & "")
135                       If strDoctor <> NVL(!开嘱医生) Then strDocPost = GetDoctorPost(rsDoctor, NVL(!开嘱医生))
136                       If !诊疗类别 & "" = "K" Then
137                           strName = GetItemInfo(rsItem, CLng(!诊疗项目ID & ""))
138                           colItem.Add !诊疗项目ID & "", "key"
139                           colItem.Add strName, "name"
140                           colItem.Add "输血项目", "type"
                              '附加项目
141                           Set colOther = New Collection
142                           Set colOtherItem = New Collection
143                           colOtherItem.Add FormatEx(NVL(!总给予量), 5), "输血量"
144                           colOtherItem.Add strDocPost, "开嘱医生职务"
145                           colOtherItem.Add "输血方法,输血量,开嘱医生职务", "keys"
                              
146                           colOther.Add colOtherItem
147                           colItem.Add colOther, "other"
                              
148                           If strIndex <> "" Then strIndex = strIndex & ","
149                           strIndex = strIndex & "K" & colList.Count
                          
150                       ElseIf !诊疗类别 & "" = "E" Then
                              '附加输血方法
151                           arrTemp = Split(strIndex, ",")
152                           For i = LBound(arrTemp) To UBound(arrTemp)
153                               Set colTemp = colList(arrTemp(i))
154                               For Each colOtherItem In colTemp("other")
155                                   colOtherItem.Add !诊疗项目ID & "", "输血方法"
156                               Next
157                           Next
                              
                              '输血方法
                              '附加项目 开嘱医生职务
158                           colItem.Add !诊疗项目ID & "", "key"
159                           colItem.Add !医嘱内容 & "", "name"
160                           colItem.Add "输血方法", "type"
                              
161                           Set colOther = New Collection
162                           Set colOtherItem = New Collection
                              
163                           colOtherItem.Add strDocPost, "开嘱医生职务"
164                           colOtherItem.Add "开嘱医生职务", "keys"
                              
165                           colOther.Add colOtherItem
166                           colItem.Add colOther, "other"

167                       End If
168                       colList.Add colItem, "K" & colList.Count
                          
169                       .MoveNext
170                   Loop
                      
171                   blnNext = False '一组医嘱已经遍历结束或已经到记录集末尾,禁止继续向下MoveNext
                      
172               ElseIf !诊疗类别 & "" = "C" Then
                      '检验项目
                      '附加项目 采集方法;标本类型;开嘱医生职务
173                   lngGroupID = Decode(!相关ID, 0, !id, !相关ID)
174                   strIndex = ""
                      '循环遍历
175                   Do While Not .EOF
176                       If lngGroupID <> Decode(!相关ID, 0, !id, !相关ID) Then Exit Do
177                       Set colItem = New Collection
178                       If !诊疗类别 & "" = "C" Then
179                           If strLisIds <> "" Then strLisIds = strLisIds & ","
180                           strLisIds = strLisIds & !诊疗项目ID
                              
181                           colItem.Add !诊疗项目ID & "", "key"
182                           colItem.Add !医嘱内容 & "", "name"
183                           colItem.Add "检验项目", "type"
                              '附加项目
184                           Set colOther = New Collection
185                           Set colOtherItem = New Collection
186                           colOtherItem.Add !标本部位 & "", "标本类型"
187                           colOtherItem.Add strDocPost, "开嘱医生职务"
188                           colOtherItem.Add "采集方法,标本类型,开嘱医生职务", "keys"
                              
189                           colOther.Add colOtherItem
190                           colItem.Add colOther, "other"
                              
191                           If strIndex <> "" Then strIndex = strIndex & ","
192                           strIndex = strIndex & "K" & colList.Count
                          
193                       ElseIf !诊疗类别 & "" = "E" And !id = lngGroupID Then
                              '附加采集方法
194                           arrTemp = Split(strIndex, ",")
195                           For i = LBound(arrTemp) To UBound(arrTemp)
196                               Set colTemp = colList(arrTemp(i))
197                               For Each colOtherItem In colTemp("other")
198                                   colOtherItem.Add !诊疗项目ID & "", "采集方法"
199                               Next
200                           Next
                              
                              '采集方法
                              '附加项目 开嘱医生职务
201                           colItem.Add !诊疗项目ID & "", "key"
202                           colItem.Add GetItemInfo(rsItem, CLng(!诊疗项目ID & "")), "name"
203                           colItem.Add "采集方法", "type"
                              
204                           Set colOther = New Collection
205                           Set colOtherItem = New Collection
                              
206                           colOtherItem.Add strDocPost, "开嘱医生职务"
207                           colOtherItem.Add "开嘱医生职务", "keys"
                              
208                           colOther.Add colOtherItem
209                           colItem.Add colOther, "other"
210                       End If
211                       colList.Add colItem, "K" & colList.Count
212                       .MoveNext
213                   Loop
214                   blnNext = False '一组医嘱已经遍历结束或已经到记录集末尾,禁止继续向下MoveNext
215               ElseIf !诊疗类别 & "" = "D" Then
                  '检查项目
                  '附加项目 部位;方法;开嘱医生职务
216                   lngGroupID = Decode(!相关ID, 0, !id, !相关ID)
217                   strName = ""
218                   Set colItem = New Collection
219                   Set colOther = New Collection
                      '循环遍历检查
220                   Do While Not .EOF
221                       If lngGroupID <> Decode(!相关ID, 0, !id, !相关ID) Then Exit Do
222                       lngId = Val(rsAdvice!相关ID & "")
223                       If lngId = 0 Then
224                           colItem.Add !诊疗项目ID & "", "key"
225                           colItem.Add "检查项目", "type"
226                           strName = !医嘱内容 & ""
227                       Else
228                           strName = !医嘱内容 & ""
                              '附加项目
229                           Set colOtherItem = New Collection
230                           colOtherItem.Add !标本部位 & "", "部位"
231                           colOtherItem.Add !检查方法 & "", "方法"
232                           colOtherItem.Add strDocPost, "开嘱医生职务"
233                           colOtherItem.Add "部位,方法,开嘱医生职务", "keys"
                              
234                           colOther.Add colOtherItem
235                       End If
236                       .MoveNext
237                   Loop
                      
238                   blnNext = False '一组医嘱已经遍历结束或已经到记录集末尾,禁止继续向下MoveNext
239                   colItem.Add strName, "name"
                      
240                   If colOther.Count = 0 Then '单条检查医嘱特殊处理
241                       Set colOtherItem = New Collection
242                       colOtherItem.Add strDocPost, "开嘱医生职务"
243                       colOtherItem.Add "开嘱医生职务", "keys"
244                   End If
245                   colItem.Add colOther, "other"
                      
246                   colList.Add colItem, "K" & colList.Count
247               Else
                      '检验指标
                      '附加项目 指标性质
                      
                      '其他医嘱项目
                      '附加项目 开嘱医生职务
248                   If !诊疗类别 & "" = "H" Then
                          '护理项目
                          '附加项目 开嘱医生职务
249                       strType = "护理项目"
250                   ElseIf !诊疗类别 & "" = "E" Then
                          '治疗项目
                          '附加项目 开嘱医生职务
251                       strType = "治疗项目"
252                   Else
253                       strType = "其他医嘱项目"
254                   End If
255                   Set colItem = New Collection
256                   Set colOther = New Collection
                      
257                   colItem.Add !诊疗项目ID & "", "key"
258                   colItem.Add !医嘱内容 & "", "name"
259                   colItem.Add strType, "type"
                          
260                   Set colOtherItem = New Collection
261                   colOtherItem.Add strDocPost, "开嘱医生职务"
262                   colOtherItem.Add "开嘱医生职务", "keys"
                      
263                   colOther.Add colOtherItem
                      
264                   colItem.Add colOther, "other"
265                   colList.Add colItem, "K" & colList.Count
266               End If
                                      
267               If blnNext Then .MoveNext
268           Loop
              '指标ID,检验指标,skey,sname
269           If strLisIds <> "" And Not gobjLIS Is Nothing Then
270               Set rsItem = Nothing
271               On Error Resume Next
272               Set rsItem = gobjLIS.GetGroupItemInfo(strLisIds)
273               On Error GoTo ErrH '结束上一忽略错误,重新开启错误捕获
274               If Not rsItem Is Nothing Then
275                   str指标ID = ""
276                   Do While Not rsItem.EOF
277                       If str指标ID <> rsItem!指标ID & "" Then
278                           If str指标ID <> "" Then
279                               colItem.Add colOther, "other"
280                               colList.Add colItem, "K" & colList.Count
                                  
281                               Set colItem = Nothing
282                               Set colOther = Nothing
283                           End If
284                           str指标ID = rsItem!指标ID & ""
                              
285                           Set colItem = New Collection
286                           colItem.Add rsItem!指标ID & "", "key"
287                           colItem.Add rsItem!检验指标 & "", "name"
288                           colItem.Add "检验指标", "type"
                               
289                           If rsItem!sname & "" <> "" And rsItem!sKey & "" <> "" Then
290                               Set colOther = New Collection
291                               Set colOtherItem = New Collection
                                  
292                               colOtherItem.Add rsItem!sname & "", rsItem!sKey & ""
293                               colOtherItem.Add rsItem!sKey & "", "keys"
                                  
294                               colOther.Add colOtherItem
295                           End If
296                       Else
297                           If rsItem!sname & "" <> "" And rsItem!sKey & "" <> "" And Not colOther Is Nothing Then
298                               Set colOtherItem = New Collection
                                  
299                               colOtherItem.Add rsItem!sname & "", rsItem!sKey & ""
300                               colOtherItem.Add rsItem!sKey & "", "keys"
                                  
301                               colOther.Add colOtherItem
302                           End If
303                       End If
                      
304                       rsItem.MoveNext
305                       If rsItem.EOF Then
306                           colItem.Add colOther, "other"
307                           colList.Add colItem, "K" & colList.Count
308                       End If
309                   Loop
310               End If
311           End If
                      
312       End With
313       If colList.Count > 0 Then
314           strInfo = GetMainJson(colList)
315       End If
316       GetMainInfo = strInfo

317       Exit Function

ErrH:
318       MsgBox "在zlCISRule.mdlPublic.GetMainInfo的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName
End Function

Private Function GetDoctorPost(ByVal rsDoctor As ADODB.Recordset, ByVal strDoctor As String) As String
'功能:获取医师专业技术职务
'参数:
    rsDoctor.Filter = "姓名='" & strDoctor & "'"
    If Not rsDoctor.EOF Then GetDoctorPost = NVL(rsDoctor!专业技术职务)
End Function

Private Function GetItemInfo(ByVal rsItem As ADODB.Recordset, ByVal lngId As Long) As String
'功能:获取医师专业技术职务
'参数:
    rsItem.Filter = "ID=" & lngId & ""
    If Not rsItem.EOF Then GetItemInfo = NVL(rsItem!名称)
End Function

Public Function GetMainJson(ByVal colList As Collection) As String
      '功能：构造主体信息
          Dim colItem As Collection
          Dim colOther As Collection
          Dim colOtherItem As Collection
          Dim arrKeys As Variant
          
          Dim strInfo As String
          Dim strTemp As String
          Dim strOther As String
          
          Dim i As Long

1         On Error GoTo ErrH

2         strInfo = ""
3         For Each colItem In colList
4             Set colOther = GetCollValue(colItem, "other")
5             strOther = ""
6             If Not colOther Is Nothing Then
7                 For Each colOtherItem In colOther
8                     arrKeys = Split(GetCollElement(colOtherItem, "keys"), ",")
9                     strTemp = ""
10                    For i = LBound(arrKeys) To UBound(arrKeys)
11                        If strTemp <> "" Then strTemp = strTemp & ","
12                        If arrKeys(i) <> "" Then
13                           strTemp = strTemp & "{\""skey\"":\""" & arrKeys(i) & "\"",\""sname\"":\""" & GetCollElement(colOtherItem, CStr(arrKeys(i))) & "\""}"
14                        End If
15                    Next
16                    If strOther <> "" Then strOther = strOther & ","
17                    strOther = strOther & "{\""value_group\"":[" & strTemp & "]}"
18                Next
19                If strOther <> "" Then strOther = "\""condition_info\"":[" & strOther & "]"
21            End If
22            If strInfo <> "" Then strInfo = strInfo & ","
23            strInfo = strInfo & "{\""key\"":\""" & GetCollElement(colItem, "key") & "\""," & vbNewLine & _
                                  "\""name\"":\""" & GetCollElement(colItem, "name") & "\""," & vbNewLine & _
                                  "\""type\"":\""" & GetCollElement(colItem, "type") & "\"""
24            If strOther = "" Then
25                strInfo = strInfo & "}"
26            Else
27                strInfo = strInfo & "," & strOther & "}"
28            End If
                                   
29        Next
30        strInfo = "\""main_info\"":[" & strInfo & "]"
31        GetMainJson = strInfo
32        Exit Function

ErrH:
33        MsgBox "在zlCISRule.mdlPublic.GetMainJson的第" & Erl() & "行出错：" & vbCrLf & _
            "错误号: " & Err.Number & vbCrLf & _
            "错误描述：" & Err.Description, vbExclamation, gstrSysName

End Function

Public Function TestJson(ByVal bytFunc As Byte) As String
    Dim strJson As String
    If bytFunc = 1 Then
        strJson = "{\""cdss_in\"":{\""patient_info\"":{\""pid\"": \""5066404\"",\""visit_id\"": \""1\""," & _
                "\""visit_no\"": \""314929\"",\""name\"": \""王琳琳\"",\""age\"": \""31岁\"",\""gender\"": \""女\"",\""marital_status\"": \""已婚\"",\""operator_id\"": \""489b7bba-31cd-4f59-8fef-c12f0570db61\"",\""operator\"": \""郑志鹏\"",\""enc_type\"": \""2\"",\""scene\"": \""1\""}," & _
                "\""main_info\"": [{\""key\"": \""168\"",\""name\"": \""注射用苄星青霉素\"",\""type\"": \""药品项目\"",\""condition_info\"":[{\""value_group\"":[{\""skey\"": \""给药途径\"",\""sname\"": \""2203\""},{\""skey\"": \""给药频率\"",\""sname\"": \""1\""},{\""skey\"": \""单次用量\"",\""sname\"": \""1\""},{\""skey\"": \""总量\"",\""sname\"": \""1\""},{\""skey\"": \""开嘱人\"",\""sname\"": \""郑志鹏\""}]}]},{\""key\"": \""2203\"",\""name\"": \""肌内注射\"",\""type\"": \""给药途径\""}]}}"
        TestJson = "{""接口json_in"":""" & strJson & """}"
        
    ElseIf bytFunc = 2 Then
        strJson = "{" & vbNewLine & _
                """businss"":""F8E7C2918A6C4060B29FE5D3FD66135A""," & vbNewLine & _
                """inquiry"":[{" & vbNewLine & _
                "    ""observ_item_id"":""F0E3D17C3CDE4FBF89FF020372D0A1EF""," & vbNewLine & _
                "    ""item_name"":""过敏体质""," & vbNewLine & _
                "    ""item_code"":""""," & vbNewLine & _
                "    ""observ_item_values"":[{" & vbNewLine & _
                "        ""item_detail_id"":""A43FD7C4166A470E9CD266FC9AC9D0B3""," & vbNewLine & _
                "        ""disp_name"":""是""," & vbNewLine & _
                "        ""default_sign"":""1""" & vbNewLine & _
                "        }, {" & vbNewLine & _
                "        ""item_detail_id"":""56E11442BA374FA8B7DF2FECE895AA13""," & vbNewLine & _
                "        ""disp_name"":""否""," & vbNewLine & _
                "        ""default_sign"":""0""" & vbNewLine & _
                "        }]" & vbNewLine & _
                "    }, {"
        strJson = strJson & "" & vbNewLine & _
                """observ_item_id"":""F14CEA53C2B646A399FD0DD491BAF0FE""," & vbNewLine & _
                """item_name"":""妊娠周期""," & vbNewLine & _
                """item_code"":""""," & vbNewLine & _
                """observ_item_values"":[{" & vbNewLine & _
                "    ""item_detail_id"":""77F1DF1C24884576BFBCCBDA39DA3D9F""," & vbNewLine & _
                "    ""disp_name"":""妊娠早期""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }, {" & vbNewLine & _
                "    ""item_detail_id"":""14345A4091194C4FAD518BDB029B1325""," & vbNewLine & _
                "    ""disp_name"":""妊娠中期""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }, {" & vbNewLine & _
                "    ""item_detail_id"":""8A64280BC22D432CB28346EA653E58F8""," & vbNewLine & _
                "    ""disp_name"":""妊娠晚期""," & vbNewLine & _
                "    ""default_sign"":""0""" & vbNewLine & _
                "    }]" & vbNewLine & _
                "}]," & vbNewLine & _
                ""
        strJson = strJson & "" & vbNewLine & _
                """messages"":[{" & vbNewLine & _
                "    ""business_name"":""药品禁忌""," & vbNewLine & _
                "    ""return_info"":""病人【王琳琳】使用药品【苄星青霉素注射剂】，禁止【肌肉注射】""," & vbNewLine & _
                "    ""key"":""2""," & vbNewLine & _
                "    ""name"":""苄星青霉素注射剂""," & vbNewLine & _
                "    ""rule_name"":""给药禁忌""," & vbNewLine & _
                "    ""taboo_level"":""警告""" & vbNewLine & _
                "    }]" & vbNewLine & _
                "}"
        TestJson = """out"":" & strJson
    ElseIf bytFunc = 3 Then
        '干预数据
        strJson = "{""businss"":""DA32302FD74541B98DC4F3F992E5B206"",""messages"":[{""business_name"":""药品禁忌"",""return_info"":""病人【王琳琳】就诊号【314929】，因【哺乳期禁止使用】【注射用苄星青霉素】"",""key"":""168"",""name"":""注射用苄星青霉素"",""rule_name"":""哺乳期禁止使用"",""taboo_level"":""禁止"",""class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""b94b6533-d96e-4836-a386-825273461144""}]}"
        TestJson = strJson
    ElseIf bytFunc = 4 Then
        '问诊数据
        strJson = "{""cdss_in"":{""patient_info"":{""pid"": ""4613704"",""visit_id"": ""1"",""visit_no"": ""303740"",""name"": ""滕凤敏"",""age"": ""71岁"",""birthday"": ""1948-03-12"",""gender"": ""女"",""marital_status"": ""已婚"",""operator_id"": ""489b7bba-31cd-4f59-8fef-c12f0570db61"",""operator"": ""郑志鹏"",""enc_type"": ""2""},""main_info"": [{""key"": ""168"",""name"": ""注射用苄星青霉素"",""type"": ""药品项目"",""condition_info"":[{""value_group"":[{""skey"": ""给药途径"",""sname"": ""144511""},{""skey"": ""给药频率"",""sname"": ""1""},{""skey"": ""单次用量"",""sname"": ""1""},{""skey"": ""总量"",""sname"": ""1""},{""skey"": ""开嘱人"",""sname"": ""郑志鹏""}]}]},{""key"": ""144511"",""name"": ""肌肉注射"",""type"": ""给药途径""}]}}"
        strJson = Replace(strJson, """", "\""")
        TestJson = "{""接口json_in"":""" & strJson & """}"
    ElseIf bytFunc = 5 Then
     '问诊加载数据
        strJson = "{""businss"":""97CEF2E3E5894591A4EB1BA3A215146E"",""inquiry"":" & vbNewLine & _
                "[{""observ_item_id"":""F0E3D17C3CDE4FBF89FF020372D0A1EF"",""item_name"":""过敏体质"",""item_code"":"""",""observ_item_values"":" & vbNewLine & _
                "[{""item_detail_id"":""A43FD7C4166A470E9CD266FC9AC9D0B3"",""disp_name"":""是"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""56E11442BA374FA8B7DF2FECE895AA13"",""disp_name"":""否"",""default_sign"":""0""}]}," & vbNewLine & _
                "{""observ_item_id"":""9EF6C13B62094E698E481868C703E634"",""item_name"":""妊娠状态"",""item_code"":"""",""observ_item_values"":" & vbNewLine & _
                "[{""item_detail_id"":""427EAB8923FF49258FC935E74BAACAE7"",""disp_name"":""是"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""CA018CB5A6914896AA4FC4D1C137164F"",""disp_name"":""否"",""default_sign"":""0""}]},"
        
        strJson = strJson & "{""observ_item_id"":""4BA5C77FE9D9408389A9BE085E30E99A"",""item_name"":""肝功能不全"",""item_code"":""""," & vbNewLine & _
                """observ_item_values"":[{""item_detail_id"":""C532D50ACF6C4DDC8D20A0039B87C2FD"",""disp_name"":""是"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""1D958F80177546F0BC06619AEDCF00A6"",""disp_name"":""否"",""default_sign"":""0""}]}," & vbNewLine & _
                "{""observ_item_id"":""1640B9E956AD4372BC353DE709455B45"",""item_name"":""肾功能不全"",""" & vbNewLine & _
                "item_code"":"""",""observ_item_values"":[{""item_detail_id"":""9EE1C7FC4EB84B28A829CC1B9654A4F1"",""disp_name"":""是"",""default_sign"":""1""}," & vbNewLine & _
                "{""item_detail_id"":""562F75B5A00849B5A5A92C0144CD49BC"",""disp_name"":""否"",""default_sign"":""0""}]}],"
        
        strJson = strJson & """messages"":[{""business_name"":""药品禁忌"",""return_info"":""病人【滕凤敏】就诊号【303740】，因【哺乳期禁止使用】【注射用苄星青霉素】""," & vbNewLine & _
                """key"":""168"",""name"":""注射用苄星青霉素"",""rule_name"":""哺乳期禁止使用"",""taboo_level"":""禁止""," & vbNewLine & _
                """class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""b94b6533-d96e-4836-a386-825273461144""}," & vbNewLine & _
                "{""business_name"":""药品禁忌"",""return_info"":""病人【滕凤敏】就诊号【303740】使用药品【注射用苄星青霉素】，慎用【肌肉注射】""," & vbNewLine & _
                """key"":""168"",""name"":""注射用苄星青霉素"",""rule_name"":""给药禁忌"",""taboo_level"":""警告""," & vbNewLine & _
                """class_id"":""46dfdfbd-fdab-429a-877c-a4dee02752e7"",""detail_id"":""72be25bb-0a1e-4aed-ab0f-ce181d87a694""}]}"
                
        TestJson = strJson
    End If
End Function

Public Function GetCollValue(ByVal colList As Collection, ByVal varRow As Variant, Optional ByVal strElement As String) As Variant
'功能：获取Json数组返回的集合数据中指定行或指定元素的值
'参数：
'  varRow=行索引或行关键字
'  strElement=元素名
'返回：
'  当未传入strElement参数时，返回指定行的集合对象；当传入strElement参数时，返回指定行指定元素的值
'  失败时返回Nothing或Empty，但不会报错

    If strElement <> "" Then
        GetCollValue = Empty
    Else
        Set GetCollValue = Nothing
    End If
    
    If colList Is Nothing Then Exit Function
    
    On Error Resume Next
    If strElement <> "" Then
        GetCollValue = colList(varRow)(strElement)
    Else
        Set GetCollValue = colList(varRow)
    End If
    Err.Clear: On Error GoTo 0
End Function

Public Function GetCollElement(ByVal colList As Collection, ByVal strElement As String) As Variant
'功能：获取集合数组中的元素值(元素值为基本数据类型)
'参数：
'  varRow=行索引或行关键字
'  strElement=元素名
'返回：
'   返回指定行指定元素的值
'   失败时返回Empty

 
    GetCollElement = Empty
    If colList Is Nothing Then Exit Function
    On Error Resume Next
    GetCollElement = colList(strElement)
    Err.Clear: On Error GoTo 0
End Function

Public Function HandleMessage(ByVal colList As Collection) As Boolean
'功能:根据返回警示级别对当前操作进行干预。
'返回值:T-禁止当前操作;F-继续当前流程
    Dim i As Long
    Dim strMsg As String
    Dim lngLevel As Byte '1-提醒;2-警告;3-禁止
    Dim lngMaxLevel As Byte
    Dim blnRet As Boolean
    
    For i = 1 To colList.Count
        If strMsg <> "" Then strMsg = strMsg & vbCrLf
        strMsg = strMsg & colList(i)("return_info")
        lngLevel = Decode(CStr(colList(i)("taboo_level")), "禁止", 3, "警告", 2, "提醒", 1, 0)
        If lngLevel > lngMaxLevel Then lngMaxLevel = lngLevel
    Next
    Select Case lngMaxLevel
    
    Case 1
        MsgBox strMsg, vbInformation + vbOKOnly, gstrSysName
    Case 2
        If MsgBox(strMsg & vbCrLf & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            blnRet = True
        End If
    Case 3
        MsgBox strMsg, vbExclamation + vbOKOnly, gstrSysName
        blnRet = True
    End Select
    HandleMessage = blnRet
End Function

Public Function GetRS(ByVal strTableName As String, ByVal strFileds As String, ByVal strInput As String, _
        Optional ByVal strWhere As String = "ID", Optional ByVal bytModel As Byte = 0, Optional ByVal bytType As Byte = 0) As Variant
'功能:返回指定表指定字段的记录集
'参数：strTableName-表名
'     strFileds
'     strInput 方式1(1个过滤条件)：ID1,ID2,...
'              方式2(2个过滤条件)：名称1,范围1;名称2,范围2;...
'             strSQL = "Select 编码, 名称, 适用范围" & vbNewLine & _
'                "From 诊疗频率项目" & vbNewLine & _
'                "Where (名称, 适用范围) In (Select /*+cardinality(B,10)*/" & vbNewLine & _
'                "                      C1, C2" & vbNewLine & _
'                "                     From Table(f_Str2list2('每天二次,1|每天三次,1', ';', ',')) B)"
'    bytModel=1 过滤条件为两列
'    当bytModel=1时： bytType=0-拆分列 C1,C2 同为字符串 =1-C1(Number),C2(Number);=2-C1(char),C2(Number);=3-C1(Number),C2(Char)
'    当bytModel=0时： bytType=0-f_num2list; bytType=1 f_Str2list


    Dim strSQL As String
    Dim strSub As String
    Dim strFun As String
    Dim arrTmp As Variant
    
    On Error GoTo ErrH
    
    If bytModel = 1 Then
        If bytType = 0 Then
            strSub = " C1,C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 1 Then
            strSub = " C1,C2 "
            strFun = "f_num2list2"
        ElseIf bytType = 2 Then
            strSub = "C1,To_Number(C2) As C2 "
            strFun = "f_Str2list2"
        ElseIf bytType = 3 Then
            strSub = " To_Number(C1) As C1,C2 "
            strFun = "f_Str2list2"
        End If
        strSQL = " Select  " & strFileds & vbNewLine & _
                " From  " & strTableName & vbNewLine & _
                " Where (" & strWhere & ") In (Select /*+cardinality(B,10)*/" & vbNewLine & _
                "                    " & strSub & vbNewLine & _
                "                     From Table(" & strFun & "([1], ';', ',')) B)"
    Else
        If bytType = 0 Then
            strFun = "f_num2list"
        ElseIf bytType = 1 Then
            strFun = "f_Str2list"
        End If
        arrTmp = Split(strInput, ",")
        If UBound(arrTmp) = 0 Or strInput = "" Then
            strSQL = "Select " & strFileds & "  From " & strTableName & " Where " & strWhere & " = [1]"
        ElseIf UBound(arrTmp) > 0 Then
            strSQL = "Select " & strFileds & vbNewLine & _
            "From " & strTableName & vbNewLine & _
            "Where " & strWhere & " In (Select /*+cardinality(A,10)*/ * From Table(" & strFun & "([1]))A )"
        End If
    End If
    Set GetRS = zlDatabase.OpenSQLRecord(strSQL, "mdlPass", strInput)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get病人诊断记录(ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str类型 As String) As ADODB.Recordset
'功能：获取病人诊断记录
'参数：lng就诊ID：门诊病人传挂号ID，住院病人传主页ID
'       诊断类型-1-西医门诊诊断;2-西医入院诊断;3-西医出院诊断;5-院内感染;6-病理诊断;7-损伤中毒码,8-术前诊断;9-术后诊断;
'        11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断;21-病原学诊断
'       记录来源:1-病历；2-入院登记；3-首页整理(门诊医生站,诊断摘要);
    Dim strSQL As String

    On Error GoTo ErrH
    strSQL = "Select a.ID,a.疾病id, a.诊断id, a.诊断描述, a.诊断次序, Nvl(b.编码, c.编码) As 编码, NVL(Nvl(b.名称, c.名称),a.诊断描述) 名称" & vbNewLine & _
             ",a.记录日期,a.记录人 " & vbNewLine & _
             "From 病人诊断记录 A, 疾病编码目录 B, 疾病诊断目录 C" & vbNewLine & _
             "Where a.病人id = [1] And a.主页id = [2] And 取消时间 Is Null And 记录来源 IN (1, 3) And Instr(',' ||[3]|| ',', ',' || 诊断类型 || ',') > 0 And a.疾病id = b.Id(+) And" & vbNewLine & _
             "      a.诊断id = c.Id(+)" & vbNewLine & _
             "Order By 记录来源, 诊断类型, 诊断次序"
    Set Get病人诊断记录 = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng病人ID, lng就诊ID, str类型)

    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub InitObjLis(ByVal lngProgram As Long)
'判断如果新版LIS部件为空就初始化
    Dim strErr As String
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = CreateObject("zlPublicHisCommLis.clsPublicHisCommLis")
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, lngProgram, gcnOracle, strErr) = False Then
                If strErr <> "" Then MsgBox "LIS部件初始化错误：" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub
