VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Access2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: Access2 全自动微粒子化学发光免疫分析仪
'厂商: BECKMAN COULTER
'作者: 陈东
'修改：王振涛
'用户: 上海曲阳医院(单向0),泽州县人民医院（双向）
Private mstrCode As String
Private mstrWaitSend As String

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
        '功能：解析数据
        '参数：
        'strSource：要解析的原始串
        'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
        '   每组检验结果以||分隔,元素之间以|分隔
        '   第0个元素：检验时间
        '   第1个元素：样本序号^是否急诊^条码号  可以只有样本序号
        '   第2个元素：检验人
        '   第3个元素：标本
        '   第4个元素：是否质控品
        '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
        '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
        'strReserved：最后不能完成解析的原始串，需要返回作后续处理
        'strCmd：如果需要，可返回向设备发送的命令
        Dim lng_S As Long, lng_E As Long, str_Code As String, var_Code As Variant, i As Integer, str_Line As String
        Dim str_日期 As String, str_标本号 As String, str_项目 As String, str_结果 As String, str_Items As String
        Dim str_本地日期 As String
    
        On Error GoTo hErr
        '应答处理
100     If InStr(strSource, Chr(&H5)) > 0 Then
102         strCmd = Chr(&H6)
104     ElseIf InStr(strSource, Chr(&HA)) > 0 Then
106         strCmd = Chr(&H6)
        End If
108     If strSource <> "" Then mstrCode = mstrCode & strSource
    
110     lng_S = InStr(mstrCode, Chr(&H5))
112     If lng_S > 0 Then
114         lng_E = InStr(lng_S, mstrCode, Chr(&H4))
        Else
116         lng_E = InStr(mstrCode, Chr(&H4))
        End If
    
118     Do While lng_S > 0 And lng_E > 0
120         str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
122         mstrCode = Mid(mstrCode, lng_E + 1)
            '-------------------------------------------------------
124         str_Items = "": str_日期 = "": str_标本号 = ""
        
126         var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
128         For i = LBound(var_Code) To UBound(var_Code)
130             str_Line = var_Code(i)
132             If str_Line Like Chr(&H5) & Chr(&H2) & "*H|*" Then
                
134                 Call ReadConf(App.Path & "\Access2.ini", "取本机时间", str_本地日期)
136                 If str_本地日期 = "1" Then
138                     str_日期 = Format(Now, "yyyy-MM-dd HH:mm")
                    Else
140                     str_日期 = Split(str_Line, "|")(13)
142                     If IsNumeric(str_日期) And Len(str_日期) >= 12 Then
144                         str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2)
                        Else
146                         str_日期 = Format(Now, "yyyy-MM-dd HH:mm")
                        End If
                    End If
148             ElseIf str_Line Like Chr(&H2) & "*Q|*" Then '双向通讯
150                 str_标本号 = Split(Split(str_Line, "|")(2), "^")(1)
152                 strCmd = "1|"
154             ElseIf str_Line Like Chr(&H2) & "*O|*" Then
156                 str_标本号 = Trim(Split(str_Line, "|")(2))
158             ElseIf str_Line Like Chr(&H2) & "*R|*^*" Then
160                 str_结果 = Trim(Split(str_Line, "|")(3))
162                 str_项目 = Trim(Split(Split(str_Line, "|")(2), "^")(3))
164                 If str_结果 <> "" And str_项目 <> "" Then str_Items = str_Items & "|" & str_项目 & "|" & str_结果
                End If
            Next
        
166         If Len(Trim(str_标本号)) >= 12 And Val(str_标本号) > 0 Then
168             str_标本号 = "^^" & str_标本号
170         ElseIf Val(str_标本号) <= 0 Then
172             str_标本号 = ""
            End If
        
174         If str_标本号 <> "" And IsDate(str_日期) Then
176             strResult = strResult & "||" & str_日期 & "|" & str_标本号
178             If str_Items <> "" Then strResult = strResult & "| |血清|0" & str_Items
            End If
            '-------------------------------------------------------
180         lng_S = InStr(mstrCode, Chr(&H5))
182         If lng_S > 0 Then
184             lng_E = InStr(lng_S, mstrCode, Chr(&H4))
            Else
186             lng_E = InStr(mstrCode, Chr(&H4))
            End If
        Loop
188     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
190     SaveLog "Analyse 第" & CStr(Erl()) & "行错误" & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5个元素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    Dim varSampleInfo As Variant
    Dim strNo As String
    Dim strItems As String
    Dim strLine As String
    Dim intloop As Integer
    Dim str_Line As String
    Dim strSendCmd As String
    Dim i As Integer
    On Error GoTo hErr
100 varSampleInfo = Split(strSample, "||")
102 SendSample = ""
104 blnSuccess = True

106 If iSendStep = 0 Then
108     For intloop = LBound(varSampleInfo) To UBound(varSampleInfo)
110         strLine = varSampleInfo(intloop)
112         If Val(Split(strLine, "|")(5)) > 0 Then
114             strNo = Split(strLine, "|")(5)
            Else
116             strNo = Split(strLine, "|")(1)
            End If

118         If strNo <> "" And UBound(Split(strLine, "|")) >= 10 Then
                '指令开始
120             str_Line = "1H|\^&|||Host LIS|||||||P|1|" & Format(Now, "yyyyMMddHHmmss") & Chr(&HD) & Chr(&H3)
122             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                '病人信息
124             str_Line = "2P|1|" & Chr(&HD) & Chr(&H3)
126             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

128             For i = 10 To UBound(Split(strLine, "|"))
130                 'strItems = Split(strLine, "|")(i)
                    '试验指令
132                 strItems = strItems & "^^^" & Split(strLine, "|")(i) & "^1\"
                    
                Next
                If strItems <> "" Then strItems = Left$(strItems, Len(strItems) - 1)
134             str_Line = "3O|1|" & strNo & "||" & strItems & "|R||||||A||||Serum" & Chr(&HD) & Chr(&H3)
136             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                '结束
138             str_Line = "4L|1|F" & Chr(&HD) & Chr(&H3)
140             strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

142             mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & strSendCmd & Chr(&H4)

            End If
        Next
144     SendSample = GetSendData
146     If mstrWaitSend <> "" Then iSendStep = 1               '开始通信
    Else
        
148     If InStr(strResponse, Chr(&H6)) > 0 Then
150         SendSample = GetSendData
152         If mstrWaitSend = "" Then iSendStep = 0
154         SaveLog "仪器应答06，通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
156     ElseIf InStr(strResponse, Chr(&H4)) > 0 Then
158         SendSample = GetSendData
160         If mstrWaitSend = "" Then iSendStep = 0
162         SaveLog "仪器应答04，表示希望暂停，根据协议仍当做06,通讯成功。发送:" & SendSample & "缓存:" & mstrWaitSend
164     ElseIf InStr(strResponse, Chr(&H15)) > 0 Then
166         SendSample = GetSendData
168         If mstrWaitSend = "" Then iSendStep = 0
170         SaveLog "仪器应答15，通讯失败。发送:" & SendSample & "缓存:" & mstrWaitSend
172     ElseIf strResponse = Chr(&H5) Then
174         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
176             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
178             mstrWaitSend = ""
            End If
180         Call SaveLog("仪器应答05，因仪器有优先权，本次发送取消")
182         iSendStep = 0
184         SendSample = Chr(&H6)
        Else
186         Call SaveLog("仪器应答" & strResponse)
        End If
188     If mstrWaitSend = "" Then iSendStep = 0    '已发送完，退出双向模式

    End If
    Exit Function
hErr:
190 SaveLog "SendSample第" & CStr(Erl()) & "行错误" & Err.Description
End Function




Private Function GetSendData() As String
    '从缓存中获取一段数据，返回。如果缓存中没有数据，返回空串
    
    Dim strSend As String
    
    If Left$(mstrWaitSend, 1) = "," Then mstrWaitSend = Mid$(mstrWaitSend, 2)
    If mstrWaitSend = "" Then Exit Function
    If InStr(mstrWaitSend, ",") <= 0 Then
        strSend = mstrWaitSend
        mstrWaitSend = ""
    Else
        strSend = Mid$(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
        mstrWaitSend = Mid$(mstrWaitSend, InStr(mstrWaitSend, ",") + 1)
    End If
    
    GetSendData = strSend
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFileSys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFileSys.FileExists(strIniFile) Then
        Set objStream = objFileSys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If arrLine(0) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub



Private Sub SaveLog(ByVal StrInput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------

    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject

    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If Dir(App.Path & "\调试.TXT") = "" Then Exit Sub
    strFileName = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"

    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (strDate & ":" & StrInput)
    objStream.Close
    Set objStream = Nothing
End Sub

Private Function GetCheckCode(ByVal str_in As String) As String
    '获取校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function




