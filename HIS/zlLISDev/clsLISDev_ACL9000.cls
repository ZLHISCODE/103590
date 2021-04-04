VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL9000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ACL9000
'厂商:贝克曼—库尔特
'作者:陈东
'用户:贵阳肿瘤医院(单项)，泽州医院(双向)
'修改人:王振涛
'修改日期:2012-12-5

Private mstrCode As String
Private mstrWaitSend As String

Public Function GetStartCmd() As String
'设备标准接口
'功能：获取设备开始发送命令
    GetStartCmd = Chr(6)
End Function

Public Function GetEndCmd() As String
'设备标准接口
'功能：获取设备停止发送命令
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    '设备标准接口
    '功能：解析数据
    '参数：
    'strSource：要解析的原始串
    'strResult：返回的检验结果(各仪器解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号 后面加 ^1 表示急诊 后加加 ^条码
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    'strReserved：最后不能完成解析的原始串，需要返回作后续处理
    'strCmd：如果需要，可返回向设备发送的命令

    Dim varLine As Variant, varItem As Variant, strItem As String, str日期 As String, str标本号 As String
    Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
    Dim strValues As String, lngLength As Long
    Dim str标本类型 As String
    Dim blnEnd As Boolean    '是否已解析完
    Dim i As Integer, str_Line As String
    On Error GoTo hErr

    '初始化返回值
100 strResult = "": strReserved = "": strCmd = ""
102 If InStr(strSource, Chr(&H5)) > 0 Then
104     strCmd = Chr(&H6)
106 ElseIf InStr(strSource, Chr(&HA)) > 0 Then
108     strCmd = Chr(&H6)
    End If

110 str标本类型 = Trim(ReadIni("ACL9000", "标本类型", App.Path & "\ACL9000.ini"))
112 If str标本类型 = "" Then str标本类型 = "血液": Call WriteIni("ACL9000", "标本类型", str标本类型, App.Path & "\ACL9000.ini")
    
114 mstrCode = mstrCode & strSource

116 If mstrCode Like "***" Then
118     blnEnd = False
        '        strReserved = strSource
120     Do While Not blnEnd
            '取一个完整的记录,剩余的存到strReserved中,下次解析
122         lngLength = InStr(mstrCode, "") - InStr(mstrCode, "")
124         strLine = Mid(mstrCode, InStr(mstrCode, ""), lngLength)
126         mstrCode = Mid(mstrCode, InStr(mstrCode, "") + 1)

128         If strLine Like "*" & Chr(&H2) & "?Q|*" Then                 '双向通讯
130             varLine = Split(strLine, vbCrLf)
132             For i = LBound(varLine) To UBound(varLine)
134                 str_Line = varLine(i)
136                 If str_Line Like Chr(&H5) & Chr(&H2) & "?H|*" Then
138                     str日期 = Split(str_Line, "|")(13)
140                     If IsNumeric(str日期) And Len(str日期) >= 12 Then
142                         str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2) & " " & Mid(str日期, 9, 2) & ":" & Mid(str日期, 11, 2)
                        Else
144                         str日期 = Format(Now, "yyyy-MM-dd HH:mm")
                        End If
146                 ElseIf str_Line Like Chr(&H2) & "?Q|*" Then
148                     If InStr(str_Line, "^") > 0 Then
150                         str标本号 = Split(Split(str_Line, "|")(2), "^")(1)
                        End If
                    End If
                Next

152             If Val(str标本号) <> 0 Then
154                 strResult = strResult & "||" & str日期 & "|^^" & Trim(str标本号)
                End If

156             If strResult <> "" Then
158                 strCmd = "1|" & strCmd                                      '双向指令
                End If

160             If Not mstrCode Like "" Then
162                 blnEnd = True
                End If

164         ElseIf strLine Like "*" & Chr(&H2) & "?R|*" Then             '解析结果

166             str日期 = "": str标本号 = "": str检验人标本质控 = "": strItem = "": strTmp = ""
                '解析
168             varLine = Split(strLine, vbCrLf)
170             strItem = ""
172             For iRow = LBound(varLine) To UBound(varLine)
174                 If InStr(varLine(iRow), "|") > 0 Then
176                     varItem = Split(varLine(iRow), "|")
178                     If UBound(varItem) >= 12 Then
180                         If Trim("" & varItem(0)) Like "*O" And Left(varItem(4), 3) = "^^^" Then    '标本号行
182                             str项目 = Replace(varItem(4), "^^^", "")
                                '一个新的标本开始,先保存上次的结果
184                             If strItem <> "" And str标本号 <> "" Then
                                    If Len(str标本号) >= 12 Then
                                        strResult = strResult & "||" & str日期 & "|^^" & str标本号 & "|" & str检验人标本质控 & strItem
                                    Else
                                        strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                                    End If
                                End If

188                             str标本号 = Trim("" & varItem(2)): str检验人标本质控 = " |" & str标本类型 & "|" & IIf(varItem(11) = "Q", "1", "0")
190                             strItem = ""

                            End If
192                         If InStr(varItem(2), "^^^") > 0 And Trim(Replace(varItem(2), "^^^", "")) = Trim(str项目) Then
194                             If Trim(varItem(3)) <> "" Then
                                    '检测结果
196                                 strItem = strItem & "|" & varItem(1) & Replace(varItem(2), "^^^", "") & "|" & varItem(3)
198                                 str日期 = Mid(varItem(12), 1, 4) & "-" & Mid(varItem(12), 5, 2) & "-" & Mid(varItem(12), 7, 2) & " " & Mid(varItem(12), 9, 2) & ":" & Mid(varItem(12), 11, 2) & ":" & Mid(varItem(12), 13, 2)
                                End If
                            End If
                        End If
                    End If
                Next

200             If strItem <> "" And str标本号 <> "" Then
202                 If Len(str标本号) >= 12 Then
204                     strResult = strResult & "||" & str日期 & "|^^" & str标本号 & "|" & str检验人标本质控 & strItem
                    Else
206                     strResult = strResult & "||" & str日期 & "|" & str标本号 & "|" & str检验人标本质控 & strItem
                    End If
208                 If Not mstrCode Like "" Then
210                     blnEnd = True
                    End If
                End If
            End If
        Loop

    End If
212 If strResult <> "" Then
214     If Left(strResult, 2) = "||" Then strResult = Mid(strResult, 3)
    End If
    Exit Sub
hErr:
216 SaveLog "Analyse第" & CStr(Erl()) & "行," & Err.Description
End Sub


Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:解析需要向仪器发送的标本信息
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
        Dim intNO As Integer
        Dim j As Integer

100     intNO = 2
102     varSampleInfo = Split(strSampleInfo, "||")
104     SendSample = ""
106     blnSuccess = True

108     If iSendStep = 0 Then
110         For intloop = LBound(varSampleInfo) To UBound(varSampleInfo)
112             strLine = varSampleInfo(intloop)
114             If Val(Split(strLine, "|")(5)) > 0 Then
116                 strNo = Split(strLine, "|")(5)
                Else
118                 strNo = Split(strLine, "|")(1)
                End If

120             If strNo <> "" And UBound(Split(strLine, "|")) >= 10 Then
                    '指令开始
122                 str_Line = "1H|\^&||||||||ACL9000||P|1|" & Format(Now, "yyyyMMddHHmmss") & Chr(&HD) & Chr(&H3)
124                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                    '病人信息
126                 str_Line = "2P|1|||||||U||||||||||||||||||||||||||" & Chr(&HD) & Chr(&H3)
128                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
130                 strItems = ""
132                 For i = 10 To UBound(Split(strLine, "|"))
134                     If InStr(strItems & ",", "," & Mid(Split(strLine, "|")(i), 2) & ",") <= 0 Then
136                         strItems = strItems & "," & Mid(Split(strLine, "|")(i), 2)
138                         intNO = intNO + 1
140                         j = j + 1
                            '试验指令
142                         str_Line = intNO & "O|" & j & "|" & strNo & "||^^^" & Mid(Split(strLine, "|")(i), 2) & "|||||||||||^|||||||||||O||||||" & Chr(&HD) & Chr(&H3)
144                         strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","
                        End If
                    Next
                    '结束
146                 intNO = intNO + 1
148                 str_Line = intNO & "L|1|N" & Chr(&HD) & Chr(&H3)
150                 strSendCmd = strSendCmd & Chr(&H2) & str_Line & GetCheckCode(str_Line) & vbCrLf & ","

152                 iSendStep = 1    '开始通信
154                 mstrWaitSend = mstrWaitSend & "," & Chr(&H5) & "," & strSendCmd & Chr(&H4)

                End If
            Next
156         SendSample = GetSendData
        Else
158         If InStr(strResponse, Chr(&H6)) > 0 Then SendSample = GetSendData            '收到06
        End If
160     If mstrWaitSend = "" Then iSendStep = 0 '退出双向通
        Exit Function
hErr:
162     SaveLog "SendSample第" & CStr(Erl()) & "行," & Err.Description
End Function

Private Function GetSendData() As String
    If mstrWaitSend <> "" Then
        If Left(mstrWaitSend, 1) = "," Then
            mstrWaitSend = Mid(mstrWaitSend, 2)
        End If
        If InStr(mstrWaitSend, ",") > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(mstrWaitSend, ",") - 1)
            mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, ","))
        Else
            GetSendData = mstrWaitSend
            mstrWaitSend = ""
        End If
    End If

End Function
Private Function GetCheckCode(ByVal str_in As String) As String
'获取校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function

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

