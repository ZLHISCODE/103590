VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Immage_800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: Immage800
'厂商:贝克曼
'作者:曾超
'用户:重庆市三院,丹东中心医院
'2012-10-17 陈东 增加双向
Private mstrCode As String
Private mstrWaitSend As String

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    'GetStartCmd = Chr(6)
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
    
        Dim varLine As Variant
        Dim iRow As Integer, str检验人标本质控 As String, strTmp As String, strLine As String, str项目 As String
        Dim strValues As String
        Dim blnEnd As Boolean '是否已解析完

        Dim strNowData As String
        Dim strData As String
        Dim intloop As Integer
        Dim varItem As Variant
        Dim strDate As String, strSampleNo As String
        Dim lng_S As Long, lng_E As Long
        Dim strItem As String
        Dim varTmp As Variant
        On Error GoTo hErr
        '初始化返回值
100     strResult = "": strReserved = "": strCmd = ""
    
102     If InStr(strSource, Chr(&H5)) > 0 Then strCmd = Chr(&H6)
104     If InStr(strSource, Chr(&HA)) > 0 Then strCmd = Chr(&H6)
    
106     mstrCode = mstrCode & strSource
    
        '没有找到开始符和结束符时退出
108     lng_S = InStr(mstrCode, Chr(&H5))  '-"H|\^&"
    '110     lng_S = lng_S - 2
110     lng_E = 0
112     If lng_S > 0 Then lng_E = InStr(lng_S, mstrCode, Chr(&H4))
    '116         lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '"L|1|N"
    '118         If lng_E > 0 Then
    '120             lng_E = InStr(lng_E, mstrCode, Chr(&HA))
    '            Else
    '122             lng_E = 0
    '            End If
    '        End If
    
114     If Not (lng_S > 0 And lng_E > 0) Then Exit Sub
116     SaveLog "800", "数据完整，开始简码", strCmd
        '取得当前一段，剩下的返回。下次再解析。
118     strNowData = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
120     mstrCode = Mid(mstrCode, lng_E + 1)
        
        '开始解析
122     varLine = Split(strNowData, vbCrLf)
124     strItem = ""
126     For intloop = LBound(varLine) To UBound(varLine)
128         If InStr(varLine(intloop), "|") > 0 Then
130             varItem = Split(varLine(intloop), "|")
132             strTmp = UCase(Right(varItem(0), 1))
134             Select Case strTmp
                    Case "H"
136                     strTmp = Trim(varItem(13))
138                     strDate = ""
140                     If Len(strTmp) >= 14 Then
142                         strDate = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                                      Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        End If
144                 Case "Q"
146                     strTmp = varItem(2)
148                     varTmp = Split(strTmp, "\")
150                     For iRow = LBound(varTmp) To UBound(varTmp)
152                         strTmp = varTmp(iRow)
154                         If IsNumeric(Mid(strTmp, 2)) And Mid(strTmp, 2) >= 12 Then
156                             strResult = strResult & "||" & Format(Now, "yyyy-MM-dd") & "|^^" & Mid$(strTmp, 2)
                            End If
                        Next
158                     If strResult <> "" Then strCmd = "1|"
160                     SaveLog "800", "双向" & strResult, strCmd
162                 Case "O"                '标本号
164                     If Not IsDate(strDate) Then
166                         strTmp = Trim(varItem(7))
168                         strDate = Mid(strTmp, 1, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & _
                                      Mid(strTmp, 9, 2) & ":" & Mid(strTmp, 11, 2) & ":" & Mid(strTmp, 13, 2)
                        End If
                        
170                     strSampleNo = Trim(Split(varItem(2), "^")(0))
172                     If IsNumeric(strSampleNo) And Len(strSampleNo) >= 12 Then strSampleNo = "^^" & strSampleNo
174                     strData = strDate & "|" & strSampleNo & "| |Blood|0"
176                 Case "R"
178                     strItem = strItem & "|" & Split(varItem(2), "^")(3) & "|" & varItem(3)
                End Select
            End If
        Next
180     If strItem <> "" And strData <> "" Then
182         strResult = strResult & "||" & strData & strItem
        End If
        
184     If strResult <> "" Then strResult = Mid(strResult, 3)
        Exit Sub
hErr:
186     SaveLog "immage800", CStr(Erl()) & "," & Err.Description, ""
End Sub

Public Function SendSample(ByVal strSample As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    ''   strSample:解析需要向仪器发送的标本信息
    '           标本之间以||分隔
    '           元素之间以|分隔
    '           第0个元素：检验时间
    '           第1个元素：样本序号
    '           第2个元素：检验人
    '           第3个元素：标本
    '           第4个元素：是否质控品
    '           第5～9元素：系统保留
    '           从第10个元素开始为需要的检验项目。
    '   iSendStep  : 步骤
    '   blnSuccess : 是否传输失败
    '   strResponse: 仪器的应答信息
    '   blnUndo    : 是否取消.
    '
    '   返回值：发到仪器的字串或指令。
    Dim varInData As Variant, iLoop As Integer, strData As String, i As Integer
    Dim strBarcode As String, strEmergency As String, strTmp As String
    Dim strChannel As String, strSend As String, strItems As String, iRow As Integer
    Dim strSegment As String, strDate As String, strSampleType As String
    Dim strSampeNo As String
    On Error GoTo hErr

100 SendSample = ""
102 blnSuccess = True
104 If iSendStep = 0 Then
        '刚进入双向模式，还没有生成发送数据

106     SaveLog "Immage800-进入双向，生成待发送数据", strSample, ""
108     varInData = Split(strSample, "||")
110     For iLoop = LBound(varInData) To UBound(varInData)
112         strData = varInData(iLoop)
114         strDate = Format(CDate(Split(strData, "|")(0)), "yyyyMMddHHmmss")
116         strSampeNo = Trim(Split(strData, "|")(1))
            '--标本 S'=血清,'U' =随机尿,'F' =脑脊髓液,“P” =血浆,'T' =定时尿,注：空默认采样类型从系统设置
118         strSampleType = Trim(Split(strData, "|")(3))
120         If InStr(strSampleType, "血浆") > 0 Then
122             strSampleType = "P"
124         ElseIf InStr(strSampleType, "定时尿") > 0 Then
126             strSampleType = "T"
128         ElseIf InStr(strSampleType, "尿") > 0 Then
130             strSampleType = "U"
132         ElseIf InStr(strSampleType, "血清") > 0 Then
134             strSampleType = "S"
            Else
136             strSampleType = ""
            End If
            '--条码
138         strBarcode = Trim(Split(strData, "|")(5))
            '---急诊
140         If Val(Split(strData, "|")(4)) = 1 Then
142             strEmergency = "S"
            Else
144             strEmergency = "R"
            End If
            '----指标
         
146         strItems = ""
148         For i = 10 To UBound(Split(strData, "|"))
150             strChannel = Trim$("" & Split(strData, "|")(i))
152             strItems = strItems & "^^^" & strChannel & "\"
            Next
        
154         strSend = ""
156         If strItems <> "" And strBarcode <> "" Then
158             If Right(strItems, 1) = "\" Then strItems = Left(strItems, Len(strItems) - 1)
160             iRow = 0
162             strSegment = Chr(&H5)
164             iRow = iRow + 1
166             strSend = iRow & "H|\^&" & Chr(&HD) & Chr(&H3)
168             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
                'P|1||A43749||Smith^Tom^J||19950315^120^H|性别M/F/U|||||医生||||||||||||位置<CR>
                'P|1||标本号|<CR>
170             iRow = iRow + 1
172             strSend = iRow & "P|1||" & strSampeNo & "|" & Chr(&HD) & Chr(&H3)
174             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
                'O|1|SID_00002^3^1^^1||^^^IGG^^N^2\^^^IGA|S|||19950316135300|||N||||S|||1^1.00|||||||O<CR>
176             iRow = iRow + 1
178             strSend = iRow & "O|1|" & strBarcode & "||" & strItems & "|" & strEmergency & "||" & _
                          IIf(InStr("T,U", strSampleType) > 0, strDate, "") & _
                          "||||N||||" & strSampleType & "|||1^1.00" & Chr(&HD) & Chr(&H3)
180             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA)
                
182             iRow = iRow + 1
184             strSend = iRow & "L|1|N" & Chr(&HD) & Chr(&H3)
186             strSegment = strSegment & Chr(&H2) & strSend & GetCheckCode(strSend) & Chr(&HD) & Chr(&HA) & Chr(&H4)
                
188             If strSegment <> "" Then mstrWaitSend = mstrWaitSend & strSegment
            End If

190         SendSample = GetSendData
192         If mstrWaitSend <> "" Then iSendStep = 1    '还有未发送数据，不退出双向模式

        Next
194     SaveLog "image800-进入双向，生成待发送数据", "发送:" & SendSample, iSendStep & "缓存:" & mstrWaitSend
    Else
        '收到仪器应答，从缓存中取得指令发送
196     SaveLog "image800-双向，发送数据 ", strResponse, iSendStep
198     If InStr(strResponse, Chr(&H6)) > 0 Then

200         SendSample = GetSendData
202         If mstrWaitSend = "" Then iSendStep = 0
204         SaveLog "image800-双向，发送数据", "发送:" & SendSample, iSendStep & "缓存:" & mstrWaitSend
206     ElseIf strResponse = Chr(&H5) Then
208         SaveLog "image800-双向，发送数据", "收到05:" & SendSample, iSendStep & "缓存:" & mstrWaitSend
210         If InStr(mstrWaitSend, Chr(&H5)) > 0 Then
212             mstrWaitSend = Mid(mstrWaitSend, InStr(mstrWaitSend, Chr(&H5)))
            Else
214             mstrWaitSend = ""
            End If
216         iSendStep = 0
218         SendSample = Chr(&H6)
        End If

    End If
    Exit Function
hErr:
220     SaveLog "image800-双向", strResponse, CStr(Erl()) & "行," & Err.Description
End Function


Private Sub SaveLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
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
    objStream.WriteLine (String(50, "≡"))
    objStream.WriteLine ("执行时间:" & strDate & "版本:" & App.major & "." & App.minor & "." & App.Revision)
    objStream.WriteLine ("Func:" & strFunc)
    objStream.WriteLine ("In:" & StrInput)
    objStream.WriteLine ("Out:" & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub
Private Function GetCheckCode(ByVal str_in As String) As String
    '计算校验码
    Dim lngAsc As Long, i As Integer
    For i = 1 To Len(str_in)
        lngAsc = lngAsc + Asc(Mid(str_in, i, 1))
    Next
    GetCheckCode = Right("0" & CStr(Hex(lngAsc Mod 256)), 2)
End Function
Private Function GetSendData() As String
    '从缓存中获取一段数据，返回。如果缓存中没有数据，返回空串

    
    If mstrWaitSend = "" Then Exit Function
    
    If Left$(mstrWaitSend, 1) = Chr(&H5) Then
        GetSendData = Chr(&H5)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H2) Then
        If InStr(1, mstrWaitSend, Chr(&HA)) > 0 Then
            GetSendData = Mid(mstrWaitSend, 1, InStr(1, mstrWaitSend, Chr(&HA)))
            mstrWaitSend = Mid$(mstrWaitSend, InStr(1, mstrWaitSend, Chr(&HA)) + 1)
        End If
    ElseIf Left$(mstrWaitSend, 1) = Chr(&H4) Then
        GetSendData = Chr(&H4)
        mstrWaitSend = Mid(mstrWaitSend, 2)
    End If

End Function

