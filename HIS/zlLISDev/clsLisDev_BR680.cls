VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BR680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BIO-RAD 680酶标仪
'厂商:伯乐
'作者:陈东
'用户:大连机车厂职工医院

Private mstrCode As String
Private mintStp As Integer   '保存调用次数
Private mstrReturn As String '保存参考波长

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
        '设备标准接口
        '功能：解析数据
        '参数：
    
        Dim intLoop As Integer, i As Integer
        Dim varline As Variant, strTmp As String
        Dim strLine As String
        Dim strData(8, 12) As String, str无结果 As String
        Dim lng_S As Long, lng_E As Long, intRow As Integer, intCol As Integer
        On Error GoTo errH
100     mstrCode = mstrCode & strSource
        strReserved = ""
102     strResult = ""
104     If mstrCode Like "*begin*end*begin*end" & Chr(&HD) & "*" Then
            WriteLog "analyse", mstrCode, strResult
            '---主波长结果
106         lng_S = InStr(mstrCode, "begin" & Chr(&HD))
108         lng_E = InStr(lng_S, mstrCode, "end" & Chr(&HD)) + 3
110         strTmp = Mid$(mstrCode, lng_S, lng_E - lng_S + 1)
112         mstrCode = Mid$(mstrCode, lng_E + 1)
            
114         varline = Split(strTmp, Chr("&H0D"))
116         intRow = 0: intCol = 0
118         For i = LBound(varline) To UBound(varline)
120             strTmp = varline(i)
122             strLine = ""
124             If Len(strTmp) = 6 * 12 Then
126                 Do While strTmp <> ""
128                     strData(intRow, intCol) = Val(Mid$(strTmp, 1, 6))
130                     intCol = intCol + 1
132                     If intCol > 11 Then
134                         intRow = intRow + 1
136                         intCol = 0
                        End If
                        'strLine = strLine & ";" & Val(Mid$(strTmp, 1, 6))
138                     strTmp = Mid$(strTmp, 7)
                    Loop
                End If

            Next
        
            '参考波长结果
142         lng_S = InStr(mstrCode, "begin" & Chr(&HD))
144         lng_E = InStr(lng_S, mstrCode, "end" & Chr(&HD)) + 3
146         strTmp = Mid$(mstrCode, lng_S, lng_E - lng_S + 1)
148         mstrCode = Mid$(mstrCode, lng_E + 1)
        
150         varline = Split(strTmp, Chr("&H0D"))
152         intRow = 0: intCol = 0
154         For i = LBound(varline) To UBound(varline)
156             strTmp = varline(i)
158             strLine = ""
160             If Len(strTmp) = 6 * 12 Then
162                 Do While strTmp <> ""
164                     strData(intRow, intCol) = Val(strData(intRow, intCol)) - Val(Mid$(strTmp, 1, 6))
166                     intCol = intCol + 1
168                     If intCol > 11 Then
170                         intRow = intRow + 1
172                         intCol = 0
                        End If
174                     strTmp = Mid$(strTmp, 7)
                    Loop
                End If

            Next

176         For intRow = 0 To 7
178             strLine = ""
180             For intCol = 0 To 11
182                 strLine = strLine & ";" & strData(intRow, intCol)
                Next
184             strResult = strResult & "|" & Mid(strLine, 2)
            Next
186         strResult = Mid$(strResult, 2)
        End If
        Exit Sub
errH:
192     WriteLog "BR680.Analyse", CStr(Erl()) & "行," & Err.Description & vbCrLf & mstrCode, strResult
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '解析需要向仪器发送的标本信息
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   第5～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。
    SendSample = ""
End Function

Public Function CmdAnalyse(strCmdType As String) As String
    '功能:  分解命令为仪器执行的命令
    Dim lngTime As Long

    Select Case strCmdType
        Case "Begin"                    '开始 出板
            CmdAnalyse = "EIA.READER AQ" & Chr(&HD) & "|" & "ERE 8073" & Chr(&HD) & "|" & "ERE 0000" & Chr(&HD)
        Case "End"                      '结束 进板
            CmdAnalyse = "EIA.READER RL" & Chr(&HD) & "|" & "ERE 0000" & Chr(&HD)
        Case "Revert"                   '回复命令
            CmdAnalyse = ""
        Case "Play"                     '整板测量
            If mintStp = 1 Then
                CmdAnalyse = "|"         '空指令
            Else
                CmdAnalyse = "EIA.READER ISTATUS" & Chr(&HD) & "|" & "ERE 0000 0 0 **.* **.*" & Chr(&HD)
            End If
        Case "ReadData_TimeOut"
            CmdAnalyse = 80
        Case "ReadData"                 '读板
            If mintStp = 1 Then
                CmdAnalyse = "|"         '空指令
                mintStp = 0
            Else
                CmdAnalyse = "EIA.READER RPLATE 1 5 H 2 4" & Chr(&HD)
                mintStp = 1
            End If
        Case "Out"                      '进板
            CmdAnalyse = ""
        Case "In"                       '出板
            CmdAnalyse = ""
        '以上指令为固定项,必须保留,如仪器没有对应指令,设为空串即可
        '----------------------------------------------------------------------------------
        Case "450"                      '波长一
            CmdAnalyse = ""
        Case "630"                      '波长二
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '进板方式,按检验仪器参数设置中的进板方式设定对应的指令
        Case "连续"                     '连续
            CmdAnalyse = ""
        Case "单步"                     '逐步
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '振板频率,按检验仪器参数设置中的振板频率设定对应的指令
        Case "低"                       '
            CmdAnalyse = ""
        Case "高"                       '
            CmdAnalyse = ""
        '----------------------------------------------------------------------------------
        '                               '空白形式,按检验仪器参数设置中的空白形式设定对应的指令
        Case "空气"
            CmdAnalyse = ""
        Case "试剂"
            CmdAnalyse = ""
    End Select
    
End Function


