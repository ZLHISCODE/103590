VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_MicroScan_AS4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Microscan AS-4 微生物分析仪
'厂商:美国德灵(DADE BEHRING INC)
'作者:曾超
'用户:长春一汽总医院
'时间:20100508
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)
    '设备标准接口
    '功能：解析数据
    Dim str_Code As String, lng_S As Long, lng_len As Long
    Dim str_Day As String, str_日期 As String, str标本号 As String, strValues As String
    Dim int_i As Integer, strRecord As String
    Dim str_WBC As String, str_RBC As String, str_PLT As String
    Dim str_Data As String
    Dim strNow As String
    Dim strLine() As String
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strItem() As String
    Dim strData As String
    
    
    
    If InStr(strSource, Chr(&H2) & Chr(&H22) & Chr(&H48) & Chr(&H22)) <= 0 Or InStr(strSource, Chr(&H4)) <= 0 Then
        strReserved = strSource
        Exit Sub
    End If
    strNow = Mid$(strSource, InStr(strSource, Chr(&H2) & Chr(&H22) & Chr(&H48) & Chr(&H22)), InStr(strSource, Chr(&H4)))
    strReserved = Mid$(strSource, InStr(strSource, Chr(&H4)))
    
    
    
    '开始解析
    strLine = Split(strNow, vbCrLf)
    For intLoop = 0 To UBound(strLine)
        '去掉回车换行
        strTmp = Replace$(strLine(intLoop), vbLf, "")
        strTmp = Replace$(strTmp, vbCr, "")
        strTmp = Replace$(strTmp, Chr(&H2), "")
        strTmp = Replace$(strTmp, Chr(&H3), "")
        
        If InStr(strTmp, ",") > 0 Then
            strItem = Split(strTmp, ",")
            Select Case Mid$(strTmp, 1, 1)
            
                Case "R"
                    If UBound(strItem) >= 12 Then
                        strData = strData & "||" & Mid(strItem(6), 1, 4) & "-" & Mid(strItem(6), 5, 2) & "-" & Mid(strItem(6), 7, 2) & " " & Format(Now, "hh:mm:ss") & _
                                "|" & Replace$(strItem(3), """", "") & "| |" & Replace$(strItem(12), """", "") & "|0"
                    End If
                Case "M"
                    '项目
                    If UBound(strItem) >= 6 Then
                        strData = strData & "|" & Replace(strItem(3), """", "") & "|" & Replace(strItem(7), """", "") & "^1^" & Replace(strItem(4), """", "")
                    End If
            End Select
        End If
    Next
    If strData <> "" Then strResult = Mid(strData, 3)
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



Private Function Graph_Text(ByVal str_bin As String) As String
    '生成图形文本数据
    Dim lng_len As Long, lng_i As Long, str_bit As String
    Dim str_Return As String
    lng_len = Len_Bin(str_bin)

    If lng_len < 1 Then
        Graph_Text = ""
        Exit Function
    End If

    For lng_i = 1 To lng_len
        str_bit = Replace(Mid_bin(str_bin, lng_i, 1, False), ",", "")
    
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format(CByte("&H" & str_bit), "000")
        End If
    Next
    Graph_Text = str_Return
End Function

