VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ZS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:航天之星ZS-2酶标仪
'厂商:上海荣盛
'作者:莫运维
'用户:钦北区人民医院
'时间:2013/12/19
Private mintStp As Integer


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
    Dim varline As Variant
    Dim strLine As String, j As Integer
    Dim strData(8, 12) As String, str无结果 As String
    Dim lng_S As Long, lng_E As Long
    Dim str结果2 As String, str结果3 As String, str结果4 As String, str结果5 As String
    Dim str结果6 As String, str结果7 As String, str结果8 As String, str结果1 As String

    On Error GoTo errH
    If mintStp = 1 Then
        strReserved = strSource
        lng_S = InStr(strReserved, "")
        If lng_S > 0 Then
            lng_E = InStr(strReserved, "")
        End If
        If lng_S > 0 And lng_E > 0 Then
            strReserved = Replace(strReserved, "-", " ")
            varline = Split(strReserved, Chr(&H20))(0)
            strLine = Split(strReserved, Chr(&H20))(1)
            If strLine > 0 Then
                For i = 1 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果1 = str结果1 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 2 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果2 = str结果2 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 3 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果3 = str结果3 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 4 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果4 = str结果4 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 5 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果5 = str结果5 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 6 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果6 = str结果6 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 7 To 96
                    strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    str结果7 = str结果7 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                    i = i + 8 - 1
                Next
                For i = 8 To 96
                    If i = 96 Then
                        strLine = Format(Mid(Replace(Split(strReserved, Chr(&H20))(i), "-", " "), 1, 4) * 0.001, "0.000")
                        str结果8 = str结果8 & ";" & Format(Mid(Replace(Split(strReserved, Chr(&H20))(i), "-", " "), 1, 4) * 0.001, "0.000")
                    Else
                        strLine = Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                        str结果8 = str结果8 & ";" & Format(Split(strReserved, Chr(&H20))(i) * 0.001, "0.000")
                        i = i + 8 - 1
                    End If
                Next
            Else
                '以下是处理酶标仪传回全是O的结果
                For i = 1 To 96
                    strLine = Format(0, "0.000")
                    str结果1 = str结果1 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 2 To 96
                    strLine = Format(0, "0.000")
                    str结果2 = str结果2 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 3 To 96
                    strLine = Format(0, "0.000")
                    str结果3 = str结果3 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 4 To 96
                    strLine = Format(0, "0.000")
                    str结果4 = str结果4 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 5 To 96
                    strLine = Format(0, "0.000")
                    str结果5 = str结果5 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 6 To 96
                    strLine = Format(0, "0.000")
                    str结果6 = str结果6 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 7 To 96
                    strLine = Format(0, "0.000")
                    str结果7 = str结果7 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
                For i = 8 To 96
                    strLine = Format(0, "0.000")
                    str结果8 = str结果8 & ";" & Format(0, "0.000")
                    i = i + 8 - 1
                Next
            End If
            strResult = Mid(str结果1, 2) & "|" & Mid(str结果2, 2) & "|" & Mid(str结果3, 2) & "|" & Mid(str结果4, 2) & "|" & Mid(str结果5, 2) & "|" & Mid(str结果6, 2) & "|" & Mid(str结果7, 2) & "|" & Mid(str结果8, 2) & "|"
        End If
    End If
    Exit Sub
errH:
    WriteLog "ZS2", "zs", Err.Number & ":" & Err.Description
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

'MsgBox strCmdType
    Select Case strCmdType
    Case "Begin"                    '开始 出板
        CmdAnalyse = ""
    Case "End"                      '结束 进板
        CmdAnalyse = ""
    Case "Revert"                   '回复命令
        CmdAnalyse = ""
    Case "Play"                     '整板测量
        If mintStp = 0 Then
            CmdAnalyse = ""
        Else
            CmdAnalyse = ""
        End If
    Case "ReadData_TimeOut"
        CmdAnalyse = 50
    Case "ReadData"                 '读板
        If mintStp = 0 Then
            CmdAnalyse = Chr(&H1B) & Chr(&H56) & Chr(&H55) & Chr(&H2) & Chr(&H4) & Chr(&HD)
            mintStp = 1
        Else
            Sleep (3000)
            CmdAnalyse = ""
            mintStp = 1
        End If
    Case "Out"                      '进板
        CmdAnalyse = ""
    Case "In"                       '出板
        CmdAnalyse = ""
        '以上指令为固定项,必须保留,如仪器没有对应指令,设为空串即可
        '----------------------------------------------------------------------------------
        '                               '波长,按检验仪器参数设置中的波长设定对应的指令
    Case "450"                      '波长一
        CmdAnalyse = ""
    Case "620"                      '波长二
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
        '----------------------------------------------------------------------------------
        '                               '空白形式,按检验仪器参数设置中的空白形式设定对应的指令
    Case "空气"
        CmdAnalyse = ""
    Case "试剂"
        CmdAnalyse = ""
    Case Else

    End Select

End Function












