VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_KJ680"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器：Bio-rad伯乐 全自动酶标仪
'型号：KJ680
'医院：保山市中医院

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
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
    Dim varLine As Variant
    Dim strLine As String, strtmp As String
    Dim intloop As Integer, i As Integer

    If InStr(1, strSource, "end") = 0 Then
        strReserved = strSource
        Exit Sub
    End If

    '开始解析
    strSource = Mid(strSource, InStr(1, strSource, "begin"), InStr(1, strSource, "end") - InStr(1, strSource, "begin"))
    varLine = Split(strSource, Chr(&HD))

    For intloop = 1 To 8
        strLine = ""
        strtmp = Split(strSource, Chr(&HD))(intloop)
        For i = 1 To 12
            strLine = strLine & ";" & Trim(Mid(strtmp, 1 + (i - 1) * 6, 6))
        Next
        If Trim(strLine) <> "" Then
            strResult = strResult & "|" & Mid(strLine, 2)
        End If
    Next

    If strResult <> "" Then strResult = Mid(strResult, 2)

    '------------------------------------------------
    Exit Sub
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
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

Private Sub MoveToNextLine(ByVal strSource, iPoint As Integer)
    '指针定位到下一个换行之后
    Do While iPoint <= Len(strSource)
        If Mid(strSource, iPoint, 1) = Chr(10) Then
            iPoint = iPoint + 1
            Exit Do
        End If
        iPoint = iPoint + 1
    Loop
End Sub

Private Function ReadSubString(ByVal strSource As String, iPoint As Integer, Optional ByVal Length As Integer = 0) As String
    '读取子串，同时指针移动到读取的字符串之后
    On Error Resume Next
    If Length = 0 Then
        ReadSubString = Mid(strSource, iPoint)
        iPoint = Len(strSource) + 1
    Else
        ReadSubString = Mid(strSource, iPoint, Length)
        iPoint = iPoint + Length
    End If
End Function

Private Function ReadLine(ByVal strSource As String, iPoint As Integer) As String
    '读取一行字符，并将字符指针移到下一行开始位置
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, Chr(10))
    If iPos = 0 Then iPos = Len(strSource) + 1
    
    ReadLine = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

Private Function ReadItemName(ByVal strSource As String, iPoint As Integer) As String
    '读取一个检验项目名称，并将字符指针移到项目名称之后
    Dim iPos As Integer
    
    iPos = InStr(iPoint, strSource, " ")
    If iPos = 0 Then ReadItemName = "": iPoint = Len(strSource) + 1: Exit Function
    
    ReadItemName = Mid(strSource, iPoint, iPos - iPoint)
    iPoint = iPos + 1
End Function

'

Public Function CmdAnalyse(strCmdType As String) As String
    '功能:  分解命令为仪器执行的命令
    
'    strControl = Me.cbo波长.Text & ";" & Me.cbo参考波长.Text & ";" & Me.cbo振板频率.Text & ";" & Me.txt振板时间 & _
'                 ";" & Me.cbo进板方式.Text & ":" & Me.cbo空白形式.Text
    
    Dim lngTime As Long
    Select Case strCmdType
        Case "Begin"                   '开始
            CmdAnalyse = ""
        Case "End"                      '结束
            CmdAnalyse = ""
        Case "Revert"                   '回复命令
            CmdAnalyse = ""
        Case "Play"                     '整板测量
            CmdAnalyse = ""
        Case "ReadData_TimeOut"
            CmdAnalyse = 90
        Case "ReadData"                  '整板测量
            CmdAnalyse = "EIA.READER RPLATE 2 0 H 2" & Chr(&HD)
        Case "450"
            CmdAnalyse = ""   '450波片
        Case "630"
            CmdAnalyse = ""   '2与波片为"450"
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
        Case "Out"                      '进板
            CmdAnalyse = ""
        Case "In"                       '出板
            CmdAnalyse = ""
        Case "SpecialConnType"
            CmdAnalyse = ""
        Case "Z5_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z10_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z15_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z20_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z25_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case "Z30_TimeOut"
            CmdAnalyse = Val(Mid$(strCmdType, 2, 2)) + 5
        Case Else
            If Left(strCmdType, 1) = "Z" Then
                '振板时间
                lngTime = Val(Mid(strCmdType, 2))
                If lngTime >= 0 And lngTime <= 60 Then
                    CmdAnalyse = "Z(" & Format(lngTime, "00") & ")"
                End If
            End If
        End Select
End Function
