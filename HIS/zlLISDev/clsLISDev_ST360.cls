VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ST360"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:KHB ST-360酶标仪
'厂商:科华
'作者:陈东
'用户:贵阳市肺科医院
Private mintStp As Integer
Private mstrReturn As String
Private mstrCode As String

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
    Dim strLine As String
    Dim strData(8, 12) As String, str无结果 As String
    Dim strCode As String, strTwo As String, strTmp As String
    Dim lngS As Long, lngE As Long, strIniFile As String
    Dim strWave As String   '返回的是一板还是二板
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream
    
    
    strIniFile = App.Path & "\St_360.ini"
    
    strWave = "NO"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "指令格式 = 一汽"
       objStream.WriteLine "数据格式 = NO"
       objStream.Close
       Set objStream = Nothing
    Else
        Set objStream = objFileSystem.OpenTextFile(strIniFile)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If strLine Like "*数据格式*=*" Then
                strWave = Trim(Split(strLine, "=")(1))
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
    End If
    
    
    For intLoop = 1 To 8
        For i = 1 To 12
            strData(intLoop - 1, i - 1) = "0.000"
        Next
    Next
    
    For intLoop = 0 To 7
        strLine = ""
        For i = 0 To 11
            strLine = strLine & ";" & strData(intLoop, i)
        Next
        If Trim(strLine) <> "" Then
            str无结果 = str无结果 & "|" & Mid(strLine, 2)
        End If
    Next

    If mstrReturn <> "" Then
        strResult = mstrReturn

        WriteLog "第二次返回仪器数据", strCode, strResult
        mstrReturn = ""
        Exit Sub
    End If
    
    mstrCode = mstrCode & strSource
    lngS = InStr(mstrCode, "A")
    If lngS > 5 Then lngS = lngS - 5
    lngE = 0
    If lngS > 0 Then
        lngE = InStr(lngS, mstrCode, "H")
        If strWave = "NO" Then
            If lngE > 0 Then
                lngE = InStr(lngE, mstrCode, "A")
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, "H")
                    If lngE > 0 Then
                        lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                    Else
                        lngE = 0
                    End If
                Else
                    lngE = 0
                End If
            Else
                lngE = 0
            End If
        Else
            If lngE > 0 Then
                lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
            Else
                lngE = 0
            End If
        End If
    End If
    
    Do While lngS > 0 And lngE > 0
        strCode = Mid$(mstrCode, lngS, lngE - lngS + 2)
        mstrCode = Mid$(mstrCode, lngE + 2)
        '--------------------------------------------------------------------
        
            '该仪器一次就将两次检测的数据发回,这里分成两步,把数据返回
             
            
            If UBound(Split(strCode, Chr(&HD) & Chr(&HA))) < 9 Then Exit Sub
            
            '开始解析
            strResult = "": strTmp = ""
            For intLoop = LBound(Split(strCode, Chr(&HD) & Chr(&HA))) To UBound(Split(strCode, Chr(&HD) & Chr(&HA)))
                strLine = Trim(Split(strCode, Chr(&HD) & Chr(&HA))(intLoop))
                If InStr(",A,B,C,D,E,F,G,H,", "," & Trim(Mid(strLine, 1, 1)) & ",") > 0 Then
                    
                    If strTwo <> "ABCDEFGH" Then
                        '第一次返回的结果
                        If InStr(strTwo, Trim(Mid(strLine, 1, 1))) <= 0 Then strTwo = strTwo & Trim(Mid(strLine, 1, 1))
                        varline = Split(strLine, " ")
                        For i = LBound(varline) + 1 To UBound(varline)
                            If i <= 12 Then
                               strData(intLoop - 1, i - 1) = Format(Val(varline(i)), "0.000")
                            End If
                        Next
                    Else
                        '第二次返回的结果
                        strTmp = ""
                        varline = Split(strLine, " ")
                        For i = LBound(varline) + 1 To UBound(varline)
                            If i <= 12 Then
                                strTmp = strTmp & ";" & Format(Val(varline(i)), "0.000")
                            End If
                        Next
                        If Trim(strTmp) <> "" Then
                            mstrReturn = mstrReturn & Mid(strTmp, 2) & "|"
                        End If
                    End If
                End If
            Next

            
            For intLoop = 0 To 7
                strLine = ""
                For i = 0 To 11
                    strLine = strLine & ";" & strData(intLoop, i)
                Next
                If Trim(strLine) <> "" Then
                    strResult = strResult & "|" & Mid(strLine, 2)
                End If
            Next
            
            
            If strResult <> "" Then
                If str无结果 = strResult Then
                    strResult = ""
                Else
                    strResult = Mid(strResult, 2)
                End If
            End If
            WriteLog "第一次返回仪器数据", strCode, strResult
        
        '--------------------------------------------------------------------
        lngS = InStr(mstrCode, "A")
        lngE = 0
        If lngS > 0 Then
            lngE = InStr(lngS, mstrCode, "H")
            If strWave = "NO" Then
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, "A")
                    If lngE > 0 Then
                        lngE = InStr(lngE, mstrCode, "H")
                        If lngE > 0 Then
                            lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                        Else
                            lngE = 0
                        End If
                    Else
                        lngE = 0
                    End If
                Else
                    lngE = 0
                End If
            Else
                If lngE > 0 Then
                    lngE = InStr(lngE, mstrCode, Chr(&HD) & Chr(&HA))
                Else
                    lngE = 0
                End If
            End If
        End If
    Loop
    
        
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
    Dim strIniFile As String
    Dim objFileSystem As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String
    Dim strType As String
    
    strIniFile = App.Path & "\St_360.ini"
    strType = "一汽"
    If Not objFileSystem.FileExists(strIniFile) Then
       Set objStream = objFileSystem.CreateTextFile(strIniFile)
       objStream.WriteLine "指令格式 = 一汽"
       objStream.WriteLine "数据格式 = NO"
       objStream.Close
       Set objStream = Nothing
    Else
        Set objStream = objFileSystem.OpenTextFile(strIniFile)
        Do While Not objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If strLine Like "*指令格式*=*" Then
                strType = Trim(Split(strLine, "=")(1))
            End If
        Loop
        objStream.Close
        Set objStream = Nothing
    End If

    Select Case strCmdType
        Case "Begin"                    '开始 出板
            CmdAnalyse = ""
        Case "End"                      '结束 进板
            CmdAnalyse = ""
        Case "Revert"                   '回复命令
            CmdAnalyse = "e_end"
        Case "Play"                     '整板测量
            If mintStp = 0 Then
                CmdAnalyse = "" 'CmdAnalyse = "e(0)|e_end"
            Else
                CmdAnalyse = ""
            End If
        Case "ReadData_TimeOut"
            CmdAnalyse = 90
        Case "ReadData"                 '读板
            If mintStp = 1 Then
                CmdAnalyse = "|"         '空指令
                mintStp = 0
            Else
                If strType = "一汽" Then
                    CmdAnalyse = "p(1,3)"
                Else
                    CmdAnalyse = "p(2,3,7)"
                End If
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
        Case "e(0)"
            CmdAnalyse = "e(0)"
        Case "e(1)"
            CmdAnalyse = "e(1)"
        Case "e(2)"
            CmdAnalyse = "e(2)"
        Case "e(3)"
            CmdAnalyse = "e(3)"
        Case "e(4)"
            CmdAnalyse = "e(4)"
        Case "e(5)"
            CmdAnalyse = "e(5)"
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




