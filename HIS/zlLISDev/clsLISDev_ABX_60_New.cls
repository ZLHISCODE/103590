VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ABX_60_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ABX Micros 60
'厂商:ABX
'作者:欧海胜
'用户:东莞莞华医院

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
    
    Dim lngS As Long, lngE As Long, strCode As String
    Dim varLine As Variant, intRow As Integer, strLine As String
    Dim str标本号 As String, str日期 As String, strItems As String, lngItem As String
    Dim StrWbc As String, StrRbc As String, StrPLt As String
    Dim strImg As String
     lngS = InStr(strSource, Chr(2))
    lngE = InStr(strSource, Chr(3))
    strReserved = strSource
    If lngS > 0 And lngE > 0 And lngE > lngS Then
        strCode = Mid(strReserved, lngS, lngE - lngS + 1)
        strReserved = Mid(strReserved, lngE + 1)
        varLine = Split(strCode, Chr(&HD))
        For intRow = LBound(varLine) To UBound(varLine)
            strLine = varLine(intRow)
            If InStr("!,2,3,4,5,6,7,8,@,A,B,C,#,%,(,),*,+,-,/,.,',"""",$,&", Mid(strLine, 1, 1)) > 0 Then
                '通道码特殊符号转换
                lngItem = Mid(strLine, 1, 1)
                If lngItem = "@" Then
                   lngItem = "PLT"
                ElseIf lngItem = "'" Then
                   lngItem = "GRANP"
                ElseIf lngItem = """" Then
                   lngItem = "LYMNA"
                ElseIf lngItem = "$" Then
                   lngItem = "MONNA"
                ElseIf lngItem = "&" Then
                   lngItem = "MONNP"
                End If
                strItems = strItems & "|" & lngItem & "|" & Val(Mid(strLine, 3))
           ElseIf Mid(strLine, 1, 1) = "q" Then
                '日期
                str日期 = Format(CDate(Mid(strLine, 9, 2) & "-" & Mid(strLine, 6, 2) & "-" & Mid(strLine, 3, 2)), "yyyy-MM-dd")
            ElseIf Mid(strLine, 1, 1) = "u" Then
                '标本号
                str标本号 = Val(Mid(strLine, 3))
           ElseIf Mid(strLine, 1, 1) = "W" Then    'WBC图形数据
                StrWbc = Mid(strLine, 2)
                StrWbc = Graph_Text(StrWbc)
            ElseIf Mid(strLine, 1, 1) = "X" Then    'RBC图形数据
                StrRbc = Mid(strLine, 2)
                StrRbc = Graph_Text(StrRbc)
            ElseIf Mid(strLine, 1, 1) = "Y" Then    'PLT图形数据
                StrPLt = Mid(strLine, 2)
                StrPLt = Graph_Text(StrPLt)
           
            End If
        Next
        
        If Val(str标本号) > 0 And strItems <> "" Then
            strResult = str日期 & "|" & str标本号 & "| |血液|0" & strItems
           If StrWbc <> "" Then strImg = strImg & "^" & "WBC;0;140,135,19-50,32.27-100,68-200,96.75-300,129-400" & StrWbc
            If StrRbc <> "" Then strImg = strImg & "^" & "RBC;0;140,135,15.45-30,55-100,103-200" & StrRbc
            If StrPLt <> "" Then strImg = strImg & "^" & "PLT;0;30,135,14-2,30-5,47-10,86-20,129-30" & StrPLt
            If strImg <> "" Then
                strResult = strResult & vbCrLf & Mid(strImg, 2)
            End If
        End If
    End If
    
    
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
    Dim str_Return As String, str_in As String
    Dim tmpStr As String, tmpStr2 As String
     '----字符转换成二进制流
    tmpStr = str_bin
    tmpStr2 = ""
    Do While tmpStr <> ""
        If Len(Hex(Asc(Mid(tmpStr, 1, 1)))) = 4 Then
            tmpStr2 = tmpStr2 & "," & Mid(Hex(Asc(Mid(tmpStr, 1, 1))), 1, 2) & "," & Mid(Hex(Asc(Mid(tmpStr, 1, 1))), 1, 2)
        Else
            tmpStr2 = tmpStr2 & "," & Hex(Asc(Mid(tmpStr, 1, 1)))
        End If
        
        tmpStr = Mid(tmpStr, 2)
    Loop
    
    
    lng_len = Len(tmpStr2)

    If lng_len < 3 Then
        Graph_Text = ""
        Exit Function
    Else
        str_in = tmpStr2
    End If

    For lng_i = 1 To lng_len / 3
        
        str_bit = Replace(Left(str_in, 3), ",", "")
        str_in = Mid(str_in, 4)
        
        If IsNull(str_bit) Then
            str_Return = str_Return & "000"
        Else
            str_Return = str_Return & ";" & Format((CByte("&H" & str_bit) - 32), "000")
        End If
    Next
    Graph_Text = str_Return
End Function


