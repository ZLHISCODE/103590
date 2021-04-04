VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_ACL_TOP700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:贝克曼ACL700TOP血凝仪
'用户:广西田东县人民医院
'作者:hhd
'单向
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

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, strCmd As String)

    Dim lng_S As Long    '开始
    Dim lng_E As Long    '结束
    Dim lng_SS As Long, lng_ss1 As Long
    '一个完整的记录
    Dim str_标本号 As String, str_日期 As String, str_检验人标本质控 As String, str_检验项目 As String, str_检验结果 As String
    Dim str_Line As String, strTmp As String
    Dim lng_标本号 As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_结果 As String
    str_检验人标本质控 = " |血液|0"
    On Error GoTo errH
    '-----应答
    strSource = Replace(strSource, Chr(&H4) & Chr(&H5), Chr(&H5))
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6)
        '    ElseIf InStr(strSource, Chr(&H4)) > 0 Then
        '        strCmd = Chr(&H6)
    ElseIf InStr(strSource, Chr(&HA)) > 0 Then
        strCmd = Chr(&H6)
    End If
    If strSource <> "" Then strCmd = Chr(&H6)

    mstrCode = mstrCode & strSource
    '---取完整的数据
    lng_S = InStr(mstrCode, Chr(&H5))   '以&H5开头，返回完整段落的开始位置
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '以&H4结束，返回完整段落的结束位置
    Else
        lng_E = 0
    End If


    Do While lng_S > 0 And lng_E > 0    '如果开始、结束都不为空
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)    '返回从开始位置到结束位置之间的字符
        mstrCode = Mid(mstrCode, lng_E + 1)
        Do While InStr(str_Code, Chr(&H17)) > 0
            lng_SS = InStr(str_Code, Chr(&H17))
            str_Code = Mid(str_Code, 1, lng_SS - 1) & Mid(str_Code, lng_SS + 7)    '取7位校验码
        Loop

        var_data = Split(str_Code, Chr(&HD))    '取每段数据
        For i = LBound(var_data) To UBound(var_data)
            str_Line = var_data(i)

            If InStr(str_Line, "H|") > 0 Then
                str_日期 = Trim(Split(str_Line, "|")(13))    '获取日期时间
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid( _
                         str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, _
                                                                             2) & ":" & Mid(str_日期, 13, 2)    '拆分日期出来“yyyy-mm-dd”
                str_日期 = Format(str_日期, "yyyy-MM-dd hh:mm:ss")
                If Not IsDate(str_日期) Then    '如果日期为空，则显示为当前日期
                    str_日期 = Format(Now, "yyyy-MM-dd hh:mm:ss")
                End If
            ElseIf InStr(str_Line, "O|") = 4 Then
                str_标本号 = Trim(Split(str_Line, "|")(2))    '从“2P|1|||||” 的|后从1开始后开始取
            ElseIf InStr(str_Line, "R|") = 4 Then    '取分组的检验项目

                strTmp = Split(str_Line, "|")(4)
                strTmp = Replace$(strTmp, "%", "1")
                strTmp = Replace$(strTmp, "/", "2")
                strTmp = Replace$(strTmp, "", "3")  ' 处理空格
                str_检验项目 = Trim(Split(str_Line, "|")(2)) & strTmp
                str_检验项目 = Replace$(str_检验项目, "^^^", "")
                'n = n + 1
                str_检验结果 = Trim(Split(str_Line, "|")(3))

                '最后结果拼接起来
                If Val(str_标本号) > 0 And str_检验项目 <> "" And str_检验结果 <> "" Then
                    str_结果 = str_结果 & "||" & str_日期 & "|" & str_标本号 & "|" & str_检验人标本质控 & "|" & str_检验项目 & "|" & str_检验结果
                End If

            End If

        Next

        lng_S = InStr(mstrCode, Chr(&H5))   '以&H5开头，返回完整段落的开始位置
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H4))  '以&H4结束，返回完整段落的结束位置
        Else
            lng_E = 0
        End If
    Loop

    strResult = Mid(str_结果, 3)
    Exit Sub
errH:
    WriteLog "ALC_TOP700", strSource, Err.Description
End Sub

Public Function SendSample(ByVal strSample As String) As String
    SendSample = ""
End Function





