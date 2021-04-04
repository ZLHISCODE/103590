VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_CS600A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:生化CS600A
'厂商:长春迪瑞
'用户:德保妇幼
'作者:hhd
'使用方法:虚拟串口接收
'时间：2013/12/19
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
    '标准函数：从文件获取检验结果
    '   strFile：包含路径的文件名
    '   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
    '   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    '   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    '   返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果

    Dim lng_S As Long    '开始
    Dim lng_E As Long    '结束
    Dim lng_SS As Long, lng_ss1 As Long
    '一个完整的记录
    Dim str_标本号 As String, str_标本号1 As String, str_日期 As String, str_检验人标本质控 As String, str_检验项目 As String, str_检验结果 As String
    Dim str_检验结果1 As String
    Dim str_Line As String
    Dim lng_标本号 As String
    Dim var_data As Variant
    Dim i As Long
    Dim str_Code As String, str_结果 As String
    str_检验人标本质控 = " |血清|0"
    
    On Error GoTo errH
    
    '-----应答
    If InStr(strSource, Chr(&H5)) > 0 Then
        strCmd = Chr(&H6&)
    End If
    If InStr(strSource, Chr(&HD) & Chr(&HA)) > 0 Then
        strCmd = Chr(&H6&)
    End If

    If strSource <> "" Then mstrCode = mstrCode & strSource

    '---取完整的数据
    lng_S = InStr(mstrCode, Chr(&H5))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H4))
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

            '获取日期时间
            If InStr(str_Line, "H|\") > 0 Then    'H|\
                str_日期 = Trim(Split(str_Line, "|")(12))
                str_日期 = Mid(str_日期, 1, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid( _
                         str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, _
                                                                             2) & ":" & Mid(str_日期, 13, 2)    '拆分日期出来“yyyy-mm-dd”
                str_日期 = Format(str_日期, "yyyy-MM-dd hh:mm:ss")

                If Not IsDate(str_日期) Then    '如果日期为空，则显示为当前日期
                    str_日期 = Format(Now, "yyyy-MM-dd hh:mm:ss")
                End If

                '取标本号
            ElseIf InStr(str_Line, "O|") = 4 Then
                str_标本号1 = Trim(Split(str_Line, "|")(2))
                str_标本号 = Trim(Val(Split(str_标本号1, "^")(1)))

                '取分组的检验项目和检验结果
            ElseIf InStr(str_Line, "R|") = 4 Then
                str_检验项目 = Replace(Replace(Trim(Split(str_Line, "|")(2)), "^^^", ""), "红细胞", "rbc")
                str_检验结果1 = Trim(Split(str_Line, "|")(3))
                str_检验结果 = IIf(str_检验结果1 = "", "*.**", str_检验结果1)

                If str_标本号 > 0 And str_检验项目 <> "" Then
                    str_结果 = IIf(Trim(str_结果) = "", "", str_结果 & "||") & str_日期 & "|" & str_标本号 & "|" & str_检验人标本质控 & "|" & str_检验项目 & "|" & str_检验结果
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
    If str_结果 <> "" Then strResult = Mid(str_结果, 3)
    Exit Sub
errH:
    WriteLog "CS600A", "CS", Err.Number & ":" & Err.Description
End Sub


