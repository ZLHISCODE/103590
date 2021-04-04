VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:微生物
'厂商:上海德记行科技发展有限公司 联系人:(沈刚 13795478025)
'作者:曾超
'用户:重庆市二院
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
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
    
    Dim aResults() As String, objFile As New Scripting.FileSystemObject
    Dim strChenel As String, str操作员 As String, str日期 As String
    Dim strFilePath As String
    
    Dim objStream As TextStream
    Dim strLine As String, varItem As Variant, varTmp As Variant
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean
    Dim varStr As Variant
    Dim lngloop As Long
    Dim strTmp As String
    Dim strData As String           '解析数据
    Dim intCount As Integer         '计数
    
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    Open strFile For Input As #1 ' 打开文件。
    Do While Not EOF(1) ' 循环至文件尾。
        Line Input #1, strLine ' 读入一行数据并将其赋予某变量。
        varStr = Split(strLine, Chr(9))
        
        '在日期范围内就退出
        If CDate(Trim(varStr(39))) < dtStart Or CDate(Trim(varStr(39))) > dtEnd Then
            Close #1
            Exit Function
        End If
        
        '替换A=1 B=2 C=3 D=4
        strTmp = Replace(Replace(Replace(Replace(varStr(7), "A", 1), "B", 2), "C", 3), "D", 4)
        strData = Trim(varStr(39)) & "|" & Trim(varStr(6) & strTmp) & "^0^| |" & varStr(9) & "|0"
        
        For lngloop = 40 To UBound(varStr)
            If Trim(varStr(lngloop)) <> "" Then
                intCount = intCount + 1
                strData = strData & "|" & Trim(varStr(lngloop))
                If intCount = 2 Then
                    strData = strData & "^^"
                    intCount = 0
                End If
            End If
        Next
        
        strData = Replace(strData, "RESIST", "R")
        strData = Replace(strData, "SUSC", "S")
        strData = Replace(strData, "INTER", "I")
        '其他类型的，暂时不处理
        strData = Replace(strData, "NOINTP", "")
        strData = Replace(strData, "NOMIC", "")
        aResults(UBound(aResults)) = strData
        ReDim Preserve aResults(UBound(aResults) + 1)
    Loop
    Close #1 ' 关闭文件。

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycleII.ResultFromFile", strFile, Err.Number & " " & Err.Description
    Close #1
End Function


Private Function TrunMonth(ByVal strMonth As String) As String
    '转换中文的月份为数字
    Select Case UCase(strMonth)
        Case "一月": TrunMonth = "01"
        Case "二月": TrunMonth = "02"
        Case "三月": TrunMonth = "03"
        Case "四月": TrunMonth = "04"
        Case "五月": TrunMonth = "05"
        Case "六月": TrunMonth = "06"
        Case "七月": TrunMonth = "07"
        Case "八月": TrunMonth = "08"
        Case "九月": TrunMonth = "09"
        Case "十月": TrunMonth = "10"
        Case "十一": TrunMonth = "11"
        Case "十二月": TrunMonth = "12"
        
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select
End Function


