VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LightCycleII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:荧光定量PCR分析仪 LightCycleII
'厂商:罗氏 Roche
'作者:陈东
'用户:德阳医院
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
    Dim str标本号 As String, str检验结果 As String, blnReturn As Boolean, i As Integer
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine Like "*Run By:*" Then
            '操作员
            str操作员 = Trim(Mid(strLine, InStr(strLine, "Run By:") + 7))
            
        ElseIf strLine Like "Run Date:*,*" Then
            '时间
            str日期 = Replace(Split(strLine, vbTab)(0), "Run Date: ", "")
            str日期 = Replace(str日期, ", ", " ")
            varTmp = Split(str日期, " ")
            str日期 = varTmp(2) & "-" & TrunMonth(varTmp(0)) & "-" & varTmp(1)
            If IsDate(str日期) = False Then
                str日期 = Format(Now, "yyyy-MM-dd")
            End If
        ElseIf InStr(strLine, vbTab) > 0 Then
            varItem = Split(strLine, vbTab)
            If UBound(varItem) = 5 Then
                str标本号 = varItem(1)
                str检验结果 = ""
                If Not (InStr(str标本号, "-") > 0 Or _
                   InStr(str标本号, "+") > 0 Or _
                   InStr(str标本号, "/") > 0 Or _
                   InStr(str标本号, "ST") > 0 Or _
                   InStr(str标本号, "CONTROL") > 0) Then
                   
                    If str标本号 Like "HB*" And Trim(varItem(3)) <> "" Then
                        str标本号 = Mid(str标本号, 3)
                        str检验结果 = "HB|" & Trim(varItem(3))
                        
                        If Len(str标本号) >= 10 Then
                            str日期 = Format(CDate(Mid(str标本号, 2, 2) & "-" & Mid(str标本号, 4, 2) & "-" & Mid(str标本号, 6, 2)), "yyyy-MM-dd")
                        End If
                        
                    ElseIf InStr(str标本号, " ") > 0 And Trim(varItem(3)) <> "" Then
                        varTmp = Split(str标本号, " ")
                        If IsNumeric(Mid(varTmp(1), 1, 1)) Then
                            str标本号 = varTmp(1)
                        Else
                            str标本号 = Mid(varTmp(1), 2)
                            
                            For i = 1 To Len(str标本号)
                                If IsNumeric(Left(str标本号, 1)) Then
                                    Exit For
                                Else
                                    str标本号 = Mid(str标本号, i + 1)
                                End If
                            Next
                        End If
                        str检验结果 = varTmp(0) & "|" & Trim(varItem(3))
                        
                        If Len(str标本号) >= 10 Then
                            str日期 = Format(CDate(Mid(str标本号, 2, 2) & "-" & Mid(str标本号, 4, 2) & "-" & Mid(str标本号, 6, 2)), "yyyy-MM-dd")
                        End If
                    End If
                    
                    If str检验结果 <> "" Then
                        blnReturn = False
                        If lngSampleNO = -1 Then
                            '取指定日期记录
                            If CDate(str日期) >= dtStart And CDate(str日期) <= dtEnd Then
                                blnReturn = True
                            End If
                        Else
                            If lngSampleNO = str标本号 Then
                                '指定了标本号
                                blnReturn = True
                            End If
                        End If
                        
                        If blnReturn Then
                            If str标本号 <> "" Then
                                If IsDate(str日期) = False Then str日期 = Format(Now(), "yyyy-MM-dd")
                                If str操作员 = "" Then str操作员 = " "
                                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "|" & str操作员 & "|血清|0|" & str检验结果
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_LightCycleII.ResultFromFile", strFile, Err.Number & " " & Err.Description
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
