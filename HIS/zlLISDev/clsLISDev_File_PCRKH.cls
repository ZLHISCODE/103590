VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_PCRKH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:PCR
'厂商:上海科华
'作者:陈东
'用户:3201医院
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
    Dim iRow As Integer '行号
    Dim varline As Variant
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    iRow = 0
    Do Until objStream.AtEndOfStream
        str检验结果 = "": str操作员 = "": str日期 = "": str标本号 = ""
        
        strLine = objStream.ReadLine
        '标本号，数据行
        
        If UBound(Split(strLine, Chr(&H1E))) = 3 Then
            strChenel = Trim(Split(strLine, Chr(&H1E))(2))
            If strChenel <> "" Then
                str操作员 = Trim(Split(strLine, Chr(&H1E))(0))
                
                str检验结果 = Trim(Split(strLine, Chr(&H1E))(1))
                
                If UBound(Split(str检验结果, " ")) > 5 Then
                    str日期 = Split(str检验结果, " ")(2) & "-" & Split(str检验结果, " ")(3) & "-" & Split(str检验结果, " ")(4)
                End If
                
                str检验结果 = Split(str检验结果, " ")(5) * 100
                If Val(str检验结果) = 0 Then
                    str检验结果 = "0.000E+0"
                Else
                    str检验结果 = Split(str检验结果, ".")(0)
                    str检验结果 = Format(Val(str检验结果) / 10 ^ (Len(str检验结果) - 1), "0.000") & "E+" & Format(Len(str检验结果) - 3, "00")
                End If
                str标本号 = ""
                Do While str操作员 <> ""
                    If IsNumeric(Right(str操作员, 1)) Then
                        str标本号 = Right(str操作员, 1) & str标本号
                        str操作员 = Left(str操作员, Len(str操作员) - 1)
                    Else
                        str操作员 = ""
                        Exit Do
                    End If
                Loop
                If Not IsDate(str日期) Then str日期 = Format(Now, "yyyy-MM-dd")
                
                If Val(str标本号) > 0 And str检验结果 <> "" And strChenel <> "" Then
                    If CDate(str日期) >= dtStart And CDate(str日期) <= dtEnd Then
                        If lngSampleNO = -1 Or lngSampleNO = Val(str标本号) Then
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = Format(CDate(str日期), "yyyy-MM-dd") & "|" & str标本号 & "| |血液|0|" & strChenel & "|" & str检验结果
                        End If
                    End If
                End If
            End If
        
        End If
    Loop
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_Anytest.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function
