VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Coda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CODA全自动酶免分析系统
'厂商:美国伯乐
'作者:陈东
'用户:泽州县医院
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strChenel As String, strDate As String   '通道,操作员,日期
    Dim strFilePath As String   '文件路径
    Dim strLine As String
    Dim strSampleID As String, strValue As String   '标本号，检验结果
    Dim lngLoop As Long
    Dim strData As String           '一个数据
    Dim intCount As Integer         '计数
     
    
    On Error GoTo errHandle
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    '
    lngLoop = DateDiff("d", dtStart, dtEnd)
    ReDim Preserve aResults(0)
    
    For intCount = 0 To lngLoop
        strDate = Format(DateAdd("d", intCount, dtStart), "ddMMyy")
        strFile = Dir(strFilePath & "\*" & strDate & "*.ast")
        Do While strFile <> ""
            
            Open strFilePath & "\" & strFile For Input As #1 ' 打开文件。
            Do While Not EOF(1) ' 循环至文件尾。
                Line Input #1, strLine ' 读入一行数据并将其赋予某变量。
                If strLine Like "P|*|*|*|*" Then
                    strSampleID = Trim(Split(strLine, "|")(3))
                ElseIf strLine Like "O|*|*|*|*" Then
                    strDate = Split(strLine, "|")(22)
                    If strDate Like "##############" Then
                        strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                    End If
                ElseIf strLine Like "R|*|*|*" Then
                    strChenel = Trim(Split(strLine, "|")(2))
                    strValue = Trim(Split(strLine, "|")(6))
                    
                    If IsDate(strDate) And strSampleID <> "" And strChenel <> "" And strValue <> "" Then
                        If strSampleID Like "##########" Then
                            strSampleID = Mid(strSampleID, 7)
                            strData = strDate & "|" & strSampleID & "| |Bold|0|"
                        Else
                            strData = strDate & "|" & strSampleID & "| |Bold|1|"
                        End If
                        If strChenel <> "" And strValue <> "" Then
                            strData = strData & strChenel & "|" & strValue
                            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                            aResults(UBound(aResults)) = strData
                        End If
                        strDate = ""
                        strSampleID = ""
                        
                    End If
                End If
            Loop
            Close #1 ' 关闭文件。
            
            '备份已读取的文件，然后删除
            If Not objFile.FolderExists(strFilePath & "\Bak") Then objFile.CreateFolder strFilePath & "\Bak"
            objFile.CopyFile strFilePath & "\" & strFile, strFilePath & "\Bak\" & strFile
            objFile.DeleteFile strFilePath & "\" & strFile
            
            strFile = Dir
        Loop
    Next
    '没有找到文件时退出
    
    
    ResultFromFile = aResults
    Exit Function
errHandle:
    'Resume
    WriteLog "CODA", strFile, Err.Number & " " & Err.Description
    Close #1
End Function


