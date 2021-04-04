VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DL_96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:DL_96
'厂商:迪尔
'作者:曾超
'用户:XXXXXXXXXXX
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
    
    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim str标本号 As String, str项目 As String
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aLine() As String
    Dim strItem As String
    
    
    On Error GoTo errH
    
    Dim rsErs As ADODB.Recordset
    Dim str日期 As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .CursorLocation = adUseClient
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function
        
        strSQL = "select * " & _
                 " from 细菌报告 "

        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        
        Do While Not rsTmp.EOF
                
            str日期 = Format("" & Now, "yyyy-MM-dd HH:mm:ss")
            str标本号 = Val("" & rsTmp("标本编号"))
            
            For intLoop = 1 To 5
                If rsTmp("细菌" & intLoop) <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |" & _
                        Mid(rsTmp("细菌" & intLoop), 1, InStr(rsTmp("细菌" & intLoop), "(") - 1) & "|0"
                    If rsTmp("药敏" & intLoop) <> "" Then
                        aLine = Split(rsTmp("药敏" & intLoop), ";")
                        strItem = ""
                        For intItem = 0 To UBound(aLine) Step 6
                            If intItem + 6 <= UBound(aLine) Then
                                strItem = strItem & "|" & aLine(intItem + 1) & "|"
                                Select Case aLine(intItem + 4)
                                    Case "耐药"
                                        strItem = strItem & "R"
                                    Case "中介"
                                        strItem = strItem & "I"
                                    Case "敏感"
                                        strItem = strItem & "S"
                                End Select
                                strItem = strItem & "^1^" & aLine(intItem + 3)
                            End If
                        Next
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & strItem
                    Else
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|1^R^"
                    End If
                End If
            Next
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    For intLoop = 0 To UBound(aResults)
        Debug.Print aResults(intLoop)
    Next
    
    Exit Function
errH:
    
    WriteLog "File_DL_96", Err.Description, ""
End Function



'0769-26208501

