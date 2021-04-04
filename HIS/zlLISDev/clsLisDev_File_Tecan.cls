VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_File_Tecan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:酶标仪
'厂商:帝肯
'作者:陈东
'用户:商丘人民医院

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
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim strDate As String, strItem As String, strNO As String, strValue As String, strOD As String, strCoutOff As String, strSCO As String
    Dim strFilePath As String, strFileName As String
    Dim arrLine As Variant
    Dim lngLoop As Long, intB As Integer, intK As Integer
    Dim lngNo As Long, intCount As Integer
    
    On Error GoTo hErr
    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    lngLoop = DateDiff("d", dtStart, dtEnd)
    ReDim aResults(0)

    For intCount = 0 To lngLoop
        strDate = Format(DateAdd("d", intCount, dtStart), "yyyy-MM-dd")
        strFileName = Dir(strFilePath & "\*_" & strDate & "_*.asc")
        
        Do While strFileName <> ""
            
            If objFile.FileExists(strFilePath & "\" & strFileName) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
                lngNo = 0
                Do While Not objStream.AtEndOfStream
                    strLine = UCase(objStream.ReadLine)
                    arrLine = Split(strLine, Chr(9))
                    lngNo = lngNo + 1
                    If UBound(arrLine) >= 5 Then
                        strItem = arrLine(0)
                        strNO = arrLine(1)
                        If strNO Like "??-??-???" Then
                            
                            intB = Val(Split(strNO, "-")(0))
                            intK = Val(Split(strNO, "-")(1))
                            
                            strNO = intK + (intB - 1) * 16
                            
                            strValue = Trim("" & arrLine(2))
                            If strValue = "" Then strValue = "-"
                            strCoutOff = arrLine(3)
                            strOD = arrLine(4)
                            
                            If Val(strCoutOff) <> 0 Then
                                strSCO = Format(Val(strOD) / Val(strCoutOff), "0.000")
                            Else
                                strSCO = 0
                            End If
                            
                            If strValue <> "" And strDate <> "" And strNO <> "" Then
                                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                                If IsNumeric(strSCO) And IsNumeric(strCoutOff) Then
                                    aResults(UBound(aResults)) = strDate & "|" & strNO & "| |血清|0|" & strItem & "|" & strValue & "^" & strOD & "^" & strCoutOff & "^" & strSCO
                                Else
                                    aResults(UBound(aResults)) = strDate & "|" & strNO & "| |血清|0|" & strItem & "|" & strValue
                                End If
                            End If
                        End If
                    End If
                Loop
                
                objStream.Close
                Set objStream = Nothing
                
                '读了就把文件放到BAK下
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFileName, strFilePath & "\Bak\" & strFileName, True
                objFile.DeleteFile strFilePath & "\" & strFileName, True
                
            End If
            
            strFileName = Dir
        Loop
    Next
    ResultFromFile = aResults
    Exit Function
hErr:
    WriteLog "Tecan", Err.Number, Err.Description
End Function

Public Function SendSample(strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    
    'iSendStep：当前发送操作的步骤。如果处理完后发送结束，则返回0，否则自动＋1返回。
    'strResponse：当前仪器的响应内容
    'strSampleInfo：发送的标本信息。
    '   标本之间以||分隔
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否急诊
    '   第5个无素：条码
    '   第6～9元素：系统保留
    '   从第10个元素开始为需要的检验项目。


    SendSample = ""
   
        
End Function

Private Function WritTxt(ByVal strFileName As String, strInfo As String)
    Dim objFileSystem As New FileSystemObject
    Dim objStream As TextStream
    
    If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
    Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strInfo)
    objStream.Close
    Set objStream = Nothing
End Function

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


