VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EH_2060B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:EH_2060B 尿沉渣
'厂商:惠生
'作者:陈东
'用户:重庆市第二人民医院
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
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As Scripting.TextStream, strLine As String

    Dim str日期 As String, str检验项目 As String, str标本号 As String, strImg As String
    Dim strFilePath As String, strFilename As String
    Dim lngS As Long, lngE As Long
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    strFilename = UCase(Dir(strFilePath & "\*.txt"))
    ReDim aResults(0)
    
    Do While strFilename <> ""
        str日期 = "": str标本号 = "": str检验项目 = "": strImg = ""
        If UCase(strFilename) Like "###########.TXT" Then
            If objFile.FileExists(strFilePath & "\" & strFilename) Then
                Set objStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine
                    
                    If strLine Like "[[]P07[]]*:*" Then
                        str标本号 = Split(strLine, ":")(1)
                    ElseIf strLine Like "[[]P08[]]*:*" Then
                        str日期 = Split(strLine, ":")(1)
                    ElseIf strLine Like "[[]P[1-4]#[]]*:*" Then
                        If Split(strLine, ":")(1) <> "" Then
                            lngS = InStr(strLine, "[") + 1
                            lngE = InStr(strLine, "]")
                            
                            str检验项目 = str检验项目 & "|" & Mid(strLine, lngS, lngE - lngS) & "|" & Split(strLine, ":")(1)
                        End If
                    ElseIf strLine Like "[[]P61[]]*:*" And UBound(Split(strLine, ":")) > 1 Then
                        strImg = strImg & "^图1;110;0;" & Split(strLine, ":")(1) & ":" & Split(strLine, ":")(2)
                    ElseIf strLine Like "[[]P62[]]*:*" And UBound(Split(strLine, ":")) > 1 Then
                        strImg = strImg & "^图2;110;0;" & Split(strLine, ":")(1) & ":" & Split(strLine, ":")(2)
                    End If
                Loop
                
                If Val(str标本号) > 0 And IsDate(str日期) And str检验项目 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |血清|0" & str检验项目
                    If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImg, 2)
                End If
                '读了就把文件放到BAK下
                objStream.Close
                Set objStream = Nothing
                If Not objFile.FolderExists(strFilePath & "\Bak") Then Call objFile.CreateFolder(strFilePath & "\Bak")
                objFile.CopyFile strFilePath & "\" & strFilename, strFilePath & "\Bak\" & strFilename, True
                objFile.DeleteFile strFilePath & "\" & strFilename, True
            End If
        End If
        strFilename = Dir
    Loop
    ResultFromFile = aResults
End Function




