VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_EStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:全自动酶免仪
'厂商:瑞士哈美顿(HAMILTON)
'作者:陈东
'用户:南京江北人民医院
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
    Dim lngResultCount As Long
    Dim str日期 As String, str检验项目 As String, str定性结果 As String, str定量结果 As String
    Dim blnOk As Boolean '是否按标本号提取数据
    Dim strFilePath As String
    Dim strCO As String 'cut-off值从config.ini取值

    
    strFilePath = UCase(objFile.GetParentFolderName(strFile))
    
    If Dir(strFile) = "" Then Exit Function
    strFile = UCase(strFile)
    
    str检验项目 = Replace(Replace(strFile, strFilePath, ""), ".TXT", "")
    If Left(str检验项目, 1) = "\" Then str检验项目 = Mid(str检验项目, 2)
    '从config读取cut-off值

    Call AnalyConf(UCase(str检验项目), strCO)
    
    Set objStream = objFile.OpenTextFile(strFile, ForReading)
    lngResultCount = -1
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, "检测时间:") > 0 Then
            str日期 = Mid(strLine, InStr(strLine, ":") + 1, 18)
            If IsDate(str日期) Then
                str日期 = Format(str日期, "yyyy-MM-dd HH:mm:ss")
            Else
                str日期 = Format(Now, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        If Val(Mid(strLine, 3, 3)) > 0 Then
            blnOk = False
            If lngSampleNO <> -1 Then
                If lngSampleNO = Val(Mid(strLine, 3, 3)) Then
                    blnOk = True
                End If
            Else
                blnOk = True
            End If
            
            If CDate(str日期) >= dtStart And CDate(str日期) <= dtEnd Then
                blnOk = True
            End If
            If blnOk Then
            
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                str定性结果 = Trim(Mid(strLine, 20, 2))
                str定性结果 = Switch(str定性结果 = "+", "阳性(+)", str定性结果 = "-", "阴性", str定性结果 = "+-", "弱阳性(+-)")
                
                If InStr(Mid(strLine, 50), ".") <= 0 Then
                    str定量结果 = Val(Mid(strLine, 30))
                Else
                    str定量结果 = Val(Mid(strLine, 50))
                End If
                If Left(str定量结果, 1) = "." Then str定量结果 = "0" & str定量结果
                aResults(lngResultCount) = str日期 & "|" & Val(Mid(strLine, 3, 3)) & "| |血液|0|" & str检验项目 & "|" & str定性结果 & "^" & str定量结果 & "^" & strCO & "^0"
            End If
            
        End If
    Loop
    
    ResultFromFile = aResults
End Function

Public Sub AnalyConf(strItem As String, strCO As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream
    On Error GoTo ErrHand
    Dim intloop As Long
    
    
    '从配置文件中读取保存在文件中的配置cotf
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Config.ini"
    If objFilesys.FileExists(strFile) Then
        Set objStream = objFilesys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            If arrLine(0) = strItem Then
               strCO = arrLine(1)
               Exit Sub
            End If
           
        Loop
        objStream.Close
        blnOpen = False
    End If
    
   
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub
