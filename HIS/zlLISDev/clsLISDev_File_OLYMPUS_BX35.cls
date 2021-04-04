VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_OLYMPUS_BX35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'仪器:奥林巴斯(OLYMPUS) BX35 显微镜
'厂商:奥林巴斯(OLYMPUS)
'作者:敬秦
'用户:邛崃市医疗中心医院
'日期:2014-12-22 09:53:13
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
    Dim strCode As String, strValue() As String, strDate As String, strIteam As String, strSample As String, strLine As String, strData As String, strChnnl As String
    Dim lngS As Long, lngE As Long, i As Integer
    Dim OBJFILE As New FileSystemObject
    Dim STRFILENAME As String, strPath As String, objStream As textStream, strFilePath As String
    Dim rsResults() As String, strImage As String
    Dim pathstr As String
On Error GoTo hErr

    If Not OBJFILE.FileExists(App.Path & "\" & "OLYMPUS_BX35.ini") Then '获取数据库信息
        Set objStream = OBJFILE.CreateTextFile(App.Path & "\" & "OLYMPUS_BX35.ini")
            objStream.WriteLine "[OLYMPUS_BX35]"
            objStream.WriteLine "filePath=" & App.Path
            objStream.Close
            Set objStream = Nothing
    End If
    strPath = ReadIni("OLYMPUS_BX35", "filePath", App.Path & "\" & "OLYMPUS_BX35.ini")
    If strPath = "" Or Right$(strPath, 1) = "\" Then
        MsgBox "配置文件设置错误或者最后含有【\】符号！", vbInformation, "接口"
    End If
    ReDim rsResults(0)
    rsResults(0) = ""
    STRFILENAME = Dir(strPath & "\", vbDirectory)       '取第一个
    Do While STRFILENAME <> ""                    '不为空就循环
       If STRFILENAME <> "." And STRFILENAME <> ".." Then   '如果不是 当前目录及上一级目录 . 及 ..
             If (GetAttr(strPath & "\" & STRFILENAME) And vbDirectory) = vbDirectory Then        '是否是目录
'                Debug.Print strFileName          '确定是目录,添加到列表,这里可以自己处理
                strDate = Split(STRFILENAME, "_")(0)
                strDate = Format(Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Right(strDate, 2) & " " & Time, "yyyy-mm-dd HH:mm:ss")
                strSample = Split(STRFILENAME, "_")(1)
                If OBJFILE.FileExists(strPath & "\" & STRFILENAME & "\report.txt") Then '查找是否存在此文件
                    Set objStream = OBJFILE.OpenTextFile(strPath & "\" & STRFILENAME & "\report.txt")
                    Do Until objStream.AtEndOfStream
                        strLine = objStream.ReadLine
                        If InStr(strLine, "<pcode") > 0 Then
                            strChnnl = Split(Split(strLine, ";")(2), "=")(1)
                            strData = Split(Split(strLine, ";")(3), "=")(1)
                            strIteam = strIteam & "|" & strChnnl & "|" & strData
                        End If
                        If InStr(strLine, "<image") > 0 Then
                            strLine = Replace$(strLine, "<", "")
                            strLine = Replace$(strLine, ">", "")
                            strImage = strImage & "^  ;113;0;" & Split(Split(strLine, " ")(1), "=")(1)
                            strImage = Replace(strImage, vbCr, "")
                        End If
                    Loop
                    If strIteam <> "" And strImage <> "" And strSample <> "" Then
                        If rsResults(UBound(rsResults)) <> "" Then ReDim Preserve rsResults(UBound(rsResults) + 1)
                        rsResults(UBound(rsResults)) = strDate & "|" & strSample & "| |血清|0" & strIteam
                        strIteam = ""
                        If strImage <> "" Then '加上图片
                            rsResults(UBound(rsResults)) = rsResults(UBound(rsResults)) & vbCrLf & strImage
                            strImage = ""
                        End If
                    End If
                End If
            End If
         End If
         STRFILENAME = Dir      '取下一个
    Loop
    ResultFromFile = rsResults
    Exit Function
hErr:
    MsgBox "zlCDLISDev|clsLISDev_File_OLYMPUS_BX35|" & Err.Description, vbInformation, "接口错误"
    Err.Clear
End Function
