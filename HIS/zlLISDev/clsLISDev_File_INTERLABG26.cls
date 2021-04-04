VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_INTERLABG26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: 电泳仪
'厂商:
'作者: 敬秦
'用户: 大安妇幼保健院 读取瑞美的数据

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
    Dim cnAccess As New ADODB.Connection, rsTmp As New ADODB.Recordset, strSQL As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject, strFilename As String, txtStream As textStream
    Dim strLine As String, arrItems() As String
    Dim strID As String, strDate As String, strValue As String, strCode As String, strItem As String
    Dim intDay As Integer, intLop As Integer, i As Integer
    Dim strIniFile As String, Mode As Integer
    Dim str标本号 As String
    
    On Error GoTo hErr
    ReDim aResults(0)
    aResults(0) = ""
    strIniFile = App.Path & "\INTERLABG26.ini"
    If Not objFile.FileExists(strIniFile) Then
       Set txtStream = objFile.CreateTextFile(strIniFile)
       txtStream.WriteLine "[Options]"
       txtStream.WriteLine "TYPE=1      ;   1:读取瑞美数据,2:读取MDB数据库"
       txtStream.Close
       Set txtStream = Nothing
    End If
    Mode = Val(ReadIni("Options", "TYPE", strIniFile, 0))
    If strFile <> "" Then strFilePath = objFile.GetParentFolderName(strFile)
    intDay = Abs(DateDiff("d", dtStart, dtEnd))
    If Mode = 1 Then
        For intLop = 0 To intDay
            strFilename = Format(dtStart + intLop, "yymmdd") & ".txt"
            If Not objFile.FileExists(strFilePath & "\" & strFilename) Then GoTo nextLine:
            Set txtStream = objFile.OpenTextFile(strFilePath & "\" & strFilename)
            Do Until txtStream.AtEndOfStream
                strLine = txtStream.ReadLine
                If InStr(strLine, Format(dtStart + intLop, "yyyy-mm-dd")) > 0 Then
                    arrItems = Split(strLine, Chr(&H9))
                    strDate = Format(arrItems(0) & " " & Time, "yyyy-mm-dd hh:mm:ss")
                    strID = arrItems(1)
                    strCode = Split(arrItems(2), ":")(1)
                    If strCode = "" Then GoTo nextLine1:
                    strValue = Replace(Split(Split(arrItems(3), ":")(1), " ")(0), "未注册", "")
                    strItem = strItem & "|" & strCode & "|" & strValue
                End If
                If InStr(strLine, "计算当前标本") > 0 Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strDate & "|" & strID & "| |血清|0" & strItem
                    strItem = ""
                End If
nextLine1:
            Loop
    '        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
    '        aResults(UBound(aResults)) = strDate & "|" & strID & "| |血清|0" & strItem
    '        strItem = ""
nextLine:
        Next
    Else
        If strFile <> "" Then strFilePath = objFile.GetParentFolderName(strFile)
        With cnAccess
            .CursorLocation = adUseClient
            .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
            .Open
            If .State = 0 Then
                Call WriteLog("调试日志", "", "未能打开MDB数据库连接！")
                Exit Function
            End If
        End With
        strSQL = "select idpaz,[data esame] as date1,fraz1 as HbA,fraz2 as HbA2,fraz3 as HbF,fraz4 as HbS from archivio where [Data esame] " & _
            "between format('" & Format(dtStart, "yyyy-mm-dd") & "','yyyy-mm-dd') and  format('" & Format(dtEnd, "yyyy-mm-dd") & "','yyyy-mm-dd')"
        Set rsTmp = cnAccess.Execute(strSQL)
        Do Until rsTmp.EOF
            str标本号 = rsTmp!idpaz
            strDate = Format(rsTmp!date1, "yyyy-mm-dd hh:mm:ss")
            strItem = "HbA|" & rsTmp!HbA & "|HbA2|" & rsTmp!HbA2 & "|HbF|" & rsTmp!HbF & "|HbS|" & rsTmp!HbS
            If str标本号 <> "" And strItem <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = strDate & "|" & str标本号 & "| |血清|0|" & strItem
            End If
            rsTmp.MoveNext
        Loop
    End If
    For i = 0 To UBound(aResults)
        Debug.Print aResults(i)
    Next
    ResultFromFile = aResults
    Exit Function
hErr:
    Call WriteLog("调试日志", "", Err.Description)
    Err.Clear
End Function



