VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_XD511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:XD511全自动尿沉渣分析仪
'型号:XD511
'厂商:上海北昂医疗技术有限公司
'作者:李建飞
'用户:上海曲阳医院罗宏伟
'日期:2010-05-12
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
    
    Dim aResults() As String
    Dim cnSQL2000 As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim str参数 As String, arr参数() As String
    Dim str检验结果 As String, str标本号 As String
    Dim str标本类型 As String, str检验日期 As String
    Dim lngResultCount As Long
    Dim iField As Integer
    Dim aPic() As String
    Dim strPic As String
    Dim strPicPath As String               '图像访问路径
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '配置文件用的是仪器商的配置文件，在技师站的参数选择中指定文件时，指定此配置文件，配置文件内容中要用到的内容如下：
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------

    '读配置文件
    Call ReadConf(strFile, "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str参数)
    arr参数 = Split(str参数, ",")
    If UBound(arr参数) <> 3 Then Exit Function
    
    With cnSQL2000
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & arr参数(1) & ";Password=" & arr参数(3) & _
        ";Persist Security Info=True;User ID=" & arr参数(2) & ";Data Source=" & arr参数(0)
        .Open
     
        If .State = adStateClosed Then Exit Function
'        strSQL = "Select smpno,smptype,daterpt,testvalues_text,c.img,c.prev,c.notes From urinerpts a, micslices b,micimages c where a.id=b.id and b.sid=c.sid" & _
'                " And daterpt>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
'                " And daterpt<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
'                IIf(lngSampleNO = -1, "", " And smpno=" & lngSampleNO)
           
        '对方觉得此仪器图形可有可无，所以目前不处理图形
        strSQL = "Select smpno,smptype,daterpt,testvalues_text From urinerpts " & _
        " Where daterpt>=Convert(DATETIME,'" & Format(dtStart, "yyyy-MM-dd") & " 00:00:00',21)" & _
        " And daterpt<=Convert(DATETIME,'" & Format(dtEnd, "yyyy-MM-dd") & " 23:59:59',21)" & _
        IIf(lngSampleNo = -1, "", " And smpno=" & lngSampleNo)
        
        
        
        Set rsTmp = cnSQL2000.Execute(strSQL)
'        If rsTmp.EOF Then
'            .Close
'            Exit Function
'        End If
'
        ReDim aResults(0)
        lngResultCount = 0

        Do Until rsTmp.EOF
            If Val("" & rsTmp("smpno")) > 0 Then '新标本
                str标本号 = Replace(Trim("" & rsTmp("smpno")), "-", "")
                str标本类型 = Trim("" & rsTmp("smptype"))
                If str标本类型 = "" Then str标本类型 = "尿液"
                str检验日期 = Format(Trim(rsTmp("daterpt")), "yyyy-MM-dd hh:mm:ss") '检验时间
                str检验结果 = Trim(ReadTestValue_Text("" & rsTmp("testvalues_text")))

                If Val(str标本号) > 0 And str检验结果 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str检验日期 & "|" & str标本号 & "| |" & str标本类型 & "|0" & str检验结果
                End If

            End If
            rsTmp.MoveNext
        Loop
    End With
        
    ResultFromFile = aResults
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As textStream

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = Trim(arrItems(intLoop, 0)) Then
                        arrItems(intLoop, 1) = Trim(arrLine(1))
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Private Function ReadTestValue_Text(ByVal strTestvalues_text As String) As String
    Dim strPcode As String
    Dim strCode As String
    Dim strValue As String
    Dim strResult As String
    Dim var_line() As String
    Dim var_Item() As String
    Dim strLine As String
    Dim strItem As String
    Dim i As Integer
    Dim j As Integer
    
    var_line = Split(strTestvalues_text, ">" & Chr(&HA) & "<")
    
    For i = LBound(var_line) To UBound(var_line)
        strLine = var_line(i)
        If InStr(var_line(i), ";") Then
            var_Item = Split(var_line(i), ";")
        
            If UBound(var_Item) > 0 Then
                For j = LBound(var_Item) To UBound(var_Item)
                    strItem = Trim(var_Item(j))
                    If strItem Like "pcode*" Then strPcode = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                    If strItem Like "code*" Then strCode = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                    If strItem Like "value*" Then strValue = Trim(Mid(strItem, InStr(strItem, "=") + 1))
                Next
                
                If strPcode <> "root" And strValue <> "" Then
                        strResult = strResult & "|" & strCode & "|" & strValue
                End If
            End If
        End If
    Next
    
    ReadTestValue_Text = strResult
End Function
