VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_FIle_LTS_E100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:LTS-E100粪便沉渣分析
'厂商:珠海丽拓
'作者:徐江
'用户:怀化市第二人民医院
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo errHand '打开错误陷阱
    '------------------------------------------------
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
    Dim strSQL As String, rsNO As ADODB.Recordset, rsItem As ADODB.Recordset, rsImage As ADODB.Recordset
    Dim strBh As String, strAnti As String, str结果 As String, strImage As String, str日期 As String, str标本号
    Dim lngResultCount As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=gsOmlipo"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "SELECT 对应样品号,检验时间,样品编号"
        strSQL = strSQL & " FROM Cases "
        strSQL = strSQL & " Where  Cases.检验时间 between CDate('" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "') and cdate( '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') Order BY 对应样品号"
                      
        Set rsNO = cnAccess.Execute(strSQL)
        If rsNO.EOF Then
            .Close
            Exit Function
        End If
        
         ReDim aResults(0)
        lngResultCount = 0
        
        Do While (Not rsNO.EOF)
           
            If Val("" & rsNO("对应样品号")) > 0 Then '新标本
                
                str日期 = Replace("" & rsNO("检验时间"), "/", "-") '检验时间
                str结果 = ""
                
                strSQL = "SELECT 样品编号,实验项目类型,结果"
                strSQL = strSQL & " FROM SpecimentCheckResult "
                strSQL = strSQL & " Where  SpecimentCheckResult.样品编号 ='" & rsNO("样品编号") & " '"
                              
                Set rsItem = cnAccess.Execute(strSQL)
                If rsItem.EOF Then
                    GoTo nextno:
                End If
                Do While (Not rsItem.EOF)
                    str结果 = str结果 & "|" & rsItem("实验项目类型") & "|" & rsItem("结果")
                    rsItem.MoveNext
                Loop
               
                strBh = "***"
        
                strBh = "" & rsNO("对应样品号")
                If strBh <> "***" And str结果 <> "" Then
                    str标本号 = "" & rsNO("对应样品号")
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(str标本号) '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血清"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str结果
                End If
                strSQL = "SELECT 样品编号,图片名称,图片路径"
                strSQL = strSQL & " FROM SpecimentResult "
                strSQL = strSQL & " Where  SpecimentResult.样品编号 ='" & rsNO("样品编号") & " '"
                              
                Set rsImage = cnAccess.Execute(strSQL)
                If rsImage.EOF Then
                    GoTo nextno:
                End If
                
                strImage = ""
                Do While (Not rsImage.EOF)
                
                    strImage = strImage & "^" & rsImage("图片名称") & ";110;0;" & rsImage("图片路径")
                    
                    rsImage.MoveNext
                Loop
                If strImage <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strImage, 2)
                End If
            End If
nextno:
            rsNO.MoveNext
        Loop
        
    End With
    Dim i As Integer
    For i = 0 To UBound(aResults)
        Debug.Print aResults(i)
    Next
    ResultFromFile = aResults
   
    '------------------------------------------------
    Exit Function
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\YFErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_FIle_LTS_E100"
    Close #lngErrFileNo
End Function











