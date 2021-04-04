VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'医院名称：祁东县人民医院
'仪器名称：免疫定量分析仪
'仪器型号：FIA-4
'仪器厂家：丰华
'作    者：刘潇宇

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
 
        On Error GoTo ResultFromFile_Err

    
        Dim varTemp As Variant '存放输入数据
        Dim cnAccess As New ADODB.Connection
        Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
        Dim aResults() As String, objFile As New Scripting.FileSystemObject
        Dim strFilePath As String, strFilename As String
        Dim strResult() As String
        Dim tempdt As String
        Dim objStream As textStream
        Dim lngResultCount As Long, strBh As String
        Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
        Dim strStartDt As String, strEndDt As String

        Dim strImgErs As String, str日期 As String
    
        Dim lngSampleID As String
        '避免出现下标错误
100     ReDim aResults(0)
102     aResults(0) = ""
104     ResultFromFile = aResults

106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};"
114         .Open
116         If .State = adStateClosed Then Exit Function
            'Sample_Pos OR Par_Id
118         strStartDt = Format(dtStart, "yyyy-MM-dd")
120         strStartDt = Mid(strStartDt, 1, 4) & Mid(strStartDt, 6, 2) & Mid(strStartDt, 9, 2)
        
122         strEndDt = Format(dtEnd, "yyyy-MM-dd")
124         strEndDt = Mid(strEndDt, 1, 4) & Mid(strEndDt, 6, 2) & Mid(strEndDt, 9, 2)
        
            'SQL区分于AutoTRFIA-2型，FIA-4多了定性结果，结果单独保存于数据库表C15行
        
126         strSQL = "select A.A0 as 记录ID, A.A2 as 记录编号, A.A17 as 检验时间,  B.C2  as 病人编号, B.C4 as 项目名称 ,B.C6 as 项目结果, b.C15 as 项目结果v" & _
                         " from h_Patient a , h_Result b " & _
                         " where a.A1=b.C2  and a.A17>=" & Chr(39) & strStartDt & Chr(39) & " And a.A17<= " & Chr(39) & strEndDt & Chr(39) & _
                         " Order By A.A2"
128         Set rsTmp = cnAccess.Execute(strSQL)
130         If rsTmp.EOF Then
132             .Close
                Exit Function
            End If
        
134            lngResultCount = -1
136         strBh = "***"
138         Do While Not rsTmp.EOF
140             If strBh <> rsTmp("记录编号") Then '新标本
142                 lngResultCount = lngResultCount + 1
144                 ReDim Preserve aResults(lngResultCount)
                
146                 aResults(lngResultCount) = Mid(rsTmp("检验时间"), 1, 4) & "-" & Mid(rsTmp("检验时间"), 5, 2) & "-" & Mid(rsTmp("检验时间"), 7, 2) '检验时间
                
148                 str日期 = Mid(rsTmp("检验时间"), 1, 4) & "-" & Mid(rsTmp("检验时间"), 5, 2) & "-" & Mid(rsTmp("检验时间"), 7, 2) 'Format(rsTmp("检验时间"), "yyyy-mm-dd")
                
150                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Val("" & rsTmp("记录编号"))   '样本序号
152                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          " "  '检验人
154                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          "血液"  '标本
156                 aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                       ' Val(rsTmp("质控")) '质控
                    
158                 strBh = rsTmp("记录编号")
                End If

160                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        rsTmp("项目名称") & "|" & IIf(IsNull(rsTmp("项目结果")) Or rsTmp("项目结果") = "", rsTmp("项目结果v"), rsTmp("项目结果")) & " " '结果

162             rsTmp.MoveNext
           Loop
164        ResultFromFile = aResults
        End With

        Exit Function

ResultFromFile_Err:
    If Error = 1 Then
        Resume
    End If
End Function








