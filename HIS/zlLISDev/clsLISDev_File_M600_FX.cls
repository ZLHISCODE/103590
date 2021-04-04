VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_M600_FX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:凝血 M600
'用户:丰县中医院

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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBH As String
    Dim blnUnion As Boolean, strimg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String, str流水号 As String
    
    On Error GoTo Errhand
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        '取今天做过的标本
        strSQL = "select distinct a.SampleNo as 流水号,SerialNo as 标本号,format(InputTime,'yyyy-MM-dd') as 日期 from ToDayDetail a, ToDayMain b " & _
        "where a.SampleNo=b.SampleNo and Format(InputTime,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(InputTime,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' " & _
        IIf(lngSampleNO = -1, "", " And INSTID_ID= '" & lngSampleNO & "'")
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        
        lngResultCount = -1
        strBH = ""
        str流水号 = ""
        Do While Not rsTmp.EOF
            If str流水号 <> rsTmp("流水号") Then '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("日期"), "yyyy-mm-dd")  '检验日期
                
                str日期 = Format(rsTmp("日期"), "yyyy-mm-dd")
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("标本号"))   '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "血液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("质控")) '质控
                    
                str流水号 = rsTmp("流水号")
            End If
          '开始查询结果
            strSQL = "select Alias as 项目,Second as 结果 " & _
                      " from ToDayDetail a, ToDayMain b where a.SampleNo=b.SampleNo and a.SampleNo= '" & str流水号 & "'" & _
                      " union all " & _
                      "select Unit2 as 项目,fResult2 as 结果 " & _
                      " from ToDayDetail a, ToDayMain b where a.SampleNo=b.SampleNo and a.SampleNo= '" & str流水号 & "' "
            Set rsErs = cnAccess.Execute(strSQL)
                Do While Not rsErs.EOF
                    If Trim(rsErs("项目")) <> "" Then
                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                             rsErs("项目") & "|" & Trim(Val(IIf(IsNull(rsErs("结果")), "0", rsErs("结果"))))
                    End If
                    rsErs.MoveNext
                Loop
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    
    Exit Function
Errhand:
    WriteLog "M600_FX", Err.Number & "：" & Err.Description, ""
End Function














