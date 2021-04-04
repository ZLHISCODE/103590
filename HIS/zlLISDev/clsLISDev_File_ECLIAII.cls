VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ECLIAII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:ECLIA- Ⅱ型  化学发光免疫分析仪
'厂商:石家庄华洋分析仪器有限公司
'作者:陈东
'用户:上海一钢医院

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
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strBh As String
    
        On Error GoTo hErr
    
100     ReDim aResults(0)
102     strFilePath = objFile.GetParentFolderName(strFile)
104     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
106     If Dir(strFile) = "" Then
108         ResultFromFile = aResults
            Exit Function
        End If
    
110     With cnAccess
112         .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
114         .Open
116         If .State = adStateClosed Then
118             ResultFromFile = aResults
                Exit Function
            End If
                
120         strSQL = "Select Format(ExamDate,'yyyy-MM-dd') as 日期,val(SampleNo) as 标本号 ,ItemName as 项目,RethickNess as 检验结果" & _
                     " From 查询1 " & _
                     " Where trim(ItemName)<>'' And trim(RethickNess)<>'' And " & _
                     " ExamDate>=#" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                     " and ExamDate<=#" & Format(dtEnd, "yyyy-MM-dd") & "# " & _
                     IIf(lngSampleNO = -1, "", " And val(SampleNo)=" & lngSampleNO) & " Order by Examdate,SampleNo"
                 
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strBh = "***"
132         Do While Not rsTmp.EOF
134             If strBh <> rsTmp("标本号") And rsTmp("标本号") <> 0 Then '新标本
136                 If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
138                 aResults(UBound(aResults)) = Format(rsTmp("日期"), "yyyy-mm-dd")  '检验时间
                
140                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsTmp("标本号"))   '样本序号
142                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "  '检验人
144                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血液"  '标本
146                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
148                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp!项目 & "|" & rsTmp!检验结果
150                 strBh = rsTmp("标本号")
                ElseIf strBh = rsTmp("标本号") And rsTmp("标本号") <> 0 Then
152                 aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp!项目 & "|" & rsTmp!检验结果
                End If
            
154             rsTmp.MoveNext
            Loop
        End With
156     ResultFromFile = aResults
        Exit Function
hErr:
158     WriteLog "ECLIAII", CStr(Erl()) & "," & Err.Description, strSQL
End Function

