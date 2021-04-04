VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_C2000A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血凝仪 C2000-A
'厂商:北京普利生有限公司
'作者:陈东
'用户:万荣县医院（太原）

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
        Dim cnAccess As New ADODB.Connection
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim strBh As String, strItem As String
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
120         strSQL = "  " & _
                    " Select Val(a.sample_id) as 标本号,a.sample_date as 日期,b.assay_id as 通道码,b.assay_result_1 as 结果,b.assay_result_2 as INR " & _
                    " from sample_info a,assay_info b" & vbNewLine & _
                    " Where a.sample_Key=B.sample_key And a.sample_id<>0 And a.sample_date>=#" & Format(dtStart, "yyyy-MM-dd") & "# And a.sample_date<=#" & Format(dtEnd, "yyyy-MM-dd") & "# " & _
                    IIf(lngSampleNo = -1, "", " And val(a.sample_id)=" & lngSampleNo) & " Order by a.sample_date,a.sample_id"
                
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            End If
 
130         strBh = "***"
132         strItem = ""
134         Do While Not rsTmp.EOF
136             If strBh <> rsTmp("标本号") & "_" & Format(rsTmp("日期"), "yyyy-mm-dd") Then '新标本
138                 If strItem <> "" Then
140                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
142                     aResults(UBound(aResults)) = strItem
                    End If
                
144                 strItem = Format(rsTmp("日期"), "yyyy-mm-dd")  '检验时间
146                 strItem = strItem & "|" & Val("" & rsTmp("标本号"))   '样本序号
148                 strItem = strItem & "| "  '检验人
150                 strItem = strItem & "|血液"  '标本
152                 strItem = strItem & "|0"
154                 strItem = strItem & "|1" & rsTmp!通道码 & "|" & rsTmp!结果
156                 If Trim("" & rsTmp!INR) <> "" And Trim$("" & rsTmp!通道码 = "1") Then
158                     strItem = strItem & "|INR|" & Trim("" & rsTmp!INR)
                    End If
160                 strBh = rsTmp("标本号") & "_" & Format(rsTmp("日期"), "yyyy-mm-dd")
                Else
162                 strItem = strItem & "|1" & rsTmp!通道码 & "|" & rsTmp!结果
164                 If Trim("" & rsTmp!INR) <> "" And Trim$("" & rsTmp!通道码 = "1") Then
166                     strItem = strItem & "|INR|" & Trim("" & rsTmp!INR)
                    End If
                End If
168             rsTmp.MoveNext
            Loop
                
170         If strItem <> "" Then
172             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
174             aResults(UBound(aResults)) = strItem
            End If
        End With
176     ResultFromFile = aResults
        Exit Function
hErr:
178     WriteLog "C2000a", CStr(Erl()) & ":" & Err.Description, ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    
End Sub




