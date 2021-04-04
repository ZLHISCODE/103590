VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_C2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:C2000-2
'厂商:
'作者:陈东
'用户:中条山医院

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
'120         strSQL = "select sentdate as 日期,samplenum as 标本号,testitem ,TestTime,TestData1,TestData2" & vbNewLine & _
'                     " From xnydata " & vbNewLine & _
'                     " Where sentdate is not null and sentdate>=#" & Format(dtStart, "yyyy-MM-dd") & "# And sentdate<=#" & Format(dtEnd, "yyyy-MM-dd") & "# " & _
'                     IIf(lngSampleNO = -1, "", " And val(samplenum)=" & lngSampleNO) & " Order by sentdate,samplenum"
120         strSQL = "select Format(sentdate,'yyyy-MM-dd') as 日期,samplenum as 标本号,testitem ,TestTime,TestData1,TestData2" & vbNewLine & _
                     " From xnydata " & vbNewLine & _
                     " Where sentdate is not null " & _
                     IIf(lngSampleNO = -1, "", " And val(samplenum)=" & lngSampleNO) & " Order by sentdate,samplenum"
            
122         Set rsTmp = cnAccess.Execute(strSQL)
124         If rsTmp.EOF Then
126             .Close
128             ResultFromFile = aResults
                Exit Function
            Else
129             rsTmp.Filter = "日期>=#" & Format(dtStart, "yyyy-MM-dd") & "# And 日期<=#" & Format(dtEnd, "yyyy-MM-dd") & "# "
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
154                 If Trim("" & rsTmp!TestTime) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "1|" & rsTmp!TestTime
                    If Trim("" & rsTmp!TestData1) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "2|" & rsTmp!TestData1
                    If Trim("" & rsTmp!TestData2) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "3|" & rsTmp!TestData2
156                 strBh = rsTmp("标本号") & "_" & Format(rsTmp("日期"), "yyyy-mm-dd")
                Else
158                 If Trim("" & rsTmp!TestTime) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "1|" & rsTmp!TestTime
                    If Trim("" & rsTmp!TestData1) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "2|" & rsTmp!TestData1
                    If Trim("" & rsTmp!TestData2) <> "" Then strItem = strItem & "|" & rsTmp!testitem & "3|" & rsTmp!TestData2
                End If
160             rsTmp.MoveNext
            Loop
                
162         If strItem <> "" Then
164             If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
166             aResults(UBound(aResults)) = strItem
            End If
        End With
168     ResultFromFile = aResults
        Exit Function
hErr:
170     WriteLog "C2000", CStr(Erl()) & ":" & Err.Description, ""
End Function

