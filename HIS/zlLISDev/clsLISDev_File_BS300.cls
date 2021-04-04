VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_BS300"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:BS300 生化分析仪
'厂商:迈瑞
'作者:陈东
'用户:盐湖区医院

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
        Dim strUID As String, strPWD As String, strDate As String, strNo As String
        Dim strItem As String, strResult As String, strValue As String
        On Error GoTo errH

106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
        
        If Dir(App.Path & "\BS300.ini") <> "" Then
            strUID = ReadIni("BS300", "UID", App.Path & "\BS300.ini", "")
            strPWD = ReadIni("BS300", "PWD", App.Path & "\BS300.ini", "BS300")
        Else
            strUID = ""
            strPWD = "BS300"
            WriteIni "BS300", "UID", "", App.Path & "\BS300.ini"
            WriteIni "BS300", "PWD", "BS300", App.Path & "\BS300.ini"
        End If
110     cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=" & strUID & ";pwd=" & strPWD
112     cnAccess.Open
114     If cnAccess.State = adStateClosed Then Exit Function
        strSQL = "SELECT A.Sample_date, A.sample_no, A.Conc, B.name" & _
                 " FROM TESTDATAREC AS A, Item AS B " & _
                 " Where (((a.ITEM_ID) = [B].[ITEM_ID])) And A.Sample_date>=" & Format(dtStart, "yyyyMMdd") & _
                 " And A.Sample_date <=" & Format(dtEnd, "yyyyMMdd") & _
                  IIf(lngSampleNO = -1, "", " And val(A.sample_no)=" & lngSampleNO) & _
                 " ORDER BY CStr(A.sample_date & '  ' & a.sample_no) "

134     ReDim aResults(0)
    
136     Set rsTmp = cnAccess.Execute(strSQL)
138     If rsTmp.EOF Then
140         cnAccess.Close
142         WriteLog "BS300", "无记录！", strSQL
            Exit Function
        End If
        strDate = "": strNo = ""
        strResult = ""
144     Do Until rsTmp.EOF
            
            If strDate <> Trim("" & rsTmp.Fields("Sample_date")) Or strNo <> Trim("" & rsTmp.Fields("Sample_NO")) Then
                
                If strResult <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strResult
                    
                End If
                

                strDate = Trim("" & rsTmp.Fields("Sample_date"))
                strNo = Trim("" & rsTmp.Fields("Sample_no"))
                
                strResult = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7, 2) & "|" & strNo & "| |Boold|0"
                
                
            End If
            strItem = Trim("" & rsTmp.Fields("Name"))
            strValue = Trim("" & rsTmp.Fields("Conc"))
    
            If strItem <> "" And strValue <> "" Then
                strResult = strResult & "|" & strItem & "|" & strValue
            End If
            
184         rsTmp.MoveNext
        Loop
        If strResult <> "" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = strResult
        End If
186     ResultFromFile = aResults
        Exit Function
errH:
188     WriteLog "BS300", CStr(Erl()) & ":" & Err.Description, ""
End Function


