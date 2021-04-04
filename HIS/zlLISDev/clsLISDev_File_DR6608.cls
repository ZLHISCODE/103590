VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DR6608"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:免疫分析仪 DR6608
'厂商:佛山达安医疗设备有限公司
'作者:陈东
'用户:稷山县医院

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
    Dim i As Integer, strDate As String, blnAdd As Boolean, strBh As String
    On Error GoTo errH
    
         
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=admin;pwd="
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function
    strSQL = "SELECT a.series_date,a.series_sampleid,a.series_itemname,a.series_value" & _
             " FROM  Series AS A Where CDate(a.series_date) >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "# And CDate(a.series_date) <=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#" & _
             " Order by a.series_date,a.series_sampleid"
'    strSQL = strSQL & "  Where cdate(A.sjrq) >=#" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "#"
'    strSQL = strSQL & "   And  cdate(A.sjrq) <=#" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "#" & _
'    IIf(lngSampleNo = -1, "", " And val(A.sxh)=" & lngSampleNo)
    WriteLog "File_DR6608", strSQL, ""
    
    ReDim aResults(0)
    
    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "File_DR6608", "无记录！", ""
        Exit Function
    End If
    strBh = ""
    Do Until rsTmp.EOF
        If IsDate(Trim("" & rsTmp!series_date)) Then
            strDate = Trim("" & rsTmp!series_date)
            blnAdd = False
            If CDate(strDate) >= CDate(Format(dtStart, "yyyy-MM-dd 00:00:00")) And CDate(strDate) <= CDate(Format(dtEnd, "yyyy-MM-dd 23:59:59")) Then
                If lngSampleNo <> -1 Then
                     If Val(rsTmp!series_sampleid) = lngSampleNo Then
                        blnAdd = True
                     End If
                Else
                    blnAdd = True
                End If
 
            End If
            
            If blnAdd Then
                
                If strBh <> "" & rsTmp!series_sampleid Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        
                    aResults(UBound(aResults)) = rsTmp("series_date")  '检验时间
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("series_sampleid")) '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |Bold|0"
                End If
                
                If Trim("" & rsTmp!series_value) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & rsTmp!series_itemname & "|" & rsTmp!series_value
                End If
                 
                strBh = "" & rsTmp!series_sampleid
            End If
        End If
        rsTmp.MoveNext
    Loop
        
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "File_DR6608", CStr(Erl()) & ":" & Err.Description, ""
End Function








