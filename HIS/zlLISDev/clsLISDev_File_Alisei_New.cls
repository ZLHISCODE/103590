VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Alisei_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:酶标仪
'厂商:alisei
'作者:陈东
'用户:烟台山医院,包头中心医院
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rs项目 As ADODB.Recordset, rs时间 As ADODB.Recordset
    Dim lngResultCount As Long, strcutoff As String, str项目 As String
    Dim str日期 As String, str定性结果 As String
    
    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "Select Format(A.Adate,'yyyy-MM-dd') As 日期, ID As 标本号, B.Name As 项目, C.Posneg, C.Result, C.Od, B.Cutoff" & vbNewLine & _
                "From Wrkl A, Test B, Pozzetti C" & vbNewLine & _
                "Where A.Sessionid = C.Sessionid And A.Wrklindex = C.Wrklindex And C.Testindex = B.Testindex And" & vbNewLine & _
                "      C.Sessionid = B.Sessionid And Format(A.Adate, 'yyyy-MM-dd') >= '" & Format(dtStart, "yyyy-MM-dd") & _
                "' And Format(A.Adate, 'yyyy-MM-dd') <= '" & Format(dtEnd, "yyyy-MM-dd") & "'"
        strSQL = strSQL & IIf(lngSampleNO = -1, "", " And ID='" & lngSampleNO & "'")
        Set rsTmp = cnAccess.Execute(strSQL)

        
        ReDim aResults(0)
        Do Until rsTmp.EOF
            
            If Val("" & rsTmp!标本号) > 0 Then
                'lngResultCount = lngResultCount + 1
'                ReDim Preserve aResults(lngResultCount)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                lngResultCount = UBound(aResults)
                str日期 = "" & rsTmp!日期
                aResults(lngResultCount) = str日期  '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp!标本号)   '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                   " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  "血液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '质控
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp!项目 & "|" & rsTmp!Posneg '结果
                aResults(lngResultCount) = aResults(lngResultCount) & "^" & Val("" & rsTmp!OD) & _
                                                                      "^" & Val("" & rsTmp!CutOff) & _
                                                                      "^" & Val("" & rsTmp!result)
            
            End If
            rsTmp.MoveNext
      Loop
    End With
    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "Alisei_new", Err.Number, Err.Description
End Function






