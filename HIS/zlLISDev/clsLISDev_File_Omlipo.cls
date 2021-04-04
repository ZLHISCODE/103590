VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Omlipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Password As String
'仪器:深圳国赛特定蛋白分析仪
'型号：Omlipo
'作者:
'用户:宁明县人民医院
'使用方法:读取mdb数据库文件
'**        ：


Public Function GetStartCmd() As String
    '第一次启动时调用
End Function

Public Function GetEndCmd() As String
    '接口关闭的时候调用
End Function

Public Function GetAnswerCmd() As String
    '定时发送的指令
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
    Dim strSQL As String, strFilePath As String
    Dim objFile As New Scripting.FileSystemObject
    Dim errdb As New ADODB.Connection
    Dim cnAccess As New ADODB.Connection, aResults() As String
    Dim rsTmp As ADODB.Recordset
    Dim strSample As String, ResultCount As Long
    Dim str日期 As String, Study_Time As String
    Dim Series_ItemName As String, Series_Value As String
    
    
    On Error GoTo errHand
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function
    Password = "gsOmlipo"
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";pwd=" & Password & ";Driver={Microsoft Access Driver (*.mdb)}"  '打开数据库
        .Open
        If errdb.State = adStateClosed Then
           strSQL = "SELECT a.csampleNo,a.cdateDiagnose ,b.iprojectInfo ,b.cresult  " & vbNewLine & _
                    "from gs_diagnoseResult a ,gs_diagnoseDetail b " & vbNewLine & _
                    "Where a.idiagnoseId=b.idiagnoseId " & _
                    " and Format(A.cdateDiagnose,'yyyy-mm-dd') Between " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " and " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & _
                    IIf(lngSampleNO = -1, "", " and  val(a.csampleNo) =" & lngSampleNO) & " order by A.csampleNo"
           Set rsTmp = cnAccess.Execute(strSQL)
            
            If rsTmp.EOF Then
                .Close
                ReDim Preserve aResults(0)
                aResults(0) = ""
                ResultFromFile = aResults
            Exit Function
                
            End If
            
            ResultCount = -1
            strSample = "***"
            Do Until rsTmp.EOF
                If strSample <> rsTmp("csampleNo") Then
                ResultCount = ResultCount + 1
                ReDim Preserve aResults(ResultCount)
                aResults(ResultCount) = Format(rsTmp("cdateDiagnose") & " " & Study_Time, "yyyy-mm-dd hh:mm:ss") '检验时间
                str日期 = Format(rsTmp("cdateDiagnose"), "yyyy-mm-dd")
                aResults(ResultCount) = aResults(ResultCount) & "|" & Val(" " & rsTmp("csampleNo")) '标本序号
                aResults(ResultCount) = aResults(ResultCount) & "|" & " " '检验人
                aResults(ResultCount) = aResults(ResultCount) & "|" & "血清" '标本
                aResults(ResultCount) = aResults(ResultCount) & "|0" '是否质控
                
                
                
                strSample = rsTmp("csampleNo")
                End If
                If rsTmp("cresult") <> "" Then
                    aResults(ResultCount) = aResults(ResultCount) & "|" & _
                    rsTmp("iprojectInfo") & "|" & Replace(rsTmp("cresult"), "↑", "")
                    
                End If
            rsTmp.MoveNext
            Loop
        End If
    End With
    ResultFromFile = aResults
    
    Exit Function
    
errHand:
    WriteLog "omlipo", Err.Number & ":" & Err.Description, ""
End Function



