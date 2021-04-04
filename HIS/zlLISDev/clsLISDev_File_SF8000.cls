VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SF8000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:凝血分析仪 SF-8000
'厂商:
'作者:陈东
'用户:贵阳肺科医院
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
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String
    Dim strChar As String
    Dim intMaxValue As Integer, strFormat As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    strFileErs = strFile
    If Dir(App.Path & "\SF8000.ini") <> "" Then
        strFormat = ReadIni("SF8000", "格式", App.Path & "\SF8000.ini", "")
    Else
        strFormat = ""
        Call WriteIni("SF8000", "格式", "", App.Path & "\SF8000.ini")
    End If
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
        If .State <> adStateClosed Then .Close
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    ReDim aResults(0)
    
    If strFileErs <> "" Then
        If strFormat = "中文字段" Then
            strSQL = "Select 标本号 as 顺序号,测试日期 as 日期,项目名称+单位 as 项目,测试结果 as 结果 From TB_TEST_RESULT " & _
                   " Where 测试日期 >= #" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                     " And 测试日期 <= #" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
                    IIf(lngSampleNO = -1, "", " And 标本号 =" & lngSampleNO)
        Else
            strSQL = "Select TEST_ID as 顺序号,RESULT_DATE as 日期,RESULT_NAME+RESULT_UNIT as 项目,RESULT_VALUE as 结果 From TB_TEST_RESULT " & _
                   " Where RESULT_DATE >= #" & Format(dtStart, "yyyy-MM-dd") & "#" & _
                     " And RESULT_DATE <= #" & Format(dtEnd, "yyyy-MM-dd") & "#" & _
                    IIf(lngSampleNO = -1, "", " And TEST_ID =" & lngSampleNO)
        End If
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
            str日期 = Format("" & rsErs("日期"), "yyyy-MM-dd")
            aResults(UBound(aResults)) = str日期   '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val("" & rsErs("顺序号"))   '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| |血清|0"  '检验人,标本,质控
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Replace(Trim("" & rsErs!项目), "g/L", "") & "|" & Val("" & rsErs!结果)

            rsErs.MoveNext
        Loop
    End If
        

    ResultFromFile = aResults
End Function




