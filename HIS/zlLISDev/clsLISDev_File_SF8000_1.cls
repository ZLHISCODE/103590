VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SF8000_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:SF-8000
'厂商:北京赛科希德科技发展有限公司 SF系列全自动凝血测试
'作者:董雨
'用户:长安区医院

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
    Dim blnUnion As Boolean
    Dim intLoop As Integer

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False
        strSQL = "SELECT 编号,max(标本号1) as 标本号,max(测试日期1) as 测试日期,MIN(测试时间1) as 测试时间,max(检验医师1) as 检验医师,max(标本类型1) as 标本类型,max(PT_sec1) as PT_sec,max(PT_INR1) as PT_INR ," _
                 & "max(PT_11) as PT_1 ,max(PT_PTR1) as PT_PTR ,max(APTT1) as APTT,max(FIB1) as FIB ,max(TT1) as TT FROM(select FORMAT(A.测试日期,'YYYYMMDD')+FORMAT(A.标本号,'000') as 编号,A.标本号 as 标本号1,A.测试日期 as 测试日期1," _
                 & "format(A.测试时间,'HH:NN:SS') as 测试时间1,B.检验医师 AS 检验医师1,B.标本类型 AS 标本类型1," _
                 & "iif(A.项目名称='PT' and A.单位='S',A.测试结果,'') as PT_sec1," _
                 & "iif(A.项目名称='PT' and A.单位='INR',A.测试结果,'') as PT_INR1," _
                 & "iif(A.项目名称='PT' and A.单位='%',A.测试结果,'') as PT_11," _
                 & "iif(A.项目名称='PT' and A.单位='PTR',A.测试结果,'') as PT_PTR1," _
                 & "iif(A.项目名称='APTT' ,A.测试结果,'') as APTT1," _
                 & "iif(A.项目名称='FIB' and A.单位='g/L',A.测试结果,'') as FIB1," _
                 & "iif(A.项目名称='TT' ,A.测试结果,'') as TT1 " _
                 & "from TB_TEST_RESULT A,TB_TEST_DATA B WHERE  " & _
                 IIf(lngSampleNO = -1, "  A.测试日期 >=#" & Format(dtStart, "yyyy-mm-dd") & "#" & _
                 " And A.测试日期 <=#" & Format(dtEnd, "yyyy-mm-dd") & "#", _
                 "  A.测试日期 =#" & Format(dtStart, "yyyy-mm-dd") & "# and  A.标本号 =" & lngSampleNO) & " and A.标本号 = B.标本号 And A.测试日期 = B.测试日期) GROUP BY 编号   ORDER BY 编号+0"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp("测试日期"), "YYYY-MM-DD ") & Format(rsTmp("测试时间"), "HH:NN:SS")
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("标本号") & "|" & IIf(IsNull(rsTmp("检验医师")), " ", rsTmp("检验医师")) & "|" & IIf(IsNull(rsTmp("标本类型")), "血清", rsTmp("标本类型")) & "|0"
            aResults(lngResultCount) = aResults(lngResultCount) & "|PT-sec|" & Nvl(rsTmp("PT_sec"), " ") & "|PT-INR|" & Nvl(rsTmp("PT_INR"), " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|PT-%|" & Nvl(rsTmp("PT_1"), " ") & "|PT_PTR|" & Nvl(rsTmp("PT_PTR"), " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|APTT|" & Nvl(rsTmp("APTT"), " ") & "|FIB|" & Nvl(rsTmp("FIB"), " ") & "|TT|" & Nvl(rsTmp("TT"), " ")
            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_SF8000", strFile, Err.Number & " " & Err.Description
End Function
