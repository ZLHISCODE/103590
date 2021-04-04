VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HTF_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HTF-1化学发光免疫分析仪     唐氏筛查
'厂商:
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
        .ConnectionString = "DBQ=" & strFile & ";User ID=admin;Password=weichang;DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = ""
        blnUnion = False

        strSQL = "SELECT 编号 as 标本号,检验医师,标本 ,检测日期,AFP,HCG,PAPPA,AFPMOM,HCGMOM,PAPPAMOM,DS,ET,NTD,ASR,AFPASR,NAFP,NHCG,NPAPPA from T_dssign WHERE  " & _
                 IIf(lngSampleNO = -1, "  检测日期 between #" & Format$(dtStart, "yyyy-mm-dd") & " 00:00:00#" & _
                 " And #" & Format$(dtEnd, "yyyy-mm-dd") & " 23:59:59#", _
                 "  检测日期 between #" & Format$(dtStart, "yyyy-mm-dd") & " 00:00:00# and #" & Format$(dtStart, "yyyy-mm-dd") & " 23:59:59# and  编号 =" & lngSampleNO) & " ORDER BY 编号+0"

        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format$(rsTmp("检测日期"), "YYYY-MM-DD HH:NN:SS")
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("标本号") & "|" & IIf(IsNull(rsTmp("检验医师")), " ", rsTmp("检验医师")) & "|" & IIf(IsNull(rsTmp("标本")), "血清", rsTmp("标本")) & "|0"
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("AFP")), "", "|AFP|" & rsTmp("AFP")) & IIf(IsNull(rsTmp("HCG")), "", "|HCG|" & rsTmp("HCG"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("PAPPA")), "", "|PAPPA|" & rsTmp("PAPPA")) & IIf(IsNull(rsTmp("AFPMOM")), "", "|AFPMOM|" & rsTmp("AFPMOM"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("HCGMOM")), "", "|HCGMOM|" & rsTmp("HCGMOM")) & IIf(IsNull(rsTmp("PAPPAMOM")), "", "|PAPPAMOM|" & rsTmp("PAPPAMOM"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("DS")), "", "|DS|" & rsTmp("DS")) & IIf(IsNull(rsTmp("ET")), "", "|ET|" & rsTmp("ET"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("NTD")), "", "|NTD|" & rsTmp("NTD")) & IIf(IsNull(rsTmp("ASR")), "", "|ASR|" & rsTmp("ASR"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("AFPASR")), "", "|AFPASR|" & rsTmp("AFPASR")) & IIf(IsNull(rsTmp("NAFP")), "", "|NAFP|" & rsTmp("NAFP"))
            aResults(lngResultCount) = aResults(lngResultCount) & IIf(IsNull(rsTmp("NHCG")), "", "|NHCG|" & rsTmp("NHCG")) & IIf(IsNull(rsTmp("NPAPPA")), "", "|NPAPPA|" & rsTmp("NPAPPA"))
            aResults(lngResultCount) = Replace$(aResults(lngResultCount), "|.", "|0.")
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_HTF_1", strFile, Err.Number & " " & Err.Description
End Function
