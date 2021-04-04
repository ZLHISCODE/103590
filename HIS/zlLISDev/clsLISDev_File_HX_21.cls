VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_HX_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HX-21细菌鉴定药敏分析仪
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
        .ConnectionString = "Provider=MSDASQL.1;Password=weichang;Persist Security Info=True;User ID=admin;Extended Properties=" _
                            & """" & "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}" & """"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False

        strSQL = "SELECT Format(B.Repdate,'yyyy-mm-dd') as 送检日期,B.Bbh as 标本号,B.Bgys as  检验人,B.jzname AS 标本类型, " & _
                 "Y.ypmc as 抗生素,iif(Y.mg='耐药', 'R',iif(Y.mg='敏感', 'S', 'I')) AS 药敏结果," & _
                 "'1' as 药敏方法,Y.MIC AS 结果 FROM Base B,Ypbase Y WHERE B.useid = Y.useid and " & _
                 IIf(lngSampleNO = -1, "  b.Repdate between #" & Format(dtStart, "yyyy-mm-dd") & " 00:00:00#" & _
                 " And #" & Format(dtEnd, "yyyy-mm-dd") & " 23:59:59#", _
                 " B.Bbh ='" & lngSampleNO & "'") & " ORDER BY B.Repdate,B.Bbh"
        'cydate  送检日期         Repdate 报告日期
        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        lngResultCount = -1

        ReDim Preserve aResults(0)
        rsTmp.MoveFirst
        Dim 送检日期, 标本号, 标本类型 As String

        Do Until rsTmp.EOF

            If 送检日期 <> Format(rsTmp("送检日期"), "YYYY-MM-DD") Or 标本号 <> Nvl(rsTmp("标本号")) Or 标本类型 <> Nvl(rsTmp("标本类型")) Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                送检日期 = Format(rsTmp("送检日期"), "YYYY-MM-DD")
                标本号 = Nvl(rsTmp("标本号"))
                标本类型 = Nvl(rsTmp("标本类型"))
                aResults(lngResultCount) = 送检日期
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("标本号"), " ") & "|" & Nvl(rsTmp("检验人"), " ") & "|" & Nvl(rsTmp("标本类型"), " ") & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("抗生素"), " ") & "|" _
                         & Nvl(rsTmp("药敏结果"), " ") & "^" & Nvl(rsTmp("药敏方法"), " ") & "^" & Nvl(rsTmp("结果"), " ")
            ElseIf Nvl(rsTmp("抗生素")) <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp("抗生素"), " ") & "|" _
                         & Nvl(rsTmp("药敏结果"), " ") & "^" & Nvl(rsTmp("药敏方法"), " ") & "^" & Nvl(rsTmp("结果"), " ")
            End If
            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_HX_21", strFile, Err.Number & " " & Err.Description
End Function
