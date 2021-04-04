VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_WanFu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:万孚超敏C反应

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
    On Error GoTo ErrHand '打开错误陷阱
    '------------------------------------------------
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsTmp1 As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strIMG As String, iField As Integer, i As Integer, iField1 As Integer
    Dim i1 As Integer
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String
    Dim strChar As String
    Dim intMaxValue As Integer
    Dim strsql1 As String
    Dim aa As Double, bb As Long, cc As Double, dd As Double, ee As Double, ff As Double
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    If Dir(strFile) = "" Then Exit Function

    With cnAccess

        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT  Swatch ,ProName ,Result ,CheckDate " & _
                 "FROM Result as A where  format(a.CheckDate,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And format(a.CheckDate,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And 序号=" & lngSampleNO)
        'strSQL = strSQL & "order by CheckDate"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If

        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF

            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)

            str日期 = rsTmp("CheckDate")
            aResults(lngResultCount) = str日期  '检验时间

            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val("" & rsTmp("Swatch"))   '样本序号

            aResults(lngResultCount) = aResults(lngResultCount) & "|" & " "   '检验人

            aResults(lngResultCount) = aResults(lngResultCount) & "|" & "血液"   '标本

            aResults(lngResultCount) = aResults(lngResultCount) & "|0"
            ' Val(rsTmp("质控")) '质控
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsTmp("ProName") & "|" & rsTmp("Result")

            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    '------------------------------------------------
    Exit Function
    '----------------
ErrHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\ZLlisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLISDev_KJ680"
    Close #lngErrFileNo
End Function
