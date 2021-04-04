VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_File_HUAYANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:HUAYANG(化洋)化学发光免疫分析仪
'厂商:化洋
'作者:曾超
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
    Dim intMaxValue As Integer
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strChar = Replace(strFile, strFilePath & "\", "")
    strChar = Mid(strChar, 4, 1)
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    lngResultCount = -1
   
    strImgErs = ""
    If strFileErs <> "" Then
        strSQL = "SELECT SampleNO, ExamDate,ItemName,ReThickNess " & _
                 " FROM Exams AS A Where Format(A.ExamDate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.ExamDate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                 "' " & IIf(lngSampleNO = -1, "", " And val(SampleNO)=" & lngSampleNO) & " order by SampleNO "
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsErs("ExamDate"), "yyyy-MM-dd")  '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs("SampleNO") & "| |血清|0|" & rsErs("ItemName") & "|" & rsErs("ReThickNess")
            rsErs.MoveNext
        Loop
    End If
    ResultFromFile = aResults
End Function

Private Function SebiaImg(ByVal strImg As String) As String
'   生成图形数据串
    Do While strImg <> ""
        If Val("&H" & Left(strImg, 4)) < 0 Then
            SebiaImg = "1;" & SebiaImg
        Else
            SebiaImg = CStr(Val("&H" & Left(strImg, 4))) & ";" & SebiaImg
        End If
        strImg = Mid(strImg, 5)
    Loop
    If Right(SebiaImg, 1) = ";" Then SebiaImg = Mid(SebiaImg, 1, Len(SebiaImg) - 1)

End Function


