VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_GFE2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:GF-EⅡ 生化分析仪
'厂商:山东高密彩虹分析仪器有限公司
'作者:罗云龙
'时间:2015-01-15
'用户:库车县第二人民医院

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
    
    On Error GoTo ErrH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
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
    
    If strFileErs <> "" Then
        strSQL = "select Cdate(sDate) As 日期,PatID As 标本号,ItemID As 项目,Value As 结果 from Result Where" & " CDate(sDate)>=CDate('" & Format(dtStart, "yyyy-MM-dd") & "') And CDate(sDate)<=CDate('" & Format(dtEnd, "yyyy-MM-dd") & "')" & _
                IIf(lngSampleNO = -1, "", " And val(PatID)=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            str日期 = rsErs("日期")
            
            aResults(lngResultCount) = Format(str日期, "yyyy-MM-dd")  '检验时间
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val("" & rsErs("标本号"))   '样本序号
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  " "  '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  "血液"  '标本
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Trim("" & rsErs!项目) & "|" & Val("" & rsErs!结果)

            rsErs.MoveNext
        Loop
    End If
    ResultFromFile = aResults
    Exit Function
ErrH:
    If Error = 1 Then
        Resume
    End If
End Function


