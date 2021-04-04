VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DS_3C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:微量无素分析仪 DS-3C
'厂商:山东东盛电子仪器有限公司
'作者:陈东
'用户:晋煤集团医院
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
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String
    Dim strChar As String
    Dim intMaxValue As Integer
    
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
        strsql = "SELECT val(Mid(A.PATNUM,9)) As 顺序号, A.PATJYSJ as 日期, " & _
                 "B.NAME as 项目, B.CONTENT as 结果 " & _
                 " FROM PAT_INFO AS A,PAT_CONT As B Where A.PATNUM=B.PATNUM And " & _
                 " A.PATJYSJ>='" & Format(dtStart, "yyyyMMdd") & _
                 "' And A.PATJYSJ<='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                IIf(lngSampleNO = -1, "", " And val(Mid(PATNUM,9))=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strsql)
        Do Until rsErs.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            str日期 = "" & rsErs("日期")
            str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2)
            aResults(lngResultCount) = Format(str日期, "yyyy-MM-dd")  '检验时间
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val("" & rsErs("顺序号"))   '样本序号
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
End Function


