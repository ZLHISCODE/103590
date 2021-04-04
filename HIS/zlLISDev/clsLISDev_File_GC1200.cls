VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_GC1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Gc1200 Υ射线放免仪
'厂商:
'作者:陈东
'用户:南京江北人民医院
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
    Dim conDBF As New ADODB.Connection
    Dim rsDbf As New ADODB.Recordset
    Dim aResults() As String
    Dim lngResultCount As Long
    Dim TableName As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strSQL As String, strFilePath As String
    Dim str日期 As String, str标本号 As String, str检验项目 As String, str定性结果 As String
    
    ReDim aResults(0)
    
    If Trim(Dir(strFile)) = "" Then Exit Function
    
    strFilePath = objFile.GetParentFolderName(strFile)
    
    TableName = UCase(Replace(strFile, strFilePath & "\", ""))
    If conDBF.State = 1 Then conDBF.Close
    conDBF.ConnectionString = "Provider=microsoft.jet.oledb.4.0;data source=" & strFilePath & ";" & _
                          "Extended Properties=dBase 5.0"
    conDBF.Open
    strSQL = "Select * From " & TableName & " Where mid(BRXMCLBM,1,8) >=  '" & Format(dtStart, "yyyyMMdd") & "' And " & _
             "  mid(BRXMCLBM,1,8) <= '" & Format(dtEnd, "yyyyMMdd") & "' " & _
              IIf(lngSampleNO = -1, "", " And  BRID=" & lngSampleNO)

    Set rsDbf = conDBF.Execute(strSQL)
    lngResultCount = -1
    Do Until rsDbf.EOF
        If Trim("" & rsDbf!result) <> "" Then
            str日期 = Mid(rsDbf!BRXMCLBM, 1, 8)
            str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7)
            str检验项目 = Trim("" & rsDbf!xmmc)
            str定性结果 = Trim("" & rsDbf!result)
            str标本号 = Trim("" & rsDbf!BRID)
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            If IsDate(CDate(str日期)) And Val(str标本号) > 0 And str检验项目 <> "" Then
                aResults(lngResultCount) = str日期 & "|" & Val(str标本号) & "| |血液|0|" & str检验项目 & "|" & str定性结果
            End If
        End If
        rsDbf.MoveNext
    Loop
    ResultFromFile = aResults
End Function

