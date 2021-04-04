VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_TaiLai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: 泰莱-1 酶标仪
'厂商: 未知
'作者: 陈东
'用户: 德阳人民医院
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
    Dim strID As String
    Dim rsErs As ADODB.Recordset
    Dim str项目结果 As String, str日期 As String, lng标本号 As Long
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "SELECT A2 as 标本号,A16 as 日期,A1  " & _
                 " From h_patient As A " & _
                 "Where A.A16 >='" & Format(dtStart, "yyyyMMdd") & "'" & _
                 " And A.A16 <='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                 IIf(lngSampleNO = -1, "", " And val(A.A2)=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        strID = "***"
        Do While Not rsTmp.EOF
            strID = "" & rsTmp("A1")
            strSQL = "Select C4,C6 From h_result As B Where B.C2='" & strID & "'"
            str项目结果 = ""
            Set rsErs = cnAccess.Execute(strSQL)
            Do Until rsErs.EOF
                If Trim("" & rsErs!C4) <> "" And Trim("" & rsErs!C6) <> "" Then
                    str项目结果 = str项目结果 & "|" & Trim("" & rsErs!C4) & "|" & Trim("" & rsErs!C6) '结果
                End If
                rsErs.MoveNext
            Loop
            
            str日期 = Trim("" & rsTmp("日期"))
            If Len(str日期) = 8 And IsNumeric(str日期) Then
                str日期 = Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2)
                If Not IsDate(str日期) Then str日期 = ""
            Else
                str日期 = ""
            End If
            lng标本号 = Val("" & rsTmp("标本号"))
            
            If str日期 <> "" And lng标本号 <> 0 And str项目结果 <> "" Then
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                aResults(UBound(aResults)) = str日期
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & lng标本号   '样本序号
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                         '检验人
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血清"                      '标本
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                              '质控
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & str项目结果
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


