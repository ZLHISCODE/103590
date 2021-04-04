VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_FIA8100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'用户：吴川人民医院
'作者：欧海胜
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
    Dim str标本号 As String, str项目 As String
    Dim str日期 As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
   
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function   'blood_para_table
        strsql = "SELECT data as 日期, nCa as 通道码,PCT as PCT结果,result as CRP结果,blood_para_table.[Ca++] as C反应结果,blood_para_table.[No] as 标本号" & _
                " FROM blood_para_table Where data>= '" & Format(dtStart, "yyyy.MM.dd") & _
                "' And data<= '" & Format(DateAdd("d", 1, dtEnd), "yyyy.MM.dd") & "' " & _
                IIf(lngSampleNO = -1, "", " And No=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        Do While Not rsTmp.EOF
                str日期 = Replace(Format("" & rsTmp("日期"), "yyyy-MM-dd hh24:mi:ss"), ".", "-")
                str标本号 = Val(rsTmp("标本号"))
                str项目 = ""
                
                    If rsTmp("通道码") = "PCT" Then
                      str项目 = "|" & rsTmp("通道码") & "|" & rsTmp("PCT结果")
                    
                    ElseIf rsTmp("通道码") = "U-CRP" And rsTmp("CRP结果") <> "" And rsTmp("C反应结果") <> "" Then
                      str项目 = "|" & rsTmp("通道码") & "|" & rsTmp("CRP结果") & "|" & "CRP" & "|" & rsTmp("C反应结果")
                    ElseIf rsTmp("通道码") = "U-CRP" And rsTmp("CRP结果") = "" And rsTmp("C反应结果") <> "" Then
                      str项目 = "|" & "CRP" & "|" & rsTmp("C反应结果")
                    ElseIf rsTmp("通道码") = "U-CRP" And rsTmp("CRP结果") <> "" And rsTmp("C反应结果") = "" Then
                      str项目 = "|" & rsTmp("通道码") & "|" & rsTmp("CRP结果")
                    End If
               
                If str项目 <> "" Then
                    If aResults(UBound(aResults)) <> "" Then
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||" & str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str标本号                       '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '检验人
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"                          '标本
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '质控
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str项目                               '项目和结果
                    
                    ElseIf aResults(UBound(aResults)) = "" Then
                    
                    aResults(UBound(aResults)) = str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str标本号                       '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "                             '检验人
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"                          '标本
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"                                  '质控
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str项目
                    End If
                    aResults(UBound(aResults)) = Replace(aResults(UBound(aResults)), "↑", "")
                End If
                rsTmp.MoveNext
            Loop
  
        
    End With
    ResultFromFile = aResults
End Function





