VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AYW9003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:全自动凝血分析仪 AYW9003
'厂商:南京瑞麦科技开发有限公司
'作者:王振涛
'用户:运城市眼科医院
'日期:2012-12-18

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBH As String, strAnti As String, str结果 As String
    Dim lngResultCount As Long, str_项目 As String

    On Error GoTo errH
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    cnAccess.ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};pwd=508"
    cnAccess.Open
    If cnAccess.State = adStateClosed Then Exit Function

    strSQL = "Select A.biaobenhao as 标本号,A.jianyanriqi as 日期,A.jianyanshijian as 时间, A.pt as PT, A.ptr AS ptr ,A.inr as inr,A.aptt as aptt,A.tt as TT,A.pt_fib AS FIB "
    strSQL = strSQL & " From bingrenziliao  A "

    strSQL = strSQL & " Where  A.jianyanriqi >=Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
    strSQL = strSQL & " And  A.jianyanriqi < Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "') " & _
             IIf(lngSampleNo = -1, "", " And val(A.biaobenhao)=" & lngSampleNo) & " order by 1"
    WriteLog "AYW9003", strSQL, ""

    strBH = "***"
    ReDim aResults(0)

    Set rsTmp = cnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        cnAccess.Close
        WriteLog "AYW9003", "无记录！", ""
        Exit Function
    End If
    Do Until rsTmp.EOF
        
        If strBH <> rsTmp("标本号") Then  '新标本
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)

            aResults(UBound(aResults)) = Format(Trim("" & rsTmp("日期")) & " " & Trim("" & rsTmp("时间")), "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         Val(rsTmp("标本号"))    '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & _
                                         "血液"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            strBH = rsTmp("标本号")
        End If

        If strBH = rsTmp("标本号") Then
            If Trim("" & rsTmp("pt")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "PT|" & Replace(Replace(Trim("" & rsTmp("pt")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|PT|" & Replace(Replace(Trim("" & rsTmp("pt")), "↑", ""), "↓", "")
                End If
            End If
            If Trim("" & rsTmp("ptr")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "PTR|" & Replace(Replace(Trim("" & rsTmp("ptr")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|PTR|" & Replace(Replace(Trim("" & rsTmp("ptr")), "↑", ""), "↓", "")
                End If
            End If
            If Trim("" & rsTmp("inr")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "INR|" & Replace(Replace(Trim("" & rsTmp("inr")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|INR|" & Replace(Replace(Trim("" & rsTmp("inr")), "↑", ""), "↓", "")
                End If
            End If
            If Trim("" & rsTmp("aptt")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "APTT|" & Replace(Replace(Trim("" & rsTmp("aptt")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|APTT|" & Replace(Replace(Trim("" & rsTmp("aptt")), "↑", ""), "↓", "")
                End If
            End If
            If Trim("" & rsTmp("tt")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "TT|" & Replace(Replace(Trim("" & rsTmp("tt")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|TT|" & Replace(Replace(Trim("" & rsTmp("tt")), "↑", ""), "↓", "")
                End If
            End If
            If Trim("" & rsTmp("FIB")) <> "" Then
                If str_项目 = "" Then
                    str_项目 = "FIB|" & Replace(Replace(Trim("" & rsTmp("fib")), "↑", ""), "↓", "")
                Else
                    str_项目 = str_项目 & "|FIB|" & Replace(Replace(Trim("" & rsTmp("fib")), "↑", ""), "↓", "")
                End If
            End If
        End If
        If str_项目 <> "" Then
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & str_项目
            str_项目 = ""
        End If

        WriteLog "AYW9003", aResults(UBound(aResults)), ""
        rsTmp.MoveNext
    Loop

    ResultFromFile = aResults
    Exit Function
errH:
    WriteLog "AYW9003", CStr(Erl()) & ":" & Err.Description, ""
End Function




