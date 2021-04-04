VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_US2020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:US2020
'厂商:重庆
'作者:曾超
'用户:长重一汽总医院
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
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intloop As Integer
    Dim strimg As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    ReDim aResults(0)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)};uid=;pwd="
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "Select * from PatientInfo A ,hua b "
        strSQL = strSQL & "  Where a.lngnumber = b.lngnumber and A.strReportDate >= Cdate('" & Format(dtStart, "yyyy-MM-dd") & "')"
        strSQL = strSQL & " And A.strReportDate< Cdate('" & Format(dtEnd + 1, "yyyy-MM-dd") & "')"
'        IIf(lngSampleNo = -1, "", " And val(A.Pat_Id)=" & lngSampleNo)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = Format(rsTmp("strReportDate"), "yyyy-mm-dd hh:mm:ss") '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(rsTmp("lngNumber"))  '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血清"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_0|" & IIf(Trim("" & rsTmp!strTP_0) = "", " ", Trim("" & rsTmp!strTP_0))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_1|" & IIf(Trim("" & rsTmp!strTP_1) = "", " ", Trim("" & rsTmp!strTP_1))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_2|" & IIf(Trim("" & rsTmp!strTP_2) = "", " ", Trim("" & rsTmp!strTP_2))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_3|" & IIf(Trim("" & rsTmp!strTP_3) = "", " ", Trim("" & rsTmp!strTP_3))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_4|" & IIf(Trim("" & rsTmp!strTP_4) = "", " ", Trim("" & rsTmp!strTP_4))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_5|" & IIf(Trim("" & rsTmp!strTP_5) = "", " ", Trim("" & rsTmp!strTP_5))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_6|" & IIf(Trim("" & rsTmp!strTP_6) = "", " ", Trim("" & rsTmp!strTP_6))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_7|" & IIf(Trim("" & rsTmp!strTP_7) = "", " ", Trim("" & rsTmp!strTP_7))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_8|" & IIf(Trim("" & rsTmp!strTP_8) = "", " ", Trim("" & rsTmp!strTP_8))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_9|" & IIf(Trim("" & rsTmp!strTP_9) = "", " ", Trim("" & rsTmp!strTP_9))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_10|" & IIf(Trim("" & rsTmp!strTP_10) = "", " ", Trim("" & rsTmp!strTP_10))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_11|" & IIf(Trim("" & rsTmp!strTP_11) = "", " ", Trim("" & rsTmp!strTP_11))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_12|" & IIf(Trim("" & rsTmp!strTP_12) = "", " ", Trim("" & rsTmp!strTP_12))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_13|" & IIf(Trim("" & rsTmp!strTP_13) = "", " ", Trim("" & rsTmp!strTP_13))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_14|" & IIf(Trim("" & rsTmp!strTP_14) = "", " ", Trim("" & rsTmp!strTP_14))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_15|" & IIf(Trim("" & rsTmp!strTP_15) = "", " ", Trim("" & rsTmp!strTP_15))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_16|" & IIf(Trim("" & rsTmp!strTP_16) = "", " ", Trim("" & rsTmp!strTP_16))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_17|" & IIf(Trim("" & rsTmp!strTP_17) = "", " ", Trim("" & rsTmp!strTP_17))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_18|" & IIf(Trim("" & rsTmp!strTP_18) = "", " ", Trim("" & rsTmp!strTP_18))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_19|" & IIf(Trim("" & rsTmp!strTP_19) = "", " ", Trim("" & rsTmp!strTP_19))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_20|" & IIf(Trim("" & rsTmp!strTP_20) = "", " ", Trim("" & rsTmp!strTP_20))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_21|" & IIf(Trim("" & rsTmp!strTP_21) = "", " ", Trim("" & rsTmp!strTP_21))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_22|" & IIf(Trim("" & rsTmp!strTP_22) = "", " ", Trim("" & rsTmp!strTP_22))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_23|" & IIf(Trim("" & rsTmp!strTP_23) = "", " ", Trim("" & rsTmp!strTP_23))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_24|" & IIf(Trim("" & rsTmp!strTP_24) = "", " ", Trim("" & rsTmp!strTP_24))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_25|" & IIf(Trim("" & rsTmp!strTP_25) = "", " ", Trim("" & rsTmp!strTP_25))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_26|" & IIf(Trim("" & rsTmp!strTP_26) = "", " ", Trim("" & rsTmp!strTP_26))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_27|" & IIf(Trim("" & rsTmp!strTP_27) = "", " ", Trim("" & rsTmp!strTP_27))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_28|" & IIf(Trim("" & rsTmp!strTP_28) = "", " ", Trim("" & rsTmp!strTP_28))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strTP_29|" & IIf(Trim("" & rsTmp!strTP_29) = "", " ", Trim("" & rsTmp!strTP_29))
            
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|stryan|" & IIf(Trim("" & rsTmp!stryan) = "", " ", Trim("" & rsTmp!stryan))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strtou|" & IIf(Trim("" & rsTmp!strtou) = "", " ", Trim("" & rsTmp!strtou))
            
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strseq|" & IIf(Trim("" & rsTmp!strseq) = "", " ", Trim("" & rsTmp!strseq))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|struro|" & IIf(Trim("" & rsTmp!struro) = "", " ", Trim("" & rsTmp!struro))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strbil|" & IIf(Trim("" & rsTmp!strbil) = "", " ", Trim("" & rsTmp!strbil))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strket|" & IIf(Trim("" & rsTmp!strket) = "", " ", Trim("" & rsTmp!strket))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strbld|" & IIf(Trim("" & rsTmp!strbld) = "", " ", Trim("" & rsTmp!strbld))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strpro|" & IIf(Trim("" & rsTmp!strpro) = "", " ", Trim("" & rsTmp!strpro))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strnit|" & IIf(Trim("" & rsTmp!strnit) = "", " ", Trim("" & rsTmp!strnit))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strleu|" & IIf(Trim("" & rsTmp!strleu) = "", " ", Trim("" & rsTmp!strleu))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strglu|" & IIf(Trim("" & rsTmp!strglu) = "", " ", Trim("" & rsTmp!strglu))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strsg|" & IIf(Trim("" & rsTmp!strsg) = "", " ", Trim("" & rsTmp!strsg))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strph|" & IIf(Trim("" & rsTmp!strph) = "", " ", Trim("" & rsTmp!strph))
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|strvc|" & IIf(Trim("" & rsTmp!strvc) = "", " ", Trim("" & rsTmp!strvc))
            
            
            '去掉单位
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("umol/L"), "")
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("Normal"), "")
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("mmol/L"), "")
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("Cell/uL"), "")
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("Ery/uL"), "")
            aResults(UBound(aResults)) = Replace$(UCase(aResults(UBound(aResults))), UCase("g/L"), "")
            
            If Dir(objFile.GetParentFolderName(strFile) & "PIC\" & Val(rsTmp("lngNumber")) & "_0.jpg") <> "" Then
                strimg = strimg & "^" & Val(rsTmp("lngNumber")) & "0;110;1;" & objFile.GetParentFolderName(strFile) & "PIC\" & Val(rsTmp("lngNumber")) & "_0.jpg"
            End If
            If Dir(objFile.GetParentFolderName(strFile) & "PIC\" & Val(rsTmp("lngNumber")) & "_1.jpg") <> "" Then
                strimg = strimg & "^" & Val(rsTmp("lngNumber")) & "1;110;1;" & objFile.GetParentFolderName(strFile) & "PIC\" & Val(rsTmp("lngNumber")) & "_1.jpg"
            End If
            
            If strimg <> "" Then
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strimg, 2)
            End If
            strimg = ""
            rsTmp.MoveNext
        Loop
    End With
    
    For intloop = 0 To UBound(aResults)
        Debug.Print aResults(intloop)
    Next
    ResultFromFile = aResults
   
End Function



