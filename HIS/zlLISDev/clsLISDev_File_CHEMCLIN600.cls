VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_CHEMCLIN600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:Chemclin600
'厂商:北京科美生物
'作者:敬秦
'用户:川大望江校区校医院

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    On Error GoTo hErr '打开错误陷阱
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
        
    
    Dim OBJFILE As New FileSystemObject, objStream As textStream
    Dim strParentFile As String, STRFILENAME As String, strFilePath As String, strFileName1 As String, str_Segment As String
    Dim strResult() As String, intLop As Integer
    Dim intDate As Integer
    Dim var_Filed As Variant, var_Code As Variant, str_日期 As String, str_标本号 As String
    Dim varItem As Variant, lng_count As Integer, str_Return As String, strValue As String, strOD As String
    Dim intloop As Integer, str_Code As String, mstrCode As String, str项目 As String
    Dim i As Integer
    
    If Len(strFile) = 0 Then Exit Function
    ReDim Preserve strResult(0)
    strResult(0) = ""
    strFilePath = OBJFILE.GetParentFolderName(strFile)
    intDate = DateDiff("d", dtStart, dtEnd)
    For intLop = 0 To intDate
        STRFILENAME = Format(DateAdd("d", intLop, dtStart), "yyyymmdd")
        strFileName1 = Dir(strFilePath & "\" & STRFILENAME & "*.res")
        Do Until STRFILENAME = ""
'            strFileName1 = Dir(strFilePath & "\" & strFileName & "*.res")
            Set objStream = OBJFILE.OpenTextFile(strFilePath & "\" & strFileName1)
            Do Until objStream.AtEndOfStream
                str_Segment = objStream.ReadLine
                str_Segment = Trim(str_Segment)
                If InStr(str_Segment, "<MethodName>") > 0 Then
                    str项目 = Trim(Mid(str_Segment, InStr(str_Segment, ">") + 1, InStr(str_Segment, "</") - 1 - Len("<MethodName>")))
                    str_Return = "": str_标本号 = "": str_日期 = "": strValue = "": strOD = ""
                ElseIf InStr(str_Segment, "<report") > 0 Then
                    str_日期 = Mid(str_Segment, InStr(str_Segment, "=") + 1, 16)
                    str_日期 = Replace(str_日期, "-", "")
                    str_日期 = Replace(str_日期, """", "")
                    str_日期 = Format(Left(str_日期, 4) & "-" & Mid(str_日期, 5, 2) & "-" & Mid(str_日期, 7, 2) & " " & Mid(str_日期, 9, 2) & ":" & Mid(str_日期, 11, 2) & ":" & Right(str_日期, 2), "yyyy-mm-dd hh:mm:ss")
                    str_Return = "": strValue = "": strOD = ""
                ElseIf InStr(str_Segment, "<sample") > 0 Then
                    var_Filed = Split(" " & str_Segment, " ")
                    For intloop = LBound(var_Filed) To UBound(var_Filed)
                        If InStr(var_Filed(intloop), "id=") > 0 Then
                           str_标本号 = Trim(var_Filed(intloop))
                           str_标本号 = Trim(Replace(str_标本号, "id=", ""))
                           str_标本号 = Replace(str_标本号, """", "")
                        ElseIf InStr(var_Filed(intloop), "result=") > 0 Then
                           strValue = Trim(var_Filed(intloop))
                           strValue = Trim(Replace(strValue, "result=", ""))
                           strValue = Trim(Replace(strValue, "&#60;", "<"))
                           strValue = Replace(strValue, """", "")
                            If Split(strValue, ";")(0) = "&gt" Then
                                strValue = ">" & Split(strValue, ";")(1)
                            End If
                            If Split(strValue, ";")(0) = "&lt" Then
                                strValue = "<" & Split(strValue, ";")(1)
                            End If
                            strOD = ""
                        ElseIf InStr(var_Filed(intloop), "value=") > 0 Then
                            strOD = Trim(var_Filed(intloop))
                            strOD = Trim(Replace(strOD, "value=", ""))
                            str_Return = strValue '& "^" & strOD & "^0^0"
    '                        StrResult = StrResult & "||" & str_日期 & "|" & str_标本号 & "| |血清|0|" & str项目 & "|" & str_Return
                            If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                            strResult(UBound(strResult)) = str_日期 & "|" & str_标本号 & "| |血清|0|" & str项目 & "|" & str_Return
                        End If
                    Next
                End If
            Loop
            strFileName1 = Dir
            If strFileName1 = "" Then Exit Do
        Loop
    Next
    ResultFromFile = strResult

    '------------------------------------------------
    Exit Function
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_CHEMCLIN600"
    Close #lngErrFileNo
End Function
