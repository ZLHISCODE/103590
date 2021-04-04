VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_U2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:U2
'医院:扬州中医院
'开发:陈玉强
'时间:2009/6/17   修改时间 2013/12/19
'修改人：王振涛
'厂家:


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
    Dim cnsql As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset

    Dim strResult As String         '结果
    Dim strItem As String           '项目（通道码)
    Dim st日期 As String, en日期 As String, str日期 As String
    Dim strSer As String, strDB As String, strPWD As String, strUserID As String
    Dim strValues As String, varValues As Variant
    Dim strItemValue As String, i As Integer
    Dim str标本号 As String
    Dim str结果 As String
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults

    On Error GoTo errH
    strSer = Val(ReadIni("U2", "SERVER", App.Path & "\U2.ini"))
    If strSer = "" Then strSer = "127.0.0.1": Call WriteIni("U2", "SERVER", "127.0.0.1", App.Path & "\U2.ini")

    strDB = Val(ReadIni("U2", "DATABASE", App.Path & "\U2.ini"))
    If strDB = "" Then strDB = "urine_20130516": Call WriteIni("U2", "DATABASE", "urine_20130516", App.Path & "\U2.ini")

    strUserID = Val(ReadIni("U2", "USERID", App.Path & "\U2.ini"))
    If strUserID = "" Then strUserID = "sa": Call WriteIni("U2", "USERID", "sa", App.Path & "\U2.ini")

    strPWD = Val(ReadIni("U2", "PWD", App.Path & "\U2.ini"))
    If strPWD = "" Then strPWD = "saa": Call WriteIni("U2", "PWD", "saa", App.Path & "\U2.ini")


    With cnsql
        Set cnsql = New ADODB.Connection
        If cnsql.State = adStateOpen Then cnsql.Close

        cnsql.ConnectionTimeout = 25
        cnsql.Provider = "sqloledb"
        cnsql.Properties("data source").Value = strSer    '服务器的名字
        cnsql.Properties("initial catalog").Value = strDB    '库名
        'conn.Properties("integrated security").Value = "SSPI" '登陆类型
        cnsql.Properties("user id").Value = strUserID  '用户名
        cnsql.Properties("password").Value = strPWD     '密码
        cnsql.Open
        '    st日期 = Format(dtStart, "yyyy-mm-dd")
        '    en日期 = Format(dtEnd, "yyyy-mm-dd") & " 23:59:59"
        strSQL = "select smpno,daterpt,testvalues_text from urinerpts where daterpt between ' " & dtStart & "' and '" & dtEnd & "'" & _
                    IIf(lngSampleNO = -1, " ", " and smpno = " & lngSampleNO)
        Set rsTmp = cnsql.Execute(strSQL)
        Do Until rsTmp.EOF
            strItemValue = ""
            varValues = Split(rsTmp!testvalues_text, Chr(10))
            str标本号 = Val(rsTmp!smpno)
            str日期 = Format(rsTmp!daterpt, "YYYY-MM-DD")
            
            For i = LBound(varValues) To UBound(varValues)
                strItem = Fun_结果(varValues(i), "code")
                Select Case strItem
                Case "PRO", "BLD", "WBC", "NIT", "GLU", "KET", "URO", "BIL", "VC", "MA", "CR", "Ca"    '去掉数值报告
                    str结果 = Fun_结果(varValues(i), "value")
                    If InStr(str结果, "(") > 0 Then
                        str结果 = Mid(str结果, 1, InStr(str结果, "(") - 1)
                    End If
                    If Fun_结果(varValues(i), "pcode") = "micro" And strItem = "WBC" Then
                        strItem = strItem & 1
                    End If
                    If str结果 <> "" Then
                        strItemValue = strItemValue & "|" & strItem & "|" & str结果
                    End If
                Case Else
                    If Fun_结果(varValues(i), "pcode") <> "root" And strItem <> "" Then
                        If strItem = "micro" Then
                            MsgBox ""
                        End If
                        strItemValue = strItemValue & "|" & strItem & "|" & Fun_结果(varValues(i), "value")
                    End If
                End Select
            Next
            
            If strItemValue <> "" Then
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                strResult = str日期 & "|" & str标本号 & "| |尿液|0" & strItemValue
                aResults(UBound(aResults)) = strResult
            End If
            
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
    Exit Function
errH:
    WriteLog "U2", "ResultFromFile", Err.Number & "    " & Err.Description
End Function



Public Function Fun_结果(ByVal StrInput As String, ByVal strKey As String) As String
    Dim str结果 As String
    Dim strVar As Variant
    Dim i As Long
    Fun_结果 = ""
    On Error GoTo errH
    If StrInput = "" Then Exit Function
    StrInput = Replace(Replace(Replace(StrInput, Chr(10), ""), Chr(13), ""), "<p", "p")
    strVar = Split(StrInput, ";")
    For i = LBound(strVar) To UBound(strVar)
        If Split(Replace(strVar(i), vbCrLf, ""), "=")(0) = strKey Then
            Fun_结果 = Trim(Split(strVar(i), "=")(1))
            Exit Function
        End If
    Next i
    Exit Function
errH:
    WriteLog "U2", "Fun_结果", Err.Number & "    " & Err.Description
End Function








