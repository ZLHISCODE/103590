VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AX4030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器厂家：西森美康
'型号：sysmex UF-500i 集成aution max AX-4030
'类型：尿机、生化
'用户：眉山市人民医院
'作者：敬秦
'时间：2015/4/2

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"), Optional ByVal lngendSampleNO As Long) As String()
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
    Dim objFile As New FileSystemObject, txtStream As textStream
    Dim strLine As String, varLine() As String, strResult() As String
    Dim intloop As Long, i As Long, j As Long, intDate As Integer
    Dim strSample As String, strDate As String, strQC As String, strChenal As String, strValue As String, strItem As String
    Dim strFileName As String, strFilePath As String
    Dim strImage As String, strImageName As String

    On Error GoTo hErr
    If strFile = "" Then Exit Function
    ReDim strResult(0)
    strResult(0) = ""
    strFilePath = objFile.GetParentFolderName(strFile)
    intDate = DateDiff("d", Format(dtStart, "yyyy-mm-dd"), Format(dtEnd, "yyyy-mm-dd"))    '循环取出日期内的数据
    For intloop = 0 To intDate
        strFileName = Dir(strFilePath & "\" & Format(dtStart + intloop, "yyyyMdd") & "*.sdf")
        Do Until strFileName = ""
            Set txtStream = objFile.OpenTextFile(strFilePath & "\" & strFileName)
            Do Until txtStream.AtEndOfStream
                strLine = txtStream.ReadLine
                varLine = Split(strLine, ",")
                If UBound(varLine) >= 8 Then
                    If varLine(0) = "0" Then    '取时间，标本号
                        strSample = varLine(3)
                        strDate = Format(varLine(1) & " " & varLine(7), "yyyy-mm-dd hh:mm:ss")
                    End If
                    If varLine(0) = "1" Then    '取仪器检验数据
                        strChenal = varLine(1)
                        strValue = IIf(varLine(3) = "", " ", varLine(3))
                        strItem = strItem & "|" & strChenal & "|" & strValue
                    End If
                ElseIf UBound(varLine) >= 4 And varLine(0) = "3" Then    '取仪器检验图像数据
                    strImageName = varLine(2)
                    strImage = strImage & "^" & strImageName & ";120;1;" & varLine(3)
                End If
            Loop
            If strSample <> "" And strItem <> "" Then    '一个标本循环完成
                If strResult(UBound(strResult)) <> "" Then ReDim Preserve strResult(UBound(strResult) + 1)
                strResult(UBound(strResult)) = strDate & "|" & strSample & "| |尿液|0" & strItem
                strItem = ""
                If strImage <> "" Then
                    strResult(UBound(strResult)) = strResult(UBound(strResult)) & vbCrLf & Mid$(strImage, 2)
                    strImage = ""
                End If
            End If
            strFileName = Dir
        Loop
    Next
    ResultFromFile = strResult
    Exit Function
hErr:
    MsgBox "错误描述：" & Err.Description, vbInformation, "中联接口"
    Err.Clear
End Function


