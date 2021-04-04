VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MB_80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'用户：钦州市第一人民医院
'仪器：MB-80
'mdb读取方式
'作者：

Option Explicit

Public Function GetStartcmd() As String
'第一次启动时调用
End Function

Public Function GetEndCmd() As String
'接口关闭的时候调用
End Function

Public Function GetAnswerCmd() As String
'定时发送的指令
End Function

Public Sub Analyse(ByVal strSource As String, _
                   strResult As String, _
                   strReserved As String, _
                   strCmd As String)
'解码函数
End Sub

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
    Dim ConnSQL As New ADODB.Connection
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim lngResultCount As Long, strBh As String
    Dim str日期 As String
    Dim str_Tmp日期 As String
    Dim str_Tmp日期2 As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strFilePath As String
    Dim strTDM As String
    
    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)

    If Len(strFilePath) = 0 Then strFilePath = App.Path

    With ConnSQL
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = ""
        strSQL = "select 样品号 as 标本号, Mid(检验日期,1,6) as 日期1, 检验项目 as 通道码, 检验结果 as 结果 from eds a  " & _
                 "where " & _
                 "Mid(检验日期,1,6) " & ">= " & Format(dtStart, "YYMMDD") & "" & _
               " and Mid(检验日期,1,6)" & "<= " & Format(dtEnd, "YYMMDD") & IIf(lngSampleNO = -1, " ", " and 样品号 = " & lngSampleNO) & " order by 样品号, 检验日期  "


        Set rsSQL = ConnSQL.Execute(strSQL)

        If rsSQL.EOF Then
            .Close
            ReDim Preserve aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If

        '                If rsTmp.EOF Then
        '            .Close
        '            Exit Function
        '        End If

        lngResultCount = -1
        strBh = "***"
        Do While Not rsSQL.EOF
            If strBh <> rsSQL("标本号") Then    '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)

                str日期 = rsSQL("日期1")
                str日期 = Mid(str日期, 1, 2) & "-" & Mid(str日期, 3, 2) & "-" & Mid(str日期, 5, 2)  '拆分日期出来“yyyy-mm-dd”
                aResults(lngResultCount) = Format(str日期, "yyyy-mm-dd")   '检验时间

                'str日期 = Format(rsSQL("ResultTime"), "yyyy-mm-dd")

                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Val(rsSQL("标本号"))  '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                         " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           "血液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"    '& _
                                                                              ' Val(rsSQL("质控")) '质控
                strBh = rsSQL("标本号")
            End If

            strTDM = rsSQL("通道码")
            strTDM = Replace(strTDM, "G-脂多糖", "G1")
            strTDM = Replace(strTDM, "1-3-β-D葡聚糖", "13BD")

            If strTDM <> "" Or IsNull(strTDM) = False Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                                           Trim(strTDM) & "|" & IIf(IsNull(rsSQL("结果")), "*", rsSQL("结果"))    '结果
            End If

            rsSQL.MoveNext

        Loop

    End With

    If rsSQL.EOF Then
        ConnSQL.Close
    End If
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "MB_80", "MB_80", Err.Number & " " & Err.Description
End Function

Public Function SendSample(strSampleInfo As String, _
                           iSendStep As Integer, _
                           blnSuccess As Boolean, _
                           Optional ByVal strResponse As String = "", _
                           Optional ByVal blnUndo As Boolean = False, _
                           Optional ByVal iType As Integer = 0) As String
'解码双向指令的
End Function








