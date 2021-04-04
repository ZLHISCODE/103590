VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MH_Micro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:美华微生物
'厂商:美华
'作者:lwf
'用户:周口永善妇科医院

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
    Dim strsql As String, rsTmp As New ADODB.Recordset
    Dim str标本号 As String, str项目 As String
    Dim intLoop As Integer
    Dim intItem As Integer
    Dim aLine() As String
    Dim strItem As String
    Dim str_bz As String
    
    
    On Error GoTo errH
    
    Dim rsErs As ADODB.Recordset
    Dim str日期 As String
    Dim strData As String

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .CursorLocation = adUseClient
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb, *.accdb)}"
        .Open
        
        If .State = adStateClosed Then Exit Function
        
       strsql = "select a.检验时间,b.标本编号,b.细菌编号,c.药物编号,c.组别,c.MIC,c.结果,c.药物类别" & _
                " from 病人信息 a,细菌结果 b,药敏结果 c where a.标本编号 = b.标本编号 and b.标本编号=c.标本编号" & _
                " and Format(a.检验时间, 'yyyy/MM/dd HH:mm:ss') >= '" & Format(dtStart, "yyyy/MM/dd HH:mm:ss") & "'" & vbNewLine & _
                " and Format(a.检验时间, 'yyyy/MM/dd HH:mm:ss') < '" & Format(dtEnd + 1, "yyyy/MM/dd HH:mm:ss") & "'"

        Set rsTmp = cnAccess.Execute(strsql)
        
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        
        Do While Not rsTmp.EOF
                
            str日期 = Format("" & rsTmp("检验时间"), "yyyy-MM-dd HH:mm:ss")
            str标本号 = Val("" & rsTmp("标本编号"))

            str_bz = ""
            
            If rsTmp("细菌编号") <> "" Then
'                If rsTmp("备注" & intLoop) <> "" Then str_bz = "|BZ|R^1^" & rsTmp("备注" & intLoop)
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    
                aResults(UBound(aResults)) = str日期 & "|" & str标本号 & "| |" & _
                    rsTmp("细菌编号") & "|0"
                If rsTmp("药物编号") <> "" Then
                    strItem = strItem & "|" & rsTmp("药物编号") & "|"
                    Select Case Nvl(rsTmp("结果"))
                        Case "耐药"
                            strItem = strItem & "R"
                        Case "中介"
                            strItem = strItem & "I"
                        Case "敏感"
                            strItem = strItem & "S"
                        Case Else
                            strItem = strItem & ""
                    End Select
                    strItem = strItem & "^1^" & rsTmp("MIC")
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & strItem
                    strItem = ""
'                Else
'                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "||R^1^" & str_bz
                End If
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
    For intLoop = 0 To UBound(aResults)
        Debug.Print aResults(intLoop)
    Next
    
    Exit Function
errH:
    
    MsgBox Err.Description
End Function



'0769-26208501

