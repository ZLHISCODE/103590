VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MDK_3200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MDK-3200 血流变
'厂商:重庆麦迪克
'用户:市三院

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
    Dim strsql As String, rsTmp As ADODB.Recordset, rs项目 As ADODB.Recordset, rs时间 As ADODB.Recordset
    Dim lngResultCount As Long, strcutoff As String, str项目 As String
    Dim str日期 As String, str定性结果 As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFile = strFilePath & "\EXPORT.MDB"
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        '查找检验批次
        strsql = "select SessionID as 序列,SessionDate as 日期 from Session" & _
                " where Format(SessionDate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(SessionDate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' "
        Set rs时间 = cnAccess.Execute(strsql)
        If rs时间.EOF Then
           .Close
           Exit Function
        End If
        lngResultCount = -1
        Do While Not rs时间.EOF
        
       '查找test表中的项目 od cutoff值
           strsql = "select SessionID as 序列,TestIndex as 序号,Name as 项目,CutOff ,DcmScad as 日期" & _
                     " from Test " & _
                     "where  Sessionid=" & rs时间("序列")
            Set rsTmp = cnAccess.Execute(strsql)
            If rsTmp.EOF Then
                .Close
                Exit Function
            End If
            
            Do While Not rsTmp.EOF
            '提取项目
            str项目 = rsTmp("项目")
            '提取检验日期
            str日期 = Format(rs时间("日期"), "yyyy-mm-dd hh:mm:ss")
            '提取本项目的cutoff值
            strcutoff = "" & rsTmp("cutoff")
            '准备开始提取结果
              strsql = "select Type as 顺序号,OD,Result as 结果" & _
                       " from Pozzetti" & _
                       " where  SessionID=" & rsTmp("序列") & " and TestIndex=" & rsTmp("序号")
              Set rs项目 = cnAccess.Execute(strsql)
                  Do While Not rs项目.EOF
                     If IsNumeric(rs项目("顺序号")) Then
                        lngResultCount = lngResultCount + 1
                        ReDim Preserve aResults(lngResultCount)
                        If str项目 = "HBeAB" Or str项目 = "HBcAB" Then
                           If rs项目("OD") < strcutoff Then
                              str定性结果 = "阳性(+)"
                           Else
                            str定性结果 = "阴性(-)"
                           End If
                         Else
                           If rs项目("OD") > strcutoff Then
                              str定性结果 = "阳性(+)"
                           Else
                            str定性结果 = "阴性(-)"
                           End If
                         End If
                        aResults(lngResultCount) = str日期  '检验时间
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rs项目("顺序号"))   '样本序号
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                           " "  '检验人
                        aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          "血液"  '标本
                       aResults(lngResultCount) = aResults(lngResultCount) & "|0" '质控
                       aResults(lngResultCount) = aResults(lngResultCount) & "|" & str项目 & "|" & str定性结果 '结果
                       aResults(lngResultCount) = aResults(lngResultCount) & _
                                     "^" & Val("" & rs项目("结果")) & "^" & strcutoff & "^" & Val(rs项目("OD")) 'cutoff od值
    
                     End If
                     rs项目.MoveNext
                  Loop
              rsTmp.MoveNext
            Loop
            rs时间.MoveNext
      Loop
    End With
    ResultFromFile = aResults
End Function








