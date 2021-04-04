VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_URIT660"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'医院名称：彭州市精神病医院
'仪器类型：酶标仪
'仪器型号：URIT-660
'仪器厂家：桂林优利特
'作    者：陈德荣
'日    期：2012-06-01
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
    Dim lngResultCount As Long, strBh As String, ingCount As Long, ingBh As Long
    Dim blnUnion As Boolean
    Dim strResult As String, strTmp As String
    Dim strOD As String
    Dim strCutOff As String
    Dim strCOV As String
    
    On Error GoTo errHandle
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    With cnAccess
        If .State = adStateOpen Then .Close
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        
        If .State = adStateClosed Then
            'MsgBox "打开数据库连接失败,请检查数据源是否正确配置！"
            Exit Function
        End If
        
        strSQL = "": strResult = "": strTmp = ""
        strSQL = "Select * From (SELECT Sample.PlateID,Sample.SampleNO, Format(Plate.TestDate+Plate.TestTime,'yyyy-mm-dd hh:mm:ss') AS Drq," & _
            "Item.Code AS ItemCode,Format(CDbl(Sample.Result),'###0.0###') As OD,Sample.Label AS Result,Sample.WellTypeID " & _
            "FROM ((Sample LEFT JOIN Plate ON Sample.PlateID=Plate.PlateID) LEFT JOIN Item ON Sample.ItemID=Item.ID) LEFT JOIN WellType ON Sample.WellTypeID=WellType.WellTypeID) A " & _
            "Where WellTypeID=2 And Drq >='" & Format(dtStart, "YYYY-MM-DD 00:00:00") & "'" & " And Drq <='" & Format(dtEnd, "YYYY-MM-DD 23:59:59") & "'" & _
            IIf(lngSampleNO = -1, "", " And val(PlateID)=" & lngSampleNO) & "  Order by SampleNO"
        
        Open "C:\OutSql.txt" For Output As #1
        Print #1, strSQL
        Close #1
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            lngResultCount = 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = ""
            ResultFromFile = aResults
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        
        ingCount = 0
                  
        Do While Not rsTmp.EOF
            ingBh = Val(Mid("" & rsTmp("SampleNO"), 7))
            
            If strBh <> Str(ingBh) Then '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(IIf(IsNull(rsTmp("Drq")), Now, rsTmp("Drq")), "yyyy-mm-dd hh:mm:ss") '检验时间
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Val(ingBh)  '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "| "     '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|血清"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0"     '质控
                strBh = Mid(Trim("" & rsTmp("SampleNO")), 7)
            End If
                
            If rsTmp("OD") > 0 Then
                strOD = IIf(Trim(rsTmp("OD")) = "", "0.000", Format(rsTmp("OD"), "###0.0###"))
                strCutOff = "0.000"
                strCOV = "0.000"
            Else
                strOD = "0.000"
                strCutOff = "0.000"
                strCOV = "0.000"
            End If
            strTmp = Trim("" & rsTmp("Result"))
            Select Case strTmp
                Case "1"
                  strResult = "阳性(+)"
                Case "2"
                  strResult = "阴性(-)"
            End Select
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                 rsTmp("ItemCode") & "|" & strResult & "^" & strOD & "^" & strCutOff & "^" & strCOV '结果
            rsTmp.MoveNext
        Loop
    End With
    For ingCount = LBound(aResults) To UBound(aResults)
      Debug.Print aResults(ingCount)
    Next
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_URIT660.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function

