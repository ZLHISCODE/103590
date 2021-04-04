VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_gc5600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
    Dim lngResultCount As Long, strBh As String, strDate As String
    Dim strChenel As String, strData As String
    Dim blnUnion As Boolean
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "": blnUnion = False
            strSQL = "Select shearrate,序号,日期,时间,visc,血浆粘度,压积,血沉,全血低切相对指数,全血高切相对指数,血沉方程K值,红细胞聚集指数,红细胞变形指数," & _
                     " 全血低切还原粘度,全血高切还原粘度,红细胞变形指数TK,红细胞刚性指数,卡松粘度,纤维蛋白原,血胆固醇,甘油三脂,高密脂蛋白,血糖," & _
                    " 血小板粘附率,体外血栓干重,红细胞电泳,血小板聚集率,体外血栓长度 " & _
                        "FROM testdata  b,visc a  where  a.testdataid=b.testdataid " & _
                     IIf(lngSampleNO = -1, " And Format(b.日期,'yyyy-mm-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "'" & _
                    " And Format(日期,'yyyy-mm-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & "'", _
                    " And B.序号=" & lngSampleNO) & _
                    " order by 日期 desc,序号"

        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            '避免出现下标错误
            ReDim aResults(0)
            aResults(0) = ""
            ResultFromFile = aResults
            Exit Function
        End If
        
        lngResultCount = -1
        Do While Not rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            strDate = Format(rsTmp("日期"), "yyyy-mm-dd")
            strDate = strDate & " " & rsTmp("时间")
            aResults(lngResultCount) = strDate '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val(rsTmp("序号")) '样本序号
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                " " '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                "血液" '标本
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                "0" '质控
                
            Select Case rsTmp("shearrate")
            Case 1
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.1|" & rsTmp("visc")
                If Not IsNull(rsTmp("血浆粘度")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|02|" & rsTmp("血浆粘度")
                If Not IsNull(rsTmp("压积")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|03|" & rsTmp("压积")
                If Not IsNull(rsTmp("血沉")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|04|" & rsTmp("血沉")
                If Not IsNull(rsTmp("全血低切相对指数")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|05|" & rsTmp("全血低切相对指数")
                If Not IsNull(rsTmp("全血高切相对指数")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|06|" & rsTmp("全血高切相对指数")
                If Not IsNull(rsTmp("血沉方程K值")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|07|" & rsTmp("血沉方程K值")
                If Not IsNull(rsTmp("红细胞聚集指数")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|08|" & rsTmp("红细胞聚集指数")
                If Not IsNull(rsTmp("红细胞变形指数")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|09|" & rsTmp("红细胞变形指数")
                If Not IsNull(rsTmp("全血低切还原粘度")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|10|" & rsTmp("全血低切还原粘度")
                If Not IsNull(rsTmp("全血高切还原粘度")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|11|" & rsTmp("全血高切还原粘度")
                If Not IsNull(rsTmp("红细胞变形指数TK")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|12|" & rsTmp("红细胞变形指数TK")
                If Not IsNull(rsTmp("红细胞刚性指数")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|13|" & rsTmp("红细胞刚性指数")
                If Not IsNull(rsTmp("卡松粘度")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|14|" & rsTmp("卡松粘度")
                If Not IsNull(rsTmp("纤维蛋白原")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|15|" & rsTmp("纤维蛋白原")
                If Not IsNull(rsTmp("血胆固醇")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|16|" & rsTmp("血胆固醇")
                If Not IsNull(rsTmp("甘油三脂")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|17|" & rsTmp("甘油三脂")
                If Not IsNull(rsTmp("高密脂蛋白")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|18|" & rsTmp("高密脂蛋白")
                If Not IsNull(rsTmp("血糖")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|19|" & rsTmp("血糖")
                If Not IsNull(rsTmp("血小板粘附率")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|20|" & rsTmp("血小板粘附率")
                If Not IsNull(rsTmp("体外血栓干重")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|21|" & rsTmp("体外血栓干重")
                If Not IsNull(rsTmp("红细胞电泳")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & rsTmp("红细胞电泳")
                If Not IsNull(rsTmp("血小板聚集率")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|23|" & rsTmp("血小板聚集率")
                If Not IsNull(rsTmp("体外血栓长度")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|24|" & rsTmp("体外血栓长度")
            Case 5
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.2|" & rsTmp("visc")
            Case 30
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.3|" & rsTmp("visc")
            Case 200
                If Not IsNull(rsTmp("visc")) Then aResults(lngResultCount) = aResults(lngResultCount) & "|01.4|" & rsTmp("visc")
            End Select
           
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function
