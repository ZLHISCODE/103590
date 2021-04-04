VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_R80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血粘度分析仪 LG-R-80F
'厂商:
'作者:陈东
'用户:洛阳202医院
'修改：王振涛 ---- 图形数据
'用户：武乡县医院

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
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean, strImg As String, iField As Integer, i As Integer
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String
    Dim strChar As String
    Dim intMaxValue As Integer
    Dim str_仪器类型 As String
    
    str_仪器类型 = Trim(ReadIni("仪器类型", "类型", App.Path & "\LG-R-80F.ini"))
    If str_仪器类型 = "" Then
        WriteIni "仪器类型", "类型", "LG-R-80F", App.Path & "\LG-R-80F.ini"
        str_仪器类型 = "LG-R-80F"
    End If
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
        
    If Dir(strFile) = "" Then Exit Function
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        strsql = "SELECT mid(A.recordnum,1,8) as 日期, mid(A.recordnum,9) as 标本号,sex as 性别, A.f1, A.f2, A.f3, A.f4, A.f5, A.f6, A.f7, A.T1, A.T2, A.T3, A.T4, A.T5, A.T6, A.T7, A.T8, A.C2, A.C3, A.C4, A.C5, Result " & _
                "FROM patient AS A where mid(A.recordnum,1,8)>='" & Format(dtStart, "yyyyMMdd") & "' And mid(A.recordnum,1,8)<='" & Format(dtEnd, "yyyyMMdd") & "'" & _
                IIf(lngSampleNO = -1, "", " And val(Mid(A.recordnum,9))=" & lngSampleNO)

        Set rsTmp = cnAccess.Execute(strsql)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("标本号") Then '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                str日期 = Trim("" & rsTmp("日期"))
                
                aResults(lngResultCount) = Format(Mid(str日期, 1, 4) & "-" & Mid(str日期, 5, 2) & "-" & Mid(str日期, 7, 2), "yyyy-MM-dd") '检验时间
                                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("标本号"))   '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "血液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("质控")) '质控
                    
                strBh = rsTmp("标本号")
            End If
            
            For iField = 3 To 21
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    iField - 2 & "|" & Val(IIf(IsNull(rsTmp.Fields(iField).Value), "0", rsTmp.Fields(iField).Value)) '结果
            Next
            If Trim("" & rsTmp!result) <> "" Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & Trim("" & rsTmp!result)
            End If
            '-- 血流变的图形数据
            '-- 图形数据X-横，Y－纵
            '-- 格式：类型;座标数据;曲线及描点数据
            '        其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
            '              曲线及描点数据:曲线1的Y坐标1,曲线1的Y座标2.....|曲线2的Y坐标1,曲线2的Y坐标2,...~描点的Y从标1,描点的Y从标2
            If str_仪器类型 = "LG-R-80F" Then
                strImg = vbCrLf & "粘度特征曲线;2;20,230|5-5,30-30,100-100,200-200|0-0,5-5,10-10,15-15,20-20;"
                
                If Trim("" & rsTmp("性别")) = "女" Then
    '                ND0 = 9.25: QB0 = 10: ND1 = 4.4: QB1 = 150
    '                ND2 = 6.5: QB2 = 10: ND3 = 3.65: QB3 = 150
                    strImg = strImg & "9.25,10,4.4,150|6.5,10,3.65,150"
                Else
    '                ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
    '                ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                    strImg = strImg & "9.58,10,4.6,150|6.8,10,3.73,150"
                End If
                strImg = strImg & "~5-" & rsTmp!f3 & ",30-" & rsTmp!f2 & ",200-" & rsTmp!f1  '描点
                strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),220,1" '坐标轴
                
                aResults(lngResultCount) = aResults(lngResultCount) & strImg
            ElseIf str_仪器类型 = "LG-R-80F_WXX" Then
                strImg = vbCrLf & "粘度特征曲线;2;35,250|1-1,40-3,80-10,120-30,180-100,230-300|0-0,5-5,10-10,15-15,20-20,25-25,30-30;"
                
                If Trim("" & rsTmp("性别")) = "女" Then
                    strImg = strImg & "11.55,12,5.4,150"
                Else
                    strImg = strImg & "11.78,12,5.6,150"
                End If
                strImg = strImg & "~"
                strImg = strImg & ";VIS(mPa.s),35,35~SHR(1/S),225,2" '坐标轴
                aResults(lngResultCount) = aResults(lngResultCount) & strImg
            End If
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


