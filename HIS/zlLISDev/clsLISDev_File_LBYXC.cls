VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LBYXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血沉仪 LBY-XC
'厂商:普利生
'作者:陈东
'用户:西安长庆(LBY-N6B)
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
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strChar = Replace(strFile, strFilePath & "\", "")
    strChar = Mid(strChar, 4, 1)
    
    strFileErs = strFilePath & "\DEsr" & Format(dtStart, "YYMM") & ".MDB"
    
    If Trim(Dir(strFileErs)) <> "" Then '血沉数据
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    lngResultCount = -1
    '-- 血沉的图形数据
    strImgErs = ""
    If strFileErs <> "" Then
        strsql = "SELECT val(Mid(TestWholeno,9)) As 顺序号, Testdate as 日期, " & _
                 "A.TestDataXC as 血沉, A.TestDataPressureResult as 红细胞压积, A.TestDataKValue as 血沉方程K值, A.TestDataVm as 最大沉降速度, A.TestDataTm as 终末时间, A.TestDataSTBS as 停滞时间, " & _
                 "A.TestDate, A.TestData1, A.TestData2, A.TestData3, A.TestData4, A.TestData5, A.TestData6, A.TestData7, A.TestData8, A.TestData9, A.TestData10, A.TestData11, A.TestData12, A.TestData13, A.TestData14, A.TestData15, A.TestData16, A.TestData17, A.TestData18, A.TestData19, A.TestData20, A.TestData21, A.TestData22, A.TestData23, A.TestData24, A.TestData25, A.TestData26, A.TestData27, A.TestData28, A.TestData29, A.TestData30" & _
                 " FROM EsrData AS A Where Format(A.Testdate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.Testdate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                 "' And val(TestDataIndex) >0 " & _
                IIf(lngSampleNO = -1, "", " And val(Mid(TestWholeno,9))=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strsql)
        Do Until rsErs.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            
            aResults(lngResultCount) = Format(rsErs("日期"), "yyyy-MM-dd")  '检验时间
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                Val("" & rsErs("顺序号"))   '样本序号
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  " "  '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                  "血液"  '标本
            aResults(lngResultCount) = aResults(lngResultCount) & "|0"
            
            aResults(lngResultCount) = aResults(lngResultCount) & "|01|" & Val("" & rsErs!红细胞压积) * 100
            aResults(lngResultCount) = aResults(lngResultCount) & "|02|" & "" & rsErs!血沉
            aResults(lngResultCount) = aResults(lngResultCount) & "|03|" & "" & rsErs!血沉方程K值
            aResults(lngResultCount) = aResults(lngResultCount) & "|04|" & "" & rsErs!最大沉降速度
            aResults(lngResultCount) = aResults(lngResultCount) & "|05|" & "" & rsErs!终末时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|06|" & "" & rsErs!停滞时间
            
            strImgErs = ""
            For i = 9 To rsErs.Fields.count - 1
                strImgErs = strImgErs & "," & rsErs.Fields(i).Value
                If rsErs.Fields(i).Value > intMaxValue Then
                    intMaxValue = rsErs.Fields(i).Value
                End If
            Next
        
            If strImgErs <> "" Then
                If intMaxValue <= 36 Then
                    strImgErs = "血沉曲线;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;" & Mid(strImgErs, 2) & ";血沉值(mm),5,36~时间(m),55,1"
                ElseIf intMaxValue <= 40 Then
                    strImgErs = "血沉曲线;3;40,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";血沉值(mm),5,40~时间(m),55,1"
                ElseIf intMaxValue <= 44 Then
                    strImgErs = "血沉曲线;3;44,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44;" & Mid(strImgErs, 2) & ";血沉值(mm),5,44~时间(m),55,1"
                ElseIf intMaxValue <= 48 Then
                    strImgErs = "血沉曲线;3;48,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48;" & Mid(strImgErs, 2) & ";血沉值(mm),5,48~时间(m),55,1"
                ElseIf intMaxValue <= 52 Then
                    strImgErs = "血沉曲线;3;52,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52;" & Mid(strImgErs, 2) & ";血沉值(mm),5,52~时间(m),55,1"
                ElseIf intMaxValue <= 56 Then
                    strImgErs = "血沉曲线;3;56,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56;" & Mid(strImgErs, 2) & ";血沉值(mm),5,56~时间(m),55,1"
                ElseIf intMaxValue <= 60 Then
                    strImgErs = "血沉曲线;3;60,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56,60-60;" & Mid(strImgErs, 2) & ";血沉值(mm),5,36~时间(m),55,1"
                End If
                
                aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImgErs
            End If
            
            rsErs.MoveNext
        Loop
    End If
        

    ResultFromFile = aResults
End Function


