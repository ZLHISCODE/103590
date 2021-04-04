VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LBYN6C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血流变 普利生 LBY-N6C
'厂商:普利生
'作者:陈东
'用户:德阳医院,南京江北人民医院,西安长庆(LBY-N6B)
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
    
    strFile = strFilePath & "\DN6" & strChar & Format(dtStart, "YYMM") & ".MDB"
    strFileErs = strFilePath & "\DEsr" & Format(dtStart, "YYMM") & ".MDB"
    
    If Dir(strFile) = "" Then Exit Function
    If Trim(Dir(strFileErs)) <> "" Then '血沉数据
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
    
    With cnAccess
        
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = "Select val(Mid(TestWholeno,9)) As 顺序号, Patientname As 姓名, PatientSex As 性别, Testdate As 日期, Patientanalysisresult As 诊断, Qxdata1 As 低切值," & vbNewLine & _
                "       Qxdata5 As 中切值, Qxdata10 As 高切值, Xjdata As 粘度, Rltcalcparam1 As 压积, Rltcalcparam2 As 血沉," & vbNewLine & _
                "       Rltcalcdata1 As 还原粘度低切值, Rltcalcdata2 As 还原粘度中切值, Rltcalcdata3 As 还原粘度高切值," & vbNewLine & _
                "       Rltcalcdata4 As 血沉方程k值, Rltcalcdata5 As 聚集指数, Rltcalcdata6 As 刚性指数, Rltcalcdata7 As 变形指数," & vbNewLine & _
                "       Rltcalcdata8 As 电泳指数, Qxstrparam1 As 低切时间, Qxstrparam5 As 中切时间, Qxstrparam10 As 高切时间" & vbNewLine & _
                "From N6" & strChar & "data Where Format(Testdate,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "' And Format(Testdate,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "' " & _
                IIf(lngSampleNO = -1, "", " And val(Mid(TestWholeno,9))=" & lngSampleNO)
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        lngResultCount = -1
        strBh = "***"
        Do While Not rsTmp.EOF
            If strBh <> rsTmp("顺序号") Then '新标本
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                
                aResults(lngResultCount) = Format(rsTmp("日期"), "yyyy-mm-dd hh:mm:ss")  '检验时间
                
                str日期 = Format(rsTmp("日期"), "yyyy-mm-dd")
                
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    Val("" & rsTmp("顺序号"))   '样本序号
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      "血液"  '标本
                aResults(lngResultCount) = aResults(lngResultCount) & "|0" '& _
                   ' Val(rsTmp("质控")) '质控
                    
                strBh = rsTmp("顺序号")
            End If
            
            For iField = 5 To 19
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                    iField - 4 & "|" & Val(IIf(IsNull(rsTmp.Fields(iField).Value), "0", rsTmp.Fields(iField).Value)) '结果
            Next
            '-- 血流变的图形数据
            '-- 图形数据X-横，Y－纵
            '-- 格式：类型;座标数据;曲线及描点数据
            '        其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
            '              曲线及描点数据:曲线1的Y坐标1,曲线1的Y座标2.....|曲线2的Y坐标1,曲线2的Y坐标2,...~描点的Y从标1,描点的Y从标2
            strImg = vbCrLf & "粘度特征曲线;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
            
            If Trim("" & rsTmp("性别")) = "女" Then
'                ND0 = 9.25: QB0 = 10: ND1 = 4.4: QB1 = 150
'                ND2 = 6.5: QB2 = 10: ND3 = 3.65: QB3 = 150
                strImg = strImg & "9.25,10,4.4,150|6.5,10,3.65,150"
            Else
'                ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
'                ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                strImg = strImg & "9.58,10,4.6,150|6.8,10,3.73,150"
            End If
            strImg = strImg & "~10-" & rsTmp!低切值 & ",60-" & rsTmp!中切值 & ",150-" & rsTmp!高切值  '描点
            strImg = strImg & ";VIS(mPa.s),15,20~SHR(1/S),195,1" '坐标轴
            
            '-- 血沉的图形数据
            strImgErs = ""
            If strFileErs <> "" Then
                strSQL = "SELECT A.SerialNo, A.PatientSample, A.TestDataIndex, A.TestDate, A.TestData1, A.TestData2, A.TestData3, A.TestData4, A.TestData5, A.TestData6, A.TestData7, A.TestData8, A.TestData9, A.TestData10, A.TestData11, A.TestData12, A.TestData13, A.TestData14, A.TestData15, A.TestData16, A.TestData17, A.TestData18, A.TestData19, A.TestData20, A.TestData21, A.TestData22, A.TestData23, A.TestData24, A.TestData25, A.TestData26, A.TestData27, A.TestData28, A.TestData29, A.TestData30" & _
                        " FROM EsrData AS A Where Format(A.Testdate,'yyyy-MM-dd')>='" & Format(rsTmp("日期"), "yyyy-mm-dd") & "' And Format(A.Testdate,'yyyy-MM-dd')<='" & Format(rsTmp("日期"), "yyyy-mm-dd") & _
                        "' And val(TestDataIndex) >0 And A.SerialNo=" & Val(rsTmp("顺序号"))
                Set rsErs = cnErs.Execute(strSQL)
                Do Until rsErs.EOF
                    For i = 4 To rsErs.Fields.count - 1
                        strImgErs = strImgErs & "," & rsErs.Fields(i).Value
                        If rsErs.Fields(i).Value > intMaxValue Then
                            intMaxValue = rsErs.Fields(i).Value
                        End If
                    Next
                    rsErs.MoveNext
                Loop
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
'                    strImgErs = "血沉曲线;3;60,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40,44-44,48-48,52-52,56-56,60-60;" & Mid(strImgErs, 2) & ";血沉值(mm),5,60~时间(m),55,1"
                End If
            End If
            
            If strImgErs <> "" Then strImg = strImg & "^" & strImgErs
            
            aResults(lngResultCount) = aResults(lngResultCount) & strImg
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function


