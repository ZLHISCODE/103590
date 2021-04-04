VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_LineGene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器: PCR 荧光定量分析仪 LineGene
'厂商:杭州博日
'作者:陈东
'用户:梅河口市受民医院
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
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
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
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strBh As String, strimg As String, strLastBH As String, strAllImg As String, strAllItems As String
    
    Dim cnErs As New ADODB.Connection
    Dim strFileErs As String, rsErs As ADODB.Recordset
    Dim strImgErs As String, str日期 As String, str项目 As String, str结果 As String
    Dim strChar As String, str接口类型 As String
    Dim intMaxValue As Integer, i As Integer
    Dim strQC As String
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strChar = Replace(strFile, strFilePath & "\", "")
    strChar = Mid(strChar, 4, 1)
    
    strFileErs = strFile
    
    If Trim(Dir(strFileErs)) <> "" Then
        With cnErs
        .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then strFileErs = ""
        End With
    Else
        strFileErs = ""
    End If
        
    ReDim aResults(0)
    
    strImgErs = ""
    strQC = "0"
    str接口类型 = Trim(ReadIni("接口", "类型", App.Path & "\LineGane.ini"))
    If str接口类型 = "" Then
        WriteIni "接口", "类型", "梅河口", App.Path & "\LineGane.ini"
        str接口类型 = "梅河口"
    End If
    
    If strFileErs <> "" Then
        strSQL = "SELECT sample_no  As 顺序号, report_date as 日期, " & _
                 "A.project_name as 项目,sample_name, A.con_Cal as 结果, ct,ct_ref,con_min,result_con, " & _
                 "A.fluorescence_1, A.fluorescence_2, A.fluorescence_3, A.fluorescence_4, A.fluorescence_5," & _
                 "A.fluorescence_6, A.fluorescence_7, A.fluorescence_8, A.fluorescence_9, A.fluorescence_10," & _
                 "A.fluorescence_11, A.fluorescence_12, A.fluorescence_13, A.fluorescence_14, A.fluorescence_15," & _
                 "A.fluorescence_16, A.fluorescence_17, A.fluorescence_18, A.fluorescence_19, A.fluorescence_20," & _
                 "A.fluorescence_21, A.fluorescence_22, A.fluorescence_23, A.fluorescence_24, A.fluorescence_25," & _
                 "A.fluorescence_26, A.fluorescence_27, A.fluorescence_28, A.fluorescence_29, A.fluorescence_30" & _
                 " FROM qpcr_data AS A Where Format(A.report_date,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.report_date,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                 "' Order by sample_no" '& _
                'IIf(lngSampleNO = -1, "", " And val(Mid(sample_no,9))=" & lngSampleNO)
        Set rsErs = cnErs.Execute(strSQL)
        Do Until rsErs.EOF
            If Val("" & rsErs("顺序号")) <> 0 Then
                str日期 = Format(rsErs("日期"), "yyyy-MM-dd")
                If strLastBH <> Trim("" & rsErs("顺序号")) And strLastBH <> "" Then
                    If strAllItems <> "" Then
                        If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                        aResults(UBound(aResults)) = str日期  '检验时间
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh)   '样本序号
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "  '检验人
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"  '标本
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & strQC
                        aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAllItems
                        If strAllImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strAllImg, 2)
                    End If
                    strAllItems = ""
                    strAllImg = ""
                    
                ElseIf strLastBH = "" Then
                    strAllItems = ""
                    strAllImg = ""
                End If
                
                strBh = Trim("" & rsErs("顺序号"))
                If Len(strBh) > 8 Then
                    strBh = Mid(strBh, 8)
                ElseIf Len(strBh) > 6 Then
                    strBh = Mid(strBh, 7)
                End If

                If str接口类型 = "梅河口" Then
                    strAllItems = strAllItems & "|" & rsErs!项目 & "|" & E_Number("" & rsErs!结果)
                    'str项目 = Trim("" & rsErs!项目)
                Else
                    str项目 = Trim("" & rsErs!sample_name)
                
                    If rsErs!result_con = 0 Then
                        '取CT值
                        If str项目 <> "HBV" Then
                            str结果 = Trim("" & rsErs!ct)
                            If str结果 = "-1" Then
                                str结果 = "∞ (阴性)"
                            ElseIf Val(str结果) > Val("" & rsErs!ct_ref) And Val("" & rsErs!ct_ref) > 0 Then
                                str结果 = str结果 & " (阴性)"
                            ElseIf str结果 <> "" Then
                                str结果 = str结果 & " (阳性)"
                            End If
                        Else
                            str项目 = ""
                        End If
                    Else
                        'HBV
                        If str项目 = "HBV" Then
                            str结果 = Trim("" & rsErs!结果)
                            If Val(str结果) < Val("" & rsErs!con_min) And Val("" & rsErs!con_min) <> 0 Then
                                str结果 = "<" & E_Number(Val(rsErs!con_min))
                            Else
                                str结果 = E_Number(Val(str结果))
                            End If
                        Else
                            str项目 = ""
                        End If
                    End If
                    
                    If str项目 <> "" And str结果 <> "" Then
                        strAllItems = strAllItems & "|" & str项目 & "|" & str结果
                    End If
                End If
                strImgErs = ""
                
                If strAllItems <> "" Then
                    For i = 6 To rsErs.Fields.count - 1
                        If rsErs.Fields(i).Value <> 0 Then
                            strImgErs = strImgErs & "," & Format((rsErs.Fields(i).Value / 6000) * 10, "0.000")
                        End If
                        If rsErs.Fields(i).Value > intMaxValue Then
                            intMaxValue = rsErs.Fields(i).Value
                        End If
                    Next
                    If strImgErs <> "" Then
                        strImgErs = "　;4;100,50,4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40|10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100;" & Mid(strImgErs, 2) & "~荧光强度,12,100|循环数,55,10"
'                        If intMaxValue <= 6000 Then
'                            strImgErs = "曲线;3;100,50|0-0,10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";荧光强度,5,100~循环数,55,1"
'                        ElseIf intMaxValue <= 10000 Then
'                            strImgErs = "曲线;3;100,50|0-0,10-10,20-20,30-30,40-40,50-50,60-60,70-70,80-80,90-90,100-100|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36,40-40;" & Mid(strImgErs, 2) & ";荧光强度,5,100~循环数,55,1"
'                        End If
                        strAllImg = strAllImg & "^" & strImgErs
                    End If
                End If
                
                strLastBH = Trim("" & rsErs("顺序号"))
            End If
            rsErs.MoveNext
        Loop
        
        
        If strAllItems <> "" Then
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            
            aResults(UBound(aResults)) = str日期  '检验时间
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & strBh   '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & " "  '检验人
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & "血液"  '标本
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & strAllItems
            If strAllImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & Mid(strAllImg, 2)
            strAllItems = ""
            strAllImg = ""
        End If
        
    End If
        

    ResultFromFile = aResults
End Function

Private Function E_Number(ByVal curIn As Double) As String
    '将数字转为科学计数法显示
    '超过1000的才转,保留三位小数
    Dim strIn As String
    strIn = curIn
    
    If Len(Split(strIn, ".")(0)) > 0 Then
        strIn = Format(Val(strIn) / 10 ^ (Len(Split(strIn, ".")(0)) - 1), "0.00") & "E+" & Format(Len(Split(strIn, ".")(0)) - 1, "00")
    End If
    E_Number = IIf(Left(strIn, 1) = ".", "0" & strIn, strIn)
End Function

