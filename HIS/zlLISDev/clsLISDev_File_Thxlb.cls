VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_Thxlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:天海血流变仪
'厂商:重庆天海
'作者:陈东
'用户:南桐
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsAnti As ADODB.Recordset
    Dim strBh As String, strAnti As String, str结果 As String
    Dim lngResultCount As Long
    Dim str参数 As String, arr参数() As String
    Dim curNd0 As Currency, curNd1 As Currency, curNd2 As Currency, curNd3 As Currency
    Dim cur高切 As Currency, cur低切 As Currency, cur中切 As Currency, str日期 As String, strImg As String
    Dim str备注 As String
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    '配置文件用的是仪器商的配置文件，在技师站的参数选择中指定文件时，指定此配置文件，配置文件内容中要用到的内容如下：
    '[BDE_MSQL_XLB]
    'Database Name = th_xlb2035
    'SERVER Name = PC081128
    'USER Name = sa
    'PASSWORD=
    '--------------------------------------------------------------------------------------------------------------------
    '读配置文件
    Call ReadConf(strFile, "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str参数)
    
    arr参数 = Split(str参数, ",")
    If UBound(arr参数) <> 3 Then Exit Function
    With cnAccess
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & arr参数(1) & ";Password=" & arr参数(3) & _
        ";Persist Security Info=True;User ID=" & arr参数(2) & ";Data Source=" & arr参数(0)
        .Open
        
        If .State = adStateClosed Then Exit Function
        strSQL = "SELECT A.SXH as 标本号, A.SEX as 性别,B.SXH as 通道号, B.VAL as 结果, A.RQ as 日期,A.BZ as 备注,B.V_Max,B.V_Min " & _
                 " FROM JZXX as A INNER JOIN JZ_ITEM as B ON A.ID = B.ID AND A.RQ >= '" & Format(dtStart, "yyyy-MM-dd") & "' And A.RQ <= '" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                  IIf(lngSampleNO = -1, "", " And A.SXH =" & lngSampleNO) & _
                  " order by A.RQ,A.SXH,B.SXH"
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        
        ReDim aResults(0)
        lngResultCount = 0
        strBh = "***"
        
        'rsTmp.MoveFirst
        Do Until rsTmp.EOF
            If Val("" & rsTmp("标本号")) > 0 Then '新标本
                

                
                If strBh <> "" & rsTmp("标本号") And strBh <> "***" Then
                    If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                    aResults(UBound(aResults)) = str日期
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh) '样本序号
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血清"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
                    aResults(UBound(aResults)) = aResults(UBound(aResults)) & str结果 & IIf(str备注 <> "", "|BZ|" & str备注, "")
                    
                    
                    strImg = vbCrLf & "粘度特征曲线;2;20,220|30-30,50-50,100-100,150-150,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
                    
                    'ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
                    'ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
                    
                    strImg = strImg & curNd0 & ",30," & curNd1 & ",200|" & curNd2 & ",30," & curNd3 & " ,200"
                    
                    strImg = strImg & "~30-" & cur低切 & ",100-" & cur中切 & ",200-" & cur高切  '描点
                    strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),200,1" '坐标轴
                    
                    If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & strImg
                    
                    strImg = ""
                    str结果 = "": curNd2 = 0: curNd0 = 0: curNd1 = 0: curNd3 = 0
                    cur低切 = 0: cur高切 = 0: cur中切 = 0
                End If
                strBh = Val("" & rsTmp("标本号"))
                str结果 = str结果 & "|" & rsTmp("通道号") & "|" & rsTmp("结果")
                str日期 = Format("" & rsTmp("日期"), "yyyy-mm-dd hh:mm:ss") '检验时间
                str备注 = Trim("" & rsTmp("备注"))
                '图形公式
                'b = ln(高切值 / 低切值) / ln(200 / 3)
                'a = 低切值/power(3,b)     //其中power(3,b)为了3的b次幂
                '
                '//a,b为计算用系数
                '
                '中切 = a * power(时间值, b)
                '=========================================================
                'ln: 以E为底的自然对数
                'jz_item :中取高切,低切,中切的值,绘出三根曲线.
                '=========================================================
                
                
                If rsTmp("通道号") = 1 Then
                    curNd2 = Val("" & rsTmp!v_max): curNd3 = Val("" & rsTmp!v_min)
                    cur高切 = Val("" & rsTmp("结果"))
                End If
                If rsTmp("通道号") = 3 Then
                    curNd0 = Val("" & rsTmp!v_max): curNd1 = Val("" & rsTmp!v_min)
                    cur低切 = Val("" & rsTmp("结果"))
                End If
                If rsTmp("通道号") = 2 Then
                    cur中切 = Val("" & rsTmp("结果"))
                End If
                

            End If
            rsTmp.MoveNext
        Loop
        
        If strBh <> "" And str结果 <> "" Then
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
            aResults(UBound(aResults)) = str日期
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(strBh) '样本序号
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|血清"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
            aResults(UBound(aResults)) = aResults(UBound(aResults)) & str结果 & IIf(str备注 <> "", "|BZ|" & str备注, "")
            
            
            strImg = vbCrLf & "粘度特征曲线;2;20,220|30-30,50-50,100-100,150-150,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;"
            
            'ND0 = 9.58: QB0 = 10: ND1 = 4.6: QB1 = 150
            'ND2 = 6.8: QB2 = 10: ND3 = 3.73: QB3 = 150
            
            strImg = strImg & curNd0 & ",30," & curNd1 & ",200|" & curNd2 & ",30," & curNd3 & " ,200"
            
            strImg = strImg & "~30-" & cur低切 & ",100-" & cur中切 & ",200-" & cur高切  '描点
            strImg = strImg & ";VIS(mPa.s),35,20~SHR(1/S),200,1" '坐标轴
            
            If strImg <> "" Then aResults(UBound(aResults)) = aResults(UBound(aResults)) & strImg
            
            strImg = ""
            str结果 = "": curNd2 = 0: curNd0 = 0: curNd1 = 0: curNd3 = 0
            cur低切 = 0: cur高切 = 0: cur中切 = 0
        End If
    End With
    
    ResultFromFile = aResults
   
End Function

Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '读配置文件
    'strIniFile :ini配置文件名
    'strItems   :要读取的字段，如：DATABASE NAME=th_xlb中的 DATABASE NAME,多个字段以,分隔
    'strReturn  :读取的返回值，如:DATABASE NAME=th_xlb中的 th_xlb ，多个反回值以,分隔
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intloop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '从配置文件中读取保存在文件中的配置cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intloop = 0 To intCount
        arrItems(intloop, 0) = Split(strItems, ",")(intloop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intloop = 0 To intCount
                    If arrLine(0) = arrItems(intloop, 0) Then
                        arrItems(intloop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intloop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intloop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub
