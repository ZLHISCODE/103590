VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AVE762A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AVE762A 尿沉渣
'厂商:长沙爱威
'作者:陈东
'用户:长治中研所

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
        Dim strBh As String, str结果 As String, str标本号 As String
        Dim lngResultCount As Long
        Dim str参数 As String, arr参数() As String
        Dim str日期 As String, iField As Integer
        Dim aPic() As String
        Dim strPic As String
        Dim strPicPath As String               '图像访问路径
        Dim intZIP As Integer
        Dim i As Integer, strImg As String, strImgTitle As String, iCount As Integer
        On Error GoTo errH
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
        '配置文件用的是仪器商的配置文件，在技师站的参数选择中指定文件时，指定此配置文件，配置文件内容中要用到的内容如下：
        '[BDE_MSQL_XLB]
        'Database Name = th_xlb2035
        'SERVER Name = PC081128
        'USER Name = sa
        'PASSWORD=
        '--------------------------------------------------------------------------------------------------------------------
    
104     strPicPath = ReadIni("AVE762A", "图片路径", App.Path & "\AVE762A.ini")
106     If strPicPath = "" Then strPicPath = "E:": Call WriteIni("AVE762A", "图片路径", "E:", App.Path & "\AVE762A.ini")
        '读配置文件
108     Call ReadConf(App.Path & "\AVE762A.ini", "SERVER NAME,DATABASE NAME,USER NAME,PASSWORD", str参数)
110     arr参数 = Split(str参数, ",")
112     If UBound(arr参数) <> 3 Then Exit Function
    
        '是否使用ZIP方式压缩文件(10.24.120)后才能使用
114     intZIP = Val(ReadIni("AVE762A", "是否压缩", App.Path & "\AVE762A.ini", "1"))
116     Call WriteIni("AVE762A", "是否压缩", CStr(intZIP), App.Path & "\AVE762A.ini")
    
118     With cnAccess
120         .ConnectionString = _
            "Provider=SQLOLEDB.1;Initial Catalog=" & arr参数(1) & ";Password=" & arr参数(3) & _
            ";Persist Security Info=True;User ID=" & arr参数(2) & ";Data Source=" & arr参数(0)
122         .Open

124         If .State = adStateClosed Then
126             WriteLog "Ave762A", "第125行", "未能连接到SQLSERVER"
                Exit Function
            End If
    '128         strSQL = "select A.id,A.chkdate,b.item0 as M1,b.item1 as M2,b.item2 as M3,b.item3 as M4,b.item4 as M5, " & _
    '                        "cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE 100-(CAST(b.Item6 AS float)/ CAST(b.Item0 AS float)) * 100 END) as decimal(5,1)) AS M6," & _
    '                        "cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE (cast(b.item6 AS float)/ cast(b.item0 AS float)) * 100 END) as decimal(5,1)) AS M7, " & _
    '                        " b.item7 as M8,b.item8 as  M9,b.item9 as M10,B.item10 as M11,B.item11 as M12,b.item12 as M13,b.item13 as M14,b.item14 as M15,b.item15 as M16," & _
    '                        "b.item16 AS M17,b.item17 AS M18,b.item18 AS M19,b.item19 as M20,b.item20 AS M21,B.iteM21 as M22,B.ITEM22 AS M23,B.ITEM23 AS M24," & _
    '                        "b.item24 AS M25,b.item25 AS M26 ,b.item26 AS M27,b.item27 AS M28,b.item28 AS ITEM29,b.item29 as M30,b.item30 as M31,b.item31 AS M32," & _
    '                        "b.item32 AS M33, b.item33 AS M34 , b.item34 AS M35 " & _
    '                        ",C.item0 as  C1,c.item1 as  C2,c.item2 as  C3,c.item3 as  C4,c.item4 as  C5,c.Item5 as  C6,c.item6 as  C7,c.item7 as  C8" & _
    '                        ",c.item8 as  C9,c.item9 as  C10,c.item10 as  C11,d.PicPrt " & vbNewLine & _
    '                        " from sick A LEFT OUTER JOIN " & _
    '                        "      chemic C ON A.ID = C.ID LEFT OUTER JOIN " & _
    '                        "      micro B  ON A.ID = B.ID LEFT OUTER JOIN " & vbNewLine & _
    '                        "      micro_out D ON A.id = d.id and a.MICROID = substring(d.id,9,4)   " & vbNewLine & _
    '                        " where  a.chkdate>='" & Format(dtStart, "yyyy.MM.dd") & "' and a.chkdate<='" & Format(dtEnd, "yyyy.MM.dd") & "'" & _
    '                        " order by A.chkdate,A.id"
    '                        "      micro_out D ON A.id = substring(d.Id,1,8) and a.MICROID = substring(d.id,9,4)   " & vbNewLine &
128                strSQL = "select A.id,A.chkdate,b.item0 as M1,b.item1 as M2,b.item2 as M3,b.item3 as M4,b.item4 as M5, " & _
                        " cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE 100-(CAST(b.Item6 AS float)/ CAST(b.Item0 AS float)) * 100 END) as decimal(5,1)) AS M6," & _
                        " cast((CASE WHEN cast(b.item0 as float)=0 THEN 0 ELSE (cast(b.item6 AS float)/ cast(b.item0 AS float)) * 100 END) as decimal(5,1)) AS M7," & _
                        "b.item7 as M8,b.item8 as  M9,b.item9 as M10,B.item10 as M11,B.item11 as M12,b.item12 as M13,b.item13 as M14,b.item14 as M15,b.item15 as M16," & _
                        "b.item16 AS M17,b.item17 AS M18,b.item18 AS M19,b.item19 as M20,b.item20 AS M21,B.iteM21 as M22,B.ITEM22 AS M23,B.ITEM23 AS M24," & _
                        "b.item24 AS M25,b.item25 AS M26 ,b.item26 AS M27,b.item27 AS M28,b.item28 AS ITEM29,b.item29 as M30,b.item30 as M31,b.item31 AS M32," & _
                        "b.item32 AS M33, b.item33 AS M34 , b.item34 AS M35" & _
                        ",C.item0 as  C1,c.item1 as  C2,c.item2 as  C3,c.item3 as  C4,c.item4 as  C5,c.Item5 as  C6,c.item6 as  C7,c.item7 as  C8" & _
                        ",c.item8 as  C9,c.item9 as  C10,c.item10 as  C11,d.PicPrt" & _
                        " from sick A LEFT OUTER JOIN" & _
                        " chemic C ON A.ID = C.ID LEFT OUTER JOIN" & _
                        " micro B  ON A.ID = B.ID LEFT OUTER JOIN" & _
                        " micro_out D ON A.id = d.id and a.MICROID = substring(d.id,9,4)" & _
                        "where  a.chkdate>='" & Format(dtStart, "yyyy.MM.dd") & "' and a.chkdate<='" & Format(DateAdd("d", 1, dtEnd), "yyyy.MM.dd") & "'" & _
                        " order by A.chkdate,A.id "
                        
130         Set rsTmp = cnAccess.Execute(strSQL)
132         If rsTmp.EOF Then
134             WriteLog "Ave762A", "第131行", Format(dtStart, "yyyy.MM.dd") & "至" & Format(dtEnd, "yyyy.MM.dd") & ",标本" & lngSampleNO & ",未找到结果"
136             .Close
                Exit Function
            End If
        
138         ReDim aResults(0)
140         lngResultCount = 0
142         strBh = "***"
        
            'rsTmp.MoveFirst
144         Do Until rsTmp.EOF
146             If Val("" & rsTmp("id")) > 0 Then '新标本
148                 strBh = "" & rsTmp("id")
150                 str日期 = Replace("" & rsTmp("chkdate"), ".", "-") '检验时间
152                 str结果 = ""
                
154                 For iField = 2 To rsTmp.Fields.count - 2
156                     If "" & rsTmp.Fields(iField).Value <> "" Then
158                         str结果 = str结果 & "|" & rsTmp.Fields(iField).Name & "|" & Replace(rsTmp.Fields(iField).Value, "mmol/l", "")
                        End If
                    Next
            
160                 If strBh <> "***" And str结果 <> "" Then
162                     str标本号 = Mid(strBh, 9)
164                     If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
166                     aResults(UBound(aResults)) = str日期
168                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|" & Val(str标本号) '样本序号
170                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "| "
172                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|尿液"
174                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & "|0"
176                     aResults(UBound(aResults)) = aResults(UBound(aResults)) & str结果
178                     strPic = ""
180                     iCount = 0
182                     For i = 1 To 6
                            '---先找JPG格式的图形
                            strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\H000" & i & ".jpg"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\L000" & i & ".jpg"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\F000" & i & ".jpg"
                            
                            '--- 没有再找BMP格式的图形
184                         If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\H000" & i & ".bmp"
186                         If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\L000" & i & ".bmp"
                            If Not gobjFSO.FileExists(strImg) Then strImg = strPicPath & "\" & Mid(strBh, 1, 8) & "\" & Mid(strBh, 9) & "\F000" & i & ".bmp"

188                         If Dir(strImg) <> "" Then
                                Dim strImgFmt As String
                                strImgFmt = "03"
                                If UCase(Right(strImg, 3)) = "JPG" Then strImgFmt = "13"
190                             If intZIP = 0 Then
192                                 strPic = strPic & "^" & strImgTitle & "000" & i & ";1" & strImgFmt & ";0;" & strImg
                                Else
194                                 strPic = strPic & "^" & strImgTitle & "000" & i & ";2" & strImgFmt & ";0;" & strImg
                                End If
196                             iCount = iCount + 1
                            Else
198                             WriteLog "Ave762A", "缺少图形：" & strImg, ""
                            End If
                        Next

200                     If strPic <> "" Then
202                         If Left(strPic, 1) = "^" Then strPic = Mid(strPic, 2)
204                         aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strPic
                        End If
                    End If
                End If
206             rsTmp.MoveNext
            Loop
        End With
    
208     ResultFromFile = aResults
        Exit Function
errH:
210 WriteLog "Ave762A", "第" & CStr(Erl()) & "行", Err.Description
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
                    If Trim(arrLine(0)) = Trim(arrItems(intloop, 0)) Then
                        arrItems(intloop, 1) = Trim(arrLine(1))
                        'Left(tempusername,   InStr(1,   tempusername,   Chr(0))   -   1)
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



