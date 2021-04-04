VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AVE764B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器厂家：爱威科技股份有限公司
'型号：AVE-764B集成AVE-752
'类型：尿液分析
'用户：遂宁市第一人民医院
'作者：敬秦
'时间：2014/4/3 15:17:08

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    On Error GoTo hErr '打开错误陷阱
    '------------------------------------------------
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
    Dim objStream As textStream, objFileSystem As New FileSystemObject
    Dim strFilePath As String
    Dim SERVERNAME As String, DATABASENAME As String, USERNAME As String, PassWord As String, PicPath As String, imgPath As String, imgName() As String
    Dim ITYP As Integer, i As Integer, j As Integer, int_name As Integer
    Dim strSQL As String, strResults() As String, strBth As String, str_img As String
    Dim cnAccess As New ADODB.Connection, rsSick As New ADODB.Recordset, rsChem As New ADODB.Recordset, rsUri As New ADODB.Recordset
    If Len(strFile) = 0 Then strFilePath = App.Path
    strFilePath = objFileSystem.GetParentFolderName(strFile)
    If Not objFileSystem.FileExists(App.Path & "\" & "AVE764B.ini") Then '获取数据库信息
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "AVE764B.ini")
            objStream.WriteLine "[AVE764B]"
            objStream.WriteLine "SERVERName = (local)"
            objStream.WriteLine "DatabaseName = AVE2"
            objStream.WriteLine "USERName = sa"
            objStream.WriteLine "PASSWORD= shuangfeng"
            objStream.WriteLine "ITYP= 0   ;0=不换算异常、正常红细胞 1=换算"
            objStream.Close
            Set objStream = Nothing
    End If
    SERVERNAME = ReadIni("AVE764B", "SERVERName", App.Path & "\" & "AVE764B.ini")
    DATABASENAME = ReadIni("AVE764B", "DatabaseName", App.Path & "\" & "AVE764B.ini")
    USERNAME = ReadIni("AVE764B", "USERName", App.Path & "\" & "AVE764B.ini")
    PassWord = ReadIni("AVE764B", "PASSWORD", App.Path & "\" & "AVE764B.ini")
    ITYP = Val(ReadIni("AVE764B", "ITYP", App.Path & "\" & "AVE764B.ini"))
    If Not objFileSystem.FileExists(App.Path & "\" & "PicAVE764B.ini") Then '获取图片保存路径
        Set objStream = objFileSystem.CreateTextFile(App.Path & "\" & "PicAVE764B.ini")
        objStream.WriteLine "[AVE764B]"
        objStream.WriteLine "PICPATH="
        objStream.Close
        Set objStream = Nothing
    End If
    PicPath = ReadIni("AVE764B", "PICPATH", App.Path & "\" & "PicAVE764B.ini")
    
    If (Trim(SERVERNAME) = "") Or (Trim(DATABASENAME) = "") Or (Trim(USERNAME) = "") Then Exit Function
    With cnAccess
        .ConnectionString = _
        "Provider=SQLOLEDB.1;Initial Catalog=" & DATABASENAME & ";Password=" & PassWord & _
        ";Persist Security Info=True;User ID=" & USERNAME & ";Data Source=" & SERVERNAME
            .Open
    End With
    If cnAccess.State = adStateClosed Then Exit Function
    ReDim Preserve strResults(0) '初始化，避免出现下标越界
    strResults(0) = ""
    i = -1
    strSQL = "select a.SickSampleID, a.SickExamineID,a.ColorSet ,a.MuddySet,convert(varchar(100),a.CheckDate,20) as checktime from AVE2.dbo.SickInfo a " & vbNewLine   '处理病人信息
    strSQL = strSQL & "where convert(varchar(100),a.CheckDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),a.CheckDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and a.SickExamineID = " & Format(lngSampleNO, "0000"))
    Set rsSick = cnAccess.Execute(strSQL) '获取尿液基本信息
    Do Until rsSick.EOF
'        If strBth <> rsSick("SickSampleID").Value Then '每条新标本产生一条记录
            i = i + 1
            ReDim Preserve strResults(i)
            strResults(i) = Format(rsSick("checktime").Value, "yyyy-mm-dd hh:mm:ss") & "|" & Format(rsSick("SickExamineID").Value, "0000") & "| |尿液|0" & "|Color|" & " " & rsSick("ColorSet").Value & "|Muddy|" & " " & rsSick("MuddySet").Value
'        End If
        rsSick.MoveNext
    Loop
    rsSick.Close
    strSQL = "select b.ChemSampleID,b.ChemExamineID,convert(varchar(100),b.ChemDate,23) as chemtime,b.GLU,b.BLD,b.LEU,b.PRO,b.NIT,b.URO,b.BIL,b.KET,b.PH,b.SG,b.VC from AVE2.dbo.ChemResult b " & vbNewLine
    strSQL = strSQL & "where convert(varchar(100),b.ChemDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),b.ChemDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and b.ChemExamineID = " & Format(lngSampleNO, "0000"))
    Set rsChem = cnAccess.Execute(strSQL) '获取干化学检验结果信息
    Do Until rsChem.EOF
        i = i + 1
        ReDim Preserve strResults(i)
        strResults(i) = Format(rsChem("chemtime").Value, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "|" & Format(rsChem("ChemExamineID").Value, "0000") & "| |尿液|0"
        For j = 3 To rsChem.Fields.count - 1
            strResults(i) = strResults(i) & "|" & rsChem.Fields.Item(j).Name & "|" & " " & rsChem(rsChem.Fields.Item(j).Name).Value
        Next
        rsChem.MoveNext
    Loop
    rsChem.Close
    strSQL = "select c.UriSampleID,c.UriExamineID,convert(varchar(100),c.UriDate,23) as uritime,c.Rbc,c.Wbc,c.Crystal,c.Epiphyte,c.Duct,c.Epithelia,c.WhackRbc,c.SingularRbc,c.Infusorial," & _
                 "c.Mucous,c.Pyoid,c.NCrystal,c.LCrystal,c.HCrystalRef,c.CCrystal,c.WCrystal,c.RbcDuct,c.WbcDuct,c.KDuct,c.TDuct,c.LDuct,c.LEpithelia," & _
                 "c.XEpithelia,c.JEpithelia,c.BigRbc,c.SmallRbc,c.JRbc,c.ZRbc,c.QRbc,c.Neutrophil,c.Lymphocyte,c.Acidophil,c.Basophil,c.Monocyte," & _
                 "c.Bacterial , c.Cocci, c.Bacilli, c.Mildew, c.Yeast, c.Parasites, c.Sperm, c.PrintPath, c.PrintImage " & _
                 "from AVE2.dbo.UriResult c " & vbNewLine
    strSQL = strSQL & "where convert(varchar(100),c.UriDate,23) >= " & "'" & Format(dtStart, "yyyy-mm-dd") & "'" & " and " & " convert(varchar(100),c.UriDate,23) <= " & "'" & Format(dtEnd, "yyyy-mm-dd") & "'" & vbNewLine
    strSQL = strSQL & IIf(lngSampleNO = -1, "", " and c.UriExamineID = " & Format(lngSampleNO, "0000"))
    Set rsUri = cnAccess.Execute(strSQL) '获取镜检结果信息
    Do Until rsUri.EOF
        i = i + 1
        ReDim Preserve strResults(i)
        strResults(i) = Format(rsUri("uritime").Value, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "|" & Format(rsUri("UriExamineID").Value, "0000") & "| |尿液|0"
        For j = 3 To rsUri.Fields.count - 1
            If j >= 3 And j <= rsUri.Fields.count - 3 Then '一般检验数据
                strResults(i) = strResults(i) & "|" & rsUri.Fields.Item(j).Name & "|" & " " & rsUri(rsUri.Fields.Item(j).Name).Value
            End If
            If j = rsUri.Fields.count - 2 Then '获取图像保存的路径
                imgPath = rsUri(rsUri.Fields.Item(j).Name).Value
            End If
            If j = rsUri.Fields.count - 1 Then '获取图像的名称
                imgName = Split(rsUri(rsUri.Fields.Item(j).Name).Value, ",")
                For int_name = LBound(imgName) To UBound(imgName)
                    If imgName(int_name) <> "" Then
                        If imgName(int_name) = "H0001" Then
                            str_img = str_img & "^" & imgName(int_name) & ";113;0;" & imgPath & imgName(int_name) & ".jpg"
                            str_img = str_img & "^" & "H0002" & ";113;0;" & imgPath & "H0002" & ".jpg"
                            str_img = str_img & "^" & "MorphA" & ";113;0;" & imgPath & "MorphA" & ".jpg"
                            str_img = str_img & "^" & "MorphC" & ";113;0;" & imgPath & "MorphC" & ".jpg"
                            str_img = str_img & "^" & "MorphS" & ";113;0;" & imgPath & "MorphS" & ".jpg"
                            str_img = str_img & "^" & "MorphT" & ";113;0;" & imgPath & "MorphT" & ".jpg"
                        End If
                    End If
                Next
                If str_img = "" Then
                    str_img = str_img & "^" & "L0001" & ";113;0;" & imgPath & "L0001" & ".jpg"
                    str_img = str_img & "^" & "L0002" & ";113;0;" & imgPath & "L0002" & ".jpg"
                    str_img = str_img & "^" & "L0003" & ";113;0;" & imgPath & "L0003" & ".jpg"
                    str_img = str_img & "^" & "L0004" & ";113;0;" & imgPath & "L0004" & ".jpg"
                    str_img = str_img & "^" & "L0005" & ";113;0;" & imgPath & "L0005" & ".jpg"
                    str_img = str_img & "^" & "L0006" & ";113;0;" & imgPath & "L0006" & ".jpg"
                End If
                If str_img <> "" Then strResults(i) = strResults(i) & vbCrLf & Mid(str_img, 2)
                str_img = ""
            End If
        Next
        rsUri.MoveNext
    Loop
    rsUri.Close
    If UBound(strResults) = 0 Then Exit Function
   
    cnAccess.Close
    ResultFromFile = strResults
    '------------------------------------------------
    Exit Function
    '----------------
hErr:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrMsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",ResultFromFile,clsLISDev_File_AVE764B"
    Close #lngErrFileNo
End Function
