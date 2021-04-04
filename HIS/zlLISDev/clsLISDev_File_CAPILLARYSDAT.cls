VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_CAPILLARYSDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:CAPILLARYS 电永仪
'厂商:Sebia
'作者:蒋淑君
'用户:广元市中心医院
'读取本地out.dat

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
    Dim lngResultCount As Long, strFilePath As String
    Dim str标本号 As String, str日期 As String
    Dim strLine As String, objFile As New Scripting.FileSystemObject, objStream As textStream
    Dim lngdotnum As Long, strflag As String, strfactor As String, strdots As String, strImgErs As String
    Dim LngType As String, Server As String, Port As String, UserId As String, Password As String, Database As String
    Dim conn As Object 'New NpgsqlConnection
    Dim strIniFile As String, strSQL As String, rsErs As New ADODB.Recordset
    Dim objFileSystem As New Scripting.FileSystemObject

    On Error GoTo errhand
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    ReDim aResults(0)
    aResults(0) = ""
    ResultFromFile = aResults

    strIniFile = App.Path & "\CAPILLARYS.ini"
    If Not objFileSystem.FileExists(strIniFile) Then
        Set objStream = objFileSystem.CreateTextFile(strIniFile)
        objStream.WriteLine "[Options]"
        objStream.WriteLine "UseMode=0   ;0读dat文件1读数据库"
        objStream.WriteLine "Server=192.168.137.90"
        objStream.WriteLine "Port=5432"
        objStream.WriteLine "UserId=postgres"
        objStream.WriteLine "Password=postgres"
        objStream.WriteLine "Database=total"
        objStream.Close
        Set objStream = Nothing
        LngType = 0
    Else
        LngType = Val(ReadIni("Options", "UseMode", strIniFile))
        Server = Val(ReadIni("Options", "Server", strIniFile))
        Port = Val(ReadIni("Options", "Port", strIniFile))
        UserId = Val(ReadIni("Options", "UserId", strIniFile))
        Password = Val(ReadIni("Options", "Password", strIniFile))
        Database = Val(ReadIni("Options", "Database", strIniFile))
    End If

    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    lngResultCount = -1

    If LngType = 0 Then
        Set objStream = objFile.OpenTextFile(strFile, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If Len(strLine) = 2912 Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)
                str标本号 = Val(Mid(strLine, 6, 15))
                str日期 = Format(Mid(Mid(strLine, 257, 8), 5, 4) & "-" & Mid(Mid(strLine, 257, 8), 3, 2) & "-" & Mid(Mid(strLine, 257, 8), 1, 2), "YYYY-MM-DD")
                aResults(lngResultCount) = str日期 & "|" & str标本号 & "| |血液|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "01" & "|" & Mid(strLine, 367, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "02" & "|" & Mid(strLine, 372, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "03" & "|" & Mid(strLine, 377, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "04" & "|" & Mid(strLine, 382, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "05" & "|" & Mid(strLine, 387, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "06" & "|" & Mid(strLine, 392, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "07" & "|" & Mid(strLine, 397, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "08" & "|" & Mid(strLine, 402, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "09" & "|" & Mid(strLine, 407, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "10" & "|" & Mid(strLine, 412, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "R01" & "|" & Mid(strLine, 549, 5)
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & Mid(strLine, 1, 1) & "R02" & "|" & Mid(strLine, 554, 5)
                lngdotnum = Val(Mid(strLine, 790, 4))
                strflag = Mid(strLine, 794, 1)
                strfactor = Mid(strLine, 795, 4)
                strImgErs = Mid(strLine, 831, 1200)
                If strImgErs <> "" Then
                    strImgErs = "　;0;4000,300,0-;" & SebiaImg(Mid(strImgErs, 5, Len(strImgErs) - 8))
                End If
                If strImgErs <> "" Then
                    aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImgErs
                End If
            End If
        Loop
    Else
        With conn
            .ConnectionString = "Server=" & Server & ";Port=" & Port & ";User Id=" & UserId & ";Password=" & Password & ";Database=" & Database & ";"
            If .State = adStateClosed Then .Close
            .Open
        End With

        strSQL = "SELECT val(seq) As 顺序号,programma as 类型, data_analisi as 日期, " & _
                 "A.fraz_1,fraz_2,fraz_3,fraz_4,fraz_5,fraz_6,fraz_7,fraz_8,fraz_9,fraz_10,rapp1,Curva " & _
                 " FROM Anagrafica AS A Where Format(A.data_analisi,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.data_analisi,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                 "' " & IIf(lngSampleNo = -1, "", " And val(seq)=" & lngSampleNo)
        Set rsErs = conn.Execute(strSQL)
        Do Until conn.EOF
            If Val("" & rsErs("顺序号")) <> 0 Then
                lngResultCount = lngResultCount + 1
                ReDim Preserve aResults(lngResultCount)

                aResults(lngResultCount) = Format(rsErs("日期"), "yyyy-MM-dd")  '检验时间

                If "" & rsErs!类型 = "S" Then
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Val("" & rsErs("顺序号"))   '样本序号 血液标本+100
                Else
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                        Val("" & rsErs("顺序号")) + 200 '样本序号
                End If
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      " "  '检验人
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                      IIf("" & rsErs!类型 = "S", "血液", "尿液")  '标本

                aResults(lngResultCount) = aResults(lngResultCount) & "|0"
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "01|" & rsErs!fraz_1
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "02|" & rsErs!fraz_2
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "03|" & rsErs!fraz_3
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "04|" & rsErs!fraz_4
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "05|" & rsErs!fraz_5

                If "" & rsErs!类型 = "S" Then
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "06|" & rsErs!rapp1
                Else
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "06|" & rsErs!fraz_6
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "07|" & rsErs!fraz_7
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "08|" & rsErs!fraz_8
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "09|" & rsErs!fraz_9
                    aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!类型 & "10|" & rsErs!fraz_10
                End If
                strImgErs = "" & rsErs!Curva

                If strImgErs <> "" Then
                    strImgErs = "　;0;4000,300,0-;" & SebiaImg(Mid(strImgErs, 5, Len(strImgErs) - 8))
                    aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImgErs
                End If
            End If
            rsErs.MoveNext
        Loop
    End If
    ResultFromFile = aResults
errhand:
    WriteLog "CAPILLARYSDAT", CStr(Erl()) & ":" & Err.Description, ""
    Err.Clear
End Function

'生成图形数据串
Private Function SebiaImg(ByVal strImg As String) As String

    Do While strImg <> ""
        If Val("&H" & Left(strImg, 4)) < 0 Then
            SebiaImg = "1;" & SebiaImg
        Else
            SebiaImg = CStr(Val("&H" & Left(strImg, 4))) & ";" & SebiaImg
        End If
        strImg = Mid(strImg, 5)
    Loop
    If Right(SebiaImg, 1) = ";" Then SebiaImg = Mid(SebiaImg, 1, Len(SebiaImg) - 1)

End Function


