VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, _
                        Optional ByVal intSaveType As Integer) As Boolean
    '外部调用接口一
    '功能：根据传入的图形数据，绘图
    DrawImg = DrawImgAndSaveFile(strType, strData, strFileName, intSaveType)
End Function

Public Function GetSampleImages(ByVal lngSampleID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
'获取一个标本的所有图片。
'
'lngsampleid : 标本ID
'strPath     : 图片保存路径,不存在则自动创建
'blnClearOldFile :是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
'strErr      : 函数返回值为空串时，返回的提示信息
'intImgType  : 返回的图片文件格式，0－cht(默认),1-jgp,2-png
'intVer      : 是新版LIS还是老版LIS在调用本函件数， 0-老版LIS（默认，从“检验图像结果”中取图形数据），1-新版LIS（从“检验报告图像”中取图形数据）

    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer, strFileType As String
    Dim StrTime As Date
    Dim strSavePath As String
    Dim strReturn As String
    Dim strFileName As String
    Dim objStream As textStream


    Dim blnFtp As Boolean       'FTP是否可用
    Static strFtpPara As String       '保存FTP参数
    Dim strFtpUser As String, strFtpPass As String, strFtpIP As String, strFtpDir As String
    Dim strFtpPath As String

    Dim intLayOut As Integer, strImageType As String
    Dim strImageData As String, strLocalFile As String, strFileNames() As String
    Dim strFileAll() As String, strFtpPathName As String, intLoop As Integer
    Dim blnDown As Boolean

    ReDim Preserve strFileNames(1)
    ReDim Preserve strFileAll(1)

100 strReturn = ""
102 If Not gblnInit Then
104     strErr = "调用本函数前，请先调用初始化函数！"
        Exit Function
    End If

106 If intSaveImgType = 1 Then
108     strFileType = ".jpg"
110 ElseIf intSaveImgType = 2 Then
112     strFileType = ".png"
    Else
114     strFileType = ".cht"
    End If
116 strSavePath = strPath
118 If Trim(strSavePath) = "" Then
120     strErr = "请指定保存图片的路径！"
        Exit Function
122 ElseIf Not gobjFSO.FolderExists(strSavePath) Then
124     strErr = strSavePath & "不存在！"
        Exit Function
    End If
126 If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
    '---------------------------------------------------
    On Error GoTo errH
128 StrTime = Now
130 If intVer = 0 Then
132     gstrSql = "select id ,标本ID,图像类型,图像位置 from 检验图像结果 where 标本id = [1] order by ID"
    Else
134     gstrSql = "select id ,标本ID,图像类型,图像位置 from 检验报告图像 where 标本id = [1] order by ID"
    End If

136 Set rsTmp = OpenSQLRecord(gstrSql, "GetSampleImage", lngSampleID)
138 If rsTmp.EOF Then
        strErr = "无图像数据！"
        Exit Function
    End If

    Call FunFtpSet(blnFtp, intVer, strFtpPara, strFtpUser, strFtpPass, strFtpIP, strFtpDir)

140 Do Until rsTmp.EOF
        If intVer = 0 Then
            strFileName = rsTmp("标本ID") & "_" & Trim$("" & rsTmp("图像类型")) & strFileType
        Else
            strFileName = rsTmp("标本ID") & "_" & Trim$("" & rsTmp("图像类型")) & strFileType
        End If
142     If blnClearOldFile = True Then
            '强制清除缓存在本地的图片
144         If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
            WriteLog "强制清除本地缓存的图片", strSavePath, strFileName
        End If

146     If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
            If blnFtp = False Then
148             If LoadImageDataTwo(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
150                 If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                End If
            Else
                strFtpPath = Trim("" & rsTmp!图像位置)
                If InStr(strFtpPath, ";") <= 0 Then
                    If LoadImageDataTwo(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
                        If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                    End If
                Else
                    If strFileNames(UBound(strFileNames)) <> "" Then
                        ReDim Preserve strFileNames(UBound(strFileNames) + 1)
                        ReDim Preserve strFileAll(UBound(strFileAll) + 1)
                    End If
                    strFtpPathName = Trim(Split(strFtpPath, ";")(1))
                    strFileNames(UBound(strFileNames)) = Split(strFtpPathName, "/")(UBound(Split(strFtpPathName, "/")))
                    strFileAll(UBound(strFileAll)) = Trim("" & rsTmp!图像类型) & ";" & strFtpPath
                End If
            End If
152         DrawIndex = DrawIndex + 1
        Else
            strReturn = strReturn & "," & strFileName
        End If
154     rsTmp.MoveNext
    Loop

    If blnFtp Then
        strLocalFile = strPath
        blnDown = True
        For intLoop = 1 To UBound(strFileNames)
            If strFileNames(intLoop) <> "" Then
                If Dir(strPath & "\" & strFileNames(intLoop)) <> "" Then
                    strFileName = Replace(strFileNames(intLoop), ".bmp", ".cht")
                    strFileName = Replace(strFileName, ".jpg", ".cht")
                    strFileName = Replace(strFileName, ".png", ".cht")
                    strFileName = Replace(strFileName, ".BMP", ".cht")
                    strFileName = Replace(strFileName, ".JPG", ".cht")
                    strFileName = Replace(strFileName, ".PNG", ".cht")
                    strFileName = Replace(strFileName, ".gif", ".cht")
                    strFileName = Replace(strFileName, ".GIF", ".cht")
                    strFileName = Replace(strFileName, ".txt", ".cht")
                    strFileName = Replace(strFileName, ".TXT", ".cht")
                    
                    strReturn = strReturn & "," & strFileName
                    blnDown = False
                End If
            End If
        Next
        If blnDown = True Then
            If DownFiles(strFtpUser, strFtpPass, strFtpIP, strFtpPathName, strLocalFile, strFileNames) = "" Then
                For intLoop = 1 To UBound(strFileNames)
                    If strFileNames(intLoop) <> "" Then
166                     strFtpPath = ""
168                     strFtpPath = strFileAll(intLoop)
                        '图像存在FTP中，从FTP中取数据
                        '图像位置的数据格式为：图像格式;FTP文件路径
                        strImageType = Trim(Split(strFtpPath, ";")(0))
262                     intLayOut = Val(Split(strFtpPath, ";")(1))
264                     strFtpPath = Trim(Split(strFtpPath, ";")(2))
266                     strImageData = ""
268                     If intLayOut >= 100 And intLayOut <= 227 Then
                            ' 图片文件，直接下载到本地
270                         strLocalFile = strPath & "\" & Split(strFtpPath, "/")(UBound(Split(strFtpPath, "/")))
278                         strImageData = intLayOut & ";" & strLocalFile
                        Else
                            ' 图形数据，需要从下载的文本文件中读取数据
280                         strLocalFile = strPath & "\" & strFileNames(intLoop)
288                         Set objStream = gobjFSO.OpenTextFile(strLocalFile, ForReading)
290                         Do Until objStream.AtEndOfLine
292                             strImageData = strImageData & objStream.ReadLine
                            Loop
294                         objStream.Close
296                         Set objStream = Nothing
298                         strImageData = Replace(Replace(Trim(strImageData), vbCr, ""), vbLf, "")
300                         strImageData = intLayOut & ";" & strImageData
302                         If gobjFSO.FileExists(strLocalFile) Then gobjFSO.DeleteFile strLocalFile
                        End If

304                     If Len(strImageData) <> 0 Then
                            strFileName = Replace(strFileNames(intLoop), ".bmp", ".cht")
                            strFileName = Replace(strFileName, ".jpg", ".cht")
                            strFileName = Replace(strFileName, ".png", ".cht")
                            strFileName = Replace(strFileName, ".BMP", ".cht")
                            strFileName = Replace(strFileName, ".JPG", ".cht")
                            strFileName = Replace(strFileName, ".PNG", ".cht")
                            strFileName = Replace(strFileName, ".gif", ".cht")
                            strFileName = Replace(strFileName, ".GIF", ".cht")
                            strFileName = Replace(strFileName, ".txt", ".cht")
                            strFileName = Replace(strFileName, ".TXT", ".cht")
                            
                            '画图并产生图形文件
306                         Call DrawImgAndSaveFile(strImageType, strImageData, strPath & "\" & strFileName, intSaveImgType)
                        End If
                        If gobjFSO.FileExists(strPath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                        '                        intLoop = 0
                        '                        Do Until intLoop > 100
                        '                            intLoop = intLoop + 1
                        '                            If gobjFSO.FileExists(strLocalFile) Then
                        '                                WriteLog "LoadImageData", "第" & intLoop & "次删除FTP下载的原始文件" & strLocalFile, ""
                        '                                Call gobjFSO.DeleteFile(strLocalFile)
                        '                            Else
                        '                                If strLocalFile <> "" Then WriteLog "LoadImageData", "FTP下载的原始文件" & strLocalFile & "已删除!", ""
                        '                                Exit Do
                        '                            End If
                        '                        Loop
                    End If
                Next
            End If
        End If
    End If
156 If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158 GetSampleImages = strReturn
    WriteLog "返回图形文件名", "", strReturn
    'Debug.Print "ID=" & lngKeyID & ",用时:" & DateDiff("s", strTime, Now)
    Exit Function
errH:
160 GetSampleImages = ""
162 strErr = Err.Description
164 WriteLog "GetSampleImages", CStr(Erl()) & "行", Err.Description
End Function

Public Function GetSampleImagesOld(ByVal lngSampleID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
        '获取一个标本的所有图片。
        '
        'lngsampleid : 标本ID
        'strPath     : 图片保存路径,不存在则自动创建
        'blnClearOldFile :是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
        'strErr      : 函数返回值为空串时，返回的提示信息
        'intImgType  : 返回的图片文件格式，0－cht(默认),1-jgp,2-png
        'intVer      : 是新版LIS还是老版LIS在调用本函件数， 0-老版LIS（默认，从“检验图像结果”中取图形数据），1-新版LIS（从“检验报告图像”中取图形数据）
    
        Dim rsTmp As New ADODB.Recordset
        Dim DrawIndex As Integer, strFileType As String
        Dim StrTime As Date
        Dim strSavePath As String
        Dim strReturn As String
        Dim strFileName As String
    
100     strReturn = ""
102     If Not gblnInit Then
104         strErr = "调用本函数前，请先调用初始化函数！"
            Exit Function
        End If
    
106     If intSaveImgType = 1 Then
108         strFileType = ".jpg"
110     ElseIf intSaveImgType = 2 Then
112         strFileType = ".png"
        Else
114         strFileType = ".cht"
        End If
116     strSavePath = strPath
118     If Trim(strSavePath) = "" Then
120         strErr = "请指定保存图片的路径！"
            Exit Function
122     ElseIf Not gobjFSO.FolderExists(strSavePath) Then
124         strErr = strSavePath & "不存在！"
            Exit Function
        End If
126     If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
        '---------------------------------------------------
        On Error GoTo errH
128     StrTime = Now
130     If intVer = 0 Then
132         gstrSql = "select id ,标本ID,图像类型 from 检验图像结果 where 标本id = [1] order by ID"
        Else
134         gstrSql = "select id ,标本ID,图像类型 from 检验报告图像 where 标本id = [1] order by ID"
        End If
    
136     Set rsTmp = OpenSQLRecord(gstrSql, "GetSampleImage", lngSampleID)
138     If rsTmp.EOF Then strErr = "无图像数据！"
        
    
140     Do Until rsTmp.EOF
            If intVer = 0 Then
                strFileName = rsTmp("标本ID") & strFileType
            Else
                strFileName = rsTmp("标本ID") & "_" & Trim$("" & rsTmp("图像类型")) & strFileType
            End If
142         If blnClearOldFile = True Then
                '强制清除缓存在本地的图片
144             If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
                WriteLog "强制清除本地缓存的图片", strSavePath, strFileName
            End If
        
146         If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
        
148             If LoadImageData(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
150                 If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
                End If
            
152             DrawIndex = DrawIndex + 1
            Else
                strReturn = strReturn & "," & strFileName
            End If
154         rsTmp.MoveNext
        Loop
156     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158     GetSampleImagesOld = strReturn
        WriteLog "返回图形文件名", "", strReturn
        'Debug.Print "ID=" & lngKeyID & ",用时:" & DateDiff("s", strTime, Now)
        Exit Function
errH:
160     GetSampleImagesOld = ""
162    strErr = Err.Description
164    WriteLog "GetSampleImages", CStr(Erl()) & "行", Err.Description
End Function

Public Function GetImage(ByVal lngImgID As Long, ByVal strPath As String, ByVal blnClearOldFile As Boolean, _
                                ByRef strErr As String, Optional intSaveImgType As Integer = 0, Optional intVer As Integer = 0) As String
        '获取指定的一张图片。
        '
        'lngImgid    : 图片ID
        'strPath     : 图片保存路径,不存在则自动创建
        'blnClearOldFile :是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
        'strErr      : 函数返回值为空串时，返回的提示信息
        'intImgType  : 返回的图片文件格式，0－cht(默认),1-jgp,2-png
        'intVer      : 是新版LIS还是老版LIS在调用本函件数， 0-老版LIS（默认，从“检验图像结果”中取图形数据），1-新版LIS（从“检验报告图像”中取图形数据）
        Dim strSavePath As String, strFileType As String
        Dim strReturn As String
        Dim rsTmp As New ADODB.Recordset
        Dim strFileName As String
100     If Not gblnInit Then
102         strErr = "调用本函数前，请先调用初始化函数！"
            Exit Function
        End If
104     If intSaveImgType = 1 Then
106         strFileType = ".jpg"
108     ElseIf intSaveImgType = 2 Then
110         strFileType = ".png"
        Else
112         strFileType = ".cht"
        End If
114     strSavePath = strPath
116     If Trim(strSavePath) = "" Then
118         strErr = "请指定保存图片的路径！"
            Exit Function
120     ElseIf Not gobjFSO.FolderExists(strSavePath) Then
122         strErr = strSavePath & "不存在！"
            Exit Function
        End If
124     If Right(strSavePath, 1) = "\" Then strSavePath = Left(strSavePath, Len(strSavePath) - 1)
        On Error GoTo errH
    
126     If intVer = 0 Then
128         gstrSql = "select id ,标本ID,图像类型 from 检验图像结果 where id = [1]"
        Else
130         gstrSql = "select id ,标本ID,图像类型 from 检验报告图像 where id = [1]"
        End If
    
132     Set rsTmp = OpenSQLRecord(gstrSql, "GetImage", lngImgID)
134     If rsTmp.EOF Then strErr = "无图像数据！"
        If intVer = 0 Then
            strFileName = rsTmp("ID") & strFileType
        Else
            strFileName = rsTmp("ID") & "_" & Trim$("" & rsTmp("图像类型")) & strFileType
        End If
136     If blnClearOldFile = True Then
            '强制清除缓存在本地的图片
138         If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then gobjFSO.DeleteFile strSavePath & "\" & strFileName, True
            WriteLog "强制清除本地缓存的图片", strSavePath, strFileName
        End If
    
140     If Not gobjFSO.FileExists(strSavePath & "\" & strFileName) Then
142         If LoadImageData(strSavePath, rsTmp("ID"), intSaveImgType, intVer, strFileName) Then
144             If gobjFSO.FileExists(strSavePath & "\" & strFileName) Then strReturn = strReturn & "," & strFileName
            Else
                strErr = "产生图形文件失败！"
            End If
        Else
            strReturn = strReturn & "," & strFileName
        End If
146     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
148     GetImage = strReturn
        WriteLog "返回图形文件名", "", strReturn
        Exit Function
errH:
150     GetImage = ""
152     strErr = Err.Description
154     WriteLog "GetImage", CStr(Erl()) & "行", Err.Description
End Function


Public Function GetSampleImgInit(ByVal lngSys As Long, ByVal cnnOra As ADODB.Connection, ByRef strErr As String) As Boolean
    '初始化连接及公共部件
    '
    Dim arrData
    Dim i As Integer
    Dim strServer As String, strUser As String, strPass As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    
    GetSampleImgInit = gblnInit
    glngSys = lngSys
    If gblnInit = True Then Exit Function
    Set gcnOracle = cnnOra
    strSQL = "Select SYS_CONTEXT('USERENV','TERMINAL') as MName From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "初始化")
    gstrComputerName = rsTmp("MName")
'    With gcnOracle
'        Set gComLib = CreateObject("zl9Comlib.clsComlib")
'        Call gComLib.initCommon(gcnOracle)
'    End With
    
    gblnInit = True
    GetSampleImgInit = True
    Exit Function
errH:
    strErr = Err.Description
    WriteLog "GetImgInit", CStr(Erl()) & "行", Err.Description
End Function

Public Function GetSampleImgExit(ByRef strErr As String) As Boolean
    '关闭连接
    On Error GoTo errH
    strErr = ""
    If gblnInit Then
        Set gComLib = Nothing
        Set gcnOracle = Nothing
        gblnInit = False
    End If
    GetSampleImgExit = True
    Exit Function
errH:
    strErr = Err.Description
    WriteLog "GetImgExit", CStr(Erl()) & "行", Err.Description
End Function


Public Function LocalImageSaveToCht(ByVal strTitle As String, ByVal strImageData As String, ByVal strSaveFileName As String) As Boolean
    '将本地图形文件生成Cht放入指定目录
    Dim strPath As String
    Dim strChtName As String
    strPath = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strPath) Then gobjFSO.CreateFolder (strPath)
    strChtName = strPath & "\" & strSaveFileName
    If gobjFSO.FileExists(strChtName) Then gobjFSO.DeleteFile strChtName, True
    
    LocalImageSaveToCht = DrawImgAndSaveFile(strTitle, strImageData, strChtName, 0)
End Function



