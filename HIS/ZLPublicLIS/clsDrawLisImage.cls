VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawLisImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFilename As String, _
                        Optional ByVal intSaveType As Integer) As Boolean
    '外部调用接口一
    '功能：根据传入的图形数据，绘图
    DrawImg = DrawImgAndSaveFile(strType, strData, strFilename, intSaveType)
End Function

Public Function GetSampleImages(ByVal lngSampleId As Long, ByVal blnClearOldFile As Boolean, Optional intSaveImgType As Integer = 0) As String
    '获取一个标本的所有图片。
    '
    'lngsampleID : 标本ID
    'strPath     : 图片保存路径,不存在则自动创建
    'blnClearOldFile :是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
    'strErr      : 函数返回值为空串时，返回的提示信息
    'intImgType  : 返回的图片文件格式，0－cht(默认),1-jgp,2-png

        Dim rsTmp       As New ADODB.Recordset
        Dim DrawIndex   As Integer
        Dim strFileType As String
        Dim strTime     As Date
        Dim strReturn   As String
        Dim strFilename As String
        Dim objFso      As New FileSystemObject
        Dim strSql      As String
        
104     strReturn = ""
    
106     If intSaveImgType = 1 Then
108         strFileType = ".jpg"
110     ElseIf intSaveImgType = 2 Then
112         strFileType = ".png"
        Else
114         strFileType = ".cht"
        End If

118     If Trim(gstrFilePath) = "" Then
120         Call ErrLog("GetSampleImages", CStr(Erl()) & "行", "请指定保存图片的路径！", "")
            Exit Function
'122     ElseIf Not objFso.FolderExists(gstrFilePath) Then
'124         Call ErrLog("GetSampleImages", CStr(Erl()) & "行", gstrFilePath & "不存在！", "")
'            Exit Function
        End If
        If Not objFso.FolderExists(gstrFilePath) Then
            objFso.CreateFolder (gstrFilePath)
        End If
126     If Right(gstrFilePath, 1) = "\" Then gstrFilePath = Left(gstrFilePath, Len(gstrFilePath) - 1)
        '---------------------------------------------------
        On Error GoTo ErrH
128     strTime = Now
130     If gblnNewLis Then
132         strSql = "select ID ,标本ID,图像类型 from 检验报告图像 where 标本ID = [1] order by ID"
            Set rsTmp = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lngSampleId)
        Else
134         strSql = "select ID ,标本ID,图像类型 from 检验图像结果 where 标本ID = [1] order by ID"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "GetSampleImage", lngSampleId)
        End If
    
136
138     If rsTmp.EOF Then Call ErrLog("GetSampleImages", CStr(Erl()) & "行", "无图像数据！", "")
        
    
140     Do Until rsTmp.EOF
            If Not gblnNewLis Then
                strFilename = rsTmp("ID") & strFileType
                'strFilename = rsTmp("标本ID") & IIf(Trim$("" & rsTmp("图像类型")) <> "", "_" & Trim$("" & rsTmp("图像类型")), "") & strFileType
            Else
                strFilename = rsTmp("标本ID") & IIf(Trim$("" & rsTmp("图像类型")) <> "", "_" & Trim$("" & rsTmp("图像类型")), "") & strFileType
            End If
142         If blnClearOldFile = True Then
                '强制清除缓存在本地的图片
144             If objFso.FileExists(gstrFilePath & "\" & strFilename) Then
                    Call objFso.DeleteFile(gstrFilePath & "\" & strFilename, True)
                    Call ErrLog("强制清除本地缓存的图片", gstrFilePath, strFilename, "")
                End If
            End If
        
146         If Not objFso.FileExists(gstrFilePath & "\" & strFilename) Then
        
148             If LoadImageData(gstrFilePath, rsTmp("ID"), 1, strFilename) Then
150                 If objFso.FileExists(gstrFilePath & "\" & strFilename) Then strReturn = strReturn & "#" & strFilename
                End If
            
152             DrawIndex = DrawIndex + 1
            Else
                strReturn = strReturn & "#" & strFilename
            End If
154         rsTmp.MoveNext
        Loop
156     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
158     GetSampleImages = strReturn
        Call ErrLog("返回图形文件名", "", strReturn, "")
        'Debug.Print "ID=" & lngKeyID & ",用时:" & DateDiff("s", strTime, Now)
        Set objFso = Nothing
        Exit Function
ErrH:
160    GetSampleImages = ""
162
164    Call ErrLog("GetSampleImages", CStr(Erl()) & "行", err.Description, "")
End Function

Public Function GetImage(ByVal lngImgID As Long, ByVal blnClearOldFile As Boolean, Optional intSaveImgType As Integer = 0) As String
        '获取指定的一张图片。
        '
        'lngImgID    : 图片ID
        'strPath     : 图片保存路径,不存在则自动创建
        'blnClearOldFile :是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
        'strErr      : 函数返回值为空串时，返回的提示信息
        'intImgType  : 返回的图片文件格式，0－cht(默认),1-jgp,2-png
        
        Dim strFileType As String
        Dim strReturn   As String
        Dim rsTmp       As ADODB.Recordset
        Dim strFilename As String
        Dim objFso      As New FileSystemObject
        Dim strSql      As String
        
104     If intSaveImgType = 1 Then
106         strFileType = ".jpg"
108     ElseIf intSaveImgType = 2 Then
110         strFileType = ".png"
        Else
112         strFileType = ".cht"
        End If

116     If Trim(gstrFilePath) = "" Then
118         Call ErrLog("GetImage", CStr(Erl()) & "行", "请指定保存图片的路径！", "")
            Exit Function
120     ElseIf Not objFso.FolderExists(gstrFilePath) Then
122         Call ErrLog("GetImage", CStr(Erl()) & "行", gstrFilePath & "不存在！", "")
            Exit Function
        End If
124     If Right(gstrFilePath, 1) = "\" Then gstrFilePath = Left(gstrFilePath, Len(gstrFilePath) - 1)
        On Error GoTo ErrH
    
126     If gblnNewLis Then
128         strSql = "select ID ,标本ID,图像类型 from 检验报告图像 where ID = [1]"
            Set rsTmp = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName, lngImgID)
        Else
130         strSql = "select ID ,标本ID,图像类型 from 检验图像结果 where ID = [1]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "GetImage", lngImgID)
        End If
    
132
134     If rsTmp.EOF Then Call ErrLog("GetImage", CStr(Erl()) & "行", "无图像数据！", "")
        If Not gblnNewLis Then
            strFilename = rsTmp("ID") & strFileType
        Else
            strFilename = rsTmp("ID") & IIf(Trim$("" & rsTmp("图像类型")) <> "", "_" & Trim$("" & rsTmp("图像类型")), "") & strFileType
        End If
136     If blnClearOldFile = True Then
            '强制清除缓存在本地的图片
138         If objFso.FileExists(gstrFilePath & "\" & strFilename) Then objFso.DeleteFile gstrFilePath & "\" & strFilename, True
            Call ErrLog("强制清除本地缓存的图片", gstrFilePath, strFilename, "")
        End If
    
140     If Not objFso.FileExists(gstrFilePath & "\" & strFilename) Then
142         If LoadImageData(gstrFilePath, rsTmp("ID"), intSaveImgType, strFilename) Then
144             If objFso.FileExists(gstrFilePath & "\" & strFilename) Then strReturn = strReturn & "," & strFilename
            Else
                Call ErrLog("GetImage", CStr(Erl()) & "行", "产生图形文件失败！", "")
            End If
        Else
            strReturn = strReturn & "," & strFilename
        End If
146     If strReturn <> "" Then strReturn = Mid(strReturn, 2)
148     GetImage = strReturn
        Call ErrLog("返回图形文件名", "", strReturn, "")
        Set objFso = Nothing
        Exit Function
ErrH:
150     GetImage = ""
152
154     Call ErrLog("GetImage", CStr(Erl()) & "行", err.Description, "")
End Function

Public Function LocalImageSaveToCht(ByVal strTitle As String, ByVal strImageData As String, ByVal strSaveFileName As String) As Boolean
    '将本地图形文件生成Cht放入指定目录
    Dim strPath     As String
    Dim strChtName  As String
    Dim objFso      As New FileSystemObject
    
    strPath = App.Path & "\LisImage"
    If Not objFso.FolderExists(strPath) Then objFso.CreateFolder (strPath)
    strChtName = strPath & "\" & strSaveFileName
    If objFso.FileExists(strChtName) Then objFso.DeleteFile strChtName, True
    
    LocalImageSaveToCht = DrawImgAndSaveFile(strTitle, strImageData, strChtName, 0)
    Set objFso = Nothing
End Function





