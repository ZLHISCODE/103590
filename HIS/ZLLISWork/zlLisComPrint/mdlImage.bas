Attribute VB_Name = "mdlImage"
Option Explicit
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Private mobjImg As Object                            'zllisdev.dll对象

Public Function ReadSampleImage(lngSampleID As Long, strChar() As String, Optional strErr As String) As Boolean
    '功能   读入标本的图像返回读出的数组
    '读图像
    Dim rsImage As ADODB.Recordset
    Dim intLoop As Integer, strReturn As String
    Dim varTmp As Variant, strDir As String
    Dim i As Integer
    
    
    On Error GoTo errH
    
    
    strErr = ""
    strDir = App.Path & "\LisImage"
    If Not gobjFSO.FolderExists(strDir) Then Call gobjFSO.CreateFolder(strDir)
    
    If mobjImg Is Nothing Then
        Set mobjImg = CreateObject("zlLisDev.clsDrawGraph")
        Call mobjImg.GetSampleImgInit(gSysInfo.SysNo, gcnOracle, strErr)
        
        If strErr <> "" Then
            Exit Function
        End If
    End If
    '标本ID
    '图片保存路径(不存在则自动创建),
    '是否清空缓存在本地的图形文件,True－每次都从数据库读文件保存到本地;False-第一次调用时从数据库读图形产生图片，之后直接使用
    '函数返回值为空串时，返回的提示信息
    '返回的图片文件格式，0－cht(默认),1-jgp,2-png
    '是新版LIS还是老版LIS在调用本函件数， 0-老版LIS（默认，从“检验图像结果”中取图形数据），1-新版LIS（从“检验报告图像”中取图形数据）
    
    strReturn = mobjImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 0)
    If strReturn = "" Then
        If strErr = "无图像数据！" Then
            strErr = ""
            ReadSampleImage = True
        ElseIf strErr = "" Then
            ReadSampleImage = True
        End If
        Exit Function
    End If
    
    varTmp = Split(strReturn, ",")

    For i = LBound(varTmp) To UBound(varTmp)
        If i > 8 Then Exit For
        If Trim("" & varTmp(i)) <> "" Then
            If Dir(strDir & "\" & Trim("" & varTmp(i))) <> "" Then strChar(i) = strDir & "\" & Trim("" & varTmp(i))
        End If
    Next
    
    ReadSampleImage = True
    Exit Function
errH:
    strErr = "出错函数(ReadSampleImage),出错信息:" & Err.Number & " " & Err.Description
End Function

Public Sub FreeImageObj()
    Dim strErr As String
    If Not mobjImg Is Nothing Then
        Call mobjImg.GetSampleImgExit(strErr)
        Set mobjImg = Nothing
    End If
End Sub


