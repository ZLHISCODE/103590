Attribute VB_Name = "mdlPublic"
Option Explicit

Private Declare Function WritePrivateProfileString _
                  Lib "kernel32" Alias "WritePrivateProfileStringA" _
                  (ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpString As Any, _
                  ByVal lpFileName As String) As Long
                  
Private Declare Function GetPrivateProfileString _
                  Lib "kernel32" Alias "GetPrivateProfileStringA" _
                  (ByVal lpApplicationName As String, _
                  ByVal lpKeyName As Any, _
                  ByVal lpDefault As String, _
                  ByVal lpReturnedString As String, _
                  ByVal nSize As Long, _
                  ByVal lpFileName As String) As Long
                  
Private mstrFileName As String

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If Err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function DynamicGet(ByVal strclass As String, ByVal strCaption As String) As Object
    On Error Resume Next
    Set DynamicGet = GetObject("", strclass)
    
    If Err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicGet = Nothing
    End If
    Err.Clear
End Function

Public Sub ResizeImg(ByVal objImg As Object, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
'功能：根据指定的位置大小显示img图片，避免图片控件超出配置文件中指定的范围

'objImg.Tag:实际图片的宽高比
    Dim dblrH As Double, dblrW As Double  '图片相对于img控件的宽和高的比例
    Dim dblW As Double, dblH As Double
    
    objImg.Tag = Format(objImg.Picture.Width / objImg.Picture.Height, "0.0000000000")
    
    dblW = objImg.Width
    dblH = objImg.Height
    
    If Val(objImg.Tag) > 1 Then         '当宽大于高时，根据宽算出高
        dblH = dblW / Val(objImg.Tag)
    Else                                '当高大于宽时，根据高算出宽
        dblW = dblH * Val(objImg.Tag)
    End If
 
    dblrH = Format(lngHeight / dblH, "0.0000000000")
    dblrW = Format(lngWidth / dblW, "0.0000000000")
    
    If dblrH > dblrW Then
        objImg.Width = dblrW * dblW
        objImg.Height = objImg.Width / Val(objImg.Tag)
    Else
        objImg.Height = dblrH * dblH
        objImg.Width = objImg.Height * Val(objImg.Tag)
    End If
    
    objImg.Left = lngLeft + lngWidth / 2 - objImg.Width / 2
    objImg.Top = lngTop + lngHeight / 2 - objImg.Height / 2
End Sub

Public Sub SetControlFont(ByVal objControl As Object, strFontProperty() As String)
'设置各个Label控件字体
    Dim i As Integer
    
    For i = 0 To UBound(strFontProperty)
        Select Case Split(strFontProperty(i), ":")(0)
            Case "字体"
                objControl.FontName = Split(strFontProperty(i), ":")(1)
            Case "字号"
                objControl.FontSize = Split(strFontProperty(i), ":")(1)
            Case "粗体"
                objControl.FontBold = CBool(Split(strFontProperty(i), ":")(1))
            Case "斜体"
                objControl.FontItalic = CBool(Split(strFontProperty(i), ":")(1))
            Case "下划线"
                objControl.FontUnderline = CBool(Split(strFontProperty(i), ":")(1))
            Case "前景色"
                objControl.ForeColor = Split(strFontProperty(i), ":")(1)
        End Select
    Next
End Sub

Public Sub SetVSFListFont(ByVal objControl As Object, ByVal lngRow As Long, strFontProperty() As String)
'设置列表某一行的字体
    Dim i As Integer
    
    If lngRow < 0 Then Exit Sub
    
    For i = 0 To UBound(strFontProperty)
        Select Case Split(strFontProperty(i), ":")(0)
            Case "字体"
                objControl.Cell(flexcpFontName, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
            Case "字号"
                objControl.Cell(flexcpFontSize, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
            Case "粗体"
                objControl.Cell(flexcpFontBold, lngRow, 0, lngRow, objControl.Cols - 1) = CBool(Split(strFontProperty(i), ":")(1))
            Case "前景色"
                objControl.Cell(flexcpForeColor, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
        End Select
    Next
End Sub

Public Sub LoadPictureInfo(ByVal objControl As Object, ByVal strPic As String)
'加载图片到指定控件
    Dim strFileName As String
    Dim arrByte() As Byte
    Dim lngCount As Long
On Error GoTo ErrorHand
    If strPic = "" Then Exit Sub
    
    '生成临时图片文件
    strFileName = App.Path & "\imgTemp.jpg"

    '16进制转换为字节数组
    ReDim arrByte(Len(strPic) / 2 - 1) As Byte
    For lngCount = LBound(arrByte) To UBound(arrByte)
        arrByte(lngCount) = CByte("&H" & Mid(strPic, lngCount * 2 + 1, 2))
    Next
    Open strFileName For Binary As #1
    Put #1, , arrByte()
    Close #1
    
    '加载图片
    objControl.Picture = LoadPicture(strFileName)
    
    '删除临时图片文件
    If gobjFile.FileExists(strFileName) Then gobjFile.DeleteFile (strFileName)
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'截取指定大小的位图
Public Function CutPicture(ByVal strImgPath As String, ByVal objResource As Object, ByVal sinleft As Single, ByVal sinTop As Single, ByVal sinWidth As Single, ByVal sinHeight As Single) As StdPicture
'objResource：图片资源
'lngRows：图片总行数
'lngShowNum:截取的行数
    objResource.Visible = False
    objResource.AutoRedraw = True
    objResource.AutoSize = True
    
    If sinleft <= 0 Then sinleft = 0
    If sinTop <= 0 Then sinTop = 0
    If sinHeight <= 0 Or sinWidth <= 0 Then Exit Function
    
    '获取本地图片
    objResource = LoadPicture(strImgPath)

    Call objResource.PaintPicture(objResource, 0, 0, objResource.ScaleWidth, objResource.ScaleHeight, _
    sinleft, sinTop, sinWidth, sinHeight, vbSrcCopy)
    
    Set CutPicture = objResource.Image

    '释放图片资源
    objResource = LoadPicture("")
End Function

Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'获取加密密码
Public Function getEncryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Long
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim strRandom As String
    Dim strBase As String
        
    i = 0
    
    lngPassWLength = Len(strPassW)
    
    strBase = GetRandom(20)
    strRandom = GetRandom(20)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
     
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassW, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strBase) Xor Asc(strRandom)
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '加密后的字串
End Function

'获取解密密码
Public Function getDecryptionPassW(ByVal strPassW As String) As String
    Dim i As Integer
    Dim lngAsc  As Integer
    Dim strTemp() As String
    Dim lngPassWLength As Integer
    Dim lngBase As Long
    Dim strRandom As String
    Dim strPassSouce As String

    i = 0
    
    strPassSouce = Mid(strPassW, 2, Len(strPassW) - 2)
    lngPassWLength = Len(strPassSouce)
    lngBase = Asc(Mid(strPassW, 1, 1))
    
    strRandom = Right(strPassW, 1)
    
    ReDim intAsc(0 To lngPassWLength - 1), strTemp(0 To lngPassWLength - 1)
    
    Do While i < lngPassWLength
        lngAsc = Asc(Mid(strPassSouce, i + 1, 1))
        lngAsc = lngAsc Xor Asc(strRandom) Xor lngBase
        strTemp(i) = Chr(lngAsc)
        i = i + 1
    Loop

    getDecryptionPassW = Join(strTemp, "") '解密后的字串
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function AnalyseComputer() As String
'获取计算机名
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FileName：Ini文件
'PathName：小节名
'KeyName：值名
'BackValue：返回值
'Default：默认字符
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadValue(strSectionName As String, strKeyName As String, _
                          Optional strDefault As String = "") As String
  Dim lngReadState As Long
  Dim strTempNum As String
  Dim strTemp As String
            
  strTemp = String$(255, Chr$(0))
  strTempNum = 255
  
  ReadValue = strDefault
            
  lngReadState = GetPrivateProfileString(strSectionName, strKeyName, strDefault, strTemp, strTempNum, mstrFileName)
                        
  If lngReadState <> 0 Then
    ReadValue = Trim(Left$(strTemp, strTempNum))
  End If
            
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'设置INI文件名称
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIniFile(ByVal strFileName As String)
  mstrFileName = strFileName
End Sub
