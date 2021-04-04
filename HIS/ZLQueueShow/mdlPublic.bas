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
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If Err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicCreate = Nothing
    End If
    Err.Clear
End Function

Public Function DynamicGet(ByVal strclass As String, ByVal strCaption As String) As Object
    On Error Resume Next
    Set DynamicGet = GetObject("", strclass)
    
    If Err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, "��ʾ"
        Set DynamicGet = Nothing
    End If
    Err.Clear
End Function

Public Sub ResizeImg(ByVal objImg As Object, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long)
'���ܣ�����ָ����λ�ô�С��ʾimgͼƬ������ͼƬ�ؼ����������ļ���ָ���ķ�Χ

'objImg.Tag:ʵ��ͼƬ�Ŀ�߱�
    Dim dblrH As Double, dblrW As Double  'ͼƬ�����img�ؼ��Ŀ�͸ߵı���
    Dim dblW As Double, dblH As Double
    
    objImg.Tag = Format(objImg.Picture.Width / objImg.Picture.Height, "0.0000000000")
    
    dblW = objImg.Width
    dblH = objImg.Height
    
    If Val(objImg.Tag) > 1 Then         '������ڸ�ʱ�����ݿ������
        dblH = dblW / Val(objImg.Tag)
    Else                                '���ߴ��ڿ�ʱ�����ݸ������
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
'���ø���Label�ؼ�����
    Dim i As Integer
    
    For i = 0 To UBound(strFontProperty)
        Select Case Split(strFontProperty(i), ":")(0)
            Case "����"
                objControl.FontName = Split(strFontProperty(i), ":")(1)
            Case "�ֺ�"
                objControl.FontSize = Split(strFontProperty(i), ":")(1)
            Case "����"
                objControl.FontBold = CBool(Split(strFontProperty(i), ":")(1))
            Case "б��"
                objControl.FontItalic = CBool(Split(strFontProperty(i), ":")(1))
            Case "�»���"
                objControl.FontUnderline = CBool(Split(strFontProperty(i), ":")(1))
            Case "ǰ��ɫ"
                objControl.ForeColor = Split(strFontProperty(i), ":")(1)
        End Select
    Next
End Sub

Public Sub SetVSFListFont(ByVal objControl As Object, ByVal lngRow As Long, strFontProperty() As String)
'�����б�ĳһ�е�����
    Dim i As Integer
    
    If lngRow < 0 Then Exit Sub
    
    For i = 0 To UBound(strFontProperty)
        Select Case Split(strFontProperty(i), ":")(0)
            Case "����"
                objControl.Cell(flexcpFontName, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
            Case "�ֺ�"
                objControl.Cell(flexcpFontSize, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
            Case "����"
                objControl.Cell(flexcpFontBold, lngRow, 0, lngRow, objControl.Cols - 1) = CBool(Split(strFontProperty(i), ":")(1))
            Case "ǰ��ɫ"
                objControl.Cell(flexcpForeColor, lngRow, 0, lngRow, objControl.Cols - 1) = Split(strFontProperty(i), ":")(1)
        End Select
    Next
End Sub

Public Sub LoadPictureInfo(ByVal objControl As Object, ByVal strPic As String)
'����ͼƬ��ָ���ؼ�
    Dim strFileName As String
    Dim arrByte() As Byte
    Dim lngCount As Long
On Error GoTo ErrorHand
    If strPic = "" Then Exit Sub
    
    '������ʱͼƬ�ļ�
    strFileName = App.Path & "\imgTemp.jpg"

    '16����ת��Ϊ�ֽ�����
    ReDim arrByte(Len(strPic) / 2 - 1) As Byte
    For lngCount = LBound(arrByte) To UBound(arrByte)
        arrByte(lngCount) = CByte("&H" & Mid(strPic, lngCount * 2 + 1, 2))
    Next
    Open strFileName For Binary As #1
    Put #1, , arrByte()
    Close #1
    
    '����ͼƬ
    objControl.Picture = LoadPicture(strFileName)
    
    'ɾ����ʱͼƬ�ļ�
    If gobjFile.FileExists(strFileName) Then gobjFile.DeleteFile (strFileName)
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

'��ȡָ����С��λͼ
Public Function CutPicture(ByVal strImgPath As String, ByVal objResource As Object, ByVal sinleft As Single, ByVal sinTop As Single, ByVal sinWidth As Single, ByVal sinHeight As Single) As StdPicture
'objResource��ͼƬ��Դ
'lngRows��ͼƬ������
'lngShowNum:��ȡ������
    objResource.Visible = False
    objResource.AutoRedraw = True
    objResource.AutoSize = True
    
    If sinleft <= 0 Then sinleft = 0
    If sinTop <= 0 Then sinTop = 0
    If sinHeight <= 0 Or sinWidth <= 0 Then Exit Function
    
    '��ȡ����ͼƬ
    objResource = LoadPicture(strImgPath)

    Call objResource.PaintPicture(objResource, 0, 0, objResource.ScaleWidth, objResource.ScaleHeight, _
    sinleft, sinTop, sinWidth, sinHeight, vbSrcCopy)
    
    Set CutPicture = objResource.Image

    '�ͷ�ͼƬ��Դ
    objResource = LoadPicture("")
End Function

Private Function GetRandom(ByVal lngBase As Long) As String
    Dim lngNum As Long
    
    Randomize 99
    
    lngNum = Fix(Rnd * lngBase)
    
    If lngNum <= 0 Then lngNum = 1
    
    GetRandom = Chr(lngNum)
End Function

'��ȡ��������
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
    
    getEncryptionPassW = strBase & Join(strTemp, "") & strRandom '���ܺ���ִ�
End Function

'��ȡ��������
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

    getDecryptionPassW = Join(strTemp, "") '���ܺ���ִ�
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function AnalyseComputer() As String
'��ȡ�������
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FileName��Ini�ļ�
'PathName��С����
'KeyName��ֵ��
'BackValue������ֵ
'Default��Ĭ���ַ�
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
'����INI�ļ�����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIniFile(ByVal strFileName As String)
  mstrFileName = strFileName
End Sub
