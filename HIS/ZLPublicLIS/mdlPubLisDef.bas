Attribute VB_Name = "mdlPubLisDef"
Option Explicit

Public strChart(9)          As Variant
Public glngModual           As Long         'ģ���
Public gblnNewLis           As Boolean      '�Ƿ�Ϊ�°�LIS
Public gstrHospital         As String
Public gstrFilePath         As String       'ͼ�񱣴�·��
Public gstrSignPath         As String
Public gbln��ʾͼƬ         As Boolean
Public glngTop              As Long
Public glngLeft             As Long
Public gbtyModel            As Integer
Public gblnNew           As Boolean      '�Ƿ�Ϊ�°�LIS
Public Const gstrͼƬ��ʽ   As String = ".cht|.GIF|.gif|.jpg|.JPG|.bmp|.BMP|.JPEG|.jpeg|.png|.PNG"

'API ����:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

Public Enum COLOR
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
End Enum

'----- ����ΪJPG��ʽ��ͼƬ
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
End Type

'�ṹ����:
Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    type As Long
    Value As Long
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal Filename As Long, clsIDEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, ID As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Function Between(X, a, b) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж�x�Ƿ���a��b֮��
    '******************************************************************************************************************
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
    '******************************************************************************************************************
    '���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    '������
    '���أ�
    '******************************************************************************************************************
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function SeveLoadImg(ByVal lng_�걾ID As Long)
    On Error GoTo ErrH
    Dim rsTmp   As ADODB.Recordset
    Dim strSql  As String
    Dim intLoop As Integer
    
'    strSQL = "select ID from ����ͼ���� where �걾ID = [1] order by ID"
    If gblnNewLis Then
        strSql = "select ID from ���鱨��ͼ�� where �걾ID = [1] order by ID"
    Else
        strSql = "select ID from ����ͼ���� where �걾ID = [1] order by ID"
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "��ȡ����ͼ��", lng_�걾ID)

    For intLoop = 1 To 9
        strChart(intLoop) = ""
    Next
    intLoop = 1
    Do Until rsTmp.EOF
        If intLoop > 9 Then Exit Do
        strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
        Debug.Print strChart(intLoop)
        Call LoadImageData(App.Path & "\", rsTmp("ID"), 1, "")
        
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    Exit Function
ErrH:
    Call ErrLog("mdlimg", "SeveLoadImg", "ͼƬ���ش���", err.Description)
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
'Public Function zlFileUnzip(ByVal strZipFile As String) As String
'    Dim strZipPath As String
'    Dim clsUnzip As New clsUnzip
'
'    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
'    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
'
'    With clsUnzip
'        .ZipFile = strZipFile
'        .UnzipFolder = strZipPath
'        .Unzip
'    End With
'    If Dir(strZipPath) <> "" Then
'        zlFileUnzip = strZipPath & Dir(strZipPath)
'    Else
'        zlFileUnzip = ""
'    End If
'End Function
'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
'Public Function zlFileZip(ByVal strFile As String, ByVal strFilename As String) As String
'    Dim strZipFile As String, lngCount As Long
'    Dim clsZip  As New clsUnzip
'
'    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
'
'    lngCount = 0
'    Do While True
'        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
'        If Dir(strZipFile) = "" Then Exit Do
'        lngCount = lngCount + 1
'    Loop
'
'    With mclsZip
'        .Encrypt = False: .AddComment = False
'        .ZipFile = strZipFile
'        .StoreFolderNames = False
'        .RecurseSubDirs = False
'        .ClearFileSpecs
'        .AddFileSpec strFile
'        .Zip
'        If (.Success) Then
'            zlFileZip = .ZipFile
'        Else
'            zlFileZip = ""
'        End If
'    End With
'End Function

'*************************************************************************
'**�� �� ����PrintRotText
'**��    �룺ByVal hDC(Long)          -
'**        ��ByVal Text(String)       -  Ҫ��ӡ������
'**        ��ByVal CenterX(Long)      -  X���ĵ����������
'**        ��ByVal CenterY(Long)      -  Y���ĵ����������
'**        ��ByVal RotDegrees(Single) -  ��ת�Ƕ�(0.0 �� 359.9999999) ��˳ʱ�룬0=ˮƽ(����ת)
'**��    ����(Boolean) -
'**������������һ��������������X,����Y���������ԽǶȻ�����ת����
'**ȫ�ֱ�����
'**����ģ�飺
'*************************************************************************
Public Function PrintRotText(ByVal hDC As Long, ByVal Text As String, ByVal CenterX As Long, ByVal CenterY As Long, ByVal RotDegrees As Single) As Boolean

Dim bOkSoFar    As Boolean      '������ʶ.
Dim hFontOld    As Long         'ԭ������
Dim hFontNew    As Long         '��������
Dim lfFont      As LOGFONT      'LOGFONT ������ṹ.
Dim ptOrigin    As POINTAPI     '���ֻ���ԭ��
Dim ptCenter    As POINTAPI     '�������ĵ�.
Dim szText      As SizeStruct   '���ֿ�Ⱥ͸߶�

    '���豸�еõ���ǰ LOGFONT �ṹ.
    hFontOld = SelectObject(hDC, GetStockObject(SYSTEM_FONT))
    
    '������豸�õ�������ɹ�...
    If hFontOld <> 0 Then
        
        '�������ȡ LOGFONT �ṹ
        bOkSoFar = (GetObjectAPI(hFontOld, Len(lfFont), lfFont) <> 0)
        
        '��ԭ��������
        Call SelectObject(hDC, hFontOld)
        
        '��λ�Ժ�ʹ��
        hFontOld = 0
    End If
    
    '����ɹ���� LOGFONT �ṹ������.
    If bOkSoFar Then
    
        '�ı����巽��ͳ���
        lfFont.lfEscapement = RotDegrees * 10
        lfFont.lfOrientation = lfFont.lfEscapement
        lfFont.lfQuality = ANTIALIASED_QUALITY
        
        '�� LOGFONT �ṹ�д������������
        hFontNew = CreateFontIndirect(lfFont)
        
        '���崴���ɹ�
        If hFontNew <> 0 Then
            
            'Select the neѡ���µ����嵽���豸
            hFontOld = SelectObject(hDC, hFontNew)
            
            '�ɹ�
            If hFontOld <> 0 Then
                
                '��ȡ�����߼���λ��С(����)
                bOkSoFar = (GetTextExtentPoint32(hDC, Text, LenB(StrConv(Text, vbFromUnicode)), szText) <> 0)
                
                '�ɹ�
                If bOkSoFar Then
                    
                    '��������ˮƽԭ��
                    With ptOrigin
                        .X = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    'ת�� CenterX, CenterY ����ṹ
                    '(��Ҫ���� RotatePoint).
                    With ptCenter
                        .X = CenterX
                        .Y = CenterY
                    End With
                    
                    '��ԭ��ѡ����ƥ��Ԥ��ѡ��
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    '���ڴ�ӡ��ת�ı������سɹ�/ʧ��
                    PrintRotText = (TextOut(hDC, ptOrigin.X, _
                      ptOrigin.Y, Text, LenB(StrConv(Text, vbFromUnicode))) <> 0)
                
                End If
                
                '�ָ����嵽ԭ���豸
                hFontNew = SelectObject(hDC, hFontOld)
            
            End If
            
            '����ڴ沢ɾ������������
            Call DeleteObject(hFontNew)
        
        End If
        
    End If
            
End Function

'*************************************************************************
'**    ��    �� ��    laviewpbt
'**    �� �� �� ��    SavePic
'**    ��    �� ��    pic(StdPicture)        -   ͼ����
'**             ��    FileName(String)       -   ����·��
'**             ��    Quality(Byte)          -   JPGͼ������
'**             ��    TIFF_ColorDepth(Long)  -   TTF��ʽ����ɫ���
'**             ��    TIFF_Compression(Long) -   TTF��ʽ��ѹ����
'**    ��    �� ��    ��
'**    �������� ��    ��ͼ�󱣴�ΪJPG��TIFF��PNG��GIF��BMP��ʽ
'**    ��    �� ��
'**    �� �� �� ��    laviewpbt
'**    ��    �� ��    2005-10-23 14.43.52
'**    ��    �� ��    Version 1.2.1
'*************************************************************************
Public Sub SavePic(ByVal pict As StdPicture, ByVal Filename As String, PicType As String, _
                    Optional ByVal Quality As Byte = 100, _
                    Optional ByVal TIFF_ColorDepth As Long = 24, _
                    Optional ByVal TIFF_Compression As Long = 6)
100    Screen.MousePointer = vbHourglass
       Dim tSI As GdiplusStartupInput
       Dim lRes As Long
       Dim lGDIP As Long
       Dim lBitmap As Long
       Dim aEncParams() As Byte
       On Error GoTo errHandle:
102    tSI.GdiplusVersion = 1   ' ��ʼ�� GDI+
104    lRes = GdiplusStartup(lGDIP, tSI)
106    If lRes = 0 Then     ' �Ӿ������ GDI+ ͼ��
108       lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
110       If lRes = 0 Then
             Dim tJpgEncoder As GUID
             Dim tParams As EncoderParameters    '��ʼ����������GUID��ʶ
112          Select Case UCase(PicType)
             Case ".JPG", "JPG", ".JPEG", "JPEG"
114             CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
116             tParams.count = 1                               ' ���ý���������
118             With tParams.Parameter ' Quality
120                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID    ' �õ�Quality������GUID��ʶ
122                .NumberOfValues = 1
124                .type = 4
126                .Value = VarPtr(Quality)
                End With
128             ReDim aEncParams(1 To Len(tParams))
130             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
132         Case ".PNG", "PNG"
134              CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
136              ReDim aEncParams(1 To Len(tParams))
138         Case ".GIF", "GIF"
140              CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
142              ReDim aEncParams(1 To Len(tParams))
144         Case ".TIFF", "TIFF"
146              CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
148              tParams.count = 2
150              ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
152              With tParams.Parameter
154                 .NumberOfValues = 1
156                 .type = 4
158                  CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID    ' �õ�ColorDepth������GUID��ʶ
160                 .Value = VarPtr(TIFF_Compression)
                End With
162             Call CopyMemory(aEncParams(1), tParams, Len(tParams))
164             With tParams.Parameter
166                 .NumberOfValues = 1
168                 .type = 4
170                  CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID    ' �õ�Compression������GUID��ʶ
172                 .Value = VarPtr(TIFF_ColorDepth)
                End With
174             Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
176         Case ".BMP", "BMP"                                              '������ǰд����ΪBMP�Ĵ��룬��Ϊ��û����GDI+
178             SavePicture pict, Filename
180             Screen.MousePointer = vbDefault
                Exit Sub
            End Select
182          lRes = GdipSaveImageToFile(lBitmap, StrPtr(Filename), tJpgEncoder, aEncParams(1))             '����ͼ��
184          GdipDisposeImage lBitmap       ' ����GDI+ͼ��
          End If
186       GdiplusShutdown lGDIP              '���� GDI+
       End If
188    Screen.MousePointer = vbDefault
190    Erase aEncParams
       Exit Sub
errHandle:
192     Screen.MousePointer = vbDefault
194     WriteLog "mdlPublic.SavePic", CStr(Erl()) & "��", err.Description
End Sub


'*************************************************************************
'**�� �� ����RotatePoint
'**��    �룺ptAxis(PointAPI)   -
'**        ��ptRotate(PointAPI) -
'**        ��fDegrees(Single)   -
'**��    ������
'**������������ǰfdegrees��ǰ����ѡ��ptRotate���ҵ�ptAxis
'**ȫ�ֱ�����
'**����ģ�飺
'*************************************************************************
Private Sub RotatePoint(ptAxis As POINTAPI, ptRotate As POINTAPI, fDegrees As Single)

' ***************************************************
' *                 RotatePoint                     *
' *                                                 *
' *  Created by: Rocky Clark (Kath-Rock Software)   *
' *                                                 *
' *  Rotate ptRotate around ptAxis, fDegrees from   *
' *  its current position.                          *
' *                                                 *
' * This procedure may be used and distributed, as  *
' * is, in your code, as long as these credits and  *
' * the code itself remain unchanged.               *
' *                                                 *
' ***************************************************

Dim fDX     As Single   'X����
Dim fDY     As Single   'Y����
Dim fRads   As Single   '����
Const dPi   As Double = 3.14159265358979  'Pi Բ����


    'ת���Ƕ�Ϊ����
    fRads = fDegrees * (dPi / 180#)
    
    '�����ĵ�������
    fDX = ptRotate.X - ptAxis.X
    fDY = ptRotate.Y - ptAxis.Y
    
    '��ת��
    ptRotate.X = ptAxis.X + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub



Public Function DeleteImge()

On Error GoTo ErrH
    Dim strFilename         As String
    Dim intLoop             As Integer
    Dim objFso              As New FileSystemObject
    Dim objFile             As File
    Dim strCreateTime       As String
    Dim strLastVisitTime    As String
    Dim strLastModityTime   As String
    Dim intLostDay          As Integer
    Dim currDate            As Date
    
    intLostDay = 20

    For intLoop = LBound(Split(gstrͼƬ��ʽ, "|")) To UBound(Split(gstrͼƬ��ʽ, "|"))
        strFilename = Dir(gstrFilePath & "\*" & Split(gstrͼƬ��ʽ, "|")(intLoop), vbNormal)  ' ��Ѱ��һ�
        
        Do While strFilename <> ""   ' ��ʼѭ����
            ' ������ǰ��Ŀ¼���ϲ�Ŀ¼��
            If strFilename <> ".." Then
                Set objFile = objFso.GetFile(gstrFilePath & "\" & strFilename)
                strCreateTime = Format(objFile.DateCreated(), "yyyy-MM-dd")
                strLastVisitTime = Format(objFile.DateLastAccessed(), "yyyy-MM-dd")
                strLastModityTime = Format(objFile.DateLastModified(), "yyyy-MM-dd")
                If CDate(strCreateTime) < CDate(Date - intLostDay) Then
                    Kill gstrFilePath & "\" & strFilename
                End If
            End If
            strFilename = Dir   ' ������һ��Ŀ¼��
        Loop
    Next
    Set objFile = Nothing
    Set objFso = Nothing
    Exit Function
ErrH:
    WriteLog "deleteimge", err.Description, ""
End Function

Public Sub ErrLog(strObj As String, strEvent As String, strErrNum As String, strErrDesc As String)
On Error GoTo ErrH
    '��������Ϣд���ļ���
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strFile As String
    If strErrNum = "9999" Then
        strFile = App.Path & "\���鱨���ӡ\������־\ErrLog" & Format(Date, "YYYYMMDD") & ".txt"
    Else
        strFile = App.Path & "\���鱨���ӡ\������־\ErrLog" & Format(Date, "YYYYMMDD") & ".Log"
    End If
    If Not objFile.FolderExists(App.Path & "\���鱨���ӡ") Then
        objFile.CreateFolder (App.Path & "\���鱨���ӡ")
    End If
    If Not objFile.FolderExists(App.Path & "\���鱨���ӡ\������־") Then
        objFile.CreateFolder (App.Path & "\���鱨���ӡ\������־")
    End If
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Set objText = objFile.OpenTextFile(strFile, ForAppending)
    objText.WriteLine "�������" & strObj
    objText.WriteLine "�¼�����" & strEvent
    objText.WriteLine "����ţ�" & strErrNum
    objText.WriteLine "����������" & strErrDesc
    objText.Close
    Set objText = Nothing
    Set objFile = Nothing
    Exit Sub
ErrH:
    err.Clear
    Exit Sub
End Sub
