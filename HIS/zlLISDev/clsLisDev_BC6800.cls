VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BC6800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ѪҺ������ BC6800
'����:����(Mindary)
'����:���
'�û�:����ʡ��ɽ����ҽԺ

Private mstrCode As String  '����������
Const mstrItem  As String = ",WBC,BAS#,BAS%,NEU#,NEU%,EOS#,EOS%,LYM#,LYM%,MON#,MON%,RBC,HGB,MCV,MCH,MCHC,RDW-CV,RDW-SD,HCT,PLT,MPV,PDW,PCT,PLCR,PLCC,IMG#,IMG%,HFC#,HFC%,PLT-I,WBC-D,WBC-B,PDW-SD,InR#,InR%,RET%,RET#,IRF,LFR,MFR,HFR,NRBC#,NRBC%,"
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0
Private Const BITMAPTYPE = &H4D42
Private Const INVALID_HANDLE_VALUE = (-1)
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long
                             
Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
                             
Private Declare Function SetPixelV _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long
                             
Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long) As Long
                             
Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     ByVal lpSecurityAttributes As Long, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long
                                     
Private Declare Function WriteFile _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpBuffer As Any, _
                                ByVal nNumberOfBytesToWrite As Long, _
                                lpNumberOfBytesWritten As Long, _
                                ByVal lpOverlapped As Long) As Long
                                
Private Declare Function CreateDIBSection _
                Lib "gdi32" (ByVal hDC As Long, _
                             pBitmapInfo As BITMAPINFO, _
                             ByVal un As Long, _
                             lplpVoid As Long, _
                             ByVal handle As Long, _
                             ByVal dw As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



'--------------------------------------------------------------------------------------------------------------------
'�����Ǳ���Ľӿں���
Public Function GetAnswerCmd() As String
    '�Զ�Ӧ��ָ��
    GetAnswerCmd = Chr(&H6)
End Function

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    Dim str_Code As String '1����������Ϣ
    Dim var_Code As Variant
    Dim lng_count As Long
    
    Dim str_�ֶηָ��� As String
    Dim str_����ָ��� As String
    Dim str_��������� As String
    Dim str_�ظ��ָ��� As String
    Dim str_ת���ַ� As String
    
    Dim str_Segment As String 'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    Dim str_Ψһ��ʶ As String '���ڱ�ʶһ����Ϣ
    Dim str_�汾 As String
    
    Dim lng_S As Long
    Dim lng_E As Long
    
    Dim str_�걾�� As String, str_���� As String
    Dim str_�ʿ� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String
    
    Dim str_Base64 As String, str_WBC As String, str_RBC As String, str_PLT As String, str_BASO As String, str_Img As String, str_RET As String, str_NRBC As String
    
    On Error GoTo errHandle
    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode & strSource
    End If
    
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        
        'Call writelislog("ԭʼ����:" & str_Code)
                 
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '��Ϣͷ
                str_�ֶηָ��� = Mid(str_Segment, 5, 1)
                str_����ָ��� = Mid(str_Segment, 6, 1)
                str_��������� = Mid(str_Segment, 7, 1)
                str_�ظ��ָ��� = Mid(str_Segment, 8, 1)
                str_ת���ַ� = Mid(str_Segment, 9, 1)
                
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_Ψһ��ʶ = var_Filed(9)
                str_�ʿ� = var_Filed(10)
                str_�汾 = var_Filed(11)
                If str_�汾 <> "2.3.1" Then
                    WriteLog "�汾����", "", ""
                End If
                str_Return = "": str_�걾�� = "": str_���� = ""
            ElseIf InStr(str_Segment, "OBR|") > 0 Then
                '�걾�ţ�����
                If Val(str_�걾��) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
                    str_Img = ""
                    If str_WBC <> "" Then str_Img = str_Img & "^DIFF;103;1;" & str_WBC
                    If str_BASO <> "" Then str_Img = str_Img & "^BASO;103;1;" & str_BASO
                    If str_RBC <> "" Then str_Img = str_Img & "^RBC;103;1;" & str_RBC
                    If str_PLT <> "" Then str_Img = str_Img & "^PLT;103;1;" & str_PLT
                    If str_RET <> "" Then str_Img = str_Img & "^RET;103;1;" & str_RET
                    If str_NRBC <> "" Then str_Img = str_Img & "^NRBC;103;1;" & str_NRBC
                    If str_Img <> "" Then strResult = strResult & vbCrLf & Mid(str_Img, 2)
                End If
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_�걾�� = var_Filed(3)
                str_���� = var_Filed(7)
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
                str_Return = ""
            ElseIf InStr(str_Segment, "OBX|") > 0 Then
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_��Ŀ = var_Filed(3)
                If InStr(mstrItem, "," & Split(str_��Ŀ, str_����ָ���)(1) & ",") > 0 Then
                    str_Return = str_Return & "|" & Split(str_��Ŀ, str_����ָ���)(1) & "|" & Trim(var_Filed(5))
                ElseIf InStr(",15056,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'RBC Histogram(ֱ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_RBC = SaveBase64Bmp("RBC", str_�걾��, str_Base64)
                ElseIf InStr(",15116,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'PLT Histogram(ֱ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_PLT = SaveBase64Bmp("PLT", str_�걾��, str_Base64)
                ElseIf InStr(",15200,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'WBC DIFF(ɢ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_WBC = SaveBase64Bmp("DIFF", str_�걾��, str_Base64)
                ElseIf InStr(",15250,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'Baso Scattergram(ɢ��ͼ)
                    str_Base64 = Split(var_Filed(5), str_����ָ���)(4)
                    str_BASO = SaveBase64Bmp("BASO", str_�걾��, str_Base64)
                ElseIf InStr(",15300,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'RET Scattergram
                    str_Base64 = Split(var_Filed(5), str_����ָ���)(4)
                    str_RET = SaveBase64Bmp("RET", str_�걾��, str_Base64)
                ElseIf InStr(",15350,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'RENRBCT Scattergram
                    str_Base64 = Split(var_Filed(5), str_����ָ���)(4)
                    str_NRBC = SaveBase64Bmp("NRBC", str_�걾��, str_Base64)
                End If
            
            End If
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
            str_Img = ""
            If str_WBC <> "" Then str_Img = str_Img & "^DIFF;103;1;" & str_WBC
            If str_BASO <> "" Then str_Img = str_Img & "^BASO;103;1;" & str_BASO
            If str_RBC <> "" Then str_Img = str_Img & "^RBC;103;1;" & str_RBC
            If str_PLT <> "" Then str_Img = str_Img & "^PLT;103;1;" & str_PLT
            If str_RET <> "" Then str_Img = str_Img & "^RET;103;1;" & str_RET
            If str_NRBC <> "" Then str_Img = str_Img & "^NRBC;103;1;" & str_NRBC
            If str_Img <> "" Then strResult = strResult & vbCrLf & Mid(str_Img, 2)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then
      strResult = Mid(strResult, 3)
    End If
    Exit Sub
errHandle:
    WriteLog "BC6800_Analyse", strSource, Err.Number & " " & Err.Description
End Sub

Public Function SendSample(ByVal strSample As String) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    SendSample = ""
End Function

'---------------------------------------------------------------------------------------------------------------
'�������ڲ�����
Private Sub TransparentBmp(ByVal strPicPath As String)
    Dim hmemDC         As Long
    Dim hmemBMP        As Long
    Dim lpmemBits      As Long
    Dim bmp_info       As BITMAPINFO
    Dim hFile          As Long
    Dim bmpfile_info   As BITMAPFILEHEADER
    Dim lpBytesWritten As Long
    Dim membmp         As Long, memdc As Long
    Dim x              As Integer, Y As Integer
    Dim cr             As Long, r As Integer, g As Integer, B As Integer
    
    frmGraph.Picture1.AutoSize = True
    frmGraph.Picture1.ScaleMode = 3 'Pixel
    If Trim(strPicPath) <> "" Then frmGraph.Picture1.Picture = LoadPicture(strPicPath) 'Load Image
    memdc = CreateCompatibleDC(frmGraph.Picture1.hDC) 'Create Compatible DC
    SelectObject memdc, frmGraph.Picture1.Picture.handle 'Select Image

    'Turn Colors:
    For x = 0 To (frmGraph.Picture1.ScaleWidth - 1)
        For Y = 0 To (frmGraph.Picture1.ScaleHeight - 1)
            cr = GetPixel(memdc, x, Y) 'Get Colors
            r = cr Mod 256
            g = (cr \ 256) Mod 256
            B = cr \ 65536
            'Reset Colors:
            SetPixelV memdc, x, Y, RGB(255 - r, 255 - g, 255 - B)
            DoEvents 'Have a rest
        Next
    Next
    'Draw Out
    BitBlt frmGraph.Picture1.hDC, 0, 0, frmGraph.Picture1.Width, frmGraph.Picture1.Height, memdc, 0, 0, vbSrcCopy

    frmGraph.Picture1.ScaleMode = vbPixels
    With bmp_info.bmiHeader
        .biSize = LenB(bmp_info.bmiHeader)
        .biWidth = frmGraph.Picture1.ScaleWidth
        .biHeight = frmGraph.Picture1.ScaleHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = .biHeight * (((.biWidth * .biBitCount + 31) And &HFFFFFFE0) \ 8)

    End With
   
    frmGraph.Picture1.Refresh
    
    hmemDC = CreateCompatibleDC(frmGraph.Picture1.hDC)
    hmemBMP = CreateDIBSection(frmGraph.Picture1.hDC, bmp_info, DIB_RGB_COLORS, lpmemBits, 0, 0)
    SelectObject hmemDC, hmemBMP
    BitBlt hmemDC, 0, 0, bmp_info.bmiHeader.biWidth, bmp_info.bmiHeader.biHeight, frmGraph.Picture1.hDC, 0, 0, vbSrcCopy
    '����ͼƬ
    hFile = CreateFile(strPicPath, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)

    If hFile <> INVALID_HANDLE_VALUE Then

        With bmpfile_info
            .bfType = BITMAPTYPE
            .bfOffBits = 14 + bmp_info.bmiHeader.biSize
            .bfSize = .bfOffBits + bmp_info.bmiHeader.biSizeImage

        End With

        WriteFile hFile, bmpfile_info.bfType, 2, lpBytesWritten, 0
        WriteFile hFile, bmpfile_info.bfSize, 12, lpBytesWritten, 0
        WriteFile hFile, bmp_info.bmiHeader, bmp_info.bmiHeader.biSize, lpBytesWritten, 0
        WriteFile hFile, ByVal lpmemBits, bmp_info.bmiHeader.biSizeImage, lpBytesWritten, 0
        CloseHandle hFile

    End If

    DeleteObject hmemBMP
    DeleteDC hmemDC
End Sub

Private Function SaveBase64Bmp(ByVal strType As String, ByVal str�걾�� As String, ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim bits(1 To 4)    As Byte
    Dim strDecode       As String
    Dim Str             As String
    Dim ByteData()      As Byte
    Dim OutPut()        As Byte
    Dim lFrom           As Long
    Dim lTo             As Long
    Dim i               As Long
    Dim lngCount As Long, strFileName As String, lngFileNum
    
    '��ȥ�س�
    Str = Replace(str2Decode, vbCrLf, "")

    'ÿ4���ַ�һ�飨4���ַ���ʾ3���֣�
    For lPtr = 1 To Len(Str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '�����ַ���BASE64�ַ����е�λ��
            iValue = InStr(1, BASE64CHR, Mid$(Str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                'û�з���
                Case 0: Exit Function
            End Select
        Next

        'ת��4��6��������Ϊ3��8������
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
        
        '�����������ʼλ��
        lFrom = lTo
        lTo = lTo + (iLen - 1) - 1

        '���¶����������
        ReDim Preserve OutPut(0 To lTo)

        For lngCount = lFrom To lTo
            OutPut(lngCount) = bits(lngCount - lFrom + 1)
        Next
        lTo = lTo + 1
    Next
   
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , OutPut
    Close lngFileNum
    
    SaveBase64Bmp = strFileName

End Function

Private Function Hex_To_Dec(ByVal Hex As String) As Long
  Dim i As Integer
  Dim B As Long
  Hex = UCase(Hex)
  For i = 1 To Len(Hex)
    Select Case Mid(Hex, Len(Hex) - i + 1, 1)
        Case "0": B = B + 16 ^ (i - 1) * 0
        Case "1": B = B + 16 ^ (i - 1) * 1
        Case "2": B = B + 16 ^ (i - 1) * 2
        Case "3": B = B + 16 ^ (i - 1) * 3
        Case "4": B = B + 16 ^ (i - 1) * 4
        Case "5": B = B + 16 ^ (i - 1) * 5
        Case "6": B = B + 16 ^ (i - 1) * 6
        Case "7": B = B + 16 ^ (i - 1) * 7
        Case "8": B = B + 16 ^ (i - 1) * 8
        Case "9": B = B + 16 ^ (i - 1) * 9
        Case "A": B = B + 16 ^ (i - 1) * 10
        Case "B": B = B + 16 ^ (i - 1) * 11
        Case "C": B = B + 16 ^ (i - 1) * 12
        Case "D": B = B + 16 ^ (i - 1) * 13
        Case "E": B = B + 16 ^ (i - 1) * 14
        Case "F": B = B + 16 ^ (i - 1) * 15
    End Select
  Next i
  Hex_To_Dec = B
End Function

'����ֱ��ͼ����
Private Function BC5600(ByVal strIn As String) As String
    Dim strImg As String
    Dim lngI As Long, lng_len As Long, Data As Long
    Dim Str As String, DataStr As String
    
    lng_len = Len(strIn)
    If lng_len < 1 Then
        BC5600 = ""
        Exit Function
    End If
    Str = strIn
    strImg = ""
    
    For lngI = 1 To Len(strIn)
        Data = AscB(Mid(strIn, lngI, 1))
        'If Chr(Mid(strIn, lngI, 1)) <> Chr(32) Then
          If strImg <> "" Then strImg = strImg & ";"
          strImg = strImg & Format(CStr(Data), "000")
       'End If
    Next
    BC5600 = strImg
End Function

Private Function Graph_Text(ByVal str_bin As String) As String
    '����ͼ���ı�����
  Dim a() As Byte, C As String, i As Integer, Tmp As String
  Dim varLine As Variant
  Dim str_Line As Long
  C = ""
  a = StrConv(str_bin, vbFromUnicode)
  For i = LBound(a) To UBound(a)
    Tmp = Hex(a(i))
      If Trim(C) <> "" Then C = C & ";"
      C = C & Tmp
  Next
  Tmp = ""
  varLine = Split(C, ";")
  For i = 1 To UBound(varLine)
    str_Line = Hex_To_Dec(varLine(i))
    If (str_Line <> 63) And (str_Line <> 32) Then
      If Tmp <> "" Then Tmp = Tmp & ";"
      Tmp = Tmp & CStr(str_Line)
    End If
  Next
  Graph_Text = Tmp
End Function

'Base64����
Private Function Base64Decode(strEncoded As String) As String '����

    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT, bRet() As Byte
    Dim i As Long, j As Long
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

'-------------------------------------------------------------------------------------------------------------------
'�������������ΪBMPͼ���ļ�
Private Function SaveBmpToFile(ByVal str_bin As String, ByVal strType As String, ByVal str�걾�� As String) As String
    '�������������ΪBMPͼ���ļ�
    Dim strFileName As String, lngFileNum As Long, lngLoop As Long
    Dim strByte As String, ByteData() As Byte
    Dim strBinData As String
    Dim lngCount As Long
    
    strBinData = str_bin
    lngCount = Len_Bin(strBinData)
    If lngCount <= 0 Then Exit Function
    ReDim ByteData(lngCount)
    
    For lngLoop = 0 To lngCount - 1
        strByte = Left(strBinData, 3)
        strBinData = Mid(strBinData, 4)
        ByteData(lngLoop) = CByte("&H" & Mid(strByte, 2))
    Next
    
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
    Close lngFileNum
    
    SaveBmpToFile = strFileName
End Function

Public Function writelislog(strlog As String)
    Const strFile As String = "LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFileName As String
    strFileName = App.Path & "\" & strFile & Format(Date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing
End Function
