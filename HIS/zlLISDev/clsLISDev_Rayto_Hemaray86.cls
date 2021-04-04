VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_Rayto_Hemaray86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'�û���տ���е�һ��ҽԺ
'�����ͺţ������׶�Hemaray86  ���ӿڻ���λͼ��ʽ����ͼƬ
'�޸��ˣ�������
'�޸�ʱ�䣺2014��3��4��
'�������ݸ�ʽ��дɢ��ͼ,λͼ�ķ�ʽȥ����ɫ����ɫ

Private mstrCode As String  '����������
Const mstrItem As String = ",WBC,BASOP,BASOA,NEUTP,NEUTA,EOSP,EOSA,LYMP,LYMA,MONOP,MONOA,RBC,HGB,MCV,MCH,MCHC,RDWCV,RDWSD,HCT,PLT,MPV,PDW,PCT,PLCR,PLCC,IGP,IGA,ALYP,ALYA,"
' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="


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
    Dim str_Code As String    '1����������Ϣ
    Dim var_Code As Variant
    Dim str_i As Variant
    Dim lng_count As Long
    Dim strShowXScale As String, strShowYScale As String, strShowThresholds As String, blnShowDot As Boolean, blnShowBorder As Boolean
    Dim strXYcolor As String, strCurveColor As String, strPicObjBackColor As String, strFillColor As String, strThresholdsColor As String

    Dim str_�ֶηָ��� As String
    Dim str_����ָ��� As String
    Dim str_��������� As String
    Dim str_�ظ��ָ��� As String
    Dim str_ת���ַ� As String

    Dim str_Segment As String    'һ����Ϣ��(��chr(&HD)�ָ�)
    Dim var_Filed As Variant   '�ֶ�
    Dim str_Ψһ��ʶ As String    '���ڱ�ʶһ����Ϣ
    Dim str_�汾 As String
    Dim lng_ͼ�θ�ʽ As Integer    'ֱ��ͼ�������ݵĴ��䷽ʽ 1����λͼ��ʽ���� 2:�����ݷ�ʽ����

    Dim lng_S As Long
    Dim lng_E As Long

    Dim str_�걾�� As String, str_���� As String
    Dim str_�ʿ� As String
    Dim str_��Ŀ As String
    Dim str_��� As String
    Dim str_Return As String

    Dim str_Base64 As String, str_FLFS As String, str_RBC As String, str_PLT As String, str_SSFS As String, Str_Img As String

    On Error GoTo errHandle
    strResult = "": str_Return = "": str_�걾�� = "": str_���� = ""

    strResult = ""
    If strSource <> "" Then
        mstrCode = mstrCode + strSource
    End If
    lng_S = InStr(mstrCode, Chr(&HB))
    lng_E = InStr(mstrCode, Chr(&H1C))
    Do While lng_S > 0 And lng_E > 0

'       writelislog (mstrCode)
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 2)
        mstrCode = Mid(mstrCode, lng_E + 2)
        var_Code = Split(str_Code, Chr(&HD))
        For lng_count = LBound(var_Code) To UBound(var_Code)
            str_Segment = var_Code(lng_count)
            str_Segment = Replace(str_Segment, Chr(&HA), "")
            If Mid(str_Segment, 2, 3) = "MSH" Then
                '��Ϣͷ
                str_�ֶηָ��� = Mid(str_Segment, 5, 1)
                str_����ָ��� = Mid(str_Segment, 8, 1)
                str_��������� = Mid(str_Segment, 11, 1)
                str_�ظ��ָ��� = Mid(str_Segment, 9, 1)
                str_ת���ַ� = Mid(str_Segment, 10, 1)

                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_Ψһ��ʶ = var_Filed(11)
                str_�ʿ� = var_Filed(12)
                str_�汾 = var_Filed(13)
                '                If str_�汾 <> "2.3.1" Then
                '                    WriteLog "�汾����", "", ""
                '                End If

                str_���� = var_Filed(8)
                str_���� = Mid(str_����, 1, 4) & "-" & Mid(str_����, 5, 2) & "-" & Mid(str_����, 7, 2) & " " & Mid(str_����, 9, 2) & ":" & Mid(str_����, 11, 2) & ":" & Mid(str_����, 13, 2)
            ElseIf InStr(str_Segment, "PID|") > 0 Then
                '�걾�ţ�����
                If Val(str_�걾��) > 0 And str_Return <> "" Then
                    strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
                    Str_Img = ""
                    If str_FLFS <> "" Then Str_Img = Str_Img & "^;103;1;" & str_FLFS
                    If str_RBC <> "" Then Str_Img = Str_Img & "^;103;1;" & str_RBC
                    If str_PLT <> "" Then Str_Img = Str_Img & "^;103;1;" & str_PLT
                    If str_SSFS <> "" Then Str_Img = Str_Img & "^;103;1;" & str_SSFS
                    If Str_Img <> "" Then strResult = strResult & vbCrLf & Mid(Str_Img, 2)
                End If
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_�걾�� = var_Filed(2)
                str_Return = ""
            ElseIf InStr(str_Segment, "OBX|") > 0 Then
                var_Filed = Split(str_Segment, str_�ֶηָ���)
                str_��Ŀ = var_Filed(3)
                If InStr(mstrItem, str_��Ŀ) > 0 Then
                    str_Return = str_Return & "|" & str_��Ŀ & "|" & Trim(var_Filed(5))
                Else
                    If InStr(var_Filed(5), "DAT") > 0 Then
                        '-------------------------�����ݷ�ʽ����------------------
                        If InStr("Sub Scattergram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_Base64 = Base64Decode(str_Base64, 1)
                            str_SSFS = drawDIFF(str_Base64, str_�걾��, "SSFS")
'                            writelislog "����ͼƬ " & str_SSFS
                        ElseIf InStr("Main Scattergram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_Base64 = Base64Decode(str_Base64, 1)
                            str_FLFS = drawDIFF(str_Base64, str_�걾��, "FLFS")
'                            writelislog "����ͼƬ " & str_FLFS
                        ElseIf InStr("RBC Histogram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_Base64 = Base64Decode(str_Base64)
                            str_RBC = Graph_Text(str_Base64)
'                            writelislog "����ͼƬRBC "
                        ElseIf InStr("PLT Histogram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_Base64 = Base64Decode(str_Base64)
                            str_PLT = Graph_Text(str_Base64)
'                            writelislog "����ͼƬPLT "
                        End If
                        If str_RBC <> "" Then Str_Img = Str_Img & "^RBC;0;200,300,55-50,100-100,150-150,200-200;" & str_RBC
                        If str_PLT <> "" Then Str_Img = Str_Img & "^PLT;0;20,30,7-2,27-5,81-10,135-15,181-20,230-25;" & str_PLT
                        str_RBC = ""
                        str_PLT = ""
                    ElseIf InStr(var_Filed(5), "BMP") > 0 Then
                        '-------------------------��λͼ��ʽ����------------------
                        If InStr("Sub Scattergram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_SSFS = SaveBase64Bmp("SSFS", str_�걾��, str_Base64)
'                            writelislog "����ͼƬ " & str_SSFS
                        ElseIf InStr("Main Scattergram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_FLFS = SaveBase64Bmp("FLFS", str_�걾��, str_Base64)
'                            writelislog "����ͼƬ " & str_FLFS
                        ElseIf InStr("RBC Histogram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_RBC = SaveBase64Bmp("RBC", str_�걾��, str_Base64)
'                            writelislog "����λͼ RBC "
                        ElseIf InStr("PLT Histogram", str_��Ŀ) > 0 Then
                            str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(1)
                            str_PLT = SaveBase64Bmp("PLT", str_�걾��, str_Base64)
'                            writelislog "����λͼ PLT "
                        End If
                        'ͼ�����ơ�ͼ���ʽ��������Ƿ�ɾ����ͼ������
                        If str_RBC <> "" Then Str_Img = Str_Img & "^RBC;103;1;" & str_RBC
                        If str_PLT <> "" Then Str_Img = Str_Img & "^PLT;103;1;" & str_PLT
                        str_RBC = ""
                        str_PLT = ""
                    End If
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
            '            Str_Img = ""
            '���ݸ�ʽ RBC��PLT��ͼ����ʽ
            If str_SSFS <> "" Then Str_Img = Str_Img & "^SSFS;103;1;" & str_SSFS
            If str_FLFS <> "" Then Str_Img = Str_Img & "^FLFS;103;1;" & str_FLFS
            If Str_Img <> "" Then strResult = strResult & vbCrLf & Mid(Str_Img, 2)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C))
    Loop

    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
'    strCmd = "MSH|^~\&|LIS||||" & Format(Time, "YYYYMMDDHHMMSS") & "||ACK^RO1| " & str_�걾�� & "|" & str_�ʿ� & "|2.3.1||||S||UNICODE|||" & vbCrLf & "MSA|AA|" & str_�걾�� & "|||||"  'Ӧ��
'    strReserved = ""
'    mstrCode = ""
errHandle:
    writelislog "Hemaray86_analyse��" & Err.Number & " " & Err.Description
End Sub

'---------------------------------------------------------------------------------------------------------------
'�������ڲ�����
Private Function Hex_To_Dec(ByVal strHex As String) As Long
    Dim i As Integer
    Dim lngB As Long
    On Error GoTo errH
    strHex = UCase(strHex)
    For i = 1 To Len(strHex)
        Select Case Mid(strHex, Len(strHex) - i + 1, 1)
        Case "0": lngB = lngB + 16 ^ (i - 1) * 0
        Case "1": lngB = lngB + 16 ^ (i - 1) * 1
        Case "2": lngB = lngB + 16 ^ (i - 1) * 2
        Case "3": lngB = lngB + 16 ^ (i - 1) * 3
        Case "4": lngB = lngB + 16 ^ (i - 1) * 4
        Case "5": lngB = lngB + 16 ^ (i - 1) * 5
        Case "6": lngB = lngB + 16 ^ (i - 1) * 6
        Case "7": lngB = lngB + 16 ^ (i - 1) * 7
        Case "8": lngB = lngB + 16 ^ (i - 1) * 8
        Case "9": lngB = lngB + 16 ^ (i - 1) * 9
        Case "A": lngB = lngB + 16 ^ (i - 1) * 10
        Case "B": lngB = lngB + 16 ^ (i - 1) * 11
        Case "C": lngB = lngB + 16 ^ (i - 1) * 12
        Case "D": lngB = lngB + 16 ^ (i - 1) * 13
        Case "E": lngB = lngB + 16 ^ (i - 1) * 14
        Case "F": lngB = lngB + 16 ^ (i - 1) * 15
        End Select
    Next i
    Hex_To_Dec = lngB
    Exit Function
errH:
    Debug.Print Err.Description
End Function

'Base64����
Private Function Base64Decode(strEncoded As String, Optional ByVal intType As Integer) As String   '����

    On Error Resume Next
    Dim arrB() As Byte, btTmp(3) As Byte, lngGet As Long, btRet() As Byte
    Dim i As Long, j As Long
    On Error GoTo errH
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim btRet(j - j \ 4 - 1)
    For i = 0 To j Step 4
        Erase btTmp
        btTmp(0) = (InStr(BASE64CHR, Chr(arrB(i))) - 1) And 63
        btTmp(1) = (InStr(BASE64CHR, Chr(arrB(i + 1))) - 1) And 63
        btTmp(2) = (InStr(BASE64CHR, Chr(arrB(i + 2))) - 1) And 63
        btTmp(3) = (InStr(BASE64CHR, Chr(arrB(i + 3))) - 1) And 63
        '��ɢ��ͼ��
        If intType = 1 Then
            If Val((i \ 4) * 3) <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3) = btTmp(0)
            End If
            If Val((i \ 4) * 3) + 1 <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3 + 1) = btTmp(1)
            End If
            If Val((i \ 4) * 3) + 2 <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3 + 2) = btTmp(2)
            End If
        Else
            lngGet = btTmp(0) * 2 ^ 18 + btTmp(1) * 2 ^ 12 + btTmp(2) * 2 ^ 6 + btTmp(3)
            '        ---------------------------------------65536
            If Val((i \ 4) * 3) <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3) = lngGet \ 65536
            End If
            If Val((i \ 4) * 3 + 1) <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3 + 1) = (lngGet And 65280) \ 256
            End If
            If Val((i \ 4) * 3 + 2) <= Val(UBound(btRet)) Then
                btRet((i \ 4) * 3 + 2) = lngGet And 255
            End If

        End If
    Next
    Base64Decode = StrConv(btRet, vbUnicode)
    Exit Function
errH:
    Debug.Print Err.Description
End Function


Public Function writelislog(strlog As String)
    Const strFile As String = "LisLog_"
    Dim objFile As New Scripting.FileSystemObject
    Dim objStream As TextStream
    Dim strFilename As String
    strFilename = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFilename) Then Call objFile.CreateTextFile(strFilename)
    Set objStream = objFile.OpenTextFile(strFilename, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing
End Function

Private Function SaveBase64Bmp(ByVal strType As String, ByVal str�걾�� As String, ByVal str2Decode As String) As String

' ******************************************************************************
'λͼ���ݽ�������
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

    Dim lPtr As Long
    Dim iValue As Integer
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim bits(1 To 4) As Byte
    Dim strDecode As String
    Dim str As String
    Dim ByteData() As Byte
    Dim OutPut() As Byte
    Dim lFrom As Long
    Dim lTo As Long
    Dim i As Long
    Dim lngCount As Long, strFilename As String, lngFileNum
    Const bytBai As Byte = &HFF
    Const bytHei As Byte = &H0
    Dim objPic As PictureBox
    On Error GoTo errH
    Dim k As Long
    k = 1
    '��ȥ�س�
    str = Replace(str2Decode, vbCr, "")
    str = Replace(str2Decode, vbLf, "")
    'ÿ4���ַ�һ�飨4���ַ���ʾ3���֣�
    For lPtr = 1 To Len(str) Step 4
        iLen = 4
        For iCtr = 0 To 3
            '�����ַ���BASE64�ַ����е�λ��
            iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
            Case 1 To 64: bits(iCtr + 1) = iValue - 1
                ' =
            Case 65
                iLen = iCtr
                Exit For
                'û�з���
            Case 0:    'Exit Function
            End Select
        Next
        k = k + 3
        'ת��4��6��������Ϊ3��8������
        bits(1) = bits(1) * &H4 + (bits(2) And &H30) \ &H10
        bits(2) = (bits(2) And &HF) * &H10 + (bits(3) And &H3C) \ &H4
        bits(3) = (bits(3) And &H3) * &H40 + bits(4)
        '����λͼ��ɫ���е����ݡ���ɫ����ռ���ݲ��þ�ȷ�����������ʡ�
        '������ɫ
        If k >= 54 And k < 61 Then
            If bits(1) <= 0 And bits(2) <= 0 And bits(3) <= 0 Then
                bits(1) = 255
                bits(2) = 255
                bits(3) = 255
            End If
        End If

        '������ɫ
        If k > 61 And k < 2000 Then
            If bits(1) >= 200 And bits(2) >= 200 And bits(3) >= 200 Then
                bits(1) = 0
                bits(2) = 0
                bits(3) = 0
            End If
        End If
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

    Dim strJPGName As String

    strFilename = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    strJPGName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFilename For Binary Access Write As lngFileNum
    Put lngFileNum, , OutPut
    Close lngFileNum

    SaveBase64Bmp = strJPGName
    Exit Function
errH:
    Debug.Print Err.Description
End Function

Private Function Graph_Text(ByVal str_bin As String) As String
'����ͼ���ı�����
    Dim btGet() As Byte, strData As String, i As Integer, strTmp As String
    Dim varLine As Variant
    Dim str_Line As Long
    On Error GoTo errH
    strData = ""
    btGet = StrConv(str_bin, vbFromUnicode)
    For i = LBound(btGet) To UBound(btGet)
        strTmp = Hex(btGet(i))
        If Trim(strData) <> "" Then strData = strData & ";"
        strData = strData & strTmp
    Next
    strTmp = ""
    varLine = Split(strData, ";")
    For i = 1 To UBound(varLine)
        str_Line = Hex_To_Dec(varLine(i))
        If (str_Line <> 63) And (str_Line <> 32) Then
            If strTmp <> "" Then strTmp = strTmp & ";"
            strTmp = strTmp & CStr(str_Line)
        End If
    Next
    Graph_Text = strTmp
    Exit Function
errH:
    Debug.Print Err.Description
End Function


Public Function drawDIFF(ByVal strImgData As String, strNO As String, strType As String) As String
    Dim x As Integer
    Dim Y As Integer
    Dim i As Long
    Dim strLine As String
    Dim strTmp As String
    Dim btPrint() As Byte
    On Error GoTo errH
    frmGraph.Picture1.Scale (0, 0)-(256, 256)
    frmGraph.Picture1.BackColor = vbWhite
    strLine = strImgData
    btPrint = StrConv(strImgData, vbFromUnicode)
    For Y = 255 To 0 Step -1
        For x = 0 To 255
            If i < UBound(btPrint) Then
                strTmp = Hex(btPrint(i))
                i = i + 1
                If Val(Replace(strTmp, ";", "")) <> 0 Then
                    frmGraph.Picture1.PSet (x, Y), vbBlack
                End If
            End If
        Next
    Next
    If Dir(App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(strNO) & "_" & strType & ".bmp") <> "" Then
        Kill App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(strNO) & "_" & strType & ".bmp"
    End If
    SavePicture frmGraph.Picture1.Image, App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(strNO) & "_" & strType & ".bmp"
    drawDIFF = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(strNO) & "_" & strType & ".bmp"
    Exit Function
errH:
    Debug.Print Err.Description
End Function





