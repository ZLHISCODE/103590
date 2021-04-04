VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_BC5600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:ѪҺ������ BC5600
'����:����(Mindary)
'����:�µ���
'�û�:˫������ҽԺ

Private mstrCode As String  '����������
Const mstrItem  As String = ",WBC,BAS#,BAS%,NEU#,NEU%,EOS#,EOS%,LYM#,LYM%,MON#,MON%,*ALY#,*ALY%,*LIC#,*LIC%,RBC,HGB,MCV,MCH,MCHC,RDW-CV,RDW-SD,HCT,PLT,MPV,PDW,PCT,"
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
    
    Dim str_Base64 As String, str_wbc As String, str_RBC As String, str_PLT As String, str_BASO As String, str_img As String
    
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
                    str_img = ""
                    If str_wbc <> "" Then str_img = str_img & "^DIFF;103;1;" & str_wbc
                    If str_BASO <> "" Then str_img = str_img & "^BASO;103;1;" & str_BASO
                    If str_RBC <> "" Then str_img = str_img & "^RBC;0;260,300,0-0,50-50,100-100,150-150,200-200,250-250,300-fL;" & str_RBC
                    If str_PLT <> "" Then str_img = str_img & "^PLT;0;260,150,0-0,21.333-5,42.66-10,63.99-15,85.32-20,106.65-25,127.98-fL;" & str_PLT
                    If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
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
                ElseIf InStr(",15200,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'WBC DIFF(ɢ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_wbc = SaveBase64Bmp("DIFF", str_�걾��, str_Base64)
                ElseIf InStr(",15050,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'RBC Histogram(ֱ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_Base64 = Base64Decode(str_Base64)
                    str_RBC = Graph_Text(str_Base64)
                ElseIf InStr(",15100,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'PLT Histogram(ֱ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_Base64 = Base64Decode(str_Base64)
                    str_PLT = Graph_Text(str_Base64)
                ElseIf InStr(",15250,", "," & Split(str_��Ŀ, str_����ָ���)(0) & ",") > 0 Then
                'Baso Scattergram(ɢ��ͼ)
                    str_Base64 = Split(Trim(var_Filed(5)), str_����ָ���)(4)
                    str_BASO = SaveBase64Bmp("BASO", str_�걾��, str_Base64)
                End If
            
            End If
        Next
        If Val(str_�걾��) > 0 And str_Return <> "" Then
            strResult = strResult & "||" & str_���� & "|" & Val(str_�걾��) & "| |Ѫ��|" & IIf(str_�ʿ� = "Q", 1, 0) & str_Return
            str_img = ""
            If str_wbc <> "" Then str_img = str_img & "^DIFF;103;1;" & str_wbc
            If str_BASO <> "" Then str_img = str_img & "^BASO;103;1;" & str_BASO
            If str_RBC <> "" Then str_img = str_img & "^RBC;0;260,300,0-0,50-50,100-100,150-150,200-200,250-250,300-fL;" & str_RBC
            If str_PLT <> "" Then str_img = str_img & "^PLT;0;260,150,0-0,21.333-5,42.66-10,63.99-15,85.32-20,106.65-25,127.98-fL;" & str_PLT
            If str_img <> "" Then strResult = strResult & vbCrLf & Mid(str_img, 2)
            'Call writelislog(strResult)
        End If
        lng_S = InStr(mstrCode, Chr(&HB))
        lng_E = InStr(mstrCode, Chr(&H1C) & Chr(&HD))
    Loop
    If strResult <> "" Then strResult = Mid(strResult, 3)
    Exit Sub
errHandle:
    WriteLog "BC5600_analyse", strSource, Err.Number & " " & Err.Description
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
   
    ReDim ByteData(UBound(OutPut))
    For i = 0 To UBound(OutPut)
        ByteData(i) = OutPut(i)
    Next
    
        For i = 1078 To UBound(OutPut) - 2 Step 2
            If OutPut(18) = &H18 And OutPut(19) = &H1 Then 'ֱ��ͼ
'               If Output(i) <> byt1 Or byte2 <> Output(i + 1) Then
'                   byt1 = Output(i): byte2 = Output(i + 1)
''                   Debug.Print byt1, byte2
'               End If
                If OutPut(i) = OutPut(i - 2) And OutPut(i + 1) = OutPut(i - 1) And OutPut(i) = OutPut(i + 2) And OutPut(i + 1) = OutPut(i + 3) And ((OutPut(i) = 0 And OutPut(i + 1) = 124) Or (OutPut(i) = 255 And OutPut(i + 1) = 127) Or (OutPut(i) = 224 And OutPut(i + 1) = 3)) Then
                    If ((i - 54) \ 560 > 0 And (i - 54) \ 560 <> 159) Then
                        'If Output(i) = Output(i - 560) And Output(i + 1) = Output(i - 559) Then 'And Output(i) = Output(i + 560) And Output(i + 1) = Output(i + 561)
                        If OutPut(i) = OutPut(i + 560) And OutPut(i + 1) = OutPut(i + 561) Then
                            ByteData(i) = 0: ByteData(i + 1) = 0
    '                        Debug.Print (i - 54) \ 560
                        End If
                    End If
                End If
            End If
            If ByteData(i) = 0 Then
               ByteData(i) = 255
            End If
            If ByteData(i + 1) = 0 Then
               ByteData(i + 1) = 255
            End If
            'output2(i) = 255 - output2(i): output2(i + 1) = 255 - output2(i + 1)
        Next i
 
    strFileName = App.Path & "\" & Format(Now, "yyyyMMdd") & "_" & Val(str�걾��) & "_" & strType & ".bmp"
    lngFileNum = FreeFile
    Open strFileName For Binary Access Write As lngFileNum
    Put lngFileNum, , ByteData
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
    strFileName = App.Path & "\" & strFile & Format(date, "yyyy-mm-dd") & ".log"
    If Not objFile.FileExists(strFileName) Then Call objFile.CreateTextFile(strFileName)
    Set objStream = objFile.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strlog)
    objStream.Close
    Set objStream = Nothing

End Function
