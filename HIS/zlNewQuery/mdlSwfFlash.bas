Attribute VB_Name = "mdlSwfFlash"
Public Type FLASHHEADER
    intIsFlashMovie As Integer      '�Ƿ���SWF�ļ�,���߷�������
    lMHeight As Long                '��Ӱ�ĸ�      Pix
    lMWidth As Long                 '��Ӱ�Ŀ�      Pix
    bColorR As Byte                 '������ɫ��Rֵ Number
    bColorG As Byte                 '������ɫ��Gֵ Number
    bColorB As Byte                 '������ɫ��Bֵ Number
    intMTotalFrames As Integer      '��Ӱ����֡�� Frames
    lMSize As Long                  '��Ӱ�Ĵ�С  ByteNumber
    intMRate As Integer             '��Ӱ���ٶ�  FPS
    bMVersion As Byte               '������Ӱ��Flash�汾
End Type

'=========================================
'���ܣ�ȡFlash�ļ���ͷ���ṹ
'�÷���
'  Dim FH As FlashHeader
'  FH = GetFlashHeader(strFlashFileName)
'����:
'�ɹ���FlashHeader�ṹ        FH.intIsFlashMovie=1
'�����ļ��Ҳ���             FH.intIsFlashMovie=-1
'      ����FlashMovie�ļ�     FH.intIsFlashMovie=0
'      δ֪����:              FH.intIsFlashMovie=2
'=========================================
Public Function GetFlashHeader(strFileName As String) As FLASHHEADER
    Dim lFileNumber As Long                      '�ļ���
    Dim b(20) As Byte
    Dim strSWFSignature As String * 3            'SWF��ǩ��
    Dim intTagSize As Integer                    '��ǩ��Ĵ�С
    Dim lMWidth As Long                          '��Ӱ�Ŀ�
    Dim lMHeight As Long                         '��Ӱ�ĸ�
    Dim bMVersion As Byte
    Dim bColorR As Byte                          '������ɫ��Rֵ Number
    Dim bColorG As Byte                          '������ɫ��Gֵ Number
    Dim bColorB As Byte                          '������ɫ��Bֵ Number
    Dim intMTotalFrames As Integer               '��Ӱ����֡�� Frames
    Dim lMSize As Long                           '��Ӱ�Ĵ�С  ByteNumber
    Dim intMRate(1)  As Byte                     '��Ӱ���ٶ�  FPS ֡ÿ��
    Dim nBites As Integer                        'һ��Tag�Ĵ�С,��ʾһ��Tagռ�е�Bitλ��
    
    Dim i As Integer
    Dim Tmpstring As String
 
    On Error GoTo errHand:
    
    '����ļ������ڣ�����-1
    If Dir(strFileName) = "" Then
        GetFlashHeader.intIsFlashMovie = -1
        Exit Function
    End If
    
     '���ļ�
    lFileNumber = FreeFile
    Open strFileName For Binary As #lFileNumber
         '��ȡǩ��
         Get #lFileNumber, , strSWFSignature
         '�������SWF�ļ�������
         If strSWFSignature <> "FWS" Then
            GetFlashHeader.intIsFlashMovie = 0
            Close #lFileNumber
            Exit Function
         End If
         
         Get #lFileNumber, , bMVersion      '�汾
         Get #lFileNumber, , lMSize         '��Ӱ��С
         Get #lFileNumber, , b()
         
        '�ھ�λ�Ŀ�ʼ(��������)��ǰ5����Ϊ�����ǩ��nBites
        '�ṹ����
        'Field      Type                Comment
        'Nbits      nBits = UB[5]       Bits in each rect value field
        'Xmin       SB[nBits]           X minimum position for rect
        'Xmax       SB[nBits]           X maximum position for rect
        'Ymin       SB[nBits]           Y minimum position for rect
        'Ymax       SB[nBits]           Y maximum position for rect
        '�Ҳ�֪������ȥǰ5��Bits�����ݣ�ͨ��And �� Or���Եģ���������ѡ������Щ���ݶ�������
        'תΪ�������ַ���,ȡǰ5��,��תΪ����,
         nBites = Bin2Dec(Left(Dec2Bin(b(0)), 5))
         intTagSize = (nBites * 4 + 5) \ 8 + 1
         '�����Ĵ�С�ڹر��ļ�֮���ټ���
         
         Get #lFileNumber, 9 + intTagSize, intMRate         '����
         Get #lFileNumber, , intMTotalFrames                '��֡��
         Get #lFileNumber, 9 + intTagSize + 6, bColorR      '������ɫR
         Get #lFileNumber, , bColorG                        '������ɫG
         Get #lFileNumber, , bColorB                        '������ɫB
    Close #lFileNumber
    
    
    'ȡ��Ӱ��ԭʼ�߶�
    'תΪ�������ַ���
    Tmpstring = ""
    For i = 0 To intTagSize - 1
        Tmpstring = Tmpstring & Dec2Bin(b(i))
    Next
    '��(����������+nBites)��ʼ,nBites��)\20
    GetFlashHeader.lMWidth = Bin2Dec(Mid(Tmpstring, 6 + nBites, nBites)) \ 20
    '�ߣ�(����������+nBites*3)��ʼ,nBites��)\20
    GetFlashHeader.lMHeight = Bin2Dec(Mid(Tmpstring, 6 + nBites * 3, nBites)) \ 20
    
    
    GetFlashHeader.intIsFlashMovie = 1
    GetFlashHeader.bMVersion = bMVersion
    GetFlashHeader.lMSize = lMSize
    GetFlashHeader.intMRate = intMRate(0) * 255 + intMRate(1)
    GetFlashHeader.intMTotalFrames = intMTotalFrames
    GetFlashHeader.bColorR = bColorR
    GetFlashHeader.bColorG = bColorG
    GetFlashHeader.bColorB = bColorB
    Exit Function
errHand:
   GetFlashHeader.intIsFlashMovie = 2
End Function

Public Function Bin2Dec(strBin As String) As Long
'���ܣ�������תΪʮ���ƺ���
'�÷���Long  bin2dec(strBin as String)
'���أ�  �����Ƶ�ʮ���� ��������Long��
'����  ����-1
    
    Dim lDec As Long
    Dim lCount As Long
    Dim i As Long
    
    On Error GoTo errHand
    lDec = 0
    If strBin = "" Then strBin = "0"

    lCount = Len(strBin)
    For i = 1 To lCount
        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
        strBin = Right(strBin, Len(strBin) - 1)
        DoEvents
    Next
    Bin2Dec = lDec
    Exit Function
errHand:
    Bin2Dec = -1
End Function

Public Function Dec2Bin(bDec As Byte) As String
'���ܣ�ʮ����תΪ�����ƺ���
'�÷���String  Dec2Bin(Bdec as Byte)
'���أ�  ʮ���ƵĶ����� �ַ���(String)
'����  ����"0"
    
    Dim strBin As String
    
    On Error GoTo Err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If

    strBin = ""

    'תΪ�ַ���
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
'        DoEvents
    Wend
    '������8λ
    If Len(strBin) < 9 Then
        While Len(strBin) < 8
            strBin = "0" & strBin
        Wend
    End If
    Dec2Bin = strBin
    
    Exit Function
Err:
   Dec2Bin = "0"
End Function

