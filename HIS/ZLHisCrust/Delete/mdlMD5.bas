Attribute VB_Name = "mdlMD5"
Option Explicit
'**************************
'����:�ļ���ȡMD5ֵģ��
'��д�޸�:ף��
'**************************

'���� HashFile("C:\APPSOFT\Apply\zlCISKernel.dll", 2 ^ 27)
'������һ�� ��׼���޷���LONG�� ��4�ֽ�32λ�� �ɴ��2^32 ��
'��VB��LONG�����з��ŵ�  ֻ��31λ���ڼ��� ����1λ���ڱ���������� ����VB LONG ����λֻ�ܵ� 2^31 = 2147483648
'���ָ�����������ǵ�32λҲ������������� �����������Ҫ�ر���  Ϊ����ӦVB ���������� ����Ĵ������������Ը���

'SIZE��ÿ��Ӱ����ļ���С ֻ����2��N�η�  ��: 2^27=2��27�η�=128M
Public Function HashFile(ByVal szFilePath As String, ByVal Size As Long, Optional ByVal Algorithm As Long = MD5, Optional ByVal Block_Size As Long = 32768) As String
    Dim hFile As Long, hMapFile As Long, lpBaseMap As Long
    Dim hCtx As Long, lRet As Long, hHash As Long, lLen As Long
    Dim i As Long, j As Long, Point As Long
    Dim FI As LARGE_INTEGER, Current As LARGE_INTEGER, CurrentPoint As Double
    Dim Temp As Long, lBlocks As Long, lLastBlock As Long, Block() As Byte
    
    '�����ļ�ָ��
    hFile = CreateFileA(szFilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        FI.lowpart = GetFileSize(hFile, FI.highpart) '�ɹ��� ��ȡ�ļ���С
        If FI.highpart > 0 Then lBlocks = ((2 ^ 32 / Size) * FI.highpart) ' ��λ   Ϊ1���� 2^32���ֽ�  Ҳ����4�ֽ��޷��ų�������ֵ
        If FI.lowpart < 0 Then        '��λ
            lBlocks = lBlocks + (2 ^ 31 / Size) '��λΪ���� ��Ȼ����2^31�η�  ��Ϊ������2^31  VB����������ʾ
            Temp = LongToUnsigned(FI.lowpart) - 2 ^ 31 'תΪ�޷������ͼ���2^31�� VB����������ʾ��������
            lLastBlock = Temp \ Size
            lBlocks = lBlocks + lLastBlock
            lLastBlock = Temp - lLastBlock * Size
        Else
            Temp = FI.lowpart \ Size
            lBlocks = lBlocks + Temp
            lLastBlock = FI.lowpart - Temp * Size
        End If
        
        
        hMapFile = CreateFileMapping(hFile, ByVal 0&, PAGE_READONLY, FI.highpart, FI.lowpart, 0) '�����ļ�ӳ�����
        lRet = CryptAcquireContextA(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
        If Err.LastDllError = &H80090016 Then lRet = CryptAcquireContextA(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
        lRet = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
        ReDim Block(Block_Size) As Byte
        
        For i = 1 To lBlocks '�ɹ������ָ����С ��ʼӰ���ļ����ڴ�ռ�
            lpBaseMap = MapViewOfFile(hMapFile, FILE_MAP_READ, Current.highpart, Current.lowpart, Size)
            If lpBaseMap Then
                Point = lpBaseMap
                For j = 1 To Size / Block_Size ' 2��N�η�  ��Ȼ����
                    
                    lRet = CryptHashData(hHash, Point, Block_Size, 0)
                    Point = Point + Block_Size
                Next
                UnmapViewOfFile (lpBaseMap)
            End If
            CurrentPoint = CurrentPoint + Size
            Current = Currency2LargeInteger(CurrentPoint / 10000@) '�����ļ��ߵ�λ
        Next
            
        If lLastBlock > 0 Then 'ӳ������
            lpBaseMap = MapViewOfFile(hMapFile, FILE_MAP_READ, Current.highpart, Current.lowpart, lLastBlock)
            If lpBaseMap Then
                Point = lpBaseMap
                Temp = lLastBlock \ Block_Size '��һ������ ������FOR ѭ�����ٴμ���
                
                For j = 1 To Temp
                    lRet = CryptHashData(hHash, Point, Block_Size, 0)
                    Point = Point + Block_Size
                Next
                Temp = lLastBlock - Temp * Block_Size
                lRet = CryptHashData(hHash, Point, Temp, 0)
                UnmapViewOfFile (lpBaseMap)
            End If
        End If
        CloseHandle (hMapFile)

        If lRet Then
            lRet = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
            If lRet Then
                ReDim hash(lLen) As Byte
                lRet = CryptGetHashParam(hHash, HP_HASHVAL, hash(0), lLen, 0)
                If lRet Then
                    For j = 0 To UBound(hash) - 1
                        HashFile = HashFile & Right$("0" & Hex$(hash(j)), 2)
                    Next
                End If
                CryptDestroyHash hHash
            End If
        End If
        CryptReleaseContext hCtx, 0
        CloseHandle (hFile)
        
        If HashFile = "" Then
            On Error Resume Next
            HashFile = MD5File(szFilePath)
        End If
    End If
End Function

Private Function Currency2LargeInteger(ByVal curDistance As Currency) As LARGE_INTEGER
    CopyMemory Currency2LargeInteger, curDistance, 8
End Function


Private Function LongToUnsigned(Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + 2 ^ 32
    Else
        LongToUnsigned = Value
    End If
End Function

Private Function MD5String(p As String) As String
    Dim R As String * 32, t As Long
    R = Space(32)
    t = Len(p)
    MDStringFix p, t, R
    MD5String = UCase(R)
End Function

Private Function MD5File(f As String) As String
    Dim R As String * 32
    R = Space(32)
    MDFile f, R
    MD5File = UCase(R)
End Function
