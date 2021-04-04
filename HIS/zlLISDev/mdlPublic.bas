Attribute VB_Name = "mdlPublic"
Option Explicit


' ***************************************************
' *             �ı���תģ��                        *
' *                                                 *
' ***************************************************

Public uDisplayDescript  As Boolean      'ѡ��ʱ��ʾ��ϸ����

'API ����:
Private Const LF_FACESIZE   As Long = 32&
Private Const SYSTEM_FONT   As Long = 13&
Private Const ANTIALIASED_QUALITY = 4

'�ṹ����:
Private Type PointAPI
    x   As Long
    Y   As Long
End Type

Private Type SizeStruct
    Width   As Long
    Height  As Long
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

'API ����:
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SizeStruct) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal LpApplicationName As String, ByVal LpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'----- ����ΪJPG��ʽ��ͼƬ
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
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

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type

Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal Filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public dX As Long, dy As Long          ' distance XY = size of snapping zone
Public X1 As Long, X2 As Long          ' co�rdinates snapping zone
Public Y1 As Long, Y2 As Long

'--------------------------------------------------------------
Dim lngTime
'       ���뺯�� �� ���ش���ʽ˵��
'    ResultFromFile ����  ���ַ������鷽ʽ���ؽ�����, һ������Ԫ�ذ���һ�������;
'    Analyse        ����  ���ַ�����ʽ���ؽ�����,ÿ���������||�ָ�
'    ÿ���������Ԫ��֮����|�ָ�,������ϸ˵��

    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������[^�Ƿ���^����]
    '              ^�Ƿ���^���� �ǿ�ѡ��,��������������ʱ��ʹ��.
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    '       ø������ʽ:���Խ��^OD^CutOff^SCO
    '
    'Analyse strCmd �����������Ҫ���ɷ������豸���͵�����
    '   1.������ָ�������|����
    '       a.�Զ����ƴ���ʽ�����Զ�Ӧ��ָ����Ҫ�̶�Ӧ��06��дΪ,06
    '   2.���������͵�ָ�������|����
    '       a.����Ҫ���͵�ָ��ǰ��ӡ�1|��������ָ���ʾ���������ȡ�걾��Ϣ�����Ǽ�����
    '       b.����Ҫ���͵�ָ��ǰ��ӡ�0|��������ָ����������������ӡ�0|����Ϊ�˺͡�1|���������֣�HL7Э�����������͵�ָ��һ�㶼������|����
    '
    
    
    '-- ����ͼ������ʱ�ĸ�ʽ:
    '����ͼ��ķ�ʽ��
    '                   1.ͼ�����ݸ���ָ�����ݺ�ʹ�ûس����з����ָ���
    '                   2.�ж��ͼ������ʱʹ��"^"���ָ�
    '                   3.����ͼ�����ݸ�ʽ: ͼ�񻭷� 0=ֱ��ͼ  1=ɢ��ͼ  2=Ѫ����ճ����������  3=Ѫ������ 4=PLT˫����ͼ 5=�����֧��˫���ߣ�XY����̶�ֵ��ֱ��ͼ 100����ΪͼƬ����
    '                     0) ֱ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ);Y1;Y2;Y3;Y4;Y5...
    '                     1) ɢ��ͼ: ͼ������;ͼ�񻭷�(0=ֱ��ͼ  1=ɢ��ͼ):
    '                        ��:00000100001000010000100010;00000100001000010000100010;
    '                        ˵��: 1.ɢ��ͼ�Ե���ʽ����ÿһ��ʹ�÷ֺ����ָ�.
    '                              2.�ж��ٸ��ֺž��ж�����
    '                              3.ÿһ���ж��ٸ�����ÿһ�еĳ�����ȷ��
    '                              4.��ͼ�ķ����Ǵ����ϱ����»�������65*65��ͼ���Ǵ�65�п�ʼ��(���ϱ߿�ʼ��)
    '                     2) ճ����������:ͼ������;ͼ�񻭷�;��������;���߼��������;�������������
    '                                   ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
    '                                   ���߼��������:ճ������1�ĸߵ�͵͵�����|ճ������2�ĸߵ�͵͵�����~���е�����,���е�����,���е�����
    '                                   �������������:Y�����������,X����,Y����~X�����������,X����,Y����
    '                        ��:ճ����������;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;VIS(mPa.s),25,20~SHR(1/S),195,1

    '                     3) Ѫ������:ͼ������;ͼ�񻭷�;��������;�������;�������������
    '                                   ����  �������ݣ�Y����,X����|X����-X������ʾ������,....|Y����-Y������ʾ������,....
    '                                   �������:Ѫ��ֵ1,Ѫ��ֵ2,....Ѫ��ֵ30
    '                                   �������������:Y�����������,X����,Y����~X�����������,X����,Y����
    '                        ��:Ѫ������;3;36,30|3-6,6-12,9-18,12-24,15-30,18-36,21-42,24-48,27-54,30-60|4-4,8-8,12-12,16-16,20-20,24-24,28-28,32-32,36-36;.5,.5,1,1,1,1.5,1.5,2,2,2,2.5,3,3,3.5,4,4.5,5.5,6.5,8,9,10.5,11.5,12.5,13.5,14.5,15.5,16.5,18,19,20;Ѫ��ֵ(mm),5,36~ʱ��(m),55,1
    '                     4) PLTͼ��ͼ������;ͼ�񻭷�;��������;�������
    '                               ���� �������ݣ�Y����,X����,X����-X������ʾ������,....[|Y����-Y������ʾ������,.....]
    '                                    �������: Y1,Y2,Y3,......|Y1,Y2,Y3,......[~Y���������,X����,Y����|X���������,X����,Y����]
    '                        ��:PLT;4;200,262;0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,3,4,4,7,7,12,12,17,17,20,20,25,25,30,30,33,33,36,36,41,41,43,43,44,44,46,46,47,47,47,47,47,47,46,46,46,46,44,44,44,44,43,43,41,41,39,39,38,38,36,36,35,35,33,33,31,31,30,30,28,28,27,27,25,25,23,23,22,22,22,22,20,20,19,19,17,17,15,15,15,15,14,14,12,12,12,12,11,11,11,11,9,9,9,9,9,9,7,7,7,7,7,7,6,6,6,6,6,6,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,3,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,7,7,9,9,8,8,9,9,12,12,16,16,22,22,26,26,30,30,35,35,36,36,37,37,39,39,42,42,44,44,46,46,46,46,44,44,43,43,40,40,37,37,37,37,37,37,39,39,37,37,36,36,32,32,29,29,25,25,23,23,22,22,22,22,21,21,19,19,18,18,16,16,16,16,15,15,15,15,15,15,14,14,12,12,11,11,9,9,9,9,8,8,8,8,7,7,7,7,7,7,7,7,7,7,8,8,7,7,7,7,5,5,4,4,4,4,2,2,4,4,4,4,2,2,2,2,4,4
    
    '                     5) ֱ��ͼ (�����֧��˫���ߣ�X,Y����̶�ֵ)������;ͼ������;Y�߶�,X����;�������ұ߿����ף����ڻ��̶ȣ�;X��̶�[|Y�̶�];����1����[|����2����...][;�������]
    '                                ����:��������: ��y��������,��,�ָ�,��������������|�ָ�
    '                                    :�������: ��x��������,��,�ŷָ�
    '
    '                        ����   RBC;5;260,310;10,50,50,10;0-0,50-50,100-100,150-150,200-200,250-250,300-fL|50-50,100-100,150-150,200-200;
    '                               000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,001,001,001,001,002,002,001,001,001,002,002,002,003,004,005,006,008,011,014,018,022,030,038,048,058,072,089,107,124,145,162,180,196,208,221,230,239,247,248,251,255,247,246,233,229,221,204,199,188,180,169,156,150,141,130,125,116,111,104,097,093,088,085,079,074,071,067,063,061,059,056,054,053,051,048,045,044,040,038,037,034,033,031
    '                               ,030,028,026,025,022,020,019,018,017,016,015,015,013,013,012,011,011,011,011,010,010,009,009,008,008,008,008,008,008,008,007,008,008,007,008,007,007,007,007,007,007,007,007,006,006,006,005,005,005,005,005,005,005,005,005,004,004,005,004,004,004,004,004,004,003,003,003,003,003,002,002,002,002,002,002,002,002,002,002,002,002,002,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,001,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,000,001,001,001,001,001,000,000,000,000,000,000,000,000;55,90
    '                     6) ֱ��ͼ (ͬһ��ͼ���ϻ�����������)������;ͼ������;��������;~(��һ������)Y1;Y2;Y3;Y4;Y5...~(�ڶ�������)Y1;Y2;Y3;Y4;Y5..
    '                                ���� �������ݣ�����߶�,���᳤��,�̶�1-��ʾֵ,�̶�2-��ʾֵ,...
    '                                     �������֣�ÿ������������ '~' �ſ�ʼ�Ա����ֲ�ͬ����
    '                        ����WBC;6;0,80;~0;0;0;0;1;1;2;3;4;6;9;13;18;23;27;31;32;30;28;24;21;17;13;11;8;6;5;4;3;2;1;1;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0~0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;1;2;2;3;4;5;5;6;6;6;6;5;4;4;3;2;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0~0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;1;1;2;2;3;5;7;9;13;17;21;25;28;31;34;37;38;38;38;37;35;33;31;29;26;24;22;19;16;13;11;8;6;4;3;2;1;1;1;1;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0
    '
    '                   100) ͼƬ����:ͼ������;ͼ�񻭷�;[��ȡ���ݺ��Ƿ�ɾ��];ȫ·��
    '                        ��:WBC Fsc;100;1;C:\tempfile.gif
    '
    '                   101-227) ͼƬ����:ͼ������;ͼ�񻭷�;[��ȡ���ݺ��Ƿ�ɾ��];ȫ·��
    '                            ����֧������ͼ�θ�ʽ,BMP,JPG,GIF
    '                            BMP��ʽͼƬ���� 100-107���
    '                            JPG��ʽͼƬ���� 110-117
    '                            GIF��ʽͼƬ���� 120-127
    '
    '                            '2��ʼ�ľ���ѹ������ZIPͼ��
    '                            BMP��ʽͼƬ���� 200-207���
    '                            JPG��ʽͼƬ���� 210-217
    '                            GIF��ʽͼƬ���� 220-227
    '
    '                            ??1-??7��ͼƬ���ݵĲ����ʽ��������ʾͼƬʱ�Ķ��뷽ʽ���á�
    '                            ����ָ��Chart�ؼ��� .ChartArea.Interior.Image.Layout ����
    '                            101= oc2dImageCentered 102=oc2dImageTiled 103=oc2dImageFitted 104=oc2dImageStretched
    '                            105=oc2dImageStretchedToWidth 106=oc2dImageStretchedToHeight 107=oc2dImageCropFitted
   
'    GetAnswerCmd        ����  �Զ����ƴ���ʽ�����Զ�Ӧ��ָ����Ҫ�̶�Ӧ��06��дΪ,06
  '
Public Sub WriteLog(ByVal strFunc As String, ByVal StrInput As String, ByVal strOutput As String)
    '------------------------------------------------------
    '--  ����:���ݵ��Ա�־,д��־����ǰĿ¼
    '------------------------------------------------------
    
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strfilename As String
    Dim objStream As textStream
    Dim objFileSystem As New FileSystemObject
    
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", 1)) = 1 Then
        If Dir(App.Path & "\����.TXT") = "" Then Exit Sub
    End If
    strfilename = App.Path & "\LisDev_" & Format(date, "yyyyMMdd") & ".LOG"
    
    If Not objFileSystem.FileExists(strfilename) Then Call objFileSystem.CreateTextFile(strfilename)
    Set objStream = objFileSystem.OpenTextFile(strfilename, ForAppending)
    
    strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
    objStream.WriteLine (String(50, "��"))
    objStream.WriteLine ("ִ��ʱ��:" & strDate & "�汾:" & App.major & "." & App.minor & "." & App.Revision)
    objStream.WriteLine ("����:" & strFunc)
    objStream.WriteLine ("  :" & StrInput)
    objStream.WriteLine ("  :" & strOutput)
    'objStream.WriteLine (String(50, "-"))
    objStream.Close
    Set objStream = Nothing
End Sub

Public Function GetStr_Section(ByVal strSource As String, ByVal strStart As String, ByVal strEnd As String) As String
    '���ܣ�ȡ�����ַ�֮������ݷ���,��ʼ�ַ��ͽ����ַ�������ͬ
    'strSource: Դ�ַ���
    'strStart : ��ʼ�ַ�
    'strEnd   �������ַ�
    '
    Dim lngLength As Long, strTmp As String, strTmpStart As String, i As Integer
    
    If strStart <> strEnd Then
        lngLength = InStr(strSource, strEnd) - InStr(strSource, strStart) + 1
    Else
        For i = -22350 To -22310
            strTmpStart = Chr(i)
            If InStr(strSource, strTmpStart) <= 0 And strStart <> strTmpStart Then
                Exit For
            End If
        Next
        strTmp = Mid(strSource, 1, InStr(strSource, strStart) - 1) & strTmpStart & Mid(strSource, InStr(strSource, strStart) + 1)
        lngLength = InStr(strTmp, strEnd) - InStr(strTmp, strTmpStart) + 1
    End If
    
    If lngLength < 0 Then
        GetStr_Section = Mid(strSource, InStr(strSource, strStart) + lngLength, Abs(lngLength))
    Else
        GetStr_Section = Mid(strSource, InStr(strSource, strStart), lngLength)
    End If
End Function
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Mid_bin(ByVal str_bin As String, lng_S As Long, Optional lng_len As Long = 0, Optional blnChar As Boolean = True) As String
    'ʮ�����ƴ���MID����
    'str_Bin :����Ķ��������ݣ���ʽΪ,FF,AA,03 �ԣ��ſ�ʼ�����ޣ���
    'lng_S   :��ʼλ��
    'lng_Len :ȡ�ĳ���
    'blnChar :�Ƿ�ת��Ϊ�ַ���ʽ����
    
    Dim varBin As Variant
    Dim lng_Loop As Long
    Dim str_Return As String

    If lng_len < 0 Then Exit Function
    If lng_S <= 0 Then Exit Function
    
    varBin = Split(str_bin, ",")
    
    If lng_S + lng_len - 1 > UBound(varBin) Then
        '����Ĵ�û����ô��
        Mid_bin = ""
        Exit Function
    End If

    If lng_len = 0 Then
        If blnChar Then
            For lng_Loop = lng_S To UBound(varBin)
                str_Return = str_Return & Chr("&H" & varBin(lng_Loop))
            Next
        Else
            str_Return = Mid(str_bin, lng_S * 3 - 2)
        End If

    Else
        If blnChar Then
            For lng_Loop = lng_S To lng_S + lng_len - 1
                str_Return = str_Return & Chr("&H" & varBin(lng_Loop))
            Next
        Else
            str_Return = Mid(str_bin, lng_S * 3 - 2, lng_len * 3)
        End If

    End If
    If str_Return <> "" Then Mid_bin = str_Return
    
End Function

Public Function Len_Bin(ByVal str_bin As String) As Long
    'ʮ�����ƴ���Len����
'    Dim varBin As Variant
'    varBin = Split(str_bin, ",")
'    Len_Bin = UBound(varBin)
    Len_Bin = Len(str_bin) / 3
End Function

Public Function Instr_Bin(ByVal str_bin As String, ByVal strChar As String, Optional ByVal lngStart As Long) As Long
    'ʮ�����Ƶ� Instr����
    Dim varBin As Variant
    Dim strFindChar As String
    Dim lngS As Long
    Dim i As Integer
    Dim strHex As String
    If Len(strChar) <= 0 Then Exit Function
    strFindChar = ""
    For i = 1 To Len(strChar)
        strHex = Hex(Asc(Mid(strChar, i, 1)))
        strFindChar = strFindChar & "," & IIf(Len(strHex) = 1, "0" & strHex, strHex)
    Next
    If lngStart > 0 Then
        lngS = InStr(lngStart + 2, str_bin, strFindChar)
    Else
        lngS = InStr(str_bin, strFindChar)
    End If
    If lngS > 0 Then
        lngS = lngS + 2
        strFindChar = Mid(str_bin, 1, lngS)
'        len(strfindchar)/3
'        varBin = Len(strFindChar) / 3 'Split(strFindChar, ",")
        Instr_Bin = Len(strFindChar) / 3
    End If
End Function

Public Function Replace_Bin(ByVal str_bin As String, ByVal strFind As String, ByVal strReplace As String) As String
    'ʮ�����Ƶġ�Replace
    Dim strFindBin As String
    Dim strReplaceBin As String
    Dim i As Long
    If str_bin = "" Then Exit Function
    
    If Len(strFindBin) <= 0 Then
        Replace_Bin = str_bin
        Exit Function
    End If
    For i = 1 To Len(strFind)
        strFindBin = strFindBin & "," & Asc(Mid(strFind, i, 1))
    Next
        
    If strReplace <> "" Then
        For i = 1 To Len(strReplace)
            strReplaceBin = strReplaceBin & "," & Asc(Mid(strReplace, i, 1))
        Next
    Else
        strReplaceBin = ""
    End If
    
    Replace_Bin = Replace(str_bin, strFindBin, strReplaceBin)
    
    
End Function


Public Sub Pause(ByVal PauseTime)
    '��ʱ,��λ��
    Dim Start As Currency
    Start = Timer   ' ���ÿ�ʼ��ͣ��ʱ�̡�
    Do While Timer < Start + PauseTime
       DoEvents   ' �������ø���������
    Loop
    
End Sub

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
Dim ptOrigin    As PointAPI     '���ֻ���ԭ��
Dim ptCenter    As PointAPI     '�������ĵ�.
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
                        .x = CenterX - (szText.Width / 2)
                        .Y = CenterY - (szText.Height / 2)
                    End With
                    
                    'ת�� CenterX, CenterY ����ṹ
                    '(��Ҫ���� RotatePoint).
                    With ptCenter
                        .x = CenterX
                        .Y = CenterY
                    End With
                    
                    '��ԭ��ѡ����ƥ��Ԥ��ѡ��
                    Call RotatePoint(ptCenter, ptOrigin, RotDegrees)
                
                    '���ڴ�ӡ��ת�ı������سɹ�/ʧ��
                    PrintRotText = (TextOut(hDC, ptOrigin.x, _
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
'**�� �� ����RotatePoint
'**��    �룺ptAxis(PointAPI)   -
'**        ��ptRotate(PointAPI) -
'**        ��fDegrees(Single)   -
'**��    ������
'**������������ǰfdegrees��ǰ����ѡ��ptRotate���ҵ�ptAxis
'**ȫ�ֱ�����
'**����ģ�飺
'*************************************************************************
Private Sub RotatePoint(ptAxis As PointAPI, ptRotate As PointAPI, fDegrees As Single)

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
    fDX = ptRotate.x - ptAxis.x
    fDY = ptRotate.Y - ptAxis.Y
    
    '��ת��
    ptRotate.x = ptAxis.x + ((fDX * Cos(fRads)) + (fDY * Sin(fRads)))
    ptRotate.Y = ptAxis.Y + -((fDX * Sin(fRads)) - (fDY * Cos(fRads)))
    
End Sub




Public Function ReadIni(strItem As String, strKey As String, strPath As String, Optional strDefault As String = "") As String
    Dim GetStr As String
    On Error GoTo errH

    GetStr = VBA.String(128, 0)
    GetPrivateProfileString strItem, strKey, strDefault, GetStr, 256, strPath
    GetStr = VBA.Replace(GetStr, VBA.Chr(0), "")
    ReadIni = GetStr
    Exit Function
errH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(strItem As String, strKey As String, strVal As String, strPath As String) As Boolean
    On Error GoTo errH
    WriteIni = True
    WritePrivateProfileString strItem, strKey, strVal, strPath
    Exit Function
errH:
    Err.Clear
    WriteIni = False
End Function
Public Function DelSapce(strLine As String) As String
    '����       ɾ������Ŀո�
    Dim intLoop  As Integer
    Dim strNow As String
    strNow = strLine
    For intLoop = 20 To 0 Step -1
        strNow = Replace(strNow, Space(intLoop), Space(1))
    Next
    DelSapce = strNow
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
108       lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap)
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
194     WriteLog "mdlPublic.SavePic", CStr(Erl()) & "��", Err.Description
End Sub


'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath) <> "" Then
        zlFileUnzip = strZipPath & Dir(strZipPath)
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String, ByVal strfilename As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function


Public Function GetIniKeyValue(ByVal strPathAndFileName As String, ByVal strItem As String, ByVal strKey As String, Optional ByVal strDefault As String) As String
        '��ȡIni�ļ��е�ֵ
        '�����ļ��������򴴽�����д��Ĭ��ֵ
        Dim objFile As New FileSystemObject
        On Error GoTo hErr
100     If Not objFile.FileExists(strPathAndFileName) Then
102         Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
104         GetIniKeyValue = strDefault
        Else
106         GetIniKeyValue = ReadIni(strItem, strKey, strPathAndFileName)
            If GetIniKeyValue = "" And strDefault <> "" Then
                Call WriteIni(strItem, strKey, strDefault, strPathAndFileName)
                GetIniKeyValue = strDefault
            End If
        End If
        Exit Function
hErr:
108     WriteLog "��ȡ" & strPathAndFileName & "�е�����," & CStr(Erl()) & "��, " & Err.Description, "", ""
End Function

