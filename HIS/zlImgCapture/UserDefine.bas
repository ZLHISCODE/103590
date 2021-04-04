Attribute VB_Name = "UserDefine"
Option Explicit
'----------������������
Public gstrSysName As String
Public gcnOracle As ADODB.Connection
Public gstrSQL As String

Public Const ATTR_������� As String = "Study Date"
Public Const ATTR_���ʱ�� As String = "Study Time"
Public Const ATTR_�������� As String = "Series Date"
Public Const ATTR_����ʱ�� As String = "Series Time"
Public Const ATTR_Ӱ����� As String = "Modality"
Public Const ATTR_�豸�� As String = "Manufacturer"
Public Const ATTR_����豸 As String = "Manufacturer's Model Name"
Public mlngAdviceID As Long, mlngSendNO As Long                         'ҽ��ID�����ͺ�
Dim iNet As New clsFtp

Public Type AutoRoutSetting                        '�洢�Զ�·�����õ�����
    type As Long                                    '1--����Ӱ���������Զ�·�ɣ�2--���ռ���豸�����Զ�·��
    strCondition As String                          '�������ݣ�Ӱ���������豸����
    strFTPDeviceNo As String                        '�Զ�·�ɵġ�Ŀ���豸�š�
End Type
Public aAutoRoutSetting() As AutoRoutSetting       '�洢�Զ�·�����õ�����
'------------------UserDefine-----------------
Public Const NUMINFILE = 30
Public Const SIZEOFLOGPALETTE = 8
Public Const SIZEOFPALETTEENTRY = 4
Public Const BZZ = 1769472

Public DefaultTitle As String

Public elapsed As Long
Public numframe As Long
Public stem As String * 256
Public frm As Integer
Public myhPalette As Long

Public MaxBoard As Integer
Public BoardName(4) As String
Public BoardTypeCode(4) As Integer


'----global variables for board--------
Public hBoard As Long               '�ɼ������
Public bActive As Byte              'ʵʱ�ɼ���־
Public total As Integer             'ȫ���ɼ�������
Public iCurrUsedNo As Long          '��ǰ�ɼ���

Public hAudio As Long
Public elapsedtimes As Long

Public SQFILE As String             '�ɼ��ļ�Ĭ������
Public dwMaxMemSize As Long         '����ڴ��С
Public dwBufSize As Long            '��󻺴��С

Public lpbi As BITMAPINFO           '�ɼ�Ŀ�����ʽ�ļ�ͷ
Public lpdib As Long                '�ɼ�Ŀ��������
Public lpMemory As Long             '�û��ڴ�ָ��

Public buf(5000000) As Byte
Public blk As tpBlockInfo           '�û��ṹָ��ṹ����

Public iNumImage As Integer         '�ɼ�֡��
Public iNum As Integer              '�࿨���ͨ���ɼ�����
Public NoCapture As Integer
Public lphBoard(5) As Long
Public iVirtCode As Integer

Public bMakeMirror As Integer       '���x,y������
Public iClientWidth As Integer      '�ͻ������
Public iClientHeight As Integer     '�ͻ����߶�

Public bMaskMode As Integer         'λ����״̬


Public sampwidth As Integer         '�ɼ����ڿ��
Public sampheight As Integer        '�ɼ����ڸ߶�

Public times As Long
Public num As Long
Public ratio As Single

Public hFile As Long

'-------------------Win Define-------------------
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_DDESHARE = &H2000

Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const OF_EXIST = &H4000
Public Const OFS_MAXPATHNAME = 128
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const BI_RGB = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const LMEM_FIXED = &H0
Public Const HELP_KEY = &H101            '  Display topic for keyword in offabData
Public Const HELP_FINDER = &HB
Public Const HELP_HELPONHELP = &H4       '  Display help on using help
Public Const BI_bitfields = 3&
Public Const HWND_DESKTOP = 0

Public Const CF_DIB = 8


Public Const PM_REMOVE = &H1

Public Const WM_KEYDOWN = &H100
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const GWL_USERDATA = (-21)

Public Const SW_SHOWNORMAL = 1
Public Const SW_RESTORE = 9

Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Type BITMAPINFOHEADER '40 bytes
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

Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved As Byte
End Type

Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(256) As PALETTEENTRY
End Type

Type POINTAPI
        x As Long
        y As Long
End Type

Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type

Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type



Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wbytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long

Declare Function GetVersion Lib "kernel32" () As Long
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function okOpenReplayDlgEx Lib "okapi32.dll" (ByVal hBoard As Long, ByVal hwnd As Long, ByVal src As Any, ByVal total As Long, lpbi As BITMAPINFOHEADER, lpdib As Byte) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal numBytes As Long, ByVal fill As Byte)
 
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long


Function exCalcCaptureRatio() As Single
    If num >= 12 Then
        times = GetTickCount() - times
        ratio = num * 1000 / times
        num = 0
    End If

    If num = 0 Then
        times = GetTickCount()
    End If
    num = num + 1

    exCalcCaptureRatio = ratio
End Function

Function exfunMin(ByVal Param1 As Long, ByVal Param2 As Long) As Long
    '����Param1��Param2�н�С������
    If Param1 > Param2 Then
        exfunMin = Param2
    Else
        exfunMin = Param1
    End If
End Function


Function exfunMax(ByVal Param1 As Long, ByVal Param2 As Long) As Long
    '����Param1��Param2�нϴ������
    If Param1 > Param2 Then
        exfunMax = Param1
    Else
        exfunMax = Param2
    End If
End Function

Function exGetBitmapdata(ByVal hBoard As Long, ByVal tgt As Long, ByVal start As Integer, lpbi As BITMAPINFO, lpdib As Byte) As Long
    Dim rRect As RECT
    Dim bits As Integer
    Dim extend As Integer
    Dim frm As Long
    Dim wbytes As Long
    Dim i As Integer
    Dim l As Long
    
    If hBoard = 0 Then
        exGetBitmapdata = 0
        Exit Function
    End If
    
    extend = okSetConvertParam(hBoard, CONVERT_FIELDEXTEND, -1)
    
    'get dib header info
    If tgt = VSCREEN Or tgt = BUFFER Then
        Dim blk As tpBlockInfo
        
        blk.lpBits = lpdib
    
        rRect.Right = -1
        If tgt = VSCREEN Then
            frm = okSetCaptureParam(hBoard, CAPTURE_SCRRGBFORMAT, GETCURRPARAM) '-1
        Else
            If tgt = BUFFER Then
                frm = okSetCaptureParam(hBoard, CAPTURE_BUFRGBFORMAT, GETCURRPARAM) '-1
            End If
        End If
        bits = exHiWord(frm)
        frm = exLoWord(frm)
        okSetTargetRect hBoard, tgt, rRect 'get current rect
        lpbi.bmiHeader.biWidth = rRect.Right - rRect.Left
        lpbi.bmiHeader.biHeight = rRect.Bottom - rRect.Top
        If tgt = VSCREEN Then 'after frozen
            rRect.Right = -1 'max. captured rect
            okSetTargetRect hBoard, VIDEO, rRect 'get video rect
            lpbi.bmiHeader.biWidth = exfunMin(lpbi.bmiHeader.biWidth, rRect.Right - rRect.Left)
            lpbi.bmiHeader.biHeight = exfunMin(lpbi.bmiHeader.biHeight, rRect.Bottom - rRect.Top)
        End If
    
    'get image to app buffer from tgt
        blk.iWidth = lpbi.bmiHeader.biWidth
        blk.iHeight = -lpbi.bmiHeader.biHeight 'to invert y
        If frm = FORM_GRAY10 Then 'special consider for high gray
            blk.iFormType = FORM_GRAY8 'must do this
            blk.iBitCount = 8
            ''okConvertRect hBoard,(TARGET)&blk,0,tgt,start,1)
        Else
            blk.iFormType = frm
            blk.iBitCount = bits
            okReadRect hBoard, tgt, start, buf(0)
            ''okTransferRect hBoard, l, 0, tgt, start, 1
        End If
    Else
        If tgt > BUFFER Then
            Dim lpblk As tpBlockInfo
            
    '        lpblk=(LPBLOCKINFO)tgt
    
            If lpblk.iBitCount = 10 Then
                bits = 8
            Else
                bits = lpblk.iBitCount
            End If
            lpbi.bmiHeader.biWidth = lpblk.iWidth
            lpbi.bmiHeader.biHeight = Abs(lpblk.iHeight)
    
            If lpdib <> lpblk.lpBits Then 'not same
    '            CopyMemory(lpdib,lpblk->lpBits,
    '                lpblk->iWidth*lpblk->iHeight*lpblk->iBitCount/8);
            End If
        End If
    End If

    lpbi.bmiHeader.biSize = 40
    lpbi.bmiHeader.biPlanes = 1
    If (frm = FORM_GRAY10 Or frm = FORM_GRAY12) Then 'if bit=10,12 convert to 8
        bits = 8
    End If

    lpbi.bmiHeader.biBitCount = bits
    wbytes = (lpbi.bmiHeader.biWidth * bits) + 31 'align
    For i = 1 To 5
        wbytes = wbytes \ 2
    Next
    For i = 1 To 2
        wbytes = wbytes * 2
    Next
    lpbi.bmiHeader.biSizeImage = wbytes * lpbi.bmiHeader.biHeight
    lpbi.bmiHeader.biClrUsed = 0

'    //special consider for 565 & 32
    If (frm = FORM_RGB565 Or frm = FORM_RGB8888) Then
        lpbi.bmiHeader.biCompression = BI_bitfields
    Else
        lpbi.bmiHeader.biCompression = 0
    End If
    If lpbi.bmiHeader.biCompression = BI_bitfields Then '565
        'DWORD   *lpmask;
'        lpmask=(DWORD *)((LPSTR)lpbi+lpbi->biSize);

'        if(bits==16) {
'            lpmask[2]=0x001f; //blue
'            lpmask[1]=0x07e0;
'            lpmask[0]=0xf800;
'        }
'        else if(bits==32) {
'            lpmask[2]=0x0000ff;
'            lpmask[1]=0x00ff00;
'            lpmask[0]=0xff0000;
'        }
'    }
    Else
        If (bits <= 8) Then
            lpbi.bmiHeader.biClrUsed = 256
            For i = 0 To lpbi.bmiHeader.biClrUsed - 1
                lpbi.bmiColors(i).rgbBlue = i
                lpbi.bmiColors(i).rgbGreen = i
                lpbi.bmiColors(i).rgbRed = i
                lpbi.bmiColors(i).rgbReserved = i
            Next
        End If
    End If
    lpbi.bmiHeader.biClrImportant = lpbi.bmiHeader.biClrUsed

    exGetBitmapdata = frm
End Function


Function exGetTargetSize(ByVal hBoard As Long, ByVal tgt As Long, wid As Integer, hei As Integer) As Long
    'get size
    Dim vrect As RECT
    Dim frm As Long
    
    
    If ((tgt = VSCREEN) Or (tgt = BUFFER)) Then
        vrect.Right = -1
        okSetTargetRect hBoard, tgt, vrect 'get current rect
        wid = vrect.Right - vrect.Left
        hei = vrect.Bottom - vrect.Top
    
        If tgt = VSCREEN Then
            frm = okSetCaptureParam(hBoard, CAPTURE_SCRRGBFORMAT, GETCURRPARAM) '-1
            'limit to video rect
            vrect.Right = -1 'max. captured rect
            okSetTargetRect hBoard, VIDEO, vrect 'get video rect
            wid = exfunMin(wid, vrect.Right - vrect.Left)
            hei = exfunMin(hei, vrect.Bottom - vrect.Top)
        Else
            If tgt = BUFFER Then
                frm = okSetCaptureParam(hBoard, CAPTURE_BUFRGBFORMAT, GETCURRPARAM) '-1
            End If
        End If
    Else
        If tgt > BUFFER Then 'from blkinfo
    ''        LPBLOCKINFO lpblk;
            ''lpblk=(LPBLOCKINFO)tgt;
    
            ''*width=lpblk->iWidth;
            ''*height=abs(lpblk->iHeight);
            ''form=MAKELONG(lpblk->iFormType,lpblk->iBitCount);
        End If
    End If
    
    
    exGetTargetSize = frm
End Function

Function exHiWord(ByVal Value As Long) As Integer
    'ȡ��Value�ĸ���
    Dim k As Long
    Dim i As Long
    Dim j As Long
    
    k = Value
    For i = 1 To 8
        k = k \ 4
    Next
    j = k: If j > 32767 Then j = j - 65536
    exHiWord = j
End Function

Function exLoByte(ByVal Value As Integer) As Byte
    'ȡ��Value�ĵ�λ
    Dim k As Long
    
    k = Value
    Do While (k > 255)
        k = k - 256
    Loop
    exLoByte = k
End Function

Function exLoWord(ByVal Value As Long) As Integer
    'ȡ��Value�ĵ���
    Dim k As Long
    Dim i As Long
    Dim j As Long
    
    k = Value
    For i = 1 To 8
        k = k \ 4
    Next
    For i = 1 To 8
        k = k * 4
    Next
    j = Value - k: If j > 32767 Then j = j - 65536
    exLoWord = j
End Function

Function exMakeLogPalette(ByVal iBits As Integer, ByVal rgbForm As Integer) As Long
    Dim hLogPal As Long
    Dim npPal As LOGPALETTE
    Dim nNumColors As Integer
    Dim i As Integer
    
    nNumColors = 1
    For i = 1 To iBits
        nNumColors = nNumColors * 2
    Next
    'npPal = LocalAlloc(LMEM_FIXED, SIZEOFLOGPALETTE + nNumColors * SIZEOFPALETTEENTRY)
    'If npPal = 0 Then exMakeLogPalette = 0: Exit Function
    
    npPal.palVersion = &H300
    npPal.palNumEntries = nNumColors
    'set palette
    If rgbForm = FORM_RGB332 Then 'rgb 332
        Dim red As Byte
        Dim green As Byte
        Dim blue As Byte
                
        red = 0
        green = 0
        blue = 0
        'For i = 0 To nNumColors - 1
        For i = 0 To 1
                npPal.palPalEntry(i).peBlue = blue
                npPal.palPalEntry(i).peGreen = green
                npPal.palPalEntry(i).peRed = red
                npPal.palPalEntry(i).peFlags = 0
                If red = 0 Then
                    red = red + 32
                    If green = 0 Then
                        green = green + 32
                        blue = blue + 64
                    End If
                End If
        Next
    Else  'if(rgbForm==FORM_GRAY8) { //gray 256
        'For i = 0 To nNumColors - 1
    
        For i = 0 To nNumColors - 1
                npPal.palPalEntry(i).peRed = i
                npPal.palPalEntry(i).peGreen = i
                npPal.palPalEntry(i).peBlue = i
                npPal.palPalEntry(i).peFlags = 0
        Next
    End If
    
    hLogPal = CreatePalette(npPal)
    exMakeLogPalette = hLogPal
End Function

Function exMakeLong(ByVal Param1 As Integer, ByVal Param2 As Long) As Long
    Dim i As Long
    
    i = Param2
    exMakeLong = i * &H10000 + Param1
End Function


Function exSetBitmapHeader(lpbi As BITMAPINFO, ByVal wid As Integer, ByVal hei As Integer, ByVal bits As Integer, ByVal frm As Integer) As Long
    'set bitmap header and bitmap info if need
    Dim wbytes As Long
    Dim i As Integer
    
    If frm = FORM_RGB555 Then
        bits = 16
    End If
    wbytes = (lpbi.bmiHeader.biWidth * bits) + 31 'align
    For i = 1 To 5
        wbytes = wbytes \ 2
    Next
    For i = 1 To 2
        wbytes = wbytes * 2
    Next
    'wbytes = width * bits
    'If (wbytes Mod 4) <> 0 Then
    '    wbytes = ((wbytes \ 4) + 1) * 4
    'End If
        
    lpbi.bmiHeader.biWidth = wid
    lpbi.bmiHeader.biHeight = hei
    
    lpbi.bmiHeader.biSize = 40
    lpbi.bmiHeader.biPlanes = 1
    
    lpbi.bmiHeader.biBitCount = bits
    lpbi.bmiHeader.biSizeImage = wbytes * lpbi.bmiHeader.biHeight
    
    lpbi.bmiHeader.biClrUsed = 0
    'special format for 555,565 & 32
    If (frm = FORM_RGB555 Or frm = FORM_RGB565 Or frm = FORM_RGB8888) Then
        lpbi.bmiHeader.biCompression = BI_bitfields
    Else
        lpbi.bmiHeader.biCompression = 0
    End If
    If lpbi.bmiHeader.biCompression = BI_bitfields Then
        If frm = FORM_RGB555 Then
            lpbi.bmiColors(0).rgbBlue = 0
            lpbi.bmiColors(0).rgbGreen = &H7C
            lpbi.bmiColors(0).rgbRed = 0
        
            lpbi.bmiColors(1).rgbBlue = &HE0
            lpbi.bmiColors(1).rgbGreen = &H3
            lpbi.bmiColors(1).rgbRed = 0
        
            lpbi.bmiColors(2).rgbBlue = &H1F
            lpbi.bmiColors(2).rgbGreen = 0
            lpbi.bmiColors(2).rgbRed = 0
        Else
            If frm = FORM_RGB565 Then
                lpbi.bmiColors(0).rgbBlue = 0
                lpbi.bmiColors(0).rgbGreen = &HF8
                lpbi.bmiColors(0).rgbRed = 0
            
                lpbi.bmiColors(1).rgbBlue = &HE0
                lpbi.bmiColors(1).rgbGreen = &H7
                lpbi.bmiColors(1).rgbRed = 0
            
                lpbi.bmiColors(2).rgbBlue = &H1F
                lpbi.bmiColors(2).rgbGreen = 0
                lpbi.bmiColors(2).rgbRed = 0
            Else
                If bits = 32 Then
                    lpbi.bmiColors(0).rgbBlue = 0
                    lpbi.bmiColors(0).rgbGreen = 0
                    lpbi.bmiColors(0).rgbRed = &HFF
                
                    lpbi.bmiColors(1).rgbBlue = 0
                    lpbi.bmiColors(1).rgbGreen = &HFF
                    lpbi.bmiColors(1).rgbRed = 0
                
                    lpbi.bmiColors(2).rgbBlue = &HFF
                    lpbi.bmiColors(2).rgbGreen = 0
                    lpbi.bmiColors(2).rgbRed = 0
                End If
            End If
        End If
    Else
        If bits <= 12 Then ' 8,10,12
            lpbi.bmiHeader.biClrUsed = 2 ^ (bits - 1)
            For i = 0 To lpbi.bmiHeader.biClrUsed - 1
                lpbi.bmiColors(i).rgbBlue = i
                lpbi.bmiColors(0).rgbGreen = i
                lpbi.bmiColors(0).rgbRed = i
                lpbi.bmiColors(0).rgbReserved = 0
            Next
        End If
    End If
    lpbi.bmiHeader.biClrImportant = lpbi.bmiHeader.biClrUsed
    
    exSetBitmapHeader = lpbi.bmiHeader.biClrUsed
End Function

Function exSetDataToDIB(ByVal src As Long, ByVal start As Integer, lpbi As BITMAPINFO, lpdib As Byte) As Long
    'set data to dib
    Dim blk As tpBlockInfo
    
    'get image to app buffer from tgt
    blk.lpBits = VarPtr(lpdib)
    blk.iWidth = lpbi.bmiHeader.biWidth
    blk.iHeight = -lpbi.bmiHeader.biHeight '!!!to invert y
    blk.iBitCount = lpbi.bmiHeader.biBitCount
    blk.lBlockStep = exLoWord(lpbi.bmiHeader.biSizeImage) 'must be set
    blk.iHiStep = exHiWord(lpbi.bmiHeader.biSizeImage) 'must be set
    If (lpbi.bmiHeader.biCompression = BI_bitfields) Then
        If lpbi.bmiColors(0).rgbGreen = &H3E0 Then '555
            blk.iFormType = FORM_RGB555
        Else
            If lpbi.bmiHeader.biBitCount = 16 Then
                blk.iFormType = FORM_RGB565
            End If
        End If
    End If
    
    exSetDataToDIB = okConvertRect(hBoard, VarPtr(blk), 0, src, start, 1)
End Function

Sub exSleep(Value As Integer)
    Dim i As Long
    
    For i = 1 To Value
        DoEvents
    Next
End Sub

Function exGetBitmapHeader(ByVal hBoard As Long, ByVal src As Long, lpbi As BITMAPINFO) As Integer
    Dim wid As Integer
    Dim hei As Integer
    Dim frm As Long
    
    frm = exGetTargetSize(hBoard, src, wid, hei)
    
    'here take form as screen forever
    If src <= 1 Then
        frm = okSetCaptureParam(hBoard, CAPTURE_SCRRGBFORMAT, GETCURRPARAM) '-1
    End If
    
    exSetBitmapHeader lpbi, wid, hei, exHiWord(frm), exLoWord(frm)
    
    exGetBitmapHeader = exLoWord(frm)
End Function




Function exConvertBitmap(lpbi As BITMAPINFOHEADER, lpbuf() As Byte) As Boolean
    'up-down image data for dib
    Dim i As Long
    Dim wbytes As Long, j As Long
    Dim lptop As Long
    Dim lpbottom As Long
    
    If lpbi.biSize <> 40 Then
        exConvertBitmap = False
        Exit Function
    End If
    If lpbi.biHeight = 0 Then
        exConvertBitmap = False
        Exit Function
    End If
    
    'inverse dib (top to bottom)
    wbytes = lpbi.biSizeImage \ lpbi.biHeight
    lptop = 0
    lpbottom = (lpbi.biHeight - 1) * wbytes
    For i = 0 To lpbi.biHeight \ 2 - 1
        For j = 0 To wbytes - 1
            exSwap buf(lptop + j), buf(lpbottom + j)
        Next
        lptop = lptop + wbytes
        lpbottom = lpbottom - wbytes
    Next
    exConvertBitmap = True
End Function


Sub exSwap(Param1 As Byte, Param2 As Byte)
    Dim dd As Byte
    
    dd = Param1
    Param1 = Param2
    Param2 = dd
End Sub


Function exTransName(Param() As Byte) As String
    Dim stem As String
    Dim i  As Integer
    
    stem = ""
    i = 0
    Do While (Param(i) <> 0)
        stem = stem + Chr(Param(i))
        i = i + 1
    Loop
    exTransName = stem
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strFormCaption As String)
'���ܣ��򿪼�¼��ͬʱ����SQL���
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String)
'���ܣ�ִ�й���ʽ��SQL���
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub SaveImages(Images As DicomImages, ByVal MainDeviceID As String, ByVal BufferDir As String, Optional iEncode As Integer = 0, Optional ByVal strImgType As String = "")
'���ܣ�����ͼ��
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer  '�����ͼ����
    Dim intSQL As Integer, rsTmp As New ADODB.Recordset
    
    Dim blnAddTmp As Boolean, blnTmp As Boolean
    Dim strAge As String, strBirth As String
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String, dtCurrent As String
    Dim strUser As String, strPwd As String
    Dim ImageType As String, CheckNo As Long, CheckDev As String
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    Dim CheckUID As String, SeriesUID As String
    Dim aPatientID() As String, lngAdviceID As Long, lngSendNO As Long 'ͼ���е�ҽ��ID��ҽ��ID_���ͺ�
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    
    dtCurrent = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select 'ftp://'||Decode(�û���,Null,'',�û���||Decode(����,Null,'',':'||����))" & _
        "||'@'||IP��ַ As Host,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL " & _
        "From Ӱ���豸Ŀ¼ " & _
        "Where �豸��=[1]"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", MainDeviceID)
    If rsTmp.EOF Then
        err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "�豸�����ô���"
    End If
    strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
    GetFtpAddress strHost, strHost, strUser, strPwd
    iNet.FuncFtpConnect strHost, strUser, strPwd
    iCount = 0
    If Images.count > 0 Then
        gstrSQL = "Select a.ҽ��ID,a.���ͺ�,a.Ӱ�����,a.����,a.����,a.Ӣ����,a.�Ա�,a.����,a.��������,a.���,a.����,a.������,a.���Ž�Ƭ,����豸,��������,c.ͼ��UID,d.ִ�м� " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,����ҽ������ d " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And a.ҽ��ID=d.ҽ��ID And a.���ͺ�=d.���ͺ� And a.���UID=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(Images(1).StudyUID))
        If Not rsTmp.EOF Then
            lngAdviceID = Nvl(rsTmp("ҽ��ID"), 0)
            lngSendNO = Nvl(rsTmp("���ͺ�"), 0)
        End If
        'ɾ��ͼ���ļ�
        Do While Not rsTmp.EOF
            RemoveFromURL strHost, strDirURL & _
                Format(Nvl(rsTmp("��������"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
                Images(1).StudyUID & "/" & rsTmp("ͼ��UID")
            rsTmp.MoveNext
        Loop
        '���¿�ʼ���
        If lngAdviceID > 0 Then
            rsTmp.MoveFirst
            
            gstrSQL = "ZL_Ӱ����_CANCEL(" & Nvl(rsTmp("ҽ��ID"), 0) & "," & Nvl(rsTmp("���ͺ�"), 0) & ")"
            ExecuteProcedure "PACSͼ�񱣴�"
            gstrSQL = "ZL_Ӱ����_BEGIN('" & Nvl(rsTmp("ִ�м�")) & "'," & Nvl(rsTmp("����"), 0) & "," & rsTmp("ҽ��ID") & "," & rsTmp("���ͺ�") & ",'" & Nvl(rsTmp("Ӱ�����")) & "','" & _
                Nvl(rsTmp("����")) & "','" & Nvl(rsTmp("Ӣ����")) & "','" & Nvl(rsTmp("�Ա�")) & "','" & _
                Nvl(rsTmp("����")) & "'," & IIf(IsNull(rsTmp("��������")), "Null", "to_Date('" & Format(rsTmp("��������"), "yyyy-MM-dd") & "','YYYY-MM-DD')") & ",'" & Nvl(rsTmp("���")) & "','" & Nvl(rsTmp("����")) & "'," & _
                Nvl(rsTmp("������"), 0) & "," & Nvl(rsTmp("���Ž�Ƭ"), 0) & ",'" & Nvl(rsTmp("����豸")) & "')"
            ExecuteProcedure "PACSͼ�񱣴�"
        End If
    End If
    
    For Each curImage In Images
        gstrSQL = "Select ͼ��UID From Ӱ����ͼ�� Where ͼ��UID=[1]" & _
            " Union All Select ͼ��UID From Ӱ����ʱͼ�� Where ͼ��UID=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(curImage.InstanceUID))
        '��ͼ��
        If rsTmp.EOF Then
            gstrSQL = "Select ���UID From Ӱ�����¼ Where ���UID=[1]" & _
                " Union All Select ���UID From Ӱ����ʱ��¼ Where ���UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(curImage.StudyUID))
            '������ID��Ӣ��������
            If rsTmp.EOF Then
                blnAddTmp = True
                aPatientID = Split(curImage.PatientID, "_")
                If UBound(aPatientID) >= 0 And lngAdviceID = 0 Then
                    lngAdviceID = Val(aPatientID(0)) ': lngSendNO = Val(aPatientID(1))
                End If
                gstrSQL = "Select Distinct A.ҽ��ID,A.���ͺ� From Ӱ�����¼ A,����ҽ������ B,����ҽ����¼ C" & _
                    " Where A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID" & _
                    " And A.ҽ��ID=[1]" & _
                    " And B.ִ��״̬=3 And B.ִ�й���=2"
                If rsTmp.State <> adStateClosed Then rsTmp.Close
                Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", lngAdviceID)
                '��HIS��д�ļ���¼��Ӧ
                If Not rsTmp.EOF Then
                    '������UID
                    gstrSQL = "ZL_Ӱ�����¼_SET(" & rsTmp(0) & "," & rsTmp(1) & ",'" & _
                        curImage.StudyUID & "','" & GetImageAttribute(curImage.Attributes, ATTR_����豸) & "'," & _
                        "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACSͼ�񱣴�"
                    blnAddTmp = False
                End If
                '������ʱ����¼
                If blnAddTmp Then
                    If IsDate(curImage.DateOfBirthAsDate) Then
                        strAge = CStr(Year(Date) - Year(curImage.DateOfBirthAsDate))
                        strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                    Else
                        strAge = "": strBirth = ""
                    End If
                    gstrSQL = "ZL_Ӱ����ʱ���_INSERT('" & strImgType & "',Null,'" & _
                        curImage.Name & "','" & curImage.Name & "','" & _
                        curImage.Sex & "','" & strAge & "'," & _
                        IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                        GetImageAttribute(curImage.Attributes, ATTR_����豸) & "','" & curImage.StudyUID & "'," & _
                        "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACSͼ�񱣴�"
                End If
            End If
            
            gstrSQL = "Select 0 As ��ʱ,��������,Ӱ�����,Nvl(����,0) As ����," & _
                "����豸,����,Ӣ����,�Ա�,Nvl(����,'-1') As ����,���UID From Ӱ�����¼ Where ���UID=[1]" & _
                " Union All Select 1 As ��ʱ,��������,Ӱ�����,Nvl(����,0) As ����," & _
                "����豸,����,Ӣ����,�Ա�,Nvl(����,'-1') As ����,���UID From Ӱ����ʱ��¼ Where ���UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(curImage.StudyUID))
            blnTmp = IIf(rsTmp(0) = 1, True, False) '���к�ͼ���Ƿ������ʱ��¼��
            dtReceived = Format(rsTmp(1), "yyyyMMdd")
            
            ImageType = Nvl(rsTmp(2)): CheckNo = rsTmp(3): CheckDev = Nvl(rsTmp(4))
            PatientName = Nvl(rsTmp(5)): EnglishName = Nvl(rsTmp(6)): Sex = Nvl(rsTmp(7)): Age = Val(rsTmp(8))
            CheckUID = Nvl(rsTmp(9))
            
            gstrSQL = "Select ����UID From " & IIf(blnTmp, "Ӱ����ʱ����", "Ӱ��������") & _
                " Where ����UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(curImage.SeriesUID))
            '�����µļ������
            If rsTmp.EOF Then
                gstrSQL = "ZL_Ӱ������_INSERT('" & curImage.StudyUID & "','" & curImage.SeriesUID & "','" & _
                    curImage.SeriesDescription & "'," & _
                    IIf(blnTmp, 1, 0) & ")"
                ExecuteProcedure "PACSͼ�񱣴�"
            End If
            
            '�����µ�ͼ��
            gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & curImage.InstanceUID & "','" & curImage.SeriesUID & "','" & _
                curImage.SeriesDescription & "'," & _
                IIf(blnTmp, 1, 0) & ")"
            ExecuteProcedure "PACSͼ�񱣴�"
            gstrSQL = "ZL_Ӱ���鱨��_ADD('" & curImage.StudyUID & "','" & curImage.InstanceUID & ".jpg')"
            ExecuteProcedure "���汨��ͼ��"
            
            '����ͼ�񵽻���Ŀ¼
            On Error Resume Next
            MkLocalDir BufferDir & dtReceived & "/" & curImage.StudyUID & "/"
            Select Case iEncode
                Case 1
                    curImage.WriteFile BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID, True, "1.2.840.10008.1.2.5"
                Case 2
                    curImage.WriteFile BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID, True
                Case Else
                    curImage.WriteFile BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            End Select
            curImage.FileExport BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID & ".jpg", "JPG"
            On Error GoTo DBError
            
            WriteToURL BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID, strHost, strDirURL & _
                dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID
            WriteToURL BufferDir & dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID & ".jpg", strHost, strDirURL & _
                dtReceived & "/" & curImage.StudyUID & "/" & curImage.InstanceUID & ".jpg"
        End If
        iCount = iCount + 1
    Next
    
    gcnOracle.CommitTrans
    iNet.FuncFtpDisConnect
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    err.Raise err.Number, "���ͼ�񱣴�"
End Sub

Public Sub WriteToURL(ByVal SrcFileName As String, ByVal DestAddress As String, ByVal DestFileName As String)
'���ܣ��������ļ����浽Զ��������
'    Dim iNet As New clsFtp, strHost As String, strUser As String, strPwd As String
    Dim objFileSystem As New Scripting.FileSystemObject
    
'    GetFtpAddress DestAddress, strHost, strUser, strPwd
'    iNet.strIPAddress = strHost: iNet.strUser = strUser: iNet.strPsw = strPwd
    
    MkDir_Remote DestAddress, DestFileName
    iNet.FuncUploadFile objFileSystem.GetParentFolderName(DestFileName), SrcFileName, objFileSystem.GetFileName(DestFileName)
End Sub

Public Sub MkDir_Remote(ByVal DestAddress As String, ByVal DestFileName As String)
'    Dim iNet As New clsFtp,
    Dim objFile As New Scripting.FileSystemObject, strPath As String
    Dim aNestPath() As Variant, i As Integer
'    Dim strHost As String, strUser As String, strPwd As String
    
    aNestPath = Array()
    
'    GetFtpAddress DestAddress, strHost, strUser, strPwd
'    iNet.strIPAddress = strHost: iNet.strUser = strUser: iNet.strPsw = strPwd
    
    strPath = objFile.GetParentFolderName(DestFileName)
    iNet.FuncFtpMkDir "/", strPath
End Sub

Public Sub RemoveFromURL(ByVal DestAddress As String, ByVal DestFileName As String)
'���ܣ��������ļ����浽Զ��������
'    Dim iNet As New clsFtp, strHost As String, strUser As String, strPwd As String
    Dim objFileSystem As New Scripting.FileSystemObject
    
'    GetFtpAddress DestAddress, strHost, strUser, strPwd
'    iNet.strIPAddress = strHost: iNet.strUser = strUser: iNet.strPsw = strPwd
    
    iNet.FuncDelFile objFileSystem.GetParentFolderName(DestFileName), objFileSystem.GetFileName(DestFileName)
End Sub
'--------------------------------------
'--��Ftp��ַ�ֽ�Ϊ�������û�������
'--------------------------------------
Private Sub GetFtpAddress(ByVal strFtpPath As String, strHost As String, strUser As String, strPwd As String)
    Dim iPos As Integer
    Dim aUser() As String
    On Error Resume Next
        
    iPos = InStr(strFtpPath, "@")
    If iPos = 0 Then '�޵�½�û���Ϣ
        strHost = Trim(Mid(strFtpPath, 7))
        If Right(strHost, 1) = "/" Then strHost = Mid(strHost, 1, Len(strHost) - 1)
        strUser = "": strPwd = ""
    Else
        strHost = Trim(Mid(strFtpPath, iPos + 1))
        If Right(strHost, 1) = "/" Then strHost = Mid(strHost, 1, Len(strHost) - 1)
        
        aUser = Split(Trim(Mid(strFtpPath, 7, iPos - 7)), ":")
        strUser = aUser(0): strPwd = aUser(1)
    End If
End Sub

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As Variant
    Dim curAttr As DicomAttribute
    
    GetImageAttribute = ""
    For Each curAttr In objAttr
        If UCase(curAttr.Description) = UCase(AttrName) Then
            If curAttr.Exists Then GetImageAttribute = curAttr.Value
            Exit For
        End If
    Next
End Function

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'���ܣ�����DicomViewer��������
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    If iCols < 1 Then iCols = 1
    If iRows < 1 Then iRows = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

Private Function funcAutoRouting(img As DicomImage, BufferDir As String, dtReceived As String, Optional iEncode As Integer = 0) As Long
    Dim i As Integer            '����ѭ���ı���
    Dim strImageType As String
    Dim strImageDevice As String
    Dim iRouting As Integer         '��ʶ���Զ�·�ɵĹ���ţ�0Ϊ������·������
    Dim rsTmp As New ADODB.Recordset
'    Dim strHost As String           'FTP�������û�������+ IP��ַ��
    Dim strDirURL As String         'FTP������Ŀ¼
    Dim strHost As String, strUser As String, strPwd As String
    Dim DestAddress As String
    
    Call subReadAutoRoutSetting
    
    iRouting = 0
    '��ȡͼ���Ӱ�����ͼ���豸��
    strImageType = GetImageAttribute(img.Attributes, ATTR_Ӱ�����)
    strImageDevice = GetImageAttribute(img.Attributes, ATTR_����豸)
    '�Աȴ洢���򣬲�ƥ�����˳�
    For i = 1 To UBound(aAutoRoutSetting)
        If aAutoRoutSetting(i).strCondition = IIf(aAutoRoutSetting(i).type = 1, strImageType, strImageDevice) Then
                iRouting = i
                Exit For
        End If
    Next
    '�洢ͼ��ָ��FTP�豸
    If iRouting <> 0 Then
        '��ȡĿ���豸��URL
        gstrSQL = "Select 'ftp://'||Decode(�û���,Null,'',�û���||Decode(����,Null,'',':'||����))" & _
        "||'@'||IP��ַ As Host,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL " & _
        "From Ӱ���豸Ŀ¼ " & _
        "Where �豸��=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", aAutoRoutSetting(iRouting).strFTPDeviceNo)
        If rsTmp.EOF Then
            err.Raise vbObjectError + 1, "PACSͼ�񱣴�", "�豸�����ô���"
        End If
        strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
        DestAddress = strHost & strDirURL
        '����ͼ��ָ��URL
        Select Case iEncode
            Case 0
                img.WriteFile BufferDir & img.InstanceUID, True, "1.2.840.10008.1.2.4.70"
            Case 1
                img.WriteFile BufferDir & img.InstanceUID, True, "1.2.840.10008.1.2.5"
            Case 2
                img.WriteFile BufferDir & img.InstanceUID, True
        End Select
        GetFtpAddress DestAddress, strHost, strUser, strPwd
        iNet.FuncFtpConnect strHost, strUser, strPwd
        WriteToURL BufferDir & img.InstanceUID, strHost, strDirURL & _
            dtReceived & "/" & img.StudyUID & "/" & img.InstanceUID
        Kill BufferDir & img.InstanceUID
        iNet.FuncFtpDisConnect
    End If
End Function

'��ȡ�Զ�·�ɵĹ���
Public Sub subReadAutoRoutSetting()
''''''''''''''''''''''''''''''''''''''''''''''''''''
'''��ע����"ZLSOFT\����ģ��\��Ʒ��\���շ���\�Զ�·��"�ж�ȡ�Զ�·�ɵĹ�������
'''��������Ϊ�����ı�ǣ�ʹ��Ӣ�Ķ��ŷֿ���
'''���1--�������ͣ�
'''���2--·�������������1���,���1=1---Ӱ����𣻱��1=2---����豸����
'''���3--·��Ŀ�ĵ��豸�š�
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSetting As String
    Dim aSettings() As String
    Dim i As Integer
    strSetting = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ�·��")
    If strSetting <> "" Then
        aSettings = Split(strSetting, ",")
        On Error GoTo err
        ReDim aAutoRoutSetting(0 To (UBound(aSettings) + 1) / 3)
        For i = 0 To UBound(aSettings) Step 3
            aAutoRoutSetting(i / 3 + 1).type = aSettings(i)
            aAutoRoutSetting(i / 3 + 1).strCondition = aSettings(i + 1)
            aAutoRoutSetting(i / 3 + 1).strFTPDeviceNo = aSettings(i + 2)
        Next
    Else
        ReDim aAutoRoutSetting(0)
    End If
    Exit Sub
err:                    '������
    ReDim aAutoRoutSetting(0)
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '����ȷ����
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '���ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '�ַ�
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '����ȷ����
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function
Public Sub SaveImage(MainDeviceID As String, Optional iEncode As Integer = 0)
    '����  ���浥��ͼ��
    'MainDeviceID       �豸��
    'iEncode            ͼ��ѹ����ʽ
    Dim strSQL As String
    Dim rsFTP As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage
    Dim BufferDir As String
    
    Dim blnTmp As Boolean
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String
    Dim strUser As String, strPwd As String
    Dim lngResult As Long           'FTP�������
    
    With frmImgCapture.DicomViewer
        If .Images.count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.count)
    End With
                
    strSQL = "Select 'ftp://'||Decode(�û���,Null,'',�û���||Decode(����,Null,'',':'||����))" & _
        "||'@'||IP��ַ As Host,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL " & _
        "From Ӱ���豸Ŀ¼ " & _
        "Where �豸��=[1]"
    Set rsFTP = OpenSQLRecord(strSQL, App.ProductName, MainDeviceID)
     'û�д洢�豸ʱ�˳�
    If rsFTP.EOF = True Then
        MsgBox "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, App.ProductName
        Exit Sub
    End If
    strHost = rsFTP("Host"): strDirURL = rsFTP("URL")
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    
    strSQL = "select ���UID ,��������  from Ӱ�����¼ where ҽ��ID = [1] and ���ͺ� = [2]"
    Set rsTmp = OpenSQLRecord(strSQL, App.ProductName, mlngAdviceID, mlngSendNO)
    If IsNull(rsTmp("���UID")) Then
        gstrSQL = "ZL_Ӱ�����¼_SET(" & mlngAdviceID & "," & mlngSendNO & ",'" & _
            ImgTmp.StudyUID & "','" & GetImageAttribute(ImgTmp.Attributes, ATTR_����豸) & "'," & _
            "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
        ExecuteProcedure "PACSͼ�񱣴�"
        dtReceived = Format(Now, "yyyyMMdd")
    Else
        dtReceived = Format(rsTmp("��������"), "yyyyMMdd")
    End If
    
    strSQL = "Select ����UID From Ӱ��������  Where ����UID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "PACSͼ�񱣴�", CStr(ImgTmp.SeriesUID))
    '�����µļ������
    If rsTmp.EOF Then
        gstrSQL = "ZL_Ӱ������_INSERT('" & ImgTmp.StudyUID & "','" & ImgTmp.SeriesUID & "','" & _
            ImgTmp.SeriesDescription & "'," & _
            IIf(blnTmp, 1, 0) & ")"
        ExecuteProcedure "PACSͼ�񱣴�"
    End If
    
    '�����µ�ͼ��
    gstrSQL = "ZL_Ӱ��ͼ��_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "','" & _
        ImgTmp.SeriesDescription & "'," & _
        IIf(blnTmp, 1, 0) & ")"
    ExecuteProcedure "PACSͼ�񱣴�"
    gstrSQL = "ZL_Ӱ���鱨��_ADD('" & ImgTmp.StudyUID & "','" & ImgTmp.InstanceUID & ".jpg')"
    ExecuteProcedure "���汨��ͼ��"
    
    '����ͼ�񵽻���Ŀ¼
    On Error Resume Next
    BufferDir = App.Path & "\TmpImage\"
    MkLocalDir BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/"
    Select Case iEncode
        Case 1
            ImgTmp.WriteFile BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.5"
        Case 2
            ImgTmp.WriteFile BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID, True
        Case Else
            ImgTmp.WriteFile BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID, True, "1.2.840.10008.1.2.4.70"
    End Select
    ImgTmp.FileExport BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID & ".jpg", "JPG"
    
    GetFtpAddress strHost, strHost, strUser, strPwd
    lngResult = iNet.FuncFtpConnect(strHost, strUser, strPwd)
    '�ж�FTP�����Ƿ�ɹ�
    If lngResult = 0 Then
        MsgBox "FTP����ʧ�ܣ�ͼ���޷����档", vbInformation, App.ProductName
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    WriteToURL BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID, strHost, strDirURL & _
        dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID
    WriteToURL BufferDir & dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID & ".jpg", strHost, strDirURL & _
        dtReceived & "/" & ImgTmp.StudyUID & "/" & ImgTmp.InstanceUID & ".jpg"
     
    gcnOracle.CommitTrans
    iNet.FuncFtpDisConnect
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    err.Raise err.Number, "���ͼ�񱣴�"
End Sub
Public Sub DeleteImage(ImagesIndex As Long, MainDeviceID As String)
    'ɾ����ǰѡ��ͼ��
    'ImagesIndex      ͼ��Index
    'MainDeviceID     �豸��
    Dim blnTmp As Boolean
    Dim strDirURL As String
    Dim dtReceived As String
    Dim ImgTmp As New DicomImage
    Dim strSQL As String
    Dim rsFTP As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strReportImage As String
    Dim varTmp As Variant
    Dim i As Integer
    Dim strHost As String, strUser As String, strPwd As String
    
    With frmImgCapture.DicomViewer
        If .Images.count < ImagesIndex Then Exit Sub
        Set ImgTmp = .Images(ImagesIndex)
    End With
                
    On Error GoTo errHand
                
    strSQL = "Select 'ftp://'||Decode(�û���,Null,'',�û���||Decode(����,Null,'',':'||����))" & _
        "||'@'||IP��ַ As Host,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL " & _
        "From Ӱ���豸Ŀ¼ " & _
        "Where �豸��=[1]"
    Set rsFTP = OpenSQLRecord(strSQL, App.ProductName, MainDeviceID)
     'û�д洢�豸ʱ�˳�
    If rsFTP.EOF = True Then
        MsgBox "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, App.ProductName
        Exit Sub
    End If
    strHost = rsFTP("Host"): strDirURL = rsFTP("URL")
    
    
    gstrSQL = "Select a.ҽ��ID,a.���ͺ�,a.Ӱ�����,a.����,a.����,a.Ӣ����,a.�Ա�,a.����,a.��������,a.���,a.����," & _
        "a.������,a.���Ž�Ƭ,����豸,��������,c.ͼ��UID,d.ִ�м�,a.����ͼ�� " & _
        "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,����ҽ������ d " & _
        "Where a.���UID=b.���UID And b.����UID=c.����UID And a.ҽ��ID=d.ҽ��ID And a.���ͺ�=d.���ͺ� And a.���UID=[1] and c.ͼ��UID = [2]"
    Set rsTmp = OpenSQLRecord(gstrSQL, "PACSͼ�񱣴�", CStr(ImgTmp.StudyUID), CStr(ImgTmp.InstanceUID))
    
    If rsTmp.EOF = True Then
        MsgBox "û���ҵ�����ɾ����ͼ��!", vbQuestion, App.ProductName
        Exit Sub
    End If
    
    If IsNull(rsTmp("����ͼ��")) Then
        Exit Sub
    End If
    varTmp = Split(rsTmp("����ͼ��"), ";")

    For i = 0 To UBound(varTmp)
        If Trim(varTmp(i)) <> ImgTmp.InstanceUID & ".jpg" Then
            strReportImage = strReportImage & ";" & varTmp(i)
        End If
    Next
    strReportImage = Mid(strReportImage, 2)
    gstrSQL = "ZL_Ӱ��ͼ��_DELETE(" & rsTmp("ҽ��ID") & "," & rsTmp("���ͺ�") & ",'" & rsTmp("ͼ��UID") & "','" & strReportImage & "')"
    
    ExecuteProcedure "Ӱ��ͼ��ɾ��"
    GetFtpAddress strHost, strHost, strUser, strPwd
    iNet.FuncFtpConnect strHost, strUser, strPwd
    
    'ɾ��ͼ���ļ�
    Do While Not rsTmp.EOF
        RemoveFromURL strHost, strDirURL & _
            Format(Nvl(rsTmp("��������"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
            ImgTmp.StudyUID & "/" & rsTmp("ͼ��UID")
        RemoveFromURL strHost, strDirURL & _
            Format(Nvl(rsTmp("��������"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
            ImgTmp.StudyUID & "/" & rsTmp("ͼ��UID") & ".jpg"
        rsTmp.MoveNext
    Loop
    iNet.FuncFtpDisConnect
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

    
Public Sub GetAllImages(dcmViewer As DicomViewer, strStudyUID As String, strSeriesID As String, _
        strCachePath As String, iCurImageIndex As Long)
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double, lngRecID As Long
    Dim curImage As DicomImage, i As Integer, iFrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim bln1stDev As Boolean, objFile As New Scripting.FileSystemObject, strFileName As String, strTmpFile As String
    Dim iNet1 As New clsFtp
    Dim iNet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    
    bln1stDev = True
    
    On Error GoTo DBError
    gstrSQL = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1," & _
        "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2,e.�豸�� as �豸��2,C.���UID,B.����UID " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And C.ҽ��ID=[1] And C.���ͺ�=[2] Order By A.ͼ���"
    Set rsTmp = OpenSQLRecord(gstrSQL, "��ȡͼ��", mlngAdviceID, mlngSendNO)
    Screen.MousePointer = vbHourglass

    With dcmViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            .MultiColumns = 1: .MultiRows = 1

            ResizeRegion rsTmp.RecordCount, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
            
            If Len(strStudyUID) = 0 Then strStudyUID = Nvl(rsTmp("���UID"))
            strSeriesID = Nvl(rsTmp("����UID"))

            lngRecID = 1
            ClearCacheFolder strCachePath
            MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
            Do While Not rsTmp.EOF
                If Dir(strCachePath & Nvl(rsTmp("URL1"))) = vbNullString Then
                    strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
'                    iNet.strIPAddress = Nvl(rsTmp("Host1")): iNet.strUser = Nvl(rsTmp("User1")): iNet.strPsw = Nvl(rsTmp("Pwd1"))
                    If strDeviceNO1 <> rsTmp("�豸��1") Then
                        strDeviceNO1 = rsTmp("�豸��1")
                        iNet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
                    End If
                    If strDeviceNO2 <> rsTmp("�豸��2") Then
                        strDeviceNO2 = rsTmp("�豸��2")
                        iNet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
                    End If
                    If iNet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                        strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                        iNet.strIPAddress = Nvl(rsTmp("Host2")): iNet.strUser = Nvl(rsTmp("User2")): iNet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call iNet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
                    End If
                End If
                Set curImage = .Images.ReadFile(strCachePath & Nvl(rsTmp("URL1")))
                DoEvents
                
                With curImage
                    .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                End With

                lngRecID = lngRecID + 1

                rsTmp.MoveNext
            Loop

            iCurImageIndex = 1: .CurrentIndex = 1
            .Images(iCurImageIndex).BorderColour = vbRed
        Else
            .MultiColumns = 1: .MultiRows = 1: iCurImageIndex = 0
        End If
    End With
    
    iNet1.FuncFtpDisConnect
    iNet2.FuncFtpDisConnect
    Screen.MousePointer = vbDefault
    
    Exit Sub

ReadURLError:
    If bln1stDev Then
        bln1stDev = False
        Resume
    Else
        If ErrCenter() = 1 Then Resume
        Screen.MousePointer = vbDefault
        Call SaveErrLog
    End If
    Exit Sub

DBError:
    If ErrCenter() = 1 Then Resume
    Screen.MousePointer = vbDefault
    Call SaveErrLog
End Sub

