Attribute VB_Name = "UserDefine"
Option Explicit
'----------部件基础参数
Public gstrSysName As String
Public gcnOracle As ADODB.Connection
Public gstrSQL As String

Public Const ATTR_检查日期 As String = "Study Date"
Public Const ATTR_检查时间 As String = "Study Time"
Public Const ATTR_序列日期 As String = "Series Date"
Public Const ATTR_序列时间 As String = "Series Time"
Public Const ATTR_影像类别 As String = "Modality"
Public Const ATTR_设备商 As String = "Manufacturer"
Public Const ATTR_检查设备 As String = "Manufacturer's Model Name"
Public mlngAdviceID As Long, mlngSendNO As Long                         '医嘱ID、发送号
Dim iNet As New clsFtp

Public Type AutoRoutSetting                        '存储自动路由设置的类型
    type As Long                                    '1--按照影像类别进行自动路由；2--按照检查设备进行自动路由
    strCondition As String                          '条件内容：影像类别或检查设备名称
    strFTPDeviceNo As String                        '自动路由的“目的设备号”
End Type
Public aAutoRoutSetting() As AutoRoutSetting       '存储自动路由设置的数组
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
Public hBoard As Long               '采集卡句柄
Public bActive As Byte              '实时采集标志
Public total As Integer             '全部采集卡数量
Public iCurrUsedNo As Long          '当前采集卡

Public hAudio As Long
Public elapsedtimes As Long

Public SQFILE As String             '采集文件默认名称
Public dwMaxMemSize As Long         '最大内存大小
Public dwBufSize As Long            '最大缓存大小

Public lpbi As BITMAPINFO           '采集目标体格式文件头
Public lpdib As Long                '采集目标体数据
Public lpMemory As Long             '用户内存指针

Public buf(5000000) As Byte
Public blk As tpBlockInfo           '用户结构指针结构变量

Public iNumImage As Integer         '采集帧数
Public iNum As Integer              '多卡或多通道采集数量
Public NoCapture As Integer
Public lphBoard(5) As Long
Public iVirtCode As Integer

Public bMakeMirror As Integer       '软件x,y方向镜象
Public iClientWidth As Integer      '客户区宽度
Public iClientHeight As Integer     '客户区高度

Public bMaskMode As Integer         '位屏蔽状态


Public sampwidth As Integer         '采集窗口宽度
Public sampheight As Integer        '采集窗口高度

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
    '返回Param1与Param2中较小的数字
    If Param1 > Param2 Then
        exfunMin = Param2
    Else
        exfunMin = Param1
    End If
End Function


Function exfunMax(ByVal Param1 As Long, ByVal Param2 As Long) As Long
    '返回Param1与Param2中较大的数字
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
    '取得Value的高字
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
    '取得Value的低位
    Dim k As Long
    
    k = Value
    Do While (k > 255)
        k = k - 256
    Loop
    exLoByte = k
End Function

Function exLoWord(ByVal Value As Long) As Integer
    '取得Value的低字
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
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Sub ExecuteProcedure(ByVal strFormCaption As String)
'功能：执行过程式的SQL语句
    Call SQLTest(App.ProductName, strFormCaption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub SaveImages(Images As DicomImages, ByVal MainDeviceID As String, ByVal BufferDir As String, Optional iEncode As Integer = 0, Optional ByVal strImgType As String = "")
'功能：保存图像
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer  '保存的图像数
    Dim intSQL As Integer, rsTmp As New ADODB.Recordset
    
    Dim blnAddTmp As Boolean, blnTmp As Boolean
    Dim strAge As String, strBirth As String
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String, dtCurrent As String
    Dim strUser As String, strPwd As String
    Dim ImageType As String, CheckNo As Long, CheckDev As String
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    Dim CheckUID As String, SeriesUID As String
    Dim aPatientID() As String, lngAdviceID As Long, lngSendNO As Long '图像中的医嘱ID：医嘱ID_发送号
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    
    dtCurrent = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    gstrSQL = "Select 'ftp://'||Decode(用户名,Null,'',用户名||Decode(密码,Null,'',':'||密码))" & _
        "||'@'||IP地址 As Host,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL " & _
        "From 影像设备目录 " & _
        "Where 设备号=[1]"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", MainDeviceID)
    If rsTmp.EOF Then
        err.Raise vbObjectError + 1, "PACS图像保存", "设备号设置错误！"
    End If
    strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
    GetFtpAddress strHost, strHost, strUser, strPwd
    iNet.FuncFtpConnect strHost, strUser, strPwd
    iCount = 0
    If Images.count > 0 Then
        gstrSQL = "Select a.医嘱ID,a.发送号,a.影像类别,a.检查号,a.姓名,a.英文名,a.性别,a.年龄,a.出生日期,a.身高,a.体重,a.病理检查,a.发放胶片,检查设备,接收日期,c.图像UID,d.执行间 " & _
            "From 影像检查记录 a,影像检查序列 b,影像检查图象 c,病人医嘱发送 d " & _
            "Where a.检查UID=b.检查UID And b.序列UID=c.序列UID And a.医嘱ID=d.医嘱ID And a.发送号=d.发送号 And a.检查UID=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(Images(1).StudyUID))
        If Not rsTmp.EOF Then
            lngAdviceID = Nvl(rsTmp("医嘱ID"), 0)
            lngSendNO = Nvl(rsTmp("发送号"), 0)
        End If
        '删除图像文件
        Do While Not rsTmp.EOF
            RemoveFromURL strHost, strDirURL & _
                Format(Nvl(rsTmp("接收日期"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
                Images(1).StudyUID & "/" & rsTmp("图像UID")
            rsTmp.MoveNext
        Loop
        '重新开始检查
        If lngAdviceID > 0 Then
            rsTmp.MoveFirst
            
            gstrSQL = "ZL_影像检查_CANCEL(" & Nvl(rsTmp("医嘱ID"), 0) & "," & Nvl(rsTmp("发送号"), 0) & ")"
            ExecuteProcedure "PACS图像保存"
            gstrSQL = "ZL_影像检查_BEGIN('" & Nvl(rsTmp("执行间")) & "'," & Nvl(rsTmp("检查号"), 0) & "," & rsTmp("医嘱ID") & "," & rsTmp("发送号") & ",'" & Nvl(rsTmp("影像类别")) & "','" & _
                Nvl(rsTmp("姓名")) & "','" & Nvl(rsTmp("英文名")) & "','" & Nvl(rsTmp("性别")) & "','" & _
                Nvl(rsTmp("年龄")) & "'," & IIf(IsNull(rsTmp("出生日期")), "Null", "to_Date('" & Format(rsTmp("出生日期"), "yyyy-MM-dd") & "','YYYY-MM-DD')") & ",'" & Nvl(rsTmp("身高")) & "','" & Nvl(rsTmp("体重")) & "'," & _
                Nvl(rsTmp("病理检查"), 0) & "," & Nvl(rsTmp("发放胶片"), 0) & ",'" & Nvl(rsTmp("检查设备")) & "')"
            ExecuteProcedure "PACS图像保存"
        End If
    End If
    
    For Each curImage In Images
        gstrSQL = "Select 图像UID From 影像检查图象 Where 图像UID=[1]" & _
            " Union All Select 图像UID From 影像临时图象 Where 图像UID=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(curImage.InstanceUID))
        '新图像
        If rsTmp.EOF Then
            gstrSQL = "Select 检查UID From 影像检查记录 Where 检查UID=[1]" & _
                " Union All Select 检查UID From 影像临时记录 Where 检查UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(curImage.StudyUID))
            '按病人ID或英文名查找
            If rsTmp.EOF Then
                blnAddTmp = True
                aPatientID = Split(curImage.PatientID, "_")
                If UBound(aPatientID) >= 0 And lngAdviceID = 0 Then
                    lngAdviceID = Val(aPatientID(0)) ': lngSendNO = Val(aPatientID(1))
                End If
                gstrSQL = "Select Distinct A.医嘱ID,A.发送号 From 影像检查记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
                    " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID" & _
                    " And A.医嘱ID=[1]" & _
                    " And B.执行状态=3 And B.执行过程=2"
                If rsTmp.State <> adStateClosed Then rsTmp.Close
                Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", lngAdviceID)
                '与HIS填写的检查记录对应
                If Not rsTmp.EOF Then
                    '填入检查UID
                    gstrSQL = "ZL_影像检查记录_SET(" & rsTmp(0) & "," & rsTmp(1) & ",'" & _
                        curImage.StudyUID & "','" & GetImageAttribute(curImage.Attributes, ATTR_检查设备) & "'," & _
                        "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACS图像保存"
                    blnAddTmp = False
                End If
                '插入临时检查记录
                If blnAddTmp Then
                    If IsDate(curImage.DateOfBirthAsDate) Then
                        strAge = CStr(Year(Date) - Year(curImage.DateOfBirthAsDate))
                        strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                    Else
                        strAge = "": strBirth = ""
                    End If
                    gstrSQL = "ZL_影像临时检查_INSERT('" & strImgType & "',Null,'" & _
                        curImage.Name & "','" & curImage.Name & "','" & _
                        curImage.Sex & "','" & strAge & "'," & _
                        IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                        GetImageAttribute(curImage.Attributes, ATTR_检查设备) & "','" & curImage.StudyUID & "'," & _
                        "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
                    ExecuteProcedure "PACS图像保存"
                End If
            End If
            
            gstrSQL = "Select 0 As 临时,接收日期,影像类别,Nvl(检查号,0) As 检查号," & _
                "检查设备,姓名,英文名,性别,Nvl(年龄,'-1') As 年龄,检查UID From 影像检查记录 Where 检查UID=[1]" & _
                " Union All Select 1 As 临时,接收日期,影像类别,Nvl(检查号,0) As 检查号," & _
                "检查设备,姓名,英文名,性别,Nvl(年龄,'-1') As 年龄,检查UID From 影像临时记录 Where 检查UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(curImage.StudyUID))
            blnTmp = IIf(rsTmp(0) = 1, True, False) '序列和图像是否放入临时记录中
            dtReceived = Format(rsTmp(1), "yyyyMMdd")
            
            ImageType = Nvl(rsTmp(2)): CheckNo = rsTmp(3): CheckDev = Nvl(rsTmp(4))
            PatientName = Nvl(rsTmp(5)): EnglishName = Nvl(rsTmp(6)): Sex = Nvl(rsTmp(7)): Age = Val(rsTmp(8))
            CheckUID = Nvl(rsTmp(9))
            
            gstrSQL = "Select 序列UID From " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                " Where 序列UID=[1]"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(curImage.SeriesUID))
            '插入新的检查序列
            If rsTmp.EOF Then
                gstrSQL = "ZL_影像序列_INSERT('" & curImage.StudyUID & "','" & curImage.SeriesUID & "','" & _
                    curImage.SeriesDescription & "'," & _
                    IIf(blnTmp, 1, 0) & ")"
                ExecuteProcedure "PACS图像保存"
            End If
            
            '插入新的图像
            gstrSQL = "ZL_影像图象_INSERT('" & curImage.InstanceUID & "','" & curImage.SeriesUID & "','" & _
                curImage.SeriesDescription & "'," & _
                IIf(blnTmp, 1, 0) & ")"
            ExecuteProcedure "PACS图像保存"
            gstrSQL = "ZL_影像检查报告_ADD('" & curImage.StudyUID & "','" & curImage.InstanceUID & ".jpg')"
            ExecuteProcedure "保存报告图像"
            
            '保存图像到缓存目录
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
    err.Raise err.Number, "检查图像保存"
End Sub

Public Sub WriteToURL(ByVal SrcFileName As String, ByVal DestAddress As String, ByVal DestFileName As String)
'功能：将本地文件保存到远程网络上
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
'功能：将本地文件保存到远程网络上
'    Dim iNet As New clsFtp, strHost As String, strUser As String, strPwd As String
    Dim objFileSystem As New Scripting.FileSystemObject
    
'    GetFtpAddress DestAddress, strHost, strUser, strPwd
'    iNet.strIPAddress = strHost: iNet.strUser = strUser: iNet.strPsw = strPwd
    
    iNet.FuncDelFile objFileSystem.GetParentFolderName(DestFileName), objFileSystem.GetFileName(DestFileName)
End Sub
'--------------------------------------
'--将Ftp地址分解为主机、用户、密码
'--------------------------------------
Private Sub GetFtpAddress(ByVal strFtpPath As String, strHost As String, strUser As String, strPwd As String)
    Dim iPos As Integer
    Dim aUser() As String
    On Error Resume Next
        
    iPos = InStr(strFtpPath, "@")
    If iPos = 0 Then '无登陆用户信息
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
'功能：计算DicomViewer的行列数
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
    Dim i As Integer            '用于循环的变量
    Dim strImageType As String
    Dim strImageDevice As String
    Dim iRouting As Integer         '标识做自动路由的规则号，0为不满足路由条件
    Dim rsTmp As New ADODB.Recordset
'    Dim strHost As String           'FTP主机的用户名密码+ IP地址串
    Dim strDirURL As String         'FTP主机的目录
    Dim strHost As String, strUser As String, strPwd As String
    Dim DestAddress As String
    
    Call subReadAutoRoutSetting
    
    iRouting = 0
    '获取图像的影像类别和检查设备名
    strImageType = GetImageAttribute(img.Attributes, ATTR_影像类别)
    strImageDevice = GetImageAttribute(img.Attributes, ATTR_检查设备)
    '对比存储规则，不匹配则退出
    For i = 1 To UBound(aAutoRoutSetting)
        If aAutoRoutSetting(i).strCondition = IIf(aAutoRoutSetting(i).type = 1, strImageType, strImageDevice) Then
                iRouting = i
                Exit For
        End If
    Next
    '存储图像到指定FTP设备
    If iRouting <> 0 Then
        '获取目的设备的URL
        gstrSQL = "Select 'ftp://'||Decode(用户名,Null,'',用户名||Decode(密码,Null,'',':'||密码))" & _
        "||'@'||IP地址 As Host,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL " & _
        "From 影像设备目录 " & _
        "Where 设备号=[1]"
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", aAutoRoutSetting(iRouting).strFTPDeviceNo)
        If rsTmp.EOF Then
            err.Raise vbObjectError + 1, "PACS图像保存", "设备号设置错误！"
        End If
        strHost = rsTmp("Host"): strDirURL = rsTmp("URL")
        DestAddress = strHost & strDirURL
        '保存图像到指定URL
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

'读取自动路由的规则
Public Sub subReadAutoRoutSetting()
''''''''''''''''''''''''''''''''''''''''''''''''''''
'''从注册表的"ZLSOFT\公共模块\产品名\接收服务\自动路由"中读取自动路由的规则设置
'''规则设置为连续的标记，使用英文逗号分开。
'''标记1--设置类型；
'''标记2--路由条件（跟标记1相关,标记1=1---影像类别；标记1=2---检查设备）；
'''标记3--路由目的地设备号。
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSetting As String
    Dim aSettings() As String
    Dim i As Integer
    strSetting = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动路由")
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
err:                    '错误处理
    ReDim aAutoRoutSetting(0)
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
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
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '不明确类型
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '字符
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '不明确类型
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function
Public Sub SaveImage(MainDeviceID As String, Optional iEncode As Integer = 0)
    '功能  保存单个图像
    'MainDeviceID       设备号
    'iEncode            图像压缩方式
    Dim strSQL As String
    Dim rsFTP As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim ImgTmp As New DicomImage
    Dim BufferDir As String
    
    Dim blnTmp As Boolean
    Dim strDirURL As String, strHost As String
    Dim dtReceived As String
    Dim strUser As String, strPwd As String
    Dim lngResult As Long           'FTP操作结果
    
    With frmImgCapture.DicomViewer
        If .Images.count <= 0 Then Exit Sub
        Set ImgTmp = .Images(.Images.count)
    End With
                
    strSQL = "Select 'ftp://'||Decode(用户名,Null,'',用户名||Decode(密码,Null,'',':'||密码))" & _
        "||'@'||IP地址 As Host,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL " & _
        "From 影像设备目录 " & _
        "Where 设备号=[1]"
    Set rsFTP = OpenSQLRecord(strSQL, App.ProductName, MainDeviceID)
     '没有存储设备时退出
    If rsFTP.EOF = True Then
        MsgBox "没有找到存储设备,请重新选择存储设备!", vbInformation, App.ProductName
        Exit Sub
    End If
    strHost = rsFTP("Host"): strDirURL = rsFTP("URL")
    
    On Error GoTo DBError
    gcnOracle.BeginTrans
    
    strSQL = "select 检查UID ,接收日期  from 影像检查记录 where 医嘱ID = [1] and 发送号 = [2]"
    Set rsTmp = OpenSQLRecord(strSQL, App.ProductName, mlngAdviceID, mlngSendNO)
    If IsNull(rsTmp("检查UID")) Then
        gstrSQL = "ZL_影像检查记录_SET(" & mlngAdviceID & "," & mlngSendNO & ",'" & _
            ImgTmp.StudyUID & "','" & GetImageAttribute(ImgTmp.Attributes, ATTR_检查设备) & "'," & _
            "to_Date('" & Format(Now, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & MainDeviceID & "')"
        ExecuteProcedure "PACS图像保存"
        dtReceived = Format(Now, "yyyyMMdd")
    Else
        dtReceived = Format(rsTmp("接收日期"), "yyyyMMdd")
    End If
    
    strSQL = "Select 序列UID From 影像检查序列  Where 序列UID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "PACS图像保存", CStr(ImgTmp.SeriesUID))
    '插入新的检查序列
    If rsTmp.EOF Then
        gstrSQL = "ZL_影像序列_INSERT('" & ImgTmp.StudyUID & "','" & ImgTmp.SeriesUID & "','" & _
            ImgTmp.SeriesDescription & "'," & _
            IIf(blnTmp, 1, 0) & ")"
        ExecuteProcedure "PACS图像保存"
    End If
    
    '插入新的图像
    gstrSQL = "ZL_影像图象_INSERT('" & ImgTmp.InstanceUID & "','" & ImgTmp.SeriesUID & "','" & _
        ImgTmp.SeriesDescription & "'," & _
        IIf(blnTmp, 1, 0) & ")"
    ExecuteProcedure "PACS图像保存"
    gstrSQL = "ZL_影像检查报告_ADD('" & ImgTmp.StudyUID & "','" & ImgTmp.InstanceUID & ".jpg')"
    ExecuteProcedure "保存报告图像"
    
    '保存图像到缓存目录
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
    '判断FTP连接是否成功
    If lngResult = 0 Then
        MsgBox "FTP连接失败，图像无法保存。", vbInformation, App.ProductName
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
    err.Raise err.Number, "检查图像保存"
End Sub
Public Sub DeleteImage(ImagesIndex As Long, MainDeviceID As String)
    '删除当前选中图像
    'ImagesIndex      图像Index
    'MainDeviceID     设备号
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
                
    strSQL = "Select 'ftp://'||Decode(用户名,Null,'',用户名||Decode(密码,Null,'',':'||密码))" & _
        "||'@'||IP地址 As Host,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL " & _
        "From 影像设备目录 " & _
        "Where 设备号=[1]"
    Set rsFTP = OpenSQLRecord(strSQL, App.ProductName, MainDeviceID)
     '没有存储设备时退出
    If rsFTP.EOF = True Then
        MsgBox "没有找到存储设备,请重新选择存储设备!", vbInformation, App.ProductName
        Exit Sub
    End If
    strHost = rsFTP("Host"): strDirURL = rsFTP("URL")
    
    
    gstrSQL = "Select a.医嘱ID,a.发送号,a.影像类别,a.检查号,a.姓名,a.英文名,a.性别,a.年龄,a.出生日期,a.身高,a.体重," & _
        "a.病理检查,a.发放胶片,检查设备,接收日期,c.图像UID,d.执行间,a.报告图象 " & _
        "From 影像检查记录 a,影像检查序列 b,影像检查图象 c,病人医嘱发送 d " & _
        "Where a.检查UID=b.检查UID And b.序列UID=c.序列UID And a.医嘱ID=d.医嘱ID And a.发送号=d.发送号 And a.检查UID=[1] and c.图像UID = [2]"
    Set rsTmp = OpenSQLRecord(gstrSQL, "PACS图像保存", CStr(ImgTmp.StudyUID), CStr(ImgTmp.InstanceUID))
    
    If rsTmp.EOF = True Then
        MsgBox "没有找到可以删除的图像!", vbQuestion, App.ProductName
        Exit Sub
    End If
    
    If IsNull(rsTmp("报告图象")) Then
        Exit Sub
    End If
    varTmp = Split(rsTmp("报告图象"), ";")

    For i = 0 To UBound(varTmp)
        If Trim(varTmp(i)) <> ImgTmp.InstanceUID & ".jpg" Then
            strReportImage = strReportImage & ";" & varTmp(i)
        End If
    Next
    strReportImage = Mid(strReportImage, 2)
    gstrSQL = "ZL_影像图象_DELETE(" & rsTmp("医嘱ID") & "," & rsTmp("发送号") & ",'" & rsTmp("图像UID") & "','" & strReportImage & "')"
    
    ExecuteProcedure "影像图像删除"
    GetFtpAddress strHost, strHost, strUser, strPwd
    iNet.FuncFtpConnect strHost, strUser, strPwd
    
    '删除图像文件
    Do While Not rsTmp.EOF
        RemoveFromURL strHost, strDirURL & _
            Format(Nvl(rsTmp("接收日期"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
            ImgTmp.StudyUID & "/" & rsTmp("图像UID")
        RemoveFromURL strHost, strDirURL & _
            Format(Nvl(rsTmp("接收日期"), CStr(zlDatabase.Currentdate)), "yyyyMMdd") & "/" & _
            ImgTmp.StudyUID & "/" & rsTmp("图像UID") & ".jpg"
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
    gstrSQL = "Select A.图像号,D.用户名 As User1,D.密码 As Pwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL1,d.设备号 as 设备号1, " & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL2,e.设备号 as 设备号2,C.检查UID,B.序列UID " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And C.医嘱ID=[1] And C.发送号=[2] Order By A.图像号"
    Set rsTmp = OpenSQLRecord(gstrSQL, "读取图像", mlngAdviceID, mlngSendNO)
    Screen.MousePointer = vbHourglass

    With dcmViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            .MultiColumns = 1: .MultiRows = 1

            ResizeRegion rsTmp.RecordCount, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
            
            If Len(strStudyUID) = 0 Then strStudyUID = Nvl(rsTmp("检查UID"))
            strSeriesID = Nvl(rsTmp("序列UID"))

            lngRecID = 1
            ClearCacheFolder strCachePath
            MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
            Do While Not rsTmp.EOF
                If Dir(strCachePath & Nvl(rsTmp("URL1"))) = vbNullString Then
                    strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
'                    iNet.strIPAddress = Nvl(rsTmp("Host1")): iNet.strUser = Nvl(rsTmp("User1")): iNet.strPsw = Nvl(rsTmp("Pwd1"))
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        iNet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
                    End If
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
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

