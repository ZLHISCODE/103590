Attribute VB_Name = "mdlApi"
Option Explicit
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
 
'常数声明
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息

Public Const GWL_EXSTYLE = (-20)

Public Const CB_GETCURSEL = &H147
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETCOUNT = &H146

Public Const SB_TOP = 6
Public Const WM_VSCROLL = &H115
Public Const BDR_SUNKENINNER = &H8
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const HH_DISPLAY_TOPIC = &H0
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCaL_MaCHINE = &H80000002
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const KLF_REORDER = &H8
Public Const LVM_SETCOLUMNWIDTH = &H101E
Public Const LVSCW_AUTOSIZE = -1
Public Const MAX_PATH = 256
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
'OpenDir函数的回调函数使用
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const WM_USER = &H400
Public Const BFFM_SETSELECTION = (WM_USER + 102)
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Const SPI_GETWORKAREA = 48

Public Const GWL_HWNDPARENT = (-8)
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO = &H24
Public Const WH_KEYBOARD = 2

 

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
    value As Long
End Type

Private Type EncoderParameters
    count As Long
    Parameter As EncoderParameter
End Type
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source



'API定义
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'系统方案设置----------------------------------
Public Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Public Declare Function SetCursorPos& Lib "user32" (ByVal X&, ByVal Y&)
Public Declare Function GetSystemMenu& Lib "user32" (ByVal hWnd&, ByVal bRevert&)
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
'打开文件夹，并设置初始路径
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const WM_NCACTIVATE = &H86
Public Const WM_CLOSE = &H10


Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As GUID) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Public Sub SaveScreen(ByVal strInfo As String, objPic As PictureBox)
    '截屏
    Dim strFileName As String, strPath As String, strTime As String
    Dim i As Integer, varPath As Variant
    
    On Error GoTo errHandle
    
    Clipboard.Clear
    keybd_event vbKeySnapshot, 0&, 0&, 0&
    DoEvents

    strPath = GetSetting("ZLSOFT", "公共全局", "程序路径", App.Path & "\")
    If InStr(strPath, "\") > 0 Then
        varPath = Split(strPath, "\")
        If UBound(varPath) >= 2 Then
            strPath = ""
            For i = 0 To 1
                strPath = strPath & varPath(i) & "\"
            Next
            
        End If
    Else
        strPath = App.Path & "\"
    End If
    strTime = Format(Now, "yyMMddHHmmss")
    strFileName = strPath & gstrDBUser & strTime & ".JPG"
    
    objPic.Picture = LoadPicture()
    objPic.Picture = Clipboard.GetData(vbCFBitmap)
    If Dir(strFileName) = "" Then
        SavePic objPic.Picture, strFileName, "jpg"
    Else
        If MsgBox("文件已存在，是否覆盖？", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            Call Kill(strFileName)
            SavePic objPic.Picture, strFileName, "jpg"
        End If
    End If
 
    If strInfo <> "" Then
        Open strPath & gstrDBUser & strTime & ".Txt" For Append As #1
        Write #1, strInfo
        Close #1
    End If
    MsgBox "截图 " & strFileName & " 已保存!", vbInformation, gstrSysName
    
    Exit Sub
errHandle:
    MsgBox "保存截图出现错误！" & vbNewLine & Err.Description, vbExclamation, gstrSysName
End Sub
Private Sub SavePic(ByVal pict As StdPicture, ByVal FileName As String, PicType As String, _
                    Optional ByVal Quality As Byte = 80, _
                    Optional ByVal TIFF_ColorDepth As Long = 24, _
                    Optional ByVal TIFF_Compression As Long = 6)
   Screen.MousePointer = vbHourglass
   Dim tSI As GdiplusStartupInput
   Dim lRes As Long
   Dim lGDIP As Long
   Dim lBitmap As Long
   Dim aEncParams() As Byte
   On Error GoTo errHandle:
   tSI.GdiplusVersion = 1   ' 初始化 GDI+
   lRes = GdiplusStartup(lGDIP, tSI)
   If lRes = 0 Then     ' 从句柄创建 GDI+ 图像
      lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
      If lRes = 0 Then
         Dim tJpgEncoder As GUID
         Dim tParams As EncoderParameters    '初始化解码器的GUID标识
         Select Case UCase(PicType)
         Case ".JPG", "JPG", ".JPEG", "JPEG"
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            tParams.count = 1                               ' 设置解码器参数
            With tParams.Parameter ' Quality
               CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID    ' 得到Quality参数的GUID标识
               .NumberOfValues = 1
               .type = 4
               .value = VarPtr(Quality)
            End With
            ReDim aEncParams(1 To Len(tParams))
            Call CopyMemory(aEncParams(1), tParams, Len(tParams))
        Case ".PNG", "PNG"
             CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
             ReDim aEncParams(1 To Len(tParams))
        Case ".GIF", "GIF"
             CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
             ReDim aEncParams(1 To Len(tParams))
        Case ".TIFF", "TIFF"
             CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
             tParams.count = 2
             ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
             With tParams.Parameter
                .NumberOfValues = 1
                .type = 4
                 CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID    ' 得到ColorDepth参数的GUID标识
                .value = VarPtr(TIFF_Compression)
            End With
            Call CopyMemory(aEncParams(1), tParams, Len(tParams))
            With tParams.Parameter
                .NumberOfValues = 1
                .type = 4
                 CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID    ' 得到Compression参数的GUID标识
                .value = VarPtr(TIFF_ColorDepth)
            End With
            Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
        Case ".BMP", "BMP"                                              '可以提前写保存为BMP的代码，因为并没有用GDI+
            SavePicture pict, FileName
            Screen.MousePointer = vbDefault
            Exit Sub
        End Select
         lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, aEncParams(1))             '保存图像
         GdipDisposeImage lBitmap       ' 销毁GDI+图像
      End If
      GdiplusShutdown lGDIP              '销毁 GDI+
   End If
   Screen.MousePointer = vbDefault
   Erase aEncParams
   Exit Sub
errHandle:
    Screen.MousePointer = vbDefault
    MsgBox "在保存图片的过程中发生错误:" & vbCrLf & vbCrLf & "错误号:  " & Err.Number & vbCrLf & "错误描述:  " & Err.Description, vbInformation Or vbOKOnly, "错误"
End Sub



