VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetFullPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const GCL_CONVERSION = 1
Const GCL_REVERSECONVERSION = 2

Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Private Const IME_ESC_MAX_KEY = &H1005
Private Const IME_ESC_IME_NAME = &H1006
Private Type CANDIDATELIST
  dwSize As Long
  dwStyle As Long
  dwCount As Long
  dwSelection As Long
  dwPageStart As Long
  dwPageSize As Long
  dwOffset(0) As Long
End Type
Private Declare Function ImmGetContext Lib "imm32" ( _
    ByVal hwnd As Long _
) As Long

Private Declare Function ImmReleaseContext Lib "imm32" ( _
    ByVal hwnd As Long, _
    ByVal hIMC As Long _
) As Long

Private Declare Function ImmGetConversionList Lib "imm32" Alias "ImmGetConversionListW" ( _
    ByVal hKL As Long, _
    ByVal hIMC As Long, _
    ByRef lpSrc As Byte, _
    ByRef lpDst As Any, _
    ByVal dwBufLen As Long, _
    ByVal uFlag As Long _
) As Long

Private Declare Function GetKeyboardLayout Lib "user32" ( _
    ByVal idThread As Long _
) As Long
Private Declare Function GetKeyboardLayoutList Lib "user32" _
   (ByVal nBuff As Long, _
    ByRef lpList As Long) As Long
    
Private Declare Function ImmEscape Lib "imm32.dll" _
    Alias "ImmEscapeA" _
   (ByVal hKL As Long, _
    ByVal hIMC As Long, _
    ByVal un As Long, _
    ByRef lpv As Any) As Long
    
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" ( _
    ByRef strString As Any _
) As Long

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(127) As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    ByRef VersionInfo As OSVERSIONINFO _
) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long _
)


Public Function ReverseConversionNew(hwnd As Long, strSource As String) As String

    Dim bySource() As Byte
    
    Dim i As Integer
    Dim arrKeyLayout() As Long
    Dim strIME As String
    
    Dim hIMC As Long
    Dim hKL As Long
    Dim lngSize As Long
    Dim lngOffset As Long
    Dim iKeyLayoutCount As Integer

    Dim byCandiateArray() As Byte
    Dim CandiateList As CANDIDATELIST

    Dim byWork() As Byte
    Dim lngResult As Long
    
    Const BUFFERSIZE As Integer = 255
    Dim osvi As OSVERSIONINFO

    Dim isChineseIme As Boolean
    
    If strSource = "" Then Exit Function

    'OSÅÐ„e
    osvi.dwOSVersionInfoSize = Len(osvi)
    lngResult = GetVersionEx(osvi)

    If osvi.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        'WindowsNTÏµ:Unicode×Ö·û¼¯
        bySource = strSource

        ReDim Preserve bySource(UBound(bySource) + 2)
    Else
        'Windows95Ïµ
        bySource = StrConv(strSource, vbFromUnicode)

        ReDim Preserve bySource(UBound(bySource) + 1)
    End If
    
        hIMC = ImmGetContext(hwnd)

    ReDim arrKeyLayout(BUFFERSIZE) As Long
    strIME = Space(BUFFERSIZE)
    iKeyLayoutCount = GetKeyboardLayoutList(BUFFERSIZE, arrKeyLayout(0))

    isChineseIme = False
    For i = 0 To iKeyLayoutCount - 1
        If ImmEscape(arrKeyLayout(i), hIMC, IME_ESC_IME_NAME, ByVal strIME) Then
            If Trim(UCase("Î¢ÈíÆ´ÒôÊäÈë·¨")) = UCase(Replace(Trim(strIME), Chr(0), "")) Then
                isChineseIme = True
                Exit For
            End If
        End If
    Next i
 
    If isChineseIme = False Then Exit Function
    hKL = arrKeyLayout(i)
'    hKL = GetKeyboardLayout(0)

    lngSize = ImmGetConversionList(hKL, hIMC, bySource(0), Null, 0, GCL_REVERSECONVERSION)

    If lngSize > 0 Then

        ReDim byCandiateArray(lngSize)

        lngSize = ImmGetConversionList(hKL, hIMC, bySource(0), byCandiateArray(0), lngSize, _
                                       GCL_REVERSECONVERSION)

        MoveMemory CandiateList, byCandiateArray(0), Len(CandiateList)

        If CandiateList.dwCount > 0 Then

            lngOffset = CandiateList.dwOffset(0)

            ReverseConversionNew = MidB(byCandiateArray, lngOffset + 1, _
                                     lstrlen(byCandiateArray(lngOffset)) * 2)

        End If

    End If

    lngResult = ImmReleaseContext(hwnd, hIMC)

End Function

Private Function funDelNum(strSource As String) As String
    Dim iTag As Long
    Dim strOne As String
    For iTag = 1 To Len(strSource)
        strOne = Mid(strSource, iTag, 1)
        If Not IsNumeric(strOne) Then
            funDelNum = funDelNum & strOne
        End If
    Next iTag
End Function

