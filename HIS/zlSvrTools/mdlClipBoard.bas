Attribute VB_Name = "mdlClipBoard"
'///////////////////////////////////////////////////////////////////////////////
'
'       模块：剪贴板操作
'       功能：剪贴板操作,复制文件目录到剪贴板
'       编写：祝庆
'       日期：2011年1月3日
'
'///////////////////////////////////////////////////////////////////////////////
Option Explicit

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
        "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'剪贴版处理函数
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd _
        As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
        As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat _
        As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" _
        (ByVal wFormat As Long) As Long

Private Declare Function DragQueryFile Lib "shell32.dll" Alias _
        "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, _
        ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal _
        hDrop As Long, lpPoint As POINTAPI) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags _
        As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As _
        Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As _
        Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As _
        Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'判断数组是否为空
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
'剪贴版数据格式定义
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17

' New shell-oriented clipboard formats
'Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
'Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
'Private Const CFSTR_NETRESOURCES As String = "Net Resource"
'Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
'Private Const CFSTR_FILECONTENTS As String = "FileContents"
'Private Const CFSTR_FILENAME As String = "FileName"
'Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
'Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"

' 内存操作定义
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const FO_COPY = &H2

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type DROPFILES
   pFiles As Long
   pt As POINTAPI
   fNC As Long
   fWide As Long
End Type

Public Function clipClear() As Boolean
'清空当前剪贴板
    Call EmptyClipboard
End Function

Public Function clipCopyFiles(File() As String) As Boolean
'复制多个文件到剪贴板
   On Error Resume Next
   Dim strData As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   strData = ""

   
   '清除剪贴版中现存的数据
   If OpenClipboard(0&) Then
        '清空当前剪贴板
        Call EmptyClipboard
        
        '判断文件数组是否为空
        If SafeArrayGetDim(File) = 0 Then Exit Function
        For i = LBound(File) To UBound(File)
            strData = strData & File(i) & vbNullChar
        Next
        
        hGlobal = GlobalAlloc(GHND, Len(df) + LenB(strData))
        
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
         
            df.pFiles = Len(df)
            Call CopyMem(ByVal lpGlobal, df, Len(df))
            Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal strData, LenB(strData))
   
            Call GlobalUnlock(hGlobal)
         
            If SetClipboardData(CF_HDROP, hGlobal) Then
                clipCopyFiles = True
            End If

        End If
        
        Call CloseClipboard
    End If
End Function


