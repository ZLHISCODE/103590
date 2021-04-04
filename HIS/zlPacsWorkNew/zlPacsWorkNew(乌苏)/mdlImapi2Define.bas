Attribute VB_Name = "mdlImapi2Define"
Option Explicit



  
  
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" ( _
    ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Public Type typArrayVariant
  uVarType As Integer
  unUsed1 As Integer
  unUsed2 As Long
  Pointer As Long
  unUsed3 As Long
End Type

Public Type typSafeArrayBound
  cElements As Long
  LowerBound As Long
End Type

Public Type typSafeArray
  nDimension As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  Pointer As Long
End Type


Public Enum enuSafeArrayMessage
  S_OK = &H0
End Enum


'完整性检测级别
Public Enum TIntergrityVerificationLevel
    ivlNone = 0
    ivlQuick = 1
    ivlFull = 2
End Enum


'刻录状态
Public Type TBurnArgs
    ElapsedTime As Long
    FreeSystemBuffer As Long
    LastReadLba As Long
    LastWrittenLba As Long
    RemainingTime As Long
    SectorCount As Long
    StartLba As Long
    TotalSystemBuffer As Long
    TotalTime As Long
    UsedSystemBuffer As Long
End Type

Public Const Lenth_Variant = 16
Public Const Offset_Variant = 8



Private Declare Function SafeArrayGetDim_array Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32" (ByVal lpSafeArray As Long) As Long
Private Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Function ObjPtr Lib "MSVBVM60.dll" Alias "VarPtr" (var As Object) As Long
Private Declare Function VarPtr Lib "MSVBVM60.dll" (var As Any) As Long
Private Declare Sub CopyMemory_array Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination() As Any, ByRef Ptr() As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal lpSafeArray As Long, lpSafeArrayBound As typSafeArrayBound) As Long







Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260



Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long




Public Function GetArrayStructPtr(hVariant As Variant) As Long

    Dim ArrayStructPtr As typArrayVariant
    
    CopyMemory ArrayStructPtr, hVariant, Len(ArrayStructPtr)
    GetArrayStructPtr = ArrayStructPtr.Pointer

End Function



Public Function mySafeArrayGetDim(ByVal hVariant As Variant) As Long

    mySafeArrayGetDim = SafeArrayGetDim(GetArrayStructPtr(hVariant))

End Function



Public Function IsArrayInit(ByVal SourceArray As Variant) As Boolean

    Dim ndim As Long
    
    IsArrayInit = False

    Select Case VarType(SourceArray)
    Case vbVariant, Is >= vbArray
    ndim = mySafeArrayGetDim(SourceArray)

    If ndim > 0 Then
        IsArrayInit = True
    End If
End Select

End Function



Public Function mySafeArrayRedim(ByVal hVariant As Variant, _
    Optional ByVal ArrayElements, Optional ByVal ArrayLowerBound)

    Dim ndim As Long
    Dim summy As Long
    Dim lpArrayBound As typSafeArrayBound

    ndim = mySafeArrayGetDim(hVariant)

    If IsMissing(ArrayLowerBound) Then
        ArrayLowerBound = LBound(hVariant, ndim)
    End If

    If IsMissing(ArrayElements) Then
        ArrayElements = UBound(hVariant, ndim) - LBound(hVariant, ndim) + 1
    End If

    lpArrayBound.cElements = ArrayElements
    lpArrayBound.LowerBound = ArrayLowerBound

    summy = SafeArrayRedim(GetArrayStructPtr(hVariant), lpArrayBound)

    If summy = S_OK Then
        mySafeArrayRedim = hVariant
    End If

End Function


Public Function TransformArrayToOneDimension(ByVal SourceArray As Variant) As Variant

    Dim TargetArray As Variant
    Dim tPointer As Long
    Dim tSA As typSafeArray
    Dim nCount As Long
    Dim i As Long
    
    TargetArray = SourceArray
    tPointer = GetArrayStructPtr(TargetArray)
    CopyMemory tSA, ByVal tPointer, Len(tSA)
    
    With tSA
        ReDim BoundItem(1 To .nDimension) As typSafeArrayBound
        
        tPointer = tPointer + Len(tSA)
        CopyMemory BoundItem(1), ByVal tPointer, Len(BoundItem(1)) * .nDimension
        
        nCount = BoundItem(1).cElements
        For i = 2 To .nDimension
            nCount = nCount * BoundItem(i).cElements
        Next i
        
        BoundItem(1).cElements = nCount
        BoundItem(1).LowerBound = 1
        
        CopyMemory ByVal tPointer, BoundItem(1), Len(BoundItem(1))
        
        .nDimension = 1
     
    End With
    
    CopyMemory ByVal tPointer - Len(tSA), tSA, Len(tSA)
    
    TransformArrayToOneDimension = TargetArray

End Function

'
'Public Function ReSetArrayPData(Arr As Variant, ByVal pNewData&) As Long
''数组pvData项指向其它位置,返回原位置
'    Dim lAryHeader&
'    CopyMemory lAryHeader, ByVal VarPtr(Arr) + 8, 4
'    CopyMemory lAryHeader, ByVal lAryHeader, 4
'    CopyMemory ReSetArrayPData, ByVal lAryHeader + 12, 4
'    CopyMemory ByVal lAryHeader + 12, pNewData, 4
'End Function
'
'
'
'
'Public Function IsArrayNoneElements(Arr As Variant) As Boolean
''判断数组是否是空
'    Dim lTemp&
'    CopyMemory lTemp, ByVal VarPtr(Arr) + 8, 4
'    If lTemp = 0 Then IsArrayNoneElements = True: Exit Function
'    CopyMemory lTemp, ByVal lTemp, 4
'    If lTemp = 0 Then IsArrayNoneElements = True: Exit Function
'    CopyMemory lTemp, ByVal lTemp + 16, 4
'    If lTemp = 0 Then IsArrayNoneElements = True: Exit Function
'End Function
'
'
'
'
'Public Function CreateArray(Arr As Variant, pData As Long, nElements As Long) As Long
''arr:要创建的数组,pdata:数据指针,nele..:元素个数
''返回一个long值,用于DeleteArray
''数组要自己定义好类型,建立的数组要用DeleteArray来删除,不要用Erase
'    Dim lAryHeader&
'    ReDim Arr(0) '释放/创建数组
'    CopyMemory lAryHeader, ByVal VarPtr(Arr) + 8, 4
'    CopyMemory lAryHeader, ByVal lAryHeader, 4
'    CopyMemory CreateArray, ByVal lAryHeader + 12, 4 '记录原地址
'    CopyMemory ByVal lAryHeader + 12, pData, 4
'    CopyMemory ByVal lAryHeader + 16, nElements, 4
'End Function
'
'
'Public Sub DeleteArray(Arr As Variant, hCreate&)
''用于删除用CreateArray创建的数组,hCreate为CreateArray的返回值
'    Dim lAryHeader&
'    CopyMemory lAryHeader, ByVal VarPtr(Arr) + 8, 4
'    CopyMemory lAryHeader, ByVal lAryHeader, 4
'    CopyMemory ByVal lAryHeader + 12, hCreate, 4 '重新赋值给原地址,以便Erase
'    CopyMemory ByVal lAryHeader + 16, 1&, 4
'    Erase Arr '释放
'End Sub
'


'打开目录对话框
Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String

    On Error GoTo ehBrowseForFolder 'Trap for errors

    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo

    'Set API properties (housed in a UDT)
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Display the browse folder...
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        'Create string of nulls so it will fill in with the path
        strPath = String(MAX_PATH, 0)

        'Retrieves the path selected, places in the null
         'character filled string
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        'Frees memory
        Call CoTaskMemFree(lngIDList)

        'Find the first instance of a null character,
         'so we can get just the path
        intNull = InStr(strPath, vbNullChar)
        'Greater than 0 means the path exists...
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If


    'Return the path name
    BrowseForFolder = strPath
    Exit Function 'Abort


ehBrowseForFolder:

    'Return no value
    BrowseForFolder = Empty

End Function

