Attribute VB_Name = "mdlImapi2Define"
Option Explicit

Public Function GetArrayStructPtr(hVariant As Variant) As Long

    Dim ArrayStructPtr As typArrayVariant
    
    CopyMemory ArrayStructPtr, hVariant, Len(ArrayStructPtr)
    GetArrayStructPtr = ArrayStructPtr.Pointer

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

