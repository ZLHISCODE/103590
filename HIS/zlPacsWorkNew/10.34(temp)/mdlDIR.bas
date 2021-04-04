Attribute VB_Name = "mdlDir"
Option Explicit

Public Enum FO_Operation
    FO_MOVE = 1
    FO_COPY = 2
    FO_DELETE = 3
    FO_RENAME = 4
End Enum

Public Enum FOFlags
    FOF_MULTIDESTFILES = &H1 'Destination specifies multiple files
    FOF_SILENT = &H4 'Don't display progress dialog
    FOF_RENAMEONCOLLISION = &H8 'Rename if destination already exists
    FOF_NOCONFIRMATION = &H10 'Don't prompt user
    FOF_WANTMAPPINGHANDLE = &H20 'Fill in hNameMappings member
    FOF_ALLOWUNDO = &H40 'Store undo information if possible
    FOF_FILESONLY = &H80 'On *.*, don't copy directories
    FOF_SIMPLEPROGRESS = &H100 'Don't show name of each file
    FOF_NOCONFIRMMKDIR = &H200 'Don't confirm making any needed dirs
End Enum

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private op As SHFILEOPSTRUCT

Public Sub DeleteFolder(sDeleteFolder As String, Optional Interface As Boolean = False)
    
    SetAttr sDeleteFolder, vbNormal
    With op
        .wFunc = FO_DELETE
        .pFrom = sDeleteFolder & "*.*"
        .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
    End With
    SHFileOperation op
    
End Sub


Public Function DirExists(ByVal strDir As String) As Boolean
'判断目录是否存在
    Dim objFs As New FileSystemObject
    
    DirExists = objFs.FolderExists(strDir)
End Function


Public Function FileExists(ByVal strFile As String) As Boolean
'判断文件是否存在
    Dim objFs As New FileSystemObject
    
    FileExists = objFs.FileExists(strFile)
End Function
