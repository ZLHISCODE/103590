Attribute VB_Name = "mdlDir"
Option Explicit

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
