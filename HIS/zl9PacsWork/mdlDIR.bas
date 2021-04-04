Attribute VB_Name = "mdlDir"
Option Explicit

Private op As SHFILEOPSTRUCT

Public Sub DeleteFolder(sDeleteFolder As String, Optional Interface As Boolean = False, Optional blnDelDir As Boolean = True)

    SetAttr sDeleteFolder, vbNormal
    With op
        .wFunc = FO_DELETE
        If blnDelDir Then
            .pFrom = sDeleteFolder & "*.*"
            .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
        Else
            '只删除目录下的文件，保留该目录
            .pFrom = sDeleteFolder & "\*.*"
            .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION + FOF_FILESONLY, FOF_NOCONFIRMATION And FOF_SILENT)
        End If
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
