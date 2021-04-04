Attribute VB_Name = "mdlDir"
Option Explicit

Private op As SHFILEOPSTRUCT

Public Sub DeleteFolder(sDeleteFolder As String, Optional Interface As Boolean = False, Optional blnDelDir As Boolean = True)
On Error Resume Next
    SetAttr sDeleteFolder, vbNormal
    With op
        .wFunc = FO_DELETE
        If blnDelDir Then
            .pFrom = sDeleteFolder & "*.*"
            .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
        Else
            'ֻɾ��Ŀ¼�µ��ļ���������Ŀ¼
            .pFrom = sDeleteFolder & "\*.*"
            .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION + FOF_FILESONLY, FOF_NOCONFIRMATION And FOF_SILENT)
        End If
    End With
    SHFileOperation op
    
End Sub


Public Sub RemoveFile(ByVal strFile As String)
    Dim objFileSys As FileSystemObject
    Dim objFile As File
On Error GoTo errhandle
    Set objFileSys = New FileSystemObject
    Set objFile = objFileSys.GetFile(strFile)

    objFile.Delete True
    
    Set objFile = Nothing
    Set objFileSys = Nothing
Exit Sub
errhandle:

End Sub


Public Function DirExists(ByVal strDir As String) As Boolean
'�ж�Ŀ¼�Ƿ����
    Dim objFs As New FileSystemObject
    
    DirExists = objFs.FolderExists(strDir)
End Function


Public Function FileExists(ByVal strFile As String) As Boolean
'�ж��ļ��Ƿ����,���ܵ���Dir����
    Dim objFs As New FileSystemObject
    
    FileExists = objFs.FileExists(strFile)
End Function


Public Function GetAppRootPath() As String
    Dim strAppRootPath As String
    strAppRootPath = App.Path
    
    If App.LogMode = 0 Then
        'Դ��ģʽ
        strAppRootPath = "C:\Appsoft\"
        If DirExists(strAppRootPath) = False Then strAppRootPath = "D:\Appsoft\"
        If DirExists(strAppRootPath) = False Then strAppRootPath = "E:\Appsoft\"
        If DirExists(strAppRootPath) = False Then strAppRootPath = "F:\Appsoft\"
    Else
        strAppRootPath = Mid(strAppRootPath, 1, InStr(UCase(strAppRootPath), "APPSOFT") + 6) & "\"
    End If
    
    GetAppRootPath = strAppRootPath
End Function















