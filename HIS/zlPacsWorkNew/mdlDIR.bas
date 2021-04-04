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
            'ֻɾ��Ŀ¼�µ��ļ���������Ŀ¼
            .pFrom = sDeleteFolder & "\*.*"
            .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION + FOF_FILESONLY, FOF_NOCONFIRMATION And FOF_SILENT)
        End If
    End With
    SHFileOperation op
    
End Sub


Public Function DirExists(ByVal strDir As String) As Boolean
'�ж�Ŀ¼�Ƿ����
    Dim objFs As New FileSystemObject
    
    DirExists = objFs.FolderExists(strDir)
End Function


Public Function FileExists(ByVal strFile As String) As Boolean
'�ж��ļ��Ƿ����
    Dim objFs As New FileSystemObject
    
    FileExists = objFs.FileExists(strFile)
End Function
