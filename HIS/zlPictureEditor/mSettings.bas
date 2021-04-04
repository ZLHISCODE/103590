Attribute VB_Name = "mSettings"
'================================================
' ���ñ���
'================================================
Option Explicit

Public Function AppPath() As String
    AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\")
End Function

Public Sub LoadMainSettings()
    With gfrmMain
        .Width = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainWidth", (Screen.Width - 12000) / 2)
        .Height = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainHeight", (Screen.Width - 9000) / 2)
        .Top = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainTop", (Screen.Height - .Height) \ 2)
        .Left = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainLeft", (Screen.Width - .Width) \ 2)
        .WindowState = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainWindowState", .WindowState)
        .LastPath = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainLastPath", vbNullString)
        .DialogPreview = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogPreview", -1)
        .DialogFitMode = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogFitMode", -1)
        .DialogJPEGquality = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogJPEGquality", 90)
    End With
End Sub
    
Public Sub LoadPanViewSettings()
    With gfPanView
        .Width = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewWidth", .Width)
        .Height = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewHeight", .Height)
        .Top = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewTop", .Top)
        .Left = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewLeft", .Left)
    End With
End Sub
    
Public Sub LoadFilterSettings()
    With gfFilter
        .Top = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterTop", .Top)
        .Left = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterLeft", .Left)
        .chkFit = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterBeforeFit", 1)
        .chkPickColor = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterBeforePickColor", 1)
        .chkResetValues = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterResetValues", 0)
        .chkNoClose = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterNoClose", 0)
    End With
End Sub

Public Sub LoadTexturizeSettings()
    With gfTexturize
        .Top = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeTop", .Top)
        .Left = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeLeft", .Left)
        .sbWeight = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeWeight", 25)
        .chkInvertTexture = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeInvert", 0)
        .chkFitMode = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFitMode", 0)
        .chkNoClose = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeNoClose", 0)
        On Error GoTo ErrPath
        .flTextures.Path = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFolder", AppPath)
        .flTextures.ListIndex = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFile", "<None>")
        On Error GoTo 0
    End With
    Exit Sub
ErrPath:
    gfTexturize.flTextures.Path = AppPath
End Sub

Public Sub LoadResizeSettings()
    With gfResize
        .chkAspectRatio = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "ResizeAspectRatio", 1)
        .chkResample = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "ResizeResample", 1)
    End With
End Sub

'========================================================================================

Public Sub SaveMainSettings()
    With gfrmMain
        If (.WindowState = vbNormal) Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainWidth", .Width
            SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainHeight", .Height
            SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainTop", .Top
            SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainLeft", .Left
        End If
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainWindowState", .WindowState
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainLastPath", .LastPath
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogPreview", .DialogPreview
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogFitMode", .DialogFitMode
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "MainDialogJPEGquality", .DialogJPEGquality
    End With
End Sub
    
Public Sub SavePanViewSettings()
    With gfPanView
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewWidth", .Width
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewHeight", .Height
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewTop", .Top
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "PanViewLeft", .Left
    End With
End Sub
    
Public Sub SaveFilterSettings()
    With gfFilter
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterTop", .Top
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterLeft", .Left
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterBeforeFit", .chkFit
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterBeforePickColor", .chkPickColor
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterResetValues", .chkResetValues
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "FilterNoClose", .chkNoClose
    End With
End Sub
    
Public Sub SaveTexturizeSettings()
    With gfTexturize
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeTop", .Top
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeLeft", .Left
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFolder", .flTextures.Path
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFile", .flTextures.ListIndex
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeWeight", .sbWeight
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeInvert", .chkInvertTexture
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeFitMode", .chkFitMode
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "TexturizeNoClose", .chkNoClose
    End With
End Sub

Public Sub SaveResizeSettings()
    With gfResize
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "ResizeAspectRatio", .chkAspectRatio
        SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\zlPictureEditor", "ResizeResample", .chkResample
    End With
End Sub
