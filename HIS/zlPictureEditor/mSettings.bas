Attribute VB_Name = "mSettings"
'================================================
' 设置保存
'================================================
Option Explicit

Public Function AppPath() As String
    AppPath = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\")
End Function

Public Sub LoadMainSettings()
    With gfrmMain
        .Width = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainWidth", (Screen.Width - 12000) / 2)
        .Height = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainHeight", (Screen.Width - 9000) / 2)
        .Top = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainTop", (Screen.Height - .Height) \ 2)
        .Left = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainLeft", (Screen.Width - .Width) \ 2)
        .WindowState = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainWindowState", .WindowState)
        .LastPath = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainLastPath", vbNullString)
        .DialogPreview = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogPreview", -1)
        .DialogFitMode = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogFitMode", -1)
        .DialogJPEGquality = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogJPEGquality", 90)
    End With
End Sub
    
Public Sub LoadPanViewSettings()
    With gfPanView
        .Width = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewWidth", .Width)
        .Height = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewHeight", .Height)
        .Top = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewTop", .Top)
        .Left = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewLeft", .Left)
    End With
End Sub
    
Public Sub LoadFilterSettings()
    With gfFilter
        .Top = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterTop", .Top)
        .Left = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterLeft", .Left)
        .chkFit = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterBeforeFit", 1)
        .chkPickColor = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterBeforePickColor", 1)
        .chkResetValues = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterResetValues", 0)
        .chkNoClose = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterNoClose", 0)
    End With
End Sub

Public Sub LoadTexturizeSettings()
    With gfTexturize
        .Top = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeTop", .Top)
        .Left = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeLeft", .Left)
        .sbWeight = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeWeight", 25)
        .chkInvertTexture = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeInvert", 0)
        .chkFitMode = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFitMode", 0)
        .chkNoClose = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeNoClose", 0)
        On Error GoTo ErrPath
        .flTextures.Path = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFolder", AppPath)
        .flTextures.ListIndex = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFile", "<None>")
        On Error GoTo 0
    End With
    Exit Sub
ErrPath:
    gfTexturize.flTextures.Path = AppPath
End Sub

Public Sub LoadResizeSettings()
    With gfResize
        .chkAspectRatio = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "ResizeAspectRatio", 1)
        .chkResample = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "ResizeResample", 1)
    End With
End Sub

'========================================================================================

Public Sub SaveMainSettings()
    With gfrmMain
        If (.WindowState = vbNormal) Then
            SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainWidth", .Width
            SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainHeight", .Height
            SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainTop", .Top
            SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainLeft", .Left
        End If
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainWindowState", .WindowState
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainLastPath", .LastPath
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogPreview", .DialogPreview
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogFitMode", .DialogFitMode
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "MainDialogJPEGquality", .DialogJPEGquality
    End With
End Sub
    
Public Sub SavePanViewSettings()
    With gfPanView
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewWidth", .Width
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewHeight", .Height
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewTop", .Top
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "PanViewLeft", .Left
    End With
End Sub
    
Public Sub SaveFilterSettings()
    With gfFilter
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterTop", .Top
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterLeft", .Left
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterBeforeFit", .chkFit
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterBeforePickColor", .chkPickColor
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterResetValues", .chkResetValues
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "FilterNoClose", .chkNoClose
    End With
End Sub
    
Public Sub SaveTexturizeSettings()
    With gfTexturize
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeTop", .Top
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeLeft", .Left
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFolder", .flTextures.Path
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFile", .flTextures.ListIndex
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeWeight", .sbWeight
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeInvert", .chkInvertTexture
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeFitMode", .chkFitMode
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "TexturizeNoClose", .chkNoClose
    End With
End Sub

Public Sub SaveResizeSettings()
    With gfResize
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "ResizeAspectRatio", .chkAspectRatio
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\zlPictureEditor", "ResizeResample", .chkResample
    End With
End Sub
