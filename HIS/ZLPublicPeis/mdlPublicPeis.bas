Attribute VB_Name = "mdlPublicPeis"
Option Explicit
'######################################################################################################################
'常量定义

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Enum Color
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HF0F0F0 '&HE0E4E7
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    
    原始单据 = 0
    冲销记录 = &HFF
    停用项目 = &H8000000C
    启用项目 = 0
    
    公共模块色 = &HC00000
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    警戒偏高背景色 = 255
    警戒偏低背景色 = 255
    复查偏高背景色 = 65280
    复查偏低背景色 = 12648384
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
End Enum

Public Const ETO_OPAQUE = 2

'系统参数信息
Public Type SYSPARAM_INFO
    系统名称 As String
    产品名称 As String
End Type


'----------------------------------------------------------------------------------------------------------------------
'全局变量申明
Public ParamInfo As SYSPARAM_INFO

Public gcnOracle As ADODB.Connection
Public gobjComLib As Object
Public gobjComFun As Object
Public gobjDatabase As Object
Public gobjReport As Object
Public glngSys As Long
Public gclsPackage As New clsPackage

Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Sub InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = gobjComLib.NVL(varValue, DefaultValue)
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Sub ShowSimpleMsg(ByVal strInfo As String)

    MsgBox strInfo, vbInformation, ParamInfo.系统名称
    
End Sub

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = gobjComFun.GetPubIcons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.Id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = bytOptions
    
End Function

'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpMain.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function
