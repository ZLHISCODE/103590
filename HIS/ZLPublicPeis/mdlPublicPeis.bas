Attribute VB_Name = "mdlPublicPeis"
Option Explicit
'######################################################################################################################
'��������

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Enum Color
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HF0F0F0 '&HE0E4E7
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    
    ԭʼ���� = 0
    ������¼ = &HFF
    ͣ����Ŀ = &H8000000C
    ������Ŀ = 0
    
    ����ģ��ɫ = &HC00000
    
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
    ����ƫ�߱���ɫ = 255
    ����ƫ�ͱ���ɫ = 255
    ����ƫ�߱���ɫ = 65280
    ����ƫ�ͱ���ɫ = 12648384
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
End Enum

Public Const ETO_OPAQUE = 2

'ϵͳ������Ϣ
Public Type SYSPARAM_INFO
    ϵͳ���� As String
    ��Ʒ���� As String
End Type


'----------------------------------------------------------------------------------------------------------------------
'ȫ�ֱ�������
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
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = gobjComLib.NVL(varValue, DefaultValue)
End Function

'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function


'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Sub ShowSimpleMsg(ByVal strInfo As String)

    MsgBox strInfo, vbInformation, ParamInfo.ϵͳ����
    
End Sub

'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
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
'���ܣ�
'������
'���أ�
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
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    
End Function

'******************************************************************************************************************
'���ܣ�
'������
'���أ�
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
'���ܣ�
'������
'���أ�
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
