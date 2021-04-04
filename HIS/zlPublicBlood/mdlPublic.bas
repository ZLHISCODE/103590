Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gstrSysName As String                'ϵͳ����
Public gstrUnitName As String               '�û���λ����
Public gstrProductName As String    '��Ʒ����

Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gobjGrid As Object

Public gstrDBUser As String
Public gstrNodeNo As String          '��ǰվ���ţ����δ��������վ�㣬��Ϊ"-"
Public gcolPrivs As Collection              '��¼�ڲ�ģ���Ȩ��

Public gstrLike As String  '��Ŀƥ�䷽��,%���
Public gblnMyStyle As Boolean 'ʹ�ø��Ի����
Public gstrIme As String '�Զ��Ŀ������뷨
Public gbytCode As Byte '�������ɷ�ʽ��0-ƴ��,1-���,2-����

Public gbytBillOpt As Byte '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
Public gbyt������˷�ʽ As Byte '49501:������˷�ʽ:0-δ��˲�������ʣ�ȱʡΪ0;1-���ʱ����������ú�ҽ��������ҽ�������ͷ��õ�����
Public gblnҽ�����ͺ�Ѫ As Boolean '��Ѫҽ�����ͺ���ܽ��з�Ѫ 0'�¿��󼴿ɷ�Ѫ;1-ҽ�����ͺ�Ѫ
Public gstrҽ���˶� As String '��ѪƤ��ҽ����Ҫ�˶� ��λ��ȡ11����һλΪ ��Ѫҽ�����ڶ�λΪ Ƥ��ҽ��
Public gbln���պ����ִ�� As Boolean 'ѪҺֻ�н��պ˶Ժ������ִ�еǼ�

Public glngOld As Long
Public glngTXTProc As Long

Public Type TYPE_USER_INFO
    id As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ���� As String
    ����ID As Long
    ������ As String
    ������ As String
    רҵ����ְ�� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Type FormRect
    MaxRect As POINTAPI
    MinRect As POINTAPI
End Type
Private mFromRect As FormRect

Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const GWL_WNDPROC = -4
Public Const WM_MOUSEWHEEL = &H20A '��������Ϣ

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
'�������洰�壨��Ļ���ľ��
Public Declare Function GetDesktopWindow Lib "user32" () As Long
'��ȡ�ͻ��������
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWHEELSCROLLLINES = 104
Public WHEEL_SCROLL_LINES As Long
'-------------------------------------------------------------------------------------------------------------------------
'��ͼ���API
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5                    '������ͼ

Public glngBooldPepWinProc As Long
Public gobjFScrollBar As Object  '����������

Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = gobjComlib.Nvl(varValue, DefaultValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
'������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    ZVal = gobjComlib.ZVal(varValue, blnForceNum)
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    SysColor2RGB = gobjComlib.os.SysColor2RGB(lngColor)
End Function

Public Function GetControlRect(ByVal lngHwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip/Pixel)
'���أ�blnTwip=True-����Twip��λ��False-�������ص�λ
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function InDesign() As Boolean
    InDesign = gobjComlib.os.IsDesinMode
End Function

Private Function Custom_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = mFromRect.MinRect.X \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = mFromRect.MinRect.Y \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = mFromRect.MaxRect.X \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = mFromRect.MaxRect.Y \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function

Public Sub SetFormSize(ByVal frmParent As Object, ByVal lngMinWidth As Long, ByVal lngMinHeight As Long, ByVal lngMaxWidth As Long, ByVal lngMaxHeight As Long, ByRef lngWindowLong As Long, Optional ByVal blnUnLoad As Boolean = False)
'���ܣ����ô����С��Χ
'˵����������ػ�ȡlngWindowLong������ж�ش���lngWindowLong
    If Not InDesign Then
        If blnUnLoad = False Then
            mFromRect.MinRect.X = lngMinWidth
            mFromRect.MinRect.Y = lngMinHeight
            mFromRect.MaxRect.X = lngMaxWidth
            mFromRect.MaxRect.Y = lngMaxHeight
            glngOld = GetWindowLong(frmParent.hWnd, GWL_WNDPROC)
            lngWindowLong = glngOld
            Call SetWindowLong(frmParent.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
        Else
            Call SetWindowLong(frmParent.hWnd, GWL_WNDPROC, IIf(lngWindowLong = 0, glngOld, lngWindowLong))
        End If
    End If
End Sub

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function IsPrivs(ByVal strPrivs As String, ByVal strPriv As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If InStr(";" & strPrivs & ";", ";" & strPriv & ";") > 0 Then
        IsPrivs = True
    Else
        IsPrivs = False
    End If
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
        
    With cbsMain.Options
        .ShowExpandButtonAlways = blnEnableCustomization
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization blnEnableCustomization

    Set cbsMain.Icons = gobjCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
End Function

Public Function NewCommandBar(objMenu As CommandBarControl, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal strParameter As String) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption)
        
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        objControl.Parameter = strParameter
        
    End With
    Set NewCommandBar = objControl
End Function

Public Function NewToolBar(objBar As CommandBar, _
                                ByVal xtpType As XTPControlType, _
                                ByVal lngID As Long, _
                                ByVal strCaption As String, _
                                Optional ByVal blnBeginGroup As Boolean, _
                                Optional ByVal lngIcon As Long = -1, _
                                Optional ByVal bytStyle As Byte = xtpButtonIconAndCaption, _
                                Optional ByVal strToolTipText As String, _
                                Optional ByVal intBefore As Integer) As CommandBarControl
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objControl As CommandBarControl
    
    With objBar.Controls
        Set objControl = .Add(xtpType, lngID, strCaption, intBefore)
        objControl.id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    Set NewToolBar = objControl
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnTrans-�Ƿ��������
    '     blnCommit-ִ������̺�,�ύ����(ǰ��:blnTrans=true)
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnTrans Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call gobjDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnCommit And blnTrans Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    Set rs = New ADODB.Recordset
    With rs
        .Fields.Append "SQL", adVarChar, 300
        .Fields.Append "Trans", adTinyInt                   '1��ʾ��ʼ;2��ʾ����
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        .Open
    End With
    SQLRecord = True
    Exit Function
ErrHand:
   If gobjComlib.ErrCenter = 1 Then
        Resume
   End If
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
   End If
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        If blnHaveTrans Then gcnOracle.BeginTrans
        rs.MoveFirst
        For intLoop = 1 To rs.RecordCount
            strSQL = CStr(rs("SQL").Value)
            Call gobjDatabase.ExecuteProcedure(strSQL, strTitle)
            rs.MoveNext
        Next
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    Exit Function
ErrHand:
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
   End If
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.id = rsTmp!id
            UserInfo.�û��� = rsTmp!User
            UserInfo.��� = rsTmp!���
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.���� = Nvl(rsTmp!����)
            UserInfo.����ID = Nvl(rsTmp!����ID, 0)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.������ = Nvl(rsTmp!������)
            UserInfo.���� = Get��Ա����
            UserInfo.רҵ����ְ�� = Nvl(rsTmp!רҵ����ְ��)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.�û���
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get��Ա����(Optional ByVal str���� As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��¼��Ա��ָ����Ա����Ա����
    '����:������Ա����,����ö��ŷ���
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If str���� <> "" Then
        strSQL = "Select B.��Ա���� From ��Ա�� A,��Ա����˵�� B Where A.ID=B.��ԱID And A.����=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", str����)
    Else
        strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��ԱID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��Ա����", UserInfo.id)
    End If
    Do While Not rsTmp.EOF
        Get��Ա���� = Get��Ա���� & "," & rsTmp!��Ա����
        rsTmp.MoveNext
    Loop
    Get��Ա���� = Mid(Get��Ա����, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����������ض���
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    
    Err = 0: On Error Resume Next
    
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Set gobjGrid = GetObject("", "zl9Comlib.clsGrid")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function

Public Sub InitLocPar()
    Dim strValue As String
    
    On Error Resume Next
    gstrLike = IIf(gobjDatabase.GetPara("����ƥ��") = 0, "%", "")
    strValue = gobjDatabase.GetPara("���뷨")
    gstrIme = IIf(strValue = "", "���Զ�����", strValue)
    gbytCode = Val(gobjDatabase.GetPara("���뷽ʽ"))
    gblnMyStyle = gobjDatabase.GetPara("ʹ�ø��Ի����") = "1"
    
    '���ѽ��ʵļ��ʵ��ݵĲ���Ȩ��:0-����,1-����,2-��ֹ��
    gbytBillOpt = Val(gobjDatabase.GetPara(23, 100))
    gbyt������˷�ʽ = Val(gobjDatabase.GetPara(185, 100))    '49501
    
    '��Ѫ��Ƥ��ҽ��ִ�к���Ҫ�˶�
    gstrҽ���˶� = gobjDatabase.GetPara(186, 100)
    '��Ѫҽ�����ͺ���ܷ�Ѫ
    gblnҽ�����ͺ�Ѫ = Val(gobjDatabase.GetPara(272, 100)) = 1
    'ѪҺ������պ������ִ�еǼ�
    gbln���պ����ִ�� = Val(gobjDatabase.GetPara(301, 100, , 1)) = 1
    If Err <> 0 Then Err.Clear
End Sub

Public Function Between(X, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    Between = gobjComlib.Between(X, a, b)
End Function

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = gobjComlib.IntEx(vNumber)
End Function

Public Function GetInsidePrivs(ByVal lngSys As Long, ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs(lngSys & "_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove lngSys & "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(lngSys, lngProg)
        gcolPrivs.Add strPrivs, lngSys & "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function


Public Sub zlRptPrint(ByVal bytMode As Byte, ByVal objvsf As Object, ByVal strText As String)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    If objvsf.Rows = objvsf.FixedRows Then Exit Sub
    Set objPrint.Body = objvsf
    
    objPrint.Title.Text = strText
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Public Function SetPaneRange(dkpM As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�����dockingpane�Ĵ�С��Χ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpM.FindPane(intPane)
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    SetPaneRange = True
End Function

Public Function SetPublicFontSize(ByRef frmMe As Object, ByVal bytSize As Byte, Optional ByVal strOther As String)
'���ܣ����ô��弰���пؼ��������С
'������frmMe=��Ҫ��������Ĵ������
'      bytSize:����Ϊ9������,0:����Ϊ9������,1,����Ϊ12������
'      strOther:�������������õĿؼ��������ļ���,��ʽΪ����������1,��������2,��������3,....
'˵����1.����漰��VsFlexGrid�ȱ��ؼ�����Ҫ�������ڵĻ������µ����п���и�
'      2.�������δ�г��������ؼ����Զ���ؼ�,��Ҫ���ض�����ָ�������С����ش���ģ������ⵥ������

    Dim objCtrol As Control, objrptCol As Object
    Dim CtlFont As StdFont
    Dim i As Long, lngOldSize As Long
    Dim lngFontSize As Long
    Dim dblRate As Double
    Dim blnDo As Boolean
    Dim strContainer As String
    
    lngFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    frmMe.FontSize = lngFontSize
    strOther = "," & strOther & ","
    blnDo = False
        
    For Each objCtrol In frmMe.Controls
        Select Case TypeName(objCtrol)
            Case "TabStrip", "Label", "ComboBox", "ListView", "OptionButton", "CheckBox", "DTPicker", "TextBox", "ReportControl", _
                "DockingPane", "CommandBars", "TabControl", "CommandButton", "Frame", "RichTextBox", "MaskEdBox", "IDKind", "PatiIdentify", "VSFlexGrid"
                blnDo = True
            Case Else
                blnDo = False
        End Select
        
        If strOther <> ",," And blnDo Then
            '����CommandBars�û��Զ���ؼ���ȡobjCtrol.Container�����
            strContainer = ""
            On Error Resume Next
            strContainer = objCtrol.Container.name
            Err.Clear: On Error GoTo 0
            If InStr(1, strOther, "," & strContainer & ",") > 0 Then
                 blnDo = False
            End If
        End If
        
        If blnDo Then
            Select Case TypeName(objCtrol)
                Case "TabStrip"
                        objCtrol.Font.Size = lngFontSize
                Case "Label"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Height = frmMe.TextHeight("��") + 20
                        'Label�����Ҫ���е���
               Case "ComboBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "ListView"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        For i = 1 To objCtrol.ColumnHeaders.Count
                            objCtrol.ColumnHeaders(i).Width = objCtrol.ColumnHeaders(i).Width * dblRate
                        Next
                Case "OptionButton"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("����" & objCtrol.Caption)
                        objCtrol.Height = objCtrol.Height * dblRate
                Case "CheckBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                Case "DTPicker"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = frmMe.TextWidth("2012-01-01    ")
                        objCtrol.Height = frmMe.TextHeight("��") + IIf(bytSize = 0, 100, 120)
                Case "TextBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��") + 60
                Case "MaskEdBox"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = objCtrol.Width * dblRate
                        objCtrol.Height = frmMe.TextHeight("��") + 90
                Case "ReportControl"
                        lngOldSize = objCtrol.PaintManager.TextFont.Size
                        dblRate = lngFontSize / lngOldSize
                        
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        Set CtlFont = objCtrol.PaintManager.TextFont
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.TextFont = CtlFont
                        For Each objrptCol In objCtrol.Columns
                            objrptCol.Width = objrptCol.Width * dblRate
                        Next
                        objCtrol.Redraw
                Case "DockingPane"
                        Set CtlFont = objCtrol.PaintManager.CaptionFont
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.CaptionFont = CtlFont
                        
                        Set CtlFont = objCtrol.TabPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.TabPaintManager.Font = CtlFont
        
                        Set CtlFont = objCtrol.PanelPaintManager.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PanelPaintManager.Font = CtlFont
                Case "CommandBars"
                        Set CtlFont = objCtrol.Options.Font
                        If CtlFont Is Nothing Then '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.Options.Font = CtlFont
                Case "TabControl"
                        Set CtlFont = objCtrol.PaintManager.Font
                        If CtlFont Is Nothing Then  '�ؼ���ʼ����ʱCtlFontΪnothing
                            Set CtlFont = frmMe.Font
                        End If
                        CtlFont.Size = lngFontSize
                        Set objCtrol.PaintManager.Font = CtlFont
                        objCtrol.PaintManager.Layout = xtpTabLayoutAutoSize
                Case "CommandButton"
                        lngOldSize = objCtrol.FontSize
                        dblRate = lngFontSize / lngOldSize
                        
                        objCtrol.FontSize = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "Frame"
                        objCtrol.FontSize = lngFontSize
                Case "IDKind"
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "PatiIdentify"
                        lngOldSize = objCtrol.Font.Size
                        dblRate = lngFontSize / lngOldSize
                        objCtrol.IDKindFont.Size = lngFontSize
                        objCtrol.objIDKind.Refrash
                        objCtrol.Font.Size = lngFontSize
                        objCtrol.Width = dblRate * objCtrol.Width
                        objCtrol.Height = dblRate * objCtrol.Height
                Case "VSFlexGrid"
                    Call gobjControl.VSFSetFontSize(objCtrol, lngFontSize)
                Case "RichTextBox"
                    Call gobjControl.RTFSetFontSize(objCtrol, lngFontSize)
            End Select
        End If
    Next
End Function

Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'�������ã�ʹ��Cbo.SeekIndex����
'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Public Function FlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
'֧�ֹ��ֵĹ���
    Dim lngScrollLines As Long
    
    If Not gobjFScrollBar Is Nothing Then
        lngScrollLines = gobjFScrollBar.LargeChange
        If lngScrollLines > gobjFScrollBar.Max Then lngScrollLines = gobjFScrollBar.Max
        
        Select Case wMsg
        Case WM_MOUSEWHEEL
            Select Case wParam
            Case -7864320  '���¹�
                If gobjFScrollBar.Value + lngScrollLines > gobjFScrollBar.Max Then
                    gobjFScrollBar.Value = gobjFScrollBar.Max
                Else
                    gobjFScrollBar.Value = gobjFScrollBar.Value + lngScrollLines
                End If
            Case 7864320   '���Ϲ�
                If gobjFScrollBar.Value - lngScrollLines < gobjFScrollBar.Min Then
                    gobjFScrollBar.Value = gobjFScrollBar.Min
                Else
                    gobjFScrollBar.Value = gobjFScrollBar.Value - lngScrollLines
                End If
            End Select
        End Select
    End If
    FlexScroll = CallWindowProc(glngBooldPepWinProc, hWnd, wMsg, wParam, lParam)
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo ErrHand

    Select Case enmRegister
    Case ע����Ϣ
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
    Case ˽��ģ��
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
    Case ˽��ȫ��
        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strKeyValue)
    Case ����ģ��
        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
    Case ����ȫ��
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
    End Select
    SetRegister = True
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo ErrHand
    
    Select Case enmRegister
    Case ע����Ϣ
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
    Case ˽��ģ��
        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
    Case ˽��ȫ��
        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.�û��� & "\" & strSection, strKey, strDefKeyValue)
    Case ����ģ��
        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
    Case ����ȫ��
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
    End Select
    GetRegister = strValue
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
