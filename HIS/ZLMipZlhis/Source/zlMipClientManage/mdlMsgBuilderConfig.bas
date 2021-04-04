Attribute VB_Name = "mdlMspManage"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------
'���Ͷ���
Public Enum Color
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E4E7
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
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
End Enum

'�û���Ϣ
Public Type USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ģ��Ȩ�� As String
    ����Ȩ�� As String
    ��λ���� As String
    �������� As String
End Type

'ϵͳ������Ϣ
Public Type SYSPARAM_INFO
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
    �շ�������Ŀƥ�� As Integer
    ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
End Type

'ѡ�����ļ�
Public Type DlgFileInfo
    iCount As Long   '�ļ���
    sPath As String  'ѡ��·��
    sFile() As String   '�ļ���
End Type

Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


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

Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
'----------------------------------------------------------------------------------------------------------------------
'ȫ�ֱ�������

Public gclsDataOracle As zlDataOracle.clsDataOracle


Public gobjFSO As New Scripting.FileSystemObject        'FSO����
Public ParamInfo As SYSPARAM_INFO
Public UserInfo As USER_INFO
Public gblnInsure As Boolean
Public gstrSQL As String
Public gblnShowInTaskBar As Boolean
Public gfrmMain As Object
Public glngTXTProc As Long                              '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngShareUseID As Long
Public gblnPersonSet As Boolean
Public gfrmPubResource As frmPubResource
Public gclsMsgBase As New clsBusiness
Public gstrSysName As String
Public glngSys As Long
Public Const M_MAX_HEIGHT = 16150  '��Ƭ�����߶�
Public Const M_MAX_WIDTH = 15005   ''��Ƭ�������
Public Const M_MIN_HEIGHT = 3900   '��Ƭ����С�߶�
Public Const M_MIN_WIDTH = 6000        '��Ƭ����С���

'����ͼ�궨��
Public Const Icon_History = 1000
Public Const Icon_Charge = 1001
Public Const Icon_Item = 1002
Public Const Icon_Report = 1003
Public Const Icon_Archives = 1004
Public Const Icon_Package = 1005
Public Const Icon_WaitPerson = 1006
Public Const Icon_NowPerson = 1007
Public Const Icon_OverPerson = 1008

Public Const Icon_Group = 1009
Public Const Icon_Single = 1010
'Private mclsUnzip As New clsUnZip

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'----------------------------------------------------------------------------------------------------------------------
'ģ���������
Public zlCommFun As New clsCommFun
Public zlDataBase As New clsDatabase
Public zlComLib As New clsComLib
Public zlControl As New clsControl


'######################################################################################################################
'�����嵥

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    '******************************************************************************************************************
    '���ܣ�ȥ��TextBox��Ĭ���Ҽ��˵�
    '������
    '���أ�
    '˵���������Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    '******************************************************************************************************************
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��½�û���Ϣ
    '������
    '���أ�
    '******************************************************************************************************************

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    Set rsTmp = zlDataBase.GetUserInfo
    
    UserInfo.���� = UserInfo.�û���
    If Not rsTmp.EOF Then
    
        UserInfo.ID = zlCommFun.NVL(rsTmp("ID").Value, 0)
        UserInfo.��� = zlCommFun.NVL(rsTmp("���").Value)
        UserInfo.����ID = zlCommFun.NVL(rsTmp("����ID").Value, 0)
        UserInfo.���� = zlCommFun.NVL(rsTmp("����").Value)
        UserInfo.���� = zlCommFun.NVL(rsTmp("����").Value)
        UserInfo.�������� = zlCommFun.NVL(rsTmp("������").Value)
        
        GetUserInfo = True

    End If
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter() = 1 Then
        Resume
    End If
    
    Call zlComLib.SaveErrLog
End Function


Public Function NVL(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '����:ȡĳ�ֶε�ֵ
    '����:rsObj          �������ֶ�
    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
    End If
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function GetImageList(Optional ByVal intIconSize As Integer = 16) As ImageList
    
'    If gfrmPubResource Is Nothing Then
'        Set GetImageList = frmPubResource.GetImageList(intIconSize)
'    Else
'        Set GetImageList = gfrmPubResource.GetImageList(intIconSize)
'    End If
    
End Function


'Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
'    '******************************************************************************************************************
'    '����:
'    '����:
'    '����:
'    '******************************************************************************************************************
'    Dim blnTran As Boolean
'    Dim intLoop As Integer
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    If rs.RecordCount > 0 Then
'        If Len(strTitle) = 0 Then strTitle = ParamInfo.ϵͳ����
'        blnTran = True
'
'        If blnHaveTrans Then gcnOracle.BeginTrans
'
'        rs.MoveFirst
'
'        For intLoop = 1 To rs.RecordCount
'
'            strSQL = CStr(rs("SQL").Value)
'            Call zlDataBase.ExecuteProcedure(strSQL, strTitle)
'
'            rs.MoveNext
'        Next
'
'        If blnHaveTrans Then gcnOracle.CommitTrans
'        blnTran = False
'    End If
'
'    SQLRecordExecute = True
'
'    Exit Function
'errHand:
'
'    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
'
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'
'
'End Function



Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '��������
    
        Call zlComLib.ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.ϵͳ��) / 100))
        
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlComLib.zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlComLib.zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlComLib.zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '����
        
        Call zlComLib.ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "0") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strFormat As String
    Dim lngNotPrintCols As Long
    Dim lngPrintCol As Long
    
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.Cols = objVsf.Cols - lngNotPrintCols
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
            lngPrintCol = lngPrintCol + 1
            objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
            objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
            If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                objPrintVsf.ColAlignment(lngPrintCol) = 4
            Else
                objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
            End If
        End If
    Next
    
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                lngPrintCol = lngPrintCol + 1
                
                If objVsf.ColDataType(lngCol) = flexDTBoolean And lngRow >= objVsf.FixedRows Then
                    objPrintVsf.TextMatrix(lngRow, lngPrintCol) = IIf(Abs(Val(objVsf.TextMatrix(lngRow, lngCol))) = 1, "��", "")
                Else
                    strFormat = objVsf.ColFormat(lngCol)
                    If strFormat = "" Then
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Trim(objVsf.TextMatrix(lngRow, lngCol))
                    Else
                        objPrintVsf.TextMatrix(lngRow, lngPrintCol) = Format(objVsf.TextMatrix(lngRow, lngCol), strFormat)
                    End If
                End If
            End If
        Next
'        Call gclsBase.SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
End Function

Public Function CommandBarInit(ByRef cbsMain As CommandBars) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
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

    Set cbsMain.Icons = zlCommFun.GetPubIcons
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
        objControl.ID = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.Type = xtpControlButton Or objControl.Type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
End Function

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.VisualTheme = ThemeOffice2003
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = True 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
        
    DockPannelInit = True
    
End Function



Public Function AppendCode(ByVal strName As String, ByVal strCode As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If strName <> "" And strCode <> "" Then
        AppendCode = "��" & strCode & "��" & strName
    Else
        AppendCode = strName
    End If
End Function

Public Function GetCode(ByVal strName As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    If strName <> "" And InStr(strName, "��") > InStr(strName, "��") Then
        GetCode = Mid(strName, InStr(strName, "��") + 1, InStr(strName, "��") - InStr(strName, "��") - 1)
    End If
    
End Function

Public Function SetPaneRange(dkpMain As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
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
Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, ParamInfo.ϵͳ����
    
End Sub

Public Sub LocationObj(ByRef objTxt As Object, Optional ByVal blnDoevents As Boolean = False)
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If blnDoevents Then DoEvents
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
    
errHand:
    
End Sub

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim strChar As String
    
    strChar = "ZXCVBNMASDFGHJKLQWERTYUIOPzxcvbnmasdfghjklqwertyuiop"
    
    Select Case bytMode
    Case 1          'ȫ����
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 2          'ȫ��ĸ
    
        For lngLoop = 1 To Len(Text)
            If InStr(strChar, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
        
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand
    
'    gblnPersonSet = (Val(zlDataBase.GetPara("ʹ�ø��Ի����")) = 1)

    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    Dim cbrControl3 As CommandBarControl
    '�����˵�����
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.Visible = cbrControl.Visible
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
            
                Set cbrControl3 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrControl3.Parameter = cbrControl2.Parameter
                cbrControl3.Visible = cbrControl2.Visible
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub CloseRecord(rs As ADODB.Recordset)
    '******************************************************************************************************************
    '����:�ر�����
    '����:
    '����:
    '******************************************************************************************************************
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Function WriteBinByFile(strFile As String, objField As Field) As Boolean
    '******************************************************************************************************************
    '����:д���ļ�
    '����:
    '����:
    '******************************************************************************************************************
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, intFile, i As Long
    Dim arrBin() As Byte
    
    On Error GoTo errH
    
    intFile = FreeFile
    Open strFile For Binary Access Read As intFile
    lngFileSize = LOF(intFile)
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    objField.Value = Null
    For i = 0 To intBolcks
        If i = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim arrBin(lngCurSize - 1) As Byte
        Get intFile, , arrBin()
        objField.AppendChunk arrBin()
    Next
    Close intFile
    WriteBinByFile = True
    Exit Function
errH:
    Close intFile
End Function

Public Function GetDlgFileInfo(strFileName As String) As DlgFileInfo
    '******************************************************************************************************************
    '����:����CommonDialog��ѡ����ļ�������·�����ļ���
    '����:strFileNameΪCommonDialog��Filename����,DlgFileInfo������һ���Զ������ͣ�����iCount������ѡ���ļ��ĸ�����sPath������ѡ
    '����:����CommonDialog��ѡ����ļ�������·�����ļ���
    '******************************************************************************************************************
    Dim sPath, tmpStr As String
    Dim sFile() As String
    Dim iCount As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    sPath = CurDir()
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '���ļ�����·������
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        'ѡ���˶���ļ�(������һ���ַ�ΪChr$(0))
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFile(iCount)
            Else
                sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·�������û��"\"��
        iCount = 1
        ReDim Preserve sFile(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFile(iCount) = tmpStr
    End If
    
    GetDlgFileInfo.iCount = iCount
    ReDim GetDlgFileInfo.sFile(iCount)
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgFileInfo.sPath = sPath
    
    For i = 1 To iCount
        GetDlgFileInfo.sFile(i) = sFile(i)
    Next
    
    Exit Function

errHandle:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    Call zlComLib.SaveErrLog
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "SQL", adVarChar, 4000
        .Fields.Append "Trans", adTinyInt                   '1��ʾ��ʼ;2��ʾ����
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500
        
        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function AddPeriodToComboBox(ByRef cbo As Object)
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    With cbo
        .Clear
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "��  ��"
        .AddItem "������"
        .AddItem "��  ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰ����"
        .AddItem "ǰһ��"
        .AddItem "ǰ����"
        .AddItem "�Զ���"
    End With
    
    AddPeriodToComboBox = True
    
End Function

Public Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '����:��ȡ����ʱ��
    '����:
    '����:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "�Զ���" Then
        '�Զ���:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), zlDataBase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), zlDataBase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "��  ��"
        GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ʱ"      '��ʱ
        GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(zlDataBase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(zlDataBase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰ����"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(zlDataBase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDataBase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function


'Public Function ShowPubSelect(ByVal frmParent As Object, _
'                                ByVal obj As Object, _
'                                ByVal mrsInfoTree As ADODB.Recordset, _
'                                ByVal strDataInfo As String, _
'                                Optional ByVal lngCX As Long = 9000, _
'                                Optional ByVal lngCY As Long = 4500) As String
'    '******************************************************************************************************************
'    '���ܣ�������+�б�ṹ,Ӧ���ڱ��ؼ�
'    '������
'    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
'    '���أ�0:ȡ��ѡ��;1:ѡ��;2:�����ݷ���
'    '******************************************************************************************************************
'
'    Dim lngX As Long
'    Dim lngY As Long
'    Dim lngObjHeight As Long
'    Dim rs As New ADODB.Recordset
'    Dim objPoint As POINTAPI
'    Dim objSelectDialog As frmEventMsgEditSelect
'
'    On Error GoTo errHand
'
'
'    If obj Is Nothing Then
'        lngX = (Screen.Width - lngCX) / 2
'        lngY = (Screen.Height - lngCY) / 2
'        lngObjHeight = 0
'    Else
'        Call ClientToScreen(obj.hWnd, objPoint)
'
'        lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
'        lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
'        lngObjHeight = obj.CellHeight
'
'    End If
'
'    Set objSelectDialog = New frmEventMsgEditSelect
'
'    ShowPubSelect = objSelectDialog.ShowSelect(frmParent, mrsInfoTree, lngX, lngY, lngCX, lngCY, strDataInfo)
'
'    Exit Function
'
'errHand:
''    If ErrCenter = 1 Then
''        Resume
''    End If
'End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003, _
                                Optional ByVal bytPosition As XTPTabPosition = xtpTabPositionTop) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With tbc

        With .PaintManager

            .Appearance = bytAppearance
            .ClientFrame = xtpTabFrameSingleLine
            .Color = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
            .BoldSelected = True
            .Position = bytPosition
        End With

'        Set .Icons = frmPubResource.imgPublic.Icons
    End With

    TabControlInit = True

End Function

