Attribute VB_Name = "mdlPubDefine"
Option Explicit

'��������
'######################################################################################################################

'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
Public Const conMenu_FilePopup = 1 '�ļ�
Public Const conMenu_ManagePopup = 2 '����
Public Const conMenu_EditPopup = 3 '�༭
Public Const conMenu_ReportPopup = 4 '����
Public Const conMenu_ViewPopup = 7 '�鿴
Public Const conMenu_ToolPopup = 8 '����
Public Const conMenu_HelpPopup = 9 '����

''�ļ��˵�
Public Const conMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Public Const conMenu_File_Preview = 102         '*Ԥ��(&V)
Public Const conMenu_File_Print = 103           '*��ӡ(&P)
Public Const conMenu_File_Excel = 104           '�����&Excel��
Public Const conMenu_File_Exit = 191            '*�˳�(&X)

'�鿴�˵�
Public Const conMenu_View_ToolBar = 701              '������(&T)
Public Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Public Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Public Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Public Const conMenu_View_StatusBar = 702            '״̬��(&S)
Public Const conMenu_View_Page = 745                '�鿴����
Public Const conMenu_View_Refresh = 791              '*ˢ��(&R)

Public Const conMenu_View_Navigatebeginning = 7401           '*��һ��(&F)
Public Const conMenu_View_Navigateleft = 7402                '*��һ��(&F)
Public Const conMenu_View_Navigateright = 7403               '*��һ��(&F)
Public Const conMenu_View_Navigateend = 7404                 '*���һ��(&F)

'�����˵�
Public Const conMenu_Help_Help = 901        '*��������(&H)
Public Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Public Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Public Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Public Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Public Const conMenu_Help_About = 991       '����(&A)��

'������������
'*********************************************************************
'CommandBar���г�������
Public Const XTP_ID_WINDOW_LIST = 35000 '�����б�
Public Const XTP_ID_TOOLBARLIST = 59392 '�������б�
Public Const ID_INDICATOR_CAPS = 59137 '״̬������д��
Public Const ID_INDICATOR_NUM = 59138 '״̬�������֣�
Public Const ID_INDICATOR_SCRL = 59139 '״̬����������

'CommandBar�����ȼ�
Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

'CommandBar�����
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_PAGEUP = &H21
Public Const VK_PAGEDOWN = &H22
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const VsModiBackColor = &HD6FFCA        'vs�ؼ����ɱ༭��Ԫ�ı���ɫ
'*********************************************************************

Public Type SYSPARAM_INFO
    ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
End Type

Public ParamInfo As SYSPARAM_INFO
Public grsData As ADODB.Recordset
Public grsPage As ADODB.Recordset
Public grsList As ADODB.Recordset
Public grsTempFile As ADODB.Recordset
Public gintStartPage As Integer
Public gstrNo As String
Public gobjRect As USERRECT
Public gobjFont As USERFONT
Public gobjPaper As USERPAPER
Public gobjDraw As Object
Public gclsDataSources As clsDataSources
Public glngVirtualPages As Long
Public gdblWaitTime As Double '��ӡPDF���

'
'######################################################################################################################

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '******************************************************************************************************************
    '���ܣ�
    '******************************************************************************************************************
    MsgBox strInfo, vbInformation, "zl9OpsFormat"
    
End Sub

Public Function CommandBarInit(ByRef cbsMain As Object, Optional ByVal blnEnableCustomization As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
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

    Set cbsMain.Icons = frmPubResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    CommandBarInit = True
    
End Function

Public Function DockPannelInit(ByRef dkpMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    DockPannelInit = True
    
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
        objControl.Id = lngID
        objControl.IconId = IIf(lngIcon = -1, lngID, lngIcon)
        objControl.BeginGroup = blnBeginGroup
        
        If strToolTipText <> "" Then objControl.ToolTipText = strToolTipText

        If objControl.type = xtpControlButton Or objControl.type = xtpControlPopup Then
            objControl.Style = bytStyle
        End If
        
    End With
    
    Set NewToolBar = objControl
    
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
    
    Select Case Control.Id
    Case conMenu_File_PrintSet '��ӡ����
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button '������
    
        For lngLoop = 2 To frmMain.cbsMain.count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_Help_Help              '��������
    
'        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((glngSys) / 100))
        
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        
        Call zlHomePage(frmMain.hwnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hwnd)
        
    Case conMenu_Help_Web_Mail '���ͷ���
        
        Call zlMailTo(frmMain.hwnd)
            
    Case conMenu_Help_About '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Select Case Control.Id
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function InsertData(ByVal str��� As String, _
                                ByVal str���� As String, _
                                Optional ByVal bytHAlignment As Byte = 1, _
                                Optional ByVal str���� As String, _
                                Optional ByVal bytVAlignment As Byte = 2, _
                                Optional ByVal blnWordWarp As Boolean, _
                                Optional ByVal intRows As Integer = 1, _
                                Optional ByVal str��־ As String = "", _
                                Optional ByVal blnAutoFit As Boolean = False, _
                                Optional ByVal blnDebug As Boolean = False, _
                                Optional ByVal strPrex As String = "A", _
                                Optional ByVal bytAngle As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    On Error GoTo errHand
    
    grsData.AddNew
    
    gstrNo = Format(Val(gstrNo) + 1, "0000000000")
    
    grsData("���").value = UCase(strPrex) & gstrNo
    grsData("����").value = IIf(blnDebug, 1, 0)
    grsData("���").value = str���
    grsData("ҳ��").value = gobjRect.Page
    grsData("����").value = str����
    grsData("����").value = str����
    grsData("X0").value = gobjRect.X0
    grsData("Y0").value = gobjRect.Y0
    grsData("X1").value = gobjRect.X1
    grsData("Y1").value = gobjRect.Y1
    grsData("B0").value = gobjRect.B0
    grsData("R0").value = gobjRect.R0
    grsData("����").value = gobjFont.Name
    grsData("ǰ��ɫ").value = gobjFont.ForeColor
    grsData("����ɫ").value = gobjFont.BackColor
    grsData("��С").value = gobjFont.size
    grsData("����").value = IIf(gobjFont.Bold, 1, 0)
    grsData("б��").value = IIf(gobjFont.Italic, 1, 0)
    grsData("�»���").value = IIf(gobjFont.Underline, 1, 0)
    grsData("�������").value = bytHAlignment                                   '1-��;2-��;3-��
    grsData("�������").value = bytVAlignment                                   '1-��;2-��;3-��
    grsData("�Զ�����").value = IIf(blnWordWarp, 1, 0)
    grsData("�������").value = IIf(gobjFont.LineWidth = 0, 1, gobjFont.LineWidth)
    grsData("��������").value = gobjFont.LineStyle
    grsData("����").value = intRows
    grsData("�Զ���Ӧ").value = IIf(blnAutoFit, 1, 0)
    grsData("��ת�Ƕ�").value = bytAngle
    
    InsertData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetLines(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As Integer
    '******************************************************************************************************************
    '���ܣ���ȡ��Ҫ����������Ϊ�п���Ҫ����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim sglSingleChar As Single
    Dim lngChars As Long

    sglSingleChar = objDraw.TextWidth("A")

    lngChars = lngCX \ sglSingleChar

    strSQL = "Select zl_GetTextRows([1],[2]) As ���� From Dual"
    Set rs = zldatabase.OpenSQLRecord(strSQL, "", strText, lngChars)
    If rs.BOF = False Then
        GetLines = zlStr.NVL(rs("����").value)
    End If

    If GetLines = 0 Then GetLines = 1
End Function

Public Function GetLineText2(ByVal objDraw As Object, ByVal strText As String, ByVal intRow As Integer, ByVal lngCX As Long) As String
    '******************************************************************************************************************
    '���ܣ���ȡָ���е����ݣ�������������������������ٸ��ַ���Ȼ����ù��̡�Get_LineText�����ָ��������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim sglSingleChar As Single
    Dim lngChars As Long

    GetLineText2 = strText

    sglSingleChar = objDraw.TextWidth("A")

    lngChars = lngCX \ sglSingleChar

    strSQL = "Select zl_GetText([1],[2],[3]) As ���ı� From Dual"
    Set rs = zldatabase.OpenSQLRecord(strSQL, "", strText, lngChars, intRow)
    If rs.BOF = False Then
        GetLineText2 = zlStr.NVL(rs("���ı�").value)
    End If

End Function

Public Function GetLineText(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡָ���е����ݣ�������������������������ٸ��ַ���Ȼ����ù��̡�Get_LineText�����ָ��������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsLine As ADODB.Recordset
    Dim strLineText  As String
    Dim lngChar As Long
    Dim lngRow As Long
    Dim strChar As String
    Dim strLastChar As String
    Dim strNextChar As String
    Dim blnCrlf As Boolean
    Dim lngTextLength As Long
    
    On Error GoTo errHand
    
    Set rsLine = New ADODB.Recordset
    With rsLine
        .Fields.Append "�к�", adBigInt
        .Fields.Append "����", adVarChar, 1000
        .Open
    End With
    
    lngRow = 0
    strChar = ""
    strLastChar = ""
    strNextChar = ""
    blnCrlf = False
    lngTextLength = Len(strText)
    
    For lngChar = 1 To lngTextLength
        
        blnCrlf = False
        
        strChar = Mid(strText, lngChar, 1)
        strLastChar = strChar
        
        Select Case Asc(strChar)
        Case 13
            '��Ҫ���ж���һ���ַ��Ƿ�Ϊ���з�
            If lngChar + 1 <= lngTextLength Then
                strNextChar = Mid(strText, lngChar + 1, 1)
                        
                If Asc(strNextChar) = 10 Then
                    '�ǻس����з�
                    
                    strLastChar = vbCrLf
                    
                    '�жϴ˻س����з�������λ�ã���Ҫ�ж�����ǰһ�ַ��ǲ���Ҳ�ǻس����з������ַ�������ǵ���Ϊ����
                    If lngChar = 1 Or strLastChar = vbCrLf Then
                        blnCrlf = True
                    End If
                    
                    'ѭ�������ۼ�1��������
                    lngChar = lngChar + 1
                End If
            
            End If
        Case 10
            'ֱ�ӻ���
            blnCrlf = True
        End Select
        
        If objDraw.TextWidth(strLineText & strChar) > lngCX Or blnCrlf Then
            
            lngRow = lngRow + 1
            rsLine.AddNew
            rsLine("�к�").value = lngRow
            rsLine("����").value = strLineText
            
            If blnCrlf = False Then
                strLineText = strChar
            Else
                strLineText = ""
            End If
        Else
            strLineText = strLineText & strChar
        End If
    Next
    
    If strLineText <> "" Then
        lngRow = lngRow + 1
        rsLine.AddNew
        rsLine("�к�").value = lngRow
        rsLine("����").value = strLineText
    End If
    
    Set GetLineText = rsLine
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AppendListData(ByVal strListName As String, ByVal bytList As Byte, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ���ӵ�Ŀ¼����
    '������
    '���أ�
    '******************************************************************************************************************
    
    grsList.AddNew
    
    grsList("Ŀ¼ҳ��").value = intPage - glngVirtualPages
    grsList("Ŀ¼����").value = strListName
    grsList("Ŀ¼����").value = bytList
    grsList("Ŀ¼����").value = 1
    
    AppendListData = True
    
End Function

Public Function CreateTmpFile(Optional ByVal strFile As String = "zl9PeisGroupRpt.tmp") As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strFileTemp As String
    Dim strTempPath As String
    
    strFileTemp = OS.TempPath
    
    CreateTmpFile = strTempPath & strFile

End Function


