Attribute VB_Name = "mdlCISAudit"
Option Explicit

'######################################################################################################################
'��������

Public Const strSplitCmb = "��"
Public Const SPI_GETWORKAREA = 48
Public Enum DbType
    T_AnsiString = 0
    T_Binary = 1
    T_Byte = 2
    T_Boolean = 3
    T_Currency = 4
    T_Date = 5
    T_DateTime = 6
    T_Decimal = 7
    T_Double = 8
    T_Guid = 9
    T_Int16 = 10
    T_Int32 = 11
    T_Int64 = 12
    T_Object = 13
    T_SByte = 14
    T_Single = 15
    T_String = 16
    T_Time = 17
    T_UInt16 = 18
    T_UInt32 = 19
    T_UInt64 = 20
    T_VarNumeric = 21
    T_AnsiStringFixedLength = 22
    T_StringFixedLength = 23
    T_xml = 25
    T_DateTime2 = 26
    T_DateTimeOffset = 27
End Enum
Public Enum COLOR
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E0E0
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
    �϶�ɫ = &HFFE0D9
    ����ģ��ɫ = &HC00000
    
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
    
End Enum

Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

'Public Const ָ���������� = "1,�ı�;2,��ֵ;3,����;4,�߼�"

'----------------------------------------------------------------------------------------------------------------------
'���Ͷ���

'�û���Ϣ
Public Type USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ģ��Ȩ�� As String
    ��λ���� As String
    �������� As String
    ���ݿ��û� As String
End Type

'ϵͳ������Ϣ
Public Type SYSPARAM_INFO
    ���ý��С��λ�� As String
    �շ�������Ŀƥ�� As String
    ����Ʊ�ݺų��� As Integer
    �շ�Ʊ�ݺų��� As Integer
    ���￨���볤�� As Integer
    ���￨��ĸǰ׺ As String
    ���￨������ʾ As Boolean
    ��Ŀ����ƥ�䷽ʽ As Integer '0-˫��;1-����
    ϵͳ�� As Long
    ϵͳ���� As String
    ��Ʒ���� As String
    ģ��� As Long
    ������ As String
    �շ�Ʊ�� As Integer
    ����Ʊ�� As Integer
    ����Ʊ���ϸ���� As Boolean
    �շ�Ʊ���ϸ���� As Boolean
    ����HIS���� As Byte
    ����RIS As Boolean
End Type

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'----------------------------------------------------------------------------------------------------------------------
'ȫ�ֱ�������
Public gcnOracle As ADODB.Connection                    '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gobjFSO As New Scripting.FileSystemObject        'FSO����
Public ParamInfo As SYSPARAM_INFO
Public glngUserId As Long                   '��ǰ�û�id
Public UserInfo As USER_INFO
Public gblnInsure As Boolean
Public gstrSQL As String
Public gblnShowInTaskBar As Boolean
Public gfrmMain As Object
Public glngTXTProc As Long                              '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngShareUseID As Long
Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjRichEPR As New cRichEPR          '�������Ĳ���
Public gobjPath As New clsCISPath           '�ٴ�·������
Public gobjEmr As Object                    '�°���Ӳ���
Public gobjJob As Object                    '�ٴ���������ZL9CISJOB
Public gobjXWHIS As Object     '�����ӿڲ���zl9XWInterface.clsHISInner
Public gobjPlugIn As Object    '�������
Public gobjLIS As Object     'Lis����
'��������
Public gstrPrivs As String
Public gstrDeptName As String
Public glngDeptId As Long
Public gstrDBUser As String
Public glngSys As Long
Public glngModul As Long
Public gstrSysName As String
Public gstrUserName As String
Public OldWindowProc As Long  ' Original window proc

'PDF��ӡ
Public gstrInputSeverName As String
Public gstrInputUser As String
Public gstrInputPwd As String
    
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

Public gclsPackage As New clsPackage
Public gstrMatchMethod      As String '���뷨���䷽ʽ
Public glngHIS����� As Long

'Private mclsUnzip As New clsUnZip

'----------------------------------------------------------------------------------------------------------------------
'ģ���������




'######################################################################################################################
'�����嵥

Public Function GetUserInfo() As Boolean
    '******************************************************************************************************************
    '����:��ȡ��½�û���Ϣ
    '����:
    '����:
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    
    UserInfo.�û��� = UserInfo.���ݿ��û�
    UserInfo.���� = UserInfo.���ݿ��û�
    
    Set rsTmp = zlDatabase.GetUserInfo
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�������� = IIf(IsNull(rsTmp!������), "", rsTmp!������)
        GetUserInfo = True
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '��ӡ����
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '��ӡ����,Ԥ������,�����Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "���ӡ�����粻�������ݣ������¼��ӣ�", vbInformation, ParamInfo.ϵͳ����
            Exit Function
        End If
        
        '���ô�ӡ��������
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("��ӡʱ��:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
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
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.ϵͳ��) / 100))
        
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '�˳�
    
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

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button            '������
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              'ͼ������
        If frmMain.cbsMain.count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '��ͼ��
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '״̬��
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function InitSysPara() As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim strTmp As String
        
    On Error GoTo errHand
    
    'Ʊ�ݺų���
    '------------------------------------------------------------------------------------------------------------------
    
    strTmp = zlDatabase.GetPara(20, ParamInfo.ϵͳ��)

    If strTmp <> "" Then
        If UBound(Split(strTmp, "|")) >= 4 Then ParamInfo.���￨���볤�� = Val(Split(strTmp, "|")(4))
    End If
    
    InitSysPara = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Public Function AdjustCodePostion(ByVal frmMain As Object, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    objTxt.Top = objTxtParent.Top + 45
    objTxt.Left = objTxtParent.Left + frmMain.TextWidth(objTxtParent.Text) + 60
    objTxt.Width = objTxtParent.Width - frmMain.TextWidth(objTxtParent.Text) - 120
    objTxt.BackColor = objTxtParent.BackColor
    
    AdjustCodePostion = True
    
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
    On Error GoTo errHand
    
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
    
errHand:
    
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
    
    On Error GoTo errHand
    
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
    
errHand:
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, ParamInfo.ϵͳ��, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************

    On Error GoTo ErrH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, ParamInfo.ϵͳ��, lngModual, blnSetup)

    Exit Function
    
ErrH:

End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "Select zlWBcode('" & strInput & "') from dual"
    Else
        strSQL = "Select zlSpellcode('" & strInput & "') from dual"
    End If
    On Error GoTo errHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        zlDatabase.OpenSQLRecord strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function GetApplyMode(ByVal StrText As String) As Byte
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    If CheckStrType(StrText, 1) And Left(ParamInfo.�շ�������Ŀƥ��, 1) = 1 Then
        '��ȫ���֣����������
            
        GetApplyMode = 1
        
    ElseIf CheckStrType(StrText, 2) And Left(ParamInfo.�շ�������Ŀƥ��, 2) = 1 Then
        '��ȫ��ĸ�����������
        
        GetApplyMode = 2
    Else
        GetApplyMode = 3
    End If
End Function

Public Function VsfInputIsCard(ByRef vsfInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��(�Ƿ�ﵽ���ų��ȣ��ڵ��ó������ж�),������ϵͳ���������Ƿ�������ʾ
    '������KeyAscii=��KeyPress�¼��е��õĲ���
    '******************************************************************************************************************
    
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, StrText As String
        
    'ˢ��ʱ����������ŵ�Ҫȡ��
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Function
    
    '����ǰ�������ʾ������(��δ��ʾ����)
    StrText = vsfInput.EditText
    If vsfInput.EditSelLength = Len(vsfInput.EditText) Then StrText = ""
    If KeyAscii = 8 Then
        If StrText <> "" Then StrText = Mid(StrText, 1, Len(StrText) - 1)
    Else
        StrText = StrText & Chr(KeyAscii)
    End If
        
    '�ж��Ƿ���ˢ��
    If IsNumeric(StrText) And IsNumeric(Left(StrText, 1)) Then  '�����������������ȫ���֣���Ϊ��ˢ��
        blnCard = True
    ElseIf KeyAscii > 32 Then
        sngNow = Timer
        If vsfInput.EditText = "" Or StrText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(StrText), "0.000") < 0.04 Then blnCard = True   '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    
'    'ˢ��ʱ�����Ƿ�������ʾ
'    If blnCard Then
'        vsfInput.PasswordChar = IIf(gobjComLib.zlDatabase.GetPara(12, lngSys) = "0", "", "*")
'    Else
'        vsfInput.PasswordChar = ""
'    End If
    
    VsfInputIsCard = blnCard
End Function



Public Sub WaitOpen(ByVal frmParent As Object, ByVal strTitle As String)
    frmPubWait.OpenWait frmParent, strTitle
End Sub

Public Sub WaitClose()
    frmPubWait.CloseWait
End Sub

Public Sub WaitInfo(ByVal strInfo As String)
    frmPubWait.WaitInfo = strInfo
End Sub

Public Sub SetMsfForeColor(ByRef msf As Object, ByVal lngRow As Long, ByVal lngColor As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intCol As Integer

    With msf

        .Row = lngRow
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellForeColor = lngColor
        Next

    End With
End Sub

Public Function SearchPrintData(ByVal objVsf As Object, ByRef objPrintVsf As Object, Optional strNotPrintCol As String = "") As Boolean
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
    
    If objPrintVsf.Cols = 0 Then Exit Function
    If strNotPrintCol <> "" Then
        lngNotPrintCols = UBound(Split(strNotPrintCol, ",")) + 1
        strNotPrintCol = "," & strNotPrintCol & ","
    End If
    
    objPrintVsf.Rows = objVsf.Rows
    objPrintVsf.FixedRows = objVsf.FixedRows
    
    objPrintVsf.Cols = 0
    lngPrintCol = -1
    For lngCol = 0 To objVsf.Cols - 1
        
        If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
            
            If InStr(strNotPrintCol, "," & lngCol & ",") = 0 Then
                
                lngPrintCol = lngPrintCol + 1
                
                objPrintVsf.Cols = lngPrintCol + 1
                
                objPrintVsf.ColWidth(lngPrintCol) = objVsf.ColWidth(lngCol)
                objPrintVsf.ColAlignmentFixed(lngPrintCol) = objVsf.ColAlignment(lngCol)
                If objVsf.ColDataType(lngCol) = flexDTBoolean Then
                    objPrintVsf.ColAlignment(lngPrintCol) = 4
                Else
                    objPrintVsf.ColAlignment(lngPrintCol) = objVsf.ColAlignment(lngCol)
                End If
            End If
        End If
    Next
    
    If objPrintVsf.Cols = 0 Then Exit Function
    
    For lngRow = 0 To objVsf.Rows - 1

        objPrintVsf.RowHeight(lngRow) = IIf(objVsf.RowHeight(lngRow) < objVsf.RowHeightMin, objVsf.RowHeightMin, objVsf.RowHeight(lngRow))
        lngPrintCol = -1
        For lngCol = 0 To objVsf.Cols - 1
            
            If objVsf.ColHidden(lngCol) = False And objVsf.TextMatrix(0, lngCol) <> "" Then
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
            End If
        Next
        Call SetMsfForeColor(objPrintVsf, lngRow, Val(objVsf.Cell(flexcpForeColor, lngRow, 1)))
    Next
    SearchPrintData = True
End Function


Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte
    '******************************************************************************************************************
    '���ܣ�������+�б�ṹ,Ӧ���ڱ��ؼ�
    '������
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '���أ�0:ȡ��ѡ��;1:ѡ��;2:�����ݷ���
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As PointAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If blnOneReturn Then
        If rsData.RecordCount = 1 Then
            Set rsResult = rsData
            ShowPubSelect = 1
            Exit Function
        End If
    End If
    
    Call ClientToScreen(obj.hWnd, objPoint)
    
    Select Case TypeName(obj)
    Case "TextBox", "CommandButton"
    
        lngX = objPoint.x * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = obj.Height + objPoint.y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        lngObjHeight = obj.Height
        
    Case Else
        lngX = objPoint.x * Screen.TwipsPerPixelX + obj.CellLeft
        lngY = objPoint.y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
        lngObjHeight = obj.CellHeight
    End Select
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function AnalyseAge(strOld As String, ByRef strAgeNumber As String, ByRef strAgeUnit As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '����:�����ݿ��б�������䰴���Ƶĸ�ʽ���ص�����
    
    Dim strTmp As Long
    
    If strOld = "��" Then Exit Function
    
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            strAgeNumber = strTmp
            strAgeUnit = "��"
        Else
            strAgeNumber = strOld
            strAgeUnit = ""
        End If
    ElseIf IsNumeric(strOld) Then
        strAgeNumber = strOld
        strAgeUnit = "��"
    Else
        strAgeNumber = strOld
        strAgeUnit = ""
    End If
    
    AnalyseAge = True
    
End Function


Public Function GetImageList(Optional ByVal intIconSize As Integer = 16) As ImageList
    Set GetImageList = frmPubResource.GetImageList(intIconSize)
End Function

Public Function CreateHelpMenu(cbsMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim objMenu As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
        
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & ParamInfo.��Ʒ����)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, ParamInfo.��Ʒ���� & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.��Ʒ���� & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With
    
    CreateHelpMenu = True
    
End Function
Public Function Get�°滤��(ByVal lng����ID, ByVal lng��ҳID As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
     On Error GoTo errHand
    strSQL = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����ϰ�����", lng����ID, lng��ҳID)
    If rsTemp.RecordCount > 0 Then
        Get�°滤�� = False
    Else
        Get�°滤�� = True
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


Public Function Have��������(ByVal lng����ID As Long, ByVal str���� As String) As Boolean
'���ܣ����ָ�������Ƿ����ָ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    On Error GoTo ErrH
    
    strSQL = "Select ����ID From ��������˵�� Where ����ID=[1] And ��������=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, str����)
    Have�������� = Not rsTmp.EOF
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetlngID(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
'���ܣ���鵱ǰ�������ڵĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select ��Ժ����ID,��Ժ����ID From ������ҳ  where ����ID =[1] and ��ҳID =[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID, lng��ҳID)
    If rsTmp.RecordCount = 1 Then
        If Val(rsTmp!��Ժ����ID) = 0 Then
            GetlngID = Val(rsTmp!��Ժ����ID)
        Else
            GetlngID = rsTmp!��Ժ����ID
        End If
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemField(ByVal strTable As String, ByVal lngID As Long, Optional ByVal strField As String) As Variant
'���ܣ���ȡָ����ָ���ֶ���Ϣ
'˵����δ����NULLֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrH
    
    If strField = "" Then
        strSQL = "Select * From " & strTable & " Where ID=[1]"
    Else
        strSQL = "Select " & strField & " From " & strTable & " Where ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISWork", lngID)
    If Not rsTmp.EOF Then
        If strField = "" Then
            Set GetItemField = rsTmp
        Else
            GetItemField = rsTmp.Fields(strField).Value
        End If
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function CheckAuditSql_IN(strSQL As String, Optional blnMsg As Boolean = False, Optional intSource As Integer = 0) As Boolean
'=���ܣ� ���SQL���书�� ͨ�� ���� �ж�����
'intSource ����Դ =0 �������ڱ�׼����ִ�У�=1������EMR��ִ��
'���أ�����ͨ������True
    Dim rsTemp          As ADODB.Recordset
    Dim zlCheck         As New clsCheck
    Dim strReturn       As String, strEMRSQL As String, strParm As String
    On Error GoTo ErrH
    If strSQL = "" Then CheckAuditSql_IN = True: Exit Function
    If strSQL = "" Then Exit Function
    strSQL = UCase(strSQL)
    If intSource = 1 Then
        'EMR��Ĳ�ѯ���ܱ�ת�ɴ�дִ�У���Щ���������Ҫ���ִ�Сд,ת�ɴ�дֻ���ڼ���Ƿ���ڸ���ɾ�����
        strEMRSQL = Replace(strSQL, "[MID]", "Hextoraw(:mid)")
        strEMRSQL = Replace(strEMRSQL, "[ALIDIN]", "Hextoraw(:alidin)")
    Else
        strSQL = Replace(strSQL, "[����ID]", "-1")
        strSQL = Replace(strSQL, "[��ҳID]", "-1")
        strSQL = "Select * From (" & strSQL & ") "
    End If
    
    If InStr(1, strSQL, "INSERT") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����д�����ݣ�", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, "UPDATE") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ����ܸ������ݣ�", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, "DELETE") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����ɾ�����ݣ�", vbCritical
        Exit Function
    ElseIf InStr(1, strSQL, ";") > 0 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����ʹ�á�;����", vbCritical
        Exit Function
    End If
    
    If intSource = 1 Then
        strParm = IIf(InStr(strEMRSQL, ":mid") = 0, "", "A^" & DbType.T_String & "^mid")
        If InStr(strEMRSQL, ":alidin") > 0 Then
            If InStr(strEMRSQL, ":mid") > 0 Then
                strParm = strParm & "|"
            End If
            strParm = strParm & "A^" & DbType.T_String & "^alidin"
        End If
        strReturn = gobjEmr.OpenSQLRecordset(strEMRSQL, strParm, rsTemp)
        If (strReturn <> "" And InStr(strReturn, "ORA-01403") = 0) Or rsTemp Is Nothing Then
            zlCheck.Msg_OK "������ݼ��ʧ��:" & vbCrLf & "��" & strReturn & "��" & vbCrLf & "������¼��������ݼ����䣡", vbExclamation
            Exit Function
        End If
    Else
        strSQL = "Select * From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", "mdlCISAudit")
        If rsTemp Is Nothing Then
            zlCheck.Msg_OK "�﷨���ʧ�ܣ�", vbCritical
            Exit Function
        Else
            If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]:ORA-01403") > 0 Then
                'û�ҵ��κ�����
            ElseIf InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") > 0 Then
                zlCheck.Msg_OK "������ݼ��ʧ��:" & vbCrLf & "��" & Mid(rsTemp.Fields(0), 23) & "��" & vbCrLf & "������¼��������ݼ����䣡", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If blnMsg Then zlCheck.Msg_OK "�﷨�����ɹ���"
    CheckAuditSql_IN = True
    Set zlCheck = Nothing
    Exit Function
ErrH:
    zlCheck.Msg_OK "�﷨�����ʧ�ܣ�" & vbCrLf & Err.Description, vbCritical
    '�ж�״̬ ���������������
    Err.Clear
    Set zlCheck = Nothing
End Function

'==============================================================================
'=���ܣ� ���SQL���书��
'==============================================================================
Public Function CheckAuditSql_OUT(strSQL As String, Optional lng����ID As Long = -1, Optional lng��ҳID As Long = -1) As String
    On Error GoTo ErrH
    strSQL = UCase(strSQL)
    strSQL = Replace(strSQL, "[����ID]", CStr(lng����ID))
    strSQL = Replace(strSQL, "[��ҳID]", CStr(lng��ҳID))
    
    CheckAuditSql_OUT = strSQL
        
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageID As Long) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2 And Nvl(���Ӵ�λ, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0) B"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ԺID", lngPatiID, lngPageID)
    If rsTemp Is Nothing Then Exit Function
    If NVL(rsTemp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTemp!ID
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEMR_MID_ALIDIN(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByRef strMid As String, ByRef strAlidin As String) As Boolean
    On Error GoTo errHandle
    Dim strReturn As String, strExtend_Tag As String, rsTemp As New ADODB.Recordset
    strExtend_Tag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtend_Tag = "" Then Exit Function
    gstrSQL = "Select Rawtohex(ID) As ID, Rawtohex(Master_Id) As Master_Id From Bz_Act_Log Where Extend_Tag = :extendtag"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtend_Tag & "^" & DbType.T_String & "^extendtag", rsTemp)
    If strReturn <> "" Then Exit Function
    strMid = rsTemp!Master_id
    strAlidin = rsTemp!ID
    
    GetEMR_MID_ALIDIN = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'==============================================================================
'=���ܣ� �����Ŀ ���ö��� �� �����������
'==============================================================================
Public Function AuditFileTran(strUsed As String, intSource As Integer) As String
    Dim strType     As String
    On Error GoTo ErrH
    Select Case strUsed
        Case "2" 'סԺ����
            strType = IIf(intSource = 0, "2", "02")
        Case "3" '������
            strType = IIf(intSource = 0, "4", "03")
        Case "4" '�����¼
            strType = "3"
        Case "6" 'ҽ������
            strType = "7"
        Case "7" '����֤��
            strType = IIf(intSource = 0, "5", "04")
        Case "8" '֪���ļ�
            strType = IIf(intSource = 0, "6", "05")
    End Select
    AuditFileTran = strType
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Function GetPPFS() As String
    On Error GoTo ErrH
    '����ƥ��
    GetPPFS = ""
    If Val(zlDatabase.GetPara("����ƥ��")) = 0 Then
        GetPPFS = "%"
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=����:ȥ���ַ����еĵ�����("'")(ConvertString)
'=��ڲ���:
'=1).sStr          ����:String
'=���ڲ���:��
'=����:ȥ���ַ���(sStr)�еĵ�����
'=����:2004-12-11
'=���:ŷ��
'=˵��:��SQL����в��ܴ�������
'==================================================================================================
Function ConvertString(ByVal sStr As String) As String
On Error GoTo ErrH
    ConvertString = Replace(sStr, "'", "")
    ConvertString = Replace(ConvertString, "��", "")
    ConvertString = Replace(ConvertString, "&", "")
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'/******************************************/
'=����:BigNote
'=����:�Ŵ�ע�ֶα༭��,�����ر༭����ı�
' ����:mStr   �Ѿ������ַ���
'      mTitle �༭���ڵı���
'=���:����
'=����:2002-04-03
'/******************************************/
Function Big_Note(mStr As String, mTitle As String, Optional bReadOnly As Boolean, Optional bSqlCheck As Boolean = False, Optional SqlSource As Integer = 0) As String
On Error GoTo ErrH
    With FrmNoteBox
        .intSource = SqlSource
        .SqlCheck = bSqlCheck
        .StrText = mStr
        .StrTile = mTitle
        .ReadOnly = bReadOnly
        .Show vbModal
        DoEvents
        Big_Note = IIf(bSqlCheck, .StrText, ConvertString(.StrText))
    End With
    Set FrmNoteBox = Nothing
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function RecordEprPrintInfo(ByVal bytMode As Byte, ByVal strRecordKey As String, ByVal lngNo As Long, Optional ByVal lngPatientKey As Long, Optional ByVal lngPatientPageKey As Long) As Boolean
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    If lngNo = 0 Then
        lngNo = 1
        strSQL = "Select Nvl(Max(��ӡ����),0)+1 As ��ӡ���� From ������ӡ��¼ Where ����id=[1] And ��ҳid=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", lngPatientKey, lngPatientPageKey)
        If rsTmp.BOF = False Then
            lngNo = rsTmp("��ӡ����").Value
        End If
    End If
    
    Select Case bytMode
    Case 1
        strSQL = "Select ����id,��ҳid,�������� From�����Ӳ�����¼ a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_������ӡ��¼_Insert(" & Val(rs("����id").Value) & "," & Val(rs("��ҳid").Value) & "," & lngNo & ",'" & rs("��������").Value & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    Case 2
        strSQL = "Zl_������ӡ��¼_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & strRecordKey & "','" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
    Case 3
        strSQL = "Select ���� From�������ļ��б� a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlCISAudit", Val(strRecordKey))
        If rs.BOF = False Then
            strSQL = "Zl_������ӡ��¼_Insert(" & lngPatientKey & "," & lngPatientPageKey & "," & lngNo & ",'" & rs("����").Value & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "mdlCISAudit")
        End If
    End Select
    
    RecordEprPrintInfo = True
    
End Function

'��ⳤ���Ƿ񳬹�����(�ֽ���)
Function ChkStrUniCode(mStr As String, mLen As Long) As String
    Dim strL        As String
On Error GoTo ErrH
    mStr = ConvertString(mStr)
    If mLen <= 0 Then
        ChkStrUniCode = mStr
        Exit Function
    Else
        strL = StrConv(mStr, vbFromUnicode)
        strL = LeftB(strL, mLen)
        ChkStrUniCode = StrConv(strL, vbUnicode)
    End If
    Exit Function
ErrH:
    Err.Clear
    ChkStrUniCode = ""
    Exit Function
End Function

Public Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid, lngNo As Long)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNo = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               'Ϊ��֧��zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(.FixedCols + i) = False
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
'                    .ColData
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
                End If
            End If
            .ColData(i) = Val(Split(arrHead(i), ",")(3)) '��������Ϊ����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Public Function zl_VsGrid_SaveToPara(ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, _
ByVal lngMoudel As Long, ByVal strParaName As String, Optional ByVal bln˽�� As Boolean = True, _
    Optional ByVal blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�������
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     lngMoudel-ģ���
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------

    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnǿ�ƻָ����� = False Then
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If

    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    zlDatabase.SetPara strParaName, strCol, glngSys, lngMoudel ', bln˽��
    zl_VsGrid_SaveToPara = True
End Function

Public Function zl_VsGrid_FromParaRestore(ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal lngMoudle As Long, _
    ByVal strParaName As String, Optional bln˽�� As Boolean = True, _
    Optional ByVal blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�Ӳ������лָ�����Ŀ��
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     lngMoudle-ģ���
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------

    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String

    If blnǿ�ƻָ����� = False Then
        If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If

    strParaValue = zlDatabase.GetPara(strParaName, glngSys, lngMoudle, "")
    If strParaValue = "" Then Exit Function
    
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...

    Err = 0: On Error GoTo errHand:

    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_VsGrid_FromParaRestore = True
    Exit Function
errHand:
End Function

Public Function GetCustomWhere(Optional ByVal lng����id As Long, _
                                    Optional ByVal str�Ա� As String, _
                                    Optional ByVal lng����ID As Long, _
                                    Optional ByVal int��ʼ���� As Integer, _
                                    Optional ByVal int�������� As Integer, _
                                    Optional ByVal str����״�� As String, _
                                    Optional ByVal strסԺ�� As String, _
                                    Optional ByVal str������ As String) As String
    '******************************************************************************************************************
    '���ܣ���ϲ������Ĳ�ѯ�Ļ�������
    '������
    '���أ����ؼǲ�ѯ����
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim strSQL As String
    
    'A ������Ϣ B������ҳ
    If lng����id > 0 Then strSQL = " And (Y1.����id,Y1.��ҳid) In (Select ����id,��ҳid From ������ϼ�¼ Where ����id=" & lng����id & ")"
    If str�Ա� <> "" Then strSQL = strSQL & " And X.�Ա�='" & str�Ա� & "'"
    If lng����ID > 0 Then strSQL = strSQL & " And Y1.��Ժ����id=" & lng����ID
    If str����״�� <> "" Then strSQL = strSQL & " And Y1.����״��='" & str����״�� & "'"
    If Val(strסԺ��) > 0 Then strSQL = strSQL & " And Y1.סԺ��='" & strסԺ�� & "'"
    If str������ <> "" Then strSQL = strSQL & " And Z.������='" & str������ & "'"
    
    If int��ʼ���� <> 0 Or int�������� <> 0 Then
'        strSQL = "Select * From (" & strSQL & ") Where ���� Between [5] And [6]"
         strSQL = strSQL & " And X.���� Between '" & int��ʼ���� & "' And '" & int�������� & "'"
    End If
    
    GetCustomWhere = strSQL
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Where����ʱ��(Optional strAlias As String) As String
    If strAlias = "" Then
        Where����ʱ�� = " (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) "
    Else
        Where����ʱ�� = " (" & strAlias & ".����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".����ʱ�� is null) "
    End If
End Function

Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡվ����������
    '���:blnAnd-�Ƿ���� And ���
    '����:str����-�������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-03-02 17:27:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strWhere As String
    Dim strAlia As String
    
'    If gbln����վ����� = False Then
        '��������
        zl_��ȡվ������ = "": Exit Function
'    End If
    
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
     zl_��ȡվ������ = strWhere
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo errHand:
    
    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
    If lng���� <> -1 Then
        lngCol = lng����
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
errHand:
End Sub

Public Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:������
    '����:vsGrid-�����е�������
    '     lngRow-��ǰ��
    '     blnBefor-��lngrow֮���֮��.true:֮��,false-֮��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT

    Err = 0: On Error GoTo errHand:

    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
errHand:
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '-------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

'������
Public Sub zl_VsGridBeforeSort(ByVal vsGrid As VSFlexGrid, ByRef Col As Long, ByRef Order As Integer, Optional strSpaceRowNotCheckCol As String = "")
    '-----------------------------------------------------------------------------------------------------------
    '����:��������(����ʱ,�������հ���)
    '���:strSpaceRowNotCheckCol-���������е���Щ��(��1,��2...)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-07-25 11:38:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngStartRow As Long, lngEndRow As Long, lngStartCol As Long, lngEndCol As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnAllowSelect As Boolean, blnAllowBigSel As Boolean
    Dim lngOldBackColor As Long
    
    If vsGrid.ExplorerBar > &H1000& Then Exit Sub
    '���浱ǰ��ѡ������
    vsGrid.GetSelection lngStartRow, lngStartCol, lngEndRow, lngEndCol
    vsGrid.Redraw = flexRDNone
    blnAllowBigSel = vsGrid.AllowBigSelection: blnAllowSelect = vsGrid.AllowSelection
    
    '������հ���
    With vsGrid
        For lngRow = .Rows - 1 To .FixedRows Step -1
            For lngCol = 0 To .Cols - 1
               If InStr(1, "," & strSpaceRowNotCheckCol & ",", "," & lngCol & ",") > 0 Then
               Else
                    If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then GoTo GoNext:
               End If
            Next
        Next
GoNext:
        If lngRow > .FixedRows Then
            
             .Select .FixedRows, Col, lngRow, Col
            .Sort = Order
        End If
        ' �ָ���ǰѡ�������
        .Select lngStartRow, lngStartCol, lngEndRow, lngEndCol
            
        .Redraw = flexRDDirect
    End With
    Order = 0
End Sub


Public Function zl_vsGrid_Para_Save(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption As String, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�Ʊ��� As Boolean = False, Optional blnHaveParaPrivs As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    If blnSaveToDataBase = False Then
        zl_vsGrid_Para_Save = True
        If blnǿ�Ʊ��� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
    End If
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    If blnSaveToDataBase Then
        zlDatabase.SetPara strKey, strCol, glngSys, lngModule, blnHaveParaPrivs
    Else
        Call SaveRegInFor(g˽��ģ��, strCaption, strKey, strCol)
    End If
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKey As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnSaveToDataBase-�Ƿ��������ݿ��б������(����������ݿ��б���,��ǿ�Ʊ���Ϊtrue,��������Ƿ�ʹ�ø��Ի������ȷ��)
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���ע���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        'ֻ���ڱ���ע����вŻᴦ����Ի�����
        zl_vsGrid_Para_Restore = True
        If blnǿ�ƻָ����� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g˽��ģ��, strCaption, strKey, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKey, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo errHand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
errHand:
End Function
 


'*********************************************************************************************************************
Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
errHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo errHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
errHand:
End Sub

Public Function GetVsGridBoolColVal(ByVal vsGrid As VSFlexGrid, lngRow As Long, lngCol As Long) As Boolean
    '------------------------------------------------------------------------------
    '����:��ȡbool�е�ֵ
    '����:�Ǹõ�Ԫ��Ϊtrue,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/28
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Err = 0: On Error GoTo errHand:
    With vsGrid
        strTemp = .TextMatrix(lngRow, lngCol)
    End With
    If UCase(strTemp) = UCase("True") Then
        GetVsGridBoolColVal = True: Exit Function
    End If
    GetVsGridBoolColVal = Val(strTemp) <> 0
    Exit Function
errHand:
End Function

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function


'################################################################################################################
'## ���ܣ�  ��������鷽���ļ�������XML�ĵ���
'##
'## ������  tvwThis     :   RTB�༭���ؼ�
'##         strFileName :   XML�ļ�����ȫ·����
'##
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Public Function ExportToXMLFile(ByRef tvwThis As TreeView, ByVal strFileName As String) As Boolean
    Dim i As Long, j As Long, k As Long
    Dim oDoc As DOMDocument             'xml�ĵ�
    Dim oRoot  As IXMLDOMElement        '���ڵ�
    Dim oNode As IXMLDOMNode            '���ڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode        '�ڵ�
    Dim oSubNode3 As IXMLDOMNode        '�ڵ�
 
    Dim strPath As String
    Dim strSolutionID As String         '����ID
    Dim strSQL As String                'SQL
    Dim rsTree As ADODB.Recordset       '�����¼��
    
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strSolutionID = Replace(tvwThis.SelectedItem.Key, "Root", "")
    
    'XML�ĵ�
    Set oDoc = New DOMDocument
    'ע��
    oDoc.appendChild oDoc.createComment(gstrSysName & "  " & _
        "����Ա:" & gstrUserName & "������:" & gstrDeptName & "��ʱ��:" & _
        Format(Now(), "YYYY��MM��DD��"))
    '���ڵ�
    Set oRoot = oDoc.createElement("SpotCheck")
    Set oDoc.documentElement = oRoot    '����Ϊ���ڵ�
    Call oRoot.setAttribute("SolutionName", tvwThis.SelectedItem.Text)
    Call oRoot.setAttribute("SolutionID", Replace(tvwThis.SelectedItem.Key, "Root", ""))
    
    strSQL = "SELECT /*+ rule */ id,�ϼ�ID,����ID,����,���� FROM ���������� Where ����ID=[1] START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID"
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "����������", strSolutionID)
    rsTree.Sort = "����"
    
    '����������
    Set oNode = CreateNode(1, oRoot, "Classify", NODE_ELEMENT, "")
    rsTree.MoveFirst
    Do Until rsTree.EOF
      '����ӽڵ�
       Set oSubNode1 = CreateNode(2, oNode, "Class", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "ID", , zlCommFun.NVL(rsTree!ID, 0)
            CreateNode 3, oSubNode1, "����ID", , zlCommFun.NVL(rsTree!����ID)
            CreateNode 3, oSubNode1, "�ϼ�ID", , zlCommFun.NVL(rsTree!�ϼ�ID)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
        rsTree.MoveNext
    Loop
    
    
    strSQL = "Select /*+ rule */ a.id,a.����id,a.����,a.����,a.����,a.��ֵ,a.����,a.˵��,a.�������,a.���ö���,a.�ļ�ID,a.���û���" & vbNewLine & _
            "  From �������Ŀ¼ a, ���������� b,������鷽�� C" & vbNewLine & _
            " Where a.����id = b.ID And b.����id = C.id And C.id =[1]"
    Set rsTree = zlDatabase.OpenSQLRecord(strSQL, "����������", strSolutionID)
    rsTree.Sort = "����id"
    
    '�������Ŀ¼
    Set oNode = CreateNode(1, oRoot, "Catalogue", NODE_ELEMENT, "")
    rsTree.MoveFirst
    Do Until rsTree.EOF
      '����ӽڵ�
        Set oSubNode1 = CreateNode(2, oNode, "Catalog", NODE_ELEMENT, "")
            CreateNode 3, oSubNode1, "ID", , zlCommFun.NVL(rsTree!ID, 0)
            CreateNode 3, oSubNode1, "����ID", , zlCommFun.NVL(rsTree!����id)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
            CreateNode 3, oSubNode1, "��ֵ", , zlCommFun.NVL(rsTree!��ֵ)
            CreateNode 3, oSubNode1, "����", , zlCommFun.NVL(rsTree!����)
            CreateNode 3, oSubNode1, "˵��", , zlCommFun.NVL(rsTree!˵��)
            CreateNode 3, oSubNode1, "�������", , zlCommFun.NVL(rsTree!�������)
            CreateNode 3, oSubNode1, "���ö���", , zlCommFun.NVL(rsTree!���ö���)
            CreateNode 3, oSubNode1, "�ļ�ID", , zlCommFun.NVL(rsTree!�ļ�ID)
            CreateNode 3, oSubNode1, "���û���", , zlCommFun.NVL(rsTree!���û���)
         rsTree.MoveNext
    Loop
 
    '�汾��Ϣ
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = oDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call oDoc.insertBefore(pi, oDoc.childNodes(0))
    'ֱ�ӱ�����ļ�����
    oDoc.Save strFileName
    
    Set oDoc = Nothing
    ExportToXMLFile = True
    Exit Function
LL:
    ExportToXMLFile = False
End Function


'################################################################################################################
'## ���ܣ�  ��XML�ļ����벡����鷽��
'##
'## ������  tvwThis     :   RTB�༭���ؼ�
'##         strFileName :   XML�ļ�����ȫ·����
'##         blnPrompt   :   �Ƿ���ʾ���븲�ǣ�Ĭ��ΪTrue
'##         blnForUndoRedo : �Ƿ�����Undo/Redo��Ĭ��ΪFalse
'##
'## ���أ�  ����ɹ�������Ture�����򷵻�False��
'################################################################################################################
Public Function ImportFromXMLFile(ByRef tvwThis As TreeView, _
    ByVal strFileName As String, _
    Optional blnPrompt As Boolean = True, _
    Optional blnForUndoRedo As Boolean = False) As Boolean
    
    Dim i As Long, j As Long, k As Long, lngSelStart As Long, lngSelEnd As Long
    
    Dim lKey As Long
    Dim oDoc As DOMDocument             'xml�ĵ�
    Dim oRoot  As IXMLDOMElement        '���ڵ�
    Dim oNode As IXMLDOMNode            '���ڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    
    Dim cllTemp As New Collection
    Dim strCurSolutionID As String      '��ǰ����ID
    Dim strSolutionName As String       'ԭ����������
    Dim strSolutionID As String         'ԭ������ID
    
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCode As String
    Dim strPath As String
    Dim strID As String
    Dim lngTypeID As Long
    Dim lngTypePrivID As Long
    Dim strTypeCode As String
    Dim strTypeName As String
    Dim lngProjectID As String
    Dim strParentCode As String
    
    
    Dim lngItemID  As Long
    Dim intItemTypeID As Integer
    Dim strItemCode As String
    Dim strItemName As String
    Dim strItemMnemonicCode As String
    Dim strItemDescription As String
    Dim strItemAudit As String
    Dim intItemUsed As Integer
    Dim intItemFileID As Integer
    Dim strItemLink As String
    Dim strPalValue As String
    Dim strNumValue As String
    
    On Error GoTo ErrH
    
    strCurSolutionID = Replace(tvwThis.SelectedItem.Key, "Root", "")
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    Set oDoc = New DOMDocument
    oDoc.Load strFileName
    '����������κ�Ԫ�أ����˳�
    If oDoc.documentElement Is Nothing Then
        Exit Function
    End If
    If blnPrompt Then
        If MsgBox("ע�⣺�����ļ���ԭ�����ݽ����ɻָ����Ƿ�������ǵ�ǰ�ļ���", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then
            Exit Function
        End If
    End If
    '��ȡ�ļ��ṹ
    Set oRoot = oDoc.selectSingleNode("SpotCheck")       'oRoot��Ϊ���ڵ�
    If oRoot Is Nothing Then MsgBox "��XML�ļ�������ȷ�Ĳ�����鷽���ļ���", vbInformation, gstrSysName: Exit Function
    
    '��ȡ������Ϣ
    On Error Resume Next
    strSolutionName = oRoot.getAttributeNode("SolutionName").Text
    strSolutionID = Val(oRoot.getAttributeNode("SolutionID").Text)
    Screen.MousePointer = vbHourglass
    
    
    '��ȡ���������� Classify:
    Set oNode = oRoot.selectSingleNode("Classify")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "ID", 0)
        If lKey > 0 Then
            If Val(GetNodeValue(oSubNode1, "�ϼ�ID", 0)) = 0 Then
                strID = "-1"
            Else
                strID = GetCllValue(cllTemp, Val(GetNodeValue(oSubNode1, "�ϼ�ID", 0)))
            End If
            
            strSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������", strID)
            If rsTemp.RecordCount = 1 Then
                    strParentCode = "" & rsTemp!����
            Else
                    strParentCode = ""
            End If
            
            If Val(GetNodeValue(oSubNode1, "�ϼ�ID", 0)) = 0 Then
               lngTypePrivID = 0
               strSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID is null"
               Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������", strID)
            Else
               lngTypePrivID = strID
               strSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID = [1]"
               Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������", strID)
            End If
            
            strCode = ""
            If rsTemp.RecordCount = 1 Then
                strCode = rsTemp!����
                strCode = IncStr(strCode)
            End If
            If strCode = "" Then
                strTypeCode = strParentCode & "01"
            Else
                strTypeCode = strCode
            End If
            strTypeName = GetNodeValue(oSubNode1, "����", "")
            lngProjectID = Val(strCurSolutionID)
            lngTypeID = zlDatabase.GetNextId("����������")
            
            AddArray cllTemp, lngTypeID & ";" & lKey
            
            strSQL = "Zl_����������_Insert (" + CStr(lngTypeID) & "," & IIf(lngTypePrivID = 0, "NULL", CStr(lngTypePrivID)) + "," + "'" + strTypeCode + "'" + "," + "'" + strTypeName + "'," & CStr(0) & "," & lngProjectID & ")"
            zlDatabase.ExecuteProcedure strSQL, "����������"
        End If
    Next


    '�������Ŀ¼
    Set oNode = oRoot.selectSingleNode("Catalogue")
    For Each oSubNode1 In oNode.childNodes
        lKey = GetNodeValue(oSubNode1, "ID", 0)
        If lKey > 0 Then
            strID = GetCllValue(cllTemp, Val(GetNodeValue(oSubNode1, "����ID", 0)))
            
            lngItemID = zlDatabase.GetNextId("�������Ŀ¼")
            intItemTypeID = Val(strID)
    
            strItemName = GetNodeValue(oSubNode1, "����", "")
            strItemMnemonicCode = GetNodeValue(oSubNode1, "����", "")
            strItemDescription = GetNodeValue(oSubNode1, "˵��", "")
            strItemAudit = GetNodeValue(oSubNode1, "�������", "")
            intItemUsed = Val(GetNodeValue(oSubNode1, "���ö���", 0))
            intItemFileID = Val(GetNodeValue(oSubNode1, "�ļ�ID", 0))
            strItemLink = GetNodeValue(oSubNode1, "���û���", "")
            strPalValue = GetNodeValue(oSubNode1, "����", "")
            strNumValue = GetNodeValue(oSubNode1, "��ֵ", "")
            
            strSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����������", strID)
   
            If rsTemp.RecordCount = 1 Then
                 strSQL = "select nvl(Max(����),0) from �������Ŀ¼ where ����id=[1] and ���� like [2] || '%'"
                 strItemCode = IncStr(zlDatabase.OpenSQLRecord(strSQL, "�������Ŀ¼", intItemTypeID, rsTemp!����).Fields(0))
                
                 If strItemCode = "1" Then
                     strItemCode = rsTemp!���� & Format(strItemCode, "0000")
                 End If
                 strItemCode = InsertNewCode(strItemCode)
            
                strSQL = "Zl_�������Ŀ¼_Insert (" + CStr(lngItemID) + "," + IIf(intItemTypeID = 0, "NULL", CStr(intItemTypeID)) + "," + "'" + strItemCode + "'" + "," + "'" + strItemName + "','" & strItemMnemonicCode & "','" & strItemDescription & "','" & strItemAudit & "'," & intItemUsed & ",'" & intItemFileID & "','" & strItemLink & "'," & Val(strPalValue) & " ," & Val(strNumValue) & ")"
                zlDatabase.ExecuteProcedure strSQL, "�������Ŀ¼"
            End If
        End If
    Next

    Screen.MousePointer = 0
    ImportFromXMLFile = True
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  ��ȡһ���ڵ��ֵ
'##
'## ������  CurNode         :   ��ǰ�ڵ����
'##         SubNodeName     :   �ӽڵ�����
'##         DefaultValue    :   Ĭ��ֵ
'################################################################################################################
Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then 'ֻ�лس����޻��з�
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then 'ֻ�л��з��޻س���
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## ���ܣ�  ����һ��XML�ڵ㲢��ֵ
'##
'## ������  TabNumber   :   �������������ʾ�ж��ٸ�Tab�Ʊ���������Ķ���
'##         Parent      :   ���ڵ�
'##         Node_Type   :   �ڵ����ͣ�Ŀǰ֧�� NODE_ELEMENT ��NODE_CDATA_SECTION ��NODE_COMMENT ��NODE_ATTRIBUTE�ȣ�
'##         Node_Name   :   �ڵ�����
'##         Node_Value  :   �ڵ��ı�
'################################################################################################################
Private Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '�ַ�����ֵ���ã���Ӱ�����ݣ���ֻӰ���Ķ����۶�
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '�����ı��ڵ�
    '�����½ڵ�
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '�����ı�ֵ
    New_Node.Text = Node_Value
    '��ӵ����ڵ�
    Parent.appendChild New_Node
    '���ĩβ�س�����Ӱ�����ݣ���ֻӰ���Ķ����۶�
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '�����ı��ڵ�
    Set CreateNode = New_Node
End Function


''''################################################################################################################
''''## ���ܣ�  ������ж����ID��
''''################################################################################################################
'''Public Sub ClearAllIDs()
'''    Dim i As Long, j As Long
'''    For i = 1 To Me.Compends.count
'''        Me.Compends(i).ID = 0
'''    Next
'''    For i = 1 To Me.Pictures.count
'''        Me.Pictures(i).ID = 0
'''    Next
'''    For i = 1 To Me.elements.count
'''        Me.elements(i).ID = 0
'''    Next
'''    For i = 1 To Me.Signs.count
'''        Me.Signs(i).ID = 0
'''    Next
'''    For i = 1 To Me.Diagnosises.count
'''        Me.Diagnosises(i).ID = 0
'''    Next
'''    For i = 1 To Me.Tables.count
'''        Me.Tables(i).ID = 0
'''        For j = 1 To Me.Tables(i).Cells.count
'''            Me.Tables(i).Cells(j).ID = 0
'''        Next
'''        For j = 1 To Me.Tables(i).elements.count
'''            Me.Tables(i).elements(j).ID = 0
'''        Next
'''        For j = 1 To Me.Tables(i).Pictures.count
'''            Me.Tables(i).Pictures(j).ID = 0
'''        Next
'''    Next
'''End Sub

Private Function GetCllValue(ByVal cllTmp As Collection, ByVal strKey As String) As String
'��ȡ�������ж�Ӧ����ID
    Dim lngNum As Long
    If cllTmp.count > 0 Then
        For lngNum = 1 To cllTmp.count
            If InStrRev(cllTmp.Item(lngNum), ";" & strKey) > 0 Then
                GetCllValue = Replace(cllTmp.Item(lngNum), ";" & strKey, "")
                Exit Function
            End If
        Next
    End If
    
End Function

'========================================================================================
'=���� �� ����ʱ�������ı����Ƿ��Ѵ���
'========================================================================================
Private Function InsertNewCode(strInCode) As String
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "select 1 from �������Ŀ¼ where ���� = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������Ŀ¼", strInCode)
    If rsTemp.RecordCount = 0 Then InsertNewCode = strInCode: Exit Function
    strInCode = IncStr(strInCode)
    InsertNewCode = InsertNewCode(strInCode)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowPubSelectTest(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte

    On Error GoTo errHand

       ShowPubSelectTest = ShowPubSelect(frmParent, obj, bytStyle, strLvw, strSavePath, strDescrible, rsData, rsResult, lngCX, lngCY, blnMuliSel, strInitKey, strFilterControl, blnOneReturn)

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CreateXWHIS(Optional ByVal blnMsg As Boolean) As Boolean
'���ܣ��ж� RIS�ӿڲ���(zl9XWInterface.clsHISInner) �Ƿ���ڣ�������
'������blnMsg������ʧ��ʱ�Ƿ���ʾ

    If Not ParamInfo.����RIS Then Exit Function
    If Not gobjXWHIS Is Nothing Then CreateXWHIS = True: Exit Function
    
    On Error Resume Next
    Set gobjXWHIS = GetObject(, "zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    On Error Resume Next
    If gobjXWHIS Is Nothing Then Set gobjXWHIS = CreateObject("zl9XWInterface.clsHISInner")
    Err.Clear: On Error GoTo 0
    
    If gobjXWHIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    CreateXWHIS = True
End Function

Public Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True: Exit Function
    
    On Error Resume Next
    Set gobjPlugIn = GetObject("", "zlPlugIn.clsPlugIn")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjPlugIn Is Nothing Then Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
    
    If Not gobjPlugIn Is Nothing Then
        Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngMod)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
        CreatePlugInOK = True
    End If
    
End Function

Public Function InitObjLis(Optional ByVal blnMsg As Boolean) As Boolean
'�ж�����°�LIS����Ϊ�վͳ�ʼ��
    Dim strErr As String
    
    If gobjLIS Is Nothing Then
        On Error Resume Next
        Set gobjLIS = GetObject(, "zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
    
        On Error Resume Next
        If gobjLIS Is Nothing Then Set gobjLIS = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        Err.Clear: On Error GoTo 0
        
        If Not gobjLIS Is Nothing Then
            If gobjLIS.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If blnMsg Then MsgBox "LIS������ʼ������" & vbCrLf & strErr, vbInformation, gstrSysName
                Set gobjLIS = Nothing
                Exit Function
            End If
        End If
    End If
    InitObjLis = True
End Function

Public Function CreateCISJOB() As Boolean
'���ܣ��ж� �ٴ�����վ����(ZL9CISJOB.CLSCISJob) �Ƿ����

    If Not gobjJob Is Nothing Then CreateCISJOB = True: Exit Function
    On Error Resume Next
    Set gobjJob = GetObject(, "ZL9CLSCISJOB.CLSCISJOB")
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    If gobjJob Is Nothing Then Set gobjJob = DynamicCreate("ZL9CISJOB.CLSCISJob", "�ٴ�����վ", False)
    Err.Clear: On Error GoTo 0
    If gobjJob Is Nothing Then Exit Function
    CreateCISJOB = True
End Function
