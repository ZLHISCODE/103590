Attribute VB_Name = "mdlDrugPurchase"
Option Explicit

Public Const GSTR_MESSAGE = "��ʾ��Ϣ"

Public gstrUser As String, gstrUserNameNew As String
Public glngUserID As Long, glngDeptID As Long
Public gbytЧ�� As Byte

Public gcnOutside As New ADODB.Connection           '�ⲿ���ݿ�����
Public gcnOracle As ADODB.Connection
Public gblnSetupFinish As Boolean
Public gobjComLib As Object

Public Const GSTR_SYSNAME = "�ɹ����ݽ����ӿ�"
Public Const GSTR_REGEDIT_PATH = "����ģ��\DrugPurchaseDBServer"
Public Const MSTR_SERVER = "localhost"
Public Const MSTR_DBNAME = "GuoYaoDB"
Public Const MSTR_USER = "sa"
Public Const MSTR_PASSWORD = ""

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

Public Enum enm_Pop_File
     File = 1
     FilePrintSet = 181
     FilePreview = 102
     FilePrint = 103
     FileExit = 191
     Edit = 3
     EditAdd = 3212
     EditDel = 3213
     EditEdit = 23
     EditIgnore = 3214
     EditProcess = 3104
     EditCurrChoose = 301
     EditCurrCancel = 302
     EditChooChoose = 303
     EditChooCancel = 304
     EditAllChoose = 305
     EditAllCancel = 306
     View = 4
     ViewRefresh = 791
     ViewFindTitle = 411
     ViewFindEdit = 412
     ViewFindButton = 413
     ViewTools = 420
     ViewToolsButton = 421
     ViewToolsLabel = 422
     ViewToolsIcon = 423
     ViewStatebar = 430
     Import = 5
     ImportTitle = 51
     ImportControl = 52
     Help = 9
     HelpHelp = 901
     HelpWeb = 902
     HelpWebhome = 9021
     HelpWebBBS = 9022
     HelpWebFeelback = 9023
     HelpAbout = 903
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Function MSSQLServerOpen(ByVal strServerName As String, ByVal strDBName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ����MS SQL Server ���ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    If Len(Trim(strUserName)) = 0 Then
        MSSQLServerOpen = False
        MsgBox "�������������ݿ���Ϣ��", vbInformation, GSTR_MESSAGE
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .ConnectionTimeout = 5
        .Open "Driver={SQL Server};Server=" & strServerName & ";Database=" & strDBName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Or Err.Number = -2147467259 Then
                MsgBox "�м����ݿ�����ʧ�ܣ�", vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            MSSQLServerOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    MSSQLServerOpen = True
    Exit Function
    
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    MSSQLServerOpen = False
    Err = 0
End Function


Public Function OraDataOpenTest(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Then
                MsgBox Mid(strError, InStr(1, strError, "[SQL Server]"), Len(strError)), vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            OraDataOpenTest = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpenTest = True
    Exit Function
    
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    OraDataOpenTest = False
    Err = 0
End Function

Public Function StringEnDeCodecn(strSource As String, MA) As String
'�ú���ֻ���������𵽼�������
'����Ϊ��Դ�ļ�������
    On Error GoTo ErrEnDeCode
    Dim X As Single, i As Integer
    Dim CHARNUM As Long, RANDOMINTEGER As Integer
    Dim SINGLECHAR As String * 1
    Dim strTmp As String
    
    If MA < 0 Then
        MA = MA * (-1)
    End If
    
    X = Rnd(-MA)
    For i = 1 To Len(strSource) Step 1                 'ȡ���ֽ�����
        SINGLECHAR = Mid(strSource, i, 1)
        CHARNUM = Asc(SINGLECHAR)
g:
        RANDOMINTEGER = Int(127 * Rnd)
        If RANDOMINTEGER < 30 Or RANDOMINTEGER > 100 Then GoTo g
        CHARNUM = CHARNUM Xor RANDOMINTEGER
        strTmp = strTmp & Chr(CHARNUM)
    Next i
    StringEnDeCodecn = strTmp
    Exit Function

ErrEnDeCode:
    StringEnDeCodecn = ""
    MsgBox Err.Number & "\" & Err.Description
End Function

Public Function GetUserNameInfo() As Boolean
'��ȡ�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            glngUserID = IIf(IsNull(!Id), 0, !Id)
            glngDeptID = IIf(IsNull(!����id), 0, !����id)
            gstrUserNameNew = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            GetUserNameInfo = True
        Else
            glngUserID = 0
            glngDeptID = 0
            gstrUserNameNew = "" '��ǰ�û�����
        End If
    End With
    rsTmp.Close

    strSQL = "Select ������, ����ֵ, ȱʡֵ From Zlparameters Where ϵͳ = [1] And Nvl(˽��, 0) = 0 And ģ�� Is Null and ������=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ȡϵͳ����", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbytЧ�� = IIf(IsNull(rsTmp!����ֵ), rsTmp!ȱʡֵ, rsTmp!����ֵ)
        Else
            gbytЧ�� = 0
        End If
    End With
    
End Function

Public Sub SelText(ByVal ctlVal As Control)
    If TypeOf ctlVal Is TextBox Then
        ctlVal.SelStart = 0
        ctlVal.SelLength = Len(ctlVal.Text)
    End If
End Sub

Public Sub InitCommandBars(ByVal cmbVal As CommandBars)
    cmbVal.VisualTheme = xtpThemeOffice2003
    With cmbVal.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cmbVal.EnableCustomization False
    cmbVal.Icons = frmPublic.imgPublic.Icons
End Sub

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

Public Sub ProviderSelecter(frmParam As Form, ByVal objParam As Object, ByVal blnClick As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strParam As String
    Dim blnCancel As Boolean
    Dim vRect As RECT

    vRect = GetControlRect(objParam.hwnd)
    If blnClick = False Then
        strParam = "%" & UCase(Trim(objParam.Text)) & "%"
        strSQL = "SELECT id, ����, ����, ����  " _
                & "FROM ��Ӧ�� a " _
                & "Where TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
                & "  and (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0)" _
                & "  and (a.���� like [1] or a.���� like [1] or a.���� like [1]) " _
                & "order by a.����"
    Else
        strSQL = "SELECT id, ����, ����, ���� " _
                & "FROM ��Ӧ�� a " _
                & "Where TO_CHAR(a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " _
                & "  and (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0)" _
                & "order by a.����"
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.ShowSQLSelect(frmParam, strSQL, 0, "��Ӧ��", False, "", "" _
              , False, False, True, vRect.Left, vRect.Top, objParam.Height, blnCancel, False, False, strParam)

    If Not rsTmp Is Nothing Then
        objParam.Text = rsTmp!����
        rsTmp.Close
    End If
    
End Sub


Public Sub InitVSF(ByVal vsfVal As VSFlexGrid, blnVal As Boolean)
'�л���ʾVSF�ı�����
    Dim strCols As String
    Dim arrCols As Variant
    Dim i As Single
    
    With vsfVal
        .Rows = 1
        .ColWidth(0) = 130 * 2                               '��һ�п�
        .ColWidth(1) = 130
        .FixedCols = 2                                       '�̶�ǰ����
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns               '����ʱ�ɵ���Columns���
        .AllowSelection = True                               '�൥Ԫѡ����ƿ���
        .SelectionMode = flexSelectionListBox                '�൥Ԫѡ�����
        .ExplorerBar = flexExSortShow
        .BackColorSel = &HC0E0FF
        .BackColorAlternate = .BackColor      '&H80000003
        '.BackColorBkg = vbWhite
    End With
    
    If blnVal Then
        '��Ʊ����
        strCols = "||ѡ��,choose,440|H_��Ӧ��ID,providerid,800|��Ӧ��,provider,1500|H_ҩƷID,id,650|�ƻ�����,plan_code,1000" & _
                  "|ҩƷ����,name,2000|ҩƷ���,spec,1500|��Ʊ����,ivqty,850,r|PDA��������,pdaqty,1100,r|����������,chkqty,1000,r" & _
                  "|��������,qty,850,r|ҩ�ⵥλ,unit,800|������,Accepter,800|������,price,1000,r" & _
                  "|��Ʊ���,iamount,1200,r|��Ʊ��,invoice,1000|��Ʊ����,idate,1000|������,producer,1000" & _
                  "|����,lot_no,1000|��������,pdate,1000|Ч��,avail_date,1000|H_DetailID,detail_id,0" & _
                  "|H_�ѵ���,imported,600|��Ϣ,mess,2000"
    Else
        '�ƻ�����
        strCols = "||ѡ��,choose,440|H_�ƻ�ID,planid,0|�ƻ�����,planno,1000|���,xh,500|H_��Ӧ��ID,providerid,800|��Ӧ��,provider,1500" & _
                  "|H_ҩƷID,id,650|ҩƷ����,name,2000|ҩƷ���,spec,1500|�ƻ�����,qty,850,r|ҩ�ⵥλ,unit,800|����,price,1000,r" & _
                  "|������,producer,1000|H_ҩ��ID,wh_id,600|ҩ��,wh,1000|H_ҩ��ID,dh_id,0|ҩ��,dh,1000|��������,edate,1000" & _
                  "|�������,cdate,1000|H_�ѵ���,imported,600|��ע,remark,3000|��Ϣ,mess,2000"
    End If
    arrCols = Split(strCols, "|")
    With vsfVal
        .Clear
        .Cols = UBound(arrCols) + 1
        For i = LBound(arrCols) To UBound(arrCols)
            If arrCols(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                .TextMatrix(0, i) = Split(arrCols(i), ",")(0)
                .ColKey(i) = Split(arrCols(i), ",")(1)
                .ColWidth(i) = Split(arrCols(i), ",")(2)
                'H_Ϊ������
                If Mid(Split(arrCols(i), ",")(0), 1, 2) = "H_" Then
                    .ColHidden(i) = True
                Else
                    .ColHidden(i) = False
                    If UBound(Split(arrCols(i), ",")) > 2 Then
                        .ColAlignment(i) = flexAlignRightCenter
                    End If
                End If
            End If
        Next
        .ColDataType(.ColIndex("choose")) = flexDTBoolean    '����ΪCheck�ؼ�
    End With
    
End Sub

Public Function CheckProvider(ByVal lngProviderID As Long) As String
'��˹�Ӧ��ID
    Dim rsTmp As New ADODB.Recordset
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord("select ���� from ��Ӧ�� where ����ʱ��>to_date('2999/12/31','yyyy/mm/dd') And id=[1]", "��˹�Ӧ��ID", lngProviderID)
'    Set rsTmp = zlDatabase.OpenSQLRecord("Select (Select ���� From ��Ӧ�� Where ����ʱ��=to_date('3000/1/1','yyyy/mm/dd') And Id=[1]) ����, " & _
'                                         "(Select count(1) From ҩƷ�б굥λ Where ��λId=[1] And ҩƷID=[2] and ����ʱ��=to_date('3000/1/1','yyyy/mm/dd')) �Ƿ��б� " & _
'                                         "from dual ", "��˹�Ӧ��ID", intProviderID, lngDrugID)
    If rsTmp.RecordCount = 1 Then
'        CheckProvider = rsTmp!���� & "|" & rsTmp!�Ƿ��б�
        CheckProvider = rsTmp!����
    End If
    rsTmp.Close
End Function

Public Sub DataLoading(ByVal vsfVal As VSFlexGrid, ByVal rsVal As ADODB.Recordset, ByVal bytTab As Byte, Optional ByVal bytMarked As Byte = 0)
    Dim i As Integer, j As Integer
    Dim strName As String, strSpec As String, strUnit As String, strProvider As String
    Dim blnGet As Boolean
    Dim dblCost As Double

    On Error GoTo errHandle
    With vsfVal
        .Rows = 1
        .Rows = rsVal.RecordCount + 1
        If rsVal.RecordCount > 0 Then rsVal.MoveFirst
        For i = 1 To rsVal.RecordCount
            strName = "": strSpec = "": strUnit = ""
            
            Err = 0: On Error Resume Next
            blnGet = GetMedicalInfo(IIf(IsNull(rsVal!ҩƷid), -1, rsVal!ҩƷid), strName, strSpec, strUnit)
            If Err <> 0 Then
                .TextMatrix(i, .ColIndex("mess")) = "��ҩƷID��" & Err.Description & "[�ⲿ���ݿ�]��"
            End If
            Err = 0: On Error GoTo errHandle
            
            .TextMatrix(i, 1) = i   '���
            '�ɹ�������
            If bytTab = 0 Then
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!Id), "", rsVal!Id)
                .TextMatrix(i, .ColIndex("planno")) = IIf(IsNull(rsVal!no), "", rsVal!no)
                .TextMatrix(i, .ColIndex("xh")) = IIf(IsNull(rsVal!���), "", rsVal!���)
                .TextMatrix(i, .ColIndex("providerid")) = IIf(IsNull(rsVal!��Ӧ��id), "", rsVal!��Ӧ��id)
                .TextMatrix(i, .ColIndex("provider")) = IIf(IsNull(rsVal!�ϴι�Ӧ��), "", rsVal!�ϴι�Ӧ��)
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!�ϴ�������), "", rsVal!�ϴ�������)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!ҩƷid), "", rsVal!ҩƷid)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rsVal!����), "", rsVal!����)
                .TextMatrix(i, .ColIndex("spec")) = IIf(IsNull(rsVal!���), "", rsVal!���)
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(rsVal!ҩ�ⵥλ), "", rsVal!ҩ�ⵥλ)
                .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(rsVal!�ƻ�����), "0", rsVal!�ƻ�����)
                '.ColFormat(.ColIndex("qty")) = "#0"
                .TextMatrix(i, .ColIndex("price")) = IIf(IsNull(rsVal!����), "0", rsVal!����)
                .ColFormat(.ColIndex("price")) = "#0.0000"
                .TextMatrix(i, .ColIndex("wh_id")) = IIf(IsNull(rsVal!ҩ��id), "", rsVal!ҩ��id)
                .TextMatrix(i, .ColIndex("wh")) = IIf(IsNull(rsVal!ҩ��), "", rsVal!ҩ��)
                .TextMatrix(i, .ColIndex("dh_id")) = IIf(IsNull(rsVal!ҩ��id), "", rsVal!ҩ��id)
                .TextMatrix(i, .ColIndex("dh")) = IIf(IsNull(rsVal!ҩ��), "", rsVal!ҩ��)
                .TextMatrix(i, .ColIndex("edate")) = IIf(IsNull(rsVal!��������), "", rsVal!��������)
                .ColFormat(.ColIndex("edate")) = "yyyy-mm-dd"
                .TextMatrix(i, .ColIndex("cdate")) = IIf(IsNull(rsVal!�������), "", rsVal!�������)
                .ColFormat(.ColIndex("cdate")) = "yyyy-mm-dd"
                
                If rsVal!�Ƿ��ϴ� = 1 Then
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .TextMatrix(i, .ColIndex("imported")) = "1,0"
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbBlue
                ElseIf .TextMatrix(i, .ColIndex("qty")) > 0 Then
                    .TextMatrix(i, .ColIndex("choose")) = 1
                    .TextMatrix(i, .ColIndex("imported")) = "1,1"
                Else
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .TextMatrix(i, .ColIndex("imported")) = "1,0"
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbRed
                End If
            '��ⵥ����
            Else
                .TextMatrix(i, .ColIndex("providerid")) = IIf(IsNull(rsVal!��Ӧ��id), "-1", rsVal!��Ӧ��id)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!ҩƷid), "-1", rsVal!ҩƷid)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(strName), "", strName)
                .TextMatrix(i, .ColIndex("spec")) = IIf(IsNull(strSpec), "", strSpec)
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(strUnit), "", strUnit)
                .TextMatrix(i, .ColIndex("ivqty")) = IIf(IsNull(rsVal!��Ʊ����), "0", rsVal!��Ʊ����)
                .ColFormat(.ColIndex("ivqty")) = "#0.000"
                .TextMatrix(i, .ColIndex("pdaqty")) = IIf(IsNull(rsVal!PDA��������), "0", rsVal!PDA��������)
                .ColFormat(.ColIndex("pdaqty")) = "#0.000"
                .TextMatrix(i, .ColIndex("chkqty")) = IIf(IsNull(rsVal!����������), "0", rsVal!����������)
                .ColFormat(.ColIndex("chkqty")) = "#0.000"
                '���ѱ�ǵ�����
                If bytMarked = 1 Then
                    .TextMatrix(i, .ColIndex("qty")) = 0 'IIf(IsNull(rsVal!��������), "0", rsVal!��������)
                Else
                    'If Val(.TextMatrix(i, .ColIndex("pdaqty"))) <= 0 Then
                    '    .TextMatrix(i, .ColIndex("qty")) = 0
                    'Else
                        .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(rsVal!PDA��������), "0", rsVal!PDA��������)
                    'End If
                End If
                .ColFormat(.ColIndex("qty")) = "#0.000"
                .ColDataType(.ColIndex("qty")) = flexDTLong
                '.TextMatrix(i, .ColIndex("price")) = IIf(IsNull(rsVal!������), "0", rsVal!������)
                
                'ҩ�ⵥλ�ĳɱ���
                Err.Clear: On Error Resume Next
                dblCost = GetCostPrice(IIf(IsNull(rsVal!ҩƷid), "-1", rsVal!ҩƷid))
                If Err <> 0 Then
                    .TextMatrix(i, .ColIndex("mess")) = "��ҩƷID��" & Err.Description & "[�ⲿ���ݿ�]��"
                    dblCost = 0
                End If
                .TextMatrix(i, .ColIndex("price")) = dblCost
                Err = 0: On Error GoTo errHandle
                
                .ColFormat(.ColIndex("price")) = "#0.0000"
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!������), "", rsVal!������)
                .TextMatrix(i, .ColIndex("lot_no")) = IIf(IsNull(rsVal!����), "", rsVal!����)
                .ColDataType(.ColIndex("lot_no")) = flexDTString
                .TextMatrix(i, .ColIndex("pdate")) = IIf(IsNull(rsVal!��������), "", rsVal!��������)
                .ColFormat(.ColIndex("pdate")) = "yyyy-mm-dd"
                .TextMatrix(i, .ColIndex("avail_date")) = IIf(IsNull(rsVal!Ч��), "", rsVal!Ч��)
                .TextMatrix(i, .ColIndex("invoice")) = IIf(IsNull(rsVal!��Ʊ��), "", rsVal!��Ʊ��)
                .ColDataType(.ColIndex("invoice")) = flexDTString
                .TextMatrix(i, .ColIndex("idate")) = IIf(IsNull(rsVal!��Ʊ����), "", rsVal!��Ʊ����)
                .ColFormat(.ColIndex("idate")) = "yyyy-mm-dd"
                '.TextMatrix(i, .ColIndex("iamount")) = IIf(IsNull(rsVal!��Ʊ���), "0", rsVal!��Ʊ���)
                .TextMatrix(i, .ColIndex("iamount")) = dblCost * IIf(IsNull(rsVal!��Ʊ����), "0", rsVal!��Ʊ����)
                .ColFormat(.ColIndex("iamount")) = "#0.0000"
                .TextMatrix(i, .ColIndex("detail_id")) = IIf(IsNull(rsVal!detail_id), "0", rsVal!detail_id)
                .TextMatrix(i, .ColIndex("plan_code")) = IIf(IsNull(rsVal!�ƻ�����), "", rsVal!�ƻ�����)
                .TextMatrix(i, .ColIndex("Accepter")) = IIf(IsNull(rsVal!������), "", rsVal!������)
                
                '��鹩Ӧ��ID
                If Trim(.TextMatrix(i, .ColIndex("providerid"))) = "" Then
                    .TextMatrix(i, .ColIndex("mess")) = "����Ӧ��ID��δ��д[�ⲿ���ݿ�]��"
                    strProvider = ""
                Else
                    Err = 0: On Error Resume Next
                    strProvider = CheckProvider(Val(.TextMatrix(i, .ColIndex("providerid"))))
                    If Err <> 0 Then
                        .TextMatrix(i, .ColIndex("mess")) = "����Ӧ��ID��" & Err.Description & "[�ⲿ���ݿ�]��"
                        strProvider = ""
                    End If
                    Err = 0: On Error GoTo errHandle
                End If
                .TextMatrix(i, .ColIndex("provider")) = strProvider
                
                'If .TextMatrix(i, .ColIndex("providerid")) = "" Or .TextMatrix(i, .ColIndex("providerid")) = "-1" Then
                If strProvider = "" Then
                    'Ϊ�����޸��ṩ��Ϣ
                    .TextMatrix(i, .ColIndex("provider")) = "��Ӧ��ID��"
                    .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
'                ElseIf Mid(strProvider, InStr(strProvider, "|") + 1, Len(strProvider)) = "0" Or Mid(strProvider, InStr(strProvider, "|") + 1, Len(strProvider)) = "" Then
'                    .TextMatrix(i, .ColIndex("provider")) = "δ�����б굥λ"
'                    .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                Else
                    'Choose�ɵ���޸�
                    If Len(Trim(strName)) = 0 Then
                        .TextMatrix(i, .ColIndex("provider")) = "ҩƷID��/��HIS����Ӧ"
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf bytMarked = 1 And Val(.TextMatrix(i, .ColIndex("pdaqty"))) <> 0 Then
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf Val(.TextMatrix(i, .ColIndex("ivqty"))) <= 0 Then 'Or Val(.TextMatrix(i, .ColIndex("qty"))) <= 0 Then
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",0"
                    ElseIf Val(.TextMatrix(i, .ColIndex("ivqty"))) > Val(.TextMatrix(i, .ColIndex("qty"))) And Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                        '��Ʊ����������������
                        If bytMarked = 0 And Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                            .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                        Else
                            .TextMatrix(i, .ColIndex("provider")) = "��Ʊ����������������"
                            .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                        End If
                    Else
                        .TextMatrix(i, .ColIndex("imported")) = IIf(IsNull(rsVal!imported), "0", rsVal!imported) & ",1"
                    End If
                End If
                
                If Mid(.TextMatrix(i, .ColIndex("imported")), 3, 1) = "1" Then
                    '��Checkѡ
                    If Val(.TextMatrix(i, .ColIndex("qty"))) > 0 Then
                        .TextMatrix(i, .ColIndex("choose")) = IIf(Len(Trim(strName)) = 0, 0, 1)
                    End If
                Else
                    '����Checkѡ
                    .TextMatrix(i, .ColIndex("choose")) = 0
                    .Cell(flexcpForeColor, i, 3, i, .Cols - 1) = vbRed
                End If
                If Left(.TextMatrix(i, .ColIndex("imported")), 1) = "1" Then
                    .Cell(flexcpForeColor, i, 3, i, .ColIndex("mess")) = vbBlue
                    .TextMatrix(i, .ColIndex("choose")) = 0
                End If
                
            End If
            rsVal.MoveNext
            .ColWidth(1) = IIf(.Rows > 0, Len(Trim(Str(.Rows))) * 130 + 70, 200)
        Next
        
    End With
    Exit Sub

errHandle:
    MsgBox "�ؼ�װ������ʱ�쳣��", vbInformation, GSTR_MESSAGE
End Sub

Public Function GetMedicalInfo(ByVal intID As Long, strName As String, strSpec As String, strUnit As String) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String

    strSQL = "select b.����,b.���,a.ҩ�ⵥλ from ҩƷ��� a, �շ���ĿĿ¼ b where a.ҩƷid=[1] and b.id=[1]  and a.ҩƷid=b.id and rownum=1 "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "", intID)
    On Error GoTo ErrHand
    If rsTmp.RecordCount = 1 Then
        strName = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        strSpec = IIf(IsNull(rsTmp!���), "", rsTmp!���)
        strUnit = IIf(IsNull(rsTmp!ҩ�ⵥλ), "", rsTmp!ҩ�ⵥλ)
    End If
    rsTmp.Close
    GetMedicalInfo = True
    Exit Function
ErrHand:
    GetMedicalInfo = False
End Function

Public Sub RefreshTVWProvider(ByVal tvwVal As TreeView, ByVal vsfVal As VSFlexGrid)
    Dim i As Long, j As Long
    Dim blnFind As Boolean
    Dim nodTmp As Node
    Dim rsTmp As New ADODB.Recordset
    
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "ID", adInteger, 18, adFldIsNullable
        .Fields.Append "Name", adVarChar, 50, adFldIsNullable
        .Open
    End With
    With tvwVal
        .Nodes.Clear
        .Nodes.Add , , "Root", "ȫ��"
        .Nodes(1).Checked = True
        .Nodes(1).Expanded = True
    End With
    '���浽RecordSet��
    With vsfVal
        For i = 1 To .Rows - 1
            blnFind = False
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If rsTmp!Id = Val(.TextMatrix(i, .ColIndex("providerid"))) Then
                    blnFind = True
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If blnFind = False And .TextMatrix(i, .ColIndex("imported")) <> "0,0" Then
                rsTmp.AddNew
                rsTmp!Id = Val(.TextMatrix(i, .ColIndex("providerid")))
                rsTmp!Name = .TextMatrix(i, .ColIndex("provider"))
            End If
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("imported")) = "0,0" Then
                rsTmp.AddNew
                rsTmp!Id = -1
                rsTmp!Name = "�����¼"
                Exit For
            End If
        Next
    End With
    '����
    With rsTmp
        .Sort = "Name"
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Set nodTmp = tvwVal.Nodes.Add("Root", tvwChild, "K" & !Id, !Name)
            nodTmp.Tag = !Id
            nodTmp.Checked = True
            .MoveNext
        Loop
    End With
End Sub

Public Function CheckRecord(ByVal vsfVal As VSFlexGrid) As Boolean
    Dim i As Integer
    With vsfVal
        For i = 1 To .Rows - 1
            If .RowHidden(i) = False And Val(.TextMatrix(i, .ColIndex("choose"))) <> 0 Then
                CheckRecord = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function GetCostPrice(ByVal lngID As Long) As Double
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    strTmp = "Select nvl(�ɱ���,0) * ҩ���װ �ɱ��� From ҩƷ��� Where ҩƷid=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strTmp, "ȡҩƷ���ĳɱ���", lngID)
    If Not rsTmp.EOF Then
        GetCostPrice = rsTmp!�ɱ���
    End If
    rsTmp.Close
End Function
