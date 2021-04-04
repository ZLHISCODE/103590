Attribute VB_Name = "mdlMain"
Option Explicit
'**************************
'       OEM����
'
'����    B0AEC9FA
'ҽҵ    D2BDD2B5
'����    CDD0C6D5
'����    D6D0C8ED
'��̩  BDF0BFB5CCA9
'ҽԺ    D2BDD4BA
'����    B1A6D0C5
'**************************
Public gobjFile         As New FileSystemObject
Public gcolOwnerConn    As New Collection
Public gblnInIDE        As Boolean  '�Ƿ�Դ��������
Public gstrOracleVer    As String 'Oracle�汾
Public gstrOracleBigVer    As String 'Oracle��汾
Public gblnTrace        As Boolean '�Ƿ����ø�����־
Public gblnTestUpgrade  As Boolean '���Խű�����
Public gblnClose11g     As Boolean '�ر�11g�ظ����ݲ������������
Public glngInterval     As Long '�����ļ��
Public glngAtuoErr      As Long '�Զ�������
Public gobjRegister     As Object 'ע����Ȩ����
'----------------------------------------------------------------------------------------
'--��װ�ű�ִ����ر�������
Private mobjLog As TextStream

Public Type SQL_DEFINE
    varName As String
    varValue As String
End Type

Public Type listFile
  Filename      As String
  FileVision    As String
  FileEditDate  As String
  FileMD5       As String
End Type

'���������б�-zq 20101213
Public Type UpdateList
  uFile() As listFile
End Type
'���ӷ�ʽ
Public Enum enuProvider
    MSODBC = 0
    OraOLEDB = 1
    OriginalConnection = 9
End Enum

Public Const G_STR_USERS As String = "'SYS','SYSTEM','SCOTT','OUTLN','DBSNMP','MTSSYS','MDSYS','ORDSYS','ORDPLUGINS','CTXSYS','ZLTOOLS','XDB','WMSYS','TSMSYS','SYSMAN','SI_INFORMTN_SCHEMA','OLAPSYS','MGMT_VIEW','MDDATA','EXFSYS','DMSYS','DIP','ANONYMOUS'"
'���˺�:����'XDB','WMSYS','TSMSYS','SYSMAN','SI_INFORMTN_SCHEMA','OLAPSYS','MGMT_VIEW','MDDATA','EXFSYS','DMSYS','DIP','ANONYMOUS'

Public gcnOracle As New ADODB.Connection     '��OraOLEDB��ʽ�򿪵Ĺ������ݿ�����
Public gcnOldOra As New ADODB.Connection    '��ODBC��ʽ�򿪵����ӣ�����ִ�нű�����OraOLEDB��ʽ�����洢���̻ᷢ��ִ�гɹ����ǹ���û�б����µ�����
Public gcnSystem As ADODB.Connection        'SYSTEM�û�����
Public gcnTools As ADODB.Connection        'ZLTools�û�����

Public gstrUserName As String               '�û���
Public gstrPassword As String               '�û������ݿ�����
Public gstrLoginPwd As String               '�û���¼ʱ���������
Public gstrLoginUserName As String          '��Ȩ�û���¼���û���
Public gstrLoginUserPwd As String           '��Ȩ�û���¼�����ݿ�����

Public gstrToolsPwd As String                  '�����ߵ�����
Public gstrSysUser As String                     'SYS�û���
Public gstrSysPwd As String                     'SYS����
Public gstrServer As String                       '��������

Public gobjFunction As Object
Public gobjReport As Object
Public gobjUsrProc As Object

Public gstrAppsoft As String                'APPSOFT·��

Public gstrSysName As String                'ϵͳ����
Public gstrProductTitle As String
Public gstrUltimatetag  As String          '�콢�棬רҵ���ʶ
Public gstrProductName As String
Public gstrDevelopers As String
Public gstrSustainer As String
Public gstrWebSustainer As String
Public gstrWebURL As String
Public gstrWebEmail As String
Public gstrע���� As String                 '�õ�ע����

Public gstrSQL    As String                 'ͨ�õ�SQL������
Public gblnCreate As Boolean                '�Ƿ��Ѿ�����������
Public gblnDBA As Boolean                   '�Ƿ�DBA
Public gblnRac As Boolean                 '�Ƿ���Rac����
Public gintInstID As Integer                  'Rac�����µ�ǰ��¼ʵ����
Public gblnOwner As Boolean                 '�Ƿ�������
Public gfrmActive As Form                   '��ǰ����Ӵ���
Public gcbsMain As CommandBars
Public gdtStart As Long
Public gstrHaveProg As String               '�������DBA��ϵͳ�����ߵ�¼�����ж��Ƿ��й����ߵ�Ȩ��
Public gblnSystemUser As Boolean            '�ж��Ƿ�Ϊϵͳ�����ߵ�¼

Public gstrComputerName As String           '��¼��ǰ�ͻ�������
Public glngSysNo As Long                    '��Ҫ���ڵ�ϵͳ��¼ʱ����¼��ǰ��¼��ϵͳ�ı��

Public gclsBase As New clsBase
Public glngTXTProc As Long


Public Const FindUserWidth = 4845   '���Ҵ��ڴ�С
Public Const FindUserHeight = 5595

Private mstrHasZltables As String  '�Ƿ���zltables���ű�
Private mstrBigTable As String   '���
Private mstrMiddleTable As String '�б�
Private mstrMiddleTableRows As String

Private Enum REGRoot
    HKEY_CLASSES_ROOT = &H80000000 '��¼Windows����ϵͳ�����������ļ��ĸ�ʽ�͹�����Ϣ����Ҫ��¼��ͬ�ļ����ļ�����׺����֮��Ӧ��Ӧ�ó��������Ӽ��ɷ�Ϊ���࣬һ�����Ѿ�ע��ĸ����ļ�����չ���������Ӽ�ǰ�涼��һ������������һ���Ǹ����ļ������й���Ϣ��
    HKEY_CURRENT_USER = &H80000001 '�˸��������˵�ǰ��¼�û����û������ļ���Ϣ����Щ��Ϣ��֤��ͬ���û���¼�����ʱ��ʹ���Լ��ĸ��Ի����ã������Լ������ǽֽ���Լ����ռ��䡢�Լ��İ�ȫ����Ȩ�޵ȡ�
    HKEY_LOCAL_MACHINE = &H80000002 '�˸��������˵�ǰ��������������ݣ���������װ��Ӳ���Լ���������á���Щ��Ϣ��Ϊ���е��û���¼ϵͳ����ġ���������ע��������Ӵ�Ҳ������Ҫ�ĸ�����
    HKEY_USERS = &H80000003 '�˸�������Ĭ���û�����Ϣ��Default�Ӽ�����������ǰ��¼�û�����Ϣ��
    HKEY_PERFORMANCE_DATA = &H80000004 '��Windows NT/2000/XPע�������Ȼû��HKEY_DYN_DATA����������ȴ������һ����Ϊ��HKEY_ PERFOR MANCE_DATA����������ϵͳ�еĶ�̬��Ϣ���Ǵ���ڴ��Ӽ��С�ϵͳ�Դ���ע���༭���޷������˼�
    HKEY_CURRENT_CONFIG = &H80000005  '�˸���ʵ������HKEY_LOCAL_MACHINE�е�һ���֣����д�ŵ��Ǽ������ǰ���ã�����ʾ������ӡ���������������Ϣ�ȡ������Ӽ���HKEY_LOCAL_ MACHINE\ Config\0001��֧�µ�������ȫһ����
    HKEY_DYN_DATA = &H80000006 '�˸����б���ÿ��ϵͳ����ʱ��������ϵͳ���ú͵�ǰ������Ϣ���������ֻ������Windows 98�С�
End Enum

'ע�����������
Private Enum REGValueType
    REG_NONE = 0                       ' No value type
    REG_SZ = 1 'Unicode���ս��ַ���
    REG_EXPAND_SZ = 2 'Unicode���ս��ַ���
    REG_BINARY = 3 '��������ֵ
    REG_DWORD = 4 '32-bit ����
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7 ' ��������ֵ��
End Enum
Private Declare Function RegQueryValueEx_ValueType Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_BINARY Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal lpcbData As Long) As Long
Private Declare Function RegSetValueEx_Long Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx_BINARY Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
'��ʾ��������
'SHwnd:���봰�ھ��(��Ϊ��������)
'htmName:��ӳ��CHM�е�htm�ļ�����

    Dim Path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\"
    If CBool(PathIsDirectory(Path)) = False Then GoTo ShowHelpErr
    strSave = "zl9server.CHM"
    Path = Trim(Path & strSave)
    If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, Path, &H0, htmName & ".htm")
    ShowHelp = True
    Exit Function

ShowHelpErr:
    err.Clear
End Function

Public Sub Main()
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    Dim strRegErr As String, strFile As String, strTmp As String
    Dim strLog As String
    Dim strCommand As String
    Dim blnAnalysis As Boolean
    Dim objclsCiph As clsCipher
    
    Call InitCommonControls

    'Ϊ��ʵ��ע�����ܣ���ȫ�ֱ������г�ʼ��
    gblnCreate = False
    gblnDBA = False
    gblnOwner = False
    gblnRac = False
    gintInstID = 0
    Set gfrmActive = Nothing
    Set gobjFunction = Nothing
    Set gobjReport = Nothing
    Set gobjUsrProc = Nothing
    
    gblnInIDE = gclsBase.InDesign
    gblnTrace = gblnInIDE
    '�Ƿ����ù����߸���
    gblnTrace = gblnTrace Or Val(GetSetting("ZLSOFT", "����ģ��\������������", "���ø���", "0")) = 1
    glngInterval = Val(GetSetting("ZLSOFT", "����ģ��\������������", "�Զ���ǨƵ��", "0"))
    glngAtuoErr = Val(GetSetting("ZLSOFT", "����ģ��\������������", "�������Դ���", "0"))
    gblnTestUpgrade = Val(GetSetting("ZLSOFT", "����ģ��\������������", "���Խű�����", "0")) = 1
    gblnClose11g = Val(GetSetting("ZLSOFT", "����ģ��\������������", "�ر�11G������", "0")) = 1
    gstrComputerName = GetMyCompterName
    If glngInterval < 100 Then glngInterval = 100
    If gblnTrace Then
        strLog = GetLogPath(LT_������־)
        If strLog <> "" Then
            Set gobjLog = gobjFile.CreateTextFile(strLog, True)
        End If
    End If
    gdtStart = Timer
    '��ȡShell�ַ���
    strCommand = CStr(Command())

    '�ж��Ƿ���ַ��������˼��ܣ��������ˣ�����н���
    If InStr(strCommand, "EncryptedLoginToken:") = 1 Then
        strCommand = Mid(strCommand, Len("EncryptedLoginToken:") + 1)
        Set objclsCiph = New clsCipher
        strCommand = objclsCiph.Decipher(MSTR_DBLINK_KEY, strCommand)
    End If
    '��ȡϵͳ��ţ�����ϵͳ�����Ϣ��ȡ��
    glngSysNo = GetSysNo(strCommand)
    If InStr(strCommand, "=") <= 0 Then
        frmSplash.ShowSplash
    End If
    Do
        If (Timer - gdtStart) > 1 Then Exit Do
        DoEvents
    Loop
    
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    err.Clear: On Error GoTo 0
    If gobjRegister Is Nothing Then
        MsgBox "����zlRegister��������ʧ�ܡ������ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        End
    End If
    
    '��鲿����MD5ֵ(����ģʽApp.Path�ǵ�ǰԴ�빤�̵�λ�ã����Բ����)
    If Not gblnInIDE Then
        strFile = App.Path & "\PUBLIC\zlRegister.Dll"
        strTmp = Md5_File_Calc(strFile)
        If strTmp <> "7F8912644328C37023F6839CDB4E7425" Then
            '10.35.90:13653ED7AF4144CAADB4CD5BF790C731
            '10.35.90SP1:F1335F5042068CF291B8141418775FD8
            MsgBox "��֤ע����Ȩ����ʧ��,�����ļ�" & strFile & "�İ汾�Ƿ�������ߵİ汾ƥ�䡣", vbExclamation, gstrSysName
            End
        End If
    End If
    

    '��ȡ��ǰ��¼ϵͳ��ţ�����ԭ�ַ����еĹ���ϵͳ��ŵĲ����޳���
    If InStr(strCommand, "=") > 0 Or glngSysNo <> -1 Then
        If Not frmUserLogin.Docmd(strCommand, blnAnalysis) Then
            If blnAnalysis = True Then  '��ʾ�Ե�һ�ַ�ʽ�����ɹ������ǵ�¼ʧ��
                '��Ϊ��ϵͳ��¼�ҵ�¼ʧ�ܣ������ṩ�ֹ�����
                If glngSysNo = -1 Then
                    frmUserLogin.Visible = True
                Else
                    Unload frmUserLogin
                    Set gcnOracle = Nothing
                End If
            Else  '��ʾ�Ե�һ�ַ�ʽ����ʧ�ܣ��ֳ���ʹ�õڶ��ַ�ʽ����
                frmUserLogin.ShowMe strCommand
                If glngSysNo <> -1 Then
                    Unload frmUserLogin
                    Set gcnOracle = Nothing
                End If
            End If
        End If
    Else
        frmUserLogin.ShowMe strCommand
    End If
    
    If InStr(strCommand, "=") <= 0 Then
        Unload frmSplash
    End If
    
    If gcnOracle.State = adStateOpen Then
        SaveSetting "ZLSOFT", "����ȫ��", "����·��", App.Path & "\" & App.EXEName & ".exe"
        If gblnCreate = False Then
            '�д��������ߣ����д���
            MsgBox "�״�����ϵͳ����Ҫ���ȴ��������ߡ�", vbExclamation, "��ʾ"
            frmSvrCreate.Show 1
        Else
            '��ODBC��ʽ�򿪵����ӣ�����ִ�нű�����OraOLEDB��ʽ�����洢���̻ᷢ��ִ�гɹ����ǹ���û�б����µ�����
            Set gcnOldOra = gobjRegister.ReGetConnection(MSODBC, strRegErr)
            If strRegErr <> "" Then
                MsgBox strRegErr, vbQuestion, "����"
                gcnOracle.Close
                End
            End If
            Call SetSQLTrace(gstrServer, gstrUserName, gcnOldOra)
            Select Case gobjRegister.zlRegInfo("��Ȩ����")
                Case "1"
                    '��ʽ
                    SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", ""
                Case "2"
                    '����
                    SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
                Case "3"
                    '����
                    SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
            End Select
            frmMDIMain.Show
        End If
    End If
End Sub

Private Function GetSysNo(ByRef strCmd As String) As Long
    '��ȡ��ǰ��¼ϵͳ��ţ�����ԭ�ַ����еĹ���ϵͳ��ŵĲ����޳���
    '��Ϊ�����¼����ϵͳ���Ϊ-1
    Dim ArrCommand() As String
    Dim strCommand As String
    Dim i As Long
    
    ArrCommand = Split(strCmd, " ")
    For i = LBound(ArrCommand) To UBound(ArrCommand)
        If UCase(ArrCommand(i)) Like "SYS=*" Then
            GetSysNo = Val(Split(ArrCommand(i), "=")(1))
        Else
            strCommand = strCommand & " " & ArrCommand(i)
        End If
    Next
    strCmd = Trim(strCommand)
    If GetSysNo = 0 Then GetSysNo = -1
End Function

Public Sub SelAll(objTxt As Control)
'���ܣ����ı���ĵ��ı�ѡ��
    If TypeName(objTxt) = "TextBox" Or TypeName(objTxt) = "ComboBox" Then
        If Trim(objTxt.Text) = "" Then Exit Sub
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Sub Getע����()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errH
    rsTemp.CursorLocation = adUseClient
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_reginfo", "��Ȩ֤��")
    gstrע���� = ""
    Do Until rsTemp.EOF
        gstrע���� = gstrע���� & IIf(IsNull(rsTemp!����), "", rsTemp!����)
        rsTemp.MoveNext
    Loop
    Exit Sub
errH:
    gstrע���� = ""
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo errH
    '���ܵ���OpenSQLRecord,��ΪOpenSQLRecordҲʹ���˸÷���
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    CurrentDate = rsTemp.Fields(0).value
    rsTemp.Close
    Exit Function
errH:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
    CurrentDate = 0
    err = 0
End Function


'��PictureBoxģ���3Dƽ�水ť
'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub DeleteAllLog(FrmObj As Form, BlnRunTimeLog As Boolean)
    Dim strRemarks As String
    Dim strNote As String
    
    '��֤��ݲ��������˵��
    If Not CheckAuditStatus(frmMDIMain.gstrLastModule, "ɾ��", strRemarks) Then Exit Sub
    Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_All_Log", IIf(BlnRunTimeLog, 1, 0))
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ��", "ɾ��������־", strRemarks)
    Call FrmObj.RefreshData
    Exit Sub
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Sub

Public Sub DeleteCurLog(FrmObj As Form, BlnRunTimeLog As Boolean)
    Dim LngDelete As Long, ItemThis As ListItem
    Dim lng�Ự�� As Long, str����վ As String, str�û��� As String, str������ As String
    Dim str�������� As String
    Dim dateʱ�� As Date, lng���� As Long, lng������� As Long
    Dim strRemarks As String
    Dim strNote As String
    
    If MsgBox("��ȷ��Ҫɾ����ѡ�����־��¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '��֤��ݲ��������˵��
    If Not CheckAuditStatus(frmMDIMain.gstrLastModule, "ɾ��", strRemarks) Then Exit Sub
    On Error Resume Next
    err = 0
    
    gcnOracle.BeginTrans
    For LngDelete = 1 To FrmObj.LvwList.ListItems.Count
        If FrmObj.LvwList.ListItems(LngDelete).Selected Then
            Set ItemThis = FrmObj.LvwList.ListItems(LngDelete)
            If BlnRunTimeLog Then
                lng�Ự�� = Val(ItemThis.Tag)
                str����վ = ItemThis.SubItems(1)
                str�û��� = ItemThis.SubItems(2)
                str������ = ItemThis.SubItems(3)
                str�������� = ItemThis.SubItems(4)
                dateʱ�� = CDate(ItemThis.SubItems(5))
                Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_Diarylog", _
                                         lng�Ự��, str�û���, str����վ, str������, str��������, dateʱ��)
                
            Else
                lng�Ự�� = Val(ItemThis.Tag)
                str����վ = ItemThis.SubItems(1)
                str�û��� = ItemThis.SubItems(2)
                lng���� = Val(IIf(ItemThis = "�洢���̴���", 1, IIf(ItemThis = "������������", 2, 3)))
                lng������� = Val(ItemThis.SubItems(4))
                dateʱ�� = CDate(ItemThis.SubItems(3))
                Call OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Delete_Errorlog", _
                                         lng�Ự��, str�û���, str����վ, lng����, lng�������, dateʱ��)
            End If
        End If
    Next
    
    If err <> 0 Then
        MsgBox "ɾ��ʱ��������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    gcnOracle.CommitTrans
    With FrmObj
    If BlnRunTimeLog Then
        If .Cbo����վ.Text <> "" Then strNote = ",����վ:" & .Cbo����վ.Text
        If .Cbo�û���.Text <> "" Then strNote = strNote & ",�û���:" & .Cbo�û���.Text
        If .Txt������.Text <> "" Then strNote = strNote & ",������:" & .Txt������.Text
        If .txt��������.Text <> "" Then strNote = strNote & ",ģ����" & .txt��������.Text
        strNote = strNote & ",��ʼʱ��:" & Format(.dtpDateStart, "yyyy-MM-dd") & ",��ֹʱ��:" & Format(.dtpDateEnd, "yyyy-MM-dd")
    Else
        If .Cbo����վ.Text <> "" Then strNote = ",����վ:" & .Cbo����վ.Text
        If .Cbo�û���.Text <> "" Then strNote = strNote & ",�û���:" & .Cbo�û���.Text
        strNote = strNote & ",��������:" & .Cbo��������.Text & ",��ʼʱ��:" & Format(.dtpDateStart, "yyyy-MM-dd") & ",��ֹʱ��:" & Format(.dtpDateEnd, "yyyy-MM-dd")
    End If
    End With
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ��", "ɾ������Ϊ��" & Mid(strNote, 2) & "����������־", strRemarks)
    Call FrmObj.RefreshData
End Sub

Public Function GetFileLineCount(ByVal txtStream As TextStream) As Long
    Do Until txtStream.AtEndOfStream
        txtStream.ReadLine
    Loop
    
    GetFileLineCount = txtStream.Line
End Function

Public Function CopyMenu(cnLink As ADODB.Connection, ByVal lngOldSys As Long, ByVal lngNewSys As Long) As Boolean
    
    On Error GoTo errHandle
    Call OpenCursor(cnLink, "ZLTOOLS.B_Popedom.Copy_menu", lngOldSys, lngNewSys)
    '���µ�������
    Call AdjustNameSequece("zltools.zlMenus", cnLink)
    CopyMenu = True
    
    Exit Function
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Function

Public Function CopyReport(cnLink As ADODB.Connection, ByVal lngOldSys As Long, ByVal lngNewSys As Long) As Boolean
    Call OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Copy_report", lngOldSys, lngNewSys)
    
    '���µ�������
    Call AdjustNameSequece("zltools.zlRPTGroups", cnLink)
    Call AdjustNameSequece("zltools.zlReports", cnLink)
    Call AdjustNameSequece("zltools.zlRPTDatas", cnLink)
    Call AdjustNameSequece("zltools.zlRPTItems", cnLink)
    
    CopyReport = True
End Function

Public Function GetOwnerName(lngSys As Long, cnLink As ADODB.Connection) As String
    Dim rsReturn As New ADODB.Recordset
    
    Set rsReturn = OpenCursor(cnLink, "ZLTOOLS.B_Public.Get_Owner_name", lngSys)
    If rsReturn.RecordCount > 0 Then
        GetOwnerName = IIf(IsNull(rsReturn.Fields(0)), "", rsReturn.Fields(0))
    Else
        GetOwnerName = ""
    End If
    
End Function

Public Sub AdjustSequence(ByVal str������ As String, cnOwner As ADODB.Connection, Optional ByVal lngSys As Long)
    Dim rsTable As ADODB.Recordset, strTable As String
    
    Set rsTable = GetSequence(str������, cnOwner)
    Do Until rsTable.EOF
        strTable = rsTable!Owner & "." & rsTable!Table_Name
        Call AdjustNameSequece(strTable, cnOwner, rsTable!Column_Name)
        DoEvents
        rsTable.MoveNext
    Loop
    If lngSys \ 100 = 1 Then
        Call Adjust����ID(cnOwner)
    End If
End Sub

Public Function GetSequence(ByVal str������ As String, cnOwner As ADODB.Connection, Optional blnCurrUser As Boolean) As ADODB.Recordset
    Dim rsSeq As New ADODB.Recordset
    Dim strSql As String
    
    '10.26���°�װ��ϵͳû����ͼ"���˷��ü�¼"
    If blnCurrUser Then
        strSql = "Select User, Decode(s.Sequence_Name, '���˷��ü�¼_ID', '���˷��ü�¼', s.Table_Name) as Table_Name, c.Column_Name, s.Sequence_Name" & vbNewLine & _
                "From User_Tab_Columns C," & vbNewLine & _
                "     (Select Sequence_Name, Decode(Sequence_Name, '���˷��ü�¼_ID', 'סԺ���ü�¼', Table_Name) Table_Name,Column_Name" & vbNewLine & _
                "       From (Select Sequence_Name, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                "                     Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name" & vbNewLine & _
                "              From User_Sequences)) S" & vbNewLine & _
                "Where c.Table_Name = s.Table_Name And c.Column_Name = s.Column_Name" & vbNewLine & _
                "Order By s.Table_Name"

    Else
        If str������ = "" Then
            strSql = " Where Sequence_Owner In (Select 'ZLTOOLS' From Dual Union Select ������ From Zlsystems)"
        Else
            strSql = " Where Sequence_Owner = '" & str������ & "'"
        End If
        
        strSql = "Select c.Owner, Decode(s.Sequence_Name, '���˷��ü�¼_ID', '���˷��ü�¼', s.Table_Name) as Table_Name, c.Column_Name, s.Sequence_Name" & vbNewLine & _
                "From All_Tab_Columns C," & vbNewLine & _
                "     (Select Sequence_Name, Sequence_Owner, Decode(Sequence_Name, '���˷��ü�¼_ID', 'סԺ���ü�¼', Table_Name) Table_Name, Column_Name" & vbNewLine & _
                "       From (Select Sequence_Name, Sequence_Owner, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                "                     Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name" & vbNewLine & _
                "              From All_Sequences" & strSql & ")) S" & vbNewLine & _
                "Where c.Table_Name = s.Table_Name And c.Column_Name = s.Column_Name And c.Owner = s.Sequence_Owner" & vbNewLine & _
                "Order By c.Owner, s.Table_Name"
    End If
    rsSeq.CursorLocation = adUseClient
    rsSeq.Open strSql, cnOwner, adOpenStatic, adLockReadOnly
    Set GetSequence = rsSeq
End Function

Public Function AdjustNameSequece(ByVal strTable As String, cnOwner As ADODB.Connection, Optional ByVal strColumn As String = "ID", Optional ByVal blnJustGetSQL As Boolean) As String
'���ܣ��������еĵ�ǰ����
'������strTable=Ҫ�����ı���,ע��Ҫʹ��"user.table"��������������
'      strColumn ���ж�Ӧ�������ֶ���,һ��ΪID,����Ϊָ���������ֶ�
'      blnJustGetSQL=��ֻ��ȡSQL
    Dim dblTableID As Double
    Dim dblSequenceID As Double
    Dim lngIncrement As Long   '������ǰ������
    Dim rsVal As New ADODB.Recordset, strSql As String, strTab As String
    Dim strReturn As String
    
    strTable = UCase(strTable)
    strTab = Mid(strTable, InStr(strTable, ".") + 1)
    dblTableID = 0
    dblSequenceID = 0
    If strTab = "������ü�¼" Or strTab = "סԺ���ü�¼" Or strTab = "���˷��ü�¼" Then
        strTab = Replace(strTable, strTab, "")    '������.
        strSql = "Select Max(MID) as MaxID From (" & _
                "Select Max(" & strColumn & ") as MID From " & strTab & "������ü�¼ " & _
                "Union All Select Max(" & strColumn & ") as MID From " & strTab & "סԺ���ü�¼)"
    Else
        strSql = "Select Max(" & strColumn & ") as MaxID From " & strTable
    End If
    
    rsVal.CursorLocation = adUseClient
    rsVal.Open strSql, cnOwner, adOpenKeyset
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTableID = CDbl(rsVal!MAXID)
        Else
            dblTableID = 0
        End If
    End If
    rsVal.Close
    
    rsVal.Open "Select " & strTable & "_" & strColumn & ".Nextval AS NextID From Dual", cnOwner, adOpenKeyset
    If Not IsNull(rsVal!NEXTID) Then
        dblSequenceID = CDbl(rsVal!NEXTID)
    Else
        dblSequenceID = 0
    End If
    rsVal.Close
    
    If dblTableID - dblSequenceID > 0 Then
        '�޸�����
        rsVal.Open "Select Increment_By From All_Sequences Where Sequence_Owner = '" & Split(strTable, ".")(0) & "' And Sequence_Name ='" & Split(strTable, ".")(1) & "_" & strColumn & "'"
        If Not rsVal.EOF Then
            lngIncrement = Nvl(rsVal!Increment_By, 1)
        Else
            lngIncrement = 1
        End If
        rsVal.Close
        strSql = "Alter Sequence " & strTable & "_" & strColumn & " Increment by " & dblTableID - dblSequenceID
        If blnJustGetSQL Then
            strReturn = "--�޸�����" & vbNewLine & strSql & ";"
        Else
            cnOwner.Execute strSql
        End If
        strSql = "Select " & strTable & "_" & strColumn & ".Nextval as NextID From Dual"
        If blnJustGetSQL Then
            strReturn = strReturn & vbNewLine & "--�ƶ�һ������" & vbNewLine & strSql & ";"
        Else
            rsVal.Open strSql, cnOwner, adOpenKeyset
        End If
        If Not IsNull(rsVal!NEXTID) Then
            dblSequenceID = CDbl(rsVal!NEXTID)
        Else
            dblSequenceID = 0
        End If
        rsVal.Close
        '��ԭ����
        cnOwner.Execute "Alter Sequence " & strTable & "_" & strColumn & " Increment by " & lngIncrement
    End If
End Function

Public Sub Adjust����ID(cnOwner As ADODB.Connection)
'----------------------------------------------
'���ܣ���Խ���ID�Բ��˽��ʼ�¼_ID�������⴦��
'----------------------------------------------
    Dim dblTableID As Double, dblTmp As Double
    Dim dblSequenceID As Double
    Dim lngIncrement As Long   '������ǰ������
    Dim rsVal As New ADODB.Recordset
    dblTableID = 0
    dblSequenceID = 0
    On Error Resume Next
    rsVal.Open "select max(����ID) as MAXID from ����Ԥ����¼", cnOwner, adOpenStatic, adLockReadOnly
    If err <> 0 Then
        '���ܸ�ϵͳ����û����Щ��
        err.Clear
        Exit Sub
    End If
    
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTableID = CDbl(rsVal!MAXID)
        Else
            dblTableID = 0
        End If
    End If
    rsVal.Close
    rsVal.Open "select max(����ID) as MAXID from ������ü�¼", cnOwner, adOpenStatic, adLockReadOnly
    If Not rsVal.EOF Then
        If Not IsNull(rsVal!MAXID) Then
            dblTmp = CDbl(rsVal!MAXID)
        Else
            dblTmp = 0
        End If
        If dblTmp > dblTableID Then
            dblTableID = dblTmp
        End If
    End If
    rsVal.Close
    
    
    rsVal.Open "select ���˽��ʼ�¼_ID.nextval AS NEXTID from dual", cnOwner, adOpenStatic, adLockReadOnly
    If Not IsNull(rsVal!NEXTID) Then
        dblSequenceID = CDbl(rsVal!NEXTID)
    Else
        dblSequenceID = 0
    End If
    
    rsVal.Close
    
    If dblTableID - dblSequenceID > 0 Then
        '�޸�����
        rsVal.Open "select INCREMENT_BY from user_sequences where SEQUENCE_NAME = '���˽��ʼ�¼_ID'"
        If Not rsVal.EOF Then
            lngIncrement = IIf(IsNull(rsVal("INCREMENT_BY")), 1, rsVal("INCREMENT_BY"))
        Else
            lngIncrement = 1
        End If
        rsVal.Close
        
        cnOwner.Execute "alter sequence ���˽��ʼ�¼_ID increment by " & (dblTableID - dblSequenceID)
        
        rsVal.Open "select ���˽��ʼ�¼_ID.nextval AS NEXTID from dual", cnOwner, adOpenStatic, adLockReadOnly
        If Not IsNull(rsVal!NEXTID) Then
            dblSequenceID = CDbl(rsVal!NEXTID)
        Else
            dblSequenceID = 0
        End If
        rsVal.Close
'        cnOwner.Execute "select " & strTable & "_ID.nextval from dual"
        '��ԭ����
        cnOwner.Execute "alter sequence ���˽��ʼ�¼_ID increment by " & lngIncrement
    End If
End Sub

Public Sub ApplyOEM(objStatus As Object)
'���״̬��Ӧ��OEM����
    Dim strOEM As String
    Dim strTmp As String
    On Error Resume Next
    
    If objStatus.Panels(1).Bevel = sbrRaised Then
         strTmp = gobjRegister.zlRegInfo("��Ʒ����")
         If strTmp <> "-" Then
             objStatus.Panels(1).Text = strTmp & "���"
             If gobjFile Is Nothing Then Set gobjFile = New FileSystemObject
             If gstrAppsoft = "" Then
                 gstrAppsoft = App.Path
                 If gblnInIDE Then
                     gstrAppsoft = "C:\APPSOFT"
                 End If
             End If
             
             If gobjFile.FileExists(gstrAppsoft & "\�����ļ�\logo_app.jpg") Then
                  Set objStatus.Panels(1).Picture = LoadPicture(gstrAppsoft & "\�����ļ�\logo_app.jpg")
             Else
                 '����״̬��ͼ���OEM����
                 If strTmp = "����" Then
                     If gobjRegister.zlRegInfo("��Ȩ����") <> "1" Then
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Try")
                     Else
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
                     End If
                 Else
                     strOEM = GetOEM(strTmp)
                     Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
                     If err <> 0 Then
                         err.Clear
                         Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
                     End If
                 End If
             End If
    
             If gobjRegister.zlRegInfo("��Ȩ����") <> "1" Then
                 If strTmp = "����" Then
                     objStatus.Panels(1).Text = ""
                 Else
                     objStatus.Panels(1).Text = strTmp & "(����)"
                 End If
             End If
             objStatus.Panels(1).ToolTipText = ""
             objStatus.Height = 360
         End If
     End If
End Sub

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str���� As String, Optional ByVal strProductName As String)
'��Ը���ͼ��Ӧ��OEM����
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If strProductName = "" Then
        strProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    End If

    If strProductName <> "����" And strProductName <> "-" Then
        '����״̬��ͼ���OEM����
        If Right(str����, 1) = "B" Then
            '��ʾ��ƷͼƬ
            blnCorp = False
            str���� = Mid(str����, 1, Len(str����) - 1)
        Else
            '��ʾ��˾�ձ�
            blnCorp = True
        End If
        
        strOEM = GetOEM(strProductName, blnCorp)
        If str���� = "Picture" Then
            Set objPicture.Picture = LoadCustomPicture(strOEM)
        ElseIf str���� = "Icon" Then
            Set objPicture.Icon = LoadCustomPicture(strOEM)
        End If
        
        If err <> 0 Then
            err.Clear
        End If
    
    End If
End Sub

Public Function LoadCustomPicture(strID As String) As StdPicture
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEMͼƬ���������� ��һ��ָ��˾�ձ꣬��һ���ǲ�Ʒ��ʶ
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function


Public Sub ReCompileProcedure(ByVal cnOwner As ADODB.Connection)
'�Ա��û��������Ѿ�ʧЧ�Ĺ��̽������±���
    Dim rsTemp As New ADODB.Recordset
    Dim lngTime As Long
    
    For lngTime = 1 To 3
        '���������Σ���Ϊ��Щ�������໥���ã�һ�α��벻�ܽ������
        'Ϊ�˿��ٵõ��б������ö���֮������ù�ϵ
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        gstrSQL = "select OBJECT_NAME from user_objects where object_type='PROCEDURE' and STATUS='INVALID'"
        rsTemp.Open gstrSQL, cnOwner, adOpenStatic, adLockReadOnly
        
        On Error Resume Next
        If rsTemp.RecordCount = 0 Then
            'û�й���ʧЧ��ֱ���˳�
            Exit Sub
        Else
            Do Until rsTemp.EOF
                '�п��ܳ���
                gstrSQL = "alter procedure " & rsTemp("OBJECT_NAME") & " compile"
                cnOwner.Execute gstrSQL
                rsTemp.MoveNext
            Loop
        End If
    Next
End Sub

Public Function LoadServer(ByRef strFileInfo As String) As Collection
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long
    Dim colServer As New Collection

    Set rsOraHome = New ADODB.Recordset
    With rsOraHome
        .Fields.Append "Name", adVarChar, 256 'Name
        .Fields.Append "VerSion", adInteger  '�汾
        .Fields.Append "Times", adInteger '�ڼ��ΰ�װ
        .Fields.Append "Server", adInteger '1-������,2-�ͻ���
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '1:��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
        arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
        If TypeName(arrTmp) = "Empty" Then
            If Is64bit Then
                strFileInfo = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle��"
            Else
                strFileInfo = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Oracle��"
            End If
        Else
            For i = LBound(arrTmp) To UBound(arrTmp)
                If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                    intVersion = 0: intTimes = 0:  intServer = 1
                    If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                        .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                        .Update
                    End If
                End If
            Next
            If UBound(arrTmp) <> -1 Then ''����Ŀ¼������Oracle_Home��Ϣ��Ĭ�϶�ȡ���
                .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
            End If
            .Sort = "VerSion Desc,Times Desc,Server"
            Do While Not .EOF
                strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle" & !name, "ORACLE_HOME")
                If strPath = "" And !name & "" = "" Then
                    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle", "ORA_CRS_HOME")
                End If
                If strPath <> "" Then
                    strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
                    If gobjFile.FileExists(strFile) Then Exit Do
                    strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                    If gobjFile.FileExists(strFile) Then Exit Do
                End If
                strFile = ""
                .MoveNext
            Loop
        End If
    End With
    If strFile = "" Then Exit Function
    strFileInfo = "�������б���Դ:" & strFile
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        strLine = Trim(strLine)
        If strLine <> "" And Left(strLine, 1) <> "#" Then
            '��ע���л����
            If InStr(strLine, "(") = 0 And InStr(strLine, ")") = 0 Then
                '���е����ݾ��Ƿ��������ˣ����������ݶ���ʼ��
                strServer = Trim(Mid(strLine, 1, InStr(strLine, "=") - 1))
                strComputer = ""
                strSID = ""
            ElseIf InStr(strLine, "(ADDRESS") > 0 Then
                '���е�������������
                If InStr(strLine, "PROTOCOL = TCP") > 0 And InStr(strLine, "PORT = ") > 0 Then
                    '�������ǵĳ���Ҫ��
                    strComputer = Mid(strLine, InStr(strLine, "HOST =") + Len("HOST ="))
                    strComputer = Trim(Mid(strComputer, 1, InStr(strComputer, ")") - 1))
                End If
            Else
                lngPos = InStr(strLine, "(SID")
                If lngPos = 0 Then
                    lngPos = InStr(strLine, "(SERVICE_NAME")
                End If
                
                If lngPos > 0 Then
                    '���е�������ʵ����
                    strSID = Mid(strLine, InStr(lngPos, strLine, "=") + 1)
                    strSID = Trim(Mid(strSID, 1, InStr(strSID, ")") - 1))
                    
                    If strServer <> "" And strComputer <> "" And strSID <> "" Then
                        '�Ѿ��õ�������Ҫ������
                        colServer.Add Array(strServer, strComputer, strSID)
                    End If
                End If
            End If
        End If
    Loop
    Close #lngFile
    
    Set LoadServer = colServer
End Function

Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'����:ͨ��OracleHome����ȡOracle��Ϣ
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*�汾Home_32Bit
    'Key_Ora*�汾_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Public Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function


Public Function ExpandEnvStr(ByVal strInput As String) As String
'���ܣ����ַ����еĻ��������滻Ϊ����ֵ
'         strInput=���������������ַ���
'���أ���ʵ�ʵ�ֵ�滻�ַ����еĻ�����������ַ���
    '// �磺 %PATH% �򷵻� "c:\;c:\windows;"
    Dim lngLen As Long, strBuf As String, strOld As String
    strOld = strInput & "  " ' ��֪ΪʲôҪ�������ַ������򷵻�ֵ������������ַ���
    strBuf = "" '// ��֧��Windows 95
    '// get the length
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, lngLen)
    '// չ���ַ���
    strBuf = String$(lngLen - 1, Chr$(0))
    lngLen = ExpandEnvironmentStrings(strOld, strBuf, LenB(strBuf))
    '// ���ػ�������
    ExpandEnvStr = TruncZero(strBuf)
End Function

Public Function ValEx(ByVal varInput As Variant) As Variant
'���ܣ�����Valֻ�������ֿ�ͷʶ��ValEx�Ե�һ�����ֽ���ʶ��
    Dim arrTmp As Variant, lngPos As Long
    If Val(varInput) = 0 Then
        varInput = varInput & ""
        If Trim(varInput) = "" Then ValEx = 0: Exit Function
        For lngPos = 1 To Len(varInput)
            If IsNumeric(Mid(varInput, lngPos, 1)) Then Exit For
        Next
        If lngPos = Len(varInput) + 1 Then
            ValEx = 0
        Else
            ValEx = Val(Mid(varInput, lngPos))
        End If
    Else
        ValEx = Val(varInput)
    End If
End Function

 Public Function Is64bit() As Boolean
    '******************************************************************************************************************
    '���ܣ��Ƿ���64λϵͳ
    '���أ�
    '******************************************************************************************************************
    Dim handle As Long
    Dim bolFunc As Boolean
        
    bolFunc = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle > 0 Then
        IsWow64Process GetCurrentProcess(), bolFunc
    End If
    Is64bit = bolFunc
End Function

Public Function GetAllSubKey(ByVal KeyRoot As Long, KeyName As String) As Variant
'����:��ȡĳ�����������
'���أ�=��������
    Dim lnghKey As Long, lngRet As Long, strName As String, lngIdx As Long
    Dim strSubKey As Variant
    strSubKey = Array()
    lngIdx = 0: strName = String(256, Chr(0))
    lngRet = RegOpenKey(KeyRoot, KeyName, lnghKey)
    If lngRet = 0 Then
        Do
            lngRet = RegEnumKey(lnghKey, lngIdx, strName, Len(strName))
            If lngRet = 0 Then
                ReDim Preserve strSubKey(UBound(strSubKey) + 1)
                strSubKey(UBound(strSubKey)) = Left(strName, InStr(strName, Chr(0)) - 1)
                lngIdx = lngIdx + 1
            End If
        Loop Until lngRet <> 0
    End If
    RegCloseKey lnghKey
    GetAllSubKey = strSubKey
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'���ܣ���ע���
    Dim i As Long                                           ' ѭ��������
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' ����򿪵�ע���ؼ���
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' ע���ؼ�����������
    Dim tmpVal As String                                    ' ע���ؼ��ֵ���ʱ�洢��
    Dim KeyValSize As Long                                  ' ע���ؼ��ֱ����ߴ�
    
    ' �� KeyRoot {HKEY_LOCAL_MACHINE...} �´�ע���ؼ���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���ؼ���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ����ߴ�
    
    '------------------------------------------------------------
    ' ����ע���ؼ��ֵ�ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' ���/�����ؼ��ֵ�ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' �����ؼ���ֵ��ת������...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������������...
    Case REG_SZ, REG_EXPAND_SZ                              ' �ַ���ע���ؼ�����������
        sKeyVal = tmpVal                                     ' �����ַ�����ֵ
    Case REG_DWORD                                          ' ���ֽ�ע���ؼ�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ�������ֵ��
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' ת�����ֽ�Ϊ�ַ���
    End Select
    
    GetKeyValue = sKeyVal                                   ' ����ֵ
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
    Exit Function                                           ' �˳�
    
GetKeyError:    ' ����������������...
    GetKeyValue = vbNullString                              ' ���÷���ֵΪ���ַ���
    rc = RegCloseKey(hKey)                                  ' �ر�ע���ؼ���
End Function

Private Function GetKeyValueInfo(ByVal strKey As String, Optional ByVal strValueName As String, Optional ByRef hRootKey As REGRoot, Optional ByRef strSubKey As String, Optional ByRef lngType As Long) As Boolean
'���ܣ����ݼ�λ��ȡ����ֵ���ӽ�,�Լ�ֵ����
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���Σ�
'          hRootKey=����
'          strSubKey=�ӽ�
'          lngType=������
'���أ��Ƿ��ȡ�ɹ�
    Dim strRoot As String, lngPos As String, hKey As Long
    Dim lngReturn As Long, strName As String * 255
    
    On Error GoTo errH
    hRootKey = 0: strSubKey = "": lngType = 0
    lngPos = InStr(strKey, "\")
    If lngPos = 0 Then Exit Function
    strRoot = Mid(strKey, 1, lngPos - 1)
    strSubKey = Mid(strKey, lngPos + 1)
    
    hRootKey = Decode(UCase(strRoot), "HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT, _
                                                                         "HKEY_CURRENT_USER", HKEY_CURRENT_USER, _
                                                                         "HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE, _
                                                                         "HKEY_USERS", HKEY_USERS, _
                                                                         "HKEY_PERFORMANCE_DATA", HKEY_PERFORMANCE_DATA, _
                                                                         "HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG, _
                                                                         "HKEY_DYN_DATA", HKEY_DYN_DATA, 0)
    If hRootKey = 0 Then Exit Function
    If lngType <> -1 Then
        'ʹ�ò�ѯ��ʽ�򿪣����м������Ͳ�ѯ
        lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, hKey)
        If lngReturn <> ERROR_SUCCESS Then
            Exit Function
        End If
        If strValueName <> "" Then
            lngReturn = RegQueryValueEx_ValueType(hKey, strValueName, ByVal 0&, lngType, ByVal strName, Len(strName))
            '�����ֶγ��������Ȳ��������Գ����˳�
            'If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (hKey): Exit Function
        End If
        RegCloseKey (hKey)
    End If
    GetKeyValueInfo = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    err.Clear
End Function

Public Function GetRegValue(ByVal strKey As String, ByVal strValueName As String, ByRef varValue As Variant, Optional blnOneString As Boolean = False) As Boolean
'���ܣ���ȡע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'           blnOneString = ��REG_EXPAND_SZ��REG_MULTI_SZ,REG_BINARY��Ч��-  True �������ص�һ�ַ������Ҳ����κδ���ֻȥ���ַ���β��
'���أ��Ƿ��ȡ�ɹ�
'˵������ǰֻ��REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ��REG_DWORD��REG_BINARYʵ���˶�ȡ��û�в�ѯ�������Զ����Ҽ���
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, strReturn As String, strTmp As String
    '������Ч��ע����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_QUERY_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ '�ַ������Ͷ�ȡ
'            lngReturn = RegQueryValueEx(lngKey, strValueName, 0, ruType, 0, lngLength)
'            If lngReturn <> ERROR_SUCCESS Then Err.Clear '���ܳ��������������
            lngLength = 1024: strBuf = Space(lngLength)
            lngReturn = RegQueryValueEx_String(lngKey, strValueName, 0, ruType, strBuf, lngLength)
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): Exit Function
            Select Case ruType
                Case REG_SZ
                    varValue = TruncZero(strBuf)
                Case REG_EXPAND_SZ ' ���价���ַ�������ѯ���������ͷ��ض���ֵ
                    If Not blnOneString Then
                        varValue = TruncZero(ExpandEnvStr(TruncZero(strBuf)))
                    Else
                        varValue = TruncZero(strBuf)
                    End If
                Case REG_MULTI_SZ ' �����ַ���
                    If Not blnOneString Then
                        If Len(strBuf) <> 0 Then ' �������Ƿǿ��ַ��������Էָ
                            strBufVar = Split(Left$(strBuf, Len(strBuf) - 1), Chr$(0))
                        Else ' ���ǿ��ַ�����Ҫ����S(0) ���������
                            ReDim strBufVar(0) As String
                        End If
                        ' ��������ֵ������һ���ַ������飿��
                        varValue = strBufVar()
                    Else
                        varValue = TruncZero(strBuf)
                    End If
            End Select
        Case REG_DWORD
            lngReturn = RegQueryValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, Len(lngBuf))
            If lngReturn <> ERROR_SUCCESS Then: RegCloseKey (lngKey): varValue = 0: Exit Function
            varValue = lngBuf
        Case REG_BINARY
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, ByVal 0, lngLength)
            If lngReturn <> ERROR_SUCCESS Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            ReDim bytBuf(lngLength - 1)
            lngReturn = RegQueryValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
            If lngReturn <> ERROR_SUCCESS Then
                RegCloseKey lngKey: Exit Function
                If blnOneString Then
                    varValue = "00"
                Else
                    ReDim bytBuf(0)
                    varValue = bytBuf()
                End If
            End If
            If lngLength <> UBound(bytBuf) + 1 Then
               ReDim Preserve bytBuf(0 To lngLength - 1) As Byte
            End If
            ' �����ַ�����ע�⣺Ҫ���ֽ��������ת����
            If blnOneString Then
                'ѭ�����ݣ����ֽ�ת��Ϊ16�����ַ���
                For i = LBound(bytBuf) To UBound(bytBuf)
                   strTmp = CStr(Hex(bytBuf(i)))
                   If (Len(strTmp) = 1) Then strTmp = "0" & strTmp
                   strReturn = strReturn & " " & strTmp
                Next i
                varValue = Trim$(strReturn)
            Else
                varValue = bytBuf()
            End If
    End Select
    RegCloseKey lngKey
    GetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function SetRegValue(ByVal strKey As String, ByVal strValueName As String, varValue As Variant) As Boolean
'���ܣ�����ע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'          strValue=����ֵ
'          strValueType=�������ͣ�Ĭ��Ϊ�ַ���
'���أ��Ƿ����óɹ�
    Dim hRootKey As REGRoot, strSubKey As String
    Dim lngReturn As Long
    Dim lngKey As Long, ruType As REGValueType
    Dim lngLength As Long, varBufData As Variant, strBufVar() As String, lngBuf As Long, bytBuf() As Byte, strBuf As String
    Dim i As Long, lb As Long, ub As Long, strReturn As String, strTmp As String
    '������Ч��ע����λ,��ȡ��������
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, ruType) Then Exit Function
    '�򿪱���
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> ERROR_SUCCESS Then
        Exit Function
    End If
    On Error GoTo errH
    Select Case ruType
        Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
            If ruType = REG_MULTI_SZ And varType(varValue) = vbArray + vbString Then 'string���飬������ϳ��ַ���
                lngLength = UBound(varValue) - LBound(varValue) + 1
                For i = LBound(varValue) To UBound(varValue)
                    strBuf = strBuf & varValue(i) & Chr$(0)
                Next
                strBuf = TruncZero(strBuf)
                lngLength = ActualLen(strBuf)
            Else
                strBuf = TruncZero(varValue)
                lngLength = ActualLen(strBuf)
            End If
            lngReturn = RegSetValueEx_String(lngKey, strValueName, ByVal 0&, ruType, ByVal strBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_DWORD
            lngBuf = Val(varValue): lngLength = Len(lngBuf)
            lngReturn = RegSetValueEx_Long(lngKey, strValueName, ByVal 0&, ruType, lngBuf, lngLength)
            If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
        Case REG_BINARY
            ' 1��varValue �� �ֽ����飬�� B()
            If varType(varValue) = vbArray + vbByte Then
                Dim binValue() As Byte, Length As Long
                bytBuf = varValue
                lngLength = UBound(bytBuf) - LBound(bytBuf) + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 2��varValue �� ���ͻ����ͣ��� 520
            ElseIf varType(varValue) = vbLong Or varType(varValue) = vbInteger Then
                lngBuf = Val(varValue): lngLength = Len(lngBuf)
                lngReturn = RegSetValueEx_Long(lngKey, strValueName, 0, ruType, lngBuf, lngLength)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            ' 3��varValue ���ַ������� "BE 3E FF AB"
            ElseIf varType(varValue) = vbString Then
                ' ת������
                Dim ByteArray() As Byte
                Dim tmpArray() As String '//ת��ASCII�ַ���16�����ֽ�
                strTmp = varValue
                ' �Կո�ָ��ַ���
                strBufVar = Split(strTmp, " ")
                lb = LBound(strBufVar): ub = UBound(strBufVar)
                ' Ϊ��̬�������ռ�
                ReDim bytBuf(lb To ub)
                ' ѭ��ת��
                For i = lb To ub - 1
                    bytBuf(i) = CByte(Val("&H" & Right$(strBufVar(i), 2)))
                Next i
                ' ע�⣺���һ����֪���ַ����������2��ʲô��Ҫ�� Left$(tmpArray(ub), 2)
                bytBuf(ub) = CByte(Val("&H" & Left$(strBufVar(ub), 2)))
                ' ������д�뵽ע���ע�⣺����� ub - lb + 1
                lngReturn = RegSetValueEx_BINARY(lngKey, strValueName, 0, ruType, bytBuf(0), ub - lb + 1)
                If lngReturn <> 0 Then RegCloseKey lngKey: Exit Function
            End If
    End Select
    RegCloseKey lngKey
    SetRegValue = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
End Function

Public Function DeleteRegValue(ByVal strKey As String, ByVal strValueName As String) As Boolean
'���ܣ�ɾ��ע�����ָ��λ�õ�ֵ
'������strKey=ע����λ���硰HKEY_CURRENT_USER\Printers\DevModePerUser"
'          strValueName=������
'���أ��Ƿ��ȡ�ɹ�
    Dim lngLength As Long, lngReturn As Long
    Dim lngKey As Long, lngType As Long
    Dim hRootKey As REGRoot, strSubKey As String
    
    '������Ч��ע����λ
    If Not GetKeyValueInfo(strKey, strValueName, hRootKey, strSubKey, -1) Then Exit Function
    '�򿪼�
    lngReturn = RegOpenKeyEx(hRootKey, strSubKey, 0, KEY_SET_VALUE, lngKey)
    If lngReturn <> 0 Then
        Exit Function
    End If
    'ɾ����
    lngReturn = RegDeleteValue(lngKey, strValueName)
    If lngReturn = 0 Then
        DeleteRegValue = True
    End If
    '�رռ�
    RegCloseKey lngKey
End Function

Public Function CheckSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String) As Boolean
'���ܣ�����ռ�������ļ��Ƿ��������û�ʹ��
'������strType    ��ռ� �����ļ�
'      strName          ��ռ�������ļ�������
'      strOwner         �����������û�����������
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "��ռ�" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2" & vbNewLine & _
                  "union " & vbNewLine & _
                  "select owner from all_indexes where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
        
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    On Error Resume Next
    rsTemp.Open gstrSQL, gcnOracle, , adLockReadOnly
    
    If rsTemp.RecordCount <= 0 Then
        'û�������û�ʹ�ã�����ɾ��
        Exit Function
    End If
    '���û�ʹ��
    CheckSpaceIsUse = True
End Function

Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
'����:���б�ؼ����н�������
'����:
'   objSet��Ҫ���õĶ���,Ŀǰֻ֧��ListView���Ժ��ټ���FlexGrid,DataGrid��
'   strColumn���д�����ʽ��"����,�п�,������ֵ,������;����,�п�,������ֵ,������"    ע����֮�����÷ֺ�
'      ���� "����,2000,0,1;����,800,0,0;����,1440,0,0"
'      ��ListView���ԣ�������Ϊ1��ʾ���в���ɾ����������Ϊ0��ʾ���п���ɾ��
'      ��FlexGridView���ԣ������Ի�Ҫ��ʾ�Ƿ����ڹ̶��У��Ա㲻�ܺ������н���˳�����
'   blnInit��True,����ʾѡ�񴰿ڣ�ֱ�ӳ�ʼ��
    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    If blnInit Then
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    End If
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


Public Function OpenCursor(ByVal cnOwner As ADODB.Connection, _
                              ByVal strPackagesName As String, _
                              ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'���ܣ����ô洢���̷��ؼ�¼��
'��Σ�strPackagesName ����ʽΪ [������.]��.������
'-----------------------------------------
    Static cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, i As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '��¼��������
    Dim varOutPar As Variant
    On Error GoTo errHandle

    '���ԭ�в���:��Ȼ�����ظ�ִ��
   
    
    cmdPackage.CommandText = "" '��Ϊ����ʱ�����������
    Do While cmdPackage.Parameters.Count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN ����
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '����
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        If cnOwner Is Nothing Then
            Set cmdPackage.ActiveConnection = gcnOracle
        Else
            Set cmdPackage.ActiveConnection = cnOwner
        End If
    Else
        If Not cnOwner Is Nothing Then
            If cmdPackage.ActiveConnection.ConnectionString <> cnOwner.ConnectionString Then
                Set cmdPackage.ActiveConnection = cnOwner
            End If
        End If
    End If
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    cmdPackage.Properties("PLSQLRSet") = False
    Exit Function
errHandle:
    If MsgBox(err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If

End Function

Public Function GetAllSystems(Optional blnAll As Boolean) As ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If blnAll Then
        gstrSQL = "Select A.���, A.����, A.������, A.�汾�� From Zlsystems A Order By A.���"
    Else
        gstrSQL = "SELECT a.���, a.����, a.������, a.�汾�� " & _
                 "       FROM Zlsystems a, " & _
                 "            (SELECT Owner " & _
                 "              FROM All_Tables " & _
                 "              WHERE Table_Name IN ('���ű�', '��Ա��', '������Ա', '�ϻ���Ա��') " & _
                 "              GROUP BY Owner " & _
                 "              HAVING COUNT(Owner) = 4) b " & _
                 "       WHERE a.������ = b.Owner " & _
                 "       ORDER BY a.���"
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenForwardOnly, adLockReadOnly
    Set GetAllSystems = rsTemp
    Exit Function
errHandle:
    MsgBox err.Description, vbCritical, gstrSysName
    
End Function


Public Function CheckHistorySpaces(ByVal cnOracle As ADODB.Connection, ByVal pgbProcess As ProgressBar, ByVal strBakOwner As String, ByVal strDbLink As String, _
                 ByVal lngSys As Long, ByVal strSysOwner As String, _
                 Optional bytCheckSys As Byte, Optional ByRef cllErrMsg As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------------------
    '����:�����ʷ���ݿռ��������Ƿ�����
    '����:cnOracle-�������ݿ�����
    '     strDbLink-������
    '     pgbProcess-������
    '     lngSys-ϵͳ��
    '     strSysOwner-ϵͳ������
    '     strBakOwner-���ݿռ��������
    '     cllErr:������Ϣ��(0-��������,1-��������,2-������Ϣ,2-�������ؼ���˵��)
    '     bytCheckSys-0-��������Ƿ���zlbakInfor���д���ϵͳ(������,1-�����������,>1��ʾȫ���:��Ҫ�Ǽ�����ͱ�
    '����:strErrMsg-������صĴ�����Ϣ
    '����:������Ϸ�,����true,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------------------
    Dim rsBakObject As New ADODB.Recordset, rsObject As New ADODB.Recordset
    Dim cllErr  As Collection, blnBakInfor As Boolean
    Dim strTemp  As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    
    If strDbLink <> "" Then
        '��������Ƿ�����
        err = 0: On Error Resume Next
        gstrSQL = "Select 1 from dual@" & strDbLink
        OpenRecordset rsBakObject, gstrSQL, "������֤", , , cnOracle
        If err <> 0 Then
            cllErr.Add Array("Զ������", strDbLink, "���Ӳ�����", "���أ���ʷ���ݿռ佫������������")
            Exit Function
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    blnBakInfor = True
    '����Ƿ������ʷ���ݿռ�
    gstrSQL = "Select ���� From zlBakTables where ϵͳ=" & lngSys
    OpenRecordset rsObject, gstrSQL, "����Ƿ������ʷ���ݿռ�", , , cnOracle
    If rsObject.EOF Then
        '��������ʷ���ݿռ䣬������ȷ��
        CheckHistorySpaces = True
        Exit Function
    End If
    
            
    If strDbLink <> "" Then
        gstrSQL = "select table_name as ����  from " & strBakOwner & ".user_tables"         '& " where  owner = '" & strBakOwnerName & "' "
    Else
        gstrSQL = "select table_name as ����  from user_tables@" & strDbLink         '& " where  owner = '" & strBakOwnerName & "' "
    End If
    
    OpenRecordset rsBakObject, gstrSQL, "��ȡ��ʷ�ռ��", , , cnOracle
    
    'cllErr����(0-��������,1-��������,2-������Ϣ,2-�������ؼ���˵��)
    
     Set cllErr = New Collection
    '���zlBakInfo���Ƿ����
    rsBakObject.Filter = "����='" & UCase("zlBakInfo") & "'"
    If rsBakObject.EOF Then
        cllErr.Add Array("��", "zlBakInfo", "������", "���أ���ʷ���ݿռ佫������������")
        blnBakInfor = False
    End If
    If (bytCheckSys = 0 Or bytCheckSys > 1) And blnBakInfor Then
        If strDbLink <> "" Then
            gstrSQL = "Select 1 From " & strBakOwner & ".zlBakInfo where ϵͳ=" & lngSys
        Else
            gstrSQL = "Select 1 From zlBakInfo@" & strDbLink & " where ϵͳ=" & lngSys
        End If
        OpenRecordset rsTemp, gstrSQL, "��ȡϵͳ", , , cnOracle
        If rsTemp.EOF Then
            cllErr.Add Array("ϵͳ����", lngSys, "ϵͳ���Ϊ:" & lngSys & "������", "���أ�Ӱ����ʷ���ݿռ����������")
        End If
        rsTemp.Close
    End If
    
    Dim lngCount As Long
    lngRow = 0
    If bytCheckSys >= 1 Then
        With rsObject
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                Call CheckHistoryTable(cnOracle, Nvl(!����), strSysOwner, strBakOwner, Nvl(!����), strDbLink, lngSys, cllErr)
                lngRow = lngRow + 1
                pgbProcess.value = lngRow \ .RecordCount * 100
                .MoveNext
            Loop
        End With
    End If
    rsBakObject.Close
    Set rsBakObject = Nothing
    rsObject.Close
    Set rsObject = Nothing
    If cllErr Is Nothing Then
    Else
    If cllErr.Count <> 0 Then Set cllErrMsg = cllErr: Exit Function
    End If
    CheckHistorySpaces = True
    Exit Function
ErrHand:
   ' Resume
    If cllErr.Count <> 0 Then cllErrMsg = cllErr
End Function


Public Sub OpenRecordset(rsTemp As ADODB.Recordset, strSql As String, ByVal strFormCaption As String, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly, _
        Optional cnOracle As ADODB.Connection = Nothing)
        '���ܣ��򿪼�¼��ͬʱ����SQL���
    On Error GoTo errHandle
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    If cnOracle Is Nothing Then
        rsTemp.Open strSql, gcnOracle, CursorType, LockType
    ElseIf cnOracle.State = 1 Then
        rsTemp.Open strSql, cnOracle, CursorType, LockType
    Else
        rsTemp.Open strSql, gcnOracle, CursorType, LockType
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function RPAD(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = strCode
    End If
    'ȡ��������ַ�
    RPAD = Replace(strTmp, Chr(0), strChar)
End Function


Public Function CheckHistoryTable(ByVal cnOracle As ADODB.Connection, _
    ByVal strTableName As String, ByVal strSysOwner As String, ByVal strBakOwnner As String, ByVal strHistoryTableName As String, strDbLinkName As String, _
    ByVal lngSys As Long, ByRef cllErr As Variant) As Boolean
    '����:���ָ�����������ʷ���ݱ�ռ�ı����Ƿ�һ��
    '����:cnOracle-���߿�����
    '     strTable_name-���߱���
    '     strSysOwner-�������ݿ��������
    '     strBakOwnner-���ݿռ��������
    '     strHistoryTableName-��ʷ����
    '     strDbLinkNameName-Զ��������
    '����:cllErr:������Ϣ��(0-��������,1-��������,2-������Ϣ,2-�������ؼ���˵��)
    '����:���ɹ�,����ture,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsColumn As New ADODB.Recordset
    Dim rsBakColumn As New ADODB.Recordset
    Dim strTemp As String
    '���ñ��Ƿ����zlbakSpaces��
    gstrSQL = "Select 1 from zltools.zlbaktables where ϵͳ=" & lngSys & " and ����='" & strHistoryTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "�����ʷ������", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        CheckHistoryTable = True
        Exit Function
    End If
    err = 0: On Error Resume Next
    If strDbLinkName = "" Then
        gstrSQL = "select table_name as ����  from  all_tables where owner='" & UCase(strBakOwnner) & "' and table_name='" & strHistoryTableName & "'"
    Else
        gstrSQL = "select table_name as ����  from user_tables@" & strDbLinkName & " where  table_name='" & strHistoryTableName & "'"
    End If
 
    OpenRecordset rsTemp, gstrSQL, "�����ʷ���Ƿ����", , , cnOracle
    
    If err = 0 Then
        If rsTemp.EOF Then
            cllErr.Add Array("��", strHistoryTableName, "������", "���أ�Ӱ����ʷ���ݿռ����������")
            GoTo LoopH:
        End If
    Else
            cllErr.Add Array("��", strHistoryTableName, "����������������", "���أ�Ӱ����ʷ���ݿռ����������")
            GoTo LoopH:
    End If
    '��ͼ����Ч�Լ��
    err = 0: On Error Resume Next
    gstrSQL = "Select 1 from H" & strHistoryTableName & " where 1=2 "
    If err <> 0 Then
        '-------------------------------------
        '��ͼ��Ч
        cllErr.Add Array("��ͼ", "H" & strHistoryTableName, "������", "���أ�Ӱ����ʷ���ݿռ����������")
    End If
    
    '�����ص����Ƿ���ȷ
    If rsColumn.State = 1 Then rsColumn.Close
    If rsBakColumn.State = 1 Then rsBakColumn.Close
    
    If strDbLinkName = "" Then
        gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
            "       From  ALL_TAB_COLUMNS" & _
            "       WHERE owner='" & strBakOwnner & "' and TABLE_NAME='" & strHistoryTableName & "'"
    Else
        gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
            "       From USER_TAB_COLUMNS@" & strDbLinkName & _
            "       WHERE TABLE_NAME='" & strHistoryTableName & "'"
    End If
    
    rsBakColumn.Open gstrSQL, cnOracle
    
    gstrSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE,DATA_DEFAULT" & _
        "       From ALL_TAB_COLUMNS" & _
        "       WHERE TABLE_NAME='" & strTableName & "' and OWNER='" & strSysOwner & "'"
    
    rsColumn.Open gstrSQL, cnOracle
                
    With rsColumn
        Do While Not .EOF
            rsBakColumn.Filter = "COLUMN_NAME='" & Nvl(!Column_Name) & "'"
            If rsBakColumn.EOF Then
                '������
                Select Case Nvl(!DATA_TYPE)
                Case "NUMBER"
                    strTemp = Nvl(!Column_Name) & " NUMBER(" & Nvl(!Data_Precision) & "," & Nvl(!Data_Scale) & ")"
                    If Not IsNull(!DATA_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !DATA_DEFAULT
                Case "VARCHAR2"
                    strTemp = Nvl(!Column_Name) & " VARCHAR2(" & Nvl(!Data_Length) & ")"
                    If Not IsNull(!DATA_DEFAULT) Then strTemp = strTemp & " DEFAULT " & !DATA_DEFAULT
                Case Else
                    strTemp = Nvl(!Column_Name) & Space(2) & Nvl(!DATA_TYPE)
                End Select
                cllErr.Add Array("���ݱ�", strHistoryTableName, "ȱ���� " & strTemp, "���أ�Ӱ����ʷ���ݿռ����������")
            Else
                '��鳤��
                    Select Case !DATA_TYPE
                    Case "NUMBER"
                        If Val(Nvl(!Data_Precision)) <> Val(Nvl(rsBakColumn!Data_Precision)) Or Val(Nvl(!Data_Scale)) <> Val(Nvl(rsBakColumn!Data_Scale)) Then
                            strTemp = Nvl(!Column_Name) & "�г���С�ڹ涨ֵ��ӦΪ��" & "NUMBER(" & Nvl(!Data_Precision) & "," & Val(Nvl(!Data_Scale)) & ")��" & _
                                     " ��Ϊ��" & "NUMBER(" & rsBakColumn!Data_Precision & "," & Val(Nvl(rsBakColumn!Data_Scale)) & ")��"
                            If Val(Nvl(!Data_Precision)) > Val(Nvl(rsBakColumn!Data_Precision)) Then
                                cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���أ�Ӱ����ʷ���ݿռ����������")
                            ElseIf Val(Nvl(!Data_Scale)) > Val(Nvl(rsBakColumn!Data_Scale)) Then
                                cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���أ����ܵ�����ʷ���ݿռ����ݾ��Ȳ���")
                            Else
                                cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���᣺������Ӱ����ʷ���ݿռ������")
                            End If
                        End If
                    Case "VARCHAR2"
                        If Val(Nvl(!Data_Length)) <> Val(Nvl(rsBakColumn!Data_Length)) Then
                            strTemp = Nvl(!Column_Name) & "�г���С�ڹ涨ֵ��ӦΪ��" & "VARCHAR2(" & Val(Nvl(!Data_Length)) & ")��" & _
                                     " ��Ϊ��" & "VARCHAR2(" & Val(Nvl(rsBakColumn!Data_Length)) & ")��"
                            If Val(Nvl(!Data_Length)) > Val(Nvl(rsBakColumn!Data_Length)) Then
                                cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���أ����ܵ�����ʷ���ݿռ����ݵĽϳ��ı��޷��洢")
                            Else
                                cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���᣺������Ӱ����ʷ���ݿռ������")
                            End If
                        End If
                    Case Else
                    End Select
                    If Nvl(!DATA_TYPE) <> Nvl(rsBakColumn!DATA_TYPE) Then
                        strTemp = Nvl(!Column_Name) & "�е����Ͳ���,ӦΪ��" & Nvl(!DATA_TYPE) & "��" & _
                                 " ��Ϊ��" & Nvl(rsBakColumn!DATA_TYPE) & "��"
                        cllErr.Add Array("���ݱ�", strHistoryTableName, strTemp, "���أ����ܵ�����ʷ���ݿռ�����ݴ洢����")
                        
                    End If
            End If
                 
            .MoveNext
        Loop
    End With
LoopH:
CheckHistoryTable = True
 
End Function

Public Sub CheckBakConstraint(ByVal cnOracle As ADODB.Connection, ByVal strBakOwner As String, ByVal strDbLinkName As String, ByVal strTableName As String, _
        ByVal strConstraintName As String, ByVal strSql As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '����:��鱸�����ݿ��Լ��
    '����:cnOracle-�������ݿ�����
    '     strDbLinkName-Զ������
    '     strTableName-����
    '     strBakOwner-�������ݿռ�������
    '     strConstraintName-Լ����
    '     strSQL-Լ����SQL���
    '����:cllErr-���ش�����Ϣ
    '---------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsConstraint As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim strColumns As String
    Dim arySql As Variant
    Dim strTemp As String
    '���ñ��Ƿ����zlbakSpaces��
 
    
    gstrSQL = "Select 1 from zltools.zlbaktables where ϵͳ=" & lngSys & " and ����='" & strTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "�����ʷ������", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Exit Sub
    End If
    
    If strDbLinkName <> "" Then
        gstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from USER_CONSTRAINTS@" & strDbLinkName & " where CONSTRAINT_NAME='" & strConstraintName & "'"
    Else
        gstrSQL = "select CONSTRAINT_TYPE,CONSTRAINT_NAME,STATUS,VALIDATED,BAD from all_CONSTRAINTS where  OWNER='" & strBakOwner & "' and CONSTRAINT_NAME='" & strConstraintName & "'"
    End If
    OpenRecordset rsConstraint, gstrSQL, "��ȡԼ��", , , cnOracle
    
    With rsConstraint
            If rsConstraint.EOF Then
                'Լ��������
                If InStr(1, strSql, " CHECK") > 0 Then
                    cllErr.Add Array("Լ��", strConstraintName, "������", "���᣺������Ӱ����ʷ���ݿռ������")
                ElseIf InStr(1, strSql, " FOREIGN ") = 0 Then
                    cllErr.Add Array("Լ��", strConstraintName, "������", "���أ����ܵ�����ʷ���ݿռ�����ݲ�һ�£�Ӱ�������ٶ�")
                End If
                Exit Sub
            End If
            If .Fields("STATUS").value <> "ENABLED" Then
                cllErr.Add Array("Լ��", strConstraintName, "��ǰ���ڽ�ֹ״̬", "���أ�������ʷ���ݿռ��Ѿ���������")
                Exit Sub
            End If
            If !VALIDATED <> "VALIDATED" Then
                cllErr.Add Array("Լ��", strConstraintName, "��ǰ������Ч״̬", "���أ�������ʷ���ݿռ������һ�����ѱ��ƻ�")
                Exit Sub
            End If
            If Not IsNull(!BAD) Then
                cllErr.Add Array("Լ��", strConstraintName, "Լ����������", "���أ�������ʷ���ݿռ����Ӳ������")
                Exit Sub
            End If
            strColumns = ""
            
            If strDbLinkName = "" Then
                gstrSQL = "" & _
                    "   Select COLUMN_NAME" & _
                    "   From all_CONS_COLUMNS" & _
                    "   where owner='" & strBakOwner & "' and CONSTRAINT_NAME='" & strConstraintName & "'" & _
                    "   order by POSITION"
            Else
                gstrSQL = "" & _
                    "   Select COLUMN_NAME" & _
                    "   From USER_CONS_COLUMNS@" & strDbLinkName & _
                    "   where CONSTRAINT_NAME='" & strConstraintName & "'" & _
                    "   order by POSITION"
            End If
            OpenRecordset rsColumns, gstrSQL, "��ȡ�������", , , cnOracle
                 
            With rsColumns
                Do While Not .EOF
                    strColumns = strColumns & "," & !Column_Name
                    .MoveNext
                Loop
            End With
            If InStr(1, strSql, " PRIMARY ") > 0 Then
                If !constraint_type <> "P" Then
                    cllErr.Add Array("Լ��", strConstraintName, "Լ�����ʹ���ӦΪ����Լ��", "���أ�����Ӱ����ʷ���ݿռ������")
                Else
                    arySql = Split(strSql, " PRIMARY ")
                    strTemp = Replace(Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "KEY", ""), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        cllErr.Add Array("Լ��", strConstraintName, "Լ���д���ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")", "���أ�����Ӱ����ʷ���ݿռ������")
                    End If
                End If
                Exit Sub
            End If
            If InStr(1, strSql, " UNIQUE") > 0 Then
                If !constraint_type <> "U" Then
                    cllErr.Add Array("Լ��", strConstraintName, "Լ�����ʹ���ӦΪΨһԼ��", "���أ�����Ӱ����ʷ���ݿռ������")
                Else
                    arySql = Split(strSql, " UNIQUE ")
                    If UBound(arySql) = 0 Then arySql = Split(strSql, " UNIQUE(")
                    strTemp = Replace(Replace(Left(arySql(1), InStr(1, arySql(1), ")") - 1), "(", ""), " ", "")
                    If strColumns <> "," & strTemp Then
                        cllErr.Add Array("Լ��", strConstraintName, "Լ���д���ӦΪ(" & strTemp & ")����Ϊ(" & Mid(strColumns, 2) & ")", "���أ�����Ӱ����ʷ���ݿռ������")
                    End If
                End If
                Exit Sub
            End If
            If InStr(1, strSql, " CHECK") > 0 Then
                If !constraint_type <> "C" Then
                    cllErr.Add Array("Լ��", strConstraintName, "Լ�����ʹ���ӦΪ���Լ��", "���أ�����Ӱ����ʷ���ݿռ������")
                End If
            End If
    End With
End Sub
Public Sub CheckBakIndex(ByVal cnOracle As ADODB.Connection, ByVal strBakOwner As String, ByVal strDbLinkName As String, ByVal strTableName As String, _
        ByVal strIndexName As String, ByVal strSql As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '����:��鱸�����ݿ��Լ��
    '����:cnOracle-�������ݿ�����
    '     strBakOwner-���ݿռ�������
    '     strDbLinkName-Զ������
    '     strTableName-����
    '     strIndexName-������
    '     strSQL-Լ����SQL���
    '����:cllErr-���ش�����Ϣ
    '---------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsIndex As New ADODB.Recordset
    Dim rsColumns As New ADODB.Recordset
    Dim strColumns As String
    Dim arySql As Variant
    Dim strTemp As String
    '���ñ��Ƿ����zlbakSpaces��
 
    On Error GoTo errHandle
    gstrSQL = "Select 1 from zltools.zlbaktables where ϵͳ=" & lngSys & " and ����='" & strTableName & "'"
    OpenRecordset rsTemp, gstrSQL, "�����ʷ������", , , cnOracle
    If rsTemp.RecordCount = 0 Then
        rsTemp.Close
        Exit Sub
    End If
    If strDbLinkName = "" Then
        gstrSQL = "select INDEX_NAME,STATUS from all_INDEXES where owner='" & strBakOwner & "' and   INDEX_NAME='" & strIndexName & "'"
    Else
        gstrSQL = "select INDEX_NAME,STATUS from USER_INDEXES@" & strDbLinkName & " where  INDEX_NAME='" & strIndexName & "'"
    End If
    OpenRecordset rsIndex, gstrSQL, "��ȡ����", , , cnOracle
    
    With rsIndex
            If rsIndex.EOF Then
                'Լ��������
                cllErr.Add Array("����", strIndexName, "������", "���أ�����Ӱ����ʷ���ݿռ�������ٶ�")
                Exit Sub
            End If
            If .Fields("STATUS").value <> "VALID" Then
                cllErr.Add Array("����", strIndexName, "��ǰ������Ч״̬", "����:����Ӱ����ʷ���ݿռ�������ٶ�")
                Exit Sub
            End If
            
            If strDbLinkName = "" Then
                strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                        " from all_IND_COLUMNS" & _
                        " where INDEX_OWNER ='" & strBakOwner & "' and INDEX_NAME='" & strIndexName & "'" & _
                        " order by COLUMN_POSITION"
            Else
                strTemp = "select TABLE_NAME,COLUMN_NAME" & _
                        " from USER_IND_COLUMNS@" & strDbLinkName & _
                        " where INDEX_NAME='" & strIndexName & "'" & _
                        " order by COLUMN_POSITION"
            End If
            
            OpenRecordset rsColumns, strTemp, "��ȡ�����������", , , cnOracle
           
            With rsColumns
                Do While Not .EOF
                    If .AbsolutePosition = 1 Then
                        strColumns = !Table_Name & "(" & !Column_Name
                    Else
                        strColumns = strColumns & "," & !Column_Name
                    End If
                    .MoveNext
                Loop
                strColumns = strColumns & ")"
            End With
            arySql = Split(strSql, " ON ")
            strTemp = Replace(Left(arySql(1), InStr(1, arySql(1), ")")), " ", "")
            If strColumns <> strTemp Then
               cllErr.Add Array("����", strIndexName, "�����д���ӦΪ��" & strTemp & "������Ϊ��" & strColumns & "��", "���أ�����Ӱ��ϵͳ�����ٶ�")
            End If
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Sub CheckBakView(ByVal cnOracle As ADODB.Connection, ByVal strOwner As String, ByVal lngSys As Long, ByRef cllErr As Variant)
    '---------------------------------------------------------------------------------
    '����:������ߵ���ʷ���ݿռ���ͼ�Ƿ����
    '����:cnOracle-�������ݿ�����
    '     strOwner-������
    '     lngSys-ϵͳ��
    '����:cllErr-���ش�����Ϣ
    '---------------------------------------------------------------------------------
    Dim rsBakTable As New ADODB.Recordset
    Dim rsView As New ADODB.Recordset
    Dim rsObject As New ADODB.Recordset
    Dim strSql As String
    '���ñ��Ƿ����zlbakSpaces��
    On Error GoTo errHandle
    strSql = "Select ���� from zltools.zlbaktables where ϵͳ=" & lngSys
    OpenRecordset rsBakTable, strSql, "�����ʷ������", , , cnOracle
    If rsBakTable.RecordCount = 0 Then
        rsBakTable.Close
        Exit Sub
    End If
    
    If gblnDBA Then
        strSql = "select VIEW_NAME from DBA_VIEWS where OWNER='" & strOwner & "'"
    Else
        strSql = "select VIEW_NAME from USER_VIEWS"
    End If
    OpenRecordset rsView, strSql, "�����ʷ������", , , cnOracle
    
    With rsBakTable
        Do While Not .EOF
            rsView.Filter = "VIEW_NAME='" & "H" & UCase(!����) & "'"
            If rsView.EOF Then
                '��ͼ������
                cllErr.Add Array("��ͼ", "H" & UCase(!����), "������", "���أ�����Ӱ����ʷ���ݿռ������ת��")
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub ExecuteProcedure(strSql As String, ByVal strFormCaption As String, Optional cnOracle As ADODB.Connection)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSql), 1) = ")" Then
        '���ԭ�в���:��Ȼ�����ظ�ִ��
'        cmdData.CommandText = "" '��Ϊ����ʱ�����������
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        'ִ�еĹ�����
        strTemp = Trim(strSql)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        'ִ�й��̲���
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '�Ƿ����ַ����ڣ��Լ����ʽ��������
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '����
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, Val(strPar))
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ��2000���ַ�����ȷ
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '����
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULLֵ�������ִ���ɼ�����������
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '����
                        If datCur = CDate(0) Then datCur = CurrentDate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULLֵ�����ַ�����ɼ�����������
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '��ѡ��������NULL������ܸı���ȱʡֵ:��˿�ѡ��������д���м�
                        GoTo NoneVarLine
                    Else '�������������ӵı��ʽ���޷�����
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        'ִ�й���
        'If cmdData.ActiveConnection Is Nothing Then
        If cnOracle Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
        Else
            Set cmdData.ActiveConnection = cnOracle '���Ƚ���
        End If
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute

    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSql = "Call " & strSql
    If InStr(strSql, "(") = 0 Then strSql = strSql & "()"
    gcnOracle.Execute strSql, , adCmdText

End Sub

Public Sub AlterUserTableSpaces(ByVal cnOracle As ADODB.Connection, ByVal strUserName As String)
    '----------------------------------------------------------------------------------------------------------------------------
    '����:�޸�ָ���û���Ĭ�ϱ�ռ�(����:10793)
    '����:cnOracle-ָ����Oracle����
    '     strUserName-ָ���û���
    '����:���˺�
    '����:2007/06/01
    '˵��:
    '   ������躺���������ʵ��һ�������,���û��ı�ռ����USERS��Temp��,����û����ĵ���Users��ռ��TMP��ռ�,
    '   �����ZLTOOLSTBS��ZLTOOLSTMP��ռ���
    '----------------------------------------------------------------------------------------------------------------------------
    '���˺�:20070531:�û��ı�ռ䲻����system,��ΪUSERS�����
    '��ΪCreate Userʱ�������û�Create TableȨ��,�Ӱ�ȫ����,ȱʡ��ռ����ΪUSERS
    
    err = 0: On Error Resume Next
    '����û����Ӧ�ı�ռ�,�������
    '9i������ȱʡȫ����ʱ��ռ�,10G��������ʱ��ռ���,�ݲ������⿼��
    gstrSQL = "Alter User " & strUserName & " Default Tablespace USERS"
    cnOracle.Execute gstrSQL
    gstrSQL = "Alter User " & strUserName & " Temporary Tablespace TEMP"
    cnOracle.Execute gstrSQL
    If err <> 0 Then
        '���ĳ�ZLToolsTBS��ռ��ZLTOOLSTMP��ռ�
        gstrSQL = "Alter User " & strUserName & " Default Tablespace ZLTOOLSTBS"
        cnOracle.Execute gstrSQL
        gstrSQL = "Alter User " & strUserName & " Temporary Tablespace ZLTOOLSTMP"
        cnOracle.Execute gstrSQL
    End If
    err.Clear
End Sub


Private Function LogTime() As String
    LogTime = "[" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "] "
End Function

Public Sub zlInitRec(ByRef rsData As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鹹�����ݼ�
    '���:rsData-���ݼ�
    '����:
    '����:
    '����:���˺�
    '����:2009-08-19 14:00:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable  '����:����,Ψһ,���,Լ��,����,����/����,��ͼ,...
        .Fields.Append "��־", adLongVarChar, 2, adFldIsNullable '1-���ڶ���,2-�����ڶ���,3-ʧЧ,4-ȱ����,5-����,6-���ڽ�ֹ״̬,7-Լ����һ��,8-��������ʱ����Ҫ�ȴ������
        .Fields.Append "��ʷ�ռ�", adLongVarChar, 2, adFldIsNullable '1-��ʷ���ݿ�����,0-�������ݿ������󲻴���
        .Fields.Append "��������", adLongVarChar, 30, adFldIsNullable '������,ȱ����...
        .Fields.Append "������Ϣ", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 4000, adFldIsNullable '�洢���̣���������̫�������ΪNULL
        .Fields.Append "�ֶ���", adLongVarChar, 50, adFldIsNullable  '��ʱ,��Ҫ���ĵ��ֶ���
        .Fields.Append "ԭ�ֶ�����", adLongVarChar, 50, adFldIsNullable  '��ʱ,��Ҫ���ĵ�ԭ�ֶ�����
        .Fields.Append "���ֶ�����", adLongVarChar, 50, adFldIsNullable  '��ʱ,��Ҫ���ĵ��ֶ���������
        .Fields.Append "ԭ�ֶγ���", adLongVarChar, 20, adFldIsNullable  '��ʱ,��Ҫ���ĵ��ֶ����ĳ���,�������С��,���Զ��ŷ���,��:16,5
        .Fields.Append "���ֶγ���", adLongVarChar, 20, adFldIsNullable  '��ʱ,��Ҫ���ĵ��ֶ����ĳ���,�������С��,���Զ��ŷ���,��:16,5
        .Fields.Append "������־", adLongVarChar, 2, adFldIsNullable  '0-δ����,1-�Ѿ�����,2-�������,4-����ִ����������Ҫ�ֹ�����
        .Fields.Append "����˵��", adLongVarChar, 500, adFldIsNullable '
        .CursorLocation = adUseClient: .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Function zlInsertRecData(ByVal rsData As ADODB.Recordset, ByVal str������ As String, ByVal str�������� As String, ByVal str���� As String, ByVal int��־ As Integer, _
       ByVal bln��ʷ�ռ� As Boolean, ByVal str������� As String, ByVal str�������� As String, ByVal str������Ϣ As String, Optional str�ֶ��� As String, Optional strԭ�ֶ����� As String, Optional str���ֶ����� As String, _
       Optional strԭ�ֶγ��� As String, Optional str���ֶγ��� As String, Optional cllProcedureExecSQLs As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�򱾵ؼ�¼���в�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-08-19 14:25:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��� As Long, str����˵�� As String, byt������־ As Byte, varData As Variant
    Dim lng���� As Long, lngС�� As Long
    
    str�������� = UCase(str��������)
    byt������־ = 0: str����˵�� = ""
    If Trim(str�ֶ���) <> "" Then
        '��Ҫ��飬�ֶ��������:
        ' 1.���Ͳ���,����ִ����������Ҫ�ֹ�����
        ' 2.�ֶγ���С����ԭ�ֶγ���,����������(����С��)����Ҫ�ֹ�����
        
        If strԭ�ֶ����� <> str���ֶ����� And (strԭ�ֶ����� <> "") Then
            '0-δ����,1-�Ѿ�����,2-�������,4-����ִ����������Ҫ�ֹ�����
            byt������־ = 4
            str����˵�� = "�ֶ����Ͳ��ԣ���������, �䶯���:" & strԭ�ֶ����� & "--->" & str���ֶ�����
        Else
            varData = Split(strԭ�ֶγ��� & ",", ",")
            lng���� = Val(varData(0)): lngС�� = Val(varData(1))
            varData = Split(str���ֶγ��� & ",", ",")
            If UCase(str���ֶ�����) = "NUMBER" Then
                '�����NUMBER�Ļ�����Ҫ����䳤��
                If lng���� > Val(varData(0)) Or lngС�� > Val(varData(1)) Then
                    byt������־ = 4
                    str����˵�� = "ԭ�ֶξ��ȴ��������ֶξ��ȣ���������, �䶯���:" & str�ֶ��� & "  NUMBER(" & lng���� & IIf(lngС�� = 0, "", "," & lngС��) & ")" & "--->" & str�ֶ��� & "  NUMBER(" & Val(varData(0)) & IIf(Val(varData(1)) = 0, "", "," & Val(varData(1))) & ")"
                End If
            ElseIf Left(UCase(Trim(str���ֶ�����)), 7) = "VARCHAR" Then
                '�ַ��ͣ����鳤��
                If lng���� > Val(varData(0)) Then
                    byt������־ = 4
                    str����˵�� = "ԭ�ֶξ��ȴ��������ֶξ��ȣ���������, �䶯���:" & str�ֶ��� & "  NUMBER(" & lng���� & ")" & "--->" & str�ֶ��� & "  NUMBER(" & Val(varData(0)) & ")"
                End If
            End If
        End If
    End If
    With rsData
        lng��� = .RecordCount + 1
        If str���� = "���" Then
            .Filter = "��������='" & str�������� & "'"
            If .RecordCount = 0 Then
                .Filter = 0
                '������,ֻ������,���ھ͸���:ԭ���ǿ��ܴ��ڼ����������
                .AddNew
            End If
        Else
            .AddNew
        End If
        !��� = lng���
        !������ = str������
        !�������� = str��������
        !���� = str����
        !��־ = int��־
        !��ʷ�ռ� = IIf(bln��ʷ�ռ�, 1, 0)
        !������� = IIf(str���� = "����/����", "", str�������)
        !�������� = str��������
        !������Ϣ = str������Ϣ
        !�ֶ��� = str�ֶ���
        !ԭ�ֶ����� = strԭ�ֶ�����
        !���ֶ����� = str���ֶ�����
        !ԭ�ֶγ��� = strԭ�ֶγ���
        !���ֶγ��� = str���ֶγ���
        !������־ = byt������־
        !����˵�� = str����˵��
        .Update
        If str���� = "����/����" Then
            cllProcedureExecSQLs.Add str�������, "K" & lng���
        End If
        .Filter = 0
    End With
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '���:strInfor-ԭ��
    '      lngStart-ֱʼλ��
    '      lngLen-����
    '����:
    '����:
    '����:���˺�
    '����:2009-08-20 12:04:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    err = 0
    On Error GoTo ErrHand:
    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
ErrHand:
    Substr = ""
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal MSG As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If MSG <> WM_CONTEXTMENU Then WndMessage = CallNewWindowProc(hwnd, MSG, wp, lp)
End Function

Public Function CallNewWindowProc(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Call CallWindowProc(glngTXTProc, hwnd, MSG, wParam, lParam)
    
    CallNewWindowProc = True
End Function

Public Function IsCharAlpha(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ�ȫ����Ӣ����ĸ����    '
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If Not ((j > 64 And j < 91) Or (j > 96 And j < 123)) Then
                IsCharAlpha = False
                Exit Function
            End If
        Next
    End If
    IsCharAlpha = True
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '���ܣ��ж�ָ���ַ����Ƿ��к���
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function SQLRecordExecute(ByVal rs As ADODB.Recordset, Optional ByVal strTitle As String, Optional ByVal blnHaveTrans As Boolean = True) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim strSql As String
    
    On Error GoTo ErrHand
    
    If rs.RecordCount > 0 Then
        If Len(strTitle) = 0 Then strTitle = gstrSysName
        blnTran = True
        
        If blnHaveTrans Then gcnOracle.BeginTrans
        
        rs.MoveFirst
    
        For intLoop = 1 To rs.RecordCount
        
            strSql = CStr(rs("SQL").value)
            
            Call ExecuteProcedure(strSql, strTitle)
            
            rs.MoveNext
        Next
    
        If blnHaveTrans Then gcnOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
ErrHand:
    
    If blnTran And blnHaveTrans Then gcnOracle.RollbackTrans
    
    MsgBox err.Description, vbCritical, gstrSysName
    
    
End Function


Public Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Function GetFileName(ByVal strFileName As String, Optional Path As String, Optional WithExt As Boolean = False) As String
'����ļ���
'strFilename �ļ�����·��
'Path ����λ��
'WithExt �Ƿ񷵻غ�׺���� True:����׺���Ʒ��� false:������׺���Ʒ���
    Dim C As String
    Dim tmpString As String
    Dim i As Integer
    Dim szlen As Integer
    Dim Cnt As Integer
    
    szlen = Len(strFileName)
    Cnt = 0
    If InStr(strFileName, "\") = 0 Then
      tmpString = strFileName
      Cnt = InStr(tmpString, ".")
      If Cnt > 0 And Not WithExt Then
          GetFileName = Left(tmpString, Cnt - 1)
      Else
          GetFileName = tmpString
      End If
    Else
      For i = szlen To 1 Step -1
        C = Mid(strFileName, i, 1)
        If C = "\" Then
          Path = Left(strFileName, szlen - Cnt)
          tmpString = Right(strFileName, Cnt)
          Cnt = InStr(tmpString, ".")
          If Cnt > 0 And Not WithExt Then
              GetFileName = Left(tmpString, Cnt - 1)
          Else
              GetFileName = tmpString
          End If
          Exit For
        Else
          Cnt = Cnt + 1
        End If
      Next i
    End If
End Function


Public Function GetWinPath() As String
    '--------------------------------------------------------------------------------------------------------------
    '--����:��ȡϵͳĿ¼
    '--------------------------------------------------------------------------------------------------------------
    Dim Buffer As String
    Dim gstrWinPath As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetWindowsDirectory(Buffer, Len(Buffer))
    gstrWinPath = Left(Buffer, rtn)
    GetWinPath = gstrWinPath
End Function

Public Function GetWinSystemPath() As String
    
    Dim Buffer As String
    Dim strSystem As String
    Dim rtn As Long
    
    Buffer = Space(MAX_PATH)
    rtn = GetSystemDirectory(Buffer, Len(Buffer))
    strSystem = Left(Buffer, rtn)
    
    GetWinSystemPath = strSystem
End Function

Public Sub LvwFlatColumnHeader(ByVal lvw As Object)
'���ܣ�ʹ��ListView���б����Ϊƽ��
    Const strHeaderClass As String = "msvb_lib_header"
    Const HDS_BUTTONS   As Long = 2
    
    Dim lngChild As Long, lngLen As Long, LngStyle As Long
    Dim strName As String * 255

    
    lngChild = GetWindow(lvw.hwnd, GW_CHILD)
    Do While lngChild <> 0
        lngLen = GetClassName(lngChild, strName, 255)
        If lngLen > 0 Then
            If Mid(strName, 1, lngLen) = strHeaderClass Then
                LngStyle = GetWindowLong(lngChild, GWL_STYLE)
                LngStyle = LngStyle And (Not HDS_BUTTONS)
                SetWindowLong lngChild, GWL_STYLE, LngStyle
                Exit Sub
            End If
        End If
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop
End Sub

Public Function CheckRushHours(ByVal strModuleNo As String, ByVal strFuncName As String) As Boolean
'���ܣ���鵱ǰʱ���Ƿ���ҵ��߷���
'������
'      strModuleNo=ģ���
'      strFuncName=��������
'���أ��Ƿ���Խ��в���

    Dim strSql As String
    Dim strTime As String
    Dim rsTemp As ADODB.Recordset
    Dim blnLimit As Boolean
    Dim dateNow As Date
    Dim strNote As String
    
    On Error GoTo errH
    CheckRushHours = True
    dateNow = CDate(Format(CurrentDate(), "HH:MM:SS"))
    strSql = "Select a.����ѡ��, a.��ʱԭ��, To_Char(b.��ʼʱ��, 'HH24:MI:SS') ��ʼʱ��, To_Char(b.����ʱ��, 'HH24:MI:SS') ����ʱ��" & vbNewLine & _
            "From Zlrunlimitset A, Zlrunlimittime B" & vbNewLine & _
            "Where a.������� = b.���� And a.ϵͳ Is Null And a.ģ�� = [1] And a.���� = [2] And b.���� = To_Char(Sysdate, 'd') - 1" & vbNewLine & _
            "Order By b.��ʼʱ��"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��ȡָ�����ܵ������ʱʱ������", strModuleNo, strFuncName)
    With rsTemp
        If .RecordCount = 0 Then Exit Function
        If Nvl(!��ʱԭ��) = "" Then
            strNote = "��ǰʱ�δ���ҵ��߷��ڣ�ʹ�ô˹��ܿ��ܻ��ϵͳʹ�����һ��Ӱ��"
        Else
            strNote = !��ʱԭ��
        End If
        Do While Not .EOF
            strTime = strTime & "��" & !��ʼʱ�� & "-" & !����ʱ��
            If dateNow > !��ʼʱ�� And dateNow < !����ʱ�� Then
                '˵����ǰʱ��������ʱ��ķ�Χ��
                blnLimit = True
            End If
            .MoveNext
        Loop
        If blnLimit = True Then
            .MoveFirst
            If !����ѡ�� = 0 Then  '������ʾ������ֹ�û����к�������
                MsgBox strNote & vbNewLine & "����ʱ�䷶Χ��" & vbNewLine & Mid(strTime, 2) & vbNewLine & "�ڽ�ֹʹ�ô˹��ܣ�", vbInformation, gstrSysName
            Else   '������ʾ��������ֹ�û����к�������
                If MsgBox(strNote & vbNewLine & "ȷ��Ҫ������", vbInformation + vbOKCancel, gstrSysName) = vbOK Then
                    blnLimit = False
                Else
                    blnLimit = True
                End If
            End If
        End If
    End With
    CheckRushHours = Not blnLimit
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Public Sub SaveAuditLog(ByVal lngType As Long, ByVal strFunction As String, ByVal strContent As String, Optional ByVal strDescription As String, Optional ByVal strModule As String)
'���ܣ�������Ҫ������־
'������lngType = �������ͣ�1-������2-�޸ģ�3-ɾ��
'      strFunction = ��������
'      strContent = ��������
'      strDescription = ����˵��
'      strModule = ����ģ��   ��һ����������ģ��ʹ��frmMDIMain.gstrLastModule���ɣ����п���һ������ģ�������һ������ģ��Ĺ��ܣ���ʱ����Ҫ�ֶ��趨һ������ģ��
    Dim strSql As String
    
    On Error GoTo errH:
    If LenB(StrConv(strContent, vbFromUnicode)) > 1024 Then
        strContent = Mid(strContent, 1, 500)
    End If
    If strModule = "" Then strModule = frmMDIMain.gstrLastModule
    strSql = "zltools.Zl_Zlauditlog_Insert('" & gstrLoginUserName & "','" & _
                                                gstrComputerName & "'," & _
                                                lngType & ",Null,'" & _
                                                strModule & "','" & _
                                                strFunction & "','" & _
                                                strContent & "','" & _
                                                strDescription & "')"
    Call ExecuteProcedure(strSql, "������Ҫ������־")
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


Public Function InstrEx(ByVal strTxt As String, ByVal strCheck As String, Optional ByVal strDeli As String) As Boolean
    '����:�Ƚ�strCheck�Ƿ������strTxt��
    'strDeli-�ַ���֮��ķָ���,Ĭ��Ϊ,
    
    If strDeli = "" Then strDeli = ","
    strTxt = strDeli & strTxt & strDeli
    strCheck = strDeli & strCheck & strDeli
    
    InstrEx = InStr(1, strTxt, strCheck) > 0
    
End Function

Public Function GetVersion() As String
'���ܣ���ȡ���ݿ�Ĵ�汾��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            GetVersion = Mid(arrTmp(1), 1, InStr(1, arrTmp(1), ".") - 1)
        End If
    End If
    
    Exit Function
errH:
    MsgBox err.Description, vbExclamation, "����"
End Function


Public Sub GetRowPos(objVsf As Object, strTxt As String, strCol As String)
'����: ���ݴ�����ַ�����λ�����
'����:strTxt-��Ҫƥ����ֶ� strCol; strCol ��Ҫƥ�����,ÿ���ֶ�֮���ö��ż�� ;objFocus-������ɺ��ȡ����Ķ���
    Dim intRow As Integer, i As Integer, j As Integer
    Dim strFiels() As String, blnResult As Boolean
    
    strFiels = Split(strCol, ",")
    blnResult = False
    '�������ݾͽ���ƥ��
    With objVsf
        '��һ��ѭ��,�ӵ�ǰ�н���ƥ��,ƥ�������һ��
        intRow = 0
        For i = .Row + 1 To .Rows - .FixedRows
            For j = 0 To UBound(strFiels)   'ѭ��ÿ����,��һ������ͼ�Ϊ��ǰ�з�������
                If (UCase(.TextMatrix(i, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(i)) = UCase(strTxt)) And .RowHidden(i) = False Then
                    blnResult = True
                    Exit For
                End If
            Next
            
            If blnResult Then '��λ����ǰ��
                intRow = i
                .Select i, 1
                .TopRow = IIf(Val(i - 10) < 0, i, i - 10)   '�����������,ȷ����λ�ڱ���м�.
                Exit Sub
            End If
        Next
        '�ڶ���ѭ��,�ӵ�һ��ƥ������ǰ��
        If .Row <> .FixedRows And intRow = 0 Then
            If MsgBox("δ�ҵ�ƥ����Ϣ,�Ƿ��ͷ����Ѱ��?", vbYesNo + vbQuestion + vbDefaultButton1, "") = vbYes Then
                For i = .FixedRows To .Row - 1
                    For j = 0 To UBound(strFiels)   'ѭ��ÿ����,��һ������ͼ�Ϊ��ǰ�з�������
                        If (UCase(.TextMatrix(i, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(i)) = UCase(strTxt)) And .RowHidden(i) = False Then
                            blnResult = True
                            Exit For
                        End If
                    Next
                    
                    If blnResult Then '��λ����ǰ��
                        intRow = i
                        .Select i, 1
                        .TopRow = IIf(Val(i - 10) < 0, i, i - 10)   '�����������,ȷ����λ�ڱ���м�.
                        Exit Sub
                    End If
                Next
            End If
        End If
        
        '���ζ�û���ҵ�,������ʾ
        If intRow = 0 Then
            For j = 0 To UBound(strFiels)   '��鵱ǰ��
                If (UCase(.TextMatrix(.Row, .ColIndex(strFiels(j)))) Like "*" & UCase(strTxt) & "*" Or UCase(.RowData(.Row)) = UCase(strTxt)) And .RowHidden(.Row) = False Then
                    blnResult = True
                    Exit For
                End If
            Next
            
            If Not blnResult Then
                MsgBox "δ�ڱ����ƥ�䵽���ݡ�", , "��ʾ"
            End If
        End If
    End With
End Sub


Public Function TranStr2Var(ByVal strTxt As String, ByVal strDeli, ByVal intLength) As Variant
'����: ������ָ�������ַ���,ת��������
    Dim varTmp As Variant, strTmp As String
    varTmp = Array()
    
    ReDim varTmp(0): varTmp(0) = strTxt
    Do While Len(strTxt) > intLength
        'ֱ��ȡָ������ǰһ���ָ�����Ϊ�������һ��Ԫ��
        strTmp = Left(strTxt, intLength)
        strTmp = Left(strTmp, InStrRev(strTmp, strDeli) - 1)
        varTmp(UBound(varTmp)) = strTmp
        
        'ԭ�ַ���ȥ����ȡ���Ĳ���
        strTxt = Mid(strTxt, Len(varTmp(UBound(varTmp))) + 2)
        
        ReDim Preserve varTmp(UBound(varTmp) + 1)
    Loop
    
    If strTxt <> "" Then
        varTmp(UBound(varTmp)) = strTxt
    End If
    
    TranStr2Var = varTmp
End Function

Public Function LoadUsers(Optional blnIncludeDBA As Boolean) As ADODB.Recordset
'����:��ȡ�û���,�������ݼ�
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select Distinct a.Username �û���, b.����, b.���� ����, b.����" & vbNewLine & _
                    "From Dba_Users A," & vbNewLine & _
                    "     (Select b.�û���, c.����, d.����, d.����" & vbNewLine & _
                    "       From �ϻ���Ա�� B, ���ű� C, ��Ա�� D, ������Ա E" & vbNewLine & _
                    "       Where b.��Աid = d.Id And e.��Աid = d.Id And e.����id = c.Id And e.ȱʡ = 1) B" & vbNewLine & _
                    "Where a.Username = b.�û���(+) And a.Account_Status Not In ('LOCKED', 'EXPIRED & LOCKED') " & vbNewLine & _
                    IIf(blnIncludeDBA, "", " And Not Exists (Select 1 From Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = Username) ") & vbNewLine & _
                    IIf(blnIncludeDBA, "", " And Not Exists (Select 1 From Dba_Sys_Privs Where Grantee = Username And Privilege = 'ADMINISTER DATABASE TRIGGER')")
                    
    Set LoadUsers = gclsBase.OpenSQLRecord(gcnOracle, strSql, "LoadUsers")
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckExist(ByVal strFields As String, ByVal strCheck As String, ByVal rsData As ADODB.Recordset) As String
'����:������ݼ����Ƿ�����ؼ�¼,��������� ,�ͷ��ز����ڵ�ֵ
'����:strFields-��Ҫ�������ֶ�,strCheck-��Ҫ�������ַ���,��","��Ϊ�ָ� ,rsData-���ݼ�
    Dim strTmp() As String, i As Integer
    Dim strResult As String
    
    strTmp = Split(strCheck, ",")
    
    For i = 0 To UBound(strTmp)
        rsData.Filter = strFields & "= '" & strTmp(i) & "'"
        If rsData.RecordCount = 0 Then
            If strResult = "" Then
                strResult = strTmp(i)
            Else
                strResult = strResult & "," & strTmp(i)
            End If
            rsData.Filter = 0
        End If
    Next
    
    CheckExist = strResult
    rsData.Filter = 0
End Function

Public Function FindUser(ByVal strUser) As String
'����:���ݴ����ֵģ����ѯ�û���,����ж�����¼,���ص�һ��,���޼�¼,���ؿ�.
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strResult As String
    
    On Error GoTo errH
    strSql = "Select Username" & vbNewLine & _
                    "From Dba_Users" & vbNewLine & _
                    "Where Username Like  '" & strUser & "%'  And Not Exists" & vbNewLine & _
                    " (Select 1 From Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = Username) And Rownum = 1"
    
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "FindUser")
    
    If rsTmp.RecordCount = 0 Then
        strResult = ""
    Else
        rsTmp.MoveFirst
        strResult = rsTmp!USERNAME
    End If
    
    FindUser = strResult
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckRAC(ByRef intInstID As Integer) As Boolean
'���ܣ�����Ƿ�ΪRAC����
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "Select Value From V$parameter Where Name = 'cluster_database'"
    On Error GoTo errH
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckRAC")
    
    If rsTmp.RecordCount > 0 And rsTmp!value = "TRUE" Then
        CheckRAC = True
        
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckRAC")
        intInstID = Val("" & rsTmp!INST_ID)
    Else
        intInstID = 0
        CheckRAC = False
    End If
    
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Function CheckSQLPlan(ByVal strSQLCheck As String, Optional ByRef vsPlan As VSFlexGrid, _
    Optional ByVal intConnect As Integer, Optional ByRef blnSuccess As Boolean) As Boolean
'����������:
'         1.���ȫ��ɨ��zlbigtable+zlbaktables��
'         2.���ͱ�ȫ��ɨ��(�����ͳ����Ϣ��User_tab_statistics:num_rows>3000(ҩƷĿ¼һ�������ֵ����) AND num_rows<100 0000��������)
'         3.��������û�����(�Ǵ��)������ϵ�����
'         4.�������ͱ�����ȫɨ�裨inex full scan��INDEX FAST FULL SCAN��
'         5.�������ͱ���Ծʽ����ɨ�裨INDEX SKIP SCAN��
'���أ�blnReturn=true ����������

    Dim rsPlan As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim j As Long, blnReturn As Boolean
    Dim rsIndex As New Recordset
    Dim rsData As ADODB.Recordset
    Dim strTable As String
    Dim rsCons_FK As New Recordset
    Dim strPar As String
    Dim strTmp As String
    Dim strAllTable As String
    
    If intConnect > 0 Then
        blnSuccess = True
        CheckSQLPlan = False
        Exit Function
    End If
    
    Set rsPlan = GetSQLPlan(strSQLCheck, intConnect)
    If Not vsPlan Is Nothing Then
        vsPlan.Redraw = flexRDNone
        vsPlan.Rows = vsPlan.FixedRows
        vsPlan.FixedAlignment(1) = flexAlignLeftCenter
    End If
    
    blnSuccess = Not rsPlan Is Nothing
    
    If Not rsPlan Is Nothing Then
        If mstrBigTable = "" Then
            '��ȡ���,�״ν����ж��Ƿ���zltables���ű�
            If mstrHasZltables = "" Then
                mstrHasZltables = CheckTblExist("ZLTABLES")
            End If
            
            '��ZLTABLES,��ȥB���C����Ϊ���,����ȡzlbigtabls��zlbaktables�еı�
            If mstrHasZltables = "True" Then
               strSql = " Select Distinct ���� From Zltables Where ���� In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3') "
            Else
                strSql = "Select Distinct ����" & vbNewLine & _
                        "From Zlbigtables" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select Distinct ���� From Zlbaktables"
            End If
           Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName)
            Do While Not rsIndex.EOF
                mstrBigTable = mstrBigTable & "," & rsIndex!����
                rsIndex.MoveNext
            Loop
            mstrBigTable = mstrBigTable & ","
        End If
        
        '��ȡ�б�ͳ����Ϣ��User_tab_statistics:num_rows>3000��
        strSql = "Select A.������,Nvl(A.����ֵ,A.ȱʡֵ) As ����ֵ " & _
                 "From zlParameters A " & _
                 "Where A.������ = '������ͱ�' And a.ϵͳ is null And a.ģ�� is null"
        Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName)
        If rsData.BOF = False Then
            strPar = Nvl(rsData("����ֵ").value, "0,0")
            If strPar <> "0,0" Then
                If strPar <> mstrMiddleTableRows Then
                    strSql = "Select Table_Name as ���� From User_Tab_Statistics Where Num_Rows > [1] And Num_Rows < [2] "
                    Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, Val(Split(strPar, ",")(0)), Val(Split(strPar, ",")(1)))
                    mstrMiddleTable = ""
                    Do While Not rsIndex.EOF
                        If InStr("," & mstrBigTable & ",", "," & rsIndex!���� & ",") = 0 Then
                            mstrMiddleTable = mstrMiddleTable & "," & rsIndex!����
                        End If
                        rsIndex.MoveNext
                    Loop
                    mstrMiddleTable = mstrMiddleTable & ","
                    mstrMiddleTableRows = strPar
                End If
            Else
                mstrMiddleTable = ""
                mstrMiddleTableRows = ""
            End If
        Else
            mstrMiddleTable = ""
            mstrMiddleTableRows = ""
        End If
        
        strAllTable = mstrMiddleTable & mstrBigTable
        
        For i = 1 To rsPlan.RecordCount
            strTmp = UCase(rsPlan!Operation & "")
            
            If Not vsPlan Is Nothing Then
                With vsPlan
                    .addItem rsPlan!Cardinality & vbTab & Trim(rsPlan!Operation) & " " & rsPlan!name & " " & IIf(rsPlan!Bytes & "" = "" And rsPlan!cost & "" = "" And rsPlan!Time & "" = "", "", " (bytes=" & rsPlan!Bytes & " cost=" & rsPlan!cost & " time=" & Format(Time / 24 / 60 / 60, "HH:MM:SS") & ")")
                    .RowOutlineLevel(.Rows - 1) = Len(rsPlan!Operation & "") - Len(LTrim(rsPlan!Operation & ""))
                    .IsSubtotal(.Rows - 1) = True
                End With
            End If
            If InStr(strTmp, "TABLE ACCESS FULL") > 0 Then
                '�ж��Ƿ��Ǵ���б�ȫɨ��
                If InStr(strAllTable, "," & rsPlan!name & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX FAST FULL SCAN") > 0 Or InStr(strTmp, "INDEX FULL SCAN") > 0 Or InStr(strTmp, "INDEX SKIP SCAN") > 0 Then
                '�ж��Ƿ��Ǵ���б�����ȫɨ�����Ծʽ����
                strTable = Split(rsPlan!name & "_", "_")(0)
                If InStr(strAllTable, "," & strTable & ",") > 0 Then
                    If Not vsPlan Is Nothing Then
                        vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
                    End If
                    blnReturn = True
                End If
            ElseIf InStr(strTmp, "INDEX RANGE SCAN") > 0 Then
                '�����ʹ���˻�����(�Ǵ��)�������
                strTable = Split(rsPlan!name & "_", "_")(0)
                
                If InStr("," & mstrBigTable & ",", "," & strTable & ",") > 0 Then
                    strSql = "Select distinct d.Table_Name, d.Index_Name, d.Column_Name,d.Column_Position" & vbNewLine & _
                        "              From User_Ind_Columns D" & vbNewLine & _
                        "              Where d.Index_Name = [1] " & vbNewLine & _
                        "              Order By d.Column_Position"
                    Set rsIndex = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, rsPlan!name & "")
                    If rsIndex.RecordCount > 0 Then
                        '�����Լ��
                        Set rsCons_FK = GetConsFK(strTable, rsPlan!object_owner & "")
                        strTmp = ""
                        Do While Not rsIndex.EOF
                            strTmp = strTmp & "," & rsIndex!Column_Name
                            rsIndex.MoveNext
                        Loop
                        rsCons_FK.Filter = "Column_Name='" & Mid(strTmp, 2) & "'"
                        If rsCons_FK.RecordCount > 0 Then
                            strTable = Split(rsCons_FK!r_Constraint_Name & "_", "_")(0)
                            
                            '��������Ǵ������Ϊ����������
                            If InStr(mstrBigTable, "," & strTable & ",") = 0 Then
                                If Not vsPlan Is Nothing Then
                                    vsPlan.Cell(flexcpForeColor, vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1) = &HFF& '��
                                End If
                                blnReturn = True
                            End If
                        End If
                    End If
                End If
            End If
            
            rsPlan.MoveNext
        Next
        
        If Not vsPlan Is Nothing Then
            vsPlan.CellBorderRange 0, 0, vsPlan.Rows - 1, 0, &H808080, 0, 0, 1, 0, 0, 0
            vsPlan.CellBorderRange vsPlan.FixedRows - 1, 0, vsPlan.FixedRows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.CellBorderRange vsPlan.Rows - 1, 0, vsPlan.Rows - 1, vsPlan.Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            vsPlan.AutoSize 0, vsPlan.Cols - 1
            vsPlan.Redraw = flexRDDirect
        End If
    End If
    
    CheckSQLPlan = blnReturn
End Function

Private Function GetSQLPlan(ByVal strSQLCheck As String, Optional ByVal intConnect As Integer = 0) As ADODB.Recordset
'���ܣ��ռ�SQL��ִ�мƻ�

    Dim strSql As String, strSID As String
    Dim rsTmp As ADODB.Recordset
    
    If strSQLCheck <> "" Then
        
        On Error Resume Next
        strSID = Time()
          
        'ִ�мƻ�
        strSql = "explain plan set statement_id = '" & strSID & "' for " & strSQLCheck
        gcnOracle.Execute strSql
        If err.Number = 0 Then
            strSql = _
                    "Select Time From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            On Error Resume Next
            Set GetSQLPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, "ִ�мƻ�", strSID)
            strSql = _
                    "Select ID, LPad(' ', Level - 1) || Operation || ' ' || Options As Operation, Object_Name As Name" & _
                    "    ,Object_Owner, Cardinality, Bytes" & vbNewLine & _
                    "    ,Cost" & IIf(err.Number = 0, ", Time ", ",0 as Time ") & vbNewLine & _
                    "From Plan_Table " & vbNewLine & _
                    "Connect By Prior ID = Parent_Id And Prior Statement_Id = Statement_Id " & vbNewLine & _
                    "Start With ID = 0 And Statement_Id = [1] " & vbNewLine & _
                    "Order By ID "
            err.Clear
            Set GetSQLPlan = gclsBase.OpenSQLRecord(gcnOracle, strSql, "ִ�мƻ�", strSID)
            gcnOracle.Execute "Delete plan_table"
        Else
            Set GetSQLPlan = Nothing
           MsgBox err.Description: err.Clear
        End If
    End If
End Function



Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '���ܣ����ݱ����жϱ��Ƿ����
    '������strTableName - Ҫ��ѯ�ı���
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error Resume Next
    strSql = "select 1 from " & strTableName & " where rownum<1 "
    Set rsData = gclsBase.OpenSQLRecord(gcnOracle, strSql, "CheckTblExist")
    
    CheckTblExist = err.Number = 0
    err.Clear
End Function

Public Function CheckAuditStatus(ByVal strModuleNo As String, ByVal strFuncName As String, ByRef strRemarks As String) As Boolean
    '���ܣ���鴫��Ĺ����Ƿ���Ҫ�������
    'strModuleNo = ģ����
    'strFuncName = ��������
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select �Ƿ������ From Zlauditlogconfig Where ϵͳ Is Null And ģ�� = [1] And ���� = [2]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, "��ȡ��Ӧģ�鹦���Ƿ���Ҫ���", strModuleNo, strFuncName)
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "����Ա�����֤ʧ�ܣ�����ϵ������Ա���ģ�鹦�������Ƿ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If !�Ƿ������ = 1 Then
            If Not frmUserCheckLogin.ShowLogin(UCT_AuditLog, , gstrUserName, , , , strRemarks) Then Exit Function
        Else
            strRemarks = ""
        End If
    End With
    CheckAuditStatus = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function GetConsFK(ByVal strFind As String, ByVal strOwner As String) As ADODB.Recordset
'���ܣ���ȡָ��������Լ����¼��
'������strFind=����
    Dim strSql As String
    Dim rsCons As New Recordset
    Dim rsCons_FK As New Recordset

    strSql = "Select" & vbNewLine & _
        "        f.Constraint_Name, f.r_Constraint_Name,e.Column_Name,e.Position" & vbNewLine & _
        "       From User_Cons_Columns E, User_Constraints F" & vbNewLine & _
        "       Where e.Constraint_Name = f.Constraint_Name And e.owner = f.owner  And f.Constraint_Type = 'R' And f.Table_Name = [1] And f.owner = [2]" & vbNewLine & _
        "       order by e.constraint_name,e.position"
    Set rsCons = gclsBase.OpenSQLRecord(gcnOracle, strSql, App.ProductName, strFind, strOwner)
    Set rsCons_FK = New ADODB.Recordset
    rsCons_FK.Fields.Append "r_Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Constraint_Name", adVarChar, 50, adFldIsNullable
    rsCons_FK.Fields.Append "Column_Name", adVarChar, 100, adFldIsNullable
    rsCons_FK.CursorLocation = adUseClient
    rsCons_FK.LockType = adLockOptimistic
    rsCons_FK.CursorType = adOpenStatic
    rsCons_FK.Open
    Do While Not rsCons.EOF
        rsCons_FK.Filter = "Constraint_Name='" & rsCons!Constraint_Name & "'"
        If rsCons_FK.RecordCount = 0 Then
            rsCons_FK.AddNew
            rsCons_FK!Constraint_Name = rsCons!Constraint_Name & ""
            rsCons_FK!r_Constraint_Name = rsCons!r_Constraint_Name & ""
            rsCons_FK!Column_Name = rsCons!Column_Name & ""
        Else
            rsCons_FK!Column_Name = rsCons_FK!Column_Name & "," & rsCons!Column_Name
        End If
        rsCons_FK.Update
        rsCons.MoveNext
    Loop
    Set GetConsFK = rsCons_FK
End Function

Public Function CheckIpValidate(ByVal strBeginIp As String, Optional ByVal strEndIp As String, Optional ByRef strErr As String) As Boolean
    '���IP�ĺϷ���
    'strBeginIp -��ʼIP strEndIp-����IP strErr-������Ϣ
    Dim arrStart As Variant, arrEnd As Variant
    Dim i As Integer
    
    If Not IsNumeric(Replace(strBeginIp, ".", "")) Then
        strErr = "IP����Ϊ����"
        Exit Function
    End If
    arrStart = Split(strBeginIp, ".")
    If UBound(arrStart) <> 3 Then
        strErr = "IP������4��IP�����"
        Exit Function
    End If

    If strEndIp <> "" Then
        If Not IsNumeric(Replace(strEndIp, ".", "")) Then
            strErr = "IP����Ϊ����"
            Exit Function
        End If
        arrEnd = Split(strEndIp, ".")
        If UBound(arrEnd) <> 3 Then
            strErr = "IP������4��IP�����"
            Exit Function
        End If
    End If
    
'    A��IP��1.0.0.0-126.0.0.255
'    B��IP��128.1.0.0--191.254.0.255
'    C��IP��192.0.1.0--223.255.254.255
'    D��IP��224.0.0.0��239.255.255.255
    '��һ��
    If arrStart(0) >= 1 And arrStart(0) <= 239 Then
        If arrEnd(0) <> "" And arrEnd(0) <> arrStart(0) Then
            strErr = "��ʼIP�����IP���׶α�����ͬ"
            Exit Function
        End If
        
        '�ڶ���
        If arrStart(1) >= 0 And arrStart(1) <= 255 Then
            If arrEnd(1) <> "" And arrEnd(1) <> arrStart(1) Then
                strErr = "��ʼIP�����IP�Ĵζα�����ͬ"
                Exit Function
            End If
        Else
            strErr = "IP�Ĵζ�ֻ�ܽ���0-255֮��"
            Exit Function
        End If
        
        '������
        If arrStart(2) >= 0 And arrStart(2) <= 255 Then
            If arrEnd(2) <> "" Then
                If arrEnd(2) >= 0 And arrEnd(2) <= 255 Then
                    If arrEnd(2) < arrStart(2) Then
                        strErr = "����IP�ĵ����α�����ڻ���ڿ�ʼIP�ĵ�����"
                        Exit Function
                    End If
                Else
                    strErr = "IP�ĵ�����ֻ�ܽ���0-255֮��"
                    Exit Function
                End If
            End If
        Else
            strErr = "IP�ĵ�����ֻ�ܽ���0-255֮��"
            Exit Function
        End If
        
        '���Ķ�
        If arrStart(3) > 0 And arrStart(3) <= 255 Then
            If arrEnd(3) <> "" Then
                If arrEnd(3) > 0 And arrEnd(3) <= 255 Then
                    If arrEnd(3) < arrStart(3) Then
                        strErr = "����IP�ĵ��Ķα�����ڻ���ڿ�ʼIP�ĵ��Ķ�"
                        Exit Function
                    End If
                Else
                    strErr = "IP�ĵ��Ķ�ֻ�ܽ���1-255֮��"
                    Exit Function
                End If
            End If
        Else
            strErr = "IP�ĵ��Ķ�ֻ�ܽ���1-255֮��"
            Exit Function
        End If
        
    Else
        strErr = "IP�׶�ֻ�ܽ���1-239֮��"
        Exit Function
    End If
    
    CheckIpValidate = True
End Function


Public Function CheckProcExist(ByVal strProc As String) As Integer
    '����:���ݴ���Ľ�������,�����������еĽ�����

    Dim intResult As Integer
    Dim uProcess As PROCESSENTRY32
    Dim lngMdlProcess As Long, strExeName As String, lngSnapShot As Long
    
    '�������̿���
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If lngSnapShot > 0 Then
        uProcess.dwSize = Len(uProcess)
        If Process32First(lngSnapShot, uProcess) Then
            Do
                strExeName = UCase(Left(Trim(uProcess.szExeFile), InStr(1, Trim(uProcess.szExeFile), vbNullChar) - 1))
                If strExeName = UCase(strProc) Then
                    intResult = intResult + 1
                End If
            Loop Until (Process32Next(lngSnapShot, uProcess) < 1)
        End If
    End If
    
    CheckProcExist = intResult
End Function

Public Sub ShowTipInfo(ByVal lngHwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long, Optional strTitle As String, Optional blnChild As Boolean)
'���ܣ���ʾ����������ʾ
'������lngHwnd=��ʾ����ԵĿؼ����,������Ϊ0ʱ������ʾ
'      strInfo=��ʾ��Ϣ,������Ϊ��ʱ������ʾ
'      blnMultiRow=��һ���ļ�������ʾ������Ϣ��ÿ�а�vbcrlf�ָ�
'      blnOutline=�Ƿ�ÿ���ı����ַ�|ǰ��������Ϊ��ٵ���һ����ʾ
'      lngMaxWidth=���ڵ���󴰶ȣ�ȱʡΪ0��ʾ�����״̬�Ĵ��������Ϊ׼
'      strTitle = ��ʾ����
'      blnChild=�Ƿ�ʹ��ChildWindowFromPoint����

    Call frmTipInfo.ShowTipInfo(lngHwnd, strInfo, blnMultiRow, blnOutline, lngMaxWidth, strTitle, blnChild)
End Sub

