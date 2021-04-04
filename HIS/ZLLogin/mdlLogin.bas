Attribute VB_Name = "mdlLogin"
Option Explicit

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000 'Forces a top-level window onto the taskbar when the window is visible.ǿ��һ���ɼ��Ķ����Ӵ�����������
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1

Public gstrSysName As String
Public gdtStart As Long
Public gobjRegister As Object               'ע����Ȩ����zlRegister
Public gcnOracle As ADODB.Connection     '�������ݿ�����
Public gstrCommand As String '������

Public gobjFile As New FileSystemObject
Public gclsLogin As clsLogin '��¼����
Public gintCallType As Integer '0-��չʾ�޸����������������,1-��ʾ�޸�����,2-��ʵ����������
Public gblnExitApp  As Boolean '�Ƿ���Ϊ�ظ����У���Ҫ�˳���������

'clsLogin���Ի���
Public gobjEmr             As Object   'EMR�°���Ӳ���
Public gstrUserName        As String   'InputUser����
Public gstrInputPwd        As String   'InputPwd����
Public gstrServerName      As String   'ServerName����
Public gstrDBUser          As String   'DBUser����
Public gblnTransPwd        As Boolean  'blnTransPwd����
Public gblnSysOwner        As Boolean  '�Ƿ�ϵͳ������
Public gstrConnString      As String   '�����ַ���
Public gstrSystems         As String   '������ѡ���ϵͳ
Public gblnCancel          As Boolean  '�Ƿ�ȡ���˳�
Public gstrMenuGroup       As String   '�˵�������
Public gstrDeptName        As String   '�û���¼��������
Public gstrStation         As String   '�û���¼����վ����
Public gstrNodeNo          As String   'վ����
Public gstrNodeName        As String   'վ������
Public gblnEMRProxy         As Boolean
Public gstrEMRPwd           As String
Public gstrEMRUser          As String

Public gblnTimer            As Boolean  '�Ƿ�ʱ�������Ŀͻ��˸��¼��
Public glngInstanceCount    As Long     'ʵ������

Public Sub SetAppBusyState()
'���������̶���δ�������ʱ���滻��ִ�������̹���ʱ�����ġ����������𡱶Ի���
    On Error Resume Next
    App.OleServerBusyMsgTitle = App.ProductName
    App.OleRequestPendingMsgTitle = App.ProductName
    
    App.OleServerBusyMsgText = "���������ڴ����������ĵȴ���"
    App.OleRequestPendingMsgText = "�������������������ĵȴ���"
    
    App.OleServerBusyTimeout = 3000
    App.OleRequestPendingTimeout = 10000
    Err.Clear
End Sub

Public Function ShowSplash(Optional ByVal blnRefresh As Boolean) As Boolean
    Dim strUnitName As String, intCount As Integer
    Dim objPic As IPictureDisp
    '��ע����л�ȡ�û�ע�������Ϣ,����û���λ���Ʋ�Ϊ��,����ʾ���ִ���
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��ʾ", "")
    strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    If blnRefresh Then
        With frmSplash
            .lblGrant = Replace(strUnitName, ";", vbCrLf)
            .lbl����֧����.Caption = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
            
            .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
            .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
            strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
            .lbl������.Caption = ""
            For intCount = 0 To UBound(Split(strUnitName, ";"))
                .lbl������.Caption = .lbl������.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
            Next
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            If gobjFile.FileExists(gstrSetupPath & "\�����ļ�\logo_login.jpg") Then
                Set objPic = LoadPicture(gstrSetupPath & "\�����ļ�\logo_login.jpg")
                .picHos.Visible = True
                .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183����
                .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323����
                .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
            Else
                .picHos.Visible = False
            End If
            If InStr(gstrCommand, "=") <= 0 Then .Show
            ShowSplash = True
        End With
    Else
        If strUnitName <> "" And strUnitName <> "-" Then
            gdtStart = Timer
            With frmSplash
                '��������Ҫ����
                '��ʱ�Ϳ�ʼ����clsComLib��ʵ��
                Call ApplyOEM_Picture(.ImgIndicate, "Picture")
                Call ApplyOEM_Picture(.imgPic, "PictureB")
                If gobjFile.FileExists(gstrSetupPath & "\�����ļ�\logo_login.jpg") Then
                    Set objPic = LoadPicture(gstrSetupPath & "\�����ļ�\logo_login.jpg")
                    .picHos.Visible = True
                    .picHos.Height = IIf(objPic.Height < 2745, objPic.Height, 2745) '183����
                    .picHos.Width = IIf(objPic.Width < 4845, objPic.Width, 4845) '323����
                    .picHos.PaintPicture objPic, 0, 0, .picHos.Width, .picHos.Height
                Else
                    .picHos.Visible = False
                End If
                If InStr(gstrCommand, "=") <= 0 Then .Show
                
                .lblGrant = Replace(strUnitName, ";", vbCrLf)
                strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "������", "")
                If Trim(strUnitName) = "" Then
                    .Label3.Visible = False
                    .lbl������.Visible = False
                Else
                    .Label3.Visible = True
                    .lbl������.Visible = True
                    .lbl������.Caption = ""
                    For intCount = 0 To UBound(Split(strUnitName, ";"))
                        .lbl������.Caption = .lbl������.Caption & Split(strUnitName, ";")(intCount) & vbCrLf
                    Next
                End If
                .LblProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒȫ��", "")
                If Len(.LblProductName) > 10 Then
                    .LblProductName.FontSize = 15.75 '����
                Else
                    .LblProductName.FontSize = 21.75 '����
                End If
                .lbl����֧���� = GetSetting("ZLSOFT", "ע����Ϣ", "����֧����", "")
                .lbltag = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒϵ��", "")
                
                If Trim$(.lbl����֧����.Caption) = "" Then
                    .Label1.Visible = False
                    .lbl����֧����.Visible = False
                Else
                    .Label1.Visible = True
                    .lbl����֧����.Visible = True
                End If
            End With
            Do
                If (Timer - gdtStart) > 1 Then Exit Do
                DoEvents
            Loop
            
            ShowSplash = True
        End If
    End If
End Function

Public Function SaveRegInfo() As Boolean
    Dim strTag As String, strTitle As String
    
    Select Case zlRegInfo("��Ȩ����")
        Case "1"
            '��ʽ
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", ""
        Case "2"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case "3"
            '����
            SaveSetting "ZLSOFT", "ע����Ϣ", "Kind", "����"
        Case Else
            '����
            MsgBox "��Ȩ���ʲ���ȷ���������˳���", vbInformation, gstrSysName
            Exit Function
    End Select
    
    gstrSysName = zlRegInfo("��Ʒ����") & "���"
    SaveSetting "ZLSOFT", "ע����Ϣ", "��ʾ", gstrSysName
    SaveSetting "ZLSOFT", "ע����Ϣ", UCase("gstrSysName"), gstrSysName
    strTag = ""
    strTitle = zlRegInfo("��Ʒ����")
    If strTitle <> "" Then
        If InStr(strTitle, "-") > 0 Then
            If Split(strTitle, "-")(1) = "Ultimate" Then
                strTag = "�콢��"
            ElseIf Split(strTitle, "-")(1) = "Professional" Then
                strTag = "רҵ��"
            End If
        End If
    End If
    strTitle = Split(strTitle, "-")(0)
    '���û�ע�������Ϣд��ע���,���´�����ʱ��ʾ
    SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", zlRegInfo("��λ����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒȫ��", strTitle
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒ����", zlRegInfo("��Ʒ����")
    SaveSetting "ZLSOFT", "ע����Ϣ", "����֧����", zlRegInfo("����֧����", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "������", zlRegInfo("��Ʒ������", , -1)
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧���̼���", zlRegInfo("֧���̼���")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��EMAIL", zlRegInfo("֧����MAIL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "WEB֧��URL", zlRegInfo("֧����URL")
    SaveSetting "ZLSOFT", "ע����Ϣ", "��Ʒϵ��", strTag
    SaveRegInfo = True
End Function

Public Function TestComponent() As Boolean
    '���û���κβ�����ʹ�ã��򷵻ؼ�
    TestComponent = False
    
    Dim strObjs As String, strCodes As String, strSQL As String
    Dim objComponent As Object
    Dim resComponent As New ADODB.Recordset
    
    On Error GoTo errH
    '--��ע����ȡ��Ȩ����--
    strObjs = GetSetting("ZLSOFT", "ע����Ϣ", "��������", "")
    If strObjs <> "" Then
        If InStr(strObjs, "'ZL9REPORT'") = 0 Then
            If CreateComponent("ZL9REPORT.ClsREPORT") Then
                strObjs = strObjs & ",'ZL9REPORT'"
                SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
            End If
        End If
        TestComponent = True
        Exit Function
    End If
    '--������Ȩ��װ����--
    strSQL = "Select Distinct ���� From (" & _
                " Select Upper(g.����) As ����" & _
                " From zlPrograms g, zlRegFunc r" & _
                " Where g.��� = r.��� And Trunc(g.ϵͳ / 100) = r.ϵͳ" & _
                " Union " & _
                " Select Upper(����) as ���� From zlPrograms Where ��� Between 10000 And 19999)"
    Set resComponent = zlDatabase.OpenSQLRecord(strSQL, "")
    With resComponent
        Do While Not .EOF
            If CreateComponent(!���� & ".Cls" & Mid(!����, 4)) Then
                strObjs = strObjs & IIf(strObjs = "", "", ",") & "'" & !���� & "'"
            End If
            .MoveNext
        Loop
    End With
    If strObjs = "" Then Exit Function
    TestComponent = True
    SaveSetting "ZLSOFT", "ע����Ϣ", "��������", strObjs
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CreateComponent(StrComponent) As Boolean
    Dim objComponent        As Object
On Error GoTo errH
    Set objComponent = CreateObject(StrComponent)
    CreateComponent = True
    Exit Function
errH:
    Err.Clear
    CreateComponent = False
    Exit Function
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

Public Function CreateRegister() As Boolean
    '����ע�Ჿ��(���ڵ�¼ʱ��ȡ���Ӷ���)
    On Error Resume Next
    Set gobjRegister = CreateObject("zlRegister.clsRegister")
    If gobjRegister Is Nothing Then
        Err.Clear
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = True
End Function

Public Function CheckPWDComplex(ByRef cnInput As ADODB.Connection, ByVal strChcekPWD As String, Optional ByRef strToolTip As String) As Boolean
'���ܣ�������븴�Ӷ�
'������cnInput=���������
'          strChcekPWD=�ȴ���������
'          strToolTip=�����ʾ����
'���أ�True-���ɹ���False-���ʧ��
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim lngLen As Long, i As Integer, intChr As Integer
    
    On Error GoTo errH
    strToolTip = ""
    strSQL = "Select ������,Nvl(����ֵ,ȱʡֵ) ����ֵ From zlOptions Where ������ in (20,21,22,23)"
    rsData.Open strSQL, cnInput
    blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
    blnComplex = False: strOterChrs = ""
    Do While Not rsData.EOF
        Select Case rsData!������
            Case 20 '�Ƿ�������볤��
                blnPwdLen = Val(rsData!����ֵ & "") = 1
            Case 21 '���볤������
                intPwdMin = Val(rsData!����ֵ & "")
            Case 22 '���볤������
                intPwdMax = Val(rsData!����ֵ & "")
            Case 23 '�Ƿ�������븴�Ӷ�
                blnComplex = Val(rsData!����ֵ & "") = 1
        End Select
        rsData.MoveNext
    Loop
    '����������ʾ
    If blnPwdLen Then
        If intPwdMin = intPwdMax Then
            strToolTip = "�������Ϊ" & intPwdMax & " λ�ַ���"
        Else
            strToolTip = "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���"
        End If
     End If
     If blnComplex Then
        If strToolTip <> "" Then
            strToolTip = strToolTip & vbNewLine & "���ٰ���һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        Else
            strToolTip = "������һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
        End If
     End If
    '���ȼ��
    lngLen = zlStr.ActualLen(strChcekPWD)
    If lngLen <> Len(strChcekPWD) Then
        MsgBox "���������˫�ֽ��ַ������飡", vbInformation, gstrSysName
        Exit Function
    End If
    If blnPwdLen Then
        If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
            If intPwdMin = intPwdMax Then
                MsgBox "�������Ϊ" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                Exit Function
            Else
                MsgBox "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    For i = 1 To Len(strChcekPWD)
        intChr = Asc(UCase(Mid(strChcekPWD, i, 1)))
        If intChr >= 32 And intChr < 127 Then
            'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
            Select Case intChr
                Case 48 To 57 '����
                    blnHaveNum = True
                Case 65 To 90 '��ĸ
                    blnAlpha = True
                Case 32, 34, 47, 64  '�ո�,˫����,/,@
                    strOterChrs = strOterChrs & Chr(intChr)
                Case Is < 48, 58 To 64, 91 To 96, Is > 122
                    blnChar = True
            End Select
        Else
            strOterChrs = strOterChrs & Chr(intChr)
        End If
    Next
    If strOterChrs <> "" Then
        MsgBox "���벻�����������ַ���" & strOterChrs, vbInformation, gstrSysName
        Exit Function
    ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
        MsgBox "����������һ�����֡�һ����ĸ��һ�������ַ���ɡ�", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPWDComplex = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox Err.Description, vbInformation, gstrSysName
End Function

Public Function CheckSysState() As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnHaveTools As Boolean, blnDBA As Boolean
    
    On Error Resume Next
    strSQL = "SELECT 1 FROM ZLTOOLS.ZLSYSTEMS WHERE ������=USER"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������")
    
    If Err.Number <> 0 Then
        blnHaveTools = False
        gclsLogin.IsSysOwner = False
        Err.Clear
    Else
        blnHaveTools = True
        gclsLogin.IsSysOwner = rsTmp.EOF
    End If

    strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ж�DBA")
    blnDBA = Not rsTmp.EOF

    If Not (blnDBA) And Not (blnHaveTools) Then
        CheckSysState = False
        MsgBox "�д��������������ߣ����Ƚ��д�����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If Not (blnDBA) And Not (gclsLogin.IsSysOwner) Then
        CheckSysState = False
        MsgBox "�������ݿ�DBA��Ӧ��ϵͳ�������ߣ�����ʹ�ñ����ߡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    If Not blnHaveTools Then
        CheckSysState = False
        MsgBox "�д��������������ߣ����Ƚ��д�����", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckSysState = True
End Function

Public Function GetMenuGroup(ByVal strCommand As String) As String
    Dim ArrCommand As Variant
    '--����Ȩ�޲˵�--
    If strCommand = "" Then
        GetMenuGroup = "ȱʡ"
    Else
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) = 0 Then
            '���������˵�����������/����ʾ���û�������ĸ�ʽ���磺zlhis/his��
            If InStr(1, ArrCommand(0), "/") = 0 And InStr(ArrCommand(0), ",") = 0 Then
                GetMenuGroup = ArrCommand(0)
            Else
                GetMenuGroup = "ȱʡ"
            End If
        Else
            '�û��������뼰�˵����
            If UBound(ArrCommand) = 2 And InStr(ArrCommand(0), "=") <= 0 Then
                GetMenuGroup = ArrCommand(2)
            Else
                GetMenuGroup = "ȱʡ"
            End If
        End If
    End If
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    If gblnTimer Then
        Set OpenSQLRecord = zlDatabase.OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    Else
        Set OpenSQLRecord = OpenSQLRecordByArray(strSQL, strTitle, arrPars)
    End If
End Function

Private Function OpenSQLRecordByArray(ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'               ��Ϊʹ�ð󶨱���,�Դ�"'"���ַ�����,����Ҫʹ��"''"��ʽ��
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'      cnOracle=����ʹ�ù�������ʱ����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim lngErrNum As Long, strErrInfo As String
    
    '������ʹ���˶�̬�ڴ������û��ʹ��/*+ XXX*/����ʾ��ʱ�Զ�����
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '�ж�ǰ���Ƿ�����IN �����򲻼�Rule
                '���ҵ����һ��SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(zlStr.FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  'ȡ����3���ַ�
                
                If strTmp = "IN(" Then '����in(select��������������ѭ�������Ƿ����û��ʹ������д����������̬�ڴ溯��
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL���󶨱�����ȫ��������Դ��" & strTitle
    End If

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next
'    If gblnSys = True Then
'        Set cmdData.ActiveConnection = gcnSysConn
'    Else
    Set cmdData.ActiveConnection = gcnOracle '���Ƚ���(���ִ��1000��Լ0.5x��)
'    End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'���ܣ�ִ�й������,���Զ��Թ��̲������а󶨱�������
'������strSQL=�������,���ܴ�����,����"������(����1,����2,...)"��
'      cnOracle=����ʹ�ù�������ʱ����
'˵�������¼���������̲�����ʹ�ð󶨱���,�����ϵĵ��÷�����
'  1.���������Ǳ��ʽ,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1,100.12*0.15,...)"
'  2.�м�û�д�����ȷ�Ŀ�ѡ����,��ʱ�����޷�����󶨱������ͺ�ֵ,��"������(����1, , ,����3,...)"
'  3.��Ϊ�ù������Զ�����,����һ��ʹ�ð󶨱���,�Դ�"'"���ַ�����,��Ҫʹ��"''"��ʽ��
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    Dim lngErrNum As Long, strErrInfo As String
    
    If Right(Trim(strSQL), 1) = ")" Then
        'ִ�еĹ�����
        strTemp = Trim(strSQL)
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
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '�ַ���
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle���ӷ�����:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '˫"''"�İ󶨱�������
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '���Ӳ�������LOBʱ������ð󶨱���ת��ΪRAWʱ����2000���ַ�Ҫ��adLongVarChar
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
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
                        If datCur = CDate(0) Then datCur = Currentdate
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
        
        '����Ա���ù���ʱ��д����
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "���� Oracle ����""" & strProc & """ʱ�����Ż�������д��ƥ�䡣ԭʼ������£�" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '����?��
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    '˵����Ϊ�˼��������ӷ�ʽ
    '1.��������adCmdStoredProc��ʽ��8i����������
    '2.�����������ʹ��{},��ʹ����û�в���ҲҪ��()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
'    Call gobjComLib.SQLTest
End Sub

Public Function IP(Optional ByVal strErr As String) As String
    '******************************************************************************************************************
    '����:ͨ��oracle��ȡ�ļ������IP��ַ
    '���:strDefaultIp_Address-ȱʡIP��ַ
    '����:
    '����:����IP��ַ
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim strIp_Address As String
    Dim strSQL As String
        
    On Error GoTo errHand
    
    strSQL = "Select Sys_Context('USERENV', 'IP_ADDRESS') as Ip_Address From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡIP��ַ")
    If rsTmp.EOF = False Then
        strIp_Address = NVL(rsTmp!Ip_Address)
    End If
    If strIp_Address = "" Then strIp_Address = OS.IP(strErr)
    If Replace(strIp_Address, " ", "") = "0.0.0.0" Then strIp_Address = ""
    IP = strIp_Address
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    strErr = strErr & IIf(strErr = "", "", "|") & Err.Description
    Err.Clear
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '���ܣ���ȡ�������ϵ�ǰ����
    '������
    '���أ�����Oracle���ڸ�ʽ�����⣬����
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngErrNum As Long, strErrInfo As String
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    Currentdate = 0
    Err = 0
End Function
