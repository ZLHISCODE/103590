VERSION 5.00
Begin VB.Form frmUserCheckLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û���֤"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmUserCheckLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4875
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtRemarks 
      Height          =   840
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "�ñ�ע��������128�����ֻ�256���ַ�"
      Top             =   1710
      Width           =   3495
   End
   Begin VB.CommandButton cmdReloadSvr 
      Caption         =   "ˢ�·�����(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   216
      TabIndex        =   10
      Top             =   2256
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.ComboBox cboServer 
      Height          =   276
      Left            =   1716
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1716
      Width           =   2592
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   1992
      Width           =   5000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2268
      TabIndex        =   6
      Top             =   2256
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3528
      TabIndex        =   7
      Top             =   2256
      Width           =   1100
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1716
      MaxLength       =   30
      TabIndex        =   1
      Top             =   900
      Width           =   2592
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1716
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1308
      Width           =   2592
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      Caption         =   "����˵��"
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   1770
      Width           =   720
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   210
      Picture         =   "frmUserCheckLogin.frx":1CFA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1092
      TabIndex        =   4
      Top             =   1776
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1092
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1272
      TabIndex        =   2
      Top             =   1368
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "    ������""Rac1(testbase)""�ķ��������������û���֤"
      Height          =   360
      Left            =   1140
      TabIndex        =   9
      Top             =   240
      Width           =   3552
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUserCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrServer As String
Private mcnOracle As ADODB.Connection '��֤�û�������
Private muctCurType As Integer
Private mstrSystems As String
Private mblnFirst As Boolean  'ΪTrue��ʾ�Ѿ�������ʾ��
Private mintTimes As Integer  '��¼���Դ���
Private mcolServer As New Collection
Private mblnOk As Boolean
Private mstrRacInfo As String 'RAC����Ϣ
Private mstrRemarks As String '��¼��ע��Ϣ

Public Function ShowLogin(Optional ByVal uctType As UserCheckType, Optional ByRef cnOracle As ADODB.Connection, _
                        Optional ByRef strUser As String, Optional ByVal strServer As String, Optional ByVal strSystems As String, _
                        Optional ByVal strRacInfo As String, Optional ByRef strRemarks As String) As Boolean
'���ܣ���֤�û���¼
'������
'          cnOracle=���ص�����
'          strUser=��֤���û�
'          strSystems=��ͨ�û���֤�� uctType=UCT_NormalUser��ʱ�޶��û�����ϵͳ��
'          strRacInfo=Rac��֤ʱ��RAC��ʶ��Ϣ�� uctType=UCT_RACInsUser��,��ʽΪ��INST_ID,DBID,Instance_Name(DBname)
'          strRemarks=��ע(uctType = UCT_AuditLog����Ҫ����ִ����Ҫ������֤���ʱ���뱸ע)
'˵������ͨ�û���¼ʱ��ϵͳ�������û��������ݿ�ʱ����֤���û���������벻�����ݿ�����
    muctCurType = uctType
    mstrUser = Decode(uctType, UCT_ZLTOOLS, "ZLTOOLS", strUser)
    mstrServer = IIf(strServer = "" And uctType <> UCT_RACInsUser, gstrServer, strServer)
    mstrRacInfo = strRacInfo
    mstrSystems = strSystems
    mstrRemarks = strRemarks
    Me.Show 1
    Set cnOracle = mcnOracle
    If uctType = UCT_NormalUser Or uctType = UCT_SysOwner Then
        strUser = mstrUser
    End If
    If uctType = UCT_AuditLog Then
        If Not mcnOracle Is Nothing Then
            mcnOracle.Close
            Set mcnOracle = Nothing
        End If
        strRemarks = mstrRemarks
        mstrRemarks = ""
    End If
    ShowLogin = mblnOk
    mblnOk = False
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Set mcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNote As String, strRemarks As String
    Dim strUser As String, strPwd As String, strServer As String
    Dim intPos As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    
    SetConState False
    If muctCurType <> UCT_AuditLog Then
        mintTimes = mintTimes + 1
    End If
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUser = Trim(txtUser.Text)
    strPwd = Trim(txtPWD.Text)
    strServer = Trim(cboServer.Text)
    strRemarks = Trim(txtRemarks.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "�������û�����"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUser) <> 1 Then
        If Mid(strUser, 1, 1) = "/" Or Mid(strUser, 1, 1) = "@" Or Mid(strUser, Len(strUser) - 1, 1) = "/" Or Mid(strUser, Len(strUser) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(strPwd) <> "" And Len(strPwd) <> 1 Then
        If Mid(strPwd, Len(strPwd) - 1, 1) = "/" Or Mid(strPwd, Len(strPwd) - 1, 1) = "@" Or Mid(strPwd, 1, 1) = "/" Or Mid(strPwd, 1, 1) = "@" Then
            txtPWD.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServer) <> "" Then
        If Mid(strServer, Len(strServer) - 1, 1) = "/" Or Mid(strServer, Len(strServer) - 1, 1) = "@" Or Mid(strServer, 1, 1) = "/" Or Mid(strServer, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If

    '�����ַ���
    intPos = InStr(1, strUser, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUser, "/", vbTextCompare)
    If intPos > 0 Then
        strPwd = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPwd, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strPwd, intPos + 1)
        strPwd = Mid(strPwd, 1, intPos - 1)
    End If
    
    If Len(Trim(strPwd)) = 0 And (muctCurType <> UCT_AuditLog Or gstrLoginUserName <> gstrUserName) Then
        strNote = "����������"
        txtPWD.SetFocus
        GoTo InputError
    End If
    
    If strRemarks = "" And muctCurType = UCT_AuditLog Then
        strNote = "�����뱸ע"
        txtRemarks.SetFocus
        GoTo InputError
    ElseIf strRemarks <> "" Then
        If StrIsValid(txtRemarks.Text, 256) = False Then
            txtRemarks.SetFocus
            SetConState
            Exit Sub
        End If
    End If
    strUser = UCase(strUser)
    
    If muctCurType <> UCT_AuditLog Or gstrLoginUserName <> gstrUserName Then
        If Not OracleOpen(strServer, strUser, strPwd, strNote) Then
            txtPWD.Text = ""
            If txtPWD.Enabled Then txtPWD.SetFocus
            SetConState
            If strNote <> "" Then GoTo InputError
            Exit Sub
        End If
    End If
    
    Select Case muctCurType
        Case UCT_ZLTOOLS
            gstrToolsPwd = strPwd
            Set gcnTools = mcnOracle
        Case UCT_CurZLBAK
        Case UCT_DBAUser
            strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "DBA�ж�")
            If rsTmp.EOF Then
                MsgBox "��ǰ�û����ǲ�����DBA��ɫ����ʹ�������û���֤��", vbInformation, gstrSysName
                txtUser.SetFocus
                Exit Sub
            End If
            gstrSysUser = strUser
            gstrSysPwd = strPwd
            Set gcnSystem = mcnOracle
        Case UCT_NormalUser
            mstrUser = strUser
        Case UCT_SysOwner
            strSQL = "Select 1 ����  From Session_Roles Where Role = 'DBA'" & vbNewLine & _
                            "Union All" & vbNewLine & _
                            "Select 1 ���� From Zltools.Zlsystems Where Upper(������) = User"
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "�����߹���Ա�ж�")
            If rsTmp.EOF Then
                MsgBox "��ǰ�û����߱������߹���ԱȨ�ޣ�", vbInformation, gstrSysName
                txtUser.SetFocus
                Exit Sub
            Else
                mstrUser = strUser
                Set gcnOracle = mcnOracle
            End If
        '��Ҫ����Ƿ���ָ�����ݿ��ָ��ʵ��
        Case UCT_RACInsUser
            arrTmp = Split(mstrRacInfo, ",")
            strSQL = "select 1" & vbNewLine & _
                    "  from v$database a" & vbNewLine & _
                    " where a.DBID = " & arrTmp(1) & vbNewLine & _
                    "   and userenv('instance') = " & arrTmp(0)
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "ָ��ʵ���ж�")
            If rsTmp.EOF Then
                MsgBox "�÷�����������Ҫ��֤��ʵ����", vbInformation, gstrSysName
                cboServer.SetFocus
                Exit Sub
            End If
    End Select
    mstrRemarks = strRemarks
    mblnOk = True
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ�ϵͳ���Զ��˳���", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If
End Sub

Private Sub cmdReloadSvr_Click()
    Dim strFileInfo As String
    Dim varItem As Variant
    Dim strServer As String
    
    strServer = cboServer.Text
    cboServer.Clear
    Set mcolServer = LoadServer(strFileInfo)
    For Each varItem In mcolServer
        cboServer.addItem varItem(0)
    Next
    cboServer.ToolTipText = strFileInfo
    cboServer.Text = strServer
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then
        If Trim(txtUser.Text) = "" Then
            cmdOK.Default = False
            If txtUser.Enabled Then txtUser.SetFocus
        Else
            If txtPWD.Enabled Then
                txtPWD.SetFocus
            Else
                cmdOK.SetFocus
            End If
        End If
        If muctCurType = UCT_AuditLog And gstrLoginUserName = gstrUserName Then
            txtRemarks.SetFocus
        End If
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPWD.Text) <> "" Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.name = "txtPWD" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If ActualLen(txtRemarks.Text) >= 256 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        SelAll txtUser
        OpenIme False
    End If
End Sub

Private Sub txtPWD_GotFocus()
    SelAll txtPWD
End Sub

Private Sub cboServer_GotFocus()
    If Me.ActiveControl Is cboServer Then
        SelAll cboServer
        OpenIme False
    End If
End Sub

Private Sub Form_Load()
    Dim strFileInfo As String
    Dim varItem As Variant

    mblnFirst = False
    mintTimes = 1
    If muctCurType = UCT_RACInsUser Then
        cmdReloadSvr.Enabled = True
        cmdReloadSvr.Visible = True
    End If
    '��ͨ�û���¼��֤
    If muctCurType = UCT_NormalUser Then
        txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="USER", Default:="")
    Else
        txtUser.Text = mstrUser
        txtUser.Enabled = mstrUser = ""
    End If
    
    If mstrServer <> "" Then
        cboServer.Locked = True
        cboServer.Text = mstrServer
    Else
        Set mcolServer = LoadServer(strFileInfo)
        For Each varItem In mcolServer
            cboServer.addItem varItem(0)
        Next
        cboServer.ToolTipText = strFileInfo
    End If

    Call ApplyOEM_Picture(Me, "Icon")
    cboServer.Enabled = False
    Select Case muctCurType
        Case UCT_ZLTOOLS
            lblNote.Caption = "    ������ZLTOOLS�����롣"
        Case UCT_CurZLBAK
            lblNote.Caption = "    ���������ʷ������롣"
        Case UCT_DBAUser
            lblNote.Caption = "    ������������ݿ�DBA��ɫ���û���"
        Case UCT_NormalUser
            lblNote.Caption = "    ������ϵͳ����Ȩ�û�������֤��"
        Case UCT_SysOwner, UCT_AuditLog
            lblNote.Caption = "    ������Ӧ��ϵͳ���������û�������֤��"
        Case UCT_RACInsUser
            cboServer.Enabled = True
            lblNote.Caption = "    ������""" & Split(mstrRacInfo, ",")(2) & """�ķ��������������û���֤"
    End Select
    
    If muctCurType = UCT_AuditLog Then
        If gstrLoginUserName <> gstrUserName Then     '��ͨ�û���¼
            '��ʼ���ؼ�λ��
            Me.Width = 5160
            Me.Height = 3690
            lblNote.Top = 390
            lblNote.Left = 915
            lblUser.Left = 660
            txtUser.Left = 1260
            txtUser.Width = txtRemarks.Width
            lblPWD.Left = 840
            txtPWD.Left = 1260
            txtPWD.Width = txtRemarks.Width
            fraSplit.Top = 2565
            cmdOK.Top = 2820
            cmdOK.Left = 2565
            cmdCancel.Top = cmdOK.Top
            cmdCancel.Left = 3660
        Else    'ϵͳ�������û���¼
            Me.Height = 2865
            If mstrRemarks <> "" Then
                Me.Caption = mstrRemarks
            Else
                Me.Caption = "����˵��"
            End If
            lblNote.Caption = "���������˵����"
            imgFlag.Visible = False
            lblRemarks.Visible = False
            lblNote.Left = 150
            lblNote.Top = 100
            txtRemarks.Left = 150
            txtRemarks.Top = lblNote.Top + lblNote.Height + 100
            txtRemarks.Width = 4560
            txtRemarks.Height = 1440
            fraSplit.Top = txtRemarks.Top + txtRemarks.Height
            cmdOK.Top = fraSplit.Top + fraSplit.Height + 50
            cmdCancel.Left = 3590
            cmdCancel.Top = cmdOK.Top
        End If
        lblDataBase.Visible = False
        cboServer.Visible = False
    Else
        txtRemarks.Visible = False
        lblRemarks.Visible = False
    End If
End Sub

Private Sub AppendText(KeyAscii As Integer)
'���ܣ���TextBox�ؼ���Text׷�����ݣ������ݵ�ǰText��ֵ���б��м������õ�������Ŀ
'������KeyAscii    ��ǰ�İ���
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '���ȵ�ǰ�û�������ַ�
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '�����ַ�ֻ�������֡�Ӣ�ĺͺ���
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With cboServer
        '��¼�ϴεĲ����λ��
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '���ݼ�������ݵõ����ܵ��б���
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdOK.Enabled = BlnState
    cmdCancel.Enabled = BlnState
End Sub

Private Function OracleOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassword As String, Optional ByRef strErr As String) As Boolean
'���ܣ� ��ָ�������ݿ�
    Dim blnOwner As Boolean, blnTransPassword As Boolean
    Dim ctTmp As enuProvider
    strErr = ""
    If muctCurType <> UCT_RACInsUser Then
        blnTransPassword = muctCurType = UCT_NormalUser Or muctCurType = UCT_SysOwner Or muctCurType = UCT_AuditLog
    Else
        blnTransPassword = Not (strUserName = "SYS" Or strUserName = "SYSTEM" Or strUserName = "ZLTOOLS")
    End If
    '�����û����ӵĻ�ȡ������ODBC���ӣ���Ϊ��������һ��Ĳ�ѯ������ִ�й��̣�ֻ��������ݿ�Ĺ���������߽ṹ����
    If Not blnTransPassword Then
        ctTmp = MSODBC
    Else
        ctTmp = OraOLEDB
    End If
    Set mcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, ctTmp, strErr, muctCurType = UCT_SysOwner)
    If mcnOracle.State = adStateClosed Then
         OracleOpen = False
        Set mcnOracle = Nothing
        If muctCurType = UCT_NormalUser Or muctCurType = UCT_SysOwner Or muctCurType = UCT_AuditLog Then
            Exit Function
        End If
    End If

    On Error GoTo ErrHand
    mstrUser = strUserName
    If muctCurType = UCT_NormalUser Then
        OracleOpen = zlGetUserInfo(mstrSystems, blnOwner)
        If Not blnOwner And Not OracleOpen Then
            MsgBox "��ʹ��Ӧ��ϵͳ����Ȩ�û�������֤��", vbOKOnly, gstrSysName
        End If
        mcnOracle.Close
        Set mcnOracle = Nothing
    Else
        OracleOpen = Not mcnOracle Is Nothing
    End If
    Exit Function
ErrHand:
    MsgBox "ע��:" & vbCrLf & "    ��½ʧ��,��ϸ������ϢΪ:" & vbCrLf & _
           "������Ϣ:" & err.Number & "-" & err.Description, vbOKOnly, gstrSysName
    OracleOpen = False
    err = 0
End Function

Private Function zlGetUserInfo(ByVal strSystems As String, Optional ByRef blnOwner As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '���û���Ϣ���蹫����������������ʹ��
    zlGetUserInfo = False
    blnOwner = False
    With rsTmp
        If .State = adStateOpen Then .Close
        strSQL = "Select S.������" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='���ű�') D" & _
                " Where Upper(S.������)=D.Owner And S.��� In (" & strSystems & ") Order by S.���"
        .Open strSQL, mcnOracle, adOpenKeyset
        If Not .EOF Then
            '��Ϊ���ܸ��û����ж��ϵͳ����ݣ�����ѭ��ȡ���
            If mstrUser = Nvl(!������) Then
                  MsgBox "ע��:" & vbCrLf & "   ��������������ݵ�½,�����������ݽ��е�½!", vbOKOnly, gstrSysName
                  blnOwner = True
                  Exit Function
            End If

            For i = 1 To .RecordCount
                strSQL = "Select R.ȱʡ,D.���� as ���ű���,D.���� as ��������,P.���,P.����,P.����" & _
                        " From " & !������ & ".�ϻ���Ա�� U," & !������ & ".��Ա�� P," & !������ & ".���ű� D," & !������ & ".������Ա R" & _
                        " Where U.��ԱID = P.ID And R.����ID = D.ID And P.ID=R.��ԱID and U.�û���=USER And (P.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.����ʱ�� Is Null) and R.ȱʡ=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSQL, mcnOracle, adOpenKeyset
                If Not rsUser.EOF Then
                    zlGetUserInfo = True
                    Exit For
                End If
                .MoveNext
            Next
        End If
        .Close
    End With
End Function



