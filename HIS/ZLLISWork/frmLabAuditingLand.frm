VERSION 5.00
Begin VB.Form frmLabAuditingLand 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����˵�½"
   ClientHeight    =   3240
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4800
   DrawMode        =   1  'Blackness
   Icon            =   "frmLabAuditingLand.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLabAuditingLand.frx":000C
   ScaleHeight     =   3240
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkClueTo 
      Caption         =   "���ʱ����Ҫ��ʾ:"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2460
      Width           =   1830
   End
   Begin VB.ComboBox cboHour 
      Height          =   300
      ItemData        =   "frmLabAuditingLand.frx":E14E
      Left            =   2160
      List            =   "frmLabAuditingLand.frx":E19A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2130
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   7
      Top             =   1995
      Width           =   4980
   End
   Begin VB.CheckBox chk����ʱ�� 
      Caption         =   "Ȩ����Ч��(Сʱ):"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2183
      Width           =   1830
   End
   Begin VB.CommandButton cmd���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3345
      TabIndex        =   6
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2235
      TabIndex        =   4
      Top             =   2730
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1695
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1500
      Width           =   2550
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1065
      Width           =   2550
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��  ��(&P)"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�����(&U)"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   1125
      Width           =   810
   End
End
Attribute VB_Name = "frmLabAuditingLand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintTimes As Integer                        '��¼���Դ���
Private mintAuditing As Integer                     '�Ƿ������Ȩ�� 0=û��Ȩ�� 1=��Ȩ�� -1��-24=��Чʱ��
Private mstrVerifyMan As String                     '������
Private mblnCancel As Boolean                       '�Ƿ���ȡ��
Private mstrLogID As String                         '��½��
Private Sub chk�޸�����_Click()
    Me.cboHour.ListIndex = 0
End Sub

Private Sub chk����ʱ��_Click()
    If chk����ʱ��.Value = 1 Then
        Me.cboHour.Enabled = True
    Else
        Me.cboHour.Enabled = False
    End If
End Sub

Private Sub cmdȷ��_Click()
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    Dim blGood As Boolean                                           '��"��˱걾"Ȩ��
    
    zlDatabase.SetPara "�Ƿ��о������Ȩ��", 0, 100, 1208
    mstrLogID = ""
    mintTimes = mintTimes + 1
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txt�û�.Text)
    strPassword = Trim(txt����.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txt�û�)) = 0 Then
        strNote = "�������û���"
        txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt�û�.SetFocus
            strNote = "�û�������"
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txt����.Enabled Then txt����.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "����������"
                    If txt����.Enabled Then txt����.SetFocus
        GoTo InputError
    End If
    
    strServerName = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
    
    If Not OraDataOpen(strServerName, UCase(strUserName), strPassword) Then
        txt����.Text = ""
        If txt����.Enabled Then txt����.SetFocus
        Exit Sub
    End If
    
    blGood = True

'    strSQL = "select user from dual "
'    Call zldatabase.OpenRecordset(rsTmp, strSQL, gstrSysName)
    
    If UCase(strUserName) = UserInfo.�û��� Then
        MsgBox "����½���û������ڵ��û�Ϊͬһ�û�,�����µ�½!", vbInformation, gstrSysName
        Me.txt���� = ""
        Me.txt�û� = ""
        Me.txt�û�.SetFocus
        Exit Sub
    End If
        
    
'    strSQL = "select ������ from zlsystems where ��� =100 and ������ = [1] "
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, strUserName)
'
'    If rsTmp.EOF = True Then
'        strSQL = "Select ���� From dba_role_privs A, zlRoleGrant B " & vbCrLf & _
'                " Where Granted_Role = B.��ɫ And grantee = [1] And Granted_Role Like 'ZL_%' " & vbCrLf & _
'                " And ϵͳ = [2] And ��� = [3] and ���� = [4] "
'        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, gstrSysName, strUserName, glngSys, glngModul, "��˱걾")
'
'        If rsTmp.EOF = False Then
'            Do Until rsTmp.EOF
'                If rsTmp("����") = "��˱걾" Then
'                    blGood = False
'                    Exit Do
'                End If
'                rsTmp.MoveNext
'            Loop
'        End If
'    End If
    
    If blGood = False Then
        MsgBox "������½���û�û��<���>Ȩ��!", vbInformation, gstrSysName
        Me.txt�û�.SetFocus
        Exit Sub
    End If
    
    If Me.chk����ʱ��.Value = 0 Then
        mintAuditing = 1
    Else
        mintAuditing = -CInt(Me.cboHour.Text)
    End If
    
    strsql = "select b.���� from �ϻ���Ա�� a ,��Ա�� b where �û��� = [1] and a.��Աid = b.id " & vbCrLf & _
             " And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UCase(strUserName))
    
    If rsTmp.EOF = False Then
        mstrLogID = UCase$(strUserName)
        zlDatabase.SetPara "�����", rsTmp("����"), 100, 1208
    Else
        MsgBox "�û���ͣ��!�����ڽ������.", vbInformation, gstrSysName
        mintAuditing = 0
    End If
    
    zlDatabase.SetPara "�Ƿ��о������Ȩ��", mintAuditing, 100, 1208
    zlDatabase.SetPara "���ʱ����Ҫ��ʾ", chkClueTo.Value, 100, 1208
    
    Unload Me
    Exit Sub

InputError:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ��Զ��˳�", vbExclamation, gstrSysName
        Call cmd����_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        Exit Sub
    End If

End Sub

Private Sub cmd����_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub Form_Activate()
'    Me.txt�û�.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If Me.ActiveControl.Name = "TXT����" Then
'            Call cmdȷ��_Click
'        Else
'            SendKeys "{Tab}"
'        End If
'    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mintTimes = 1
    Me.chk����ʱ��.Value = zlDatabase.GetPara("frmLabAuditingLand_ʱ��", 100, 1208, 0)
    Me.cboHour.ListIndex = zlDatabase.GetPara("frmLabAuditingLand_ʱ��", 100, 1208, 0)
    If Me.chk����ʱ��.Value = 0 Then
        Me.cboHour.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "frmLabAuditingLand_ʱ��", Me.chk����ʱ��.Value, 100, 1208
    zlDatabase.SetPara "frmLabAuditingLand_ʱ��", Me.cboHour.ListIndex, 100, 1208
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = Len(Me.txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�û�_GotFocus()
    Me.txt�û�.SelStart = 0: Me.txt�û�.SelLength = Len(Me.txt�û�.Text)
End Sub

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranPasswd = strNew

End Function

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strsql As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim objzlRegister As Object
    
    
    '����������,�����zlRegister������µ�ע�᷽ʽ,��������ϵ�ע�᷽ʽ
    On Error GoTo errhandOld
    If objzlRegister Is Nothing Then Set objzlRegister = CreateObject("zlRegister.clsRegister")
    If Not objzlRegister.LoginValidate(strServerName, strUserName, strUserPwd, strError) Then
        If strError <> "" Then
            MsgBox strError, vbInformation, "��½"
        End If
        Exit Function
    End If
    OraDataOpen = True
    Exit Function
errhandOld:


    On Error Resume Next
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strUserPwd, TranPasswd(strUserPwd))
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷�ע�ᡣ", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo ErrHand
    cnOracle.Close
    Set cnOracle = Nothing
        
    OraDataOpen = True
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub txt�û�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Public Sub ShowMe(objfrm As Object, strReportMan As String, blnCancel As Boolean, strLogID As String)
    '����               ��ʾ����˵�½����
    '����               Objfrm ���������
    '                   strReportMan ������
    mstrVerifyMan = strReportMan
    mblnCancel = False
    Me.Show vbModal, objfrm
    blnCancel = mblnCancel
    strLogID = mstrLogID
End Sub


