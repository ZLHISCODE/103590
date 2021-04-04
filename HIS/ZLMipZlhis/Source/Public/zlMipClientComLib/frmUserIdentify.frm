VERSION 5.00
Begin VB.Form frmUserIdentify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û���֤"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   6
      Top             =   1335
      Width           =   5025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   3
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1755
      TabIndex        =   2
      Top             =   1590
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      TabIndex        =   0
      Top             =   555
      Width           =   1920
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����֤���������û���������"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1335
      TabIndex        =   7
      Top             =   105
      Width           =   2520
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserIdentify.frx":000C
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1500
      TabIndex        =   5
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   615
      Width           =   540
   End
End
Attribute VB_Name = "frmUserIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNote As String
Private mlngSys As Long
Private mlngProgID As Long
Private mstrFunc As String

Private mcnNew As ADODB.Connection
Private mstrServer As String
Private mstrUserName As String
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection) As String
'������strNote=��ʾ��Ϣ(���)
'      lngProgID=�������
'      strFunc=��Ȩ����
'      cnNew=Ҫ���ص�����,��Ҫ����ʱ,���봫���Nothing�Ķ���,������Ҫ�ɵ��ó���ر����ӣ�����ǵ�ǰ��¼�û�,����Nothing
'���أ��ɹ�������Ա����
    mstrNote = strNote
    mlngSys = lngSys
    mlngProgID = lngProgId
    mstrFunc = strFunc
    
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrUserName
        If Not cnNew Is Nothing Then
            Set cnNew = mcnNew
        ElseIf Not mcnNew Is Nothing Then
            mcnNew.Close
            Set mcnNew = Nothing
        End If
    Else
        Set cnNew = Nothing
    End If
End Function

Private Sub cmdOK_Click()
    Dim strUser As String
    Dim strPass As String
    
    strUser = Trim(txtUser.Text)
    strPass = Trim(txtPass.Text)
    
    '��Ч�ַ���Ч��
    If strUser = "" Then
        MsgBox "�������û�����", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If InStr(strUser, "/") > 0 Or InStr(strUser, "@") > 0 Then
        MsgBox "��������Ч���û��������������롣", vbInformation, gstrSysName
        txtUser.SetFocus: Exit Sub
    End If
    If strPass = "" Then
        MsgBox "���������롣", vbInformation, gstrSysName
        txtPass.SetFocus: Exit Sub
    End If
    If InStr(strPass, "/") > 0 Or InStr(strPass, "@") > 0 Then
        MsgBox "��������Ч�����룬���������롣", vbInformation, gstrSysName
        txtPass.Text = "": txtPass.SetFocus: Exit Sub
    End If
    
    If Not OpenOracle(strUser, TranPasswd(strPass)) Then Exit Sub
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IdentifyUser", txtUser.Text)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Trim(txtUser.Text) <> "" Then txtPass.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If Me.ActiveControl Is txtPass Then
            Call cmdOK_Click
        Else
            Call gobjComLib.zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    mstrUserName = ""
    Set mcnNew = Nothing
    mstrServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "Server", "")
    'txtUser.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IdentifyUser", "")
    
    If mstrNote <> "" Then lblNote.Caption = mstrNote
End Sub

Private Sub txtUser_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtUser)
End Sub

Private Sub txtPass_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    cmdCancel.Enabled = blnEnabled
    cmdOK.Enabled = blnEnabled
    Screen.MousePointer = IIf(Not blnEnabled, 11, 0)
End Sub

Private Function IsOwner(ByVal strUser As String) As Boolean
'���ܣ��ж�ָ���û��Ƿ�ǰʹ��ϵͳ��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From zlSystems Where ������=[1] And ���=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser, mlngSys)
    IsOwner = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function OpenOracle(ByVal strUser As String, ByVal strPass As String) As Boolean
'���ܣ���֤�û�,�������û���������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    Dim strUserName As String
    
    Call SetEnabled(False)
    strUser = UCase(strUser)
    
    On Error GoTo errH
    
    '����û���
    strSQL = "Select UserName From All_Users Where UserName=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rsTmp.EOF Then
        MsgBox "���û������ڡ�", vbInformation, gstrSysName
        Call SetEnabled(True)
        txtPass.Text = "": txtUser.SetFocus
        Exit Function
    End If
    
    '�������
    On Error Resume Next
    Set mcnNew = New ADODB.Connection
    mcnNew.Provider = "MSDataShape"
    mcnNew.Open "Driver={Microsoft ODBC for Oracle};Server=" & mstrServer, strUser, strPass
    strError = Err.Description
    Err.Clear: On Error GoTo errH
    If strError <> "" Then
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
            MsgBox "�û�" & strUser & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            MsgBox "�û�������������޷�ͨ����֤��", vbInformation, gstrSysName
        Else
            MsgBox strError, vbInformation, gstrSysName
        End If
        Call SetEnabled(True)
        txtPass.Text = "": txtPass.SetFocus
        Set mcnNew = Nothing: Exit Function
    End If
    
    '����ϻ��û�
    strSQL = "Select B.���� From �ϻ���Ա�� A,��Ա�� B Where (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) And A.��ԱID=B.ID And Upper(A.�û���)=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strUser)
    If rsTmp.EOF Then
        MsgBox "���û�δ���ö�Ӧ����Ա��Ϣ��", vbInformation, gstrSysName
        Call SetEnabled(True)
        txtPass.Text = "": txtUser.SetFocus
        Exit Function
    End If
    strUserName = rsTmp!����
    
    '���Ȩ��
    If mstrFunc <> "" Then
        If Not IsOwner(strUser) Then
            strSQL = _
                " Select 1 From (" & _
                "   Select Granted_Role From DBA_Role_Privs Where Granted_Role Like 'ZL_%' And Grantee='" & strUser & "'" & _
                " ) A,zlRoleGrant B " & _
                " Where A.Granted_Role=B.��ɫ And B.ϵͳ=[1] And B.���=[2] And B.����=[3]"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngSys, mlngProgID, mstrFunc)
            If rsTmp.EOF Then
                MsgBox "���û�û��Ȩ�޽��в�����", vbInformation, gstrSysName
                Call SetEnabled(True)
                txtPass.Text = "": txtUser.SetFocus
                Exit Function
            End If
        End If
    End If
    
    '����ǵ�ǰ�û�����Ҫʹ�õ���������
    If strUser = UCase(gstrDBUser) Then
        mcnNew.Close: Set mcnNew = Nothing
    End If
    mstrUserName = strUserName
    Call SetEnabled(True)
    OpenOracle = True
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function TranPasswd(strOld As String) As String
'���ܣ�����ת������
'������strOld��ԭ����
'���أ��������ɵ�����
    Dim iBit As Integer, StrBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        StrBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(StrBit = "0", "W", StrBit = "1", "I", StrBit = "2", "N", StrBit = "3", "T", StrBit = "4", "E", StrBit = "5", "R", StrBit = "6", "P", StrBit = "7", "L", StrBit = "8", "U", StrBit = "9", "M", _
                   StrBit = "A", "H", StrBit = "B", "T", StrBit = "C", "I", StrBit = "D", "O", StrBit = "E", "K", StrBit = "F", "V", StrBit = "G", "A", StrBit = "H", "N", StrBit = "I", "F", StrBit = "J", "J", _
                   StrBit = "K", "B", StrBit = "L", "U", StrBit = "M", "Y", StrBit = "N", "G", StrBit = "O", "P", StrBit = "P", "W", StrBit = "Q", "R", StrBit = "R", "M", StrBit = "S", "E", StrBit = "T", "S", _
                   StrBit = "U", "T", StrBit = "V", "Q", StrBit = "W", "L", StrBit = "X", "Z", StrBit = "Y", "C", StrBit = "Z", "X", True, StrBit)
        Case 2
            strNew = strNew & _
                Switch(StrBit = "0", "7", StrBit = "1", "M", StrBit = "2", "3", StrBit = "3", "A", StrBit = "4", "N", StrBit = "5", "F", StrBit = "6", "O", StrBit = "7", "4", StrBit = "8", "K", StrBit = "9", "Y", _
                   StrBit = "A", "6", StrBit = "B", "J", StrBit = "C", "H", StrBit = "D", "9", StrBit = "E", "G", StrBit = "F", "E", StrBit = "G", "Q", StrBit = "H", "1", StrBit = "I", "T", StrBit = "J", "C", _
                   StrBit = "K", "U", StrBit = "L", "P", StrBit = "M", "B", StrBit = "N", "Z", StrBit = "O", "0", StrBit = "P", "V", StrBit = "Q", "I", StrBit = "R", "W", StrBit = "S", "X", StrBit = "T", "L", _
                   StrBit = "U", "5", StrBit = "V", "R", StrBit = "W", "D", StrBit = "X", "2", StrBit = "Y", "S", StrBit = "Z", "8", True, StrBit)
        Case 0
            strNew = strNew & _
                Switch(StrBit = "0", "6", StrBit = "1", "J", StrBit = "2", "H", StrBit = "3", "9", StrBit = "4", "G", StrBit = "5", "E", StrBit = "6", "Q", StrBit = "7", "1", StrBit = "8", "X", StrBit = "9", "L", _
                   StrBit = "A", "S", StrBit = "B", "8", StrBit = "C", "5", StrBit = "D", "R", StrBit = "E", "7", StrBit = "F", "M", StrBit = "G", "3", StrBit = "H", "A", StrBit = "I", "N", StrBit = "J", "F", _
                   StrBit = "K", "O", StrBit = "L", "4", StrBit = "M", "K", StrBit = "N", "Y", StrBit = "O", "D", StrBit = "P", "2", StrBit = "Q", "T", StrBit = "R", "C", StrBit = "S", "U", StrBit = "T", "P", _
                   StrBit = "U", "B", StrBit = "V", "Z", StrBit = "W", "0", StrBit = "X", "V", StrBit = "Y", "I", StrBit = "Z", "W", True, StrBit)
        End Select
    Next
    TranPasswd = strNew
End Function
