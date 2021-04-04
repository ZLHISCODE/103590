VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserAuthentication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û���֤"
   ClientHeight    =   3375
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "frmUserAuthentication.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4965
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtRemarks 
      Height          =   840
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "�ñ�ע��������128�����ֻ�256���ַ�"
      Top             =   1755
      Width           =   3315
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   45
      Left            =   2715
      TabIndex        =   8
      Top             =   1020
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   79
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Left            =   0
      TabIndex        =   6
      Top             =   2655
      Width           =   5145
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2340
      TabIndex        =   4
      Top             =   2910
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3445
      TabIndex        =   5
      Top             =   2910
      Width           =   1100
   End
   Begin VB.TextBox txtUser 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1230
      MaxLength       =   30
      TabIndex        =   1
      Top             =   900
      Width           =   3315
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1230
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3315
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      Caption         =   "��ע˵��"
      Height          =   180
      Left            =   390
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   210
      Picture         =   "frmUserAuthentication.frx":1CFA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   750
      TabIndex        =   2
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "������Ӧ��ϵͳ���������û�������֤��"
      Height          =   210
      Left            =   1230
      TabIndex        =   7
      Top             =   390
      Width           =   3555
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUserAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrRemarks As String

Public Function ShowLogin(ByVal strDBAUser As String, ByRef strRemarks As String) As Boolean
'���ܣ���֤�û���¼
'������
'      strDBAUser=ϵͳ������
'      strRemarks=�ص���ע��Ϣ
'˵������ͨ�û���¼ʱ��ϵͳ�������û��������ݿ�ʱ����֤
    txtUser.Text = strDBAUser
    Me.Show vbModal
    strRemarks = mstrRemarks
    ShowLogin = mblnOK
    mblnOk = False
    mstrRemarks = ""
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNote As String
    Dim strRemarks As String
    Dim strUser As String, strPwd As String
    Dim strSQL As String

    SetConState False
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUser = Trim(txtUser.Text)
    strPwd = Trim(txtPWD.Text)
    strRemarks = Trim(txtRemarks.Text)
    
    '��Ч�ַ���Ч��
    If strPwd <> "" And Len(strPwd) <> 1 Then
        If Mid(strPwd, Len(strPwd) - 1, 1) = "/" Or Mid(strPwd, Len(strPwd) - 1, 1) = "@" Or Mid(strPwd, 1, 1) = "/" Or Mid(strPwd, 1, 1) = "@" Then
            txtPWD.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    
    If Len(strPwd) = 0 Then
        strNote = "����������"
        txtPWD.SetFocus
        GoTo InputError
    End If
    
    If Len(strRemarks) = 0 Then
        strNote = "�����뱸ע"
        txtRemarks.SetFocus
        GoTo InputError
    End If
    
    strUser = UCase(strUser)
    
    If Not OracleOpen(strUser, strPwd, strNote) Then
        txtPWD.Text = ""
        If txtPWD.Enabled Then txtPWD.SetFocus
        SetConState
        If strNote <> "" Then GoTo InputError
        Exit Sub
    End If
    
    mstrRemarks = strRemarks
    mblnOK = True
    Unload Me
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
    SetConState
    Exit Sub
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

Private Sub Form_Load()
    glngWndProc = SetWindowLong(txtRemarks.hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong txtRemarks.hwnd, GWL_WNDPROC, glngWndProc
End Sub

Private Sub txtPwd_GotFocus()
    SelAll txtPWD
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdOK.Enabled = BlnState
    cmdCancel.Enabled = BlnState
End Sub

Private Function OracleOpen(ByVal strUserName As String, ByVal strPassword As String, Optional ByRef strErr As String) As Boolean
'���ܣ� ��ָ�������ݿ�
    Dim cnOracle As ADODB.Connection '��֤�û�������
    Dim blnTransPassword As Boolean
    Dim ctTmp As enuProvider
    strErr = ""
    blnTransPassword = Not (strUserName = "SYS" Or strUserName = "SYSTEM" Or strUserName = "ZLTOOLS")
    '�����û����ӵĻ�ȡ������ODBC���ӣ���Ϊ��������һ��Ĳ�ѯ������ִ�й��̣�ֻ��������ݿ�Ĺ����������߽ṹ����
    If Not blnTransPassword Then
        ctTmp = MSODBC
    Else
        ctTmp = OraOLEDB
    End If
    Set cnOracle = gobjRegister.GetConnection(gstrServer, strUserName, strPassword, blnTransPassword, ctTmp, strErr, False)
    If cnOracle.State = adStateClosed Then
        OracleOpen = False
    Else
        OracleOpen = True
        cnOracle.Close
    End If
    Set cnOracle = Nothing
    Exit Function
End Function

Private Sub txtPWD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPWD.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPWD.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPWD_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPWD.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If ActualLen(txtRemarks.Text) >= 256 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRemarks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtRemarks.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtRemarks.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtRemarks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtRemarks.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub