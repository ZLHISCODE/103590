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
Private mcnNewOLEDB As ADODB.Connection
Private mstrServer As String
Private mstrUserName As String
Private mblnOK As Boolean
Private mblnDefaultPreUser As Boolean
Private mblnDBUser As Boolean
Private mstrDBUser  As String

Public Function ShowMe(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDefaultPreUser As Boolean, Optional ByVal blnDBUser As Boolean, Optional ByRef strDBUser As String, Optional cnNewOLEDB As ADODB.Connection) As String
'������strNote=��ʾ��Ϣ(���)
'      lngProgID=�������
'      strFunc=��Ȩ����,blnDBUserΪTrueʱΪָ���û�
'      cnNew=Ҫ���ص�����,blnDBUse=falseʱ��,���봫���Nothing�Ķ���,������Ҫ�ɵ��ó���ر����ӣ�����ǵ�ǰ��¼�û�,����Nothing
'            blnDBUse=trueʱ������ն��󣬷��ش򿪵����Ӷ���
'      blnDefaultPreUser-ȱʡ��ʾ�ϴε�¼��
'      blnDBUser=�����ݿ��û�ֱ����֤��¼�������ظ��û�������������������û�������ʱ����lngProgId��strFunc��blnDefaultPreUser
'      strDBUser=�����������������ݿ��û�
'      cnNewOLEDB=��Ҫ��ȡ��OLEDB���ӣ���CNNEW����ͬһ�û����������Ӳ�ͬ�����ò�������Nothingʱ���ŷ��� OLEDB���ӣ����򲻷���
'���أ��ɹ�������Ա����
'      strDBUser=��������ݿ��û�
    mstrNote = strNote
    mlngSys = lngSys
    mlngProgID = lngProgId
    mblnDefaultPreUser = blnDefaultPreUser
    mblnDBUser = blnDBUser
    mstrDBUser = strDBUser
    mstrFunc = ""
    mstrUserName = ""
    Set mcnNewOLEDB = cnNewOLEDB
    If mblnDBUser Then
        mstrUserName = strFunc
    Else
        mstrFunc = strFunc
    End If
    
    Me.Show 1, frmParent
    If mblnOK Then
        ShowMe = mstrUserName
        If blnDBUser Then
            If Not mcnNew Is Nothing Then
                Set cnNew = mcnNew
                Set cnNewOLEDB = mcnNewOLEDB
            End If
        Else
            If Not cnNew Is Nothing Then
                Set cnNew = mcnNew
                Set cnNewOLEDB = mcnNewOLEDB
            ElseIf Not mcnNew Is Nothing Then
                mcnNew.Close
                Set mcnNew = Nothing
                If Not mcnNewOLEDB Is Nothing Then mcnNewOLEDB.Close
                Set mcnNewOLEDB = Nothing
            End If
        End If
        strDBUser = mstrDBUser
    Else
        Set cnNew = Nothing
        Set cnNewOLEDB = Nothing
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

    If Not OpenOracle(strUser, strPass) Then Exit Sub
    mstrDBUser = UCase(strUser)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IdentifyUser", txtUser.Text)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mstrUserName = ""
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
   
    Set mcnNew = Nothing
    mstrServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "Server", "")
    
    If mblnDBUser Then
        txtUser.Text = mstrUserName
        txtUser.Enabled = False
    Else
        If mblnDefaultPreUser Then
            txtUser.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IdentifyUser", "")
        End If
    End If
    
    If mstrNote <> "" Then lblNote.Caption = mstrNote
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = False Then mstrUserName = ""
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

Private Function OpenOracle(ByVal strUser As String, ByVal strPass As String) As Boolean
'���ܣ���֤�û�,�������û���������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    Dim strUserName As String
    
    Call SetEnabled(False)
    strUser = UCase(strUser)
    
    On Error GoTo errh
    
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
    Set mcnNew = gobjRegister.GetConnection(mstrServer, strUser, strPass, Not mblnDBUser, , , False)
    If mcnNew.State = adStateClosed Then
        Call SetEnabled(True)
        txtPass.Text = "": txtPass.SetFocus
        Set mcnNew = Nothing: Exit Function
    End If
    If Not mcnNewOLEDB Is Nothing Then
        Set mcnNewOLEDB = gobjRegister.GetConnection(mstrServer, strUser, strPass, Not mblnDBUser, OraOLEDB, , False)
    End If
    If mblnDBUser Then
        mstrUserName = strUser
    Else
        '����ϻ��û�
        strSQL = "Select B.���� From �ϻ���Ա�� A,��Ա�� B Where (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) And A.��ԱID=B.ID And A.�û���=[1]"
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
            If gobjComLib.SystemOwner(mlngSys) <> strUser Then
                strSQL = _
                    " Select 1 From zlUserRoles A,zlRoleGrant B " & _
                    " Where A.��ɫ=B.��ɫ And B.ϵͳ=[1] And B.���=[2] And B.����=[3] And A.�û� = [4]"
                Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngSys, mlngProgID, mstrFunc, strUser)
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
    End If
    
    Call SetEnabled(True)
    OpenOracle = True
    Exit Function
errh:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

