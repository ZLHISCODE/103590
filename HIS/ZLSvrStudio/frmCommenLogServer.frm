VERSION 5.00
Begin VB.Form frmCommenLogServer 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��־����������"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   Icon            =   "frmCommenLogServer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdReset 
      Caption         =   "����(&R)"
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   4365
      Width           =   1100
   End
   Begin VB.Frame fraLogUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��¼��־���ʺ�"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   285
      TabIndex        =   4
      Top             =   2760
      Width           =   3975
      Begin VB.CheckBox chkTrans 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "����ת��"
         Height          =   255
         Left            =   315
         TabIndex        =   7
         ToolTipText     =   "�Ƿ������������ZLHIS�����㷨ת�������ӵ����ݿ�"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox txtPWD 
         Height          =   320
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   32
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtUser 
         Height          =   320
         Left            =   1320
         MaxLength       =   32
         TabIndex        =   5
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblPWD 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����"
         Height          =   180
         Left            =   690
         TabIndex        =   19
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�û���"
         Height          =   180
         Left            =   510
         TabIndex        =   18
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame fraServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��־��������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   285
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
      Begin VB.TextBox txtSID 
         Height          =   320
         Left            =   1320
         MaxLength       =   32
         TabIndex        =   3
         Top             =   1100
         Width           =   1575
      End
      Begin VB.TextBox txtPort 
         Height          =   320
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   2
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txtIP 
         Height          =   320
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   1
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label lblSID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "ʵ����"
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�˿ں�"
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "IP��ַ"
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   14
      Top             =   960
      Width           =   6195
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   4575
      TabIndex        =   11
      Top             =   0
      Width           =   4575
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.δ������־������ʱ�Ե�ǰ��������Ϊ��־��������"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   13
         Top             =   240
         Width           =   4860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "2.δ����д��־�û�ʱ��ZLUA�û���Ϊд��־�û���"
         Height          =   450
         Index           =   0
         Left            =   165
         TabIndex        =   12
         Top             =   525
         Width           =   4830
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&Q)"
      Height          =   350
      Left            =   3165
      TabIndex        =   10
      Top             =   4365
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   1905
      TabIndex        =   9
      Top             =   4365
      Width           =   1100
   End
End
Attribute VB_Name = "frmCommenLogServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK          As Boolean

Public Function ShowMe() As Boolean
    mblnOK = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strServer As String, strUser As String, strPwd As String, strCommand As String
    Dim strSQL As String, strTemp As String, strError As String
    Dim cnOracle As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim i As Long
    
    On Error GoTo errH
    
    If txtIP.Text <> "" Or txtPort.Text <> "" Or txtSID.Text <> "" Then
        If txtIP.Text <> "" And txtPort.Text <> "" And txtSID.Text <> "" Then
        Else
            If txtIP.Text = "" Then
                MsgBox "��������־������IP��ַ", vbInformation, gstrSysName
                txtIP.SetFocus
                Exit Sub
            End If
            If txtPort.Text = "" Then
                MsgBox "��������־�������˿�", vbInformation, gstrSysName
                txtPort.SetFocus
                Exit Sub
            End If
            If txtSID.Text = "" Then
                MsgBox "��������־������ʵ����", vbInformation, gstrSysName
                txtSID.SetFocus
                Exit Sub
            End If
        End If
        strTemp = CheckIP("��������Ч��IP��ַ��", txtIP.Text)
        If strTemp <> "" Then
            MsgBox strTemp, vbInformation, gstrSysName
            txtIP.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(txtPort.Text) Then
            MsgBox "��������Ч�Ķ˿ں�", vbInformation, gstrSysName
            txtPort.SetFocus
            Exit Sub
        Else
            If val(txtPort.Text) > 6535 Or val(txtPort.Text) < 0 Then
                MsgBox "�˿ں�Ӧ��0��6534֮��", vbInformation, gstrSysName
                txtPort.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtUser.Text <> "" Or txtPWD.Text <> "" Or chkTrans.value <> 0 Then
        If txtUser.Text <> "" And txtPWD.Text <> "" Then
            '...
        Else
            If txtUser.Text = "" Then
                MsgBox "�������¼��־���û�", vbInformation, gstrSysName
                txtUser.SetFocus
                Exit Sub
            End If
            If txtPWD.Text = "" Then
                MsgBox "�������¼��־���û�����", vbInformation, gstrSysName
                txtPWD.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtIP.Text <> "" And txtPort.Text <> "" And txtSID.Text <> "" Then
        strServer = Trim(txtIP.Text) & ":" & val(txtPort.Text) & "/" & Trim(txtSID.Text)
        strCommand = "SERVER=" & strServer
    Else
        strServer = gstrServer
    End If
    If txtUser.Text <> "" And txtPWD.Text <> "" Then
        strUser = txtUser.Text
        strPwd = txtPWD.Text
        blnTrans = chkTrans.value <> 0
        strCommand = strCommand & " USER=" & strUser & " PASS=" & strPwd & " TRANS=" & IIf(blnTrans, 1, 0)
    Else
        strUser = "ZLUA"
        strPwd = Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY))
        blnTrans = False
    End If
    strCommand = Trim(strCommand)
    Set cnOracle = gobjRegister.GetConnection(strServer, strUser, strPwd, blnTrans, OraOLEDB, strError, False)
    If cnOracle.State = adStateClosed Then
        MsgBox "δ��������־��������������־���������û��Ƿ�������ȷ���Լ���־�û�״̬�Ƿ�����������" & strError, vbInformation, gstrSysName
        Exit Sub
    End If
    strSQL = "Select Table_Name, Privilege" & vbNewLine & _
            "From (Select 'ZLLOGCATEGORY' Table_Name, 'SELECT' Privilege" & vbNewLine & _
            "       From Dual" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'ZLLOGSET' Table_Name, 'SELECT' Privilege" & vbNewLine & _
            "       From Dual" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'ZLLOGINFO' Table_Name, 'SELECT' Privilege" & vbNewLine & _
            "       From Dual" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'ZLLOGSET_EDIT' Table_Name, 'EXECUTE' Privilege" & vbNewLine & _
            "       From Dual" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'ZLLOGCATEGORY_EDIT' Table_Name, 'EXECUTE' Privilege" & vbNewLine & _
            "       From Dual" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 'ZLLOGINFO_INSERT' Table_Name, 'EXECUTE' Privilege" & vbNewLine & _
            "       From Dual)" & vbNewLine & _
            "Minus (Select Table_Name, Privilege" & vbNewLine & _
            "       From All_Tab_Privs" & vbNewLine & _
            "       Where Table_Name In ('ZLLOGCATEGORY', 'ZLLOGSET', 'ZLLOGINFO','ZLLOGCATEGORY_EDIT','ZLLOGSET_EDIT', 'ZLLOGINFO_INSERT') And" & vbNewLine & _
            "             Grantee In ('PUBLIC', USER))"
    Set rsTmp = gclsBase.OpenSQLRecord(cnOracle, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        strError = strUser & "�û�ȱʧ���¶���Ȩ�ޣ�"
        
        Do While Not rsTmp.EOF
            strError = strError & vbNewLine & rsTmp!Privilege & " On " & rsTmp!Table_Name
            rsTmp.MoveNext
        Loop
        MsgBox strError & vbNewLine & "��Ը��û�������Ȩ�����л�������Ȩ�޵��û�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not gclsBase.UpdateZLReginfo("��־������", Sm4EncryptEcb(strCommand), val("2-����")) Then
        Exit Sub
    End If
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    MsgBox "�����־�������Լ��û��Ϸ�״̬���ִ���" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdReset_Click()
    Dim ctlItem As Control
    For Each ctlItem In Me.Controls
        If TypeOf ctlItem Is TextBox Then
            ctlItem.Text = ""
        ElseIf TypeOf ctlItem Is CheckBox Then
            ctlItem.value = 0
        End If
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL          As String, rsTmp            As ADODB.Recordset
    Dim strCommand      As String
    Dim strServer   As String, strUser      As String, strPass  As String, blnTrans As Boolean
    Dim arrTmp      As Variant, i           As Long
    
    On Error GoTo errH
    strSQL = "Select Max(����) ���� From zlRegInfo Where ��Ŀ = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "��־������")
    If Not IsNull(rsTmp!����) Then
        If rsTmp!���� & "" Like "ZLSV*:*" Then
            strCommand = Sm4DecryptEcb(rsTmp!����)
        Else
            strCommand = rsTmp!����
        End If

        arrTmp = Split(strCommand, " ")
        For i = LBound(arrTmp) To UBound(arrTmp)
            If Trim(arrTmp(i)) <> "" Then
                If arrTmp(i) Like "USER=*" Then
                    strUser = Mid(arrTmp(i), Len("USER=*"))
                ElseIf arrTmp(i) Like "PASS=*" Then
                    strPass = Mid(arrTmp(i), Len("PASS=*"))
'                ElseIf arrTmp(i) Like "TRANS=*" Then
'                    blnTrans = val(Mid(arrTmp(i), Len("TRANS=*"))) = 1
                ElseIf arrTmp(i) Like "SERVER=*" Then
                    strServer = Mid(arrTmp(i), Len("SERVER=*"))
                Else
                    If LenB(strServer) = 0 Then
                        strServer = arrTmp(i)
                    End If
                End If
            End If
        Next
        If InStr(strServer, "/") > 0 Then
            arrTmp = Split(strServer, "/")
            txtSID.Text = arrTmp(1)
            If InStr(arrTmp(0), ":") > 0 Then
                arrTmp = Split(arrTmp(0), ":")
                txtIP.Text = arrTmp(0)
                txtPort.Text = arrTmp(1)
            Else
                txtIP.Text = arrTmp(0)
                txtPort.Text = "1521"
            End If
        End If
        txtUser.Text = strUser
        txtPWD.Text = strPass
'        chkTrans.value = IIf(blnTrans, 1, 0)
    End If
    
    'zlLog�Ѿ�����Ϊ����������������������zlRegister�������ӣ�����ת�����߼�û��ʲô���壨��ʱȡ����
    chkTrans.value = 0
    chkTrans.Visible = False
    
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub txtIp_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "#" Or Chr(KeyAscii) = ".") Then
        If Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "#" Then
        If Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtSID_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) Like "[:/]" Or UCase(Chr(KeyAscii)) Like "[A-Z]" Or Chr(KeyAscii) Like "#") Then
        If Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
End Sub
