VERSION 5.00
Begin VB.Form frmCommenLogServer 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��־�������༭"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmCommenLogServer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLogUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "д��־�û�"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   4200
      Width           =   5175
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   960
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   960
         TabIndex        =   21
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����"
         Height          =   180
         Left            =   480
         TabIndex        =   24
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�û�"
         Height          =   180
         Left            =   480
         TabIndex        =   23
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��־�����������û�"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      TabIndex        =   15
      Top             =   2880
      Width           =   5175
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�û�"
         Height          =   180
         Left            =   480
         TabIndex        =   19
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����"
         Height          =   180
         Left            =   480
         TabIndex        =   18
         Top             =   810
         Width           =   360
      End
   End
   Begin VB.Frame fraServer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��־������"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtSID 
         Height          =   350
         Left            =   960
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtPort 
         Height          =   350
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtIP 
         Height          =   350
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblSID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "SID"
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   1290
         Width           =   270
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Port"
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   810
         Width           =   360
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "IP"
         Height          =   180
         Left            =   480
         TabIndex        =   9
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   7
      Top             =   990
      Width           =   5835
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   5205
      TabIndex        =   3
      Top             =   0
      Width           =   5205
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������״̬�����ú�����ϴ������ø��ͻ���"
         Height          =   180
         Index           =   2
         Left            =   1365
         TabIndex        =   6
         Top             =   675
         Width           =   3600
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ţ�Ψһȷ��һ���������ı�ʶ"
         Height          =   180
         Index           =   1
         Left            =   1365
         TabIndex        =   5
         Top             =   135
         Width           =   2700
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "Ĭ�Ϸ�������ֻ����һ��Ĭ��ȱʡ������"
         Height          =   225
         Index           =   0
         Left            =   1365
         TabIndex        =   4
         Top             =   405
         Width           =   3780
      End
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   0
      Left            =   -360
      TabIndex        =   2
      Top             =   5760
      Width           =   5835
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&Q)"
      Height          =   350
      Left            =   3525
      TabIndex        =   1
      Top             =   5925
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2265
      TabIndex        =   0
      Top             =   5925
      Width           =   1100
   End
End
Attribute VB_Name = "frmCommenLogServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=================================================================
'ģ�����
'=================================================================
Private mblnOk              As Boolean      '�Ƿ����ݴ���ɹ�
Private mlngServerNo        As Long         '���������
Private mblnHaveDefault     As Boolean      '�Ƿ����Ĭ�Ϸ�����
Private mblnChange          As Boolean
Private mblnCollect         As Boolean      '�Ƿ��ռ�������
Private mstrFileType        As String       '�ռ�����
Private mblnLoad            As Boolean      '�Ƿ����ݼ�����
Private Enum ServerState
    SS_ͣ�� = 1
    SS_���� = 0
End Enum

Private Enum ServerType
    ST_���� = 0
    ST_FTP = 1
End Enum
'=================================================================
'�����ӿ�
'=================================================================
Public Function ShowMe(ByVal lngServerNO As Long, ByVal blnHaveDefault As Boolean) As Boolean
'���ܣ��������ݵ������޸�
'intServerNO=Ҫ�༭�ķ�������ţ�=0��ʾ��������
'blnHaveDefault=�Ѿ�����Ĭ������������
'���أ�True-�ɹ���false-ʧ��
    mlngServerNo = lngServerNO
    mblnHaveDefault = blnHaveDefault
    mblnCollect = False
    mstrFileType = ""
    mblnOk = False
    mblnChange = False
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOk
End Function

'=================================================================
'˽�з���
'=================================================================
Private Sub chkDefault_Click()

    If chkDefault.Tag <> "" Or mblnLoad Then Exit Sub
    chkDefault.Tag = "������"
    If Not mblnHaveDefault Then
        chkDefault.value = 1
        Call MsgBox("�״������������������Ҫ���÷�����������ΪĬ��ȱʡ������������ȡ����", vbInformation, gstrSysName)
        chkDefault.Tag = ""
        Exit Sub
    End If
    optServerState(SS_ͣ��).Enabled = chkDefault.value = 1
    optServerState(SS_����).Enabled = chkDefault.value = 1
    If chkDefault.value = 1 Then
        optServerState(SS_����).value = True
    End If
    mblnChange = True
    chkDefault.Tag = ""
End Sub

Private Sub cmdCancel_Click()
    If mblnChange Then
        If MsgBox("�Ƿ������ǰ�༭���ݣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim objConn As clsConnect, strErr As String
    Dim strSQL  As String
    On Error GoTo ErrH
    '������
    If txtServerPath.Text = "" Then
        MsgBox "������" & IIf(optServerType(ST_����).value, "����Ŀ¼", "IP��ַ") & " !", vbInformation, gstrSysName
        txtServerPath.SetFocus
        Exit Sub
    End If
    If ActualLen(txtServerPath.Text) > txtServerPath.MaxLength Then
        MsgBox IIf(optServerType(ST_����).value, "����Ŀ¼", "IP��ַ") & "����" & txtServerPath.MaxLength & "λ�ַ����ȣ����������롣", vbInformation, gstrSysName
        txtServerPath.SetFocus
        Exit Sub
    End If
    
    If txtUser.Text = "" Then
        MsgBox "�������û��� !", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    If ActualLen(txtUser.Text) > txtUser.MaxLength Then
        MsgBox "�û�������" & txtUser.MaxLength & "λ�ַ����ȣ����������롣", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    
    If txtPWD.Text = "" Then
        MsgBox "���������� !", vbInformation, gstrSysName
        txtPWD.SetFocus
        Exit Sub
    End If
    If ActualLen(txtPWD.Text) > txtPWD.MaxLength Then
        MsgBox "���볬��" & txtPWD.MaxLength & "λ�ַ����ȣ����������롣", vbInformation, gstrSysName
        txtPWD.SetFocus
        Exit Sub
    End If
    
    If txtPort.Text = "" And txtPort.Enabled Then
        MsgBox "������˿ں� !", vbInformation, gstrSysName
        txtPort.SetFocus
        Exit Sub
    End If
    If MsgBox("�Ƿ��������У�飿", vbYesNo + vbInformation + vbDefaultButton1, gstrSysName) = vbYes Then
        Set objConn = New clsConnect
        If objConn.ToConnect(IIf(optServerType(ST_����).value, SCT_Share, SCT_FTP), txtServerPath.Text, txtUser.Text, txtPWD.Text, val(txtPort.Text), "", False, strErr) Then
            Call objConn.CloseConnect
        Else
            MsgBox "���Ӳ���ʧ�ܣ���Ϣ��" & strErr, vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    strSQL = "Zl_Zlupgradeserver_Update(1," & mlngServerNo & "," & IIf(optServerType(ST_����).value, 0, 1) & ",'" & Trim(txtServerPath.Text) & "','" & Trim(txtUser.Text) & "'," & SQLAdjust(Cipher(Trim(txtPWD.Text))) & "," & ZVal(txtPort.Text) & "," & IIf(optServerState(SS_����).value, 1, 0) & "," & IIf(chkDefault.value, 1, 0) & "," & IIf(optServerState(SS_����).value, 0, IIf(mblnCollect, 1, 0)) & "," & SQLAdjust(IIf(optServerState(SS_����).value, "", mstrFileType)) & "," & IIf(gblnDelFileServer, "NULL", SQLAdjust(Trim(txtPWD.Text))) & ")"
    Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    If mlngServerNo = 0 Then
        '������Ҫ������־
        Call SaveAuditLog(1, "�ļ�����������-����", "�������Ϊ" & txtno.Text & "���ļ�������")
    Else
        '������Ҫ������־
        Call SaveAuditLog(2, "�ļ�����������-�޸�", "�޸ı��Ϊ" & mlngServerNo & "���ļ�������")
    End If
    mblnOk = True
    Unload Me
    Exit Sub
ErrH:
    If 1 = 0 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdServerPath_Click()
    Dim strFolderName As String
    On Error Resume Next

    strFolderName = OpenFolder(Me, "ѡ�����²���������Ŀ¼")
    If Len(strFolderName) = 3 Then
        MsgBox "����ѡ���Ŀ¼(" & strFolderName & ")!", vbInformation, gstrSysName
        Exit Sub
    End If
    If InStr(1, strFolderName, "\\") <> 0 Then
        txtServerPath.Text = strFolderName
    Else
        txtServerPath.Text = "\\" & GetMyCompterName & Mid(strFolderName, 3)
    End If
End Sub

Private Sub Form_Activate()
    If mlngServerNo = 0 Then
        txtServerPath.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    mblnLoad = True
    On Error GoTo ErrH
    
    HookDefend txtPWD.hwnd
    
    If mlngServerNo <> 0 Then
        strSQL = "Select ���, ����, λ��, �û���, ����, �˿�, �Ƿ�����, �Ƿ�ȱʡ,�Ƿ��ռ�,�ռ����� From ZLTOOLS.Zlupgradeserver Where ���=[1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngServerNo)
        If rsTmp.EOF Then
            If MsgBox("��ǰ�����Ѿ���ɾ�����Ƿ��������ݣ�", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                mlngServerNo = 0
            Else
                On Error Resume Next
                Unload Me
                Exit Sub
            End If
        End If
    End If
    If mlngServerNo = 0 Then
        Me.Caption = "�����ļ�������"
        strSQL = "Select Nvl(Max(���), 0) + 1 ��� From Zlupgradeserver"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngServerNo)
        txtno.Text = rsTmp!���
        imgCaption.Picture = imgList.ListImages("NEW").Picture
    Else
        Me.Caption = "�޸��ļ�������"
        imgCaption.Picture = imgList.ListImages("EDIT").Picture
        optServerState(SS_ͣ��).value = val(rsTmp!�Ƿ����� & "") = 0 And val(rsTmp!�Ƿ�ȱʡ & "") = 0 And val(rsTmp!�Ƿ��ռ� & "") = 0
        optServerType(val(rsTmp!���� & "")) = True
        txtServerPath.Text = rsTmp!λ�� & ""
        txtUser.Text = rsTmp!�û��� & ""
        txtPWD.Text = Decipher(rsTmp!���� & "")
        txtPort.Text = rsTmp!�˿� & ""
        mblnCollect = val(rsTmp!�Ƿ��ռ� & "") = 1
        mstrFileType = rsTmp!�ռ����� & ""
        chkDefault.value = val(rsTmp!�Ƿ�ȱʡ & "")
    End If
    If Not mblnHaveDefault Then
        chkDefault.value = 1
        chkDefault.Enabled = False
        optServerState(SS_����).value = True
        optServerState(SS_����).Enabled = False
        optServerState(SS_ͣ��).Enabled = False
    End If
    mblnLoad = False
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub optServerState_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub optServerType_Click(Index As Integer)
    mblnChange = True
    lblServerPath.Caption = IIf(Index = ST_����, "����Ŀ¼", "IP��ַ")
    lblServerPath.Left = lblServerType.Left + lblServerType.Width - lblServerPath.Width
    cmdServerPath.Visible = Index = ST_����
    txtPort.Enabled = Index = ST_FTP
    If Not txtPort.Enabled Then
        txtPort.Text = ""
    Else
        txtPort.Text = "24"
    End If
End Sub

Private Sub txtPort_Change()
    mblnChange = True
End Sub

Private Sub txtPort_GotFocus()
    Call gclsBase.TxtSelAll(txtPort)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPWD_Change()
    mblnChange = True
End Sub

Private Sub txtPWD_GotFocus()
    Call gclsBase.TxtSelAll(txtPWD)
End Sub

Private Sub txtServerPath_Change()
    mblnChange = True
End Sub

Private Sub txtUser_Change()
    mblnChange = True
End Sub

Private Sub txtUser_GotFocus()
    Call gclsBase.TxtSelAll(txtUser)
End Sub

