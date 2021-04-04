VERSION 5.00
Begin VB.Form frmCaseTendSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ǩ��"
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5295
   Icon            =   "frmCaseTendSign.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   75
      Width           =   3765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2850
      TabIndex        =   8
      Top             =   1875
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   7
      Top             =   1875
      Width           =   1095
   End
   Begin VB.OptionButton optName 
      Caption         =   "��ǰ�û�(&U)"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   645
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "ָ���û�(&S)"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   1320
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   4
      Top             =   945
      Width           =   3480
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -345
      TabIndex        =   3
      Top             =   1770
      Width           =   5670
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1530
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1290
      Width           =   1995
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   3750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1365
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -345
      TabIndex        =   0
      Top             =   495
      Width           =   5805
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������(&L)"
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   135
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1530
      TabIndex        =   11
      Top             =   645
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��(&P)"
      Height          =   180
      Left            =   435
      TabIndex        =   10
      Top             =   1350
      Width           =   990
   End
End
Attribute VB_Name = "frmCaseTendSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'######################################################################################################################

Private frmParent As Object                 '������
Private Sign As cEPRSign                    'ǩ������

Private lngCertID As Long                   '֤��ID
Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mblnOk As Boolean

Private strSource As String                 '����ǩ����Դ�ַ���
Private UserSignLevel As EPRSignLevelEnum   '��ǰ�û���ǩ������
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlngUnitID As Long                  '��ǰ�����Ĳ���ID
Private mstrSource As String                'ǩ��ԭ����
Private mstr״̬ As String
Private mstrPrivs As String

'######################################################################################################################

'����ǩ��ʹ�ó��ϣ�
'26  ����ǩ��ʹ�ó���(4λ�ַ�) �Բ�ͬ�����Ƿ�ʹ�õ���ǩ�����п���,����λ���ֱ�Ϊ:����,סԺ,ҽ��,���� 0-������,1-����

Public Function ShowMe(ByRef objParent As Object, ByVal strPrivs As String, ByVal sSource As String, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngUnitID As Long, Optional str״̬ As String) As cEPRSign
    '******************************************************************************************************************
    '���ܣ� ��ʾǩ������
    '������ edtThis     :IN     �༭���ؼ�
    '       fParent     :IN     ������
    '       strSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
    '******************************************************************************************************************
    
    Set Sign = New cEPRSign
    Set frmParent = objParent
    strSource = sSource
    mstrPrivs = strPrivs
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlngUnitID = lngUnitID
    mstr״̬ = str״̬
    UserSignLevel = GetUserSignLevel(glngUserId, , mlng����ID, mlng��ҳID)  '��ȡ�û�ǩ������
    
    '����ǩ����������ʼ����ǩ������
    cmbLevel.AddItem "1 - ��ʿ"
    cmbLevel.AddItem "3 - ��ʿ��"
    cmbLevel.ListIndex = 0
    If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
    
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    'mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 4, 1))     '����,סԺ,ҽ��,���� (1111),Ϊ��Ĭ�ϲ�������ģʽ
    lblUserName.Caption = gstrUserName
    
    Call RefControls
    
    If mstr״̬ <> "" Then
        Call cmdOk_Click
    Else
        Me.Show vbModal, frmParent
    End If
    
    If mblnOk Then
        str״̬ = mstr״̬
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Private Function Validation() As Boolean
    '******************************************************************************************************************
    '
    '���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
    '
    '******************************************************************************************************************
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String, strʱ�����Ϣ As String
    Dim SignLevel As EPRSignLevelEnum, strErr As String
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        '--------------------------------------------------------------------------------------------------------------
        If chkEsign.Value = vbUnchecked Then
            '����ǩ��
            strUserName = gstrUserName
            lngUserID = glngUserId
        ElseIf chkEsign.Value = vbChecked Then
            '����ǩ��
            Err.Clear
            If gobjTendESign Is Nothing Then
                On Error Resume Next
                Set gobjTendESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err.Clear: strSign = ""
                On Error GoTo 0
                If Not gobjTendESign Is Nothing Then
                    Call gobjTendESign.Initialize(gcnOracle, glngSys)
                End If
            End If
            If gobjTendESign Is Nothing Then
                MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
                GoTo LL
            End If
            lngCertID = 0
            strSign = gobjTendESign.signature(strSource, UCase(gcnOracle.Properties(23)), lngCertID, strʱ���, , strʱ�����Ϣ) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
            strUserName = gstrUserName
            lngUserID = glngUserId
        End If
        SignLevel = GetUserSignLevel(lngUserID, , mlng����ID, mlng��ҳID) '��ȡָ���û���ǩ������
    Else
        '--------------------------------------------------------------------------------------------------------------
        If chkEsign.Value = vbUnchecked Then
            '����ǩ��
            If gobjRegister Is Nothing Then Set gobjRegister = DynamicCreate("zlRegister.clsRegister", "������֤���")
            If Not gobjRegister.LoginValidate("", txtName, txtPass, strErr) Then
                Validation = False
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��" & strErr, vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        ElseIf chkEsign.Value = vbChecked Then
            '����ǩ��
            Err.Clear
            If gobjTendESign Is Nothing Then
                On Error Resume Next
                Set gobjTendESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err.Clear: strSign = ""
                On Error GoTo 0
                If Not gobjTendESign Is Nothing Then
                    Call gobjTendESign.Initialize(gcnOracle, glngSys)
                End If
            End If
            If gobjTendESign Is Nothing Then
                MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
                GoTo LL
            End If
            lngCertID = 0
            strSign = gobjTendESign.signature(strSource, UCase(txtName), lngCertID, strʱ���, , strʱ�����Ϣ) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
        
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select ID,���� From ��Ա�� p Where ID=(Select ��ԱID From �ϻ���Ա�� Where �û���='" & UCase(txtName) & "') And (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If Not rsTemp.EOF Then
            strUserName = rsTemp.Fields("����") '�û�����
            lngUserID = rsTemp.Fields("ID")     '�û�ID
        End If
        rsTemp.Close
        SignLevel = GetUserSignLevel(lngUserID, strUserName, mlng����ID, mlng��ҳID) '��ȡָ���û���ǩ������
    End If
    
    If SignLevel < Val(cmbLevel.Text) Then
        MsgBox "ָ���û�û��ǩ��Ȩ�޻����ְ��δ�ﵽǩ������������������֤��Ϣ��", vbInformation + vbOKOnly, gstrSysName
        GoTo LL
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Sign.���� = strUserName
    Sign.ǩ������ = Val(cmbLevel.Text)
    If Sign.ǩ������ > cprSL_���� Then Sign.ǩ������ = cprSL_����
    Sign.ǩ����Ϣ = strSign
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ������ = 1
    Sign.֤��ID = IIf(Sign.ǩ����ʽ = 2, lngCertID, 0)
    Sign.ʱ��� = strʱ���
    Sign.ʱ�����Ϣ = strʱ�����Ϣ
    
    Validation = True
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
LL:
    Err = 0: On Error Resume Next
    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    ElseIf txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    Else
        optName(0).SetFocus
    End If
End Function

'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Sub RefControls()
    Dim arrData
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '63955:������,2013-09-16,���õ��ǩ�������ҵ�ǰǩ���Ĳ��������õĵ���ǩ�����ò����в���ʹ�õ���ǩ��
    '˵�������û�����õ���ǩ����Ҫ���õĲ���,��˵�����õ���ǩ���Ĳ���Ϊ���в���
    
    If mstr״̬ <> "" And InStr(1, mstr״̬, "|") <> 0 Then
        arrData = Split(mstr״̬, "|")
        cmbLevel.ListIndex = Val(arrData(0))
        optName(0).Value = arrData(1)
        optName(1).Value = arrData(2)
        txtName.Text = arrData(3)
        txtPass.Text = arrData(4)
        mlngPassType = Val(arrData(5))
    Else
        gstrSQL = "Select Zl_Fun_Getsignpar([1],[2]) ����ǩ�� From dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ǩ�����ò���", 4, mlngUnitID)
        If rsTemp.RecordCount > 0 Then
            mlngPassType = Val(NVL(rsTemp!����ǩ��, 0))
        Else
            mlngPassType = 0
        End If
        If mlngPassType = 1 Then
            If CertificateStoped(gstrUserName) = True Then mlngPassType = 0
        End If
    End If

    
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.Value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1������
            chkEsign.Value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2�����߽Կ�
        End Select
    Else
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.Value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1������
            chkEsign.Value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2�����߽Կ�
            txtPass.Enabled = (chkEsign.Value = vbUnchecked)
        End Select
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CertificateStoped(ByVal strName As String) As Boolean
'���ܣ����ǩ���˵�֤���Ƿ�ͣ�ã�ͣ�õĻ�����ʹ�õ���ǩ��
    On Error Resume Next
    CertificateStoped = True
    Err.Clear
    If gobjTendESign Is Nothing Then
        Set gobjTendESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
    End If
    
    If gobjTendESign Is Nothing Then Exit Function
    CertificateStoped = gobjTendESign.CertificateStoped(strName)
    If Err <> 0 Then Err.Clear
End Function

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Validation Then
        mstr״̬ = cmbLevel.ListIndex & "|" & optName(0).Value & "|" & optName(1).Value & "|" & txtName.Text & "|" & txtPass.Text & "|" & chkEsign.Value
        
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub optPassType_Click(Index As Integer)
    If Index = 1 Then
        txtPass.Enabled = True
        If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus
    Else
        txtPass.Enabled = False
    End If
End Sub

Private Sub optPassType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus:  Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab):  Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


