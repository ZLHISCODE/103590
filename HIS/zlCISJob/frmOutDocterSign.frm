VERSION 5.00
Begin VB.Form frmOutDocterSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��дǩ��"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmOutDocterSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2580
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Index           =   0
      Left            =   -270
      TabIndex        =   18
      Top             =   510
      Width           =   5985
   End
   Begin VB.CheckBox chkPreText 
      Caption         =   "��ǩ��������Ϊǰ׺����(&P)"
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   1950
      Width           =   2565
   End
   Begin VB.CheckBox chkHandSign 
      Caption         =   "��ʾ��ǩλ��(&H)"
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   4170
      TabIndex        =   7
      Top             =   1380
      Width           =   1365
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1305
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   1785
      Width           =   5985
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   1995
   End
   Begin VB.OptionButton optName 
      Caption         =   "ָ���û�(&U)"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.OptionButton optName 
      Caption         =   "��ǰ�û�(&C)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   13
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4365
      TabIndex        =   12
      Top             =   2430
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1995
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��ʱ��(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   3255
      Width           =   5475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��Ч��Ԥ��:"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û�����(&P)"
      Height          =   180
      Left            =   510
      TabIndex        =   5
      Top             =   1365
      Width           =   990
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1605
      TabIndex        =   14
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ������(&L)"
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmOutDocterSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frmParent As Object                 '������
Private mlngPatiID As Long
Private mlngPatiPageID As Long
Private mstrSource As String                 '����ǩ����Դ�ַ���
Private mstrPatiSign As String              'ǩ����ʾ������
Private mintSign As Integer

Private Sign As cEPRSign                    'ǩ������
Private objESign As Object                  '����ǩ���ӿڲ���
Private lngCertID As Long                   '֤��ID
Private lngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mblnOK As Boolean
Private mobjRegister As Object  '������֤�������ò�������ʱ��������ΪNothing,�������չ��ܻ���

Private UserSignLevel As EPRSignLevelEnum   '��ǰ�û���ǩ������

Private Enum EPRDocTypeEnum
    cpr���ﲡ�� = 1
    cprסԺ���� = 2
    cpr�����¼ = 3
    cpr������ = 4
    cpr������� = 5
    cpr֪���ļ� = 6
    cpr���Ʊ��� = 7             '���Ƶ��ݣ�����
    cpr�������� = 8             '���Ƶ��ݣ�����
End Enum

Public Enum PatiFromEnum
    cprPF_���� = 1              '1-���
    cprPF_סԺ = 2              '2-סԺ��
    cprPF_���� = 3              '3-������
    cprPF_��� = 4              '4-���
End Enum

'ǩ��״̬
Public Enum EPRSignLevelEnum
    cprSL_�հ� = 0              'δǩ��
    cprSL_���� = 1              '����ҽʦǩ��
    cprSL_���� = 2              '����ҽʦǩ��
    cprSL_���� = 3              '����ҽʦǩ��
    cprSL_���� = 4              '���ߣ�ǩ�����𲻰�����ֻ��ʾ��Ա��������ְ�ƣ��Ա���������ҽʦ
End Enum

'################################################################################################################
'## ���ܣ�  ��ʾ������
'##
'## ������  edtThis     :IN     �༭���ؼ�
'##         fParent     :IN     ������
'##         strSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
'################################################################################################################
Public Function ShowMe(ByRef fParent As Object, ByVal strSource As String, _
    lngPatiID As Long, lngPatiPageID As Long) As cEPRSign
    
    Dim bytFileKind As Byte, bytPatiSource As Byte
    Dim lngStart As Long, strPreText As String
    
    Set frmParent = fParent
    mstrSource = strSource
    mlngPatiID = lngPatiID
    mlngPatiPageID = lngPatiPageID
    
    
    bytFileKind = cpr���ﲡ��
    bytPatiSource = cprPF_����
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "����ʾ"
    Me.cboTime.AddItem Format(Now(), "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(Now(), "yyyy��MM��dd�� hh:mm")
    
    mintSign = zlDatabase.GetPara("SignShow", glngSys, 1070, 0)
    
    UserSignLevel = GetUserSignLevel(UserInfo.ID, , mlngPatiID, mlngPatiPageID)  '��ȡ�û�ǩ������
    '����ǩ����������ʼ����ǩ������
    Select Case bytFileKind
    Case cpr������
        cmbLevel.AddItem "1 - ��ʿ"
        cmbLevel.AddItem "3 - ��ʿ��"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
    Case cpr���Ʊ���
        cmbLevel.AddItem "1 - ҽ��"
        cmbLevel.AddItem "2 - ����"
        cmbLevel.AddItem "3 - ����"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - ����ҽʦ"
        cmbLevel.AddItem "2 - ����ҽʦ"
        cmbLevel.AddItem "3 - ������ҽʦ"
        cmbLevel.AddItem "4 - ����ҽʦ"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 2
        If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 3
    End Select
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    Dim lS As Long
    Select Case bytFileKind
    Case cpr���ﲡ��
        lS = 1
    Case cprסԺ����
        lS = 2
    Case cpr���Ʊ���
        lS = 3
    Case cpr������
        lS = 4
    Case Else
        Select Case bytPatiSource
        Case cprPF_����
            lS = 1
        Case cprPF_סԺ
            lS = 2
        Case Else
            lS = 2  '������סԺΪ׼
        End Select
    End Select
    
    lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), lS, 1)) '����,סԺ,ҽ��,���� (1111),Ϊ��Ĭ�ϲ�������ģʽ
    lblUserName.Caption = UserInfo.����
    
    chkEsign.Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    chkHandSign.Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    Call RefControls
    
    Me.Show vbModal, frmParent
    If mblnOK Then
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    
    Dim rs As ADODB.Recordset, strSQL As String
    Dim lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    err = 0: On Error GoTo ErrHand
    strSQL = "Select g.����" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, �ϻ���Ա�� p" & vbNewLine & _
            "Where r.Grantee = p.�û��� And g.��ɫ = r.Granted_Role And g.ϵͳ = [2] And g.��� = [3] And g.���� = [4] And" & vbNewLine & _
            "      p.��Աid = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As ���� From �ϻ���Ա�� p Where �û��� = [5] And p.��Աid = [1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUserID, glngSys, 1070, "ǩ��Ȩ", UserInfo.�û���)
    If rs.RecordCount <= 0 Then GetUserSignLevel = cprSL_�հ�: Exit Function
    
    strSQL = "select Ƹ�μ���ְ��,ǩ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("Ƹ�μ���ְ��"), 0)
        If mintSign = 1 Then mstrPatiSign = "" & rs!ǩ��
    End If
    Select Case lngR    '1 ����  2 ����  3 �м�  4 ����/ʦ��  5 Ա/ʿ  9 ��Ƹ
    Case 1: lngLevel1 = cprSL_����
    Case 2: lngLevel1 = cprSL_����
    Case 3: lngLevel1 = cprSL_����
    Case Else: lngLevel1 = cprSL_����
    End Select
    rs.Close
    
    If lngPatiID > 0 Then
        strSQL = "Select ����ҽʦ, ����ҽʦ, ����ҽʦ " & _
            " From ���˱䶯��¼ " & _
            " Where ����ID = [1] And ��ҳID = [2] And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1) " & _
            "       And ��ʼʱ�� Is Not Null And Nvl(���Ӵ�λ, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_����
        Else
            If rs.Fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = cprSL_����
            ElseIf rs.Fields("����ҽʦ") = IIf(strUserName = "", UserInfo.����, strUserName) Then
                lngLevel2 = cprSL_����
            Else
                lngLevel2 = cprSL_����
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function

ErrHand:
    GetUserSignLevel = cprSL_�հ�
End Function

'################################################################################################################
'## ���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String
    Dim SignLevel As EPRSignLevelEnum
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        If chkEsign.Value = vbChecked Then
            '����ǩ��
            err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
            lngCertID = 0
            strSign = objESign.signature(mstrSource, UCase(gcnOracle.Properties(23)), lngCertID, strʱ���) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
        strUserName = IIf(mstrPatiSign = "", UserInfo.����, mstrPatiSign)
        lngUserID = UserInfo.ID
    Else
        If chkEsign.Value = vbUnchecked Then
            '����ǩ��
            If Not CreateRegister() Then
                GoTo LL
            End If
            If Not mobjRegister.LoginValidate(mobjRegister.GetServerName, txtName.Text, txtPass.Text, "") Then
                Validation = False
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        ElseIf chkEsign.Value = vbChecked Then
            '����ǩ��
            err.Clear: On Error Resume Next
            If objESign Is Nothing Then
                Set objESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0: strSign = ""
            End If
            If Not objESign Is Nothing Then
                Call objESign.Initialize(gcnOracle, glngSys)
            End If
            lngCertID = 0
            strSign = objESign.signature(mstrSource, UCase(txtName), lngCertID, strʱ���) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
        
        strSQL = "Select ID,����,ǩ�� From ��Ա�� p Where ID=(Select ��ԱID From �ϻ���Ա�� Where �û���=[1])"
        On Error GoTo errH
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Sign-GetUserInfo", UCase(txtName.Text))
        If Not rsTemp.EOF Then
            If mintSign = 1 Then strUserName = "" & rsTemp!ǩ��
            If strUserName = "" Then strUserName = rsTemp.Fields("����")
            lngUserID = rsTemp.Fields("ID")     '�û�ID
        End If
    End If
    SignLevel = GetUserSignLevel(lngUserID, strUserName, mlngPatiID, mlngPatiPageID) '��ȡָ���û���ǩ������
    
    If SignLevel < cprSL_���� And SignLevel < Val(cmbLevel.Text) Then
        MsgBox "ָ���û�ǩ�����𲻹���������������֤��Ϣ��", vbInformation, gstrSysName
        GoTo LL
    End If
    
    Set Sign = New cEPRSign
    
    Sign.���� = strUserName
    Sign.ǩ������ = Val(cmbLevel.Text)
    If Sign.ǩ������ > cprSL_���� Then Sign.ǩ������ = cprSL_����
    If Me.chkPreText.Value = vbChecked Then
        Sign.ǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        Sign.ǰ������ = ""
    End If
    Sign.ǩ����Ϣ = strSign   '����ǩ����ǩ����Ϣ�洢����Ҫ��ֵ���ֶ��У�
    Sign.��ʾ��ǩ = (chkHandSign.Value = vbChecked)
    Sign.ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
    Sign.ǩ������ = 1
    Sign.֤��ID = IIf(Sign.ǩ����ʽ = 2, lngCertID, 0)
    Sign.ǩ��ʱ�� = zlDatabase.Currentdate()
    Select Case Me.cboTime.ListIndex
    Case 1: Sign.��ʾʱ�� = "yyyy-MM-dd hh:mm"
    Case 2: Sign.��ʾʱ�� = "yyyy��MM��dd�� hh:mm"
    Case Else: Sign.��ʾʱ�� = ""
    End Select
    Sign.ʱ��� = strʱ���
    
    Validation = True
    Exit Function

LL:
    err = 0: On Error Resume Next
    If txtName.Enabled And txtName.Visible Then
        txtName.SetFocus
    ElseIf txtPass.Enabled And txtPass.Visible Then
        txtPass.SetFocus
    Else
        optName(0).SetFocus
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Sub RefControls()
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case lngPassType
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
        Select Case lngPassType
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
End Sub

Private Sub cboTime_Click()
     Call Preview
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Validation Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Preview()
    Dim strText As String, bln��ǩ As Boolean, strǰ������ As String
    If Me.chkPreText.Value = vbChecked Then
        strǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        strǰ������ = ""
    End If
    bln��ǩ = (chkHandSign.Value = vbChecked)
    strText = strǰ������ & IIf(mstrPatiSign = "", UserInfo.����, mstrPatiSign) & IIf(bln��ǩ, "����ǩ��_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "��" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lngPassType = 2 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.Value
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", chkHandSign.Value
    Set objESign = Nothing
    Set mobjRegister = Nothing
End Sub

Private Sub optName_Click(Index As Integer)
    Call RefControls
    If Index = 1 Then
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus
    End If
End Sub

Private Sub optName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
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
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkEsign.Value = vbUnchecked Then
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call ZLCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
        End If
    End If
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtNames_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function CreateRegister() As Boolean
    '����ע�Ჿ��
    If Not mobjRegister Is Nothing Then CreateRegister = True: Exit Function
    On Error Resume Next
    Set mobjRegister = CreateObject("zlRegister.clsRegister")
    If mobjRegister Is Nothing Then
        err.Clear
        MsgBox "����zlRegister��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ", vbExclamation, gstrSysName
        Exit Function
    End If
    CreateRegister = True
End Function
