VERSION 5.00
Begin VB.Form frmSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��дǩ��"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2580
      Width           =   2310
   End
   Begin VB.Frame fraLine 
      Height          =   30
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
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1605
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   3840
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
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4365
      TabIndex        =   12
      Top             =   1935
      Width           =   1095
   End
   Begin VB.ComboBox cmbLevel 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4110
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
      Left            =   240
      TabIndex        =   17
      Top             =   3255
      Width           =   5235
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
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSign As cTabSign                     'ǩ������
Private fParent As Object                    'ǩ������
Private mlngCertID As Long                   '֤��ID
Private mblnOK As Boolean
Private UserSignLevel As EPRSignLevel   '��ǰ�û���ǩ������

'################################################################################################################
'## ���ܣ�  ��ʾ������
'##
'##         fParent     :IN     ������
'##         strSource   :IN     ����ǩ����Դ�ַ��������ı�����ȡ��ȥ��ǩ����٣�
'################################################################################################################
Public Function ShowMe(ByVal strSignKey As String, ByVal frmParent As Object) As cTabSign
    On Error GoTo errHand
    mblnOK = False
    Set fParent = frmParent
    Set mSign = fParent.Document.Signs("K" & strSignKey)
    cboTime.AddItem "����ʾ"
    cboTime.AddItem Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm")
    cboTime.AddItem Format(zlDatabase.Currentdate, "yyyy��MM��dd�� hh:mm")
    
    '����ǩ����������ʼ����ǩ������
    UserSignLevel = fParent.Document.ǩ������
    Select Case fParent.Document.EPRPatiRecInfo.��������
     Case Tab������
        cmbLevel.AddItem "1 - ��ʿ"
        cmbLevel.AddItem "3 - ��ʿ��"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 1
    Case Tab���Ʊ���
        cmbLevel.AddItem "1 - ҽ��"
        cmbLevel.AddItem "2 - ����"
        cmbLevel.AddItem "3 - ����"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 1
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 2
    Case Else
        cmbLevel.AddItem "1 - ����ҽʦ"
        cmbLevel.AddItem "2 - ����ҽʦ"
        cmbLevel.AddItem "3 - ������ҽʦ"
        cmbLevel.AddItem "4 - ����ҽʦ"
        cmbLevel.ListIndex = 0
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 1
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 2
        If UserSignLevel >= TabSL_���� Then cmbLevel.ListIndex = 3
    End Select
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��
    Dim lS As Long
    Select Case fParent.Document.EPRPatiRecInfo.��������
      Case Tab���ﲡ��
        lS = 1
    Case TabסԺ����
        lS = 2
    Case Tab���Ʊ���
        lS = 3
    Case Tab������
        lS = 4
    Case Else
        Select Case fParent.Document.EPRPatiRecInfo.������Դ
        Case TabPF_����
            lS = 1
        Case TabPF_סԺ
            lS = 2
        Case Else
            lS = 2  '������סԺΪ׼
        End Select
    End Select
        
    lblUserName.Caption = UserInfo.����
    
    chkHandSign.Value = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    chkPreText.Value = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "chkPreText", vbUnchecked)
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    Call RefControls
    
    Me.Show vbModal, frmParent

    If mblnOK Then
        Set ShowMe = mSign
    Else
        Set ShowMe = Nothing
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String, objSignPic As Object, strʱ�����Ϣ As String
    Dim SignLevel As EPRSignLevel, strSource As String, strFile As String, objFileStream As TextStream, strErr As String
    
    'strSourceΪ��Ҫ����ǩ����ԭ���ı�,��Doc����
    If chkEsign.Value = vbChecked Then
        Call fParent.Document.BuildXmlFile(strFile, True, True)
        If gobjFSO.FileExists(strFile) = False Then
            MsgBox "�������Ŀ¼�Ƿ����ļ�д��Ȩ�ޣ�", vbInformation, gstrSysName
            Exit Function
        End If
        Set objFileStream = gobjFSO.OpenTextFile(strFile, ForReading)
        strSource = objFileStream.ReadAll
        Set objFileStream = Nothing
        gobjFSO.DeleteFile strFile, True
    End If
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    
    If optName(0).Value Then
        If chkEsign.Value = vbUnchecked Then
            '����ǩ��
        ElseIf chkEsign.Value = vbChecked Then
            '����ǩ��
            Err.Clear: On Error Resume Next
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            mlngCertID = 0
            strSign = gobjESign.signature(strSource, UCase(UserInfo.�û���), mlngCertID, strʱ���, objSignPic, strʱ�����Ϣ) '����ǩ����Ϣ,mlngCertID����ǩ��ʹ�õ�֤��ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
        strUserName = UserInfo.����
        lngUserID = UserInfo.ID
        SignLevel = CInt(UserSignLevel)
    Else
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
            Err.Clear: On Error Resume Next
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0: strSign = ""
            End If
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            mlngCertID = 0
            strSign = gobjESign.signature(strSource, UCase(txtName), mlngCertID, strʱ���, objSignPic, strʱ�����Ϣ) '����ǩ����Ϣ,mlngCertID����ǩ��ʹ�õ�֤���¼ID
            If strSign = "" Then
                MsgBox "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
        
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select b.Id, b.����, b.ǩ��" & vbNewLine & _
                    "From �ϻ���Ա�� A, ��Ա�� B" & vbNewLine & _
                    "Where a.�û��� =[1] And a.��Աid = b.Id And" & vbNewLine & _
                    "      Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo", UCase(txtName))
        If Not rsTemp.EOF Then
            strUserName = rsTemp!����  '�û�����
            lngUserID = rsTemp!ID      '�û�ID
        End If
        rsTemp.Close
        SignLevel = GetUserSignLevel(lngUserID, strUserName, fParent.Document.EPRPatiRecInfo.����ID, fParent.Document.EPRPatiRecInfo.��ҳID)  '��ȡָ���û���ǩ������
    End If
    
    If SignLevel < TabSL_���� And SignLevel < Val(cmbLevel.Text) Then
        MsgBox "ָ���û�ǩ�����𲻹���������������֤��Ϣ��", vbInformation, gstrSysName
        GoTo LL
    End If
    
    With mSign
        .���� = strUserName
        .ǩ������ = Val(cmbLevel.Text)
        If .ǩ������ > TabSL_���� Then .ǩ������ = TabSL_����
        If chkPreText.Value = vbChecked Then
            .ǰ������ = Trim(Mid(cmbLevel.Text, 4)) & "��"
        Else
            .ǰ������ = ""
        End If
        .ǩ����Ϣ = strSign   '����ǩ����ǩ����Ϣ�洢����Ҫ��ֵ���ֶ��У�
        .��ʾ��ǩ = (chkHandSign.Value = vbChecked)
        .ǩ����ʽ = IIf(chkEsign.Value = vbUnchecked, 1, 2)
        .ǩ������ = 1
        .֤��ID = IIf(.ǩ����ʽ = 2, mlngCertID, 0) '����ǩ��
        .ʱ��� = strʱ���                         '����ǩ��
        .ǩ��ʱ�� = zlDatabase.Currentdate()
        Select Case cboTime.ListIndex
        Case 1: .��ʾʱ�� = "yyyy-MM-dd hh:mm"
        Case 2: .��ʾʱ�� = "yyyy��MM��dd�� hh:mm"
        Case Else: .��ʾʱ�� = ""
        End Select
    End With
    
    Validation = True
    Exit Function

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
    If optName(0).Value Then
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case gbytEsign
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
        Select Case gbytEsign
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
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkEsign_Click()
    txtPass.Enabled = (chkEsign.Value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).Value, False, txtPass.Enabled)
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub chkEsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHandSign_Click()
     Call Preview
End Sub

Private Sub chkHandSign_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub chkPreText_Click()
    Call Preview
End Sub

Private Sub chkPreText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
End Sub

Private Sub cmbLevel_Click()
    Call Preview
End Sub

Private Sub cmbLevel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
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
    strText = strǰ������ & UserInfo.���� & IIf(bln��ǩ, "����ǩ��_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "��" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
    If Err.Number <> 0 Then
        MsgBox Me.Caption & vbCrLf & Err.Description & "   ���ͼ��֪ͨ����Ա��", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "cboTime", cboTime.ListIndex
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "chkPreText", chkPreText.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\" & App.ProductName & "\" & Me.Name, "cmbLevel", cmbLevel.ListIndex
    Set fParent = Nothing
    Err.Clear
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
            If txtPass.Enabled And txtPass.Visible Then zlControl.TxtSelAll txtPass: txtPass.SetFocus: Call Preview: Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab): Call Preview: Exit Sub
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
