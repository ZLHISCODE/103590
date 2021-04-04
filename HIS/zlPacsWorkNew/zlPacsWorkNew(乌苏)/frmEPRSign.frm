VERSION 5.00
Begin VB.Form frmEPRSign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��дǩ��"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5700
   Icon            =   "frmEPRSign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "���ǩ��"
      Height          =   350
      Index           =   2
      Left            =   4320
      TabIndex        =   19
      Top             =   2430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "���ǩ��"
      Height          =   350
      Index           =   1
      Left            =   4320
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
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
      TabIndex        =   17
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
      TabStop         =   0   'False
      Top             =   2257
      Width           =   1695
   End
   Begin VB.CheckBox chkEsign 
      Caption         =   "����ǩ��(&E)"
      Height          =   195
      Left            =   4170
      TabIndex        =   7
      TabStop         =   0   'False
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
      TabIndex        =   14
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
      Left            =   4320
      TabIndex        =   12
      Top             =   2840
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
      TabIndex        =   16
      Top             =   3255
      Width           =   5235
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ǩ��Ч��Ԥ��:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
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
      TabIndex        =   13
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
Attribute VB_Name = "frmEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sign As cEPRSign                    'ǩ������


Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�
Private mblnOK As Boolean
Private mlngPatientID As Long
Private mlngPageID As Long
Private mlngReportID As Long                '����ID
Private mint��ʼ�� As Integer               '���α���ǩ���Ŀ�ʼ��

Private mstrPrivs As String                 'Ȩ���ַ���

Private UserSignLevel As EPRSignLevelEnum   '��ǰ�û���ǩ������

'################################################################################################################
'## ���ܣ�  ��ʾ������
'##
'##         fParent     :IN     ������
'##         lngMaxSignLevel :IN  ���α����Ѿ��е�����ǩ������
'################################################################################################################
Public Function ShowMe(ByRef fParent As Object, ByVal lngPassType As Long, ByVal lngReportID As Long, ByVal lngPatientID As Long, ByVal lngPageID As Long, _
     ByVal strPrivs As String, ByVal lngMaxSignLevel As Long, int��ʼ�� As Integer) As cEPRSign
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    mlngPatientID = lngPatientID
    mlngPageID = lngPageID
    mstrPrivs = strPrivs
    mlngReportID = lngReportID
    mint��ʼ�� = int��ʼ��
    mlngPassType = lngPassType
    mblnOK = False
    
    Me.cboTime.Clear
    Me.cboTime.AddItem "����ʾ"
    Me.cboTime.AddItem Format(curDate, "yyyy-MM-dd hh:mm")
    Me.cboTime.AddItem Format(curDate, "yyyy��MM��dd�� hh:mm")
    
    '���ݵ�ǰ�����ǩ��״̬��ȷ���Ƿ���ʾ�����ǩ������ť�����Ѿ��й�һ�����ǩ���󣬲�����ʾ���ǩ����ť
    If lngMaxSignLevel > 1 Then
        cmdOK(1).Visible = False
    Else
        cmdOK(1).Visible = True
    End If
    
    '����Ȩ�ޣ�ȷ���Ƿ���ʾ�����ǩ������ť
    If InStr(mstrPrivs, "PACS�����޶�") > 0 Then
        cmdOK(2).Visible = True
    Else
        cmdOK(2).Visible = False
    End If

    Set Sign = New cEPRSign
    
    UserSignLevel = GetUserSignLevel(UserInfo.ID, , lngPatientID, lngPageID)    '��ȡ�û�ǩ������
    '����ǩ����������ʼ����ǩ������
    cmbLevel.AddItem "1 - ҽ��"
    cmbLevel.AddItem "2 - ����"
    cmbLevel.AddItem "3 - ����"
    cmbLevel.ListIndex = 0
    If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
    If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 2
    
    '��ȡ��ǰǩ����ʽ��ϵͳ����26��,���Ʊ����Ǵ� 3��ʼ
    'lngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 7, 1))  '����,סԺ,ҽ��,����,ҩƷ,LIS,PACS (1111111),Ϊ��Ĭ�ϲ�������ģʽ
    lblUserName.Caption = UserInfo.����
    
    chkEsign.value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", vbUnchecked)
    chkHandSign.value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", vbUnchecked)
    
    Dim intFormat As Integer
    intFormat = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", 0))
    If intFormat >= 0 And intFormat < Me.cboTime.ListCount Then Me.cboTime.ListIndex = intFormat
    
    'ˢ�½����и����ؼ�����ʾ״̬
    Call RefControls
    
    Me.Show vbModal, fParent
    If mblnOK Then
        Set ShowMe = Sign
    Else
        Set ShowMe = Nothing
    End If
End Function

'################################################################################################################
'## ���ܣ�  ����ǩ�����ڲ�ǩ���鲢ˢ����ʾ����֤�����������ǩ����
'################################################################################################################
Private Function Validation() As Boolean
    On Error GoTo LL
    Dim strUserName As String, lngUserID As Long, strSign As String, strʱ��� As String, strʱ���Base64 As String
    Dim objESign As Object                  '����ǩ���ӿڲ���
    Dim lngCertID As Long                   '֤��ID
    Dim SignLevel As EPRSignLevelEnum
    Dim strSource As String     '����Դ��
    Dim intRule As Integer      '����Դ����֯����
    Dim strDBUser As String     '�����ǩ�����������ݿ��û���
    Dim objPic As StdPicture
    
    txtName = Trim(txtName)
    txtPass = Trim(txtPass)
    strUserName = ""
    intRule = 1
    
    'ʹ�õ�ǰ�û�ǩ��
    If optName(0).value Then
        If chkEsign.value = vbChecked Then
            '����ǩ��
            strDBUser = UCase(gcnOracle.Properties(23))
        End If
        
        strUserName = UserInfo.����
        lngUserID = UserInfo.ID
        '��ȡ��ǰ�û���ǩ������
        SignLevel = GetUserSignLevel(lngUserID, , mlngPatientID, mlngPageID)
    Else
    'ʹ��ָ���û�ǩ��
        If chkEsign.value = vbUnchecked Then
            '����ǩ��
            If Not OraDataOpen(txtName, IIf(UCase(txtName) = "SYS" Or UCase(txtName) = "SYSTEM", txtPass, TranPasswd(txtPass))) Then
                MsgBoxD Me, "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If

        strDBUser = UCase(txtName)
        
        '�����ݿ��л�ȡָ���û�ǩ����ʽ��ǩ����������ID
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select ����,ID From ��Ա�� p Where ID=(Select ��ԱID From �ϻ���Ա�� Where �û���='" & strDBUser & "')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Sign-GetUserInfo")
        If Not rsTemp.EOF Then
            strUserName = rsTemp.Fields("����") '�û�����
            lngUserID = rsTemp.Fields("ID")     '�û�ID
        End If
        rsTemp.Close
        
        '��ȡָ���û���ǩ������
        SignLevel = GetUserSignLevel(lngUserID, strUserName, mlngPatientID, mlngPageID)
    End If
    
    If SignLevel < cprSL_���� And SignLevel < Val(cmbLevel.Text) Then
        MsgBoxD Me, "ָ���û�ǩ�����𲻹���������������֤��Ϣ��", vbInformation, gstrSysName
        GoTo LL
    End If
    
    
    Sign.���� = strUserName
    Sign.ǩ������ = Val(cmbLevel.Text)
    If Sign.ǩ������ > cprSL_���� Then Sign.ǩ������ = cprSL_����
    If Me.chkPreText.value = vbChecked Then
        Sign.ǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        Sign.ǰ������ = ""
    End If
    
    Sign.��ʾ��ǩ = (chkHandSign.value = vbChecked)
    Sign.ǩ����ʽ = IIf(chkEsign.value = vbUnchecked, 1, 2)
    Sign.ǩ������ = intRule
    Sign.ǩ��ʱ�� = zlDatabase.Currentdate()
    Sign.��ʼ�� = mint��ʼ��
    
    Select Case Me.cboTime.ListIndex
        Case 1: Sign.��ʾʱ�� = "yyyy-MM-dd hh:mm"
        Case 2: Sign.��ʾʱ�� = "yyyy��MM��dd�� hh:mm"
        Case Else: Sign.��ʾʱ�� = ""
    End Select
    
    '���������ǩ��������ȡԴ�ģ���Դ�ļ���
    If chkEsign.value = vbChecked Then
        '��������ǩ������
        err.Clear: On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If err <> 0 Then err = 0: strSign = ""
        End If
        
        '��ʼ������ǩ������
        If Not objESign Is Nothing Then
            If objESign.Initialize(gcnOracle, glngSys) = False Then
                MsgBoxD Me, "����֤���ʼ��ʧ�ܣ���ʹ����ȷ������֤��ǩ����", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If

        '�ȼ������֤�����½�û��Ƿ�һ��
        If objESign.CheckCertificate(strDBUser) = False Then
            '����֤��ͣ��ʱ����ʹ������ǩ����Դ�Ľ���ǩ�����ܣ������Լ���ǩ������
            If Not objESign.CertificateStoped(UserInfo.����) Then
                'Validation = True
                Exit Function
            End If
        Else
            '��ȡǩ����Դ��
            intRule = GetSignSourceString(1, mlngReportID, mint��ʼ��, False, Sign, strSource)
            If intRule = 0 Then
                'Դ����ȡʧ�ܣ��˳�ǩ��
                MsgBoxD Me, "���α���汾Ϊ" & mint��ʼ�� & "��ǩ��Դ����ȡʧ�ܣ��޷�ǩ����", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
            
            lngCertID = 0
            
            'ʹ������ǩ����Դ�Ľ���ǩ������
            '���أ�ǩ����ϢstrSign-���ܺ��Դ�ģ�lngCertID-ǩ��ʹ�õ�֤���¼ID��strʱ��� --ǩ��֮���ʱ���
            strSign = objESign.signature(strSource, strDBUser, lngCertID, strʱ���, objPic, strʱ���Base64)
            If strSign = "" Then
                MsgBoxD Me, "��֤ʧ�ܣ�������������֤��Ϣ��", vbInformation + vbOKOnly, "��дǩ��"
                GoTo LL
            End If
        End If
    End If
     
    '֤��ID����ͨ��ǩ�������صģ�����¼��ǩ����֤��ID�ֶΣ�����ֶα����ڡ��������ԡ��У����Զ������Ե�������ǩ��ǰ�����ģ�������Ϊǩ����Դ��
    Sign.֤��ID = IIf(Sign.ǩ����ʽ = 2, lngCertID, 0)
    
    'ǩ����Ϣ�������ͨ������ǩ������֮���������Ϣ
    Sign.ǩ����Ϣ = strSign   '����ǩ����ǩ����Ϣ�洢����Ҫ��ֵ���ֶ��У�
    
    'ǩ��ʱ�����
    Sign.ʱ��� = strʱ���
    
    'ʱ���base64����
    Sign.ʱ�����Ϣ = strʱ���Base64
    
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
End Function

'################################################################################################################
'## ���ܣ�  ��֤�û��������Ƿ���ȷ
'################################################################################################################
Private Function OraDataOpen(ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    Dim strSQL As String
    Dim strError As String
    Dim Cn As New ADODB.Connection
    
    On Error Resume Next
    err = 0
    With Cn
        If .State = adStateOpen Then .Close
'        .Provider = "MSDataShape"
        .Open gcnOracle.ConnectionString, strUserName, strUserPwd
        If err <> 0 Then
            OraDataOpen = False
            Exit Function
        End If
        .Close
    End With
    Set Cn = Nothing
    OraDataOpen = True
    Exit Function
errHand:
    Set Cn = Nothing
    OraDataOpen = False
    err = 0
End Function

'################################################################################################################
'## ���ܣ�  ����ת������
'##
'## ������  strOld  :ԭ����
'##
'## ���أ�  �������ɵ�����
'################################################################################################################
Public Function TranPasswd(strOld As String) As String
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

'################################################################################################################
'## ���ܣ�  ˢ�¿ؼ�
'################################################################################################################
Private Sub RefControls()
    If optName(0).value Then
        'ʹ�õ�ǰ�û�ǩ��
        txtName.Enabled = False
        txtPass.Enabled = False
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.value = vbUnchecked
            chkEsign.Visible = False
        Case 1
            '1������
            chkEsign.value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2�����߽Կ�
        End Select
    Else
        'ʹ��ָ���û�ǩ��
        chkEsign.Enabled = True
        txtPass.Enabled = True
        txtName.Enabled = True
        Select Case mlngPassType
        Case 0
            '����ǩ��
            chkEsign.value = vbUnchecked
            txtPass.Enabled = True
        Case 1
            '1������
            chkEsign.value = vbChecked
            chkEsign.Left = txtPass.Left
            Me.Label2.Visible = False
            chkEsign.Visible = True
            chkEsign.Enabled = False
            txtPass.Visible = False
        Case 2
            '2�����߽Կ�
            txtPass.Enabled = (chkEsign.value = vbUnchecked)
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
    txtPass.Enabled = (chkEsign.value = vbUnchecked)
    txtPass.Enabled = IIf(optName(0).value, False, txtPass.Enabled)
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

Private Sub Preview()
    Dim strText As String, bln��ǩ As Boolean, strǰ������ As String
    If Me.chkPreText.value = vbChecked Then
        strǰ������ = Trim(Mid(Me.cmbLevel.Text, 4)) & "��"
    Else
        strǰ������ = ""
    End If
    bln��ǩ = (chkHandSign.value = vbChecked)
    strText = strǰ������ & UserInfo.���� & IIf(bln��ǩ, "����ǩ��_____________", "")
    If Me.cboTime.ListIndex > 0 Then
        strText = strText & "��" & Me.cboTime.Text
    End If
    lblPreview.Caption = strText
    
End Sub

Private Sub cmdOK_Click(Index As Integer)
    '����ǩ�����ͣ��Զ�����ǩ������
    If optName(0).value = True Then
        If Index = 1 Then   '���ǩ�����Զ���ǩ�������趨Ϊ��ҽ����
            If cmbLevel.ListIndex <> 0 Then cmbLevel.ListIndex = 0
        ElseIf Index = 2 Then   '���ǩ���������ǰtxtLevel��ѡ����ǡ�ҽ�������Զ�����Ϊ��ߵ�ǩ�����𣬷����޸�
            If UserSignLevel < cprSL_���� Then
                MsgBoxD Me, "�����߱����ǩ����Ƹ��ְ�����顣"
                Exit Sub
            End If
            If cmbLevel.ListIndex = 0 Then
                If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 1
                If UserSignLevel >= cprSL_���� Then cmbLevel.ListIndex = 2
            End If
        End If
    End If
    
    '���������Ч�ԣ�ͬʱ��Դ�Ľ��м���
    If Validation Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
        Me.Tag = "1st."
        Me.cmbLevel.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlngPassType = 2 Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkEsign", chkEsign.value
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "chkHandSign", chkHandSign.value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "cboTime", chkHandSign.value
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
        If chkEsign.value = vbUnchecked Then
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
