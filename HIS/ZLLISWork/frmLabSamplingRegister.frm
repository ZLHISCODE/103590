VERSION 5.00
Begin VB.Form frmLabSamplingRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Ǽ�"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8040
   Icon            =   "frmLabSamplingRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   6540
      TabIndex        =   26
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Height          =   345
      Left            =   5100
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkConcatenation 
      Caption         =   "���浱ǰ��Ŀ��������"
      Height          =   225
      Left            =   30
      TabIndex        =   14
      Top             =   2220
      Width           =   2295
   End
   Begin VB.Frame FraPatientInfo 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   7965
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   1350
         Width           =   4455
      End
      Begin VB.ComboBox cboִ�п��� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   210
         Width           =   1635
      End
      Begin VB.ComboBox cbo�Ա� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6852
         Left            =   3210
         List            =   "frmLabSamplingRegister.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   2
         Top             =   210
         Width           =   435
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6856
         Left            =   4770
         List            =   "frmLabSamplingRegister.frx":6869
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox txtBed 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
      Begin VB.TextBox txtPatientDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3570
         TabIndex        =   7
         Top             =   600
         Width           =   4185
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtҽ������ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   6525
      End
      Begin VB.ComboBox cbo�������� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabSamplingRegister.frx":6885
         Left            =   900
         List            =   "frmLabSamplingRegister.frx":6887
         TabIndex        =   8
         Top             =   960
         Width           =   1635
      End
      Begin VB.ComboBox cboҽ�� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   9
         Top             =   960
         Width           =   1785
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   16
         Top             =   1710
         Width           =   285
      End
      Begin VB.TextBox txt����1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5550
         MaxLength       =   5
         TabIndex        =   4
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        λ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   150
         TabIndex        =   28
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   5460
         TabIndex        =   27
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3915
         TabIndex        =   25
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2790
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڿ���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2790
         TabIndex        =   23
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6330
         TabIndex        =   22
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ʶ ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   645
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��       ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   18
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   2790
         TabIndex        =   17
         Top             =   990
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLabSamplingRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsRelativeAdvice As ADODB.Recordset                             '�Ǽǵ����ҽ��
Private PatientType As Integer, mlng����ID As Long, mstrNO As String    '�����շѵ��ݺ�
Private mlngCapID As Long                                               '�ɼ���ĿID
Private mlngReqDept As Long, mstrReqDoctor As String                    'Ĭ�ϵĵǼǿ��Һ�ҽ��
Private mlngKey As Long                                                 'ID
Private mblnSaveAdvice As Boolean                                       '�Ƿ���Ҫ����ҽ���������޸���Ժ���˱걾��Ϣ
Private mstrKeys As String                                              '��ǰ���յ�����ҽ��ID
Private mblnBarCode As Boolean                                          '����
Private iInputType As Integer
Private mstrExtData  As String                                           '�Ǽǵ�������Ŀ��Ϣ
Private mbln΢������Ŀ As Boolean
Private mlngDeptID As Long                                              '����ID

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cbo��������_GotFocus()
    Call zlControl.TxtSelAll(cbo��������)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo��������.ListIndex <> -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex): Exit Sub '��ѡ��
    If cbo��������.Text = "" Then '������
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cbo��������.Text))
    'ȫԺ�ٴ�����
    strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo��������.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        If Not zlControl.CboLocate(cbo��������, rsTmp!����) Then
            cbo��������.Text = ""
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Me.cbo��������.ListIndex > -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cboҽ��_Click()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboҽ��.ListIndex <> -1 Then mstrReqDoctor = Me.cboҽ��.Text: Exit Sub '��ѡ��
    If cboҽ��.Text = "" Then '������
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cboҽ��.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cboҽ��.Text = rsTmp!����
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    If Len(Trim(Me.cboҽ��.Text)) > 0 Then mstrReqDoctor = Me.cboҽ��.Text
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboִ�п���_Click()
    mlngDeptID = cboִ�п���.ItemData(cboִ�п���.ListIndex)
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not ValidAdvice Then Exit Sub
        
    mlngKey = SaveAdviceData
    If mlngKey = 0 Then
        MsgBox "����ʧ��", vbInformation, gstrSysName
        Exit Sub
    Else
        If Me.chkConcatenation.Value = 1 Then
'            Me.txt����.Text = "": Me.txt����.Tag = "":
            Me.txt����.SetFocus
            If Me.cbo��������.ListIndex > -1 Then
                mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
            End If
            If Me.cboҽ��.ListIndex > -1 Then
                mstrReqDoctor = Me.cboҽ��.ItemData(Me.cboҽ��.ListIndex)
            End If
        Else
'            Me.txt����.Text = "": Me.txt����.Tag = "":
            txtUnit.Text = "": Me.txtҽ������.Text = "": Me.txtҽ������.Tag = "": Me.txt����.SetFocus
        End If
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim strExtData As String
    Dim rsTmp As New ADODB.Recordset
    
    strExtData = frmLabSamplingSelect.ShowMe(Me, mlngDeptID)
    If strExtData <> "" Then
        '��ȡ�ɼ���ʽ
        Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
        If rsTmp Is Nothing Then
            MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
            Exit Sub
        End If
        mlngCapID = rsTmp("ID")
        Call AdviceSet�������(3, strExtData)
        txtҽ������.Text = Get�����������(2, "")
        txtҽ������.Text = txtҽ������.Text & "(" & Split(strExtData, ";")(1) & ")"
    End If
End Sub

Private Sub Form_Load()
    InitDepts                     'ȡ�ÿ��Һ��Ա�
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlDatabase.SetPara "�ɼ�����վ�Ǽ�", chkConcatenation.Value, 100, 1211
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        If Len(Trim(Me.cbo��������.Text)) <= 0 Then
            Me.cbo��������.SetFocus
        ElseIf Len(Trim(Me.cboҽ��.Text)) <= 0 Then
            Me.cboҽ��.SetFocus
'        ElseIf Len(Trim(Me.cboִ�п���.Text)) <= 0 Then
'            Me.cboִ�п���.SetFocus
        ElseIf Len(Trim(Me.txtҽ������.Text)) <= 0 Then
            Me.txtҽ������.SetFocus
        Else
            Me.cmdOK.SetFocus
        End If
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txt����_Validate(Cancel As Boolean)
    Dim strInput As String
    Dim rsTmp As New ADODB.Recordset, i As Integer
    Dim strField As String
    Dim strBarCode As String
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim intSelect As Integer
    Dim strAge As String
    Dim aAge() As String
    
    If Len(Trim(txt����)) = 0 Then Exit Sub
    If txt���� = txt����.Tag Then Exit Sub
    
    Call AdjustEditState(True)

    
    mblnSaveAdvice = True
    Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)
    
    Me.cbo��������.ListIndex = -1
    Me.cboҽ��.ListIndex = -1
'    Me.txtҽ������.Text = ""
    
    '��ʼ������Ϣ
    Set rsTmp = GetPatient(txt����)
    strBarCode = txt����
    If rsTmp.EOF Then
        mlng����ID = 0
        '�Ǽ��²���
        mstrKeys = ""
        Me.txt���� = "": Me.txt����1 = "": Me.cboAge.ListIndex = 0
        Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Me.txtID = "": Me.txtBed = ""
        '���������Ժ�ڲ��ˣ����������
        If InStr("+-*./", Left(Me.txt����.Text, 1)) > 0 Or mblnBarCode Then
            Me.txt����.Text = "": Cancel = True
            Exit Sub
        End If
        PatientType = 1
        '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
        If mlngReqDept > 0 Then
            cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
            Me.cboҽ��.Text = mstrReqDoctor
        End If
    Else
        On Error Resume Next
        Me.txt����.Text = Nvl(rsTmp("����"))
        Me.txt����.Text = "": Me.txt����1.Text = ""
        strAge = IIf(IsNull(rsTmp("����")), "", rsTmp("����")): If Me.txt���� = "0" Then Me.txt���� = ""
        
        strAge = Replace(strAge, "Сʱ", "ʱ")
        strAge = Replace(strAge, "����", "��")

        If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
            If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
                Me.txt����.Text = ""
                Me.cboAge.Text = Trim(strAge)
            Else
                strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
                aAge = Split(strAge, ";")
                If UBound(aAge) = 1 Then
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                Else
                    Me.txt����.Text = Val(aAge(0))
                    Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
                    Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
                End If
            End If
        Else
            Me.txt����.Text = ""
            Me.cboAge.ListIndex = 0
        End If
'        Me.txt���� = IIf(IsNull(rsTmp("����")), "", Val(rsTmp("����"))): If Me.txt���� = "0" Then Me.txt���� = ""
'        Me.cboAge.Text = IIf(IsNull(rsTmp("����")), "��", Replace(rsTmp("����"), Val(rsTmp("����")), ""))
        If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
        Me.cbo�Ա� = Nvl(rsTmp("�Ա�")) ' CombIndex(cbo�Ա�, Nvl(rsTmp("�Ա�")))
        
        mlng����ID = Nvl(rsTmp("����ID"), 0): PatientType = Nvl(rsTmp("PatientType"), 1)
            
        '����Ĭ�Ͽ������ҡ�ҽ��
        cbo��������.ListIndex = FindComboItem(cbo��������, Nvl(rsTmp("���˿���"), 0))
        
        '���˵�λ
        txtUnit.Text = Nvl(rsTmp("������λ"))
        DoEvents
        
        strField = ""
        strField = rsTmp.Fields("ҽ��").Name
        If strField = "ҽ��" Then
            Me.cboҽ��.Text = Nvl(rsTmp("ҽ��"))
            For i = 0 To Me.cboҽ��.ListCount - 1
                If Me.cboҽ��.List(i) Like Nvl(rsTmp("ҽ��")) Then
                    Me.cboҽ��.ListIndex = i
                    Exit For
                End If
            Next
        End If
        '��ʾ���˿���
        strSQL = "Select ���� From ���ű� Where ID=[1]"
        Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Nvl(rsTmp("���˿���"), 0)))
        If rsDept.EOF Then
            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
        Else
            Me.txtPatientDept.Text = rsDept("����"): Me.txtPatientDept.Tag = Nvl(rsTmp("���˿���"), 0)
        End If
        Me.txtID = Nvl(rsTmp("סԺ��")): If Len(Me.txtID) = 0 Then Me.txtID = Nvl(rsTmp("�����"))
        Me.txtBed = Nvl(rsTmp("��ǰ����"))
    
        '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
        If Me.cbo��������.ListIndex = -1 And mlngReqDept > 0 Then
            cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
            Me.cboҽ��.Text = mstrReqDoctor
        End If
    End If
    txt����.Tag = txt����.Text
    Me.cbo�Ա�.Tag = "����"
End Sub

Private Sub txtҽ������_GotFocus()
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txtҽ������.Text = txtҽ������.Tag Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        
        With txtҽ������
            Set rsTmp = SelectDiagItem()
        End With
        
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        If AdviceInput(rsTmp) Then
            DoEvents
            '��ʾ��ȱʡ���õ�ֵ
            txtҽ������.Tag = txtҽ������.Text
            Me.cmdOK.SetFocus
        Else
            DoEvents
            '�ָ�ԭֵ
            txtҽ������.Text = txtҽ������.Tag
            zlControl.TxtSelAll txtҽ������

            txtҽ������.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub
Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cboҽ��.Clear
    
    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
        
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ��.AddItem rsTmp!����
            cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
            
            If rsTmp!ID = UserInfo.ID And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = cboҽ��.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
    End If
End Sub
Public Sub ShowMe(Objfrm As Object)
    Me.Show vbModal, Objfrm
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strOldText As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('����'))" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With Me.cboִ�п���
        .AddItem ""
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then
            .ListIndex = 0
        End If
    End With
    
    
    strOldText = Me.cbo��������.Text
    Me.cbo��������.Clear
    
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        cbo��������.AddItem rsTmp!����
        cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    
    On Error Resume Next
    Me.cbo��������.Text = strOldText
    If cbo��������.ListCount > 0 And Me.cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    
    
    
    
     '�Ա�
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    If Not rsTmp Is Nothing Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    chkConcatenation.Value = zlDatabase.GetPara("�ɼ�����վ�Ǽ�", 100, 1211, 0)
    
    InitDepts = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '����:              �����༭״̬
    'Me.txt����.Enabled = blEnable
    cbo�Ա�.Enabled = blEnable
    txt����.Enabled = blEnable
    txt����1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo��������.Enabled = blEnable
    cboҽ��.Enabled = blEnable
    txtҽ������.Enabled = blEnable
    cmdSelect.Enabled = blEnable
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
'���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
    Dim strSQL As String, i As Long
    Dim strNO As String, str���� As String, lng����ID As Long
    Dim strSeek As String
    
    On Error GoTo errH
    
    If BlnIsNumber(strCode) Then
    'Ԥ�����뵥������
        mblnBarCode = True
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,B.��ҳID,B.���˿���id As ���˿���,B.����ҽ�� As ҽ��," & gConst_������Ϣ_���� & _
            " From ������Ϣ A,����ҽ����¼ B,����ҽ������ C Where A.����ID=B.����ID+0 And B.ID=C.ҽ��ID+0" & _
            " And C.��������=[1]"
        Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
        Exit Function
    End If
    mblnBarCode = False
    
    strSeek = strCode
    '�жϵ�ǰ����ģʽ
    If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) And iInputType = -1 Then 'ˢ��
        iInputType = 0
    ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
        iInputType = 1
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
        iInputType = 2
        strSeek = Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����
        iInputType = 3
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '�Һŵ�
        iInputType = 4
        strSeek = Mid(strCode, 2)
    ElseIf Left(strCode, 1) = "/" Then '�շѵ��ݺ�
        iInputType = 5
        strSeek = Mid(strCode, 2)
    ElseIf Not IsNumeric(Mid(strCode, 2)) Then '��������
        iInputType = 6
        strSeek = Replace(strCode, "(Ӥ��)", "")
    End If
    
    If iInputType = 0 Then 'ˢ��
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���,B.ִ���� As ҽ��," & gConst_������Ϣ_���� & _
            " From ������Ϣ A,���˹Һż�¼ B Where A.���￨��=[1] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and (b.����ID is null or (b.��¼���� =1 and b.��¼״̬ =1)) "
'            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 1 Then '����ID
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,0) As ���˿���," & gConst_������Ϣ_���� & _
            " From ������Ϣ A Where A.����ID=[2]"
    ElseIf iInputType = 2 Then 'סԺ��
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.��Ժ����ID,0),A.��ǰ����id) As ���˿���,B.סԺҽʦ As ҽ��," & gConst_������Ϣ_���� & _
            " From ������Ϣ A,������ҳ B Where A.סԺ��=[2] And A.����ID=B.����ID" ' And A.��ǰ����id IS NOT NULL And B.��Ժ���� Is NULL"
    ElseIf iInputType = 3 Then '�����
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���,B.ִ���� As ҽ��," & gConst_������Ϣ_���� & _
            " From ������Ϣ A,���˹Һż�¼ B Where A.�����=[2] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and (b.����ID is null or (b.��¼���� =1 and b.��¼״̬ =1)) "
'            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
    ElseIf iInputType = 4 Then '�Һŵ�
        strNO = GetFullNO(strSeek, 12)
        strSQL = "Select 1 As PatientType,0 As ��ҳID,Nvl(B.ִ�в���ID,0) As ���˿���,B.ִ���� As ҽ��," & gConst_������Ϣ_���� & _
            " From ������Ϣ A,������ü�¼ B " & _
            " Where B.��¼����=4 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID"
    ElseIf iInputType = 5 Then '�շѵ��ݺ�
        strNO = GetFullNO(strSeek, 13): mstrNO = strNO
        
        strSQL = "Select 1 As PatientType,0 As ��ҳID,B.��������ID As ���˿���,B.������ As ҽ��,B.����,B.�Ա�,B.����," & _
            "A.����ID,A.��λ�绰,A.������λ,A.��λ�ʱ�,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�����,A.���֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ," & _
            "A.����,A.����״��,A.����,A.ְҵ From ������Ϣ A,������ü�¼ B" & _
            " Where Mod(B.��¼����,10)=1 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID(+) Order By B.����ID" ' And B.ҽ����� Is Null"
    Else '��������
        strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,0) As ���˿���," & gConst_������Ϣ_���� & _
            " From ������Ϣ A Where A.����=[1] and 1 = 2 " '�������������Ĳ��˵��²��˴���
    End If
    
    Set GetPatient = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSeek, Val(strSeek), strNO)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SelectDiagItem() As ADODB.Recordset
'ѡ�������Ŀ
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    strSQL = "Select Distinct A.ID,A.����,A.����,nvl(A.���㵥λ,'��') As ���㵥λ,nvl(A.�걾��λ,' ') As �걾��λ," + _
        "Decode(A.���,'H',Decode(A.��������,'1','����ȼ�','������')," + _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨',4,'��ҩ�÷�','����')," + _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','����'),A.��������) As ��Ŀ����,A.��� As ���ID,A.ID As ������ĿID,nvl(ִ��Ƶ��,0) As ִ��Ƶ��ID,nvl(���㷽ʽ,0) As ���㷽ʽID,nvl(ִ�а���,0) As ִ�а���ID,nvl(�Ƽ�����,0) As �Ƽ�����ID,nvl(ִ�п���,0) As ִ�п���ID "
    strSQL = strSQL + "From ������ĿĿ¼ A,������Ŀ���� C,����ִ�п��� D Where A.ID=C.������ĿID And A.ID=D.������ĿID And A.���='C' "   'And D.ִ�п���ID=" & mlngDeptID
    strSQL = strSQL + " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        "And A.������� IN(" & PatientType & ",3) And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And (A.���� Like '" + txtҽ������ + "%' Or Upper(A.����) Like '" + txtҽ������ + "%' Or Upper(C.����) Like '" + UCase(txtҽ������) + "%')"
            
    Call ClientToScreen(txtҽ������.Hwnd, objPoint)
    Set SelectDiagItem = zlDatabase.ShowSelect(Me, strSQL, 0, "ѡ��������Ŀ", True, Me.txtҽ������.Text, "", True, True, True, objPoint.x * 15, objPoint.Y * 15, Me.txtҽ������.Height, False, True)
End Function
Private Function AdviceInput(Optional rsInput As ADODB.Recordset = Nothing) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'���أ�����¼���Ƿ���Ч
    Dim rsTmp As ADODB.Recordset
    Dim strHelpText As String
    Dim strSQL As String
    Dim t_Pati As TYPE_PatiInfoEx
    Dim blnOk As Boolean
    Dim strExtData As String
    
    On Error GoTo errH

    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    If Not rsInput Is Nothing Then txtҽ������.Text = rsInput!����    '��ʱ��ʾ

    '��Ҫ����������ݵ�һЩ��Ŀ
    '---------------------------------------------------------------------------------------------------------------
    '������Ŀѡ�����걾
    strHelpText = "������Ŀ"
    If Not rsInput Is Nothing Then
        strExtData = rsInput!������Ŀid & ";" & rsInput!�걾��λ    '��������Ŀ
    Else
        strExtData = mstrExtData    '��������Ŀ
    End If
    
    On Error Resume Next
    '�ӿڸ��죺int����û�д������ڴ�Ϊ0�� bytUseType ��ǰû�������ڴ�Ϊ0
    blnOk = frmAdviceEditEx.ShowMe(Me, Me.txtҽ������.Hwnd, t_Pati, 0, 4, 0, 1, PatientType, , , , 0, strExtData, , , , , True, mlngDeptID)
    On Error GoTo errH

    If Not blnOk Then Exit Function
    If strExtData = "" Or Mid(strExtData, 1, 1) = ";" Then Exit Function
    
    '��ȡ�ɼ���ʽ
    Set rsTmp = SelectCap(Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp Is Nothing Then
        MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    mlngCapID = rsTmp("ID")
    
    strSQL = "Select C.��Ŀ��� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C " & _
        "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Split(Split(strExtData, ";")(0), ",")(0))
    If rsTmp.EOF Then
        mbln΢������Ŀ = False
    Else
        mbln΢������Ŀ = IIf(Nvl(rsTmp("��Ŀ���"), 0) = 2, True, False)
    End If
    
    mstrExtData = strExtData
    
    
    Call AdviceSet�������(3, mstrExtData)
    txtҽ������.Text = Get�����������(2, "")
    txtҽ������.Text = txtҽ������.Text & "(" & Split(mstrExtData, ";")(1) & ")"
    
    '����ҽ��
    On Error Resume Next
    If Me.cboҽ��.Text = "" Then Me.cboҽ��.ListIndex = 0
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    
    If Len(strNO) >= 8 Then
        GetFullNO = Right(strNO, 8)
        Exit Function
    ElseIf Len(strNO) = 7 Then
        GetFullNO = PreFixNO & strNO
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNO
    
    strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SelectCap(Optional ByVal lngItemID As Long = 0) As ADODB.Recordset
'��ȡ�ɼ���ʽ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim tmpRect As RECT
    
    On Error GoTo DBError
        
    strSQL = "Select Distinct A.ID,A.����,A.���� " + _
        "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
        " And A.���='E' And A.��������='6'" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
        " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
        IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
        " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
        " And D.��ĿID=" & lngItemID
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select Distinct A.ID,A.����,A.���� " + _
            "From ������ĿĿ¼ A Where " + _
            " A.���='E' And A.��������='6'" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
            " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
            IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
            " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set SelectCap = rsTmp
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '���������Ŀ
    strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
    
    If strDataIDs <> "" Then
        If Not rsRelativeAdvice Is Nothing Then
            rsRelativeAdvice.Close
        Else
            Set rsRelativeAdvice = New ADODB.Recordset
        End If
        strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
        "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п���,�������� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        Set rsRelativeAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        If Not rsRelativeAdvice Is Nothing Then rsRelativeAdvice.Close: Set rsRelativeAdvice = Nothing
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
'���ܣ��������ɼ���������ݵ�ҽ������
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
    Dim lngBegin As Long, i As Long
    Dim str���� As String, strTmp As String
    Dim strDate As String
    
    If rsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
        
    rsRelativeAdvice.MoveFirst
    Do While Not rsRelativeAdvice.EOF
        If Len(Trim(rsRelativeAdvice("����"))) > 0 Then
            strTmp = strTmp & "," & rsRelativeAdvice("����")
        End If
        
        rsRelativeAdvice.MoveNext
    Loop
    
    If strTmp <> "" Then
        Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
    Else
        Get����������� = txtMainAdvice
    End If
End Function
'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
    ValidAdvice = True
    
    On Error Resume Next
    If txt����.Text = "" Then
        ValidAdvice = False
        MsgBox "�����벡�˵�������", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.����
        txt����.SetFocus: Exit Function
    End If
    
    If Len(Trim(Me.txtҽ������)) = 0 Then
        ValidAdvice = False
        MsgBox "��������������Ŀ��", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.ҽ������
        Me.txtҽ������.SetFocus: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 Then
        ValidAdvice = False
        MsgBox "��ָ���������ң�", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.��������
        Me.cbo��������.SetFocus: Exit Function
    End If
'    If Me.cboִ�п���.ListIndex = -1 Then
'        ValidAdvice = False
'        MsgBox "��ָ��ִ�п���!", vbInformation, gstrSysName: DoEvents
'        Me.cboִ�п���.SetFocus: Exit Function
'    End If
    If Len(Trim(Me.cboҽ��.Text)) = 0 Then
        ValidAdvice = False
        MsgBox "��ָ������ҽ����", vbInformation, gstrSysName: DoEvents
'        mintFocusItem = FocusItem.ҽ��
        Me.cboҽ��.SetFocus: Exit Function
    End If
End Function
Private Function SaveAdviceData() As Long
    Dim strSQL As String, strDate As String, strNO As String
    Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
    Dim iMaxSeq As Integer, iSendSeq As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim lng��������ID As Long, lng����ID As Long, strDoctor As String, i As Integer
    Dim strִ�п���ID As String, strִ�п���ID1 As String, lngDept As Long
    Dim rsCard As ADODB.Recordset
    Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer, tmpintִ������ As Integer
    Dim rsDept As ADODB.Recordset
    Dim intPatientSource As Integer                     '������Դ
    Dim lngJ As Long, strCostType As String
    
    Dim strAge As String
    Dim strInfo As String
    Dim lngTmp As Long
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    
    '���没����Ϣ
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    If PatientType = 1 Then '���ﲡ��
        If mlng����ID > 0 Then '���еĲ���
'            strSQL = _
                "zl_�ҺŲ��˲���_INSERT(3," & mlng����ID & ",Null," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
                "'�Է�','�Է�'," & _
                "'','',''," & _
                "'','','',0,'','','','',''," & strDate & ",NULL)"
        Else '�²���
            If txt����.Locked = False Then
                strAge = txt����.Text
                If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt����1.Text
                strInfo = CheckAge(strAge)
                If InStr(1, strInfo, "|") > 0 Then
                    lngTmp = Val(Split(strInfo, "|")(0)) '1��ֹ,0��ʾ
                    strInfo = Split(strInfo, "|")(1)
                    If lngTmp = 1 Then
                        MsgBox strInfo, vbInformation, gstrSysName
                        gcnOracle.RollbackTrans
                        If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
                    End If
                End If
            End If
            '��ӻ�ȡĬ�Ϸѱ�
            strSQL = "select ����,ȱʡ��־ from �ѱ� order by ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork")
            Do While Not rsTmp.EOF
                lngJ = lngJ + 1
                If lngJ = 1 Then
                    strCostType = rsTmp("����")
                End If
                If rsTmp("ȱʡ��־") = 1 Then
                    strCostType = rsTmp("����")
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            If strCostType = "" Then strCostType = "�Է�"
            
            mlng����ID = zlDatabase.GetNextNo(1)
            strSQL = _
                "zl_�ҺŲ��˲���_INSERT(1," & mlng����ID & ",Null," & _
                "'',''," & _
                "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
                "'" & strCostType & "','" & strCostType & "'," & _
                "'','',''," & _
                "'','','" & Me.txtUnit.Text & "',0,'','','','',''," & strDate & ",NULL)"
            zlDatabase.ExecuteProcedure strSQL, "������Ϣ����"
        End If
    End If
    '����ҽ��������
    lngAdviceID = zlDatabase.GetNextId("����ҽ����¼")
    iMaxSeq = 0
    
    lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    strDoctor = NeedName(Me.cboҽ��.Text)
    
    If rsRelativeAdvice.RecordCount = 0 Then
        strִ�п���ID = mlngDeptID
    Else
        'PatientType
        If mlng����ID > 0 Then
            strSQL = "select  ִ�п���ID from  ����ִ�п��� where ������Դ = [1] and ������ĿID = [2] "
        Else
            strSQL = "select ִ�п���id from ����ִ�п��� where ������Ŀid = [2]"
        End If
        rsRelativeAdvice.MoveFirst
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, PatientType, CLng(rsRelativeAdvice("Id")))
        strִ�п���ID = Val(Nvl(rsTmp("ִ�п���ID")))
    End If
    
    'ѡ����ִ�п��Ұ�ִ�п��ҽ���
    If Me.cboִ�п���.Text <> "" Then
        strִ�п���ID = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex)
    End If
    
    iSendSeq = 1
    '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
    tmplngClinicID = mlngCapID
    'ȡ�ɼ���ʽ��ִ�в���
    strִ�п���ID1 = UserInfo.����ID
    
    lngSendNO = zlDatabase.GetNextNo(10)
    strNO = zlDatabase.GetNextNo(IIf(PatientType = 2, 14, 13))
    
    '�������ҽ��
    If Not rsRelativeAdvice Is Nothing Then
        i = 2
        rsRelativeAdvice.MoveFirst
        Do While Not rsRelativeAdvice.EOF
            lngTmpID = zlDatabase.GetNextId("����ҽ����¼")
            With rsRelativeAdvice
                strSQL = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                    (iMaxSeq + i) & ",3," & mlng����ID & ",NULL," & _
                    "0,1," & _
                    "1,'" & .Fields("���") & "'," & _
                    .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                    "'" & Replace(.Fields("����"), "'", "''") & "',''," & _
                    "'" & .Fields("�걾��λ") & "','һ����',NULL,NULL,'',NULL," & _
                    .Fields("�Ƽ�����") & "," & _
                    strִ�п���ID & "," & _
                    .Fields("ִ�п���") & ",0," & strDate & ",NULL," & _
                    IIf(Me.txtPatientDept.Tag = 0, lng��������ID, Me.txtPatientDept.Tag) & "," & lng��������ID & ",'" & strDoctor & "'," & _
                    "Sysdate,'',Null)"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                iSendSeq = iSendSeq + 1
                strSQL = "ZL_����ҽ������_Insert(" & _
                    lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                    iSendSeq & ",NULL,NULL,NULL," & _
                    "Sysdate+1/(24*3600)," & _
                    "0," & strִ�п���ID & ",0,0)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                i = i + 1
                .MoveNext
            End With
        Loop
    End If
    '��������Ĳɼ���ʽ�ŵ����
    iMaxSeq = iMaxSeq + 1
    strSQL = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
        iMaxSeq & ",3," & mlng����ID & ",NULL," & _
        "0,1," & _
        "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
        "'" & Replace(Me.txtҽ������, "'", "''") & "',''," & _
        "'','һ����',NULL,NULL,'',NULL,2," & _
        strִ�п���ID1 & ",3,0," & strDate & ",NULL," & _
        IIf(Me.txtPatientDept.Tag = 0, lng��������ID, Me.txtPatientDept.Tag) & "," & lng��������ID & ",'" & strDoctor & "'," & _
        "Sysdate,'',Null)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    iSendSeq = iSendSeq + 1
    '������ҽ��
    strSQL = "ZL_����ҽ������_Insert(" & _
        lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
        iSendSeq & ",NULL,NULL,NULL," & _
        "Sysdate+1/(24*3600)," & _
        "0," & strִ�п���ID & ",0,1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveAdviceData = mlng����ID
    gcnOracle.CommitTrans
    
    Exit Function
ErrHand:
    mlng����ID = 0
    gcnOracle.RollbackTrans
'    Err.Raise Err.Number, "�걾����"
    Exit Function
End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function
