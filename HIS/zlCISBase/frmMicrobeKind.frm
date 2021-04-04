VERSION 5.00
Begin VB.Form frmMicrobeKind 
   BorderStyle     =   0  'None
   Caption         =   "ϸ������"
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtӢ�� 
      Height          =   300
      Left            =   975
      MaxLength       =   60
      TabIndex        =   6
      Top             =   915
      Width           =   3870
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   975
      MaxLength       =   60
      TabIndex        =   2
      Top             =   510
      Width           =   3870
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   975
      MaxLength       =   13
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
   Begin VB.TextBox txt��д 
      Height          =   300
      Left            =   975
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label lblӢ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   975
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   585
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ͱ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   195
      Width           =   720
   End
   Begin VB.Label lbl��д 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ����д"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   1365
      Width           =   720
   End
End
Attribute VB_Name = "frmMicrobeKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngKindId As Long          '��ǰ��ʾ������id

Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function zlRefresh(lngKindId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    mlngKindId = lngKindId
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = ""
    If lngKindId = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ����, ��������, Ӣ������, ���� From ����ϸ������ Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKindId)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("��������").DefinedSize
        Me.txtӢ��.MaxLength = .Fields("Ӣ������").DefinedSize
        Me.txt��д.MaxLength = .Fields("����").DefinedSize
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !��������
            Me.txtӢ��.Text = "" & !Ӣ������
            Me.txt��д.Text = "" & !����
        End If
    End With
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngKindId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngKindId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
        
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(����), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From ����ϸ������"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlEditStart")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            
'            Call SQLTest
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
        End With
        
        '���������Ĭ��ֵ
        Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = ""
    End If

    mlngKindId = lngKindId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250)
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Call Me.zlRefresh(mlngKindId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "�������������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "�������Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtӢ��.Text), vbFromUnicode)) > Me.txtӢ��.MaxLength Then
        MsgBox "Ӣ�����Ƴ��������" & Me.txtӢ��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txtӢ��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��д.Text), vbFromUnicode)) > Me.txt��д.MaxLength Then
        MsgBox "��д���������" & Me.txt��д.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt��д.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '���ݱ��������֯
    If Me.Tag = "����" Then
        lngNewId = zldatabase.GetNextId("����ϸ������")
    Else
        lngNewId = mlngKindId
    End If

    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txtӢ��.Text) & "','" & Trim(Me.txt��д.Text) & "'"
    
    If Me.Tag = "����" Then
        gstrSql = "Zl_����ϸ������_Insert(" & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_����ϸ������_Update(" & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngKindId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    zlEditSave = mlngKindId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
 
Private Sub Form_Load()
    mlngKindId = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��д_GotFocus()
    Me.txt��д.SelStart = 0: Me.txt��д.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��д_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
