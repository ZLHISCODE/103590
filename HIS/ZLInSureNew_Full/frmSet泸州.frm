VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txt��ַ 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   13
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CheckBox chkԶ�� 
      Caption         =   "ͨ��ҽ��������������֤(&M)"
      Height          =   285
      Left            =   210
      TabIndex        =   11
      Top             =   3030
      Width           =   2775
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽ��������"
      Height          =   1545
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   10
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   9
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   5
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   8
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   4
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4470
      TabIndex        =   15
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4470
      TabIndex        =   14
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fraҽԺ 
      Caption         =   "ҽԺ��Ϣ"
      Height          =   945
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4155
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2595
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽԺ����(&G)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�����ĵ�ַ(&A)"
      Height          =   180
      Left            =   510
      TabIndex        =   12
      Top             =   3420
      Width           =   1350
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private mlng���� As Long
Private mlng���� As Long

Private Sub chkԶ��_Click()
    If chkԶ��.Value = 1 Then
        txt��ַ.BackColor = TxtEdit(Textҽ��������).BackColor
        txt��ַ.Enabled = True
    Else
        txt��ַ.BackColor = Me.BackColor
        txt��ַ.Enabled = False
    End If
End Sub

Private Sub cmb����_Click()
    mblnChange = True
End Sub

Private Sub cmdTest_Click()
    If OraDataOpen(gcn����, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
    
    If TxtEdit(Textҽ������).Tag = TxtEdit(Textҽ������).Text Then
        cmb����.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    cmb����.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    If cmb����.Text = "" Then
        MsgBox "������ҽԺ����", vbInformation, gstrSysName
        If cmb����.Enabled = True Then cmb����.SetFocus
        Exit Function
    End If
    
    For lngCount = TxtEdit.LBound To TxtEdit.UBound
        If zlCommFun.StrIsValid(TxtEdit(lngCount).Text, TxtEdit(lngCount).MaxLength, TxtEdit(lngCount).hwnd) = False Then
            zlControl.TxtSelAll TxtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If txt��ַ.Enabled = True Then
        If zlCommFun.StrIsValid(txt��ַ.Text, , txt��ַ.hwnd, "ҽ�����ĵ�ַ") = False Then
            Exit Function
        End If
        If Trim(txt��ַ.Text) = "" Then
            MsgBox "������ҽ�����ĵ�ַ��", vbInformation, gstrSysName
            zlControl.TxtSelAll txt��ַ
            txt��ַ.SetFocus
            Exit Function
        End If
    End If
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽԺ����','" & cmb����.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ���û���','" & TxtEdit(textҽ���û�).Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ���û�����','" & _
            IIf(TxtEdit(Textҽ������).Tag = "", "", EncryptStr(TxtEdit(Textҽ������).Tag, 256, True)) & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ��������','" & TxtEdit(Textҽ��������).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'���������֤','" & IIf(chkԶ��.Value = 1, "��", "") & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",null,'ҽ�����ĵ�ַ','" & IIf(chkԶ��.Value = 1, txt��ַ.Text, "") & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If gcn����.State = adStateClosed Then
        If OraDataOpen(gcn����, TxtEdit(Textҽ��������).Text, TxtEdit(textҽ���û�).Text, TxtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        TxtEdit(Index).Tag = TxtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If gcn����.State = adStateOpen Then gcn����.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
End Sub

Public Function ��������(ByVal lng���� As Long, ByVal lng���� As Long) As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    mlng���� = lng����
    mlng���� = lng����
    
    On Error GoTo errHandle
    
    cmb����.AddItem "33.���ȼ׼�"
    cmb����.AddItem "32.�����Ҽ�"
    cmb����.AddItem "23.���ȼ׼�"
    cmb����.AddItem "22.�����Ҽ�"
    cmb����.AddItem "13.һ�ȼ׼�"
    cmb����.AddItem "12.һ���Ҽ�"
    cmb����.AddItem "03.����ҽ��"
    cmb����.AddItem "0.�޼� "
    
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����, lng����)
    
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽԺ����"
                SetComboByText cmb����, str����ֵ, False, " "
            Case "���������֤"
                chkԶ��.Value = IIf(str����ֵ = "��", 1, 0)
            Case "ҽ�����ĵ�ַ"
                txt��ַ.Text = str����ֵ
            Case "ҽ���û���"
                TxtEdit(textҽ���û�).Text = str����ֵ
            Case "ҽ��������"
                TxtEdit(Textҽ��������).Text = str����ֵ
            Case "ҽ���û�����"
                TxtEdit(Textҽ������).Text = "        "    '������
                If str����ֵ = "" Then
                    TxtEdit(Textҽ������).Tag = ""
                Else
                    TxtEdit(Textҽ������).Tag = EncryptStr(str����ֵ, 256, False)
                End If
        End Select
        
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    
    �������� = mblnOK
    If mblnOK = False Then
        '����ʧ�ܣ��ر����ӡ��������û�����������û�������������
        If gcn����.State = adStateOpen Then gcn����.Close
    End If
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt��ַ_GotFocus()
    zlControl.TxtSelAll txt��ַ
End Sub
