VERSION 5.00
Begin VB.Form frmSet�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Caption         =   "����������"
      Height          =   1545
      Left            =   60
      TabIndex        =   21
      Top             =   2625
      Width           =   4695
      Begin VB.CommandButton cmdBA 
         Caption         =   "����(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   17
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   12
         Top             =   330
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1110
         Width           =   2145
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&A)"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   11
         Top             =   390
         Width           =   735
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   13
         Top             =   780
         Width           =   555
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����Դ(&C)"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   15
         Top             =   1170
         Width           =   735
      End
   End
   Begin VB.TextBox txtNumber 
      Height          =   300
      Left            =   1380
      TabIndex        =   10
      Top             =   2175
      Width           =   3345
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1545
      Left            =   60
      TabIndex        =   20
      Top             =   90
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1110
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   1
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   6
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����Դ(&S)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   4
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   19
      Top             =   765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4890
      TabIndex        =   18
      Top             =   285
      Width           =   1100
   End
   Begin VB.ComboBox cbo���õ��� 
      Height          =   300
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1755
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ͳ������(&N)"
      Height          =   195
      Left            =   315
      TabIndex        =   9
      Top             =   2250
      Width           =   930
   End
   Begin VB.Label lbl���õ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���õ���(&Q)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   315
      TabIndex        =   7
      Top             =   1815
      Width           =   930
   End
End
Attribute VB_Name = "frmSet��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mblnChangePassword As Boolean  '���뱻�޸Ĺ�
Private cnTest As ADODB.Connection

Private Sub cmdBA_Click()
    Set cnTest = New ADODB.Connection
    If cnTest.State = adStateOpen Then cnTest.Close
    On Error Resume Next
    cnTest.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(txtEdit(4).Text) & ";User ID=" & Trim(txtEdit(5).Text) & ";Data Source=" & Trim(txtEdit(3).Text) & ";Persist Security Info=True"
    cnTest.CursorLocation = adUseClient
    cnTest.Open
    If Err <> 0 Then
        MsgBox "��������������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    MsgBox "�������������ӳɹ�", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
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

Private Sub cmdTest_Click()
    If gcn����.State = adStateOpen Then gcn����.Close
    On Error Resume Next
    If cbo���õ���.ListIndex = 1 Then
        gcn����.ConnectionString = "Provider=MSDASQL.1;Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text)
    Else
        gcn����.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(txtEdit(1).Tag) & ";User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text) & ";Persist Security Info=True"
    End If
    gcn����.CursorLocation = adUseClient
    gcn����.Open
    If Err <> 0 Then
        MsgBox "ҽ��ǰ�÷���������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    MsgBox "ҽ��ǰ�÷��������ӳɹ�", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    '���ж��ַ��ĺϷ���
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    '�����ӽ��в���
    If gcn����.State = adStateClosed Then
        On Error Resume Next
        If gcn����.State = adStateOpen Then gcn����.Close
        If cbo���õ���.ListIndex = 1 Then
            gcn����.ConnectionString = "Provider=MSDASQL.1;Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text)
        Else
            gcn����.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(txtEdit(1).Tag) & ";User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text) & ";Persist Security Info=True"
        End If
        gcn����.CursorLocation = adUseClient
        gcn����.Open
        
        If Err <> 0 Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    IsValid = True
End Function

Public Function ��������() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    Dim int���õ��� As Integer
    
    mblnOK = False
    On Error GoTo errHandle
    
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_��������)
    
    int���õ��� = 0
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "�û���"
                txtEdit(0).Text = str����ֵ
            Case "������"
                txtEdit(2).Text = str����ֵ
            Case "�û�����"
                txtEdit(1).Text = "        "    '������
                txtEdit(1).Tag = str����ֵ
            Case "���õ���"
                int���õ��� = Val(str����ֵ)
            Case "ͳ������"
                txtNumber = str����ֵ
            Case "�����û���"
                txtEdit(5).Text = str����ֵ
            Case "�����û�����"
                txtEdit(4).Text = str����ֵ
            Case "����������"
                txtEdit(3).Text = str����ֵ
        End Select
        rsTemp.MoveNext
    Loop
'    If txtEdit(4).Text = "" Then txtEdit(4).Enabled = True
    On Error Resume Next
    With cbo���õ���
        .Clear
        .AddItem "����ʹ��(ORACLE����)"
        .AddItem "�۷�ְ��ҽԺ(SYBASE����)"
        .ListIndex = int���õ���
    End With
    
    mblnChange = False
    mblnChangePassword = False
    frmSet��������.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�������� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'�û���','" & txtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'�û�����','" & txtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'������','" & txtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'���õ���','" & cbo���õ���.ListIndex & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_insert(" & TYPE_�������� & ",null,'ͳ������','" & txtNumber.Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'�����û���','" & txtEdit(5).Text & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'�����û�����','" & txtEdit(4).Text & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�������� & ",null,'����������','" & txtEdit(3).Text & "',8)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
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
    If Index = 1 Then
        txtEdit(1).Tag = txtEdit(1).Text
        mblnChangePassword = True
    End If
    
    '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
    If gcn����.State = adStateOpen Then gcn����.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub
