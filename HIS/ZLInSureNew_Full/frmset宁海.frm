VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmset����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   8
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3450
      TabIndex        =   9
      Top             =   1950
      Width           =   1100
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1605
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   4515
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1935
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3330
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
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
Dim mblnTest As Boolean
Dim mcnTest As New ADODB.Connection

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    If Not mblnTest Then MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
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

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
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
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
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
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Public Function ��������() As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    Dim strҽԺ���� As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    'ȡ���ղ���
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����)
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ���û���"
                txtEdit(textҽ���û�) = str����ֵ
            Case "ҽ��������"
                txtEdit(Textҽ��������) = str����ֵ
            Case "ҽ���û�����"
                txtEdit(Textҽ������).Text = "        "    '������
                txtEdit(Textҽ������).Tag = str����ֵ
        End Select
        rsTemp.MoveNext
    Loop
    
    mblnChange = False
    frmSet����.Show vbModal, frmҽ�����
    
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
