VERSION 5.00
Begin VB.Form frmSet�˰� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkסԺ��ϸ 
      Caption         =   "סԺ���㲹����ϸ"
      Height          =   300
      Left            =   840
      TabIndex        =   8
      Top             =   1875
      Width           =   1755
   End
   Begin VB.TextBox txt�ȴ� 
      Height          =   300
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "400"
      Top             =   1140
      Width           =   465
   End
   Begin VB.CheckBox chk�Զ� 
      Caption         =   "�Զ�������������(&C)"
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   1620
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   -150
      TabIndex        =   3
      Top             =   810
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   2
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2430
      TabIndex        =   1
      Top             =   2430
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   -150
      TabIndex        =   0
      Top             =   2220
      Width           =   5265
   End
   Begin VB.Label lblInif 
      AutoSize        =   -1  'True
      Caption         =   "��ҽ���з�������ȴ�ʱ��       ��"
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1230
      Width           =   2970
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmSet�˰�.frx":0000
      Top             =   210
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "������صĲ���."
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   5
      Top             =   420
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet�˰�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    If Val(txt�ȴ�.Text) <= 10 Then
        ShowMsgbox "�ȴ�ʱ�䲻��С��10��"
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�˰� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˰� & ",null,'�Զ���������','" & IIf(chk�Զ�.Value = 1, 1, 0) & "',0)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˰� & ",null,'����ȴ�ʱ��','" & Val(txt�ȴ�.Text) & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�˰� & ",null,'סԺ���㲹����ϸ','" & IIf(chkסԺ��ϸ.Value = 1, 1, 0) & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�˰�
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "�Զ���������"
                chk�Զ�.Value = IIf(Nvl(!����ֵ, 1) = 1, 1, 0)
            Case "����ȴ�ʱ��"
                txt�ȴ�.Text = Nvl(!����ֵ, 400)
            Case "סԺ���㲹����ϸ"
                chkסԺ��ϸ.Value = IIf(Nvl(!����ֵ, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With

End Sub

Public Function ��������() As Boolean
    frmSet�˰�.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function
End Function


Private Sub txt�ȴ�_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�ȴ�, KeyAscii, m����ʽ
End Sub
