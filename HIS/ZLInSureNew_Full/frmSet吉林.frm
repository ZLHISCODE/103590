VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSet����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   5265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3420
      TabIndex        =   2
      Top             =   2340
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   5265
   End
   Begin VB.CheckBox chk��ϸ 
      Caption         =   "��ϸʱʵ�ϴ�(&D)"
      Height          =   210
      Left            =   870
      TabIndex        =   0
      Top             =   1500
      Width           =   3375
   End
   Begin VB.Label lbl 
      Caption         =   "������صĲ���."
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   570
      Width           =   7125
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   270
      Picture         =   "frmSet����.frx":014A
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmSet����"
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
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���� & ",null,'��ϸʱʵ�ϴ�','" & IIf(chk��ϸ.Value = 1, 1, 0) & "',0)"
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
    
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_����
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "��ϸʱʵ�ϴ�"
                chk��ϸ.Value = IIf(Nvl(!����ֵ, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With

End Sub

Public Function ��������() As Boolean
    frmSet����.Show vbModal, frmҽ�����
    �������� = mblnOK
    Exit Function
End Function

