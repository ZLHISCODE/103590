VERSION 5.00
Begin VB.Form Frm��ɽ_��ʾ 
   Caption         =   "��ʾ"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3405
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Cmdȷ�� 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txt˵�� 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Lbl˵�� 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frm��ɽ_��ʾ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstr���� As String
Public Function �������ڸ���_��ɽ(ByVal lng���� As Long, ByVal str˵�� As String) As String
    '����:
    '����:lng���� 1 ��ʾ�·ݴ���2 ��ʾ���ڴ���
    '����: ������ȷ������

    If lng���� = 1 Then
       Me.Lbl˵��.Caption = "�ò��˵ĳ����·ݣ�����Ϊ��" & str˵�� & "��������������ȷ��ֵ��"
    Else
       Me.Lbl˵��.Caption = "�ò��˵ĳ������ڣ�����Ϊ��" & str˵�� & "��������������ȷ��ֵ��"
    End If
    Me.Show 1
    �������ڸ���_��ɽ = mstr����
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Cmdȷ��_Click()
    If txt˵��.Text <> "" Then
       mstr���� = Me.txt˵��.Text
       Unload Me
    Else
       MsgBox "��������ȷ����ֵ��"
       Exit Sub
    End If
End Sub

Private Sub Form_Activate()

    Me.txt˵��.SetFocus
    
End Sub

