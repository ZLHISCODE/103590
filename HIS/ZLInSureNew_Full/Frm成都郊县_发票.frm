VERSION 5.00
Begin VB.Form Frm�ɶ�����_��Ʊ 
   Caption         =   "��Ʊ�Ŵ���"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4395
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��"
      Height          =   405
      Left            =   1680
      TabIndex        =   2
      Top             =   1620
      Width           =   975
   End
   Begin VB.TextBox TXT��Ʊ�� 
      Height          =   405
      Left            =   690
      TabIndex        =   1
      Top             =   930
      Width           =   3105
   End
   Begin VB.Label lbl˵�� 
      Caption         =   "�����뷢Ʊ���룺"
      BeginProperty Font 
         Name            =   "��������"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   2325
   End
End
Attribute VB_Name = "Frm�ɶ�����_��Ʊ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstr���� As String
Public Function ��Ʊ��() As String
    '����:
    '����:lng���� 1 ��ʾ�·ݴ���2 ��ʾ���ڴ���
    '����: ������ȷ������

    Me.Show 1
    ��Ʊ�� = mstr����
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdȷ��_Click()
    If TXT��Ʊ��.Text <> "" Then
       mstr���� = Me.TXT��Ʊ��.Text
       Unload Me
    Else
       MsgBox "�����뷢Ʊ�š�"
       Exit Sub
    End If
End Sub

Private Sub Form_Activate()

    Me.TXT��Ʊ��.SetFocus
    
End Sub


