VERSION 5.00
Begin VB.Form frmPacsInterfacePar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPacsInterfacePar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lstPar 
      Height          =   3765
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   375
      Left            =   1125
      TabIndex        =   0
      Top             =   4140
      Width           =   1095
   End
End
Attribute VB_Name = "frmPacsInterfacePar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsOk As Boolean
Private mstrSelectPar As String

Public Function ShowMe(objParent As Object) As String
    Call Me.Show(1, objParent)
    
    If mblnIsOk Then
        ShowMe = mstrSelectPar
    End If
End Function

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    mblnIsOk = False
    
    Call Me.Hide
Exit Sub
errHandle:
    'If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim i As Long
    
    For i = 0 To lstPar.ListCount - 1
        If lstPar.Selected(i) Then
            mstrSelectPar = lstPar.list(i)
            Exit For
        End If
    Next i
    
    mblnIsOk = True
    Call Me.Hide
Exit Sub
errHandle:
    'If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    mblnIsOk = False
    
    Call LoadInputPar
Exit Sub
errHandle:
    'If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadInputPar()
'�����ѡ���¼�����
    Dim i As Long
    
    lstPar.AddItem "[[�û���]]"
    lstPar.AddItem "[[ϵͳ��]]"
    lstPar.AddItem "[[ģ���]]"
    lstPar.AddItem "[[����ID]]"
    lstPar.AddItem "[[����ID]]"
    lstPar.AddItem "[[ҽ��ID]]"
    lstPar.AddItem "[[���ͺ�]]"
    lstPar.AddItem "[[����]]"
    lstPar.AddItem "[[�����]]"
    lstPar.AddItem "[[סԺ��]]"
    lstPar.AddItem "[[���֤��]]"
    lstPar.AddItem "[[Ӱ�����]]"
End Sub

Private Sub lstPar_DblClick()
    Call cmdSure_Click
End Sub
