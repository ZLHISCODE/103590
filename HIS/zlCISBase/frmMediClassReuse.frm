VERSION 5.00
Begin VB.Form frmMediClassReuse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ��������"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3795
   Icon            =   "frmMediClassReuse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ø÷���Ŀ¼ʱ�Ƿ�ͬʱ�������²���"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chk���ù�� 
         Caption         =   "���ø÷��������й��"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chk����Ʒ�� 
         Caption         =   "���ø÷���������Ʒ��"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chk������Ŀ¼ 
         Caption         =   "���ø÷�����������Ŀ¼"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmMediClassReuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng����id As Long
Private mstr���� As String


Public Sub ShowForm(ByVal lng����id As Long, ByVal str���� As String)
    mlng����id = lng����id
    mstr���� = str����
    
    frmMediClassReuse.Show vbModal
    Exit Sub
End Sub


Private Sub chk����Ʒ��_Click()
    If chk����Ʒ��.Value = 1 Then
        chk���ù��.Enabled = True
    Else
        chk���ù��.Value = 0
        chk���ù��.Enabled = False
    End If
End Sub

Private Sub chk������Ŀ¼_Click()
    If chk������Ŀ¼.Value = 1 Then
        chk����Ʒ��.Enabled = True
    Else
        chk����Ʒ��.Value = 0
        chk����Ʒ��.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim int������Ŀ¼ As Integer
    Dim int����Ʒ�� As Integer
    Dim int���ù�� As Integer
    
    int������Ŀ¼ = chk������Ŀ¼.Value
    
    If chk����Ʒ��.Enabled Then
        int����Ʒ�� = chk����Ʒ��.Value
    End If
    
    If chk���ù��.Enabled Then
        int���ù�� = chk���ù��.Value
    End If
    
    On Error GoTo ErrHand
    
    gstrSql = "Zl_���Ʒ���Ŀ¼_ҩƷ��������(" & mlng����id & "," & Val(mstr����) & "," & int������Ŀ¼ & "," & int����Ʒ�� & "," & int���ù�� & " )"
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


