VERSION 5.00
Begin VB.Form frmLabMBTemplet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����ģ��"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   6
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2310
      TabIndex        =   5
      Top             =   1380
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -30
      TabIndex        =   4
      Top             =   1140
      Width           =   5385
   End
   Begin VB.TextBox txt���� 
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   690
      Width           =   2865
   End
   Begin VB.TextBox txt��� 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   757
      Width           =   450
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "���:"
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   247
      Width           =   450
   End
End
Attribute VB_Name = "frmLabMBTemplet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrResult As String
Public Sub ShowMe(Objfrm As Object, strResult As String)
    mstrResult = strResult
    Me.Show vbModal, Objfrm
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '����ģ��
    
    '��������Ƿ���ȷ
    If Len(Trim(Me.txt���)) < 1 Then
        MsgBox "��������!", vbInformation
        Me.txt���.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Me.txt���) = False Then
        MsgBox "��ű���Ϊ���ݣ����޸�!", vbInformation
        Me.txt���.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txt����)) < 1 Then
        MsgBox "����������!", vbInformation
        Me.txt����.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errH
   
    gstrSql = "Zl_����ø��ģ��_Insert(" & zlDatabase.GetNextId("����ø��ģ��") & "," & Val(Me.txt���) & ",'" & _
              Me.txt���� & "','" & Split(mstrResult, "|")(0) & "','" & Mid(mstrResult, InStr(mstrResult, "|") + 1) & "')"
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    MsgBox "�������!", vbInformation
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "select nvl(max(���),0)+ 1  as ��� from ����ø��ģ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.txt���.Text = rsTmp(0)
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = Len(Me.txt���.Text)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    Else
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = Len(Me.txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
