VERSION 5.00
Begin VB.Form frmDueGo 
   Caption         =   "��λ����"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   5085
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   9
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   8
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   100
      Width           =   3495
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   5
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtInvoice 
         Height          =   300
         Left            =   1470
         TabIndex        =   7
         Top             =   1320
         Width           =   1635
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1470
         MaxLength       =   100
         TabIndex        =   3
         Top             =   615
         Width           =   1635
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1470
         MaxLength       =   18
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʵ��ݺ�(&3)"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   1020
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ʊ�ݺ�(&4)"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   1380
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&2)"
         Height          =   180
         Left            =   780
         TabIndex        =   2
         Top             =   675
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&1)"
         Height          =   180
         Left            =   600
         TabIndex        =   0
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmDueGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdOK_Click()
    If txtNO.Text = "" And txtInvoice.Text = "" And txtסԺ��.Text = "" And txt����.Text = "" Then
        MsgBox "�������趨һ��������", vbInformation, gstrSysName
        txtסԺ��.SetFocus: Exit Sub
    End If
    '����:30532
    If InStr(1, txtNO.Text, "[") > 0 Then
        MsgBox "���ݺ��к��÷Ƿ��ַ�[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txtNO.Text, "]") > 0 Then
        MsgBox "���ݺ��к��÷Ƿ��ַ�[]", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
    End If
    If InStr(1, txt����.Text, "[") > 0 Then
        MsgBox "�����к��÷Ƿ��ַ�[]", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If InStr(1, txt����.Text, "]") > 0 Then
        MsgBox "�����к��÷Ƿ��ַ�[]", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    txtסԺ��.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
    If InStr(1, "[]", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub Form_Load()
    gblnOK = False
    txtInvoice.MaxLength = gbytFactLength
End Sub

Private Sub txtInvoice_GotFocus()
    zlControl.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_Validate(Cancel As Boolean)
    txtInvoice.Text = Trim(txtInvoice.Text)
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    txtNO.Text = Trim(txtNO.Text)
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 15)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNO, KeyAscii, m�ı�ʽ
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = Trim(txt����.Text)
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txtסԺ��_Validate(Cancel As Boolean)
    txtסԺ��.Text = Trim(txtסԺ��.Text)
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

