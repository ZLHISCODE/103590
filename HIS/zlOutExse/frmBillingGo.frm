VERSION 5.00
Begin VB.Form frmBillingGo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ����"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   4770
      Begin VB.OptionButton optHead 
         Caption         =   "����"
         Height          =   195
         Left            =   2985
         TabIndex        =   3
         Top             =   750
         Width           =   660
      End
      Begin VB.OptionButton optCur 
         Caption         =   "����"
         Height          =   195
         Left            =   3705
         TabIndex        =   4
         Top             =   750
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   930
         MaxLength       =   100
         TabIndex        =   2
         Top             =   675
         Width           =   1365
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   8
         TabIndex        =   0
         Top             =   255
         Width           =   1365
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         MaxLength       =   15
         TabIndex        =   1
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   525
         TabIndex        =   10
         Top             =   735
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   345
         TabIndex        =   9
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   180
         Left            =   2520
         TabIndex        =   8
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2400
      TabIndex        =   5
      Top             =   1275
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   1275
      Width           =   1100
   End
End
Attribute VB_Name = "frmBillingGo"
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
    If txtNO.Text = "" And txt����ID.Text = "" And txt����.Text = "" Then
        MsgBox "�������趨һ��������", vbInformation, gstrSysName
        txtNO.SetFocus: Exit Sub
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
    txtNO.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If InStr(1, "[]", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    gblnOK = False
End Sub

Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNO, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNO_LostFocus()
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 14)
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
