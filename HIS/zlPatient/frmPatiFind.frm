VERSION 5.00
Begin VB.Form frmPatiFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ����"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   10
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4380
      TabIndex        =   9
      Top             =   270
      Width           =   1100
   End
   Begin VB.Frame fraBdr 
      Height          =   2280
      Left            =   120
      TabIndex        =   11
      Top             =   15
      Width           =   4155
      Begin VB.OptionButton optHead 
         Caption         =   "����"
         Height          =   195
         Left            =   2460
         TabIndex        =   7
         Top             =   1950
         Width           =   660
      End
      Begin VB.OptionButton optCur 
         Caption         =   "����"
         Height          =   195
         Left            =   3180
         TabIndex        =   8
         Top             =   1950
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1095
         Width           =   1110
      End
      Begin VB.TextBox txt���֤ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1515
         Width           =   3030
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   900
         TabIndex        =   4
         Top             =   1095
         Width           =   1110
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         MaxLength       =   18
         TabIndex        =   3
         Top             =   675
         Width           =   1110
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   18
         TabIndex        =   2
         Top             =   675
         Width           =   1110
      End
      Begin VB.TextBox txt���￨ 
         BackColor       =   &H00EBFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2820
         TabIndex        =   1
         Top             =   255
         Width           =   1110
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         TabIndex        =   0
         Top             =   255
         Width           =   1110
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��"
         Height          =   180
         Left            =   2250
         TabIndex        =   18
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label lbl���֤ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   1575
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2250
         TabIndex        =   15
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   300
         TabIndex        =   14
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lbl���￨ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨"
         Height          =   180
         Left            =   2250
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lbl����ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Height          =   180
         Left            =   300
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmPatiFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytType As Byte

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub



Private Sub cmdOK_Click()
    If txt����ID.Text = "" And txt���￨.Text = "" And txt�����.Text = "" And txtסԺ��.Text = "" And txt����.Text = "" And txt����.Text = "" And txt���֤.Text = "" Then
        MsgBox "����������һ����λ������", vbInformation, gstrSysName
        txt����ID.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Activate()
    txt����ID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    If glngSys Like "8??" Then
        lbl�����.Visible = False
        lblסԺ��.Visible = False
        txt�����.Visible = False
        txtסԺ��.Visible = False
        lbl����.Visible = False
        txt����.Visible = False
        
        lbl����ID.Caption = "�ͻ�ID"
        lbl���￨.Caption = "��Ա��"
        
        lbl����.Top = lbl����.Top - 420
        txt����.Top = txt����.Top - 420
        lbl���֤.Top = lbl���֤.Top - 420
        txt���֤.Top = txt���֤.Top - 420
        
        optHead.Top = optHead.Top - 420
        optCur.Top = optCur.Top - 420
        fraBdr.Height = fraBdr.Height - 420
        Me.Height = Me.Height - 420
    End If

    txt���￨.Enabled = gblnShowCard
    If Not txt���￨.Enabled Then txt���￨.BackColor = Me.BackColor
    
    Select Case mbytType
        Case 0 '���в���
        Case 1 '��Ժ����
            txt�����.Enabled = False
            txt�����.BackColor = Me.BackColor
        Case 2 '��Ժ����
            txt�����.Enabled = False
            txt�����.BackColor = Me.BackColor
            txt����.Enabled = False
            txt����.BackColor = Me.BackColor
        Case 3 '���ﲡ��
            txtסԺ��.Enabled = False
            txtסԺ��.BackColor = Me.BackColor
            txt����.Enabled = False
            txt����.BackColor = Me.BackColor
    End Select
End Sub

Private Sub txt����ID_GotFocus()
    zlControl.TxtSelAll txt����ID
End Sub

Private Sub txt����ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("' " & Chr(8), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���￨_GotFocus()
    zlControl.TxtSelAll txt���￨
End Sub

Private Sub txt���￨_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt���֤_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
End Sub
'����29712 by lesfeng 2010-05-11
Private Sub txt����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("[]:��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtסԺ��_GotFocus()
    zlControl.TxtSelAll txtסԺ��
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt���֤_GotFocus()
    zlControl.TxtSelAll txt���֤
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
