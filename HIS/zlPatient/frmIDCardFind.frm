VERSION 5.00
Begin VB.Form frmIDCardFind 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��λ����"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   90
      TabIndex        =   9
      Top             =   15
      Width           =   4920
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   990
         MaxLength       =   8
         TabIndex        =   0
         Top             =   255
         Width           =   1275
      End
      Begin VB.TextBox txtCard 
         BackColor       =   &H00EBFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   990
         TabIndex        =   1
         Top             =   690
         Width           =   1275
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   1560
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3225
         MaxLength       =   10
         TabIndex        =   4
         Top             =   675
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Height          =   570
         Left            =   2880
         TabIndex        =   15
         Top             =   975
         Width           =   1950
         Begin VB.OptionButton optCur 
            Caption         =   "����"
            Height          =   195
            Left            =   945
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton optHead 
            Caption         =   "����"
            Height          =   195
            Left            =   225
            TabIndex        =   5
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3225
         MaxLength       =   18
         TabIndex        =   3
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   345
         TabIndex        =   14
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   525
         TabIndex        =   13
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   345
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2760
         TabIndex        =   10
         Top             =   735
         Width           =   360
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2580
         TabIndex        =   11
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2430
      TabIndex        =   7
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3750
      TabIndex        =   8
      Top             =   1755
      Width           =   1100
   End
End
Attribute VB_Name = "frmIDCardFind"
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
    If txtNO.Text = "" And txtCard.Text = "" And cbo����Ա.ListIndex = 0 And txt����.Text = "" And txtסԺ��.Text = "" Then
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
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    
    gblnOK = False

    If Not gblnShowCard Then txtCard.PasswordChar = "*"
    'txtCard.MaxLength = gbytCardNOLen
    
    cbo����Ա.AddItem ""
    cbo����Ա.ListIndex = 0
    
    
    Set rsTmp = GetPersonnel("�����Ǽ���", True)
    For i = 1 To rsTmp.RecordCount
        cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Next
    
    txtCard.Enabled = gblnShowCard
    If Not txtCard.Enabled Then txtCard.BackColor = Me.BackColor
    
    If glngSys Like "8??" Then
        lblסԺ��.Visible = False
        txtסԺ��.Visible = False
    End If
End Sub

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����Ա.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����Ա.hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����Ա.ListIndex = lngIdx
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount <> 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub txtCard_GotFocus()
    SelAll txtCard
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
        If Len(txtCard.Text) = gbytCardNOLen - 1 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtNO_GotFocus()
    SelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text) Or txtNO.SelStart = 0) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

Private Sub txtNO_LostFocus()
    If txtNO.Text <> "" Then txtNO.Text = GetFullNO(txtNO.Text, 16)
End Sub

Private Sub txt����_GotFocus()
    SelAll txt����
End Sub
'����29712 by lesfeng 2010-05-11
Private Sub txt����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("[]:��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtסԺ��_GotFocus()
    SelAll txtסԺ��
End Sub

Private Sub txtסԺ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
