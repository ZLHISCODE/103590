VERSION 5.00
Begin VB.Form frmDiagTitle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ο���������"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkProof 
      Caption         =   "�ñ���ο�Ϊ��֤����(&S)"
      Height          =   285
      Left            =   1650
      TabIndex        =   4
      Top             =   1770
      Width           =   2940
   End
   Begin VB.Frame fraTier 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1650
      TabIndex        =   12
      Top             =   1470
      Width           =   2940
      Begin VB.OptionButton optTier 
         Caption         =   "һ������(&1)"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTier 
         Caption         =   "��������(&2)"
         Height          =   210
         Index           =   1
         Left            =   1635
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkMethod 
      Caption         =   "�����ö�Ӧ�����ƴ�ʩ(&M)"
      Height          =   285
      Left            =   1650
      TabIndex        =   5
      Top             =   2100
      Width           =   2940
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "����"
      Top             =   975
      Width           =   2940
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   7
      Top             =   2715
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2730
      TabIndex        =   6
      Top             =   2715
      Width           =   1100
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -30
      TabIndex        =   9
      Top             =   2535
      Width           =   5745
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2715
      Width           =   1100
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   975
      TabIndex        =   0
      Top             =   1050
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmDiagTitle.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "������ϲο����ݵ�С���⣬���ڲο��Ķ����㣻�����Ͳο����ݺ����Ǻϡ�"
      Height          =   345
      Left            =   975
      TabIndex        =   11
      Top             =   210
      Width           =   4170
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "��ҽ"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   405
      TabIndex        =   10
      Top             =   555
      Width           =   360
   End
End
Attribute VB_Name = "frmDiagTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strLefts As String   '�Ѿ����ڵ�ǰ��ı���
Public strRights As String  '�Ѿ����ڵĺ���ı���
Public strTitle As String   '�༭�����ı���
Dim intCount As Integer

Private Sub chkMethod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkProof_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    strTitle = ""
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim aryItems() As String
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "�����������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If Me.chkProof.Value = 1 And Me.optTier(1).Value Then
        If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > 4 Then
            MsgBox "������֤���ⲻ�ܳ���4�ĳ�������", vbExclamation, gstrSysName
            Me.txtName.SetFocus
            Exit Sub
        End If
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���ⳬ��" & Me.txtName.MaxLength & "�ĳ�������", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    '�ظ��Լ��
    aryItems = Split(Mid(strLefts & strRights, 2), ";")
    For intCount = LBound(aryItems) To UBound(aryItems)
        If Split(aryItems(intCount), ",")(2) = Trim(Me.txtName.Text) Then
            MsgBox "�ñ����Ѿ������ڲο���", vbExclamation, gstrSysName
            Me.txtName.SetFocus
            Exit Sub
        End If
    Next
    '���涨��ʽ��֯�༭����Ŀ
    strTitle = Me.Tag & Trim(Me.txtName.Text) & "," & _
            IIf(Me.chkProof.Value = 1, 1, 0) & "," & IIf(Me.optTier(0).Value, 1, 2) & "," & IIf(Me.chkMethod.Value = 1, 1, 0)
    Me.Hide
End Sub

Private Sub optTier_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub


