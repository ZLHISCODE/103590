VERSION 5.00
Begin VB.Form frmSet��Ϫũҽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmSet��Ϫũҽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtIC�˿ں� 
      Height          =   300
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   1560
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2790
      TabIndex        =   9
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   8
      Top             =   2100
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   350
      Left            =   180
      TabIndex        =   7
      Top             =   2100
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "ǰ�û�IP���˿�����"
      Height          =   1275
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txt�˿ں� 
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "8801"
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox txtIP��ַ 
         Height          =   300
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "192.168.168.168"
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lbl�˿ں� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�˿ں�(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   3
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblIP��ַ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IP��ַ(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   540
         TabIndex        =   1
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.Label lblIC�˿ں� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IC�˿ں�(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1620
      Width           =   990
   End
End
Attribute VB_Name = "frmSet��Ϫũҽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    If Trim(txtIC�˿ں�.Text) = "" Then
        MsgBox "�˿ںŲ���Ϊ�գ�", vbInformation, gstrSysName
        txt�˿ں�.SetFocus
        Exit Sub
    End If
    If Val(txtIC�˿ں�.Text) < 1 Or Val(txtIC�˿ں�.Text) > 5 Then
        MsgBox "�˿ںŲ���С��1���ߴ���5", vbInformation, gstrSysName
        txt�˿ں�.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_��Ϫũҽ & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��Ϫũҽ & ",NULL,'IP��ַ','" & txtIP��ַ.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��Ϫũҽ & ",NULL,'�˿ں�','" & txt�˿ں�.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_��Ϫũҽ & ",NULL,'IC�˿ں�','" & txtIC�˿ں�.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTest_Click()
    Dim lngPort As Long
    Dim strIP As String
    '�����Ƿ����ӵ�ͨ
    strIP = txtIP��ַ.Text
    lngPort = Val(txt�˿ں�.Text)
    Call CXNY_SetRemoteServerAddr(lngPort, strIP)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ղ���", TYPE_��Ϫũҽ)
    
    With rsTemp
        Do While Not .EOF
            Select Case !������
            Case "IP��ַ"
                txtIP��ַ.Text = Nvl(!����ֵ, "127.0.0.1")
            Case "�˿ں�"
                txt�˿ں�.Text = Nvl(!����ֵ, "8801")
            Case "IC�˿ں�"
                txtIC�˿ں�.Text = Nvl(!����ֵ, 1)
            End Select
            .MoveNext
        Loop
    End With
End Sub

Private Sub txtIP��ַ_GotFocus()
    txtIP��ַ.SelStart = 0
    txtIP��ַ.SelLength = Len(txtIP��ַ.Text)
End Sub

Private Sub txtIP��ַ_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub txtIP��ַ_Validate(Cancel As Boolean)
    Dim arrIP
    Dim intCOUNT As Integer, intDO As Integer
    '���IP�����Ƿ�Ϸ�
    If Trim(txtIP��ַ.Text) = "" Then Exit Sub
    
    On Error GoTo errHand
    arrIP = Split(txtIP��ַ.Text, ".")
    intCOUNT = UBound(arrIP)
    If intCOUNT > 3 Then
        MsgBox "������Ϸ���IP��ַ��", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    
    For intDO = 0 To 3
        If Val(arrIP(intDO)) > 255 Then
            MsgBox "������Ϸ���IP��ַ��", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    Next
    Exit Sub
errHand:
    Cancel = True
    MsgBox "������Ϸ���IP��ַ��", vbInformation, gstrSysName
End Sub

Private Sub Txt�˿ں�_GotFocus()
    txt�˿ں�.SelStart = 0
    txt�˿ں�.SelLength = Len(txt�˿ں�.Text)
End Sub

Private Sub txt�˿ں�_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Public Function ShowME() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowME = mblnReturn
End Function
