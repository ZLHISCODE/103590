VERSION 5.00
Begin VB.Form frmҩƷ����_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ����"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmҩƷ����_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdҩƷ 
      Caption         =   "��"
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txt��ע 
      Height          =   2085
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   7125
   End
   Begin VB.TextBox txt��λ 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   893
      Width           =   1245
   End
   Begin VB.TextBox txt��� 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5310
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   53
      Width           =   2685
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   7125
   End
   Begin VB.TextBox txt�ۼ� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   893
      Width           =   1305
   End
   Begin VB.TextBox txt���� 
      Height          =   315
      Left            =   4590
      TabIndex        =   5
      Top             =   893
      Width           =   975
   End
   Begin VB.TextBox txt��� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   893
      Width           =   1515
   End
   Begin VB.TextBox txtҩƷ���� 
      Height          =   315
      Left            =   870
      TabIndex        =   0
      Top             =   53
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   315
      Left            =   5850
      TabIndex        =   8
      Top             =   3690
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   315
      Left            =   6900
      TabIndex        =   9
      Top             =   3690
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   30
      TabIndex        =   10
      Top             =   3450
      Width           =   8175
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "�������漰������������׼������ʱ����ָ��������������������óɰ�������ʽ����ʱ��ϵͳ���Զ�����"
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   8
      Left            =   60
      TabIndex        =   20
      Top             =   3630
      Width           =   5430
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ע"
      Height          =   180
      Index           =   7
      Left            =   480
      TabIndex        =   18
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ۼ۽��"
      Enabled         =   0   'False
      Height          =   180
      Index           =   6
      Left            =   5790
      TabIndex        =   17
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ۼ�"
      Enabled         =   0   'False
      Height          =   180
      Index           =   5
      Left            =   2370
      TabIndex        =   16
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ۼ۵�λ"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   15
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   4260
      TabIndex        =   14
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   510
      TabIndex        =   13
      Top             =   540
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmҩƷ����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    If txtҩƷ����.Tag = "" Then ShowMsgbox "��ȷ��ҩƷ��Ϣ��": txtҩƷ����.SetFocus: Exit Sub
    If Val(txt����.Text) <= 0 Then ShowMsgbox "ҩƷ�����������0��": txt����.SetFocus: Exit Sub
    gstrSQL = "Zl_��ҩ����Ŀ¼_����_Insert(" & mint���� & ",'" & txtҩƷ����.Tag & "','" & txt����.Text & "','" & txt��ע.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call frm�������Լ�����ҩ����_����.mnuViewRefresh_Click
    ShowMsgbox "ҩƷ�������óɹ���"
    If Me.Tag = "����" Then
        txtҩƷ����.Text = "": txt��ע.Text = "": txt����.Text = "": txt���.Text = ""
        txt���.Text = ""
        txt���.Text = ""
        txt��λ.Text = "": txt�ۼ�.Text = ""
        txt����.Text = 1: txtҩƷ����.Tag = "": txtҩƷ����.SetFocus
    Else
        Unload Me
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdҩƷ_Click()
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select  ID, Null As �ϼ�id, 0 As ĩ��, ���� As ����,  ����, ����, Null As ���, Null As ����," & _
              "      Null As �ۼ۵�λ, -null As �ۼ� From ҩƷ��;���� Union All " & _
              "Select B.ҩƷid As ID, A.��;����id As �ϼ�id, 1 As ĩ��, D.���� As ����, B.����, B.����, B.���, B.����, B.�ۼ۵�λ," & _
              "      C.�ּ� As �ۼ� From ҩƷ��Ϣ A, ҩƷĿ¼ B, �շѼ�Ŀ C, ҩƷ��;���� D " & _
              "Where A.��;����id = D.ID And A.ҩ��id = B.ҩ��id And B.ҩƷid = C.�շ�ϸĿid And " & _
              "     (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 2, "ҩƷĿ¼", , txtҩƷ����.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txtҩƷ����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
            txt���.Text = Nvl(rsTemp!���): txt����.Text = Nvl(rsTemp!����)
            txt��λ.Text = Nvl(rsTemp!�ۼ۵�λ): txt�ۼ�.Text = Format(Nvl(rsTemp!�ۼ�), "0.00000")
            txt����.Text = 1: txtҩƷ����.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt����.Text = 0: txtҩƷ����.Tag = ""
        End If
    Else
        txt����.Text = 0: txtҩƷ����.Tag = ""
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Sub ShowME(ByVal intinsure As Integer, ByVal strMode As String, ByVal strҩƷID As String)
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    Me.Caption = Me.Caption & "-" & strMode
    Me.Tag = strMode
    mint���� = intinsure
    If strMode = "�޸�" Then
        gstrSQL = "Select A.ҩƷID,A.����, A.����, A.���, A.����, A.�ۼ۵�λ, trim(to_char(B.����,'900090.00')) As ����, " & _
              "      trim(to_char(C.�ּ�,'900090.00000'))  As �ۼ�, trim(to_char(Nvl(B.����, 0) * Nvl(C.�ּ�, 0),'90009990.00')) As �ۼ۽��,B.��ע " & _
              "From ҩƷĿ¼ A, ��ҩ����Ŀ¼_���� B, �շѼ�Ŀ C " & _
              "Where A.ҩƷid = B.ҩƷid And B.ҩƷid = C.�շ�ϸĿID And B.����=[1]" & _
              " And (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.ҩƷID='" & strҩƷID & "'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
        txtҩƷ����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
        txt���.Text = Nvl(rsTemp!���): txt����.Text = Nvl(rsTemp!����)
        txt��λ.Text = Nvl(rsTemp!�ۼ۵�λ): txt�ۼ�.Text = Format(Nvl(rsTemp!�ۼ�), "0.00000")
        txt����.Text = Nvl(rsTemp!����): txtҩƷ����.Tag = Nvl(rsTemp!ҩƷid)
        txt��ע.Text = Nvl(rsTemp!��ע): txtҩƷ����.Enabled = False: cmdҩƷ.Visible = False
    End If
    Me.Show 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_Change()
    txt���.Text = Format(Val(txt�ۼ�.Text) * Val(txt����.Text), "0.00000")
End Sub

Private Sub txt����_GotFocus()
     zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt����, KeyAscii, m���ʽ)
End Sub

Private Sub txtҩƷ����_GotFocus()
    zlControl.TxtSelAll txtҩƷ����
End Sub

Private Sub txtҩƷ����_KeyPress(KeyAscii As Integer)
     On Error GoTo errHand
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Distinct B.ҩƷid As ID,Null As �ϼ�ID,D.���� As ����, B.����, B.����, B.���, B.����," & _
              "       B.�ۼ۵�λ, C.�ּ� As �ۼ� From ҩƷ��Ϣ A, ҩƷĿ¼ B, �շѼ�Ŀ C, ҩƷ��;���� D, �շѱ��� E " & _
              "Where A.��;����id = D.ID And A.ҩ��id = B.ҩ��id And B.ҩƷid = C.�շ�ϸĿid And " & _
              "     (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.ҩƷid = E.�շ�ϸĿid " & _
              "     And (Upper(B.����) Like '%" & UCase(txtҩƷ����.Text) & "%' Or  Upper(B.����) Like '%" & UCase(txtҩƷ����.Text) & "%' " & _
              "     Or Upper(E.����) Like '%" & UCase(txtҩƷ����.Text) & "%')"
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҩƷĿ¼", , txtҩƷ����.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txtҩƷ����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
            txt���.Text = Nvl(rsTemp!���): txt����.Text = Nvl(rsTemp!����)
            txt��λ.Text = Nvl(rsTemp!�ۼ۵�λ): txt�ۼ�.Text = Format(Nvl(rsTemp!�ۼ�), "0.00000")
            txt����.Text = 1: txtҩƷ����.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt����.Text = 0: txtҩƷ����.Tag = ""
            txtҩƷ����.SetFocus
        End If
    Else
        txt����.Text = 0: txtҩƷ����.Tag = ""
        txtҩƷ����.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
