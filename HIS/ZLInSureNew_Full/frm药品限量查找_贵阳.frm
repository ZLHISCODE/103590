VERSION 5.00
Begin VB.Form frmҩƷ��������_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmҩƷ��������_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.TextBox txtҩƷ��Ϣ 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ��Ϣ"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "��������ҩƷ���롢ҩƷ���ơ�������в���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   705
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   4260
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmҩƷ��������_����"
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
    Dim rsTemp As New ADODB.Recordset, strWhere As String
    Me.MousePointer = vbHourglass
    strWhere = "  And (Upper(A.����) Like '%" & UCase(txtҩƷ��Ϣ.Text) & "%' Or  Upper(A.����) Like '%" & UCase(txtҩƷ��Ϣ.Text) & "%' " & _
              "     Or Upper(D.����) Like '%" & UCase(txtҩƷ��Ϣ.Text) & "%')"
    gstrSQL = "Select Distinct A.ҩƷID,A.����, A.����, A.���, A.����, A.�ۼ۵�λ, trim(to_char(B.����,'900090.00')) As ����, " & _
              "      trim(to_char(C.�ּ�,'900090.00000'))  As �ۼ�, trim(to_char(Nvl(B.����, 0) * Nvl(C.�ּ�, 0),'90009990.00')) As �ۼ۽��,B.��ע " & _
              "From ҩƷĿ¼ A, ��ҩ����Ŀ¼_���� B, �շѼ�Ŀ C,�շѱ��� D " & _
              "Where A.ҩƷid = B.ҩƷid And B.ҩƷid = C.�շ�ϸĿID And B.����=[1] And B.ҩƷID=D.�շ�ϸĿID " & _
              " And (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) " & strWhere & " Order By A.���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    Set frm�������Լ�����ҩ����_����.mshBill.DataSource = rsTemp
    Call CenterTableCaption(frm�������Լ�����ҩ����_����.mshBill)
    frm�������Լ�����ҩ����_����.mshBill.ColWidth(0) = 0
    Call frm�������Լ�����ҩ����_����.SetMenu
    Me.MousePointer = vbDefault
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub ShowME(ByVal intinsure As Integer)
    mint���� = intinsure
    Me.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
