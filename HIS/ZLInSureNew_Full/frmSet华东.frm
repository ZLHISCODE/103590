VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk��λ�� 
      Caption         =   "��λ��(&B)"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   990
      Width           =   1155
   End
   Begin VB.Frame fra��λ�� 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   180
      TabIndex        =   4
      Top             =   990
      Width           =   4605
      Begin VB.TextBox txt�Է��� 
         Height          =   300
         Left            =   3060
         TabIndex        =   10
         Top             =   870
         Width           =   1365
      End
      Begin VB.TextBox txt��λ���޶� 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   7
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "    ����λ�ѳ����޶�󣬶��ಿ�ֶ�ӦΪ�Է��룬�����µļ�¼�ϴ�"
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   330
         Width           =   4125
      End
      Begin VB.Label lbl��λ���Է��� 
         AutoSize        =   -1  'True
         Caption         =   "�Է���(&Z)"
         Height          =   180
         Left            =   2130
         TabIndex        =   9
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1740
         TabIndex        =   8
         Top             =   930
         Width           =   180
      End
      Begin VB.Label lbl��λ���޶� 
         AutoSize        =   -1  'True
         Caption         =   "�޶�(&X)"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   930
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   12
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2415
      TabIndex        =   11
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdBrower 
      Caption         =   "���(&B)"
      Height          =   350
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   945
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   3660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ָ���ļ����λ��(&L)"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1890
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���� As Long, mblnReturn As Boolean

Public Function ShowME(ByVal lng���� As Long) As Boolean
    mlng���� = lng����
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub chk��λ��_Click()
    On Error Resume Next
    fra��λ��.Enabled = (chk��λ��.Value = 1)
    If fra��λ��.Enabled Then
        txt��λ���޶�.SetFocus
    Else
        txt��λ���޶�.Text = ""
        txt�Է���.Text = ""
    End If
End Sub

Private Sub cmdBrower_Click()
    txtPath.Text = BrowPath(Me.hwnd, "��ѡ���ļ����λ�ã�")
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtPath.Text) = "" Then Exit Sub
    If chk��λ��.Value = 1 Then
        If Val(txt��λ���޶�.Text) <= 0 Then
            MsgBox "��λ���޶��С�ڵ����㣡", vbInformation, gstrSysName
            Exit Sub
        End If
        If Trim(txt�Է���.Text) = "" Then
            MsgBox "�Է��벻��Ϊ�գ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'�ļ����λ��','" & txtPath.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'��λ���޶�','" & txt��λ���޶�.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'��λ���Է���','" & txt�Է���.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mstrSavePath = txtPath.Text
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        If rsTemp!������ = "�ļ����λ��" Then txtPath.Text = strTemp
        If rsTemp!������ = "��λ���޶�" Then txt��λ���޶�.Text = strTemp
        If rsTemp!������ = "��λ���Է���" Then txt�Է���.Text = strTemp
        rsTemp.MoveNext
    Loop
    If Val(txt��λ���޶�.Text) <> 0 Then chk��λ��.Value = 1
End Sub
