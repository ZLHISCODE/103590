VERSION 5.00
Begin VB.Form frmStuffQueryParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   3060
   ClientLeft      =   3585
   ClientTop       =   4680
   ClientWidth     =   4770
   Icon            =   "frmStuffQueryParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   -90
      TabIndex        =   12
      Top             =   2175
      Width           =   5070
   End
   Begin VB.OptionButton Opt��λ1 
      Caption         =   "��ɢװ��λ��ʾ���(&1)"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   2370
   End
   Begin VB.OptionButton Opt��λ2 
      Caption         =   "�ð�װ��λ��ʾ���(&2)"
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   975
      Width           =   2205
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   -45
      TabIndex        =   11
      Top             =   660
      Width           =   5070
   End
   Begin VB.CheckBox Chk����ͣ�ò��� 
      Caption         =   "����ͣ�ò���(&S)"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1725
      Width           =   1950
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CheckBox chk����� 
      Caption         =   "ֻ��ʾ�п�������Ĳ���(&L)"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   1395
      Width           =   2730
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2460
      TabIndex        =   7
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   8
      Top             =   2505
      Width           =   1100
   End
   Begin VB.TextBox TxtЧ�ڱ��� 
      Height          =   300
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "3"
      Top             =   1665
      Width           =   300
   End
   Begin VB.Label lbl 
      Caption         =   "�Կ���ѯ������ʾ���á�"
      Height          =   240
      Left            =   765
      TabIndex        =   10
      Top             =   390
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   135
      Picture         =   "frmStuffQueryParaSet.frx":1CFA
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Lbl�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4095
      TabIndex        =   6
      Top             =   1725
      Width           =   180
   End
   Begin VB.Label LblЧ�ڱ��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ч�ڱ���(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2595
      TabIndex        =   4
      Top             =   1725
      Width           =   990
   End
End
Attribute VB_Name = "frmStuffQueryParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnBootUp As Boolean '�����ɹ���
Private mlngModule As Long
Private mstrPrivs As String
'ע��:ѡ������һ����λ,����ⵥ���Դ˵�λ��ʾ
Public Sub ��������(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '-----------------------------------------------------------------------------------------------
    '����:�����������
    '����:frmMain-������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '����:
    '����:���˺�
    '����:2007/12/24
    '-----------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Sub
Private Sub Chk����ͣ�ò���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/24
    '------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    Call zlDatabase.SetPara("���ĵ�λ", IIf(Me.Opt��λ2.Value = True, "1", "0"), glngSys, mlngModule)  '
    Call zlDatabase.SetPara("ֻ��ʾ�п������", IIf(chk�����.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("����ͣ������", IIf(Chk����ͣ�ò���.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("��������", Val(TxtЧ�ڱ���.Text), glngSys, mlngModule)
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub Cmd����_Click()
    If Val(TxtЧ�ڱ���.Text) < 0 Then
        MsgBox "Ч�ڱ�������С���㣡", vbInformation, gstrSysName
        TxtЧ�ڱ���.SetFocus
        Exit Sub
    End If
    If SaveSet = False Then Exit Sub
    frmStuffQuery.mblnDo = True
    Unload Me
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Cmdȡ��_Click
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim blnHavePriv As Boolean
    blnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "��������")
    RestoreWinState Me
    If Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, , Array(Opt��λ1, Opt��λ2), blnHavePriv)) = 0 Then
        Opt��λ1.Value = True
    Else
        Opt��λ2.Value = True
    End If
    Me.chk�����.Value = IIf(Val(zlDatabase.GetPara("ֻ��ʾ�п������", glngSys, mlngModule, , Array(chk�����), blnHavePriv)) = 1, 1, 0)
    Me.TxtЧ�ڱ���.Text = Val(zlDatabase.GetPara("��������", glngSys, mlngModule, 3, Array(TxtЧ�ڱ���), blnHavePriv))
    Chk����ͣ�ò���.Value = IIf(Val(zlDatabase.GetPara("����ͣ������", glngSys, mlngModule, , Array(TxtЧ�ڱ���, LblЧ�ڱ���, Lbl��), blnHavePriv)) = 1, 1, 0)
    mblnBootUp = True
End Sub

Private Sub Opt��λ1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub Opt��λ2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub TxtЧ�ڱ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtЧ�ڱ���_KeyPress(KeyAscii As Integer)
   zlControl.TxtCheckKeyPress TxtЧ�ڱ���, KeyAscii, m����ʽ
End Sub
