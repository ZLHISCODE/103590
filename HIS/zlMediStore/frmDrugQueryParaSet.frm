VERSION 5.00
Begin VB.Form frmDrugQueryParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   3270
   ClientLeft      =   3585
   ClientTop       =   4680
   ClientWidth     =   4770
   Icon            =   "frmDrugQueryParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox Chk����ͣ��ҩƷ 
      Caption         =   "����ͣ��ҩƷ(&S)"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2910
      Width           =   2730
   End
   Begin VB.TextBox TxtЧ�ڱ��� 
      Height          =   300
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "3"
      Top             =   2160
      Width           =   300
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   3420
      TabIndex        =   12
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CheckBox chk����� 
      Caption         =   "ֻ��ʾ�п��������ҩƷ(&L)"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2580
      Width           =   2730
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   10
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3420
      TabIndex        =   11
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame Fra�������� 
      Caption         =   "��ʾ��λ����"
      Height          =   1905
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3000
      Begin VB.OptionButton Opt��λ2 
         Caption         =   "���ﵥλ(&2)"
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1305
      End
      Begin VB.OptionButton Opt��λ4 
         Caption         =   "סԺ��λ(&4)"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1305
      End
      Begin VB.OptionButton Opt��λ3 
         Caption         =   "ҩ�ⵥλ(&3)"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
      End
      Begin VB.OptionButton Opt��λ1 
         Caption         =   "�ۼ۵�λ(&1)"
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Label Lbl�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1740
      TabIndex        =   7
      Top             =   2220
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
      Left            =   240
      TabIndex        =   5
      Top             =   2220
      Width           =   990
   End
End
Attribute VB_Name = "frmDrugQueryParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrBillStyle As String
Private IntChoose As Integer
Private BlnBootUp As Boolean '�����ɹ���
Private mstrPrivs As String
Private mblnSetPara As Boolean      '�Ƿ���в�������Ȩ��
'ע��:ѡ������һ����λ,����ⵥ���Դ˵�λ��ʾ

Public Property Get In_Ȩ��() As String
    In_Ȩ�� = mstrPrivs
End Property

Public Property Let In_Ȩ��(ByVal vNewValue As String)
    mstrPrivs = vNewValue
End Property
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd����_Click()
    If Val(TxtЧ�ڱ���.Text) < 0 Then
        MsgBox "Ч�ڱ�������С���㣡", vbInformation, gstrSysName
        TxtЧ�ڱ���.SetFocus
        Exit Sub
    End If

    If Opt��λ1.Value = True Then IntChoose = 1
    If Opt��λ2.Value = True Then IntChoose = 2
    If Opt��λ3.Value = True Then IntChoose = 3
    If Opt��λ4.Value = True Then IntChoose = 4

    zlDataBase.SetPara "��λ", IntChoose, glngSys, 1309
    zlDataBase.SetPara "�Ƿ���ʾ�޿��ҩƷ", chk�����.Value, glngSys, 1309
    zlDataBase.SetPara "Ч�ڱ�������", Val(TxtЧ�ڱ���.Text), glngSys, 1309
    zlDataBase.SetPara "�Ƿ���ʾͣ��ҩƷ", Chk����ͣ��ҩƷ.Value, glngSys, 1309

    frmDrugQuery.BlnDO = True
    Unload Me
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Cmdȡ��_Click
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    
    Dim bln��� As Boolean
    Dim intMonths As Integer
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "��������")

    IntChoose = Val(zlDataBase.GetPara("��λ", glngSys, 1309, 3, Array(Fra��������), mblnSetPara))
    bln��� = (zlDataBase.GetPara("�Ƿ���ʾ�޿��ҩƷ", glngSys, 1309, 0, Array(chk�����), mblnSetPara) = 1)
    intMonths = Val(zlDataBase.GetPara("Ч�ڱ�������", glngSys, 1309, 3, Array(TxtЧ�ڱ���), mblnSetPara))
    Chk����ͣ��ҩƷ.Value = Val(zlDataBase.GetPara("�Ƿ���ʾͣ��ҩƷ", glngSys, 1309, 0, Array(Chk����ͣ��ҩƷ), mblnSetPara))
    
    Select Case IntChoose
        Case 1
            Opt��λ1.Value = True
        Case 2
            Opt��λ2.Value = True
        Case 3
            Opt��λ3.Value = True
        Case 4
            Opt��λ4.Value = True
    End Select
    Me.chk�����.Value = IIf(bln���, 1, 0)
    Me.TxtЧ�ڱ��� = intMonths
    
    If glngSys \ 100 = 8 Then
        Opt��λ2.Visible = False
        Opt��λ4.Visible = False
        Opt��λ3.Caption = "�ɹ���λ(&3)"
        If Opt��λ3.Value = 0 And Opt��λ1.Value = 0 Then
            Opt��λ1.Value = 1
        End If
    End If
    
    BlnBootUp = True
End Sub

Private Sub TxtЧ�ڱ���_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub
