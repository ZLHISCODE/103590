VERSION 5.00
Begin VB.Form frm���ֱ�׼�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frm���ֱ�׼��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   972
      Top             =   1170
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   162
      TabIndex        =   5
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2217
      TabIndex        =   2
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3447
      TabIndex        =   3
      Top             =   1695
      Width           =   1100
   End
   Begin VB.Frame fra1 
      Caption         =   "ϵͳ����"
      Height          =   1380
      Left            =   147
      TabIndex        =   4
      Top             =   135
      Width           =   4380
      Begin VB.CheckBox chk���� 
         Caption         =   "Ҫ�󲡰���ҳ��Ŀ��������(&P)"
         Height          =   210
         Index           =   91
         Left            =   975
         TabIndex        =   1
         Top             =   810
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���ֵȼ��Զ�д�벡����ҳ(&Y)"
         Height          =   210
         Index           =   90
         Left            =   975
         TabIndex        =   0
         Top             =   435
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   315
         Picture         =   "frm���ֱ�׼��������.frx":000C
         Top             =   450
         Width           =   480
      End
   End
End
Attribute VB_Name = "frm���ֱ�׼��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'==============================================================================
'=���ܣ� ���������������'��ʾTips
'==============================================================================
Private Sub chk����_GotFocus(Index As Integer)
    On Error GoTo errH
    If Index = 90 Then
        ShowTips fra1, chk����(90), "�Ƿ��ڲ�����ҳ�ĵȼ�Ϊ��ʱ�������ֽ���ȼ��Զ�д�벡����ҳ��Ĭ��Ϊ��", "���ֵȼ��Զ�д�벡����ҳ"
    ElseIf Index = 91 Then
        ShowTips fra1, chk����(91), "�Ƿ�Ҫ�󲡰���ҳ��Ŀ��������֡�Ĭ��Ϊ�ǡ�", "Ҫ�󲡰���ҳ��Ŀ��������"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ȡ���˳�
'==============================================================================
Private Sub cmdCancel_Click()
    On Error GoTo errH
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ȷ���������
'==============================================================================
Private Sub cmdOK_Click()
    On Error GoTo errH
    If Save����() = False Then Exit Sub
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڿؼ���ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڳ�ʼ��
'==============================================================================
Private Sub Form_Load()
    chk����(90).Value = zlDatabase.GetPara(90, glngSys)
    chk����(91).Value = zlDatabase.GetPara(91, glngSys)
End Sub

'==============================================================================
'=����:����༭�����ݵ�������ϵͳ������صı���
'=����ֵ:�ɹ�����True,����ΪFalse
'==============================================================================
Private Function Save����() As Boolean
    Dim i           As Integer
    
    On Error GoTo errH
    
    Save���� = False
    gcnOracle.BeginTrans
    For i = 90 To 91
        Call zlDatabase.SetPara(i, IIf(chk����(i).Value = 1, 1, 0), ParamInfo.ϵͳ��, 0)
    Next
    gcnOracle.CommitTrans
    Save���� = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=����:�ؼ�ͨ��Tips��ʾ
'==============================================================================
Private Sub ShowTips(ctl0 As Control, ctl As Control, str���� As String, Optional str���� As String = "��ʾ��Ϣ", Optional lngʱ�� As Long = 3000)
    Dim X           As Single
    Dim Y           As Single
    
    On Error GoTo errH
    
    X = (ctl.Left + ctl.Width / 2 + ctl0.Left) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height + ctl0.Top) / Screen.TwipsPerPixelY
    If Len(str����) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lngʱ��
        tipPopup1.Title = str����
        tipPopup1.Text = str����
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
