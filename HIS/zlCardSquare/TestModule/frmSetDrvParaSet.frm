VERSION 5.00
Begin VB.Form frmSetDrvParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�豸����"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4125
      TabIndex        =   8
      Top             =   915
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4125
      TabIndex        =   7
      Top             =   435
      Width           =   1100
   End
   Begin VB.Frame fraSet 
      Caption         =   "�豸����"
      Height          =   1695
      Left            =   135
      TabIndex        =   0
      Top             =   255
      Width           =   3855
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "�Զ�ʶ��"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1170
         Width           =   1095
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "300"
         ToolTipText     =   "��С300����"
         Top             =   1125
         Width           =   495
      End
      Begin VB.ComboBox cboCom 
         Height          =   300
         ItemData        =   "frmSetDrvParaSet.frx":0000
         Left            =   1440
         List            =   "frmSetDrvParaSet.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label lblSet 
         Caption         =   "ͨѶ�˿�"
         Height          =   225
         Left            =   600
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
      Begin VB.Label lbltitle 
         Caption         =   "�Զ�ʶ����"
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "����"
         Height          =   225
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSetDrvParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngCardNo As Long
Private Sub chkAutoRead_Click()
    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        txtInterval.Text = Val(GetSetting("ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ���", 300))
    Else
        txtInterval.Enabled = False
        txtInterval.Text = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdOK_Click()
    Dim i As Integer
    SaveSetting "ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�˿�", cboCom.ListIndex
    SaveSetting "ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ���", Val(txtInterval.Text)
    SaveSetting "ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ", Val(chkAutoRead.Value)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call gobjCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTmp As Integer
    Dim bln�Զ���ȡ As Boolean
    cboCom.Clear
    With cboCom
        .AddItem "Com1"
        .AddItem "Com2"
        .AddItem "Com3"
        .AddItem "Com4"
        .AddItem "Com5"
        .AddItem "Com6"
        .AddItem "Com7"
        .AddItem "Com8"
    End With
    cboCom.ListIndex = 0
 
    i = Val(GetSetting("ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�˿�", 0))
    If i > 0 And i <= cboCom.ListCount Then cboCom.ListIndex = i

    If bln�Զ���ȡ = True Then
        chkAutoRead.Enabled = False
        txtInterval.Enabled = False
    Else
        chkAutoRead.Value = Val(GetSetting("ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ", 1))
    End If

    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        intTmp = Val(GetSetting("ZLSOFT", "����ȫ��\SquareCard\" & mlngCardNo, "�Զ���ȡ���", 300))
    Else
        txtInterval.Enabled = False
        intTmp = 0
    End If
    txtInterval.Text = IIf(intTmp < 300, 300, intTmp)
End Sub
Public Sub ShowMe(ByVal frmMain As Form, ByVal lngCardNo As Long)
    mlngCardNo = lngCardNo
    Me.Show 1, frmMain
End Sub






