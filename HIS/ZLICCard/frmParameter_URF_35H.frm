VERSION 5.00
Begin VB.Form frmParameter_URF_35H 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IC���豸����"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5385
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraSet 
      Caption         =   "�豸����"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cbo�� 
         Height          =   300
         ItemData        =   "frmParameter_URF_35H.frx":0000
         Left            =   2040
         List            =   "frmParameter_URF_35H.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txt_MW_Len 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "7"
         Top             =   825
         Width           =   495
      End
      Begin VB.TextBox txt_MW_SAddr 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "1"
         Top             =   825
         Width           =   495
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmParameter_URF_35H.frx":0004
         Left            =   480
         List            =   "frmParameter_URF_35H.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "300"
         ToolTipText     =   "��С300����"
         Top             =   1282
         Width           =   495
      End
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "�Զ�ʶ��"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "��"
         Height          =   225
         Left            =   2880
         TabIndex        =   16
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "��"
         Height          =   225
         Left            =   1800
         TabIndex        =   15
         Top             =   405
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   225
         Left            =   1320
         TabIndex        =   13
         Top             =   405
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "����"
         Height          =   225
         Index           =   2
         Left            =   3240
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "�Զ�ʶ����"
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "����"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   855
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "������ʼ��ַ"
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   863
         Width           =   1095
      End
      Begin VB.Label lblSet 
         Caption         =   "��"
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   1100
   End
End
Attribute VB_Name = "frmParameter_URF_35H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mIntCardNo As Integer

Private Sub chkAutoRead_Click()
    If chkAutoRead.value = 1 Then
        txtInterval.Enabled = True
        txtInterval.Text = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard" & mIntCardNo, "�Զ���ȡ���", 300))
    Else
        txtInterval.Enabled = False
        txtInterval.Text = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "����", cbo����.ListIndex
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "��", cbo��.ListIndex
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "��ʼ��ַ", Val(txt_MW_SAddr.Text)
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "����", Val(txt_MW_Len.Text)
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "�Զ���ȡ���", Val(txtInterval.Text)
    SaveSetting "ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "�Զ���ȡ", Val(chkAutoRead.value)
    For i = 1 To Cards.Count
        If Item(i).���� = mIntCardNo Then
            Item(i).�Ƿ��Զ���ȡ = Val(chkAutoRead.value)
        End If
    Next
    Call frmCardSelect.LoadData(Cards, False)
    frmTimer.tmrMain.Interval = Val(txtInterval.Text)
    frmTimer.tmrMain.Enabled = False
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim intTmp As Integer
    Dim bln�Զ���ȡ As Boolean
    cbo����.Clear
    With cbo����
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "6"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .AddItem "13"
        .AddItem "14"
        .AddItem "15"
    End With
    cbo����.ListIndex = 0
    
    cbo��.Clear
    With cbo��
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
    End With
    cbo��.ListIndex = 0
    
    mIntCardNo = Val(frmCardSelect.vfgList.TextMatrix(frmCardSelect.vfgList.Row, frmCardSelect.vfgList.ColIndex("����")))
    
    i = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "����", 0))
    If i > 0 And i <= cbo����.ListCount Then cbo����.ListIndex = i
    
    i = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "��", 0))
    If i > 0 And i <= cbo��.ListCount Then cbo��.ListIndex = i

    If mIntCardNo = 9 Then
        fraSet.Enabled = True
        txt_MW_SAddr.Text = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "��ʼ��ַ", 1))
        txt_MW_Len.Text = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "����", 7))
    Else
        txt_MW_SAddr.Enabled = False
        txt_MW_Len.Enabled = False
    End If

    For i = 1 To Cards.Count
        If Item(i).�Ƿ��Զ���ȡ = 1 And Item(i).���� <> mIntCardNo Then bln�Զ���ȡ = True
    Next
    If bln�Զ���ȡ = True Then
        chkAutoRead.Enabled = False
        txtInterval.Enabled = False
    Else
        chkAutoRead.value = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "�Զ���ȡ", 1))
    End If
    
    If chkAutoRead.value = 1 Then
        txtInterval.Enabled = True
        intTmp = Val(GetSetting("ZLSOFT", "����ȫ��\ICCard\" & mIntCardNo, "�Զ���ȡ���", 300))
    Else
        txtInterval.Enabled = False
        intTmp = 0
    End If
    txtInterval.Text = IIf(intTmp < 300, 300, intTmp)
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtInterval_Validate(Cancel As Boolean)
    If txtInterval.Text < 300 Then Cancel = True
End Sub

Private Sub txt_MW_Len_GotFocus()
    txt_MW_Len.SelStart = 0
    txt_MW_Len.SelLength = Len(txt_MW_Len)
End Sub

Private Sub txt_MW_Len_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt_MW_SAddr_GotFocus()
    txt_MW_SAddr.SelStart = 0
    txt_MW_SAddr.SelLength = Len(txt_MW_SAddr)
End Sub

Private Sub txt_MW_SAddr_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Public Sub ShowMe(ByVal intCardType As Integer)
    mIntCardNo = intCardType
    Me.Show vbModal
End Sub
