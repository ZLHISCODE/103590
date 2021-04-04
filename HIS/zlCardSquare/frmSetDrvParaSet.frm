VERSION 5.00
Begin VB.Form frmSetDrvParaSet 
   Caption         =   "�豸��������"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5490
   Icon            =   "frmSetDrvParaSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5490
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSet 
      Caption         =   "�豸����"
      Height          =   1695
      Left            =   180
      TabIndex        =   2
      Top             =   165
      Width           =   3855
      Begin VB.ComboBox cboCom 
         Height          =   300
         ItemData        =   "frmSetDrvParaSet.frx":030A
         Left            =   1440
         List            =   "frmSetDrvParaSet.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   1230
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "300"
         ToolTipText     =   "��С300����"
         Top             =   1125
         Width           =   495
      End
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "�Զ�ʶ��"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "����"
         Height          =   225
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "�Զ�ʶ����"
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblSet 
         Caption         =   "ͨѶ�˿�"
         Height          =   225
         Left            =   600
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   1
      Top             =   345
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4170
      TabIndex        =   0
      Top             =   825
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetDrvParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCardTypeNo As String
Private mbytCardType As Byte    '0-���ѿ�;1-ҽ�ƿ�
Private mstrҽ�ƿ� As String 'ҽ�ƿ�,���ѿ�ʱΪ��
Private Sub chkAutoRead_Click()
    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        txtInterval.Text = Val(GetSetting("ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�Զ���ȡ���", 300))
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
    Dim objYLCards As clsCards
    Dim objYlCardObjs As clsCardObjects
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Sub
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Sub
    
    SaveSetting "ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�˿�", cboCom.ListIndex
    SaveSetting "ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�Զ���ȡ���", Val(txtInterval.Text)
    SaveSetting "ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�Զ���ȡ", Val(chkAutoRead.Value)
    If mbytCardType = 1 Then
        For i = 1 To objYLCards.Count
            If objYLCards.Item(i).�ӿ���� = Val(mstrCardTypeNo) Then
                objYLCards.Item(i).�Ƿ��Զ���ȡ = Val(chkAutoRead.Value)
            End If
        Next
        For i = 1 To objYlCardObjs.Count
            If objYlCardObjs.Item(i).�ӿ���� = Val(mstrCardTypeNo) Then
                objYlCardObjs.Item(i).CardPreporty.�Ƿ��Զ���ȡ = Val(chkAutoRead.Value)
            End If
        Next
    Else
        For i = 1 To gObjXFCards.Count
            If gObjXFCards.Item(i).�ӿڱ��� = mstrCardTypeNo Then
                gObjXFCards.Item(i).�Ƿ��Զ���ȡ = Val(chkAutoRead.Value)
            End If
        Next
    End If
    Call frmCardSelect.LoadData
    frmCardBrush.tmrMain.Interval = Val(txtInterval.Text)
    frmCardBrush.tmrMain.Enabled = False
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
 
    i = Val(GetSetting("ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�˿�", 0))
    If i > 0 And i <= cboCom.ListCount Then cboCom.ListIndex = i
    If bln�Զ���ȡ = True Then
        chkAutoRead.Enabled = False
        txtInterval.Enabled = False
    Else
        chkAutoRead.Value = Val(GetSetting("ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�Զ���ȡ", 1))
    End If

    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        intTmp = Val(GetSetting("ZLSOFT", mstrҽ�ƿ� & mstrCardTypeNo, "�Զ���ȡ���", 300))
    Else
        txtInterval.Enabled = False
        intTmp = 0
    End If
    txtInterval.Text = IIf(intTmp < 300, 300, intTmp)
End Sub
Public Sub ShowMe(ByVal frmMain As Form, ByVal strCardTypeNo As String, Optional bytCardType As Byte = 1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���ѿ���ҽ�ƿ����豸���ô���
    '���:frmMain-���õ�������
    '       strCardTypeNo-������(���ѿ�Ϊ�ӿ����;ҽ�ƿ���ҽ�ƿ����ID)
    '       bytCardType-1��ʾ���ѿ�;2��ʾҽ�ƿ�
    '����:���˺�
    '����:2011-05-25 11:57:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardTypeNo = strCardTypeNo: mbytCardType = bytCardType
    mstrҽ�ƿ� = "����ģ��\zlSquareCard\" & IIf(mbytCardType = 2, "ҽ�ƿ�\", "")
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
End Sub
