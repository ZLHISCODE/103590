VERSION 5.00
Begin VB.Form frmDeviceSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����������"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chk�������� 
      Caption         =   "�����������"
      Height          =   300
      Left            =   405
      TabIndex        =   14
      Top             =   285
      Width           =   1500
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "����(&T)"
      Height          =   390
      Left            =   3630
      TabIndex        =   13
      Top             =   1410
      Width           =   1065
   End
   Begin VB.Frame fraKeyboard 
      Caption         =   "�˿�����"
      Height          =   2595
      Left            =   285
      TabIndex        =   2
      Top             =   345
      Width           =   3165
      Begin VB.ComboBox cboDataBit 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1965
         Width           =   1650
      End
      Begin VB.ComboBox CboStopBit 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1560
         Width           =   1650
      End
      Begin VB.ComboBox cboCheckBit 
         Height          =   300
         ItemData        =   "frmDeviceSet.frx":0000
         Left            =   1065
         List            =   "frmDeviceSet.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1185
         Width           =   1650
      End
      Begin VB.ComboBox cboPt 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   795
         Width           =   1650
      End
      Begin VB.ComboBox cboCom 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   1650
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����λ"
         Height          =   180
         Index           =   4
         Left            =   435
         TabIndex        =   12
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ֹͣλ"
         Height          =   180
         Index           =   3
         Left            =   435
         TabIndex        =   10
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��żУ��λ"
         Height          =   180
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   6
         Top             =   855
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ͨѶ�˿�"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   4
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3630
      TabIndex        =   1
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   0
      Top             =   930
      Width           =   1100
   End
End
Attribute VB_Name = "frmDeviceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-07-28 10:42:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strReg As String, i As Long, j As Long
   i = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "�˿�", "0"))
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
    If i >= 0 And i <= cboCom.ListCount - 1 Then cboCom.ListIndex = i
    
    i = Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "������", "9600"))
    With cboPt
        .Clear
        .AddItem "75": If i = 75 Then .ListIndex = .NewIndex
        .AddItem "110": If i = 110 Then .ListIndex = .NewIndex
        .AddItem "134": If i = 134 Then .ListIndex = .NewIndex
        .AddItem "150": If i = 150 Then .ListIndex = .NewIndex
        .AddItem "300": If i = 300 Then .ListIndex = .NewIndex
        .AddItem "600": If i = 600 Then .ListIndex = .NewIndex
        .AddItem "1200": If i = 1200 Then .ListIndex = .NewIndex
        .AddItem "2400": If i = 2400 Then .ListIndex = .NewIndex
        .AddItem "4800": If i = 4800 Then .ListIndex = .NewIndex
        .AddItem "9600": If i = 9600 Then .ListIndex = .NewIndex: j = .NewIndex
        .AddItem "14400": If i = 14400 Then .ListIndex = .NewIndex
        .AddItem "19200": If i = 19200 Then .ListIndex = .NewIndex
        .AddItem "38400": If i = 38400 Then .ListIndex = .NewIndex
        .AddItem "43000": If i = 43000 Then .ListIndex = .NewIndex
        .AddItem "56000": If i = 56000 Then .ListIndex = .NewIndex
        .AddItem "57600": If i = 57600 Then .ListIndex = .NewIndex
        .AddItem "115200": If i = 115200 Then .ListIndex = .NewIndex
        .AddItem "128000": If i = 128000 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = j
    End With
    strReg = Trim(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "��ż����λ", ""))
    With cboCheckBit
        .Clear
        .AddItem "��": If strReg = "��" Then .ListIndex = .NewIndex
        .AddItem "��": If strReg = "��" Then .ListIndex = .NewIndex
        .AddItem "ż": If strReg = "ż" Then .ListIndex = .NewIndex
        .AddItem "��־": If strReg = "��־" Then .ListIndex = .NewIndex
        .AddItem "�ո�": If strReg = "�ո�" Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    strReg = Trim(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "��ż����λ", "��"))
    With cboCheckBit
        .Clear
        .AddItem "��": If strReg = "��" Then .ListIndex = .NewIndex
        .AddItem "��": If strReg = "��" Then .ListIndex = .NewIndex
        .AddItem "ż": If strReg = "ż" Then .ListIndex = .NewIndex
        .AddItem "��־": If strReg = "��־" Then .ListIndex = .NewIndex
        .AddItem "�ո�": If strReg = "�ո�" Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    strReg = Trim(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "ֹͣλ", "1"))
    With CboStopBit
        .Clear
        .AddItem 1: If Val(strReg) = 1 Then .ListIndex = .NewIndex
        .AddItem 1.5: If Val(strReg) = 1.5 Then .ListIndex = .NewIndex
        .AddItem 2: If Val(strReg) = 2 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    strReg = Trim(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "����λ", "1"))
    
    With cboDataBit
        .Clear
       .AddItem 4: If Val(strReg) = 4 Then .ListIndex = .NewIndex
       .AddItem 5: If Val(strReg) = 5 Then .ListIndex = .NewIndex
       .AddItem 6: If Val(strReg) = 6 Then .ListIndex = .NewIndex
       .AddItem 7: If Val(strReg) = 7 Then .ListIndex = .NewIndex
       .AddItem 8: If Val(strReg) = 8 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    chk��������.Value = IIf(Val(GetSetting("ZLSOFT", "����ģ��\zlKeyboard", "����", "0")) = 1, 1, 0)
    Call chk��������_Click
End Sub
Private Sub SavePata()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2011-07-28 10:43:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
     SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "����", chk��������.Value
    gblnStartKeyboard = IIf(chk��������.Value = 1, True, False)
    If chk��������.Value = 1 Then
        SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "�˿�", cboCom.ListIndex
        SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "������", cboPt.Text
        SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "��ż����λ", cboCheckBit.Text
        SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "ֹͣλ", CboStopBit.Text
        SaveSetting "ZLSOFT", "����ģ��\zlKeyboard", "����λ", cboDataBit.Text
    End If
End Sub

Private Sub chk��������_Click()
    fraKeyboard.Enabled = chk��������.Value = 1
    cboCheckBit.Enabled = fraKeyboard.Enabled
    cboCom.Enabled = fraKeyboard.Enabled
    cboDataBit.Enabled = fraKeyboard.Enabled
    cboPt.Enabled = fraKeyboard.Enabled
    CboStopBit.Enabled = fraKeyboard.Enabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call SavePata
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim cllKeyboard As clsKeyboard
    Set cllKeyboard = New clsKeyboard
    Call SavePata
    Call cllKeyboard.OpenPassKeyoardInput(Me, Nothing, False)
    Call cllKeyboard.ColsePassKeyoardInput(Me, Nothing)
    Set cllKeyboard = Nothing
End Sub

Private Sub Form_Load()
    Call InitData
End Sub


