VERSION 5.00
Begin VB.Form frmParameterSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IC卡设备设置"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5415
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame fraSet 
      Caption         =   "设备配置"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "自动识别"
         Height          =   225
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtInterval 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "300"
         ToolTipText     =   "最小300毫秒"
         Top             =   1282
         Width           =   495
      End
      Begin VB.ComboBox cboCom 
         Height          =   300
         ItemData        =   "frmParameterSet.frx":0000
         Left            =   1440
         List            =   "frmParameterSet.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txt_MW_SAddr 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "0"
         Top             =   825
         Width           =   495
      End
      Begin VB.TextBox txt_MW_Len 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "8"
         Top             =   825
         Width           =   495
      End
      Begin VB.Label lblSet 
         Caption         =   "通讯端口"
         Height          =   225
         Left            =   600
         TabIndex        =   10
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "数据起始地址"
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   863
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "长度"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   855
         Width           =   375
      End
      Begin VB.Label lbltitle 
         Caption         =   "自动识别间隔"
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "毫秒"
         Height          =   225
         Index           =   2
         Left            =   3240
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmParameterSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mIntCardNo As Integer

Private Sub chkAutoRead_Click()
    If chkAutoRead.value = 1 Then
        txtInterval.Enabled = True
        txtInterval.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard" & mIntCardNo, "自动读取间隔", 300))
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
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "端口", cboCom.ListIndex
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "起始地址", Val(txt_MW_SAddr.Text)
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "长度", Val(txt_MW_Len.Text)
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "自动读取间隔", Val(txtInterval.Text)
    SaveSetting "ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "自动读取", Val(chkAutoRead.value)
    For i = 1 To Cards.Count
        If Item(i).编码 = mIntCardNo Then
            Item(i).是否自动读取 = Val(chkAutoRead.value)
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
    Dim bln自动读取 As Boolean
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
    
    mIntCardNo = Val(frmCardSelect.vfgList.TextMatrix(frmCardSelect.vfgList.Row, frmCardSelect.vfgList.ColIndex("编码")))
    
    i = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "端口", 0))
    If i > 0 And i <= cboCom.ListCount Then cboCom.ListIndex = i

    If mIntCardNo = 4 Or mIntCardNo = 10 Or mIntCardNo = 11 Or mIntCardNo = 12 Then
        fraSet.Enabled = True
        txt_MW_SAddr.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "起始地址", 32))
        txt_MW_Len.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "长度", 10))
    ElseIf mIntCardNo = 13 Then
        fraSet.Enabled = True
        txt_MW_SAddr.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "起始地址", 2))
        txt_MW_Len.Text = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "长度", 16))
    Else
        txt_MW_SAddr.Enabled = False
        txt_MW_Len.Enabled = False
    End If

    For i = 1 To Cards.Count
        If Item(i).是否自动读取 = 1 And Item(i).编码 <> mIntCardNo Then bln自动读取 = True
    Next
    If bln自动读取 = True Then
        chkAutoRead.Enabled = False
        txtInterval.Enabled = False
    Else
        chkAutoRead.value = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "自动读取", 1))
    End If
    
    If chkAutoRead.value = 1 Then
        txtInterval.Enabled = True
        intTmp = Val(GetSetting("ZLSOFT", "公共全局\ICCard\" & mIntCardNo, "自动读取间隔", 300))
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




