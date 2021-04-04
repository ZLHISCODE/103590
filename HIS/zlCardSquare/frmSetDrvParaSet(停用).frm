VERSION 5.00
Begin VB.Form frmSetDrvParaSet 
   Caption         =   "设备参数设置"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5490
   Icon            =   "frmSetDrvParaSet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5490
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSet 
      Caption         =   "设备配置"
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
         ToolTipText     =   "最小300毫秒"
         Top             =   1125
         Width           =   495
      End
      Begin VB.CheckBox chkAutoRead 
         Caption         =   "自动识别"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Caption         =   "毫秒"
         Height          =   225
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbltitle 
         Caption         =   "自动识别间隔"
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblSet 
         Caption         =   "通讯端口"
         Height          =   225
         Left            =   600
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   1
      Top             =   345
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
Private mbytCardType As Byte    '0-消费卡;1-医疗卡
Private mstr医疗卡 As String '医疗卡,消费卡时为空
Private Sub chkAutoRead_Click()
    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        txtInterval.Text = Val(GetSetting("ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "自动读取间隔", 300))
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
    
    SaveSetting "ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "端口", cboCom.ListIndex
    SaveSetting "ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "自动读取间隔", Val(txtInterval.Text)
    SaveSetting "ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "自动读取", Val(chkAutoRead.Value)
    If mbytCardType = 1 Then
        For i = 1 To objYLCards.Count
            If objYLCards.Item(i).接口序号 = Val(mstrCardTypeNo) Then
                objYLCards.Item(i).是否自动读取 = Val(chkAutoRead.Value)
            End If
        Next
        For i = 1 To objYlCardObjs.Count
            If objYlCardObjs.Item(i).接口序号 = Val(mstrCardTypeNo) Then
                objYlCardObjs.Item(i).CardPreporty.是否自动读取 = Val(chkAutoRead.Value)
            End If
        Next
    Else
        For i = 1 To gObjXFCards.Count
            If gObjXFCards.Item(i).接口编码 = mstrCardTypeNo Then
                gObjXFCards.Item(i).是否自动读取 = Val(chkAutoRead.Value)
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
 
    i = Val(GetSetting("ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "端口", 0))
    If i > 0 And i <= cboCom.ListCount Then cboCom.ListIndex = i
    If bln自动读取 = True Then
        chkAutoRead.Enabled = False
        txtInterval.Enabled = False
    Else
        chkAutoRead.Value = Val(GetSetting("ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "自动读取", 1))
    End If

    If chkAutoRead.Value = 1 Then
        txtInterval.Enabled = True
        intTmp = Val(GetSetting("ZLSOFT", mstr医疗卡 & mstrCardTypeNo, "自动读取间隔", 300))
    Else
        txtInterval.Enabled = False
        intTmp = 0
    End If
    txtInterval.Text = IIf(intTmp < 300, 300, intTmp)
End Sub
Public Sub ShowMe(ByVal frmMain As Form, ByVal strCardTypeNo As String, Optional bytCardType As Byte = 1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示消费卡或医疗卡的设备配置窗体
    '入参:frmMain-调用的主窗体
    '       strCardTypeNo-卡类别号(消费卡为接口序号;医疗卡类医疗卡类别ID)
    '       bytCardType-1表示消费卡;2表示医疗卡
    '编制:刘兴洪
    '日期:2011-05-25 11:57:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardTypeNo = strCardTypeNo: mbytCardType = bytCardType
    mstr医疗卡 = "公共模块\zlSquareCard\" & IIf(mbytCardType = 2, "医疗卡\", "")
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
End Sub
