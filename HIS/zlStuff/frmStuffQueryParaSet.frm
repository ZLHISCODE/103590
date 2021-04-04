VERSION 5.00
Begin VB.Form frmStuffQueryParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   -90
      TabIndex        =   12
      Top             =   2175
      Width           =   5070
   End
   Begin VB.OptionButton Opt单位1 
      Caption         =   "用散装单位显示库存(&1)"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   2370
   End
   Begin VB.OptionButton Opt单位2 
      Caption         =   "用包装单位显示库存(&2)"
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
   Begin VB.CheckBox Chk包含停用材料 
      Caption         =   "包含停用材料(&S)"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1725
      Width           =   1950
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   9
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CheckBox chk库存数 
      Caption         =   "只显示有库存数量的材料(&L)"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   1395
      Width           =   2730
   End
   Begin VB.CommandButton Cmd保存 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2460
      TabIndex        =   7
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton Cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   8
      Top             =   2505
      Width           =   1100
   End
   Begin VB.TextBox Txt效期报警 
      Height          =   300
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "3"
      Top             =   1665
      Width           =   300
   End
   Begin VB.Label lbl 
      Caption         =   "对库存查询进行显示设置。"
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
   Begin VB.Label Lbl月 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4095
      TabIndex        =   6
      Top             =   1725
      Width           =   180
   End
   Begin VB.Label Lbl效期报警 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "效期报警(&E)"
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
Private mblnBootUp As Boolean '启动成功否
Private mlngModule As Long
Private mstrPrivs As String
'注意:选择其中一个单位,则入库单据以此单位显示
Public Sub 参数设置(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '-----------------------------------------------------------------------------------------------
    '功能:参数设置入口
    '参数:frmMain-主窗口
    '     lngModule-模块号
    '     strPrivs-权限串
    '返回:
    '编制:刘兴宏
    '日期:2007/12/24
    '-----------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
End Sub
Private Sub Chk包含停用材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk库存数_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    Call zlDatabase.SetPara("卫材单位", IIf(Me.Opt单位2.Value = True, "1", "0"), glngSys, mlngModule)  '
    Call zlDatabase.SetPara("只显示有库存卫材", IIf(chk库存数.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("包含停用卫材", IIf(Chk包含停用材料.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("报警月数", Val(Txt效期报警.Text), glngSys, mlngModule)
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub Cmd保存_Click()
    If Val(Txt效期报警.Text) < 0 Then
        MsgBox "效期报警不能小于零！", vbInformation, gstrSysName
        Txt效期报警.SetFocus
        Exit Sub
    End If
    If SaveSet = False Then Exit Sub
    frmStuffQuery.mblnDo = True
    Unload Me
End Sub

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Cmd取消_Click
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim blnHavePriv As Boolean
    blnHavePriv = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    RestoreWinState Me
    If Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, , Array(Opt单位1, Opt单位2), blnHavePriv)) = 0 Then
        Opt单位1.Value = True
    Else
        Opt单位2.Value = True
    End If
    Me.chk库存数.Value = IIf(Val(zlDatabase.GetPara("只显示有库存卫材", glngSys, mlngModule, , Array(chk库存数), blnHavePriv)) = 1, 1, 0)
    Me.Txt效期报警.Text = Val(zlDatabase.GetPara("报警月数", glngSys, mlngModule, 3, Array(Txt效期报警), blnHavePriv))
    Chk包含停用材料.Value = IIf(Val(zlDatabase.GetPara("包含停用卫材", glngSys, mlngModule, , Array(Txt效期报警, Lbl效期报警, Lbl月), blnHavePriv)) = 1, 1, 0)
    mblnBootUp = True
End Sub

Private Sub Opt单位1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub Opt单位2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub Txt效期报警_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt效期报警_KeyPress(KeyAscii As Integer)
   zlControl.TxtCheckKeyPress Txt效期报警, KeyAscii, m数字式
End Sub
