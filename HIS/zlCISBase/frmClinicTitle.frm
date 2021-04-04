VERSION 5.00
Begin VB.Form frmCureRTitle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参考标题设置"
   ClientHeight    =   2970
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTier 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1650
      TabIndex        =   10
      Top             =   1470
      Width           =   2940
      Begin VB.OptionButton optTier 
         Caption         =   "一级标题(&1)"
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTier 
         Caption         =   "二级标题(&2)"
         Height          =   210
         Index           =   1
         Left            =   1635
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkBan 
      Caption         =   "可设置项目对应的禁忌症(&B)"
      Height          =   285
      Left            =   1650
      TabIndex        =   4
      Top             =   1785
      Width           =   2940
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1650
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "名称"
      Top             =   975
      Width           =   2940
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2730
      TabIndex        =   5
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -30
      TabIndex        =   8
      Top             =   2310
      Width           =   5745
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "标题(&N)"
      Height          =   180
      Left            =   975
      TabIndex        =   0
      Top             =   1050
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmClinicTitle.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    设置应用参考内容的小标题；为了方便参考阅读与应用，尽量和参考内容含义吻合。"
      Height          =   345
      Left            =   975
      TabIndex        =   9
      Top             =   210
      Width           =   4230
   End
End
Attribute VB_Name = "frmCureRTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strLefts As String   '已经存在的前面的标题
Public strRights As String  '已经存在的后面的标题
Public strTitle As String   '编辑产生的标题
Dim intCount As Integer

Private Sub chkBan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    strTitle = ""
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim aryItems() As String
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "标题必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "标题超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    '重复性检查
    aryItems = Split(Mid(strLefts & strRights, 2), ";")
    For intCount = LBound(aryItems) To UBound(aryItems)
        If Split(aryItems(intCount), ",")(2) = Trim(Me.txtName.Text) Then
            MsgBox "该标题已经包含在参考中", vbExclamation, gstrSysName
            Me.txtName.SetFocus
            Exit Sub
        End If
    Next
    '按规定格式组织编辑的项目
    strTitle = Me.Tag & Trim(Me.txtName.Text) & "," & IIf(Me.optTier(0).Value, 1, 2) & "," & IIf(Me.chkBan.Value = 1, 1, 0)
    Me.Hide
End Sub

Private Sub optTier_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub
