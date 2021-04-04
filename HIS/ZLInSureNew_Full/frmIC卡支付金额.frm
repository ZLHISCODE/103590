VERSION 5.00
Begin VB.Form frmIC卡支付金额 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "指定IC卡支付金额"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2472
      TabIndex        =   6
      Top             =   1815
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   1302
      TabIndex        =   5
      Top             =   1815
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   -30
      TabIndex        =   4
      Top             =   1650
      Width           =   4950
   End
   Begin VB.TextBox txt支付金额 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2385
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1230
      Width           =   1725
   End
   Begin VB.Label lbl自付金额 
      AutoSize        =   -1  'True
      Caption         =   "结算费用中自付部分为[0.00元]"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   240
      Width           =   2520
   End
   Begin VB.Label lblMSG 
      AutoSize        =   -1  'True
      Caption         =   "指定IC卡支付金额："
      Height          =   180
      Left            =   765
      TabIndex        =   2
      Top             =   1305
      Width           =   1620
   End
   Begin VB.Label lblIC卡余额 
      AutoSize        =   -1  'True
      Caption         =   "IC卡余额为[0.00元]"
      Height          =   180
      Left            =   300
      TabIndex        =   1
      Top             =   950
      Width           =   1620
   End
   Begin VB.Label lbl中心余额 
      AutoSize        =   -1  'True
      Caption         =   "中心个人帐户余额为[0.00元]"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   595
      Width           =   2340
   End
End
Attribute VB_Name = "frmIC卡支付金额"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function get_IC支付金额(cur自付金额 As Currency, cur中心余额 As Currency, curIC卡余额 As Currency) As Currency
    lbl自付金额.Caption = "结算费用中自付部分为[" & Format(cur自付金额, "0.00") & "元]"
    lbl自付金额.Tag = cur自付金额
    lbl中心余额.Caption = "中心个人帐户余额为[" & Format(cur中心余额, "0.00") & "元]"
    lbl中心余额.Tag = cur中心余额
    lblIC卡余额.Caption = "IC卡余额为[" & Format(curIC卡余额, "0.00") & "元]"
    lblIC卡余额.Tag = curIC卡余额
    
    If cur自付金额 < cur中心余额 + curIC卡余额 Then
        txt支付金额.Text = Format(cur自付金额, "0.00")
    Else
        txt支付金额.Text = Format(cur中心余额 + curIC卡余额, "0.00")
    End If
    txt支付金额.Tag = cur中心余额 + curIC卡余额
    
    Me.Show vbModal
    get_IC支付金额 = CCur(txt支付金额.Text)
End Function

Private Sub cmdCancel_Click()
    txt支付金额.Text = -1
End Sub

Private Sub cmdOK_Click()
    txt支付金额.SetFocus
    If Not IsNumeric(txt支付金额.Text) Then
        MsgBox "请输入IC卡支付金额。", vbInformation, Me.Caption
        Exit Sub
    End If
    If CCur(txt支付金额.Text) < 0 Then
        MsgBox "IC卡支付金额不能为负数。", vbInformation, Me.Caption
        Exit Sub
    End If
    If Len(Format(txt支付金额.Text, "0.00")) > 12 Then
        MsgBox "支付金额的整数部分应小于12位，小数部分应小于2位。", vbInformation, Me.Caption
        Exit Sub
    End If
    If CCur(txt支付金额.Text) > CCur(txt支付金额.Tag) Then
        MsgBox "支付金额超出了中心帐户余额与IC卡余额的和。", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub txt支付金额_GotFocus()
    txt支付金额.SelStart = 0
    txt支付金额.SelLength = Len(txt支付金额.Text)
End Sub
