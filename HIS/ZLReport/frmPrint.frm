VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报表打印"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2820
      TabIndex        =   10
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2820
      TabIndex        =   11
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "副本"
      Height          =   660
      Left            =   105
      TabIndex        =   13
      Top             =   1770
      Width           =   2610
      Begin MSComCtl2.UpDown udCopy 
         Height          =   300
         Left            =   1485
         TabIndex        =   9
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCopy"
         BuddyDispid     =   196613
         OrigLeft        =   1935
         OrigTop         =   240
         OrigRight       =   2175
         OrigBottom      =   585
         Max             =   255
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印份数"
         Height          =   180
         Left            =   285
         TabIndex        =   14
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "范围"
      Height          =   1620
      Left            =   105
      TabIndex        =   12
      Top             =   75
      Width           =   2595
      Begin VB.CheckBox chkOrder 
         Caption         =   "按倒序打印"
         Enabled         =   0   'False
         Height          =   195
         Left            =   570
         TabIndex        =   7
         Top             =   1290
         Width           =   1200
      End
      Begin VB.OptionButton optPage 
         Caption         =   "偶数页"
         Height          =   180
         Index           =   4
         Left            =   1200
         TabIndex        =   6
         Top             =   970
         Width           =   840
      End
      Begin VB.OptionButton optPage 
         Caption         =   "奇数页"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   970
         Width           =   840
      End
      Begin VB.TextBox txtEnd 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   4
         Top             =   570
         Width           =   450
      End
      Begin VB.TextBox txtBegin 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   720
         MaxLength       =   8
         TabIndex        =   3
         Top             =   570
         Width           =   450
      End
      Begin VB.OptionButton optPage 
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   620
         Width           =   270
      End
      Begin VB.OptionButton optPage 
         Caption         =   "当前页"
         Height          =   180
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   285
         Width           =   840
      End
      Begin VB.OptionButton optPage 
         Caption         =   "所有页"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从      页到      页"
         Height          =   180
         Left            =   510
         TabIndex        =   15
         Top             =   630
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstr编号 As String '报表编号
Public mbln票据 As Boolean '是否票据
Public mintMax As Integer '入:最大页号
Public mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtBegin.Enabled And Not IsNumeric(txtBegin.Text) Then
        MsgBox "非法的开始页号！", vbInformation, App.Title
        txtBegin.SetFocus: Exit Sub
    End If
    If txtEnd.Enabled And Not IsNumeric(txtEnd.Text) Then
        MsgBox "非法的结束页号！", vbInformation, App.Title
        txtEnd.SetFocus: Exit Sub
    End If
    If txtBegin.Enabled And (CLng(txtBegin.Text) < 1 Or CLng(txtBegin.Text) > mintMax) Then
        MsgBox "开始页号必须在 1-" & mintMax & " 之间！", vbInformation, App.Title
        txtBegin.SetFocus: Exit Sub
    End If
    If txtEnd.Enabled And (CLng(txtEnd.Text) < 1 Or CLng(txtEnd.Text) > mintMax) Then
        MsgBox "结束页号必须在 1-" & mintMax & " 之间！", vbInformation, App.Title
        txtEnd.SetFocus: Exit Sub
    End If
    If txtBegin.Enabled And CLng(txtEnd.Text) < CLng(txtBegin.Text) Then
        MsgBox "结束页号不能小于开始页号！", vbInformation, App.Title
        txtEnd.SetFocus: Exit Sub
    End If
    mblnOK = True
    Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Form_Load()
    mblnOK = False
    txtBegin.Text = 1: txtEnd.Text = mintMax
    If mintMax = 1 Then optPage(4).Enabled = False
    txtCopy.Text = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & mstr编号, "PaperCopy", 1))
    If Val(txtCopy.Text) < 1 Then txtCopy.Text = 1
    
    '如果是票据，则只能打印1份
    If mbln票据 Then
        txtCopy.Enabled = False
        udCopy.Enabled = False
        txtCopy.Text = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr编号 = Empty
    mbln票据 = Empty
    mintMax = Empty
End Sub

Private Sub optPage_Click(Index As Integer)
    If Index = 2 Then
        txtBegin.Enabled = True
        txtEnd.Enabled = True
        txtBegin.SetFocus
    Else
        txtBegin.Enabled = False
        txtEnd.Enabled = False
    End If
    
    chkOrder.Enabled = Index = 3 Or Index = 4
    If Not chkOrder.Enabled Then
        chkOrder.Value = 0
    Else
        chkOrder.Value = IIF(Index = 3, 0, 1)
    End If
End Sub

Private Sub txtBegin_GotFocus()
    SelAll txtBegin
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEnd_GotFocus()
    SelAll txtEnd
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
