VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "副本"
      Height          =   705
      Left            =   90
      TabIndex        =   9
      Top             =   1620
      Width           =   2085
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   270
         Left            =   1410
         TabIndex        =   12
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtNumber"
         BuddyDispid     =   196611
         OrigLeft        =   1470
         OrigTop         =   270
         OrigRight       =   1710
         OrigBottom      =   465
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNumber 
         Height          =   270
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "份数(&C):"
         Height          =   180
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.TextBox txtBegin 
      Height          =   285
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1110
      Width           =   495
   End
   Begin VB.OptionButton optRange 
      Caption         =   "由第"
      Height          =   345
      Left            =   390
      TabIndex        =   5
      Top             =   1080
      Width           =   705
   End
   Begin VB.OptionButton optCurrent 
      Caption         =   "当前页(&R)"
      Height          =   345
      Left            =   390
      TabIndex        =   3
      Top             =   690
      Width           =   1185
   End
   Begin VB.OptionButton optAll 
      Caption         =   "全部(&A)"
      Height          =   345
      Left            =   390
      TabIndex        =   2
      Top             =   270
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "范围"
      Height          =   1515
      Left            =   90
      TabIndex        =   8
      Top             =   60
      Width           =   3225
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   2100
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "页到第      页(&S)"
         Height          =   195
         Left            =   1530
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   1
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   3510
      TabIndex        =   0
      Top             =   180
      Width           =   1100
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnPrint As Boolean
'本窗体用于选择打印范围

Public Function PrintData() As Boolean
    mblnPrint = False
    Me.Show 1
    PrintData = mblnPrint
End Function

Private Sub cmdCancel_Click()
    mblnPrint = False
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If optAll Then
        Call RealPrint(1, gintRowTotal * gintColTotal)
    ElseIf optCurrent Then
        Call RealPrint(gintPage, gintPage)
    Else
        If CInt(txtBegin.Text) < 1 Or CInt(txtBegin.Text) > gintRowTotal * gintColTotal Then
            MsgBox "输出的页码超出范围了。", vbCritical, gstrSysName
            txtBegin.SelStart = 0
            txtBegin.SelLength = 5
            txtBegin.SetFocus
            Exit Sub
        End If
        If CInt(txtEnd.Text) < 1 Or CInt(txtEnd.Text) > gintRowTotal * gintColTotal Then
            MsgBox "输出的页码超出范围了。", vbCritical, gstrSysName
            txtEnd.SelStart = 0
            txtEnd.SelLength = 5
            txtEnd.SetFocus
            Exit Sub
        End If
        If CInt(txtEnd.Text) < CInt(txtBegin.Text) Then
            MsgBox "结束页码超过了开始页码。", vbCritical, gstrSysName
            txtBegin.SelStart = 0
            txtBegin.SelLength = 5
            txtBegin.SetFocus
            Exit Sub
        End If
        gintCopies = Val(txtNumber.Text)
        If gintCopies < 1 Then gintCopies = 1
        If gintCopies > 99 Then gintCopies = 99
        Call RealPrint(CInt(txtBegin.Text), CInt(txtEnd.Text))
    End If
    
    mblnPrint = True
    Unload Me
End Sub

Private Sub Form_Load()
    txtNumber.Text = CStr(gintCopies)
    txtBegin.Text = 1
    txtEnd.Text = CStr(gintColTotal * gintRowTotal)
End Sub

Private Sub txtBegin_Change()
    optRange.Value = True
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEnd_Change()
    optRange.Value = True
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumber_Change()
    If txtNumber.Text = "" Then txtNumber.Text = "1"
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub
