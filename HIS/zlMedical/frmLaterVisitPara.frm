VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLaterVisitPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5160
   Icon            =   "frmLaterVisitPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   4
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   3
      Top             =   135
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   3810
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Left            =   2925
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt"
         BuddyDispid     =   196613
         OrigLeft        =   3345
         OrigTop         =   1035
         OrigRight       =   3585
         OrigBottom      =   1365
         Max             =   12
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   1665
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "随访间隔(&3)                      个月"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Top             =   375
         Width           =   3330
      End
   End
End
Attribute VB_Name = "frmLaterVisitPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mlngLoop As Long
Private mblnOK As Boolean

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    
    mblnOK = False
    
    Set mfrmMain = frmMain
    '初始化
   
    txt.Text = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "随访间隔", 1))
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long

    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "随访间隔", Val(txt.Text))
    
    mblnOK = True

    Unload Me
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

