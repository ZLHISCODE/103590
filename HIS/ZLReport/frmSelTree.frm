VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelTree 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmSelTree.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4725
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCmd 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   3405
      ScaleHeight     =   4575
      ScaleWidth      =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1320
      Begin MSComctlLib.ImageList img16 
         Left            =   315
         Top             =   1650
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelTree.frx":014A
               Key             =   "ReportNode"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelTree.frx":02A4
               Key             =   "GroupNode"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelTree.frx":03FE
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelTree.frx":0998
               Key             =   "Path"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelTree.frx":0F32
               Key             =   "App"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   750
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1100
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   4530
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   7990
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      PathSeparator   =   "."
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSelTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'参数说明：入
'1;tvw.Tag=没有选择时的提示信息
'2:Node.Tag="",表示可选,<>"",则不可选,同时为提示信息

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If tvw.SelectedItem Is Nothing Then
        If tvw.Tag <> "" Then
            MsgBox tvw.Tag, vbInformation, App.Title: tvw.SetFocus: Exit Sub
        Else
            MsgBox "无可选择的内容。", vbInformation, App.Title: tvw.SetFocus: Exit Sub
        End If
    End If
    If tvw.SelectedItem.Tag <> "" Then
        MsgBox tvw.SelectedItem.Tag, vbInformation, App.Title: tvw.SetFocus: Exit Sub
    End If
    gblnOK = True
    Hide
End Sub

Private Sub Form_Load()
    gblnOK = False
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tvw.Left = 0
    tvw.Top = 0
    tvw.Width = Me.ScaleWidth - picCmd.Width
    tvw.Height = Me.ScaleHeight
    
    If Not tvw.SelectedItem Is Nothing Then tvw.SelectedItem.EnsureVisible
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub tvw_DblClick()
    cmdOK_Click
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK_Click
End Sub
