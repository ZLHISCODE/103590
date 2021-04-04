VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInsTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插入表格"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   3720
   Icon            =   "frmInsTable.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   1260
      TabIndex        =   4
      Top             =   1395
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   2520
      TabIndex        =   5
      Top             =   1395
      Width           =   1050
   End
   Begin MSComCtl2.UpDown upCol 
      Height          =   270
      Left            =   3315
      TabIndex        =   1
      Top             =   510
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCol"
      BuddyDispid     =   196610
      OrigLeft        =   4365
      OrigTop         =   495
      OrigRight       =   4620
      OrigBottom      =   735
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   1080
      TabIndex        =   6
      Top             =   225
      Width           =   2490
   End
   Begin VB.TextBox txtCol 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   2430
      TabIndex        =   0
      Text            =   "1"
      Top             =   510
      Width           =   885
   End
   Begin VB.TextBox txtRow 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   2430
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   885
   End
   Begin MSComCtl2.UpDown upRow 
      Height          =   270
      Left            =   3315
      TabIndex        =   3
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtRow"
      BuddyDispid     =   196609
      OrigLeft        =   4410
      OrigTop         =   855
      OrigRight       =   4665
      OrigBottom      =   1095
      Max             =   9999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "表格尺寸"
      ForeColor       =   &H000052D9&
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   225
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "列数(&C):"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   540
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "行数(&R):"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   885
      Width           =   1185
   End
End
Attribute VB_Name = "frmInsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean, lRow As Long, lCol As Long

Public Function ShowMe(ByRef frmParent As Object, ByRef R As Long, C As Long) As Boolean
    Dim lRows As Long, lCols As Long
    lRows = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Rows", 3)
    lCols = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Cols", 3)
    txtRow.Text = lRows
    txtCol.Text = lCols
    
    Me.Show vbModal, frmParent
    
    If mblnOK Then
        R = lRow
        C = lCol
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Rows", lRow
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Cols", lCol
    End If
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    lRow = Val(txtRow)
    lCol = Val(txtCol)
    mblnOK = True
    Unload Me
End Sub

Private Sub txtCol_GotFocus()
    txtCol.SelStart = 0
    txtCol.SelLength = Len(txtCol)
End Sub

Private Sub txtRow_GotFocus()
    txtRow.SelStart = 0
    txtRow.SelLength = Len(txtRow)
End Sub
