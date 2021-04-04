VERSION 5.00
Begin VB.Form frmMediClassReuse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品分类启用"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3795
   Icon            =   "frmMediClassReuse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "启用该分类目录时是否同时进行以下操作"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chk启用规格 
         Caption         =   "启用该分类下所有规格"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chk启用品种 
         Caption         =   "启用该分类下所有品种"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chk启用子目录 
         Caption         =   "启用该分类下所有子目录"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmMediClassReuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng分类id As Long
Private mstr类型 As String


Public Sub ShowForm(ByVal lng分类id As Long, ByVal str类型 As String)
    mlng分类id = lng分类id
    mstr类型 = str类型
    
    frmMediClassReuse.Show vbModal
    Exit Sub
End Sub


Private Sub chk启用品种_Click()
    If chk启用品种.Value = 1 Then
        chk启用规格.Enabled = True
    Else
        chk启用规格.Value = 0
        chk启用规格.Enabled = False
    End If
End Sub

Private Sub chk启用子目录_Click()
    If chk启用子目录.Value = 1 Then
        chk启用品种.Enabled = True
    Else
        chk启用品种.Value = 0
        chk启用品种.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim int启用子目录 As Integer
    Dim int启用品种 As Integer
    Dim int启用规格 As Integer
    
    int启用子目录 = chk启用子目录.Value
    
    If chk启用品种.Enabled Then
        int启用品种 = chk启用品种.Value
    End If
    
    If chk启用规格.Enabled Then
        int启用规格 = chk启用规格.Value
    End If
    
    On Error GoTo ErrHand
    
    gstrSql = "Zl_诊疗分类目录_药品分类启用(" & mlng分类id & "," & Val(mstr类型) & "," & int启用子目录 & "," & int启用品种 & "," & int启用规格 & " )"
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


