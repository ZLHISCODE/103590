VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholArchivesLocate 
   Caption         =   "过滤"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   Icon            =   "frmPatholArchivesLocate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4335
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 认(&S)"
      Height          =   400
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame framFilter 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtPatholNum 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   780
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   155058179
         CurrentDate     =   40921
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2640
         TabIndex        =   3
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   155058179
         CurrentDate     =   40921
      End
      Begin VB.Label Label1 
         Caption         =   "姓    名："
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "报到时间："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "到"
         Height          =   255
         Left            =   2430
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "病 理 号："
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholArchivesLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnOk As Boolean


Public Sub ShowFilterWindow(owner As Object)
    blnOk = False
    
    Me.Show 1, owner
End Sub

Private Sub cmdCancel_Click()
    blnOk = False
    
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
    blnOk = True
    
    Call Me.Hide
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
    Call RestoreWinState(Me, App.ProductName)
    
    curDate = zlDatabase.Currentdate
    
    dtpStart.value = curDate
    dtpEnd.value = curDate
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub
