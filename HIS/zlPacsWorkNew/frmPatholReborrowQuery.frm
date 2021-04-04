VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholReborrowQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "借阅查询"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   Icon            =   "frmPatholReborrowQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtBorrowNum 
         Height          =   300
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtCardNum 
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   330
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd 00:00:00"
         Format          =   163119107
         CurrentDate     =   40884
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   330
         Left            =   2775
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd 23:59:59"
         Format          =   163119107
         CurrentDate     =   40884
      End
      Begin VB.Label Label1 
         Caption         =   "到"
         Height          =   255
         Left            =   2565
         TabIndex        =   11
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "借阅日期："
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "借 阅 号："
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   760
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "证 件 号："
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "姓    名："
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1740
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查 询(&Q)"
      Height          =   400
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatholReborrowQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mrsArchivesClass As ADODB.Recordset


Public mblnIsOk As Boolean

Public dtStartDate As Date
Public dtEndDate As Date

Public strBorrowId As String
Public strCardNo As String
Public strBorrowName As String


Public Sub ShowBorrowQueryWindow(ByVal lngDefaultQueryDays As Long, owner As Object)
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    
    If dtpStartDate.value = dtpEndDate.value And Format(dtpEndDate.value, "yymmdd") = Format(curDate, "yymmdd") Then
        dtpStartDate.value = Format(curDate - lngDefaultQueryDays, "yy-mm-dd 00:00:00")
    End If
    
    mblnIsOk = False
    
    Me.Show 1, owner
End Sub


Private Sub cmdCancel_Click()
    Call Me.Hide
End Sub


Private Sub cmdQuery_Click()
On Error GoTo errHandle
    dtStartDate = dtpStartDate.value
    dtEndDate = dtpEndDate.value
    
    strBorrowId = txtBorrowNum.Text
    
    strCardNo = txtCardNum.Text
    strBorrowName = txtName.Text
    
    mblnIsOk = True
    
    Call Me.Hide
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim curDate As Date
    
    Call RestoreWinState(Me, App.ProductName)
    
    curDate = zlDatabase.Currentdate
    
    dtpStartDate.value = curDate
    dtpEndDate.value = curDate
    
'    dtpStartDate = zlDatabase.Currentdate
'    dtEndDate = zlDatabase.Currentdate
    
    strBorrowId = ""
    
    strCardNo = ""
    strBorrowName = ""
    
    mblnIsOk = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub
