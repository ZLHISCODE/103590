VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatientHistoryFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   Icon            =   "frmPatientHistoryFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   3330
      TabIndex        =   7
      Top             =   1740
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   2160
      TabIndex        =   6
      Top             =   1740
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件"
      Height          =   1515
      Left            =   90
      TabIndex        =   8
      Top             =   60
      Width           =   4305
      Begin VB.TextBox TxtPatient 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   270
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker DTPBegin 
         Height          =   300
         Left            =   1290
         TabIndex        =   3
         Top             =   660
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   529
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   38257
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   300
         Left            =   1290
         TabIndex        =   5
         Top             =   1020
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   529
         _Version        =   393216
         Format          =   92930049
         CurrentDate     =   38257
      End
      Begin VB.Label Label3 
         Caption         =   "结束日期(&E)"
         Height          =   165
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "开始日期(&S)"
         Height          =   165
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病人(&P)"
         Height          =   180
         Left            =   585
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmPatientHistoryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    '退出
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '传递要过滤的字串
    With frmPatientHistoryQuery
        .GetFilterStr Me.TxtPatient.Text, Me.DTPBegin, Me.DTPEND
    End With
    Unload Me
End Sub

Private Sub DTPBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub DTPEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub TxtPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
