VERSION 5.00
Begin VB.Form frmPatholReportDelay_Select 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "延迟原因"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5070
   Icon            =   "frmPatholReportDelay_Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   3600
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtOther 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkJF 
         Caption         =   "需缴费"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkFZBL 
         Caption         =   "需分子病理"
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkTSRS 
         Caption         =   "需特殊染色"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkMYZH 
         Caption         =   "需免疫组化"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkBQC 
         Caption         =   "需补取材"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkTG 
         Caption         =   "需脱钙"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkLQ 
         Caption         =   "需连切"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkCQ 
         Caption         =   "需重切"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkSQ 
         Caption         =   "需深切"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label labRecordInf 
      Caption         =   "其他："
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1755
      Width           =   615
   End
End
Attribute VB_Name = "frmPatholReportDelay_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOk As Boolean



Public Sub ShowReasonWindow(ByVal strReason As String, owner As Form)
    Dim strCurReason As String
    strCurReason = strReason
    
    chkJF.value = IIf(InStr(1, strCurReason, "需缴费") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需缴费", "")
    
    chkTG.value = IIf(InStr(1, strCurReason, "需脱钙") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需脱钙", "")
    
    chkBQC.value = IIf(InStr(1, strCurReason, "需补取材") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需补取材", "")
    
    chkSQ.value = IIf(InStr(1, strCurReason, "需深切") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需深切", "")
    
    chkCQ.value = IIf(InStr(1, strCurReason, "需重切") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需重切", "")
    
    chkLQ.value = IIf(InStr(1, strCurReason, "需连切") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需连切", "")
    
    chkMYZH.value = IIf(InStr(1, strCurReason, "需免疫组化") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需免疫组化", "")
    
    chkFZBL.value = IIf(InStr(1, strCurReason, "需分子病理") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需分子病理", "")
    
    chkTSRS.value = IIf(InStr(1, strCurReason, "需特殊染色") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "需特殊染色", "")
    
    txtOther.Text = Replace(strCurReason, "、", "")
    
    Call Me.Show(1, owner)
End Sub



Private Sub cmdCancel_Click()
    IsOk = False
    
    Me.Hide
End Sub



Private Sub Command1_Click()
    IsOk = True
    Me.Hide
End Sub



Private Sub Form_Initialize()
    IsOk = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
