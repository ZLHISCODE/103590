VERSION 5.00
Begin VB.Form frmPatholReportDelay_Select 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ӳ�ԭ��"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
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
         Caption         =   "��ɷ�"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkFZBL 
         Caption         =   "����Ӳ���"
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkTSRS 
         Caption         =   "������Ⱦɫ"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkMYZH 
         Caption         =   "�������黯"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkBQC 
         Caption         =   "�貹ȡ��"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkTG 
         Caption         =   "���Ѹ�"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox chkLQ 
         Caption         =   "������"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkCQ 
         Caption         =   "������"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkSQ 
         Caption         =   "������"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label labRecordInf 
      Caption         =   "������"
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
    
    chkJF.value = IIf(InStr(1, strCurReason, "��ɷ�") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "��ɷ�", "")
    
    chkTG.value = IIf(InStr(1, strCurReason, "���Ѹ�") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "���Ѹ�", "")
    
    chkBQC.value = IIf(InStr(1, strCurReason, "�貹ȡ��") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "�貹ȡ��", "")
    
    chkSQ.value = IIf(InStr(1, strCurReason, "������") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "������", "")
    
    chkCQ.value = IIf(InStr(1, strCurReason, "������") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "������", "")
    
    chkLQ.value = IIf(InStr(1, strCurReason, "������") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "������", "")
    
    chkMYZH.value = IIf(InStr(1, strCurReason, "�������黯") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "�������黯", "")
    
    chkFZBL.value = IIf(InStr(1, strCurReason, "����Ӳ���") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "����Ӳ���", "")
    
    chkTSRS.value = IIf(InStr(1, strCurReason, "������Ⱦɫ") > 0, 1, 0)
    strCurReason = Replace(strCurReason, "������Ⱦɫ", "")
    
    txtOther.Text = Replace(strCurReason, "��", "")
    
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
