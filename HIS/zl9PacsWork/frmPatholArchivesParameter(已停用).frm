VERSION 5.00
Begin VB.Form frmPatholArchivesParameter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4695
   Icon            =   "frmPatholArchivesParameter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cbxReportName 
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtQueryDays 
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Text            =   "30"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "��"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "������ǩ��Ӧ�������ƣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "������¼Ĭ�ϲ�ѯ������"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPatholArchivesParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngDefaultQueryDays As Long
Public strLabelReportName As String


Public Sub ShowParameterWindow(ByVal lngCurDefaultQueryDays As Long, ByVal strCurReportName As String, owner As Object)
    lngDefaultQueryDays = lngCurDefaultQueryDays
    strLabelReportName = strCurReportName
    
    txtQueryDays.Text = lngDefaultQueryDays
    cbxReportName.Text = strLabelReportName
    
    Call Me.Show(1, owner)
End Sub


Private Sub cmdCancel_Click()
'ȡ������
On Error GoTo errHandle
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
'ȷ������
On Error GoTo errHandle
    lngDefaultQueryDays = Val(txtQueryDays.Text)
    strLabelReportName = cbxReportName.Text
    
    Call zlDatabase.SetPara("����Ĭ�ϲ�ѯ����", Val(txtQueryDays.Text), glngSys, G_LNG_PATHOLARCHIVES_NUM)
    Call zlDatabase.SetPara("������ǩ��������", cbxReportName.Text, glngSys, G_LNG_PATHOLARCHIVES_NUM)
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHandle
    Call SaveWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
