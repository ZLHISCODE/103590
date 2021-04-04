VERSION 5.00
Begin VB.Form frmPatholReborrowParameter 
   Caption         =   "��������"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4710
   Icon            =   "frmPatholReborrowParameter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4710
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkSurePrint 
         Caption         =   "����ȷ�Ϻ��Զ���ӡ���Ļ�ִ��"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtQueryDays 
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cbxReportName 
         Height          =   300
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "���ļ�¼Ĭ�ϲ�ѯ������"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "���Ļ�ִ��Ӧ�������ƣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   760
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "��"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   280
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatholReborrowParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngDefaultQueryDays As Long
Public strLabelReportName As String
Public blnIsAutoPrint   As Boolean


Public Sub ShowParameterWindow(ByVal lngCurDefaultQueryDays As Long, ByVal strCurReportName As String, _
    ByVal blnCurIsAutoPrint As Boolean, owner As Object)
    
    lngDefaultQueryDays = lngCurDefaultQueryDays
    strLabelReportName = strCurReportName
    
    txtQueryDays.Text = lngDefaultQueryDays
    cbxReportName.Text = strLabelReportName
    chkSurePrint.value = IIf(blnCurIsAutoPrint, 1, 0)
    
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
    blnIsAutoPrint = chkSurePrint.value
    
    Call zlDatabase.SetPara("����Ĭ�ϲ�ѯ����", Val(txtQueryDays.Text), glngSys, G_LNG_PATHOLBORROW_NUM)
    Call zlDatabase.SetPara("���Ļ�ִ��������", cbxReportName.Text, glngSys, G_LNG_PATHOLBORROW_NUM)
    Call zlDatabase.SetPara("����ȷ�Ϻ��Զ���ӡ��ִ", IIf(chkSurePrint.value = 0, 0, 1), glngSys, G_LNG_PATHOLBORROW_NUM)
    
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
