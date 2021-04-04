VERSION 5.00
Begin VB.Form frmPatholReborrowParameter 
   Caption         =   "参数配置"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4710
   Icon            =   "frmPatholReborrowParameter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4710
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkSurePrint 
         Caption         =   "借阅确认后自动打印借阅回执单"
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
         Caption         =   "借阅记录默认查询天数："
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "借阅回执对应报表名称："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   760
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "天"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   280
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
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
'取消设置
On Error GoTo errHandle
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
'确认设置
On Error GoTo errHandle
    lngDefaultQueryDays = Val(txtQueryDays.Text)
    strLabelReportName = cbxReportName.Text
    blnIsAutoPrint = chkSurePrint.value
    
    Call zlDatabase.SetPara("借阅默认查询天数", Val(txtQueryDays.Text), glngSys, G_LNG_PATHOLBORROW_NUM)
    Call zlDatabase.SetPara("借阅回执报表名称", cbxReportName.Text, glngSys, G_LNG_PATHOLBORROW_NUM)
    Call zlDatabase.SetPara("借阅确认后自动打印回执", IIf(chkSurePrint.value = 0, 0, 1), glngSys, G_LNG_PATHOLBORROW_NUM)
    
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
