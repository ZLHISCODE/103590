VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自定义报表密码计算"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "frmPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4155
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   345
      Left            =   2490
      TabIndex        =   3
      Top             =   1665
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1155
      Width           =   2850
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1005
      TabIndex        =   1
      Top             =   630
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1005
      TabIndex        =   0
      Top             =   195
      Width           =   2850
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      Height          =   180
      Left            =   540
      TabIndex        =   6
      Top             =   1215
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报表名称"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   690
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报表编号"
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjReport As Object

Private Sub Command1_Click()
    If Me.Text1.Text = "" Or Me.Text2.Text = "" Then Exit Sub
    Me.Text3.Text = mobjReport.GenReportPass(Me.Text1.Text, Me.Text2.Text)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set mobjReport = CreateObject("zl9Report.clsReport")
    If Err.Number <> 0 Then
        MsgBox "报表工具初始化失败。", vbInformation, "错误"
        Me.Command1.Enabled = False: Exit Sub
    End If
    On Error GoTo 0
    
    Call mobjReport.InitOracle(gcnOracle)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjReport = Nothing
End Sub
