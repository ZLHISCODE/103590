VERSION 5.00
Begin VB.Form frmFindByName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmFindByName.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -390
      TabIndex        =   5
      Top             =   1455
      Width           =   5310
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3195
      TabIndex        =   4
      Top             =   1605
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1995
      TabIndex        =   3
      Top             =   1605
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1785
      MaxLength       =   12
      TabIndex        =   2
      Top             =   960
      Width           =   2280
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Caption         =   "#"
      Height          =   630
      Left            =   1005
      TabIndex        =   0
      Top             =   210
      Width           =   3525
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   0
      Left            =   255
      Picture         =   "frmFindByName.frx":058A
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "#"
      Height          =   180
      Left            =   990
      TabIndex        =   1
      Top             =   1020
      Width           =   90
   End
End
Attribute VB_Name = "frmFindByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum nameObject
    name报表 = 0
End Enum

Dim mstrName As String
Dim mstrCaption As String

Private Sub cmdCancel_Click()
    mstrName = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtName) = "" Then
        MsgBox mstrName & "名不能为空。", vbExclamation, App.Title
        txtName.SetFocus
        Exit Sub
    End If
    If InStr(txtName, "'") > 0 Or InStr(txtName, """") > 0 Then
        MsgBox mstrName & "名不能含有单引号和双引号。", vbExclamation, App.Title
        txtName.SetFocus
        Exit Sub
    End If
    mstrName = UCase(Trim(txtName.Text))
    Unload Me
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Public Function GetName(ByVal name As nameObject) As String
    mstrName = ""
    img(name).Visible = True
    Select Case name
        Case name报表
            txtName.MaxLength = 40
            mstrName = "报表"
            lblName.Caption = "报表名称"
            frmFindByName.Caption = "报表查找"
            lblTitle.Caption = "请输入要查找的报表名称,支持模糊查找."
            txtName.Text = ""
    End Select
    
    frmFindByName.Show vbModal, frmMain
    GetName = mstrName
End Function

