VERSION 5.00
Begin VB.Form frmNameEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmNameEdit.frx":0000
   LinkTopic       =   "Form1"
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
      Height          =   570
      Index           =   3
      Left            =   105
      Picture         =   "frmNameEdit.frx":000C
      Stretch         =   -1  'True
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image img 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "frmNameEdit.frx":685E
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "frmNameEdit.frx":6A68
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
      Height          =   480
      Index           =   0
      Left            =   180
      Picture         =   "frmNameEdit.frx":78AA
      Top             =   180
      Visible         =   0   'False
      Width           =   480
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
Attribute VB_Name = "frmNameEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum nameObject
    name角色 = 0
    name菜单 = 1
    name模块 = 2
    name组名 = 3
End Enum

Private mstrName As String
Private mstrCaption As String

Private Sub cmdCancel_Click()
    mstrName = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtName) = "" Then
        MsgBox mstrName & "名不能为空。", vbExclamation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If InStr(txtName, "'") > 0 Or InStr(txtName, """") > 0 Then
        MsgBox mstrName & "名不能含有单引号和双引号。", vbExclamation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    mstrName = UCase(Trim(txtName.Text))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        mstrName = ""
    End If
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Public Function GetName(ByVal name As nameObject) As String
    mstrName = ""
    img(name).Visible = True
    Select Case name
        Case name角色
            txtName.MaxLength = 12
            mstrName = "角色"
            lblName.Caption = "角色名称"
            frmNameEdit.Caption = "角色增加"
            lblTitle.Caption = "增加的角色不具有任何权限，只有返回上级窗体分配权限以后才能使用。角色名称尽可能表达你将分配的权限。"
            txtName.Text = "新增" & mstrName
        Case name菜单
            txtName.MaxLength = 30
            mstrName = "菜单"
            lblName.Caption = "菜单名称"
            frmNameEdit.Caption = "菜单增加"
            lblTitle.Caption = "新增一个菜单体系实际上就是把当前选择的菜单体系克隆一份，它们拥有相同的结构和内容。（注：缺省菜单体系是不能修改的）"
            txtName.Text = "新增" & mstrName
        Case name模块
            txtName.MaxLength = 40
            mstrName = "模块"
            lblName.Caption = "名称序号"
            frmNameEdit.Caption = "模块查找"
            lblTitle.Caption = "请输入要查找的模块名称或模块序号,支持模糊查找."
            txtName.Text = ""
        Case name组名
            txtName.MaxLength = 30
            mstrName = "组名"
            lblName.Caption = "组名称"
            frmNameEdit.Caption = "组增加"
            lblTitle.Caption = "请输入要新建的组名称"
            txtName.Text = ""
    End Select
    
    frmNameEdit.Show vbModal, frmMDIMain
    GetName = mstrName
End Function

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call cmdOK_Click
    End If
End Sub
