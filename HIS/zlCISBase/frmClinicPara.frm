VERSION 5.00
Begin VB.Form frmClinicPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmClinicPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   8
      Top             =   3615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2010
      TabIndex        =   7
      Top             =   3615
      Width           =   1100
   End
   Begin VB.Frame fraAddMode 
      Caption         =   " 1、项目增加操作模式"
      Height          =   1365
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4155
      Begin VB.OptionButton opt增加模式 
         Caption         =   "单项增加(保存后关闭编辑)"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2580
      End
      Begin VB.OptionButton opt增加模式 
         Caption         =   "连续增加(保存后自动增加项目)"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 2、项目允许应用范围控制"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1695
      Width           =   4155
      Begin VB.CheckBox chk应用范围 
         Caption         =   "允许应用于同级所有项目"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "允许应用于同分类所有项目"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chk应用范围 
         Caption         =   "允许应用于同类别所有项目"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmClinicPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrPrivs As String

Public Sub ShowMe(ByVal frmParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim str应用范围 As String
    
    If Me.opt增加模式(0).Value = True Then
        Call zlDatabase.SetPara("诊疗项目连续增加", 0, glngSys, 1054)
    Else
        Call zlDatabase.SetPara("诊疗项目连续增加", 1, glngSys, 1054)
    End If
    
    str应用范围 = IIf(chk应用范围(0).Value = 1, "1", "0")
    str应用范围 = str应用范围 & IIf(chk应用范围(1).Value = 1, "1", "0")
    str应用范围 = str应用范围 & IIf(chk应用范围(2).Value = 1, "1", "0")
    
    Call zlDatabase.SetPara("项目应用范围", str应用范围, glngSys, 1054)
    
    Unload Me
End Sub

Private Sub Form_Load()
    '根据用户权限，装入控件
    Dim lngValues As Long
    Dim str应用范围 As String
    Dim blnSetPara As Boolean
    
    blnSetPara = zlStr.IsHavePrivs(mstrPrivs, "参数设置")
    
    lngValues = Val(zlDatabase.GetPara("诊疗项目连续增加", glngSys, 1054, 0, Array(Me.opt增加模式(0), Me.opt增加模式(1)), blnSetPara))
    str应用范围 = zlDatabase.GetPara("项目应用范围", glngSys, 1054, "000", Array(chk应用范围(0), chk应用范围(1), chk应用范围(2)), blnSetPara)
    
    If lngValues = 0 Then
        Me.opt增加模式(0).Value = True: Me.opt增加模式(1).Value = False
    Else
        Me.opt增加模式(0).Value = False: Me.opt增加模式(1).Value = True
    End If
    
    If Val(Mid(str应用范围, 1, 1)) = 1 Then
        chk应用范围(0).Value = 1
    End If
    
    If Val(Mid(str应用范围, 2, 1)) = 1 Then
        chk应用范围(1).Value = 1
    End If
    
    If Val(Mid(str应用范围, 3, 1)) = 1 Then
        chk应用范围(2).Value = 1
    End If
End Sub

