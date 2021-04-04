VERSION 5.00
Begin VB.Form frm3DSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "三维重建设置"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frm3DSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUse3D 
      Caption         =   "启动3D重建"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame frm3DSetup 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox chkAutoDecompress 
         Caption         =   "自动解压缩"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "PACS中图像使用JPEG压缩时，勾选此选项"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txt3DPara 
         Height          =   350
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   4455
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "灌注成像"
            Height          =   255
            Index           =   6
            Left            =   3240
            TabIndex        =   15
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "表面重建"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   14
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "虚拟内窥镜"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "MMPR"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "MPR"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chk3DFunc 
            Caption         =   "容积重建"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   8
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txt3DExeDir 
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "3D功能："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "3D参数："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   768
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "3D程序路径："
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   288
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   3600
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   1100
   End
End
Attribute VB_Name = "frm3DSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String             '权限字符串

Public Sub ShowMe(frmParent As Form, strPrivs As String)
    mstrPrivs = strPrivs
    Me.Show 1, frmParent
End Sub

Private Sub chkUse3D_Click()
    If chkUse3D.value = 0 Then
        frm3DSetup.Enabled = False
    Else
        frm3DSetup.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim str3DFuncs As String
    Dim i As Integer
    
    '保存3D参数
    zlDatabase.SetPara "启用三维重建", chkUse3D.value, glngSys, 1290, CheckPopedom(mstrPrivs, "参数设置")
    zlDatabase.SetPara "3D程序路径", txt3DExeDir.Text, glngSys, 1290, CheckPopedom(mstrPrivs, "参数设置")
    zlDatabase.SetPara "3D参数", txt3DPara.Text, glngSys, 1290, CheckPopedom(mstrPrivs, "参数设置")
    zlDatabase.SetPara "3D自动解压缩", chkAutoDecompress.value, glngSys, 1290, CheckPopedom(mstrPrivs, "参数设置")
    
    For i = 1 To 6
        If chk3DFunc(i).value = 1 Then str3DFuncs = str3DFuncs & "," & i
    Next i
    zlDatabase.SetPara "3D功能", str3DFuncs, glngSys, 1290, CheckPopedom(mstrPrivs, "参数设置")
    Unload Me
End Sub

Private Sub Form_Load()
    Dim str3DFuncs As String
    Dim str3DFunc() As String
    Dim i As Integer
    Dim j As Integer
    Dim i3DFunc As Integer
    
    '初始化默认值
    For i = 1 To 6
        chk3DFunc(i).value = 0
    Next i
    frm3DSetup.Enabled = False
    
    '读取3D的参数
    chkUse3D.value = Val(zlDatabase.GetPara("启用三维重建", glngSys, 1290, 0))
    txt3DExeDir.Text = zlDatabase.GetPara("3D程序路径", glngSys, 1290, "")
    txt3DPara.Text = zlDatabase.GetPara("3D参数", glngSys, 1290, "")
    str3DFuncs = zlDatabase.GetPara("3D功能", glngSys, 1290, "")
    chkAutoDecompress.value = Val(zlDatabase.GetPara("3D自动解压缩", glngSys, 1290, 0))
    
    If str3DFuncs <> "" Then
        str3DFunc = Split(str3DFuncs, ",")
            For j = 1 To UBound(str3DFunc)
                i3DFunc = Val(str3DFunc(j))
                If i3DFunc >= 1 And i3DFunc <= 6 Then
                    chk3DFunc(i3DFunc).value = 1
                End If
            Next j
    End If
End Sub


