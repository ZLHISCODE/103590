VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSchSetTimeTableColoe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查预约--时间表颜色设置"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmSchSetTimeTableColoe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   720
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "…"
      Height          =   255
      Index           =   4
      Left            =   3975
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "…"
      Height          =   255
      Index           =   3
      Left            =   3975
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "…"
      Height          =   255
      Index           =   2
      Left            =   3975
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "…"
      Height          =   255
      Index           =   0
      Left            =   3975
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "…"
      Height          =   255
      Index           =   1
      Left            =   3975
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lab 
      Caption         =   "预约标签颜色：已过号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   2880
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "预约标签颜色：已完成"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   2880
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "预约标签颜色：已预约"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   2880
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "时间表颜色：工作时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2880
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lab 
      Caption         =   "时间表颜色：休息时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.Shape shpColor 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   2880
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmSchSetTimeTableColoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngColorTabRest As Long    '预约时间表，休息时间颜色
Private mlngColorTabWork As Long    '预约时间表，工作时间颜色
Private mlngColorLblWaiting As Long '预约标签，预约等候颜色
Private mlngColorLblDone As Long    '预约标签，完成颜色
Private mlngColorLblPassed As Long  '预约标签，过号颜色

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click(Index As Integer)
    dlgColor.Color = shpColor(Index).FillColor
    dlgColor.ShowColor
    shpColor(Index).FillColor = dlgColor.Color
End Sub

Private Sub cmdOK_Click()
    
    Call zlDatabase.SetPara("检查预约时间表工作时间颜色", shpColor(0).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("检查预约时间表休息时间颜色", shpColor(1).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("检查预约标签已预约颜色", shpColor(2).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("检查预约标签已完成颜色", shpColor(3).FillColor, glngSys, 1292)
    Call zlDatabase.SetPara("检查预约标签已过号颜色", shpColor(4).FillColor, glngSys, 1292)
    Unload Me
End Sub

Private Sub Form_Load()
        
    Call LoadColors
    
    shpColor(0).FillColor = mlngColorTabWork
    shpColor(1).FillColor = mlngColorTabRest
    shpColor(2).FillColor = mlngColorLblWaiting
    shpColor(3).FillColor = mlngColorLblDone
    shpColor(4).FillColor = mlngColorLblPassed
    
End Sub

Private Sub LoadColors()
'------------------------------------------------
'功能：装载时间表的表格格式和基础内容
'参数：
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    '从数据库中读取设置过的颜色
    mlngColorTabWork = zlDatabase.GetPara("检查预约时间表工作时间颜色", glngSys, 1292, "8421376")
    mlngColorTabRest = zlDatabase.GetPara("检查预约时间表休息时间颜色", glngSys, 1292, "16777215")
    mlngColorLblWaiting = zlDatabase.GetPara("检查预约标签已预约颜色", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("检查预约标签已完成颜色", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("检查预约标签已过号颜色", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

