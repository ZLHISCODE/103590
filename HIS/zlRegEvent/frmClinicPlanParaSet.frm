VERSION 5.00
Begin VB.Form frmClinicPlanParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   Icon            =   "frmClinicPlanParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSplit 
      Height          =   4365
      Left            =   2520
      TabIndex        =   4
      Top             =   -150
      Width           =   25
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "周出诊表打印设置(&4)"
      Height          =   405
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1950
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "月出诊表打印设置(&3)"
      Height          =   405
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1380
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "固定出诊表打印设置(&2)"
      Height          =   405
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "预约清单打印设置(&1)"
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "退出(&E)"
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmClinicPlanParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModul As Long
Private mblnOk As Boolean

Public Function ShowMe(frmParent As Form, ByVal lngModul As Long, _
    ByVal strPrivs As String) As Boolean
    '程序入口
    mstrPrivs = strPrivs: mlngModul = lngModul
    
    On Error Resume Next
    mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(index As Integer)
    On Error GoTo ErrHandler
    Select Case index
    Case 0: '预约清单打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_4", Me)
    Case 1: '固定出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_1", Me)
    Case 2: '月出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_2", Me)
    Case 3: '周出诊表打印方式
      Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1114_3", Me)
    Case Else:
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

