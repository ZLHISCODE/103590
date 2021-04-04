VERSION 5.00
Begin VB.Form frm清单管理Set 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frm清单管理Set.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   2760
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fra单位 
      Caption         =   "药品部分单位"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton opt单位 
         Caption         =   "药库单位(&4)"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "住院单位(&3)"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "门诊单位(&2)"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "售价单位(&1)"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm清单管理Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub
 
Public Sub 参数设置(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '-------------------------------------------------------------------------------------------
    '功能:提供给上级窗体调用
    '参数:frmMain-父窗体对象
    '     lngModule-模块号
    '     strPrivs-权限串
    '返回:
    '编制:lesfeng
    '修改:2010/02/25
    '-------------------------------------------------------------------------------------------
    mlngModule = lngModule:    mstrPrivs = strPrivs
    mblnHavePriv = IsHavePrivs(mstrPrivs, "参数设置")
    
    Call InitDate
    Me.Show vbModal, frmMain
    
End Sub

Sub InitDate()
    ''''''''''''''''''''''''''''''''''
    '功能               初使化数据
    ''''''''''''''''''''''''''''''''''
    Dim strTmp As String
    Dim i As Long
     
    '选中默认单位
    strTmp = Trim(zlDatabase.GetPara("单位", glngSys, mlngModule, , Array(fra单位, opt单位(0), opt单位(1), opt单位(2), opt单位(3)), mblnHavePriv))
    Select Case strTmp
    Case "0"
        opt单位(0).Value = True
    Case "1"
        opt单位(1).Value = True
    Case "2"
        opt单位(2).Value = True
    Case "3"
        opt单位(3).Value = True
    End Select
End Sub

Private Function SaveDate() As Boolean
    '------------------------------------------------------------------------------------------------
    '功能       保存数据
    '------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String
    
    For i = 0 To opt单位.Count - 1
        If opt单位(i).Value Then
            strTmp = i
        End If
    Next
    
    Err = 0: On Error GoTo ErrHand:
    Call zlDatabase.SetPara("单位", strTmp, glngSys, mlngModule, IIf(opt单位(0).Enabled = True, True, False))
    SaveDate = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Function

Private Sub cmdHelp_Click()
    '功能:调用帮助
    '修改:lesfeng
    '日期:2010-02-25
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdOK_Click()
    If SaveDate = False Then Exit Sub
    Unload Me
End Sub

Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

