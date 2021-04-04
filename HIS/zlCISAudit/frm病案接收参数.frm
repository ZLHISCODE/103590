VERSION 5.00
Begin VB.Form frm病案接收参数 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病案接收参数"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frm病案接收参数.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk打印 
      Caption         =   "接收后打印接收清单(B)"
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1650
      TabIndex        =   3
      Top             =   2175
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2835
      TabIndex        =   2
      Top             =   2175
      Width           =   1100
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   1905
      Width           =   4245
   End
   Begin VB.Frame fraSplit 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   4245
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frm病案接收参数.frx":030A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "  请选择控制相关的病案接收管理数据的相关参数．"
      Height          =   405
      Left            =   765
      TabIndex        =   4
      Top             =   135
      Width           =   3105
   End
End
Attribute VB_Name = "frm病案接收参数"
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
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/19
    '------------------------------------------------------------------------------------------
    On Error GoTo ErrHand
    If mlngModule = 201 Then
        
    Else
        Call zlDatabase.SetPara("打印接收清单", IIf(chk打印.Value = 1, "1", "0"), glngSys, mlngModule)
    End If
    SaveSet = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOK_Click()
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Public Sub 参数设置(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '------------------------------------------------------------------------------------
    '功能:参数设置入口
    '参数:
    '返回:
    '编制:刘兴宏
    '修改:2007/12/21
    '------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnHavePriv = IsHavePrivs(mstrPrivs, "参数设置")
    
    If mlngModule = 201 Then
        chk打印.Visible = False
    Else
        chk打印.Value = IIf(Val(zlDatabase.GetPara("打印接收清单", glngSys, mlngModule, , Array(chk打印), mblnHavePriv)) = 1, 1, 0)
        chk打印.Visible = True
    End If
    
    frm病案接收参数.Show 1, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    If gbln共享 = False Or gSystemPara.bln联机 = False Then
'        ChkAuto.Visible = False
'        cmdOK.Top = chk病因.Top + chk病因.Height + 100
'        cmdCancel.Top = cmdOK.Top
'        Me.Height = cmdOK.Top + cmdOK.Height + 600
'    End If
'
'End Sub


