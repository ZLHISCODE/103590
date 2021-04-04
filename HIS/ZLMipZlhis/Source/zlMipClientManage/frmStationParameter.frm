VERSION 5.00
Begin VB.Form frmStationParameter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6480
   Icon            =   "frmStationParameter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "启用ZLHIS消息服务"
      Height          =   1125
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   6330
      Begin VB.CheckBox chk 
         Caption         =   "启用消息服务（ZLHIS产品）"
         Height          =   195
         Index           =   0
         Left            =   1065
         TabIndex        =   13
         Top             =   720
         Width           =   3390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "启用或禁用ZLHIS产品的消息服务功能"
         Height          =   180
         Index           =   0
         Left            =   1065
         TabIndex        =   14
         Top             =   315
         Width           =   2970
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   0
         Left            =   195
         Picture         =   "frmStationParameter.frx":6852
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   4095
      TabIndex        =   8
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5265
      TabIndex        =   9
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "连接消息集成平台"
      Height          =   2280
      Left            =   60
      TabIndex        =   10
      Top             =   1335
      Width           =   6330
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1515
         TabIndex        =   1
         Top             =   750
         Width           =   3030
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1110
         Width           =   3030
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1515
         TabIndex        =   5
         Top             =   1485
         Width           =   3030
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   10
         Left            =   1515
         TabIndex        =   7
         Top             =   1860
         Width           =   3030
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "如:6066"
         Height          =   180
         Index           =   2
         Left            =   4665
         TabIndex        =   16
         Top             =   1905
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "如:192.168.2.24"
         Height          =   180
         Index           =   1
         Left            =   4665
         TabIndex        =   15
         Top             =   1530
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户"
         Height          =   180
         Index           =   7
         Left            =   1080
         TabIndex        =   0
         Top             =   795
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         Height          =   180
         Index           =   8
         Left            =   1080
         TabIndex        =   2
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "地址"
         Height          =   180
         Index           =   9
         Left            =   1080
         TabIndex        =   4
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "端口"
         Height          =   180
         Index           =   10
         Left            =   1080
         TabIndex        =   6
         Top             =   1905
         Width           =   360
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   225
         Picture         =   "frmStationParameter.frx":81D4
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "设置连接消息集成平台所需要的用户、IP地址及端口号（默认）"
         Height          =   180
         Index           =   11
         Left            =   1065
         TabIndex        =   11
         Top             =   375
         Width           =   5040
      End
   End
End
Attribute VB_Name = "frmStationParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'变量定义
Private mblnDataChanged As Boolean
Private mstrTitle As String

'######################################################################################################################
'接口方法

Public Function ShowConfigDialog(ByVal frmParent As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strPara As String
    Dim varPara As Variant
    
    chk(0).Value = Val(gclsBusiness.ParameterRead(16))
    
    strPara = gclsBusiness.ParameterRead(18)
    If strPara <> "" Then
        varPara = Split(strPara, ";")
        txt(7).Text = varPara(0)
        txt(8).Text = varPara(1)
        txt(9).Text = varPara(2)
        txt(10).Text = Val(varPara(3))
    End If

    mblnDataChanged = False
    
    Me.Show 1, frmParent
        
    ShowConfigDialog = mblnDataChanged
    
End Function

Private Sub chk_Click(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsPara As ADODB.Recordset
    Dim strPara As String
    Dim blnRet As Boolean
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "参数号", 16)
    Call zlCommFun.SetParameter(rsPara, "参数名", "启用消息服务")
    Call zlCommFun.SetParameter(rsPara, "参数值", chk(0).Value)
    blnRet = gclsBusiness.ParameterEdit("UPDATE", rsPara)
    
    If blnRet Then
        strPara = txt(7).Text & ";" & txt(8).Text & ";" & txt(9).Text & ";" & txt(10).Text
        Call zlCommFun.SetParameter(rsPara, "参数号", 18)
        Call zlCommFun.SetParameter(rsPara, "参数名", "连接消息集成平台参数")
        Call zlCommFun.SetParameter(rsPara, "参数值", strPara)
        blnRet = gclsBusiness.ParameterEdit("UPDATE", rsPara)
    End If
    
    If blnRet Then
        mblnDataChanged = False
        Unload Me
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

