VERSION 5.00
Begin VB.Form frmReceiveLogPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   Icon            =   "frmReceiveLogPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Top             =   2220
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5250
      TabIndex        =   4
      Top             =   2220
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "自动清除日志"
      Height          =   1905
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   6330
      Begin VB.ComboBox cboPeiord 
         Height          =   300
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间范围"
         Height          =   180
         Index           =   9
         Left            =   1080
         TabIndex        =   3
         Top             =   1140
         Width           =   720
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   225
         Picture         =   "frmReceiveLogPara.frx":000C
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "设置自动清除指定时间范围之前的日志，在进入日志查阅以及发送消息时自动执行。"
         Height          =   465
         Index           =   11
         Left            =   1065
         TabIndex        =   2
         Top             =   375
         Width           =   5130
      End
   End
End
Attribute VB_Name = "frmReceiveLogPara"
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
    
        
    With cboPeiord
        .Clear
        .AddItem "三天前"
        .AddItem "半月前"
        .AddItem "一月前"
        .AddItem "二月前"
        .AddItem "三月前"
        .AddItem "半年前"
        .AddItem "自定义"
    End With
    If cboPeiord.ListCount > 0 And cboPeiord.ListIndex = -1 Then cboPeiord.ListIndex = 0
    
    strPara = gclsBusiness.ParameterRead(11)
    If strPara <> "" Then
'        varPara = Split(strPara, ";")
'        txt(7).Text = varPara(0)
'        txt(8).Text = varPara(1)
'        txt(9).Text = varPara(2)
'        txt(10).Text = Val(varPara(3))
    End If

    mblnDataChanged = False
    
    Me.Show 1, frmParent
        
    ShowConfigDialog = mblnDataChanged
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsPara As ADODB.Recordset
    Dim strPara As String
    
    On Error GoTo errHand
    
'    strPara = txt(7).Text & ";" & txt(8).Text & ";" & txt(9).Text & ";" & txt(10).Text
    
    Set rsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(rsPara, "参数号", 11)
    Call zlCommFun.SetParameter(rsPara, "参数名", "连接消息服务平台参数")
    Call zlCommFun.SetParameter(rsPara, "参数值", strPara)
    
    If gclsBusiness.ParameterEdit("UPDATE", rsPara) Then
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



