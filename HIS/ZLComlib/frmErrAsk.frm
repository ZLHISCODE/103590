VERSION 5.00
Begin VB.Form frmErrAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提示"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00E1E1FF&
      Caption         =   "结束(&E)"
      Height          =   350
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdCopyScreen 
      Caption         =   "截图(&S)"
      Height          =   350
      Left            =   3195
      TabIndex        =   7
      Top             =   1380
      Width           =   900
   End
   Begin VB.TextBox txtHelp 
      Height          =   1635
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmErrAsk.frx":0000
      Top             =   1845
      Width           =   4275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2280
      TabIndex        =   5
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "重试(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1365
      TabIndex        =   4
      Top             =   1380
      Width           =   900
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3390
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblAsk 
      AutoSize        =   -1  'True
      Caption         =   "再试一次吗？"
      Height          =   180
      Left            =   975
      TabIndex        =   3
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label lblNote 
      Caption         =   "    可能是其他用户的独占或重新安装了操作系统带来的错误，排除独占使用因素仍不能运行，则需部分重装本系统。"
      Height          =   585
      Left            =   975
      TabIndex        =   2
      Top             =   360
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "说明："
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "错误编号："
      Height          =   180
      Left            =   1965
      TabIndex        =   0
      Top             =   150
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmErrAsk.frx":0065
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mtmrConnect  As clsTimer '自动重连Timer
Attribute mtmrConnect.VB_VarHelpID = -1
Private mbytReturn As Byte

Private Sub cmdCancel_Click()
    mbytReturn = 0
    Unload Me
End Sub

Private Sub cmdCopyScreen_Click()
    Call SaveScreen(txtHelp.Text, picS)
End Sub

Private Sub cmdEnd_Click()
    If cmdRetry.Caption = "重连(&S)" Then
        TerminateProcess GetCurrentProcess, 0
    Else
       If CheckAdoConnction(False) = False Then
         cmdCancel_Click
       End If
    End If
End Sub

Private Sub cmdRetry_Click()
    mbytReturn = 1
    Unload Me
End Sub

Public Function ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String, ByVal blnConnect As Boolean) As Byte
'功能：显示错误提示窗口，可以选择重试
'参数：lngErrNum   错误编号
'      strNote     错误内容
'      strErrInfo  详细的错误信息
'返回：下一步操作的批示。1-再试；0-取消
    mbytReturn = 0
    
    lblNumber.Caption = "序号：" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo
    
    If blnConnect Then
        cmdRetry.Caption = "重连(&S)"
        cmdRetry.Visible = True
        cmdEnd.Visible = True
        txtHelp.Text = "与ORACLE的网络连接已经断开," & vbNewLine & "可以在网络恢复后手动进行重新连接,请点(重连)!"
        
        
        '开启自动重连功能,默认10秒重新连接
        Set mtmrConnect = New clsTimer
        mtmrConnect.Interval = 10000
    Else
        If lngErrNum = -2147467259 _
            And (InStr(strErrInfo, "E_FAIL") > 0 _
                Or InStr(UCase(strErrInfo), "UNKNOW") > 0 _
                Or strErrInfo = "未指定的错误") Then
            '数据提供程序或其他服务返回一个 E_FAIL 状态
            cmdEnd.Visible = True
        ElseIf lngErrNum = -2147217900 _
            And (InStr(strErrInfo, "ORA-00028") > 0 _
                Or InStr(strErrInfo, "ORA-01012") > 0 _
                Or InStr(strErrInfo, "ORA-03113") > 0) Then
            'ORA-00028: 您的会话己被终止
            'ORA-01012: 没有登录
            'ORA-03113: 通信通道的文件结束
            cmdEnd.Visible = True
        End If
    End If

    frmErrAsk.Show vbModal
    If blnConnect Then
        cmdRetry.Caption = "重试(&O)"
    End If
    ShowEdit = mbytReturn
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        mbytReturn = 0
        Unload Me
    End If
End Sub

Private Sub mtmrConnect_ThatTime()
    If CheckAdoConnction(False) = False Then
      If ObjPtr(mtmrConnect) > 0 Then
        mtmrConnect.Interval = 0
      End If
      cmdCancel_Click
    End If
End Sub

