VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmComm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timWsk 
      Left            =   1560
      Top             =   1650
   End
   Begin MSWinsockLib.Winsock wskSend 
      Left            =   1020
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'创    建:蔡青松
'创建时间:2017/4/24
'模块功能:用于放置winsock控件，程序运行时不显示界面。
'winsock与VB运行不是同进程，vb不会等待winsock返回状态，使用timer控件用来循环检测winsock 的链接状态
'---------------------------------------------------------------------------------------

Option Explicit

Private mstrSend As String  '要发送的消息

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:向zlLisMessage发送消息
'入    参:
'           strSend     要发送的消息内容
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub funSendMessage(ByVal strSend As String)

1         On Error GoTo funSendMessage_Error

2         mstrSend = strSend
3         If wskSend.State <> sckClosed Then wskSend.Close '如果winsock不是断开状态，则先断开winsock
4         wskSend.Connect gstrIP, glngPort   '链接服务端  服务端IP 端口
5         Me.timWsk.Interval = 200    '每200毫秒检测一次winsock的链接状态

6         Exit Sub
funSendMessage_Error:
7         Call writeErrLog("zlPublicLIS", "frmComm", "执行(funSendMessage)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
8         Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/24
'功    能:每200毫秒检测一次winsock的链接状态
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub timWsk_Timer()
          '一旦winsock已连接，就发送消息，并且停止对winsock状态的检测

1         On Error GoTo timWsk_Timer_Error

2         If wskSend.State = sckConnected Then
3             wskSend.SendData mstrSend
4             DoEvents '转让CUP权限
5             Me.timWsk.Interval = 0
              
              '发送完消息之后主动断开连接
6             If wskSend.State <> sckClosed Then
7                 wskSend.Close
8             End If
              
9         End If


10        Exit Sub
timWsk_Timer_Error:
11        Call writeErrLog("zlPublicLIS", "frmComm", "执行(timWsk_Timer)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
12        Err.Clear

End Sub

