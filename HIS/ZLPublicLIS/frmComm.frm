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
   StartUpPosition =   3  '����ȱʡ
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
'��    ��:������
'����ʱ��:2017/4/24
'ģ�鹦��:���ڷ���winsock�ؼ�����������ʱ����ʾ���档
'winsock��VB���в���ͬ���̣�vb����ȴ�winsock����״̬��ʹ��timer�ؼ�����ѭ�����winsock ������״̬
'---------------------------------------------------------------------------------------

Option Explicit

Private mstrSend As String  'Ҫ���͵���Ϣ

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:��zlLisMessage������Ϣ
'��    ��:
'           strSend     Ҫ���͵���Ϣ����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub funSendMessage(ByVal strSend As String)

1         On Error GoTo funSendMessage_Error

2         mstrSend = strSend
3         If wskSend.State <> sckClosed Then wskSend.Close '���winsock���ǶϿ�״̬�����ȶϿ�winsock
4         wskSend.Connect gstrIP, glngPort   '���ӷ����  �����IP �˿�
5         Me.timWsk.Interval = 200    'ÿ200������һ��winsock������״̬

6         Exit Sub
funSendMessage_Error:
7         Call writeErrLog("zlPublicLIS", "frmComm", "ִ��(funSendMessage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
8         Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/4/24
'��    ��:ÿ200������һ��winsock������״̬
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub timWsk_Timer()
          'һ��winsock�����ӣ��ͷ�����Ϣ������ֹͣ��winsock״̬�ļ��

1         On Error GoTo timWsk_Timer_Error

2         If wskSend.State = sckConnected Then
3             wskSend.SendData mstrSend
4             DoEvents 'ת��CUPȨ��
5             Me.timWsk.Interval = 0
              
              '��������Ϣ֮�������Ͽ�����
6             If wskSend.State <> sckClosed Then
7                 wskSend.Close
8             End If
              
9         End If


10        Exit Sub
timWsk_Timer_Error:
11        Call writeErrLog("zlPublicLIS", "frmComm", "ִ��(timWsk_Timer)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
12        Err.Clear

End Sub

