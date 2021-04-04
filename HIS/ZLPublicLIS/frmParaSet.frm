VERSION 5.00
Begin VB.Form frmParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6330
   Icon            =   "frmParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6330
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3780
      TabIndex        =   2
      Top             =   1590
      Width           =   970
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&Q)"
      Height          =   350
      Left            =   5070
      TabIndex        =   3
      Top             =   1590
      Width           =   970
   End
   Begin VB.Frame fraWsk 
      Caption         =   "通讯参数"
      Height          =   1425
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   6075
      Begin VB.CheckBox chkStart 
         Caption         =   "启用通讯"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         TabIndex        =   8
         Top             =   270
         Width           =   1425
      End
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   0
         Top             =   750
         Width           =   1695
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4290
         TabIndex        =   1
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label lblShow 
         Caption         =   "启用通讯之后，在进行业务操作时会向服务端发送刷新请求"
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1950
         TabIndex        =   7
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label lblInfor 
         AutoSize        =   -1  'True
         Caption         =   "服务端IP"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   805
         Width           =   960
      End
      Begin VB.Label lblInfor 
         AutoSize        =   -1  'True
         Caption         =   "通讯端口"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3210
         TabIndex        =   5
         Top             =   805
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK   As Boolean  '点击确定
Private mstrFontName As String
Private mlngFontSize As Long
Private mlngFontColor As Long
Private mblnFontBold As Boolean
Private mblnFontItalic As Boolean
Private mblnFontUnderline As Boolean
Private mblnFontStrikethru As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
          Dim strIP As String
          Dim lngPort As Integer
          Dim objOpt As OptionButton
          
1         On Error GoTo cmdOk_Click_Error

2         mblnOK = True
          '先判断是否为ip地址
3         strIP = Trim(txtIp.Text)
4         If Not IsIP(strIP) Then
5             MsgBox "输入的IP地址不能转为IP！", vbInformation, "提示"
6             mblnOK = False
7             Exit Sub
8         End If
          
9         If IsNumeric(Trim(txtPort.Text)) And Val(Trim(txtPort.Text)) > 0 Then
10            lngPort = Trim(txtPort.Text)
11        Else
12            MsgBox "端口请输入大于0的数字！", vbInformation, "提示"
13            mblnOK = False
14            Exit Sub
15        End If
              
              
          '修改配置文件
16        Call SaveSet(strIP, lngPort, Me.chkStart.Value)
          
          '保存之后立即读取，以保证参数能够立即生效
17        Call InitPara
          
18        Unload Me


19        Exit Sub
cmdOk_Click_Error:
20        Call writeErrLog("zlPublicLIS", "frmParaSet", "执行(cmdOk_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Sub


Private Sub Form_Load()
    Dim strIP As String
    Dim lngPort As Integer
    Dim intProtocol As Integer
    
    '获取参数
    Call InitPara
        
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/20
'功    能:保存配置文件
'入    参:
'           strIP           IP
'           lngPort         端口
'           IntStart        是否启用通讯,0=未启用，1=启用
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SaveSet(ByVal strIP As String, ByVal lngPort As Integer, ByVal intStart As Integer)
          Dim strPara As String
          
1         On Error GoTo SaveSet_Error

2         strPara = intStart & "|" & strIP & "|" & lngPort
3         Call ComSetPara("LIS远程通讯参数", strPara, 2500, 2500)


4         Exit Sub
SaveSet_Error:
5         Call writeErrLog("zlPublicLIS", "frmParaSet", "执行(SaveSet)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
6         Err.Clear
          
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/4/18
'功    能:读取配置文件
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Function InitPara()
          Dim strPara As String
          '通讯参数

1         On Error GoTo InitPara_Error

2         strPara = ComGetPara("LIS远程通讯参数", 2500, 2500, "0|127.0.0.1|8888")
3         Me.chkStart.Value = Split(strPara, "|")(0)
4         Me.txtIp.Text = Split(strPara, "|")(1)
5         Me.txtPort.Text = Split(strPara, "|")(2)
          
6         Exit Function
InitPara_Error:
7         Call writeErrLog("zlPublicLIS", "frmParaSet", "执行(InitPara)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
8         Err.Clear
             
End Function


Private Sub txtIp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPort.SetFocus
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOk.SetFocus
    End If
End Sub
