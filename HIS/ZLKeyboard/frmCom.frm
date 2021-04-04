VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmCom 
   Caption         =   "frmCom"
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3090
   Begin MSCommLib.MSComm msCommKeyBoard 
      Left            =   210
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnInitCom As Boolean   '是否初始化Com接口成功
Private WithEvents mTxtPass As TextBox
Attribute mTxtPass.VB_VarHelpID = -1

Private Sub Form_Load()
     Call InitComProperty
     mblnInitCom = InitComm
End Sub

Private Sub msCommKeyBoard_OnComm()
    Dim strKeyChar As String
    '接收数据
    If mTxtPass Is Nothing Then Exit Sub
    strKeyChar = msCommKeyBoard.Input
    If strKeyChar = "" Then Exit Sub
    '清除字符
    If Asc(strKeyChar) = 8 Then mTxtPass.Text = "": Exit Sub
    If Asc(strKeyChar) >= Asc("0") And Asc(strKeyChar) <= Asc("9") Then
        mTxtPass.Text = mTxtPass.Text & strKeyChar
        mTxtPass.SelStart = Len(mTxtPass.Text)
    End If
    If Asc(strKeyChar) = 13 Then PressKey vbKeyReturn
End Sub
Private Sub mTxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
        Exit Sub
    End If
    KeyAscii = 0
End Sub
Public Function OpenPassKeyoardInput(ByVal frmMain As Object, _
    ByVal objPassCtl As Object, Optional blnAffirmPass As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码建盘输入
    '入参:frmMain-调用的主窗体
    '       objPassCtl-输入的密码控件
    '       blnAffirmPass-False:请输入密码;true:请输入确认密码
    '出参:
    '返回:打开成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:30:54
    '--------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mblnInitCom = False Then Exit Function
    With msCommKeyBoard
        If .PortOpen = False Then .PortOpen = True
        If blnAffirmPass Then
            Call HexSend("81H")  '你好,请再次输入密码
        Else
            Call HexSend("82H") '你好,请输入密码
        End If
    End With
    Set mTxtPass = objPassCtl
    OpenPassKeyoardInput = True
    Exit Function
errHandle:
End Function
Public Function ColsePassKeyoardInput(ByVal frmMain As Object, ByVal objPassCtl As Object) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭密码建盘输入
    '入参:frmMain-调用的主窗体
    '       objPassCtl-输入的密码控件
    '出参:
    '返回:关闭成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-28 16:07:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mblnInitCom = False Then Exit Function
    With msCommKeyBoard
        If .PortOpen = True Then .PortOpen = False
    End With
    Set mTxtPass = Nothing
    ColsePassKeyoardInput = True
    Exit Function
errHandle:
End Function
Private Function InitComm() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开端口
    '编制:刘兴洪
    '日期:2011-07-28 14:35:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSet As String
    On Error GoTo errHandle
     
     strSet = g_Com_Property.lng波特率
     strSet = strSet & "," & Switch(g_Com_Property.str奇偶检验位 = "无", "n", g_Com_Property.str奇偶检验位 = "奇", "o", g_Com_Property.str奇偶检验位 = "偶", "e", g_Com_Property.str奇偶检验位 = "空格", " ", True, "n")
     strSet = strSet & "," & g_Com_Property.int数据位
     strSet = strSet & "," & g_Com_Property.int停止位
     With msCommKeyBoard
        .CommPort = g_Com_Property.int端口号
        .Settings = strSet
        .InputLen = 1      '返回接收缓冲区中等待的字符数,该属性在设计时无效
        .RThreshold = 1 '接收到1个字节数据就立即触发OnComm()事件
     End With
    InitComm = True
    Exit Function
errHandle:
End Function
Private Sub HexSend(ByVal strSend As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按照十六进制输出数据
    '编制:刘兴洪
    '日期:2011-07-28 15:25:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intOutPutLen As Integer    '发送数据的长度
    Dim strOutdata As String          '发送数据暂存
    Dim bytSendArr() As Byte       '发送数组
    Dim strTempSave As String '数据暂存
    Dim intCount As Integer
    Dim i As Integer
    Err = 0: On Error Resume Next
    strOutdata = UCase(Replace(strSend, " ", ""))         '去掉空格，变成大写字母
    intOutPutLen = Len(strOutdata)            '数据的长度
    For i = 0 To intOutPutLen
        strTempSave = Mid(strOutdata, i + 1, 1)          '取一位数据
        If (Asc(strTempSave) >= Asc("0") And Asc(strTempSave) <= Asc("9")) _
        Or (Asc(strTempSave) >= 65 And Asc(strTempSave) <= 70) Then
            intCount = intCount + 1
        Else
            Exit For
        End If
    Next
    If intCount Mod 2 <> 0 Then            '判断十六进制数据是否为双数
        intCount = intCount - 1           '不是双数则减去1
    End If
    strOutdata = Left(strOutdata, intCount)       '取出有效的十六进制数据
    ReDim bytSendArr(intCount / 2 - 1)        '重新定义数组长度
    For i = 0 To intCount / 2 - 1
        bytSendArr(i) = Val("&H" + Mid(strOutdata, i * 2 + 1, 2)) '取出数据转换成十六进制并存放到数组中
    Next
     msCommKeyBoard.Output = bytSendArr          '发送数据
End Sub


