Attribute VB_Name = "mdlKeyboard"
Option Explicit
Public gobjCom As MSComm
Public Type Ty_Com_Property
    int端口号 As Integer   '端口号
    lng波特率 As Long
    str奇偶检验位 As String
    int停止位 As Integer
    int数据位 As Integer
End Type
Public g_Com_Property As Ty_Com_Property
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public gblnStartKeyboard As Boolean '是否启用密码键盘

Public Sub InitComProperty()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化属性
    '编制:刘兴洪
    '日期:2011-07-28 14:46:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With g_Com_Property
        .int端口号 = Val(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "端口", 0)) + 1
        .int数据位 = Val(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "数据位", "6"))
        .int停止位 = Val(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "停止位", "1"))
        .lng波特率 = Val(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "波特率", "9600"))
        .str奇偶检验位 = Trim(GetSetting("ZLSOFT", "公共模块\zlKeyboard", "奇偶较验位", "无"))
    End With
End Sub

Public Sub PressKey(bytKey As Byte)
    '功能：向键盘发送一个键,类似SendKey
    '参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub
