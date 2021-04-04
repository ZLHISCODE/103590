Attribute VB_Name = "Module1"
Option Explicit

'初始化动态库函数
'   serverIp:服务器IP地址
'   serverPort:服务器端口号
'   hwnd:父窗体句柄
'   callFun:回调函数指针,设置为空时消息传递
'   obj:对象指针
'返回:true(成功);false(失败)
Private Declare Function CEC_Initialize Lib "E:\CecDeviceToHis.dll" (ByVal serverIp As String, ByVal serverPort As Long, ByVal hwnd As Long, ByVal callFun As Long, ByVal object As Long) As Boolean

Private Function CEC_HisSetDataToCec(ByVal nMonitorNo As Long, ByVal nCmd As Long, ByVal obj As Object) As Boolean
    Form1.Text1.Text = nMonitorNo
    CEC_HisSetDataToCec = True
End Function

Public Function Initialize(ByVal hwnd As Long) As Boolean
    CEC_Initialize "192.168.1.200", 5000, hwnd, AddressOf CEC_HisSetDataToCec, 0
    Initialize = True
End Function
