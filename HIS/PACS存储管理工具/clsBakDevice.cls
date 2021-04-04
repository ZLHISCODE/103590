VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBakDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strDevID As String               '设备号
Public strDevName As String             '设备名称
Public strIP As String                  '设备IP，网络地址
Public strUser As String                '用户名
Public strPasswd As String              '密码
Public lngType As Long                  '设备类型,1是存储设备
Public strVirtualPath As String         'FTP虚拟目录
Public strLocalPath As String           '本机目录
Public lngStatus As Long                '设备状态，0-在线设备，1-离线或备份设备

