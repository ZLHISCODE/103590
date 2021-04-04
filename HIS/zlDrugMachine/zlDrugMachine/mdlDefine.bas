Attribute VB_Name = "mdlDefine"
Option Explicit

Public Type TYPE_PARAMS
    输出日志 As Boolean
    详细日志 As Boolean
    保存日志天数 As Integer
End Type

Public Const GSTR_CONFIG_FILE As String = "zlDrugMachine.cfg"
Public Const GLNG_SYS As Long = 100
Public Const GLNG_MODULE As Long = 9010

Public gcnThird As ADODB.Connection

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type

Private Const MAX_IP = 5                                                    'To make a buffer... i dont think you have more than 5 ip on your pc..
Public Type IPINFO
    dwAddr As Long                                                          ' IP address
    dwIndex As Long                                                         ' interface index
    dwMask As Long                                                          ' subnet mask
    dwBCastAddr As Long                                                     ' broadcast address
    dwReasmSize  As Long                                                    ' assembly size
    unused1 As Integer                                                      ' not currently used
    unused2 As Integer                                                      '; not currently used
End Type
Public Type MIB_IPADDRTABLE
    dEntrys As Long                                                         'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO                                               'array of IP address entries
End Type

Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Type TYP_YUYAMA
    MacNO As Integer             '机器号
    BottomLevel As Boolean       '最低层
    SendIDs As String            '发送过的收发ID
End Type
Public gtypYUYAMA As TYP_YUYAMA
