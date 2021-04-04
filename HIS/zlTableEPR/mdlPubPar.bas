Attribute VB_Name = "mdlPubPar"
Option Explicit
Public gcnOracle As ADODB.Connection      '全局连接
Public gstrSQL As String                  '全局共用
Public gstrDbOwner As String              '数据库拥有者
Public glngSys As Long                    '系统编号
Public gstrProductName As String          '程序名称
Public gstrSysName As String              '系统名称
Public gstrAviPath As String              'AVI路径
Public gstrVersion As String              '版本
Public gstrMatch As String                '匹配模式
Public gobjFSO As New Scripting.FileSystemObject    'FSO对象
Public gbytEsign As Byte              '是否启用电子签名 0-密码；1－数字                '
Public gAllFont As Collection
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO
