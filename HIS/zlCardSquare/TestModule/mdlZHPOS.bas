Attribute VB_Name = "mdlZHPOS"
Option Explicit
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例

