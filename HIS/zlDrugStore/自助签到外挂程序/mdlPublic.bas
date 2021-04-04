Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrSql As String
Public gcnOracle As ADODB.Connection
Public glngStock As Long                    '当前药房ID
Public gstrStockName As String              '当前药房名称
Public glngCardTypeID As Long                   '当前刷卡的类别ID

Const GCONS_COMMEN_BUTTON = "武汉普仁医院 欢迎使用自助签到系统"


