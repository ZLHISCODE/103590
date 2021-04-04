VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPubDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type USER_INFO
    strID As String     '用户ID
    strName As String   '姓名
    strSignName As String  '签名
    strUserID      As String '身份证号
    strCertSn    As String   '证书序号
    strCert   As String      '证书内容
    strCertDN As String      '证书DN
    strEncCert As String
    strCertID As String
    strPicCode As String    'BASE64签章图片编码
End Type
