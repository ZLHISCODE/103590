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
    strID As String     '�û�ID
    strName As String   '����
    strSignName As String  'ǩ��
    strUserID      As String '����֤��
    strCertSn    As String   '֤�����
    strCert   As String      '֤������
    strCertDN As String      '֤��DN
    strEncCert As String
    strCertID As String
    strPicCode As String    'BASE64ǩ��ͼƬ����
End Type