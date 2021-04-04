VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrName As String
Private mbytType As Byte
Private mcnDB As ADODB.Connection
Private mcnWS As Object
Private mstrConnect As String
Private mblnStatus As Boolean
Private mstrURL As String, mstrUser As String, mstrPassword As String

'连接对象名称
Public Property Get Name() As String
    Name = mstrName
End Property
Public Property Let Name(ByVal strName As String)
    mstrName = strName
End Property

'连接类型
Public Property Get LinkType() As Byte
    LinkType = mbytType
End Property
Public Property Let LinkType(ByVal bytType As Byte)
    mbytType = bytType
End Property

'DB连接对象
Public Property Get DBConnect() As ADODB.Connection
    Set DBConnect = mcnDB
End Property
Public Property Let DBConnect(ByVal cnDB As ADODB.Connection)
    Set mcnDB = cnDB
End Property

'WebServices连接对象
Public Property Get WSConnect() As Object
    Set WSConnect = mcnWS
End Property

'其他连接
Public Property Get Connect() As String
    Connect = mstrConnect
End Property
Public Property Let Connect(ByVal strConn As String)
    Const STR_URL = "URL"
    Const STR_USER = "USER"
    Const STR_PWD = "PWD"
    
    Dim arrTmp As Variant
    Dim strKey As String, strValue As String
    Dim i As Integer
    
    mstrConnect = strConn
    mblnStatus = True
    
    On Error GoTo errHandle
    If LinkType = enuLinkType.WEBServices Then
        arrTmp = Split(mstrConnect, GSTR_SEPARAT_CHILD)
        For i = LBound(arrTmp) To UBound(arrTmp)
            strKey = UCase(Left(arrTmp(i), InStr(arrTmp(i), "=") - 1))
            strValue = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
            Select Case strKey
                Case STR_URL
                    '地址
                    mstrURL = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                Case STR_USER
                    '用户
                    mstrUser = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                Case STR_PWD, "PASSWORD"
                    '口令
                    mstrPassword = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
            End Select
        Next
        
        'WebServices
        CreateWebServices mstrURL, mcnWS
        
    End If
    Exit Property
    
errHandle:
    mblnStatus = False
    gstrMessage = Err.Description
End Property

'连接状态
Public Property Get Status() As Boolean
    Status = mblnStatus
End Property

'连接用户
Public Property Get User() As String
    User = mstrUser
End Property

'连接密码
Public Property Get Password() As String
    Password = mstrPassword
End Property

'连接地址
Public Property Get URL() As String
    URL = mstrURL
End Property
