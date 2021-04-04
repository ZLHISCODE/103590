VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsESign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const mstrSysName = "签名"

Private mobjESign As Object
Private mstrServerName As String

Public IsConnection As Boolean

'用户信息
Private mstrUserName As String
Private mstrUserPwd As String

'数据库实例名
Public Property Get ServerName() As String
    ServerName = mstrServerName
End Property

Public Property Let ServerName(ByVal vNewValue As String)
    mstrServerName = vNewValue
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'指定用户的用户名
Public Property Get UserName() As String
    UserName = mstrUserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    mstrUserName = vNewValue
End Property

'指定用户的密码
Public Property Get UserPwd() As String
    UserPwd = mstrUserPwd
End Property

Public Property Let UserPwd(ByVal vNewValue As String)
    mstrUserPwd = vNewValue
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'初始化数字签名部件
Public Function Initialize(ByVal lngSys As Long) As Boolean
On Error GoTo ErrorHand

    If mobjESign Is Nothing Then
        mobjESign = CreateObject("zl9ESign.clsESign")

        If mobjESign Is Nothing Then
            MsgBox "部件zl9ESign.dll创建失败，请联系管理员!", vbExclamation, mstrSysName
            Exit Function
        End If
    End If

    '初始化数字签名部件
    If Not mobjESign Is Nothing Then
        If mobjESign.Initialize(gcnOracle, lngSys) = False Then
            MsgBox "数字证书初始化失败，请使用正确的数字证书签名。", vbInformation + vbOKOnly, mstrSysName
            Exit Function
        End If
    End If

    Initialize = True
    Exit Function
ErrorHand:
    Initialize = False
    Err = 0
End Function

'检查数字证书跟登陆用户是否一致
Public Function CheckCertificate(ByVal strDbUser As String) As Boolean
    CheckCertificate = False

On Error GoTo ErrorHand

    If Not mobjESign Is Nothing Then
        CheckCertificate = mobjESign.CheckCertificate(strDbUser)
    End If
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'检查证书是否停用
Public Function CertificateStoped(ByVal strDBName As String) As Boolean
    CertificateStoped = False

On Error GoTo ErrorHand

    If Not mobjESign Is Nothing Then
        CertificateStoped = mobjESign.CertificateStoped(strDBName)
    End If
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'使用数字签名对源文进行签名加密
'返回：签名信息strSign-加密后的源文；lngCertID-签名使用的证书记录ID；str时间戳 --签名之后的时间戳
Public Function Signature(ByVal strSource As String, _
                          ByVal strDbUser As String, _
                          lngCertID As Long, _
                          Optional strTimeStamp As String, _
                          Optional objSignPic As StdPicture, _
                          Optional strTimeStampCode As String) As String
    Dim strSign As String
    
On Error GoTo ErrorHand
    
    If Not mobjESign Is Nothing Then
        strSign = mobjESign.Signature(strSource, strDbUser, lngCertID, strTimeStamp, objSignPic, strTimeStampCode)
        If strSign = "" Then MsgBox "验证失败！请重新输入验证信息！", vbInformation + vbOKOnly, mstrSysName
    End If

    Signature = strSign
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

'验证签名,不需要插入USB-Key
Public Function VerifySignature(ByVal strSource As String, ByVal lngSignID As Long, ByVal intSignType As Integer) As Boolean
    VerifySignature = False
    
On Error GoTo ErrorHand
    
    If Not mobjESign Is Nothing Then
        VerifySignature = mobjESign.VerifySignature(strSource, lngSignID, intSignType)
    End If
    
    Exit Function
ErrorHand:
    MsgBox Err.Description, vbExclamation, mstrSysName
    Err = 0
End Function

Private Sub Class_Initialize()
    mstrServerName = ""
    mstrUserName = ""
    mstrUserPwd = ""
    IsConnection = False
End Sub

Private Sub Class_Terminate()
    If Not mobjESign Is Nothing Then mobjESign = Nothing
End Sub

Public Function OraDataOpen() As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    Dim strUserPwd As String

    On Error GoTo ErrorHand
    
    IsConnection = False
    OraDataOpen = False
    
    '转换密码
    If Not (UCase(mstrUserName) = "SYS" Or UCase(mstrUserName) = "SYSTEM") Then
        strUserPwd = TranPasswd(mstrUserPwd)
    Else
        strUserPwd = mstrUserPwd
    End If

    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close

        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & ServerName, UserName, strUserPwd

        If Err <> 0 Then
            MsgBox "帐户验证失败！可能是帐户名或密码不正确", vbExclamation, mstrSysName
            Exit Function
        End If
    End With
    
    IsConnection = True
    OraDataOpen = True
    Exit Function

ErrorHand:
    MsgBox "帐户验证失败！可能是帐户名或密码不正确", vbExclamation, mstrSysName
    Err = 0
End Function
