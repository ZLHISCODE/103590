Attribute VB_Name = "mdlPacs"
Option Explicit

''''''''插件说明''''''''''''''''''''''
'1、CallPACSView.dll和CallPACSView.lib两个文件是上海岱嘉提供给HIS医生站调用的接口文件。
'2、这两个接口文件需要放到Windows/System32目录下，不需要注册。
'3、本接口中直接写了岱嘉在大连医大一院的RIS服务器IP：192.9.200.6， WEB服务器IP：192.9.200.9，用户名，密码等，因此只能在医院内网环境中运行。
''''''''''''''''''''''''''''''''''''''''''''''''''

'初始化pacs连接
Public Declare Function InitPACSConnection Lib "CallPACSView.dll" ( _
    ByVal strRisIp As String, _
    ByVal strRisUser As String, _
    ByVal strRisPwd As String, _
    ByVal strRisDbName As String, _
    ByVal strPacsIp As String, _
    ByVal strPacsUser As String, _
    ByVal strPacsPwd As String, _
    ByVal strPacsDbName As String _
    ) As String


'调阅pacs影像
Public Declare Function CallPACSView Lib "CallPACSView.dll" ( _
    ByVal strAdviceId As String, _
    ByVal strWebIp As String, _
    ByVal strWebUser As String, _
    ByVal strWebPwd As String, _
    ByVal blnIsOpenImage As Boolean _
    ) As String

Public Const gstrFunc_PACS影像调阅 = "PACS影像调阅"

Private Const 岱嘉RIS服务器IP = "192.9.200.6"
Private Const 岱嘉RIS用户名 = "zlhis"
Private Const 岱嘉RIS密码 = "his"
Private Const 岱嘉RIS数据库名 = "UniRISCDB"
'RIS服务器超机用户名 sa,密码 ats

Private Const 岱嘉PACS服务器IP = "192.9.200.6"
Private Const 岱嘉PACS用户名 = "zlhis"
Private Const 岱嘉PACS密码 = "his"
Private Const 岱嘉PACS数据库名 = "DICOMDB"
'PACS服务器超机用户名 sa,密码 ats

Private Const 岱嘉WEB服务器IP = "192.9.200.9"
Private Const 岱嘉WEB用户名 = "user"
Private Const 岱嘉WEB密码 = "1"


Private blnInitPacsConnection As Boolean        '是否需要初始化PACS连接

Public Function InitPacs() As Boolean
'初始化岱嘉的PACS数据库连接

    Dim strErr As String
    
    On Error GoTo err
    
    InitPacs = False
    
    strErr = InitPACSConnection(岱嘉RIS服务器IP, 岱嘉RIS用户名, 岱嘉RIS密码, 岱嘉RIS数据库名, _
                                岱嘉PACS服务器IP, 岱嘉PACS用户名, 岱嘉PACS密码, 岱嘉PACS数据库名)


    If strErr <> "成功" Then
        MsgBox strErr, vbOKOnly, "PACS影像接口"
        Exit Function
    End If
    
    InitPacs = True
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACS影像接口"
    err.Clear
End Function

Public Function ShowPacsViewer(ByVal varKeyId As Variant) As Boolean
'调用岱嘉的CallPACSView函数，从IE显示WEB版本的PACS图像浏览器
    Dim strErr As String
    
    On Error GoTo err
    
    ShowPacsViewer = False
    
    '先初始化
    '只针对门诊医生站和住院医生站才初始化PACS调阅图像的插件
    If blnInitPacsConnection = False Then
        blnInitPacsConnection = InitPacs
    End If
        
    If blnInitPacsConnection = True Then
        strErr = CallPACSView(CStr(varKeyId), 岱嘉WEB服务器IP, 岱嘉WEB用户名, 岱嘉WEB密码, False)
        If strErr <> "成功" Then
            MsgBox strErr, vbOKOnly, "PACS影像接口"
            Exit Function
        End If
    
        ShowPacsViewer = True
    End If
    
    Exit Function
err:
    MsgBox err.Description, vbOKOnly, "PACS影像接口"
    err.Clear
End Function
