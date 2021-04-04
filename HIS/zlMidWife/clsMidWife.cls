VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMidWife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mobjFrm As frmBrowser

Public Function GetForm() As Object
'功能：返回助产信息的窗体对象
    Set GetForm = mobjFrm
End Function

Public Sub RefreshForm(ByVal lngPatiID As Long, ByVal lngPageID As Long)
'功能：刷新指定病人的助产信息界面
'参数：lngPatiID-病人id,lngPageID-主页ID
    Call mobjFrm.RefreshForm(lngPatiID, lngPageID)
End Sub

Public Function Initialize(ByRef cnOracle As ADODB.Connection, Optional ByVal lngPid As Long, Optional ByVal lngTid As Long) As Boolean
'功能：接口初始化
'参数：cnOracle=登录ZLHIS导航台的活动连接
    Dim rsTemp As ADODB.Recordset, strSql As String, strTemp As String
    Dim varPar As Variant, strIP As String, strPubUserName As String, strPwd As String, strTime As String
    
    On Error GoTo errHand
    
    If gstrURLLogin = "" Then
        Set gcnOracle = cnOracle
        
        strSql = "Select user From dual"
        Set rsTemp = gcnOracle.Execute(strSql)
                
        If rsTemp.RecordCount > 0 Then gstrUser = rsTemp!user
        'gstrUser = "ZLHIS"
            
        strTemp = GetSysPar(325, 100)
        'varPar = Split("192.168.0.60:8080/ords+127", "+")
        varPar = Split(strTemp, "+")
        If UBound(varPar) < 1 Or varPar(0) = "" Then
            MsgBox "没有对助产士系统进行初始化配置，请在系统公共参数中设置助产士工作站的相关参数！", vbInformation, "助产信息"
            Exit Function
        End If
        strIP = "http://" & varPar(0) & "/f?p=" & varPar(1)
        
        strPubUserName = "zlsol_pub"
        strPwd = Replace(strTemp, varPar(0) & "+" & varPar(1) & "+", "")
        strPwd = zlStr.Sm4DecryptEcb(strPwd)
        strPwd = GetEncrypt(strPwd)
        strTime = GetEncrypt(Replace(Now, " ", "-"))
        
        gstrURL = strIP & ":14:[SESSION_TOKEN]::NO:RP:P14_PID,P14_TID:"
        gstrURLLogin = strIP & ":LOGIN_DESKTOP" & "&username=" & strPubUserName & "&pwd=" & strPwd & "&key=" & strTime & "&pid=" & lngPid & "&tid=" & lngTid & "&hisname=" & gstrUser
        glngPatiID = lngPid
        glngPageID = lngTid
    End If
    
    mobjFrm.InitLoad
    
    Initialize = True
    Exit Function
    
errHand:
    MsgBox Err.Description, vbExclamation, "助产信息初始化"
    Initialize = False
End Function

Private Sub Class_Initialize()
    '即使类的Instancing属性设置为5-MultiUse(多个对象共享一个类的实例),而不是3-SingleUse（每个对象对应一个类的实例），也会调用
    '相当于是对象的初始化，而不一定是类的实例初始化
    glngInstance = glngInstance + 1
    Set mobjFrm = New frmBrowser    '支持医生站、护士站分别同时开启界面
End Sub

Private Sub Class_Terminate()
    glngInstance = glngInstance - 1
    Set mobjFrm = Nothing
    
    If glngInstance = 0 Then
        Set gcnOracle = Nothing
    End If
End Sub
