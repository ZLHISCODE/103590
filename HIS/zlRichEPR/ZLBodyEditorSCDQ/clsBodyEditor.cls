VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mfrmCaseTendBody As frmCaseTendBody
Attribute mfrmCaseTendBody.VB_VarHelpID = -1
Public Event zlAfterPrint()

Public Function GetTendBody() As Object
    Set GetTendBody = frmCaseTendBody
End Function

Public Function GetNewTendBody() As Object
    Set GetNewTendBody = New frmCaseTendBody
End Function

Public Sub InitBodyEditor(ByVal lngSys As Long, ByVal cnOracle As ADODB.Connection)
    glngSys = lngSys
    Set gcnOracle = cnOracle
    
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("产品名称"), Default:="中联")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    Call GetUserInfo
End Sub

Public Function GetCaseTendBodyPara() As Object
    Set GetCaseTendBodyPara = frmCaseTendBodyPara
End Function

Private Sub Class_Initialize()
    Set mfrmCaseTendBody = New frmCaseTendBody
End Sub

Private Sub Class_Terminate()
    Unload mfrmCaseTendBody
End Sub

Private Sub mfrmCaseTendBody_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub


