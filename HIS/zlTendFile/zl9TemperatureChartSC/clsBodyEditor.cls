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

Private mfrmPrintSet As Object
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
    Set mfrmPrintSet = New frmPrintSet
End Sub

Private Sub Class_Terminate()
    If Not (mfrmCaseTendBody Is Nothing) Then Unload mfrmCaseTendBody
    If Not (mfrmPrintSet Is Nothing) Then Unload mfrmPrintSet
    Set mfrmCaseTendBody = Nothing
    Set mfrmPrintSet = Nothing
End Sub

Private Sub mfrmCaseTendBody_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

Public Sub zlPrintSet(ByVal frmParent As Object, Optional ByVal bytMode As Byte = 1)
'-----------------------------------------------------------------------------------------------------------------------------
'功能:体温单打印设置
'-----------------------------------------------------------------------------------------------------------------------------
    Call mfrmPrintSet.ShowMe(frmParent, bytMode)
End Sub

Public Sub BodyMutilEditor(ByVal objParent As Object, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0)
'调用体温单批量录入窗体
    Dim frmBodyFileMutil As New frmTendBodyMutilEditor

    Call frmBodyFileMutil.ShowMe(objParent, lngDeptID, strPrivs)
End Sub

Public Function zlCurvePrint(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String, Optional strPage As String, Optional strParam As String = "") As Boolean
    '---------------------------------------------------------------------------
    '功能:体温单打印
    'strPage intPrintRange=1 是需提供打印开始页号页结束页号 格式为 0;1 表示开始页号1 结束页号 2
    'strParam 新窗体调用时要传入 文件ID;病人ID;主页ID;科室ID;婴儿
    '---------------------------------------------------------------------------
    zlCurvePrint = mfrmCaseTendBody.BodyEdit.PrintState(intPrintRange, blnPrint, lngBeginY, intPageNo, strPrintDevice, strPage, strParam)
End Function
