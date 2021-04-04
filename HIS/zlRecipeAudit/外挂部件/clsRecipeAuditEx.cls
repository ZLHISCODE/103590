VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecipeAuditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Init(ByVal cnOracle As ADODB.Connection, ByVal bytClass As Byte, _
    ByVal strIDs As String) As Boolean
'功能：初始化外挂接口
'参数：
'  cnOracle：ZLHIS端的数据库连接对象
'  bytClass：1-门诊；2-住院
'  strIDs：待审查的给药途径医嘱ID；格式：给药途径医嘱ID[,给药途径医嘱ID[...]]
'返回：True成功；False失败

    gblnInit = False
    
    If cnOracle Is Nothing Then
        MsgBox "数据库连接参数对象不正确！", vbInformation
        Exit Function
    End If
    If bytClass < 1 Or bytClass > 2 Then
        MsgBox "指定门诊或住院的参数不正确！", vbInformation
        Exit Function
    End If
    If strIDs = "" Then
        MsgBox "医嘱参数无！", vbInformation
        Exit Function
    End If
    
    Set gcnOracle = cnOracle
    gbytClass = bytClass
    gstrIDs = strIDs
    
    gblnInit = True
    Init = True
    
End Function

Public Function Check(ByVal strFunNO As String, ByRef strMedicalID As String, ByRef strErr As String) As Boolean
'功能：审查指定功能编号，通过Init方法传入的药嘱记录是否合格
'参数：
'  strFunNO：功能编码
'  strMedicalID（实参）：合格/不合格的药嘱ID
'  strErr（实参）：异常信息
'返回：True合格；False不合格

    If gblnInit = False Then
        MsgBox "请先初始化接口！", vbInformation
        Exit Function
    End If
    
    If LCase(strFunNO) = LCase(GSTR_CODE_中药注射剂) Then
        Check = F_中药注射剂(strMedicalID, strErr)
    'ElseIf 新审查项目 Then
    '  Check = F_新方法(...)    '新方法请在mdlRecipeAuditEx模块中新增逻辑代码，以及定义方法的编码常量，详见“GSTR_CODE_中药注射剂”常量的定义
    End If
    
End Function

Public Function CloseWindows() As Boolean
'功能:关闭部件里所有的窗体
'返回:True成功；False失败

    Dim frmThis As Form
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

