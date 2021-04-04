Attribute VB_Name = "mdlCommon"
Option Explicit

Public gstrSysName As String                '系统名称
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String    '产品名称
Public gstrSQL As String
Public glngSys As Long
Public glngMainModule As Long '调用者的模块号
Public gstrMainPrivs As String '调用者的相关权限
Public gcnOracle As ADODB.Connection
Public grsStockCheck As ADODB.Recordset      '库存检查
Public gstrDBUser As String '所有者

'公共对象定义
Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjDatabase As Object
Public gstrNodeNo As String '站点名

'接口要使用到的系统参数
Public Type Type_SysParms
    P9_费用金额保留位数 As Integer
    P150_药品出库优先算法 As Integer
    P157_费用单价保留位数 As Integer
End Type
Public gtype_UserSysParms As Type_SysParms     '系统参数

Public Enum StockCheck
    不检查 = 0
    不足提醒 = 1
    不足禁止 = 2
End Enum

'用户信息------------------------
Public Type TYPE_USER_INFO
    用户ID As Long
    用户编码 As String
    用户姓名 As String
    用户简码 As String
    部门ID As Long
    部门编码 As String
    部门名称 As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-05-15 15:34:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.gobjDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function
 
Public Sub GetSysParms()
    '取系统参数值
    
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    
    gstrSQL = "Select 参数号, 参数值, 缺省值 From Zlparameters Where 系统 = 100 And Nvl(私有, 0) = 0 And 模块 Is Null Order By 参数号 "
    Set rsTemp = gobjDatabase.OpenSQLRecord(gstrSQL, "GetSysParms")
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.Filter = "参数号=9"
        If Not rsTemp.EOF Then gtype_UserSysParms.P9_费用金额保留位数 = Val(NVL(rsTemp!参数值, rsTemp!缺省值))
        
        rsTemp.Filter = "参数号=150"
        If Not rsTemp.EOF Then gtype_UserSysParms.P150_药品出库优先算法 = Val(NVL(rsTemp!参数值, rsTemp!缺省值))
        
        rsTemp.Filter = "参数号=157"
        If Not rsTemp.EOF Then gtype_UserSysParms.P157_费用单价保留位数 = Val(NVL(rsTemp!参数值, rsTemp!缺省值))
    End If
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    '功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub GetStockCheckRule()
    '取库存检查规则
    
    gstrSQL = "Select 库房id, 检查方式 From 药品出库检查 "
    Set grsStockCheck = gobjDatabase.OpenSQLRecord(gstrSQL, "GetStockCheckRule")
    
End Sub

Public Sub GetUserInfo()
    Dim rsUser As ADODB.Recordset
    
    Set rsUser = gobjDatabase.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            UserInfo.用户ID = !Id
            UserInfo.用户编码 = !编号
            UserInfo.用户姓名 = IIf(IsNull(!姓名), "", !姓名)
            UserInfo.用户简码 = IIf(IsNull(!简码), "", !简码)
            UserInfo.部门ID = !部门ID
            UserInfo.部门编码 = !部门码
            UserInfo.部门名称 = !部门名
        Else
            UserInfo.用户ID = 0
            UserInfo.用户编码 = ""
            UserInfo.用户姓名 = ""
            UserInfo.用户简码 = ""
            UserInfo.部门ID = 0
            UserInfo.部门编码 = ""
            UserInfo.部门名称 = ""
        End If
    End With
End Sub
