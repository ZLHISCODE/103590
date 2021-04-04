Attribute VB_Name = "mdlPubComLib"
Option Explicit


Private mblnInit As Boolean                                        '公共部件是否已初始化

Public gcnOracle As New ADODB.Connection     '公共数据库连接
Public gdtStart     As Long                  '启动时间，用以判断闪现屏幕的等待时间

Public zl9ComLib As Object
Public zlDatabase As Object
Public zlCommFun As Object
Public zlControl As Object

Public Type TYPE_SYS_INFO   '-----------应用程序信息 及 注册信息
    AppName As String       '系统名称 (产品简称+软件，如中联软件，医业软件)
    ShortName As String     '产品简名
    AppTitle As String      '系统标题，产品全称
    
    Version As String       '系统版本
    AviPath As String       'AVI文件路径
    
    UnitName  As String     '用户单位名称
    Supporter As String     '技术支持商
    Develop As String       '开发商
    SupporterWEB As String  '支持商WEB简名
    SupporterMail As String '支持商邮件
    SupporterURL As String  '支持商网址
    ProductLine  As String  '产品系列，[标准版],[大客户版]
    
    SysNo       As Long     '系统编号
    ModlNo      As Long     '模块号
End Type

Public Type TYPE_SYS_PARAMETER    '系统参数
    Privs        As String  '模块权限
    
    MachineCount As Integer '仪器数量
    blnEmerge    As Boolean '是否区分急诊
    BuffDir      As String   '本地缓存记录集的缓存目录
    InvaidWord   As String   '需去掉的非常字符
    intCA        As Integer  'CA中心编号
    strMatch     As String   '输入匹配
    
    LogLevel     As LOGTYPE  '日志记录等级 3-错误，4-警告 6-提示 7-调试
    strDevList  As String     '本机连接仪器的列表
    
    ftpSetup    As String    'FTP设置
    
End Type

Public Type TYPE_USER_INFO
    ID As Long          '人员ID
    DeptID As Long      '人员对应的部门ID
    DeptName As String  '人员对应的部门名称
    No As String        '人员编号
    Name As String      '人员姓名
    Code As String      '人员简码
    DBUser As String    '人员对应的数据库用户名
End Type

Public UserInfo As TYPE_USER_INFO
Public gSysInfo As TYPE_SYS_INFO
Public gSysParameter As TYPE_SYS_PARAMETER

'调用 公共部件ComLib的一些公共函数过程
Public Function ComOpenSQL(ByVal strSQL As String, ByVal strTitle As String, _
    ParamArray arrInput() As Variant) As ADODB.Recordset
    '功能：通过ComLib对象打开带参数SQL的记录集
    
    Dim lngCount As Long
    Dim var(30) As Variant
    
    If Not mblnInit Then Exit Function
    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        Err.Raise -2147483645, , "不支持超过30个参数的SQL！"
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    Set ComOpenSQL = zlDatabase.OpenSQLRecord(strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))

End Function

Public Function ComExecuteProc(strSQL As String, ByVal strFormCaption As String) As String
    '功能：执行过程语句,并自动对过程参数进行绑定变量处理
    '返回：无错误返回空串，否则返回错误提示
    If Not mblnInit Then Exit Function
    Call zlDatabase.ExecuteProcedure(strSQL, strFormCaption)
    Exit Function
End Function

Public Function ComInitComLib(ByRef strErr As String) As Boolean
    '初始化公共部件,在程序启动时调用
    Dim strSQL As String
    On Error GoTo errH
    ComInitComLib = False
    If mblnInit Then
        ComInitComLib = True
        Exit Function
    End If
    
    Set zl9ComLib = CreateObject("zl9ComLib.clsComLib")
    zl9ComLib.InitCommon gcnOracle

    Set zlDatabase = zl9ComLib.zlDatabase
    Set zlCommFun = zl9ComLib.zlCommFun
    Set zlControl = zl9ComLib.zlControl
    
'    If zl9ComLib.RegCheck = False Then
'        strErr = "注册信息验证未通过！"
'        Exit Function
'    End If
    
    '如果发行码无效（为空或为"-"），则退出
    gSysInfo.ShortName = zl9ComLib.zlRegInfo("产品简名")
    gSysInfo.UnitName = zl9ComLib.zlRegInfo("单位名称", , -1)
    gSysInfo.AppName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gSysInfo.Version = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gSysInfo.AviPath = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gSysInfo.AppTitle = zl9ComLib.zlRegInfo("产品标题")

     
    gSysInfo.SysNo = 100   '系统号
    gSysInfo.ModlNo = 1208  '模块号
    
'    strSQL = "Zl_Createsynonyms(" & gSysInfo.SysNo & ")"
'    zlDatabase.ExecuteProcedure strSQL, "创建同义词"

    ComInitComLib = True
    mblnInit = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
 
End Function

Public Function ComGetPrivs(ByVal lngSys As Long, ByVal lngModul As Long) As String
    '读取模块权限
   ComGetPrivs = zl9ComLib.GetPrivFunc(lngSys, lngModul)
End Function


Public Function ComGetSysParameter(strErr As String) As Boolean
    '读取系统参数
    On Error GoTo errH
    ComGetSysParameter = False
    gSysParameter.InvaidWord = "`#@$%&|\{}[]?;""'"
    
    gSysParameter.ftpSetup = zlDatabase.GetPara("FTP设置", gSysInfo.SysNo, gSysInfo.ModlNo, "")
    ComGetSysParameter = True
    Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
'    Call SaveLog("GetSysPra", LOG_ERR, Err.Number, strErr)
End Function

Public Function ComSetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '设置参数
    ComSetPara = zlDatabase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    '取参数
    ComGetPara = zlDatabase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function
Public Function ComGetUserInfo(ByRef strErr As String) As Boolean
    '功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    zl9ComLib.SetDbUser UserInfo.DBUser
    Set rsTmp = zlDatabase.GetUserInfo

    If Not rsTmp.EOF Then
        UserInfo.ID = Val("" & rsTmp!ID)
        UserInfo.No = Trim("" & rsTmp!编号)
        UserInfo.DeptID = Val("" & rsTmp!部门ID)
        UserInfo.DeptName = Trim("" & rsTmp!部门名)
        UserInfo.Code = Trim("" & rsTmp!简码)
        UserInfo.Name = Trim("" & rsTmp!姓名)
    
        ComGetUserInfo = True
 
    End If
            
Exit Function
errH:
    strErr = Err.Number & " " & Err.Description
'    Call SaveLog("ComGetUserInfo", LOG_ERR, Err.Number, strErr)
End Function

Public Function ComGetNextID(ByVal strTableName As String) As Long
    '取表名对应的序列
    ComGetNextID = zlDatabase.GetNextId(strTableName)
End Function

Public Function ComOEMPicture(objPic As Object, strAttribute As String, Optional strProductName As String)
    '取OEM图片
    On Error GoTo errH
    Call zl9ComLib.ApplyOEM_Picture(objPic, strAttribute, strProductName)
    Exit Function
errH:
'    Call SaveLog("OEMPicture", LOG_ERR, Err.Number, Err.Description)
End Function

Public Function ComGetLike(ByVal strTable As String, ByVal strField As String, ByVal strInput As String) As String
    '获取Like条件
    ComGetLike = zlCommFun.GetLike(strTable, strField, strInput)
End Function

Public Function ComPressKey(bytKey As Byte)
    '执行PressKey功能
     Call zlCommFun.PressKey(bytKey)
End Function

Public Function ComGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '获取简码函数
   ComGetSymbol = zlCommFun.zlGetSymbol(strInput, bytIsWB)
End Function

Public Function ComIncStr(ByVal strVal As String) As String
    ComIncStr = zlCommFun.IncStr(strVal)
End Function

Public Function ComErrCenter() As Byte
    '错误处理中心
    ComErrCenter = zl9ComLib.ErrCenter
End Function

Public Function ComCurrDate() As Date
    '取服务器当前日期时间
    ComCurrDate = zlDatabase.Currentdate
End Function







