VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function gOpenDataBase(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '===================================================================================
    '功能:                                  打开数据库连接
    '参数;
    'strServerName                          服务器连接字串
    'strUserName                            用户名
    'strUserPwd                             密码
    '返回                                   True=连接成功 False=连接失败
    '===================================================================================
    gOpenDataBase = OraDataOpen(strServerName, strUserName, strUserPwd)
End Function

Public Function gGetApplication(strPatientID As String) As String
    '=========================================================================================
    '功能:                              得到病人申请单的记录集
    '参数
    'strPatientID                       数字为就诊卡号或条码、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号、“\”健康号
    '返回                               返回字串内容:"医嘱ID^姓名^性别^年龄^病人来源^门诊号^住院号^申请科室编码^申请科室名称^开嘱医生^开嘱时间^医保类别"
    '                                   为空表示没有找到记录
    '=========================================================================================
    gGetApplication = GetApplication(strPatientID)
End Function

Public Function gOraDataClose() As Boolean
    '==============================================================
    '功能                               '关闭数据库连接
    '==============================================================
    gOraDataClose = OraDataClose
End Function

Public Function gInsertReport(lngID As Long, strReportPath As String, strErrInfo As String) As Boolean
    '===================================================================
    '功能                               插入报告到HIS
    '参数
    'lngID                              医嘱ID
    'strReportPath                      报告路径
    '===================================================================
    gInsertReport = InsertReport(lngID, strReportPath, strErrInfo)
End Function

Public Function gDeleteReport(lngID As Long) As Boolean
    '===================================================================
    '功能                               删除报告
    '参数
    'lngID                              医嘱ID
    '===================================================================
    gDeleteReport = DeleteReport(lngID)
End Function

Public Function gGetClinicItem(lngAdivce As Long) As String
    '===================================================================
    '功能                               取得要做的诊疗项目内容
    '参数
    'lngAdivce                          医嘱ID
    '返回                               字串格式:诊疗项目ID^诊疗项目编码^诊疗项目名称^执行科室编码^执行科室名称^单价^金额^是否收费
    '===================================================================
    gGetClinicItem = GetClinicItem(lngAdivce)
End Function

Public Function gGetItemList(lngClinicID As Long) As String
    '===================================================================
    '功能                               取得诊疗项目的指标明细
    '参数
    'lngClinicID                        诊疗项目ID
    '返回                               字串格式:编码^中文名^英文名^单价
    '===================================================================
    gGetItemList = GetItemList(lngClinicID)
End Function

Public Function gSetRegister(lngAdivce As Long, intTag As Integer) As Boolean
    '=====================================================================
    '功能                               标本核收或取消核收
    '参数
    'lngAdivce                          医嘱ID
    'intTag                             1=核收 0=取消核收
    '=====================================================================
    gSetRegister = SetRegister(lngAdivce, intTag)
End Function

Public Function gGetAllItem(Optional strFindItem As String) As String()
    '=====================================================================
    '功能                               取得所有的诊疗项目编码和名称
    '参数
    'strItem                            可选，查找编码和名称相同的诊疗项目项目
    '返回                               查找到的诊疗项目数组
    '=====================================================================
    gGetAllItem = GetAllItem(strFindItem)
End Function

Public Function gTestResults(ByVal lngID As Long, ByVal strTestName As String, ByVal strTestTime As String, ByVal strTestResults As String) As String
    '===================================================================
    '功能                               返回检验结果到体检系统
    '参数
    'lngID                              医嘱ID
    'strTestName                        检验人
    'strTestTime                        检验时间，格式 2009-01-01 10:30:01
    'strTestResults                     医嘱ID对应的检验结果，可以对多少个检验指标一起处理，详细格式如下：
    '
    '                                     诊治项目id;检验结果1;单位1;结果参1考;结果标志1|诊治项目id;检验结果2;单位2;结果参考2;结果标志2......
    '
    '                                     其中，结果标志在 “偏低,偏高,异常,空串”中选择一个返回。
    '返回: 空，表示更新成功，非空，表示错误信息。
    '===================================================================
    gTestResults = UpdateTestResults(lngID, strTestName, strTestTime, strTestResults)
End Function

Public Function gzlLisRegister(ByVal lngDevID As Long, ByVal lngID As Long, ByVal strSampleNo As String, ByRef strErrInfo As String) As Boolean
    '===================================================================
    '功能                               核收申请到zllis
    '参数
    'lngDevID                           仪器ID
    'lngID                              医嘱ID
    'strSampleNo                        标本号


    '===================================================================
    gzlLisRegister = zlLISRegister(lngDevID, lngID, strSampleNo, strErrInfo)
End Function

Public Function gzlLisUnRegister(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
    '===================================================================
    '功能                               取消在zllis中已核收的标本
    '参数
    'lngDevID                           仪器ID
    'lngID                              医嘱ID
    'strSampleNo                        标本号

    '返回: True，表示更新成功，false，表示失败，信息在strErrInfo变量中。
    '===================================================================
    gzlLisUnRegister = zlLisUnRegister(lngID, strErrInfo)
End Function

Public Function gZLLisInsterReport(ByVal lngID As Long, ByVal strItems As String, ByRef strErrInfo As String) As Boolean
    '写入已审核报告
    gZLLisInsterReport = ZLLisInsterReport(lngID, strItems, strErrInfo)
End Function

Public Function gzlLisUnAudit(ByVal lngID As Long, ByRef strErrInfo As String) As Boolean
    '取消已审报告
    gzlLisUnAudit = zlLisUnAudit(lngID, strErrInfo)
End Function

Public Function gGetAllDevice(ByRef strErrInfo As String) As String
    '取检验仪器
    gGetAllDevice = GetAllDevice(strErrInfo)
End Function

Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
End Sub

