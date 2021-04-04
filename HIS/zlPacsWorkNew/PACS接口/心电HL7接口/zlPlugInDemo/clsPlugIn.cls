VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------
'功能：外挂功能接口部件，用于一些工作站集成外挂调用功能，并传入相应的业务参数
'编制人：周韬
'编制时间：2010-07-22
'说明：
'  2010-07-23:集成到了门诊/住院医嘱虚拟模块中
'---------------------------------------------------------------------------

Public Sub Initialize(ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngModual As Long)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
    
    Set gcnOracle = cnOracle
    
End Sub

Public Function GetFuncNames(ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal strReserve As String) As String
'功能：获取外挂功能列表；该接口一般在初始化之后调用一次，调用时机和频度可能因不同的主程序而不同
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      strReserve=保留参数,用于扩展使用
'返回：外挂功能列表，格式为"功能名1,功能名2,功能名3,..."
    
    Select Case lngModual
        Case 1252
            GetFuncNames = gstrFunc_PACS影像调阅 & "," & gstrFunc_PACS报告调阅 & "," & gstrFunc_MUSE心电结果调阅
        Case 1253
            GetFuncNames = gstrFunc_PACS影像调阅 & "," & gstrFunc_PACS报告调阅 & "," & gstrFunc_MUSE心电结果调阅
    End Select
End Function

Public Sub ExecuteFunc(ByVal lngSys As Long, ByVal lngModual As Long, _
    ByVal strFuncName As String, ByVal lngPatiID As Long, ByVal varRecId As Variant, _
    ByVal varKeyId As Variant, Optional ByVal strReserve As String)
'功能：执行某个外挂功能
'参数：lngSys,lngModual=当前调用接口的上级系统号及模块号
'      strFuncName=要执行的功能名，该功能名之前由GetFuncNames函数返回
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
'      strReserve=保留参数,用于扩展使用
    On Error GoTo err
    
    Select Case lngModual
        Case 1252
            '门诊医生站
            Select Case strFuncName
                Case gstrFunc_PACS影像调阅
                    Call ShowPacsViewer(varKeyId, 1)
                Case gstrFunc_PACS报告调阅
                    Call ShowPacsViewer(varKeyId, 2)
                Case gstrFunc_MUSE心电结果调阅
                    Call ShowMUSEViewer(varKeyId)
            End Select
        Case 1253
            '住院医生站
            Select Case strFuncName
                Case gstrFunc_PACS影像调阅
                    Call ShowPacsViewer(varKeyId, 1)
                Case gstrFunc_PACS报告调阅
                    Call ShowPacsViewer(varKeyId, 2)
                Case gstrFunc_MUSE心电结果调阅
                    Call ShowMUSEViewer(varKeyId)
            End Select
    End Select
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "系统消息"
    Exit Sub
End Sub

Public Sub Terminate(ByVal lngSys As Long, ByVal lngModual As Long)
'功能：接口终止
'参数：lngSys,lngModual=当前调用接口的主程序系统号及模块号
    Call PacsRelease
End Sub
