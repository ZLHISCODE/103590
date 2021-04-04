VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'常量定义
'######################################################################################################################
Private mstrSQL As String
Private mstrTitle As String

Public Function Get_PeisPersonReportBill(ByVal lngPersonKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************

    On Error GoTo errHand

    mstrSQL = _
        "Select A.开始时间, A.结束时间, B.任务名称, A.任务id, A.病人id" & vbNewLine & _
        "From 体检任务人员 A, 体检任务记录 B" & vbNewLine & _
        "Where A.病人id = [1] And A.体检状态 > 1 And A.任务id = B.ID And A.开始时间  is not Null " & vbNewLine & _
        "Order By A.开始时间 Desc"

    Set Get_PeisPersonReportBill = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngPersonKey)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function Get_PeisPersonItem(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long) As ADODB.Recordset
    
    On Error GoTo errHand
    
    mstrSQL = _
            "Select Distinct '报告' As 图标, a.清单id, b.Id, b.名称 As 项目, a.采集方式id, a.标本形态, a.备注说明, c.排列顺序" & vbNewLine & _
            "From 体检任务发送 A, 体检项目目录 B, 体检项目顺序 C" & vbNewLine & _
            "Where a.任务id = [1] And a.病人id = [2] And a.体检项目id = b.Id And b.Id = c.体检项目id(+)" & vbNewLine & _
            "Order By b.Id, c.排列顺序"



    Set Get_PeisPersonItem = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngTaskKey, lngPersonKey)

    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function



Public Function Get_PeisPersonConclusion(ByVal lngStation As Long, ByVal lngTaskKey As Long, ByVal lngPersonKey As Long, Optional ByVal lng清单id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case lngStation
        Case 1
            mstrSQL = _
                    "Select ID, 结论描述, 记录序号" & vbNewLine & _
                    "From 体检任务结论" & vbNewLine & _
                    "Where 结论性质 = 1 And 任务id = [1] And 病人id = [2] And 清单id=[3] And 记录性质 = 1" & vbNewLine & _
                    "Order By 记录序号"
        Case 2
            mstrSQL = _
                    "Select ID, 结论描述, 记录序号" & vbNewLine & _
                    "From 体检任务结论" & vbNewLine & _
                    "Where 结论性质 = 2 And 任务id = [1] And 病人id = [2] And 记录性质 = 1" & vbNewLine & _
                    "Order By 记录序号"
    End Select
    
    Set Get_PeisPersonConclusion = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngTaskKey, lngPersonKey, lng清单id)
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function get_PeisPersonResult(ByVal lng任务id As Long, ByVal lng人员id As Long, ByVal lng清单id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '说明：
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = _
        "Select a.清单id, a.体检项目id As ID, b.名称 As 指标, a.结果, a.报警 As 提示, a.参考" & vbNewLine & _
        "From 体检任务结果 A, 体检指标目录 B" & vbNewLine & _
        "Where a.任务id = [1] And a.病人id = [2] And a.清单id = [3] And a.体检指标id = b.Id"

    
    Set get_PeisPersonResult = gobjDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng任务id, lng人员id, lng清单id)
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

