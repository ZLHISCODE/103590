Attribute VB_Name = "mdlReadDataPatiPath"
Option Explicit
'---------------------------------------------------------------------------------------
'本模块只负责数据层面的访问
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : ReadPathPhase
' Author    : YWJ
' Date      : 2019-04-29
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function ReadPathPhase(ByVal lngPatiPathID As Long, ByVal lngPhaseBranchId As Long) As ADODB.Recordset
'参数:
'lngPatiPathID-病人路径记录ID
'lngPhaseBranchId-阶段分支ID
    Dim strSQL As String
    
    On Error GoTo errH
    '阶段排序时用 NVL(c.序号,b.序号) 是为了处理备用分支序列排序的问题，取值b.序号 是因为界面上需要显示是第几个分支。（取分支路径的序号时，取其其一阶段的序号加上分支路径的序号）
    If lngPhaseBranchId = 0 Then
        strSQL = _
        "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,b.路径ID,b.开始天数 " & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id " & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 临床路径阶段 B,临床路径阶段 C,病人路径评估 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And g.路径记录id(+) = a.路径记录id And g.阶段id(+) = a.阶段id And g.日期(+) = a.日期 " & vbNewLine & _
                 "Order By 日期,g.登记时间,NVL(c.序号,b.序号)"
    Else
        strSQL = _
        "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期, To_Char(a.日期, 'day') 星期, b.名称 As 阶段名, b.序号, b.说明, b.父id,b.路径ID,b.开始天数 " & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id " & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 临床路径阶段 B,临床路径阶段 C,临床路径分支 D,临床路径阶段 E,临床路径阶段 F,病人路径评估 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And b.分支id=d.id(+) and d.前一阶段id=e.id(+) And e.父id=f.id(+) And g.路径记录id(+) = a.路径记录id And g.阶段id(+) = a.阶段id And g.日期(+) = a.日期 " & vbNewLine & _
                 "Order By 日期,g.登记时间, Decode(b.分支ID,Null,NVL(c.序号,b.序号),NVL(c.序号,b.序号)+NVL(f.序号,e.序号))"
    End If

    Set ReadPathPhase = zlDatabase.OpenSQLRecord(strSQL, "ReadPathPhase", lngPatiPathID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReadPathPhaseNoEvaluate
' Author    : YWJ
' Date      : 2019-04-29
' Purpose   :获取待评估的路径阶段
'---------------------------------------------------------------------------------------
'
Public Function ReadPathPhaseNoEvaluate(ByVal lngPatiPathID As Long, ByVal lngPhaseBranchId As Long) As ADODB.Recordset
'参数:
'lngPatiPathID-路径记录ID
'lngPhaseBranchId-阶段分支ID
    Dim strSQL As String
    
    On Error GoTo errH
    '阶段排序时用 NVL(c.序号,b.序号) 是为了处理备用分支序列排序的问题，取值b.序号 是因为界面上需要显示是第几个分支。（取分支路径的序号时，取其其一阶段的序号加上分支路径的序号）
    '病人路径评估
    If lngPhaseBranchId = 0 Then
        strSQL = "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期" & vbNewLine & _
                "From (Select a.阶段id, a.天数, a.日期, a.路径记录id" & vbNewLine & _
                "       From 病人路径执行 A" & vbNewLine & _
                "       Where a.路径记录id = [1]" & vbNewLine & _
                "       Group By a.阶段id, a.天数, a.日期, a.路径记录id) A, 临床路径阶段 B, 临床路径阶段 C, 病人路径评估 G" & vbNewLine & _
                "Where a.阶段id = b.Id And b.父id = c.Id(+) And g.路径记录id(+) = a.路径记录id And g.阶段id(+) = a.阶段id And g.日期(+) = a.日期 And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From 病人路径评估 P Where p.路径记录id = a.路径记录id And p.阶段id = a.阶段id And p.日期 = a.日期)" & vbNewLine & _
                "Order By 日期, g.登记时间, Nvl(c.序号, b.序号)"
                 
    Else
        strSQL = _
        "Select a.阶段id, a.天数, To_Char(a.日期, 'yyyy-mm-dd') 日期" & vbNewLine & _
                 "From (Select a.阶段id, a.天数, a.日期,a.路径记录id " & vbNewLine & _
                 "       From 病人路径执行 A" & vbNewLine & _
                 "       Where a.路径记录id = [1]" & vbNewLine & _
                 "       Group By a.阶段id, a.天数, a.日期,a.路径记录id) A, 临床路径阶段 B,临床路径阶段 C,临床路径分支 D,临床路径阶段 E,临床路径阶段 F,病人路径评估 G" & vbNewLine & _
                 "Where a.阶段id = b.Id And b.父id=c.id(+) And b.分支id=d.id(+) and d.前一阶段id=e.id(+) And e.父id=f.id(+) And g.路径记录id(+) = a.路径记录id And g.阶段id(+) = a.阶段id And g.日期(+) = a.日期 And" & vbNewLine & _
                "      Not Exists" & vbNewLine & _
                " (Select 1 From 病人路径评估 P Where p.路径记录id = a.路径记录id And p.阶段id = a.阶段id And p.日期 = a.日期)" & vbNewLine & _
                 "Order By 日期,g.登记时间, Decode(b.分支ID,Null,NVL(c.序号,b.序号),NVL(c.序号,b.序号)+NVL(f.序号,e.序号))"
    End If

    Set ReadPathPhaseNoEvaluate = zlDatabase.OpenSQLRecord(strSQL, "ReadPathPhaseNoEvaluate", lngPatiPathID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDiagType
' Author    : YWJ
' Date      : 2019-05-08
' Purpose   : 获取病人本次就诊编码类别
'---------------------------------------------------------------------------------------
Public Function GetDiagType(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As ADODB.Recordset
'参数:
'   lngPatiID   -病人ID
'   lngVisitID  -主页ID
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select Distinct Nvl(a.编码类别, 'D') As 编码类别 From 病人诊断记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set GetDiagType = zlDatabase.OpenSQLRecord(strSQL, "GetDiagType", lngPatiID, lngVisitID)
   
   Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
