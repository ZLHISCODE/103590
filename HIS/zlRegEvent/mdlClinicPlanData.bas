Attribute VB_Name = "mdlClinicPlanData"
Option Explicit
Public grsWorkTime As ADODB.Recordset '所有上班时段，缓存
Public grsUnit As ADODB.Recordset '所有合作单位与预约方式，缓存
Public Function LpadTime(ByVal strStartTime As String, ByVal strEndTime As String, ByRef dtStartDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:左补齐日期
    '入参:strStartTime-开始时间,格式为HH:MM:SS
    '     strEndTime-开始时间,格式为HH:MM:SS
    '出参:dtStartDate-开始时间,即yyyy-mm-dd hh:mm:ss
    '     dtEndDate-终止时间,即yyyy-mm-dd hh:mm:ss
    '返回:如果补齐成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-24 14:50:32
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCurDate As String
    Dim dtStart As Date, dtEnd As Date
    
    On Error GoTo errHandle
    
    strCurDate = Format(Date, "yyyy-mm-dd")
    If strStartTime = "" Or strEndTime = "" Then Exit Function
    strStartTime = Format(strStartTime, "HH:MM:SS")
    strEndTime = Format(strEndTime, "HH:MM:SS")
    
    dtStart = CDate(strCurDate & " " & strStartTime)
    dtEnd = CDate(strCurDate & " " & strEndTime)
    If dtStart >= dtEnd Then dtEnd = dtEnd + 1
    
    dtStartDate = dtStart
    dtEndDate = dtEnd
    LpadTime = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetWorkTimeRange(str时间段 As String, ByVal str站点 As String, ByVal str号类 As String) As 上班时段
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据时间段名获取上班时段对象信息
    '入参:str时间段-时间段名称
    '返回:返回上班时段信息
    '编制:刘兴洪
    '日期:2016-03-24 16:03:25
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj上班时段 As 上班时段, strStart As String, strEnd As String
    Dim rsWorkTime As ADODB.Recordset
    On Error GoTo errHandle
    
    '时间段, 开始时间, 终止时间, 缺省时间, 提前时间, 提前颜色, nvl(站点,'-') as 站点, nvl(号类,'-') as 号类, 序号, 出诊预留时间, 休息时段
    Set rsWorkTime = GetWorkTimeRec
    '1.看有无本号源适用的时段
    If str站点 <> "" And str号类 <> "" Then
        rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='" & str站点 & "' And 号类='" & str号类 & "'"
        If rsWorkTime.EOF Then
            rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='" & str站点 & "' And 号类='-'"
            If rsWorkTime.EOF Then
                rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='" & str号类 & "'"
                If rsWorkTime.EOF Then rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='-'"
            End If
        End If
    ElseIf str站点 <> "" And str号类 = "" Then
        rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='" & str站点 & "' And 号类='-'"
        If rsWorkTime.EOF Then rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='-'"
    ElseIf str站点 = "" And str号类 <> "" Then
        rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='" & str号类 & "'"
        If rsWorkTime.EOF Then rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='-'"
    Else
        rsWorkTime.Filter = rsWorkTime.Filter = "时间段='" & str时间段 & "' And 站点='-' And 号类='-'"
    End If
    '存在站点或号类的
    If Not rsWorkTime.EOF Then
        Set obj上班时段 = New 上班时段
        With obj上班时段
            .时间段 = str时间段
            .出诊预留时间 = Val(Nvl(rsWorkTime!出诊预留时间))
            .开始时间 = Format(rsWorkTime!开始时间, "yyyy-mm-dd HH:MM:SS")
            .结束时间 = Format(rsWorkTime!终止时间, "yyyy-mm-dd HH:MM:SS")
            .缺省预约时间 = Format(rsWorkTime!缺省时间, "yyyy-mm-dd HH:MM:SS")
            .提前挂号时间 = Nvl(rsWorkTime!提前时间)
            .休息时段 = Nvl(rsWorkTime!休息时段)
        End With
        Set GetWorkTimeRange = obj上班时段
        Exit Function
    End If
   Set GetWorkTimeRange = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 
End Function

Public Function GetClinicRecordFromSignalSource(ByVal lng号源Id As Long) As 出诊记录集
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据号源ID获取出诊记录集
    '入参:lng号源ID-号源ID
    '返回:返回出诊记录集
    '编制:刘兴洪
    '日期:2016-03-22 17:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录集 As 出诊记录集, obj出诊记录 As 出诊记录
    Dim obj分诊诊室集 As 分诊诊室集, obj分诊诊室 As 分诊诊室
    Dim obj号序信息集 As 号序信息集, obj号序信息 As 号序信息
    Dim obj合作单位控制集 As 合作单位控制集, obj合作单位控制 As 合作单位控制
    Dim obj上班时段  As 上班时段
    Dim rsControl As ADODB.Recordset, rsWorkTime As ADODB.Recordset, rsNum As ADODB.Recordset, rsUnitControl As ADODB.Recordset
    Dim dtDate As Date, strTemp As String, strSQL As String
    Dim rsRoom As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set obj出诊记录集 = New 出诊记录集
    Set obj出诊记录 = New 出诊记录
    Set rsWorkTime = GetWorkTimeRec
    strSQL = " " & _
    "   Select a.Id,C.号类,C.出诊频次,C.科室ID,C.医生ID,C.医生姓名," & vbNewLine & _
    "          a.上班时段, a.限号数, a.限约数, a.是否序号控制, a.是否分时段, a.预约控制," & vbNewLine & _
    "          a.是否独占, a.分诊方式, a.诊室id, b.名称 As 门诊诊室, d.站点 " & vbNewLine & _
    "   From 临床出诊号源 C, 临床出诊号源限制 A, 门诊诊室 B, 部门表 D" & vbNewLine & _
    "   Where c.Id = a.号源id And a.诊室id = b.Id(+) And c.科室ID = d.ID And c.Id = [1]"
    Set rsControl = zlDatabase.OpenSQLRecord(strSQL, "获取号源控制限制信息", lng号源Id)

    strSQL = " " & _
    "   Select b.限制id, b.诊室id, c.名称 As 门诊诊室 " & _
    "   From 临床出诊号源限制 A, 临床出诊号源诊室 B, 门诊诊室 C " & _
    "   Where a.号源id = [1] And a.Id = b.限制id And B.诊室id = c.Id"
    Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, "获临床出诊号源诊室信息", lng号源Id)

    strSQL = " " & _
    "   Select b.限制id, b.序号, b.开始时间, b.终止时间, b.限制数量, b.是否预约 " & _
    "   From 临床出诊号源限制 A, 临床出诊号源时段 B " & _
    "   Where a.号源id = [1] And a.Id = b.限制id"
    Set rsNum = zlDatabase.OpenSQLRecord(strSQL, "获临床出诊号源时段信息", lng号源Id)
    

    strSQL = "" & _
    "Select  b.限制id, b.类型, b.性质, b.名称, b.序号, b.控制方式, b.数量, c.开始时间, c.终止时间 " & _
    "   From 临床出诊号源限制 A, 临床出诊号源控制 B, 临床出诊时段 C" & _
    "   Where a.号源id = [1] And a.Id = b.限制id And b.限制ID = c.限制ID(+) And b.序号 = c.序号(+)"
 
    Set rsUnitControl = zlDatabase.OpenSQLRecord(strSQL, "获临床出诊号源控制信息", lng号源Id)
    
    dtDate = zlDatabase.Currentdate
    With rsControl
        Do While Not .EOF
            If Nvl(rsControl!上班时段) <> "" Then
                Set obj上班时段 = GetWorkTimeRange(Nvl(rsControl!上班时段), Nvl(rsControl!站点), Nvl(rsControl!号类))
                Set obj出诊记录 = New 出诊记录
                rsWorkTime.Filter = "时间段='" & Nvl(rsControl!上班时段) & "'"
                If Not rsWorkTime.EOF Then
                    obj出诊记录.出诊日期 = Format(obj上班时段.开始时间, "yyyy-mm-dd")
                    obj出诊记录.分诊方式 = Val(Nvl(rsControl!分诊方式))
                    obj出诊记录.记录ID = Val(Nvl(rsControl!id))
                   Set obj出诊记录.上班时段 = obj上班时段
                    obj出诊记录.开始时间 = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(rsWorkTime!开始时间, "HH:MM:SS"))
                    If Format(rsWorkTime!开始时间, "yyyy-mm-dd HH:MM:SS") >= Format(rsWorkTime!终止时间, "yyyy-mm-dd HH:MM:SS") Then
                        obj出诊记录.终止时间 = CDate(Format(dtDate + 1, "yyyy-mm-dd") & " " & Format(rsWorkTime!终止时间, "HH:MM:SS"))
                    Else
                        obj出诊记录.终止时间 = CDate(Format(dtDate, "yyyy-mm-dd") & " " & Format(rsWorkTime!终止时间, "HH:MM:SS"))
                    End If
                    obj出诊记录.时间段 = Nvl(rsControl!上班时段)
                    obj出诊记录.是否分时段 = Val(Nvl(rsControl!是否分时段)) = 1
                    obj出诊记录.是否独占 = Val(Nvl(rsControl!是否独占)) = 1
                    obj出诊记录.是否序号控制 = Val(Nvl(rsControl!是否序号控制)) = 1
                    obj出诊记录.替诊医生 = ""
                    obj出诊记录.科室ID = Val(Nvl(rsControl!科室ID))
                    obj出诊记录.医生ID = Val(Nvl(rsControl!医生ID))
                    obj出诊记录.医生姓名 = Nvl(rsControl!医生姓名)
                    obj出诊记录.限号数 = Val(Nvl(rsControl!限号数))
                    obj出诊记录.限约数 = Val(Nvl(rsControl!限约数))
                    obj出诊记录.已挂数 = 0
                    obj出诊记录.已约数 = 0
                    obj出诊记录.预约控制 = Val(Nvl(rsControl!预约控制))
                    Set obj分诊诊室集 = New 分诊诊室集
                    obj分诊诊室集.分诊方式 = Val(Nvl(rsControl!分诊方式))
                    obj分诊诊室集.医生姓名 = Nvl(rsControl!医生姓名)
                    '1.加载诊室
                    rsRoom.Filter = "限制ID=" & Val(Nvl(rsControl!id))
                    If rsRoom.RecordCount <> 0 Then rsRoom.MoveFirst
                    Do While Not rsRoom.EOF
                        Set obj分诊诊室 = New 分诊诊室
                        obj分诊诊室.诊室ID = Val(Nvl(rsRoom!诊室ID))
                        obj分诊诊室.诊室名称 = Nvl(rsRoom!门诊诊室)
                        
                        obj分诊诊室集.AddItem obj分诊诊室, "K" & obj分诊诊室.诊室ID
                        
                        rsRoom.MoveNext
                    Loop
                   Set obj出诊记录.安排门诊诊室集 = obj分诊诊室集
                   
                   '2.加载号序信息集
                    Set obj号序信息集 = New 号序信息集
                    rsNum.Filter = "限制ID=" & Val(Nvl(rsControl!id))
                    If rsNum.RecordCount <> 0 Then rsNum.MoveFirst
                    Do While Not rsNum.EOF
                        Set obj号序信息 = New 号序信息
                        obj号序信息.序号 = Val(Nvl(rsNum!序号))
                        obj号序信息.开始时间 = Format(rsNum!开始时间, "yyyy-mm-dd HH:MM:SS")
                        obj号序信息.终止时间 = Format(rsNum!终止时间, "yyyy-mm-dd HH:MM:SS")
                        obj号序信息.是否预约 = Val(Nvl(rsNum!是否预约)) = 1
                        obj号序信息.数量 = Val(Nvl(rsNum!限制数量))
                       
                        obj号序信息集.AddItem obj号序信息
                        rsNum.MoveNext
                    Loop
                    'Set obj号序信息集.上班时段 = obj上班时段
                    obj号序信息集.出诊频次 = Val(Nvl(rsControl!出诊频次))
                    obj号序信息集.时间段 = obj出诊记录.时间段
                    obj号序信息集.是否分时段 = obj出诊记录.是否分时段
                    obj号序信息集.是否序号控制 = obj出诊记录.是否序号控制
                    obj号序信息集.限号数 = obj出诊记录.限号数
                    obj号序信息集.限约数 = obj出诊记录.限约数
                    obj号序信息集.预约控制 = obj出诊记录.预约控制
                    
                   Set obj出诊记录.号序信息集 = obj号序信息集
                   '3.加载三方单位控制
                    Set obj合作单位控制集 = New 合作单位控制集
                    obj合作单位控制集.是否独占 = Val(Nvl(rsControl!是否独占))
                    Set obj号序信息集 = Nothing
                    strTemp = ""
                    
                    
                    rsUnitControl.Filter = "限制ID=" & Val(Nvl(rsControl!id))
                    rsUnitControl.Sort = "类型,性质,名称,序号"
                    If rsUnitControl.RecordCount <> 0 Then
                        rsUnitControl.MoveFirst
                        obj合作单位控制集.预约控制方式 = Val(Nvl(rsUnitControl!控制方式))
                    End If
                    Do While Not rsUnitControl.EOF
                        If strTemp <> Nvl(rsUnitControl!类型) & "-" & Nvl(rsUnitControl!性质) & "-" & Nvl(rsUnitControl!名称) Then
                            If Not obj号序信息集 Is Nothing Then
                                Set obj合作单位控制.号序信息集 = obj号序信息集
                                obj合作单位控制集.AddItem obj合作单位控制, "K" & obj合作单位控制.合作单位名称
                            End If
                            Set obj合作单位控制 = New 合作单位控制
                            obj合作单位控制.合作单位名称 = Nvl(rsUnitControl!名称)
                            obj合作单位控制.类型 = Val(Nvl(rsUnitControl!类型))
                            obj合作单位控制.预约控制方式 = Val(Nvl(rsUnitControl!控制方式))
                            Set obj号序信息集 = New 号序信息集
                            
                            strTemp = Nvl(rsUnitControl!类型) & "-" & Nvl(rsUnitControl!性质) & "-" & Nvl(rsUnitControl!名称)
                        End If
                        Set obj号序信息 = New 号序信息
                        obj号序信息.序号 = Val(Nvl(rsUnitControl!序号))
                        
                        obj号序信息.开始时间 = Format(Nvl(rsUnitControl!开始时间), "yyyy-mm-dd HH:MM:SS")
                        obj号序信息.终止时间 = Format(Nvl(rsUnitControl!终止时间), "yyyy-mm-dd HH:MM:SS")
                        obj号序信息.数量 = Val(Nvl(rsUnitControl!数量))
                        obj号序信息.是否预约 = 1 '都是预约用
                        obj号序信息集.AddItem obj号序信息
                        rsUnitControl.MoveNext
                    Loop
                    If Not obj号序信息集 Is Nothing Then
                        Set obj合作单位控制.号序信息集 = obj号序信息集
                        obj合作单位控制集.AddItem obj合作单位控制, "K" & obj合作单位控制.合作单位名称
                        
                    End If
                   
                    Set obj出诊记录.合作单位控制集 = obj合作单位控制集
                   obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
                End If
            End If
            .MoveNext
        Loop
        obj出诊记录集.出诊日期 = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    End With
    
    Set GetClinicRecordFromSignalSource = obj出诊记录集
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function GetWorkTimeRec() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取上班时间段的记录集
    '入参:
    '返回:上班时间段记录集
    '编制:刘兴洪
    '日期:2016-03-22 16:18:06
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
    strSQL = "Select 时间段, To_Char(Sysdate, 'yyyy-mm-dd') || ' ' || To_Char(开始时间, 'HH24:MI:SS') As 开始时间," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(开始时间, 'hh24:mi:ss') >= To_Char(终止时间, 'hh24:mi:ss') Then 1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(终止时间, 'HH24:MI:SS') As 终止时间," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(开始时间, 'hh24:mi:ss') > To_Char(Nvl(缺省时间,开始时间), 'hh24:mi:ss') Then  1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(Nvl(缺省时间,开始时间), 'HH24:MI:SS') As 缺省时间," & vbNewLine & _
            "        To_Char(Sysdate + Case When To_Char(开始时间, 'hh24:mi:ss') < To_Char(Nvl(提前时间,开始时间), 'hh24:mi:ss') Then -1 Else 0 End, 'yyyy-mm-dd') || ' ' || To_Char(Nvl(提前时间,开始时间), 'HH24:MI:SS') As 提前时间," & vbNewLine & _
            "        Nvl(站点, '-') As 站点, Nvl(号类, '-') As 号类, 出诊预留时间, 休息时段" & vbNewLine & _
            " From 时间段"

    If grsWorkTime Is Nothing Then
        Set grsWorkTime = zlDatabase.OpenSQLRecord(strSQL, "获取上班时段")
    ElseIf grsWorkTime.State <> adStateOpen Then
        Set grsWorkTime = zlDatabase.OpenSQLRecord(strSQL, "获取上班时段")
    End If
    
    Set GetWorkTimeRec = grsWorkTime
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function GetUnitAll() As ADODB.Recordset
    '功能：获取所有挂号合作单位和预约方式
    '入参：
    '   strStationNo:站点编号
    '   strSignalType:号类
    Dim strSQL As String
    
    On Error GoTo errHandler
    strSQL = "Select 1 As 类型, 名称 From 挂号合作单位" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 2 As 类型, 名称 From 预约方式"
    If grsUnit Is Nothing Then
        Set grsUnit = zlDatabase.OpenSQLRecord(strSQL, "获取所有挂号合作单位和预约方式")
    ElseIf grsUnit.State <> adStateOpen Then
        Set grsUnit = zlDatabase.OpenSQLRecord(strSQL, "获取所有挂号合作单位和预约方式")
    End If
    grsUnit.MoveFirst
    Set GetUnitAll = grsUnit
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDoctorRooms(ByVal lng科室ID As Long) As ADODB.Recordset
    '功能：根据适用科室ID获取门诊诊室
    '入参：
    '   lng科室ID:科室ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    strSQL = "Select a.Id As 诊室ID, a.编码, a.名称, b.缺省标志" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = [1]" & vbNewLine & _
            "       And (a.站点 Is Null Or a.站点=(Select 站点 From 部门表 Where id = [1]))" & vbNewLine & _
            " Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊诊室", lng科室ID)
    
    If rsTemp.RecordCount = 0 Then
        strSQL = "Select a.Id As 诊室ID, a.编码, a.名称, 0 As 缺省标志" & vbNewLine & _
                " From 门诊诊室 A" & vbNewLine & _
                " Where a.站点 Is Null Or a.站点=(Select 站点 From 部门表 Where id = [1])" & vbNewLine & _
                " Order By 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊诊室", lng科室ID)
    End If
    
    Set GetDoctorRooms = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSignalSource(Optional ByVal str排班方式 As String, _
    Optional ByVal lng号源Id As Long) As ADODB.Recordset
    '功能：获取临床出诊号源
    '入参：
    '   str排班方式:0-固定排班;1-按月排班;2-按周排班 多个用逗号分隔
    '   lng号源ID:临床出诊号源.ID
    Dim strSQL As String, strWhere As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    If str排班方式 <> "" Then strWhere = " And Instr([1], a.排班方式) > 0"
    If lng号源Id <> 0 Then strWhere = " And a.ID=[2]"
    strSQL = "Select a.Id, a.号类, a.号码, a.科室id, b.名称 As 科室名称, a.项目id, c.名称 As 项目名称, a.医生id, a.医生姓名," & vbNewLine & _
            "        d.专业技术职务 As 医生职称, a.是否建病案, a.预约天数, a.出诊频次, a.假日控制状态, a.是否假日换休," & vbNewLine & _
            "        a.是否临床排班, a.排班方式, a.是否删除, a.建档时间, a.撤档时间, b.站点" & vbNewLine & _
            " From 临床出诊号源 A, 部门表 B, 收费项目目录 C, 人员表 D" & vbNewLine & _
            " Where a.科室id = b.Id And a.项目id = c.Id And a.医生ID = d.ID(+)" & strWhere & vbNewLine & _
            "       And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(b.站点, [4]), Nvl([3], '-')) = Nvl([3], '-')" & vbNewLine & _
            " Order By a.号码 asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取临床出诊号源", str排班方式, lng号源Id, gstrNodeNo, gVisitPlan_ModulePara.str号源维护站点)
    
    Set GetSignalSource = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetNextPlanDate(frmParent As Object, ByVal byt排班方式 As Byte, _
    ByRef intYear As Integer, ByRef intMonth As Integer, Optional ByRef intWeek As Integer, _
    Optional ByVal lng人员ID As Long, Optional ByVal blnShowSelect As Boolean = True) As Boolean
    '功能：确定下一个排班的年月周
    '参数：
    '   byt排班方式：1-按月排班;2-按周排班
    '   lng人员ID：无"所有科室"权限时传入
    '返回：数组，0-年份，1-月份，2-周数
    '说明：由人员ID确定号源时，通过所拥有号源所在最后一个出诊表来确定时间
    Dim strSQL As String, rsData As ADODB.Recordset
    Dim dtStart As Date, dtEnd As Date, dtCur As Date
    Dim blnFind As Boolean, strWhere As String
    Dim intYearTemp As Integer, intMonthTemp As Integer, intWeekTemp As Integer
    
    Err = 0: On Error GoTo errHandler
    intYear = 0: intMonth = 0: intWeek = 0
'    If lng人员ID <> 0 Then
'        strWhere = "Exists" & vbNewLine & _
'                " (Select 1" & vbNewLine & _
'                "       From 临床出诊安排 M, 临床出诊号源 N, 部门表 P" & vbNewLine & _
'                "       Where m.出诊id = a.Id And m.号源id = n.Id And n.科室id = p.Id" & vbNewLine & _
'                "             And Nvl(n.是否临床排班, 0) = 1 And Nvl(n.排班方式, 0) = [1]" & vbNewLine & _
'                "             And Nvl(n.是否删除, 0) = 0 And (n.撤档时间 Is Null Or n.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
'                "             And (p.站点 = '" & gstrNodeNo & "' Or p.站点 Is Null)" & vbNewLine & _
'                "             And Exists (Select 1 From 部门人员 Where 部门id = n.科室id And 人员id = [2]))"
'        '必须要大于最后一个已发布的出诊表
'        strWhere = " And (" & strWhere & " Or a.发布时间 Is Not Null)"
'    End If
    '排班方式：0-固定排班;1-按月排班;2-按周排班;3-模板
    strSQL = "Select a.年份, a.月份, a.周数" & vbNewLine & _
            " From 临床出诊表 A" & vbNewLine & _
            " Where a.排班方式=[1] " & strWhere & vbNewLine & _
            "       And Nvl(a.站点,'-') = Nvl([3],'-')" & vbNewLine & _
            " Order By a.年份 Desc, a.月份 Desc, a.周数 Desc"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "获取大于当前时间的年月周", byt排班方式, lng人员ID, gstrNodeNo)
    
    dtStart = zlDatabase.Currentdate()
    If rsData.EOF Then
        '数据库中无出诊表，由用户确定出诊表的日期
        GoTo SetNewDate:
    Else
        intYear = Val(Nvl(rsData!年份))
        intMonth = Val(Nvl(rsData!月份))
        If byt排班方式 = 2 Then intWeek = Val(Nvl(rsData!周数))
    End If
        
    If byt排班方式 = 1 Then
        If intMonth = 12 Then '当前为本年的最后一个月,则下月为下一年的1月
            intYear = intYear + 1: intMonth = 1
        Else
            intMonth = intMonth + 1
        End If
    Else
        If GetWeekCount(intYear, intMonth) = intWeek Then '当前为本月的最后一周,则下周为下月的第一周
            If intMonth = 12 Then '当前为本年的最后一个月,则下月为下一年的1月
                intYear = intYear + 1: intMonth = 1
            Else
                intMonth = intMonth + 1
            End If
            intWeek = 1
        Else
            intWeek = intWeek + 1
        End If
    End If
    
    '小于当前时间时，由用户确定出诊表的日期
    intYearTemp = Year(dtStart): intMonthTemp = Month(dtStart): intWeekTemp = GetDateWeek(dtStart)
    If intYear < intYearTemp Then
        GoTo SetNewDate:
    ElseIf intYear = intYearTemp Then
        If intMonth < intMonthTemp Then
            GoTo SetNewDate:
        ElseIf intMonth = intMonthTemp Then
            If intWeek < intWeekTemp And byt排班方式 = 2 Then
                GoTo SetNewDate:
            End If
        End If
    End If
    GetNextPlanDate = True
    Exit Function
    
SetNewDate:
    intYear = 0: intMonth = 0: intWeek = 0
    '由用户确定出诊表的日期
    If blnShowSelect Then
        Dim frm As New frmClinicSetNewPlanDate
        If frm.ShowMe(frmParent, byt排班方式, dtStart, intYear, intMonth, intWeek) = False Then Exit Function
    End If
    GetNextPlanDate = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetWorkTimes(Optional ByVal strStationNo As String, _
    Optional ByVal strSignalType As String) As ADODB.Recordset
    '功能：获取上班时段
    '入参：
    '   strStationNo:站点编号
    '   strSignalType:号类
    Dim rsTemp As New ADODB.Recordset, strFilter As String
    Dim str时间段 As String, lngCount As Long

    On Error GoTo errHandler
    '过滤条件
    strFilter = "(站点='-' And 号类='-')"
    If strStationNo <> "" Then
        strFilter = strFilter & " OR (站点='" & strStationNo & "' And 号类='-')"
    End If
    If strSignalType <> "" Then
        strFilter = strFilter & " OR (站点='-' And 号类='" & strSignalType & "')"
    End If
    If strStationNo <> "" And strSignalType <> "" Then
        strFilter = strFilter & " OR (站点='" & strStationNo & "' And 号类='" & strSignalType & "')"
    End If
    
    Set rsTemp = GetWorkTimeRec() '取缓存数据
    rsTemp.Filter = strFilter

    Set GetWorkTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitPlan(ByVal lng安排ID As Long, Optional ByVal lng出诊ID As Long) As ADODB.Recordset
    '功能：获取临床出诊安排
    '入参：
    '   lng安排ID:安排ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If lng安排ID = 0 Then
        strSQL = "Select b.Id As 出诊id, b.排班方式, b.出诊表名, b.年份, b.月份, b.周数, b.应用范围, b.科室id, b.备注, b.发布人, b.发布时间," & vbNewLine & _
                "        b.模板类型,Null As 安排id, Null As 号源id, Null As 项目id, Null As 项目名称, Null As 医生id," & _
                "        Null As 医生姓名, Null As 医生职称, Null As 排班规则, Null As 是否周六出诊, Null As 是否周日出诊," & vbNewLine & _
                "        null As 开始时间, null As 终止时间, Null As 操作员姓名, Null As 登记时间, Null As 是否临时安排" & vbNewLine & _
                " From 临床出诊表 B" & vbNewLine & _
                " Where b.Id = [2] And Rownum < 2"
    Else
        strSQL = "Select b.Id As 出诊id, b.排班方式, b.出诊表名, b.年份, b.月份, b.周数," & vbNewLine & _
                "        b.应用范围, b.科室id, b.备注, b.发布人, b.发布时间,b.模板类型," & vbNewLine & _
                "        a.Id As 安排id, a.号源id, a.项目id, c.名称 As 项目名称, a.医生id, a.医生姓名,d.专业技术职务 As 医生职称," & vbNewLine & _
                "        a.排班规则, a.是否周六出诊, a.是否周日出诊, a.开始时间, a.终止时间, a.操作员姓名, a.登记时间, a.是否临时安排" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊表 B, 收费项目目录 C, 人员表 D" & vbNewLine & _
                " Where a.出诊id = b.Id And a.项目ID = c.ID And a.医生ID = d.ID(+) And a.Id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取安排数据", lng安排ID, lng出诊ID)
    Set GetVisitPlan = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitItems(ByVal lng安排ID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取临床出诊限制或记录项目
    '入参：
    '   lng安排ID:安排ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select Distinct To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊记录 B" & vbNewLine & _
                " Where a.Id = b.安排id And a.Id = [1] And b.上班时段 Is Not Null"
    Else
        strSQL = "Select Distinct b.限制项目 As 出诊日期" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊限制 B" & vbNewLine & _
                " Where a.Id = b.安排id And a.Id = [1] And b.上班时段 Is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊项目", lng安排ID)
    Set GetVisitItems = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimes(ByVal lng安排ID As Long, Optional ByVal str项目 As String, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取临床出诊时段
    '入参：
    '   lng安排ID:安排ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select b.ID As 记录ID,To_Char(b.出诊日期,'yyyy-mm-dd') As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, b.开始时间, b.终止时间," & vbNewLine & _
                "        b.限号数, b.已挂数, b.限约数, b.已约数, b.分诊方式, b.预约控制, b.替诊医生姓名, b.科室ID" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊记录 B" & vbNewLine & _
                " Where a.Id = b.安排id And a.Id = [1] And b.上班时段 Is Not Null"
        If str项目 <> "" Then
            strSQL = strSQL & " And b.出诊日期=to_date([2],'yyyy-mm-dd')"
        End If
    Else
        strSQL = "Select b.ID as 记录ID,b.限制项目 As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, NULL as 开始时间, NULL as 终止时间," & vbNewLine & _
                "        b.限号数, 0 as 已挂数, b.限约数, 0 as 已约数, b.分诊方式, b.预约控制, '' As 替诊医生姓名,0 As 科室ID" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊限制 B" & vbNewLine & _
                " Where a.Id = b.安排id(+) And a.Id = [1] And b.上班时段 Is Not Null"
        If str项目 <> "" Then
            strSQL = strSQL & " And b.限制项目=[2]"
        End If
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", lng安排ID, str项目)
    Set GetVisitTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPreviousVisitTimes(ByVal lng号源Id As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取上次有效临床出诊时段
    '入参：
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select 记录id, 出诊日期, 上班时段, 是否分时段, 是否序号控制, 开始时间, 终止时间, 限号数," & vbNewLine & _
                "        已挂数, 限约数, 已约数, 分诊方式, 预约控制, 替诊医生姓名, 科室ID" & vbNewLine & _
                " From (Select b.Id As 记录id, b.出诊日期, b.上班时段, b.是否分时段, b.是否序号控制," & vbNewLine & _
                "              b.开始时间, b.终止时间, b.限号数, b.已挂数, b.限约数, b.已约数, b.分诊方式, b.预约控制," & vbNewLine & _
                "              b.替诊医生姓名, b.科室ID, Row_Number() Over(Partition By b.上班时段 Order By b.Id Desc) As 组号" & vbNewLine & _
                "        From 临床出诊安排 A, 临床出诊记录 B, 临床出诊表 C" & vbNewLine & _
                "        Where a.Id = b.安排id And a.出诊id = c.Id And a.号源ID=[1] And c.发布时间 Is Not Null)" & vbNewLine & _
                " Where 组号 = 1"
    Else
        strSQL = "Select 记录id, 出诊日期, 上班时段, 是否分时段, 是否序号控制, 开始时间, 终止时间, 限号数," & vbNewLine & _
                "        已挂数, 限约数, 已约数, 分诊方式, 预约控制, 替诊医生姓名, 科室ID" & vbNewLine & _
                " From (Select b.Id As 记录id, b.限制项目 As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制," & vbNewLine & _
                "              NULL as 开始时间, NULL as 终止时间, b.限号数, 0 As 已挂数, b.限约数, 0 As 已约数, b.分诊方式, b.预约控制," & vbNewLine & _
                "              '' As 替诊医生姓名, 0 As 科室ID, Row_Number() Over(Partition By b.上班时段 Order By b.Id Desc) As 组号" & vbNewLine & _
                "        From 临床出诊安排 A, 临床出诊限制 B, 临床出诊表 C" & vbNewLine & _
                "        Where a.Id = b.安排id And a.出诊id = c.Id And a.号源ID=[1] And c.发布时间 Is Not Null)" & vbNewLine & _
                " Where 组号 = 1"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", lng号源Id)
    Set GetPreviousVisitTimes = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTime(ByVal lng记录ID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取临床出诊单个时段
    '入参：
    '   lng记录ID:记录ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select b.ID As 记录ID,b.出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, b.开始时间, b.终止时间," & vbNewLine & _
                "        b.限号数, b.已挂数, b.限约数, b.已约数, b.分诊方式, b.预约控制, b.替诊医生姓名, " & vbNewLine & _
                "        b.科室ID, b.项目ID, c.名称 As 项目名称, b.医生ID, b.医生姓名, b.是否独占, " & vbNewLine & _
                "        b.停诊开始时间, b.停诊终止时间, b.停诊原因, b.是否临时出诊" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊记录 B, 收费项目目录 C" & vbNewLine & _
                " Where a.Id = b.安排id And b.项目ID = c.ID And b.Id = [1] And b.上班时段 Is Not Null"
    Else
        strSQL = "Select b.ID as 记录ID,b.限制项目 As 出诊日期, b.上班时段, b.是否分时段, b.是否序号控制, NULL as 开始时间, NULL as 终止时间," & vbNewLine & _
                "        b.限号数, 0 as 已挂数, b.限约数, 0 as 已约数, b.分诊方式, b.预约控制, NULL As 替诊医生姓名," & vbNewLine & _
                "        0 As 科室ID, 0 As 项目ID, '' As 项目名称, 0 As 医生ID, '' As 医生姓名, 0 As 是否独占, " & vbNewLine & _
                "        NULL As 停诊开始时间, NULL As 停诊终止时间, NULL As 停诊原因, 0 As 是否临时出诊" & vbNewLine & _
                " From 临床出诊安排 A, 临床出诊限制 B" & vbNewLine & _
                " Where a.Id = b.安排id(+) And b.Id = [1] And b.上班时段 Is Not Null"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取出诊时段", lng记录ID)
    Set GetVisitTime = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitRooms(ByVal lngID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取临床出诊诊室
    '入参：
    '   lngID:记录ID/限制ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandler
    If blnRecord Then
        strSQL = "Select a.诊室ID, b.名称" & vbNewLine & _
                " From 临床出诊诊室记录 A, 门诊诊室 B" & vbNewLine & _
                " Where a.诊室id = b.Id And a.记录ID = [1]"
    Else
        strSQL = "Select a.诊室ID, b.名称" & vbNewLine & _
                " From 临床出诊诊室 A, 门诊诊室 B" & vbNewLine & _
                " Where a.诊室id = b.Id And a.限制id = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取临床出诊诊室", lngID)
    Set GetVisitRooms = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeInterval(ByVal lngID As Long, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取号序信息
    '入参：
    '   lngID:记录ID/限制ID
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        '"       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _'开始时间与终止时间相等的是加号的序号
        strSQL = "Select a.序号, a.开始时间, a. 终止时间, a.数量, a.是否预约, a.是否停诊" & vbNewLine & _
                " From 临床出诊序号控制 A,临床出诊记录 B" & vbNewLine & _
                " Where a.记录ID =b.ID And b.ID=[1] " & vbNewLine & _
                "       And (a.开始时间 <> a. 终止时间 Or a.开始时间 Is Null And a. 终止时间 Is Null)" & vbNewLine & _
                "       And (Not(Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0) Or Nvl(b.是否分时段,0)=1 And Nvl(b.是否序号控制,0)=0 And a.预约顺序号 IS NULL)"
    Else
        strSQL = "Select a.序号, a.开始时间, a. 终止时间, a.限制数量  As 数量, a.是否预约, 0 As 是否停诊" & vbNewLine & _
                " From 临床出诊时段 A" & vbNewLine & _
                " Where a.限制ID = [1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取号序信息", lngID)
    Set GetTimeInterval = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetUnitReg(ByVal lngID As Long, ByVal str合作单位 As String, ByVal byt类型 As Byte, Optional ByVal blnRecord As Boolean) As ADODB.Recordset
    '功能：获取合作单位号序信息
    '入参：
    '   lngID:记录ID/限制ID
    '   str合作单位:合作单位名称
    '   byt类型:类型，0-合作单位，1-预约方式
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    If blnRecord Then
        '"       And (b.开始时间 <> b. 终止时间 Or b.开始时间 Is Null And b. 终止时间 Is Null)" & vbNewLine & _'开始时间与终止时间相等的是加号的序号
        strSQL = "Select a.控制方式, a.序号, b.开始时间, b.终止时间, a.数量, b.是否预约, b.是否停诊" & vbNewLine & _
                " From 临床出诊挂号控制记录 A, 临床出诊序号控制 B,临床出诊记录 C" & vbNewLine & _
                " Where a.记录id = b.记录id(+) And a.序号 = b.序号(+)" & vbNewLine & _
                "       And (b.开始时间 <> b. 终止时间 Or b.开始时间 Is Null And b. 终止时间 Is Null)" & vbNewLine & _
                "       And a.记录id = c.ID And c.ID=[1] And a.名称 = [2] and Nvl(a.类型,0) = [3]" & vbNewLine & _
                "       And (Not(Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0) Or Nvl(c.是否分时段,0)=1 And Nvl(c.是否序号控制,0)=0 And b.预约顺序号 IS NULL)"
    Else
        strSQL = "Select a.控制方式, a.序号, b.开始时间, b.终止时间, a.数量, b.是否预约, 0 As 是否停诊" & vbNewLine & _
                " From 临床出诊挂号控制 A, 临床出诊时段 B" & vbNewLine & _
                " Where a.限制ID = b.限制ID(+) And a.序号 = b.序号(+) " & vbNewLine & _
                "       And a.限制ID = [1] And a.名称 = [2] and Nvl(a.类型,0) = [3]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取合作单位挂号控制", lngID, str合作单位, byt类型)

    Set GetUnitReg = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetStopVisit(ByVal lng号源Id As Long, ByVal dt开始时间 As Date, dt终止时间 As Date, _
    Optional ByVal blnAllHoliday As Boolean = True) As ADODB.Recordset
    '功能：获取停诊记录（节假日、停诊安排）
    '入参：
    '   lng号源ID:号源ID
    '   dt开始时间、dt终止时间:查询时间范围
    '   blnAllHoliday:所有假日，不根据“假日控制状态”判断
    Dim strSQL As String, rsTemp As New ADODB.Recordset

    On Error GoTo errHandler
    '停诊安排
    strSQL = "Select 1 As 类型, a.开始时间, a.终止时间, a.停诊原因" & vbNewLine & _
            " From 临床出诊停诊记录 A, 临床出诊号源 B" & vbNewLine & _
            " Where a.申请人 = b.医生姓名 And a.记录id Is Null And a.审批时间 Is Not Null And a.取消人 Is Null" & vbNewLine & _
            "       And b.Id = [1] And b.医生id Is Not Null And Not (a.开始时间 > [3] Or a.终止时间 < [2])" & vbNewLine
    '节假日(全部，这里不根据"临床出诊号源.假日控制状态"判断，在发布安排时处理)
    strSQL = strSQL & _
            " Union All" & vbNewLine & _
            " Select 2 As 类型, 开始日期, 终止日期, 节日名称" & vbNewLine & _
            " From 法定假日表" & vbNewLine & _
            " Where 性质 = 0" & vbNewLine & _
            "       And Not (开始日期 > [3] Or 终止日期 < [2])"
    If blnAllHoliday = False Then
        strSQL = strSQL & vbNewLine & _
                "   And Exists (Select 1 From 临床出诊号源 Where ID = [1] And Nvl(假日控制状态, 0) = 0)"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取停诊记录", lng号源Id, dt开始时间, dt终止时间)

    Set GetStopVisit = rsTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get出诊记录(ByVal lng号源Id As Long, ByVal lng记录ID As Long, ByVal blnRecord As Boolean, _
    ByRef obj出诊号源 As 出诊号源, ByRef obj出诊记录 As 出诊记录) As Boolean
    Dim rsSignalSource As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim obj合作单位 As 合作单位控制, obj合作单位控制 As 合作单位控制
    Dim obj所有分诊诊室集 As 分诊诊室集
    Dim obj所有上班时段集 As 上班时段集, obj出诊记录集 As 出诊记录集
    Dim obj所有合作单位 As 合作单位控制集
    
    Err = 0: On Error GoTo errHandler
    '号源信息,科室、医生、项目取安排中的
    Set rsSignalSource = GetSignalSource("", lng号源Id)
    If rsSignalSource.RecordCount = 0 Then
        MsgBox "未发现号源信息，请刷新数据后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    Set obj出诊号源 = GetSignalSourceObject(rsSignalSource)
    If obj出诊号源 Is Nothing Then
        MsgBox "获取号源信息出错，请重试！", vbInformation, gstrSysName
        Exit Function
    End If
    
     Set obj所有上班时段集 = GetWorkTimesObjects(GetWorkTimes(obj出诊号源.站点, obj出诊号源.号类))
    '出诊记录
    Set obj出诊记录集 = GetVisitTimesObjects(GetVisitTime(lng记录ID, blnRecord))
    If obj出诊记录集.Count = 0 Then
        MsgBox "获取出诊记录出错，请重试！", vbInformation, gstrSysName: Exit Function
    End If
    
    Set obj出诊记录 = obj出诊记录集(1)
    With obj出诊记录
        '号源信息,科室、医生、项目取安排中的
        obj出诊号源.项目ID = .项目ID
        obj出诊号源.项目名称 = .项目名称
        obj出诊号源.医生ID = .医生ID
        obj出诊号源.医生姓名 = .医生姓名
    
        If obj所有上班时段集.Exits("K" & .时间段) Then
            Set .上班时段 = obj所有上班时段集("K" & .时间段).Clone
        Else
            Set .上班时段 = New 上班时段
        End If
        
        '出诊诊室
        Set .安排门诊诊室集 = GetVisitRoomsObjects(GetVisitRooms(.记录ID, blnRecord))
        .安排门诊诊室集.分诊方式 = .分诊方式
        .安排门诊诊室集.医生姓名 = obj出诊号源.医生姓名
        
        '号序信息
        Set .号序信息集 = GetTimeIntervalObjects(GetTimeInterval(.记录ID, blnRecord))
        
        .号序信息集.出诊频次 = obj出诊号源.出诊频次
        .号序信息集.是否分时段 = .是否分时段
        .号序信息集.是否序号控制 = .是否序号控制
        .号序信息集.限号数 = .限号数
        .号序信息集.限约数 = .限约数
        .号序信息集.预约控制 = .预约控制
        .号序信息集.时间段 = .时间段
        
        '合作单位控制
        Set .合作单位控制集 = New 合作单位控制集
        .合作单位控制集.是否独占 = .是否独占
        Set obj所有合作单位 = GetUnitsObjects(GetUnitAll())
        For Each obj合作单位 In obj所有合作单位
            Set rsTemp = GetUnitReg(.记录ID, obj合作单位.合作单位名称, obj合作单位.类型, blnRecord)
            If Not rsTemp.EOF Then
                Set obj合作单位控制 = New 合作单位控制
                obj合作单位控制.类型 = obj合作单位.类型
                obj合作单位控制.合作单位名称 = obj合作单位.合作单位名称
                obj合作单位控制.预约控制方式 = Val(Nvl(rsTemp!控制方式))
                Set obj合作单位控制.号序信息集 = GetTimeIntervalObjects(rsTemp)
                .合作单位控制集.AddItem obj合作单位控制
            End If
        Next
    End With
    Get出诊记录 = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Get预约天数(ByVal lng出诊ID As Long, Optional lng号源Id As Long) As Integer
    '固定安排最大预约天数
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '1.号源的最大预约天数
    If lng号源Id = 0 Then
        strSQL = "Select Nvl(Max(b.预约天数), 0) As 预约天数" & vbNewLine & _
                "        From 临床出诊安排 A, 临床出诊号源 B" & vbNewLine & _
                "        Where a.号源id = b.Id And a.出诊id = [1]" & vbNewLine
    Else
        strSQL = "Select Nvl(Max(b.预约天数), 0) As 预约天数" & vbNewLine & _
                "        From 临床出诊号源 B" & vbNewLine & _
                "        Where B.id = [2]" & vbNewLine
    End If
    '2.预约方式的最大预约天数
    strSQL = strSQL & _
            " Union All" & vbNewLine & _
            " Select Max(预约天数) As 预约天数 From 预约方式" & vbNewLine
    
    strSQL = "Select 预约天数 From (" & strSQL & ") Where 预约天数 > 0 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预约天数", lng出诊ID, lng号源Id)
    If Not rsTemp.EOF Then
        Get预约天数 = Val(Nvl(rsTemp!预约天数))
    End If
    
    '3.系统参数"挂号允许预约天数"
    If Get预约天数 = 0 Then Get预约天数 = Val(zlDatabase.GetPara(Val("66-挂号允许预约天数"), glngSys))
    '4.缺省七天
    If Get预约天数 = 0 Then Get预约天数 = 7
    
    '104266
    '以半天为单位,如果参数“号源开放时间”在12:00:00-23:59:59期间的，则开放预约天数+1天
    strSQL = "Select Zl_Fun_Getappointmentdays As 加预约天数 From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断此刻是否多开放一天")
    If Not rsTemp.EOF Then
        Get预约天数 = Get预约天数 + Val(Nvl(rsTemp!加预约天数))
    End If
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetVisitedRecord(ByVal lng号源Id As Long, _
    ByVal str开始时间 As String, ByVal str终止时间 As String) As ADODB.Recordset
    '获取当前号源已设置了的出诊日期
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    '排除按天设置安排的月模板
    strSQL = "Select b.出诊ID,b.ID as 安排ID,b.号源ID,a.出诊日期" & vbNewLine & _
            " From 临床出诊记录 A,临床出诊安排 B,临床出诊表 C" & vbNewLine & _
            " Where a.安排ID=b.ID And a.号源ID=[1] And a.出诊日期 Between [2] And [3]" & vbNewLine & _
            "       And c.ID=b.出诊ID And Nvl(c.排班方式,0) In (0,1,2)"
    Set GetVisitedRecord = zlDatabase.OpenSQLRecord(strSQL, "获取号源已设置了的出诊记录", lng号源Id, _
        CDate(str开始时间), CDate(str终止时间))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitedRecordByDate(ByVal lng安排ID As Long, ByVal str出诊日期 As String) As ADODB.Recordset
    '获取当前安排在指定日期的出诊记录
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    strSQL = "Select a.Id, a.上班时段," & vbNewLine & _
            "        Max(Decode(b.Id, Null, 0, 1)) As 已使用," & vbNewLine & _
            "        Max(Decode(a.停诊开始时间, Null, 0, 1)) As 已停诊" & vbNewLine & _
            " From 临床出诊记录 A, 病人挂号记录 B" & vbNewLine & _
            " Where a.Id = b.出诊记录id(+) And a.安排id = [1] And a.出诊日期 = [2]" & vbNewLine & _
            " Group By a.Id, a.上班时段"
    Set GetVisitedRecordByDate = _
        zlDatabase.OpenSQLRecord(strSQL, "获取当前安排已用于挂号预约的出诊记录", lng安排ID, CDate(str出诊日期))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckExistRecord(ByVal lng号源Id As Long, ByVal strApply As String, _
    Optional ByVal obj出诊安排 As 出诊安排, Optional ByVal blnMonthTemplet As Boolean, _
    Optional ByVal lng安排ID As Long) As Boolean
    '检查被应用的日期内是否已有出诊记录
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim ObjItem As 出诊记录集
    Dim varDate As Variant, i As Integer
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    If strApply = "" Then Exit Function
    If blnMonthTemplet Then
        strSQL = "Select /*+cardinality(b,10)*/ 1" & _
                " From 临床出诊限制 A, Table(f_Str2list([2], '|')) B" & _
                " Where a.安排ID = [1] And a.限制项目 = b.Column_Value And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查被应用的日期内是否已有出诊记录", lng安排ID, strApply)
        CheckExistRecord = Not rsTemp.EOF
        Exit Function
    End If
    
    If lng号源Id <> 0 Then
        strSQL = "Select /*+cardinality(b,10)*/ 1" & _
                " From 临床出诊记录 A, Table(f_Str2list([2], '|')) B" & _
                " Where a.号源ID = [1] And a.出诊日期 = To_Date(b.Column_Value, 'yyyy-mm-dd') And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查被应用的日期内是否已有出诊记录", lng号源Id, strApply)
        CheckExistRecord = Not rsTemp.EOF
        Exit Function
    End If
    
    If obj出诊安排 Is Nothing Then Exit Function
    varDate = Split(strApply, "|")
    For i = 0 To UBound(varDate)
        If blnFind Then Exit For
        If Not obj出诊安排.已保存出诊安排 Is Nothing Then
            For Each ObjItem In obj出诊安排.已保存出诊安排
                If obj出诊安排(1).出诊日期 <> ObjItem.出诊日期 Then
                    If IsDate(varDate(i)) Then
                        If DateDiff("d", ObjItem.出诊日期, varDate(i)) = 0 Then blnFind = True: Exit For
                    Else
                        If ObjItem.出诊日期 = varDate(i) Then blnFind = True: Exit For
                    End If
                End If
            Next
        End If
        If blnFind Then Exit For
        If Not obj出诊安排.未保存出诊安排 Is Nothing Then
            For Each ObjItem In obj出诊安排.未保存出诊安排
                If obj出诊安排(1).出诊日期 <> ObjItem.出诊日期 Then
                    If IsDate(varDate(i)) Then
                        If DateDiff("d", ObjItem.出诊日期, varDate(i)) = 0 Then blnFind = True: Exit For
                    Else
                        If ObjItem.出诊日期 = varDate(i) Then blnFind = True: Exit For
                    End If
                End If
            Next
        End If
    Next
    CheckExistRecord = blnFind
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function SearchStopVisitReason(frmFrom As Object, objControl As Object, ByVal strInput As String) As String
    '功能:模糊查找，弹出停诊原因选择列表
    '参数:
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput = "" Then Exit Function
    vRect = zlControl.GetControlRect(objControl.Hwnd)
    '去掉"'"
    strInput = Replace(strInput, "'", " ")
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If IsNumeric(strInput) Then '输入全是数字时只匹配编码
            strWhere = " Where 编码 Like Upper([1])"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then '输入全是字母时只匹配简码
            strWhere = " Where 简码 Like Upper([1])"
        Else
            strWhere = " Where 编码 Like Upper([1]) Or 名称 Like [1] Or 简码 Like Upper([1])"
        End If
    End If
    
    strSQL = "Select RowNum As ID, 编码, 名称 From 常用停诊原因" & strWhere
    Set rsTemp = zlDatabase.ShowSQLSelect(frmFrom, strSQL, 0, "停诊原因", False, _
                   "", "", False, False, True, vRect.Left, vRect.Top, objControl.Height, blnCancel, True, False, strKey)
                   
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> adStateOpen Then Exit Function
    
    SearchStopVisitReason = Nvl(rsTemp!名称)
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get预约挂号记录(ByVal lng号源Id As Long, _
    ByVal dtStartTime As Date, ByVal dtEndTime As Date) As ADODB.Recordset
    '获取指定号源在指定时间范围内的预约挂号数据，相同星期的只获取一条记录
    Dim strSQL As String
    
    On Error GoTo errHandler
    strSQL = "Select Decode(To_Char(出诊日期, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) As 限制项目," & vbNewLine & _
            "        ID As 记录id, 出诊日期, 上班时段, 开始时间, 终止时间, 是否独占" & vbNewLine & _
            " From (Select a.Id, a.出诊日期, a.上班时段, a.开始时间, a.终止时间, 是否独占," & vbNewLine & _
            "               Row_Number() Over(Partition By To_Char(a.出诊日期, 'D'), a.上班时段 Order By a.出诊日期) As 行号" & vbNewLine & _
            "        From 临床出诊记录 A, 病人挂号记录 B" & vbNewLine & _
            "        Where a.Id = b.出诊记录id And a.上班时段 Is Not Null And a.号源id = [1] And a.出诊日期 Between [2] And [3])" & vbNewLine & _
            " Where 行号 < 2" & vbNewLine & _
            " Order By To_Char(出诊日期, 'D'), To_Date(To_Char(开始时间, 'hh24:mi:ss'), 'hh24:mi:ss')"
    Set Get预约挂号记录 = zlDatabase.OpenSQLRecord(strSQL, "获取有预约挂号数据的出诊记录", lng号源Id, _
        CDate(Format(dtStartTime, "yyyy-mm-dd")), CDate(Format(dtEndTime, "yyyy-mm-dd")))
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
