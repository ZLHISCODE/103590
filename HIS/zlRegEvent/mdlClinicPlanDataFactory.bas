Attribute VB_Name = "mdlClinicPlanDataFactory"
Option Explicit

Public Function GetSignalSourceObject(ByVal rsRecord As ADODB.Recordset) As 出诊号源
    '功能:将记录集的当前记录转换为"出诊号源"对象
    '入参：
    '   rsRecord - 号源记录集
    Dim obj出诊号源 As New 出诊号源
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj出诊号源
            .ID = Val(Nvl(rsRecord!ID))
            .号类 = Nvl(rsRecord!号类)
            .号码 = Nvl(rsRecord!号码)
            .科室ID = Val(Nvl(rsRecord!科室ID))
            .科室名称 = Nvl(rsRecord!科室名称)
            .项目ID = Val(Nvl(rsRecord!项目ID))
            .项目名称 = Nvl(rsRecord!项目名称)
            .医生ID = Val(Nvl(rsRecord!医生ID))
            .医生姓名 = Nvl(rsRecord!医生姓名)
            .医生职称 = Nvl(rsRecord!医生职称)
            .是否建病案 = Val(Nvl(rsRecord!是否建病案)) = 1
            .预约天数 = Val(Nvl(rsRecord!预约天数))
            .出诊频次 = Val(Nvl(rsRecord!出诊频次))
            .假日控制状态 = Val(Nvl(rsRecord!假日控制状态))
            .是否临床排班 = Val(Nvl(rsRecord!是否临床排班)) = 1
            .排班方式 = Val(Nvl(rsRecord!排班方式))
            .是否删除 = Val(Nvl(rsRecord!是否删除)) = 1
            .建档时间 = Format(Nvl(rsRecord!建档时间), "yyyy-mm-dd hh:mm:ss")
            .撤档时间 = Format(Nvl(rsRecord!撤档时间), "yyyy-mm-dd hh:mm:ss")
            .站点 = Nvl(rsRecord!站点)
        End With
    End If
    Set GetSignalSourceObject = obj出诊号源
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetWorkTimesObjects(ByVal rsRecord As ADODB.Recordset) As 上班时段集
    '功能:将记录集转换为"上班时段集"对象
    '入参：
    '   rsRecord - 上班时段记录
    Dim obj上班时段集 As New 上班时段集, obj上班时段 As 上班时段
    Err = 0: On Error GoTo errHandler
        
    '特殊处理
    '取相同时段的第一个
    rsRecord.Sort = "站点 Desc,号类 Desc"
    If rsRecord.RecordCount > 0 Then rsRecord.MoveFirst
    Do While Not rsRecord.EOF
        If obj上班时段集.Exits("K" & Nvl(rsRecord!时间段)) = False Then
            Set obj上班时段 = New 上班时段
            With obj上班时段
                .时间段 = Nvl(rsRecord!时间段)
                .开始时间 = Format(Nvl(rsRecord!开始时间), "yyyy-mm-dd hh:mm:ss")
                .结束时间 = Format(Nvl(rsRecord!终止时间), "yyyy-mm-dd hh:mm:ss")
                .缺省预约时间 = Format(Nvl(rsRecord!缺省时间), "yyyy-mm-dd hh:mm:ss")
                .提前挂号时间 = Format(Nvl(rsRecord!提前时间), "yyyy-mm-dd hh:mm:ss")
                .出诊预留时间 = Val(Nvl(rsRecord!出诊预留时间))
                .休息时段 = Nvl(rsRecord!休息时段)
            End With
            obj上班时段集.AddItem obj上班时段, "K" & Nvl(rsRecord!时间段)
        End If
        rsRecord.MoveNext
    Loop
    Set GetWorkTimesObjects = obj上班时段集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetUnitsObjects(ByVal rsRecord As ADODB.Recordset) As 合作单位控制集
    '功能:将记录集转换为"合作单位控制集"对象
    '入参：
    '   rsRecord - 合作单位控制记录
    Dim obj合作单位控制集 As New 合作单位控制集, obj合作单位控制 As 合作单位控制
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj合作单位控制 = New 合作单位控制
        With obj合作单位控制
            .类型 = Nvl(rsRecord!类型)
            .合作单位名称 = Nvl(rsRecord!名称)
        End With
        obj合作单位控制集.AddItem obj合作单位控制, "K" & Nvl(rsRecord!名称)
        rsRecord.MoveNext
    Loop
    Set GetUnitsObjects = obj合作单位控制集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetVisitPlanObjects(ByVal rsRecord As ADODB.Recordset) As 出诊安排
    '功能:将记录集转换为"出诊安排"对象
    '入参
    Dim obj出诊安排 As New 出诊安排
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj出诊安排
            .出诊ID = Val(Nvl(rsRecord!出诊ID))
            .出诊表名 = Nvl(rsRecord!医生姓名)
            .排班方式 = Val(Nvl(rsRecord!排班方式))
            .年份 = Val(Nvl(rsRecord!年份))
            .月份 = Val(Nvl(rsRecord!月份))
            .周数 = Val(Nvl(rsRecord!周数))
            .应用范围 = Val(Nvl(rsRecord!应用范围))
            .科室ID = Val(Nvl(rsRecord!科室ID))
            .备注 = Nvl(rsRecord!备注)
            .发布人 = Nvl(rsRecord!发布人)
            .发布时间 = Format(Nvl(rsRecord!发布时间), "yyyy-mm-dd hh:mm:ss")
            .模板类型 = Val(Nvl(rsRecord!模板类型))
            
            .安排ID = Val(Nvl(rsRecord!安排ID))
            .项目ID = Val(Nvl(rsRecord!项目ID))
            .项目名称 = Nvl(rsRecord!项目名称)
            .医生ID = Val(Nvl(rsRecord!医生ID))
            .医生姓名 = Nvl(rsRecord!医生姓名)
            .医生职称 = Nvl(rsRecord!医生职称)
            .排班规则 = Val(Nvl(rsRecord!排班规则))
            .周六不出诊 = Val(Nvl(rsRecord!是否周六出诊)) = 0
            .周日不出诊 = Val(Nvl(rsRecord!是否周日出诊)) = 0
            .开始时间 = Format(Nvl(rsRecord!开始时间), "yyyy-mm-dd hh:mm:ss")
            .终止时间 = Format(Nvl(rsRecord!终止时间), "yyyy-mm-dd hh:mm:ss")
            .操作员姓名 = Nvl(rsRecord!操作员姓名)
            .登记时间 = Format(Nvl(rsRecord!登记时间), "yyyy-mm-dd hh:mm:ss")
            .是否临时安排 = Val(Nvl(rsRecord!是否临时安排)) = 1
        End With
    End If
    Set GetVisitPlanObjects = obj出诊安排
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimesObjects(ByVal rsRecord As ADODB.Recordset) As 出诊记录集
    '功能:将记录集转换为"出诊记录集"对象
    '入参：
    Dim obj出诊记录集 As New 出诊记录集, obj出诊记录 As 出诊记录
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj出诊记录 = GetVisitTimesObject(rsRecord)
        obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
        rsRecord.MoveNext
    Loop
    Set GetVisitTimesObjects = obj出诊记录集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitTimesObject(ByVal rsRecord As ADODB.Recordset) As 出诊记录
    '功能:将记录集转换为"出诊记录集"对象
    '入参：
    Dim obj出诊记录 As New 出诊记录
    
    Err = 0: On Error GoTo errHandler
    If Not rsRecord.EOF Then
        With obj出诊记录
            .记录ID = Val(Nvl(rsRecord!记录ID))
            .时间段 = Nvl(rsRecord!上班时段)
            .是否分时段 = Val(Nvl(rsRecord!是否分时段)) = 1
            .是否序号控制 = Val(Nvl(rsRecord!是否序号控制)) = 1
            .限号数 = Val(Nvl(rsRecord!限号数))
            '预约控制：0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约;
            If Val(Nvl(rsRecord!预约控制)) = 1 Then
                .限约数 = 0
            Else
                .限约数 = IIf(Val(Nvl(rsRecord!限约数)) = 0, Val(Nvl(rsRecord!限号数)), Val(Nvl(rsRecord!限约数)))
            End If
            .预约控制 = Val(Nvl(rsRecord!预约控制))
            .分诊方式 = Val(Nvl(rsRecord!分诊方式))
            .预约控制 = Val(Nvl(rsRecord!预约控制))
            .出诊日期 = Format(Nvl(rsRecord!出诊日期), "yyyy-mm-dd hh:mm:ss")
            .已挂数 = Val(Nvl(rsRecord!已挂数))
            .已约数 = Val(Nvl(rsRecord!已约数))
            .替诊医生 = Nvl(rsRecord!替诊医生姓名)
            .科室ID = Val(Nvl(rsRecord!科室ID))
            .项目ID = Val(Nvl(rsRecord!项目ID))
            .项目名称 = Nvl(rsRecord!项目名称)
            .医生ID = Val(Nvl(rsRecord!医生ID))
            .医生姓名 = Nvl(rsRecord!医生姓名)
            .开始时间 = Format(Nvl(rsRecord!开始时间), "yyyy-mm-dd hh:mm:ss")
            .终止时间 = Format(Nvl(rsRecord!终止时间), "yyyy-mm-dd hh:mm:ss")
            .是否独占 = Val(Nvl(rsRecord!是否独占)) = 1
            .停诊开始时间 = Format(Nvl(rsRecord!停诊开始时间), "yyyy-mm-dd hh:mm:ss")
            .停诊终止时间 = Format(Nvl(rsRecord!停诊终止时间), "yyyy-mm-dd hh:mm:ss")
            .停诊原因 = Nvl(rsRecord!停诊原因)
            .是否临时出诊 = Val(Nvl(rsRecord!是否临时出诊)) = 1
        End With
    End If
    Set GetVisitTimesObject = obj出诊记录
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVisitRoomsObjects(ByVal rsRecord As ADODB.Recordset) As 分诊诊室集
    '功能:将记录集转换为"分诊诊室集"对象
    '入参：
    Dim obj分诊诊室集 As New 分诊诊室集, obj分诊诊室 As 分诊诊室
    
    Err = 0: On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj分诊诊室 = New 分诊诊室
        With obj分诊诊室
            .诊室ID = Nvl(rsRecord!诊室ID)
            .诊室名称 = Nvl(rsRecord!名称)
        End With
        obj分诊诊室集.AddItem obj分诊诊室, "K" & Nvl(rsRecord!诊室ID)
        rsRecord.MoveNext
    Loop
    Set GetVisitRoomsObjects = obj分诊诊室集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeIntervalObjects(ByVal rsRecord As ADODB.Recordset) As 号序信息集
    '功能:将记录集转换为"号序信息集"对象
    '入参：
    Dim obj号序信息集 As New 号序信息集, obj号序信息 As 号序信息

    On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj号序信息 = GetTimeIntervalObject(rsRecord)
        If Not obj号序信息 Is Nothing Then
            obj号序信息集.AddItem obj号序信息
        End If
        rsRecord.MoveNext
    Loop
    Set GetTimeIntervalObjects = obj号序信息集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetTimeIntervalObject(ByVal rsRecord As ADODB.Recordset) As 号序信息
    '功能:将记录集转换为"号序信息"对象
    '入参：
    Dim obj号序信息 As 号序信息

    On Error GoTo errHandler
    If rsRecord.EOF Then Exit Function
    Set obj号序信息 = New 号序信息
    With obj号序信息
        .序号 = Nvl(rsRecord!序号)
        .开始时间 = Format(Nvl(rsRecord!开始时间), "yyyy-mm-dd hh:mm:ss")
        .终止时间 = Format(Nvl(rsRecord!终止时间), "yyyy-mm-dd hh:mm:ss")
        .数量 = Val(Nvl(rsRecord!数量))
        .是否预约 = Val(Nvl(rsRecord!是否预约)) = 1
        .是否停诊 = Val(Nvl(rsRecord!是否停诊)) = 1
    End With
    Set GetTimeIntervalObject = obj号序信息
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ChangeCurPlan(obj出诊安排 As 出诊安排, ByVal strNewItem As String, _
    Optional ByVal blnRuleChanged As Boolean)
    '当前出诊项目已改变，调整未保存出诊安排集合，并提取当前安排
    Dim ObjItem As 出诊记录集, strKey As String
    Dim objTemp As 出诊记录集
    
    On Error GoTo Errhand
    If obj出诊安排.未保存出诊安排 Is Nothing Then Set obj出诊安排.未保存出诊安排 = New 出诊安排
    If obj出诊安排.已保存出诊安排 Is Nothing Then Set obj出诊安排.已保存出诊安排 = New 出诊安排
    '模板时，规则变化时要清空未保存安排
    If blnRuleChanged Then
        obj出诊安排.未保存出诊安排.RemoveAll
    Else
        For Each ObjItem In obj出诊安排
            Set objTemp = ObjItem.Clone
            If objTemp.出诊日期 <> "" Then  '无项目的不保存
                strKey = GetPlanKey(objTemp.出诊日期)
                
                '在未保存记录集中存在，则先删除
                If obj出诊安排.未保存出诊安排.Exits(strKey) Then obj出诊安排.未保存出诊安排.Remove strKey
                
                '一个时段都没有的不保存
                If objTemp.Count = 0 Then
                    '如果在已保存记录集中存在，但本次无时段表示已删除
                    If obj出诊安排.已保存出诊安排.Exits(strKey) Then
                        obj出诊安排.已保存出诊安排(strKey).是否删除 = True
                    End If
                ElseIf objTemp.是否修改 Then '未修改的不保存
                    obj出诊安排.未保存出诊安排.AddItem objTemp, strKey
                End If
            End If
        Next
    End If
    
    '切换当前出诊安排
    obj出诊安排.RemoveAll
    strKey = GetPlanKey(strNewItem)
    If obj出诊安排.未保存出诊安排.Exits(strKey) Then
        obj出诊安排.AddItem obj出诊安排.未保存出诊安排(strKey).Clone, strKey
    ElseIf obj出诊安排.已保存出诊安排.Exits(strKey) Then
        If obj出诊安排.已保存出诊安排(strKey).是否删除 Then
            Set ObjItem = New 出诊记录集
            ObjItem.出诊日期 = strNewItem
            obj出诊安排.AddItem ObjItem, strKey
        Else
            obj出诊安排.AddItem obj出诊安排.已保存出诊安排(strKey).Clone, strKey
        End If
    ElseIf obj出诊安排.已保存出诊安排.Count > 0 _
        And obj出诊安排.已保存出诊安排.排班规则 = obj出诊安排.排班规则 _
        And InStr("4,5", obj出诊安排.排班规则) > 0 Then
        '不改变安排，只改变出诊项目
        Set ObjItem = obj出诊安排.已保存出诊安排(1).Clone
        ObjItem.出诊日期 = strNewItem
        obj出诊安排.AddItem ObjItem, strKey
    Else
        Set ObjItem = New 出诊记录集
        ObjItem.出诊日期 = strNewItem
        obj出诊安排.AddItem ObjItem, strKey
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetStopVisitObjects(ByVal rsRecord As ADODB.Recordset) As 停诊记录集
    '功能:将停诊记录集转换为"停诊记录集"对象
    '入参：
    Dim obj停诊记录集 As New 停诊记录集, obj停诊记录 As 停诊记录

    On Error GoTo errHandler
    Do While Not rsRecord.EOF
        Set obj停诊记录 = New 停诊记录
        With obj停诊记录
            .开始时间 = Format(Nvl(rsRecord!开始时间), "yyyy-mm-dd hh:mm:ss")
            .终止时间 = Format(Nvl(rsRecord!终止时间), "yyyy-mm-dd hh:mm:ss")
            .停诊原因 = Nvl(rsRecord!停诊原因)
            .类型 = Nvl(rsRecord!类型)
        
            obj停诊记录集.AddItem obj停诊记录
        End With
        rsRecord.MoveNext
    Loop
    Set GetStopVisitObjects = obj停诊记录集
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
