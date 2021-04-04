Attribute VB_Name = "mdlClinicPlanFun"
Option Explicit
'挂号安排功能清单
Public Enum RegistPlanFun
    Pane_FunFace = 1
    Pane_Face = 2
    
    Pane_WorkTime = 11 '工作时间管理
    Pane_Holiday = 12 '节假日管理
    Pane_DoctorOffice = 13 '门诊诊室设置
    Pane_SignalSource = 14 '号源管理置
    
    Pane_StopPlan = 21 '停诊管理
    Pane_FixedPlan = 22 '固定安排
    Pane_PlanTemplet = 23 '安排模板
    Pane_MonthPlan = 24 '月安排
    Pane_WeekPlan = 25 '周安排
    Pane_MonthTemplet = 26 '按天设置出诊安排的月模板
End Enum

Public Sub ZlUpdatePlanMenu(frmParent As Object, cbsMain As Object, Optional ByVal bytFun As Byte, Optional ByVal lng人员ID As Long)
    '设置菜单名称
    Dim cbrControl As CommandBarControl
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim strMonthMenu As String, strWeekMenu As String
    
    On Error Resume Next
    If cbsMain Is Nothing Then Exit Sub
    If GetNextPlanDate(frmParent, 1, intYear, intMonth, intWeek, lng人员ID, False) = False Then Exit Sub
    If intMonth = 0 Then
        strMonthMenu = "生成月出诊表"
    Else
        strMonthMenu = "生成" & intMonth & "月出诊表"
    End If
    If bytFun = 1 Then
        Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu & "(&N)"
        Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu
    End If
    
    Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextMonthNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu & "(&N)"
    Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextMonthNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strMonthMenu
    
    If GetNextPlanDate(frmParent, 2, intYear, intMonth, intWeek, lng人员ID, False) = False Then Exit Sub
    If intWeek = 0 Then
        strWeekMenu = "生成周出诊表"
    Else
        strWeekMenu = "生成" & intMonth & "月第" & intWeek & "周出诊表"
    End If
    If bytFun <> 1 Then
        Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu & "(&N)"
        Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextNewPlan, , True)
        If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu
    End If
    
    Set cbrControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_Edit_NextWeekNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu & "(&W)"
    Set cbrControl = cbsMain(2).Controls.Find(, conMenu_Edit_NextWeekNewPlan, , True)
    If Not cbrControl Is Nothing Then cbrControl.Caption = strWeekMenu
    cbsMain.RecalcLayout
End Sub

Public Function CheckIsHavePlan(ByVal byt排班方式 As Byte, ByVal lngUserID As Long, _
    Optional ByVal dt开始时间 As Date, Optional ByVal dt终止时间 As Date, _
    Optional ByVal blnDeleteFixedPlan As Boolean) As Boolean
    '检查当前操作员是否有可操作的号源
    '入参：
    '   byt排班方式 - 0:固定排班,1:月排班,2:周排班,3:模板
    '   lngUserID - 用户ID
    '   dt开始时间、dt终止时间 -  新出诊表的时间范围
    '   blnDeleteFixedPlan - 是否删除无挂号预约的出诊记录
    Dim strWhere As String, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lngUserID > 0 Then
        strWhere = " And Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [2])"
    End If
    Select Case byt排班方式
    Case 0 '固定出诊表
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D" & vbNewLine & _
                " Where a.科室ID = b.ID And a.医生ID = c.ID(+) And a.项目ID = d.ID And a.排班方式 = 0 And Nvl(a.是否删除, 0) = 0" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
                "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
                "       And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null)" & vbNewLine & _
                "       And Nvl(Nvl(b.站点, [3]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                "       And Not Exists(Select 1 From 临床出诊安排 P,临床出诊表 Q" & vbNewLine & _
                "                      Where p.出诊ID = q.ID And p.号源ID = a.ID And q.排班方式 = 0)" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源", gstrNodeNo, lngUserID, gVisitPlan_ModulePara.str号源维护站点)
    Case 1, 2
        If byt排班方式 = 1 Then
            strWhere = " And 排班方式 = 1" & vbNewLine & strWhere
        Else
            '1.当前号源为周排班且在出诊表时间范围内无出诊记录
            '2.当前号源已由周排班调整为了月排班，调整前已在出诊表所在月有了周排班，则出诊表所在月剩下的部分将继续按周进行排班
            strWhere = " And (a.排班方式 = 2 And Not Exists (Select 1" & vbNewLine & _
                    "           From 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                    "           Where p.出诊id = q.Id And p.号源id = a.Id" & vbNewLine & _
                    "               And Not(p.终止时间<[5] Or p.开始时间>Last_Day([3])) And q.排班方式 = 1)" & vbNewLine & _
                    "       Or a.排班方式 = 1 And Exists (Select 1" & vbNewLine & _
                    "           From 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                    "           Where p.出诊id = q.Id And p.号源id = a.Id" & vbNewLine & _
                    "               And Not(p.终止时间<[5] Or p.开始时间>Last_Day([3])) And q.排班方式 = 2))" & vbNewLine & _
                    strWhere
        End If
        
        '号源在该出诊表时间范围内无出诊记录
        strWhere = _
                "  And Not Exists" & vbNewLine & _
                "       (Select 1" & vbNewLine & _
                "        From 临床出诊记录 O, 临床出诊安排 P, 临床出诊表 Q" & vbNewLine & _
                "        Where o.安排id = p.Id And p.出诊id = q.Id And p.号源id = a.Id" & _
                "              And o.出诊日期 Between [3] And [4] " & vbNewLine & _
                         IIf(blnDeleteFixedPlan, _
                "              And (q.排班方式 In (1, 2) Or q.排班方式 = 0 And (Nvl(o.已挂数, 0) <> 0 Or Nvl(o.已约数, 0) <> 0))", "") & ")" & vbNewLine & _
                strWhere
        
        strSQL = "Select 1" & vbNewLine & _
            " From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D" & vbNewLine & _
            " Where a.科室ID = b.ID And a.医生ID = c.ID(+) And a.项目ID = d.ID And Nvl(a.是否删除, 0) = 0" & vbNewLine & _
            "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
            "       And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
            "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
            "       And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null)" & vbNewLine & _
            "       And Nvl(Nvl(b.站点, [6]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                    strWhere & vbNewLine & _
            "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源", gstrNodeNo, lngUserID, dt开始时间, dt终止时间, _
            CDate(Format(dt开始时间, "yyyy-mm-01")), gVisitPlan_ModulePara.str号源维护站点)
    Case 3 '模板
        strSQL = "Select 1" & vbNewLine & _
                " From 临床出诊号源 A, 部门表 B, 人员表 C, 收费项目目录 D" & vbNewLine & _
                " Where a.科室ID = b.ID And a.医生ID = c.ID(+) And a.项目ID = d.ID And a.排班方式 In (1, 2) And Nvl(a.是否删除, 0) = 0" & vbNewLine & _
                "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                "       And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null)" & vbNewLine & _
                "       And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)" & vbNewLine & _
                "       And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null)" & vbNewLine & _
                "       And Nvl(Nvl(b.站点, [3]), Nvl([1], '-')) = Nvl([1], '-')" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源", gstrNodeNo, lngUserID, gVisitPlan_ModulePara.str号源维护站点)
    End Select
    CheckIsHavePlan = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckExistsFixedToOthers(ByVal byt排班方式 As Byte, ByVal lngUserID As Long, _
     ByVal dt开始时间 As Date, ByRef strDelInfo As String, ByRef strNotDelInfo As String) As Boolean
    '检查满足排班方式的号源是否有由固定排班转换过来的
    '入参：
    '   byt排班方式 - 1:月排班,2:周排班
    '   lngUserID - 用户ID
    '   dt开始时间 -  新出诊表的开始时间
    '出错：
    '   strDelInfo - 可删除出诊记录用于当前排班方式的号源,格式：号码-科室(医生姓名) + vbCrLf + 号码-科室(医生姓名) + vbCrLf + ...
    '   strNotDelInfo - 不可删除出诊记录不能用于当前排班方式的号源,格式：号码-科室(医生姓名) + vbCrLf + 号码-科室(医生姓名) + vbCrLf + ...
    Dim strWhere As String, strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strDelInfo = "": strNotDelInfo = ""
    If lngUserID > 0 Then
        strWhere = " And Nvl(d.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = d.科室id And 人员id = [4])"
    End If
    strSQL = "Select Max(Decode(e.Id, Null, 0, 1)) As 类型, d.号码, Max(f.名称) As 科室, Max(d.医生姓名) As 医生姓名" & vbNewLine & _
            " From 临床出诊记录 A, 临床出诊安排 B, 临床出诊表 C, 临床出诊号源 D, 病人挂号记录 E, 部门表 F" & vbNewLine & _
            " Where a.安排id = b.Id And b.出诊id = c.Id And b.号源id = d.Id And a.Id = e.出诊记录id(+) And d.科室id = f.Id And c.排班方式 = 0" & vbNewLine & _
            "       And a.出诊日期 >= [3] And Nvl(d.是否删除, 0) = 0" & vbNewLine & _
            "       And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And Nvl(d.排班方式, 0) = [2]" & vbNewLine & _
            "       And Nvl(Nvl(f.站点,[5]),Nvl([1],'-')) = Nvl([1],'-')" & vbNewLine & _
                    strWhere & _
            " Group By d.号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查号源", gstrNodeNo, byt排班方式, dt开始时间, lngUserID, gVisitPlan_ModulePara.str号源维护站点)
    Do While Not rsTemp.EOF
        If Val(Nvl(rsTemp!类型)) = 0 Then
            strDelInfo = strDelInfo & vbCrLf & "  " & Nvl(rsTemp!号码) & "-" & Nvl(rsTemp!科室) & IIf(Nvl(rsTemp!医生姓名) = "", "", "(" & Nvl(rsTemp!医生姓名) & ")")
        End If
        If Val(Nvl(rsTemp!类型)) = 1 Then
            strNotDelInfo = strNotDelInfo & vbCrLf & "  " & Nvl(rsTemp!号码) & "-" & Nvl(rsTemp!科室) & IIf(Nvl(rsTemp!医生姓名) = "", "", "(" & Nvl(rsTemp!医生姓名) & ")")
        End If
        rsTemp.MoveNext
    Loop
    If strDelInfo <> "" Then strDelInfo = Mid(strDelInfo, 4)
    If strNotDelInfo <> "" Then strNotDelInfo = Mid(strNotDelInfo, 4)
    CheckExistsFixedToOthers = Not (strDelInfo = "" And strNotDelInfo = "")
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlDeletePlan(ByVal lng出诊ID As Long, Optional ByVal lngUserID As Long) As Boolean
    '功能：删除出诊表
    '入参：
    '   lngUserID - 用户ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    
    'Zl_临床出诊表_Delete
    strSQL = "Zl_临床出诊表_Delete("
    '  Id_In       临床出诊表.Id%Type
    strSQL = strSQL & "" & lng出诊ID & ","
    '  人员id_In 人员表.Id%Type := Null
    strSQL = strSQL & "" & lngUserID & ","
    '  站点_In   部门表.站点%Type
    strSQL = strSQL & "'" & gstrNodeNo & "')"
    zlDatabase.ExecuteProcedure strSQL, "删除出诊表"
    
    ZlDeletePlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlClearPlan(ByVal lng安排ID As Long, ByVal strItem As String, _
    Optional ByVal blnRecord As Boolean) As Boolean
    '功能：清除某一天的安排
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If lng安排ID = 0 Or strItem = "" Then Exit Function
    
    'Zl_临床出诊上班时段_Delete(
    strSQL = "Zl_临床出诊上班时段_Delete("
    '安排id_In   临床出诊限制.安排id%Type,
    strSQL = strSQL & "" & lng安排ID & ","
    '项目_In     临床出诊限制.限制项目%Type,
    strSQL = strSQL & "'" & strItem & "',"
    '出诊记录_In Number := 0,
    strSQL = strSQL & "" & IIf(blnRecord, 1, 0) & ","
    '上班时段_In     临床出诊限制.上班时段%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '删除出诊安排_In Number:=0
    strSQL = strSQL & "" & 1 & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "清除安排"
    ZlClearPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlClearPlanBatch(ByVal lng出诊ID As Long, Optional ByVal lng号源Id As Long, _
    Optional ByVal lng人员ID As Long, Optional ByVal lng安排ID As Long, Optional ByVal blnTempPlan As Boolean) As Boolean
    '功能：清除所有号源安排，或者删除某一个号源的所有安排
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function

    'Zl_临床出诊安排_BatchDelete(
    strSQL = "Zl_临床出诊安排_BatchDelete("
    '出诊id_In 临床出诊表.Id%Type,
    strSQL = strSQL & "" & lng出诊ID & ","
    '人员id_In 人员表.Id%Type := 0,
    strSQL = strSQL & "" & lng人员ID & ","
    '站点_In   部门表.站点%Type := Null,
    strSQL = strSQL & "" & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '号源id_In 临床出诊安排.号源id%Type := 0
    strSQL = strSQL & "" & lng号源Id & ","
    '安排id_In 临床出诊安排.Id%Type := 0
    strSQL = strSQL & "" & lng安排ID & ","
    '临时安排_In 临床出诊安排.是否临时安排%Type := 0
    strSQL = strSQL & "" & IIf(blnTempPlan, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, "批量清除安排"
    ZlClearPlanBatch = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlPlanApplyTo(ByVal bytType As Byte, ByVal lng原安排ID As Long, ByVal str原项目 As String, _
    ByVal lng新安排ID As Long, ByVal str新项目 As String, Optional ByVal blnTemp As Boolean) As Boolean
    '功能：应用于其它安排
    '参数：
    '   bytType 0-模板或固定出诊表,1-出诊记录
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandler
    If lng原安排ID = 0 Or lng新安排ID = 0 _
        Or str原项目 = "" Or str新项目 = "" Then Exit Function
    
    'Zl_临床出诊安排_Applyto(
    strSQL = "Zl_临床出诊安排_Applyto("
    '应用类型_In     Number,--0-模板或固定出诊表,1-出诊记录
    strSQL = strSQL & "" & bytType & ","
    '原id_In         临床出诊安排.Id%Type,
    strSQL = strSQL & "" & lng原安排ID & ","
    '原项目_In       Varchar2,
    strSQL = strSQL & "'" & str原项目 & "',"
    '新id_In         临床出诊安排.Id%Type,
    strSQL = strSQL & "" & lng新安排ID & ","
    '新项目_In       Varchar2,--应用于的项目（多个用"|"分隔）
    strSQL = strSQL & "'" & str新项目 & "',"
    '是否临时出诊_In Number:=0
    strSQL = strSQL & "" & IIf(blnTemp, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, "应用于其它安排"
    ZlPlanApplyTo = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZlBatchSNControl(ByVal lng出诊ID As Long, ByVal bytStart As Boolean, _
    Optional ByVal lng人员ID As Long) As Boolean
    '功能：全部启用序号控制或者全部取消序号控制
    '参数：
    '   bytStart True-启用,False-停用
    '   blnRecord True-出诊记录,False-出诊限制
    Dim strSQL As String

    On Error GoTo errHandler
    If lng出诊ID = 0 Then Exit Function
    'Zl_临床出诊安排_序号控制(
    strSQL = "Zl_临床出诊安排_序号控制("
    '出诊id_In   临床出诊表.Id%Type,
    strSQL = strSQL & "" & lng出诊ID & ","
    '序号控制_In 临床出诊限制.是否序号控制%Type,
    strSQL = strSQL & "" & IIf(bytStart, 1, 0) & ","
    '站点_In     部门表.站点%Type := Null,
    strSQL = strSQL & "" & IIf(gstrNodeNo = "", "NULL", "'" & gstrNodeNo & "'") & ","
    '人员id_In   人员表.Id%Type := 0
    strSQL = strSQL & "" & ZVal(lng人员ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, "设置序号控制"
    
    ZlBatchSNControl = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZlBatchLockPlan(ByVal str记录IDs As String, ByVal blnUnlock As Boolean) As Boolean
    '锁定出诊记录
    '入参：
    '   blnUnlock 是否解锁,True-解锁,False-加锁
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    If str记录IDs = "" Then Exit Function
    If blnUnlock Then
        'Zl_临床出诊记录_Batchlock
        '  -- Ids_In 批量加锁或解锁，多个用逗号分隔
        strSQL = "Zl_临床出诊记录_Batchlock("
        '  Ids_In      Varchar2,
        strSQL = strSQL & "'" & str记录IDs & "',"
        '  取消锁定_In Number:=0
        strSQL = strSQL & "" & 1 & ")"
        zlDatabase.ExecuteProcedure strSQL, "批量加锁"
    Else
        'Zl_临床出诊记录_Batchlock
        '  -- Ids_In 批量加锁或解锁，多个用逗号分隔
        strSQL = "Zl_临床出诊记录_Batchlock("
        '  Ids_In      Varchar2,
        strSQL = strSQL & "'" & str记录IDs & "')"
        '  取消锁定_In Number:=0
        zlDatabase.ExecuteProcedure strSQL, "批量解锁"
    End If
    ZlBatchLockPlan = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExistsPlanOnVisitTable(ByVal byt排班方式 As Byte, ByVal lng出诊ID As Long, _
    Optional ByVal lngUserID As Long) As Boolean
    '检查当前月/周出诊表中是否存在有效的安排
    '入参：
    '   byt排班方式 1-月排班,2-周排班
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo errHandler
    strSQL = "" & _
        " Select 1" & vbNewLine & _
        " From 临床出诊号源 A, 临床出诊安排 B, 临床出诊记录 C, 部门表 D" & vbNewLine & _
        " Where a.Id = b.号源id and b.id=c.安排id And a.科室id = d.Id And a.排班方式 =[3] And Nvl(a.是否删除, 0) = 0" & vbNewLine & _
        "       And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
        "       And (Nvl([2], 0) = 0 Or (Nvl(a.是否临床排班, 0) = 1 And Exists (Select 1 From 部门人员 Where 部门id = a.科室id And 人员id = [2])))" & vbNewLine & _
        "       And b.出诊id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前出诊表中是否存在有效的安排", lng出诊ID, lngUserID, byt排班方式)
    ExistsPlanOnVisitTable = Not rsTemp.EOF
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get周出诊表ID(ByVal intYear As Integer, ByVal intMonth As Integer, _
    ByVal intWeek As Integer) As Long
    '根据年月周获取周安排的出诊表ID
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    strSQL = "Select b.ID" & vbNewLine & _
            " From 临床出诊表 B" & vbNewLine & _
            " Where Nvl(排班方式, 0) = 2 And 年份 = [1] And 月份 = [2] And 周数 = [3] And Nvl(站点,'-') = Nvl([4],'-')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据年月周获取周安排的出诊表ID", intYear, intMonth, intWeek, gstrNodeNo)
    If Not rsTemp.EOF Then
        Get周出诊表ID = Val(Nvl(rsTemp!id))
    End If
End Function

Public Function GetNewPlanInfo(ByVal frmParent As Form, ByVal strPrivs As String, ByVal blnMonth As Boolean, _
    ByRef strCurrentPlanKey As String, ByRef blnDeletePlan As Boolean) As Collection
    '获取和检查下一个出诊表的信息
    '入参：
    '   blnMonth - 是否月出诊表
    '出参：
    '   strCurrentPlanKey - 出诊表Key值，用于定位
    '   blnDeletePlan - 是否删除以未用于挂号预约的固定出诊安排
    '返回：新出诊表信息，Array(年份,月份,周数,开始日期,结束日期)
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim varDateRange As Variant, dtStart As Date, dtEnd As Date
    Dim cllPlan As New Collection 'Array(年份,月份,周数,开始日期,结束日期)
    Dim intElseYear As Integer, intElseMonth As Integer, intElseWeek As Integer
    Dim dtStartTemp As Date, dtEndTemp As Date
    Dim strDelInfo As String, strNotDelInfo As String, strInfo As String
    
    Err = 0: On Error GoTo errHandler
    strCurrentPlanKey = "": blnDeletePlan = False
    
    If GetNextPlanDate(frmParent, IIf(blnMonth, 1, 2), intYear, intMonth, intWeek, _
        IIf(HavePrivs(strPrivs, "所有科室"), 0, UserInfo.id)) = False Then
        MsgBox "确定下一个出诊表的日期时出错，请重试！", vbInformation, gstrSysName
        Exit Function
    End If
    'XX月出诊表节点：K2_年份_月份
    'XX周出诊表节点：K3_年份_月份_周数
    If blnMonth Then
        strCurrentPlanKey = "K2_" & intYear & "_" & intMonth
    Else
        strCurrentPlanKey = "K3_" & intYear & "_" & intMonth & "_" & intWeek
    End If
    
    varDateRange = GetDateRange(intYear, intMonth, intWeek)
    dtStart = varDateRange(0): dtEnd = varDateRange(1)
    'Array(年份,月份,周数,开始日期,结束日期)
    cllPlan.Add Array(intYear, intMonth, intWeek, dtStart, dtEnd)
    
    '整周跨月的周出诊表的同步处理
    If blnMonth = False Then
        dtStartTemp = dtStart: dtEndTemp = dtEnd
        If IsDoubleMonthWeekPlan(intElseYear, intElseMonth, intElseWeek, dtStartTemp, dtEndTemp) Then
            dtStart = dtStartTemp: dtEnd = dtEndTemp
            
            varDateRange = GetDateRange(intElseYear, intElseMonth, intElseWeek)
            'Array(年份,月份,周数,开始日期,结束日期)
            cllPlan.Add Array(intElseYear, intElseMonth, intElseWeek, varDateRange(0), varDateRange(1))
        Else
            intElseYear = 0: intElseMonth = 0: intElseWeek = 0
        End If
    End If
    
    If CheckExistsFixedToOthers(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "所有科室"), 0, UserInfo.id), dtStart, strDelInfo, strNotDelInfo) Then
        strInfo = "提示：" & vbCrLf & _
                  "     当前可按" & IIf(blnMonth, "月", "周") & "排班的部分号源在时间(" & Format(dtStart, "yyyy-mm-dd") & ")以后存在固定出诊安排。"
        If strDelInfo <> "" Then
            strInfo = strInfo & vbCrLf & "以下号源可删除这部分出诊安排，然后按" & IIf(blnMonth, "月", "周") & "排班：" & vbCrLf & strDelInfo & vbCrLf & _
                "是否删除这部分出诊安排，以便其可按" & IIf(blnMonth, "月", "周") & "排班？"
            If strNotDelInfo <> "" Then
                strInfo = strInfo & vbCrLf & vbCrLf & _
                    "另：" & vbCrLf & _
                    "     以下号源因为这部分出诊安排中的部分已用于挂号预约，不能按" & IIf(blnMonth, "月", "周") & "排班：" & vbCrLf & strNotDelInfo
            End If
            blnDeletePlan = MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
        Else
            strInfo = strInfo & vbCrLf & "以下号源因为这部分出诊安排中的部分已用于挂号预约，不能按" & IIf(blnMonth, "月", "周") & "排班：" & vbCrLf & strNotDelInfo & vbCrLf & _
                "是否继续？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    '检查是否有按月/周排班的有效号源
    If cllPlan.Count = 2 Then
        'Array(年份,月份,周数,开始日期,结束日期)
        dtStart = cllPlan(2)(3): dtEnd = cllPlan(2)(4)
        If CheckIsHavePlan(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "所有科室"), 0, UserInfo.id), _
            dtStart, dtEnd, blnDeletePlan) = False Then
            cllPlan.Remove 2
        End If
    End If
    'Array(年份,月份,周数,开始日期,结束日期)
    dtStart = cllPlan(1)(3): dtEnd = cllPlan(1)(4)
    If CheckIsHavePlan(IIf(blnMonth, 1, 2), IIf(zlStr.IsHavePrivs(strPrivs, "所有科室"), 0, UserInfo.id), _
        dtStart, dtEnd, blnDeletePlan) = False Then
        cllPlan.Remove 1
    End If
    
    If cllPlan.Count = 0 Then
        MsgBox "当前无按" & IIf(blnMonth, "月", "周") & "排班的有效号源，请先到“基础设置>临床号源管理”中添加！", vbInformation, gstrSysName
        Exit Function
    End If
    Set GetNewPlanInfo = cllPlan
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

