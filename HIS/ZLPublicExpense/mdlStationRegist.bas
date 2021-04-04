Attribute VB_Name = "mdlStationRegist"
Option Explicit
'医生站挂号窗体代码

Public Function InitRegist(ByVal lngSys As Long, ByVal lngModul As Long, ByVal cnOracle As ADODB.Connection, ByVal strDbUser As String, _
                Optional objRegist As clsRegist, _
                Optional objExseSvr As clsExpenceSvr, _
                Optional objService As clsService) As Boolean
    '初始化挂号
    Dim strDept As String
    On Error GoTo errH:
    Set objRegist = New clsRegist
    If objRegist.zlInitCommon(lngSys, cnOracle, strDbUser) = False Then Exit Function
    
    Set objExseSvr = New clsExpenceSvr
    If objExseSvr.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    Set objService = New zlPublicExpense.clsService
    If objService.zlInitCommon(lngSys, lngModul, cnOracle, strDbUser) = False Then Exit Function
    
    InitRegist = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function ReadLastAppoint(ByVal lng安排ID As Long, ByVal lng计划ID As Long, _
                                ByVal lng记录ID As Long, ByVal datDay As Date, _
                                ByVal bln分时段 As Boolean, strLastAppoint As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    If lng记录ID <> 0 Then
        If bln分时段 Then
            strSQL = "Select 号序, 发生时间, 开始时间, 终止时间" & vbNewLine & _
                        "From (Select b.序号 as 号序, a.发生时间, a.登记时间, b.开始时间, b.终止时间" & vbNewLine & _
                        "       From 病人挂号记录 a, 临床出诊序号控制 b" & vbNewLine & _
                        "       Where a.出诊记录id = [1] And a.记录状态 = 1 And a.出诊记录id = b.记录id And (a.号序 = b.序号 or a.号序 = zl_To_number(b.备注)) And a.发生时间 Between [2] And [3]" & vbNewLine & _
                        "       Order By 登记时间 Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        Else
            strSQL = "Select 号序, 发生时间" & vbNewLine & _
                        "From (Select 号序, 发生时间, 登记时间" & vbNewLine & _
                        "       From 病人挂号记录" & vbNewLine & _
                        "       Where 出诊记录id = [1] And 记录状态 = 1 And 发生时间 Between [2] And [3]" & vbNewLine & _
                        "       Order By 登记时间 Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ReadLastAppoint", lng记录ID, CDate(Format(datDay, "yyyy-MM-dd")), CDate(Format(datDay, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    Else
        If bln分时段 Then
            strSQL = "Select 号序, 发生时间, 开始时间, 终止时间" & vbNewLine & _
                        "From (Select a.号序, a.发生时间, a.登记时间, c.开始时间, c.结束时间 As 终止时间" & vbNewLine & _
                        "       From 病人挂号记录 a, 挂号安排 b, 挂号安排时段 c" & vbNewLine & _
                        "       Where b.Id = [1] And a.号别 = b.号码 And a.记录状态 = 1 And b.Id = c.安排id And a.号序 = c.序号(+) And" & vbNewLine & _
                        "             c.星期(+) = Decode(To_Char([2], 'D')," & vbNewLine & _
                        "                              '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) And" & vbNewLine & _
                        "             发生时间 Between [2] And [3]" & vbNewLine & _
                        "       Order By 登记时间 Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        Else
            strSQL = "Select 号序, 发生时间" & vbNewLine & _
                        "From (Select a.号序, a.发生时间, a.登记时间" & vbNewLine & _
                        "       From 病人挂号记录 a, 挂号安排 b" & vbNewLine & _
                        "       Where b.Id = [1] And a.号别 = b.号码 And a.记录状态 = 1 And 发生时间 Between [2] And" & vbNewLine & _
                        "             [3]" & vbNewLine & _
                        "       Order By 登记时间 Desc)" & vbNewLine & _
                        "Where Rownum < 2"
        End If
        If lng计划ID <> 0 Then
            strSQL = Replace(strSQL, "挂号安排时段", "挂号计划时段")
            strSQL = Replace(strSQL, "挂号安排", "挂号安排计划")
            strSQL = Replace(strSQL, "c.安排id", "c.计划id")
        End If
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "ReadLastAppoint", IIf(lng计划ID > 0, lng计划ID, lng安排ID), CDate(Format(datDay, "yyyy-MM-dd")), CDate(Format(datDay, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    End If
    
    If Not rsTemp.EOF Then
        If bln分时段 Then
            strLastAppoint = Nvl(rsTemp!号序) & "(" & Format(Nvl(rsTemp!开始时间), "HH:MM") & "-" & Format(Nvl(rsTemp!终止时间), "HH:MM") & ")"
        Else
            strLastAppoint = Nvl(rsTemp!号序) & "(" & Format(Nvl(rsTemp!发生时间), "HH:MM") & ")"
        End If
    End If
    ReadLastAppoint = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GetSNState(ByVal bytRegistMode As Byte, lng记录ID As Long, Optional str号别 As String, Optional datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    If bytRegistMode = 0 Then
        Set GetSNState = GetSNState_Visit(lng记录ID)
    Else
        Set GetSNState = GetSNState_Normal(str号别, datThis, lngSN)
    End If
End Function

Private Function GetSNState_Normal(str号别 As String, datThis As Date, Optional lngSN As Long) As ADODB.Recordset
    Dim strSQL           As String
    Dim datStart         As Date
    Dim datEnd           As Date
    On Error GoTo errH
    datStart = CDate(Format(datThis, "yyyy-MM-dd"))
    datEnd = DateAdd("s", -1, DateAdd("d", 1, datStart))
    strSQL = "    " & vbNewLine & " Select 序号,状态,操作员姓名,Nvl(预约,0) as 预约,TO_Char(日期,'hh24:mi:ss') as 日期  "
    strSQL = strSQL & vbNewLine & " From 挂号序号状态 "
    strSQL = strSQL & vbNewLine & " Where 号码=[1]"
    strSQL = strSQL & vbNewLine & IIf(datThis = CDate(0), " And 日期 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 ", " And 日期 Between  [2] And [3]")
    strSQL = strSQL & vbNewLine & IIf(lngSN > 0, " And 序号=[4]", "")
    Set GetSNState_Normal = gobjDatabase.OpenSQLRecord(strSQL, "GetSNState_Normal", str号别, datStart, datEnd, lngSN)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function GetSNState_Visit(lng记录ID As Long) As ADODB.Recordset
    Dim strSQL           As String
    On Error GoTo errH

    strSQL = "    " & vbNewLine & " Select A.序号,A.挂号状态,A.操作员姓名,Decode(A.挂号状态,2,1,0) as 预约,To_Char(B.出诊日期,'hh24:mi:ss') as 日期  "
    strSQL = strSQL & vbNewLine & " From 临床出诊序号控制 A, 临床出诊记录 B "
    strSQL = strSQL & vbNewLine & " Where B.ID=[1] And B.ID=A.记录ID"
    Set GetSNState_Visit = gobjDatabase.OpenSQLRecord(strSQL, "GetSNState_Visit", lng记录ID)

    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetAll医生(rsDoctor As ADODB.Recordset) As Boolean
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
            " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
            " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = '医生' And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order By a.简码 Desc"
    Set rsDoctor = gobjDatabase.OpenSQLRecord(strSQL, "GetAll医生")
    GetAll医生 = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlGetActiveViewSql(ByVal bytRegistMode As Byte) As String
    Dim strSQL As String
    
    If bytRegistMode = 0 Then
        strSQL = _
        "       Select   Havedata, 安排id" & vbNewLine & _
        "       From (" & vbNewLine & _
        "               Select 1 As Havedata, b.Id As 安排id " & vbNewLine & _
        "               From 挂号安排时段 A, 挂号安排 B" & vbNewLine & _
        "               Where B.号码=[1] And A.安排id = b.ID " & _
        "                And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
        "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
        "                       And Not Exists" & vbNewLine & _
        "                     (Select 1 From 挂号安排计划 C " & vbNewLine & _
        "                         Where c.安排id = b.Id And c.审核时间 Is Not Null And [2] Between " & _
        "                               Nvl(c.生效时间, [2]) And" & _
        "                          c.失效时间)" & vbNewLine & _
        "               Union All " & vbNewLine & _
        "               Select 1 As Havedata, c.Id As 安排id" & vbNewLine & _
        "               From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C,(" & vbNewLine & _
        "                   SELECT MAX(a.生效时间 ) 生效 FROM 挂号安排计划 a,挂号安排 B  WHERE a.安排Id=b.ID AND b.号码=[1] AND a.审核时间 IS NOT NULL" & vbNewLine & _
        "             And [2] Between nvl(a.生效时间,to_date('1900-01-01','yyyy-mm-dd')) And a.失效时间" & vbNewLine & _
        "           ) D  " & vbNewLine & _
        "               Where  C.号码=[1] And c.Id = b.安排id And b.Id = a.计划id And b.生效时间=d.生效 And b.审核时间 Is Not Null" & _
        "                    And   Decode(To_Char([2], 'D'), '1', '周日', '2'," & _
        "                   '周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =a.星期 " & vbNewLine & _
        "                       And [2] Between Nvl(b.生效时间,[2]) And b.失效时间) B"
    Else
        strSQL = "Select 1 From 临床出诊记录 Where ID=[1] And Nvl(是否分时段,0)=1 "
    End If
    zlGetActiveViewSql = strSQL
End Function

Public Function zlGetTimeSnSql(ByVal bytRegistMode As Byte) As String
    Dim strSQL As String
    
    If bytRegistMode = 0 Then
        strSQL = "" & _
        " Select Distinct a.序号 As ID, A.序号,To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
        " From 挂号安排时段 A, 挂号安排 B, 挂号安排限制 C" & vbNewLine & _
        " Where a.安排id = b.Id And b.号码 = [1] And" & vbNewLine & _
        " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
        "      Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',Null) = a.星期(+)  " & _
        "      And b.Id = c.安排Id(+) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',Null) = c.限制项目(+)" & _
        "      And Not Exists (Select Count(1) From 挂号序号状态 Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like Rpad(a.序号, Length(a.序号)+ length(Nvl(c.限号数,0)), '_')) Having Count(1) - a.限制数量 >= 0) " & _
        "      And Not Exists (Select 1 From 挂号安排计划 E Where e.安排id = b.Id And e.审核时间 Is Not Null And [2] Between Nvl(e.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And e.失效时间)" & _
        "      And Not Exists (Select 1 From 合作单位安排控制 Where 安排id = b.Id And 序号 = a.序号 And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二','4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = 限制项目)"
        
        strSQL = strSQL & " Union " & _
        "Select Distinct a.序号 As ID,A.序号,To_Char(a.开始时间, 'hh24:mi') As 开始时间, To_Char(a.结束时间, 'hh24:mi') As 结束时间" & vbNewLine & _
        "From 挂号计划时段 A, 挂号安排计划 B, 挂号安排 C, 挂号计划限制 E," & vbNewLine & _
        "     (Select Max(a.生效时间) 生效" & vbNewLine & _
        "       From 挂号安排计划 A, 挂号安排 B" & vbNewLine & _
        "       Where a.安排id = b.Id And b.号码 = [1] And a.审核时间 Is Not Null And" & vbNewLine & _
        "             [2] Between Nvl(a.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & vbNewLine & _
        "             a.失效时间) D" & vbNewLine & _
        "Where a.计划id = b.Id And b.安排id = c.Id And c.号码 = [1] And b.生效时间 = d.生效 And b.审核时间 Is Not Null And" & vbNewLine & _
        " Decode(Sign(Sysdate - To_Date(To_Char([2], 'YYYY-MM-DD') || ' ' || To_Char(a.开始时间, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS')), -1, 0, 1) <> 1 And" & _
        "      [2] Between Nvl(b.生效时间, To_Date('1900-01-01', 'yyyy-mm-dd')) And b.失效时间" & vbNewLine & _
        "      And b.Id = e.计划Id(+) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六',Null) = e.限制项目(+)" & vbNewLine & _
        "      And Not Exists" & vbNewLine & _
        " (Select Count(1)" & vbNewLine & _
        "       From 挂号序号状态" & vbNewLine & _
        "       Where Trunc(日期) = [2] And 号码 = b.号码 And (序号 = a.序号 Or 序号 Like Rpad(a.序号, Length(a.序号)+ length(Nvl(e.限号数,0)), '_')) Having" & vbNewLine & _
        "        Count(1) - a.限制数量 >= 0) And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5'," & vbNewLine & _
        "                                           '周四', '6', '周五', '7', '周六', Null) = a.星期(+) And Not Exists" & vbNewLine & _
        " (Select 1 From 合作单位计划控制 Where 计划id = b.Id And 序号 = a.序号 And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = 限制项目)" & vbNewLine & _
        "Order By 开始时间"
    Else
        strSQL = "" & _
        " Select Rownum As Id, 序号, To_Char(开始时间, 'hh24') || ':00' As 时间点, To_Char(开始时间, 'hh24:mi') As 开始时间," & vbNewLine & _
        "       To_Char(终止时间, 'hh24:mi') As 结束时间, 开始时间 As 详细开始时间, 终止时间 As 详细结束时间 " & vbNewLine & _
        " From 临床出诊序号控制 A" & vbNewLine & _
        " Where 记录id = [1] And Nvl(挂号状态,0) = 0 And Nvl(是否预约,0)=1 And Trunc(开始时间) = [2] And Not Exists " & vbNewLine & _
        "(Select 1 From 临床出诊挂号控制记录 B Where b.记录id = a.记录id And b.控制方式 = 3 And a.序号 = b.序号)" & vbNewLine & _
        "Order By 详细开始时间"
    End If
    zlGetTimeSnSql = strSQL
End Function

Public Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
                                ByVal rsExpenses As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL  As String
    
    Err = 0: On Error GoTo errHand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            rsMoney.Filter = "收费类别='" & Nvl(rsItems!类别, "无") & "'"
            If rsMoney.EOF Then
                .AddNew
            End If
            !收费类别 = Nvl(rsItems!类别, "无")
            Do While Not rsIncomes.EOF
                !金额 = Val(Nvl(!金额)) + Val(Nvl(rsIncomes!实收))
                rsIncomes.MoveNext
            Loop
            .Update
            rsItems.MoveNext
        Loop
        
        If Not rsExpenses Is Nothing Then
            If rsExpenses.RecordCount > 0 Then rsExpenses.MoveFirst
            Do While Not rsExpenses.EOF
                rsMoney.Filter = "收费类别='" & Nvl(rsExpenses!类别, "无") & "'"
                If rsMoney.EOF Then
                    .AddNew
                End If
                !收费类别 = Nvl(rsExpenses!类别, "无")
                !金额 = Val(Nvl(!金额)) + Val(Nvl(rsExpenses!实收))
                .Update
                rsExpenses.MoveNext
            Loop
        End If
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CreatePublicPatient(ByVal frmMain As Form, objPubPatient As clsInterFacePatient) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建zlPublicPatient部件
    '返回:创建成功,返回True,否则返回False
    '编制:冉俊明
    '日期:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPubPatient = New clsInterFacePatient
    If objPubPatient.Init(frmMain, glngSys, glngModul, gcnOracle, gstrDBUser) = False Then Exit Function
    CreatePublicPatient = True
End Function

Public Function zlInsure_Check(ByVal str保险结算 As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查当前的医保是否需要较对
    '入参:str保险结算-保险结算
    '       strAdvance-医保返回的结算
    '出参:
    '返回:需要较对,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varTemp As Variant
    Dim varData As Variant
    Dim varData1 As Variant
    Dim varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str保险结算 <> strAdvance) Then Exit Function
    '正式结算前后,结算方式和结算金额未发生变化时不校对
    blnMedicareCheck = True
    varData = Split(str保险结算, "|")
    varData1 = Split(strAdvance, "|")
    If UBound(varData) = UBound(varData1) Then
    
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, ",")
            
            For j = 0 To UBound(varData1)
                strTmp = varData1(j)
                varTemp1 = Split(strTmp, ",")
                
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next

    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlMakeBillRecord(ByVal lng病人ID As Long, ByVal str费别 As String, ByVal bln急诊 As Boolean, _
            ByVal blnHav病历费 As Boolean, _
            ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset, _
            ByVal rsExpense As ADODB.Recordset, ByVal datDate As Date, _
            ByRef rsDetail As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据挂号收费项目，生成医保记录集明细信息(以售价单位)
    '入参: datDate:结算时间,
    '出参:rsDetail-返回的医保相关明细数据
    '      单据序号(1--n),费别,NO,序号,实际票号,结算时间,病人ID,收费类别,收据费目,计算单位,开单人,
    '      收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,保险编码,摘要,开单部门ID,
    '      执行部门ID
    '返回:医保相关数据的数据集()
    '编制:刘兴洪
    '日期:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNo As String
    Dim i As Long, j As Long, lngSort As Long
    
    Err = 0: On Error GoTo errHand:
    
    Set rsDetail = New ADODB.Recordset
    If rsItems Is Nothing Or rsIncomes Is Nothing Then
        MsgBox "请先选择挂号项目", vbInformation, gstrSysName
        Exit Function
    End If
    
    rsDetail.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "序号", adBigInt, , adFldIsNullable '问题:42961
    rsDetail.Fields.Append "实际票号", adVarChar, 20, adFldIsNullable
    rsDetail.Fields.Append "结算时间", adDBTimeStamp, , adFldIsNullable
    rsDetail.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsDetail.Fields.Append "收据费目", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsDetail.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "数量", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "单价", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "实收金额", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "统筹金额", adDouble, , adFldIsNullable
    rsDetail.Fields.Append "保险支付大类ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "是否医保", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "保险编码", adVarChar, 50, adFldIsNullable
    rsDetail.Fields.Append "摘要", adVarChar, 200, adFldIsNullable
    rsDetail.Fields.Append "是否急诊", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsDetail.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsDetail.CursorLocation = adUseClient
    rsDetail.LockType = adLockOptimistic
    rsDetail.CursorType = adOpenStatic
    rsDetail.Open
    
    If Not blnHav病历费 Then rsItems.Filter = "性质 <> 3"
    If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
            lngSort = lngSort + 1
            rsDetail.AddNew
            rsDetail!费别 = str费别
            rsDetail!序号 = lngSort
            rsDetail!结算时间 = datDate
            rsDetail!病人ID = lng病人ID
            rsDetail!收费类别 = rsItems!类别
            rsDetail!收据费目 = rsIncomes!收据费目
            rsDetail!收费细目ID = rsIncomes!项目ID
            rsDetail!计算单位 = rsItems!计算单位
            rsDetail!数量 = rsItems!数次
            rsDetail!单价 = rsIncomes!单价
            rsDetail!实收金额 = rsIncomes!实收
            rsDetail!统筹金额 = rsIncomes!统筹金额
            rsDetail!保险支付大类ID = rsItems!保险大类ID
            rsDetail!是否医保 = rsItems!保险项目否
            rsDetail!保险编码 = rsItems!保险编码
            rsDetail!摘要 = Null
            rsDetail!是否急诊 = bln急诊
            rsDetail!开单部门ID = UserInfo.部门ID
            rsDetail!执行部门ID = rsItems!执行科室ID
            rsDetail!开单人 = UserInfo.姓名
            rsDetail.Update
            rsIncomes.MoveNext
        Next
        rsItems.MoveNext
    Next
    
    If Not rsExpense Is Nothing Then
        If lngSort <> 0 Then lngSort = lngSort - 1
        '141815:李南春，2019/6/10，预结算时先定位行标
        If rsExpense.RecordCount <> 0 Then rsExpense.MoveFirst
        For i = 1 To rsExpense.RecordCount
            lngSort = lngSort + 1
            rsDetail.AddNew
            rsDetail!费别 = str费别
            rsDetail!序号 = lngSort
            rsDetail!结算时间 = datDate
            rsDetail!病人ID = lng病人ID
            rsDetail!收费类别 = rsExpense!类别
            rsDetail!收据费目 = rsExpense!收据费目
            rsDetail!收费细目ID = rsExpense!项目ID
            rsDetail!计算单位 = rsExpense!计算单位
            rsDetail!数量 = rsExpense!数次
            rsDetail!单价 = rsExpense!单价
            rsDetail!实收金额 = rsExpense!实收
            rsDetail!统筹金额 = rsExpense!统筹金额
            rsDetail!保险支付大类ID = rsExpense!保险大类ID
            rsDetail!是否医保 = rsExpense!保险项目否
            rsDetail!保险编码 = rsExpense!保险编码
            rsDetail!摘要 = Null
            rsDetail!是否急诊 = bln急诊
            rsDetail!开单部门ID = UserInfo.部门ID
            rsDetail!执行部门ID = rsExpense!执行科室ID
            rsDetail!开单人 = UserInfo.姓名
            rsDetail.Update
            rsExpense.MoveNext
        Next
    End If
    If rsDetail.RecordCount > 0 Then rsDetail.MoveFirst
    zlMakeBillRecord = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function AddPayToList(ByVal objPay As clsPayInfo, ByVal vsfPay As VSFlexGrid, _
                Optional ByVal bln异常重收 As Boolean, Optional ByVal byt支付类型 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将结算信息更新到支付列表中
    '入参:objPayInfo-结算信息
    '返回:
    '编制:李南春
    '日期:2019/1/29 9:42:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim objSubPay As clsSubPayInfo
    If objPay Is Nothing Then AddPayToList = True: Exit Function
    If objPay.支付金额 = 0 Then AddPayToList = True: Exit Function
    With vsfPay
        If objPay.Count > 0 Then
            .RemoveItem objPay.PayRow
            For Each objSubPay In objPay
                .Rows = .Rows + 1
                .RowData(objPay.PayRow) = objPay.结算性质
                .TextMatrix(.Rows - 1, .ColIndex("支付方式")) = objSubPay.结算方式
                .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(objSubPay.结算金额, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = objSubPay.结算方式
                .TextMatrix(.Rows - 1, .ColIndex("结算号码")) = objSubPay.结算号码
                .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = objPay.接口序号
                .TextMatrix(.Rows - 1, .ColIndex("消费卡")) = IIf(objPay.消费卡, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("消费卡ID")) = objPay.消费卡ID
                .TextMatrix(.Rows - 1, .ColIndex("卡号")) = objPay.卡号
                .TextMatrix(.Rows - 1, .ColIndex("关联交易ID")) = objPay.关联交易ID
                .TextMatrix(.Rows - 1, .ColIndex("修改")) = IIf(byt支付类型 = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("金额修改")) = IIf(byt支付类型 = 0, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 2, 1)
                .Cell(flexcpData, .Rows - 1, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 1, 0) '固定
                .TextMatrix(.Rows - 1, .ColIndex("独立结算")) = IIf(objPay.独立结算, 1, 0)
                If byt支付类型 = 2 Then
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                Else
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = 0
                End If
            Next
        Else
            '如果指定列无效，则插入到最后一列中
            If objPay.PayRow = 0 Or objPay.PayRow > .Rows - 1 Then
                objPay.PayRow = 0
                For lngRow = 1 To .Rows - 1
                    If Trim(.TextMatrix(lngRow, .ColIndex("支付方式"))) = "" Then
                        objPay.PayRow = lngRow: Exit For
                    End If
                Next
                If objPay.PayRow = 0 Then
                    objPay.PayRow = .Rows
                    .Rows = .Rows + 1
                End If
            ElseIf bln异常重收 Then
                .RemoveItem objPay.PayRow
                objPay.PayRow = .Rows
                .Rows = .Rows + 1
            End If
            .RowData(objPay.PayRow) = objPay.结算性质
            .TextMatrix(objPay.PayRow, .ColIndex("支付方式")) = objPay.名称
            .TextMatrix(objPay.PayRow, .ColIndex("金额")) = Format(Val(.TextMatrix(objPay.PayRow, .ColIndex("金额"))) + objPay.支付金额, "0.00")
            .TextMatrix(objPay.PayRow, .ColIndex("结算方式")) = objPay.结算方式
            .TextMatrix(objPay.PayRow, .ColIndex("结算号码")) = objPay.结算号码
            .TextMatrix(objPay.PayRow, .ColIndex("卡类别ID")) = objPay.接口序号
            .TextMatrix(objPay.PayRow, .ColIndex("消费卡")) = IIf(objPay.消费卡, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("消费卡ID")) = objPay.消费卡ID
            .TextMatrix(objPay.PayRow, .ColIndex("卡号")) = objPay.卡号
            .TextMatrix(objPay.PayRow, .ColIndex("关联交易ID")) = objPay.关联交易ID
            .TextMatrix(objPay.PayRow, .ColIndex("修改")) = IIf(byt支付类型 = 1, 0, 1)
            .TextMatrix(objPay.PayRow, .ColIndex("金额修改")) = IIf(byt支付类型 = 0, 1, 0)
            .TextMatrix(objPay.PayRow, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 2, 1)
            .Cell(flexcpData, objPay.PayRow, .ColIndex("校对标志")) = IIf(byt支付类型 = 1, 1, 0) '固定
            .TextMatrix(objPay.PayRow, .ColIndex("独立结算")) = IIf(objPay.独立结算, 1, 0)
            If byt支付类型 = 2 Then
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, objPay.PayRow, 0, objPay.PayRow, .Cols - 1) = 0
            End If
        End If
    End With
End Function

Public Function GetPayInfo(ByVal colCardPayMode As Collection, ByVal str结算方式 As String, _
                            objPayInfo As clsPayInfo) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据结算名称获取结算信息
    ' 入参 : colCardPayMode:支付信息集合，窗体加载支付方式时初始化
    '      : str结算方式 :需要获取的结算方式
    ' 出参 : objPayInfo：包括结算性质、结算性质、结算方式、接口序号、是否消费卡、支付类型
    '      : bln独立结算:结算方式是否能与其他三方卡共同结算
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2018/11/20 09:30
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    If objPayInfo Is Nothing Then Set objPayInfo = New clsPayInfo
    If str结算方式 = "" Then GetPayInfo = True: Exit Function
    
    If str结算方式 = "预交金" Or str结算方式 = "退预交" Then
        objPayInfo.名称 = str结算方式
        objPayInfo.支付类型 = Pay_AccountPay
        objPayInfo.结算性质 = 11
        GetPayInfo = True: Exit Function
    End If
    
    '优先充集合中查找,但存在一个问题，医疗卡名称和消费卡名称相同
    'colCardPayMode:名称，性质，结算方式，卡类别ID，是否消费卡，是否独立结算
    For i = 1 To colCardPayMode.Count
        If colCardPayMode(i)(0) = str结算方式 Then
            objPayInfo.名称 = str结算方式
            objPayInfo.结算性质 = Val(colCardPayMode(i)(1))
            objPayInfo.结算方式 = colCardPayMode(i)(2)
            objPayInfo.接口序号 = Val(colCardPayMode(i)(3))
            objPayInfo.消费卡 = Val(colCardPayMode(i)(4)) = 1
            objPayInfo.独立结算 = Val(colCardPayMode(i)(5)) = 1
            If objPayInfo.接口序号 > 0 Then
                If objPayInfo.消费卡 Then
                    objPayInfo.支付类型 = Pay_SquarePay
                Else
                    objPayInfo.支付类型 = Pay_ThreePay
                End If
            Else
                objPayInfo.支付类型 = Pay_CashPay
            End If
            GetPayInfo = True
            Exit Function
        End If
    Next
    ' 什么时候选中了结算方式但又是缓存中没有的
    MsgBox "无效的结算方式", vbInformation, gstrSysName
    Exit Function
    
    strSQL = "Select 性质 From 结算方式 Where 名称=[1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str结算方式)
    If Not rsTemp.EOF Then
        objPayInfo.名称 = str结算方式
        objPayInfo.结算性质 = Val(rsTemp!性质)
        objPayInfo.结算方式 = objPayInfo.名称
    Else
'        strSQL = "Select 1 As 类型, a.Id, b.名称, b.性质, A.是否独立结算" & vbNewLine & _
'                "From 医疗卡类别 a, 结算方式 b" & vbNewLine & _
'                "Where a.结算方式 = b.名称 And a.名称 = [1]" & vbNewLine & _
'                "Union" & vbNewLine & _
'                "Select 2 As 类型, c.编号 As Id, d.名称, d.性质, 0 as 是否独立结算" & vbNewLine & _
'                "From 消费卡类别目录 c, 结算方式 d" & vbNewLine & _
'                "Where c.结算方式 = d.名称 And c.名称 = [1]"
'
'        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "GetPayInfo", str结算方式)
'        If rsTemp.EOF Then
'            MsgBox str结算方式 & "是无效的结算方式，请选择其他支付方式。", vbInformation, gstrSysName
'            Exit Function
'        End If
'        objPayInfo.名称 = str结算方式
'        objPayInfo.结算性质 = Val(Nvl(rsTemp!性质))
'        objPayInfo.结算方式 = Nvl(rsTemp!名称)
'        objPayInfo.接口序号 = Val(Nvl(rsTemp!ID))
'        objPayInfo.消费卡 = Val(Nvl(rsTemp!类型)) = 2
'        objPayInfo.独立结算 = Val(Nvl(rsTemp!是否独立结算)) = 1
'        If objPayInfo.接口序号 > 0 Then
'            If objPayInfo.消费卡 Then
'                objPayInfo.支付类型 = Pay_SquarePay
'            Else
'                objPayInfo.支付类型 = Pay_ThreePay
'            End If
'        Else
'            objPayInfo.支付类型 = Pay_CashPay
'        End If
    End If
    GetPayInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGetBalanceSQLByVsf(ByVal strNo As String, lng结帐ID As Long, ByVal vsfPay As VSFlexGrid, _
                                    cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 :
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/2/26 16:29
    '---------------------------------------------------------------------------------------
    Dim dbl金额 As Double
    Dim str结算方式 As String, str结算信息 As String, strSQL As String
    Dim PayType As gPagePay
    Dim i As Long
    On Error GoTo errH
    If cllPro Is Nothing Then Set cllPro = New Collection
    With vsfPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("校对标志"))) = 1 And (Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1 Or Val(.TextMatrix(i, .ColIndex("卡类别ID"))) = 0) Then
                dbl金额 = Val(.TextMatrix(i, .ColIndex("金额")))
                str结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
                If FormatEx(dbl金额, 6) > 0 Then
                    If Val(.TextMatrix(i, .ColIndex("消费卡"))) = 1 Then
                        str结算信息 = str结算方式 & "," & dbl金额
                        PayType = Pay_SquarePay
                    Else
                        str结算信息 = str结算方式 & "," & dbl金额 & "," & .TextMatrix(i, .ColIndex("结算号码")) & ", "
                        PayType = Pay_CashPay
                    End If
                    strSQL = zlGetRegFeeModifySQL(strNo, lng结帐ID, str结算信息, PayType, , , , , _
                                Val(.TextMatrix(i, .ColIndex("卡类别ID"))), .TextMatrix(i, .ColIndex("卡号")))
                    Call zlAddArray(cllPro, strSQL)
                End If
            End If
        Next
    End With
    zlGetBalanceSQLByVsf = True
    Exit Function
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistSql(ByVal int挂号模式 As Integer, _
                ByVal lng病人ID As Long, ByVal str门诊号 As String, ByVal str姓名 As String, _
                ByVal str性别 As String, ByVal str年龄 As String, ByVal str付款方式 As String, _
                ByVal str费别 As String, ByVal str单据号 As String, ByVal lng执行部门ID As Long, _
                ByVal str发生时间 As String, _
                ByVal str登记时间 As String, ByVal str医生姓名 As String, ByVal byt急诊 As Byte, _
                ByVal str号别 As String, ByVal str诊室 As String, _
                ByVal str摘要 As String, ByVal bln预约挂号 As Boolean, ByVal byt复诊 As Byte, _
                ByVal lng号序 As Long, ByVal int社区 As Integer, ByVal bln预约接收 As Boolean, _
                ByVal str预约方式 As String, ByVal int操作类型 As Integer, ByVal int险类 As Integer, _
                ByVal lng挂号项目ID As Long, Optional ByVal lng出诊记录ID As Long, Optional ByVal int结算模式 As Integer, _
                Optional ByVal str收费单 As String, Optional ByVal str交易流水号 As String, _
                Optional ByVal str交易说明 As String, Optional ByVal str合作单位 As String) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人挂号记录SQL
    ' 入参 : 挂号信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    
    If int挂号模式 = 0 Then
        strSQL = "Zl_病人挂号记录_Insert_S("
    Else
        strSQL = "Zl_病人挂号记录_出诊_Insert_S("
    End If
    '    病人id_In        病人挂号记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(lng病人ID) & ","
    '    门诊号_In        病人挂号记录.门诊号%Type,
    strSQL = strSQL & "" & IIf(str门诊号 = "", "NULL", str门诊号) & ","
    '    姓名_In          病人挂号记录.姓名%Type,
    strSQL = strSQL & "'" & str姓名 & "',"
    '    性别_In          病人挂号记录.性别%Type,
    strSQL = strSQL & "'" & str性别 & "',"
    '    年龄_In          病人挂号记录.年龄%Type,
    strSQL = strSQL & "'" & str年龄 & "',"
    '    付款方式_In      病人挂号记录.医疗付款方式%Type, --用于存放病人的医疗付款方式名称
    strSQL = strSQL & "'" & str付款方式 & "',"
    '    费别_In          病人挂号记录.费别%Type,
    strSQL = strSQL & "'" & str费别 & "',"
    '    单据号_In        病人挂号记录.No%Type,
    strSQL = strSQL & "'" & str单据号 & "',"
    '    执行部门id_In    病人挂号记录.执行部门ID%Type,
    strSQL = strSQL & "" & lng执行部门ID & ","
    '    操作员编号_In    病人挂号记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    病人挂号记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    发生时间_In      病人挂号记录.发生时间%Type,
    strSQL = strSQL & "" & "To_Date('" & str发生时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    登记时间_In      病人挂号记录.登记时间%Type,
    strSQL = strSQL & "" & "To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    医生姓名_In      挂号安排.医生姓名%Type,
    strSQL = strSQL & "'" & str医生姓名 & "',"
    '    急诊_In          Number,
    strSQL = strSQL & "" & byt急诊 & ","
    '    号别_In          挂号安排.号码%Type,
    strSQL = strSQL & "'" & str号别 & "',"
    '    诊室_In          病人挂号记录.诊室%Type,
    strSQL = strSQL & "'" & str诊室 & "',"
    '    摘要_In          病人挂号记录.摘要%Type, --预约挂号摘要信息
    strSQL = strSQL & "'" & str摘要 & "',"
    '    预约挂号_In      Number := 0, --预约挂号时用(记录状态=0,发生时间为预约时间),此时不需要传入结算相关参数
    strSQL = strSQL & "" & IIf(bln预约挂号, 1, 0) & ","
    '    复诊_In          病人挂号记录.复诊%Type := 0,
    strSQL = strSQL & "" & byt复诊 & ","
    '    号序_In          挂号序号状态.序号%Type := Null, --预约时填入费用记录的发药窗口字段,挂号时填入挂号记录
    strSQL = strSQL & "" & ZVal(lng号序) & ","
    '    社区_In          病人挂号记录.社区%Type := Null,
    strSQL = strSQL & "" & ZVal(int社区) & ","
    '    预约接收_In      Number := 0,
    strSQL = strSQL & "" & IIf(bln预约接收, 1, 0) & ","
    '    预约方式_In      预约方式.名称%Type := Null,
    strSQL = strSQL & "'" & str预约方式 & "',"
    '    操作类型_In      Number := 0,
    strSQL = strSQL & "" & int操作类型 & ","
    '    险类_In          病人挂号记录.险类%Type := Null,
    strSQL = strSQL & "" & ZVal(int险类) & ","
    '    挂号项目ID_In    病人挂号记录.挂号项目ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng挂号项目ID) & ","
    '    出诊记录id_In    病人挂号记录.出诊记录ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng出诊记录ID) & ","
    '    结算模式_In      病人挂号记录.结算模式%Type := 0,
    strSQL = strSQL & "" & int结算模式 & ","
    '    收费单_In        病人挂号记录.收费单%Type := Null,
    strSQL = strSQL & "'" & str收费单 & "',"
    '    交易流水号_In    病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '    交易说明_In      病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '    合作单位_In      病人预交记录.合作单位%Type := Null
    strSQL = strSQL & "'" & str合作单位 & "')"
    
    zlGetRegistSql = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistCollectSql(ByVal str医生姓名 As String, ByVal lng医生ID As Long, _
                ByVal lng项目id As Long, ByVal lng执行部门ID As Long, ByVal str发生时间 As String, ByVal byt预约标志 As Byte, _
                ByVal str号别 As String, Optional ByVal lng记录ID As Long) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人挂号汇总SQL
    ' 入参 : 挂号信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    
     strSQL = "zl_病人挂号汇总_Update("
    '  医生姓名_In   挂号安排.医生姓名%Type,
    strSQL = strSQL & "'" & str医生姓名 & "',"
    '  医生id_In     挂号安排.医生id%Type,
    strSQL = strSQL & "" & ZVal(lng医生ID) & ","
    '  收费细目id_In 门诊费用记录.收费细目id%Type,
    strSQL = strSQL & "" & lng项目id & ","
    '  执行部门id_In 门诊费用记录.执行部门id%Type,
    strSQL = strSQL & "" & lng执行部门ID & ","
    '  发生时间_In   门诊费用记录.发生时间%Type,
    strSQL = strSQL & "" & "To_Date('" & str发生时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
    strSQL = strSQL & byt预约标志 & ","
    '  号码_In       挂号安排.号码%Type := Null
    strSQL = strSQL & "'" & str号别 & "',"
    '  三方调用_In   Number := 0
    strSQL = strSQL & "" & "Null" & ","
    '  出诊记录id_In 临床出诊记录.Id%Type := Null
    strSQL = strSQL & "" & ZVal(lng记录ID) & ")"
    
    zlGetRegistCollectSql = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegistFeeSql(ByVal cllPro As Collection, ByVal lng病人ID As Long, ByVal str门诊号 As String, ByVal str姓名 As String, _
                ByVal str性别 As String, ByVal str年龄 As String, ByVal str付款方式 As String, _
                ByVal str费别 As String, ByVal str单据号 As String, ByVal str票据号 As String, _
                ByVal lng序号 As Long, ByVal int价格父号 As Long, ByVal int从属父号 As Long, _
                ByVal str收费类别 As String, ByVal lng收费细目id As Long, ByVal int数次 As Integer, _
                ByVal dbl标准单价 As Double, ByVal lng收入项目id As Long, ByVal str收据费目 As String, _
                ByVal dbl应收金额 As Double, ByVal dbl实收金额 As Double, ByVal lng病人科室ID As Long, _
                ByVal lng开单部门ID As Long, ByVal lng执行部门ID As Long, ByVal str登记时间 As String, _
                ByVal str发生时间 As String, ByVal str医生姓名 As String, ByVal lng结帐ID As Long, _
                ByVal dbl结帐金额 As Double, ByVal lng保险大类ID As Long, ByVal int保险项目否 As Integer, _
                ByVal dbl统筹金额 As Double, ByVal str保险编码 As String, ByVal bln病历费 As Boolean, _
                ByVal byt急诊 As Byte, ByVal str号别 As String, ByVal str诊室 As String, _
                ByVal lng号序 As Long, ByVal bln预约挂号 As Boolean, ByVal str预约方式 As String, _
                ByVal str摘要 As String, ByVal int计费方式 As Integer, Optional ByVal str收费单 As String, _
                Optional ByVal str计算单位 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人挂号费用SQL
    ' 入参 : 挂号费用信息
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/10/31 20:47
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHand
    If cllPro Is Nothing Then Set cllPro = New Collection
    
    strSQL = "Zl_病人挂号费用_Insert_S("
    '    病人id_In        病人挂号记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(lng病人ID) & ","
    '    门诊号_In        病人挂号记录.门诊号%Type,
    strSQL = strSQL & "" & IIf(str门诊号 = "", "NULL", str门诊号) & ","
    '    姓名_In          病人挂号记录.姓名%Type,
    strSQL = strSQL & "'" & str姓名 & "',"
    '    性别_In          病人挂号记录.性别%Type,
    strSQL = strSQL & "'" & str性别 & "',"
    '    年龄_In          病人挂号记录.年龄%Type,
    strSQL = strSQL & "'" & str年龄 & "',"
    '    付款方式_In      病人挂号记录.医疗付款方式%Type, --用于存放病人的医疗付款方式编号
    strSQL = strSQL & "'" & str付款方式 & "',"
    '    费别_In          病人挂号记录.费别%Type,
    strSQL = strSQL & "'" & str费别 & "',"
    '    单据号_In        病人挂号记录.No%Type,
    strSQL = strSQL & "'" & str单据号 & "',"
    '    票据号_In        门诊费用记录.实际票号%Type,
    strSQL = strSQL & "'" & str票据号 & "',"
    '    序号_In          门诊费用记录.序号%Type,
    strSQL = strSQL & "" & lng序号 & ","
    '    价格父号_In      门诊费用记录.价格父号%Type,
    strSQL = strSQL & "" & ZVal(int价格父号) & ","
    '    从属父号_In      门诊费用记录.从属父号%Type,
    strSQL = strSQL & "" & ZVal(int从属父号) & ","
    '    收费类别_In      门诊费用记录.收费类别%Type,
    strSQL = strSQL & "'" & str收费类别 & "',"
    '    收费细目id_In    门诊费用记录.收费细目id%Type,
    strSQL = strSQL & "" & lng收费细目id & ","
    '    数次_In          门诊费用记录.数次%Type,
    strSQL = strSQL & "" & int数次 & ","
    '    标准单价_In      门诊费用记录.标准单价%Type,
    strSQL = strSQL & "" & dbl标准单价 & ","
    '    收入项目id_In    门诊费用记录.收入项目id%Type,
    strSQL = strSQL & "" & lng收入项目id & ","
    '    收据费目_In      门诊费用记录.收据费目%Type,
    strSQL = strSQL & "'" & str收据费目 & "',"
    '    应收金额_In      门诊费用记录.应收金额%Type,
    strSQL = strSQL & "" & IIf(str收费单 <> "", 0, dbl应收金额) & ","
    '    实收金额_In      门诊费用记录.实收金额%Type,
    strSQL = strSQL & "" & IIf(str收费单 <> "", 0, dbl实收金额) & ","
    '    病人科室id_In    门诊费用记录.病人科室id%Type,
    strSQL = strSQL & "" & lng病人科室ID & ","
    '    开单部门id_In    门诊费用记录.开单部门id%Type,
    strSQL = strSQL & "" & lng开单部门ID & ","
    '    执行部门id_In    门诊费用记录.执行部门id%Type,
    strSQL = strSQL & "" & lng执行部门ID & ","
    '    操作员编号_In    门诊费用记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '    操作员姓名_In    门诊费用记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '    登记时间_In      门诊费用记录.登记时间%Type,
    strSQL = strSQL & "" & "To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    发生时间_In      门诊费用记录.发生时间%Type,
    strSQL = strSQL & "" & "To_Date('" & str发生时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
    '    医生姓名_In      门诊费用记录.执行人%Type,
    strSQL = strSQL & "'" & str医生姓名 & "',"
    '    结帐id_In        门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & ZVal(lng结帐ID) & ","
    '    结帐金额_In      病人预交记录.冲预交%Type, --挂号时现金支付部份金额,序号为1传入.
    strSQL = strSQL & "" & IIf(str收费单 <> "", 0, dbl结帐金额) & ","
    '    保险大类id_In    门诊费用记录.保险大类id%Type,
    strSQL = strSQL & "" & ZVal(lng保险大类ID) & ","
    '    保险项目否_In    门诊费用记录.保险项目否%Type,
    strSQL = strSQL & "" & ZVal(int保险项目否) & ","
    '    统筹金额_In      门诊费用记录.统筹金额%Type,
    strSQL = strSQL & "" & ZVal(dbl统筹金额) & ","
    '    保险编码_In      门诊费用记录.保险编码%Type,
    strSQL = strSQL & "'" & str保险编码 & "',"
    '    病历费_In Number, --该条记录是否病历工本费
    strSQL = strSQL & "" & IIf(bln病历费, 1, 0) & ","
    '    急诊_In          Number,
    strSQL = strSQL & "" & byt急诊 & ","
    '    号别_In          门诊费用记录.计算单位%Type,
    strSQL = strSQL & "'" & str号别 & "',"
    '    诊室_In          门诊费用记录.发药窗口%Type,
    strSQL = strSQL & "'" & str诊室 & "',"
    '    号序_In          门诊费用记录.发药窗口%Type,
    strSQL = strSQL & "" & ZVal(lng号序) & ","
    '    预约挂号_In      Number := 0,
    strSQL = strSQL & "" & IIf(bln预约挂号, 1, 0) & ","
    '    预约方式_In      预约方式.名称%Type := Null,
    strSQL = strSQL & "'" & str预约方式 & "',"
    '    摘要_In          门诊费用记录.摘要%Type, --预约挂号摘要信息
    strSQL = strSQL & "'" & str摘要 & "',"
    '    计费方式_In      Number := 0,
    strSQL = strSQL & "" & int计费方式 & ","
    '    收费单_In        门诊费用记录.No%Type := Null
    strSQL = strSQL & "'" & str收费单 & "')"
    zlAddArray cllPro, strSQL
    
    If str收费单 <> "" Then
        strSQL = "Zl_门诊划价记录_Insert_S("
        '    No_In           门诊费用记录.No%Type,
        strSQL = strSQL & "'" & str收费单 & "',"
        '    序号_In         门诊费用记录.序号%Type,
        strSQL = strSQL & "" & lng序号 & ","
        '    病人id_In       门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '    主页id_In       住院费用记录.主页id%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    标识号_In       门诊费用记录.标识号%Type,
        strSQL = strSQL & "" & IIf(str门诊号 = "", "NULL", str门诊号) & ","
        '    付款方式_In     门诊费用记录.付款方式%Type,
        strSQL = strSQL & "'" & str付款方式 & "',"
        '    姓名_In         门诊费用记录.姓名%Type,
        strSQL = strSQL & "'" & str姓名 & "',"
        '    性别_In         门诊费用记录.性别%Type,
        strSQL = strSQL & "'" & str性别 & "',"
        '    年龄_In         门诊费用记录.年龄%Type,
        strSQL = strSQL & "'" & str年龄 & "',"
        '    费别_In         门诊费用记录.费别%Type,
        strSQL = strSQL & "'" & str费别 & "',"
        '    加班标志_In     门诊费用记录.加班标志%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    病人科室id_In   门诊费用记录.病人科室id%Type,
        strSQL = strSQL & "" & lng病人科室ID & ","
        '    开单部门id_In   门诊费用记录.开单部门id%Type,
        strSQL = strSQL & "" & lng开单部门ID & ","
        '    开单人_In       门诊费用记录.开单人%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    从属父号_In     门诊费用记录.从属父号%Type,
        strSQL = strSQL & "" & ZVal(int从属父号) & ","
        '    收费细目id_In   门诊费用记录.收费细目id%Type,
        strSQL = strSQL & "" & lng收费细目id & ","
        '    收费类别_In     门诊费用记录.收费类别%Type,
        strSQL = strSQL & "'" & str收费类别 & "',"
        '    计算单位_In     门诊费用记录.计算单位%Type,
        strSQL = strSQL & "'" & str计算单位 & "',"
        '    发药窗口_In     门诊费用记录.发药窗口%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    付数_In         门诊费用记录.付数%Type,
        strSQL = strSQL & "" & 1 & ","
        '    数次_In         门诊费用记录.数次%Type,
        strSQL = strSQL & "" & int数次 & ","
        '    附加标志_In     门诊费用记录.附加标志%Type,
        strSQL = strSQL & "" & "NULL" & ","
        '    执行部门id_In   门诊费用记录.执行部门id%Type,
        strSQL = strSQL & "" & lng执行部门ID & ","
        '    价格父号_In     门诊费用记录.价格父号%Type,
        strSQL = strSQL & "" & ZVal(int价格父号) & ","
        '    收入项目id_In   门诊费用记录.收入项目id%Type,
        strSQL = strSQL & "" & lng收入项目id & ","
        '    收据费目_In     门诊费用记录.收据费目%Type,
        strSQL = strSQL & "'" & str收据费目 & "',"
        '    标准单价_In     门诊费用记录.标准单价%Type,
        strSQL = strSQL & "" & dbl标准单价 & ","
        '    应收金额_In     门诊费用记录.应收金额%Type,
        strSQL = strSQL & "" & dbl应收金额 & ","
        '    实收金额_In     门诊费用记录.实收金额%Type,
        strSQL = strSQL & "" & dbl实收金额 & ","
        '    发生时间_In     门诊费用记录.发生时间%Type,
        strSQL = strSQL & "" & "To_Date('" & str发生时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
        '    登记时间_In     门诊费用记录.登记时间%Type,
        strSQL = strSQL & "" & "To_Date('" & str登记时间 & "','YYYY-MM-DD HH24:MI:SS')" & ","
        '    操作员姓名_In   门诊费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '    费用id_In       门诊费用记录.Id%Type := Null,
        strSQL = strSQL & "" & "NULL" & ","
        '    费用摘要_In     门诊费用记录.摘要%Type := Null,
        strSQL = strSQL & "'" & "挂号:" & str单据号 & "')"
        zlAddArray cllPro, strSQL
    End If
    zlGetRegistFeeSql = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetRegFeeModifySQL(ByVal strNo As String, ByVal lng结帐ID As Long, _
                ByVal strBalance As String, Optional ByVal Pay结算类型 As gPagePay = Pay_CashPay, _
                Optional ByVal int校对标志 As Integer = 2, Optional ByVal bln完成结算 As Boolean, _
                Optional ByVal bln连续更新 As Boolean, _
                Optional ByVal lng关联ID As Long, Optional ByVal lng卡类别ID As Long, _
                Optional ByVal str卡号 As String, Optional str交易流水号 As String, _
                Optional ByVal str交易说明 As String, Optional ByVal bln普通结算 As Boolean) As String
    Dim strSQL As String
    '更新校对标志，完成挂号收费
    strSQL = "Zl_病人挂号收费_Modify_S("
    '单据号_In         门诊费用记录.No%Type
    strSQL = strSQL & "'" & strNo & "',"
    '结帐id_In         门诊费用记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '结算信息_In       Varchar2,
    strSQL = strSQL & "'" & strBalance & "',"
    '结算类型_In       Number := 0,
    strSQL = strSQL & "" & Pay结算类型 & ","
    '完成标志_In       Number := 0,
    strSQL = strSQL & "" & IIf(bln完成结算, 1, 0) & ","
    '连续更新_In       Number := 0,
    strSQL = strSQL & "" & IIf(bln连续更新, 1, 0) & ","
    '关联交易ID_In     病人预交记录.关联交易ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng关联ID) & ","
    '卡类别ID_In       病人预交记录.卡类别ID%Type := Null,
    strSQL = strSQL & "" & ZVal(lng卡类别ID) & ","
    '卡号_In           病人预交记录.卡号%Type := Null,
    strSQL = strSQL & "'" & str卡号 & "',"
    '交易流水号_In     病人预交记录.交易流水号%Type := Null,
    strSQL = strSQL & "'" & str交易流水号 & "',"
    '交易说明_In       病人预交记录.交易说明%Type := Null,
    strSQL = strSQL & "'" & str交易说明 & "',"
    '普通结算_In       Number := 0
    strSQL = strSQL & "" & IIf(bln普通结算, 1, 0) & ","
    '校对标志_In       Number := 2
    strSQL = strSQL & "" & int校对标志 & ")"
    
    zlGetRegFeeModifySQL = strSQL
End Function
    
Public Function zlGetRegDoneSQL(ByVal strNo As String, ByVal bln在院 As Boolean, ByVal bln预约 As Boolean, _
                Optional ByVal bln生成队列 As Boolean) As String
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取完成挂号过程SQL
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/1 20:15
    '---------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHand
    
    strSQL = "Zl_病人挂号记录_完成挂号_S("
    '    单据号_In     门诊费用记录.No%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '    在院病人_In   Number := 0,
    strSQL = strSQL & "" & IIf(bln在院, 1, 0) & ","
    '    预约标志_In   Number := 0,
    strSQL = strSQL & "" & IIf(bln预约, 1, 0) & ","
    '    生成队列_In Number:=0
    strSQL = strSQL & "" & IIf(bln生成队列, 1, 0) & ")"

    zlGetRegDoneSQL = strSQL
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlGetCancelSql(ByVal lngReg结帐ID As Long, cllBack As Collection, _
                Optional ByVal bln预约 As Boolean, Optional ByVal bln仅删除结算方式 As Boolean) As Boolean
    Dim strSQL As String
    Set cllBack = New Collection
    
    If lngReg结帐ID <> 0 Then
        strSQL = "Zl_病人挂号记录_Cancel(" & lngReg结帐ID & ", " & IIf(bln预约, 2, 0) & "," & IIf(bln仅删除结算方式, 1, 0) & ")"
        zlAddArray cllBack, strSQL
    End If
    
End Function

Public Function zlReadAddrInfo(ByVal objService As clsService, _
                            ByVal objCtrl As PatiAddress, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                            ByVal intTYPE As Integer, Optional ByVal strAddress As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定的病人地址信息到控件中
    '入参:objCtrl-结构化地址控件,intType -地址类型1-出生地，2-籍贯,3-现住址,4-户口地址,5-联系人地址，6-单位地址
    '返回:
    '编制:李南春
    '日期:2015/12/3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str区划代码 As String, str地址_省 As String, str地址_市 As String
    Dim str地址_县 As String, str地址_乡 As String, str地址_其他 As String
    On Error GoTo errHandle
    If objService Is Nothing Then Exit Function
    If lng病人ID = 0 Then zlReadAddrInfo = True: Exit Function
    If objService.zlPatiSvr_GetPatiAddrssInfo(lng病人ID, lng主页ID, intTYPE, str地址_省, str地址_市, _
                    str地址_县, str地址_乡, str地址_其他, str区划代码) Then
         Call objCtrl.LoadStructAdress(str地址_省, str地址_市, str地址_县, str地址_乡, str地址_其他)
    Else
        objCtrl.value = strAddress
    End If
    zlReadAddrInfo = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Zl_Calc_Age(ByVal lng病人ID As Long, ByVal str出生日期 As String, Optional ByVal str计算日期 As String) As String
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errH
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    Zl_Calc_Age = objOneCardComLib.Zl_Calc_Age(lng病人ID, str出生日期, str计算日期)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Public Function SimilarIDs(str身份证号 As String) As String
    '功能：检查病人是否存在相似信息
    '返回：相似记录的病人ID串,如"234,235,236"
    Dim i As Integer
    Dim cllPati As Collection
    Dim cllFilter As New Collection
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errH
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    cllFilter.Add Array("身份证号", str身份证号)
    If objOneCardComLib.zlGetPatiInfsByFilter(2, False, cllFilter, cllPati) = False Then Exit Function
    If cllPati Is Nothing Then Exit Function
    
    For i = 1 To cllPati.Count
        SimilarIDs = SimilarIDs & "|ID:" & cllPati("病人ID")
        SimilarIDs = SimilarIDs & ",姓名:" & cllPati("姓名")
        SimilarIDs = SimilarIDs & ",门诊号:" & IIf(cllPati("门诊号") = "", "无", cllPati("门诊号"))
        SimilarIDs = SimilarIDs & ",身份证号:" & IIf(cllPati("身份证号") = "", cllPati("身份证号"), "未登记")
        SimilarIDs = SimilarIDs & ",地址:" & IIf(cllPati("家庭地址") = "", cllPati("家庭地址"), "未登记")
        SimilarIDs = SimilarIDs & ",登记日期:" & Format(cllPati("登记时间"), "YYYY-MM-DD")
    Next
    SimilarIDs = Mid(SimilarIDs, 2)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function GetCardByName(ByVal strCardName As String, ByVal bln消费卡 As Boolean, ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 根据卡类别名称获取卡对象
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/25 16:28
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlGetCardFromTypeName(strCardName, bln消费卡, objCard) = False Then Exit Function
    If objCard Is Nothing Then GetCardByName = True
End Function

Public Function GetPatiIDByName(ByVal frmMain As Object, ByVal objControl As Object, ByVal strName As String, _
    ByVal str性别 As String, ByRef lngPatiID As Long, _
    Optional ByVal blnCont住院 As Boolean, Optional ByVal blnAddNewPati As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息
    '入参:objControl-调用的控件
    '     strName-病人信息
    '     frmMain-调用的主窗体
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUserCancel As Boolean
    Dim rsPati As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim intNameDays As Integer
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHand
    If LenB(strName) < 4 Then Exit Function
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    intNameDays = Val(gobjDatabase.GetPara("姓名查找天数", glngSys, 9000, 0))
    
    If objOneCardComLib.zlGetPatiIdFromPatiName(objControl, strName, lngPatiID, frmMain, intNameDays, , IIf(blnCont住院, 0, 1), 1, blnUserCancel, blnAddNewPati) = False Then Exit Function
    If blnUserCancel Then Exit Function
    If blnAddNewPati Then GetPatiIDByName = True: Exit Function

    GetPatiIDByName = lngPatiID <> 0
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function CheckUsed门诊号(ByVal lng病人ID As Long, ByVal str门诊号 As String, _
                ByRef blnUsedByOther As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 检查门诊号是否被使用
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/12 11:33
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    CheckUsed门诊号 = objOneCardComLib.zlCheckOutNoIsExist(lng病人ID, str门诊号, blnUsedByOther)
End Function

Public Function CheckMobile(str手机号 As String, Optional ByVal lng病人ID As Long, _
                    Optional ByVal blnShowMsg As Boolean = True, Optional ByRef strErrMsg As String) As Boolean
    '功能：判断指定手机号是否是正确的手机号格式以及是否已经存在于数据库中
    '入参：str手机号-进行检查的手机号
    '      lng病人ID - 检查手机号重复性，不需要检查就传0
    '      blnShowMsg-错误时是否显示提示
    '出参：strErrMsg -错误信息
    '返回：手机号允许使用-true；否则返回False
    Dim blnUsedByOther As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMobileRange As String
    Dim blnQuery As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errH
    '127941:李南春,2018/8/10,根据手机号段检查手机号是否合法
    If str手机号 = "" Then CheckMobile = True: Exit Function
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    strSQL = "Select Max(检查结果) As 检查结果" & vbNewLine & _
                "From (Select Decode(号码长度, Length([1]), 1, 2) As 检查结果" & vbNewLine & _
                "       From 手机号常用号段表" & vbNewLine & _
                "       Where 号段 = Substr([1], 1, Length(号段)))"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "手机号检查", str手机号)
    
    If rsTmp.RecordCount = 0 Then
        strErrMsg = "未能在【手机号常用号段表】检索到输入的手机号格式，请重新录入！"
    ElseIf Val(Nvl(rsTmp!检查结果)) = 0 Then
        strErrMsg = "未能在【手机号常用号段表】检索到输入的手机号格式，请重新录入！"
    ElseIf Val(Nvl(rsTmp!检查结果)) = 2 Then
        strErrMsg = "输入的手机号位数不正确，请重新录入！"
    End If
    
    If gSysPara.bln检查手机号重复 Then
        If objOneCardComLib.zlCheckPhoneIsExist(lng病人ID, str手机号, blnUsedByOther, Not blnShowMsg, strErrMsg) = False Then Exit Function
        If blnUsedByOther Then
            strErrMsg = "输入的手机号与其他病人重复，是否确定录入？"
            blnQuery = True
        End If
    End If
    
    If strErrMsg <> "" Then
        If Not blnShowMsg Then Exit Function
        If blnQuery Then
            If MsgBox(strErrMsg, vbQuestion + vbYesNo, gstrSysName) <> vbYes Then
                strErrMsg = "": Exit Function
            End If
            strErrMsg = ""
        Else
            MsgBox strErrMsg, vbInformation, gstrSysName
            strErrMsg = "": Exit Function
        End If
    End If
    CheckMobile = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function CheckUsedMCNO(ByVal strMCNO As String) As Boolean
    '功能:检查医保号是否已存在
    Dim blnUsed As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    If objOneCardComLib.zlCheckMCNOIsExist(strMCNO, blnUsed) = False Then Exit Function

    If blnUsed Then
        MsgBox "医保号" & strMCNO & "已存在，请检查。", vbInformation, gstrSysName
    End If
    CheckUsedMCNO = Not blnUsed
End Function

Public Function GetPatientInfo(objPati As clsPatientInfo, ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal str查询方式 As String, ByVal lng卡类别ID As Long, ByVal strInput As String, _
                Optional ByVal blnCard As Boolean, Optional ByVal blnCont在院 As Boolean = True, _
                Optional ByVal blnSeekName As Boolean, Optional ByVal strName As String, Optional ByVal strSex As String, _
                Optional ByRef strPassWord As String, Optional ByRef blnUserCancel As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByRef strValidTime As String, Optional ByVal bln密码验证 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人信息
    ' 入参 : int查询方式：0：按病人id查找;-1:卡号模式查找;>0 按卡类别查找
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/6 16:19
    '---------------------------------------------------------------------------------------
    Dim blnValidTime As Boolean '验证卡有效终止时间
    Dim cllPati As Collection, cllOtherFindCons As Collection
    Dim strErrMsg As String
    Dim lngDefaultCardTypeID As Long, lng病人ID As Long
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHand
    strPassWord = ""
    Set objPati = New clsPatientInfo
    Set cllOtherFindCons = New Collection
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If str查询方式 = "病人ID" Then
        lng病人ID = Val(strInput)
    ElseIf str查询方式 = "姓名" Or str查询方式 = "姓名或就诊卡" Then
        If blnCard Then
            If objOneCardComLib.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, , objControl, frmMain, , , , blnUserCancel, , blnValidTime, intCardStatus, strValidTime) = False Then lng病人ID = 0
            
        ElseIf blnSeekName Then
            If GetPatiIDByName(frmMain, objControl, strInput, strSex, lng病人ID, blnCont在院, True) = False Then Exit Function
            If lng病人ID = 0 Then GetPatientInfo = True: Exit Function  '新病人
        End If
        If lng病人ID = 0 Then
            If objOneCardComLib.zlIsMobileNo(strInput) Then
                lng卡类别ID = 0
                str查询方式 = "手机号"
            End If
        End If
    ElseIf str查询方式 = "手机号" Then
        If objOneCardComLib.zlIsMobileNo(strInput) = False Then Exit Function
    ElseIf lng卡类别ID > 0 Then
        str查询方式 = lng卡类别ID
        blnValidTime = True
    End If
    
    If (str查询方式 = "姓名" Or str查询方式 = "姓名或就诊卡") And Not blnCard And lng病人ID = 0 Then
        GetPatientInfo = True: Exit Function   '新病人
    End If
    If lng病人ID = 0 Then
       If objOneCardComLib.zlGetPatiID(str查询方式, strInput, False, lng病人ID, strPassWord, , , objControl, frmMain, , , , blnUserCancel, , blnValidTime, intCardStatus, strValidTime) = False Then Exit Function
    End If
    If lng病人ID = 0 Then Exit Function
    
    If strPassWord <> "" Then
        If Not VerifyPassWord(frmMain, strPassWord) Then
            MsgBox "病人身份验证失败！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If objOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati, strErrMsg) = False Then lng病人ID = 0: Exit Function
    If Not blnCont在院 And objPati.在院 Then
        If objOneCardComLib.zlGetInpatiState(objPati.病人ID, objPati.主页ID, , cllPati) Then
            If Val(cllPati("病人状态")) = 0 Then
                Set objPati = New clsPatientInfo: lng病人ID = 0: Exit Function
            End If
        End If
    End If
    GetPatientInfo = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function GetPatiID(ByVal frmMain As Object, ByVal objControl As Object, _
                ByVal strCardTypes As String, ByVal strInput As String, ByRef lng病人ID As Long, _
                Optional ByVal strCardPassWord As String, Optional ByVal blnUserCancel As Boolean, _
                Optional ByVal blnNotShowErr As Boolean, Optional ByRef intCardStatus As Integer, _
                Optional ByVal blnShowMergePati As Boolean, Optional ByRef strValidTime As String) As Boolean
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlGetPatiID(strCardTypes, strInput, blnNotShowErr, lng病人ID, strCardPassWord, , , objControl, frmMain, blnShowMergePati, True, , blnUserCancel, , True, intCardStatus, strValidTime) = False Then Exit Function
End Function

Public Function GetPatientOtherInfo(ByVal lng病人ID As Long, cllDrug As Collection, _
                                cllImmune As Collection, cllOther As Collection, cllContact As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人健康页信息
    ' 入参 :
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 16:48
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    On Error GoTo errHand
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    If objOneCardComLib.zlGetPatiOtherInforFromPatiID(lng病人ID, , , , True, True, , True, _
                    , cllDrug, cllImmune, , cllOther, cllContact) = False Then Exit Function
    GetPatientOtherInfo = True
    Exit Function
errHand:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Public Function zlUpdateOutMedRec(ByVal lng病人ID As Long, Optional ByVal str病案号 As String, Optional ByVal str建立日期 As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 更新门诊病案，没有记录时新增记录,不传str病案号时删除病案记录
    ' 入参 : str病案号-门诊号
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/21 13:53
    '---------------------------------------------------------------------------------------
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib

    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    
    zlUpdateOutMedRec = objOneCardComLib.zlUpdateOutMedRec(lng病人ID, str病案号, str建立日期)
End Function

Public Function zlCheckMzLgPatiUseDeposit(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊留观病人是否能使用门诊预交
    '入参:lng病人ID-病人ID
    '出参:
    '返回:是门诊留观能够使用门诊预交返回true,否则返回False
    '编制:李南春
    '日期:2019/9/6 15:57:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnLimitDeposit As Boolean, blnMzLgPati As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errHandle
    blnLimitDeposit = Val(gobjDatabase.GetPara(Val("323-门诊留观病人预交款使用控制"), glngSys)) <> 0
    If Not blnLimitDeposit Then zlCheckMzLgPatiUseDeposit = True: Exit Function
    
    If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
    If objOneCardComLib.zlCheckMzLgPati(lng病人ID, lng主页ID, blnMzLgPati, True) = False Then Exit Function
    zlCheckMzLgPatiUseDeposit = Not blnMzLgPati
    
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetMoneyInfoRegist(lng病人ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的门诊预交剩余额
    '入参:
    '       curModiMoney=修改时,原单据的当前病人的费用合计
    '       int类型:类型(0-门诊和住院共用;1-门诊;2-住院),-1表示所有
    '       bytModiMoneyType-修改费用的类别(在按类别统计时有效)
    '       blnFamilyMoney-是否读取家属余额
    '出参:
    '返回:病人剩余额
    '编制:刘兴洪
    '日期:2011-07-21 15:33:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, blnFamilyMoney As Boolean
    Dim strSQL As String, strFamilyIds As String
    Dim objOneCardComLib As zlOneCardComLib.clsOneCardComLib
    
    On Error GoTo errH
    blnFamilyMoney = True
    
    If blnFamilyMoney Then
        If zlGetOneCardComLibObject(Nothing, glngModul, objOneCardComLib) = False Then Exit Function
        Call objOneCardComLib.ZlGetPatiFamilyMember(1, lng病人ID, strFamilyIds)
    End If
    
    strSQL = "Select " & IIf(blnFamilyMoney, "0 As 家属,", "") & _
            "       Nvl(费用余额,0) As 费用余额,Nvl(预交余额,0) As 预交余额" & _
            " From 病人余额" & _
            " Where 性质=1 And 病人ID=[1] And 类型 = 1"
    '79868,读取病人家属余额
    If blnFamilyMoney And strFamilyIds <> "" Then
        strSQL = strSQL & " Union All " & _
                " Select /*+cardinality(B,10) */ " & IIf(blnFamilyMoney, "1 As 家属,", "") & _
                "       Nvl(a.费用余额, 0) As 费用余额, Nvl(a.预交余额, 0) As 预交余额" & _
                " From 病人余额 A, (Select Column_Value as 家属ID from Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) B" & _
                " Where a.病人id = b.家属id And a.性质 = 1 And a.类型 = 1 "
    End If

    strSQL = "Select " & IIf(blnFamilyMoney, "家属,", "") & _
            "       nvl(Sum(费用余额),0) as 费用余额,nvl(Sum(预交余额),0) as 预交余额 " & _
            " From (" & strSQL & ")" & vbCrLf & _
                IIf(blnFamilyMoney, " Group by 家属", "")
    
    Set GetMoneyInfoRegist = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, strFamilyIds)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHand
    strSQL = "Select Zl1_Fun_Getreturnvisit([1],[2]) As 复诊标志 From Dual"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "复诊检查", lng病人ID, lng执行部门ID)
    Check复诊 = Val(Nvl(rsTmp!复诊标志)) = 1
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ZValStr(ByVal lngTmp As Long) As String
    ZValStr = IIf(lngTmp = 0, "", lngTmp)
End Function
