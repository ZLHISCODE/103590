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

'######################################################################################################################
Private mstrSQL As String
Private mrsTmp As ADODB.Recordset
Private mstrTitle As String

Private mstrConvertObject As String

'######################################################################################################################


Public Function GetOperationPerson(Optional ByVal strFilter As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取操作人员
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
    If strFilter = "" Then
        mstrSQL = "SELECT   Distinct A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 As 科室," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 " & _
                        "AND (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                        "AND B.人员性质 In ('医生') " & _
                        "ORDER BY A.编号"
    Else
                        
        mstrSQL = "SELECT   Distinct A.编号," & _
                           "A.姓名," & _
                           "A.简码," & _
                           "C.名称 AS 科室," & _
                           "A.ID " & _
                    "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D " & _
                    "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                        "AND b.人员性质 In ('医生')  "

        mstrSQL = mstrSQL & " AND (A.编号 LIKE [1] OR A.姓名 LIKE [1] OR A.简码 LIKE [1]) "

        mstrSQL = mstrSQL & " Order By A.编号 "
        
        strFilter = "%" & strFilter & "%"
                    
    End If
    
    Set GetOperationPerson = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, strFilter)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Public Function GetDeptSelect(Optional ByVal strFilter As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
                        
    If strFilter = "" Then
        mstrSQL = "SELECT   A.编码," & _
                           "A.名称," & _
                           "A.简码," & _
                           "A.ID,0 As 选择 " & _
                    "FROM 部门表 A,部门性质说明 B " & _
                    "WHERE A.ID=B.部门id " & _
                        "AND B.工作性质 In ('临床') And b.服务对象 In (2,3) And ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) " & _
                        "ORDER BY A.编码"
    Else
                        
        mstrSQL = "SELECT   A.编码," & _
                           "A.名称," & _
                           "A.简码," & _
                           "A.ID " & _
                    "FROM 部门表 A,部门性质说明 B " & _
                    "WHERE A.ID=B.部门id " & _
                        "AND B.工作性质 In ('临床') And b.服务对象 In (2,3)  And ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) "

        mstrSQL = mstrSQL & " AND (A.编码 LIKE [1] OR A.名称 LIKE [1] OR A.简码 LIKE [1]) "

        mstrSQL = mstrSQL & " Order By A.编码 "
        
        strFilter = "%" & strFilter & "%"
                    
    End If
    
    Set GetDeptSelect = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, strFilter)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetInsureKind() As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取保险类别
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "SELECT 项目提示,序号,名称,说明,医院编码,是否固定,是否禁止,具有中心,医保部件,外挂 FROM 保险类别 a Order By a.序号"
    Set GetInsureKind = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetPatient(ByVal bytFilterMode As Byte, Optional ByVal strFilter As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取病人
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    '1-等待审查;2-拒绝审查;3-正在审查;4-审查反馈;5-审查归档
    On Error GoTo errHand
    mstrSQL = "Select Trim(To_Char(B.病人id))||'-'||Trim(To_Char(B.主页id)) As ID,B.病人id,B.主页id,To_Char(B.入院日期,'yyyy-mm-dd hh24:mi') As 入院时间,To_Char(B.出院日期,'yyyy-mm-dd hh24:mi') As 出院时间,C.名称 As 出院科室,A.姓名, A.性别, A.出生日期, A.年龄, A.民族, A.身份证号, A.婚姻状况, A.工作单位 " & _
                "From 病人信息 A,病案主页 B,部门表 C " & _
                "Where A.病人id=B.病人id And C.ID=B.出院科室ID And B.出院日期 Is Not Null And B.病案状态=5 And A.病人id In " & _
                      "(          Select 病人id From 病人信息 Where [1] = 1 And 就诊卡号 = Upper([2]) " & _
                       "Union All Select 病人id From 病人信息 Where [1] = 2 And 病人ID = zl_To_Number([2]) " & _
                       "Union All Select 病人id From 病人信息 Where [1] = 3 And 住院号 = zl_To_Number([2]) " & _
                       "Union All Select 病人id From 病人信息 Where [1] = 4 And 门诊号 = zl_To_Number([2]) " & _
                       "Union All Select 病人id From 病人信息 Where [1] = 5 And 当前床号 = Upper([2]) " & _
                       "Union All Select 病人id From 住院病案记录 Where [1] = 7 And 病案号 = Upper([2]) " & _
                       "Union All Select 病人id From 病人信息 Where [1] = 6 And 姓名 Like '%' || [2] || '%') "

    Set GetPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, bytFilterMode, strFilter)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDeptPatient(ByVal strStart As String, ByVal strEnd As String, _
                            ByVal str当前病况 As String, ByVal str出院情况 As String, _
                            ByVal int病人类型 As Integer, ByVal str医保种类 As String, _
                            ByVal str住院医师 As String, ByVal str疾病名称 As String, _
                            ByVal str检查类型 As String, ByVal str药品信息 As String, _
                            ByVal str医嘱开始时间 As String, ByVal str医嘱结束时间 As String, ByVal blnNCommit As Boolean) As ADODB.Recordset
'功能：读取科室在院病人
    Dim strTmp As String
    On Error GoTo errHand
    
    mstrSQL = "Select /*+ Rule*/Decode(b.封存时间, Null," & vbNewLine & _
                "               Decode(b.病案状态, Null," & vbNewLine & _
                "                       Decode((Select Count(ID)" & vbNewLine & _
                "                               From 病案反馈记录" & vbNewLine & _
                "                               Where 病人id = a.病人id And 主页id = b.主页id And 记录性质 = 1), 0, 'patient', 'aduitDone'), 13," & vbNewLine & _
                "                       'aduitpatient2', 14, 'questionpatient2', 16, 'dealpatient2', 'patient'), 'lockpatient') As 图标," & vbNewLine & _
                "       Decode(b.路径状态, Null, '', 0, '', 'lujin') As 路径, b.数据转出," & vbNewLine & _
                "       Trim(To_Char(a.病人id)) || '-' || Trim(To_Char(b.主页id)) As ID, a.病人id, b.主页id, a.就诊卡号, b.住院号, a.姓名, a.性别, a.年龄," & vbNewLine & _
                "       Nvl(b.住院医师, b.门诊医师) As 住院医师, e.名称 As 护理等级, b.病案状态, d.名称 As 出院科室, b.出院科室id As 出院科室id," & vbNewLine & _
                "       Decode(b.病案状态, Null," & vbNewLine & _
                "               Decode((Select Count(ID) From 病案反馈记录 Where 病人id = a.病人id And 主页id = b.主页id And 记录性质 = 1), 0, Null, '抽查完成')," & vbNewLine & _
                "               13, '正在抽查', 14, '抽查反馈', 16, '抽查整改') As 审查状态, b.出院病床 As 床号, b.入院日期, b.出院日期, b.状态, b.封存时间, b.病案状态 As 病案状态值" & vbNewLine & _
                "From 病人信息 A, 病案主页 B, 部门表 D, 护理等级 E" & vbNewLine & _
                "Where a.病人id = b.病人id And b.护理等级id = e.序号(+) And d.Id = b.出院科室id And Nvl(b.主页id, 0) <> 0 And Nvl(b.状态, 0) <> 1"
    
    If blnNCommit Then '是否包含出院但未提交病案的病人
        mstrSQL = mstrSQL & vbCrLf & " And ((a.在院 = 1 And b.出院日期 Is Null) Or (b.出院日期 Between [1] And [2] And Nvl(b.病案状态, 0) In (0, 13, 14, 16)))"
    Else
        mstrSQL = mstrSQL & vbCrLf & " And a.在院 = 1 And b.出院日期 Is Null"
    End If

    If str当前病况 <> "" Then mstrSQL = mstrSQL & " And b.当前病况=[3] "
    
    Select Case int病人类型
        Case 0          '所有病人
        Case 1          '非医保病人
            mstrSQL = mstrSQL & " And b.险类 Is Null "
        Case 2          '医保病人
            mstrSQL = mstrSQL & " And b.险类 Is Not Null "
            If str医保种类 <> "" Then mstrSQL = mstrSQL & " And b.险类 In (" & str医保种类 & ") "
    End Select

    '住院医师
    If Trim(str住院医师) <> "" Then mstrSQL = mstrSQL & vbCrLf & " And nvl(b.住院医师,b.门诊医师) = [5]"

    '疾病名称
    If Trim(str疾病名称) <> "" Then
        mstrSQL = mstrSQL & vbCrLf & "And Exists ("
        strTmp = strTmp & vbCrLf & "select 1 from 病人诊断记录 x where 记录来源 in(2,3)  And x.病人ID = B.病人ID and x.主页ID = B.主页ID"
        strTmp = strTmp & vbCrLf & "And 诊断类型 in (1,2,3,5,6,7,11,12,13) and 疾病ID = [6]"
        strTmp = strTmp & vbCrLf & ")"
    End If
    '检查类型
    If str检查类型 <> "" Then
        '医嘱开始时间、医嘱结束时间
        If str药品信息 <> "" Then '存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 诊疗类别 in ('D','5','6','7') And f.病人ID = B.病人ID and f.主页ID = B.主页ID"
            strTmp = strTmp & vbCrLf & "And 开嘱时间 Between [7] and [8]"
            strTmp = strTmp & vbCrLf & "And 诊疗项目ID in (select ID from 诊疗项目目录 where 类别='D' and 操作类型 = [9])"
            strTmp = strTmp & vbCrLf & "And 收费细目ID =[10]"
            strTmp = strTmp & vbCrLf & ")"
        Else '不存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 开嘱时间 Between [7] and [8]"
            strTmp = strTmp & vbCrLf & "and 诊疗类别='D' and 诊疗项目ID in (select ID from 诊疗项目目录 where 类别='D' and 操作类型 = [9]) And f.病人ID = B.病人ID and f.主页ID = B.主页ID"
            strTmp = strTmp & vbCrLf & ")"
        End If
    Else
        '只有药品信息
        If str药品信息 <> "" Then '存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 诊疗类别 in ('5','6','7') And f.病人ID = B.病人ID and f.主页ID = B.主页ID"
            strTmp = strTmp & vbCrLf & "And 收费细目ID =[10]"
            strTmp = strTmp & vbCrLf & ")"
        End If
    End If
    
    If strTmp <> "" Then mstrSQL = mstrSQL & strTmp
    
    If str出院情况 <> "" Then
        mstrSQL = mstrSQL & " And Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[4])"
    End If
    
    Select Case int病人类型
        Case 0          '所有病人
        Case 1          '非医保病人
            mstrSQL = mstrSQL & " And b.险类 Is Null "
        Case 2          '医保病人
            mstrSQL = mstrSQL & " And b.险类 Is Not Null "
            If str医保种类 <> "" Then
                mstrSQL = mstrSQL & " And b.险类 In (" & str医保种类 & ") "
            End If
    End Select
    
    If blnNCommit Then
        mstrSQL = "Select * From (" & mstrSQL & ") Order by nvl(出院日期,sysdate) desc ,出院科室id,床号"
    End If
    
    Set GetDeptPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, CDate(strStart), CDate(strEnd), str当前病况, str出院情况, _
        str住院医师, Val(str疾病名称), CDate(str医嘱开始时间), CDate(str医嘱结束时间), _
        str检查类型, Val(str药品信息))
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAduitPatient(ByVal bytApplyMode As Byte, ByVal strStart As String, ByVal strEnd As String, _
                            ByVal strState As String, _
                            ByVal lng病人ID As Long, _
                            ByVal lng主页ID As Long, _
                            ByVal str出院情况 As String, _
                            ByVal int病人类型 As Integer, _
                            ByVal str医保种类 As String, _
                            ByVal str住院医师 As String, _
                            ByVal str疾病名称 As String, _
                            ByVal str检查类型 As String, _
                            ByVal str药品信息 As String, _
                            ByVal str医嘱开始时间 As String, _
                            ByVal str医嘱结束时间 As String _
                            ) As ADODB.Recordset
                            
    '******************************************************************************************************************
    '功能：读取审查/归档病人
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim varState    As Variant
    Dim strTmp      As String
    '1-接收待审;2-拒绝接收;3-正在审查;4-审查反馈;5-审查归档;6-提交待收
    On Error GoTo errHand
    varState = Split(strState, ";")

    mstrSQL = _
        "Select  0 as 选择, Decode(b.封存时间,Null, Decode(b.病案状态,1,'waitpatient',10,'waitpatient',2,'refusepatient',3,'aduitpatient',4,'questionpatient',5,'archivepatient',6,'dealpatient',13,'aduitpatient2',14,'questionpatient2',16,'dealpatient2'),'lockpatient') As 图标," & _
        "Decode(b.路径状态,Null,'',0,'','lujin') as 路径,b.数据转出," & _
        "Decode(b.病案状态,1,'等待审查',10,'接收待审',2,'拒绝审查',3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改',13,'正在抽查',14,'抽查反馈',16,'抽查整改') As 病案状态,b.出院科室id," & _
        "C.ID,c.提交人,c.提交时间,c.接收人,c.接收时间,c.归档人,c.归档时间,c.拒审人,c.拒审时间,c.拒审理由," & _
        " A.病人ID,B.主页ID,A.就诊卡号,B.住院号,A.姓名,A.性别,b.年龄 ,nvl(b.住院医师,b.门诊医师) as 住院医师,e.名称 as 护理等级," & _
        " LPAD(B.出院病床,10,' ') as 床号,B.入院日期,B.出院日期,b.病案状态 as 病案状态值,b.封存时间,d.名称 As 出院科室,nvl(f.反馈条数,0) as 反馈条数,Decode(nvl(nvl(f.反馈条数,0) -f.完成反馈条数,0),0,1,0) as 反馈完成,d.撤档时间 " & _
        " From 病案提交记录 c,病人信息 a,病案主页 b,部门表 d ,护理等级 e,  (Select A.提交ID,Sum(A.反馈条数) as 反馈条数,Sum(A.完成反馈条数) as 完成反馈条数  From (Select 提交ID,Count(ID) as 反馈条数,0 as 完成反馈条数 From 病案反馈记录 where  提交ID is not null Group by 提交ID union All  Select 提交ID,0 as 反馈条数,Count(记录状态) as 完成反馈条数 From 病案反馈记录 where  提交ID is not Null Group by 提交ID,记录状态 having 记录状态=3) A  Group by A.提交ID) F " & _
        " Where c.提交时间 Between [1] And [2] And  c.记录状态<>2 And c.病人id=b.病人id And c.主页id=b.主页id And c.ID = f.提交ID(+) And A.病人ID=B.病人ID And d.ID=b.出院科室id  And b.护理等级ID = e.序号(+)" & _
        " And  b.病案状态 In ([3],[4],[5],[6],[10],[17]) "
        
    If lng病人ID > 0 Then mstrSQL = mstrSQL & vbCrLf & " And c.病人id=[7] And c.主页id=[8] "
    If str出院情况 <> "" Then
        mstrSQL = mstrSQL & vbCrLf & " And  Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[9])"
    End If
    
    Select Case int病人类型
        Case 0          '所有病人
        Case 1          '非医保病人
            mstrSQL = mstrSQL & vbCrLf & " And b.险类 Is Null "
        Case 2          '医保病人
            mstrSQL = mstrSQL & vbCrLf & " And b.险类 Is Not Null "
            If str医保种类 <> "" Then
                mstrSQL = mstrSQL & vbCrLf & " And b.险类 In (" & str医保种类 & ") "
            End If
    End Select
    '添加的查询
    '住院医师
    If Trim(str住院医师) <> "" Then
        mstrSQL = mstrSQL & vbCrLf & " And nvl(b.住院医师,b.门诊医师) = [11]"
    End If
    '疾病名称
    If Trim(str疾病名称) <> "" Then
        mstrSQL = mstrSQL & vbCrLf & "And Exists ("
        strTmp = strTmp & vbCrLf & "select 1 from 病人诊断记录 x where 记录来源=2  And x.病人ID = c.病人ID and x.主页ID = c.主页ID"
        strTmp = strTmp & vbCrLf & "And 诊断类型 in (3,13) and 疾病ID = [12]"
        strTmp = strTmp & vbCrLf & ")"
    End If
    '检查类型
    If str检查类型 <> "" Then
        '医嘱开始时间、医嘱结束时间
        If str药品信息 <> "" Then '存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 诊疗类别 in ('D','5','6','7') And f.病人ID = c.病人ID and f.主页ID = c.主页ID"
            strTmp = strTmp & vbCrLf & "And 开嘱时间 Between [13] and [14]"
            strTmp = strTmp & vbCrLf & "And 诊疗项目ID in (select ID from 诊疗项目目录 where 类别='D' and 操作类型 = [15])"
            strTmp = strTmp & vbCrLf & "And 收费细目ID =[16]"
            strTmp = strTmp & vbCrLf & ")"
        Else '不存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 开嘱时间 Between [13] and [14]"
            strTmp = strTmp & vbCrLf & "and 诊疗类别='D' and 诊疗项目ID in (select ID from 诊疗项目目录 where 类别='D' and 操作类型 = [15]) And f.病人ID = c.病人ID and f.主页ID = c.主页ID"
            strTmp = strTmp & vbCrLf & ")"
        End If
    Else
        '只有药品信息
        If str药品信息 <> "" Then '存在药品信息
            strTmp = strTmp & vbCrLf & "And Exists ("
            strTmp = strTmp & vbCrLf & "select 1 from 病人医嘱记录 f where 病人来源 = 2 And 诊疗类别 in ('5','6','7') And f.病人ID = c.病人ID and f.主页ID = c.主页ID"
            strTmp = strTmp & vbCrLf & "And 收费细目ID =[16]"
            strTmp = strTmp & vbCrLf & ")"
        End If
    End If
    
    If strTmp <> "" Then mstrSQL = mstrSQL & strTmp
    
    If Val(varState(0)) = 1 Then varState(0) = 10
    If Val(varState(1)) = 1 Then varState(1) = 5
    If Val(varState(2)) = 1 Then varState(2) = 3
    If Val(varState(3)) = 1 Then varState(3) = 4
    If Val(varState(4)) = 1 Then varState(4) = 6
    If Val(varState(5)) = 1 Then varState(5) = 1
    
    Set GetAduitPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, _
            CDate(strStart), CDate(strEnd), Val(varState(0)), Val(varState(1)), Val(varState(2)), _
            Val(varState(3)), lng病人ID, lng主页ID, str出院情况, Val(varState(4)), _
            str住院医师, Val(str疾病名称), CDate(str医嘱开始时间), CDate(str医嘱结束时间), str检查类型, _
            Val(str药品信息), Val(varState(5)))
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetCustomPatient(Optional ByVal lng病种id As Long, _
                                    Optional ByVal str性别 As String, _
                                    Optional ByVal lng科室ID As Long, _
                                    Optional ByVal int开始年龄 As Integer, _
                                    Optional ByVal int结束年龄 As Integer, _
                                    Optional ByVal str婚姻状况 As String, _
                                    Optional ByVal str住院号 As String, _
                                    Optional ByVal str病案号 As String, _
                                    Optional ByVal strCustom As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取病人
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "Select Trim(To_Char(B.病人id))||'-'||Trim(To_Char(B.主页id)) As ID,B.病人id,B.主页id,b.住院号,d.病案号,b.主页id As 住院次数,B.入院日期 As 入院时间,B.出院日期 As 出院时间,b.数据转出,C.名称 As 出院科室,A.姓名, A.性别, A.出生日期, " & _
                "Decode(a.出生日期,Null,Decode(b.年龄,Null,0,Decode(Trim(Substr(b.年龄,Length(b.年龄),1)),'岁',Zl_To_Number(Substr(b.年龄,1,Length(b.年龄)-1)),'月',Zl_To_Number(Substr(b.年龄,1,Length(b.年龄)-1))/12,'天',Zl_To_Number(Substr(b.年龄,1,Length(b.年龄)-1))/365,Zl_To_Number(b.年龄))),Trunc(Months_between(Sysdate,a.出生日期)/12)) As 年龄," & _
                "A.民族, A.身份证号, A.婚姻状况, A.工作单位 " & _
                "From 病人信息 A,病案主页 B,部门表 C,住院病案记录 D " & _
                "Where A.病人id=B.病人id And C.ID=B.出院科室ID And b.病案状态=5 And d.病人id=b.病人id And (d.主页id=b.主页id Or d.主页id Is Null) " & strCustom
    
    If lng病种id > 0 Then mstrSQL = mstrSQL & " And (b.病人id,b.主页id) In (Select 病人id,主页id From 病人诊断记录 Where 疾病id=[1])"
    If str性别 <> "" Then mstrSQL = mstrSQL & " And a.性别=[2]"
    If lng科室ID > 0 Then mstrSQL = mstrSQL & " And b.出院科室id=[3]"
    If str婚姻状况 <> "" Then mstrSQL = mstrSQL & " And b.婚姻状况=[4]"
    If Val(str住院号) > 0 Then mstrSQL = mstrSQL & " And b.住院号=[7]"
    If str病案号 <> "" Then mstrSQL = mstrSQL & " And d.病案号=[8]"
    
    If int开始年龄 <> 0 Or int开始年龄 <> 0 Then
        mstrSQL = "Select * From (" & mstrSQL & ") Where 年龄 Between [5] And [6]"
    End If
    
    Set GetCustomPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病种id, str性别, lng科室ID, str婚姻状况, int开始年龄, int结束年龄, Val(str住院号), str病案号)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetBorrow(ByVal bytFilterMode As Byte, _
                        Optional ByVal lngKey As Long, _
                        Optional ByVal strStartNo As String, _
                        Optional ByVal strEndNo As String, _
                        Optional ByVal str申请人 As String, _
                        Optional ByVal str批准人 As String, _
                        Optional ByVal str拒借人 As String, _
                        Optional ByVal str登记开始时间 As String, _
                        Optional ByVal str登记结束时间 As String, _
                        Optional ByVal str批准开始时间 As String, _
                        Optional ByVal str批准结束时间 As String, _
                        Optional ByVal str拒借开始时间 As String, _
                        Optional ByVal str拒借结束时间 As String, _
                        Optional ByVal bln新申请 As Boolean = True, _
                        Optional ByVal bln已批准 As Boolean = True, _
                        Optional ByVal bln已拒绝 As Boolean = True, _
                        Optional ByVal bln已归还 As Boolean = True, _
                        Optional ByVal str住院号 As String, _
                        Optional ByVal str病人姓名 As String, _
                        Optional ByVal str申请理由 As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取电子病案借阅记录
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim dt登记开始时间 As Date
    Dim dt登记结束时间 As Date
    Dim dt批准开始时间 As Date
    Dim dt批准结束时间 As Date
    Dim dt拒借开始时间 As Date
    Dim dt拒借结束时间 As Date
    
    Dim strWhere As String
    On Error GoTo errHand
    
    Select Case bytFilterMode
    Case 1          '读取指定的借阅申请单据
        mstrSQL = "Select Decode(记录状态,1,'借阅申请',2,'批准申请',3,'拒绝申请',4,'归还接收') As 图标,ID,记录状态,No,申请人,申请时间,申请理由,申请期限,批准人,批准时间,借阅时间,借阅期限,拒借人,拒借理由,拒借时间,登记时间,收回人,归还时间 " & _
                    "From 病案借阅记录 Where ID=[1]"
        
        Set GetBorrow = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
    Case 2          '读取单据清单
        strWhere = ""
        
        dt登记开始时间 = Empty
        dt登记结束时间 = Empty
        dt批准开始时间 = Empty
        dt批准结束时间 = Empty
        dt拒借开始时间 = Empty
        dt拒借结束时间 = Empty
        
        If str登记开始时间 <> "" Then
            strWhere = strWhere & " And A.登记时间 Between [1] And [2] "
            dt登记开始时间 = CDate(str登记开始时间)
            dt登记结束时间 = CDate(str登记结束时间)
        End If
        
        If str批准开始时间 <> "" Then
            strWhere = strWhere & " And A.批准时间 Between [3] And [4] "
            dt批准开始时间 = CDate(str批准开始时间)
            dt批准结束时间 = CDate(str批准结束时间)
        End If
        
        If str拒借开始时间 <> "" Then
            strWhere = strWhere & " And A.拒借时间 Between [5] And [6] "
            dt拒借开始时间 = CDate(str拒借开始时间)
            dt拒借结束时间 = CDate(str拒借结束时间)
        End If
        
        If strStartNo <> "" And strEndNo <> "" Then
            strWhere = strWhere & " And No Between [7] And [8] "
        ElseIf strStartNo <> "" Then
            strWhere = strWhere & " And No = [7] "
        ElseIf strEndNo <> "" Then
            strWhere = strWhere & " And No = [8] "
        End If
        
        If str申请人 <> "" Then strWhere = strWhere & " And A.申请人 Like [9] "
        If str批准人 <> "" Then strWhere = strWhere & " And A.批准人 Like [10] "
        If str拒借人 <> "" Then strWhere = strWhere & " And A.拒借人 Like [11] "
        strWhere = strWhere & " And A.记录状态 In ([12],[13],[14],[15])"
        
        If str住院号 <> "" Then strWhere = strWhere & " And C.住院号 Like [16] "
        If str病人姓名 <> "" Then strWhere = strWhere & " And D.姓名 Like [17] "
        If str申请理由 <> "" Then strWhere = strWhere & " And A.申请理由 Like [18] "
         
        If str住院号 = "" And str病人姓名 = "" And str申请理由 = "" Then
            mstrSQL = "Select Decode(A.记录状态,1,'借阅申请',2,'批准申请',3,'拒绝申请',4,'归还接收') As 图标,A.ID,A.记录状态,A.No,A.申请人,A.申请时间,A.申请理由,A.申请期限,A.批准人,A.批准时间,A.借阅时间,A.借阅期限,A.拒借人,A.拒借理由,A.拒借时间,A.登记时间,收回人,归还时间 " & _
                        "From 病案借阅记录 A Where 1=1 " & strWhere
        Else
            mstrSQL = "Select Decode(A.记录状态, 1, '借阅申请', 2, '批准申请', 3, '拒绝申请', 4, '归还接收') As 图标,A.ID, A.记录状态, A.No, A.申请人, A.申请时间, A.申请理由, A.申请期限, A.批准人," & vbNewLine & _
                        "       A.批准时间, A.借阅时间,A.借阅期限, A.拒借人,A.拒借理由,  A.拒借时间, A.登记时间, B.病人ID, C.住院号, D.姓名,收回人,归还时间" & vbNewLine & _
                        "  From 病案借阅记录 A," & vbNewLine & _
                        "       病案借阅内容 B," & vbNewLine & _
                        "       病案主页     C," & vbNewLine & _
                        "       病人信息     D," & vbNewLine & _
                        "       住院病案记录 E" & vbNewLine & _
                        " Where A.id = B.借阅ID" & vbNewLine & _
                        "   And B.病人ID = C.病人ID" & vbNewLine & _
                        "   And B.主页id = C.主页ID" & vbNewLine & _
                        "   And B.病人id = D.病人id" & vbNewLine & _
                        "   And E.病人id = B.病人id" & vbNewLine & _
                        "   And (e.主页id = b.主页id Or e.主页id Is Null) " & strWhere
        End If
        
        str申请人 = "%" & str申请人 & "%"
        str批准人 = "%" & str批准人 & "%"
        str拒借人 = "%" & str拒借人 & "%"
        
        If str住院号 = "" Then
            str住院号 = ""
        Else
            str住院号 = "%" & str住院号 & "%"
        End If
        
        If str病人姓名 = "" Then
            str病人姓名 = ""
        Else
            str病人姓名 = "%" & str病人姓名 & "%"
        End If
        
        If str申请理由 = "" Then
            str申请理由 = ""
        Else
            str申请理由 = "%" & str申请理由 & "%"
        End If
        
        Set GetBorrow = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, dt登记开始时间, dt登记结束时间, dt批准开始时间, dt批准结束时间, dt拒借开始时间, dt拒借结束时间, strStartNo, strEndNo, str申请人, str批准人, str拒借人, _
                                                IIf(bln新申请, 1, 0), IIf(bln已批准, 2, 0), IIf(bln已拒绝, 3, 0), IIf(bln已归还, 4, 0), str住院号, str病人姓名, str申请理由)
    
    End Select
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBorrowPerson(ByVal lngKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取电子病案借阅的人员
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    
    mstrSQL = "Select 'person' As 图标,b.ID,b.姓名,b.编号,d.名称 As 科室 From 病案借阅人员 a,人员表 b,部门人员 c,部门表 d Where a.借阅id=[1] And a.人员id=b.ID And b.ID=c.人员id And c.缺省=1 And d.ID=c.部门id"

    Set GetBorrowPerson = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
             
End Function

Public Function GetBorrowPatient(Optional ByVal lngKey As Long, _
                                Optional ByVal strStartNo As String, _
                                Optional ByVal strEndNo As String, _
                                Optional ByVal str申请人 As String, _
                                Optional ByVal str批准人 As String, _
                                Optional ByVal str拒借人 As String, _
                                Optional ByVal str登记开始时间 As String, _
                                Optional ByVal str登记结束时间 As String, _
                                Optional ByVal str批准开始时间 As String, _
                                Optional ByVal str批准结束时间 As String, _
                                Optional ByVal str拒借开始时间 As String, _
                                Optional ByVal str拒借结束时间 As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取电子病案借阅的病人
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim dt登记开始时间 As Date
    Dim dt登记结束时间 As Date
    Dim dt批准开始时间 As Date
    Dim dt批准结束时间 As Date
    Dim dt拒借开始时间 As Date
    Dim dt拒借结束时间 As Date
    Dim strWhere As String
    On Error GoTo errHand
    
    strWhere = ""
    
    dt登记开始时间 = Empty
    dt登记结束时间 = Empty
    dt批准开始时间 = Empty
    dt批准结束时间 = Empty
    dt拒借开始时间 = Empty
    dt拒借结束时间 = Empty
    
    If str登记开始时间 <> "" Then
        strWhere = strWhere & " And 登记时间 Between [1] And [2] "
        dt登记开始时间 = CDate(str登记开始时间)
        dt登记结束时间 = CDate(str登记结束时间)
    End If
    
    If str批准开始时间 <> "" Then
        strWhere = strWhere & " And 批准时间 Between [3] And [4] "
        dt批准开始时间 = CDate(str批准开始时间)
        dt批准结束时间 = CDate(str批准结束时间)
    End If
    
    If str拒借开始时间 <> "" Then
        strWhere = strWhere & " And 拒借时间 Between [5] And [6] "
        dt拒借开始时间 = CDate(str拒借开始时间)
        dt拒借结束时间 = CDate(str拒借结束时间)
    End If
    
    If strStartNo <> "" And strEndNo <> "" Then
        strWhere = strWhere & " And No Between [7] And [8] "
    ElseIf strStartNo <> "" Then
        strWhere = strWhere & " And No = [7] "
    ElseIf strEndNo <> "" Then
        strWhere = strWhere & " And No = [8] "
    End If
    
    If str申请人 <> "" Then strWhere = strWhere & " And 申请人 Like [9] "
    If str批准人 <> "" Then strWhere = strWhere & " And 批准人 Like [10] "
    If str拒借人 <> "" Then strWhere = strWhere & " And 拒借人 Like [11] "
    
    mstrSQL = "Select ID From 病案借阅记录 Where 记录状态=2 " & strWhere
    
    str申请人 = "%" & str申请人 & "%"
    str批准人 = "%" & str批准人 & "%"
    str拒借人 = "%" & str拒借人 & "%"

    If lngKey = 0 Then
        
        mstrSQL = "Select 'archivepatient' As 图标,Decode(b.路径状态,Null,'',0,'','lujin') as 路径,E.No,E.申请人,E.申请时间,E.申请理由,b.病人id As ID,b.病人id,b.主页id,b.住院号,g.病案号,b.主页id As 住院次数,b.数据转出,LPAD(b.出院病床,10,' ') as 床号,b.入院日期 As 入院时间,b.出院日期 As 出院时间,c.名称 As 出院科室,b.出院科室ID as 出院科室ID,d.姓名, d.性别, d.年龄, d.身份证号,d.就诊卡号, d.婚姻状况,Z.病人ID as 借出状态,Z.申请人 as 借出申请人,Decode(Z.病人ID,'','在院','借出') as 病案存储状态 " & _
                    "From 病案借阅内容 A,病案主页 B,部门表 C,病人信息 D,病案借阅记录 E,病案借阅人员 F,住院病案记录 g,(Select Max(A1.申请人) as 申请人,Max(B1.病人ID) as 病人ID,Max(B1.主页ID) as 主页ID From 病案借阅记录 A1 ,病案借阅内容 B1,病案借阅人员 C1 Where A1.ID = B1.借阅ID And A1.ID = C1.借阅ID And A1.记录状态=2 Group By B1.病人ID,B1.主页ID,A1.申请人) Z " & _
                    "Where b.病人id=a.病人id And b.主页id=a.主页id And g.病人id(+)=b.病人id And (g.主页id=b.主页id Or g.主页id Is Null) And d.病人id=b.病人id  And Z.病人ID(+) = a.病人ID And Z.主页ID(+)=a.主页ID And c.ID=b.出院科室ID And E.ID=F.借阅id And F.人员id=[12] And A.借阅id=E.ID And A.借阅id In (" & mstrSQL & ") And (E.借阅时间 <= sysdate And E.借阅期限 >= sysdate)"
                    
        Set GetBorrowPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, dt登记开始时间, dt登记结束时间, dt批准开始时间, dt批准结束时间, dt拒借开始时间, dt拒借结束时间, strStartNo, strEndNo, str申请人, str批准人, str拒借人, UserInfo.ID)
    Else
        mstrSQL = "Select 'archivepatient' As 图标,Decode(b.路径状态,Null,'',0,'','lujin') as 路径,E.No,E.申请人,E.申请时间,E.申请理由,b.病人id As ID,b.病人id,b.主页id,b.住院号,g.病案号,b.主页id As 住院次数,b.数据转出,LPAD(b.出院病床,10,' ') as 床号,b.入院日期 As 入院时间,b.出院日期 As 出院时间,c.名称 As 出院科室,b.出院科室ID as 出院科室ID,d.姓名, d.性别, d.年龄, d.身份证号,d.就诊卡号, d.婚姻状况,Z.病人ID as 借出状态,Z.申请人 as 借出申请人,Decode(Z.病人ID,'','在院','借出') as 病案存储状态 " & _
                    "From 病案借阅内容 A,病案主页 B,部门表 C,病人信息 D,病案借阅记录 E,住院病案记录 g,(Select Max(A1.申请人) as 申请人,Max(B1.病人ID) as 病人ID,Max(B1.主页ID) as 主页ID From 病案借阅记录 A1 ,病案借阅内容 B1,病案借阅人员 C1 Where A1.ID = B1.借阅ID And A1.ID = C1.借阅ID And A1.记录状态=2 Group By B1.病人ID,B1.主页ID,A1.申请人) Z " & _
                    "Where a.借阅id=[1] And b.病人id=a.病人id And b.主页id=a.主页id And g.病人id(+)=b.病人id And (g.主页id=b.主页id Or g.主页id Is Null) And d.病人id=b.病人id  And Z.病人ID(+) = a.病人ID  And Z.主页ID(+)=a.主页ID And c.ID=b.出院科室ID And A.借阅id=E.ID " 'And (E.借阅时间 <= sysdate And E.借阅期限 >= sysdate)
        Set GetBorrowPatient = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDisease(Optional ByVal strFilter As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取疾病编码目录
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim StrText As String
    Dim strTmp As String
    On Error GoTo errHand
    
    If strFilter = "" Then
        mstrSQL = "SELECT ID," & _
                    "上级ID," & _
                    "0 AS 末级," & _
                    "NULL AS 编码," & _
                    "名称," & _
                    "NULL AS 简码,Null As 附码 " & _
            "FROM 疾病编码分类 " & _
            "WHERE 类别='D' " & _
            "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
            "UNION ALL " & _
            "SELECT A.ID, " & _
                    "A.分类id AS 上级ID, " & _
                    "1 AS 末级, " & _
                    "A.编码, " & _
                    "A.名称, " & _
                    "A.简码,a.附码 " & _
            "FROM 疾病编码目录 A " & _
            "WHERE 类别='D' "
    Else
        mstrSQL = "SELECT   编码," & _
                           "名称," & _
                           "简码," & _
                           "附码," & _
                           "ID " & _
                    "FROM 疾病编码目录 " & _
                    "WHERE 类别='D' "
                        
        StrText = strFilter & "%"
        strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strFilter, "%" & strFilter)
                
        Select Case GetApplyMode(strFilter)
        Case 1      '是全数字，按编码查找
            mstrSQL = mstrSQL & " And 编码 Like [1] "
        Case 2      '是全字母，按简码查找
            mstrSQL = mstrSQL & " And 简码 Like [2] "
        Case Else
            mstrSQL = mstrSQL & "AND (编码 Like [1] OR 名称 Like [2] OR 简码 Like [2])"
        End Select
    End If
    Set GetDisease = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, StrText, strTmp)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBaseCode(ByVal strTable As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取基础字典编码
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mstrSQL = "SELECT a.编码||'-'||a.名称 As 显示名称,a.* FROM " & strTable & " a Order By a.编码"
    Set GetBaseCode = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDept(ByVal str工作性质 As String, Optional byt服务对象 As Byte = 2, Optional ByVal blnAllDept As Boolean = True, Optional ByVal blnShowStopDept As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取指定工作性质的部门
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim strStopDeptSql As String
    If blnShowStopDept Then
        strStopDeptSql = ""
    Else
        strStopDeptSql = " and ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) "
    End If
    If blnAllDept Then
        mstrSQL = "SELECT a.编码||'-'||a.名称 As 显示名称,a.ID,a.上级ID,a.编码,a.名称,a.简码,a.位置,a.末级,a.建档时间,a.撤档时间,a.站点,a.环境类别 FROM 部门表 a Where Exists (Select 1 From 部门性质说明 b Where b.工作性质=[1] And a.ID=b.部门id  And b.服务对象 In ([2],3)) " & strStopDeptSql & " Order By a.编码"
    Else
        mstrSQL = "SELECT a.编码||'-'||a.名称 As 显示名称,a.ID,a.上级ID,a.编码,a.名称,a.简码,a.位置,a.末级,a.建档时间,a.撤档时间,a.站点,a.环境类别 FROM 部门表 a Where Exists (Select 1 From 部门性质说明 b Where b.工作性质=[1] And a.ID=b.部门id  And b.服务对象 In ([2],3)) And a.ID In (Select 部门id From 部门人员 Where 人员id=[3]) " & strStopDeptSql & " Order By a.编码"
    End If
    
    Set GetDept = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, str工作性质, byt服务对象, UserInfo.ID)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetCISStruct(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, ByVal blnDataMove As Boolean, Optional ByVal strType As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取电子病案的结构
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim arySerial As Variant
    Dim strTmp As String
    Dim strSerial(1 To 9) As String
    Dim intCount As Integer
    Dim strSQL As String
    Dim strSQL1 As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnPath As Boolean '是否有权限显示临床路径
    
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    On Error GoTo errHand
    
    strTmp = Trim(zlDatabase.GetPara("档案排序顺序", ParamInfo.系统号, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    For intCount = 0 To UBound(arySerial)
        strSerial(Val(arySerial(intCount))) = intCount
    Next
    
    '病人科室存在可用的临床路径时，显示临床路径记录
    '先判断是否有"临床路径应用" 序号=1256
    If GetPrivFunc(glngSys, 1256) <> "" Then
        blnPath = GetHavePath(lng科室ID) 'mlng科室ID
    End If
    
    mstrSQL = _
        "Select *" & vbNewLine & _
        "From (Select 'R5' As ID, '' As 上级id,NULL as EPRID, '首页记录' As 名称, '' As 参数,1 As 末级,'object_first' As 图标,[7] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R2' As ID, '' As 上级id,NULL as EPRID, '住院病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[4] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R3' As ID, '' As 上级id,NULL as EPRID, '护理病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[5] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R4' As ID, '' As 上级id,NULL as EPRID, '护理记录' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[6] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R1' As ID, '' As 上级id,NULL as EPRID, '住院医嘱' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,[3] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R6' As ID, '' As 上级id,NULL as EPRID, '医嘱报告' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[8] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R7' As ID, '' As 上级id,NULL as EPRID, '疾病证明' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[9] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R8' As ID, '' As 上级id,NULL as EPRID,'知情文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[10] As 排序 " & vbNewLine & _
        "       From Dual " & vbNewLine & _
        IIf(blnPath, " Union All Select 'R9' As ID, '' As 上级id,NULL as EPRID, '临床路径' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[11] As 排序 From Dual" & vbNewLine, "") & vbNewLine & _
        IIf(strType = "住院证", " Union All Select 'R10' As ID, '' As 上级id,NULL as EPRID, '住院证' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,'99' As 排序 From Dual" & vbNewLine, "")
    mstrSQL = mstrSQL & "Union All" & vbNewLine & _
                "       Select *" & vbNewLine & _
                "       From (Select a.上级id || 'K' || Trim(To_Char(a.Id)) || ',' || Trim(To_Char(Nvl(a.医嘱id, 0))) || ',0' As ID, 上级id," & vbNewLine & _
                "                     Decode(上级id, 'R6', b.诊疗类别, Null) As Eprid," & vbNewLine & _
                "                     Decode(a.医嘱id, Null, a.名称, '<' || b.医嘱内容 || '>' || a.名称) || '【创建:' ||" & vbNewLine & _
                "                      To_Char(a.创建时间, 'yyyy-mm-dd hh24:mi') || Decode(a.最后版本, 1, '书写：', '修订：') || a.保存人 || '在' ||" & vbNewLine & _
                "                      To_Char(a.保存时间, 'yyyy-mm-dd hh24:mi') || Decode(Nvl(a.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') ||" & vbNewLine & _
                "                      Decode(a.完成时间, Null, '】', ',完成:' || To_Char(a.完成时间, 'yyyy-mm-dd hh24:mi') || '】') As 名称," & vbNewLine & _
                "                     Trim(To_Char(a.Id)) || ';' || Decode(a.医嘱id, Null, '0', Trim(To_Char(a.医嘱id)))||';0;' || NVL(a.RISID,0) As 参数, 1 As 末级," & vbNewLine & _
                "                     Decode(病历种类, 2, 'object_case', 3, 'object_case', 7, 'object_report', 'object_file') As 图标, 排序" & vbNewLine & _
                "              From (Select a.Id, Decode(病历种类, 2, 'R2', 4, 'R3', 7, 'R6', 6, 'R8', 5, 'R7') As 上级id, 签名级别, 保存时间, 保存人, 最后版本," & vbNewLine & _
                "                            a.病历名称 As 名称, c.医嘱id,C.RISID, a.病历种类, a.创建时间, a.完成时间, To_Char(a.创建时间, 'yyyy-mm-dd hh24:mi:ss') As 排序" & vbNewLine & _
                "                     From 电子病历记录 A," & vbNewLine & _
                "                          (Select a.Id, Min(c.医嘱id) 医嘱id,MIN(c.RISID) as RISID " & vbNewLine & _
                "                            From 电子病历记录 A, 病人医嘱报告 C" & vbNewLine & _
                "                            Where a.病人id = [1] And a.主页id = [2] And a.Id = c.病历id(+)" & vbNewLine & _
                "                            Group By a.Id) C" & vbNewLine & _
                "                     Where a.Id = c.Id And 病历种类 In (2, 3, 4, 5, 6, 7)" & vbNewLine & _
                "                     Order By a.科室id) A, 病人医嘱记录 B" & vbNewLine & _
                "              Where a.医嘱id = b.Id(+)" & vbNewLine & _
                "              Order By b.诊疗类别 Desc)"
    
    '新版PACS报告编辑器
    mstrSQL = mstrSQL & " Union All " & vbNewLine & _
                        "Select ID, 上级id, 'D' as Eprid, 名称, 参数, 末级, 图标, 排序" & vbNewLine & _
                        "From (Select Rawtohex(b.检查报告id) ID, 'R6' 上级id, Null As Eprid, '<' || a.医嘱内容 || '>' 名称," & vbNewLine & _
                        "              '0;' || a.Id || ';' || b.检查报告id || ';' || NVL(b.RISID,0) 参数, 1 末级, 'object_report' 图标," & vbNewLine & _
                        "              To_Char(a.开嘱时间, 'yyyy-mm-dd hh24:mi:ss') As 排序" & vbNewLine & _
                        "       From 病人医嘱记录 A, 病人医嘱报告 B" & vbNewLine & _
                        "       Where a.病人id = [1] And a.主页id = [2] And a.Id = b.医嘱id And Rawtohex(b.检查报告id) Is Not Null" & vbNewLine & _
                        "       Order By a.开嘱时间)"
    
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "检查是否存在老版数据", lng病人ID, lng主页ID)
    If rsTemp.RecordCount > 0 Then
        mstrSQL = mstrSQL & " Union All" & vbNewLine & _
                            "Select 'R4K' || Trim(To_Char(a.Id)) || ',0,' || Trim(To_Char(a.科室id)) || ',' || Nvl(a.婴儿,0) As ID, 'R4' As 上级id," & vbNewLine & _
                            "              'R4K' || Trim(To_Char(a.Eprid)) || ',0,' || Trim(To_Char(a.科室id)) As Eprid," & vbNewLine & _
                            "              a.名称 || '(' || b.名称 || '：' || To_Char(a.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' ||" & vbNewLine & _
                            "               To_Char(a.截止, 'yyyy-mm-dd hh24:mi') || ')' As 名称," & vbNewLine & _
                            "              Trim(To_Char(b.Id)) || '_' || Nvl(a.婴儿,0) || ';' || Trim(To_Char(Nvl(保留, 0))) || ';' || To_Char(a.开始, 'yyyy-mm-dd hh24:mi') ||" & vbNewLine & _
                            "               ' ～ ' || To_Char(a.截止, 'yyyy-mm-dd hh24:mi') || ';' || Trim(To_Char(a.Id)) || ';' || Nvl(a.婴儿,0) As 参数," & vbNewLine & _
                            "              1 As 末级, 'object_tend' As 图标, To_Char(a.开始, 'yyyy-mm-dd hh24:mi:ss') As 排序" & vbNewLine & _
                            "       From (Select f.Id, f.Id As Eprid, f.编号, f.名称, r.开始, r.截止, r.科室id, f.保留, r.婴儿" & vbNewLine & _
                            "              From (Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & vbNewLine & _
                            "                     From 病历文件列表" & vbNewLine & _
                            "                     Where 种类 = 3 And 保留 < 0" & vbNewLine & _
                            "                     Union All" & vbNewLine & _
                            "                     Select l.Id, l.编号, l.名称, f.报表 As 护理级别, l.通用, a.科室id, l.保留" & vbNewLine & _
                            "                     From 病历文件列表 L, 病历页面格式 F, 病历应用科室 A" & vbNewLine & _
                            "                     Where l.种类 = 3 And l.保留 = 0 And l.种类 = f.种类 And l.编号 = f.编号 And l.Id = a.文件id(+)) F," & vbNewLine & _
                            "                   (Select r.科室id, r.婴儿, Nvl(Min(r.护理级别), 3) As 护理级别, Min(r.发生时间) As 开始, Max(r.发生时间) As 截止" & vbNewLine & _
                            "                     From 病人护理记录 R" & vbNewLine & _
                            "                     Where r.病人来源 = 2 And r.病人id = [1] And Nvl(r.主页id, 0) = [2]" & vbNewLine & _
                            "                     Group By r.科室id, r.婴儿) R" & vbNewLine & _
                            "              Where (f.通用 = 1 Or f.通用 = 2 And f.科室id = r.科室id) And f.护理级别 >= r.护理级别) A, 部门表 B" & vbNewLine & _
                            "       Where a.科室id = b.Id "
    Else
        '67839:刘鹏飞,2013-11-25,修正子查询F.ID该文R.ID，调用新版记录单、体温单应该传入的是文件ID,而不是格式ID
        mstrSQL = mstrSQL & " Union All" & _
                   " Select 'R4K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R4' As 上级id,'R4K'||Trim(To_Char(A.EPRID))||',0,'||Trim(To_Char(A.科室Id)) as EPRID ," & vbNewLine & _
                   "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                   "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿)) As 参数," & vbNewLine & _
                   "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                   " From (" & vbNewLine & _
                   "   Select R.ID,F.ID as EPRID , F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                   "   From (" & vbNewLine & _
                   "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                   "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                   "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                   "" & vbNewLine & _
                   "       ) F,(" & vbNewLine & _
                   "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                   "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                   "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                   "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                   "       ) R" & vbNewLine & _
                   "       Where F.ID=R.格式ID" & vbNewLine & _
                   "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿"
    End If
    
    If strType = "住院证" Then
        mstrSQL = mstrSQL & vbNewLine & "  Union All " & _
        " Select * From (Select ID||'' as ID, 上级id, Eprid, 名称, 参数, 末级, 图标, 排序" & vbNewLine & _
            "From (Select a.Id, 'R10' As 上级id, c.编号 As Eprid, c.名称||'【操作时间：'|| To_Char(D.发送时间, 'yyyy-mm-dd hh24:mi')||' 操作人：'||D.发送人||'】' as 名称," & vbNewLine & _
            " d.No || ';' || d.记录性质 As 参数, 1 As 末级, 'Folder' As 图标, To_Char(d.发送时间, 'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
            "       From 病人医嘱记录 A, 病人医嘱发送 D, 病历单据应用 B, 病历文件列表 C" & vbNewLine & _
            "       Where a.Id = d.医嘱id And a.病人id = [1] And a.诊疗项目id = b.诊疗项目id And b.病历文件id = c.Id And c.种类 = 7 And" & vbNewLine & _
            "             b.应用场合 = 1 And c.名称 Like '%住院证%'" & vbNewLine & _
            "       Order By D.发送时间 Desc)) Where Rownum<10 "

    End If
    
    mstrSQL = mstrSQL & ")" & vbNewLine & _
        " Order By Decode(上级id,Null,' ',上级id),排序"
        
    If lng病人ID > 0 Then
        '只处理住院病人
        If blnDataMove Then
            mstrSQL = Replace(mstrSQL, "电子病历记录", "H电子病历记录")
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
            mstrSQL = Replace(mstrSQL, "病人医嘱报告", "H病人医嘱报告")
            mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
        End If
    End If
    
    Set GetCISStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病人ID, lng主页ID, strSerial(1), strSerial(2), strSerial(3), strSerial(4), strSerial(5), strSerial(6), strSerial(7), strSerial(8), strSerial(9))
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FuncPrintBatch( _
    ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngDept As Long, ByVal strRange As String, ByVal strRegRange As String, _
    ByRef clsDockAduit As clsDockAduits, ByRef clsPath As clsDockPath, ByRef clsTendsNew As clsTendFile, ByVal blnPDF As Boolean, _
    ByVal strParPath As String, ByVal strName As String, ByVal lngInNo As Long, ByRef objfrmMain As Object, _
    ByRef strInfo As String, ByVal blnDataMove As Boolean, ByVal strPrinterName As String, ByVal blnDoctorAdvice As Boolean, _
    ByRef strPrintDocIDs As String, Optional ByVal rsMedRec As Recordset, Optional ByRef lngPage As Long, Optional ByVal strReport As String, _
    Optional ByRef objRichEMR As Object)
    '功能：多个病人批量打印
    '参数:
    'lngPatiID-病人ID
    'lngPageID-主页ID
    'lngDept-科室ID
    'strRange-勾选的电子病案分类
    'strRegRange -勾选的电子病案首页具体子项目
    'clsDockAduit-电子病历对象 zlRichEPR.clsDockAduits
    'clsPath-临床路径对象Set clsPath = New zlCISPath.clsDockPath
    'clsTendsNew -护理对象 传人前需要对它初始化 New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    'strParPath-PDF输出文件路径
    'blnPDF-是否PDF输出
    'strName-病人姓名
    'lngInNo-病人住院号
    'strFrmCaption-调用窗体名称
    'strInfo -打印进度提示信息
    'blnDataMove-T-转出数据,F-未转出数据
    'strPrinterName-打印设备名称
    'blnDoctorAdvice- True-医嘱单,False-医嘱本
    'strPrintDocIDs-共享病历子文档只打印一份儿
    'rsMedRec-单个病人打印时，可勾选具体打印文件
    'lngPage-用于返回打印成功的份数(首页\医嘱\临床路径算一份)
    'strReport -电子病案打印时,是否按指定报表打印检验检查报告,未设置对应报表时传空串。  格式:检验报表(报表名称,系统号,报表编号);检查(报表名称,系统号,报表编号)
    Dim intSel As Integer
    
    Dim lngNo As Long, i As Long
    Dim lng科室ID As Long
    
    Dim strPath As String
    Dim strReportNO As String, strFilter As String
    Dim strFileName As String, strEprName As String
    Dim str检验医嘱IDs As String, str检查医嘱IDs As String
    
    Dim rs As ADODB.Recordset
    Dim varParam As Variant, varTmp As Variant
    Dim objPacsDoc As Object
    
    Dim blnNewTends As Boolean
    Dim blnNEWLIS As Boolean, blnModel As Boolean
    On Error GoTo ErrH
    If blnPDF Then
        strPath = strParPath & "\" & strName & "_" & lngInNo
        If Not gobjFSO.FolderExists(strPath) Then
            Call gobjFSO.CreateFolder(strPath)
        End If
    Else
        lngNo = 1
        gstrSQL = "Select Nvl(Max(打印次数),0)+1 As 打印次数 From 病案打印记录 Where 病人id=[1] And 主页id=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, objfrmMain.Caption, lngPatiID, lngPageID)
        If rs.BOF = False Then
            lngNo = rs!打印次数
        End If
    End If
    
    '读取记录
    If rsMedRec Is Nothing Then
        Set rs = gclsPackage.GetCISStruct(lngPatiID, lngPageID, lngDept, blnDataMove)
        If strReport <> "" Then
            varParam = Split(strReport, ";")
            If UBound(varParam) = 2 Then
                blnNEWLIS = Val(varParam(2)) = 1
            End If
        End If
    Else
        Set rs = rsMedRec
        ' 格式:检验报表(报表名称,系统号,报表编号);检查(报表名称,系统号,报表编号)
        If strReport <> "" Then
            varParam = Split(strReport, ";")
            If varParam(0) <> "" Then '检验
                rs.Filter = "是否选择=1 And 上级ID='R6' And Eprid ='E'"
                str检验医嘱IDs = ""
                For i = 1 To rs.RecordCount
                    str检验医嘱IDs = str检验医嘱IDs & "," & Split(rs!参数 & "", ";")(1)
                    rs.MoveNext
                Next
                If str检验医嘱IDs <> "" Then
                    str检验医嘱IDs = Mid(str检验医嘱IDs, 2)
                    varTmp = Split(varParam(0), ",")
                    Call ReportOpen(gcnOracle, varTmp(1), varTmp(2), objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "医嘱ID=" & str检验医嘱IDs, 2)
                    lngPage = lngPage + (i - 1)
                End If
            Else
                strFilter = " OR (是否选择=1 And 上级ID='R6' And Eprid ='E')"
            End If
            If varParam(1) <> "" Then '检查
                rs.Filter = "是否选择=1 And 上级ID='R6' And Eprid ='D'"
                str检查医嘱IDs = ""
                For i = 1 To rs.RecordCount
                    str检查医嘱IDs = str检查医嘱IDs & "," & Split(rs!参数 & "", ";")(1)
                    rs.MoveNext
                Next
                If str检查医嘱IDs <> "" Then
                    str检查医嘱IDs = Mid(str检查医嘱IDs, 2)
                    varTmp = Split(varParam(1), ",")
                    Call ReportOpen(gcnOracle, varTmp(1), varTmp(2), objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "医嘱ID=" & str检查医嘱IDs, 2)
                    lngPage = lngPage + (i - 1)
                End If
            Else
                strFilter = strFilter & " OR (是否选择=1 And 上级ID='R6' And Eprid ='D')"
            End If
            rs.Filter = "(是否选择=1 And 上级ID<>'R6') OR (是否选择=1 And 上级ID=NULL)" & strFilter
            If UBound(varParam) = 2 Then
                blnNEWLIS = Val(varParam(2)) = 1
            End If
        Else
            rs.Filter = "是否选择=1"
        End If
        If rs.RecordCount > 0 Then rs.MoveFirst
    End If
    
    Do Until rs.EOF
        If NVL(rs("上级id").Value) = "" Then
            If InStr(strRange, "," & rs("ID").Value & ",") > 0 Then
                Select Case rs("ID").Value
                Case "R5"               '首页
                    lng科室ID = GetlngID(lngPatiID, lngPageID)
                    Select Case Val(zlDatabase.GetPara("病案首页标准", glngSys, 1261, "0"))
                    Case 0 '卫生部标准
                        If Have部门性质(lng科室ID, "中医科") Then
                            strReportNO = "ZL1_INSIDE_1261_4"
                        Else
                            strReportNO = "ZL1_INSIDE_1261_1"
                        End If
                    Case 1    '四川省标准
                        If Have部门性质(lng科室ID, "中医科") Then
                            strReportNO = "ZL1_INSIDE_1261_6"
                        Else
                            strReportNO = "ZL1_INSIDE_1261_5"
                        End If
                    Case 2    '云南省标准
                        If Have部门性质(lng科室ID, "中医科") Then
                            strReportNO = "ZL1_INSIDE_1261_8"
                        Else
                            strReportNO = "ZL1_INSIDE_1261_7"
                        End If
                    Case 3     '湖南省标准
                        If Have部门性质(lng科室ID, "中医科") Then
                            strReportNO = "ZL1_INSIDE_1261_10"
                        Else
                            strReportNO = "ZL1_INSIDE_1261_9"
                        End If
                    Case Else '当期修改时未定义
                        If Have部门性质(lng科室ID, "中医科") Then
                            strReportNO = "ZL1_INSIDE_1261_4"
                        Else
                            strReportNO = "ZL1_INSIDE_1261_1"
                        End If
                    End Select
                    strInfo = "开始输出" & strName & "首页"
                    
                    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & strReportNO, "Printer", IIf(blnPDF, "TinyPDF", strPrinterName))
                    If InStr("," & strRegRange & ",", ",5,") > 0 Then '正面
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页正面.PDF"
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=1", "PDF=" & strFileName, 4)
                        Else
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=1", 2)
                        End If
                    End If
                    
                    If InStr("," & strRegRange & ",", ",52,") > 0 Then '反面
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页反面.PDF"
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=2", "PDF=" & strFileName, 4)
                        Else
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=2", 2)
                        End If
                    End If
                    
                    If InStr("," & strRegRange & ",", ",53,") > 0 Then '附一
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页附页一.PDF"
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=3", "PDF=" & strFileName, 4)
                        Else
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=3", 2)
                        End If
                    End If
                    
                    If InStr("," & strRegRange & ",", ",54,") > 0 Then '附二
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页附页二.PDF"
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=4", "PDF=" & strFileName, 4)
                        Else
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "ReportFormat=4", 2)
                        End If
                    End If
                    
                    If Not blnPDF Then
                        Call RecordEprPrintInfo(2, "首页记录", lngNo, lngPatiID, lngPageID)
                    End If
                    lngPage = lngPage + 1
                Case "R1"               '医嘱
                    strInfo = "开始输出" & strName & "医嘱"
                    If blnDoctorAdvice Then
                        
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_医嘱.PDF"
                            Call zlCommFun.PDFFile(strFileName)
                        End If
                        '先打印长嘱
                        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_1", "Printer", IIf(blnPDF, "TinyPDF", strPrinterName))
                        Call gobjKernel.zlPrintAdvice(objfrmMain, lngPatiID, lngPageID, 0, 0, IIf(blnPDF, "TinyPDF", strPrinterName), 2)
                        
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_临嘱.PDF"
                            Call zlCommFun.PDFFile(strFileName)
                        End If
                        '再打印临嘱
                        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & "ZL1_INSIDE_1254_2", "Printer", IIf(blnPDF, "TinyPDF", strPrinterName))
                        Call gobjKernel.zlPrintAdvice(objfrmMain, lngPatiID, lngPageID, 0, 1, IIf(blnPDF, "TinyPDF", strPrinterName), 2)
                        
                        '记录打印
                        Call RecordEprPrintInfo(2, "住院医嘱", lngNo, lngPatiID, lngPageID)
                    Else
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_医嘱.PDF"
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, "ZL1_INSIDE_1560", objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, "PDF=" & strFileName, 4)
                        Else
                            Call ReportOpen(gcnOracle, ParamInfo.系统号, "ZL1_INSIDE_1560", objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, 2)
                        End If
                    End If
                    lngPage = lngPage + 1
                Case "R9"               '临床路径
                    strInfo = "开始输出" & strName & "临床路径"
                    If blnPDF Then
                        strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_临床路径.PDF"
                        Call clsPath.zlFuncPathTableOutPut(4, True, strFileName, lngPatiID, lngPageID, IIf(blnPDF, "TinyPDF", strPrinterName))
                    Else
                        Call clsPath.zlRefreshReadOnly(lngPatiID, lngPageID)
                        Call clsPath.zlFuncPathTableOutPut(1, True, "", 0, 0, IIf(blnPDF, "TinyPDF", strPrinterName))
                        Call RecordEprPrintInfo(2, "临床路径", lngNo, lngPatiID, lngPageID)
                    End If
                    lngPage = lngPage + 1
                End Select
            End If
        Else
            If InStr(strRange, "," & rs("上级id").Value & ",") > 0 Then
                If InStr(rs!ID, "R") = 0 And Len(rs!ID) >= 32 Then
                    'EMR病历预览
                    If Not objRichEMR Is Nothing Then
                        If InStr(rs!参数, "|") > 0 Then
                            Call objRichEMR.zlShowDoc(Split(rs!参数, "|")(0), Split(rs!参数, "|")(1))
                        Else
                            Call objRichEMR.zlShowDoc(rs!参数, "")
                        End If
                        Err.Number = 0: On Error Resume Next
                        Call objRichEMR.zlPrintDoc(False, strPrinterName)
                        If Err.Number = 0 Then lngPage = lngPage + 1
                        If Err.Number = 450 Then
                            Err.Number = 0
                            Call objRichEMR.zlPrintDoc(False)
                            If Err.Number = 0 Then lngPage = lngPage + 1
                        End If
                        Err.Clear: On Error GoTo 0
                    End If
                Else
                varParam = Split(rs("参数").Value, ";")

                Select Case rs("上级id").Value
                Case "R2"               '住院病历
                    If InStr("," & strPrintDocIDs, "," & Val(varParam(0)) & ",") = 0 Then '本次没打过
                        strEprName = Split(rs("名称").Value, "【")(0)
                        strInfo = "开始输出" & strName & "病历:" & "住院病历_" & Val(varParam(0))
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                            Call zlCommFun.PDFFile(strFileName)
                        End If
                        Call clsDockAduit.zlPrintDocument(3, 2, Val(varParam(0)), IIf(blnPDF, "TinyPDF", strPrinterName))
                        
                        If Not blnPDF Then
                            Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
                        End If
                    End If
                    lngPage = lngPage + 1
                Case "R3"               '护理病历
                    If InStr("," & strPrintDocIDs, "," & Val(varParam(0)) & ",") = 0 Then '本次没打过
                        strEprName = Split(rs("名称").Value, "【")(0)
                        strInfo = "开始输出" & strName & "病历:" & "护理病历_" & Val(varParam(0))
                        If blnPDF Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                            Call zlCommFun.PDFFile(strFileName)
                        End If
                        Call clsDockAduit.zlPrintDocument(3, 2, Val(varParam(0)), IIf(blnPDF, "TinyPDF", strPrinterName))
                        
                        If Not blnPDF Then
                            Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
                        End If
                    End If
                    lngPage = lngPage + 1
                Case "R4"               '护理记录
                    blnNewTends = Get新版护理(lngPatiID, lngPageID)
                    If blnNewTends = False Then
                        varParam = Split(rs("参数").Value, ";")
                        If UBound(varParam) >= 1 Then
                            If Val(varParam(1)) = -1 Then '体温单
                                strInfo = "开始输出" & strName & "体温单_" & Val(varParam(0))
                                If blnPDF Then
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_体温单_" & Val(varParam(0)) & ".PDF"
                                    Call zlCommFun.PDFFile(strFileName)
                                End If
                                Call clsDockAduit.zlRefreshTendBody(lngPatiID, lngPageID, Val(Split(varParam(0), "_")(0)), Val(varParam(4)))
                                Call clsDockAduit.zlPrintDocument(1, 2, , IIf(blnPDF, "TinyPDF", strPrinterName))
                                
                                If Not blnPDF Then
                                    Call RecordEprPrintInfo(2, "体温单", lngNo, lngPatiID, lngPageID)
                                End If
                            Else '护理记录
                                strInfo = "开始输出" & strName & "护理记录" & Val(varParam(3))
                                If blnPDF Then
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_护理记录_" & Val(varParam(3)) & ".PDF"
                                    Call zlCommFun.PDFFile(strFileName)
                                End If
                                Call clsDockAduit.zlRefresh(3, Val(varParam(3)), lngPatiID, lngPageID, Val(Split(varParam(0), "_")(0)), CStr(varParam(2)), , Val(varParam(4)))
                                Call clsDockAduit.zlPrintDocument(2, 2, , IIf(blnPDF, "TinyPDF", strPrinterName))
                                
                                If Not blnPDF Then
                                    Call RecordEprPrintInfo(3, Val(varParam(3)), lngNo, lngPatiID, lngPageID)
                                End If
                            End If
                        End If
                    Else
                        '此参数保存 保留
                        varParam = Split(rs("参数").Value, ";")
                        If UBound(varParam) >= 1 Then
                            Select Case Val(varParam(1))
                                Case -1 '体温单
                                    intSel = 1
                                Case 1  '产程图
                                    intSel = 3
                                Case Else '记录单
                                    intSel = 2
                            End Select
                            strInfo = "开始输出" & strName & Decode(intSel, 1, "体温单", 2, "护理记录", "产程图") & "_" & Val(varParam(3))
                            If blnPDF Then
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & Decode(intSel, 1, "体温单", 2, "护理记录", "产程图") & "_" & Val(varParam(3)) & ".PDF"
                                Call zlCommFun.PDFFile(strFileName)
                            End If
                            Call clsTendsNew.zlPrintDocument(lngPatiID, lngPageID, Val(varParam(4)), Val(varParam(0)), Val(varParam(3)), intSel, IIf(blnPDF, "TinyPDF", strPrinterName))
                            
                            If Not blnPDF Then
                                Call RecordEprPrintInfo(2, Decode(intSel, 1, "体温单", 2, "护理记录", "产程图"), lngNo, lngPatiID, lngPageID)
                            End If
                        End If
                    End If
                    lngPage = lngPage + 1
                Case "R6"               '医嘱报告
                    strEprName = Split(rs("名称").Value, "【")(0)
                    If UBound(Split(strEprName, ">")) > 0 Then
                        strEprName = Split(strEprName, ">")(1)
                    End If
                    If Val(varParam(0)) = 0 Then
                        strInfo = "开始输出" & strName & "医嘱报告" & varParam(2)
                    Else
                        strInfo = "开始输出" & strName & "医嘱报告" & Val(varParam(0))
                    End If
                        blnModel = False
                        If blnNEWLIS And UCase(rs!Eprid & "") = "E" Then
                            '新版LIS
                            If InitObjLis(False) Then
                                blnModel = gobjLIS.PrintReport(objfrmMain, Val(varParam(1)), lngPatiID, 2, strInfo)
                                If blnModel Then lngPage = lngPage + 1
                            End If
                        End If
                        If Not blnModel Then
                    If blnPDF Then
                        If Val(varParam(0)) = 0 Then
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & varParam(2) & ".PDF"
                        Else
                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                        End If
                        Call zlCommFun.PDFFile(strFileName)
                    End If
                    
                    If Val(varParam(3)) <> 0 Then  'Val(varParam(3)) RISID
                    '新网RIS 返回值 0-成功
                        If Not gobjXWHIS Is Nothing Then
                            Call gobjXWHIS.HISPrintReportByAppno(objfrmMain.hWnd, Val(varParam(1)), Val(varParam(3)), strPrinterName)
                        End If
                    ElseIf Val(varParam(0)) <> 0 Then
                        Call clsDockAduit.zlPrintDocument(4, 2, Val(varParam(0)), IIf(blnPDF, "TinyPDF|" & strFileName, strPrinterName))
                        If Not blnPDF Then
                                Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
                        End If
                    Else
                        If objPacsDoc Is Nothing Then
                            Set objPacsDoc = DynamicCreate("zlPublicPACS.clsPublicPacs", "新版PACS编辑器", False)
                            Call objPacsDoc.InitInterface(gcnOracle, gstrDBUser)
                        End If
                        Call objPacsDoc.PrintReport(varParam(2), IIf(blnPDF, "TinyPDF", strPrinterName))
                    End If
                    lngPage = lngPage + 1
                        End If
                Case "R7"               '疾病证明
                    strEprName = Split(rs("名称").Value, "【")(0)
                    strInfo = "开始输出" & strName & "疾病证明" & Val(varParam(0))
                    If blnPDF Then
                        strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName)
                    End If
                    Call clsDockAduit.zlPrintDocument(3, 2, Val(varParam(0)), IIf(blnPDF, "TinyPDF", strPrinterName))
                    
                    If Not blnPDF Then
                        Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
                    End If
                    lngPage = lngPage + 1
                Case "R8"               '知情文件
                    strEprName = Split(rs("名称").Value, "【")(0)
                    strInfo = "开始输出" & strName & "知情文件" & Val(varParam(0))
                    If blnPDF Then
                        strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                        Call zlCommFun.PDFFile(strFileName)
                    End If
                    Call clsDockAduit.zlPrintDocument(3, 2, Val(varParam(0)), IIf(blnPDF, "TinyPDF", strPrinterName))
                    
                    If Not blnPDF Then
                        Call RecordEprPrintInfo(1, Val(varParam(0)), lngNo)
                    End If
                    lngPage = lngPage + 1
                Case "R10"       '电子病案打印 住院证
                    strInfo = "开始输出" & Split(rs!名称 & "", "【")(0)
                    strEprName = "ZLCISBILL" & Format(rs!Eprid, "00000") & "-1"
                    If UBound(varParam) >= 1 Then
                        Call ReportOpen(gcnOracle, glngSys, strEprName, objfrmMain, "NO=" & varParam(0), "性质=" & varParam(1), "医嘱ID=0", "PrintEmpty=1", 2)
                    End If
                    lngPage = lngPage + 1
                Case "R11"   '电子病案打印 其他报表
                    strInfo = "开始输出" & rs!名称 & ""
                    If UBound(varParam) >= 2 Then
                        strReportNO = varParam(2)
                        Call ReportOpen(gcnOracle, 0, strReportNO, objfrmMain, "病人id=" & lngPatiID, "主页id=" & lngPageID, 2)
                        lngPage = lngPage + 1
                    End If
                End Select
            End If
        End If
        End If
        rs.MoveNext
    Loop
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset, strExtendTag As String, strReturn As String
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    gstrSQL = "Select Decode(e.Kind, '02', 'R2', '03', 'R3', '04', 'R7', '05', 'R8', 'R2') 上级id," & vbNewLine & _
                "       Nvl(d.Subdoc_Id, Rawtohex(b.Id)) As ID, d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       Nvl(d.Subdoc_Title, b.Title) ||" & vbNewLine & _
                "        Decode(d.Completor, Null, ''," & vbNewLine & _
                "               '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称," & vbNewLine & _
                "       Rawtohex(b.Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As 参数, 'object_case' As 图标,B.TITLE " & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select Distinct a.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
                "       From (Select ID" & vbNewLine & _
                "              From Bz_Act_Log" & vbNewLine & _
                "              Where Extend_Tag = :etag" & vbNewLine & _
                "              Union" & vbNewLine & _
                "              Select b.Id" & vbNewLine & _
                "              From Bz_Act_Log F, Bz_Act_Log B" & vbNewLine & _
                "              Where f.Extend_Tag = :etag And b.Basiclog_Id = f.Id) A, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Id = c.Actlog_Id And c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And d.Real_Doc_Id = b.Id And d.Antetype_Id = e.Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Id, Null, d.Antetype_Id, b.Antetype_Id) = b.Antetype_Id And" & vbNewLine & _
                "      Decode(d.Subdoc_Title, Null, e.Title, d.Subdoc_Title) = e.Title" & vbNewLine & _
                "Order By e.Code, b.Creat_Time, d.Complete_Time"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtendTag & "^" & DbType.T_String & "^etag", rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, ParamInfo.产品名称
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEmrLIST(ByVal lngPatiID As Long, ByVal lngPageID As Long) As ADODB.Recordset
'功能:电子病案打印新版病历列表专用
    Dim rsTemp As New ADODB.Recordset, strExtendTag As String, strReturn As String
    
    If gobjEmr Is Nothing Then Set GetEmrLIST = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageID)
    If strExtendTag = "" Then Set GetEmrLIST = Nothing: Exit Function

    '上级ID，ID，名称，参数，图标
    gstrSQL = "Select Rawtohex(b.Id) As ID, b.Title As 名称,b.Creat_Time" & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select ID" & vbNewLine & _
                "       From Bz_Act_Log" & vbNewLine & _
                "       Where Extend_Tag = :etag" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select b.Id From Bz_Act_Log F, Bz_Act_Log B Where f.Extend_Tag = :etag And b.Basiclog_Id = f.Id) D," & vbNewLine & _
                "     Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And b.Antetype_Id = e.Id" & vbNewLine & _
                "Order By e.Code, b.Creat_Time"

    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtendTag & "^" & DbType.T_String & "^etag", rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, ParamInfo.产品名称
        Set GetEmrLIST = Nothing: Exit Function
    End If
    
    Set GetEmrLIST = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetQuestion(rsCondition As ADODB.Recordset, ByVal strDept As String, ByVal bytApplyMode As Byte, Optional ByVal lngKey As Long, Optional ByVal strStart As String, Optional ByVal strEnd As String, Optional ByVal lng病人ID As Long = -1, Optional ByVal lng主页ID As Long, Optional ByVal lngCur次数 As Long, Optional ByVal str开始时间 As String, Optional ByVal str结束时间 As String, Optional ByVal str反馈人 As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strFilter As String
    
    Dim varState As Variant
    
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    
    '形成查病人的子ＳＱＬ语句
    mstrSQL = "" '每次清空全局变量
    
    If bytApplyMode > 1 Then
        '1.审查状态
        '------------------------------------------------------------------------------------------------------------------
        
        strSQL = _
            "Select B.病人id, B.主页id, B.险类,b.出院科室id" & vbNewLine & _
            "From (Select B.病人id, B.主页id, B.险类,b.出院科室id " & vbNewLine & _
            "       From 病案提交记录 C, 病案主页 B" & vbNewLine & _
            "       Where C.提交时间 Between [6] And [7] And C.病人id = B.病人id And C.主页id = B.主页id And" & vbNewLine & _
            "             B.病案状态 In ([12], [13], [14], [15])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select B.病人id, B.主页id, B.险类,b.出院科室id " & vbNewLine & _
            "       From 病案提交记录 C, 病案主页 B" & vbNewLine & _
            "       Where C.提交时间 Between [8] And [9] And C.病人id = B.病人id And C.主页id = B.主页id And B.病案状态 = 5) B"
        
        If ParamRead(rsCondition, "出院情况") <> "" Then
            strSQL = strSQL & " Where Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[16])"
        End If
        
        '2.在院,出院
        '------------------------------------------------------------------------------------------------------------------
        
        If ParamRead(rsCondition, "出院情况") = "" Then
            
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And a.在院=1 "
                    
            If ParamRead(rsCondition, "当前病况") <> "" Then strSQL = strSQL & " And b.当前病况=[17] "
                        
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And b.出院日期 Between [10] And [11]"
                    
        Else
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And b.出院日期 Between [10] And [11] " & vbNewLine & _
                        " And Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[16])"
        End If
        
        strSQL = "Select b.病人id,b.主页id From (" & strSQL & ") b,Table (Cast(f_Num2List([18]) As zlTools.t_NumList)) f Where b.出院科室id=f.Column_Value "
        Select Case Val(ParamRead(rsCondition, "病人类型"))
        Case 1          '非医保病人
            strSQL = strSQL & " And b.险类 Is Null "
        Case 2          '医保病人
            strSQL = strSQL & " And b.险类 Is Not Null "
            If ParamRead(rsCondition, "医保种类") <> "" Then
                strSQL = strSQL & " And b.险类 In (" & ParamRead(rsCondition, "医保种类") & ") "
            End If
        End Select
        
        strTmp = Val(ParamRead(rsCondition, "接收待审")) & ";" & Val(ParamRead(rsCondition, "拒绝接收")) & ";" & Val(ParamRead(rsCondition, "正在审查")) & ";" & Val(ParamRead(rsCondition, "审查反馈"))
        varState = Split(strTmp, ";")
        If Val(varState(0)) = 1 Then varState(0) = 1
        If Val(varState(1)) = 1 Then varState(1) = 2
        If Val(varState(2)) = 1 Then varState(2) = 3
        If Val(varState(3)) = 1 Then varState(3) = 4
    
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    strTmp = "Decode(a.反馈对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件',9,'临床路径') As 反馈对象"
    
    '按次数查找在院审查zq
    If lngCur次数 = 0 Then
        If str反馈人 = "" Then
            strFilter = "A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = "A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    ElseIf lngCur次数 = 1 Then
         If str反馈人 = "" Then
            strFilter = " (A.反馈次数 is null or A.反馈次数 =[19]) And A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = " (A.反馈次数 is null or A.反馈次数 =[19]) And A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    Else
        If str反馈人 = "" Then
            strFilter = " A.反馈次数 =[19] And A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = " A.反馈次数 =[19] And A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    End If
    
    mstrSQL = ""
    Select Case bytApplyMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1, 0                     '指定
        gstrSQL = "Select length(文件ID) 长度 From 病案反馈记录 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lngKey)
        If Not rsTemp.EOF Then
        If NVL(rsTemp!长度, 0) = 32 Then
            mstrSQL = "Select Decode(a.反馈对象, 1, 'object_advice', 2, 'object_case', 3, 'object_case', 4, 'object_tend', 5, 'object_first', 6," & vbNewLine & _
                    "               'object_report', 'object_file') As 图标, a.Id, a.相关id, a.提交id, a.病人id, a.主页id, a.反馈对象 As 反馈对象id," & vbNewLine & _
                    "       Decode(a.反馈对象, 1, '住院医嘱', 2, '住院病历', 3, '护理病历', 4, '护理记录', 5, '首页记录', 6, '医嘱报告', 7, '疾病证明', 8, '知情文件', 9, '临床路径') As 反馈对象," & vbNewLine & _
                    "       a.记录性质, a.记录状态, a.反馈意见, a.反馈项目id, a.反馈人, a.反馈时间, a.处理期限, a.评分级别, a.处理说明, a.处理人, a.处理时间, a.反馈记录 As 文件名称," & vbNewLine & _
                    "       a.文件id || Decode(a.子文档id, Null, '', '|') || Nvl(a.子文档id, '') 文件id, a.医嘱id, a.科室id, a.分制, a.分值, a.补充说明, a.反馈次数," & vbNewLine & _
                    "       a.反馈记录, e.姓名, f.名称 As 科室" & vbNewLine & _
                    "From 病案反馈记录 A, 病案主页 D, 病人信息 E, 部门表 F" & vbNewLine & _
                    "Where a.Id = [1] And a.病人id = d.病人id And a.主页id = d.主页id And e.病人id = d.病人id And f.Id = d.出院科室id"
        End If
        End If
        
        If mstrSQL = "" Then
            mstrSQL = _
                "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,a.ID, a.相关id, a.提交id, a.病人id, a.主页id,a.反馈对象 As 反馈对象id," & strTmp & ", a.记录性质, a.记录状态, a.反馈意见, a.反馈项目id, a.反馈人, a.反馈时间, a.处理期限,a.评分级别," & vbNewLine & _
                "       a.处理说明, a.处理人, a.处理时间,Decode(a.反馈对象,4,b.名称,c.病历名称) As 文件名称,a.文件id,a.医嘱id,a.科室id,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录,e.姓名,f.名称 As 科室 " & vbNewLine & _
                "From 病案反馈记录 A,病历文件列表 b,电子病历记录 c,病案主页 d,病人信息 e,部门表 f " & vbNewLine & _
                "Where A.ID = [1] And a.文件id=b.ID(+) And a.文件id=c.ID(+) And a.病人id=d.病人id And a.主页id=d.主页id And e.病人id=d.病人id And f.ID=d.出院科室id"
        End If
    Case 2                      '等待修改
        
        mstrSQL = _
            "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈记录,a.反馈次数, C.姓名,D.名称 As 科室" & vbNewLine & _
            "From 病案反馈记录 A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
            "Where " & strFilter & "A.记录状态 = 1 And A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And d.ID=b.出院科室id"
        If lng病人ID > -1 Then
            mstrSQL = mstrSQL & " And a.病人id=[4] And a.主页id=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3                      '等待复查

        mstrSQL = _
            "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈记录,a.反馈次数, C.姓名,D.名称 As 科室" & vbNewLine & _
            "From 病案反馈记录 A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
            "Where A.记录状态 = 2 And A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"

        If lng病人ID > -1 Then
            mstrSQL = mstrSQL & " And a.病人id=[4] And a.主页id=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 4                      '结束问题
        
        If lng病人ID > -1 Then
            mstrSQL = _
                "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录, C.姓名,D.名称 As 科室" & vbNewLine & _
                "From (Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈记录 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] And a.记录状态 = 3" & IIf(lng病人ID > -1, " And a.病人id=[4] And a.主页id=[5] ", "") & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈历史 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] " & IIf(lng病人ID > -1, " And a.病人id=[4] And a.主页id=[5] ", "") & ") A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"
        Else
            mstrSQL = _
                "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录, C.姓名,D.名称 As 科室" & vbNewLine & _
                "From (Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈记录 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] And a.记录状态 = 3" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈历史 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3]) A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"
            
            mstrSQL = "Select * From (" & mstrSQL & ") A Where (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
        
    End Select
    
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    
    If bytApplyMode > 1 Then
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng病人ID, lng主页ID, CDate(ParamRead(rsCondition, "审查开始时间")), CDate(ParamRead(rsCondition, "审查结束时间")), CDate(ParamRead(rsCondition, "归档开始时间")), CDate(ParamRead(rsCondition, "归档结束时间")), CDate(ParamRead(rsCondition, "出院开始时间")), CDate(ParamRead(rsCondition, "出院结束时间")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "出院情况"), ParamRead(rsCondition, "当前病况"), strDept, lngCur次数, CDate(str开始时间), CDate(str结束时间)) ', UCase(str反馈人)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng病人ID, lng主页ID, CDate(ParamRead(rsCondition, "审查开始时间")), CDate(ParamRead(rsCondition, "审查结束时间")), CDate(ParamRead(rsCondition, "归档开始时间")), CDate(ParamRead(rsCondition, "归档结束时间")), CDate(ParamRead(rsCondition, "出院开始时间")), CDate(ParamRead(rsCondition, "出院结束时间")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "出院情况"), ParamRead(rsCondition, "当前病况"), strDept, lngCur次数, CDate(str开始时间), CDate(str结束时间))
        End If
    Else
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng病人ID, lng主页ID)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng病人ID, lng主页ID)
        End If
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetAduitPatientSafeKeep(ByVal strTxt As String, ByVal strStart As String, ByVal strEnd As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取封存病案
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim varState    As Variant
    Dim strTmp      As String
    On Error GoTo errHand
    If strTxt = "" Then
        strTmp = ""
    Else
        strTmp = " And f.封存理由 Like '%" & strTxt & "%'"
    End If
    
    mstrSQL = _
       "Select 0 as 选择," & vbNewLine & _
        "       Decode(b.封存时间,Null, Decode(b.病案状态,1,'waitpatient',10,'waitpatient',2,'refusepatient',3,'aduitpatient',4,'questionpatient',5,'archivepatient', 6,'dealpatient',13, 'aduitpatient2',14,'questionpatient2',16,'dealpatient2'),'safekeeping') As 图标," & vbNewLine & _
        "       Decode(b.路径状态,Null,'',0,'','lujin') as 路径,b.数据转出,Decode(b.病案状态,1,'等待审查',10,'接收待审', 2, '拒绝审查', 3,'正在审查',4,'审查反馈',5,'审查归档',6,'审查整改', 13, '正在抽查',14,'抽查反馈',16,'抽查整改') As 审查状态,Decode(B.出院日期,Null,'在院病人','出院病人') as 病人状态," & vbNewLine & _
        "       b.出院科室id,C.ID,c.提交人,c.提交时间,c.接收人,c.接收时间,c.归档人,c.归档时间,c.拒审人,c.拒审时间,c.拒审理由,A.病人ID,B.主页ID,A.就诊卡号,B.住院号,A.姓名,A.性别,b.年龄,nvl(b.住院医师, b.门诊医师) as 住院医师,e.名称 as 护理等级, LPAD(B.出院病床, 10, ' ') as 床号,B.入院日期, B.出院日期,b.病案状态 as 病案状态值,b.封存时间,d.名称 As 出院科室,f.封存理由,f.封存人 " & vbNewLine & _
        "  From 病案提交记录 c, 病人信息 a, 病案主页 b, 部门表 d, 护理等级 e,病案封存记录 f" & vbNewLine & _
        " Where c.病人id(+) = b.病人id And c.主页id(+) = b.主页id And A.病人ID = B.病人ID  And b.病人id =f.病人id and b.主页id =f.主页id  And f.记录状态=1 " & strTmp & " And  d.ID = b.出院科室id And b.护理等级ID = e.序号(+) And b.封存时间 is not null And b.封存时间 Between [2] And  [3]  order by b.出院日期"
        
    Set GetAduitPatientSafeKeep = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, CStr(strTxt), CDate(strStart), CDate(strEnd))
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEMRFile(ByVal strKey As String) As ADODB.Recordset
Dim strDocid As String, strSubDocid As String, strReturn As String, rsTemp As ADODB.Recordset
    On Error GoTo errHand
    If gobjEmr Is Nothing Then Exit Function
    If InStr(strKey, "|") = 0 Then
        strDocid = strKey
    Else
        strDocid = Split(strKey, "|")(0)
        strSubDocid = Split(strKey, "|")(1)
    End If
    
    gstrSQL = "Select Nvl(b.Subdoc_Title, a.Title) 名称" & vbNewLine & _
                "From Bz_Doc_Log A, Bz_Doc_Tasks B" & vbNewLine & _
                "Where a.Id = Hextoraw(:docid) And a.Id = b.Real_Doc_Id" & IIf(strSubDocid = "", "", " And b.Subdoc_Id =:subdocid")
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strDocid & "^" & DbType.T_String & "^docid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^" & DbType.T_String & "^subdocid"), rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, ParamInfo.产品名称
        Set GetEMRFile = Nothing: Exit Function
    End If
    
    Set GetEMRFile = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function GetEPRFile(ByVal lngKey As Long, Optional ByVal lng医嘱id As Long) As ADODB.Recordset
    On Error GoTo errHand
    If lng医嘱id > 0 Then
        mstrSQL = "Select  '<'||c.医嘱内容||'>' || a.病历名称  As 名称 From 电子病历记录 a,病人医嘱报告 b,病人医嘱记录 c Where a.ID=[1] And a.ID=b.病历ID And b.医嘱id=[2] And b.医嘱id=c.ID"
    Else
        mstrSQL = "Select 病历名称 As 名称 From 电子病历记录 a Where a.ID=[1]"
    End If
    
    Set GetEPRFile = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, lng医嘱id)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEPRFileStruct(ByVal lngKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "Select 名称 From 病历文件列表 a Where a.ID=[1]"

    Set GetEPRFileStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDocumentLocation(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "Select 病案状态 as 病案状态值,Decode(出院日期,Null,Decode(状态,3,1,0),1) As 状态 From 病案主页 a Where a.病人id=[1] And a.主页id=[2]"
    Set GetDocumentLocation = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病人ID, lng主页ID)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetGradeStandard(Optional ByVal strFilter As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    If strFilter = "" Then
        mstrSQL = _
            "Select A.ID, A.上级id, Decode(A.上级id, Null, A.名称, A.描述) As 名称, Decode(A.上级id, Null, 0, 1) As 末级" & vbNewLine & _
            "From 病案评分标准 A, 病案评分方案 B" & vbNewLine & _
            "Where A.方案id = B.ID And B.选用 = 1" & vbNewLine & _
            "Start With A.上级id Is Null" & vbNewLine & _
            "Connect By Prior A.ID = A.上级id"
    Else
                        
        mstrSQL = _
            "Select A.ID, A.描述 As 名称" & vbNewLine & _
            "From 病案评分标准 A, 病案评分方案 B" & vbNewLine & _
            "Where A.方案id = B.ID And B.选用 = 1 And A.上级id Is Not Null And A.描述 Like [1]" & vbNewLine & _
            "Order By a.上级id,a.序号"

        strFilter = "%" & strFilter & "%"
                    
    End If
    
    Set GetGradeStandard = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, strFilter)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetApplyMode(ByVal StrText As String) As Byte
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    If CheckStrType(StrText, 1) And Left(ParamInfo.收费诊疗项目匹配, 1) = 1 Then
        '是全数字，按编码查找
            
        GetApplyMode = 1
        
    ElseIf CheckStrType(StrText, 2) And Left(ParamInfo.收费诊疗项目匹配, 2) = 1 Then
        '是全字母，按简码查找
        GetApplyMode = 2
    Else
        GetApplyMode = 3
    End If
End Function


Public Function GetDictTableData(ByVal strTable As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    Select Case strTable
    Case "借阅理由"
        mstrSQL = "Select 编码,名称,简码,缺省标志,编码||'-'||名称 As 显示名称 From 借阅理由"
    End Select
    
    Set GetDictTableData = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMedicalExits() As ADODB.Recordset
    '******************************************************************************************************************
    '功能:检查是否安装了病案系统
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    mstrSQL = "Select 编号 From zlsystems where 编号=300"
    
    Set GetMedicalExits = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetHavePath(ByVal lng部门ID As Long) As Boolean
'功能：检查指定科室或病区是否有可用的临床路径
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select a.Id" & vbNewLine & _
            "From 临床路径目录 A, 临床路径版本 B, 临床路径科室 C," & vbNewLine & _
            "     (Select 科室id From 病区科室对应 Where 病区id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From 部门表 Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.路径id And a.最新版本 = b.版本号 And a.Id = c.路径id(+) And (c.科室id = d.科室id or c.科室id is null) And Rownum < 2"
    On Error GoTo ErrH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng部门ID)
    GetHavePath = Not rsTmp.EOF
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetProjectUse(ByVal lng方案ID As Long) As ADODB.Recordset
'功能:检查该方案是否在病案反馈记录中已经被使用过
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(A.ID) as 条数 From 病案审查目录 A,病案审查分类 B,病案审查方案 C,病案反馈记录 D" & vbNewLine & _
            "Where A.分类ID= B.ID And B.方案ID = C.ID And A.id =D.反馈项目ID And C.ID=[1] And Rownum >0"
    Set GetProjectUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng方案ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemUse(ByVal lng项目ID As Long) As ADODB.Recordset
'功能:检查该项目是否在病案反馈记录中已经被使用过
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(*) as 条数 From 病案反馈记录 where 反馈项目ID=[1]"
    Set GetItemUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng项目ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRelevanceID(ByVal lng关联ID As Long) As ADODB.Recordset
'功能:检查该反馈记录中是否有相关联的ID,获取最主要反馈记录。
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select 相关ID From 病案反馈记录 Where ID =[1]"
    Set GetRelevanceID = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng关联ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetFeedback(ByVal str反馈人 As String) As String
'功能:反馈人查询串
    Dim strTemp() As String
    Dim lngRow As Long
    Dim strFeedback As String
    Dim strTempSql As String
    If InStrRev(str反馈人, ",", -1) Then
        strTemp = Split(str反馈人, ",")
        strFeedback = ""
        For lngRow = 0 To UBound(strTemp)
            strFeedback = strFeedback & "'" & strTemp(lngRow) & "'" & ","
        Next
    End If
    
    If strFeedback <> "" Then
        If Right(strFeedback, 1) = "," Then
              strTempSql = " A.反馈人 in (" & Left(strFeedback, Len(strFeedback) - 1) & ") And "
              GetFeedback = strTempSql
        End If
    Else
        If str反馈人 <> "" Then
            strTempSql = " A.反馈人 = '" & str反馈人 & "' And "
            GetFeedback = strTempSql
        End If
    End If
End Function

Public Function GetExamineStartUse() As Boolean
'功能:检查是否有启动了的审查方案
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select ID From 病案审查方案 Where 启用时间 is not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    If rsTmp.RecordCount > 0 Then
        GetExamineStartUse = True
    Else
        GetExamineStartUse = False
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Class_Initialize()
    mstrTitle = "数据包类"
    
End Sub

Private Sub Class_Terminate()
    Set mrsTmp = Nothing
End Sub


