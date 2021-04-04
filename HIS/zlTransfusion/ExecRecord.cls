VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExecRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ExecutiveGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function GetExecGroups(ByVal objPati As cPatient, ByVal lng执行部门ID As Long, ByVal int执行类别 As Integer, _
    ByVal dateBeging As Date, ByVal DateEnd As Date) As Boolean
    
    Dim strSQL As String, i As Integer
    Dim rsExecGroups As ADODB.Recordset, rsGroup As ADODB.Recordset, rsBill As ADODB.Recordset
    Dim objexecGroup As ExecutiveGroup, objGroup As Group, objBIlls As Bills, lng时间 As Long, Date上次执行时间 As Date
    Dim objBIll As Bill, lng总耗时 As Long, lng液体量 As Long, int已执行次数 As Integer
    Dim rsTmp As ADODB.Recordset, str说明 As String, lng_Count As Long
    Dim int计费状态 As Integer, cur金额 As Currency
    
    On Error GoTo errHandle
    
    If objPati.病人来源 = 1 Then
        '门诊留观
        strSQL = "Select distinct A.流水号, A.配药人, A.执行时间, A.组数, A.接单人, A.滴系数," & _
                 "  Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗') As 执行分类,d.计费状态" & vbNewLine & _
                 "From 病人医嘱发送 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱执行 A" & vbNewLine & _
                 "Where A.医嘱id = D.医嘱id And A.发送号 = D.发送号 And A.医嘱id = B.ID And B.诊疗项目id = C.ID And B.主页id=[5] And " & vbNewLine & _
                 "   Nvl(A.组次,0) <= 1 And (A.流水号 is not Null or A.流水号<>0) And D.发送时间 Between [3] And [4] And " & _
                 "   D.执行部门id = [2] And B.病人id = [1] " & vbCr & _
                 "Order by A.流水号 "
        Set rsExecGroups = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", objPati.病人ID, lng执行部门ID, dateBeging, DateEnd, objPati.单据ID)
    Else
        strSQL = "Select distinct A.流水号, A.配药人, A.执行时间, A.组数, A.接单人, A.滴系数,Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗') As 执行分类,d.计费状态" & vbNewLine & _
                "From 病人医嘱发送 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱执行 A" & vbNewLine & _
                "Where A.医嘱id = D.医嘱id And A.发送号 = D.发送号 And A.医嘱id = B.ID And B.诊疗项目id = C.ID And B.挂号单=[5] And " & vbNewLine & _
                "   Nvl(A.组次,0) <= 1 And (A.流水号 is not Null or A.流水号<>0) And D.发送时间 Between [3] And [4] And D.执行部门id = [2] And B.病人id = [1] Order by A.流水号 "
        Set rsExecGroups = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", objPati.病人ID, lng执行部门ID, dateBeging, DateEnd, objPati.挂号单)
    End If
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next
    Do Until rsExecGroups.EOF
    
        Set objexecGroup = New ExecutiveGroup
        objexecGroup.流水号 = zlCommFun.NVL(rsExecGroups.Fields("流水号"), 0)
        objexecGroup.配药人 = zlCommFun.NVL(rsExecGroups.Fields("配药人"))
        objexecGroup.执行时间 = zlCommFun.NVL(rsExecGroups.Fields("执行时间"), CDate(0))
        objexecGroup.组数 = zlCommFun.NVL(rsExecGroups.Fields("组数"), 0)
        objexecGroup.接单人 = zlCommFun.NVL(rsExecGroups.Fields("接单人"))
        objexecGroup.滴系数 = zlCommFun.NVL(rsExecGroups.Fields("滴系数"), 20)
        objexecGroup.执行分类 = zlCommFun.NVL(rsExecGroups.Fields("执行分类"), "0-治疗")
        
        
        lng总耗时 = 0
        
        '2012-07-17 显示零费用
        strSQL = "Select d.流水号, d.组次, d.滴速, d.液体量, a.医嘱id, d.执行人,d.核对人, a.发送号, a.发送数次, d.本次数次, b.上次执行时间, d.要求时间," & vbNewLine & _
                "       c.操作类型,Decode(c.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗') As 执行分类, d.滴速, b.医嘱内容 As 用法, b.皮试结果, b.执行频次, b.频率次数," & vbNewLine & _
                "       a.执行状态, a.计费状态, b.收费细目id, a.记录性质, b.执行性质, Sum(Nvl(Decode(e.记录状态,9,0,3,0,e.实收金额), 0)) As 金额, e.记录状态" & vbNewLine & _
                "From 病人医嘱执行 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A, 门诊费用记录 E" & vbNewLine & _
                "Where a.医嘱id = d.医嘱id And a.发送号 = d.发送号 And b.诊疗项目id = c.Id And a.医嘱id = b.Id And c.类别 = 'E' And a.执行部门id = [2] And" & vbNewLine & _
                "      a.发送时间 Between [3] And [4] And d.流水号 = [5] And b.病人来源 = [6] And b.病人id = [1] And a.医嘱id = e.医嘱序号(+) And" & vbNewLine & _
                "      a.记录序号 = e.序号(+) And a.No = e.No(+) And e.记录状态(+) <> 2" & vbNewLine & _
                "Group By d.流水号, d.组次, d.滴速, d.液体量, a.医嘱id, d.执行人,d.核对人, a.发送号, a.发送数次, d.本次数次, b.上次执行时间, d.要求时间," & vbNewLine & _
                "         c.操作类型,Decode(c.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗'), d.滴速, b.医嘱内容, b.皮试结果, b.执行频次, b.频率次数, a.执行状态, a.计费状态," & vbNewLine & _
                "         b.收费细目id, a.记录性质, b.执行性质, e.记录状态" & vbNewLine & _
                "Order By d.组次"

        Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", _
                                               objPati.病人ID, lng执行部门ID, dateBeging, DateEnd, objexecGroup.流水号, _
                                               IIf(objPati.病人来源 = 1, 2, 1))
        With rsGroup
        Do Until .EOF
            strSQL = "Select A.说明 From 病人医嘱执行 A Where A.医嘱id = [1] And A.发送号 = [2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", Val(.Fields("医嘱ID")), Val(.Fields("发送号")))
            str说明 = ""
            lng_Count = 0
            Do Until rsTmp.EOF
                lng_Count = lng_Count + 1
                str说明 = str说明 & rsTmp!说明
                If lng_Count <> rsTmp.RecordCount Then
                    str说明 = str说明 & vbNewLine
                End If
                rsTmp.MoveNext
            Loop
            
            Select Case .Fields("执行分类")
            Case "0-治疗"
                '治疗  区分给药途径和非给药途径
                If Val(.Fields("操作类型") & "") = 2 Then
                    '当是给药途径时和 输液,注射 查询的SQL相同
                    strSQL = "Select 滴速, 液体量, 类别, Id, Rownum 顺序, 执行频次, 医嘱内容, 单次用量, 计算单位, 单次容量, 执行频次,实收金额,计费状态,记录状态,执行性质" & vbNewLine & _
                            "From (Select Distinct d.滴速, d.液体量, c.类别, b.Id, b.医嘱内容, b.单次用量, c.计算单位," & vbNewLine & _
                            "        Decode(Sign(Instr(',5,6,', ',' || c.类别 || ',')), 1, ( b.单次用量/g.剂量系数) * g.容量, 0) As 单次容量," & vbNewLine & _
                            "        b.执行频次 , b.序号,Decode(E.记录状态,9,0,3,0,E.实收金额) as 实收金额,f.计费状态,e.记录状态,b.执行性质" & vbNewLine & _
                            "      From 病人医嘱发送 f,门诊费用记录 e,病人医嘱执行 d, 药品规格 g, 诊疗项目目录 c, 病人医嘱记录 b, 病人医嘱发送 a" & vbNewLine & _
                            "      Where b.id=f.医嘱id and f.发送号=[5] and f.医嘱id=E.医嘱序号(+) And F.NO=E.NO(+) And E.记录状态(+)<>2 and a.医嘱id = d.医嘱id And a.发送号 = d.发送号 And b.收费细目id = g.药品id(+) And b.诊疗项目id = c.Id And" & vbNewLine & _
                            "        a.医嘱id = b.相关id And a.执行部门id = [1] And" & vbNewLine & _
                            "        a.发送时间 Between [2] And [3] And a.医嘱id = [4] And a.发送号 = [5] And" & vbNewLine & _
                            "        b.病人来源 = [8] And b.病人id = [6] And D.流水号=[7]" & vbNewLine & _
                            "      Order By b.序号 ) "
                Else
                    strSQL = "Select 液体量,类别,Id,Rownum 顺序,执行频次,医嘱内容,单次用量,计算单位,单次容量,执行频次,实收金额,计费状态,记录状态,执行性质 " & vbNewLine & _
                            "From (" & vbNewLine & _
                            "  Select Distinct 0 as 液体量,C.类别, B.ID, B.执行频次, B.医嘱内容 || '；' || b1.医嘱内容 医嘱内容, B1.单次用量, C.计算单位, " & _
                            "      0 As 单次容量, B.序号, Decode(D.记录状态,9,0,3,0,D.实收金额) as 实收金额,a.计费状态,d.记录状态,b.执行性质" & vbNewLine & _
                            "  From  门诊费用记录 D,诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱记录 B1, 病人医嘱发送 A" & vbNewLine & _
                            "  Where A.医嘱id=D.医嘱序号(+) And A.NO=D.NO(+) and D.记录状态(+)<>2 and B1.诊疗项目id = C.ID(+) And A.医嘱id = B.ID And " & vbNewLine & _
                            "      A.医嘱id = B1.相关ID(+) And A.执行部门id = [1] And A.发送时间 Between [2] And [3] And A.医嘱id = [4] And " & vbNewLine & _
                            "      A.发送号 = [5] And B.病人来源 = [8] And B.病人id = [6]" & vbNewLine & _
                            "  Order By B.序号 ) "
                End If
            Case "3-皮试"
                '皮试
                strSQL = "Select  液体量,类别,Id,Rownum 顺序,执行频次,医嘱内容,单次用量,计算单位,单次容量,执行频次,实收金额,计费状态,记录状态,执行性质" & vbNewLine & _
                        "From (" & vbNewLine & _
                        "  Select Distinct 0 as 液体量,C.类别, B.ID,  B.执行频次, B.医嘱内容, B.单次用量, C.计算单位, 0 As 单次容量, B.序号, " & _
                        "    Decode(D.记录状态,9,0,3,0,D.实收金额) as 实收金额,a.计费状态,d.记录状态,b.执行性质" & vbNewLine & _
                        "  From  门诊费用记录 D,诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                        "  Where A.医嘱id=D.医嘱序号(+) And A.NO=D.NO(+) and D.记录状态(+)<>2 and B.诊疗项目id = C.ID And A.医嘱id = B.ID And A.执行部门id = [1] And" & vbNewLine & _
                        "      A.发送时间 Between [2] And [3] And A.医嘱id = [4] And A.发送号 = [5] And B.病人来源 = [8] And" & vbNewLine & _
                        "      B.病人id = [6]" & vbNewLine & _
                        "  Order By B.序号 ) "
            Case Else
                '输液,注射
                strSQL = "Select 滴速, 液体量, 类别, Id, Rownum 顺序, 执行频次, 医嘱内容, 单次用量, 计算单位, 单次容量, 执行频次,实收金额,计费状态,记录状态,执行性质" & vbNewLine & _
                        "From (Select Distinct d.滴速, d.液体量, c.类别, b.Id, b.医嘱内容, b.单次用量, c.计算单位," & vbNewLine & _
                        "        Decode(Sign(Instr(',5,6,', ',' || c.类别 || ',')), 1, ( b.单次用量/g.剂量系数) * g.容量, 0) As 单次容量," & vbNewLine & _
                        "        b.执行频次 , b.序号,Decode(E.记录状态,9,0,3,0,E.实收金额) as 实收金额,f.计费状态,e.记录状态,b.执行性质" & vbNewLine & _
                        "      From 病人医嘱发送 f,门诊费用记录 e,病人医嘱执行 d, 药品规格 g, 诊疗项目目录 c, 病人医嘱记录 b, 病人医嘱发送 a" & vbNewLine & _
                        "      Where b.id=f.医嘱id and f.发送号=[5] and f.医嘱id=E.医嘱序号(+) And F.NO=E.NO(+) And E.记录状态(+)<>2 and a.医嘱id = d.医嘱id And a.发送号 = d.发送号 And b.收费细目id = g.药品id(+) And b.诊疗项目id = c.Id And" & vbNewLine & _
                        "        a.医嘱id = b.相关id And a.执行部门id = [1] And" & vbNewLine & _
                        "        a.发送时间 Between [2] And [3] And a.医嘱id = [4] And a.发送号 = [5] And" & vbNewLine & _
                        "        b.病人来源 = [8] And b.病人id = [6] And D.流水号=[7]" & vbNewLine & _
                        "      Order By b.序号 ) "

            End Select
            Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "ExecRecord.GetExecGroups", lng执行部门ID, dateBeging, DateEnd, _
                                                  Val(.Fields("医嘱ID")), Val(.Fields("发送号")), _
                                                  objPati.病人ID, objexecGroup.流水号, _
                                                  IIf(objPati.病人来源 = 1, 2, 1))
            
            Set objBIlls = New Bills
            lng时间 = 0
            lng液体量 = 0
            Do Until rsBill.EOF
                If rsBill.Fields("顺序") = 1 Then
                    lng时间 = CacleTransTime(Val(rsBill.Fields("液体量")), objexecGroup.滴系数, .Fields("滴速"))
                    lng液体量 = Val(rsBill.Fields("液体量"))
                Else
                    lng时间 = 0
                    lng液体量 = 0
                End If
                int计费状态 = Val("" & rsBill.Fields("计费状态"))

                cur金额 = Val("" & rsBill.Fields("实收金额"))
                
                If Val("" & rsBill.Fields("记录状态")) = 0 Then cur金额 = 0
                
                If Val("" & rsBill.Fields("执行性质")) = 5 And int计费状态 <> -1 Then
                    int计费状态 = -1 '院外带药的，等同为不需计费
                ElseIf Val("" & rsBill.Fields("记录状态")) = 1 And cur金额 = 0 Then
                    int计费状态 = -2 '零费用
                ElseIf Val("" & rsBill.Fields("记录状态")) = 3 Then
                    int计费状态 = -3 '已退费
                End If
                
                Call objBIlls.Add(objPati.病人ID, Val(rsBill.Fields("ID")), zlCommFun.NVL(rsBill.Fields("医嘱内容")), zlCommFun.NVL(rsBill.Fields("单次用量"), 0), zlCommFun.NVL(rsBill.Fields("计算单位")), lng液体量, _
                             lng时间, str说明, Val(rsBill.Fields("顺序")), cur金额, int计费状态, Val(.Fields("医嘱ID")) & "_" & rsBill.Fields("ID"))
                rsBill.MoveNext
            Loop
            Date上次执行时间 = zlCommFun.NVL(.Fields("上次执行时间"), CDate(0))
            
            For Each objBIll In objBIlls
                lng总耗时 = lng总耗时 + objBIll.时间
            Next
            
            strSQL = "Select Sum(nvl(本次数次,0)) as 已执行数次 From 病人医嘱执行 Where 医嘱id=[1] And 发送号=[2]"
            Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "", Val(.Fields("医嘱ID")), Val(.Fields("发送号")))
            If rsBill.RecordCount > 0 Then
                int已执行次数 = IIf(IsNull(rsBill.Fields("已执行数次")), 0, Val(rsBill.Fields("已执行数次"))) - Val(.Fields("本次数次"))
            End If
            

            int计费状态 = Val("" & .Fields("计费状态"))
            cur金额 = Val("" & .Fields("金额"))
            
            If Val("" & .Fields("执行性质")) = 5 And int计费状态 <> -1 Then
                int计费状态 = -1 '院外带药的，等同为不需计费
            ElseIf Val("" & .Fields("记录状态")) = 1 And cur金额 = 0 Then
                int计费状态 = -2 '零费用
            ElseIf Val("" & .Fields("记录状态")) = 3 Then
                int计费状态 = -3 '已退费
            End If
            
            If objexecGroup.Item(Val(.Fields("医嘱ID")) & "_" & .Fields("发送号")) Is Nothing Then
                Call objexecGroup.Add(Val(.Fields("医嘱ID")), Val(.Fields("发送号")), Val(.Fields("执行状态")), zlCommFun.NVL(.Fields("执行人")), CStr(.Fields("用法")), zlCommFun.NVL(.Fields("频率次数"), 0), _
                                      Val(.Fields("滴速")), .Fields("执行频次"), zlCommFun.NVL(.Fields("皮试结果")), CStr(.Fields("执行分类")), int已执行次数, _
                                      Val(.Fields("发送数次")), Val(.Fields("本次数次")), Val(.Fields("组次")), objBIlls, Date上次执行时间, CDate(.Fields("要求时间")), _
                                      int计费状态, Val("" & .Fields("记录性质")), "" & .Fields("核对人"), Val(.Fields("医嘱ID")) & "_" & .Fields("发送号"))
            End If
            
            .MoveNext
        Loop
        End With
        objexecGroup.总耗时 = lng总耗时
        
        Call Add(objexecGroup, zlCommFun.NVL(rsExecGroups.Fields("流水号"), 0))
        rsExecGroups.MoveNext
    Loop
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal ObjExecutiveGroup As ExecutiveGroup, ByVal sKey As String) As ExecutiveGroup
    '创建新对象
    Dim objNewMember As ExecutiveGroup
    Set objNewMember = ObjExecutiveGroup
    On Error GoTo errH
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
errH:
    If Err.Number = 457 Then
        Exit Function
    ElseIf ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Property Get Item(vntIndexKey As Variant) As ExecutiveGroup
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub



