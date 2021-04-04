VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Group"
Attribute VB_Ext_KEY = "Member0" ,"Group"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvar选择组数 As Integer
Public p本次执行时间 As Date
Public p执行摘要 As String

Public p流水号 As String
Public p配药人 As String
Public p滴系数 As Integer
Public p接单人 As String
Public p耗时 As Integer
Public p提醒 As Integer

'局部变量，保存集合
Private mCol As Collection

Public Function GetGroups(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long, ByVal int执行类别 As Integer, _
        ByVal datAdviceB As Date, ByVal datAdviceE As Date, ByVal str挂号单 As String, _
        ByVal strKey As String, ByVal bytPatiFrom As Byte) As Boolean
'功能: 取可执行的项目加入对象
    
    Dim strSQL As String
    Dim rsGroups As New ADODB.Recordset
    Dim rsBill As New ADODB.Recordset
    Dim objBIlls As Bills, cur金额 As Currency
   
    Dim lng已执行数次 As Long, dat上次执行时间 As Date
    Dim int默认滴速 As Integer, lng时间 As Long
    Dim str医生滴速  As String, strTmp As String
    Dim int计费状态  As Integer
    
    On Error GoTo errHandle
    int默认滴速 = Val(zlDatabase.GetPara("默认滴速", glngSys, 1264))
    If int默认滴速 < 10 Or int默认滴速 > 100 Then int默认滴速 = 40
    '取所有的执行单,加到groups对象中
    
    If bytPatiFrom = 1 Then
        '门诊留观
        strSQL = "Select A.医嘱id, A.发送号, A.发送数次,c.操作类型,Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗') As 执行分类," & vbNewLine & _
                " B.医嘱内容 As 用法, B.皮试结果, B.执行频次, B.频率次数, A.执行状态, A.计费状态,B.收费细目id, b.医生嘱托, D.记录状态, B.执行性质, " & vbNewLine & _
                " Sum(Decode(Sign(D.记录状态), 0, 0, Decode(d.费用状态,1,0,nvl(D.实收金额,0)))) 实收金额" & vbNewLine & _
                " From 门诊费用记录 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                " Where  A.NO=D.No(+) And A.医嘱id = D.医嘱序号(+) And B.诊疗项目id = C.ID And A.医嘱id = B.ID And C.类别='E' And A.执行部门id = [2] " & vbNewLine & _
                " And nvl(C.执行分类,0) = [3] And D.记录状态(+)<>2 And A.发送时间 Between [4] And [5] And B.病人来源 = 2 " & vbNewLine & _
                " And B.病人id = [1] and b.主页id=[6] " & vbNewLine & _
                " Group By A.医嘱id, A.发送号, A.发送数次,c.操作类型,Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗'), B.医嘱内容," & vbNewLine & _
                " B.皮试结果, B.执行频次, B.频率次数, A.执行状态, A.计费状态,B.收费细目id, b.医生嘱托, D.记录状态, B.执行性质" & _
                " Order by 医嘱ID"
        Set rsGroups = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", lng病人ID, lng执行部门ID, int执行类别, datAdviceB, datAdviceE, Val(Split(strKey, "_")(1)))
    Else
        strSQL = "Select A.医嘱id, A.发送号, A.发送数次,c.操作类型,Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗') As 执行分类," & vbNewLine & _
                " B.医嘱内容 As 用法, B.皮试结果, B.执行频次, B.频率次数, A.执行状态, A.计费状态,B.收费细目id, b.医生嘱托, D.记录状态, B.执行性质, " & vbNewLine & _
                " Sum(Decode(Sign(D.记录状态), 0, 0, Decode(d.费用状态,1,0,nvl(D.实收金额,0)))) 实收金额" & vbNewLine & _
                " From 门诊费用记录 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                " Where  A.NO=D.No(+) And A.医嘱id = D.医嘱序号(+) And B.诊疗项目id = C.ID And A.医嘱id = B.ID And C.类别='E' And " & _
                " A.执行部门id = [2] And B.挂号单=[6] And nvl(C.执行分类,0) = [3] And " & vbNewLine & _
                " D.记录状态(+)<>2 And A.发送时间 Between [4] And [5] And B.病人来源 = 1 And B.病人id = [1]" & vbNewLine & _
                " Group By A.医嘱id, A.发送号, A.发送数次,c.操作类型,Decode(C.执行分类, 1, '1-输液', 2, '2-注射', 3, '3-皮试', '0-治疗'), B.医嘱内容," & vbNewLine & _
                " B.皮试结果, B.执行频次, B.频率次数, A.执行状态, A.计费状态,B.收费细目id, b.医生嘱托, D.记录状态, B.执行性质" & _
                " Order by 医嘱ID"
        Set rsGroups = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", lng病人ID, lng执行部门ID, int执行类别, datAdviceB, datAdviceE, str挂号单)
    End If
            
    cur金额 = 0
    Do Until rsGroups.EOF
        lng已执行数次 = 0
        dat上次执行时间 = CDate(0)
        '在加到group对象之前,还要取药品单据和执行情况,存到bills和execrecords中
        
        '添加bills对象
        Select Case int执行类别
        Case 0
            '治疗  区分给药途径和非给药途径
            If Val(rsGroups.Fields("操作类型") & "") = 2 Then
                '当是给药途径时和 输液,注射 查询的SQL相同
                strSQL = "Select C.类别, B.ID, Rownum 顺序, B.执行频次, B.医嘱内容, B.单次用量, C.计算单位," & vbNewLine & _
                        "       Decode(Sign(Instr(',5,6,', ',' || C.类别 || ',')), 1, ( B.单次用量/G.剂量系数) * G.容量, 0) As 单次容量," & vbNewLine & _
                        "       B.执行频次, B.医生嘱托,D.记录状态,Decode(D.费用状态,1,0,D.实收金额) as 实收金额,E.计费状态,B.执行性质" & vbNewLine & _
                        "From 病人医嘱发送 E,门诊费用记录 D,药品规格 G, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                        "Where B.id=E.医嘱ID and E.NO=D.NO(+) And  E.医嘱id = D.医嘱序号(+) And D.记录状态(+)<>2 And B.收费细目id = G.药品id(+) And B.诊疗项目id = C.ID" & _
                        " And A.医嘱id = B.相关id And A.医嘱id=[1] And A.发送号=[2] And A.执行部门id=[3] and E.发送号=[2]" & vbNewLine & _
                        "Order By B.序号"
            Else
                strSQL = "Select m.类别,m.ID,m.执行频次,m.医嘱内容,m.单次用量,m.计算单位,m.医生嘱托,m.记录状态,m.实收金额,m.计费状态,m.执行性质,Rownum As 顺序,0 As 单次容量 " & vbNewLine & _
                        "From (Select C.类别, B.Id, B.执行频次, B.医嘱内容 || '；' || B1.医嘱内容 医嘱内容, B1.单次用量, C.计算单位, B.医生嘱托, D.记录状态, " & _
                        "             Sum(Decode(D.费用状态,1,0,D.实收金额)) As 实收金额, A.计费状态, B.执行性质 " & vbNewLine & _
                        "       From 门诊费用记录 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱记录 B1, 病人医嘱发送 A " & vbNewLine & _
                        "       Where A.No = D.No(+) And A.医嘱id = D.医嘱序号(+) And D.记录状态(+) <> 2 And B1.诊疗项目id = C.Id(+) And " & vbNewLine & _
                        "             A.医嘱id = B.Id And A.医嘱id = B1.相关ID(+) And A.医嘱id=[1] And A.发送号=[2] And A.执行部门id=[3]" & vbNewLine & _
                        "       Group By C.类别, B.Id, B.执行频次, B.医嘱内容, B1.单次用量, C.计算单位, B.执行频次, B.医生嘱托, D.记录状态, A.计费状态, B.执行性质, b1.医嘱内容" & vbNewLine & _
                        "      ) M " & vbNewLine & _
                        "Order By 顺序 "
            End If            
        Case 3
            '皮试
            strSQL = "Select m.类别,m.ID,m.执行频次,m.医嘱内容,m.单次用量,m.计算单位,m.医生嘱托,m.记录状态,m.实收金额,m.计费状态,m.执行性质, Rownum As 顺序, 0 As 单次容量" & vbNewLine & _
                    "From (Select C.类别, B.Id, B.执行频次, B.医嘱内容, B.单次用量, C.计算单位, B.医生嘱托, D.记录状态, " & _
                    "             Sum(Decode(D.费用状态,1,0,D.实收金额)) As 实收金额, A.计费状态, B.执行性质" & vbNewLine & _
                    "       From 门诊费用记录 D, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                    "       Where A.No = D.No(+) And A.医嘱id = D.医嘱序号(+) And D.记录状态(+) <> 2 And B.诊疗项目id = C.Id And A.医嘱id = B.Id" & vbNewLine & _
                    "            And A.医嘱id=[1] And A.发送号=[2] And A.执行部门id=[3] " & vbNewLine & _
                    "       Group By C.类别, B.Id, B.执行频次, B.医嘱内容, B.单次用量, C.计算单位, B.执行频次, B.医生嘱托, D.记录状态, A.计费状态, B.执行性质) M" & vbNewLine & _
                    "Order By 顺序"
        
        Case Else
            '输液,注射
            strSQL = "Select C.类别, B.ID, Rownum 顺序, B.执行频次, B.医嘱内容, B.单次用量, C.计算单位," & vbNewLine & _
                    "       Decode(Sign(Instr(',5,6,', ',' || C.类别 || ',')), 1, ( B.单次用量/G.剂量系数) * G.容量, 0) As 单次容量," & vbNewLine & _
                    "       B.执行频次, B.医生嘱托,D.记录状态,Decode(D.费用状态,1,0,D.实收金额) as 实收金额,E.计费状态,B.执行性质" & vbNewLine & _
                    "From 病人医嘱发送 E,门诊费用记录 D,药品规格 G, 诊疗项目目录 C, 病人医嘱记录 B, 病人医嘱发送 A" & vbNewLine & _
                    "Where B.id=E.医嘱ID and E.NO=D.NO(+) And  E.医嘱id = D.医嘱序号(+) And D.记录状态(+)<>2 And B.收费细目id = G.药品id(+) And B.诊疗项目id = C.ID" & _
                    " And A.医嘱id = B.相关id And A.医嘱id=[1] And A.发送号=[2] And A.执行部门id=[3] and E.发送号=[2]" & vbNewLine & _
                    "Order By B.序号"

        End Select
        If int执行类别 = 1 Then
            '医生嘱托 如有 X滴/分钟 的，据此更改默认滴速
            str医生滴速 = "" & rsGroups.Fields("医生嘱托")
            If InStr(str医生滴速, "滴/分钟") > 0 Then
                If Val(str医生滴速) > 10 And Val(str医生滴速) < 100 Then
                    int默认滴速 = Val(str医生滴速)
                End If
            End If
        End If
        
        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", Val(rsGroups.Fields("医嘱ID")), Val(rsGroups.Fields("发送号")), lng执行部门ID)
        With rsBill
            Set objBIlls = New Bills
            Do Until .EOF
                
                lng时间 = CacleTransTime(Val(zlCommFun.NVL(.Fields("单次容量"), 0)), p滴系数, int默认滴速)
                int计费状态 = Val("" & .Fields("计费状态"))
                If Val("" & .Fields("执行性质")) = 5 And int计费状态 <> -1 Then
                    int计费状态 = -1 '院外带药的，等同为不需计费
                ElseIf int计费状态 > 0 Then
                    '已计费的,判断是否零费用
                    If .Fields("记录状态") = 1 And Val("" & .Fields("实收金额")) = 0 Then
                        int计费状态 = -2
                    ElseIf .Fields("记录状态") = 3 Then
                        int计费状态 = -3    '已退费
                    End If
                End If
                Call objBIlls.Add(lng病人ID, Val(.Fields("ID")), zlCommFun.NVL(.Fields("医嘱内容")), zlCommFun.NVL(.Fields("单次用量"), 0), _
                            zlCommFun.NVL(.Fields("计算单位")), Val(zlCommFun.NVL(.Fields("单次容量"), 0)), _
                            lng时间, zlCommFun.NVL(.Fields("医生嘱托")), Val(.Fields("顺序")), _
                            IIf(Val("" & .Fields("记录状态")) = 0, 0, Val("" & .Fields("实收金额"))), _
                            int计费状态, Val(rsGroups.Fields("医嘱ID")) & "_" & .Fields("ID"))
                
                .MoveNext
            Loop
        End With
        

        strSQL = "Select  max(执行时间) as 上次执行时间,sum(nvl(B.本次数次,0)) as 执行数次" & vbNewLine & _
                "From 病人医嘱执行 B" & vbNewLine & _
                "Where B.医嘱id = [1] And B.发送号 = [2] "
        Set rsBill = zlDatabase.OpenSQLRecord(strSQL, "Groups.getGroups", Val(rsGroups.Fields("医嘱ID")), Val(rsGroups.Fields("发送号")))
        With rsBill

        Do Until .EOF
            lng已执行数次 = zlCommFun.NVL(.Fields("执行数次"), 0)
            dat上次执行时间 = zlCommFun.NVL(.Fields("上次执行时间"), CDate(0))
            .MoveNext
        Loop
        End With
        
        '将可执行的项目 加入group对象
        With rsGroups
            If Val(.Fields("发送数次")) - lng已执行数次 > 0 Then
                '加上 执行项目(如静脉输液) 本身的金额
                cur金额 = Val("" & .Fields("实收金额"))
                int计费状态 = Val("" & .Fields("计费状态"))
                
                If Val("" & .Fields("执行性质")) = 5 And int计费状态 <> -1 Then
                    int计费状态 = -1 '院外带药的，等同为不需计费
                ElseIf int计费状态 > 0 Then
                    '已计费的,判断是否零费用
                    If .Fields("记录状态") = 1 And cur金额 = 0 Then
                        int计费状态 = -2       '
                    ElseIf .Fields("记录状态") = 3 Then
                        int计费状态 = -3
                    End If
                End If
                If Me.Item(CStr(.Fields("医嘱ID") & "_" & .Fields("发送号"))) Is Nothing Then
                    Call Add(.Fields("医嘱ID"), .Fields("发送号"), Val(.Fields("发送数次")), lng已执行数次, dat上次执行时间, _
                              .Fields("执行分类"), zlCommFun.NVL(.Fields("用法")), zlCommFun.NVL(.Fields("皮试结果")), _
                              zlCommFun.NVL(.Fields("执行频次"), ""), zlCommFun.NVL(.Fields("频率次数"), 0), _
                              int默认滴速, zlCommFun.NVL(.Fields("执行状态"), 0), cur金额, objBIlls, int计费状态, _
                              .Fields("医嘱ID") & "_" & .Fields("发送号"))
                              
                    Call GetOtherValue(.Fields("医嘱ID"), .Fields("发送号"))
                End If
           End If
        End With
        
        rsGroups.MoveNext
    Loop
    GetGroups = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal 执行医嘱ID As Long, ByVal 发送号 As Long, ByVal 发送数次 As Integer, ByVal 已执行数次 As Integer, ByVal 上次执行时间 As Date, _
        ByVal 执行分类 As String, ByVal 用法 As String, ByVal 皮试结果 As String, ByVal 执行频次 As String, ByVal 频率次数 As Integer, _
        ByVal 滴速 As Integer, ByVal 执行状态 As Integer, ByVal 收费金额 As Currency, ByVal clsBills As Bills, ByVal 计费状态 As Integer, _
        Optional sKey As String) As Group
    '创建新对象
    Dim objNewMember As Group, objNewBills As Bills
    Set objNewMember = New Group
    
    '设置传入方法的属性
    objNewMember.执行医嘱ID = 执行医嘱ID
    objNewMember.发送号 = 发送号
    objNewMember.发送数次 = 发送数次
    objNewMember.已执行数次 = 已执行数次
    objNewMember.上次执行时间 = 上次执行时间
    objNewMember.执行分类 = 执行分类
    objNewMember.用法 = 用法
    objNewMember.皮试结果 = 皮试结果
    objNewMember.执行频次 = 执行频次
    objNewMember.频率次数 = 频率次数
    objNewMember.执行状态 = 执行状态
    objNewMember.滴速 = 滴速
    objNewMember.收费金额 = 收费金额
    objNewMember.计费状态 = 计费状态
    If clsBills Is Nothing Then
        Set objNewBills = New Bills
    Else
        Set objNewBills = clsBills
    End If
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
        Call objNewMember.AddBills(objNewBills, 0)
    Else
        mCol.Add objNewMember, sKey
        Call objNewMember.AddBills(objNewBills, sKey)
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Let 选择组数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行医嘱ID = 5
    mvar选择组数 = vData
End Property

Public Property Get 选择组数() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.执行医嘱ID
    选择组数 = mvar选择组数
End Property

Public Property Get Item(vntIndexKey As Variant) As Group
Attribute Item.VB_UserMemId = 0
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
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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

Public Sub CheckGroup(ByVal strBillKey As String, blnCheck As Boolean)
    '选择了其中一组数据时执行
    Dim lng组次 As Long
    If Me.选择组数 < 0 Then Me.选择组数 = 0
    Me.选择组数 = Me.选择组数 + IIf(blnCheck, 1, -1)
    If blnCheck Then
        Me.Item(strBillKey).组次 = Me.选择组数
    Else
        lng组次 = Me.Item(strBillKey).组次
        Me.Item(strBillKey).组次 = 0
    End If
    
    If Me.选择组数 > 0 Then
        Call Sort(lng组次)
    End If
    
End Sub

Private Sub Sort(ByVal lng组次 As Long)
    Dim bln要减 As Boolean
    Dim objGroup As Group
    Dim lng组 As Long
    For Each objGroup In Me
        If objGroup.组次 > Me.选择组数 Then
            bln要减 = True
        End If
    Next
    If bln要减 Then
        For Each objGroup In Me
            If objGroup.组次 <> 0 Then
                If Me.Item(objGroup.执行医嘱ID & "_" & objGroup.发送号).组次 > lng组次 Then
                    lng组 = Me.Item(objGroup.执行医嘱ID & "_" & objGroup.发送号).组次
                    Me.Item(objGroup.执行医嘱ID & "_" & objGroup.发送号).组次 = IIf(objGroup.组次 - 1 < 0, 0, IIf(objGroup.组次 = 1, 1, objGroup.组次 - 1))
                End If
            End If
        Next
        If lng组 > 0 Then
            Call Sort(lng组)
        End If
    End If
End Sub

Public Function SelectGroupThingNew(ByVal lngExcDeptID As Long, Optional ByVal blnPrint As Boolean, _
            Optional ByVal intType As Integer, Optional ByVal frmParent As Object, _
            Optional ByVal blnView As Boolean, Optional ByVal blnLabel As Boolean = False, _
            Optional ByVal blnWristband As Boolean = False) As Boolean
'功能:    记录执行情况，保存已选择的执行内容到病人医嘱执行表
'参数
'  lngExcDeptID : 执行部门ID 47323
'  blnPrint：打印
'  intType：报表的序号；
'  frmParent：主窗体对象
'  blnView：True预览；False不预览
'  blnLabel：True打印输液瓶签；False不打印
'  blnWristband：True打印输液腕带；False不打印
'返回：True成功；Fals失败

    Dim objGroup As Group
    Dim strSQL As String, lng流水号 As Long, rsTmp As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim strRpt As String
    Dim objBIlls As Bills, i As Integer, Y As Integer
    Dim objBIll As Bill
    Dim cur数量 As Currency
    Dim arrSQL As Variant
    
    On Error GoTo errH
    
    '取流水号
    strSQL = "Select 病人医嘱执行_流水号.nextval as 流水号 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Group.SelectGroupThingNew")
    If rsTmp.EOF Then
        MsgBox "取流水号失败，请检查！", vbInformation, gstrSysName
    Else
        lng流水号 = zlCommFun.NVL(rsTmp.Fields("流水号"), 1)
    End If
    
    arrSQL = Array()
    
    For Each objGroup In Me
        With objGroup
        If .组次 > 0 Then
            '0-未执行;1-完全执行;2-拒绝执行;3-正在执行(今后可能分解为若干实际步骤)
            If .执行状态 = 1 Then
                MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '本次执行时间应晚于上次执行时间
            Call GetOtherValue(.执行医嘱ID, .发送号)
            If .本次数次 <= 0 Then .本次数次 = 1
            If p本次执行时间 <= .上次执行时间 Then
                MsgBox "本次执行时间应晚于上次执行时间 " & Format(.上次执行时间, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If .发送数次 > .已执行数次 Then
                '执行人不填，在完成时填，填了表示完成
                'ZL_病人医嘱执行_Insert(医嘱ID,发送号,要求时间,本次数次,执行摘要,执行人,执行时间)
                strSQL = "ZL_病人医嘱执行_Insert(" & .执行医嘱ID & "," & .发送号 & "," & _
                    "To_Date('" & Format(.要求时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    Val(.本次数次) & ",'" & DelInvalidChar(p执行摘要, "'") & "',''," & _
                    "To_Date('" & Format(p本次执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0,0,1,Null,Null,Null," & lngExcDeptID & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL '
                strSQL = "ZL_病人医嘱执行_Transfusion(" & .执行医嘱ID & "," & .发送号 & "," & _
                         "To_Date('" & Format(p本次执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & lng流水号 & ",'" & p配药人 & "'," & _
                         Me.选择组数 & "," & .组次 & "," & .滴速 & "," & p滴系数 & "," & .液体量 & ",'" & DelInvalidChar(.说明, "'") & "','" & p接单人 & "'," & p耗时 & "," & p提醒 & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                For i = 1 To objGroup.BillsCount
                    Set objBIlls = objGroup.BillsItem(i)
                    For Y = 1 To objBIlls.Count
                        Set objBIll = objBIlls.Item(Y)
                        '检查暂存药品否，如果有，则填写使用记录
                                
                        strSQL = "Select B.NO, B.序号, B.药品id,C.发送数次, D.剂量系数, Sum(Nvl(B.数量* B.入出系数, 0)) As 寄存数量, max(d.门诊可否分零) 门诊可否分零 " & vbNewLine & _
                                "From 药品规格 D, 病人医嘱发送 C, 暂存药品记录 B, 病人医嘱记录 A" & vbNewLine & _
                                "Where C.医嘱id = B.医嘱id And C.发送号 = B.发送号 And D.药品id = B.药品id And" & vbNewLine & _
                                "      B.药品id = A.收费细目id And A.ID = B.医嘱id And A.ID = [1] And B.发送号 = [2]" & vbNewLine & _
                                "Group By B.NO, B.序号, B.药品id, C.发送数次 ,D.剂量系数" & vbNewLine & _
                                "Having Sum(Nvl(B.数量, 0)) >0"

                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查暂存药品", objBIll.医嘱ID, .发送号)
                        If Not rsTmp.EOF Then
                            '本次执行的数量不能大于寄存数量
                            If zlCommFun.NVL(rsTmp!门诊可否分零, 0) = 0 Then
                                cur数量 = .本次数次 / .发送数次 * (Val("" & rsTmp.Fields("发送数次")) / rsTmp.Fields("剂量系数"))
                            Else
                                cur数量 = .本次数次 * Abs(Int(0 - objBIll.单量 / rsTmp!剂量系数))
                            End If
                            
                            If rsTmp.Fields("寄存数量") - cur数量 < 0 Then cur数量 = rsTmp.Fields("寄存数量")
                            
                            strSQL = "Zl_暂存药品记录_Adviceused('" & rsTmp.Fields("NO") & "'," & rsTmp.Fields("序号") & "," & objBIll.医嘱ID & "," & .发送号 & "," & _
                                     rsTmp.Fields("药品ID") & "," & cur数量 & ",'" & p接单人 & "',To_Date('" & Format(p本次执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = strSQL
                        End If
                    Next
                Next
            Else
                MsgBox "医嘱 " & .用法 & " 允许执行 " & .发送数次 & " 次，当前已经执行了 " & .已执行数次 & " 次，不能再执行。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        End With
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "Groups.SelectGroupThingNew")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    SelectGroupThingNew = True

    If lng流水号 <> 0 Then
        If blnPrint Then
            strRpt = "ZL1_BILL_1264_" & intType
            Call ReportOpen(gcnOracle, glngSys, strRpt, frmParent, "接单流水号=" & lng流水号, IIf(blnView, 1, 2))
        End If
        If blnLabel Then
            '打印输液瓶签
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_4", frmParent, "接单流水号=" & lng流水号, IIf(blnView, 1, 2))
        End If
        If blnWristband Then
            '打印输液腕带
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1264_5", frmParent, "接单流水号=" & lng流水号, IIf(blnView, 1, 2))
        End If
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetOtherValue(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long)
'功能：取取 上次执行时间,已执行次数,上次组次，要求时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objGroup As Group
    Dim strData As String
 
    On Error GoTo errH
    
    Set objGroup = Me.Item(lng医嘱ID & "_" & lng发送号)
    If objGroup Is Nothing Then
        MsgBox "未找到医嘱，程序异常！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '新方式处理
    strSQL = "Select 上次执行时间,上次要求时间,上次组次,已执行数次 " & vbNewLine & _
        " From (Select 执行时间 as 上次执行时间, 要求时间 as 上次要求时间, 组次 as 上次组次," & vbNewLine & _
        " Count(本次数次) Over(Partition By 医嘱id, 发送号) as 已执行数次 " & vbNewLine & _
        " From 病人医嘱执行 " & vbNewLine & _
        " Where 医嘱id = [1] And 发送号 = [2] " & vbNewLine & _
        " Order By 流水号 Desc) " & vbNewLine & _
        " Where Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Groups.GetOtherValue1", lng医嘱ID, lng发送号)
    If rsTmp.EOF Then
        objGroup.上次执行时间 = CDate(0)
        objGroup.已执行数次 = 0
        objGroup.上次组次 = 0
    Else
        objGroup.上次执行时间 = zlCommFun.NVL(rsTmp!上次执行时间, CDate(0))
        objGroup.已执行数次 = zlCommFun.NVL(rsTmp!已执行数次, 0)
        objGroup.上次组次 = zlCommFun.NVL(rsTmp!上次组次, 0)
        strData = Format(zlCommFun.NVL(rsTmp!上次要求时间), "yyyy-MM-dd HH:mm:ss")
    End If
    objGroup.要求时间 = GetAdvice要求时间(lng医嘱ID, lng发送号, strData)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RecalculateTime(ByVal dat时间 As Date, ByVal str间隔单位 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal int发送次数 As Integer) As String
'功能：按固定一个开始时间点，重新计算执行时间方案
'参数：
'返回：执行时间方案

    Const DBL_START As Double = 0.333333    '即08：00
    
    Dim i As Integer
    Dim dblBase As Double
    Dim strTimes As String
    
    Select Case Trim(str间隔单位)
        Case "周"
            dblBase = int频率间隔 * 7 / int频率次数
        Case "天"
            dblBase = int频率间隔 / int频率次数
        Case "小时"
            dblBase = int频率间隔 / int频率次数 / 24
    End Select
    
    If Format(dat时间, "yyyy-mm-dd hh:mm:ss") = Format(dat时间, "yyyy-mm-dd 00:00:00") Then
        dat时间 = dat时间 + DBL_START   '如果没有时间，固定按08：00
    End If
    
    For i = 0 To int发送次数 - 1
        strTimes = strTimes & "," & Format(dat时间 + i * dblBase, "yyyy-mm-dd hh:mm:ss")
    Next
    If strTimes <> "" Then strTimes = Mid(strTimes, 2)

    RecalculateTime = strTimes
End Function

Private Function GetAdvice要求时间(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal strCurDate As String) As Date
'功能：根据医嘱ID，发送号和上次执行时间计算本次要求时间
'      医嘱ID，发送号，
'      strCurDate 上次执行时间，为空表示从未执行 固定格式：yyyy-MM-dd HH:mm:ss
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim dat要求时间 As Date
    Dim strPause As String
    Dim strTime  As String
    Dim vDate As Date
    Dim arrTime As Variant
    Dim i As Long
    
    On Error GoTo errH
    
    '计算本次执行应该的要求时间
    strSQL = "Select A.首次时间,A.末次时间," & _
        " B.开始执行时间,B.执行终止时间,B.上次执行时间,B.执行时间方案," & _
        " B.执行频次,B.频率次数,B.频率间隔,B.间隔单位,B.医嘱期效" & _
        " From 病人医嘱发送 A,病人医嘱记录 B" & _
        " Where A.医嘱ID=B.ID And A.医嘱ID=[1] And A.发送号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Groups.GetAdvice要求时间", lng医嘱ID, lng发送号)
        
    If rsTmp!执行频次 & "" = "一次性" Or rsTmp!执行频次 & "" = "需要时" Then
        '为一次性执行的临嘱
        dat要求时间 = rsTmp!开始执行时间
    ElseIf strCurDate = "" Then
        '第一次执行时,就为首次时间
        dat要求时间 = rsTmp!首次时间
    Else
        '根据执行频率分解时间
        If Val(rsTmp!医嘱期效 & "") = 0 Then
            '长嘱才会有暂停时间段
            strPause = GetAdvicePause(lng医嘱ID)
        End If
        If IsNull(rsTmp!执行时间方案) And (NVL(rsTmp!频率次数, 0) = 0 Or NVL(rsTmp!频率间隔, 0) = 0 Or IsNull(rsTmp!间隔单位)) Then
            '持续性长嘱
            strTime = ""
            vDate = Format(rsTmp!首次时间, "yyyy-MM-dd")
            Do While vDate <= Format(rsTmp!末次时间, "yyyy-MM-dd")
                If Not DateIsPause(vDate, strPause) Then
                    If Format(vDate, "yyyy-MM-dd") > Format(strCurDate, "yyyy-MM-dd") And strTime = "" Then
                        strTime = Format(vDate, "yyyy-MM-dd")
                    End If
                End If
                vDate = vDate + 1
            Loop
            If strTime = "" Then strTime = "0"
            dat要求时间 = CDate(strTime)
        Else
            vDate = Calc本周期开始时间(rsTmp!开始执行时间, rsTmp!首次时间, rsTmp!频率间隔, rsTmp!间隔单位)
            strTime = Calc段内分解时间(vDate, rsTmp!末次时间, strPause, rsTmp!执行时间方案, rsTmp!频率次数, rsTmp!频率间隔, rsTmp!间隔单位, rsTmp!开始执行时间)
            arrTime = Split(strTime, ",")
            dat要求时间 = rsTmp!开始执行时间
            For i = 0 To UBound(arrTime)
                If arrTime(i) > strCurDate Then
                    dat要求时间 = CDate(arrTime(i))
                    Exit For '以第一个时间为要求时间
                End If
            Next
            If i > UBound(arrTime) Then
                dat要求时间 = CDate(arrTime(0))
            End If
        End If
    End If
    GetAdvice要求时间 = dat要求时间
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
