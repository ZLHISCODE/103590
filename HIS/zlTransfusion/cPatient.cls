VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cPatients"
Option Explicit

'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制

Private mvar病人科室 As String '局部复制
Private mvar门诊号 As String '局部复制
Private mvar床位号 As String '局部复制
Private mvar接受时间 As Date '局部复制
Private mvar座位号 As String '局部复制
Private mvar顺序号 As String '局部复制
Private mvar排队状态 As String '局部复制
Private mvar就诊卡号 As String '局部复制
Private mvar费别 As String '局部复制
Private mvar门诊诊断 As String '局部复制
Private mvar医生 As String '局部复制
Private mvar挂号单 As String '局部复制
Private mvar加权号 As Long
Private mvar呼叫标志 As Integer
Private mvar挂号时间 As Date
Private mvar身份证号  As String
Private mvarIC卡号 As String

Private mvar开始时间 As Date
Private mvar开始操作员 As String
Private mvar结束时间 As Date
Private mvar结束操作员 As String
Private mvar穿刺台 As String
Private mlng单据ID As Long
Private mbyt病人来源 As Byte

Private mvarKey As String

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let 挂号单(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.挂号单 = 5
    mvar挂号单 = vData
End Property

Public Property Get 挂号单() As String
'检索属性值时使用，位于赋值语句的右边。
' X.挂号单
    挂号单 = mvar挂号单
End Property

Public Property Let 单据ID(ByVal lngVar As Long)
'门诊为挂号ID；门诊留观为主页ID
    mlng单据ID = lngVar
End Property
Public Property Get 单据ID() As Long
    单据ID = mlng单据ID
End Property

Public Property Let 病人来源(ByVal bytVar As Byte)
'0-门诊；1-门诊留观
    mbyt病人来源 = bytVar
End Property
Public Property Get 病人来源() As Byte
    病人来源 = mbyt病人来源
End Property

Public Property Get 加权号() As Long
    加权号 = mvar加权号
End Property

Public Property Let 加权号(ByVal vData As Long)
    mvar加权号 = vData
End Property

Public Property Let 医生(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医生 = 5
    mvar医生 = vData
End Property


Public Property Get 医生() As String
'检索属性值时使用，位于赋值语句的右边。
' X.医生
    医生 = mvar医生
End Property



Public Property Let 门诊诊断(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊诊断 = 5
    mvar门诊诊断 = vData
End Property


Public Property Get 门诊诊断() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊诊断
    门诊诊断 = mvar门诊诊断
End Property



Public Property Let 费别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.费别 = 5
    mvar费别 = vData
End Property


Public Property Get 费别() As String
'检索属性值时使用，位于赋值语句的右边。
' X.费别
    费别 = mvar费别
End Property



Public Property Let 就诊卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar就诊卡号 = vData
End Property


Public Property Get 就诊卡号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.就诊卡号
    就诊卡号 = mvar就诊卡号
End Property



Public Property Let 排队状态(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行状态 = 5
    mvar排队状态 = vData
End Property


Public Property Get 排队状态() As String
'检索属性值时使用，位于赋值语句的右边。
' X.执行状态
    排队状态 = mvar排队状态
End Property

Public Property Let 顺序号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.顺序号 = 5
    mvar顺序号 = vData
End Property


Public Property Get 顺序号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.顺序号
    顺序号 = mvar顺序号
End Property



Public Property Let 座位号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.座位号 = 5
    mvar座位号 = vData
End Property


Public Property Get 座位号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.座位号
    座位号 = mvar座位号
End Property

Public Property Let 接受时间(ByVal vData As Date)
    mvar接受时间 = vData
End Property


Public Property Get 接受时间() As Date
    接受时间 = mvar接受时间
End Property

Public Property Let 挂号时间(ByVal vData As Date)
    mvar挂号时间 = vData
End Property


Public Property Get 挂号时间() As Date
    挂号时间 = mvar挂号时间
End Property


Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mvar门诊号 = vData
End Property


Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊号
    门诊号 = mvar门诊号
End Property


Public Property Let 病人科室(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人科室 = 5
    mvar病人科室 = vData
    Exit Property
End Property


Public Property Get 病人科室() As String
'检索属性值时使用，位于赋值语句的右边。
' X.科室
    病人科室 = mvar病人科室
    Exit Property
End Property

Public Property Let 年龄(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
    Exit Property
End Property


Public Property Get 年龄() As String

'检索属性值时使用，位于赋值语句的右边。
' X.年龄
    年龄 = mvar年龄
    Exit Property
End Property



Public Property Let 性别(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
    Exit Property
End Property


Public Property Get 性别() As String

'检索属性值时使用，位于赋值语句的右边。
' X.性别
    性别 = mvar性别
    Exit Property
End Property



Public Property Let 姓名(ByVal vData As String)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
    Exit Property
End Property


Public Property Get 姓名() As String

'检索属性值时使用，位于赋值语句的右边。
' X.姓名
    姓名 = mvar姓名
    Exit Property
End Property



Public Property Let 病人ID(ByVal vData As Long)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
    Exit Property
End Property


Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.病人ID
    病人ID = mvar病人ID
    Exit Property
End Property

Public Function Get顺序号() As String
    '取当天最大号+1
    Dim str日期 As String, lng序号 As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "Select To_Char(Sysdate, 'yymmdd') As 日期, Max(顺序号) 顺序号" & vbNewLine & _
            "From 排队记录" & vbNewLine & _
            "Where 日期 Between To_Date(To_Char(Sysdate, 'yyyy-mm-dd'), 'yyyy-mm-dd') And" & vbNewLine & _
            "           To_Date(To_Char(Sysdate + 1, 'yyyy-mm-dd'), 'yyyy-mm-dd')"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "cPatients.Get顺序号")
    If rsTmp.RecordCount > 0 Then
        str日期 = rsTmp.Fields("日期")
        lng序号 = Val(zlCommFun.NVL(rsTmp.Fields("顺序号"), 0)) + 1
    Else
        str日期 = Format(zlDatabase.Currentdate, "yyMMdd")
        lng序号 = 1
    End If
    Get顺序号 = str日期 & "_" & Format(lng序号, "000")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub AddQueue(ByVal lng科室ID As Long)
    Dim strSQL As String, lng顺序号 As Long
    On Error GoTo errHandle
    顺序号 = Get顺序号
    lng顺序号 = Val(Split(顺序号, "_")(1))
    strSQL = "ZL_排队记录_AddQueue(" & mvar病人ID & "," & lng科室ID & ",'" & CStr(lng顺序号) & "','" & Me.挂号单 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.AddQueue")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function UpdateState(ByVal strState As String, ByVal lng科室ID As Long, Optional ByVal blnShowMsg As Boolean = True) As Boolean
    '修改状态
    '0-未到 1-待配液 2-弃号 3-退号 4-结束
    
    '5-待穿刺 6-待执行 7-执行中
    
    '此处不能从 0-未到,改为 1-待配液等,这是由AddQueue过程完成
    Dim intState As Integer, lng顺序号 As Long, strSQL As String
    Dim strInfo As String, blnCurDay As Boolean '是否当天
    Dim blnYes As Boolean
    On Error GoTo errHandle
    
    If Me.排队状态 = strState Then Exit Function    '状态相同，不用修改
    
    intState = Val(Mid(strState, 1, 1))
    blnCurDay = Format(Me.接受时间, "yyyy-MM-dd") = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Select Case intState
        
    Case 1  '1-待配液
        blnYes = False
        'If (Val(Mid(Me.排队状态, 1, 1)) = 3 Or Val(Mid(Me.排队状态, 1, 1)) = 1) And blnCurDay Then
        If Val(Mid(Me.排队状态, 1, 1)) = 1 And blnCurDay Then
            '1-待配液 不能直接改成 1－待配液
            If blnShowMsg Then MsgBox "不能由[" & Me.排队状态 & "]更改为[1-待配液]状态！", vbQuestion, "修改排队状态"
            Exit Function
        Else
            '0,2,3,4,5,6,7 可改为待配液,不是当天的话，什么状态都可以改为待配液
            
            If Val(Me.排队状态) <> 0 Then
                If blnShowMsg Then
                    If MsgBox("准备将[" & Me.姓名 & "]的排队状态从[" & Me.排队状态 & "]更改为[1-待配液]状态，请确认。", vbInformation + vbYesNo + vbDefaultButton2, "修改排队状态") = vbYes Then
                        blnYes = True
                    End If
                Else
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes Then
                lng顺序号 = Val(Split(Me.顺序号, "_")(1))
                If Me.病人来源 = 1 Then
                    '门诊留观
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & ",1,Null," & Me.单据ID & ")"
                Else
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & ",1,'" & Me.挂号单 & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.排队状态 = "1-待配液"
                UpdateState = True
                Exit Function
            End If
        End If
    Case 2, 3, 4 '2-弃号
        blnYes = False
        If Val(Mid(Me.排队状态, 1, 1)) = 0 Or (Val(Mid(Me.排队状态, 1, 1)) >= 2 And Val(Mid(Me.排队状态, 1, 1)) <= 4) Then
            
            If blnShowMsg Then MsgBox "不能由[" & Me.排队状态 & "]更改为[" & strState & "]状态！", vbQuestion, "修改排队状态"
            Exit Function
        Else
            '1,5,6,7 均可弃号,退号,结束
            If blnShowMsg Then
                If MsgBox("准备将[" & Me.姓名 & "]的排队状态从[" & Me.排队状态 & "]更改为[" & strState & "]状态，请确认。", vbInformation + vbYesNo + vbDefaultButton2, "修改排队状态") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes = True Then
                lng顺序号 = Val(Split(Me.顺序号, "_")(1))
                If Me.病人来源 = 1 Then
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",Null," & Me.单据ID & ")"
                Else
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",'" & Me.挂号单 & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                
                'strSQL = "Zl_排队记录_Startend(2," & lng科室ID & ",'" & Me.挂号单 & "',sysdate,'" & UserInfo.姓名 & "')"
                If Me.病人来源 = 1 Then
                    strSQL = "Zl_排队记录_Startend(2," & lng科室ID & "," & Me.病人ID & ",Null," & Me.单据ID & ",'" & UserInfo.姓名 & "')"
                Else
                    strSQL = "Zl_排队记录_Startend(2," & lng科室ID & "," & Me.病人ID & ",'" & Me.挂号单 & "',Null,'" & UserInfo.姓名 & "')"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "填写结束时间")

                Me.排队状态 = strState
                UpdateState = True
                Exit Function
            End If
        End If

    Case 5, 6  '5-待穿刺(要叫号的)，6－不叫号的
        '只能由1－待配液转为 5,6
        blnYes = False
        
        If Val(Mid(Me.排队状态, 1, 1)) = 1 Or Val(Mid(Me.排队状态, 1, 1)) = 7 Then
        
            If blnShowMsg Then
                If MsgBox("准备将[" & Me.姓名 & "]的排队状态从[" & Me.排队状态 & "]更改为[" & strState & "]状态，请确认。", vbInformation + vbYesNo + vbDefaultButton2, "修改排队状态") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            If blnYes = True Then
                lng顺序号 = Val(Split(Me.顺序号, "_")(1))
                If Me.病人来源 = 1 Then
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",Null," & Me.单据ID & ")"
                Else
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",'" & Me.挂号单 & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.排队状态 = strState
                UpdateState = True
                Exit Function
            End If
        Else
            If blnShowMsg Then MsgBox "不能由[" & Me.排队状态 & "]更改为[" & strState & "]状态！", vbQuestion, "修改排队状态"
            Exit Function
        End If
    Case 7  '7-执行中
        blnYes = False
        
        If InStr("0,1,2,3,4,5,6,", Val(Mid(Me.排队状态, 1, 1))) > 0 Then
            
            If blnShowMsg Then
                If MsgBox("准备将[" & Me.姓名 & "]的排队状态从[" & Me.排队状态 & "]更改为[" & strState & "]状态，请确认。", vbInformation + vbYesNo + vbDefaultButton2, "修改排队状态") = vbYes Then
                    blnYes = True
                End If
            Else
                blnYes = True
            End If
            
            If blnYes = True Then
                lng顺序号 = Val(Split(Me.顺序号, "_")(1))
                If Me.病人来源 = 1 Then
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",Null," & Me.单据ID & ")"
                Else
                    strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",'" & Me.挂号单 & "',Null)"
                End If
                Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSrate")
                Me.排队状态 = strState
                UpdateState = True
                Exit Function
            End If
        Else
            If blnShowMsg Then MsgBox "不能由[" & Me.排队状态 & "]更改为[" & strState & "]状态！", vbQuestion, "修改排队状态"
            Exit Function
        End If
        
    Case Else
        If blnShowMsg Then MsgBox "非法的状态值。", vbQuestion, "修改排队状态", gstrSysName
        Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function UpdateSequence(ByVal lng科室ID As Long) As Boolean
    Dim intState As Integer, lng顺序号 As Long, strSQL As String
    On Error GoTo errHandle
    
    intState = Val(Mid(Me.排队状态, 1, 1))
    If intState = 1 Then
        lng顺序号 = Val(Split(Me.顺序号, "_")(1))
        If Me.病人来源 = 1 Then
            strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",Null," & Me.单据ID & ")"
        Else
            strSQL = "ZL_排队记录_Update(" & Me.病人ID & "," & lng科室ID & "," & lng顺序号 & "," & Me.加权号 & "," & intState & ",'" & Me.挂号单 & "', Null)"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.UpdateSequence")
        UpdateSequence = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Public Property Let 呼叫标志(ByVal vData As Integer)

'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar呼叫标志 = vData
    Exit Property
End Property

Public Property Get 呼叫标志() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.病人ID
    呼叫标志 = mvar呼叫标志
    Exit Property
End Property

Public Property Let 身份证号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mvar身份证号 = vData
End Property


Public Property Get 身份证号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊号
    身份证号 = mvar身份证号
End Property

Public Property Let IC卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊号 = 5
    mvarIC卡号 = vData
End Property


Public Property Get IC卡号() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊号
    IC卡号 = mvarIC卡号
End Property

Public Property Let 开始操作员(ByVal vData As String)
    mvar开始操作员 = vData
End Property
Public Property Get 开始操作员() As String
    开始操作员 = mvar开始操作员
End Property

Public Property Let 结束操作员(ByVal vData As String)
    mvar结束操作员 = vData
End Property
Public Property Get 结束操作员() As String
    结束操作员 = mvar结束操作员
End Property

Public Property Let 穿刺台(ByVal vData As String)
    mvar穿刺台 = vData
End Property
Public Property Get 穿刺台() As String
    穿刺台 = mvar穿刺台
End Property

Public Property Let 开始时间(ByVal vData As Date)
    mvar开始时间 = vData
End Property
Public Property Get 开始时间() As Date
    开始时间 = mvar开始时间
End Property

Public Property Let 结束时间(ByVal vData As Date)
    mvar结束时间 = vData
End Property
Public Property Get 结束时间() As Date
    结束时间 = mvar结束时间
End Property

Public Function SetCallTag(ByVal lng科室ID As Long) As Boolean
    Dim strSQL As String
    On Error GoTo errHandle
    
    If Me.病人来源 = 1 Then
        strSQL = "Zl_排队记录_Setcalltag(" & Me.病人ID & "," & lng科室ID & ",Null, " & Me.单据ID & ")"
    Else
        strSQL = "Zl_排队记录_Setcalltag(" & Me.病人ID & "," & lng科室ID & ",'" & Me.挂号单 & "', Null)"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "cPatient.SetCallTag")
    
    SetCallTag = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


