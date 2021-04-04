VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************************************************************************
'接口说明:
'1.收费接口
'1.1 zlDeviceSetup：配置执行端扣费的支付方式
'1.2 zlSquareAffirm:调用执行端付费

'*********************************************************************************************************************************************
'内部方法及过程说明
'1.GetCardSquareObject:创建结算对象
'*********************************************************************************************************************************************
Private mobjSquare As Object

Private Function GetCardSquareObject(ByRef objCardSquare_Out As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建卡对象
    '入参:
    '出参:objCardSquare_Out-返回对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-06-27 14:10:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjSquare Is Nothing Then
        Set objCardSquare_Out = mobjSquare: GetCardSquareObject = True
        Exit Function
    End If
    
    err = 0: On Error Resume Next
    Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    If err <> 0 Then
        err = 0: On Error GoTo 0
         MsgBox "创建zl9CardSquare对象失败，请检查该部件是否存在或注册", vbInformation + vbOKOnly, "提示"
        Exit Function
    End If
    
    If mobjSquare.zlInitComponents(Nothing, glngModule, glngSys, gstrDBUser, gcnOracle, False) = False Then
         MsgBox "初始化失败,请检查相关环境变量是正确或与HIS厂商联系", vbInformation + vbOKOnly, "提示"
         Set mobjSquare = Nothing
         Exit Function
    End If
    
    Set objCardSquare_Out = mobjSquare: GetCardSquareObject = True
End Function

Public Function zlSquareAffirm(ByVal lng病人ID As Long, ByVal str医嘱信息 As String, ByVal strFeeNos As String) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 调用执行端付费
    ' 入参 : str医嘱信息-格式“医嘱IDs|执行部门ID”根据执行部门对应的医嘱提取划价单
    '        strFeeNos-本次付费的划价单，如果医嘱也传递了划价单，则将两者合并
    ' 出参 :
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/6/27 09:46
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, str医嘱IDs As String
    Dim lng执行部门ID As Long
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    If str医嘱信息 = "" And strFeeNos = "" Then Exit Function
    If GetCardSquareObject(mobjSquare) = False Then Exit Function
    
    str医嘱IDs = Split(str医嘱信息 & "|", "|")(0)
    lng执行部门ID = Val(Split(str医嘱信息 & "|", "|")(1))
    strNos = strFeeNos
    
    If str医嘱信息 <> "" Then
        If InStr(str医嘱信息, ",") > 0 Then '有多个医嘱ID的情况
            strSQL = "Select Distinct NO From (" & vbNewLine & _
                    "With 医嘱 As (Select /*+cardinality(B,10) */" & vbNewLine & _
                    "                 Column_Value" & vbNewLine & _
                    "                From Table(f_Num2list([2])) b)" & vbNewLine & _
                    "Select a.No" & vbNewLine & _
                    "From 门诊费用记录 a, 医嘱 c" & vbNewLine & _
                    "Where a.医嘱序号 = c.Column_Value  And a.记录性质 = 1 And (a.记录状态 = 0 Or a.记录状态 = 1 And a.结帐id Is Null) And" & vbNewLine & _
                    "      a.病人id = [1] And a.执行部门id = [3] And Nvl(a.费用状态, 0) = 0" & vbNewLine & _
                    "Union all" & vbNewLine & _
                    "Select a.No" & vbNewLine & _
                    "From 门诊费用记录 a, 病人医嘱记录 b, 医嘱 c" & vbNewLine & _
                    "Where a.医嘱序号 = b.ID And b.相关id = c.Column_Value And a.记录性质 = 1 And (a.记录状态 = 0 Or a.记录状态 = 1 And a.结帐id Is Null) And" & vbNewLine & _
                    "      a.病人id = [1] And a.执行部门id = [3] And Nvl(a.费用状态, 0) = 0)"
        Else
            strSQL = "Select Distinct No" & vbNewLine & _
                    "From 门诊费用记录" & vbNewLine & _
                    "Where 医嘱序号 In (Select Id From 病人医嘱记录 Where Id = [2] Union All " & vbNewLine & _
                    "                   Select Id From 病人医嘱记录 Where 相关id = [2]) And 记录性质 = 1 And (记录状态 = 0 Or 记录状态 = 1 And 结帐id Is Null) And" & vbNewLine & _
                    "      病人id = [1] And 执行部门id = [3] And Nvl(费用状态, 0) = 0"
        End If

        Set rsData = gzlComLib.zlDatabase.OpenSQLRecord(strSQL, "zlSoftCISInterface", lng病人ID, str医嘱IDs, lng执行部门ID)
        Do While Not rsData.EOF
            If InStr("," & strNos & ",", "," & Nvl(rsData!NO) & ",") = 0 Then
                strNos = strNos & "," & Nvl(rsData!NO)
            End If
            rsData.MoveNext
        Loop
        If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    End If
    
    If strNos = "" Then Exit Function
    
    If mobjSquare.zlSquareAffirm(Nothing, glngModule, "", lng病人ID, , , 1, strNos) = False Then
        Exit Function
    End If
    zlSquareAffirm = True
    Exit Function
errH:
    MsgBox err.Description
End Function

Public Sub zlDeviceSetup()
    '---------------------------------------------------------------------------------------
    ' 功能 : 配置执行端扣费的支付方式
    ' 编制 : 李南春
    ' 日期 : 2019/6/27 09:45
    '---------------------------------------------------------------------------------------
    On Error GoTo errH
    Call gzlComLib.zlCommFun.DeviceSetup(Nothing, glngSys, glngModule)
    Exit Sub
errH:
    MsgBox err.Description
End Sub

Private Sub Class_Terminate()
    If Not mobjSquare Is Nothing Then Set mobjSquare = Nothing
End Sub
