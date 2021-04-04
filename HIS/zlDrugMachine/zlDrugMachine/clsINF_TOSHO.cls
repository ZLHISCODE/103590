VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINF_TOSHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
'说明：第三方接口（TOSHO）类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'------------------------------------------------------------------------------

Implements clsINF

Private mobjBase As New clsINF
Private mcnThird As ADODB.Connection

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Set mobjBase = Nothing
    Set mcnThird = Nothing
End Sub

Private Property Let clsINF_Action(ByVal RHS As Boolean)
'说明：开启接口为活动状态
    mobjBase.Action = RHS
    
    On Error GoTo hErr
    
    If RHS Then
        If mcnThird Is Nothing Then
            mobjBase.Action = False
            mobjBase.mobjLog.Add "连接TOSHO服务器失败", 1
            Exit Property
        End If
        
        mobjBase.mobjLog.Add "开始连接TOSHO服务器", 1
        mcnThird.Open mobjBase.URL
        mobjBase.mobjLog.Add "成功连接TOSHO服务器", 1
    Else
        Set mcnThird = Nothing
    End If
    
    mobjBase.mobjLog.Save
    Exit Property
    
hErr:
    mobjBase.Action = False
    mobjBase.mobjLog.Add Err.Number & "：" & Err.Description, 1
    mobjBase.mobjLog.Add "连接TOSHO服务器异常", 1
    mobjBase.mobjLog.Save
End Property

Private Property Get clsINF_Action() As Boolean
    clsINF_Action = mobjBase.Action
End Property

Private Function clsINF_FilterData(ByVal strCode As String, ByVal intType As Integer, ByVal rsData As ADODB.Recordset, rsFinish As ADODB.Recordset) As Boolean
    clsINF_FilterData = mobjBase.FilterData(strCode, intType, rsData, rsFinish)
End Function

Private Function clsINF_GetData(ByVal intType As Integer, ByVal strData As String, blnFinish As Boolean) As ADODB.Recordset
    Set clsINF_GetData = mobjBase.GetData(intType, strData, blnFinish)
    
    '如果不使用基类的记录集对象，请将前面代码屏蔽，再以下面编写新的代码
    'TODO：
End Function

Private Property Let clsINF_KEY(ByVal RHS As String)
    mobjBase.KEY = RHS
End Property

Private Property Get clsINF_KEY() As String
    clsINF_KEY = mobjBase.KEY
End Property

Private Property Set clsINF_mobjComLib(ByVal RHS As Object)
    Set mobjBase.mobjComLib = RHS
End Property

Private Property Get clsINF_mobjComLib() As Object
    Set clsINF_mobjComLib = mobjBase.mobjComLib
End Property

Private Property Set clsINF_mobjLog(ByVal RHS As clsLog)
    Set mobjBase.mobjLog = RHS
End Property

Private Property Get clsINF_mobjLog() As clsLog
    Set clsINF_mobjLog = mobjBase.mobjLog
End Property

Private Property Let clsINF_mstrNOs(ByVal RHS As String)
    mobjBase.mstrNOs = RHS
End Property

Private Property Get clsINF_mstrNOs() As String
    clsINF_mstrNOs = mobjBase.mstrNOs
End Property

Private Function clsINF_SendData(ByVal intAppType As Integer, ByVal intType As Integer, ByVal intIO As Integer, _
    ByVal strCode As String, ByVal rsData As ADODB.Recordset, strMessage As String, _
    Optional ByRef strWinInfo As String) As Boolean
'功能、参数、返回、说明，请参见clsINF基类的SendData方法

    Dim strSQL As String
    Dim colProc As New Collection

    On Error GoTo hErr
    
    If rsData.State <> adStateOpen Then Exit Function
    
    mobjBase.mobjLog.Add "开始传送数据", 1
    
    Select Case intType
    Case 1 To 3
        rsData.Filter = "停用=0"
    Case 4 To 5
        rsData.Filter = ""
    Case 21
        rsData.Filter = "是否传送=1"
    Case Else
        GoTo makExit
    End Select
    
    If rsData.RecordCount <= 0 Then
makExit:
        mobjBase.mobjLog.Add "无数据可传送", 1
        mobjBase.mobjLog.Save
        clsINF_SendData = True
        Exit Function
    End If
    
    With rsData
        Do While .EOF = False
            '该存储过程的名称固定
            Select Case intType
            Case Val("3-药品目录")
                strSQL = "proHISDrug("
                strSQL = strSQL & "'" & mdlDrugMachine.NVL(!药品id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!通用名) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!规格) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!药品剂型) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!剂量系数)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!剂量单位) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!住院包装)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!住院单位) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!生产商编码) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!生产商) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!通用名简码) & "'"
                strSQL = strSQL & ")"
                
            Case Val("21-配药")
                strSQL = "proHISDosage("
                strSQL = strSQL & "'" & mdlDrugMachine.NVL(!病区id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!病区) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!库房id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!住院号) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!病人id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!病人姓名) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!开嘱医生) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!床号) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!用法) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!药品id) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!通用名) & "'"
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!药品规格) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!剂量系数)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!单量单位) & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!住院单位数量)
                strSQL = strSQL & ",'" & mdlDrugMachine.NVL(!执行时间) & "'"
                strSQL = strSQL & ",'" & IIf(Val(mdlDrugMachine.NVL(!优先级)) = 1, "1", "0") & "'"
                strSQL = strSQL & "," & mdlDrugMachine.NVL(!医嘱id)
                strSQL = strSQL & ")"
                
            Case Else
                strSQL = ""
            End Select
            
            If strSQL <> "" Then
                mdlDrugMachine.AddArray colProc, strSQL
            End If
            
            .MoveNext
        Loop
    End With
    
    '传送数据至存储过程
    On Error GoTo makFail
    Call mdlDrugMachine.ExecuteProcedureBeach(colProc, "", mcnThird, mobjBase.mobjLog)
    On Error GoTo hErr
    
    '标记数据传送成功
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, True)
    
    mobjBase.mobjLog.Add "完成传送数据", 1
    mobjBase.mobjLog.Save
    
    '执行基类的代码
    clsINF_SendData = mobjBase.SendData(intAppType, intType, intIO, strCode, rsData, strMessage)
    Exit Function
    
makFail:
    '标记数据传送失败
    Call mdlDrugMachine.TransmitFlag(intAppType, intType, intIO, rsData, mobjBase, False)
    
hErr:
    mobjBase.mobjLog.Add Err.Number & ":" & Err.Description, 1
    mobjBase.mobjLog.Add "传送数据失败", 1
    mobjBase.mobjLog.Save
End Function

Private Property Let clsINF_Token(ByVal RHS As String)
    '
End Property

Private Property Get clsINF_Token() As String
    '
End Property

Private Property Let clsINF_URL(ByVal RHS As String)
    mobjBase.URL = RHS
    
    Set mcnThird = New ADODB.Connection
End Property

Private Property Get clsINF_URL() As String
    clsINF_URL = mobjBase.URL
End Property

Private Property Let clsINF_UserName(ByVal RHS As String)
    mobjBase.UserName = RHS
End Property

Private Property Get clsINF_UserName() As String
    clsINF_UserName = mobjBase.UserName
End Property
