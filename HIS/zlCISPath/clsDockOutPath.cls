VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockOutPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean) '要求查看报告
Public Event Activate()                                                     '子窗体激活
Public Event RequestRefresh(ByVal lngPathState As Long)                     '要求主窗体刷新,-1表示未导入状态
Public Event StatusTextUpdate(ByVal Text As String)                         '要求更新主窗体状态栏文字

Private WithEvents mfrmPathTable As frmPathTableOut
Attribute mfrmPathTable.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mfrmPathTable = New frmPathTableOut
End Sub

Private Sub Class_Terminate()
    Unload mfrmPathTable
    Set mfrmPathTable = Nothing
End Sub

Private Sub mfrmPathTable_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(报告ID, CanPrint)
End Sub

Private Sub mfrmPathTable_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmPathTable_RequestRefresh(ByVal lngPathState As Long)
    RaiseEvent RequestRefresh(lngPathState)
End Sub

Private Sub mfrmPathTable_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
    Set zlGetForm = mfrmPathTable
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      CommandBars=仅用于查看时可以不传(传入Nothing)
    Call mfrmPathTable.zlDefCommandBars(frmParent, CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmPathTable.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mfrmPathTable.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmPathTable.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号ID As Long, ByVal str挂号NO As String, ByVal lng科室ID As Long, ByVal int病人状态 As Integer, Optional ByVal blnMoved As Boolean, _
     Optional ByVal blnForceRefresh As Boolean)
'功能：刷新路径表数据
'参数：int病人状态= 病人挂号记录里面的执行状态：0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
'      blnForceRefresh=强制刷新路径表
'      lngState =病人的状态
'      lng医护科室ID=医护站界面ID
    Call mfrmPathTable.zlRefresh(lng病人ID, lng挂号ID, str挂号NO, lng科室ID, int病人状态, blnMoved, blnForceRefresh, 0)
End Sub

Public Sub zlRefreshReadOnly(ByVal lng病人ID As Long, ByVal lng挂号ID As Long)
'功能：外部程序调用,显示只读的路径表单
'注意：如果不是医生站或护士站调用，之前必须先调clsCISPath.InitCISPath
    Dim lng病区ID As Long, lng科室ID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String

'''    strSql = "Select 当前病区id,出院科室id,状态,数据转出 From 病人挂号记录 Where 病人ID = [1] And ID = [2]"
'''    On Error GoTo errH
'''    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "zlRefreshReadOnly", lng病人ID, lng挂号ID)
'''    If rsTmp.RecordCount = 0 Then Exit Sub
'''
'''    Call mfrmPathTable.zlRefresh(lng病人ID, lng挂号ID, Val("" & rsTmp!当前病区id), Val("" & rsTmp!出院科室id), _
'''            Val("" & rsTmp!状态), Val("" & rsTmp!数据转出) = 1, False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function zlExecPathSend(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'功能：执行路径生成(医生站和护士站调用)
'参数：blnIsCancel，没有路径可生成时，用户是否取消了评估。true=取消
'      lngType:1-医嘱编辑界面调用，则评估后不继续生成，因为医嘱编辑界面不能再调用医嘱编辑。
    zlExecPathSend = mfrmPathTable.FuncSendItem(blnIsCancel, lngType)
End Function

Public Function zlExecEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'功能：执行取消评估(医生站)
'参数：blnPrompt=是否弹出询问提示
    zlExecEvaluateCancel = mfrmPathTable.FuncEvaluateCancel(blnPrompt, blnRefresh)
End Function

Public Function zlExePathAppendItem(ByVal strItemType As String, ByVal strAdviceIDs As String, ByVal lng执行ID As Long, ByVal datDate As Date) As Boolean
'功能：添加路径外项目(医生站)
    zlExePathAppendItem = mfrmPathTable.FuncAppendItem(1, strItemType, strAdviceIDs, lng执行ID, datDate)
End Function

Public Function zlImportPath() As Boolean
    zlImportPath = mfrmPathTable.FuncImport
End Function

Public Function zlShowAltAdvice(ByVal frmParent As Object, ByVal lng路径项目ID As Long, Optional ByVal strSelectedIDs As String, Optional ByVal intMode As Integer = 0) As String
'功能：路径生成时，根据当前行的路径项目ID，显示备选医嘱
'参数：strSelectedIds，已经选中的备选医嘱IDs
'返回：选择后的选中的备选医嘱IDs
    zlShowAltAdvice = frmAltAdviceSelect.ShowSelect(frmParent, lng路径项目ID, strSelectedIDs, intMode, 1)
End Function

Public Function zlCheckPathDiagOut(ByVal lng挂号ID As Long, ByVal lng诊断来源 As Long, ByVal lngDiagType As Long, ByVal lngDiag As Long, ByVal lng诊断ID As Long) As Boolean
'功能：检查门诊路径对应的诊断是否允许修改
'参数：lng挂号ID；挂号ID,lngDiagType：诊断类型,lngDiag=疾病ID
'返回值:F-不允许修改;T-允许修改
    zlCheckPathDiagOut = CheckPathDiagOut(lng挂号ID, lng诊断来源, lngDiagType, lngDiag, lng诊断ID)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置路径表清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    Call mfrmPathTable.SetFontSize(bytSize)
End Sub

Public Sub zlAddOutPathItem(ByVal strAdviceIDs As String, ByVal lng病人ID As Long, ByVal lng挂号ID As Long)
'功能：回退医嘱为"停止"状态的医嘱时，根据回退医嘱ID串，生成路径外项目
'参数：strAdviceIDs，回退的医嘱ID串
'      lng病人Id-病人ID
'      lng挂号Id-挂号ID
    Call AddOutPathItem(strAdviceIDs, 2, lng病人ID, lng挂号ID)
End Sub

Public Sub zlFuncPathTableOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal lng病人ID As Long, Optional ByVal lng挂号ID As Long, Optional ByVal strDeviceName As String)
'功能：临床路径表单打印
'参数：bytStyle=1-打印,2-预览,3-输出到Excel,4-输出到PDF
'     blnIsSetup-批量打印不进行打印前设置
'     当bytStyle=4时，需要传入strPDFFile=PDF输出默认路径,包含文件名、后缀
'     如果外部未调用zlRefresh或zlRefreshReadOnly的话，则需要传入病人ID,挂号ID
'     strDeviceName-外部指定打印机时传入打印机名称
    If lng病人ID <> 0 And lng挂号ID <> 0 Then
        Call zlRefreshReadOnly(lng病人ID, lng挂号ID)
    End If
    Call mfrmPathTable.zlPrintOutPut(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub
