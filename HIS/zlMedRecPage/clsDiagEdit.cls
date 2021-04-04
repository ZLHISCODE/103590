VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmDiagEdit As frmDiagEdit
Private mclsProperty As clsProperty                   '属性类，用来缓存部分变量，隔离数据
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
Public Event Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String) '首页关闭事件
'首页是否是打开状态，只读属性
Public Property Get IsOpen() As Boolean
    IsOpen = mclsProperty.IsOpen
End Property
'首页是否合并过,只读属性
Public Property Get IsMerged() As Boolean
    IsMerged = mclsProperty.IsMerged
End Property
'出院首要诊断是否发生改变,只读属性
Public Property Get IsDiagChange() As Boolean
    IsDiagChange = mclsProperty.IsDiagChange
End Property
'IsDiagInput:指定诊断的首页诊断是否输入
Public Property Get IsDiagInput() As Boolean
    IsDiagInput = mclsProperty.IsDiagInput
End Property

Public Sub InitDiagEdit(cnMain As ADODB.Connection, ByVal lngSys As Long, Optional ByVal lngModule As Long, Optional ByRef objMip As zl9ComLib.clsMipModule)
'功能：初始化
'参数：objMip=消息对象
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
    Set mclsProperty = New clsProperty
    Call mclsProperty.InitProperty(f诊断选择, lngSys, lngModule, objMip)
End Sub


Public Function ShowDiagEdit(frmParent As Object, ByVal lng标识ID As Long, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, _
                        ByVal int病人来源 As Integer, ByVal lng开单科室ID As Long, ByRef str诊断IDs As String, _
                        ByRef str诊断S As String, ByVal bytSize As Byte, Optional ByVal lng医嘱组ID As Long) As Boolean
'功能：调用诊断编辑界面
'参数：lng病人ID=病人ID
'      lng就诊ID=住院:主页ID,门诊：挂号单ID
'      int病人来源=1-门诊，2-住院
'      lng开单科室ID=病人所在科室，诊断使用
'      lng标识ID =用于区分各个申请单的标识，用于保存相应的诊断
'      str诊断IDs=该申请单相关的诊断ID,多个诊断时诊断ID以逗号分割
'      str诊断S =选择的诊断的字符串，多个诊断描述以逗号分割
'      bytSize=0-9号字体，1-12号字体
'      lng医嘱组ID=申请单对应的组医嘱ID，已经生成医嘱的申请单修改时传入
'返回： ShowDiagEdit= 是确定还是取消
'        str诊断S=返回诊断描述字符串，供申请单使用
    Dim arrTmp As Variant, i As Long
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.Module = IIf(int病人来源 = 2, p住院医生站, p门诊医生站)
    mclsProperty.PatiType = IIf(int病人来源 = 2, PF_住院, PF_门诊)
    If Not mclsProperty.CurrentForm Is mfrmDiagEdit Or mclsProperty.CurrentForm Is Nothing Then
        Set mfrmDiagEdit = New frmDiagEdit
        Set mclsProperty.CurrentForm = mfrmDiagEdit
    End If
    mclsProperty.Privs = "" '自动读取
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.IsComfirmInfect = False
    mclsProperty.病人ID = lng病人ID
    mclsProperty.主页ID = lng就诊ID
    mclsProperty.出院科室ID = lng开单科室ID
    mclsProperty.FontSize = IIf(bytSize = 0, 9, 12)
    mclsProperty.AplyMark = lng标识ID
    mclsProperty.AdviceID = lng医嘱组ID
    mclsProperty.DiagRowIDs = str诊断IDs
    mclsProperty.DiagNames = str诊断S
    mclsProperty.IsOK = mfrmDiagEdit.ShowMe
    Call mclsProperty.InitCacheRecInfo
    str诊断IDs = mclsProperty.DiagRowIDs
    str诊断S = mclsProperty.DiagNames
    '清空原有该申请单关联诊断的缓存，缓存新关联的诊断
    If mclsProperty.IsOK Then
        Call DeleteApplyDiag(lng标识ID)
        If str诊断IDs <> "" Then
            arrTmp = Split(str诊断IDs, ",")
            With mclsProperty.DiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !诊断ID = Val(arrTmp(i))
                    !标识ID = lng标识ID
                    !类型 = 0
                    .Update
                Next
            End With
        End If
    End If
    ShowDiagEdit = mclsProperty.IsOK
End Function

Public Function DeleteApplyDiag(ByVal lng标识ID As Long) As Boolean
'功能：删除申请单时，清除该申请单关联的诊断
    With mclsProperty.DiagConn
        .Filter = "标识ID=" & lng标识ID
        .Sort = "诊断ID"
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
    End With
End Function

Public Function ConfirmInfectiousDiseases(ByVal frmParent As Object, ByVal lng标识ID As Long, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal int病人来源 As Integer, ByVal lng开单科室ID As Long, ByVal rsInfec As ADODB.Recordset, Optional ByRef str诊断IDs As String, _
                        Optional ByRef str诊断S As String, Optional ByVal bytSize As Byte) As Boolean
'功能：确认传染病
'参数：lng病人ID=病人ID
'      lng就诊ID=住院:主页ID,门诊：挂号单ID
'      lng标识ID =用于区分各个申请单的标识，用于保存相应的诊断
'      int病人来源=1-门诊，2-住院
'      lng开单科室ID=病人所在科室，诊断使用
'      bytSize=0-9号字体，1-12号字体
'      str诊断IDs=该申请单相关的诊断ID,多个诊断时诊断ID以逗号分割[暂时未使用，以后扩展使用】
'      str诊断S =选择的诊断的字符串，多个诊断描述以逗号分割[暂时未使用，以后扩展使用】
'返回： ShowDiagEdit= 是确定还是取消
'        str诊断S=返回诊断描述字符串，供申请单使用[暂时未使用，以后扩展使用】
    Dim arrTmp As Variant, i As Long
    If Not gclsMain Is Me Or gclsMain Is Nothing Then
        Set gclsMain = Me
    End If
    If Not gclsPros Is mclsProperty Or gclsPros Is Nothing Then
        Set gclsPros = mclsProperty
    End If
    mclsProperty.Module = IIf(int病人来源 = 2, p住院医生站, p门诊医生站)
    mclsProperty.PatiType = IIf(int病人来源 = 2, PF_住院, PF_门诊)
    mclsProperty.IsComfirmInfect = True
    If Not mclsProperty.CurrentForm Is mfrmDiagEdit Or mclsProperty.CurrentForm Is Nothing Then
        Set mfrmDiagEdit = New frmDiagEdit
        Set mclsProperty.CurrentForm = mfrmDiagEdit
    End If
    mclsProperty.Privs = "" '自动读取
    mclsProperty.RefreshPara
    Set mclsProperty.MainForm = frmParent
    mclsProperty.病人ID = lng病人ID
    mclsProperty.主页ID = lng就诊ID
    mclsProperty.出院科室ID = lng开单科室ID
    mclsProperty.FontSize = IIf(bytSize = 0, 9, 12)
    mclsProperty.AplyMark = lng标识ID
    mclsProperty.DiagRowIDs = str诊断IDs
    mclsProperty.DiagNames = str诊断S
    Do While Not rsInfec.EOF
        With mclsProperty.DiagConn
            .AddNew
            !标识ID = lng标识ID
            !类型 = 1
            !疾病目录ID = Val(rsInfec!疾病id & "")
            !诊断目录ID = Val(rsInfec!诊断ID & "")
            .Update
        End With
        rsInfec.MoveNext
    Loop
    mclsProperty.IsOK = mfrmDiagEdit.ShowMe
    Call mclsProperty.InitCacheRecInfo
    str诊断IDs = mclsProperty.DiagRowIDs
    str诊断S = mclsProperty.DiagNames
    '清空原有该申请单关联诊断的缓存，缓存新关联的诊断
    If mclsProperty.IsOK Then
        Call DeleteApplyDiag(lng标识ID)
        If str诊断IDs <> "" Then
            arrTmp = Split(str诊断IDs, ",")
            With mclsProperty.DiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !诊断ID = Val(arrTmp(i))
                    !标识ID = lng标识ID
                    !类型 = 0
                    .Update
                Next
            End With
        End If
    End If
    ConfirmInfectiousDiseases = mclsProperty.IsOK
End Function

Public Sub Closed(ByVal blnEditCancel As Boolean, ByVal str疾病ID As String, ByVal str诊断ID As String, ByVal strTag As String)
'功能：调用事件
' strTag=附加信息，现在存储门诊病人照片文件的路径，以后扩展时，以|分割
    RaiseEvent Closed(blnEditCancel, str疾病ID, str诊断ID, strTag)
End Sub

Private Sub Class_Terminate()
    If Not mfrmDiagEdit Is Nothing Then
        Unload mfrmDiagEdit
        Set mfrmDiagEdit = Nothing
    End If
    Set mclsProperty = Nothing
End Sub
