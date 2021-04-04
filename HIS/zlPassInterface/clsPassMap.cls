VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsVSCOL"
Option Explicit

'保持属性值的局部变量
Private mVSCOL As clsVSCOL  '局部复制
Private mlngModel As PASS_MODEL '局部复制
Private mvsAdvice As Object  '暂存医嘱对象
Private mobjCmdBar As Object         '菜单对象
Private mbytUseType As Byte              'bytUseType =1 临床路径生成医嘱时调用住院医嘱编辑界面
Private mfrmMain As Object               '主窗体
Private mDiags As clsDiags         '门诊诊断对象 门诊传人
Private mPassPati As clsPatient      '病人信息
Private mint场合 As Integer        ' 调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)

Public Property Set PassPati(ByVal clsPati As clsPatient)
    Set mPassPati = clsPati
End Property

Public Property Get PassPati() As clsPatient
    Set PassPati = mPassPati
End Property

Public Property Set frmMain(ByVal vData As Object)
    Set mfrmMain = vData
End Property

Public Property Get frmMain() As Object
    Set frmMain = mfrmMain
End Property

Public Property Let bytUseType(ByVal bytData As Byte)
    mbytUseType = bytData
End Property

Public Property Get bytUseType() As Byte
     bytUseType = mbytUseType
End Property

Public Property Set Diags(ByVal clsData As clsDiags)
    Set mDiags = clsData
End Property

Public Property Get Diags() As clsDiags
    Set Diags = mDiags
End Property

Public Property Let lngModel(ByVal PMData As PASS_MODEL)
    mlngModel = PMData
End Property

Public Property Get lngModel() As PASS_MODEL
    lngModel = mlngModel
End Property

Public Property Set VSCOL(ByRef clsVSData As clsVSCOL)
    Set mVSCOL = clsVSData
End Property

Public Property Get VSCOL() As clsVSCOL
    Set VSCOL = mVSCOL
End Property

Public Property Set vsAdvice(ByRef vsData As Object)
    Set mvsAdvice = vsData
End Property

Public Property Get vsAdvice() As Object
    Set vsAdvice = mvsAdvice
End Property

Public Property Set objCmdBar(ByRef objData As Object)
    Set mobjCmdBar = objData
End Property

Public Property Get objCmdBar() As Object
    Set objCmdBar = mobjCmdBar
End Property

Public Property Let int场合(ByVal intData As Integer)
    mint场合 = intData
End Property

Public Property Get int场合() As Integer
     int场合 = mint场合
End Property

Public Function GetVSCOL(Optional ByVal intCOLID As Integer = -1, Optional ByVal intCOL相关ID As Integer = -1, Optional ByVal intCOL诊疗类别 As Integer = -1, _
    Optional ByVal intCOL诊疗项目ID As Integer = -1, Optional ByVal intCOL收费细目ID As Integer = -1, Optional ByVal intCOL医嘱内容 As Integer = -1, _
    Optional ByVal intCOL期效 As Integer = -1, Optional ByVal intCOL单量 As Integer = -1, Optional ByVal int单量单位 As Integer = -1, _
    Optional ByVal intCOL用法 As Integer = -1, Optional ByVal intCOL天数 As Integer = -1, Optional ByVal intCOL婴儿 As Integer = -1, _
    Optional ByVal intCOL开嘱时间 As Integer = -1, Optional ByVal intCOL开嘱医生 As Integer = -1, Optional ByVal intCOL开始时间 As Integer = -1, _
    Optional ByVal intCOL开嘱科室ID As Integer = -1, Optional ByVal intCOL终止时间 As Integer = -1, Optional ByVal intCOL频率 As Integer = -1, _
    Optional ByVal intCOL频率次数 As Integer = -1, Optional ByVal intCOL频率间隔 As Integer = -1, Optional ByVal intCOL间隔单位 As Integer = -1, _
    Optional ByVal intCOL警示 As Integer = -1, Optional ByVal intCOL序号 As Integer = -1, Optional ByVal intCOL状态 As Integer = -1, _
    Optional ByVal intCOLEDIT As Integer = -1, Optional ByVal intCOL病人ID As Integer = -1, Optional ByVal intCOL主页ID As Integer = -1, _
    Optional ByVal intCOL选择 As Integer = -1, Optional ByVal intCOL执行性质 As Integer = -1, Optional ByVal intCOL药品名称 As Integer = -1, _
    Optional ByVal intCOL单据 As Integer = -1, Optional ByVal intCOLNO As Integer = -1, Optional ByVal intCOL住院号 As Integer = -1, _
    Optional ByVal intCOL门诊号 As Integer = -1, Optional ByVal intCOL审查结果 As Integer = -1, Optional ByVal intCOL总量 As Integer = -1, _
    Optional ByVal intcol总量单位 As Integer = -1, Optional ByVal intcol医嘱嘱托 As Integer = -1, Optional ByVal intcol用药目的 As Integer = -1, _
    Optional ByVal intCol操作类型 As Integer, Optional ByVal intCol禁忌药品说明 As Integer, Optional ByVal intcol用药理由 As Integer, _
    Optional ByVal intCol标志 As Integer, Optional ByVal intCol处方号 As Integer, Optional ByVal intCol执行分类 As Integer, _
    Optional ByVal intCol执行科室ID As Integer) As clsVSCOL
    
    Dim objVsCOL As clsVSCOL
    
    Set objVsCOL = New clsVSCOL
    
    With objVsCOL
        .intCOLID = intCOLID
        .intCOL相关ID = intCOL相关ID
        .intCOL诊疗类别 = intCOL诊疗类别
        .intCOL诊疗项目ID = intCOL诊疗项目ID
        .intCOL收费细目ID = intCOL收费细目ID
        .intCOL医嘱内容 = intCOL医嘱内容
        .intCOL期效 = intCOL期效
        .intCOL单量 = intCOL单量
        .intCOL单量单位 = int单量单位
        .intCOL用法 = intCOL用法
        
        .intCOL天数 = intCOL天数
        .intCOL婴儿 = intCOL婴儿
        .intCOL开嘱时间 = intCOL开嘱时间
        .intCOL开嘱医生 = intCOL开嘱医生
        .intCOL开始时间 = intCOL开始时间
        .intCOL开嘱科室ID = intCOL开嘱科室ID
        .intCOL终止时间 = intCOL终止时间
        .intCOL频率 = intCOL频率
        .intCOL频率次数 = intCOL频率次数
        .intCOL频率间隔 = intCOL频率间隔
        
        .intCOL间隔单位 = intCOL间隔单位
        .intCOL警示 = intCOL警示
        .intCOL序号 = intCOL序号
        .intCOL状态 = intCOL状态
        .intCOLEDIT = intCOLEDIT
        .intCOL病人ID = intCOL病人ID
        .intCOL主页ID = intCOL主页ID
        .intCOL选择 = intCOL选择       'gbytUseType =1时传人
        
        .intCOL执行性质 = intCOL执行性质   '
        .intCOL药品名称 = intCOL药品名称   '中西成药名称
        '为药品模块提供
        .intCOL单据 = intCOL单据
        .intCOLNO = intCOLNO
        .intCOL住院号 = intCOL住院号
        .intCOL门诊号 = intCOL门诊号
        .intCOL审查结果 = intCOL审查结果
        .intCOL总量 = intCOL总量
        .intcol总量单位 = intcol总量单位
        .intcol医嘱嘱托 = intcol医嘱嘱托
        .intcol用药目的 = intcol用药目的
        .intcol用药理由 = intcol用药理由
        .intCol操作类型 = intCol操作类型
        .intCol禁忌药品说明 = intCol禁忌药品说明
        .intCol标志 = intCol标志
        .intCol处方号 = intCol处方号
        .intCol执行分类 = intCol执行分类
        .intCol执行科室ID = intCol执行科室ID
    End With
    Set GetVSCOL = objVsCOL
End Function

Public Function GetPatient(Optional lng病人ID As Long = 0, Optional ByVal lng主页ID As Long = 0) As clsPatient
    Dim objPati As clsPatient
    Set objPati = New clsPatient
    objPati.lng病人ID = lng病人ID
    objPati.lng主页ID = lng主页ID
    
    Set GetPatient = objPati
    Set objPati = Nothing
End Function

Public Function GetDiags() As clsDiags
    Set GetDiags = New clsDiags
End Function

Private Sub Class_Terminate()
    Set mVSCOL = Nothing
    Set mvsAdvice = Nothing
    Set mobjCmdBar = Nothing
    Set mfrmMain = Nothing
    Set mDiags = Nothing
End Sub
