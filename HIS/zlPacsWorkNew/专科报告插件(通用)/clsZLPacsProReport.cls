VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZLPacsProReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'************************************************************************************************************
'
'该插件执行原理：
'
'1.主界面初始化时，如果在参数中设置了所需的专科报告，程序将创建该插件，并调用InitPlugin接口
'2.当选择检查或者检查状态改变时将调用该插件的Refresh方法
'3.当主界面选择保存报告按钮时，将调用插件的GetElementString方法获取专科报告数据
'4.当报告保存后，将修改插件的pModified属性为False
'
'当使用要素对专科报告进行保存时，必须按照如下格式组织专科报告内容，格式如下：
'
'[[@]]要素名称[[;]]专科报告内容
'
'要素名称默认为“专科报告”要素，即使用该方法之前，必须在对应的诊疗单据中添加“专科报告”的要素，
'如果需要保存自定义的要素，则需要先在诊疗单据中添加所需的要素，然后在保存专科报告时，使用相同的要素名称进行保存。
'
'另外当使用要素对数据进行保存时，电子病历在解析要素内容时，不能使用回车，
'因此如需使用回车换行符，必须使用自定义报表提取要素内容，进行预览打印。
'
'注意：如果不想通过此原理保存专科报告则可以使用初始化参数时，objParent对象所提供的方法
'      objParent对象的方法介绍部分在InitPlugin方法中说明
'************************************************************************************************************



Private mfrmProReport As Object



Public Sub InitPlugin(cnOracle As Object, ByRef objParent As Object)
'***************************************************************************
'cnOracle:Oracle连接对象
'objParent:父窗体对象
'
'---------------------------------------------------------------------------
'
'gobjParent对象包含方法有SendReport，ClearReport，GetReport
'
'SendReport方法定义为(该方法将把内容发送到对应的报告编辑框中)
'public sub SendReport(byval strDescription as string, byval strResult as string, ByVal strAdvice As String)
'
'
'ClearReport方法定义为(该方法将清除对应编辑框中的报告内容)
'public sub ClearReport(byval blnClearDescription as boolean, byval blnClearResult as boolean, ByVal blnClearAdvice As boolean)
'
'
'GetReport方法定义为(该方法将从对应编辑框中获取报告内容)
'public sub GetReport(byref strDescription as string, byref strResult as string, byref strAdvice As String)
'
'
'注意：以上参数中strDescription表示检查描述，strResult表示检查结果，strAdvice表示诊断建议
'
'
'如果成功，返回空字符，失败则返回错误信息
'***************************************************************************
    
    Set gcnOracle = cnOracle
    Set gobjParent = objParent
    
    Call zl9ComLib.InitCommon(gcnOracle)
End Sub


'记录修改状态
Property Get pModified() As Boolean
    pModified = gModified
End Property

Property Let pModified(value As Boolean)
    gModified = value
End Property

'控件名称
Property Get Name() As Boolean
    Name = mfrmProReport.Name
End Property

Property Let Name(value As Boolean)
    mfrmProReport.Name = value
End Property


'Left
Property Get Left() As Long
    Left = mfrmProReport.Left
End Property

Property Let Left(value As Long)
    mfrmProReport.Left = value
End Property

'Top
Property Get Top() As Long
    Top = mfrmProReport.Top
End Property

Property Let Top(value As Long)
    mfrmProReport.Top = value
End Property


'Width
Property Get Width() As Long
    Width = mfrmProReport.Width
End Property

Property Let Width(value As Long)
    mfrmProReport.Width = value
End Property


'Height
Property Get Height() As Long
    Height = mfrmProReport.Height
End Property

Property Let Height(value As Long)
    mfrmProReport.Height = value
End Property


'刷新专科报告数据
Public Function Refresh(ByVal lngAdviceID As Long, ByVal lngReportID As Long, _
    ByVal blnEditable As Boolean, ByVal blnMoved As Boolean) As String
    glngAdviceId = lngAdviceID
    glngReportId = lngReportID
    gblnEditable = blnEditable
    gblnMoved = blnMoved
    
    Call mfrmProReport.zlRefresh
End Function




'返回专科报告内容
Public Function GetElementString() As String
    GetElementString = ""
    
    If mfrmProReport Is Nothing Then Exit Function
    
    GetElementString = mfrmProReport.GetElementString
End Function


Public Sub Resize(ByVal lngLeft As Long, ByVal lngTop As Long, _
    ByVal lngWidth As Long, ByVal lngHeight As Long)
    If mfrmProReport Is Nothing Then Exit Sub
    
    mfrmProReport.Left = 0
    mfrmProReport.Top = 0
    
    mfrmProReport.Width = lngWidth
    mfrmProReport.Height = lngHeight
End Sub

'专科报告句柄
Public Function hWnd() As Long
    hWnd = 0
    
    If mfrmProReport Is Nothing Then Exit Function
    
    hWnd = mfrmProReport.hWnd
End Function



Private Sub Class_Initialize()
    Set mfrmProReport = New frmReportUS
End Sub

'释放专科报告对象
Private Sub Class_Terminate()
    Unload mfrmProReport
    Set mfrmProReport = Nothing
End Sub
