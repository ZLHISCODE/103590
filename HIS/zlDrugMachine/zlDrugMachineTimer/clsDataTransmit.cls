VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTransmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'-----------------------------------------
'说明：定时数据传送类
'作者：YuZhiYong
'邮箱：yzyong@zlsoft.cn
'-----------------------------------------

'自定义事件，回事日志信息到管理工具
Public Event AfterTransmit(ByVal strLog As String)

Private WithEvents mfrmMain As frmMain                  '主窗体对象
Attribute mfrmMain.VB_VarHelpID = -1
Private ParamsInfo As mdlDefine.TYPE_PARAMS             '参数
Private mblnAction As Boolean
Private mblnTransmitting As Boolean
Private mobjComLib As Object
Private mclsLog As clsLog
Private mstrData As String
Private mstrUser As String

Public Sub Transmit(ByVal strData As String)
'功能：基础数据的传送
'参数：
'  strData：发送的数据
'    格式：[接口编号]|[业务类别]|[业务数据]
'    接口编号：接口编号1[;接口编号...]
'    业务类别：
'        1-部门信息；
'        2-人员信息；
'        3-药品目录；
'        4-药品库存与库位；
'        5-发药窗口；
'    业务数据：
'        1-部门：部门性质1;部门性质2;…
'        2-人员：人员性质1;人员性质2;…
'        3-药品：剂型1;剂型2;…
'        4-库存：库房id1;库房id2;…
'        5-窗口：库房id1;库房id2;…

    Dim strTmp As String

    If mfrmMain.timTransmit.Tag = "1" Then
        'Timer正在执行
        strTmp = "定时数据正在传送，基础数据传送将稍后进行"
        mclsLog.Add strTmp
        mclsLog.Add strData, , Val("1-详细日志")
        RaiseEvent AfterTransmit(strTmp)
        mfrmMain.BaseData = strData
    Else
        'Timer未在执行
        mclsLog.Add vbNewLine & "" & Now
        
        strTmp = "开始传送基础数据，定时器暂停"
        mclsLog.Add strTmp
        mclsLog.Add strData, , Val("1-详细日志")
        RaiseEvent AfterTransmit(strTmp)
        
        mblnTransmitting = True
        mfrmMain.timTransmit.Enabled = False    '暂停Timer事件
        
        Call mfrmMain.TransBase(strData)
        
        mfrmMain.timTransmit.Enabled = True     '恢复Timer事件
        mblnTransmitting = False
        
        strTmp = "完成传送基础数据，定时器恢复"
        mclsLog.Add strTmp
        RaiseEvent AfterTransmit(strTmp)
    End If

    mclsLog.Save

End Sub

Public Property Get Transmitting() As Boolean
'数据正在传送
    Transmitting = mblnTransmitting
End Property

Public Property Get SupportData() As String
'定时传送的业务数据列表字符串（只读）
    If mfrmMain Is Nothing Then
        SupportData = ""
    Else
        SupportData = mfrmMain.SupportData
    End If
End Property

Public Property Get TimerAction() As Boolean
'定时传送启停
    TimerAction = mblnAction
End Property
Public Property Let TimerAction(ByVal blnAction As Boolean)
    mblnAction = blnAction
    
    If mblnAction Then
    
        '创建定时窗体
        Set mfrmMain = New frmMain
        If mfrmMain.ShowMe(mstrUser, mobjComLib, mclsLog) = False Then
            mblnAction = False
            GoTo makClose
        End If
        
        Call mfrmMain.ReadParams
        
        '调整定时器
        With mfrmMain.timTransmit
            .Enabled = False
            .Interval = Val("60000/毫秒")
            .Enabled = True
        End With
        mfrmMain.Tag = Timer()  '开始时间
        
    Else
    
makClose:
        If Not mfrmMain Is Nothing Then
            mfrmMain.timTransmit.Enabled = False
        End If
        Set mfrmMain = Nothing
        
    End If
        
End Property

Public Function Init(ByVal strUser As String, ByRef objComLib As Object) As Boolean
'参数：
'  strUser：ZLHIS用户名

    Dim strTmp As String
    
    mstrUser = strUser
    
    If Trim(mstrUser) = "" Then
        strTmp = "初始化定时部件时，未传送ZLHIS用户名"
        RaiseEvent AfterTransmit(strTmp)
        Exit Function
    End If
    
    If objComLib Is Nothing Then
        strTmp = "初始化定时部件时，请确保zl9ComLib部件已创建"
        RaiseEvent AfterTransmit(strTmp)
        Exit Function
    End If
    
    Set mobjComLib = objComLib
End Function

Public Sub ReadParams()
'功能：读取参数，并保存到变量
    
    Dim objXML As New clsXML
    Dim strFile As String, strTmp As String
    
    If LCase(App.Path) Like "*\apply" Then
        strFile = App.Path & "\zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*\apply\*" Then
        strFile = Left(App.Path, InStr(LCase(App.Path), "\apply\") + Len("\apply\") - 1) & "zlDrugMachine.cfg"
    ElseIf LCase(App.Path) Like "*zldrugmachinemanage*" Or LCase(App.Path) Like "*zldrugmachine\*" Or LCase(App.Path) Like "*zldrugmachine" Then
        strFile = Replace(App.Path, App.EXEName, "") & "zlDrugMachineManage\zlDrugMachine.cfg"
    End If
    
    If objXML.OpenXMLFile(strFile) = False Then
        strTmp = "管理工具的参数文件不正确！" & vbNewLine & strFile
        RaiseEvent AfterTransmit(strTmp)
    End If

    With ParamsInfo
        .定时周期 = Val(GetParameter(objXML, "cycle", "5"))
        .有效天数 = Val(GetParameter(objXML, "validdays", "2"))
        .显示最大行数 = Val(GetParameter(objXML, "viewlines", "200"))
        .输出日志 = Val(GetParameter(objXML, "output", "0")) = 1
        .详细日志 = Val(GetParameter(objXML, "detailed", "0")) = 1
        .保存日志天数 = Val(GetParameter(objXML, "savedays", "7"))
        .业务数据 = Trim(GetParameter(objXML, "businessdata", ""))
        
        If .有效天数 > 7 Then .有效天数 = 5
    End With
    
    objXML.CloseXMLDocument
    Set objXML = Nothing
    
    '同时更新窗体的参数值
    If Not mfrmMain Is Nothing Then
        Call mfrmMain.ReadParams
    End If
End Sub

Private Sub Class_Initialize()
    Set mfrmMain = New frmMain

    Call ReadParams
    Set mclsLog = New clsLog
    With mclsLog
        .Path = App.Path
        .FileName = App.EXEName
        .Action = ParamsInfo.输出日志
        .Type_ = IIf(ParamsInfo.详细日志, 1, 0)
        .SaveDays = ParamsInfo.保存日志天数
    End With
End Sub

Private Sub Class_Terminate()
    Dim frmThis As Form
    
    mclsLog.Save
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    Set mfrmMain = Nothing
    Set mclsLog = Nothing
    Set mobjComLib = Nothing
End Sub

Private Sub mfrmMain_AfterTransmit(ByVal strLog As String, ByVal blnTransmitting As Boolean)
    mblnTransmitting = blnTransmitting
    RaiseEvent AfterTransmit(strLog)
End Sub


 
