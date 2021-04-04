VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCard"
Option Explicit

Private mObjCard As clsCardDev  '卡接口对象
Public gcnOracle As ADODB.Connection '调用医保和一卡通接口需要使用.
Private mstrInterFace As String '读卡时选择的接口名称

Private mfrmTimer As frmTimer
Private mlngHwnd As Long

Public Event ShowICCardInfo(ByVal strCardNO As String)

Friend Function GetParent() As Long
    GetParent = mlngHwnd
End Function

Public Sub SetParent(ByVal lngHwnd As Long)
    mlngHwnd = lngHwnd
End Sub

Public Function GetSpare(Optional ByRef strInterFace As String) As Currency
    If Not mObjCard Is Nothing Then
        GetSpare = mObjCard.GetSpare
        strInterFace = mstrInterFace
    End If
End Function

Public Function ReturnSwap(ByVal strCardNO As String, ByVal lngHOSCode As Long, ByVal strTransFlow As String, ByVal dblMoney As Double) As Boolean
'注意,调用本接口之前应先调读卡接口,以确定调用哪一个接口（有多个接口时）
    If Not mObjCard Is Nothing Then
        ReturnSwap = mObjCard.ReturnSwap(strCardNO, lngHOSCode, strTransFlow, dblMoney)
    End If
End Function

Public Function PaymentSwap(ByVal dblMoney As Double, ByRef dblBalance As Double, ByRef intCardType As Integer, _
    ByVal lngHOSCode As Long, ByRef strCardNO As String, ByRef strTransFlow As String, ByVal strBalanceID As String, ByVal strPatiID As String) As Boolean
    If Not mObjCard Is Nothing Then
        PaymentSwap = mObjCard.PaymentSwap(dblMoney, dblBalance, intCardType, lngHOSCode, strCardNO, strTransFlow, strBalanceID, strPatiID)
    End If
End Function

Public Function UploadSwap(ByVal lng病人ID As Long, ByRef strOut As String) As Boolean
    '函数模板，完成与中心数据上传的工作，目前仅应用于宁波一卡通，完成病人信息及卡信息的上传
    On Error GoTo errHand
    
    If mObjCard Is Nothing Then
        Set mObjCard = CreateObject("zlICCard.clsIC_NBYKT")
        If mObjCard Is Nothing Then
            gLastErr = "读卡接口对象未创建！"
            Exit Function
        End If
        Set gcnConnect = gcnOracle
        If Not mObjCard.initCard Then Exit Function
    End If
    UploadSwap = mObjCard.UploadSwap(lng病人ID, strOut)
    Exit Function
errHand:
    gLastErr = err.Description
End Function

Friend Sub ShowICCardInfo(ByVal strCardNO As String)
    RaiseEvent ShowICCardInfo(strCardNO)
End Sub

Public Function Read_Card(Optional ByVal FrmMain As Object) As String
    '功能-读卡功能
    '入参:卡类型
    '出参:卡号,如为空串,表示读卡失败
    Dim blnCancel As Boolean
    Dim strClsCard As String '卡接口类名,
    Dim objCard As clsCard, int险类 As Integer, i As Integer, intReturn As Integer
    Dim objInsure As Object, intType As Integer
    
    On Error GoTo errHandle
    
    Set gcnConnect = gcnOracle
    '显示本已启用的卡
    For Each objCard In Cards
        If objCard.启用 = True Then
            i = i + 1
            int险类 = objCard.险类
            strClsCard = objCard.接口程序名
            mstrInterFace = objCard.名称
        End If
    Next
    If i = 0 Then
        mfrmTimer.tmrMain.Enabled = False
        MsgBox "当前没有启用任何读卡接口！", vbInformation, "IC卡接口"
        Exit Function
    End If
    If i <> 1 Then
        '有多个启用卡,要用户选择
        If Cards.Count > 0 Then
            intType = frmCardSelect.SelectCard(gCol, i, FrmMain)
            
            For Each objCard In Cards
                If objCard.编码 = intType Then
                    int险类 = objCard.险类
                    strClsCard = objCard.接口程序名
                    mstrInterFace = objCard.名称
                    gLastErr = ""
                    Exit For
                Else
                    int险类 = 0
                    strClsCard = ""
                End If
                
            Next
            If strClsCard = "" Then
                gLastErr = "用户取消选择或无可用读卡器！"
            End If
        Else
            gLastErr = "无可用读卡器！"
            GoTo errSow
        End If
    End If
    
    '-----------------------------
    If strClsCard <> "" Then
        If int险类 <> 0 Then
            If Not gcnOracle Is Nothing Then
                If gcnOracle.State = 1 Then
                    Set objInsure = CreateObject("zl9Insure.clsInsure")
                    Call objInsure.InitOracle(gcnOracle)
                    Read_Card = objInsure.readcard(int险类, intReturn)
                    If intReturn <> 1 Then
                        gLastErr = Read_Card
                        Read_Card = ""
                    End If
                Else
                    Read_Card = ""
                    gLastErr = "数据库连接未打开，不能进行医保读卡！"
                End If
            Else
                Read_Card = ""
                gLastErr = "未初始化数据库连接，不能进行医保读卡！"
            End If
        Else
            If i <> 1 Then
                Set mObjCard = CreateObject(strClsCard)
            End If
            If mObjCard Is Nothing Then
                Read_Card = ""
                gLastErr = "读卡接口对象未创建！"
                GoTo errSow
            End If
            If mObjCard.initCard Then
                Read_Card = mObjCard.readcard("", blnCancel)
            End If
            mObjCard.CloseCard
        End If
    End If
    
errSow:
    '如果读卡返回空且取消为假,说明读卡失败
    If gbln自动读取 = False Then
        If Read_Card = "" And blnCancel = False Then
            If gLastErr = "" Then
                gLastErr = "读卡时发生未知错误！"
            End If
            Call WritLog("ICCard.read_card", "", gLastErr)
            MsgBox gLastErr, vbExclamation, "IC卡接口"
        End If
    End If
    Exit Function
errHandle:
    Read_Card = ""
    If Not mObjCard Is Nothing Then mObjCard.CloseCard
    gLastErr = "读卡时发生错误。" & vbNewLine & err.Description
    Call WritLog("ICCard.read_card", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC卡接口"
End Function

Public Function Set_Card(Optional ByVal FrmMain As Object) As Boolean
    '功能-卡类型有参数设置
    Dim objCard As clsCard
    On Error GoTo errHandle
    If Cards.Count > 0 Then
        '前台调用时，显示所有可用的读卡器，供用户选择，然后进行设置。
        Call frmCardSelect.SelectCard(gCol, -1, FrmMain)
        For Each objCard In Cards
            objCard.启用 = GetSetting("ZLSOFT", "公共模块\zlICCard", objCard.编码, 0) = 1
        Next
        gDebug = GetSetting("ZLSOFT", "公共模块\zlICCard", "调试", 0) = 1
        Set_Card = True
    Else
        gLastErr = "无可用的读卡器！"
    End If
    Exit Function
errHandle:
    gLastErr = "设置读卡器发生错误。" & vbNewLine & err.Description
    Call WritLog("ICCard.Set_Card", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC卡接口"
End Function

Public Property Get ErrorMessage() As String
    '检索属性值时使用，位于赋值语句的右边。
    'Syntax: Debug.Print X.ErrorMessage
    ErrorMessage = gLastErr
End Property



Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    If gbln自动读取 = True Then
        If mfrmTimer Is Nothing Then Exit Sub
        mfrmTimer.tmrMain.Enabled = blnEnabled
    End If
End Sub

'----------------------
'以下是内部过程
'----------------------

Private Sub Class_Initialize()
    Dim intCount自动读取 As Integer
    Dim intCount启用 As Integer
    Dim i As Integer
    Dim strClsCard As String
    
    '当创建 clsInterfaceCard 类时，创建 mclsCard 对象
    On Error GoTo errHandle
    Set gCol = New Collection
    Set gcnOracle = Nothing
    Set mObjCard = Nothing
    mstrInterFace = ""
    
    '--------------
    '暂时手工加入一个测试卡，要改为从配置文件中读取。
    
    Call initCards
    
    For i = 1 To Cards.Count
        If Item(i).是否自动读取 = 1 And Item(i).启用 = True Then    '判断当前卡是否启用并且自动读取
            Set mfrmTimer = New frmTimer
            Load mfrmTimer
            mfrmTimer.tmrMain.Interval = Val(GetSetting("ZLSOFT", "公共全局\ICCard", "自动读取间隔" & Item(i).编码, 300))
            intCount自动读取 = intCount自动读取 + 1
            intCount启用 = intCount启用 + 1
            strClsCard = Item(i).接口程序名 '获取接口程序名
        ElseIf Item(i).启用 = True Then
            intCount启用 = intCount启用 + 1
            strClsCard = Item(i).接口程序名 '获取接口程序名
        End If
    Next
    gbln自动读取 = intCount自动读取 = 1 And intCount启用 = 1
    
    If intCount启用 = 1 Then   '自动读卡
        Set mObjCard = CreateObject(strClsCard)
    End If
    If gbln自动读取 = True Then Call mfrmTimer.Init(Me)
    
    gDebug = GetSetting("ZLSOFT", "公共模块\zlICCard", "调试", 0) = 1
    
    '初始化公共部件
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    End If
    
    
    Exit Sub
errHandle:
    gLastErr = "初始化类失败!" & vbNewLine & err.Description
    Call WritLog("ICCard.Initialize", "", gLastErr)
    MsgBox gLastErr, vbExclamation, "IC卡接口"
End Sub

