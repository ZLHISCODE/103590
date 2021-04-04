VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Initialize(cnMain As ADODB.Connection) As Boolean
'功能：初始化社区公共接口部件，不涉及具体的社区接口调用(因为是动态初始化)
'参数：cnMain=数据库连接对象
'返回：如果有启用的社区接口，则返回True，否则表示没有启用任何社区接口
    Dim strSQL As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        
    '读取社区目录缓存(所有的)
    On Error GoTo errH
    strSQL = "Select * From 社区目录 Order by 序号"
    Set grsCommunity = zlDatabase.OpenSQLRecord(strSQL, "Initialize")
    If grsCommunity.EOF Then Exit Function
    grsCommunity.Filter = "启用=1"
    If grsCommunity.EOF Then Exit Function
    
    Initialize = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Terminate() As Boolean
'功能：终止社区接口，在HIS调用主程序退出时调用
'返回：执行成功/失败
    Dim blnReturn As Boolean, i As Integer
    
    blnReturn = True
    
    '将用过的具体社区接口都进行终止
    For i = 1 To gcolCommunity.Count
        blnReturn = blnReturn And gcolCommunity(i).Terminate
    Next
    
    If Not grsCommunity Is Nothing Then
        If grsCommunity.State = 1 Then grsCommunity.Close
        Set grsCommunity = Nothing
    End If
    Set gobjCommunity = Nothing
    Set gcolCommunity = Nothing
    
    Terminate = blnReturn
End Function

Public Function Setup(ByVal int社区 As Integer) As Boolean
'功能：对指定的社区接口参数进行设置
    Set gobjCommunity = GetCommunity(int社区)
    
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        Setup = gobjCommunity.Setup(int社区)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Identify(ByVal lngSys As Long, ByVal lngModual As Long, _
    int社区 As Integer, str社区号 As String, colInfo As Collection, _
    Optional lng病人ID As Long, Optional ByVal lng就诊ID As Long) As Boolean
'功能：社区身份验证，验证失败返回False(根据平台接口对社区病人信息及状态的判断)
'参数：
'     lngSys,lngModual(In)：调用该接口的系统及模块号，如100-标准版,1111-挂号。
'     int社区(Out)：该病人所属的社区编号(如1-成都市青羊区)
'         如果系统中只启用了一个社区，则可以直接取得，如果启用了多个，则弹出窗体人为确定当前病人所属的社区。
'         在具体的社区接口中，该函数参数为传入(In)。
'     str社区号(Out)：该病人在该社区中的唯一编号，如ABC123
'     colInfo(Out)：包含病人各个信息项目内容的集合，Key为"_字段名"，如"_姓名"；失败时返回空集(非Nothing)
'     lng病人ID(In/Out)：可选参数，如果已确定病人或者补充身份验证时传入。可用为Out参数。
'     lng就诊ID(In)：可选参数，补充身份验证时使用；对应"挂号ID"或"主页ID"，可通过调用模块场合确定。
'返回：如果验证成功则返回True及相关Out参数

    If int社区 = 0 Then
        int社区 = frmSelect.ShowMe
        If int社区 = 0 Then Exit Function '取消选择或其他原因
    End If
    
    Set gobjCommunity = GetCommunity(int社区)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        Set colInfo = New Collection
        Identify = gobjCommunity.Identify(lngSys, lngModual, int社区, str社区号, colInfo, lng病人ID, lng就诊ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetCommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer) As String
'功能：返回特定社区在特定模块中的可调用功能，用于调用模块的社区菜单显示
'参数：都是传入参数
'      lngSys,lngModual=系统号，模块号
'      int社区=特定的社区；当医院可对多个社区的病人就诊时，用于区分不同社区的功能
'返回：功能描述串="功能号1,功能名1,功能说明1;功能号2,功能名2,功能说明2;..."
'      功能号为>0的整数，可以不连续；功能说明可为空，但需加上间隔符；功能名和功能说明中不能包含半角的间隔符
'      如果某个功能需要在病人接诊时自动调用(也可手工调用),则在前面加上"Auto:"前缀，如"Auto:1,档案查看,查看社区档案;...",这种功能只允许有一个
'      返回空表示不提供功能
    
    Set gobjCommunity = GetCommunity(int社区)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        GetCommunityFunc = gobjCommunity.GetCommunityFunc(lngSys, lngModual, int社区)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CommunityFunc(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int功能号 As Integer, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：社区功能调用，调用平台接口中提供的社区病人业务功能。
'参数：都是传入参数
'      int功能号：由"GetCommunityFunc"返回的预定的一个编号序列，接口中根据这个编号再调用平台接口的相关功能。
'      int社区、str社区号、lng病人ID、lng就诊ID：当前病人的就诊标识，其中就诊ID对应"挂号ID"或"主页ID"，可通过调用模块场合确定
    
    Set gobjCommunity = GetCommunity(int社区)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        CommunityFunc = gobjCommunity.CommunityFunc(lngSys, lngModual, int功能号, int社区, str社区号, lng病人ID, lng就诊ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function UpdateInfo(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：社区档案同步，HIS中的社区病人信息发生变化时，调用平台接口对中心档案库的数据进行同步。
'参数：都是传入参数
'说明：该接口应在HIS对应功能的同一个数据库事务中调用，如果调用失败，则HIS中回退事务。
    Set gobjCommunity = GetCommunity(int社区)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        UpdateInfo = gobjCommunity.UpdateInfo(lngSys, lngModual, int社区, str社区号, lng病人ID, lng就诊ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ClinicSubmit(ByVal lngSys As Long, ByVal lngModual As Long, ByVal int社区 As Integer, ByVal str社区号 As String, ByVal lng病人ID As Long, ByVal lng就诊ID As Long) As Boolean
'功能：就诊信息提交，社区病人完成本次就诊时(比如医生站完成就诊)，调用接口将本次就诊过程中的业务数据信息上传到中心档案库。
'参数：都是传入参数
'说明：该接口应在HIS对应功能的同一个数据库事务中调用，如果调用失败，则HIS中回退事务。
    Set gobjCommunity = GetCommunity(int社区)
    If Not gobjCommunity Is Nothing Then
        On Error GoTo errH
        ClinicSubmit = gobjCommunity.ClinicSubmit(lngSys, lngModual, int社区, str社区号, lng病人ID, lng就诊ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
