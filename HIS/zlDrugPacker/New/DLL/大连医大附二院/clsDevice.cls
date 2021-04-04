VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'设备基础信息
Private mlngID As Long                      '设备ID
Private mstrNO As String                    '编码
Private mstrName As String                  '名称
Private mstrModel As String                 '型号
Private mstrManufacturer As String          '制造商
Private mlngDeptID As Long                  '使用部门ID
Private mstrDeptName As String              '使用部门名称
Private mbytLinkType As Byte                '连接类型
Private mstrLinkDescribe As String           '连接内容
Private mblnStart As Boolean                '设备启用状态

'设备参数信息
Private mbytServiceObject As Byte           '服务对象
Private mbytDispenseFunc As Byte            '预配药响应
Private mbytSendFunc As Byte                '发送响应
Private mstrBill As String                  '单据类型
Private mstrDrugForm As String              '剂型

'设备数据交换/连接信息
Private mblnStatus As Boolean               '连接状态
Private mcnDB As New ADODB.Connection       '数据库连接
Private mcnWS As Object                     'WebService连接
Private mstrConnect As String
Private mstrURL As String, mstrUser As String, mstrPassword As String

'设备对象ID，与设备ID相同
Public Property Get ID() As Long
    ID = mlngID
End Property
Public Property Let ID(ByVal lngID As Long)
    Dim rsData As ADODB.Recordset
    Dim strSQL As String, strKey As String ', strValue As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    mlngID = lngID
    
    On Error GoTo errHandle
    strSQL = "Select a.Id, a.编码, a.名称, a.型号, a.制造商, a.使用部门ID, a.连接类型, a.连接内容, a.是否启用, a.服务对象, b.参数id, b.参数值 " & _
             "From 药房发药设备 A, 药房设备参数 B " & _
             "Where a.Id = b.设备id(+) And a.Id = [1] And b.参数id(+) = 1 "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(strSQL, "获取设备连接信息", lngID)
    If rsData.EOF = False Then
        LinkType = gobjComLib.zlcommfun.NVL(rsData!连接类型, 1) - 1
        LinkDescribe = gobjComLib.zlcommfun.NVL(rsData!连接内容)
        Name = rsData!名称
        DeptID = rsData!使用部门id
        Start = gobjComLib.zlcommfun.NVL(rsData!是否启用, 0) = 1
        ServiceObject = gobjComLib.zlcommfun.NVL(rsData!服务对象, 0)
        DrugForm = gobjComLib.zlcommfun.NVL(rsData!参数值)
    Else
        LinkType = 0
        LinkDescribe = ""
        Name = ""
        DeptID = 0
        Start = False
        ServiceObject = 0
        DrugForm = ""
    End If
    rsData.Close
    
    '建立连接
    mblnStatus = False
    Select Case LinkType
    Case enuLinkType.DB
        If LinkDescribe <> "" Then
            'DBConnect = New ADODB.Connection
            With DBConnect
                .ConnectionString = LinkDescribe
                .Open
            End With
            mblnStatus = True
        End If
    Case enuLinkType.WEBServices
        If LinkDescribe <> "" Then
            arrTmp = Split(LinkDescribe, GSTR_SEPARAT_CHILD)
            For i = LBound(arrTmp) To UBound(arrTmp)
                strKey = UCase(Left(arrTmp(i), InStr(arrTmp(i), "=") - 1))
                'strValue = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                Select Case strKey
                Case "URL"
                    '地址
                    mstrURL = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                Case "USER"
                    '用户
                    mstrUser = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                Case "PWD", "PASSWORD"
                    '口令
                    mstrPassword = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
                End Select
            Next
            CreateWebServices mstrURL, mcnWS
            
            mblnStatus = Not mcnWS Is Nothing
        End If
    Case enuLinkType.Directory
        If LinkDescribe <> "" Then
            mblnStatus = Dir(LinkDescribe, vbDirectory) <> ""
        End If
    End Select
    
    Exit Property
    
errHandle:
    gstrMessage = Err.Description
    If gobjComLib.ErrCenter = 1 Then Resume
End Property

'设备对象编码
Public Property Get NO() As String
    NO = mstrNO
End Property
Public Property Let NO(ByVal strNO As String)
    mstrNO = strNO
End Property

'设备对象名称
Public Property Get Name() As String
    Name = mstrName
End Property
Public Property Let Name(ByVal strName As String)
    mstrName = strName
End Property

'设备对象型号
Public Property Get Model() As String
    Model = mstrModel
End Property
Public Property Let Model(ByVal strModel As String)
    mstrModel = strModel
End Property

'设备对象制造商
Public Property Get Manufacturer() As String
    Manufacturer = mstrManufacturer
End Property
Public Property Let Manufacturer(ByVal strManufacturer As String)
    mstrManufacturer = strManufacturer
End Property

'设备使用部门ID
Public Property Get DeptID() As Long
    DeptID = mlngDeptID
End Property
Public Property Let DeptID(ByVal lngDeptID As Long)
    mlngDeptID = lngDeptID
End Property

''设备使用部门名称
'Public Property Get DeptName() As String
'    DeptName = mstrDeptName
'End Property
'Public Property Let DeptName(ByVal strDeptName As String)
'    mstrDeptName = strDeptName
'End Property

'设备连接类型
Public Property Get LinkType() As Byte
    LinkType = mbytLinkType
End Property
Private Property Let LinkType(ByVal bytLinkType As Byte)
    mbytLinkType = bytLinkType
End Property

'设备连接内容
Private Property Get LinkDescribe() As String
    LinkDescribe = mstrLinkDescribe
End Property
Private Property Let LinkDescribe(ByVal strLinkDescribe As String)
    mstrLinkDescribe = strLinkDescribe
End Property

'设备启用状态
Public Property Get Start() As Boolean
    Start = mblnStart
End Property
Public Property Let Start(ByVal blnStart As Boolean)
    mblnStart = blnStart
End Property

'设备支持的服务对象
Public Property Get ServiceObject() As Byte
    ServiceObject = mbytServiceObject
End Property
Public Property Let ServiceObject(ByVal bytServiceObject As Byte)
    mbytServiceObject = bytServiceObject
End Property

''预配药响应
'Public Property Get DispenseFunc() As Byte
'    DispenseFunc = mbytDispenseFunc
'End Property
'Public Property Let DispenseFunc(ByVal bytDispenseFunc As Byte)
'    mbytDispenseFunc = bytDispenseFunc
'End Property
'
''发送响应
'Public Property Get SendFunc() As Byte
'    SendFunc = mbytSendFunc
'End Property
'Public Property Let SendFunc(ByVal bytSendFunc As Byte)
'    mbytSendFunc = bytSendFunc
'End Property

''设备支持的单据类型(住院)
'Public Property Get Bill() As String
'    Bill = mstrBill
'End Property
'Public Property Let Bill(ByVal strBill As String)
'    mstrBill = strBill
'End Property

'设备支持的药品剂型
Public Property Get DrugForm() As String
    DrugForm = mstrDrugForm
End Property
Public Property Let DrugForm(ByVal strDrugForm As String)
    mstrDrugForm = strDrugForm
End Property


'设备数据连接状态
Public Property Get Status() As Boolean
    Status = mblnStatus
End Property

'DB连接对象
Public Property Get DBConnect() As ADODB.Connection
    Set DBConnect = mcnDB
End Property
Public Property Let DBConnect(ByVal cnDB As ADODB.Connection)
    Set mcnDB = cnDB
End Property

'WebServices连接对象
Public Property Get WSConnect() As Object
    Set WSConnect = mcnWS
End Property

''其他连接
'Public Property Get Connect() As String
'    Connect = mstrConnect
'End Property
'Public Property Let Connect(ByVal strConn As String)
'    Const STR_URL = "URL"
'    Const STR_USER = "USER"
'    Const STR_PWD = "PWD"
'
'    Dim arrTmp As Variant
'    Dim strKey As String, strValue As String
'    Dim i As Integer
'
'    mstrConnect = strConn
'    mblnStatus = True
'
'    On Error GoTo errHandle
'    If LinkType = enuLinkType.WEBServices Then
'        arrTmp = Split(mstrConnect, GSTR_SEPARAT_CHILD)
'        For i = LBound(arrTmp) To UBound(arrTmp)
'            strKey = UCase(Left(arrTmp(i), InStr(arrTmp(i), "=") - 1))
'            strValue = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
'            Select Case strKey
'                Case STR_URL
'                    '地址
'                    mstrURL = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
'                Case STR_USER
'                    '用户
'                    mstrUser = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
'                Case STR_PWD, "PASSWORD"
'                    '口令
'                    mstrPassword = Mid(arrTmp(i), InStr(arrTmp(i), "=") + 1)
'            End Select
'        Next
'
'        'WebServices
'        CreateWebServices mstrURL, mcnWS
'
'    End If
'    Exit Property
'
'errHandle:
'    mblnStatus = False
'    gstrMessage = Err.Description
'End Property


'连接用户
Public Property Get User() As String
    User = mstrUser
End Property

'连接密码
Public Property Get Password() As String
    Password = mstrPassword
End Property

'连接地址
Public Property Get URL() As String
    URL = mstrURL
End Property
