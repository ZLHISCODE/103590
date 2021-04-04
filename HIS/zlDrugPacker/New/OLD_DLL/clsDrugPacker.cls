VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrugPacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Set gcolConn = New Collection
    Set gcolDevice = New Collection
End Sub

Private Sub Class_Terminate()
    Dim frmItem As Form
    
    On Error Resume Next
    For Each frmItem In Forms
        Unload frmItem
    Next
    
    Set gcolDevice = Nothing
    Set gcolConn = Nothing
    Set frmItem = Nothing
    Set gobjComLib = Nothing
End Sub

Public Function InitCommon(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, _
        ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal blnInitLink As Boolean) As Boolean
'功能：初始化方法
'参数：
'  cnOracle：HIS的DB连接对象
'  frmOwner：主调窗体对象
'  lngSys：主调系统号
'  lngmodule：主调模块号
'  blnInitLink：是否初始化自动化设备的连接。通常情况在上传、下载业务设置为True。

    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strKey As String
    Dim frmTmp As frmSetting
    Dim blnCalled As Boolean
    Dim blnMenuItem As Boolean
    
    On Error GoTo errHandle
    
    Set gobjConn = cnOracle
    Set gfrmOwner = frmOwner
    glngSys = lngSys
    glngModule = lngModule
    
    '创建zl9Comlib部件
    On Error Resume Next
    'Set gobjComLib = New clsComLib
    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
    If Err <> 0 Or gobjComLib Is Nothing Then
        gstrMessage = "引用zl9ComLib部件失败！" & vbNewLine & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo errHandle
    
    '注册医院名称
    gstrRegHospital = gobjComLib.zlRegInfo("单位名称", , -1)
    
    '获取用户信息
    Call GetUserInfo
    
    '获取接口的权限
    '?????
    gstrPrivs = ";设备开关;上传参数设置;药品信息上传;药品库存上传;"
    
    '获取设备参数
    Call GetDeviceParam
    
    '嵌入菜单
    If Not gfrmOwner.mnuDrugPackerItems Is Nothing Then
        If InStr(gstrPrivs, ";药品信息上传;") > 0 Then
            Call SetMenuItem(enuMenuNo.药品信息)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";药品库存上传;") > 0 Then
            Call SetMenuItem(enuMenuNo.药品库存)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";设备开关;") > 0 Then
            Call SetMenuItem(enuMenuNo.设备开关)
            blnMenuItem = True
        End If
        If InStr(gstrPrivs, ";上传参数设置;") > 0 Then
            Call SetMenuItem(enuMenuNo.上传设置)
            blnMenuItem = True
        End If
        '显示菜单
        If blnMenuItem Then Call ShowMenuItem
    End If
    
    If blnInitLink = False Then
        InitCommon = True
        Exit Function
    End If
    
    
    
    
makInit:
    '创建自动化药房的连接
    On Error GoTo errSQL
    gstrSQL = "select ID, 名称, 连接类型, 连接内容 from 药房设备连接 "
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "获取自动化系统的连接")
    On Error GoTo errHandle
    
    If rsTmp.EOF Then
        rsTmp.Close
        
        If blnCalled = False Then
            
            If MsgBox("尚未对自动化系统设置连接，现在是否设置？", vbInformation + vbYesNo + vbDefaultButton1, GSTR_INTERFACE_NAME) = vbNo Then
                Exit Function
            End If
        
            '调用连接设备窗体
            Set frmTmp = New frmSetting
            frmTmp.Show vbModal, frmOwner
            Set frmTmp = Nothing
            
            blnCalled = True
            GoTo makInit
        
        End If
        
    Else
        
        Do While Not rsTmp.EOF
            strKey = rsTmp!名称
            
            gcolConn.Add New clsConnect, strKey
            gcolConn(strKey).Name = strKey
            gcolConn(strKey).LinkType = gobjComLib.zlCommFun.NVL(rsTmp!连接类型, 0)
            
            Select Case gcolConn(strKey).LinkType
                Case enuLinkType.DB
                    With gcolConn(strKey)
                        .DBConnect = New ADODB.Connection
                        On Error Resume Next
                        .DBConnect.Open rsTmp!连接内容
                        If Err <> 0 Then
                            .Status = False
                            gstrMessage = "连接名：" & strKey & vbNewLine & _
                                          "内容：" & Err.Description
                        Else
                            .Status = True
                        End If
                        Err.Clear: On Error GoTo 0
                    End With
                Case enuLinkType.WEBServices, enuLinkType.Directory
                    With gcolConn(strKey)
                        .Connect = rsTmp!连接内容
                        If .Status = False Then
                            gstrMessage = "连接名：" & strKey & vbNewLine & _
                                          "内容：" & gstrMessage
                        End If
                    End With
            End Select
            rsTmp.MoveNext
        Loop
        
    End If
    rsTmp.Close
    
    InitCommon = True
    Exit Function
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
    Exit Function
    
errSQL:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Public Property Get Message() As String
   Message = gstrMessage
End Property

'Public Sub SetInterface(ByVal cnOracle As ADODB.Connection, ByVal frmOwner As Object, ByVal lngSys As Long, ByVal lngModule As Long)
Public Sub SetInterface(ByVal bytType As Byte, Optional ByVal strKey As String)
'功能：各类功能设置
'参数：
'  bytType：
'    0：接口参数设置
'    1：上传参数设置
'    2：设备开关

    Dim frmTmp As Form
    
    Select Case bytType
        Case 1
            Set frmTmp = New frmDeviceSetting
            frmTmp.ShowMe 53
        Case 2
            Set frmTmp = New frmDeviceState
            frmTmp.ShowMe 53
        Case Else
            Set frmTmp = New frmSetting
            frmTmp.Show vbModal, gfrmOwner
    End Select
    Set frmTmp = Nothing
    
End Sub

Public Function Upload(ByVal lngModule As Long, ByVal bytType As Byte, Optional ByVal varKey As Variant) As Boolean
'功能：通知接口按业务需要上传数据
'参数：
'   lngModule：模块号，用于判断业务
'   bytType：
'       0: 药品基本信息上传
'       1: 门诊处方上传 (配药)
'       2: 门诊发药通知 (发药)
'       3: 住院药品医嘱上传 (配、发药)
'       4: 药品库存上传
'   varKey：
'       当bytType=0时，varKey不填写
'       当bytType=1时，varKey表示“单据;库房ID;NO”；
'       格式：“单据;库房ID;NO[|单据;库房ID;NO][|...]”
'       当bytType=2时，同bytType=1
'       当bytType=3时，varKey表示药品收发ID；
'       格式：“药品收发ID[,药品收发ID][,...]”
'返回：True成功；False失败
    
    Dim rsData As ADODB.Recordset
    Dim frmDrug As frmDrugInfo
    
    glngModule = lngModule
            
    '1.判断当前业务是否需要上传数据
    If GetJudge_IsNeedUpload(lngModule, bytType) = False Then Exit Function
        
    '2.确认设备ID，可能会有多个
        
    '3.上传数据处理，如果存在多个设备，循环处理
    Select Case bytType
        Case 0  '药品基础信息上传
            Set frmDrug = New frmDrugInfo
            frmDrug.Show vbModal, gfrmOwner
            Set frmDrug = Nothing
            
            Upload = True
            Exit Function
        Case 1  '处方明细上传
            '获取HIS数据，并上传
            mdlProcessData.SetUpload bytType, varKey
        Case 2  '发药信息上传(处方概要信息)
            '
        Case 3  '主要药品医嘱信息上传
            '
        Case 4  '药品库存信息上传
            '
        Case Else
            '
    End Select
    
    Upload = True
    Exit Function
    
errHandle:
    '
End Function
