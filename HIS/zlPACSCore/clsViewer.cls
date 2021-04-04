VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'继承窗体的事件
Public Event AfterSaveReportImage(strStudyUID As String)
Public Event AfterSaveOuterImage(strStudyUID As String)

'定义带事件的窗体
Private WithEvents mfrmViewer As frmViewer
Attribute mfrmViewer.VB_VarHelpID = -1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''直接调用观片器的接口
''''胡涛
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub zlShowViewer(ByVal f As Object)
    blLocalRun = True
    If Dir(App.Path & IIf(Len(App.Path) > 3, "\LocalZLPACS.mdb", "LocalZLPACS.mdb")) = "" Then
        If CreateAccessDB = True Then
            '''''''''''''''''[打开数据库连接]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            strADOcn = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & IIf(Len(App.Path) > 3, "\LocalZLPACS.mdb", "LocalZLPACS.mdb")  ''定义ADO数据库联接的参数
            Set cnAccess = New ADODB.Connection
            cnAccess.Open strADOcn
            InitData
        Else
            MsgBox "建立本机数据库失败,不能运行观片站!", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        '''''''''''''''''[打开数据库连接]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        strADOcn = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & IIf(Len(App.Path) > 3, "\LocalZLPACS.mdb", "LocalZLPACS.mdb")  ''定义ADO数据库联接的参数
        Set cnAccess = New ADODB.Connection
        cnAccess.Open strADOcn
    End If
    mfrmViewer.Show , f
End Sub

Public Function zlHasImage(lngOrderID As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：传递一个医嘱ID给接口方法,根据这个医嘱ID进行判断是否有影像图片
'参数：lngOrderID--病人医嘱ID编号
'返回：true--该病人有影像图片  false--该病人没有影像图片
'修改人：罗冠骁
'修改时间：2012-6-14
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

On Error GoTo errHandle
    strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否有影像图片", lngOrderID)
    
    '如果有 检查UID 的记录说明数据库中有图像，则返回True，反之返回false
    If rsTemp.RecordCount < 1 Then
        zlHasImage = False
        Exit Function
    End If
    
    zlHasImage = IIf(Nvl(rsTemp!检查UID) <> 0, True, False)
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function OpenViewer(ByVal strImages As String, ByVal lngOrderID As Long, DBConnect As ADODB.Connection, blnMoved As Boolean, Optional ByVal bAdd As Boolean = False, _
    Optional intImageInterval As Integer = 0, Optional lngSys As Long = 0) As Long
    
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    blLocalRun = False      '初始化为网络版
    
    PstrBufferImagePath = App.Path & "\TmpImage\"
    
    If lngSys <> 0 Then
        glngSys = lngSys
    Else
        glngSys = 100
    End If
    
    '得到权限字串,从虚拟模块1289读取权限
    mstrPrivs = GetPrivFunc(glngSys, 1289)
    
    Set gcnOracle = DBConnect
    Set rsTmp = zl9ComLib.zlDatabase.GetUserInfo
    glngUserID = rsTmp("ID")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    Set frmMain = mfrmViewer
    
    '如果图片不是增加进入观片站，则先删除观片站中原有所有图片，并初始化部分参数
    If Not bAdd Then
        mfrmViewer.subKillPicture True  '删除原有图像，并且隐藏序列切换按钮
        mfrmViewer.MSFViewer.Rows = 1
        mfrmViewer.MSFViewer.Cols = 16
    End If
    
    '显示空的观片窗体,如果是最小化的，自动最大化显示出来
    If mfrmViewer.WindowState = 1 Then
        mfrmViewer.WindowState = 2
    End If
    
    '初始化鼠标左右分配的按键按钮
    Call mfrmViewer.InitComButtonChecked
    DoEvents
    
    If funGetImageStrings(lngOrderID, blnMoved, strImages, intImageInterval) = False Then
        OpenViewer = mfrmViewer.hwnd
        Exit Function
    End If
    
    '把需要显示的图像加载进入VIEWER,把Viewer显示到窗口中，显示Viewer相关的滚动条
    Call subShowForm(mfrmViewer)
    
    '显示缩略图
    Call subShowMiniImages(mfrmViewer)
    
    OpenViewer = mfrmViewer.hwnd
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Public Sub CallOpenViewerSimple(ByVal strImages As String, lngOrderID As Long, _
    strDBConnection As String, blnMoved As Boolean, Optional ByVal bAdd As Boolean = False, _
    Optional intImageInterval As Integer = 0, Optional lngSys As Long = 0, Optional blnReconnectDB As Boolean = False)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：供其他开发程序调用的接口，比如C#，显示观片站，保存成本地缓存，并优先从本地缓存中读取图像
'       根据lngOrderID和strImages从数据库中查找图像文件串，打开一个检查的图像，调用一次本接口
'       由于使用了数据库连接串，仅支持 10.35.10之后的HIS版本
'输入参数说明：
'       strImages--图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
'       lngOrderID -- 医嘱ID
'       strDBConnection－－数据库连接串，包含“服务名[+]用户名[+]密码[+]密码是否转换”，连接符是三个字符“[+]”
'                          当“密码”是用户登录密码时，“密码是否转换”=1；当“密码”是数据库登录密码时，“密码是否转换”=0
'       blnMoved -- 数据是否被转储
'       bAdd－－新图像是增加进观片站，还是替换原观片站所有图像，True为增加，Fasle为替换
'       intImageInterval -- 打开图像的间隔，只对打开全部序列,且序列中图像数量>100时有效
'       lngSys --- 系统序号
'       blnReconnectDB --- 是否重新连接数据库。第一次打开观片时自动连接数据库，之后再打开观片，由blnReconnectDB参数决定是否重新连接数据库
'                           =True，使用strDBConnection参数重新连接数据库；=False，不再重新连接数据库，使用观片部件现在的数据库连接
'2017-5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim objRegister As Object                  '10.35.10之后的注册对象
    Dim objzlComLib As Object
    Dim DBConnect As ADODB.Connection
    Dim strDBPassword As String
    Dim strDBServer As String
    Dim blnTransPassword As Boolean
    Dim strDBUser As String
    Dim strError As String

    On Error GoTo err
    
    If gcnOracle Is Nothing Or blnReconnectDB = True Then
    
        '创建注册对象，连接数据库
        Set objRegister = GetObject("", "zlRegister.clsRegister")
        If objRegister Is Nothing Then
            MsgBox "程序版本太低，无法注册，无法启动观片，只支持10.35.10以上版本。"
            Exit Sub
        End If
        
        '从参数解析用户名密码，连接数据库
            
        '要新建一个数据库连接
        Set DBConnect = New ADODB.Connection
        
        '从参数解析数据库连接信息
        If UBound(Split(strDBConnection, "[+]")) = 3 Then
            strDBServer = Split(strDBConnection, "[+]")(0)
            strDBUser = Split(strDBConnection, "[+]")(1)
            strDBPassword = Split(strDBConnection, "[+]")(2)
            blnTransPassword = (Split(strDBConnection, "[+]")(3) = 1)
            
            '连接数据库
            Set DBConnect = objRegister.GetConnection(strDBServer, strDBUser, strDBPassword, blnTransPassword, , strError, True)
            If DBConnect.State <> adStateOpen Then
                MsgBox "连接数据库出错，请检查数据库连接配置。", vbOKOnly
                Exit Sub
            End If
        Else
            MsgBox "数据库参数格式不正确，请检查。", vbOKOnly
            Exit Sub
        End If
        
        '初始化zlComLib
        Set objzlComLib = GetObject("", "zl9ComLib.clsComLib")

        Call objzlComLib.InitCommon(DBConnect)
    Else
        Set DBConnect = gcnOracle
    End If

    Call CallOpenViewer(strImages, lngOrderID, Nothing, DBConnect, blnMoved, bAdd, intImageInterval, lngSys)
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub CallOpenViewer(ByVal strImages As String, lngOrderID As Long, ByVal objParent As Object, _
    DBConnect As ADODB.Connection, blnMoved As Boolean, Optional ByVal bAdd As Boolean = False, _
    Optional intImageInterval As Integer = 0, Optional lngSys As Long = 0)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：供影像医技站等工作站调用的接口，显示观片站，保存成本地缓存，并优先从本地缓存中读取图像
'       根据lngOrderID和strImages从数据库中查找图像文件串，打开一个检查的图像，调用一次本接口
'输入参数说明：
'       strImages--图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
'       lngOrderID -- 医嘱ID
'       objParent－－父窗体
'       DBConnect－－数据库连接
'       blnMoved -- 数据是否被转储
'       bAdd－－新图像是增加进观片站，还是替换原观片站所有图像，True为增加，Fasle为替换
'       intImageInterval -- 打开图像的间隔，只对打开全部序列,且序列中图像数量>100时有效
'       lngSys --- 系统序号
'2009-7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    blLocalRun = False      '初始化为网络版
    
    PstrBufferImagePath = App.Path & "\TmpImage\"
    
    If lngSys <> 0 Then
        glngSys = lngSys
    Else
        glngSys = 100
    End If
    
    '得到权限字串,从虚拟模块1289读取权限
    mstrPrivs = GetPrivFunc(glngSys, 1289)
    
    Set gcnOracle = DBConnect
    Set rsTmp = zl9ComLib.zlDatabase.GetUserInfo
    glngUserID = rsTmp("ID")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    Set frmMain = mfrmViewer
    
    '如果图片不是增加进入观片站，则先删除观片站中原有所有图片，并初始化部分参数
    If Not bAdd Then
        mfrmViewer.subKillPicture True  '删除原有图像，并且隐藏序列切换按钮
        mfrmViewer.MSFViewer.Rows = 1
        mfrmViewer.MSFViewer.Cols = 16
    End If
    
    '显示空的观片窗体,如果是最小化的，自动最大化显示出来
    If mfrmViewer.WindowState = 1 Then
        mfrmViewer.WindowState = 2
    End If
    
    '检查是否注册成功,注册失败则关闭观片窗口
    If mfrmViewer.blnLogined = False Then
        Call Closefrom
        Exit Sub
    End If
    
    '初始化鼠标左右分配的按键按钮
    Call mfrmViewer.InitComButtonChecked
    DoEvents
    
    If funGetImageStrings(lngOrderID, blnMoved, strImages, intImageInterval) = False Then
        Exit Sub
    End If
    
    '调用Activex EXE下载图像，传入的参数为 strImages，cnOracle。或者直接传FTP和共享目录串
    '后面再增加这个zlDownloadImage.exe
    
    '把需要显示的图像加载进入VIEWER,把Viewer显示到窗口中，显示Viewer相关的滚动条
    Call subShowForm(mfrmViewer)
    
    '显示缩略图
    Call subShowMiniImages(mfrmViewer)
    
    '因为这个窗口可能是模态窗口，所以在全部事情都完成后，最后才显示窗口
    mfrmViewer.ShowMe objParent
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Closefrom()
    '关闭窗口
    Unload mfrmViewer
End Sub
Public Function GetLeft() As Long
    GetLeft = mfrmViewer.left
End Function
Public Function GetTop() As Long
    GetTop = mfrmViewer.top
End Function
Public Function Getwidth() As Long
    Getwidth = mfrmViewer.width
End Function
Public Function GetHeight() As Long
    GetHeight = mfrmViewer.height
End Function
Public Function SetViewerFocus()
    mfrmViewer.SetFocus
End Function
Private Function CreateAccessDB() As Boolean
    '创建MDB如不成功则退出
    
    Dim tbl As New Table
    Dim cat As New ADOX.Catalog
    Dim kyForeign  As New ADOX.Key
    Dim col As New ADOX.Column
    Dim TmpPath As String
    
    On Error GoTo errh
    
    TmpPath = App.Path & IIf(Len(App.Path) > 3, "\LocalZLPACS.mdb", "LocalZLPACS.mdb")
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TmpPath
    
    '--影像颜色清单
    tbl.Name = "影像颜色清单"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "序号", adInteger
    tbl.Columns("序号").Properties("AutoIncrement") = True
    tbl.Columns.Append "颜色", adVarWChar, 200
    tbl.Columns("颜色").Attributes = adColNullable
    tbl.Columns.Append "颜色内容", adLongVarWChar
    tbl.Columns("颜色内容").Attributes = adColNullable
    tbl.Columns.Append "系统方案", adInteger
    tbl.Columns("系统方案").Attributes = adColNullable
    
    kyForeign.Name = "序号_PK"
    kyForeign.Columns.Append "序号"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    
    '--影像打印格式
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像打印格式"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "编号", adInteger
    tbl.Columns("编号").Properties("AutoIncrement") = True
    tbl.Columns.Append "格式", adVarWChar, 200
    tbl.Columns("格式").Attributes = adColNullable
    tbl.Columns.Append "参数1", adInteger
    tbl.Columns("参数1").Attributes = adColNullable
    tbl.Columns.Append "参数2", adInteger
    tbl.Columns("参数2").Attributes = adColNullable
    tbl.Columns.Append "参数3", adInteger
    tbl.Columns("参数3").Attributes = adColNullable
    tbl.Columns.Append "参数4", adInteger
    tbl.Columns("参数4").Attributes = adColNullable
    tbl.Columns.Append "参数5", adInteger
    tbl.Columns("参数5").Attributes = adColNullable
    tbl.Columns.Append "参数6", adInteger
    tbl.Columns("参数6").Attributes = adColNullable
    tbl.Columns.Append "参数7", adInteger
    tbl.Columns("参数7").Attributes = adColNullable
    tbl.Columns.Append "格式标识", adVarWChar, 200
    tbl.Columns("格式标识").Attributes = adColNullable
    
    kyForeign.Name = "编号_PK"
    kyForeign.Columns.Append "编号"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像打印机设置
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像打印机设置"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "打印机名", adVarWChar, 200
    tbl.Columns("打印机名").Attributes = adColNullable
    tbl.Columns.Append "IP地址", adVarWChar, 200
    tbl.Columns("IP地址").Attributes = adColNullable
    tbl.Columns.Append "端口号", adInteger
    tbl.Columns("端口号").Attributes = adColNullable
    tbl.Columns.Append "AE名称", adVarWChar, 200
    tbl.Columns("AE名称").Attributes = adColNullable
    tbl.Columns.Append "打印格式", adVarWChar, 200
    tbl.Columns("打印格式").Attributes = adColNullable
    tbl.Columns.Append "优先级", adVarWChar, 200
    tbl.Columns("优先级").Attributes = adColNullable
    tbl.Columns.Append "打印份数", adInteger
    tbl.Columns("打印份数").Attributes = adColNullable
    tbl.Columns.Append "介质", adVarWChar, 200
    tbl.Columns("介质").Attributes = adColNullable
    tbl.Columns.Append "方向", adVarWChar, 200
    tbl.Columns("方向").Attributes = adColNullable
    tbl.Columns.Append "胶片规格", adVarWChar, 200
    tbl.Columns("胶片规格").Attributes = adColNullable
    tbl.Columns.Append "选用片盒", adVarWChar, 200
    tbl.Columns("选用片盒").Attributes = adColNullable
    tbl.Columns.Append "分辨率", adVarWChar, 200
    tbl.Columns("分辨率").Attributes = adColNullable
    tbl.Columns.Append "放大模式", adVarWChar, 200
    tbl.Columns("放大模式").Attributes = adColNullable
    tbl.Columns.Append "平滑模式", adVarWChar, 200
    tbl.Columns("平滑模式").Attributes = adColNullable
    tbl.Columns.Append "修整", adVarWChar, 200
    tbl.Columns("修整").Attributes = adColNullable
    tbl.Columns.Append "最小密度", adVarWChar, 200
    tbl.Columns("最小密度").Attributes = adColNullable
    tbl.Columns.Append "最大密度", adVarWChar, 200
    tbl.Columns("最大密度").Attributes = adColNullable
    tbl.Columns.Append "空白密度", adVarWChar, 200
    tbl.Columns("空白密度").Attributes = adColNullable
    tbl.Columns.Append "边框密度", adVarWChar, 200
    tbl.Columns("边框密度").Attributes = adColNullable
    tbl.Columns.Append "极性", adVarWChar, 200
    tbl.Columns("极性").Attributes = adColNullable
    tbl.Columns.Append "图像位数", adInteger
    tbl.Columns("图像位数").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像胶片规格
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像胶片规格"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "编号", adInteger
    tbl.Columns("编号").Properties("AutoIncrement") = True
    tbl.Columns.Append "胶片宽度", adVarWChar, 200
    tbl.Columns("胶片宽度").Attributes = adColNullable
    tbl.Columns.Append "胶片长度", adVarWChar, 200
    tbl.Columns("胶片长度").Attributes = adColNullable
    tbl.Columns.Append "单位", adVarWChar, 200
    tbl.Columns("单位").Attributes = adColNullable
    tbl.Columns.Append "规格标识", adVarWChar, 200
    tbl.Columns("规格标识").Attributes = adColNullable
    
    kyForeign.Name = "编号_PK"
    kyForeign.Columns.Append "编号"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像标注存储表
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像标注存储表"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "编号", adInteger
    tbl.Columns("编号").Properties("AutoIncrement") = True
    tbl.Columns.Append "VGROUP", adVarWChar, 200
    tbl.Columns("VGROUP").Attributes = adColNullable
    tbl.Columns.Append "ELEMENT", adVarWChar, 200
    tbl.Columns("ELEMENT").Attributes = adColNullable
    tbl.Columns.Append "VR", adVarWChar, 200
    tbl.Columns("VR").Attributes = adColNullable
    tbl.Columns.Append "标注属性", adVarWChar, 200
    tbl.Columns("标注属性").Attributes = adColNullable
    
    kyForeign.Name = "编号_PK"
    kyForeign.Columns.Append "编号"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--错误日志
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像错误日志"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "产生时间", adDate
    tbl.Columns("产生时间").Attributes = adColNullable
    tbl.Columns.Append "错误号", adInteger
    tbl.Columns("错误号").Attributes = adColNullable
    tbl.Columns.Append "错误信息", adWChar
    tbl.Columns("错误信息").Attributes = adColNullable
    tbl.Columns.Append "错误类型", adInteger
    tbl.Columns("错误类型").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像图像信息表
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像图像信息表"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "开始地址", adVarWChar, 200
    tbl.Columns("开始地址").Attributes = adColNullable
    tbl.Columns.Append "结束地址", adVarWChar, 200
    tbl.Columns("结束地址").Attributes = adColNullable
    tbl.Columns.Append "英文名称", adVarWChar, 200
    tbl.Columns("英文名称").Attributes = adColNullable
    tbl.Columns.Append "中文名称", adVarWChar, 200
    tbl.Columns("中文名称").Attributes = adColNullable
    tbl.Columns.Append "中文简称", adVarWChar, 200
    tbl.Columns("中文简称").Attributes = adColNullable
    tbl.Columns.Append "英文简称", adVarWChar, 200
    tbl.Columns("英文简称").Attributes = adColNullable
    tbl.Columns.Append "常用", adInteger
    tbl.Columns("常用").Attributes = adColNullable
    tbl.Columns.Append "被选用", adInteger
    tbl.Columns("被选用").Attributes = adColNullable
    tbl.Columns.Append "位置", adInteger
    tbl.Columns("位置").Attributes = adColNullable
    tbl.Columns.Append "角内序号", adInteger
    tbl.Columns("角内序号").Attributes = adColNullable
    tbl.Columns.Append "使用计算", adInteger
    tbl.Columns("使用计算").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像胶片打印字体
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像胶片打印字体"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "影像类别", adVarWChar, 200
    
    tbl.Columns.Append "字体大小", adInteger
    tbl.Columns("字体大小").Attributes = adColNullable
    tbl.Columns.Append "是否随图像缩放", adInteger
    tbl.Columns("是否随图像缩放").Attributes = adColNullable
    
    kyForeign.Name = "影像类别_PK"
    kyForeign.Columns.Append "影像类别"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像鼠标按钮分配
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像鼠标按钮分配"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    
    tbl.Columns.Append "直线", adVarWChar, 200
    tbl.Columns("直线").Attributes = adColNullable
    tbl.Columns.Append "矩形", adVarWChar, 200
    tbl.Columns("矩形").Attributes = adColNullable
    tbl.Columns.Append "椭圆", adVarWChar, 200
    tbl.Columns("椭圆").Attributes = adColNullable
    tbl.Columns.Append "箭头", adVarWChar, 200
    tbl.Columns("箭头").Attributes = adColNullable
    tbl.Columns.Append "多边形", adVarWChar, 200
    tbl.Columns("多边形").Attributes = adColNullable
    tbl.Columns.Append "多边线", adVarWChar, 200
    tbl.Columns("多边线").Attributes = adColNullable
    tbl.Columns.Append "角度", adVarWChar, 200
    tbl.Columns("角度").Attributes = adColNullable
    tbl.Columns.Append "文字", adVarWChar, 200
    tbl.Columns("文字").Attributes = adColNullable
    tbl.Columns.Append "穿梭定位", adVarWChar, 200
    tbl.Columns("穿梭定位").Attributes = adColNullable
    tbl.Columns.Append "窗宽窗位", adVarWChar, 200
    tbl.Columns("窗宽窗位").Attributes = adColNullable
    tbl.Columns.Append "漫游", adVarWChar, 200
    tbl.Columns("漫游").Attributes = adColNullable
    tbl.Columns.Append "缩放", adVarWChar, 200
    tbl.Columns("缩放").Attributes = adColNullable
    tbl.Columns.Append "裁剪_标注调整", adVarWChar, 200
    tbl.Columns("裁剪_标注调整").Attributes = adColNullable
    tbl.Columns.Append "自适应调窗", adVarWChar, 200
    tbl.Columns("自适应调窗").Attributes = adColNullable
    tbl.Columns.Append "三维鼠标", adVarWChar, 200
    tbl.Columns("三维鼠标").Attributes = adColNullable
    tbl.Columns.Append "画标注", adVarWChar, 200
    tbl.Columns("画标注").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像图像消隐表
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像图像消隐表"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "影像类型", adVarWChar, 200
    tbl.Columns("影像类型").Attributes = adColNullable
    tbl.Columns.Append "消隐类型", adVarWChar, 200
    tbl.Columns("消隐类型").Attributes = adColNullable
    tbl.Columns.Append "圆心X", adInteger
    tbl.Columns("圆心X").Attributes = adColNullable
    tbl.Columns.Append "圆心Y", adInteger
    tbl.Columns("圆心Y").Attributes = adColNullable
    tbl.Columns.Append "圆形半径", adInteger
    tbl.Columns("圆形半径").Attributes = adColNullable
    tbl.Columns.Append "矩形左边界", adInteger
    tbl.Columns("矩形左边界").Attributes = adColNullable
    tbl.Columns.Append "矩形右边界", adInteger
    tbl.Columns("矩形右边界").Attributes = adColNullable
    tbl.Columns.Append "矩形上边界", adInteger
    tbl.Columns("矩形上边界").Attributes = adColNullable
    tbl.Columns.Append "矩形下边界", adInteger
    tbl.Columns("矩形下边界").Attributes = adColNullable
    tbl.Columns.Append "多边形顶点", adVarWChar, 200
    tbl.Columns("多边形顶点").Attributes = adColNullable
    tbl.Columns.Append "消隐颜色", adInteger
    tbl.Columns("消隐颜色").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像界面参数表
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像界面参数表"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "正常图像边框颜色", adVarWChar, 200
    tbl.Columns("正常图像边框颜色").Attributes = adColNullable
    tbl.Columns.Append "正常图像边框线型", adVarWChar, 200
    tbl.Columns("正常图像边框线型").Attributes = adColNullable
    tbl.Columns.Append "正常图像边框线宽", adVarWChar, 200
    tbl.Columns("正常图像边框线宽").Attributes = adColNullable
    tbl.Columns.Append "选中图像边框颜色", adVarWChar, 200
    tbl.Columns("选中图像边框颜色").Attributes = adColNullable
    tbl.Columns.Append "选中序列边框颜色", adVarWChar, 200
    tbl.Columns("选中序列边框颜色").Attributes = adColNullable
    tbl.Columns.Append "选中图像边框线型", adVarWChar, 200
    tbl.Columns("选中图像边框线型").Attributes = adColNullable
    tbl.Columns.Append "选中图像边框线宽", adVarWChar, 200
    tbl.Columns("选中图像边框线宽").Attributes = adColNullable
    tbl.Columns.Append "图像标记颜色", adVarWChar, 200
    tbl.Columns("图像标记颜色").Attributes = adColNullable
    tbl.Columns.Append "图像标记大小", adVarWChar, 200
    tbl.Columns("图像标记大小").Attributes = adColNullable
    tbl.Columns.Append "标注选择句柄颜色", adVarWChar, 200
    tbl.Columns("标注选择句柄颜色").Attributes = adColNullable
    tbl.Columns.Append "标注选择句柄大小", adVarWChar, 200
    tbl.Columns("标注选择句柄大小").Attributes = adColNullable
    tbl.Columns.Append "定位线颜色", adVarWChar, 200
    tbl.Columns("定位线颜色").Attributes = adColNullable
    tbl.Columns.Append "定位线线型", adVarWChar, 200
    tbl.Columns("定位线线型").Attributes = adColNullable
    tbl.Columns.Append "定位线间距", adVarWChar, 200
    tbl.Columns("定位线间距").Attributes = adColNullable
    tbl.Columns.Append "序列间间隔", adVarWChar, 200
    tbl.Columns("序列间间隔").Attributes = adColNullable
    tbl.Columns.Append "横向最大序列", adVarWChar, 200
    tbl.Columns("横向最大序列").Attributes = adColNullable
    tbl.Columns.Append "纵向最大序列", adVarWChar, 200
    tbl.Columns("纵向最大序列").Attributes = adColNullable
    tbl.Columns.Append "图像间距", adVarWChar, 200
    tbl.Columns("图像间距").Attributes = adColNullable
    tbl.Columns.Append "显示多余边框", adVarWChar, 200
    tbl.Columns("显示多余边框").Attributes = adColNullable
    tbl.Columns.Append "背景颜色", adVarWChar, 200
    tbl.Columns("背景颜色").Attributes = adColNullable
    tbl.Columns.Append "程序背景颜色", adVarWChar, 200
    tbl.Columns("程序背景颜色").Attributes = adColNullable
    tbl.Columns.Append "标注正常颜色", adVarWChar, 200
    tbl.Columns("标注正常颜色").Attributes = adColNullable
    tbl.Columns.Append "标注正常线型", adVarWChar, 200
    tbl.Columns("标注正常线型").Attributes = adColNullable
    tbl.Columns.Append "标注正常线宽", adVarWChar, 200
    tbl.Columns("标注正常线宽").Attributes = adColNullable
    tbl.Columns.Append "标注选中颜色", adVarWChar, 200
    tbl.Columns("标注选中颜色").Attributes = adColNullable
    tbl.Columns.Append "标注选中线型", adVarWChar, 200
    tbl.Columns("标注选中线型").Attributes = adColNullable
    tbl.Columns.Append "标注选中线宽", adVarWChar, 200
    tbl.Columns("标注选中线宽").Attributes = adColNullable
    tbl.Columns.Append "标注文字大小", adVarWChar, 200
    tbl.Columns("标注文字大小").Attributes = adColNullable
    tbl.Columns.Append "测量显示面积", adVarWChar, 200
    tbl.Columns("测量显示面积").Attributes = adColNullable
    tbl.Columns.Append "测量显示平均值", adVarWChar, 200
    tbl.Columns("测量显示平均值").Attributes = adColNullable
    tbl.Columns.Append "测量显示均方差", adVarWChar, 200
    tbl.Columns("测量显示均方差").Attributes = adColNullable
    tbl.Columns.Append "测量显示中文", adVarWChar, 200
    tbl.Columns("测量显示中文").Attributes = adColNullable
    tbl.Columns.Append "文字X方向偏移", adVarWChar, 200
    tbl.Columns("文字X方向偏移").Attributes = adColNullable
    tbl.Columns.Append "文字Y方向偏移", adVarWChar, 200
    tbl.Columns("文字Y方向偏移").Attributes = adColNullable
    tbl.Columns.Append "文字随图像缩放", adVarWChar, 200
    tbl.Columns("文字随图像缩放").Attributes = adColNullable
    tbl.Columns.Append "显示体位标记", adVarWChar, 200
    tbl.Columns("显示体位标记").Attributes = adColNullable
    tbl.Columns.Append "中文体位标记", adVarWChar, 200
    tbl.Columns("中文体位标记").Attributes = adColNullable
    tbl.Columns.Append "显示标尺", adVarWChar, 200
    tbl.Columns("显示标尺").Attributes = adColNullable
    tbl.Columns.Append "标尺左右边距", adVarWChar, 200
    tbl.Columns("标尺左右边距").Attributes = adColNullable
    tbl.Columns.Append "标尺上下边距", adVarWChar, 200
    tbl.Columns("标尺上下边距").Attributes = adColNullable
    tbl.Columns.Append "标尺宽度", adVarWChar, 200
    tbl.Columns("标尺宽度").Attributes = adColNullable
    tbl.Columns.Append "标尺高度", adVarWChar, 200
    tbl.Columns("标尺高度").Attributes = adColNullable
    tbl.Columns.Append "标尺线宽", adVarWChar, 200
    tbl.Columns("标尺线宽").Attributes = adColNullable
    tbl.Columns.Append "标尺颜色", adVarWChar, 200
    tbl.Columns("标尺颜色").Attributes = adColNullable
    tbl.Columns.Append "窗宽窗位位置", adVarWChar, 200
    tbl.Columns("窗宽窗位位置").Attributes = adColNullable
    tbl.Columns.Append "鼠标穿梭步长", adVarWChar, 200
    tbl.Columns("鼠标穿梭步长").Attributes = adColNullable
    tbl.Columns.Append "鼠标漫游步长", adVarWChar, 200
    tbl.Columns("鼠标漫游步长").Attributes = adColNullable
    tbl.Columns.Append "鼠标调窗步长", adVarWChar, 200
    tbl.Columns("鼠标调窗步长").Attributes = adColNullable
    tbl.Columns.Append "鼠标缩放步长", adVarWChar, 200
    tbl.Columns("鼠标缩放步长").Attributes = adColNullable
    tbl.Columns.Append "病人信息上下边距", adVarWChar, 200
    tbl.Columns("病人信息上下边距").Attributes = adColNullable
    tbl.Columns.Append "病人信息左右边距", adVarWChar, 200
    tbl.Columns("病人信息左右边距").Attributes = adColNullable
    tbl.Columns.Append "病人信息颜色", adVarWChar, 200
    tbl.Columns("病人信息颜色").Attributes = adColNullable
    tbl.Columns.Append "病人信息显示最小值", adVarWChar, 200
    tbl.Columns("病人信息显示最小值").Attributes = adColNullable
    tbl.Columns.Append "病人信息随图像缩放", adVarWChar, 200
    tbl.Columns("病人信息随图像缩放").Attributes = adColNullable
    tbl.Columns.Append "病人信息字体", adVarWChar, 200
    tbl.Columns("病人信息字体").Attributes = adColNullable
    tbl.Columns.Append "病人信息题头", adVarWChar, 200
    tbl.Columns("病人信息题头").Attributes = adColNullable
    tbl.Columns.Append "直接照相", adVarWChar, 200
    tbl.Columns("直接照相").Attributes = adColNullable
    tbl.Columns.Append "工具栏图标大小", adVarWChar, 200
    tbl.Columns("工具栏图标大小").Attributes = adColNullable
    tbl.Columns.Append "工具栏位置", adVarWChar, 200
    tbl.Columns("工具栏位置").Attributes = adColNullable
    tbl.Columns.Append "工具栏显示", adVarWChar, 200
    tbl.Columns("工具栏显示").Attributes = adColNullable
    tbl.Columns.Append "状态栏字体大小", adVarWChar, 200
    tbl.Columns("状态栏字体大小").Attributes = adColNullable
    tbl.Columns.Append "正常血管阈值", adVarWChar, 200
    tbl.Columns("正常血管阈值").Attributes = adColNullable
    tbl.Columns.Append "狭窄血管阈值", adVarWChar, 200
    tbl.Columns("狭窄血管阈值").Attributes = adColNullable
    tbl.Columns.Append "血管壁宽度", adVarWChar, 200
    tbl.Columns("血管壁宽度").Attributes = adColNullable
    tbl.Columns.Append "测量显示周长", adVarWChar, 200
    tbl.Columns("测量显示周长").Attributes = adColNullable

    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像预设窗宽窗位
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像预设窗宽窗位"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "影像类型", adVarWChar, 200
    tbl.Columns("影像类型").Attributes = adColNullable
    tbl.Columns.Append "快捷键", adInteger
    tbl.Columns("快捷键").Attributes = adColNullable
    tbl.Columns.Append "窗口名称", adVarWChar, 200
    tbl.Columns("窗口名称").Attributes = adColNullable
    tbl.Columns.Append "窗口英文名", adVarWChar, 200
    tbl.Columns("窗口英文名").Attributes = adColNullable
    tbl.Columns.Append "窗宽", adInteger
    tbl.Columns("窗宽").Attributes = adColNullable
    tbl.Columns.Append "窗位", adInteger
    tbl.Columns("窗位").Attributes = adColNullable
    tbl.Columns.Append "是否默认", adInteger
    tbl.Columns("是否默认").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    
    '--影像屏幕布局
    Set tbl = Nothing
    Set kyForeign = Nothing
    tbl.Name = "影像屏幕布局"
    
    Set tbl.ParentCatalog = cat
    tbl.Columns.Append "ID", adInteger
    tbl.Columns("ID").Properties("AutoIncrement") = True
    tbl.Columns.Append "影像类型", adVarWChar, 200
    tbl.Columns("影像类型").Attributes = adColNullable
    tbl.Columns.Append "自动序列布局", adInteger
    tbl.Columns("自动序列布局").Attributes = adColNullable
    tbl.Columns.Append "自动图像布局", adInteger
    tbl.Columns("自动图像布局").Attributes = adColNullable
    tbl.Columns.Append "序列行数", adInteger
    tbl.Columns("序列行数").Attributes = adColNullable
    tbl.Columns.Append "序列列数", adInteger
    tbl.Columns("序列列数").Attributes = adColNullable
    tbl.Columns.Append "图像行数", adInteger
    tbl.Columns("图像行数").Attributes = adColNullable
    tbl.Columns.Append "图像列数", adInteger
    tbl.Columns("图像列数").Attributes = adColNullable
    tbl.Columns.Append "自动反白", adInteger
    tbl.Columns("自动反白").Attributes = adColNullable
    tbl.Columns.Append "显示病人信息", adInteger
    tbl.Columns("显示病人信息").Attributes = adColNullable
    tbl.Columns.Append "选择定位线", adInteger
    tbl.Columns("选择定位线").Attributes = adColNullable
    tbl.Columns.Append "选择序列同步", adInteger
    tbl.Columns("选择序列同步").Attributes = adColNullable
    tbl.Columns.Append "插值模式", adInteger
    tbl.Columns("插值模式").Attributes = adColNullable
    tbl.Columns.Append "图像排序", adInteger
    tbl.Columns("图像排序").Attributes = adColNullable
    
    kyForeign.Name = "ID_PK"
    kyForeign.Columns.Append "ID"
    
    tbl.Keys.Append kyForeign
    
    cat.Tables.Append tbl
    CreateAccessDB = True
    
    Exit Function
errh:
    Kill TmpPath
    MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
End Function


Private Sub InitData()
    Dim strSQL As Variant
    
    '影像颜色清单
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (1,'10_Percent','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;"
    strSQL = strSQL & "37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;"
    strSQL = strSQL & "0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;"
    strSQL = strSQL & "0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;"
    strSQL = strSQL & "0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;"
    strSQL = strSQL & "0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;"
    strSQL = strSQL & "0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;"
    strSQL = strSQL & "0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;0,214,20,1;"
    strSQL = strSQL & "183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;"
    strSQL = strSQL & "183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;"
    strSQL = strSQL & "183,214,0,1;183,214,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;"
    strSQL = strSQL & "255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;"
    strSQL = strSQL & "255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;"
    strSQL = strSQL & "255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;"
    strSQL = strSQL & "255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;"
    strSQL = strSQL & "255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;"
    strSQL = strSQL & "255,5,5,1;255,5,5,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;"
    strSQL = strSQL & "255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;"
    strSQL = strSQL & "255,172,209,1;255,172,209,1;255,172,209,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (2,'Thermal','0,0,0,1;0,0,0,1;3,0,0,1;6,0,0,1;10,0,0,1;13,0,0,1;17,0,0,1;20,0,0,1;24,1,0,1;"
    strSQL = strSQL & "27,1,0,1;30,1,0,1;33,1,0,1;36,2,0,1;39,2,0,1;42,2,0,1;45,2,0,1;48,2,0,1;51,2,0,1;54,2,0,1;57,2,0,1;60,3,0,1;63,3,0,1;66,3,0,1;69,3,0,1;72,4,0,1;"
    strSQL = strSQL & "75,4,0,1;78,4,0,1;81,4,0,1;84,4,0,1;85,5,0,1;86,6,0,1;87,7,0,1;88,9,0,1;90,9,0,1;92,10,0,1;94,11,0,1;96,12,0,1;97,13,0,1;98,14,0,1;99,15,0,1;"
    strSQL = strSQL & "100,17,0,1;102,17,0,1;104,18,0,1;106,18,0,1;108,19,0,1;109,20,0,1;110,21,0,1;111,22,0,1;112,24,0,1;114,24,0,1;116,25,0,1;118,26,0,1;120,27,0,1;"
    strSQL = strSQL & "121,28,0,1;122,29,0,1;123,30,0,1;124,32,0,1;126,32,0,1;128,33,0,1;130,33,0,1;132,34,0,1;133,35,0,1;134,36,0,1;135,37,0,1;136,39,0,1;138,39,1,1;"
    strSQL = strSQL & "140,40,2,1;142,40,3,1;144,41,4,1;145,42,4,1;146,43,4,1;147,44,4,1;148,46,4,1;150,46,4,1;152,47,4,1;154,48,4,1;156,49,4,1;157,50,4,1;158,51,4,1;"
    strSQL = strSQL & "159,52,4,1;160,54,4,1;162,54,5,1;164,54,6,1;166,54,7,1;168,55,8,1;169,56,8,1;170,57,8,1;171,58,8,1;172,60,8,1;174,60,8,1;176,61,8,1;178,62,8,1;"
    strSQL = strSQL & "180,63,8,1;181,64,8,1;182,65,8,1;183,66,8,1;184,68,8,1;185,69,8,1;186,70,8,1;187,71,8,1;188,73,8,1;189,74,8,1;190,75,8,1;191,76,8,1;192,77,8,1;"
    strSQL = strSQL & "193,78,8,1;194,79,8,1;195,80,8,1;196,82,8,1;198,82,9,1;200,83,10,1;202,83,11,1;204,84,12,1;205,85,12,1;206,86,12,1;207,87,12,1;208,89,12,1;"
    strSQL = strSQL & "209,90,12,1;210,91,12,1;211,92,12,1;212,94,12,1;213,95,12,1;214,96,12,1;215,97,12,1;216,98,12,1;217,99,12,1;218,100,12,1;219,101,12,1;"
    strSQL = strSQL & "220,103,12,1;222,103,12,1;224,104,12,1;226,105,12,1;228,106,12,1;229,107,13,1;230,108,14,1;231,109,15,1;232,110,16,1;233,111,16,1;234,112,16,1;"
    strSQL = strSQL & "235,113,16,1;236,115,16,1;237,116,16,1;238,117,16,1;239,118,16,1;240,119,16,1;241,120,16,1;242,121,16,1;243,122,16,1;244,124,16,1;246,124,17,1;"
    strSQL = strSQL & "248,125,18,1;250,125,19,1;252,126,20,1;252,127,20,1;252,129,20,1;252,131,20,1;252,133,20,1;252,134,20,1;252,136,20,1;252,138,20,1;252,140,20,1;"
    strSQL = strSQL & "252,141,21,1;252,143,22,1;252,144,23,1;252,146,24,1;252,147,24,1;252,149,24,1;252,151,24,1;252,153,24,1;252,154,25,1;252,156,26,1;252,157,27,1;"
    strSQL = strSQL & "252,159,28,1;252,160,30,1;252,161,32,1;252,162,34,1;252,164,36,1;252,165,38,1;252,166,40,1;252,167,42,1;252,169,44,1;252,170,46,1;252,172,48,1;"
    strSQL = strSQL & "252,173,50,1;252,175,52,1;252,176,54,1;252,177,56,1;252,178,58,1;252,180,60,1;252,181,61,1;252,183,62,1;252,184,63,1;252,186,64,1;252,187,66,1;"
    strSQL = strSQL & "252,188,68,1;252,189,70,1;252,191,72,1;252,192,74,1;252,193,76,1;252,194,78,1;252,196,80,1;252,197,82,1;252,199,84,1;252,200,86,1;252,202,88,1;"
    strSQL = strSQL & "252,203,90,1;252,204,92,1;252,205,94,1;252,207,96,1;252,208,98,1;252,209,100,1;252,210,102,1;252,212,104,1;252,213,107,1;252,214,110,1;252,215,113,1;"
    strSQL = strSQL & "252,217,116,1;252,218,119,1;252,219,122,1;252,220,125,1;252,221,128,1;252,222,131,1;252,223,134,1;252,224,137,1;252,226,140,1;252,227,143,1;252,228,146,1;"
    strSQL = strSQL & "252,229,149,1;252,230,152,1;252,231,155,1;252,232,158,1;252,233,161,1;252,235,164,1;252,236,167,1;252,237,170,1;252,238,173,1;252,239,176,1;252,240,180,1;"
    strSQL = strSQL & "252,241,184,1;252,242,188,1;252,243,192,1;252,243,197,1;252,244,202,1;252,245,207,1;252,246,212,1;252,246,217,1;252,247,222,1;252,248,227,1;252,249,232,1;"
    strSQL = strSQL & "252,249,237,1;252,250,242,1;252,251,247,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (3,'16color','0,0,0,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;"
    strSQL = strSQL & "0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,168,1;0,0,168,1;"
    strSQL = strSQL & "0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;0,0,168,1;"
    strSQL = strSQL & "0,0,168,1;0,0,168,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;"
    strSQL = strSQL & "0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,84,255,1;0,168,255,1;"
    strSQL = strSQL & "0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;"
    strSQL = strSQL & "0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;"
    strSQL = strSQL & "128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;128,255,0,1;"
    strSQL = strSQL & "255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;"
    strSQL = strSQL & "255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;"
    strSQL = strSQL & "255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;255,180,0,1;"
    strSQL = strSQL & "255,180,0,1;255,180,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;"
    strSQL = strSQL & "255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,64,0,1;255,64,0,1;"
    strSQL = strSQL & "255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;255,64,0,1;"
    strSQL = strSQL & "255,64,0,1;255,64,0,1;255,64,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;"
    strSQL = strSQL & "220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;220,190,190,1;"
    strSQL = strSQL & "220,190,190,1;220,190,190,1;220,190,190,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;"
    strSQL = strSQL & "220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;220,220,220,1;"
    strSQL = strSQL & "220,220,220,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (4,'5percent','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;19,0,60,1;"
    strSQL = strSQL & "19,0,60,1;19,0,60,1;19,0,60,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;"
    strSQL = strSQL & "37,0,118,1;37,0,118,1;37,0,118,1;37,0,118,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;"
    strSQL = strSQL & "0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,12,153,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;"
    strSQL = strSQL & "0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,23,188,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;"
    strSQL = strSQL & "0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,46,138,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;"
    strSQL = strSQL & "0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,115,88,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;"
    strSQL = strSQL & "0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,54,1;0,165,20,1;0,165,20,1;0,165,20,1;"
    strSQL = strSQL & "0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;0,165,20,1;92,214,0,1;92,214,0,1;"
    strSQL = strSQL & "92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;92,214,0,1;183,214,0,1;"
    strSQL = strSQL & "183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;183,214,0,1;"
    strSQL = strSQL & "183,214,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;219,198,0,1;"
    strSQL = strSQL & "219,198,0,1;219,198,0,1;219,198,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;"
    strSQL = strSQL & "255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;"
    strSQL = strSQL & "255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,138,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;"
    strSQL = strSQL & "255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,93,0,1;255,49,3,1;255,49,3,1;255,49,3,1;"
    strSQL = strSQL & "255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,49,3,1;255,5,5,1;255,5,5,1;"
    strSQL = strSQL & "255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,5,5,1;255,89,107,1;"
    strSQL = strSQL & "255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;255,89,107,1;"
    strSQL = strSQL & "255,89,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;255,172,209,1;"
    strSQL = strSQL & "255,172,209,1;255,172,209,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (5,'8color','0,0,0,1;0,0,0,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;"
    strSQL = strSQL & "0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;"
    strSQL = strSQL & "0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,0,84,1;0,168,255,1;0,168,255,1;"
    strSQL = strSQL & "0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;"
    strSQL = strSQL & "0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;"
    strSQL = strSQL & "0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,168,255,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;"
    strSQL = strSQL & "255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;"
    strSQL = strSQL & "255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;255,220,0,1;"
    strSQL = strSQL & "255,220,0,1;255,220,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;"
    strSQL = strSQL & "255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;"
    strSQL = strSQL & "255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;255,128,0,1;"
    strSQL = strSQL & "255,128,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,190,190,1;254,190,190,1;253,190,190,1;252,190,190,1;"
    strSQL = strSQL & "251,190,190,1;250,190,190,1;249,190,190,1;248,190,190,1;247,190,190,1;246,190,190,1;245,190,190,1;244,190,190,1;243,190,190,1;"
    strSQL = strSQL & "242,190,190,1;241,190,190,1;240,190,190,1;238,190,190,1;237,190,190,1;236,190,190,1;235,190,190,1;234,190,190,1;233,190,190,1;"
    strSQL = strSQL & "232,190,190,1;231,190,190,1;230,190,190,1;229,190,190,1;228,190,190,1;227,190,190,1;226,190,190,1;225,190,190,1;224,190,190,1;"
    strSQL = strSQL & "220,190,190,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (6,'blue','0,0,0,1;0,0,9,1;0,0,19,1;0,0,29,1;0,1,38,1;0,1,47,1;"
    strSQL = strSQL & "0,2,55,1;0,3,63,1;0,4,71,1;0,5,79,1;0,6,87,1;0,7,94,1;0,8,100,1;0,9,107,1;0,11,113,1;0,12,118,1;0,13,124,1;0,15,128,1;0,16,133,1;"
    strSQL = strSQL & "0,17,137,1;0,19,141,1;0,20,144,1;0,21,147,1;0,22,149,1;0,24,151,1;0,25,153,1;0,26,154,1;0,27,155,1;0,28,156,1;0,29,157,1;0,30,158,1;"
    strSQL = strSQL & "0,31,159,1;0,32,160,1;0,33,161,1;0,34,162,1;0,35,163,1;0,36,164,1;0,37,165,1;0,38,166,1;0,39,167,1;0,40,168,1;0,41,169,1;0,42,170,1;"
    strSQL = strSQL & "0,43,171,1;0,44,172,1;0,45,173,1;0,46,174,1;0,47,175,1;0,48,176,1;0,49,177,1;0,50,178,1;0,51,179,1;0,52,180,1;0,53,181,1;0,54,182,1;"
    strSQL = strSQL & "0,55,183,1;0,56,184,1;0,57,185,1;0,58,186,1;0,59,187,1;0,60,188,1;0,61,189,1;0,62,190,1;0,63,191,1;0,64,192,1;0,65,193,1;0,66,194,1;"
    strSQL = strSQL & "0,67,195,1;0,68,196,1;0,69,197,1;0,70,198,1;0,71,199,1;0,72,200,1;0,73,201,1;0,74,202,1;0,75,203,1;0,76,204,1;0,77,205,1;0,78,206,1;"
    strSQL = strSQL & "0,79,207,1;0,80,208,1;0,81,209,1;0,82,210,1;0,83,211,1;0,84,212,1;0,85,213,1;0,86,214,1;0,87,215,1;0,88,216,1;0,89,217,1;0,90,218,1;"
    strSQL = strSQL & "0,91,219,1;0,92,220,1;0,93,221,1;0,94,222,1;0,95,223,1;0,96,224,1;0,97,225,1;0,98,226,1;0,99,227,1;0,100,228,1;0,101,229,1;"
    strSQL = strSQL & "0,102,230,1;0,103,231,1;0,104,232,1;0,105,233,1;0,106,234,1;0,107,235,1;0,108,236,1;0,109,237,1;0,110,238,1;0,111,239,1;0,112,240,1;"
    strSQL = strSQL & "0,113,241,1;0,114,242,1;0,115,243,1;0,116,244,1;0,117,245,1;0,118,246,1;0,119,247,1;0,120,248,1;0,121,249,1;0,122,250,1;0,123,251,1;"
    strSQL = strSQL & "0,124,252,1;0,125,253,1;0,126,254,1;0,127,255,1;0,128,255,1;0,129,255,1;0,130,255,1;0,131,255,1;0,132,255,1;0,133,255,1;0,134,255,1;"
    strSQL = strSQL & "0,135,255,1;0,136,255,1;0,137,255,1;0,138,255,1;0,139,255,1;0,140,255,1;0,141,255,1;0,142,255,1;0,143,255,1;0,144,255,1;0,145,255,1;"
    strSQL = strSQL & "0,146,255,1;0,147,255,1;0,148,255,1;0,149,255,1;0,150,255,1;0,151,255,1;0,152,255,1;0,153,255,1;0,154,255,1;0,155,255,1;0,156,255,1;"
    strSQL = strSQL & "0,157,255,1;0,158,255,1;0,159,255,1;0,160,255,1;0,161,255,1;0,162,255,1;0,163,255,1;0,164,255,1;0,165,255,1;0,166,255,1;0,167,255,1;"
    strSQL = strSQL & "0,168,255,1;0,169,255,1;0,170,255,1;0,171,255,1;0,172,255,1;0,173,255,1;0,174,255,1;0,175,255,1;0,176,255,1;0,177,255,1;0,178,255,1;"
    strSQL = strSQL & "0,179,255,1;0,180,255,1;0,181,255,1;0,182,255,1;0,183,255,1;0,184,255,1;0,185,255,1;0,186,255,1;0,187,255,1;0,188,255,1;0,189,255,1;"
    strSQL = strSQL & "0,190,255,1;0,191,255,1;0,192,255,1;0,193,255,1;0,194,255,1;0,195,255,1;0,196,255,1;0,197,255,1;0,198,255,1;0,199,255,1;0,200,255,1;"
    strSQL = strSQL & "0,201,255,1;0,202,255,1;0,203,255,1;0,204,255,1;0,205,255,1;0,206,255,1;0,207,255,1;0,208,255,1;0,209,255,1;0,210,255,1;0,211,255,1;"
    strSQL = strSQL & "0,212,255,1;0,213,255,1;0,214,255,1;0,215,255,1;0,216,255,1;0,217,255,1;0,218,255,1;0,219,255,1;0,220,255,1;0,221,255,1;0,222,255,1;"
    strSQL = strSQL & "0,223,255,1;0,224,255,1;0,225,255,1;0,226,255,1;0,227,255,1;0,228,255,1;0,229,255,1;0,230,255,1;0,231,255,1;2,233,255,1;4,234,255,1;"
    strSQL = strSQL & "7,235,255,1;11,236,255,1;16,238,255,1;22,239,255,1;28,240,255,1;35,242,255,1;43,243,255,1;52,244,255,1;61,246,255,1;71,247,255,1;"
    strSQL = strSQL & "82,248,255,1;94,249,255,1;107,250,255,1;120,251,255,1;134,252,255,1;149,253,255,1;165,254,255,1;181,254,255,1;199,255,255,1;"
    strSQL = strSQL & "217,255,255,1;236,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (7,'Blue_Green_Red','0,0,0,1;2,0,4,1;4,0,8,1;7,0,12,1;9,0,15,1;"
    strSQL = strSQL & "11,0,19,1;13,0,23,1;15,0,27,1;18,0,31,1;20,0,35,1;22,0,38,1;24,0,42,1;26,0,46,1;29,0,50,1;31,0,54,1;33,0,58,1;35,0,61,1;37,0,65,1;"
    strSQL = strSQL & "39,0,69,1;42,0,73,1;44,0,77,1;46,0,81,1;48,0,84,1;50,0,88,1;53,0,92,1;55,0,96,1;57,0,100,1;59,0,104,1;61,0,107,1;64,0,111,1;"
    strSQL = strSQL & "66,0,115,1;68,0,119,1;66,0,123,1;64,0,128,1;62,0,132,1;60,0,136,1;57,0,140,1;55,0,145,1;53,0,149,1;51,0,153,1;49,0,157,1;"
    strSQL = strSQL & "47,0,162,1;45,0,166,1;43,0,170,1;40,0,174,1;38,0,179,1;36,0,183,1;34,0,187,1;32,0,191,1;30,0,196,1;28,0,200,1;26,0,204,1;"
    strSQL = strSQL & "23,0,208,1;21,0,213,1;19,0,217,1;17,0,221,1;15,0,225,1;13,0,230,1;11,0,234,1;9,0,238,1;6,0,242,1;4,0,247,1;2,0,251,1;0,0,255,1;"
    strSQL = strSQL & "0,5,249,1;0,9,244,1;0,14,238,1;0,18,232,1;0,23,227,1;0,27,221,1;0,32,215,1;0,36,210,1;0,41,204,1;0,45,198,1;0,50,193,1;0,54,187,1;"
    strSQL = strSQL & "0,59,181,1;0,63,176,1;0,68,170,1;0,73,164,1;0,77,159,1;0,82,153,1;0,86,147,1;0,91,142,1;0,95,136,1;0,100,130,1;0,104,125,1;0,109,119,1;"
    strSQL = strSQL & "0,113,113,1;0,118,108,1;0,122,102,1;0,127,96,1;0,131,91,1;0,136,85,1;0,141,79,1;0,145,74,1;0,150,68,1;0,154,62,1;0,159,57,1;0,163,51,1;"
    strSQL = strSQL & "0,168,45,1;0,172,40,1;0,177,34,1;0,181,28,1;0,186,23,1;0,190,17,1;0,195,11,1;0,199,6,1;0,204,0,1;9,206,0,1;18,208,0,1;26,209,0,1;"
    strSQL = strSQL & "35,211,0,1;44,213,0,1;53,215,0,1;62,216,0,1;70,218,0,1;79,220,0,1;88,222,0,1;97,223,0,1;106,225,0,1;114,227,0,1;123,229,0,1;"
    strSQL = strSQL & "132,230,0,1;141,232,0,1;149,234,0,1;158,236,0,1;167,237,0,1;176,239,0,1;185,241,0,1;193,243,0,1;202,244,0,1;211,246,0,1;220,248,0,1;"
    strSQL = strSQL & "229,250,0,1;237,251,0,1;246,253,0,1;255,255,0,1;255,253,0,1;255,251,0,1;255,249,0,1;255,246,0,1;255,244,0,1;255,242,0,1;255,240,0,1;"
    strSQL = strSQL & "255,238,0,1;255,236,0,1;255,233,0,1;255,231,0,1;255,229,0,1;255,227,0,1;255,225,0,1;255,223,0,1;255,220,0,1;255,218,0,1;255,216,0,1;"
    strSQL = strSQL & "255,214,0,1;255,212,0,1;255,210,0,1;255,207,0,1;255,205,0,1;255,203,0,1;255,201,0,1;255,199,0,1;255,197,0,1;255,194,0,1;255,192,0,1;"
    strSQL = strSQL & "255,190,0,1;255,188,0,1;255,186,0,1;255,184,0,1;255,182,0,1;255,179,0,1;255,177,0,1;255,175,0,1;255,173,0,1;255,171,0,1;255,169,0,1;"
    strSQL = strSQL & "255,166,0,1;255,164,0,1;255,162,0,1;255,160,0,1;255,158,0,1;255,156,0,1;255,153,0,1;255,151,0,1;255,149,0,1;255,147,0,1;255,145,0,1;"
    strSQL = strSQL & "255,143,0,1;255,140,0,1;255,138,0,1;255,136,0,1;255,134,0,1;255,132,0,1;255,130,0,1;255,127,0,1;255,125,0,1;255,123,0,1;255,121,0,1;"
    strSQL = strSQL & "255,119,0,1;255,117,0,1;255,115,0,1;255,112,0,1;255,110,0,1;255,108,0,1;255,106,0,1;255,104,0,1;255,102,0,1;255,99,0,1;255,97,0,1;"
    strSQL = strSQL & "255,95,0,1;255,93,0,1;255,91,0,1;255,89,0,1;255,86,0,1;255,84,0,1;255,82,0,1;255,80,0,1;255,78,0,1;255,76,0,1;255,73,0,1;255,71,0,1;"
    strSQL = strSQL & "255,69,0,1;255,67,0,1;255,65,0,1;255,63,0,1;255,61,0,1;255,58,0,1;255,56,0,1;255,54,0,1;255,52,0,1;255,50,0,1;255,48,0,1;255,45,0,1;"
    strSQL = strSQL & "255,43,0,1;255,41,0,1;255,39,0,1;255,37,0,1;255,35,0,1;255,32,0,1;255,30,0,1;255,28,0,1;255,26,0,1;255,24,0,1;255,22,0,1;255,19,0,1;"
    strSQL = strSQL & "255,17,0,1;255,15,0,1;255,13,0,1;255,11,0,1;255,9,0,1;255,6,0,1;255,4,0,1;255,2,0,1;255,0,0,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (8,'blue_orange','0,0,0,1;1,2,2,1;2,4,4,1;3,5,6,1;4,7,7,1;5,9,9,1;"
    strSQL = strSQL & "6,11,11,1;7,13,13,1;8,15,15,1;9,16,17,1;11,18,19,1;12,20,21,1;13,22,22,1;14,24,24,1;15,25,26,1;16,27,28,1;17,29,30,1;18,31,32,1;"
    strSQL = strSQL & "19,33,34,1;20,35,36,1;21,36,37,1;22,38,39,1;23,40,41,1;24,42,43,1;25,44,45,1;26,46,47,1;27,47,49,1;28,49,51,1;29,51,52,1;30,53,54,1;"
    strSQL = strSQL & "32,55,56,1;33,56,58,1;34,58,60,1;35,60,62,1;36,62,64,1;37,64,66,1;38,66,67,1;39,67,69,1;40,69,71,1;41,71,73,1;42,70,74,1;43,68,76,1;"
    strSQL = strSQL & "43,67,77,1;44,66,78,1;45,65,79,1;46,63,81,1;46,62,82,1;47,61,83,1;48,60,84,1;49,58,86,1;49,57,87,1;50,56,88,1;51,55,89,1;52,53,91,1;"
    strSQL = strSQL & "52,52,92,1;53,51,93,1;54,49,95,1;55,48,96,1;55,47,97,1;56,46,98,1;57,44,100,1;58,43,101,1;58,42,102,1;59,41,103,1;60,39,105,1;"
    strSQL = strSQL & "61,38,106,1;61,37,107,1;62,35,108,1;63,34,110,1;64,33,111,1;64,32,112,1;65,30,114,1;66,29,115,1;67,28,116,1;67,27,117,1;"
    strSQL = strSQL & "68,25,119,1;69,24,120,1;70,23,121,1;70,22,122,1;71,20,124,1;72,19,125,1;73,18,126,1;73,16,128,1;74,15,129,1;75,14,130,1;76,13,131,1;"
    strSQL = strSQL & "76,11,133,1;77,10,134,1;78,9,135,1;79,8,136,1;79,6,138,1;80,5,139,1;81,4,140,1;82,3,141,1;82,1,143,1;83,0,144,1;85,2,144,1;87,3,143,1;"
    strSQL = strSQL & "89,5,143,1;91,7,142,1;93,9,142,1;95,10,141,1;97,12,141,1;99,14,140,1;102,16,140,1;104,17,140,1;106,19,139,1;108,21,139,1;110,22,138,1;"
    strSQL = strSQL & "112,24,138,1;114,26,137,1;116,28,137,1;118,29,136,1;120,31,136,1;122,33,136,1;124,35,135,1;126,36,135,1;128,38,134,1;130,40,134,1;"
    strSQL = strSQL & "132,42,133,1;134,43,133,1;137,45,133,1;139,47,132,1;141,48,132,1;143,50,131,1;145,52,131,1;147,54,130,1;149,55,130,1;151,57,129,1;"
    strSQL = strSQL & "153,59,129,1;155,61,129,1;157,62,128,1;159,64,128,1;161,66,127,1;163,68,127,1;165,69,126,1;167,71,126,1;169,73,125,1;171,74,125,1;"
    strSQL = strSQL & "174,76,125,1;176,78,124,1;178,80,124,1;180,81,123,1;182,83,123,1;184,85,122,1;186,87,122,1;188,88,121,1;190,90,121,1;191,90,119,1;"
    strSQL = strSQL & "192,91,117,1;193,91,115,1;194,92,113,1;195,92,111,1;197,92,109,1;198,93,107,1;199,93,105,1;200,94,103,1;201,94,101,1;202,94,99,1;"
    strSQL = strSQL & "203,95,97,1;204,95,95,1;205,96,93,1;206,96,91,1;207,96,89,1;208,97,87,1;210,97,85,1;211,98,83,1;212,98,81,1;213,98,79,1;214,99,77,1;"
    strSQL = strSQL & "215,99,75,1;216,100,73,1;217,100,71,1;218,100,69,1;219,101,67,1;220,101,65,1;221,102,63,1;223,102,61,1;224,102,58,1;225,103,56,1;"
    strSQL = strSQL & "226,103,54,1;227,104,52,1;228,104,50,1;229,104,48,1;230,105,46,1;231,105,44,1;232,106,42,1;233,106,40,1;234,106,38,1;236,107,36,1;"
    strSQL = strSQL & "237,107,34,1;238,108,32,1;239,108,30,1;240,108,28,1;241,109,26,1;242,109,24,1;243,110,22,1;244,110,20,1;245,110,18,1;246,111,16,1;"
    strSQL = strSQL & "247,111,14,1;249,112,12,1;250,112,10,1;251,112,8,1;252,113,6,1;253,113,4,1;254,114,2,1;255,114,0,1;255,117,5,1;255,120,11,1;"
    strSQL = strSQL & "255,123,16,1;255,126,21,1;255,129,27,1;255,132,32,1;255,135,37,1;255,138,43,1;255,140,48,1;255,143,53,1;255,146,58,1;255,149,64,1;"
    strSQL = strSQL & "255,152,69,1;255,155,74,1;255,158,80,1;255,161,85,1;255,164,90,1;255,167,96,1;255,170,101,1;255,173,106,1;255,176,112,1;"
    strSQL = strSQL & "255,179,117,1;255,182,122,1;255,185,128,1;255,187,133,1;255,190,138,1;255,193,143,1;255,196,149,1;255,199,154,1;255,202,159,1;"
    strSQL = strSQL & "255,205,165,1;255,208,170,1;255,211,175,1;255,214,181,1;255,217,186,1;255,220,191,1;255,223,197,1;255,226,202,1;255,229,207,1;"
    strSQL = strSQL & "255,232,213,1;255,234,218,1;255,237,223,1;255,240,228,1;255,243,234,1;255,246,239,1;255,249,244,1;255,252,250,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (9,'blue_orange_icb','0,0,0,1;0,1,0,1;0,2,0,1;0,3,0,1;0,4,0,1;0,5,0,1;"
    strSQL = strSQL & "0,6,0,1;0,7,0,1;0,8,0,1;0,9,0,1;0,10,0,1;0,11,0,1;0,12,0,1;0,13,0,1;0,14,0,1;0,15,0,1;0,16,0,1;0,17,0,1;0,18,0,1;0,19,0,1;0,20,0,1;"
    strSQL = strSQL & "0,21,0,1;0,22,0,1;0,23,0,1;0,24,0,1;0,25,0,1;0,26,0,1;0,27,0,1;0,28,0,1;0,29,0,1;0,30,0,1;0,31,0,1;0,32,0,1;0,33,0,1;0,34,0,1;"
    strSQL = strSQL & "0,35,0,1;0,36,0,1;0,37,0,1;0,38,0,1;0,39,0,1;0,40,0,1;0,41,0,1;0,42,0,1;0,43,0,1;0,44,0,1;0,45,0,1;0,46,0,1;0,47,0,1;0,48,0,1;"
    strSQL = strSQL & "0,49,0,1;0,50,0,1;0,51,0,1;0,52,0,1;0,53,0,1;0,54,0,1;0,55,0,1;0,56,0,1;0,57,0,1;0,58,0,1;0,59,0,1;0,60,0,1;0,61,0,1;0,62,0,1;"
    strSQL = strSQL & "0,63,0,1;1,63,1,1;3,62,3,1;5,61,5,1;7,60,7,1;9,59,9,1;11,58,11,1;13,57,13,1;15,56,15,1;17,55,17,1;19,54,19,1;21,53,21,1;23,52,23,1;"
    strSQL = strSQL & "25,51,25,1;27,50,27,1;29,49,29,1;31,48,31,1;33,47,33,1;35,46,35,1;37,45,37,1;39,44,39,1;41,43,41,1;43,42,43,1;45,41,45,1;47,40,47,1;"
    strSQL = strSQL & "49,39,49,1;51,38,51,1;53,37,53,1;55,36,55,1;57,35,57,1;59,34,59,1;61,33,61,1;63,32,63,1;65,31,65,1;67,30,67,1;69,29,69,1;71,28,71,1;"
    strSQL = strSQL & "73,27,73,1;75,26,75,1;77,25,77,1;79,24,79,1;81,23,81,1;83,22,83,1;85,21,85,1;87,20,87,1;89,19,89,1;91,18,91,1;93,17,93,1;95,16,95,1;"
    strSQL = strSQL & "97,15,97,1;99,14,99,1;101,13,101,1;103,12,103,1;105,11,105,1;107,10,107,1;109,9,109,1;111,8,111,1;113,7,113,1;115,6,115,1;117,5,117,1;"
    strSQL = strSQL & "119,4,119,1;121,3,121,1;123,2,123,1;125,1,125,1;127,0,127,1;129,0,125,1;131,2,123,1;133,4,121,1;135,6,119,1;137,8,117,1;139,10,115,1;"
    strSQL = strSQL & "141,12,113,1;143,14,111,1;145,16,109,1;147,18,107,1;149,20,105,1;151,22,103,1;153,24,101,1;155,26,99,1;157,28,97,1;159,30,95,1;"
    strSQL = strSQL & "161,32,93,1;163,34,91,1;165,36,89,1;167,38,87,1;169,40,85,1;171,42,83,1;173,44,81,1;175,46,79,1;177,48,77,1;179,50,75,1;181,52,73,1;"
    strSQL = strSQL & "183,54,71,1;185,56,69,1;187,58,67,1;189,60,65,1;191,62,63,1;193,64,61,1;195,66,59,1;197,68,57,1;199,70,55,1;201,72,53,1;203,74,51,1;"
    strSQL = strSQL & "205,76,49,1;207,78,47,1;209,80,45,1;211,82,43,1;213,84,41,1;215,86,39,1;217,88,37,1;219,90,35,1;221,92,33,1;223,94,31,1;225,96,29,1;"
    strSQL = strSQL & "227,98,27,1;229,100,25,1;231,102,23,1;233,104,21,1;235,106,19,1;237,108,17,1;239,110,15,1;241,112,13,1;243,114,11,1;245,116,9,1;"
    strSQL = strSQL & "247,118,7,1;249,120,5,1;251,122,3,1;253,124,1,1;255,126,0,1;255,128,3,1;255,130,7,1;255,132,11,1;255,134,15,1;255,136,19,1;"
    strSQL = strSQL & "255,138,23,1;255,140,27,1;255,142,31,1;255,144,35,1;255,146,39,1;255,148,43,1;255,150,47,1;255,152,51,1;255,154,55,1;255,156,59,1;"
    strSQL = strSQL & "255,158,63,1;255,160,67,1;255,162,71,1;255,164,75,1;255,166,79,1;255,168,83,1;255,170,87,1;255,172,91,1;255,174,95,1;255,176,99,1;"
    strSQL = strSQL & "255,178,103,1;255,180,107,1;255,182,111,1;255,184,115,1;255,186,119,1;255,188,123,1;255,190,127,1;255,192,131,1;255,194,135,1;"
    strSQL = strSQL & "255,196,139,1;255,198,143,1;255,200,147,1;255,202,151,1;255,204,155,1;255,206,159,1;255,208,163,1;255,210,167,1;255,212,171,1;"
    strSQL = strSQL & "255,214,175,1;255,216,179,1;255,218,183,1;255,220,187,1;255,222,191,1;255,224,195,1;255,226,199,1;255,228,203,1;255,230,207,1;"
    strSQL = strSQL & "255,232,211,1;255,234,215,1;255,236,219,1;255,238,223,1;255,240,227,1;255,242,231,1;255,244,235,1;255,246,239,1;255,248,243,1;"
    strSQL = strSQL & "255,250,247,1;255,252,251,1;255,254,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (10,'Blue_Red_Yellow','0,2,0,1;0,2,24,1;0,0,72,1;0,0,75,1;0,0,78,1;"
    strSQL = strSQL & "0,0,82,1;0,0,85,1;0,0,88,1;0,0,92,1;0,0,95,1;0,0,98,1;0,0,102,1;0,0,108,1;0,0,115,1;0,0,122,1;0,0,130,1;0,0,144,1;0,0,151,1;0,0,158,1;"
    strSQL = strSQL & "0,0,166,1;0,0,180,1;0,0,187,1;0,0,194,1;0,0,202,1;0,0,216,1;0,0,223,1;0,0,230,1;0,0,238,1;0,0,252,1;5,0,250,1;10,0,248,1;15,0,246,1;"
    strSQL = strSQL & "24,0,240,1;29,0,235,1;34,0,230,1;40,0,225,1;45,0,220,1;51,0,214,1;56,0,209,1;62,0,204,1;72,0,192,1;77,0,188,1;82,0,183,1;87,0,179,1;"
    strSQL = strSQL & "96,0,168,1;101,0,164,1;106,0,159,1;111,0,155,1;120,0,144,1;125,0,140,1;130,0,135,1;135,0,131,1;144,0,120,1;149,0,116,1;154,0,111,1;"
    strSQL = strSQL & "159,0,107,1;168,0,96,1;177,0,92,1;187,0,87,1;197,0,83,1;216,0,72,1;216,0,68,1;216,0,63,1;216,0,59,1;216,0,48,1;221,0,44,1;226,0,39,1;"
    strSQL = strSQL & "231,0,35,1;240,0,24,1;242,0,20,1;245,0,15,1;247,0,11,1;252,0,0,1;252,0,0,1;252,0,0,1;252,1,0,1;252,1,0,1;252,2,0,1;251,4,0,1;251,6,0,1;"
    strSQL = strSQL & "248,9,0,1;248,10,0,1;248,12,0,1;248,13,0,1;248,16,0,1;248,18,0,1;247,20,0,1;247,22,0,1;244,25,0,1;244,27,0,1;243,29,0,1;243,31,0,1;"
    strSQL = strSQL & "240,34,0,1;240,36,0,1;239,38,0,1;239,40,0,1;236,43,0,1;236,44,0,1;236,46,0,1;236,47,0,1;236,50,0,1;236,52,0,1;235,54,0,1;235,56,0,1;"
    strSQL = strSQL & "232,59,0,1;232,60,0,1;231,62,0,1;231,64,0,1;228,67,0,1;228,68,0,1;228,70,0,1;228,71,0,1;228,74,0,1;228,76,0,1;227,78,0,1;227,80,0,1;"
    strSQL = strSQL & "224,83,0,1;224,85,0,1;223,87,0,1;223,89,0,1;220,92,0,1;220,94,0,1;219,96,0,1;219,98,0,1;216,101,0,1;216,102,0,1;216,104,0,1;"
    strSQL = strSQL & "216,105,0,1;216,108,0,1;216,109,0,1;215,111,0,1;215,113,0,1;212,116,0,1;212,118,0,1;211,120,0,1;211,122,0,1;208,125,0,1;208,127,0,1;"
    strSQL = strSQL & "207,129,0,1;207,131,0,1;204,134,0,1;204,135,0,1;204,137,0,1;204,138,0,1;204,141,0,1;204,143,0,1;203,145,0,1;203,147,0,1;200,150,0,1;"
    strSQL = strSQL & "200,152,0,1;199,154,0,1;199,156,0,1;196,159,0,1;196,160,0,1;196,161,0,1;196,163,0,1;196,165,0,1;196,167,0,1;195,169,0,1;195,171,0,1;"
    strSQL = strSQL & "192,174,0,1;192,176,0,1;191,178,0,1;191,180,0,1;188,183,0,1;188,185,0,1;187,187,0,1;187,189,0,1;184,192,0,1;184,193,0,1;184,195,0,1;"
    strSQL = strSQL & "184,196,0,1;184,199,0,1;184,201,0,1;183,203,0,1;183,205,0,1;180,208,0,1;180,209,0,1;179,211,0,1;179,213,0,1;176,216,0,1;176,216,0,1;"
    strSQL = strSQL & "176,216,0,1;176,216,0,1;176,216,0,1;176,219,0,1;175,222,0,1;175,226,0,1;172,232,0,1;172,234,0,1;171,236,0,1;171,238,0,1;168,241,0,1;"
    strSQL = strSQL & "168,243,0,1;167,245,0,1;167,247,0,1;164,250,0,1;164,251,0,1;164,252,0,1;164,253,0,1;164,255,0,1;164,255,0,1;163,255,0,1;163,255,0,1;"
    strSQL = strSQL & "160,255,0,1;160,255,0,1;159,255,0,1;159,255,0,1;156,255,0,1;156,255,0,1;155,255,0,1;155,255,0,1;152,255,0,1;152,255,0,1;152,255,0,1;"
    strSQL = strSQL & "152,255,0,1;152,255,0,1;152,255,0,1;151,255,0,1;151,255,0,1;148,255,0,1;150,255,6,1;153,255,13,1;155,255,19,1;160,255,32,1;"
    strSQL = strSQL & "162,255,38,1;165,255,45,1;167,255,51,1;172,255,64,1;175,255,70,1;178,255,77,1;182,255,83,1;188,255,96,1;190,255,102,1;193,255,109,1;"
    strSQL = strSQL & "195,255,115,1;200,255,128,1;202,255,133,1;205,255,139,1;207,255,145,1;212,255,156,1;214,255,162,1;217,255,169,1;219,255,175,1;"
    strSQL = strSQL & "224,255,188,1;227,255,194,1;230,255,201,1;234,255,207,1;240,255,220,1;242,255,226,1;245,255,233,1;247,255,239,1;252,255,252,1;"
    strSQL = strSQL & "253,255,253,1;254,255,254,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (11,'Blue-Orange','0,0,0,1;0,1,0,1;0,2,0,1;0,3,0,1;0,4,0,1;0,5,0,1;"
    strSQL = strSQL & "0,6,0,1;0,7,0,1;0,8,0,1;0,9,0,1;0,10,0,1;0,11,0,1;0,12,0,1;0,13,0,1;0,14,0,1;0,15,0,1;0,16,0,1;0,17,0,1;0,18,0,1;0,19,0,1;0,20,0,1;"
    strSQL = strSQL & "0,21,0,1;0,22,0,1;0,23,0,1;0,24,0,1;0,25,0,1;0,26,0,1;0,27,0,1;0,28,0,1;0,29,0,1;0,30,0,1;0,31,0,1;0,32,0,1;0,33,0,1;0,34,0,1;"
    strSQL = strSQL & "0,35,0,1;0,36,0,1;0,37,0,1;0,38,0,1;0,39,0,1;0,40,0,1;0,41,0,1;0,42,0,1;0,43,0,1;0,44,0,1;0,45,0,1;0,46,0,1;0,47,0,1;0,48,0,1;"
    strSQL = strSQL & "0,49,0,1;0,50,0,1;0,51,0,1;0,52,0,1;0,53,0,1;0,54,0,1;0,55,0,1;0,56,0,1;0,57,0,1;0,58,0,1;0,59,0,1;0,60,0,1;0,61,0,1;0,62,0,1;"
    strSQL = strSQL & "0,63,0,1;1,63,1,1;3,62,3,1;5,61,5,1;7,60,7,1;9,59,9,1;11,58,11,1;13,57,13,1;15,56,15,1;17,55,17,1;19,54,19,1;21,53,21,1;23,52,23,1;"
    strSQL = strSQL & "25,51,25,1;27,50,27,1;29,49,29,1;31,48,31,1;33,47,33,1;35,46,35,1;37,45,37,1;39,44,39,1;41,43,41,1;43,42,43,1;45,41,45,1;47,40,47,1;"
    strSQL = strSQL & "49,39,49,1;51,38,51,1;53,37,53,1;55,36,55,1;57,35,57,1;59,34,59,1;61,33,61,1;63,32,63,1;65,31,65,1;67,30,67,1;69,29,69,1;71,28,71,1;"
    strSQL = strSQL & "73,27,73,1;75,26,75,1;77,25,77,1;79,24,79,1;81,23,81,1;83,22,83,1;85,21,85,1;87,20,87,1;89,19,89,1;91,18,91,1;93,17,93,1;95,16,95,1;"
    strSQL = strSQL & "97,15,97,1;99,14,99,1;101,13,101,1;103,12,103,1;105,11,105,1;107,10,107,1;109,9,109,1;111,8,111,1;113,7,113,1;115,6,115,1;117,5,117,1;"
    strSQL = strSQL & "119,4,119,1;121,3,121,1;123,2,123,1;125,1,125,1;127,0,127,1;129,0,125,1;131,2,123,1;133,4,121,1;135,6,119,1;137,8,117,1;139,10,115,1;"
    strSQL = strSQL & "141,12,113,1;143,14,111,1;145,16,109,1;147,18,107,1;149,20,105,1;151,22,103,1;153,24,101,1;155,26,99,1;157,28,97,1;159,30,95,1;"
    strSQL = strSQL & "161,32,93,1;163,34,91,1;165,36,89,1;167,38,87,1;169,40,85,1;171,42,83,1;173,44,81,1;175,46,79,1;177,48,77,1;179,50,75,1;181,52,73,1;"
    strSQL = strSQL & "183,54,71,1;185,56,69,1;187,58,67,1;189,60,65,1;191,62,63,1;193,64,61,1;195,66,59,1;197,68,57,1;199,70,55,1;201,72,53,1;203,74,51,1;"
    strSQL = strSQL & "205,76,49,1;207,78,47,1;209,80,45,1;211,82,43,1;213,84,41,1;215,86,39,1;217,88,37,1;219,90,35,1;221,92,33,1;223,94,31,1;225,96,29,1;"
    strSQL = strSQL & "227,98,27,1;229,100,25,1;231,102,23,1;233,104,21,1;235,106,19,1;237,108,17,1;239,110,15,1;241,112,13,1;243,114,11,1;245,116,9,1;"
    strSQL = strSQL & "247,118,7,1;249,120,5,1;251,122,3,1;253,124,1,1;255,126,0,1;255,128,3,1;255,130,7,1;255,132,11,1;255,134,15,1;255,136,19,1;"
    strSQL = strSQL & "255,138,23,1;255,140,27,1;255,142,31,1;255,144,35,1;255,146,39,1;255,148,43,1;255,150,47,1;255,152,51,1;255,154,55,1;255,156,59,1;"
    strSQL = strSQL & "255,158,63,1;255,160,67,1;255,162,71,1;255,164,75,1;255,166,79,1;255,168,83,1;255,170,87,1;255,172,91,1;255,174,95,1;255,176,99,1;"
    strSQL = strSQL & "255,178,103,1;255,180,107,1;255,182,111,1;255,184,115,1;255,186,119,1;255,188,123,1;255,190,127,1;255,192,131,1;255,194,135,1;"
    strSQL = strSQL & "255,196,139,1;255,198,143,1;255,200,147,1;255,202,151,1;255,204,155,1;255,206,159,1;255,208,163,1;255,210,167,1;255,212,171,1;"
    strSQL = strSQL & "255,214,175,1;255,216,179,1;255,218,183,1;255,220,187,1;255,222,191,1;255,224,195,1;255,226,199,1;255,228,203,1;255,230,207,1;"
    strSQL = strSQL & "255,232,211,1;255,234,215,1;255,236,219,1;255,238,223,1;255,240,227,1;255,242,231,1;255,244,235,1;255,246,239,1;255,248,243,1;"
    strSQL = strSQL & "255,250,247,1;255,252,251,1;255,254,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (12,'Bullseye','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;"
    strSQL = strSQL & "160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;"
    strSQL = strSQL & "160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;"
    strSQL = strSQL & "160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;"
    strSQL = strSQL & "160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;160,32,240,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,222,173,1;"
    strSQL = strSQL & "255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;"
    strSQL = strSQL & "255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;"
    strSQL = strSQL & "255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;"
    strSQL = strSQL & "255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;"
    strSQL = strSQL & "255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;255,222,173,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (13,'CEqual','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,6,7,1;0,13,14,1;0,19,21,1;0,26,28,1;0,33,35,1;"
    strSQL = strSQL & "0,39,42,1;0,46,49,1;0,53,57,1;0,59,64,1;0,66,71,1;0,72,78,1;0,79,85,1;0,86,92,1;0,92,99,1;0,99,106,1;0,106,114,1;0,106,116,1;"
    strSQL = strSQL & "0,107,119,1;0,107,122,1;0,108,124,1;0,108,127,1;0,109,130,1;0,109,133,1;0,110,135,1;0,111,138,1;0,111,141,1;0,112,144,1;0,112,146,1;"
    strSQL = strSQL & "0,113,149,1;0,113,152,1;0,114,155,1;0,115,158,1;2,114,160,1;5,114,163,1;8,114,166,1;11,114,169,1;14,113,172,1;17,113,174,1;"
    strSQL = strSQL & "20,113,177,1;23,113,180,1;26,112,183,1;29,112,186,1;32,112,188,1;35,112,191,1;38,111,194,1;41,111,197,1;44,111,200,1;47,111,207,1;"
    strSQL = strSQL & "51,111,210,1;55,111,213,1;59,111,216,1;63,111,220,1;67,111,223,1;71,111,226,1;75,111,229,1;79,111,233,1;83,111,236,1;87,111,239,1;"
    strSQL = strSQL & "91,111,242,1;95,111,246,1;99,111,249,1;103,111,252,1;107,111,255,1;112,112,255,1;114,105,255,1;116,98,255,1;118,91,255,1;120,84,255,1;"
    strSQL = strSQL & "122,77,255,1;124,70,255,1;126,63,255,1;129,56,255,1;131,49,255,1;133,42,255,1;135,35,255,1;137,28,255,1;139,21,255,1;141,14,255,1;"
    strSQL = strSQL & "143,7,255,1;146,0,255,1;148,0,252,1;151,0,250,1;154,0,248,1;157,0,246,1;160,0,244,1;163,0,242,1;166,0,240,1;169,0,238,1;172,0,235,1;"
    strSQL = strSQL & "175,0,233,1;177,0,231,1;180,0,229,1;183,0,227,1;186,0,225,1;189,0,223,1;175,1,221,1;177,6,216,1;180,11,211,1;183,17,206,1;"
    strSQL = strSQL & "186,22,202,1;189,27,197,1;192,33,192,1;195,38,188,1;198,44,183,1;201,49,178,1;204,54,174,1;207,60,169,1;210,65,164,1;213,70,160,1;"
    strSQL = strSQL & "216,76,155,1;219,81,150,1;222,87,146,1;222,88,143,1;223,90,140,1;224,91,138,1;225,93,135,1;226,95,132,1;227,96,130,1;228,98,127,1;"
    strSQL = strSQL & "229,100,125,1;230,101,122,1;231,103,119,1;232,104,117,1;233,106,114,1;234,108,111,1;235,109,109,1;236,111,106,1;237,113,104,1;"
    strSQL = strSQL & "238,116,97,1;239,119,91,1;240,122,84,1;241,126,78,1;242,129,71,1;243,132,65,1;244,136,58,1;246,139,52,1;247,142,45,1;248,146,39,1;"
    strSQL = strSQL & "249,149,32,1;250,152,26,1;251,156,19,1;252,159,13,1;253,162,6,1;255,166,0,1;255,170,0,1;255,174,0,1;255,178,0,1;255,182,0,1;"
    strSQL = strSQL & "255,186,0,1;255,191,0,1;255,195,0,1;255,199,0,1;255,203,0,1;255,207,0,1;255,212,0,1;255,216,0,1;255,220,0,1;255,224,0,1;255,228,0,1;"
    strSQL = strSQL & "255,213,0,1;255,215,11,1;255,217,22,1;255,219,33,1;255,221,45,1;255,223,56,1;255,225,67,1;255,227,79,1;255,229,90,1;255,231,101,1;"
    strSQL = strSQL & "255,233,113,1;255,235,124,1;255,237,135,1;255,239,147,1;255,241,158,1;255,243,169,1;255,246,148,1;255,246,151,1;255,247,155,1;"
    strSQL = strSQL & "255,247,159,1;255,248,163,1;255,248,167,1;255,249,171,1;255,249,175,1;255,250,179,1;255,251,182,1;255,251,186,1;255,252,190,1;"
    strSQL = strSQL & "255,252,194,1;255,253,198,1;255,253,202,1;255,254,206,1;255,255,210,1;255,255,212,1;255,255,215,1;255,255,218,1;255,255,221,1;"
    strSQL = strSQL & "255,255,224,1;255,255,226,1;255,255,229,1;255,255,232,1;255,255,235,1;255,255,238,1;255,255,240,1;255,255,243,1;255,255,246,1;"
    strSQL = strSQL & "255,255,249,1;255,255,252,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (14,'colder','0,0,0,1;19,19,19,1;37,38,38,1;53,56,56,1;68,74,74,1;"
    strSQL = strSQL & "81,90,90,1;93,106,106,1;104,121,121,1;114,135,135,1;122,148,148,1;129,161,161,1;135,173,173,1;140,184,184,1;144,194,194,1;"
    strSQL = strSQL & "147,203,203,1;150,212,212,1;151,220,220,1;151,227,227,1;151,233,233,1;150,239,239,1;148,244,244,1;145,248,248,1;142,251,251,1;"
    strSQL = strSQL & "139,253,253,1;135,255,255,1;130,255,255,1;125,255,255,1;120,255,255,1;115,255,255,1;110,255,255,1;105,255,255,1;100,255,255,1;"
    strSQL = strSQL & "95,255,255,1;90,255,255,1;85,255,255,1;80,255,255,1;75,255,255,1;70,255,255,1;65,255,255,1;60,255,255,1;55,255,255,1;50,255,255,1;"
    strSQL = strSQL & "45,255,255,1;40,255,255,1;35,255,255,1;30,255,255,1;25,255,255,1;20,255,255,1;15,255,255,1;10,255,255,1;5,255,255,1;0,255,255,1;"
    strSQL = strSQL & "2,253,255,1;5,250,255,1;7,248,255,1;10,245,255,1;12,243,255,1;15,240,255,1;17,238,255,1;20,235,255,1;22,233,255,1;25,230,255,1;"
    strSQL = strSQL & "27,228,255,1;30,225,255,1;32,223,255,1;35,220,255,1;37,218,255,1;40,215,255,1;42,213,255,1;45,210,255,1;47,208,255,1;50,205,255,1;"
    strSQL = strSQL & "52,203,255,1;55,200,255,1;57,198,255,1;60,195,255,1;62,193,255,1;65,190,255,1;67,188,255,1;70,185,255,1;72,183,255,1;75,180,255,1;"
    strSQL = strSQL & "77,178,255,1;80,175,255,1;82,173,255,1;85,170,255,1;87,168,255,1;90,165,255,1;92,163,255,1;95,160,255,1;97,158,255,1;100,155,255,1;"
    strSQL = strSQL & "102,153,255,1;105,150,255,1;107,148,255,1;110,145,255,1;112,143,255,1;115,140,255,1;117,138,255,1;120,135,255,1;122,133,255,1;"
    strSQL = strSQL & "125,130,255,1;128,128,255,1;130,125,255,1;133,122,255,1;135,120,255,1;138,117,255,1;140,115,255,1;143,112,255,1;145,110,255,1;"
    strSQL = strSQL & "148,107,255,1;150,105,255,1;153,102,255,1;155,100,255,1;158,97,255,1;160,95,255,1;163,92,255,1;165,90,255,1;168,87,255,1;170,85,255,1;"
    strSQL = strSQL & "173,82,255,1;175,80,255,1;178,77,255,1;180,75,255,1;183,72,255,1;185,70,255,1;188,67,255,1;190,65,255,1;193,62,255,1;195,60,255,1;"
    strSQL = strSQL & "198,57,255,1;200,55,255,1;203,52,255,1;205,50,255,1;208,47,255,1;210,45,255,1;213,42,255,1;215,40,255,1;218,37,255,1;220,35,255,1;"
    strSQL = strSQL & "223,32,255,1;225,30,255,1;228,27,255,1;230,25,255,1;233,22,255,1;235,20,255,1;238,17,255,1;240,15,255,1;243,12,255,1;245,10,255,1;"
    strSQL = strSQL & "248,7,255,1;250,5,255,1;253,2,255,1;255,0,255,1;255,2,253,1;255,5,250,1;255,7,248,1;255,10,245,1;255,12,243,1;255,15,240,1;"
    strSQL = strSQL & "255,17,238,1;255,20,235,1;255,22,233,1;255,25,230,1;255,27,228,1;255,30,225,1;255,32,223,1;255,35,220,1;255,37,218,1;255,40,215,1;"
    strSQL = strSQL & "255,42,213,1;255,45,210,1;255,47,208,1;255,50,205,1;255,52,203,1;255,55,200,1;255,57,198,1;255,60,195,1;255,62,193,1;255,65,190,1;"
    strSQL = strSQL & "255,67,188,1;255,70,185,1;255,72,183,1;255,75,180,1;255,77,178,1;255,80,175,1;255,82,173,1;255,85,170,1;255,87,168,1;255,90,165,1;"
    strSQL = strSQL & "255,92,163,1;255,95,160,1;255,97,158,1;255,100,155,1;255,102,153,1;255,105,150,1;255,107,148,1;255,110,145,1;255,112,143,1;"
    strSQL = strSQL & "255,115,140,1;255,117,138,1;255,120,135,1;255,122,133,1;255,125,130,1;255,128,127,1;255,130,125,1;255,133,122,1;255,135,120,1;"
    strSQL = strSQL & "255,138,117,1;255,140,115,1;255,143,112,1;255,145,110,1;255,148,107,1;255,150,105,1;255,153,102,1;255,155,100,1;255,158,97,1;"
    strSQL = strSQL & "255,160,95,1;255,163,92,1;255,165,90,1;255,168,87,1;255,170,85,1;255,173,82,1;255,175,80,1;255,178,77,1;255,180,75,1;255,183,72,1;"
    strSQL = strSQL & "255,185,70,1;255,188,67,1;255,190,65,1;255,193,62,1;255,195,60,1;255,198,59,1;255,201,58,1;255,204,59,1;255,208,59,1;255,211,61,1;"
    strSQL = strSQL & "255,214,63,1;255,218,66,1;255,221,70,1;255,224,74,1;255,228,80,1;255,231,86,1;255,234,93,1;255,237,101,1;255,240,110,1;255,242,120,1;"
    strSQL = strSQL & "255,245,131,1;255,247,143,1;255,249,155,1;255,251,169,1;255,253,184,1;255,254,200,1;255,255,218,1;255,255,236,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (15,'cool','0,0,2,1;0,2,4,1;0,4,6,1;0,6,8,1;0,8,10,1;0,10,12,1;0,12,14,1;"
    strSQL = strSQL & "0,14,16,1;0,16,18,1;0,18,20,1;0,20,22,1;0,22,24,1;0,24,26,1;0,26,28,1;0,28,30,1;0,30,32,1;0,33,34,1;0,35,36,1;0,37,38,1;0,39,40,1;"
    strSQL = strSQL & "0,41,42,1;0,43,44,1;0,45,46,1;0,47,48,1;0,49,50,1;0,51,52,1;0,53,54,1;0,55,56,1;0,57,58,1;0,59,60,1;0,61,62,1;0,63,64,1;0,65,66,1;"
    strSQL = strSQL & "0,67,68,1;0,69,70,1;0,71,72,1;0,73,74,1;0,75,76,1;0,77,78,1;0,79,80,1;0,81,82,1;0,83,84,1;0,85,86,1;0,87,88,1;0,89,90,1;0,91,92,1;"
    strSQL = strSQL & "0,93,94,1;0,95,96,1;0,98,98,1;0,100,100,1;0,102,102,1;0,104,104,1;0,106,106,1;0,108,108,1;0,110,110,1;0,112,112,1;0,114,114,1;"
    strSQL = strSQL & "0,116,116,1;0,118,118,1;0,120,120,1;0,122,122,1;0,124,124,1;0,126,126,1;2,128,128,1;4,126,129,1;6,124,131,1;8,122,133,1;10,120,135,1;"
    strSQL = strSQL & "12,118,137,1;14,116,139,1;16,114,141,1;18,112,143,1;20,110,145,1;22,108,147,1;24,106,149,1;26,104,151,1;28,102,153,1;30,100,155,1;"
    strSQL = strSQL & "32,98,157,1;34,96,159,1;36,95,161,1;38,93,163,1;40,91,165,1;42,89,167,1;43,87,169,1;45,85,171,1;47,83,173,1;49,81,175,1;51,79,177,1;"
    strSQL = strSQL & "53,77,179,1;55,75,181,1;57,73,183,1;59,71,185,1;61,69,187,1;63,67,189,1;65,65,191,1;67,63,193,1;69,61,195,1;71,59,197,1;73,57,199,1;"
    strSQL = strSQL & "75,55,201,1;77,53,203,1;79,51,205,1;81,49,207,1;83,47,209,1;85,45,211,1;87,43,213,1;89,41,215,1;91,39,217,1;93,37,219,1;95,35,221,1;"
    strSQL = strSQL & "97,33,223,1;99,32,225,1;101,30,227,1;103,28,229,1;105,26,231,1;107,24,233,1;109,22,235,1;111,20,237,1;113,18,239,1;115,16,241,1;"
    strSQL = strSQL & "117,14,243,1;119,12,245,1;121,10,247,1;123,8,249,1;125,6,251,1;127,4,253,1;128,2,255,1;130,0,251,1;132,2,247,1;134,4,243,1;136,6,239,1;"
    strSQL = strSQL & "138,8,235,1;140,10,231,1;142,12,227,1;144,14,223,1;146,16,219,1;148,18,215,1;150,20,211,1;152,22,207,1;154,24,203,1;156,26,199,1;"
    strSQL = strSQL & "158,28,195,1;160,30,191,1;162,32,187,1;164,34,183,1;166,36,179,1;168,38,175,1;170,40,171,1;172,42,167,1;174,44,163,1;176,46,159,1;"
    strSQL = strSQL & "178,48,155,1;180,50,151,1;182,52,147,1;184,54,143,1;186,56,139,1;188,58,135,1;190,60,131,1;192,62,128,1;194,64,124,1;196,66,120,1;"
    strSQL = strSQL & "198,68,116,1;200,70,112,1;202,72,108,1;204,74,104,1;206,76,100,1;208,78,96,1;210,80,92,1;212,82,88,1;213,84,84,1;215,86,80,1;"
    strSQL = strSQL & "217,88,76,1;219,90,72,1;221,92,68,1;223,94,64,1;225,96,60,1;227,98,56,1;229,100,52,1;231,102,48,1;233,104,44,1;235,106,40,1;"
    strSQL = strSQL & "237,108,36,1;239,110,32,1;241,112,28,1;243,114,24,1;245,116,20,1;247,118,16,1;248,120,12,1;251,122,8,1;253,124,4,1;255,126,0,1;"
    strSQL = strSQL & "255,129,4,1;255,131,8,1;255,133,12,1;255,135,16,1;255,137,20,1;255,139,24,1;255,141,28,1;255,143,32,1;255,145,36,1;255,147,40,1;"
    strSQL = strSQL & "255,149,45,1;255,151,49,1;255,153,53,1;255,155,57,1;255,157,61,1;255,159,65,1;255,161,69,1;255,163,73,1;255,165,77,1;255,167,81,1;"
    strSQL = strSQL & "255,169,85,1;255,171,89,1;255,173,93,1;255,175,97,1;255,177,101,1;255,179,105,1;255,181,109,1;255,183,113,1;255,185,117,1;"
    strSQL = strSQL & "255,187,121,1;255,189,125,1;255,191,130,1;255,193,134,1;255,195,138,1;255,197,142,1;255,199,146,1;255,201,150,1;255,203,154,1;"
    strSQL = strSQL & "255,205,158,1;255,207,162,1;255,209,166,1;255,211,170,1;255,213,174,1;255,215,178,1;255,217,182,1;255,219,186,1;255,221,190,1;"
    strSQL = strSQL & "255,223,194,1;255,225,198,1;255,227,202,1;255,229,206,1;255,231,210,1;255,233,215,1;255,235,219,1;255,237,223,1;255,239,227,1;"
    strSQL = strSQL & "255,241,231,1;255,243,235,1;255,245,239,1;255,247,243,1;255,249,247,1;255,251,251,1;255,253,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (16,'Royal','0,0,0,1;0,0,6,1;0,0,12,1;0,0,17,1;0,0,23,1;0,0,29,1;0,0,35,1;"
    strSQL = strSQL & "0,0,41,1;0,0,46,1;0,0,52,1;0,0,58,1;0,0,64,1;0,0,70,1;0,0,75,1;0,0,81,1;0,0,85,1;0,0,89,1;0,0,93,1;0,0,98,1;0,0,104,1;0,0,109,1;"
    strSQL = strSQL & "0,0,115,1;0,0,121,1;0,0,126,1;0,0,132,1;0,0,137,1;0,0,143,1;0,0,149,1;0,0,154,1;0,0,160,1;0,0,165,1;0,0,169,1;0,0,174,1;0,6,177,1;"
    strSQL = strSQL & "0,12,182,1;0,17,188,1;0,23,193,1;0,29,199,1;0,35,205,1;0,41,210,1;0,46,216,1;0,52,221,1;0,58,227,1;0,64,233,1;0,70,238,1;0,75,244,1;"
    strSQL = strSQL & "0,81,249,1;0,85,255,1;0,89,255,1;0,93,255,1;0,98,255,1;0,104,255,1;0,109,255,1;0,115,255,1;0,121,255,1;0,126,255,1;0,132,255,1;"
    strSQL = strSQL & "0,137,255,1;0,143,255,1;0,149,255,1;0,154,255,1;0,160,255,1;0,165,255,1;0,169,255,1;0,174,255,1;0,177,255,1;0,182,255,1;0,188,255,1;"
    strSQL = strSQL & "0,193,255,1;0,199,255,1;0,205,255,1;0,210,255,1;0,216,255,1;0,221,255,1;0,227,255,1;0,233,255,1;0,238,255,1;0,244,255,1;0,249,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,239,1;9,255,223,1;17,255,207,1;26,255,191,1;35,255,175,1;"
    strSQL = strSQL & "44,255,159,1;52,255,143,1;61,255,127,1;70,255,111,1;79,255,95,1;87,255,79,1;96,255,63,1;105,255,47,1;114,255,31,1;122,255,15,1;"
    strSQL = strSQL & "128,255,0,1;134,255,0,1;139,253,0,1;148,251,0,1;156,249,0,1;164,246,0,1;172,244,0,1;181,242,0,1;189,240,0,1;197,238,0,1;205,236,0,1;"
    strSQL = strSQL & "214,234,0,1;222,232,0,1;230,229,0,1;238,227,0,1;247,225,0,1;255,223,0,1;255,223,0,1;255,220,0,1;255,218,0,1;255,215,0,1;255,212,0,1;"
    strSQL = strSQL & "255,210,0,1;255,207,0,1;255,204,0,1;255,202,0,1;255,199,0,1;255,196,0,1;255,194,0,1;255,191,0,1;255,188,0,1;255,186,0,1;255,184,0,1;"
    strSQL = strSQL & "255,182,0,1;255,180,0,1;255,176,0,1;255,173,0,1;255,169,0,1;255,166,0,1;255,162,0,1;255,159,0,1;255,155,0,1;255,152,0,1;255,148,0,1;"
    strSQL = strSQL & "255,145,0,1;255,141,0,1;255,138,0,1;255,134,0,1;255,129,0,1;255,131,0,1;255,127,0,1;255,122,0,1;255,118,0,1;255,114,0,1;255,110,0,1;"
    strSQL = strSQL & "255,105,0,1;255,101,0,1;255,97,0,1;255,93,0,1;255,88,0,1;255,84,0,1;255,80,0,1;255,76,0,1;255,71,0,1;255,66,0,1;255,68,0,1;"
    strSQL = strSQL & "255,63,0,1;255,58,0,1;255,54,0,1;255,49,0,1;255,45,0,1;255,40,0,1;255,36,0,1;255,31,0,1;255,27,0,1;255,22,0,1;255,18,0,1;255,13,0,1;"
    strSQL = strSQL & "255,9,0,1;255,4,0,1;255,0,0,1;255,0,0,1;253,13,13,1;251,25,25,1;249,38,38,1;246,51,51,1;244,64,64,1;242,76,76,1;240,89,89,1;"
    strSQL = strSQL & "238,102,102,1;236,115,115,1;234,127,127,1;232,140,140,1;229,153,153,1;227,166,166,1;225,178,178,1;223,183,183,1;223,188,188,1;"
    strSQL = strSQL & "223,193,193,1;223,195,195,1;223,197,197,1;223,200,200,1;223,202,202,1;223,204,204,1;223,206,206,1;223,208,208,1;223,210,210,1;"
    strSQL = strSQL & "223,212,212,1;223,214,214,1;223,217,217,1;223,219,219,1;223,221,221,1;223,223,223,1;223,223,223,1;225,225,225,1;227,227,227,1;"
    strSQL = strSQL & "229,229,229,1;232,232,232,1;234,234,234,1;236,236,236,1;238,238,238,1;240,240,240,1;242,242,242,1;244,244,244,1;246,246,246,1;"
    strSQL = strSQL & "249,249,249,1;251,251,251,1;253,253,253,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (17,'16_Color','0,0,0,1;0,0,0,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;"
    strSQL = strSQL & "0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,87,1;0,0,84,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;"
    strSQL = strSQL & "0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,171,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;"
    strSQL = strSQL & "0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,87,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;"
    strSQL = strSQL & "0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;0,171,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;"
    strSQL = strSQL & "131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;131,255,0,1;"
    strSQL = strSQL & "131,255,0,1;131,255,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;"
    strSQL = strSQL & "255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,223,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;"
    strSQL = strSQL & "255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;255,183,0,1;"
    strSQL = strSQL & "255,183,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;"
    strSQL = strSQL & "255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,131,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;"
    strSQL = strSQL & "255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,67,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;"
    strSQL = strSQL & "223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,191,191,1;223,223,223,1;"
    strSQL = strSQL & "223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;"
    strSQL = strSQL & "223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;223,223,223,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (18,'Ceretec','0,0,0,1;0,0,3,1;0,0,7,1;0,0,11,1;0,0,15,1;0,0,19,1;"
    strSQL = strSQL & "0,0,23,1;0,0,27,1;0,0,31,1;0,0,35,1;0,0,39,1;0,0,43,1;0,0,47,1;0,0,51,1;0,0,55,1;0,0,59,1;0,0,63,1;0,0,67,1;0,0,71,1;0,0,75,1;"
    strSQL = strSQL & "0,0,79,1;0,0,83,1;0,0,87,1;0,0,91,1;0,0,95,1;0,0,99,1;0,0,103,1;0,0,107,1;0,0,111,1;0,0,115,1;0,0,119,1;0,0,123,1;0,0,127,1;"
    strSQL = strSQL & "0,0,131,1;0,0,135,1;0,0,139,1;0,0,143,1;0,0,147,1;0,0,151,1;0,0,155,1;0,0,159,1;0,0,163,1;0,0,167,1;0,0,171,1;0,0,175,1;0,0,179,1;"
    strSQL = strSQL & "0,0,183,1;0,0,187,1;0,0,191,1;0,0,195,1;0,0,199,1;0,0,203,1;0,0,207,1;0,0,211,1;0,0,215,1;0,0,219,1;0,0,223,1;0,0,227,1;0,0,231,1;"
    strSQL = strSQL & "0,0,235,1;0,0,239,1;0,0,243,1;0,0,247,1;0,0,251,1;3,3,255,1;7,6,251,1;11,8,247,1;15,11,243,1;19,13,239,1;23,16,235,1;27,18,231,1;"
    strSQL = strSQL & "31,21,227,1;35,23,223,1;39,26,219,1;43,29,215,1;47,31,211,1;51,34,207,1;55,36,203,1;59,39,199,1;63,41,195,1;67,44,191,1;71,47,187,1;"
    strSQL = strSQL & "75,49,183,1;79,52,179,1;83,54,175,1;87,57,171,1;91,59,167,1;95,62,163,1;99,64,159,1;103,67,155,1;107,70,151,1;111,72,147,1;115,75,143,1;"
    strSQL = strSQL & "119,77,139,1;123,80,135,1;127,82,131,1;131,85,127,1;135,87,123,1;139,90,119,1;143,93,115,1;147,95,111,1;151,98,107,1;155,100,103,1;"
    strSQL = strSQL & "159,103,99,1;163,105,95,1;167,108,91,1;171,111,87,1;175,113,83,1;179,116,79,1;183,118,75,1;187,121,71,1;191,123,67,1;195,126,63,1;"
    strSQL = strSQL & "199,128,59,1;203,131,55,1;207,134,51,1;211,136,47,1;215,139,43,1;219,141,39,1;223,144,35,1;227,146,31,1;231,149,27,1;235,151,23,1;"
    strSQL = strSQL & "239,154,19,1;243,157,15,1;247,159,11,1;251,162,7,1;255,164,3,1;255,161,0,1;255,158,0,1;255,156,0,1;255,153,0,1;255,151,0,1;"
    strSQL = strSQL & "255,148,0,1;255,146,0,1;255,143,0,1;255,141,0,1;255,138,0,1;255,135,0,1;255,133,0,1;255,130,0,1;255,128,0,1;255,125,0,1;255,123,0,1;"
    strSQL = strSQL & "255,120,0,1;255,117,0,1;255,115,0,1;255,112,0,1;255,110,0,1;255,107,0,1;255,105,0,1;255,102,0,1;255,100,0,1;255,97,0,1;255,94,0,1;"
    strSQL = strSQL & "255,92,0,1;255,89,0,1;255,87,0,1;255,84,0,1;255,82,0,1;255,79,0,1;255,77,0,1;255,74,0,1;255,71,0,1;255,69,0,1;255,66,0,1;255,64,0,1;"
    strSQL = strSQL & "255,61,0,1;255,59,0,1;255,56,0,1;255,53,0,1;255,51,0,1;255,48,0,1;255,46,0,1;255,43,0,1;255,41,0,1;255,38,0,1;255,36,0,1;255,33,0,1;"
    strSQL = strSQL & "255,30,0,1;255,28,0,1;255,25,0,1;255,23,0,1;255,20,0,1;255,18,0,1;255,15,0,1;255,13,0,1;255,10,0,1;255,7,0,1;255,5,0,1;255,2,0,1;"
    strSQL = strSQL & "255,0,0,1;255,3,3,1;255,7,7,1;255,11,11,1;255,15,15,1;255,19,19,1;255,23,23,1;255,27,27,1;255,31,31,1;255,35,35,1;255,39,39,1;"
    strSQL = strSQL & "255,43,43,1;255,47,47,1;255,51,51,1;255,55,55,1;255,59,59,1;255,63,63,1;255,67,67,1;255,71,71,1;255,75,75,1;255,79,79,1;255,83,83,1;"
    strSQL = strSQL & "255,87,87,1;255,91,91,1;255,95,95,1;255,99,99,1;255,103,103,1;255,107,107,1;255,111,111,1;255,115,115,1;255,119,119,1;255,123,123,1;"
    strSQL = strSQL & "255,127,127,1;255,131,131,1;255,135,135,1;255,139,139,1;255,143,143,1;255,147,147,1;255,151,151,1;255,155,155,1;255,159,159,1;"
    strSQL = strSQL & "255,163,163,1;255,167,167,1;255,171,171,1;255,175,175,1;255,179,179,1;255,183,183,1;255,187,187,1;255,191,191,1;255,195,195,1;"
    strSQL = strSQL & "255,199,199,1;255,203,203,1;255,207,207,1;255,211,211,1;255,215,215,1;255,219,219,1;255,223,223,1;255,227,227,1;255,231,231,1;"
    strSQL = strSQL & "255,235,235,1;255,239,239,1;255,243,243,1;255,247,247,1;255,251,251,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (19,'gray','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;1,1,1,1;1,1,1,1;2,2,2,1;"
    strSQL = strSQL & "3,3,3,1;4,4,4,1;5,5,5,1;6,6,6,1;7,7,7,1;8,8,8,1;9,9,9,1;11,11,11,1;12,12,12,1;13,13,13,1;15,15,15,1;16,16,16,1;17,17,17,1;19,19,19,1;"
    strSQL = strSQL & "20,20,20,1;21,21,21,1;22,22,22,1;24,24,24,1;25,25,25,1;26,26,26,1;27,27,27,1;28,28,28,1;29,29,29,1;30,30,30,1;31,31,31,1;32,32,32,1;"
    strSQL = strSQL & "33,33,33,1;34,34,34,1;35,35,35,1;36,36,36,1;37,37,37,1;38,38,38,1;39,39,39,1;40,40,40,1;41,41,41,1;42,42,42,1;43,43,43,1;44,44,44,1;"
    strSQL = strSQL & "45,45,45,1;46,46,46,1;47,47,47,1;48,48,48,1;49,49,49,1;50,50,50,1;51,51,51,1;52,52,52,1;53,53,53,1;54,54,54,1;55,55,55,1;56,56,56,1;"
    strSQL = strSQL & "57,57,57,1;58,58,58,1;59,59,59,1;60,60,60,1;61,61,61,1;62,62,62,1;63,63,63,1;64,64,64,1;65,65,65,1;66,66,66,1;67,67,67,1;68,68,68,1;"
    strSQL = strSQL & "69,69,69,1;70,70,70,1;71,71,71,1;72,72,72,1;73,73,73,1;74,74,74,1;75,75,75,1;76,76,76,1;77,77,77,1;78,78,78,1;79,79,79,1;80,80,80,1;"
    strSQL = strSQL & "81,81,81,1;82,82,82,1;83,83,83,1;84,84,84,1;85,85,85,1;86,86,86,1;87,87,87,1;88,88,88,1;89,89,89,1;90,90,90,1;91,91,91,1;92,92,92,1;"
    strSQL = strSQL & "93,93,93,1;94,94,94,1;95,95,95,1;96,96,96,1;97,97,97,1;98,98,98,1;99,99,99,1;100,100,100,1;101,101,101,1;102,102,102,1;103,103,103,1;"
    strSQL = strSQL & "104,104,104,1;105,105,105,1;106,106,106,1;107,107,107,1;108,108,108,1;109,109,109,1;110,110,110,1;111,111,111,1;112,112,112,1;"
    strSQL = strSQL & "113,113,113,1;114,114,114,1;115,115,115,1;116,116,116,1;117,117,117,1;118,118,118,1;119,119,119,1;120,120,120,1;121,121,121,1;"
    strSQL = strSQL & "122,122,122,1;123,123,123,1;124,124,124,1;125,125,125,1;126,126,126,1;127,127,127,1;128,128,128,1;129,129,129,1;130,130,130,1;"
    strSQL = strSQL & "131,131,131,1;132,132,132,1;133,133,133,1;134,134,134,1;135,135,135,1;136,136,136,1;137,137,137,1;138,138,138,1;139,139,139,1;"
    strSQL = strSQL & "140,140,140,1;141,141,141,1;142,142,142,1;143,143,143,1;144,144,144,1;145,145,145,1;146,146,146,1;147,147,147,1;148,148,148,1;"
    strSQL = strSQL & "149,149,149,1;150,150,150,1;151,151,151,1;152,152,152,1;153,153,153,1;154,154,154,1;155,155,155,1;156,156,156,1;157,157,157,1;"
    strSQL = strSQL & "158,158,158,1;159,159,159,1;160,160,160,1;161,161,161,1;162,162,162,1;163,163,163,1;164,164,164,1;165,165,165,1;166,166,166,1;"
    strSQL = strSQL & "167,167,167,1;168,168,168,1;169,169,169,1;170,170,170,1;171,171,171,1;172,172,172,1;173,173,173,1;174,174,174,1;175,175,175,1;"
    strSQL = strSQL & "176,176,176,1;177,177,177,1;178,178,178,1;179,179,179,1;180,180,180,1;181,181,181,1;182,182,182,1;183,183,183,1;184,184,184,1;"
    strSQL = strSQL & "185,185,185,1;186,186,186,1;187,187,187,1;188,188,188,1;189,189,189,1;190,190,190,1;191,191,191,1;192,192,192,1;193,193,193,1;"
    strSQL = strSQL & "194,194,194,1;195,195,195,1;196,196,196,1;197,197,197,1;198,198,198,1;199,199,199,1;200,200,200,1;201,201,201,1;202,202,202,1;"
    strSQL = strSQL & "203,203,203,1;204,204,204,1;205,205,205,1;206,206,206,1;207,207,207,1;208,208,208,1;209,209,209,1;210,210,210,1;211,211,211,1;"
    strSQL = strSQL & "212,212,212,1;213,213,213,1;214,214,214,1;215,215,215,1;216,216,216,1;217,217,217,1;218,218,218,1;219,219,219,1;220,220,220,1;"
    strSQL = strSQL & "221,221,221,1;222,222,222,1;223,223,223,1;224,224,224,1;225,225,225,1;226,226,226,1;227,227,227,1;228,228,228,1;229,229,229,1;"
    strSQL = strSQL & "230,230,230,1;231,231,231,1;233,233,233,1;234,234,234,1;235,235,235,1;236,236,236,1;238,238,238,1;239,239,239,1;240,240,240,1;"
    strSQL = strSQL & "242,242,242,1;243,243,243,1;244,244,244,1;246,246,246,1;247,247,247,1;248,248,248,1;249,249,249,1;250,250,250,1;251,251,251,1;"
    strSQL = strSQL & "252,252,252,1;253,253,253,1;254,254,254,1;254,254,254,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (20,'ge-16','0,0,0,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;"
    strSQL = strSQL & "0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,16,32,1;0,16,32,1;0,16,32,1;"
    strSQL = strSQL & "0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;"
    strSQL = strSQL & "0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;"
    strSQL = strSQL & "0,24,48,1;0,24,48,1;0,24,48,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;"
    strSQL = strSQL & "0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;"
    strSQL = strSQL & "31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;63,16,96,1;63,16,96,1;"
    strSQL = strSQL & "63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;"
    strSQL = strSQL & "63,16,96,1;63,16,96,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;"
    strSQL = strSQL & "95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;"
    strSQL = strSQL & "127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;"
    strSQL = strSQL & "159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;"
    strSQL = strSQL & "159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;"
    strSQL = strSQL & "191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;223,74,31,1;"
    strSQL = strSQL & "223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;"
    strSQL = strSQL & "223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;"
    strSQL = strSQL & "255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,138,67,1;255,138,67,1;"
    strSQL = strSQL & "255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;"
    strSQL = strSQL & "255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;"
    strSQL = strSQL & "255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;"
    strSQL = strSQL & "255,178,131,1;255,178,131,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;"
    strSQL = strSQL & "255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (21,'hot','0,0,0,1;2,0,0,1;5,0,0,1;7,0,0,1;10,0,0,1;13,0,0,1;15,0,0,1;"
    strSQL = strSQL & "18,0,0,1;21,0,0,1;23,0,0,1;26,0,0,1;29,0,0,1;31,0,0,1;34,0,0,1;37,0,0,1;39,0,0,1;42,0,0,1;45,0,0,1;47,0,0,1;50,0,0,1;53,0,0,1;"
    strSQL = strSQL & "55,0,0,1;58,0,0,1;61,0,0,1;63,0,0,1;66,0,0,1;69,0,0,1;71,0,0,1;74,0,0,1;77,0,0,1;79,0,0,1;82,0,0,1;85,0,0,1;87,0,0,1;90,0,0,1;"
    strSQL = strSQL & "92,0,0,1;95,0,0,1;98,0,0,1;100,0,0,1;103,0,0,1;106,0,0,1;108,0,0,1;111,0,0,1;114,0,0,1;116,0,0,1;119,0,0,1;122,0,0,1;124,0,0,1;"
    strSQL = strSQL & "127,0,0,1;130,0,0,1;132,0,0,1;135,0,0,1;138,0,0,1;140,0,0,1;143,0,0,1;146,0,0,1;148,0,0,1;151,0,0,1;154,0,0,1;156,0,0,1;159,0,0,1;"
    strSQL = strSQL & "162,0,0,1;164,0,0,1;167,0,0,1;170,0,0,1;172,0,0,1;175,0,0,1;177,0,0,1;180,0,0,1;183,0,0,1;185,0,0,1;188,0,0,1;191,0,0,1;193,0,0,1;"
    strSQL = strSQL & "196,0,0,1;199,0,0,1;201,0,0,1;204,0,0,1;207,0,0,1;209,0,0,1;212,0,0,1;215,0,0,1;217,0,0,1;220,0,0,1;223,0,0,1;225,0,0,1;228,0,0,1;"
    strSQL = strSQL & "231,0,0,1;233,0,0,1;236,0,0,1;239,0,0,1;241,0,0,1;244,0,0,1;247,0,0,1;249,0,0,1;252,0,0,1;255,0,0,1;255,2,0,1;255,5,0,1;255,7,0,1;"
    strSQL = strSQL & "255,10,0,1;255,13,0,1;255,15,0,1;255,18,0,1;255,21,0,1;255,23,0,1;255,26,0,1;255,29,0,1;255,31,0,1;255,34,0,1;255,37,0,1;255,39,0,1;"
    strSQL = strSQL & "255,42,0,1;255,45,0,1;255,47,0,1;255,50,0,1;255,53,0,1;255,55,0,1;255,58,0,1;255,61,0,1;255,63,0,1;255,66,0,1;255,69,0,1;255,71,0,1;"
    strSQL = strSQL & "255,74,0,1;255,77,0,1;255,79,0,1;255,82,0,1;255,85,0,1;255,87,0,1;255,90,0,1;255,92,0,1;255,95,0,1;255,98,0,1;255,100,0,1;255,103,0,1;"
    strSQL = strSQL & "255,106,0,1;255,108,0,1;255,111,0,1;255,114,0,1;255,116,0,1;255,119,0,1;255,122,0,1;255,124,0,1;255,127,0,1;255,130,0,1;255,132,0,1;"
    strSQL = strSQL & "255,135,0,1;255,138,0,1;255,140,0,1;255,143,0,1;255,146,0,1;255,148,0,1;255,151,0,1;255,154,0,1;255,156,0,1;255,159,0,1;255,162,0,1;"
    strSQL = strSQL & "255,164,0,1;255,167,0,1;255,170,0,1;255,172,0,1;255,175,0,1;255,177,0,1;255,180,0,1;255,183,0,1;255,185,0,1;255,188,0,1;255,191,0,1;"
    strSQL = strSQL & "255,193,0,1;255,196,0,1;255,199,0,1;255,201,0,1;255,204,0,1;255,207,0,1;255,209,0,1;255,212,0,1;255,215,0,1;255,217,0,1;255,220,0,1;"
    strSQL = strSQL & "255,223,0,1;255,225,0,1;255,228,0,1;255,231,0,1;255,233,0,1;255,236,0,1;255,239,0,1;255,241,0,1;255,244,0,1;255,247,0,1;255,249,0,1;"
    strSQL = strSQL & "255,252,0,1;255,255,0,1;255,255,3,1;255,255,7,1;255,255,11,1;255,255,15,1;255,255,19,1;255,255,23,1;255,255,27,1;255,255,31,1;"
    strSQL = strSQL & "255,255,35,1;255,255,39,1;255,255,43,1;255,255,47,1;255,255,51,1;255,255,55,1;255,255,59,1;255,255,63,1;255,255,67,1;255,255,71,1;"
    strSQL = strSQL & "255,255,75,1;255,255,79,1;255,255,83,1;255,255,87,1;255,255,91,1;255,255,95,1;255,255,99,1;255,255,103,1;255,255,107,1;255,255,111,1;"
    strSQL = strSQL & "255,255,115,1;255,255,119,1;255,255,123,1;255,255,127,1;255,255,131,1;255,255,135,1;255,255,139,1;255,255,143,1;255,255,147,1;"
    strSQL = strSQL & "255,255,151,1;255,255,155,1;255,255,159,1;255,255,163,1;255,255,167,1;255,255,171,1;255,255,175,1;255,255,179,1;255,255,183,1;"
    strSQL = strSQL & "255,255,187,1;255,255,191,1;255,255,195,1;255,255,199,1;255,255,203,1;255,255,207,1;255,255,211,1;255,255,215,1;255,255,219,1;"
    strSQL = strSQL & "255,255,223,1;255,255,227,1;255,255,231,1;255,255,235,1;255,255,239,1;255,255,243,1;255,255,247,1;255,255,251,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (22,'hot_Iron','0,0,0,1;0,0,3,1;0,0,7,1;0,0,11,1;0,0,15,1;0,0,19,1;"
    strSQL = strSQL & "0,0,23,1;0,0,27,1;0,0,31,1;0,0,35,1;0,0,39,1;0,0,43,1;0,0,47,1;0,0,51,1;0,0,55,1;0,0,59,1;0,0,63,1;0,0,67,1;0,0,71,1;0,0,75,1;"
    strSQL = strSQL & "0,0,79,1;0,0,83,1;0,0,87,1;0,0,91,1;0,0,95,1;0,0,99,1;0,0,103,1;0,0,107,1;0,0,111,1;0,0,115,1;0,0,119,1;0,0,123,1;0,0,127,1;"
    strSQL = strSQL & "0,0,131,1;0,0,135,1;0,0,139,1;0,0,143,1;0,0,147,1;0,0,151,1;0,0,155,1;0,0,159,1;0,0,163,1;0,0,167,1;0,0,171,1;0,0,175,1;0,0,179,1;"
    strSQL = strSQL & "0,0,183,1;0,0,187,1;0,0,191,1;0,0,195,1;0,0,199,1;0,0,203,1;0,0,207,1;0,0,211,1;0,0,215,1;0,0,219,1;0,0,223,1;0,0,227,1;0,0,231,1;"
    strSQL = strSQL & "0,0,235,1;0,0,239,1;0,0,243,1;0,0,247,1;0,0,251,1;0,0,255,1;0,0,255,1;3,0,251,1;7,0,247,1;11,0,243,1;15,0,239,1;19,0,235,1;23,0,231,1;"
    strSQL = strSQL & "27,0,227,1;31,0,223,1;35,0,219,1;39,0,215,1;43,0,211,1;47,0,207,1;51,0,203,1;55,0,199,1;59,0,195,1;63,0,191,1;67,0,187,1;71,0,183,1;"
    strSQL = strSQL & "75,0,179,1;79,0,175,1;83,0,171,1;87,0,167,1;91,0,163,1;95,0,159,1;99,0,155,1;103,0,151,1;107,0,147,1;111,0,143,1;115,0,139,1;"
    strSQL = strSQL & "119,0,135,1;123,0,131,1;127,0,127,1;131,0,123,1;135,0,119,1;139,0,115,1;143,0,111,1;147,0,107,1;151,0,103,1;155,0,99,1;159,0,95,1;"
    strSQL = strSQL & "163,0,91,1;167,0,87,1;171,0,83,1;175,0,79,1;179,0,75,1;183,0,71,1;187,0,67,1;191,0,63,1;195,0,59,1;199,0,55,1;203,0,51,1;207,0,47,1;"
    strSQL = strSQL & "211,0,43,1;215,0,39,1;219,0,35,1;223,0,31,1;227,0,27,1;231,0,23,1;235,0,19,1;239,0,15,1;243,0,11,1;247,0,7,1;251,0,3,1;255,0,0,1;"
    strSQL = strSQL & "255,3,0,1;255,7,0,1;255,11,0,1;255,15,0,1;255,19,0,1;255,23,0,1;255,27,0,1;255,31,0,1;255,35,0,1;255,39,0,1;255,43,0,1;255,47,0,1;"
    strSQL = strSQL & "255,51,0,1;255,55,0,1;255,59,0,1;255,63,0,1;255,67,0,1;255,71,0,1;255,75,0,1;255,79,0,1;255,83,0,1;255,87,0,1;255,91,0,1;255,95,0,1;"
    strSQL = strSQL & "255,99,0,1;255,103,0,1;255,107,0,1;255,111,0,1;255,115,0,1;255,119,0,1;255,123,0,1;255,127,0,1;255,131,0,1;255,135,0,1;255,139,0,1;"
    strSQL = strSQL & "255,143,0,1;255,147,0,1;255,151,0,1;255,155,0,1;255,159,0,1;255,163,0,1;255,167,0,1;255,171,0,1;255,175,0,1;255,179,0,1;255,183,0,1;"
    strSQL = strSQL & "255,187,0,1;255,191,0,1;255,195,0,1;255,199,0,1;255,203,0,1;255,207,0,1;255,211,0,1;255,215,0,1;255,219,0,1;255,223,0,1;255,227,0,1;"
    strSQL = strSQL & "255,231,0,1;255,235,0,1;255,239,0,1;255,243,0,1;255,247,0,1;255,251,0,1;255,255,0,1;255,255,4,1;255,255,8,1;255,255,12,1;255,255,16,1;"
    strSQL = strSQL & "255,255,20,1;255,255,24,1;255,255,28,1;255,255,32,1;255,255,36,1;255,255,40,1;255,255,44,1;255,255,48,1;255,255,52,1;255,255,56,1;"
    strSQL = strSQL & "255,255,60,1;255,255,64,1;255,255,68,1;255,255,72,1;255,255,76,1;255,255,80,1;255,255,85,1;255,255,89,1;255,255,93,1;255,255,97,1;"
    strSQL = strSQL & "255,255,101,1;255,255,105,1;255,255,109,1;255,255,113,1;255,255,117,1;255,255,121,1;255,255,125,1;255,255,129,1;255,255,133,1;"
    strSQL = strSQL & "255,255,137,1;255,255,141,1;255,255,145,1;255,255,149,1;255,255,153,1;255,255,157,1;255,255,161,1;255,255,165,1;255,255,170,1;"
    strSQL = strSQL & "255,255,174,1;255,255,178,1;255,255,182,1;255,255,186,1;255,255,190,1;255,255,194,1;255,255,198,1;255,255,202,1;255,255,206,1;"
    strSQL = strSQL & "255,255,210,1;255,255,214,1;255,255,218,1;255,255,222,1;255,255,226,1;255,255,230,1;255,255,234,1;255,255,238,1;255,255,242,1;"
    strSQL = strSQL & "255,255,246,1;255,255,250,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (23,'hotter','0,0,0,1;19,19,0,1;38,37,0,1;56,55,1,1;74,71,2,1;90,86,4,1;"
    strSQL = strSQL & "106,100,6,1;121,112,8,1;135,124,10,1;148,135,13,1;161,145,15,1;173,154,18,1;184,162,21,1;194,169,24,1;203,175,27,1;212,181,31,1;"
    strSQL = strSQL & "220,185,34,1;227,189,37,1;233,192,41,1;239,194,44,1;244,196,47,1;248,196,51,1;251,197,54,1;253,196,57,1;255,195,60,1;255,193,62,1;"
    strSQL = strSQL & "255,190,65,1;255,188,67,1;255,185,70,1;255,183,72,1;255,180,75,1;255,178,77,1;255,175,80,1;255,173,82,1;255,170,85,1;255,168,87,1;"
    strSQL = strSQL & "255,165,90,1;255,163,92,1;255,160,95,1;255,158,97,1;255,155,100,1;255,153,102,1;255,150,105,1;255,148,107,1;255,145,110,1;255,143,112,1;"
    strSQL = strSQL & "255,140,115,1;255,138,117,1;255,135,120,1;255,133,122,1;255,130,125,1;255,128,128,1;255,125,130,1;255,122,133,1;255,120,135,1;"
    strSQL = strSQL & "255,117,138,1;255,115,140,1;255,112,143,1;255,110,145,1;255,107,148,1;255,105,150,1;255,102,153,1;255,100,155,1;255,97,158,1;"
    strSQL = strSQL & "255,95,160,1;255,92,163,1;255,90,165,1;255,87,168,1;255,85,170,1;255,82,173,1;255,80,175,1;255,77,178,1;255,75,180,1;255,72,183,1;"
    strSQL = strSQL & "255,70,185,1;255,67,188,1;255,65,190,1;255,62,193,1;255,60,195,1;255,57,198,1;255,55,200,1;255,52,203,1;255,50,205,1;255,47,208,1;"
    strSQL = strSQL & "255,45,210,1;255,42,213,1;255,40,215,1;255,37,218,1;255,35,220,1;255,32,223,1;255,30,225,1;255,27,228,1;255,25,230,1;255,22,233,1;"
    strSQL = strSQL & "255,20,235,1;255,17,238,1;255,15,240,1;255,12,243,1;255,10,245,1;255,7,248,1;255,5,250,1;255,2,253,1;255,0,255,1;253,2,255,1;"
    strSQL = strSQL & "250,5,255,1;248,7,255,1;245,10,255,1;243,12,255,1;240,15,255,1;238,17,255,1;235,20,255,1;233,22,255,1;230,25,255,1;228,27,255,1;"
    strSQL = strSQL & "225,30,255,1;223,32,255,1;220,35,255,1;218,37,255,1;215,40,255,1;213,42,255,1;210,45,255,1;208,47,255,1;205,50,255,1;203,52,255,1;"
    strSQL = strSQL & "200,55,255,1;198,57,255,1;195,60,255,1;193,62,255,1;190,65,255,1;188,67,255,1;185,70,255,1;183,72,255,1;180,75,255,1;178,77,255,1;"
    strSQL = strSQL & "175,80,255,1;173,82,255,1;170,85,255,1;168,87,255,1;165,90,255,1;163,92,255,1;160,95,255,1;158,97,255,1;155,100,255,1;153,102,255,1;"
    strSQL = strSQL & "150,105,255,1;148,107,255,1;145,110,255,1;143,112,255,1;140,115,255,1;138,117,255,1;135,120,255,1;133,122,255,1;130,125,255,1;"
    strSQL = strSQL & "127,128,255,1;125,130,255,1;122,133,255,1;120,135,255,1;117,138,255,1;115,140,255,1;112,143,255,1;110,145,255,1;107,148,255,1;"
    strSQL = strSQL & "105,150,255,1;102,153,255,1;100,155,255,1;97,158,255,1;95,160,255,1;92,163,255,1;90,165,255,1;87,168,255,1;85,170,255,1;82,173,255,1;"
    strSQL = strSQL & "80,175,255,1;77,178,255,1;75,180,255,1;72,183,255,1;70,185,255,1;67,188,255,1;65,190,255,1;62,193,255,1;60,195,255,1;57,198,255,1;"
    strSQL = strSQL & "55,200,255,1;52,203,255,1;50,205,255,1;47,208,255,1;45,210,255,1;42,213,255,1;40,215,255,1;37,218,255,1;35,220,255,1;32,223,255,1;"
    strSQL = strSQL & "30,225,255,1;27,228,255,1;25,230,255,1;22,233,255,1;20,235,255,1;17,238,255,1;15,240,255,1;12,243,255,1;10,245,255,1;7,248,255,1;"
    strSQL = strSQL & "5,250,255,1;2,253,255,1;0,255,255,1;5,255,255,1;10,255,255,1;15,255,255,1;20,255,255,1;25,255,255,1;30,255,255,1;35,255,255,1;"
    strSQL = strSQL & "40,255,255,1;45,255,255,1;50,255,255,1;55,255,255,1;60,255,255,1;65,255,255,1;70,255,255,1;75,255,255,1;80,255,255,1;85,255,255,1;"
    strSQL = strSQL & "90,255,255,1;95,255,255,1;100,255,255,1;105,255,255,1;110,255,255,1;115,255,255,1;120,255,255,1;125,255,255,1;130,255,255,1;"
    strSQL = strSQL & "135,255,255,1;141,255,255,1;147,255,255,1;153,255,255,1;160,255,255,1;166,255,255,1;173,255,255,1;180,255,255,1;186,255,255,1;"
    strSQL = strSQL & "193,255,255,1;200,255,255,1;206,255,255,1;212,255,255,1;218,255,255,1;224,255,255,1;229,255,255,1;234,255,255,1;239,255,255,1;"
    strSQL = strSQL & "243,255,255,1;247,255,255,1;250,255,255,1;252,255,255,1;254,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (24,'Ice','0,0,0,1;0,0,0,1;0,0,0,1;0,1,1,1;0,2,2,1;0,3,3,1;0,5,5,1;"
    strSQL = strSQL & "0,6,6,1;0,8,8,1;0,10,10,1;0,12,12,1;0,14,14,1;0,17,17,1;0,19,19,1;0,22,22,1;0,25,25,1;0,27,27,1;0,30,30,1;0,33,33,1;0,35,35,1;"
    strSQL = strSQL & "0,38,38,1;0,40,40,1;0,43,43,1;0,45,45,1;0,48,48,1;0,50,50,1;0,52,52,1;0,54,54,1;0,56,56,1;0,58,58,1;0,60,60,1;0,62,62,1;0,64,64,1;"
    strSQL = strSQL & "0,66,66,1;0,68,68,1;0,70,70,1;0,72,72,1;0,74,74,1;0,76,76,1;0,78,78,1;0,80,80,1;0,82,82,1;0,84,84,1;0,86,86,1;0,88,88,1;0,90,90,1;"
    strSQL = strSQL & "0,92,92,1;0,94,94,1;0,96,96,1;0,98,98,1;0,100,100,1;0,102,102,1;0,104,104,1;0,106,106,1;0,108,108,1;0,110,110,1;0,112,112,1;"
    strSQL = strSQL & "0,114,114,1;0,116,116,1;0,118,118,1;0,120,120,1;0,122,122,1;0,124,124,1;0,126,126,1;0,128,128,1;0,130,130,1;0,132,132,1;0,134,134,1;"
    strSQL = strSQL & "0,136,136,1;0,138,138,1;0,140,140,1;0,142,142,1;0,144,144,1;0,146,146,1;0,148,148,1;0,150,150,1;0,152,152,1;0,154,154,1;0,156,156,1;"
    strSQL = strSQL & "0,158,158,1;0,160,160,1;0,162,162,1;0,164,164,1;0,166,166,1;0,168,168,1;0,170,170,1;0,172,172,1;0,174,174,1;0,176,176,1;0,178,178,1;"
    strSQL = strSQL & "0,180,180,1;0,182,182,1;0,184,184,1;0,186,186,1;0,188,188,1;0,190,190,1;0,192,192,1;0,194,194,1;0,196,196,1;0,198,198,1;0,200,200,1;"
    strSQL = strSQL & "0,202,202,1;0,204,204,1;0,206,206,1;0,208,208,1;0,210,210,1;0,212,212,1;0,214,214,1;0,216,216,1;0,218,218,1;0,220,220,1;0,222,222,1;"
    strSQL = strSQL & "0,224,224,1;0,226,226,1;0,228,228,1;0,230,230,1;0,232,232,1;0,234,234,1;0,236,236,1;0,238,238,1;0,240,240,1;0,242,242,1;0,244,244,1;"
    strSQL = strSQL & "0,246,246,1;0,248,248,1;0,250,250,1;0,252,252,1;0,254,254,1;1,254,255,1;3,252,255,1;5,250,255,1;7,248,255,1;9,246,255,1;11,244,255,1;"
    strSQL = strSQL & "13,242,255,1;15,240,255,1;17,238,255,1;19,236,255,1;21,234,255,1;23,232,255,1;25,230,255,1;27,228,255,1;29,226,255,1;31,224,255,1;"
    strSQL = strSQL & "33,222,255,1;35,220,255,1;37,218,255,1;39,216,255,1;41,214,255,1;43,212,255,1;45,210,255,1;47,208,255,1;49,206,255,1;51,204,255,1;"
    strSQL = strSQL & "53,202,255,1;55,200,255,1;57,198,255,1;59,196,255,1;61,194,255,1;63,192,255,1;65,190,255,1;67,188,255,1;69,186,255,1;71,184,255,1;"
    strSQL = strSQL & "73,182,255,1;75,180,255,1;77,178,255,1;79,176,255,1;81,174,255,1;83,172,255,1;85,170,255,1;87,168,255,1;89,166,255,1;91,164,255,1;"
    strSQL = strSQL & "93,162,255,1;95,160,255,1;97,158,255,1;99,156,255,1;101,154,255,1;103,152,255,1;105,150,255,1;107,148,255,1;109,146,255,1;111,144,255,1;"
    strSQL = strSQL & "113,142,255,1;115,140,255,1;117,138,255,1;119,136,255,1;121,134,255,1;123,132,255,1;125,130,255,1;127,128,255,1;129,126,255,1;"
    strSQL = strSQL & "131,124,255,1;133,122,255,1;135,120,255,1;137,118,255,1;139,116,255,1;141,114,255,1;143,112,255,1;145,110,255,1;147,108,255,1;"
    strSQL = strSQL & "149,106,255,1;151,104,255,1;153,102,255,1;155,100,255,1;157,98,255,1;159,96,255,1;161,94,255,1;163,92,255,1;165,90,255,1;"
    strSQL = strSQL & "167,88,255,1;169,86,255,1;171,84,255,1;173,82,255,1;175,80,255,1;177,78,255,1;179,76,255,1;181,74,255,1;183,72,255,1;"
    strSQL = strSQL & "185,70,255,1;187,68,255,1;189,66,255,1;191,64,255,1;193,62,255,1;195,60,255,1;197,58,255,1;199,56,255,1;201,54,255,1;"
    strSQL = strSQL & "203,52,255,1;205,50,255,1;207,48,255,1;210,48,255,1;212,48,255,1;215,48,255,1;217,50,255,1;220,52,255,1;222,55,255,1;"
    strSQL = strSQL & "225,58,255,1;228,63,255,1;230,68,255,1;233,74,255,1;236,81,255,1;238,89,255,1;241,97,255,1;243,107,255,1;245,117,255,1;"
    strSQL = strSQL & "247,129,255,1;249,141,255,1;250,154,255,1;252,169,255,1;253,184,255,1;254,200,255,1;255,218,255,1;255,236,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (25,'ge_256','0,0,0,1;0,1,1,1;0,1,2,1;0,2,3,1;0,2,4,1;0,3,5,1;0,3,6,1;"
    strSQL = strSQL & "0,4,7,1;0,4,8,1;0,5,9,1;0,5,10,1;0,6,11,1;0,6,12,1;0,7,13,1;0,7,14,1;0,8,15,1;0,8,16,1;0,9,17,1;0,9,18,1;0,10,19,1;0,10,20,1;"
    strSQL = strSQL & "0,11,21,1;0,11,22,1;0,12,23,1;0,12,24,1;0,13,25,1;0,13,26,1;0,14,27,1;0,14,28,1;0,15,29,1;0,15,30,1;0,16,31,1;0,16,32,1;0,17,33,1;"
    strSQL = strSQL & "0,17,34,1;0,18,35,1;0,18,36,1;0,19,37,1;0,19,38,1;0,20,39,1;0,20,40,1;0,21,41,1;0,21,42,1;0,22,43,1;0,22,44,1;0,23,45,1;0,23,46,1;"
    strSQL = strSQL & "0,24,47,1;0,24,48,1;0,25,49,1;0,25,50,1;0,26,51,1;0,26,52,1;0,27,53,1;0,27,54,1;0,28,55,1;0,28,56,1;0,29,57,1;0,29,58,1;0,30,59,1;"
    strSQL = strSQL & "0,30,60,1;0,31,61,1;0,31,62,1;0,32,63,1;0,32,64,1;1,31,65,1;3,31,66,1;5,30,67,1;7,30,68,1;9,29,69,1;11,29,70,1;13,28,71,1;15,28,72,1;"
    strSQL = strSQL & "17,27,73,1;19,27,74,1;21,26,75,1;23,26,76,1;25,25,77,1;27,25,78,1;29,24,79,1;31,24,80,1;33,23,81,1;35,23,82,1;37,22,83,1;39,22,84,1;"
    strSQL = strSQL & "41,21,85,1;43,21,86,1;45,20,87,1;47,20,88,1;49,19,89,1;51,19,90,1;53,18,91,1;55,18,92,1;57,17,93,1;59,17,94,1;61,16,95,1;63,16,96,1;"
    strSQL = strSQL & "65,15,97,1;67,15,98,1;69,14,99,1;71,14,100,1;73,13,101,1;75,13,102,1;77,12,103,1;79,12,104,1;81,11,105,1;83,11,106,1;85,10,107,1;"
    strSQL = strSQL & "87,10,108,1;89,9,109,1;91,9,110,1;93,8,111,1;95,8,112,1;97,7,113,1;99,7,114,1;101,6,115,1;103,6,116,1;105,5,117,1;107,5,118,1;"
    strSQL = strSQL & "109,4,119,1;111,4,120,1;113,3,121,1;115,3,122,1;117,2,123,1;119,2,124,1;121,1,125,1;123,1,126,1;125,0,127,1;127,2,127,1;129,3,125,1;"
    strSQL = strSQL & "131,5,123,1;133,6,121,1;135,8,119,1;137,9,117,1;139,11,115,1;141,12,113,1;143,14,111,1;145,15,109,1;147,17,107,1;149,18,105,1;"
    strSQL = strSQL & "151,20,103,1;153,21,101,1;155,23,99,1;157,24,97,1;159,26,95,1;161,27,93,1;163,29,91,1;165,30,89,1;167,32,87,1;169,33,85,1;171,35,83,1;"
    strSQL = strSQL & "173,36,81,1;175,38,79,1;177,39,77,1;179,41,75,1;181,42,73,1;183,44,71,1;185,45,69,1;187,47,67,1;189,48,65,1;191,50,63,1;193,51,61,1;"
    strSQL = strSQL & "195,53,59,1;197,54,57,1;199,56,55,1;201,57,53,1;203,59,51,1;205,60,49,1;207,62,47,1;209,63,45,1;211,65,43,1;213,66,41,1;215,68,39,1;"
    strSQL = strSQL & "217,69,37,1;219,71,35,1;221,72,33,1;223,74,31,1;225,75,29,1;227,77,27,1;229,78,25,1;231,80,23,1;233,81,21,1;235,83,19,1;237,84,17,1;"
    strSQL = strSQL & "239,86,15,1;241,87,13,1;243,89,11,1;245,90,9,1;247,92,7,1;249,93,5,1;251,95,3,1;253,95,0,1;255,98,3,1;255,100,7,1;255,103,11,1;"
    strSQL = strSQL & "255,105,15,1;255,108,19,1;255,110,23,1;255,113,27,1;255,115,31,1;255,118,35,1;255,120,39,1;255,123,43,1;255,125,47,1;255,128,51,1;"
    strSQL = strSQL & "255,130,55,1;255,133,59,1;255,135,63,1;255,138,67,1;255,140,71,1;255,143,75,1;255,145,79,1;255,148,83,1;255,150,87,1;255,153,91,1;"
    strSQL = strSQL & "255,155,95,1;255,158,99,1;255,160,103,1;255,163,107,1;255,165,111,1;255,168,115,1;255,170,119,1;255,173,123,1;255,175,127,1;"
    strSQL = strSQL & "255,178,131,1;255,180,135,1;255,183,139,1;255,185,143,1;255,188,147,1;255,190,151,1;255,193,155,1;255,195,159,1;255,198,163,1;"
    strSQL = strSQL & "255,200,167,1;255,203,171,1;255,205,175,1;255,208,179,1;255,210,183,1;255,213,187,1;255,215,191,1;255,218,195,1;255,220,199,1;"
    strSQL = strSQL & "255,223,203,1;255,225,207,1;255,228,211,1;255,230,215,1;255,233,219,1;255,235,223,1;255,238,227,1;255,240,231,1;255,243,235,1;"
    strSQL = strSQL & "255,245,239,1;255,248,243,1;255,250,247,1;255,253,251,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (26,'Isocontour','0,2,0,1;0,2,6,1;0,2,73,1;0,2,109,1;0,2,146,1;0,2,182,1;"
    strSQL = strSQL & "0,0,255,1;0,0,246,1;0,0,237,1;0,0,227,1;0,0,218,1;0,0,208,1;0,0,199,1;0,0,190,1;0,0,180,1;0,0,171,1;0,0,160,1;0,0,152,1;0,0,143,1;"
    strSQL = strSQL & "0,0,133,1;0,0,124,1;0,0,114,1;0,0,105,1;0,0,96,1;0,0,86,1;0,0,77,1;0,0,67,1;0,0,58,1;0,0,49,1;0,0,39,1;0,0,30,1;0,0,20,1;0,0,0,1;"
    strSQL = strSQL & "0,8,0,1;0,16,0,1;0,24,0,1;0,32,0,1;0,40,0,1;0,48,0,1;0,56,0,1;0,64,0,1;0,72,0,1;0,80,0,1;0,88,0,1;0,96,0,1;0,104,0,1;0,112,0,1;0,120,0,1;"
    strSQL = strSQL & "0,128,0,1;0,136,0,1;0,144,0,1;0,152,0,1;0,160,0,1;0,168,0,1;0,176,0,1;0,184,0,1;0,192,0,1;0,200,0,1;0,208,0,1;0,216,0,1;0,224,0,1;"
    strSQL = strSQL & "0,232,0,1;0,240,0,1;0,255,0,1;0,0,0,1;7,5,0,1;14,10,0,1;22,15,0,1;29,20,0,1;37,25,0,1;44,30,0,1;52,35,0,1;59,40,0,1;67,45,0,1;74,50,0,1;"
    strSQL = strSQL & "82,55,0,1;89,60,0,1;97,65,0,1;104,70,0,1;112,75,0,1;126,85,0,1;133,89,0,1;141,93,0,1;149,98,0,1;157,102,0,1;165,106,0,1;173,111,0,1;"
    strSQL = strSQL & "181,115,0,1;189,120,0,1;197,124,0,1;205,128,0,1;213,133,0,1;221,137,0,1;229,141,0,1;237,146,0,1;252,154,0,1;0,0,0,1;1,0,0,1;2,0,0,1;"
    strSQL = strSQL & "3,0,0,1;4,0,0,1;5,0,0,1;6,0,0,1;7,0,0,1;8,0,0,1;9,0,0,1;10,0,0,1;11,0,0,1;12,0,0,1;14,0,0,1;16,0,0,1;18,0,0,1;20,0,0,1;22,0,0,1;"
    strSQL = strSQL & "24,0,0,1;27,0,0,1;29,0,0,1;31,0,0,1;33,0,0,1;35,0,0,1;37,0,0,1;41,0,0,1;44,0,0,1;47,0,0,1;50,0,0,1;53,0,0,1;56,0,0,1;59,0,0,1;"
    strSQL = strSQL & "62,1,0,1;65,1,0,1;68,1,0,1;71,1,0,1;74,1,0,1;77,1,0,1;82,1,0,1;85,1,0,1;89,1,0,1;92,1,0,1;96,1,0,1;100,1,0,1;103,1,0,1;107,1,1,1;"
    strSQL = strSQL & "110,1,1,1;114,1,1,1;118,1,1,1;121,1,1,1;128,1,1,1;131,1,1,1;135,1,1,1;139,1,1,1;142,1,1,1;146,1,1,1;150,1,1,1;153,1,1,1;157,1,1,1;"
    strSQL = strSQL & "161,1,1,1;164,1,1,1;168,1,1,1;175,1,1,1;178,1,1,1;181,1,1,1;184,1,1,1;187,1,1,1;191,1,1,1;194,1,1,1;197,2,1,1;200,2,1,1;204,2,1,1;"
    strSQL = strSQL & "207,2,1,1;210,2,1,1;216,2,1,1;218,2,1,1;220,2,1,1;222,2,1,1;224,2,1,1;227,2,1,1;229,2,1,1;231,2,2,1;233,2,2,1;236,2,2,1;238,2,2,1;"
    strSQL = strSQL & "240,2,2,1;244,2,2,1;244,2,2,1;245,2,2,1;246,2,2,1;247,2,2,1;247,2,2,1;248,2,2,1;249,2,2,1;250,2,2,1;251,2,2,1;251,2,2,1;252,2,2,1;"
    strSQL = strSQL & "253,2,2,1;254,2,1,1;254,2,3,1;254,2,5,1;254,2,7,1;254,2,9,1;254,2,11,1;254,2,13,1;254,2,15,1;254,2,17,1;254,2,20,1;254,2,22,1;"
    strSQL = strSQL & "254,2,24,1;254,2,26,1;254,2,28,1;254,2,30,1;254,2,32,1;254,2,36,1;254,2,36,1;254,3,38,1;254,4,40,1;255,5,42,1;255,7,45,1;255,9,47,1;"
    strSQL = strSQL & "255,11,50,1;255,13,53,1;255,17,58,1;255,20,60,1;255,23,63,1;255,26,66,1;255,29,68,1;255,34,73,1;255,38,77,1;255,42,81,1;255,47,85,1;"
    strSQL = strSQL & "255,55,92,1;255,59,95,1;255,64,99,1;255,69,103,1;255,73,106,1;255,82,113,1;255,88,118,1;255,94,123,1;255,100,128,1;255,112,138,1;"
    strSQL = strSQL & "255,118,143,1;255,125,148,1;255,132,154,1;255,145,164,1;255,152,170,1;255,159,176,1;255,166,182,1;255,180,193,1;255,187,199,1;"
    strSQL = strSQL & "255,195,205,1;255,202,211,1;255,217,223,1;255,223,228,1;255,230,234,1;255,236,239,1;255,243,245,1;255,249,250,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (27,'linear','0,0,0,1;0,0,19,1;0,0,37,1;0,1,55,1;0,2,71,1;0,3,86,1;"
    strSQL = strSQL & "0,5,101,1;0,6,114,1;0,8,126,1;0,10,138,1;0,12,148,1;0,14,158,1;0,17,166,1;0,19,174,1;0,22,181,1;0,25,187,1;0,27,192,1;0,30,197,1;"
    strSQL = strSQL & "0,33,200,1;0,35,203,1;0,38,205,1;0,40,207,1;0,43,207,1;0,45,207,1;0,48,207,1;0,50,205,1;0,52,203,1;0,54,201,1;0,56,199,1;0,58,197,1;"
    strSQL = strSQL & "0,60,195,1;0,62,193,1;0,64,191,1;0,66,189,1;0,68,187,1;0,70,185,1;0,72,183,1;0,74,181,1;0,76,179,1;0,78,177,1;0,80,175,1;0,82,173,1;"
    strSQL = strSQL & "0,84,171,1;0,86,169,1;0,88,167,1;0,90,165,1;0,92,163,1;0,94,161,1;0,96,159,1;0,98,157,1;0,100,155,1;0,102,153,1;0,104,151,1;"
    strSQL = strSQL & "0,106,149,1;0,108,147,1;0,110,145,1;0,112,143,1;0,114,141,1;0,116,139,1;0,118,137,1;0,120,135,1;0,122,133,1;0,124,131,1;0,126,129,1;"
    strSQL = strSQL & "0,128,127,1;0,130,125,1;0,132,123,1;0,134,121,1;0,136,119,1;0,138,117,1;0,140,115,1;0,142,113,1;0,144,111,1;0,146,109,1;0,148,107,1;"
    strSQL = strSQL & "0,150,105,1;0,152,103,1;0,154,101,1;0,156,99,1;0,158,97,1;0,160,95,1;0,162,93,1;0,164,91,1;0,166,89,1;0,168,87,1;0,170,85,1;0,172,83,1;"
    strSQL = strSQL & "0,174,81,1;0,176,79,1;0,178,77,1;0,180,75,1;0,182,73,1;0,184,71,1;0,186,69,1;0,188,67,1;0,190,65,1;0,192,63,1;0,194,61,1;0,196,59,1;"
    strSQL = strSQL & "0,198,57,1;0,200,55,1;0,202,53,1;0,204,51,1;0,206,49,1;0,208,47,1;0,210,45,1;0,212,43,1;0,214,41,1;0,216,39,1;0,218,37,1;0,220,35,1;"
    strSQL = strSQL & "0,222,33,1;0,224,31,1;0,226,29,1;0,228,27,1;0,230,25,1;0,232,23,1;0,234,21,1;0,236,19,1;0,238,17,1;0,240,15,1;0,242,13,1;0,244,11,1;"
    strSQL = strSQL & "0,246,9,1;0,248,7,1;0,250,5,1;0,252,3,1;0,254,1,1;1,254,0,1;3,252,0,1;5,250,0,1;7,248,0,1;9,246,0,1;11,244,0,1;13,242,0,1;15,240,0,1;"
    strSQL = strSQL & "17,238,0,1;19,236,0,1;21,234,0,1;23,232,0,1;25,230,0,1;27,228,0,1;29,226,0,1;31,224,0,1;33,222,0,1;35,220,0,1;37,218,0,1;39,216,0,1;"
    strSQL = strSQL & "41,214,0,1;43,212,0,1;45,210,0,1;47,208,0,1;49,206,0,1;51,204,0,1;53,202,0,1;55,200,0,1;57,198,0,1;59,196,0,1;61,194,0,1;63,192,0,1;"
    strSQL = strSQL & "65,190,0,1;67,188,0,1;69,186,0,1;71,184,0,1;73,182,0,1;75,180,0,1;77,178,0,1;79,176,0,1;81,174,0,1;83,172,0,1;85,170,0,1;87,168,0,1;"
    strSQL = strSQL & "89,166,0,1;91,164,0,1;93,162,0,1;95,160,0,1;97,158,0,1;99,156,0,1;101,154,0,1;103,152,0,1;105,150,0,1;107,148,0,1;109,146,0,1;"
    strSQL = strSQL & "111,144,0,1;113,142,0,1;115,140,0,1;117,138,0,1;119,136,0,1;121,134,0,1;123,132,0,1;125,130,0,1;127,128,0,1;129,126,0,1;131,124,0,1;"
    strSQL = strSQL & "133,122,0,1;135,120,0,1;137,118,0,1;139,116,0,1;141,114,0,1;143,112,0,1;145,110,0,1;147,108,0,1;149,106,0,1;151,104,0,1;153,102,0,1;"
    strSQL = strSQL & "155,100,0,1;157,98,0,1;159,96,0,1;161,94,0,1;163,92,0,1;165,90,0,1;167,88,0,1;169,86,0,1;171,84,0,1;173,82,0,1;175,80,0,1;177,78,0,1;"
    strSQL = strSQL & "179,76,0,1;181,74,0,1;183,72,0,1;185,70,0,1;187,68,0,1;189,66,0,1;191,64,0,1;193,62,0,1;195,60,0,1;197,58,0,1;199,56,0,1;201,54,0,1;"
    strSQL = strSQL & "203,52,0,1;205,50,0,1;207,48,0,1;210,48,2,1;212,48,4,1;215,48,7,1;217,50,11,1;220,52,16,1;222,55,22,1;225,58,28,1;228,63,35,1;"
    strSQL = strSQL & "230,68,43,1;233,74,52,1;236,81,61,1;238,89,71,1;241,97,82,1;243,107,94,1;245,117,107,1;247,129,120,1;249,141,134,1;250,154,149,1;"
    strSQL = strSQL & "252,169,165,1;253,184,181,1;254,200,199,1;255,218,217,1;255,236,236,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (28,'MMC','0,0,2,1;0,2,4,1;0,4,6,1;0,6,8,1;0,8,10,1;0,10,12,1;0,12,14,1;"
    strSQL = strSQL & "0,14,16,1;0,16,18,1;0,18,20,1;0,20,22,1;0,22,24,1;0,24,26,1;0,26,28,1;0,28,30,1;0,30,32,1;0,33,34,1;0,35,36,1;0,37,38,1;0,39,40,1;"
    strSQL = strSQL & "0,41,42,1;0,43,44,1;0,45,46,1;0,47,48,1;0,49,50,1;0,51,52,1;0,53,54,1;0,55,56,1;0,57,58,1;0,59,60,1;0,61,62,1;0,63,64,1;0,65,66,1;"
    strSQL = strSQL & "0,67,68,1;0,69,70,1;0,71,72,1;0,73,74,1;0,75,76,1;0,77,78,1;0,79,80,1;0,81,82,1;0,83,84,1;0,85,86,1;0,87,88,1;0,89,90,1;0,91,92,1;"
    strSQL = strSQL & "0,93,94,1;0,95,96,1;0,98,98,1;0,100,100,1;0,102,102,1;0,104,104,1;0,106,106,1;0,108,108,1;0,110,110,1;0,112,112,1;0,114,114,1;"
    strSQL = strSQL & "0,116,116,1;0,118,118,1;0,120,120,1;0,122,122,1;0,124,124,1;0,126,126,1;2,128,128,1;4,126,129,1;6,124,131,1;8,122,133,1;10,120,135,1;"
    strSQL = strSQL & "12,118,137,1;14,116,139,1;16,114,141,1;18,112,143,1;20,110,145,1;22,108,147,1;24,106,149,1;26,104,151,1;28,102,153,1;30,100,155,1;"
    strSQL = strSQL & "32,98,157,1;34,96,159,1;36,95,161,1;38,93,163,1;40,91,165,1;42,89,167,1;43,87,169,1;45,85,171,1;47,83,173,1;49,81,175,1;51,79,177,1;"
    strSQL = strSQL & "53,77,179,1;55,75,181,1;57,73,183,1;59,71,185,1;61,69,187,1;63,67,189,1;65,65,191,1;67,63,193,1;69,61,195,1;71,59,197,1;73,57,199,1;"
    strSQL = strSQL & "75,55,201,1;77,53,203,1;79,51,205,1;81,49,207,1;83,47,209,1;85,45,211,1;87,43,213,1;89,41,215,1;91,39,217,1;93,37,219,1;95,35,221,1;"
    strSQL = strSQL & "97,33,223,1;99,32,225,1;101,30,227,1;103,28,229,1;105,26,231,1;107,24,233,1;109,22,235,1;111,20,237,1;113,18,239,1;115,16,241,1;"
    strSQL = strSQL & "117,14,243,1;119,12,245,1;121,10,247,1;123,8,249,1;125,6,251,1;127,4,253,1;128,2,255,1;130,0,251,1;132,2,247,1;134,4,243,1;"
    strSQL = strSQL & "136,6,239,1;138,8,235,1;140,10,231,1;142,12,227,1;144,14,223,1;146,16,219,1;148,18,215,1;150,20,211,1;152,22,207,1;154,24,203,1;"
    strSQL = strSQL & "156,26,199,1;158,28,195,1;160,30,191,1;162,32,187,1;164,34,183,1;166,36,179,1;168,38,175,1;170,40,171,1;172,42,167,1;174,44,163,1;"
    strSQL = strSQL & "176,46,159,1;178,48,155,1;180,50,151,1;182,52,147,1;184,54,143,1;186,56,139,1;188,58,135,1;190,60,131,1;192,62,128,1;194,64,124,1;"
    strSQL = strSQL & "196,66,120,1;198,68,116,1;200,70,112,1;202,72,108,1;204,74,104,1;206,76,100,1;208,78,96,1;210,80,92,1;212,82,88,1;213,84,84,1;"
    strSQL = strSQL & "215,86,80,1;217,88,76,1;219,90,72,1;221,92,68,1;223,94,64,1;225,96,60,1;227,98,56,1;229,100,52,1;231,102,48,1;233,104,44,1;"
    strSQL = strSQL & "235,106,40,1;237,108,36,1;239,110,32,1;241,112,28,1;243,114,24,1;245,116,20,1;247,118,16,1;248,120,12,1;251,122,8,1;253,124,4,1;"
    strSQL = strSQL & "255,126,0,1;255,129,4,1;255,131,8,1;255,133,12,1;255,135,16,1;255,137,20,1;255,139,24,1;255,141,28,1;255,143,32,1;255,145,36,1;"
    strSQL = strSQL & "255,147,40,1;255,149,45,1;255,151,49,1;255,153,53,1;255,155,57,1;255,157,61,1;255,159,65,1;255,161,69,1;255,163,73,1;255,165,77,1;"
    strSQL = strSQL & "255,167,81,1;255,169,85,1;255,171,89,1;255,173,93,1;255,175,97,1;255,177,101,1;255,179,105,1;255,181,109,1;255,183,113,1;"
    strSQL = strSQL & "255,185,117,1;255,187,121,1;255,189,125,1;255,191,130,1;255,193,134,1;255,195,138,1;255,197,142,1;255,199,146,1;255,201,150,1;"
    strSQL = strSQL & "255,203,154,1;255,205,158,1;255,207,162,1;255,209,166,1;255,211,170,1;255,213,174,1;255,215,178,1;255,217,182,1;255,219,186,1;"
    strSQL = strSQL & "255,221,190,1;255,223,194,1;255,225,198,1;255,227,202,1;255,229,206,1;255,231,210,1;255,233,215,1;255,235,219,1;255,237,223,1;"
    strSQL = strSQL & "255,239,227,1;255,241,231,1;255,243,235,1;255,245,239,1;255,247,243,1;255,249,247,1;255,251,251,1;255,253,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (29,'multi','0,0,0,1;0,0,0,1;0,0,1,1;0,0,4,1;0,0,6,1;0,0,10,1;0,0,15,1;"
    strSQL = strSQL & "0,0,19,1;0,0,25,1;0,0,31,1;0,0,37,1;0,0,44,1;0,0,52,1;0,0,59,1;0,0,67,1;0,0,75,1;0,0,83,1;0,0,91,1;0,0,99,1;0,0,107,1;0,0,114,1;"
    strSQL = strSQL & "0,0,122,1;0,0,130,1;0,0,137,1;0,0,144,1;0,0,150,1;0,0,156,1;0,0,162,1;0,0,168,1;0,0,174,1;0,0,180,1;0,0,186,1;0,0,192,1;0,0,198,1;"
    strSQL = strSQL & "0,0,204,1;0,0,210,1;0,0,216,1;0,0,222,1;0,0,228,1;0,0,234,1;0,0,240,1;0,0,246,1;0,0,252,1;0,3,255,1;0,9,255,1;0,15,255,1;0,21,255,1;"
    strSQL = strSQL & "0,27,255,1;0,33,255,1;0,39,255,1;0,45,255,1;0,51,255,1;0,57,255,1;0,63,255,1;0,69,255,1;0,75,255,1;0,81,255,1;0,87,255,1;0,93,255,1;"
    strSQL = strSQL & "0,99,255,1;0,105,255,1;0,111,255,1;0,117,255,1;0,123,255,1;0,129,255,1;0,135,255,1;0,141,255,1;0,147,255,1;0,153,255,1;0,159,255,1;"
    strSQL = strSQL & "0,165,255,1;0,171,255,1;0,177,255,1;0,183,255,1;0,189,255,1;0,195,255,1;0,201,255,1;0,207,255,1;0,213,255,1;0,219,255,1;0,225,255,1;"
    strSQL = strSQL & "0,231,255,1;0,237,255,1;0,243,255,1;0,249,255,1;0,255,255,1;0,255,249,1;0,255,243,1;0,255,237,1;0,255,231,1;0,255,225,1;0,255,219,1;"
    strSQL = strSQL & "0,255,213,1;0,255,207,1;0,255,201,1;0,255,195,1;0,255,189,1;0,255,183,1;0,255,177,1;0,255,171,1;0,255,165,1;0,255,159,1;0,255,153,1;"
    strSQL = strSQL & "0,255,147,1;0,255,141,1;0,255,135,1;0,255,129,1;0,255,123,1;0,255,117,1;0,255,111,1;0,255,105,1;0,255,99,1;0,255,93,1;0,255,87,1;"
    strSQL = strSQL & "0,255,81,1;0,255,75,1;0,255,69,1;0,255,63,1;0,255,57,1;0,255,51,1;0,255,45,1;0,255,39,1;0,255,33,1;0,255,27,1;0,255,21,1;0,255,15,1;"
    strSQL = strSQL & "0,255,9,1;0,255,3,1;3,255,0,1;9,255,0,1;15,255,0,1;21,255,0,1;27,255,0,1;33,255,0,1;39,255,0,1;45,255,0,1;51,255,0,1;57,255,0,1;"
    strSQL = strSQL & "63,255,0,1;69,255,0,1;75,255,0,1;81,255,0,1;87,255,0,1;93,255,0,1;99,255,0,1;105,255,0,1;111,255,0,1;117,255,0,1;123,255,0,1;"
    strSQL = strSQL & "129,255,0,1;135,255,0,1;141,255,0,1;147,255,0,1;153,255,0,1;159,255,0,1;165,255,0,1;171,255,0,1;177,255,0,1;183,255,0,1;189,255,0,1;"
    strSQL = strSQL & "195,255,0,1;201,255,0,1;207,255,0,1;213,255,0,1;219,255,0,1;225,255,0,1;231,255,0,1;237,255,0,1;243,255,0,1;249,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,249,0,1;255,243,0,1;255,237,0,1;255,231,0,1;255,225,0,1;255,219,0,1;255,213,0,1;255,207,0,1;255,201,0,1;255,195,0,1;255,189,0,1;"
    strSQL = strSQL & "255,183,0,1;255,177,0,1;255,171,0,1;255,165,0,1;255,159,0,1;255,153,0,1;255,147,0,1;255,141,0,1;255,135,0,1;255,129,0,1;255,123,0,1;"
    strSQL = strSQL & "255,117,0,1;255,111,0,1;255,105,0,1;255,99,0,1;255,93,0,1;255,87,0,1;255,81,0,1;255,75,0,1;255,69,0,1;255,63,0,1;255,57,0,1;"
    strSQL = strSQL & "255,51,0,1;255,45,0,1;255,39,0,1;255,33,0,1;255,27,0,1;255,21,0,1;255,15,0,1;255,9,0,1;255,3,0,1;255,3,3,1;255,9,9,1;255,15,15,1;"
    strSQL = strSQL & "255,21,21,1;255,27,27,1;255,33,33,1;255,39,39,1;255,45,45,1;255,51,51,1;255,57,57,1;255,63,63,1;255,69,69,1;255,75,75,1;255,81,81,1;"
    strSQL = strSQL & "255,87,87,1;255,93,93,1;255,99,99,1;255,105,105,1;255,111,111,1;255,118,118,1;255,125,125,1;255,133,133,1;255,141,141,1;255,148,148,1;"
    strSQL = strSQL & "255,156,156,1;255,164,164,1;255,172,172,1;255,180,180,1;255,188,188,1;255,196,196,1;255,203,203,1;255,211,211,1;255,218,218,1;"
    strSQL = strSQL & "255,224,224,1;255,230,230,1;255,236,236,1;255,240,240,1;255,245,245,1;255,249,249,1;255,251,251,1;255,254,254,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (30,'Nathalie','0,0,101,1;0,0,105,1;0,0,110,1;0,0,115,1;0,0,120,1;"
    strSQL = strSQL & "0,0,125,1;0,0,129,1;0,0,134,1;0,0,139,1;0,0,144,1;0,0,149,1;0,0,153,1;0,0,158,1;0,0,163,1;0,0,168,1;0,0,173,1;0,0,178,1;0,0,182,1;"
    strSQL = strSQL & "0,0,187,1;0,8,192,1;0,16,197,1;0,24,202,1;0,32,206,1;0,40,211,1;0,48,216,1;0,56,221,1;0,65,226,1;0,73,230,1;0,81,235,1;0,89,240,1;"
    strSQL = strSQL & "0,97,245,1;0,105,250,1;0,113,255,1;0,121,255,1;0,121,251,1;0,121,247,1;0,121,243,1;0,121,239,1;0,121,235,1;0,121,231,1;0,121,227,1;"
    strSQL = strSQL & "0,121,223,1;0,121,219,1;0,121,215,1;0,121,211,1;0,121,207,1;0,121,203,1;0,121,199,1;0,121,195,1;0,121,191,1;0,121,187,1;0,128,183,1;"
    strSQL = strSQL & "0,135,179,1;0,142,175,1;0,149,171,1;0,155,167,1;0,162,163,1;0,169,159,1;0,176,155,1;0,183,151,1;0,190,147,1;0,196,143,1;0,203,139,1;"
    strSQL = strSQL & "0,210,135,1;0,217,131,1;0,224,128,1;0,224,118,1;0,224,109,1;0,224,100,1;0,224,91,1;0,224,82,1;0,224,73,1;0,224,63,1;0,224,54,1;"
    strSQL = strSQL & "0,224,45,1;0,224,36,1;0,224,27,1;0,224,18,1;0,224,9,1;0,224,0,1;0,224,0,1;0,224,0,1;0,224,0,1;0,226,0,1;0,228,0,1;0,230,0,1;0,232,0,1;"
    strSQL = strSQL & "0,234,0,1;0,236,0,1;0,238,0,1;0,240,0,1;0,242,0,1;0,244,0,1;0,246,0,1;0,248,0,1;0,250,0,1;0,252,0,1;0,255,0,1;7,255,0,1;15,255,0,1;"
    strSQL = strSQL & "23,255,0,1;31,255,0,1;39,255,5,1;47,255,11,1;55,255,16,1;63,255,22,1;71,255,27,1;79,255,33,1;87,255,39,1;95,255,44,1;103,255,50,1;"
    strSQL = strSQL & "111,255,55,1;119,255,61,1;127,255,66,1;135,255,66,1;143,252,66,1;151,250,66,1;159,248,66,1;167,246,66,1;175,244,66,1;183,242,66,1;"
    strSQL = strSQL & "191,240,66,1;199,238,66,1;207,236,66,1;215,234,66,1;223,232,66,1;231,230,66,1;239,228,66,1;247,226,66,1;255,224,66,1;255,224,61,1;"
    strSQL = strSQL & "255,224,55,1;255,224,50,1;255,224,44,1;255,224,39,1;255,224,33,1;255,224,27,1;255,224,22,1;255,224,16,1;255,224,11,1;255,224,5,1;"
    strSQL = strSQL & "255,224,0,1;255,224,0,1;255,224,0,1;255,224,0,1;255,224,0,1;255,224,0,1;249,212,0,1;244,201,0,1;238,189,0,1;233,178,0,1;227,166,0,1;"
    strSQL = strSQL & "222,155,0,1;216,143,0,1;211,132,0,1;206,120,0,1;206,109,0,1;206,97,0,1;206,86,0,1;206,74,0,1;206,63,0,1;206,51,0,1;206,51,17,1;"
    strSQL = strSQL & "206,51,9,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;"
    strSQL = strSQL & "206,52,0,1;206,52,0,1;206,52,0,1;206,52,0,1;206,48,0,1;206,45,0,1;206,41,0,1;206,38,0,1;206,34,0,1;206,31,0,1;211,27,0,1;216,24,0,1;"
    strSQL = strSQL & "222,20,0,1;227,17,0,1;233,13,0,1;238,10,0,1;244,6,0,1;249,3,8,1;255,0,17,1;255,0,25,1;255,0,34,1;255,0,42,1;255,0,51,1;255,0,59,1;"
    strSQL = strSQL & "255,0,68,1;255,0,76,1;255,0,85,1;255,0,93,1;255,0,102,1;255,0,110,1;255,0,119,1;255,0,128,1;255,0,132,1;255,0,137,1;255,0,142,1;"
    strSQL = strSQL & "255,0,147,1;255,0,152,1;255,0,156,1;255,0,161,1;255,0,166,1;255,0,171,1;255,0,176,1;255,0,180,1;255,0,185,1;255,0,190,1;255,0,195,1;"
    strSQL = strSQL & "255,0,200,1;255,0,205,1;255,0,205,1;255,0,205,1;255,0,205,1;255,14,205,1;255,28,205,1;255,42,205,1;248,56,205,1;242,70,205,1;"
    strSQL = strSQL & "236,84,205,1;230,99,205,1;224,113,205,1;218,127,205,1;212,141,205,1;206,155,205,1;206,169,205,1;206,183,205,1;206,198,205,1;"
    strSQL = strSQL & "206,201,205,1;206,204,205,1;206,207,205,1;206,210,205,1;206,213,205,1;206,216,205,1;206,219,208,1;206,222,211,1;206,225,214,1;"
    strSQL = strSQL & "210,228,217,1;214,231,220,1;218,234,223,1;222,237,226,1;226,240,229,1;230,243,232,1;234,246,235,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (31,'PIXEF','0,0,0,1;0,0,0,1;0,0,0,1;0,1,1,1;0,4,4,1;0,5,5,1;0,8,8,1;"
    strSQL = strSQL & "0,11,11,1;0,16,16,1;0,17,17,1;0,20,20,1;0,21,21,1;0,24,24,1;0,24,24,1;0,24,24,1;0,27,27,1;0,32,32,1;0,33,33,1;0,36,36,1;0,39,39,1;"
    strSQL = strSQL & "0,44,44,1;0,45,45,1;0,48,48,1;0,49,49,1;0,52,52,1;0,53,53,1;0,56,56,1;0,57,57,1;0,60,60,1;0,61,61,1;0,64,64,1;0,65,65,1;0,68,68,1;"
    strSQL = strSQL & "0,71,71,1;0,76,76,1;0,77,77,1;0,80,80,1;0,81,81,1;0,84,84,1;0,85,85,1;0,88,88,1;0,89,89,1;0,91,91,1;0,93,93,1;0,96,96,1;0,99,99,1;"
    strSQL = strSQL & "0,104,104,1;0,105,105,1;0,108,108,1;0,109,109,1;0,112,112,1;0,113,113,1;0,116,116,1;0,117,117,1;0,120,120,1;0,121,121,1;0,124,124,1;"
    strSQL = strSQL & "0,127,127,1;0,132,132,1;0,133,133,1;0,136,136,1;0,137,137,1;0,140,140,1;0,141,141,1;0,144,144,1;0,145,145,1;0,148,148,1;0,149,149,1;"
    strSQL = strSQL & "0,152,152,1;0,153,153,1;0,156,156,1;0,159,159,1;0,164,164,1;0,165,165,1;0,168,168,1;0,169,169,1;0,172,172,1;0,173,173,1;0,176,176,1;"
    strSQL = strSQL & "0,177,177,1;0,180,180,1;0,181,181,1;0,184,184,1;0,187,187,1;0,192,192,1;0,193,193,1;0,196,196,1;0,197,197,1;0,200,200,1;0,201,201,1;"
    strSQL = strSQL & "0,204,204,1;0,205,205,1;0,208,208,1;0,209,209,1;0,212,212,1;0,213,213,1;0,216,216,1;0,219,219,1;0,224,224,1;0,225,225,1;0,228,228,1;"
    strSQL = strSQL & "0,229,229,1;0,232,232,1;0,233,233,1;0,236,236,1;0,237,237,1;0,240,240,1;0,241,241,1;0,244,244,1;0,247,247,1;0,252,252,1;0,243,243,1;"
    strSQL = strSQL & "0,224,224,1;0,215,215,1;0,196,196,1;0,187,187,1;0,168,168,1;0,159,159,1;0,140,140,1;0,131,131,1;0,112,112,1;0,103,103,1;0,84,84,1;"
    strSQL = strSQL & "0,75,75,1;0,56,56,1;0,47,47,1;0,28,28,1;0,19,19,1;0,0,0,1;5,5,0,1;16,16,0,1;20,20,0,1;28,28,0,1;33,33,0,1;44,44,0,1;49,49,0,1;"
    strSQL = strSQL & "60,60,0,1;65,65,0,1;76,76,0,1;81,81,0,1;92,92,0,1;97,97,0,1;108,108,0,1;113,113,0,1;124,124,0,1;129,129,0,1;140,140,0,1;146,146,0,1;"
    strSQL = strSQL & "153,153,0,1;159,159,0,1;172,172,0,1;177,177,0,1;188,188,0,1;193,193,0,1;204,204,0,1;209,209,0,1;220,220,0,1;225,225,0,1;236,236,0,1;"
    strSQL = strSQL & "241,241,0,1;252,252,0,1;252,228,0,1;252,203,0,1;252,179,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;"
    strSQL = strSQL & "252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,136,24,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,185,132,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (32,'PIXRED','0,0,0,1;0,0,0,1;0,0,0,1;1,0,1,1;4,0,4,1;5,0,5,1;8,0,8,1;"
    strSQL = strSQL & "11,0,11,1;16,0,16,1;17,0,17,1;20,0,20,1;21,0,21,1;24,0,24,1;24,0,24,1;24,0,24,1;27,0,27,1;32,0,32,1;33,0,33,1;36,0,36,1;39,0,39,1;"
    strSQL = strSQL & "44,0,44,1;45,0,45,1;48,0,48,1;49,0,49,1;52,0,52,1;53,0,53,1;56,0,56,1;57,0,57,1;60,0,60,1;61,0,61,1;64,0,64,1;65,0,65,1;68,0,68,1;"
    strSQL = strSQL & "71,0,71,1;76,0,76,1;77,0,77,1;80,0,80,1;81,0,81,1;84,0,84,1;85,0,85,1;88,0,88,1;89,0,89,1;91,0,91,1;93,0,93,1;96,0,96,1;99,0,99,1;"
    strSQL = strSQL & "104,0,104,1;105,0,105,1;108,0,108,1;109,0,109,1;112,0,112,1;113,0,113,1;116,0,116,1;117,0,117,1;120,0,120,1;121,0,121,1;124,0,124,1;"
    strSQL = strSQL & "127,0,127,1;132,0,132,1;133,0,133,1;136,0,136,1;137,0,137,1;140,0,140,1;141,0,141,1;144,0,144,1;145,0,145,1;148,0,148,1;149,0,149,1;"
    strSQL = strSQL & "152,0,152,1;153,0,153,1;156,0,156,1;159,0,159,1;164,0,164,1;165,0,165,1;168,0,168,1;169,0,169,1;172,0,172,1;173,0,173,1;176,0,176,1;"
    strSQL = strSQL & "177,0,177,1;180,0,180,1;181,0,181,1;184,0,184,1;187,0,187,1;192,0,192,1;193,0,193,1;196,0,196,1;197,0,197,1;200,0,200,1;201,0,201,1;"
    strSQL = strSQL & "204,0,204,1;205,0,205,1;208,0,208,1;209,0,209,1;212,0,212,1;213,0,213,1;216,0,216,1;219,0,219,1;224,0,224,1;225,0,225,1;228,0,228,1;"
    strSQL = strSQL & "229,0,229,1;232,0,232,1;233,0,233,1;236,0,236,1;237,0,237,1;240,0,240,1;241,0,241,1;244,0,244,1;247,0,247,1;252,0,252,1;243,0,243,1;"
    strSQL = strSQL & "224,0,224,1;215,0,215,1;196,0,196,1;187,0,187,1;168,0,168,1;159,0,159,1;140,0,140,1;131,0,131,1;112,0,112,1;103,0,103,1;84,0,84,1;"
    strSQL = strSQL & "75,0,75,1;56,0,56,1;47,0,47,1;28,0,28,1;19,0,19,1;0,0,0,1;5,5,0,1;16,16,0,1;20,20,0,1;28,28,0,1;33,33,0,1;44,44,0,1;49,49,0,1;"
    strSQL = strSQL & "60,60,0,1;65,65,0,1;76,76,0,1;81,81,0,1;92,92,0,1;97,97,0,1;108,108,0,1;113,113,0,1;124,124,0,1;129,129,0,1;140,140,0,1;146,146,0,1;"
    strSQL = strSQL & "153,153,0,1;159,159,0,1;172,172,0,1;177,177,0,1;188,188,0,1;193,193,0,1;204,204,0,1;209,209,0,1;220,220,0,1;225,225,0,1;236,236,0,1;"
    strSQL = strSQL & "241,241,0,1;252,252,0,1;252,228,0,1;252,203,0,1;252,179,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;"
    strSQL = strSQL & "252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,128,0,1;252,136,24,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;252,152,72,1;"
    strSQL = strSQL & "252,185,132,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;"
    strSQL = strSQL & "252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;252,252,252,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (33,'prism','0,0,0,1;0,0,19,1;0,1,38,1;0,2,56,1;0,4,74,1;0,7,90,1;"
    strSQL = strSQL & "0,10,106,1;0,13,121,1;0,16,135,1;0,21,148,1;0,25,161,1;0,29,173,1;0,34,184,1;0,39,194,1;0,44,203,1;0,50,212,1;0,55,220,1;0,60,227,1;"
    strSQL = strSQL & "0,66,233,1;0,71,239,1;0,76,244,1;0,81,248,1;0,86,251,1;0,91,253,1;0,96,255,1;0,100,255,1;0,104,255,1;0,108,255,1;0,112,255,1;"
    strSQL = strSQL & "0,116,255,1;0,120,255,1;0,124,255,1;0,128,255,1;0,132,255,1;0,136,255,1;0,140,255,1;0,144,255,1;0,148,255,1;0,152,255,1;0,156,255,1;"
    strSQL = strSQL & "0,160,255,1;0,164,255,1;0,168,255,1;0,172,255,1;0,176,255,1;0,180,255,1;0,184,255,1;0,188,255,1;0,192,255,1;0,196,255,1;0,200,255,1;"
    strSQL = strSQL & "0,204,255,1;0,208,255,1;0,212,255,1;0,216,255,1;0,220,255,1;0,224,255,1;0,228,255,1;0,232,255,1;0,236,255,1;0,240,255,1;0,244,255,1;"
    strSQL = strSQL & "0,248,255,1;0,252,255,1;0,255,254,1;0,255,250,1;0,255,246,1;0,255,242,1;0,255,238,1;0,255,234,1;0,255,230,1;0,255,226,1;0,255,222,1;"
    strSQL = strSQL & "0,255,218,1;0,255,214,1;0,255,210,1;0,255,206,1;0,255,202,1;0,255,198,1;0,255,194,1;0,255,190,1;0,255,186,1;0,255,182,1;0,255,178,1;"
    strSQL = strSQL & "0,255,174,1;0,255,170,1;0,255,166,1;0,255,162,1;0,255,158,1;0,255,154,1;0,255,150,1;0,255,146,1;0,255,142,1;0,255,138,1;0,255,134,1;"
    strSQL = strSQL & "0,255,130,1;0,255,126,1;0,255,122,1;0,255,118,1;0,255,114,1;0,255,110,1;0,255,106,1;0,255,102,1;0,255,98,1;0,255,94,1;0,255,90,1;"
    strSQL = strSQL & "0,255,86,1;0,255,82,1;0,255,78,1;0,255,74,1;0,255,70,1;0,255,66,1;0,255,62,1;0,255,58,1;0,255,54,1;0,255,50,1;0,255,46,1;0,255,42,1;"
    strSQL = strSQL & "0,255,38,1;0,255,34,1;0,255,30,1;0,255,26,1;0,255,22,1;0,255,18,1;0,255,14,1;0,255,10,1;0,255,6,1;0,255,2,1;2,255,0,1;6,255,0,1;"
    strSQL = strSQL & "10,255,0,1;14,255,0,1;18,255,0,1;22,255,0,1;26,255,0,1;30,255,0,1;34,255,0,1;38,255,0,1;42,255,0,1;46,255,0,1;50,255,0,1;54,255,0,1;"
    strSQL = strSQL & "58,255,0,1;62,255,0,1;66,255,0,1;70,255,0,1;74,255,0,1;78,255,0,1;82,255,0,1;86,255,0,1;90,255,0,1;94,255,0,1;98,255,0,1;102,255,0,1;"
    strSQL = strSQL & "106,255,0,1;110,255,0,1;114,255,0,1;118,255,0,1;122,255,0,1;126,255,0,1;130,255,0,1;134,255,0,1;138,255,0,1;142,255,0,1;146,255,0,1;"
    strSQL = strSQL & "150,255,0,1;154,255,0,1;158,255,0,1;162,255,0,1;166,255,0,1;170,255,0,1;174,255,0,1;178,255,0,1;182,255,0,1;186,255,0,1;190,255,0,1;"
    strSQL = strSQL & "194,255,0,1;198,255,0,1;202,255,0,1;206,255,0,1;210,255,0,1;214,255,0,1;218,255,0,1;222,255,0,1;226,255,0,1;230,255,0,1;234,255,0,1;"
    strSQL = strSQL & "238,255,0,1;242,255,0,1;246,255,0,1;250,255,0,1;254,255,0,1;255,252,0,1;255,248,0,1;255,244,0,1;255,240,0,1;255,236,0,1;255,232,0,1;"
    strSQL = strSQL & "255,228,0,1;255,224,0,1;255,220,0,1;255,216,0,1;255,212,0,1;255,208,0,1;255,204,0,1;255,200,0,1;255,196,0,1;255,192,0,1;255,188,0,1;"
    strSQL = strSQL & "255,184,0,1;255,180,0,1;255,176,0,1;255,172,0,1;255,168,0,1;255,164,0,1;255,160,0,1;255,156,0,1;255,152,0,1;255,148,0,1;255,144,0,1;"
    strSQL = strSQL & "255,140,0,1;255,136,0,1;255,132,0,1;255,128,0,1;255,124,0,1;255,120,0,1;255,116,0,1;255,112,0,1;255,108,0,1;255,104,0,1;255,100,0,1;"
    strSQL = strSQL & "255,96,0,1;255,93,2,1;255,91,4,1;255,89,7,1;255,88,11,1;255,87,16,1;255,88,22,1;255,89,28,1;255,90,35,1;255,93,43,1;255,96,52,1;"
    strSQL = strSQL & "255,101,61,1;255,106,71,1;255,112,82,1;255,119,94,1;255,128,107,1;255,137,120,1;255,148,134,1;255,159,149,1;255,172,165,1;"
    strSQL = strSQL & "255,186,181,1;255,201,199,1;255,218,217,1;255,236,236,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (34,'GEM-16','0,0,0,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;"
    strSQL = strSQL & "0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,8,16,1;0,16,32,1;0,16,32,1;0,16,32,1;"
    strSQL = strSQL & "0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;0,16,32,1;"
    strSQL = strSQL & "0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;0,24,48,1;"
    strSQL = strSQL & "0,24,48,1;0,24,48,1;0,24,48,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;"
    strSQL = strSQL & "0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;0,32,64,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;"
    strSQL = strSQL & "31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;31,24,80,1;63,16,96,1;63,16,96,1;"
    strSQL = strSQL & "63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;63,16,96,1;"
    strSQL = strSQL & "63,16,96,1;63,16,96,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;"
    strSQL = strSQL & "95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;95,8,112,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;"
    strSQL = strSQL & "127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;"
    strSQL = strSQL & "159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;"
    strSQL = strSQL & "159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;159,25,95,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;"
    strSQL = strSQL & "191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;191,50,63,1;223,74,31,1;"
    strSQL = strSQL & "223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;"
    strSQL = strSQL & "223,74,31,1;223,74,31,1;223,74,31,1;223,74,31,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;"
    strSQL = strSQL & "255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,98,3,1;255,138,67,1;255,138,67,1;"
    strSQL = strSQL & "255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;"
    strSQL = strSQL & "255,138,67,1;255,138,67,1;255,138,67,1;255,138,67,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;"
    strSQL = strSQL & "255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;255,178,131,1;"
    strSQL = strSQL & "255,178,131,1;255,178,131,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;"
    strSQL = strSQL & "255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;255,218,195,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (35,'Rainbow','0,0,0,1;2,2,2,1;4,4,4,1;6,6,6,1;8,8,8,1;10,10,10,1;"
    strSQL = strSQL & "12,12,12,1;14,14,14,1;16,16,16,1;18,18,18,1;20,20,20,1;22,22,22,1;24,24,24,1;26,26,26,1;28,28,28,1;30,30,30,1;32,32,32,1;34,34,34,1;"
    strSQL = strSQL & "36,36,36,1;38,38,38,1;40,40,40,1;42,42,42,1;44,44,44,1;46,46,46,1;48,48,48,1;50,50,50,1;52,52,52,1;54,54,54,1;56,56,56,1;58,58,58,1;"
    strSQL = strSQL & "60,60,60,1;62,62,62,1;64,64,64,1;66,66,66,1;68,68,68,1;70,70,70,1;72,72,72,1;74,74,74,1;76,76,76,1;78,78,78,1;80,80,80,1;82,82,82,1;"
    strSQL = strSQL & "84,84,84,1;86,86,86,1;88,88,88,1;90,90,90,1;92,92,92,1;94,94,94,1;96,96,96,1;98,98,98,1;100,100,100,1;79,47,79,1;82,49,82,1;"
    strSQL = strSQL & "85,51,85,1;88,53,88,1;91,55,91,1;94,57,94,1;97,59,97,1;100,61,100,1;103,63,103,1;106,65,106,1;109,67,109,1;112,69,112,1;"
    strSQL = strSQL & "115,71,115,1;118,73,118,1;121,75,121,1;124,77,124,1;127,79,127,1;130,81,130,1;133,83,133,1;136,85,136,1;139,87,139,1;142,89,142,1;"
    strSQL = strSQL & "147,91,147,1;151,93,151,1;155,94,155,1;159,95,159,1;153,91,163,1;147,87,167,1;141,83,171,1;135,79,175,1;129,75,179,1;123,71,183,1;"
    strSQL = strSQL & "117,67,187,1;111,63,191,1;105,59,195,1;99,55,199,1;93,51,203,1;87,47,207,1;81,43,211,1;75,39,215,1;69,35,219,1;63,31,223,1;"
    strSQL = strSQL & "57,27,227,1;51,23,231,1;45,19,235,1;39,16,239,1;33,13,243,1;27,10,247,1;15,6,250,1;9,3,253,1;0,1,254,1;0,0,255,1;0,10,240,1;"
    strSQL = strSQL & "0,20,230,1;0,30,220,1;0,40,210,1;0,50,200,1;0,60,190,1;0,70,180,1;0,80,170,1;0,90,160,1;0,100,150,1;0,110,140,1;0,120,130,1;"
    strSQL = strSQL & "0,130,120,1;0,140,110,1;0,150,100,1;0,160,90,1;0,170,80,1;0,180,70,1;0,190,60,1;0,200,50,1;0,210,40,1;0,220,30,1;0,230,20,1;"
    strSQL = strSQL & "0,240,10,1;0,255,0,1;10,250,0,1;20,240,0,1;30,230,0,1;40,220,0,1;50,210,0,1;60,200,0,1;70,190,0,1;80,180,0,1;90,170,0,1;"
    strSQL = strSQL & "100,160,0,1;110,150,0,1;120,140,0,1;130,130,0,1;140,120,0,1;150,110,0,1;160,100,0,1;170,90,0,1;180,80,0,1;190,70,0,1;"
    strSQL = strSQL & "200,60,0,1;210,50,0,1;220,40,0,1;230,30,0,1;240,20,0,1;250,10,0,1;255,0,0,1;252,2,2,1;250,4,4,1;248,6,6,1;246,8,8,1;"
    strSQL = strSQL & "244,10,10,1;242,12,12,1;240,14,14,1;238,16,16,1;236,18,18,1;234,20,20,1;232,22,22,1;230,24,24,1;228,26,26,1;226,28,28,1;"
    strSQL = strSQL & "224,30,30,1;222,32,32,1;220,34,34,1;218,36,36,1;216,38,38,1;214,40,40,1;212,42,42,1;210,44,44,1;208,46,46,1;206,48,48,1;"
    strSQL = strSQL & "204,50,50,1;206,58,49,1;208,66,48,1;210,74,46,1;212,81,44,1;214,89,42,1;216,97,40,1;218,105,38,1;220,113,36,1;222,121,34,1;"
    strSQL = strSQL & "224,129,32,1;226,137,30,1;228,145,28,1;230,153,26,1;232,161,24,1;234,169,22,1;236,177,20,1;238,185,18,1;240,193,16,1;"
    strSQL = strSQL & "242,201,14,1;244,209,12,1;246,217,10,1;248,225,8,1;250,233,6,1;252,241,4,1;254,249,2,1;255,255,0,1;255,255,6,1;255,255,12,1;"
    strSQL = strSQL & "255,255,18,1;255,255,24,1;255,255,30,1;255,255,36,1;255,255,42,1;255,255,48,1;255,255,54,1;255,255,60,1;255,255,66,1;"
    strSQL = strSQL & "255,255,72,1;255,255,78,1;255,255,84,1;255,255,90,1;255,255,96,1;255,255,102,1;255,255,108,1;255,255,114,1;255,255,120,1;"
    strSQL = strSQL & "255,255,125,1;255,255,130,1;255,255,135,1;255,255,140,1;255,255,145,1;252,252,145,1;252,252,149,1;252,252,153,1;252,252,157,1;"
    strSQL = strSQL & "252,252,161,1;252,252,165,1;252,252,169,1;252,252,173,1;252,252,177,1;252,252,181,1;252,252,185,1;252,252,189,1;252,252,193,1;"
    strSQL = strSQL & "252,252,197,1;252,252,201,1;252,252,205,1;252,252,209,1;252,252,213,1;252,252,218,1;252,252,223,1;252,252,228,1;252,252,233,1;"
    strSQL = strSQL & "252,252,238,1;252,252,243,1;252,252,248,1;252,252,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (36,'GEM-256','0,0,0,1;0,1,1,1;0,1,2,1;0,2,3,1;0,2,4,1;0,3,5,1;0,3,6,1;"
    strSQL = strSQL & "0,4,7,1;0,4,8,1;0,5,9,1;0,5,10,1;0,6,11,1;0,6,12,1;0,7,13,1;0,7,14,1;0,8,15,1;0,8,16,1;0,9,17,1;0,9,18,1;0,10,19,1;0,10,20,1;"
    strSQL = strSQL & "0,11,21,1;0,11,22,1;0,12,23,1;0,12,24,1;0,13,25,1;0,13,26,1;0,14,27,1;0,14,28,1;0,15,29,1;0,15,30,1;0,16,31,1;0,16,32,1;0,17,33,1;"
    strSQL = strSQL & "0,17,34,1;0,18,35,1;0,18,36,1;0,19,37,1;0,19,38,1;0,20,39,1;0,20,40,1;0,21,41,1;0,21,42,1;0,22,43,1;0,22,44,1;0,23,45,1;0,23,46,1;"
    strSQL = strSQL & "0,24,47,1;0,24,48,1;0,25,49,1;0,25,50,1;0,26,51,1;0,26,52,1;0,27,53,1;0,27,54,1;0,28,55,1;0,28,56,1;0,29,57,1;0,29,58,1;0,30,59,1;"
    strSQL = strSQL & "0,30,60,1;0,31,61,1;0,31,62,1;0,32,63,1;0,32,64,1;1,31,65,1;3,31,66,1;5,30,67,1;7,30,68,1;9,29,69,1;11,29,70,1;13,28,71,1;15,28,72,1;"
    strSQL = strSQL & "17,27,73,1;19,27,74,1;21,26,75,1;23,26,76,1;25,25,77,1;27,25,78,1;29,24,79,1;31,24,80,1;33,23,81,1;35,23,82,1;37,22,83,1;39,22,84,1;"
    strSQL = strSQL & "41,21,85,1;43,21,86,1;45,20,87,1;47,20,88,1;49,19,89,1;51,19,90,1;53,18,91,1;55,18,92,1;57,17,93,1;59,17,94,1;61,16,95,1;63,16,96,1;"
    strSQL = strSQL & "65,15,97,1;67,15,98,1;69,14,99,1;71,14,100,1;73,13,101,1;75,13,102,1;77,12,103,1;79,12,104,1;81,11,105,1;83,11,106,1;85,10,107,1;"
    strSQL = strSQL & "87,10,108,1;89,9,109,1;91,9,110,1;93,8,111,1;95,8,112,1;97,7,113,1;99,7,114,1;101,6,115,1;103,6,116,1;105,5,117,1;107,5,118,1;"
    strSQL = strSQL & "109,4,119,1;111,4,120,1;113,3,121,1;115,3,122,1;117,2,123,1;119,2,124,1;121,1,125,1;123,1,126,1;125,0,127,1;127,2,127,1;129,3,125,1;"
    strSQL = strSQL & "131,5,123,1;133,6,121,1;135,8,119,1;137,9,117,1;139,11,115,1;141,12,113,1;143,14,111,1;145,15,109,1;147,17,107,1;149,18,105,1;"
    strSQL = strSQL & "151,20,103,1;153,21,101,1;155,23,99,1;157,24,97,1;159,26,95,1;161,27,93,1;163,29,91,1;165,30,89,1;167,32,87,1;169,33,85,1;"
    strSQL = strSQL & "171,35,83,1;173,36,81,1;175,38,79,1;177,39,77,1;179,41,75,1;181,42,73,1;183,44,71,1;185,45,69,1;187,47,67,1;189,48,65,1;"
    strSQL = strSQL & "191,50,63,1;193,51,61,1;195,53,59,1;197,54,57,1;199,56,55,1;201,57,53,1;203,59,51,1;205,60,49,1;207,62,47,1;209,63,45,1;211,65,43,1;"
    strSQL = strSQL & "213,66,41,1;215,68,39,1;217,69,37,1;219,71,35,1;221,72,33,1;223,74,31,1;225,75,29,1;227,77,27,1;229,78,25,1;231,80,23,1;233,81,21,1;"
    strSQL = strSQL & "235,83,19,1;237,84,17,1;239,86,15,1;241,87,13,1;243,89,11,1;245,90,9,1;247,92,7,1;249,93,5,1;251,95,3,1;253,95,0,1;255,98,3,1;"
    strSQL = strSQL & "255,100,7,1;255,103,11,1;255,105,15,1;255,108,19,1;255,110,23,1;255,113,27,1;255,115,31,1;255,118,35,1;255,120,39,1;255,123,43,1;"
    strSQL = strSQL & "255,125,47,1;255,128,51,1;255,130,55,1;255,133,59,1;255,135,63,1;255,138,67,1;255,140,71,1;255,143,75,1;255,145,79,1;255,148,83,1;"
    strSQL = strSQL & "255,150,87,1;255,153,91,1;255,155,95,1;255,158,99,1;255,160,103,1;255,163,107,1;255,165,111,1;255,168,115,1;255,170,119,1;"
    strSQL = strSQL & "255,173,123,1;255,175,127,1;255,178,131,1;255,180,135,1;255,183,139,1;255,185,143,1;255,188,147,1;255,190,151,1;255,193,155,1;"
    strSQL = strSQL & "255,195,159,1;255,198,163,1;255,200,167,1;255,203,171,1;255,205,175,1;255,208,179,1;255,210,183,1;255,213,187,1;255,215,191,1;"
    strSQL = strSQL & "255,218,195,1;255,220,199,1;255,223,203,1;255,225,207,1;255,228,211,1;255,230,215,1;255,233,219,1;255,235,223,1;255,238,227,1;"
    strSQL = strSQL & "255,240,231,1;255,243,235,1;255,245,239,1;255,248,243,1;255,250,247,1;255,253,251,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (37,'S_Pet','0,0,0,1;3,0,3,1;6,0,7,1;8,0,12,1;9,0,16,1;12,0,20,1;"
    strSQL = strSQL & "15,0,24,1;17,0,29,1;18,0,33,1;21,0,37,1;24,0,41,1;26,0,46,1;27,0,50,1;30,0,54,1;33,0,58,1;34,0,63,1;37,0,67,1;40,0,71,1;42,0,75,1;"
    strSQL = strSQL & "43,0,80,1;46,0,84,1;49,0,88,1;51,0,92,1;52,0,97,1;55,0,101,1;58,0,105,1;60,0,109,1;61,0,114,1;64,0,118,1;67,0,122,1;68,0,127,1;"
    strSQL = strSQL & "68,0,127,1;67,0,131,1;64,0,135,1;61,0,139,1;60,0,143,1;58,0,147,1;55,0,151,1;52,0,155,1;51,0,159,1;49,0,163,1;46,0,167,1;43,0,171,1;"
    strSQL = strSQL & "42,0,175,1;40,0,179,1;37,0,183,1;34,0,187,1;33,0,191,1;30,0,195,1;27,0,199,1;26,0,203,1;24,0,207,1;21,0,211,1;18,0,215,1;17,0,219,1;"
    strSQL = strSQL & "15,0,223,1;12,0,227,1;9,0,231,1;8,0,235,1;6,0,239,1;3,0,243,1;1,0,247,1;0,0,255,1;0,0,255,1;0,4,247,1;0,9,243,1;0,13,239,1;0,18,235,1;"
    strSQL = strSQL & "0,22,231,1;0,27,227,1;0,31,223,1;0,36,219,1;0,40,215,1;0,45,211,1;0,49,207,1;0,54,203,1;0,58,199,1;0,63,195,1;0,67,191,1;0,72,187,1;"
    strSQL = strSQL & "0,76,183,1;0,81,179,1;0,85,175,1;0,90,171,1;0,94,167,1;0,99,163,1;0,103,159,1;0,108,155,1;0,112,151,1;0,117,147,1;0,121,143,1;"
    strSQL = strSQL & "0,126,139,1;0,130,135,1;0,135,131,1;0,139,127,1;0,144,122,1;0,148,118,1;0,153,114,1;0,157,109,1;0,162,105,1;0,166,101,1;0,171,97,1;"
    strSQL = strSQL & "0,175,92,1;0,180,88,1;0,184,84,1;0,189,80,1;0,193,75,1;0,204,71,1;0,204,67,1;9,206,63,1;18,208,58,1;27,210,54,1;36,212,50,1;45,214,46,1;"
    strSQL = strSQL & "54,216,41,1;63,218,37,1;72,220,33,1;81,222,29,1;90,224,24,1;99,226,20,1;108,228,16,1;117,230,12,1;126,232,10,1;135,234,7,1;144,236,4,1;"
    strSQL = strSQL & "153,238,1,1;162,240,0,1;171,242,0,1;180,243,0,1;189,245,0,1;198,246,0,1;207,248,0,1;216,249,0,1;225,251,0,1;234,252,0,1;243,254,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,254,0,1;255,252,0,1;255,250,0,1;255,248,0,1;255,246,0,1;255,244,0,1;255,242,0,1;255,240,0,1;255,238,0,1;"
    strSQL = strSQL & "255,236,0,1;255,234,0,1;255,232,0,1;255,230,0,1;255,228,0,1;255,226,0,1;255,224,0,1;255,222,0,1;255,220,0,1;255,218,0,1;255,216,0,1;"
    strSQL = strSQL & "255,214,0,1;255,212,0,1;255,210,0,1;255,208,0,1;255,206,0,1;255,204,0,1;255,202,0,1;255,200,0,1;255,198,0,1;255,196,0,1;255,194,0,1;"
    strSQL = strSQL & "255,192,0,1;255,190,0,1;255,188,0,1;255,186,0,1;255,184,0,1;255,182,0,1;255,180,0,1;255,178,0,1;255,176,0,1;255,174,0,1;255,172,0,1;"
    strSQL = strSQL & "255,170,0,1;255,168,0,1;255,166,0,1;255,164,0,1;255,162,0,1;255,160,0,1;255,158,0,1;255,156,0,1;255,154,0,1;255,152,0,1;255,150,0,1;"
    strSQL = strSQL & "255,148,0,1;255,146,0,1;255,144,0,1;255,142,0,1;255,140,0,1;255,138,0,1;255,136,0,1;255,134,0,1;255,132,0,1;255,130,0,1;255,128,0,1;"
    strSQL = strSQL & "255,126,0,1;255,124,0,1;255,122,0,1;255,120,0,1;255,118,0,1;255,116,0,1;255,114,0,1;255,112,0,1;255,110,0,1;255,108,0,1;255,106,0,1;"
    strSQL = strSQL & "255,104,0,1;255,102,0,1;255,100,0,1;255,98,0,1;255,96,0,1;255,94,0,1;255,92,0,1;255,90,0,1;255,88,0,1;255,86,0,1;255,84,0,1;255,82,0,1;"
    strSQL = strSQL & "255,80,0,1;255,78,0,1;255,76,0,1;255,74,0,1;255,72,0,1;255,69,0,1;255,66,0,1;255,64,0,1;255,61,0,1;255,58,0,1;255,55,0,1;255,53,0,1;"
    strSQL = strSQL & "255,50,0,1;255,47,0,1;255,44,0,1;255,42,0,1;255,39,0,1;255,36,0,1;255,33,0,1;255,31,0,1;255,28,0,1;255,25,0,1;255,22,0,1;255,20,0,1;"
    strSQL = strSQL & "255,17,0,1;255,14,0,1;255,11,0,1;255,9,0,1;255,6,0,1;255,3,0,1;255,0,0,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (38,'smart','0,0,0,1;1,1,1,1;1,1,1,1;2,2,2,1;3,3,3,1;4,4,4,1;4,4,4,1;"
    strSQL = strSQL & "5,5,5,1;6,6,6,1;7,7,7,1;7,7,7,1;8,8,8,1;9,9,9,1;10,10,10,1;10,10,10,1;11,11,11,1;12,12,12,1;13,13,13,1;13,13,13,1;14,14,14,1;15,15,15,1;"
    strSQL = strSQL & "15,15,15,1;16,16,16,1;17,17,17,1;18,18,18,1;18,18,18,1;19,19,19,1;20,20,20,1;21,21,21,1;21,21,21,1;22,22,22,1;23,23,23,1;24,24,24,1;"
    strSQL = strSQL & "24,24,24,1;25,25,25,1;26,26,26,1;26,26,26,1;27,27,27,1;28,28,28,1;29,29,29,1;29,29,29,1;30,30,30,1;31,31,31,1;32,32,32,1;32,32,32,1;"
    strSQL = strSQL & "33,33,33,1;34,34,34,1;35,35,35,1;35,35,35,1;36,36,36,1;37,37,37,1;38,38,38,1;38,38,38,1;39,39,39,1;40,40,40,1;40,40,40,1;41,41,41,1;"
    strSQL = strSQL & "42,42,42,1;43,43,43,1;43,43,43,1;44,44,44,1;45,45,45,1;46,46,46,1;46,46,46,1;47,47,47,1;48,48,48,1;49,49,49,1;49,49,49,1;50,50,50,1;"
    strSQL = strSQL & "51,51,51,1;51,51,51,1;52,52,52,1;53,53,53,1;54,54,54,1;54,54,54,1;55,55,55,1;56,56,56,1;57,57,57,1;57,57,57,1;58,58,58,1;59,59,59,1;"
    strSQL = strSQL & "60,60,60,1;60,60,60,1;61,61,61,1;62,62,62,1;63,63,63,1;63,63,63,1;64,64,64,1;65,65,65,1;65,65,65,1;66,66,66,1;67,67,67,1;68,68,68,1;"
    strSQL = strSQL & "68,68,68,1;69,69,69,1;70,70,70,1;71,71,71,1;71,71,71,1;72,72,72,1;73,73,73,1;74,74,74,1;74,74,74,1;75,75,75,1;76,71,71,1;77,68,66,1;"
    strSQL = strSQL & "77,64,61,1;78,61,57,1;79,57,52,1;80,54,47,1;80,50,42,1;81,46,39,1;82,44,37,1;83,42,35,1;83,41,33,1;84,39,31,1;85,37,29,1;86,35,27,1;"
    strSQL = strSQL & "86,33,25,1;87,31,23,1;88,29,21,1;89,28,20,1;89,26,19,1;90,25,18,1;91,25,18,1;92,24,17,1;92,23,17,1;93,23,16,1;94,22,16,1;95,22,16,1;"
    strSQL = strSQL & "95,21,15,1;96,21,15,1;97,21,14,1;98,20,13,1;98,20,12,1;99,19,12,1;100,19,12,1;101,19,11,1;101,18,11,1;102,18,10,1;103,17,10,1;102,17,"
    strSQL = strSQL & "10,1;105,17,9,1;106,16,8,1;107,16,7,1;109,16,6,1;111,15,5,1;112,15,5,1;113,15,4,1;114,15,3,1;116,14,3,1;117,14,3,1;119,14,2,1;"
    strSQL = strSQL & "122,15,2,1;125,15,1,1;128,15,1,1;130,15,1,1;133,15,0,1;136,15,0,1;139,15,0,1;142,15,0,1;145,16,0,1;149,17,0,1;152,19,0,1;156,20,0,1;"
    strSQL = strSQL & "159,22,0,1;163,23,0,1;166,25,0,1;189,26,0,1;173,28,0,1;177,30,0,1;179,30,0,1;182,30,0,1;185,31,0,1;189,32,0,1;192,33,0,1;195,35,0,1;"
    strSQL = strSQL & "197,36,0,1;200,38,0,1;202,40,0,1;204,42,0,1;206,46,0,1;208,49,0,1;210,52,0,1;211,54,0,1;213,56,0,1;214,58,0,1;216,61,0,1;217,64,0,1;"
    strSQL = strSQL & "219,67,0,1;221,71,0,1;222,74,1,1;223,78,1,1;225,83,1,1;223,88,2,1;227,92,2,1;228,97,3,1;229,100,3,1;230,103,3,1;230,107,4,1;231,110,4,1;"
    strSQL = strSQL & "232,113,5,1;234,121,5,1;235,130,5,1;237,136,6,1;239,144,6,1;240,156,7,1;241,162,7,1;241,167,7,1;242,173,8,1;243,179,8,1;243,185,9,1;"
    strSQL = strSQL & "244,187,9,1;246,190,10,1;247,193,10,1;248,196,11,1;249,199,11,1;249,202,12,1;250,205,12,1;250,208,13,1;251,211,13,1;251,214,14,1;"
    strSQL = strSQL & "252,216,16,1;253,217,17,1;253,218,19,1;254,219,21,1;254,220,23,1;255,221,27,1;255,223,31,1;255,223,35,1;255,225,39,1;255,226,42,1;"
    strSQL = strSQL & "255,227,46,1;255,228,49,1;255,228,52,1;255,229,55,1;255,230,58,1;255,230,66,1;255,231,74,1;255,231,82,1;255,232,89,1;255,232,101,1;"
    strSQL = strSQL & "255,236,113,1;255,239,125,1;255,241,136,1;255,243,148,1;255,246,164,1;255,246,179,1;255,247,187,1;255,248,199,1;255,249,210,1;"
    strSQL = strSQL & "255,250,226,1;255,251,242,1;255,252,246,1;255,253,250,1;255,254,253,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (39,'Sopha','0,0,101,1;0,0,105,1;0,0,110,1;0,0,115,1;0,0,120,1;0,0,125,1;"
    strSQL = strSQL & "0,0,129,1;0,0,134,1;0,0,139,1;0,0,144,1;0,0,149,1;0,0,153,1;0,0,158,1;0,0,163,1;0,0,168,1;0,0,173,1;0,0,178,1;0,0,182,1;0,0,187,1;"
    strSQL = strSQL & "0,8,192,1;0,16,197,1;0,24,202,1;0,32,206,1;0,40,211,1;0,48,216,1;0,56,221,1;0,65,226,1;0,73,230,1;0,81,235,1;0,89,240,1;0,97,245,1;"
    strSQL = strSQL & "0,105,250,1;0,113,255,1;0,121,255,1;0,121,251,1;0,121,247,1;0,121,243,1;0,121,239,1;0,121,235,1;0,121,231,1;0,121,227,1;0,121,223,1;"
    strSQL = strSQL & "0,121,219,1;0,121,215,1;0,121,211,1;0,121,207,1;0,121,203,1;0,121,199,1;0,121,195,1;0,121,191,1;0,121,187,1;0,128,183,1;0,135,179,1;"
    strSQL = strSQL & "0,142,175,1;0,149,171,1;0,155,167,1;0,162,163,1;0,169,159,1;0,176,155,1;0,183,151,1;0,190,147,1;0,196,143,1;0,203,139,1;0,210,135,1;"
    strSQL = strSQL & "0,217,131,1;0,224,128,1;0,224,118,1;0,224,109,1;0,224,100,1;0,224,91,1;0,224,82,1;0,224,73,1;0,224,63,1;0,224,54,1;0,224,45,1;"
    strSQL = strSQL & "0,224,36,1;0,224,27,1;0,224,18,1;0,224,9,1;0,224,0,1;0,224,0,1;0,224,0,1;0,224,0,1;0,226,0,1;0,228,0,1;0,230,0,1;0,232,0,1;0,234,0,1;"
    strSQL = strSQL & "0,236,0,1;0,238,0,1;0,240,0,1;0,242,0,1;0,244,0,1;0,246,0,1;0,248,0,1;0,250,0,1;0,252,0,1;0,255,0,1;7,255,0,1;15,255,0,1;23,255,0,1;"
    strSQL = strSQL & "31,255,0,1;39,255,5,1;47,255,11,1;55,255,16,1;63,255,22,1;71,255,27,1;79,255,33,1;87,255,39,1;95,255,44,1;103,255,50,1;111,255,55,1;"
    strSQL = strSQL & "119,255,61,1;127,255,66,1;135,255,66,1;143,252,66,1;151,250,66,1;159,248,66,1;167,246,66,1;175,244,66,1;183,242,66,1;191,240,66,1;"
    strSQL = strSQL & "199,238,66,1;207,236,66,1;215,234,66,1;223,232,66,1;231,230,66,1;239,228,66,1;247,226,66,1;255,224,66,1;255,224,61,1;255,224,55,1;"
    strSQL = strSQL & "255,224,50,1;255,224,44,1;255,224,39,1;255,224,33,1;255,224,27,1;255,224,22,1;255,224,16,1;255,224,11,1;255,224,5,1;255,224,0,1;"
    strSQL = strSQL & "255,224,0,1;255,224,0,1;255,224,0,1;255,224,0,1;255,224,0,1;249,212,0,1;244,201,0,1;238,189,0,1;233,178,0,1;227,166,0,1;222,155,0,1;"
    strSQL = strSQL & "216,143,0,1;211,132,0,1;206,120,0,1;206,109,0,1;206,97,0,1;206,86,0,1;206,74,0,1;206,63,0,1;206,51,0,1;206,51,17,1;206,51,9,1;"
    strSQL = strSQL & "206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,51,0,1;206,52,0,1;"
    strSQL = strSQL & "206,52,0,1;206,52,0,1;206,52,0,1;206,48,0,1;206,45,0,1;206,41,0,1;206,38,0,1;206,34,0,1;206,31,0,1;211,27,0,1;216,24,0,1;222,20,0,1;"
    strSQL = strSQL & "227,17,0,1;233,13,0,1;238,10,0,1;244,6,0,1;249,3,8,1;255,0,17,1;255,0,25,1;255,0,34,1;255,0,42,1;255,0,51,1;255,0,59,1;255,0,68,1;"
    strSQL = strSQL & "255,0,76,1;255,0,85,1;255,0,93,1;255,0,102,1;255,0,110,1;255,0,119,1;255,0,128,1;255,0,132,1;255,0,137,1;255,0,142,1;255,0,147,1;"
    strSQL = strSQL & "255,0,152,1;255,0,156,1;255,0,161,1;255,0,166,1;255,0,171,1;255,0,176,1;255,0,180,1;255,0,185,1;255,0,190,1;255,0,195,1;255,0,200,1;"
    strSQL = strSQL & "255,0,205,1;255,0,205,1;255,0,205,1;255,0,205,1;255,14,205,1;255,28,205,1;255,42,205,1;248,56,205,1;242,70,205,1;236,84,205,1;"
    strSQL = strSQL & "230,99,205,1;224,113,205,1;218,127,205,1;212,141,205,1;206,155,205,1;206,169,205,1;206,183,205,1;206,198,205,1;206,201,205,1;"
    strSQL = strSQL & "206,204,205,1;206,207,205,1;206,210,205,1;206,213,205,1;206,216,205,1;206,219,208,1;206,222,211,1;206,225,214,1;210,228,217,1;"
    strSQL = strSQL & "214,231,220,1;218,234,223,1;222,237,226,1;226,240,229,1;230,243,232,1;234,246,235,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (40,'step10','70,141,255,1;70,141,255,1;0,0,128,1;0,0,128,1;0,0,128,1;"
    strSQL = strSQL & "0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;"
    strSQL = strSQL & "0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;"
    strSQL = strSQL & "0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;0,94,0,1;"
    strSQL = strSQL & "0,94,0,1;0,94,0,1;0,94,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;"
    strSQL = strSQL & "64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;"
    strSQL = strSQL & "64,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;"
    strSQL = strSQL & "200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;"
    strSQL = strSQL & "200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;175,175,175,1;"
    strSQL = strSQL & "175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;"
    strSQL = strSQL & "175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;"
    strSQL = strSQL & "175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;175,175,175,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (41,'step20','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;"
    strSQL = strSQL & "0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,127,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;"
    strSQL = strSQL & "0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,191,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;"
    strSQL = strSQL & "0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,127,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;"
    strSQL = strSQL & "0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,191,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;"
    strSQL = strSQL & "127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;127,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;"
    strSQL = strSQL & "191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;191,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;"
    strSQL = strSQL & "127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;127,0,127,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;"
    strSQL = strSQL & "191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;191,0,191,1;255,0,255,1;255,0,255,1;255,0,255,1;"
    strSQL = strSQL & "255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;255,0,255,1;0,127,127,1;"
    strSQL = strSQL & "0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;0,127,127,1;"
    strSQL = strSQL & "0,127,127,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;0,191,191,1;"
    strSQL = strSQL & "0,191,191,1;0,191,191,1;0,191,191,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;"
    strSQL = strSQL & "0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;0,255,255,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;"
    strSQL = strSQL & "127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;127,63,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;"
    strSQL = strSQL & "191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;191,127,0,1;255,191,0,1;255,191,0,1;"
    strSQL = strSQL & "255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;255,191,0,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (42,'step4','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;"
    strSQL = strSQL & "165,42,42,1;165,42,42,1;165,42,42,1;165,42,42,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;"
    strSQL = strSQL & "219,147,112,1;219,147,112,1;219,147,112,1;219,147,112,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (43,'step5','0,0,0,1;0,0,254,1;0,0,255,1;0,0,254,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,254,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,254,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,254,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,254,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (44,'Thal_16','0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;"
    strSQL = strSQL & "0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,0,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;"
    strSQL = strSQL & "0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,128,1;0,0,192,1;0,0,192,1;0,0,192,1;"
    strSQL = strSQL & "0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;0,0,192,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "0,0,255,1;0,0,255,1;0,0,255,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;"
    strSQL = strSQL & "0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,64,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;"
    strSQL = strSQL & "0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,128,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;"
    strSQL = strSQL & "0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,192,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;0,255,0,1;"
    strSQL = strSQL & "64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;64,0,0,1;"
    strSQL = strSQL & "64,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;"
    strSQL = strSQL & "128,0,0,1;128,0,0,1;128,0,0,1;128,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;"
    strSQL = strSQL & "192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;192,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;"
    strSQL = strSQL & "255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;255,0,0,1;200,64,0,1;200,64,0,1;200,64,0,1;"
    strSQL = strSQL & "200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;200,64,0,1;"
    strSQL = strSQL & "200,64,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;"
    strSQL = strSQL & "200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;200,128,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;255,192,0,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (45,'Thal_256','0,0,0,1;0,0,8,1;0,0,16,1;0,0,24,1;0,0,32,1;0,0,40,1;"
    strSQL = strSQL & "0,0,48,1;0,0,56,1;0,0,64,1;0,0,72,1;0,0,80,1;0,0,88,1;0,0,96,1;0,0,104,1;0,0,112,1;0,0,120,1;0,0,128,1;0,0,132,1;0,0,136,1;0,0,140,1;"
    strSQL = strSQL & "0,0,144,1;0,0,148,1;0,0,152,1;0,0,156,1;0,0,160,1;0,0,164,1;0,0,168,1;0,0,172,1;0,0,176,1;0,0,180,1;0,0,184,1;0,0,188,1;0,0,192,1;"
    strSQL = strSQL & "0,0,196,1;0,0,200,1;0,0,204,1;0,0,208,1;0,0,212,1;0,0,216,1;0,0,220,1;0,0,224,1;0,0,228,1;0,0,232,1;0,0,236,1;0,0,240,1;0,0,244,1;"
    strSQL = strSQL & "0,0,248,1;0,0,252,1;0,0,255,1;0,4,240,1;0,8,224,1;0,12,208,1;0,16,192,1;0,20,176,1;0,24,160,1;0,28,144,1;0,32,128,1;0,36,112,1;"
    strSQL = strSQL & "0,40,96,1;0,44,80,1;0,48,64,1;0,52,48,1;0,56,32,1;0,60,16,1;0,64,0,1;0,68,0,1;0,72,0,1;0,76,0,1;0,80,0,1;0,84,0,1;0,88,0,1;0,92,0,1;"
    strSQL = strSQL & "0,96,0,1;0,100,0,1;0,104,0,1;0,108,0,1;0,112,0,1;0,116,0,1;0,120,0,1;0,124,0,1;0,128,0,1;0,132,0,1;0,136,0,1;0,140,0,1;0,144,0,1;"
    strSQL = strSQL & "0,148,0,1;0,152,0,1;0,156,0,1;0,160,0,1;0,164,0,1;0,168,0,1;0,172,0,1;0,176,0,1;0,180,0,1;0,184,0,1;0,188,0,1;0,192,0,1;0,196,0,1;"
    strSQL = strSQL & "0,200,0,1;0,204,0,1;0,208,0,1;0,212,0,1;0,216,0,1;0,220,0,1;0,224,0,1;0,228,0,1;0,232,0,1;0,236,0,1;0,240,0,1;0,244,0,1;0,248,0,1;"
    strSQL = strSQL & "0,252,0,1;0,255,0,1;4,240,0,1;8,224,0,1;12,208,0,1;16,192,0,1;20,176,0,1;24,160,0,1;28,144,0,1;32,128,0,1;36,112,0,1;40,96,0,1;"
    strSQL = strSQL & "44,80,0,1;48,64,0,1;52,48,0,1;56,32,0,1;60,16,0,1;64,0,0,1;68,0,0,1;72,0,0,1;76,0,0,1;80,0,0,1;84,0,0,1;88,0,0,1;92,0,0,1;96,0,0,1;"
    strSQL = strSQL & "100,0,0,1;104,0,0,1;108,0,0,1;112,0,0,1;116,0,0,1;120,0,0,1;124,0,0,1;128,0,0,1;132,0,0,1;136,0,0,1;140,0,0,1;144,0,0,1;148,0,0,1;"
    strSQL = strSQL & "152,0,0,1;156,0,0,1;160,0,0,1;164,0,0,1;168,0,0,1;172,0,0,1;176,0,0,1;180,0,0,1;184,0,0,1;188,0,0,1;192,0,0,1;196,0,0,1;200,0,0,1;"
    strSQL = strSQL & "204,0,0,1;208,0,0,1;212,0,0,1;216,0,0,1;220,0,0,1;224,0,0,1;228,0,0,1;232,0,0,1;236,0,0,1;240,0,0,1;244,0,0,1;248,0,0,1;252,0,0,1;"
    strSQL = strSQL & "255,0,0,1;251,4,0,1;248,8,0,1;244,12,0,1;240,16,0,1;237,20,0,1;233,24,0,1;230,28,0,1;226,32,0,1;222,36,0,1;219,40,0,1;215,44,0,1;"
    strSQL = strSQL & "211,48,0,1;209,52,0,1;206,56,0,1;202,60,0,1;200,64,0,1;200,68,0,1;200,72,0,1;200,76,0,1;200,80,0,1;200,84,0,1;200,88,0,1;200,92,0,1;"
    strSQL = strSQL & "200,96,0,1;200,100,0,1;200,104,0,1;200,108,0,1;200,112,0,1;200,116,0,1;200,120,0,1;200,124,0,1;200,128,0,1;202,132,0,1;206,136,0,1;"
    strSQL = strSQL & "209,140,0,1;211,144,0,1;215,148,0,1;219,152,0,1;222,156,0,1;226,160,0,1;230,164,0,1;233,168,0,1;237,172,0,1;240,176,0,1;244,180,0,1;"
    strSQL = strSQL & "248,184,0,1;251,188,0,1;255,192,0,1;255,196,16,1;255,200,32,1;255,204,48,1;255,208,64,1;255,212,80,1;255,216,96,1;255,220,112,1;"
    strSQL = strSQL & "255,224,128,1;255,228,144,1;255,232,160,1;255,236,176,1;255,240,192,1;255,244,208,1;255,248,224,1;255,252,240,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;"
    strSQL = strSQL & "255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (46,'Thermal-SP','0,0,0,1;0,0,0,1;2,2,2,1;3,0,0,1;4,4,4,1;10,0,0,1;"
    strSQL = strSQL & "6,6,6,1;17,0,0,1;8,8,8,1;24,1,0,1;10,10,10,1;30,1,0,1;12,12,12,1;36,2,0,1;14,14,14,1;42,2,0,1;16,16,16,1;48,2,0,1;18,18,18,1;"
    strSQL = strSQL & "54,2,0,1;20,20,20,1;60,3,0,1;22,22,22,1;66,3,0,1;24,24,24,1;72,4,0,1;26,26,26,1;78,4,0,1;28,28,28,1;84,4,0,1;30,30,30,1;86,6,0,1;"
    strSQL = strSQL & "32,32,32,1;88,9,0,1;34,34,34,1;92,10,0,1;36,36,36,1;96,12,0,1;38,38,38,1;98,14,0,1;40,40,40,1;100,17,0,1;42,42,42,1;104,18,0,1;"
    strSQL = strSQL & "44,44,44,1;108,19,0,1;46,46,46,1;110,21,0,1;48,48,48,1;112,24,0,1;50,50,50,1;116,25,0,1;52,52,52,1;120,27,0,1;54,54,54,1;122,29,0,1;"
    strSQL = strSQL & "56,56,56,1;124,32,0,1;58,58,58,1;128,33,0,1;60,60,60,1;132,34,0,1;62,62,62,1;134,36,0,1;64,64,64,1;136,39,0,1;66,66,66,1;140,40,2,1;"
    strSQL = strSQL & "68,68,68,1;144,41,4,1;70,70,70,1;146,43,4,1;72,72,72,1;148,46,4,1;74,74,74,1;152,47,4,1;76,76,76,1;156,49,4,1;78,78,78,1;158,51,4,1;"
    strSQL = strSQL & "80,80,80,1;160,54,4,1;82,82,82,1;164,54,6,1;84,84,84,1;168,55,8,1;86,86,86,1;170,57,8,1;88,88,88,1;172,60,8,1;90,90,90,1;176,61,8,1;"
    strSQL = strSQL & "92,92,92,1;180,63,8,1;94,94,94,1;182,65,8,1;96,96,96,1;184,68,8,1;98,98,98,1;186,70,8,1;100,100,100,1;188,73,8,1;102,102,102,1;"
    strSQL = strSQL & "190,75,8,1;104,104,104,1;192,77,8,1;106,106,106,1;194,79,8,1;108,108,108,1;196,82,8,1;110,110,110,1;200,83,10,1;112,112,112,1;"
    strSQL = strSQL & "204,84,12,1;114,114,114,1;206,86,12,1;116,116,116,1;208,89,12,1;118,118,118,1;210,91,12,1;120,120,120,1;212,94,12,1;122,122,122,1;"
    strSQL = strSQL & "214,96,12,1;124,124,124,1;216,98,12,1;126,126,126,1;218,100,12,1;128,128,128,1;220,103,12,1;130,130,130,1;224,104,12,1;132,132,132,1;"
    strSQL = strSQL & "228,106,12,1;134,134,134,1;230,108,14,1;136,136,136,1;232,110,16,1;138,138,138,1;234,112,16,1;140,140,140,1;236,115,16,1;142,142,142,1;"
    strSQL = strSQL & "238,117,16,1;144,144,144,1;240,119,16,1;146,146,146,1;242,121,16,1;148,148,148,1;244,124,16,1;150,150,150,1;248,125,18,1;152,152,152,1;"
    strSQL = strSQL & "252,126,20,1;154,154,154,1;252,129,20,1;156,156,156,1;252,133,20,1;158,158,158,1;252,136,20,1;160,160,160,1;252,140,20,1;162,162,162,1;"
    strSQL = strSQL & "252,143,22,1;164,164,164,1;252,146,24,1;166,166,166,1;252,149,24,1;168,168,168,1;252,153,24,1;170,170,170,1;252,156,26,1;172,172,172,1;"
    strSQL = strSQL & "252,159,28,1;174,174,174,1;252,161,32,1;176,176,176,1;252,164,36,1;178,178,178,1;252,166,40,1;180,180,180,1;252,169,44,1;182,182,182,1;"
    strSQL = strSQL & "252,172,48,1;184,184,184,1;252,175,52,1;186,186,186,1;252,177,56,1;188,188,188,1;252,180,60,1;190,190,190,1;252,183,62,1;192,192,192,1;"
    strSQL = strSQL & "252,186,64,1;194,194,194,1;252,188,68,1;196,196,196,1;252,191,72,1;198,198,198,1;252,193,76,1;200,200,200,1;252,196,80,1;202,202,202,1;"
    strSQL = strSQL & "252,199,84,1;204,204,204,1;252,202,88,1;206,206,206,1;252,204,92,1;208,208,208,1;252,207,96,1;210,210,210,1;252,209,100,1;212,212,212,1;"
    strSQL = strSQL & "252,212,104,1;214,214,214,1;252,214,110,1;216,216,216,1;252,217,116,1;218,218,218,1;252,219,122,1;220,220,220,1;252,221,128,1;"
    strSQL = strSQL & "222,222,222,1;252,223,134,1;224,224,224,1;252,226,140,1;226,226,226,1;252,228,146,1;228,228,228,1;252,230,152,1;230,230,230,1;"
    strSQL = strSQL & "252,232,158,1;232,232,232,1;252,235,164,1;234,234,234,1;252,237,170,1;236,236,236,1;252,239,176,1;238,238,238,1;252,241,184,1;"
    strSQL = strSQL & "240,240,240,1;252,243,192,1;242,242,242,1;252,244,202,1;244,244,244,1;252,246,212,1;246,246,246,1;252,247,222,1;248,248,248,1;"
    strSQL = strSQL & "252,249,232,1;250,250,250,1;252,250,242,1;252,252,252,1;252,252,252,1;254,254,254,1;252,252,252,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (47,'warm','0,0,0,1;0,0,0,1;0,0,0,1;0,0,1,1;0,0,1,1;0,0,2,1;0,0,3,1;"
    strSQL = strSQL & "0,0,4,1;0,0,5,1;0,0,6,1;0,0,8,1;0,0,10,1;0,0,12,1;0,0,14,1;0,0,16,1;0,0,18,1;1,0,21,1;1,0,23,1;1,0,26,1;1,0,29,1;1,0,32,1;1,0,35,1;"
    strSQL = strSQL & "1,0,39,1;1,0,42,1;2,0,46,1;2,0,50,1;2,0,54,1;2,1,58,1;3,1,63,1;3,1,67,1;3,1,72,1;4,1,77,1;4,1,89,1;4,1,95,1;5,1,101,1;5,1,107,1;"
    strSQL = strSQL & "6,1,113,1;6,1,118,1;7,2,124,1;7,2,129,1;8,2,134,1;8,2,139,1;9,2,143,1;10,2,148,1;10,2,152,1;11,3,156,1;12,3,159,1;13,3,163,1;"
    strSQL = strSQL & "14,3,166,1;14,0,169,1;15,0,172,1;16,0,175,1;    17,0,177,1;18,0,180,1;19,0,182,1;20,0,184,1;22,0,186,1;23,0,187,1;24,0,188,1;"
    strSQL = strSQL & "25,0,189,1;27,0,190,1;28,1,191,1;29,2,192,1;31,3,192,1;32,0,191,1;34,0,191,1;35,0,191,1;37,0,190,1;39,0,189,1;40,0,188,1;42,1,187,1;"
    strSQL = strSQL & "44,1,185,1;46,1,184,1;48,1,182,1;50,2,180,1;52,2,177,1;54,2,175,1;56,3,172,1;58,3,169,1;61,4,165,1;63,4,162,1;65,5,158,1;68,6,154,1;"
    strSQL = strSQL & "70,6,150,1;73,7,145,1;76,8,141,1;78,8,136,1;81,9,130,1;84,10,125,1;87,11,119,1;90,12,114,1;93,13,108,1;96,14,85,1;99,15,78,1;102,16,72,1;"
    strSQL = strSQL & "105,17,66,1;109,18,61,1;112,19,56,1;116,20,51,1;119,21,46,1;123,22,42,1;127,24,38,1;131,25,34,1;134,26,30,1;138,28,27,1;142,29,24,1;"
    strSQL = strSQL & "146,31,21,1;151,32,19,1;155,33,16,1;155,35,14,1;160,37,12,1;164,38,11,1;168,40,9,1;172,42,8,1;176,43,6,1;180,45,5,1;184,47,4,1;"
    strSQL = strSQL & "187,49,3,1;191,50,3,1;194,52,2,1;197,54,2,1;200,56,1,1;203,58,1,1;206,60,1,1;209,62,0,1;212,64,0,1;214,66,0,1;217,69,0,1;219,71,0,1;"
    strSQL = strSQL & "221,73,0,1;223,75,0,1;225,78,0,1;227,80,0,1;229,82,0,1;231,85,0,1;233,87,0,1;234,90,0,1;236,92,0,1;237,95,0,1;239,97,0,1;240,100,0,1;"
    strSQL = strSQL & "241,103,0,1;243,105,1,1;244,108,1,1;245,111,1,1;246,113,1,1;247,116,1,1;247,119,1,1;248,122,2,1;249,125,2,1;250,128,2,1;250,131,3,1;"
    strSQL = strSQL & "251,129,3,1;251,133,3,1;252,137,4,1;252,141,4,1;253,145,4,1;253,149,5,1;253,153,5,1;254,156,6,1;255,160,7,1;255,164,7,1;255,167,8,1;"
    strSQL = strSQL & "255,170,9,1;255,174,9,1;255,177,10,1;255,180,11,1;255,183,12,1;255,186,13,1;255,189,14,1;255,192,15,1;255,195,16,1;255,198,17,1;"
    strSQL = strSQL & "255,201,18,1;255,204,19,1;255,206,20,1;255,209,21,1;255,211,23,1;255,214,24,1;255,216,26,1;255,218,27,1;255,220,29,1;255,222,30,1;"
    strSQL = strSQL & "255,225,32,1;255,227,34,1;255,228,36,1;255,230,37,1;255,232,39,1;255,234,41,1;255,236,44,1;255,237,46,1;255,239,48,1;255,240,50,1;"
    strSQL = strSQL & "255,241,52,1;255,243,55,1;255,244,57,1;255,245,60,1;255,246,63,1;255,247,65,1;255,248,68,1;255,249,71,1;255,250,74,1;255,251,77,1;"
    strSQL = strSQL & "255,252,80,1;255,252,83,1;255,253,87,1;255,253,90,1;255,254,93,1;255,254,97,1;255,254,101,1;255,255,104,1;255,255,108,1;255,255,112,1;"
    strSQL = strSQL & "255,255,116,1;255,255,120,1;255,255,124,1;255,255,129,1;255,255,133,1;255,255,138,1;255,255,142,1;255,255,147,1;255,255,152,1;"
    strSQL = strSQL & "255,255,157,1;255,255,162,1;255,255,165,1;255,255,171,1;255,255,176,1;255,255,182,1;255,255,187,1;255,255,192,1;255,255,196,1;"
    strSQL = strSQL & "255,255,201,1;255,255,205,1;255,255,210,1;255,255,214,1;255,255,218,1;255,255,221,1;255,255,225,1;255,255,228,1;255,255,231,1;"
    strSQL = strSQL & "255,255,234,1;255,255,237,1;255,255,239,1;255,255,242,1;255,255,244,1;255,255,246,1;255,255,247,1;255,255,249,1;255,255,250,1;"
    strSQL = strSQL & "255,255,252,1;255,255,253,1;255,255,254,1;255,255,254,1;255,255,255,1;255,255,255,1;255,255,255,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (48,'White_Blue_Green_Red','255,255,255,1;250,252,255,1;246,250,255,1;"
    strSQL = strSQL & "242,248,255,1;238,246,255,1;234,244,255,1;230,242,255,1;226,240,255,1;222,238,255,1;218,236,255,1;214,234,255,1;210,232,255,1;"
    strSQL = strSQL & "206,230,255,1;202,228,255,1;198,226,255,1;194,224,255,1;190,222,255,1;186,220,255,1;182,218,255,1;178,216,255,1;174,214,255,1;"
    strSQL = strSQL & "169,212,255,1;165,210,255,1;161,208,255,1;157,206,255,1;153,204,255,1;149,202,255,1;145,200,255,1;141,198,255,1;137,196,255,1;"
    strSQL = strSQL & "133,194,255,1;129,192,255,1;125,190,255,1;121,188,255,1;117,186,255,1;113,184,255,1;109,182,255,1;105,180,255,1;101,178,255,1;"
    strSQL = strSQL & "97,176,255,1;93,174,255,1;89,172,255,1;84,170,255,1;80,168,255,1;76,166,255,1;72,164,255,1;68,162,255,1;64,160,255,1;60,158,255,1;"
    strSQL = strSQL & "56,156,255,1;52,154,255,1;48,152,255,1;44,150,255,1;40,148,255,1;36,146,255,1;32,144,255,1;28,142,255,1;24,140,255,1;20,138,255,1;"
    strSQL = strSQL & "16,136,255,1;12,134,255,1;8,132,255,1;4,130,255,1;0,128,255,1;0,125,255,1;0,128,251,1;0,130,247,1;0,132,243,1;0,134,239,1;0,136,235,1;"
    strSQL = strSQL & "0,138,231,1;0,140,227,1;0,142,223,1;0,144,219,1;0,146,215,1;0,148,211,1;0,150,207,1;0,152,203,1;0,154,199,1;0,156,195,1;0,158,191,1;"
    strSQL = strSQL & "0,160,187,1;0,162,183,1;0,164,179,1;0,166,175,1;0,168,171,1;0,170,167,1;0,172,163,1;0,174,159,1;0,176,155,1;0,178,151,1;0,180,147,1;"
    strSQL = strSQL & "0,182,143,1;0,184,139,1;0,186,135,1;0,188,131,1;0,190,127,1;0,192,123,1;0,194,119,1;0,196,115,1;0,198,111,1;0,200,107,1;0,202,103,1;"
    strSQL = strSQL & "0,204,99,1;0,206,95,1;0,208,91,1;0,210,87,1;0,212,83,1;0,214,79,1;0,216,75,1;0,218,71,1;0,220,67,1;0,222,63,1;0,224,59,1;0,226,55,1;"
    strSQL = strSQL & "0,228,51,1;0,230,47,1;0,232,43,1;0,234,39,1;0,236,35,1;0,238,31,1;0,240,27,1;0,242,23,1;0,244,19,1;0,246,15,1;0,248,11,1;0,250,7,1;"
    strSQL = strSQL & "0,252,3,1;0,255,0,1;3,255,2,1;7,255,4,1;11,255,6,1;15,255,8,1;19,255,10,1;23,255,12,1;27,255,14,1;31,255,16,1;35,255,18,1;39,255,20,1;"
    strSQL = strSQL & "43,255,22,1;47,255,24,1;51,255,26,1;55,255,28,1;59,255,30,1;63,255,32,1;67,255,34,1;71,255,36,1;75,255,38,1;79,255,40,1;83,255,42,1;"
    strSQL = strSQL & "87,255,44,1;91,255,46,1;95,255,48,1;99,255,50,1;103,255,52,1;107,255,54,1;111,255,56,1;115,255,58,1;119,255,60,1;123,255,62,1;"
    strSQL = strSQL & "127,255,64,1;131,255,66,1;135,255,68,1;139,255,70,1;143,255,72,1;147,255,74,1;151,255,76,1;155,255,78,1;159,255,80,1;163,255,82,1;"
    strSQL = strSQL & "167,255,84,1;171,255,86,1;175,255,88,1;179,255,90,1;183,255,92,1;187,255,94,1;191,255,96,1;195,255,98,1;199,255,100,1;203,255,102,1;"
    strSQL = strSQL & "207,255,104,1;211,255,106,1;215,255,108,1;219,255,110,1;223,255,112,1;227,255,114,1;231,255,116,1;235,255,118,1;239,255,120,1;"
    strSQL = strSQL & "243,255,122,1;247,255,124,1;251,255,126,1;255,255,128,1;255,251,126,1;255,247,124,1;255,243,122,1;255,239,120,1;255,235,118,1;"
    strSQL = strSQL & "255,231,116,1;255,227,114,1;255,223,112,1;255,219,110,1;255,215,108,1;255,211,106,1;255,207,104,1;255,203,102,1;255,199,100,1;"
    strSQL = strSQL & "255,195,98,1;255,191,96,1;255,187,94,1;255,183,92,1;255,179,90,1;255,175,88,1;255,171,86,1;255,167,84,1;255,163,82,1;255,159,80,1;"
    strSQL = strSQL & "255,155,78,1;255,151,76,1;255,147,74,1;255,143,72,1;255,139,70,1;255,135,68,1;255,131,66,1;255,127,64,1;255,123,62,1;255,119,60,1;"
    strSQL = strSQL & "255,115,58,1;255,111,56,1;255,107,54,1;255,103,52,1;255,99,50,1;255,95,48,1;255,91,46,1;255,87,44,1;255,83,42,1;255,79,40,1;"
    strSQL = strSQL & "255,75,38,1;255,71,36,1;255,67,34,1;255,63,32,1;255,59,30,1;255,55,28,1;255,51,26,1;255,47,24,1;255,43,22,1;255,39,20,1;"
    strSQL = strSQL & "255,35,18,1;255,31,16,1;255,27,14,1;255,23,12,1;255,19,10,1;255,15,8,1;255,11,6,1;255,7,4,1;255,3,2,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (49,'Yellow_WFG','0,0,0,1;0,0,0,1;106,0,124,1;104,0,126,1;103,0,128,1;"
    strSQL = strSQL & "101,0,130,1;99,0,132,1;97,0,134,1;96,0,136,1;94,0,138,1;92,0,141,1;91,0,143,1;89,0,145,1;87,0,147,1;85,0,149,1;84,0,151,1;82,0,153,1;"
    strSQL = strSQL & "80,0,155,1;79,0,157,1;77,0,160,1;75,0,162,1;73,0,164,1;72,0,166,1;70,0,168,1;68,0,170,1;67,0,172,1;65,0,174,1;63,0,176,1;61,0,179,1;"
    strSQL = strSQL & "60,0,181,1;58,0,183,1;56,0,185,1;55,0,187,1;53,0,189,1;51,0,191,1;49,0,193,1;48,0,195,1;46,0,198,1;44,0,200,1;42,0,202,1;41,0,204,1;"
    strSQL = strSQL & "39,0,206,1;37,0,208,1;36,0,210,1;34,0,212,1;32,0,214,1;30,0,217,1;29,0,219,1;27,0,221,1;25,0,223,1;24,0,225,1;22,0,227,1;20,0,229,1;"
    strSQL = strSQL & "18,0,231,1;17,0,233,1;15,0,236,1;13,0,238,1;12,0,240,1;10,0,242,1;8,0,244,1;6,0,246,1;5,0,248,1;3,0,250,1;1,0,252,1;0,0,255,1;0,0,255,1;"
    strSQL = strSQL & "2,0,251,1;5,0,247,1;8,0,243,1;11,0,239,1;14,0,235,1;17,0,231,1;20,0,227,1;23,0,223,1;26,0,219,1;29,0,215,1;32,0,211,1;35,0,207,1;"
    strSQL = strSQL & "38,0,203,1;41,0,199,1;44,0,195,1;47,0,191,1;50,0,187,1;53,0,183,1;56,0,179,1;59,0,175,1;62,0,171,1;65,0,167,1;68,0,163,1;71,0,159,1;"
    strSQL = strSQL & "74,0,155,1;77,0,151,1;80,0,147,1;83,0,143,1;85,0,139,1;88,0,135,1;91,0,131,1;94,0,127,1;97,0,123,1;100,0,119,1;103,0,115,1;106,0,111,1;"
    strSQL = strSQL & "109,3,107,1;112,6,103,1;115,9,99,1;118,12,95,1;121,15,91,1;124,19,87,1;127,22,83,1;130,25,79,1;133,28,75,1;136,31,71,1;139,35,67,1;"
    strSQL = strSQL & "142,38,63,1;145,41,59,1;148,44,55,1;151,47,51,1;154,51,47,1;157,54,43,1;160,57,39,1;163,60,35,1;166,63,31,1;169,66,27,1;171,70,23,1;"
    strSQL = strSQL & "174,73,19,1;177,76,15,1;180,79,11,1;183,82,7,1;186,86,3,1;189,89,0,1;192,92,0,1;195,95,0,1;198,98,0,1;201,102,0,1;204,105,0,1;"
    strSQL = strSQL & "207,108,0,1;210,111,0,1;213,114,0,1;216,117,0,1;219,121,0,1;222,124,0,1;225,127,0,1;228,130,0,1;231,133,0,1;234,137,0,1;237,140,0,1;"
    strSQL = strSQL & "240,143,0,1;243,146,0,1;246,149,0,1;249,153,0,1;252,156,0,1;255,159,0,1;255,162,0,1;255,165,0,1;255,168,0,1;255,172,0,1;255,175,0,1;"
    strSQL = strSQL & "255,178,0,1;255,181,0,1;255,184,0,1;255,188,0,1;255,191,0,1;255,194,0,1;255,197,0,1;255,200,0,1;255,204,0,1;255,207,0,1;255,210,0,1;"
    strSQL = strSQL & "255,213,0,1;255,216,0,1;255,219,0,1;255,223,0,1;255,226,0,1;255,229,0,1;255,232,0,1;255,235,0,1;255,239,0,1;255,242,0,1;255,245,0,1;"
    strSQL = strSQL & "255,248,0,1;255,251,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;255,255,0,1;"
    strSQL = strSQL & "255,255,0,1;253,255,4,1;251,255,9,1;249,255,13,1;248,255,18,1;246,255,23,1;244,255,27,1;242,255,32,1;241,255,37,1;239,255,41,1;"
    strSQL = strSQL & "237,255,46,1;235,255,51,1;234,255,55,1;232,255,60,1;230,255,64,1;228,255,69,1;227,255,74,1;225,255,78,1;223,255,83,1;221,255,88,1;"
    strSQL = strSQL & "220,255,92,1;221,255,97,1;222,255,101,1;223,255,106,1;224,255,111,1;225,255,115,1;226,255,120,1;227,255,125,1;228,255,129,1;"
    strSQL = strSQL & "229,255,134,1;230,255,139,1;231,255,143,1;232,255,148,1;233,255,152,1;234,255,157,1;235,255,162,1;236,255,166,1;237,255,171,1;"
    strSQL = strSQL & "238,255,176,1;239,255,180,1;240,255,185,1;241,255,190,1;242,255,194,1;243,255,199,1;244,255,204,1;245,255,208,1;246,255,213,1;"
    strSQL = strSQL & "247,255,217,1;248,255,222,1;249,255,227,1;250,255,231,1;251,255,236,1;252,255,241,1;253,255,245,1;254,255,250,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (50,'PRISM-DWS','0,0,0,1;1,1,1,1;2,2,2,1;3,3,3,1;4,4,4,1;5,5,5,1;6,6,6,1;"
    strSQL = strSQL & "7,7,7,1;8,8,8,1;9,9,9,1;10,10,10,1;11,11,11,1;12,12,12,1;13,13,13,1;14,14,14,1;15,15,15,1;16,16,16,1;17,17,17,1;18,18,18,1;19,19,19,1;"
    strSQL = strSQL & "20,20,20,1;21,21,21,1;22,22,22,1;23,23,23,1;24,24,24,1;25,25,25,1;26,26,26,1;27,27,27,1;28,28,28,1;29,29,29,1;30,30,30,1;31,31,31,1;"
    strSQL = strSQL & "32,32,32,1;33,33,33,1;34,34,34,1;35,35,35,1;36,36,36,1;37,37,37,1;38,38,38,1;39,39,39,1;40,40,40,1;41,41,41,1;42,42,42,1;43,43,43,1;"
    strSQL = strSQL & "44,44,44,1;45,45,45,1;46,46,46,1;47,47,47,1;48,48,48,1;49,49,49,1;50,50,50,1;51,51,51,1;52,52,52,1;53,53,53,1;54,54,54,1;55,55,55,1;"
    strSQL = strSQL & "56,56,56,1;57,57,57,1;58,58,58,1;59,59,59,1;60,60,60,1;61,61,61,1;62,62,62,1;63,63,63,1;64,64,64,1;65,65,65,1;66,66,66,1;67,67,67,1;"
    strSQL = strSQL & "68,68,68,1;69,69,69,1;70,70,70,1;71,71,71,1;72,72,72,1;73,73,73,1;74,74,74,1;75,75,75,1;76,76,76,1;77,77,77,1;78,78,78,1;79,79,79,1;"
    strSQL = strSQL & "80,80,80,1;81,81,81,1;82,82,82,1;83,83,83,1;84,84,84,1;85,85,85,1;86,86,86,1;87,87,87,1;88,88,88,1;89,89,89,1;90,90,90,1;91,91,91,1;"
    strSQL = strSQL & "92,92,92,1;93,93,93,1;94,94,94,1;95,95,95,1;96,96,96,1;97,97,97,1;98,98,98,1;99,99,99,1;100,100,100,1;101,93,95,1;101,86,91,1;"
    strSQL = strSQL & "101,80,86,1;101,74,82,1;101,68,79,1;101,63,75,1;101,59,72,1;101,55,68,1;101,51,66,1;102,48,63,1;102,46,60,1;102,44,58,1;102,42,56,1;"
    strSQL = strSQL & "103,41,54,1;103,40,52,1;104,40,50,1;104,38,48,1;105,37,47,1;106,36,46,1;106,35,45,1;107,34,44,1;108,33,43,1;109,32,42,1;110,30,41,1;"
    strSQL = strSQL & "112,29,40,1;113,28,40,1;115,27,40,1;116,26,39,1;118,25,39,1;120,23,39,1;122,22,38,1;123,21,38,1;126,20,38,1;129,19,38,1;131,18,38,1;"
    strSQL = strSQL & "133,16,38,1;136,15,38,1;139,14,38,1;141,15,36,1;143,17,33,1;145,19,31,1;148,21,29,1;150,23,26,1;152,25,24,1;155,27,22,1;157,29,19,1;"
    strSQL = strSQL & "160,31,17,1;162,33,15,1;164,35,12,1;166,37,10,1;168,39,8,1;172,41,5,1;174,43,3,1;177,45,4,1;179,47,3,1;181,49,3,1;183,51,1,1;"
    strSQL = strSQL & "186,53,0,1;188,55,0,1;190,57,0,1;193,59,0,1;195,61,3,1;197,63,0,1;200,65,0,1;202,67,3,1;204,69,1,1;206,71,1,1;209,73,5,1;211,75,5,1;"
    strSQL = strSQL & "213,76,3,1;216,80,3,1;218,84,3,1;221,88,2,1;221,92,1,1;221,96,1,1;221,100,7,1;222,104,5,1;222,108,5,1;223,113,3,1;223,117,2,1;"
    strSQL = strSQL & "224,121,0,1;225,125,0,1;225,129,0,1;226,133,0,1;226,137,0,1;227,141,0,1;228,145,0,1;229,149,0,1;229,153,0,1;230,157,0,1;"
    strSQL = strSQL & "230,161,0,1;231,165,0,1;231,169,0,1;232,173,0,1;233,177,0,1;233,181,0,1;234,186,0,1;235,190,0,1;235,194,0,1;236,198,0,1;236,202,0,1;"
    strSQL = strSQL & "237,206,0,1;237,210,0,1;238,214,0,1;239,218,0,1;240,221,0,1;240,223,0,1;241,224,0,1;242,225,0,1;243,226,0,1;244,227,0,1;244,227,0,1;"
    strSQL = strSQL & "245,229,0,1;246,229,0,1;246,231,0,1;247,231,0,1;248,233,0,1;249,233,0,1;249,235,0,1;253,235,0,1;253,235,0,1;253,234,0,1;253,235,0,1;"
    strSQL = strSQL & "253,235,0,1;253,237,0,1;253,237,0,1;253,239,0,1;253,239,0,1;253,240,0,1;253,241,0,1;253,241,0,1;253,242,0,1;253,243,0,1;253,243,0,1;"
    strSQL = strSQL & "253,245,0,1;253,245,0,1;253,246,0,1;253,247,0,1;253,247,0,1;253,248,31,1;253,249,61,1;253,249,92,1;253,251,123,1;253,251,153,1;"
    strSQL = strSQL & "253,253,205,1;253,253,211,1;253,253,217,1;253,253,223,1;253,253,229,1;253,253,235,1;253,253,241,1;253,253,247,1;253,253,253,1;"
    strSQL = strSQL & "254,1,128,1;255,0,128,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (51,'InvertGray','255,255,255,1;254,254,254,1;253,253,253,1;252,252,252,1;"
    strSQL = strSQL & "251,251,251,1;250,250,250,1;249,249,249,1;248,248,248,1;247,247,247,1;246,246,246,1;245,245,245,1;244,244,244,1;243,243,243,1;"
    strSQL = strSQL & "242,242,242,1;241,241,241,1;240,240,240,1;239,239,239,1;238,238,238,1;237,237,237,1;236,236,236,1;235,235,235,1;234,234,234,1;"
    strSQL = strSQL & "233,233,233,1;232,232,232,1;231,231,231,1;230,230,230,1;229,229,229,1;228,228,228,1;227,227,227,1;226,226,226,1;225,225,225,1;"
    strSQL = strSQL & "224,224,224,1;223,223,223,1;222,222,222,1;221,221,221,1;220,220,220,1;219,219,219,1;218,218,218,1;217,217,217,1;216,216,216,1;"
    strSQL = strSQL & "215,215,215,1;214,214,214,1;213,213,213,1;212,212,212,1;211,211,211,1;210,210,210,1;209,209,209,1;208,208,208,1;207,207,207,1;"
    strSQL = strSQL & "206,206,206,1;205,205,205,1;204,204,204,1;203,203,203,1;202,202,202,1;201,201,201,1;200,200,200,1;199,199,199,1;198,198,198,1;"
    strSQL = strSQL & "197,197,197,1;196,196,196,1;195,195,195,1;194,194,194,1;193,193,193,1;192,192,192,1;191,191,191,1;190,190,190,1;189,189,189,1;"
    strSQL = strSQL & "188,188,188,1;187,187,187,1;186,186,186,1;185,185,185,1;184,184,184,1;183,183,183,1;182,182,182,1;181,181,181,1;180,180,180,1;"
    strSQL = strSQL & "179,179,179,1;178,178,178,1;177,177,177,1;176,176,176,1;175,175,175,1;174,174,174,1;173,173,173,1;172,172,172,1;171,171,171,1;"
    strSQL = strSQL & "170,170,170,1;169,169,169,1;168,168,168,1;167,167,167,1;166,166,166,1;165,165,165,1;164,164,164,1;163,163,163,1;162,162,162,1;"
    strSQL = strSQL & "161,161,161,1;160,160,160,1;159,159,159,1;158,158,158,1;157,157,157,1;156,156,156,1;155,155,155,1;154,154,154,1;153,153,153,1;"
    strSQL = strSQL & "152,152,152,1;151,151,151,1;150,150,150,1;149,149,149,1;148,148,148,1;147,147,147,1;146,146,146,1;145,145,145,1;144,144,144,1;"
    strSQL = strSQL & "143,143,143,1;142,142,142,1;141,141,141,1;140,140,140,1;139,139,139,1;138,138,138,1;137,137,137,1;136,136,136,1;135,135,135,1;"
    strSQL = strSQL & "134,134,134,1;133,133,133,1;132,132,132,1;131,131,131,1;130,130,130,1;129,129,129,1;128,128,128,1;127,127,127,1;126,126,126,1;"
    strSQL = strSQL & "125,125,125,1;124,124,124,1;123,123,123,1;122,122,122,1;121,121,121,1;120,120,120,1;119,119,119,1;118,118,118,1;117,117,117,1;"
    strSQL = strSQL & "116,116,116,1;115,115,115,1;114,114,114,1;113,113,113,1;112,112,112,1;111,111,111,1;110,110,110,1;109,109,109,1;108,108,108,1;"
    strSQL = strSQL & "107,107,107,1;106,106,106,1;105,105,105,1;104,104,104,1;103,103,103,1;102,102,102,1;101,101,101,1;100,100,100,1;99,99,99,1;"
    strSQL = strSQL & "98,98,98,1;97,97,97,1;96,96,96,1;95,95,95,1;94,94,94,1;93,93,93,1;92,92,92,1;91,91,91,1;90,90,90,1;89,89,89,1;88,88,88,1;"
    strSQL = strSQL & "87,87,87,1;86,86,86,1;85,85,85,1;84,84,84,1;83,83,83,1;82,82,82,1;81,81,81,1;80,80,80,1;79,79,79,1;78,78,78,1;77,77,77,1;"
    strSQL = strSQL & "76,76,76,1;75,75,75,1;74,74,74,1;73,73,73,1;72,72,72,1;71,71,71,1;70,70,70,1;69,69,69,1;68,68,68,1;67,67,67,1;66,66,66,1;"
    strSQL = strSQL & "65,65,65,1;64,64,64,1;63,63,63,1;62,62,62,1;61,61,61,1;60,60,60,1;59,59,59,1;58,58,58,1;57,57,57,1;56,56,56,1;55,55,55,1;"
    strSQL = strSQL & "54,54,54,1;53,53,53,1;52,52,52,1;51,51,51,1;50,50,50,1;49,49,49,1;48,48,48,1;47,47,47,1;46,46,46,1;45,45,45,1;44,44,44,1;"
    strSQL = strSQL & "43,43,43,1;42,42,42,1;41,41,41,1;40,40,40,1;39,39,39,1;38,38,38,1;37,37,37,1;36,36,36,1;35,35,35,1;34,34,34,1;33,33,33,1;"
    strSQL = strSQL & "32,32,32,1;31,31,31,1;30,30,30,1;29,29,29,1;28,28,28,1;27,27,27,1;26,26,26,1;25,25,25,1;24,24,24,1;23,23,23,1;22,22,22,1;"
    strSQL = strSQL & "21,21,21,1;20,20,20,1;19,19,19,1;18,18,18,1;17,17,17,1;16,16,16,1;15,15,15,1;14,14,14,1;13,13,13,1;12,12,12,1;11,11,11,1;"
    strSQL = strSQL & "10,10,10,1;9,9,9,1;8,8,8,1;7,7,7,1;6,6,6,1;5,5,5,1;4,4,4,1;3,3,3,1;2,2,2,1;1,1,1,1;0,0,0,1;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (62,'dfffff','0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;"
    strSQL = strSQL & "0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;"
    strSQL = strSQL & "0,0,0,0;0,0,0,0;0,0,0,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;"
    strSQL = strSQL & "37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;"
    strSQL = strSQL & "37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;"
    strSQL = strSQL & "0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;"
    strSQL = strSQL & "0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;"
    strSQL = strSQL & "0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;"
    strSQL = strSQL & "0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,214,20,0;0,214,20,0;0,214,20,0;"
    strSQL = strSQL & "0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;"
    strSQL = strSQL & "0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;183,214,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;"
    strSQL = strSQL & "255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;"
    strSQL = strSQL & "255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;"
    strSQL = strSQL & "255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;"
    strSQL = strSQL & "255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,5,5,0;255,5,5,0;"
    strSQL = strSQL & "255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;"
    strSQL = strSQL & "255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;255,172,209,0;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (63,'ddeeee','0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;"
    strSQL = strSQL & "0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;0,0,0,0;"
    strSQL = strSQL & "0,0,0,0;0,0,0,0;0,0,0,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;"
    strSQL = strSQL & "37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;"
    strSQL = strSQL & "37,0,118,0;37,0,118,0;37,0,118,0;37,0,118,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;"
    strSQL = strSQL & "0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;"
    strSQL = strSQL & "0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,23,188,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;"
    strSQL = strSQL & "0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;"
    strSQL = strSQL & "0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,115,88,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;"
    strSQL = strSQL & "0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;"
    strSQL = strSQL & "0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;0,214,20,0;183,214,0,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;183,214,0,0;"
    strSQL = strSQL & "183,214,0,0;183,214,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;"
    strSQL = strSQL & "255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;"
    strSQL = strSQL & "255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,183,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;"
    strSQL = strSQL & "255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;"
    strSQL = strSQL & "255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,93,0,0;255,5,5,0;255,5,5,0;255,5,5,0;"
    strSQL = strSQL & "255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;"
    strSQL = strSQL & "255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,5,5,0;255,172,209,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;255,172,209,0;"
    strSQL = strSQL & "255,172,209,0;255,172,209,0;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (64,'blue_se','0,0,0,0;0,0,9,0;0,0,19,0;0,0,29,0;0,1,38,0;0,1,47,0;"
    strSQL = strSQL & "0,2,55,0;0,3,63,0;0,4,71,0;0,5,79,0;0,6,87,0;0,7,94,0;0,8,100,0;0,9,107,0;0,11,113,0;0,12,118,0;0,13,124,0;0,15,128,0;0,16,133,0;"
    strSQL = strSQL & "0,17,137,0;0,19,141,0;0,20,144,0;0,21,147,0;0,22,149,0;0,24,151,0;0,25,153,0;0,26,154,0;0,27,155,0;0,28,156,0;0,29,157,0;0,30,158,0;"
    strSQL = strSQL & "0,31,159,0;0,32,160,0;0,33,161,0;0,34,162,0;0,35,163,0;0,36,164,0;0,37,165,0;0,38,166,0;0,39,167,0;0,40,168,0;0,41,169,0;0,42,170,0;"
    strSQL = strSQL & "0,43,171,0;0,44,172,0;0,45,173,0;0,46,174,0;0,47,175,0;0,48,176,0;0,49,177,0;0,50,178,0;0,51,179,0;0,52,180,0;0,53,181,0;0,54,182,0;"
    strSQL = strSQL & "0,55,183,0;0,56,184,0;0,57,185,0;0,58,186,0;0,59,187,0;0,60,188,0;0,61,189,0;0,62,190,0;0,63,191,0;0,64,192,0;0,65,193,0;0,66,194,0;"
    strSQL = strSQL & "0,67,195,0;0,68,196,0;0,69,197,0;0,70,198,0;0,71,199,0;0,72,200,0;0,73,201,0;0,74,202,0;0,75,203,0;0,76,204,0;0,77,205,0;0,78,206,0;"
    strSQL = strSQL & "0,79,207,0;0,80,208,0;0,81,209,0;0,82,210,0;0,83,211,0;0,84,212,0;0,85,213,0;0,86,214,0;0,87,215,0;0,88,216,0;0,89,217,0;0,90,218,0;"
    strSQL = strSQL & "0,91,219,0;0,92,220,0;0,93,221,0;0,94,222,0;0,95,223,0;0,96,224,0;0,97,225,0;0,98,226,0;0,99,227,0;0,100,228,0;0,101,229,0;0,102,230,0;"
    strSQL = strSQL & "0,103,231,0;0,104,232,0;0,105,233,0;0,106,234,0;0,107,235,0;0,108,236,0;0,109,237,0;0,110,238,0;0,111,239,0;0,112,240,0;0,113,241,0;"
    strSQL = strSQL & "0,114,242,0;0,115,243,0;0,116,244,0;0,117,245,0;0,118,246,0;0,119,247,0;0,120,248,0;0,121,249,0;0,122,250,0;0,123,251,0;0,124,252,0;"
    strSQL = strSQL & "0,125,253,0;0,126,254,0;0,127,255,0;0,128,255,0;0,129,255,0;0,130,255,0;0,131,255,0;0,132,255,0;0,133,255,0;0,134,255,0;0,135,255,0;"
    strSQL = strSQL & "0,136,255,0;0,137,255,0;0,138,255,0;0,139,255,0;0,140,255,0;0,141,255,0;0,142,255,0;0,143,255,0;0,144,255,0;0,145,255,0;0,146,255,0;"
    strSQL = strSQL & "0,147,255,0;0,148,255,0;0,149,255,0;0,150,255,0;0,151,255,0;0,152,255,0;0,153,255,0;0,154,255,0;0,155,255,0;0,156,255,0;0,157,255,0;"
    strSQL = strSQL & "0,158,255,0;0,159,255,0;0,160,255,0;0,161,255,0;0,162,255,0;0,163,255,0;0,164,255,0;0,165,255,0;0,166,255,0;0,167,255,0;0,168,255,0;"
    strSQL = strSQL & "0,169,255,0;0,170,255,0;0,171,255,0;0,172,255,0;0,173,255,0;0,174,255,0;0,175,255,0;0,176,255,0;0,177,255,0;0,178,255,0;0,179,255,0;"
    strSQL = strSQL & "0,180,255,0;0,181,255,0;0,182,255,0;0,183,255,0;0,184,255,0;0,185,255,0;0,186,255,0;0,187,255,0;0,188,255,0;0,189,255,0;0,190,255,0;"
    strSQL = strSQL & "0,191,255,0;0,192,255,0;0,193,255,0;0,194,255,0;0,195,255,0;0,196,255,0;0,197,255,0;0,198,255,0;0,199,255,0;0,200,255,0;0,201,255,0;"
    strSQL = strSQL & "0,202,255,0;0,203,255,0;0,204,255,0;0,205,255,0;0,206,255,0;0,207,255,0;0,208,255,0;0,209,255,0;0,210,255,0;0,211,255,0;0,212,255,0;"
    strSQL = strSQL & "0,213,255,0;0,214,255,0;0,215,255,0;0,216,255,0;0,217,255,0;0,218,255,0;0,219,255,0;0,220,255,0;0,221,255,0;0,222,255,0;0,223,255,0;"
    strSQL = strSQL & "0,224,255,0;0,225,255,0;0,226,255,0;0,227,255,0;0,228,255,0;0,229,255,0;0,230,255,0;0,231,255,0;2,233,255,0;4,234,255,0;7,235,255,0;"
    strSQL = strSQL & "11,236,255,0;16,238,255,0;22,239,255,0;28,240,255,0;35,242,255,0;43,243,255,0;52,244,255,0;61,246,255,0;71,247,255,0;82,248,255,0;"
    strSQL = strSQL & "94,249,255,0;107,250,255,0;120,251,255,0;134,252,255,0;149,253,255,0;165,254,255,0;181,254,255,0;199,255,255,0;217,255,255,0;"
    strSQL = strSQL & "236,255,255,0;255,255,255,0;',1);"
    cnAccess.Execute strSQL
    
    
    strSQL = "insert into 影像颜色清单 (序号,颜色,颜色内容,系统方案) values (67,'blue_green','0,0,0,0;0,0,9,0;0,0,19,0;0,0,29,0;0,1,38,0;0,1,47,0;"
    strSQL = strSQL & "0,2,55,0;0,3,63,0;255,255,128,0;251,253,126,0;248,251,124,0;244,250,123,0;241,248,121,0;237,246,119,0;234,244,117,0;230,243,116,0;"
    strSQL = strSQL & "227,241,114,0;223,239,112,0;220,237,110,0;216,236,108,0;212,234,107,0;209,232,105,0;205,230,103,0;202,229,101,0;198,227,100,0;"
    strSQL = strSQL & "195,225,98,0;191,223,96,0;188,221,94,0;184,220,92,0;181,218,91,0;177,216,89,0;174,214,87,0;170,213,85,0;166,211,84,0;163,209,82,0;"
    strSQL = strSQL & "159,207,80,0;156,206,78,0;152,204,76,0;149,202,75,0;145,200,73,0;142,199,71,0;138,197,69,0;135,195,68,0;131,193,66,0;128,192,64,0;"
    strSQL = strSQL & "124,190,62,0;120,188,60,0;117,186,59,0;113,184,57,0;110,183,55,0;106,181,53,0;103,179,52,0;99,177,50,0;96,176,48,0;92,174,46,0;"
    strSQL = strSQL & "89,172,44,0;85,170,43,0;81,169,41,0;78,167,39,0;74,165,37,0;71,163,36,0;67,162,34,0;64,160,32,0;60,158,30,0;57,156,28,0;53,154,27,0;"
    strSQL = strSQL & "50,153,25,0;46,151,23,0;42,149,21,0;39,147,20,0;35,146,18,0;32,144,16,0;28,142,14,0;25,140,12,0;21,139,11,0;18,137,9,0;14,135,7,0;"
    strSQL = strSQL & "11,133,5,0;7,132,4,0;4,130,2,0;0,128,0,0;0,81,209,0;0,82,210,0;0,83,211,0;0,84,212,0;0,85,213,0;0,86,214,0;0,87,215,0;0,88,216,0;"
    strSQL = strSQL & "0,89,217,0;0,90,218,0;0,91,219,0;0,92,220,0;0,93,221,0;0,94,222,0;0,95,223,0;0,96,224,0;0,97,225,0;128,128,128,0;131,131,128,0;"
    strSQL = strSQL & "134,134,128,0;137,137,128,0;140,140,128,0;143,143,128,0;146,146,128,0;149,149,128,0;152,152,128,0;155,155,128,0;158,158,128,0;"
    strSQL = strSQL & "161,161,128,0;164,164,128,0;167,167,128,0;170,170,128,0;173,173,128,0;176,176,128,0;179,179,128,0;182,182,128,0;185,185,128,0;"
    strSQL = strSQL & "188,188,128,0;192,192,128,0;195,195,128,0;198,198,128,0;201,201,128,0;204,204,128,0;207,207,128,0;210,210,128,0;213,213,128,0;"
    strSQL = strSQL & "216,216,128,0;219,219,128,0;222,222,128,0;225,225,128,0;228,228,128,0;231,231,128,0;234,234,128,0;237,237,128,0;240,240,128,0;"
    strSQL = strSQL & "243,243,128,0;246,246,128,0;249,249,128,0;252,252,128,0;255,255,128,0;0,141,255,0;0,142,255,0;0,143,255,0;0,144,255,0;0,145,255,0;"
    strSQL = strSQL & "0,146,255,0;0,147,255,0;0,148,255,0;0,149,255,0;0,150,255,0;0,151,255,0;0,152,255,0;0,153,255,0;0,154,255,0;0,155,255,0;0,156,255,0;"
    strSQL = strSQL & "0,157,255,0;0,158,255,0;0,159,255,0;255,128,255,0;249,122,255,0;242,115,255,0;236,109,255,0;230,102,255,0;223,96,255,0;217,90,255,0;"
    strSQL = strSQL & "211,83,255,0;204,77,255,0;198,70,255,0;192,64,255,0;185,58,255,0;179,51,255,0;172,45,255,0;166,38,255,0;160,32,255,0;153,26,255,0;"
    strSQL = strSQL & "147,19,255,0;141,13,255,0;134,6,255,0;128,0,255,0;0,181,255,0;0,182,255,0;0,183,255,0;0,184,255,0;0,185,255,0;0,186,255,0;0,187,255,0;"
    strSQL = strSQL & "0,188,255,0;0,189,255,0;0,190,255,0;0,191,255,0;0,192,255,0;0,193,255,0;0,194,255,0;0,195,255,0;0,196,255,0;0,197,255,0;0,198,255,0;"
    strSQL = strSQL & "0,199,255,0;0,200,255,0;0,201,255,0;0,202,255,0;0,203,255,0;0,204,255,0;0,205,255,0;0,206,255,0;0,207,255,0;0,208,255,0;0,209,255,0;"
    strSQL = strSQL & "0,210,255,0;0,211,255,0;0,212,255,0;0,213,255,0;0,214,255,0;0,215,255,0;0,216,255,0;0,217,255,0;0,218,255,0;0,219,255,0;0,220,255,0;"
    strSQL = strSQL & "0,221,255,0;0,222,255,0;0,223,255,0;0,224,255,0;0,225,255,0;0,226,255,0;0,227,255,0;0,228,255,0;0,229,255,0;0,230,255,0;0,231,255,0;"
    strSQL = strSQL & "2,233,255,0;4,234,255,0;7,235,255,0;11,236,255,0;16,238,255,0;22,239,255,0;28,240,255,0;35,242,255,0;43,243,255,0;52,244,255,0;"
    strSQL = strSQL & "61,246,255,0;71,247,255,0;82,248,255,0;94,249,255,0;107,250,255,0;120,251,255,0;134,252,255,0;149,253,255,0;165,254,255,0;"
    strSQL = strSQL & "181,254,255,0;199,255,255,0;217,255,255,0;236,255,255,0;255,255,255,0;',1);"
    cnAccess.Execute strSQL
    
    '--影像打印格式
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (1,'STANDARD',1,1,0,0,0,0,0,'STANDARD\1,1');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (2,'STANDARD',1,2,0,0,0,0,0,'STANDARD\1,2');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (3,'STANDARD',2,2,0,0,0,0,0,'STANDARD\2,2');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (7,'STANDARD',5,5,0,0,0,0,0,'STANDARD\5,5');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (8,'STANDARD',2,1,0,0,0,0,0,'STANDARD\2,1');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (9,'STANDARD',2,3,0,0,0,0,0,'STANDARD\2,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (10,'STANDARD',3,2,0,0,0,0,0,'STANDARD\3,2');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (11,'STANDARD',3,3,0,0,0,0,0,'STANDARD\3,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (12,'STANDARD',3,4,0,0,0,0,0,'STANDARD\3,4');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (13,'STANDARD',4,3,0,0,0,0,0,'STANDARD\4,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (14,'STANDARD',4,4,0,0,0,0,0,'STANDARD\4,4');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (15,'STANDARD',5,3,0,0,0,0,0,'STANDARD\5,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (16,'STANDARD',5,4,0,0,0,0,0,'STANDARD\5,4');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (17,'STANDARD',5,5,0,0,0,0,0,'STANDARD\5,5');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (18,'STANDARD',6,4,0,0,0,0,0,'STANDARD\6,4');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (19,'ROW',1,2,0,0,0,0,0,'ROW\1,2');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (20,'ROW',2,1,0,0,0,0,0,'ROW\2,1');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (22,'ROW',2,3,0,0,0,0,0,'ROW\2,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (23,'ROW',3,2,0,0,0,0,0,'ROW\3,2');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (24,'COL',2,3,0,0,0,0,0,'COL\2,3');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像打印格式 (编号,格式,参数1,参数2,参数3,参数4,参数5,参数6,参数7,格式标识) values (25,'COL',3,2,0,0,0,0,0,'COL\3,2');"
    cnAccess.Execute strSQL
    
    '--影像胶片规格
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (1,'8','10','IN','8INX10IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (2,'8.5','11','IN','8_5INX11IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (3,'10','12','IN','10INX12IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (4,'10','14','IN','10INX14IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (5,'11','14','IN','11INX14IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (6,'11','17','IN','11INX17IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (7,'14','14','IN','14INX14IN');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像胶片规格 (编号,胶片宽度,胶片长度,单位,规格标识) values (8,'14','17','IN','14INX17IN');"
    cnAccess.Execute strSQL
    
    '--影像标注存储表
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (1,'6001','10','LO','TPRODUCER');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (2,'6001','20','LO','TSUM');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (3,'6001','10C0','SQ','TALL');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (4,'6001','2000','SQ','TONE');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (5,'6001','2001','LO','TINDEX');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (6,'6001','2002','LO','Left');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (7,'6001','2003','LO','Top');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (8,'6001','2004','LO','Width');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (9,'6001','2005','LO','Height');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (10,'6001','2006','LO','ImageTied');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (11,'6001','2007','LO','LabelType');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (12,'6001','2008','LO','Alignment');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (13,'6001','2009','LO','AnchorImageTied');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (14,'6001','2010','LO','AnchorX');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (15,'6001','2011','LO','AnchorY');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (16,'6001','2012','LO','Angle');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (17,'6001','2013','LO','AutoSize');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (18,'6001','2014','LO','BackColour');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (19,'6001','2015','LO','BackStyle');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (20,'6001','2016','LO','BoundingPolygon');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (21,'6001','2017','LO','FontName');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (22,'6001','2018','LO','FontSize');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (23,'6001','2019','LO','ForeColour');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (24,'6001','2020','LO','LineStyle');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (25,'6001','2021','LO','LineWidth');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (26,'6001','2022','LO','Margin');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (27,'6001','2023','LO','Outline');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (28,'6001','2024','LO','Points');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (29,'6001','2025','LO','RotateTextWithImage');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (30,'6001','2026','LO','ScaleFontSize');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (31,'6001','2027','LO','ScaleWithCell');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (32,'6001','2028','LO','Shadow');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (33,'6001','2029','LO','ShowAnchor');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (34,'6001','2030','LO','ShowTextBox');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (35,'6001','2031','LO','Tag');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (36,'6001','2032','LO','TagObject');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (37,'6001','2033','LO','Text');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (38,'6001','2034','LO','Transparent');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (39,'6001','2035','LO','Visible');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (40,'6001','2036','LO','XOR');"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像标注存储表 (编号,VGroup,Element,VR,标注属性) values (41,'6001','2037','LO','PointsCount');"
    cnAccess.Execute strSQL
    
    '--影像图像信息表
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (1,'0008','0020','Study Date','检查日期','检查日期','Study Date',-1,-1,4,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (2,'0008','0022','Acquisition Date','拍片日期','拍片日期','Acq. Date',-1,0,1,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (3,'0008','0023','Image Date','图像日期','图像日期','Image Date',-1,0,2,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (4,'0008','0030','Study Time','检查时间','检查时间','Study Time',-1,-1,4,5,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (5,'0008','0032','Acquisition Time','拍片时间','拍片时间','Acq. Time',-1,0,1,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (6,'0008','0033','Image Time','图像时间','图像时间','Image Time',-1,0,1,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (7,'0008','0060','Modality','影像类别','影像类别','Modality',-1,-1,1,6,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (8,'0008','0070','Manufacturer','制造商','制造商','Manufacturer',-1,-1,2,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (9,'0008','0080','Institution Name','医院名称','医院','Institution',-1,-1,4,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (10,'0008','1010','Station Name','工作站名','工作站','Station',-1,0,4,2,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (11,'0008','1090','Manufacturer''s Model Name','设备名称','设备','Model Name',-1,0,1,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (12,'0010','0010','Patient''s Name','病人姓名','姓名','Name',-1,-1,4,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (13,'0010','0020','Patient ID','病人ID','病人ID','ID',-1,-1,4,6,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (14,'0010','0030','Patient''s Birth Date','病人生日','生日','Birth Date',-1,0,3,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (15,'0010','0040','Patient''s Sex','病人性别','性别','Sex',-1,-1,4,2,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (16,'0010','1010','Patient''s Age','病人年龄','年龄','Age',-1,-1,4,3,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (17,'0018','0015','Body Part Examined','检查部位','部位','Part',-1,-1,2,3,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (18,'0018','0050','Slice Thickness','层厚','层厚','Thickness',-1,-1,3,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (19,'0018','0060','KVP','电压','电压','KVP',-1,-1,2,6,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (20,'0018','0090','Data Collection Diameter','数据采集直径','采集直径','Collection Diameter',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (21,'0018','1100','econstruction Diameter','重建直径','重建直径','econstruction Diameter',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (22,'0018','1120','Gantry/Detector Tilt','设备倾角','倾角','Tilt',-1,-1,3,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (23,'0018','1130','Table Height','平台高度','平台高度','Table Height',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (24,'0018','1140','Rotation Direction','旋转方向','旋转方向','Rotation Direction',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (25,'0018','1150','Exposure Time','曝光时间','曝光时间','Exposure Time',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (26,'0018','1151','X-ray Tube Current','X线球管电流','电流','Current',-1,-1,2,5,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (27,'0018','1152','Exposure','曝光','曝光','Exposure',-1,-1,2,8,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (28,'0018','5100','Patient Position','病人位置','病人位置','Patient Position',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (29,'0020','000D','Study Instance UID','检查UID','检查UID','Study UID',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (30,'0020','000E','Series Instance UID','序列UID','序列UID','Series UID',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (31,'0020','0010','Study ID','检查ID','检查ID','Study ID',-1,-1,1,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (32,'0020','0011','Series Number','序列号','序列号','Series Number',-1,-1,1,2,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (33,'0020','0012','Acquisition Number','获取编号','获取编号','Acquisition Number',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (34,'0020','0013','Image Number','图像号','图像号','Image Number',-1,-1,1,3,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (35,'0020','0032','Image Position (Patient)','图像位置（病人）','图像位置','Image Position',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (36,'0020','0037','Image Orientation (Patient)','图像方向（病人）','图像方向','Image Orientation',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (37,'0020','1041','Slice Location','切片位置','切片位置','Slice Location',-1,-1,3,2,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (38,'0028','0002','Samples per Pixel','像素点的采样数','采样点数','Samples per Pixel',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (39,'0028','0004','Photometric Interpretation','光度解释','光度解释','Photometric',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (40,'0028','0010','Rows','行数','行数','Rows',-1,-1,1,4,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (41,'0028','0011','Columns','列数','列数','Columns',-1,-1,1,5,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (42,'0028','0030','Pixel Spacing','像素点间距','像素点距','Pixel Spacing',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (43,'0028','0100','Bits Allocated','图像位数','图像位数','Bits Allocated',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (44,'0028','0101','Bits Stored','图像存储的位数','图像位数','Bits Stored',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (45,'0028','0102','High Bit','图像存储的高位','图像高位','High Bit',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (46,'0028','0103','Pixel Representation','像素表述','像素表述','Pixel Representation',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (47,'0028','0106','Smallest Image Pixel Value','最小像素值','最小像素值','Smallest Pixel',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (48,'0028','0107','Largest Image Pixel Value','最大像素值','最大像素值','Largest Pixel',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (49,'0028','1050','Window Center','窗位','窗位','Center',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (50,'0028','1051','Window Width','窗宽','窗宽','Width',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (51,'0008','0016','SOP Class UID','SOP类UID','SOP UID','SOP UID',0,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (52,'0008','0018','SOP Instance UID','SOP实例UID','实例UID','Instance UID',-1,0,1,7,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (53,'0018','0080','Repetition Time','循环时间','TR','TR',-1,-1,2,1,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (54,'0018','0081','Echo Time','响应时间','TE','TE',-1,-1,2,2,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (55,'0018','0091','Echo Train Length','ET','ET','ET',-1,-1,2,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (56,'0008','103E','Series Description','序列描述','序列描述','Series',-1,-1,1,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (57,'0018','0088','Spacing Between Slices','层间距','层间距','Spacing Between Slices',-1,0,0,0,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (59,'1','1','cal','层厚层距','层厚层距','thk',-1,-1,2,7,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (60,'1','1','cal','视野FOV','视野FOV','FOV',-1,-1,3,3,0);"
    cnAccess.Execute strSQL
    strSQL = "insert into 影像图像信息表 (ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算) values (61,'2','2','cal','aaaa','aaaa','USER',-1,-1,1,7,0);"
    cnAccess.Execute strSQL
    
    '--影像界面参数表
    strSQL = "insert into 影像界面参数表 (ID,正常图像边框颜色,正常图像边框线型,正常图像边框线宽,选中图像边框颜色,选中序列边框颜色,选中图像边框线型,选中图像边框线宽,图像标记颜色,图像标记大小,标注选择句柄颜色,标注选择句柄大小,定位线颜色,定位线线型,定位线间距,序列间间隔,横向最大序列,纵向最大序列,图像间距,显示多余边框,背景颜色,程序背景颜色,标注正常颜色,标注正常线型,标注正常线宽,标注选中颜色,标注选中线型,标注选中线宽,标注文字大小,测量显示面积,测量显示平均值,测量显示均方差,测量显示中文,文字X方向偏移,文字Y方向偏移,文字随图像缩放,显示体位标记,中文体位标记,显示标尺,标尺左右边距,标尺上下边距,标尺宽度,标尺高度,标尺线宽,标尺颜色,窗宽窗位位置,鼠标穿梭步长,鼠标漫游步长,鼠标调窗步长,鼠标缩放步长,病人信息上下边距,病人信息左右边距,病人信息颜色,病人信息显示最小值,病人信息随图像缩放,病人信息字体,病人信息题头,直接照相,工具栏图标大小,工具栏位置,工具栏显示,状态栏字体大小,正常血管阈值,狭窄血管阈值,血管壁宽度,测量显示周长)" & _
             "VALUES (0,16777215,0,1,16777215,16777088,0,1,16777215,10,16777215,8,16777215,3,7,50,8,8,4,0,986895,131586,16777215,0,4,16777215,0,0,12,1,1,-1,-1,10,8,0,1010,1,1000,36,210,30,600,3,16777215,2,8,10,10,5,4,1,16777215,200,0,10,1,1,24,1,-1,9,51,50,10,-1);"
    cnAccess.Execute strSQL
    
    '--影像鼠标按钮分配
    strSQL = "INSERT INTO 影像鼠标按钮分配 (ID,直线,矩形,椭圆,箭头,多边形,多边线,角度,文字,穿梭定位,窗宽窗位,漫游,缩放,裁剪_标注调整,自适应调窗,三维鼠标,画标注)" & _
             "VALUES (0,'1,1,0,0,miLabelline,343','2,1,0,0,miLabelRectangle,344','3,1,0,0,miLabelEllipse,339','4,1,0,0,miLabelArrowhead,338','5,1,0,0,miLabelPolygon,342','6,1,0,0,miLabelPolyLine,341','7,1,0,0,miLabelAngle,340','8,1,0,0,miLabeltext,337','101,1,0,0,miStack,308','102,1,0,0,miWidthLevel,314','103,1,0,0,miCruise,309','104,1,0,0,miZoom,311','201,1,0,0,No,0','105,1,0,0,miAutoWidthLevel,315','106,1,0,0,mi3dCursor,321','20,1,0,0,no,0');"
    cnAccess.Execute strSQL
End Sub

Private Sub Class_Initialize()
    Set mfrmViewer = New frmViewer
End Sub

Private Sub Class_Terminate()
    Set mfrmViewer = Nothing
End Sub

Private Sub mfrmViewer_AfterSaveOuterImage(strCheckUID As String)
    RaiseEvent AfterSaveOuterImage(strCheckUID)
End Sub

Private Sub mfrmViewer_AfterSaveReportImage(strCheckUID As String)
    RaiseEvent AfterSaveReportImage(strCheckUID)
End Sub

Private Function funGetImageStrings(lngOrderID As Long, blnMoved As Boolean, strImageStrings As String, _
    intImageInterval As Integer) As Boolean
'------------------------------------------------
'功能：根据strImageStrings和intImageInterval，组织ZLSeriesInfos结构
'参数： lngOrderID -- 医嘱ID
'       blnMoved -- 数据是否被转储
'       strImageStrings -- 需要打开的图像串，规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象，空字符串表示打开这次检查的全部图像
'       intImageInterval -- 当选择打开全部图像，且图像数量大于100时，提取图像的时候的间隔
'返回：成功True，失败False
'时间：2009-7
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strImageSeries() As String
    Dim strImagesSeg() As String
    Dim iSeries As Integer
    Dim iSegment As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    
    Dim strSeriesUID As String
    Dim blnNewSeries As Boolean
    Dim iCurrentImage As Integer
    Dim iPrivImage As Integer
    Dim iCurrentSeries As Integer
    Dim oneSeriesInfo As clsSeriesInfo
    Dim oneImageInfo As clsImageInfo
    Dim strImageFilter As String
    Dim intCurrentInterval As Integer   '记录当前图像的间隔
    Dim blnGetImageWithInterval As Boolean  '当前序列是否根据间隔提取图像
    Dim strFilter As String
    
    Dim strSortKey As String
    
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errHandle
    
    If strImageStrings = "" Then
        strSQL = "Select 序列UID from 影像检查记录 a,影像检查序列 b where a.医嘱ID=[1] and a.检查UID=b.检查UID and ((b.序列描述<>'视频录像' and b.序列描述<>'音频数据') or b.序列描述 is null) order by b.序列号"
        
        
        '如果有转储标志，则读取转储的历史表
        If blnMoved Then
            strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取序列UID", lngOrderID)
        
        If rsTemp.EOF = True Then
            MsgBox "没有图像，请先选择图像后再打开观片站。", vbOKOnly, gstrSysName
            Exit Function
        End If
        
        rsTemp.MoveFirst
        While rsTemp.EOF = False
            If strImageStrings = "" Then
                strImageStrings = rsTemp("序列UID") & "|全部"
            Else
                strImageStrings = strImageStrings & "+" & rsTemp("序列UID") & "|全部"
            End If
            rsTemp.MoveNext
        Wend
    End If
    
    '组织图像分段检索的SQL语句

    '解析打开的图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
    strImageSeries = Split(strImageStrings, "+")    '拆分序列
    strFilter = ""
    For iSeries = 0 To UBound(strImageSeries)       '循环每一个序列
        '提取序列UID
        strSeriesUID = Split(strImageSeries(iSeries), "|")(0)
        If Split(strImageSeries(iSeries), "|")(1) = "全部" Then
            '打开本序列的全部图像
            strFilter = IIf(strFilter = "", "", strFilter & " OR ") & " A.序列UID='" & strSeriesUID & "'"
        Else
            '打开本序列的部分图像
            strImagesSeg = Split(Split(strImageSeries(iSeries), "|")(1), ";")
            strImageFilter = ""
            For iSegment = 0 To UBound(strImagesSeg)
                iStart = Split(strImagesSeg(iSegment), "-")(0)
                iEnd = Split(strImagesSeg(iSegment), "-")(1)
                
                If iStart = iEnd Then
                    strImageFilter = IIf(strImageFilter = "", "", strImageFilter & " OR ") & " A.图像号 = " & iStart
                Else
                    strImageFilter = IIf(strImageFilter = "", "", strImageFilter & " OR ") & " (A.图像号 >= " & iStart & " And A.图像号 <=" & iEnd & ")"
                End If
            Next iSegment
            strFilter = IIf(strFilter = "", "", strFilter & " OR ") & " (A.序列UID='" & strSeriesUID & "'" & " And (" & strImageFilter & "))"
        End If
    Next iSeries
    
    '查找需要打开的所有图象信息
    strSQL = "Select /*+RULE*/ A.动态图,A.图像号,C.检查UID,B.序列UID,B.序列号,D.FTP用户名 As FTPUser1,D.FTP密码 As FTPPwd1," & _
        "D.IP地址 As Host1,'/'||D.Ftp目录||'/' As FtpPath1,d.设备号 as 设备号1, C.影像类别, " & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'\')" & _
        "||C.检查UID As Path,A.图像UID, " & "E.FTP用户名 As FTPUser2,E.FTP密码 As FTPPwd2," & _
        "E.IP地址 As Host2,'/'||E.Ftp目录||'/' As FtpPath2," & " e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2,A.胶片打印, " & _
        "A.采集时间,A.图像时间,A.层厚,A.图像位置病人,A.图像方向病人,A.参考帧UID,A.切片位置,A.行数,A.列数,A.像素距离, " & _
        "f.姓名,f.性别,f.年龄,c.英文名,c.检查号,c.医嘱ID " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E, " & _
        " (Select x.姓名,x.性别,x.年龄,y.Id As 医嘱ID From 病人信息 x,病人医嘱记录 y Where y.id= [1] And x.病人ID = y.病人ID ) f " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "and f.医嘱ID = c.医嘱ID and ((B.序列描述<>'视频录像' and b.序列描述<>'音频数据') or B.序列描述 is null) And C.医嘱ID=[1] "
    strSQL = strSQL & "and (" & strFilter & ") Order By b.序列号,A.图像号 "

    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSQL = Replace(strSQL, "/*+RULE*/", "")
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取影像文件", lngOrderID)
    
    If rsTemp.EOF = True Then
        MsgBox "没有图像，请先选择图像后再打开观片站。", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '记录最新一次的检查UID
    PstrCheckUID = Nvl(rsTemp!检查UID)
    
    '初始化FTP参数并连接FTP
    If Nvl(rsTemp!设备号1) <> "" Then
        PstrFtpHost = Nvl(rsTemp!Host1)
        PstrFtpUser = Nvl(rsTemp!FTPUser1)
        PstrFtpPwd = Nvl(rsTemp!FTPPwd1)
        PstrFtpPath = Nvl(rsTemp!FtpPath1) & Nvl(rsTemp!Path)
    Else                                    '设备1中没有保存图像，改读设备2的信息
        PstrFtpHost = Nvl(rsTemp!Host2)
        PstrFtpUser = Nvl(rsTemp!FTPUser2)
        PstrFtpPwd = Nvl(rsTemp!FTPPwd2)
        PstrFtpPath = Nvl(rsTemp!FtpPath2) & Nvl(rsTemp!Path)
    End If
    
    '根据图像序列UID来装载图像
    '解析打开的图象号,规则是“序列UID1|1-3;5-27;33-100+序列UID2|全部”,全部表示打开全部图象
    For iSeries = 0 To UBound(strImageSeries)   '循环每一个序列
        
        '提取序列UID
        strSeriesUID = Split(strImageSeries(iSeries), "|")(0)
        
        rsTemp.Filter = "序列UID=" & strSeriesUID
        If rsTemp.EOF Then GoTo continue
        
        rsTemp.MoveFirst
        
        If Split(strImageSeries(iSeries), "|")(1) = "全部" And intImageInterval <> 0 And rsTemp.RecordCount > 100 Then   '根据间隔提取图像
            blnGetImageWithInterval = True
        Else    '提取全部图像
            blnGetImageWithInterval = False
        End If
        
        '查找该序列是否已经存在,是否需要创建新序列
        blnNewSeries = True
        For i = 1 To ZLSeriesInfos.Count
            If ZLSeriesInfos(i).SeriesUID = strSeriesUID Then
                blnNewSeries = False
                iCurrentSeries = i
                Exit For
            End If
        Next i
        '如果不存在，则增加该序列，存在则直接在该序列中查找图像是否存在
        '序列不排序,图像需要按照图像号排序
        If blnNewSeries = True Then
            Set oneSeriesInfo = funGetNewSeriesInfo
            oneSeriesInfo.StudyUID = rsTemp!检查UID
            oneSeriesInfo.SeriesUID = rsTemp!序列UID
            oneSeriesInfo.SeriesNo = rsTemp!序列号
            oneSeriesInfo.strModality = rsTemp!影像类别
            oneSeriesInfo.MultiFrame = Nvl(rsTemp!动态图, 0)
            oneSeriesInfo.strCName = Nvl(rsTemp!姓名)
            oneSeriesInfo.strEName = Nvl(rsTemp!英文名)
            oneSeriesInfo.strAge = Nvl(rsTemp!年龄)
            oneSeriesInfo.strSex = Nvl(rsTemp!性别)
            oneSeriesInfo.strStudyID = Nvl(rsTemp!检查号)
            oneSeriesInfo.strOrderID = Nvl(rsTemp!医嘱ID)
            If Nvl(rsTemp!设备号1) <> "" Then
                oneSeriesInfo.strHostIP = rsTemp!Host1
                oneSeriesInfo.strFTPDir = rsTemp!FtpPath1
                oneSeriesInfo.strFTPPasw = rsTemp!FTPPwd1
                oneSeriesInfo.strFTPUser = rsTemp!FTPUser1
                oneSeriesInfo.strShareDir = Nvl(rsTemp!共享目录1)
                oneSeriesInfo.strShareDirUser = Nvl(rsTemp!共享目录用户名1)
                oneSeriesInfo.strShareDirPasw = Nvl(rsTemp!共享目录密码1)
            ElseIf Nvl(rsTemp!设备号2) <> "" Then
                oneSeriesInfo.strHostIP = rsTemp!Host2
                oneSeriesInfo.strFTPDir = rsTemp!FtpPath2
                oneSeriesInfo.strFTPPasw = rsTemp!FTPPwd2
                oneSeriesInfo.strFTPUser = rsTemp!FTPUser2
                oneSeriesInfo.strShareDir = Nvl(rsTemp!共享目录2)
                oneSeriesInfo.strShareDirUser = Nvl(rsTemp!共享目录用户名2)
                oneSeriesInfo.strShareDirPasw = Nvl(rsTemp!共享目录密码2)
            End If
            oneSeriesInfo.strSaveDir = rsTemp!Path
            '读取预设窗宽窗位
            For i = 1 To UBound(aPresetWinWL, 2)
                If UCase(aPresetWinWL(3, i).strModality) = UCase(oneSeriesInfo.strModality) Then
                    For j = 3 To 12
                        If aPresetWinWL(j, i).bInUse And aPresetWinWL(j, i).intDefault = 1 Then
                            oneSeriesInfo.lngWinWidth = aPresetWinWL(j, i).lngWinWidth
                            oneSeriesInfo.lngWinLevel = aPresetWinWL(j, i).lngWinLevel
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
            
            ZLSeriesInfos.Add oneSeriesInfo, CStr(ZLSeriesInfos.Count + 1)
            iCurrentSeries = ZLSeriesInfos.Count
            
            '添加图像
            intCurrentInterval = 0
            While rsTemp.EOF = False
                '提取一个图像
                Set oneImageInfo = funGetNewImageInfo
                oneImageInfo.InstanceUID = rsTemp!图像UID
                oneImageInfo.ImageNo = rsTemp!图像号
                oneImageInfo.ImageName = oneImageInfo.InstanceUID
                oneImageInfo.AcquisitionTime = Format(Nvl(rsTemp!采集时间), "yyyy-MM-dd HH:MM:SS")
                oneImageInfo.ImageTime = Format(Nvl(rsTemp!图像时间), "yyyy-MM-dd HH:MM:SS")
                oneImageInfo.SliceThickness = Nvl(rsTemp!层厚)
                oneImageInfo.ImagePositionPatient = Nvl(rsTemp!图像位置病人)
                oneImageInfo.ImageOrientationPatient = Nvl(rsTemp!图像方向病人)
                oneImageInfo.FrameOfReferenceUID = Nvl(rsTemp!参考帧UID)
                oneImageInfo.SliceLocation = Nvl(rsTemp!切片位置)
                oneImageInfo.Rows = Nvl(rsTemp!行数)
                oneImageInfo.Columns = Nvl(rsTemp!列数)
                oneImageInfo.PixelSpacing = Nvl(rsTemp!像素距离)
                oneImageInfo.blnPrinted = IIf(Nvl(rsTemp!胶片打印, 0) = 1, True, False)
                oneImageInfo.ImgTag = Nvl(rsTemp!动态图, 0)
                
                ZLSeriesInfos(iCurrentSeries).ImageInfos.Add oneImageInfo
                
                '如果按照间隔提取图像，按照间隔向后移动数据集
                If blnGetImageWithInterval Then '按间隔提取图像
                    For i = 1 To intImageInterval
                        If rsTemp.EOF = False Then rsTemp.MoveNext
                    Next i
                Else
                    rsTemp.MoveNext
                End If
            Wend
        Else
            '序列已经存在，则当前增加进来的图像就属于追加的
            intCurrentInterval = 0
            
            While rsTemp.EOF = False
                '先循环当前图像，查找图像是否存在
                iCurrentImage = 0
                iPrivImage = 0
                strSortKey = Nvl(rsTemp!图像号)
                
                For i = 1 To ZLSeriesInfos(iCurrentSeries).ImageInfos.Count
                    If ZLSeriesInfos(iCurrentSeries).ImageInfos(i).ImageNo < Val(strSortKey) Then
                        iPrivImage = i
                    End If
                    
                    If ZLSeriesInfos(iCurrentSeries).ImageInfos(i).InstanceUID = rsTemp!图像UID Then
                        iCurrentImage = i
                        Exit For
                    End If
                Next i
                
                '图像不存在才增加,增加的位置是iPrivImage之后
                If iCurrentImage = 0 Then
                    Set oneImageInfo = funGetNewImageInfo
                    oneImageInfo.InstanceUID = rsTemp!图像UID
                    oneImageInfo.ImageNo = rsTemp!图像号
                    oneImageInfo.ImageName = oneImageInfo.InstanceUID
                    oneImageInfo.AcquisitionTime = Format(Nvl(rsTemp!采集时间), "yyyy-MM-dd HH:MM:SS")
                    oneImageInfo.ImageTime = Format(Nvl(rsTemp!图像时间), "yyyy-MM-dd HH:MM:SS")
                    oneImageInfo.SliceThickness = Nvl(rsTemp!层厚)
                    oneImageInfo.ImagePositionPatient = Nvl(rsTemp!图像位置病人)
                    oneImageInfo.ImageOrientationPatient = Nvl(rsTemp!图像方向病人)
                    oneImageInfo.FrameOfReferenceUID = Nvl(rsTemp!参考帧UID)
                    oneImageInfo.SliceLocation = Nvl(rsTemp!切片位置)
                    oneImageInfo.Rows = Nvl(rsTemp!行数)
                    oneImageInfo.Columns = Nvl(rsTemp!列数)
                    oneImageInfo.PixelSpacing = Nvl(rsTemp!像素距离)
                    oneImageInfo.blnPrinted = IIf(Nvl(rsTemp!胶片打印, 0) = 1, True, False)
                    oneImageInfo.ImgTag = Nvl(rsTemp!动态图, 0)
                
                    If iPrivImage = 0 Then
                        '在第一个图的前面追加
                        ZLSeriesInfos(iCurrentSeries).ImageInfos.Add oneImageInfo, , 1
                    Else
                        ZLSeriesInfos(iCurrentSeries).ImageInfos.Add oneImageInfo, , , iPrivImage
                    End If
                End If
                
                '如果按照间隔提取图像，按照间隔向后移动数据集
                If blnGetImageWithInterval Then '按间隔提取图像
                    For i = 1 To intImageInterval
                        If rsTemp.EOF = False Then rsTemp.MoveNext
                    Next i
                Else
                    rsTemp.MoveNext
                End If
            Wend
        End If
continue:
    Next iSeries
    
    funGetImageStrings = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

