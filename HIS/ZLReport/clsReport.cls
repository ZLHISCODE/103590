VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private arrForms() As Variant
Private mblnDataIsEmpty As Boolean
Private mrsFormatInfo As ADODB.Recordset

'事件定义
Public Event CommitCondition(ByVal strNO As String, strCond As String, frmParent As Object) '报表条件确定

Public Event ResetCondition(ByVal strNO As String, strCond As String, blnInhere As Boolean, frmParent As Object) '报表条件重置

Public Event EnterCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String) '进入某个表格的一个单元
Public Event LevelCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String) '离开某个表格的一个单元
Public Event EnterRow(ByVal strNO As String, Row As Integer, RowText As String, Sheet As Object) '进入某个表格的一行
Public Event EnterCol(ByVal strNO As String, Col As Integer, Sheet As Object) '进入某个表格的一列
Public Event ColResize(ByVal strNO As String, Col As Integer, NewWidth As Single, OldWidth As Single) '改变某个表格的列宽

Public Event BeforeReportLoad(ByVal strNO As String, Form As Object) '开始读取报表
Public Event AfterReportLoad(ByVal strNO As String, Form As Object) '读取报表结束
Public Event ReportActive(ByVal strNO As String, Form As Object) '报表窗体激活

Public Event BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant) '报表打印之前
Public Event PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal id As Long) '打印任意表格的行时
Public Event AfterPrint(ByVal ReportNum As String) '报表打印之后

Public Event PaperMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, blnPop As Boolean) '报表纸张上鼠标按下
Public Event PaperMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single) '报表纸张上鼠标移动
Public Event PaperMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single) '报表纸张上鼠标松开

Public Event SheetDblClick(ByVal strNO As String, Sheet As Object, frmParent As Object)

Public Event SheetMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '报表表格上鼠标按下
Public Event SheetMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '报表表格上鼠标移动
Public Event SheetMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object) '报表表格上鼠标松开

Public Event PaperScroll(ByVal strNO As String, Way As Byte, Percent As Single) '报表横纵向滚动
Public Event SheetScroll(ByVal strNO As String, Sheet As Object) '某个表格内部滚动

Public Event ReportUnload(ByVal strNO As String, frmReport As Object) '某个表格内部滚动
Private mobjfrmReport As frmReport

Public Property Get DataIsEmpty() As Boolean
'属性功能：直接调用报表打印时，判断报表数据是否为空。
'注：该属性只在直接调用报表打印时有用。
    DataIsEmpty = mblnDataIsEmpty
End Property

Friend Property Let DataIsEmpty(ByVal blnDataIsEmpty As Boolean)
'说明：想做成对外只读属性，实际测试好象无效。
    mblnDataIsEmpty = blnDataIsEmpty
End Property

'==事件激活过程(外部隐藏)=====================================================================================
Friend Sub Act_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
'功能：报表直接输出到打印机时，在页数计算完毕打印之前激活
'参数：ReportNum=报表编号
'      TotalPages=全部打印时的总页数，-1表示只打印当前页，-2表示打印指定范围的页
'      Cancel=是否取消打印
'      arrBill=实际要打印的票据号数组“Split("A,B,C",",")”，缺省为非数组/Empty
    RaiseEvent BeforePrint(ReportNum, TotalPages, Cancel, arrBill)
End Sub

Friend Sub Act_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal id As Long)
'功能：打印任意表格的行时激活
'参数：Page=页号(1-n),Row=行号(1-n),ID=数据源中的ID标识(不一定绑定到了表格上)
'说明：如果该行单元格一个都没有数据，则认为不打印，该行不激活本事件
    RaiseEvent PrintSheetRow(ReportNum, Sheet, Page, Row, id)
End Sub

Friend Sub Act_AfterPrint(ByVal ReportNum As String)
'功能：报表直接输出到打印机之后激活
'参数：ReportNum=报表编号
    RaiseEvent AfterPrint(ReportNum)
End Sub

Friend Sub Act_ColResize(ByVal strNO As String, Col As Integer, NewWidth As Single, OldWidth As Single)
'功能：改变某个表格的列宽时激活
'参数：可以设置新的改变后列宽度
    RaiseEvent ColResize(strNO, Col, NewWidth, OldWidth)
End Sub

Friend Sub Act_SheetScroll(ByVal strNO As String, Sheet As Object)
'功能：某个表格内部滚动时激活
'参数：可以设置Sheet的内容
    RaiseEvent SheetScroll(strNO, Sheet)
End Sub

Friend Sub Act_PaperScroll(ByVal strNO As String, Way As Byte, Percent As Single)
'功能：报表横纵向滚动时激活
'参数：Way=0:横向，1:纵向,Percent:当前滚动位置百分比,可以更改
    RaiseEvent PaperScroll(strNO, Way, Percent)
End Sub

Friend Sub Act_PaperMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能：报表纸张上鼠标松开时激活
    RaiseEvent PaperMouseUp(strNO, Button, Shift, X, Y)
End Sub

Friend Sub Act_PaperMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能：报表纸张上鼠标移动时激活
    RaiseEvent PaperMouseMove(strNO, Button, Shift, X, Y)
End Sub

Friend Sub Act_PaperMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, blnPop As Boolean)
'功能：报表纸张上鼠标按下时激活
    RaiseEvent PaperMouseDown(strNO, Button, Shift, X, Y, blnPop)
End Sub

Friend Sub Act_SheetDblClick(ByVal strNO As String, Sheet As Object, frmParent As Object)
'功能：报表表格上鼠标按下时激活
    RaiseEvent SheetDblClick(strNO, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseUp(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'功能：报表表格上鼠标松开时激活
    RaiseEvent SheetMouseUp(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseMove(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'功能：报表表格上鼠标移动时激活
    RaiseEvent SheetMouseMove(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_SheetMouseDown(ByVal strNO As String, Button As Integer, Shift As Integer, X As Single, Y As Single, Sheet As Object, frmParent As Object)
'功能：报表表格上鼠标按下时激活
    RaiseEvent SheetMouseDown(strNO, Button, Shift, X, Y, Sheet, frmParent)
End Sub

Friend Sub Act_ReportActive(ByVal strNO As String, Form As Object)
'功能：在报表窗体激活时激活
'参数：可以改变Form值设置报表
    RaiseEvent ReportActive(strNO, Form)
End Sub

Friend Sub Act_AfterReportLoad(ByVal strNO As String, Form As Object)
'功能：读取报表结束时激活
'参数：可以改变Form值设置报表
    RaiseEvent AfterReportLoad(strNO, Form)
End Sub

Friend Sub Act_BeforeReportLoad(ByVal strNO As String, Form As Object)
'功能：开始读取报表时激活
'参数：可以改变Form值设置报表
    RaiseEvent BeforeReportLoad(strNO, Form)
End Sub

Friend Sub Act_EnterCol(ByVal strNO As String, Col As Integer, Sheet As Object)
'功能：进入某个表格列时激活
'参数：可以改变Sheet值设置表格
    RaiseEvent EnterCol(strNO, Col, Sheet)
End Sub

Friend Sub Act_EnterRow(ByVal strNO As String, Row As Integer, RowText As String, Sheet As Object)
'功能：进入某个表格行时激活,RowText="内容|内容|....."
'参数：可以改变Sheet值设置表格
    RaiseEvent EnterRow(strNO, Row, RowText, Sheet)
End Sub

Friend Sub Act_LevelCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String)
'功能：离开某个表格单元格时激活
    RaiseEvent LevelCell(strNO, Row, Col, Text)
End Sub

Friend Sub Act_EnterCell(ByVal strNO As String, Row As Integer, Col As Integer, Text As String)
'功能：进入某个表格单元格时激活
'参数：可以改变Row,Col,Text值设置表格
    RaiseEvent EnterCell(strNO, Row, Col, Text)
End Sub

Friend Sub Act_ReportUnload(ByVal strNO As String, Form As Object)
'功能：读取报表结束时激活
'参数：可以改变Form值设置报表
    RaiseEvent ReportUnload(strNO, Form)
End Sub

Friend Sub Act_CommitCondition(ByVal strNO As String, strCond As String, frmParent As Object)
'功能：条件确定后查询前激活
'参数：
'   strNO=当前打开的报表编号
'   strCond=用于DLL与调用程序条件的互相专递,格式为"参数名=参数值|参数名=参数值...."
'           如果报表具有多个格式,则包含"ReportFormat=x",返回时也应该处理,不然就保持当前格式不变
'   frmParent=父窗体(报表窗体)
    RaiseEvent CommitCondition(strNO, strCond, frmParent)
End Sub

Friend Sub Act_ResetCondition(ByVal strNO As String, strCond As String, blnInhere As Boolean, frmParent As Object)
'功能：重置条件时激活
'参数：
'   strNO=当前打开的报表编号
'   strCond=用于DLL与调用程序条件的互相专递,格式为"参数名=参数值|参数名=参数值...."
'           如果报表具有多个格式,则包含"ReportFormat=x",返回时也应该处理,不然就保持当前格式不变
'   blnInHere=是否调用DLL中原有的条件输入窗体
'   frmParent=父窗体(报表窗体)
    RaiseEvent ResetCondition(strNO, strCond, blnInhere, frmParent)
End Sub
'===================================================================================================

Public Function ReportMan(ByVal cnMain As ADODB.Connection, Optional frmParent As Object _
    , Optional strUser As String, Optional ByRef frmDBConnect As Object) As Boolean
'功能：自定义报表管理
'参数：
'     cnMain:主程序的数据库链接
'     frmParent:主窗体
'     strUser:管理工具进入的登录用户名
'     frmDBConnect：数据连接管理
'返回：调用是否成功

    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Screen.MousePointer = vbHourglass
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    gstrLoginUser = strUser
    gstrLoginUserName = GetTheUserName(strUser)
    gstrComputerName = ComputerName
    Set gfrmMain = frmParent
        
    Call InitPar
    gblnAutoConnect = GetAutoConnect()
    
    '获取用户权限对象(只处理一次)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    '初始化数据连接的记录集对象
    gblnManagementTool = True
    Call mdlPublic.SetControlDBConnect(grsConnect)
    
    '数据连接管理窗体
    Set gfrmDBConnect = frmDBConnect
    
    If Not frmParent Is Nothing Then
        Load frmMain
        Err.Clear
        On Error Resume Next
        frmMain.Show , frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmMain.Show 1
        End If
    Else
        frmMain.Show
    End If
    
    ReportMan = True
    Screen.MousePointer = vbDefault
End Function

Public Function ReportImport(ByVal strFile As String, ByVal cnMain As ADODB.Connection, _
    Optional ByVal strNum As String, Optional ByVal blnOnlyData As Boolean) As Boolean
'功能：导入一个报表文件
'参数：
'      strFile:要导入的报表文件名
'      cnMain:调用程序的数据库连接
'      strNum:覆盖固定报表时用,要覆盖的报表编号
'      blnOnlyData:覆盖固定报表时用,是否只导入数据源
'返回：导入是否成功
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then Exit Function
    
    If strNum <> "" Then
        strSQL = "Select ID From zlReports Where 编号=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "ReportImport", strNum)
        If rsTmp.EOF Then Exit Function
        ReportImport = ImportReport(strFile, rsTmp!id, blnOnlyData) <> ""
    Else
        ReportImport = ImportReport(strFile) <> ""
    End If
End Function

Public Function ReportOpen(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, ParamArray arrPar() As Variant) As Boolean
'功能：按报表发布的程序号或报表编号打开报表。
'参数：
'     lngSys=报表所属的系统,当为共享系统报表时,传入为0
'     varReport:需要执行的报表编号或对应的程序ID(当报表为独立菜单程序时才使用程序ID)
'     cnMain:主程序的数据库连接
'     frmParent:主窗体
'     arrPar():1.报表参数,格式为"参数名1=参数值1"、"参数数据源1"...,用是否有等号区分是报表参数还是附加报表数据。
'              2.对于选择器或下接框类型的参数,如果要想直接指定显示值/绑定值,可以以"参数名=显示值|绑定值"的格式传入
'                也可只传入"参数名=绑定值"的方式，程序会自动尝试去寻找对应的显示值。
'              3.对于有多个格式的报表,可以在参数中传入"ReportFormat=x",以确定要打开的报表格式,对报表组无效
'              4.在直接打印时,可以在参数中传入"PrintEmpty=1/0",表明当前格式中的所有表格无数据时,是否可以打印
'              5.如果明确打印或预览时需要进行套打，则在参数中传入"PressWork=1"，或"PressWorkFirst=1".
'                表示全部或首页套打，这样就只会对表格中的数据进行输出，对表头、表格线及其他元素不作输出。
'              6.如果打印页号时要确定起始计算基准页，可传入"StartPageNum=x"(x>=1)参数，[页号]标签会在此基础上计算。
'              7.最后附加参数:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF
'              8.输出到Excel时，可以通过"ExcelFile=***.xls"参数来指定输出文件，指定了输出文件时不会自动打开
'              9.当打开的是报表组时,可以将所有报表的参数依次传入
'返回：调用是否成功
'说明：1.当使用程序号打开报表时,报表必须已经发布。当打开未发布的报表时,必须手工在数据库上增加权限。
'      2.参数其个数由报表定义时决定,每个形式为"参数名=参数值",如:"部门名称='内一科','妇产科'","开始时间=2000-01-02","姓名=张三","部门ID=23"。
'      3.这种方式允许主程序自已设计参数输入窗体,以加强使用者输入参数的灵活性及适用性(如输入操作员姓名时,可以加一个人员选择器)。
'      4.因为报表有多个数据源,每个数据源中有多个参数,这些参数名称唯一,序号却是乱的,所以传参数时要以名称为键。
'      5.参数如没有传或没有传完(包括报表格式参数),则系统会自动弹出参数输入窗体对未传参数进行输入。
    
    Dim bytStyle As Byte
    Dim blnCache As Boolean, blnDo As Boolean
    Dim intTmp As Integer, intState As Integer
    Dim strInfo As String, strGroup As String
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then Exit Function
    
    bytStyle = 0
    If UBound(arrPar) >= 0 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    gblnSilentMode = GetSilentMode(arrPar) And (bytStyle = Val("2-打印") Or bytStyle = Val("4-PDF"))
    
    Set gfrmMain = frmParent
    glngSys = lngSys
    
    mblnDataIsEmpty = False
    
    Call InitPar
    
    '参数及显示方式
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
'        If IsNumeric(arrPar(UBound(arrPar))) Then
'            bytStyle = CByte(arrPar(UBound(arrPar)))
'        End If
    End If
    If bytStyle = 3 Then
        If zlRegInfo("授权性质") <> "1" Then
            MsgBox "试用或测试版本不能使用该功能。", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '初始化数据连接的记录集对象
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup, strInfo) Then
        '打开报表组
        If glngGroup = 0 Then Exit Function
        
        '报表状态
        strGroup = "组"
        GoSub proState
        If blnDo = False Then Exit Function
        
        '执行报表组
        Set gobjReport = Nothing
        ReportOpen = ShowReport(frmParent, Me, bytStyle)
    Else
        '合法性检查
        If Not CheckPass(glngGroup) Then
            MsgBox "报表:[" & grsReport!编号 & "]" & grsReport!名称 & " 数据错误，不能执行该报表！", vbInformation, App.Title
            Exit Function
        End If
        '权限判断
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then '仅报表才提示,票据不提示
                MsgBox "你没有权限执行报表:[" & grsReport!编号 & "]" & grsReport!名称 & "，请确定系统是否正确授权！", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        '报表状态
        strGroup = ""
        GoSub proState
        If blnDo = False Then Exit Function
        
        '读取报表
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.编号 = grsReport!编号 And gobjReport.系统 = Nvl(grsReport!系统, 0) _
                And Format(gobjReport.修改时间, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!修改时间, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '打开报表
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        ReportOpen = ShowReport(frmParent, Me, bytStyle)
    End If
    
    ReportOpen = ReportOpen And Not gblnError
    Exit Function
    
proState:
    intState = mdlPublic.ReportStateSwitch(lngSys, varReport, strGroup <> "")
    
    blnDo = False
    Select Case intState
    Case Val("0-报表不存在或未发布")
        If strInfo = "" Then
            MsgBox mdlPublic.FormatString("报表[1]不存在，请联系管理员！", strGroup), vbInformation, App.Title
        Else
'            MsgBox mdlPublic.FormatString("“[1]”报表[2]未发布，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
            blnDo = True
        End If
    Case Val("1-报表启用中")
        blnDo = True
    Case Val("2-报表停用中")
        MsgBox mdlPublic.FormatString("“[1]”报表[2]停用中，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
    End Select
    Return
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ReportTaskBegin()
'功能：开始打印，用于多报表在单一任务中打印
'说明：需要最先调用该函数，再用ReportOpen，且调用ReportOpen时一定要以打印方式
    If Not gblnSingleTask Then
        gblnSingleTask = True
    End If
End Sub

Public Sub ReportTaskEnd()
'功能：结束打印，用于多报表在单一任务中打印
'说明：调用完ReportOpen之后，最后才调用该函数
    If gblnSingleTask Then
        Printer.EndDoc
        gblnSingleTask = False
    End If
End Sub

Public Function ReportPrintSet(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, Optional frmParent As Object) As Boolean
'功能：设置本地报表打印机(打印时打印机顺序：本地-服务器-当前缺省)
'参数：
'     varReport:需要设置的报表功能序号(程序ID)或报表编号
'     cnMain:主程序的数据库连接
'     frmParent:主窗体
'返回：调用是否成功
    Dim lngGroup As Long, intTmp As Integer
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    If InitRegister = False Then Exit Function
    
    If Not isGroup(lngSys, varReport, lngGroup) Then
        '权限判断
        intTmp = ReportCanQuery(lngGroup)
        If intTmp <> 0 Then
            If intTmp <> 3 Then
                MsgBox "你没有足够的权限执行该项操作，请确定是否正确授权！", vbInformation, App.Title
            End If
            Exit Function
        End If
    End If
    
    ReportPrintSet = ReportLocalSet(lngSys, varReport, True, , frmParent)
    If Not frmParent Is Nothing Then frmParent.Refresh
End Function

Public Function ReportDesign(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal varReport As Variant, frmParent As Object, Optional blnModiData As Boolean) As Boolean
'功能：在指定对象中显示报表内
'参数：lngSys=系统编号,共享报表传0
'      varReport:报表编号
'      frmParent=调用父窗口
'      blnModiData=是否允许修改数据源(缺省为否)
'返回：报表是否被修改
    Dim rsTmp As New ADODB.Recordset
    Dim frmNew As frmDesign
    Dim strSQL As String, i As Integer
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    glngSys = lngSys
    
    If InitRegister = False Then Exit Function
    
    Call InitPar
        
    '获取用户权限对象(只处理一次)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then
        Set grsObject = Nothing
        Set grsObject = UserObject
        If grsObject Is Nothing Then Exit Function
    End If
        
    '读取报表信息
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where Nvl(系统,0)=[3] And 编号=[1]"
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where Nvl(系统,0)=[3] And 程序ID=[2]"
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = OpenSQLRecord(strSQL, "ReportDesign", UCase(varReport), Val(varReport), lngSys)
    If rsTmp.EOF Then Exit Function
    
    '--
    Set frmNew = New frmDesign
    frmNew.lngRPTID = rsTmp!id
    frmNew.mblnNotModiData = Not blnModiData
    
    On Error Resume Next
    frmNew.Show 1, frmParent
    On Error GoTo 0
    
    ReportDesign = gblnModi
End Function

Public Sub ReportShow(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal varReport As Variant, objParent As Object)
'功能：在指定对象中显示报表内容(目前没有使用)
'参数：lngSys=系统编号,共享报表传0
'      varReport:报表编号
'      objParent=要显示报表的容器(一般为PictureBox或Form)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    If cnMain Is Nothing Then Exit Sub
    If cnMain.State = adStateClosed Then Exit Sub
    If Not ChartInstall Then Exit Sub
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then Exit Sub
        
    '读取报表信息
    If TypeName(varReport) = "String" Then
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where Nvl(系统,0)=[3] And 编号=[1]"
    Else
        strSQL = "Select ID,编号,名称,说明,密码,打印机,进纸,票据,打印方式,系统,程序ID,功能,修改时间,发布时间,禁止开始时间,禁止结束时间 From zlReports Where Nvl(系统,0)=[3] And 程序ID=[2]"
    End If
    'Set rsTmp = New ADODB.Recordset
    Set rsTmp = OpenSQLRecord(strSQL, "ReportDesign", UCase(varReport), Val(varReport), lngSys)
    If rsTmp.EOF Then Exit Sub
    
    '确定一个未使用的位置
    For i = 0 To UBound(arrForms)
        If arrForms(i) Is Nothing Then Exit For
    Next
    If i > UBound(arrForms) Then
        ReDim Preserve arrForms(UBound(arrForms) + 1)
        i = UBound(arrForms)
    End If
    
    '显示报表
    Set arrForms(i) = New frmShowDesign
    arrForms(i).mlngRPTID = rsTmp!id
    Load arrForms(i)
    
    SetParent arrForms(i).hwnd, objParent.hwnd
    
    arrForms(i).Tag = varReport
    
    '必须用API定位
    MoveWindow arrForms(i).hwnd, 0, 0, arrForms(i).ScaleX(objParent.ScaleWidth, vbTwips, vbPixels), arrForms(i).ScaleY(objParent.ScaleHeight, vbTwips, vbPixels), 0
    
    arrForms(i).Show
End Sub

Public Sub ReportWindowResize(ByVal varReport As Variant, ByVal W As Long, ByVal H As Long)
'功能：当容器尺寸发生变化时,同步更改报表显示的尺寸
'参数：varReport:报表编号
'      W,H=新的高度,宽度(vbTwips Units)
    Dim i As Integer
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            If arrForms(i).Tag = varReport Then
                MoveWindow arrForms(i).hwnd, 0, 0, arrForms(i).ScaleX(W, vbTwips, vbPixels), arrForms(i).ScaleY(H, vbTwips, vbPixels), 0
            End If
        End If
    Next
End Sub

Public Sub ReportWindowUnload(ByVal varReport As Variant)
'功能：卸载指定的显示报表的窗体
'参数：varReport:报表编号
    Dim i As Integer
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            If arrForms(i).Tag = varReport Then
                Unload arrForms(i)
                Set arrForms(i) = Nothing
            End If
        End If
    Next
End Sub

Public Function ReportNew(cnMain As ADODB.Connection, lngSys As Long, strNum As String, strName As String, strSQL As String, ParamArray arrPar() As Variant) As Boolean
'功能：新增一张固定未发布报表
'参数：lngSys=系统编号
'      strNum=报表编号
'      strName=报表名称
'      strSQL=数据源,其中可能包含参数的书写
'      arrPar=如果数据源中包含参数,则传具体的描述。格式为"参数名;类型;缺省值",其中缺省值可为空,顺序依据数据源中定义的参数序号顺序。
'返回：如果数据源不能正确分析,则会返回失败。
'说明：1.缺省纸张为A4,纵向,打印机在打印前调用相关函数设置。
'      2.报表缺省的标题为报表名称
'      3.根据参数传入顺序,将各个参数作为表上项目组织
'      4.根据数据源内容组织成清册表输出
    Dim blnTran As Boolean, lngNewID As Long
    Dim objReport As Report, objPars As RPTPars
    Dim strObject As String, strField As String
    Dim i As Integer, j As Integer, k As Integer
    Dim strTmp As String, strR As String
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    
    If InitRegister = False Then Exit Function
    
    '获取用户权限对象(只处理一次)
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    
    '生成报表内容
    Set objReport = New Report
    objReport.编号 = strNum
    objReport.名称 = strName
    objReport.进纸 = 15
    
    '数据源
    '参数
    Set objPars = New RPTPars
    If UBound(arrPar) <> -1 Then
        For i = 0 To UBound(arrPar)
            strTmp = CStr(arrPar(i))
            objPars.Add "", CByte(i), CStr(Split(strTmp, ";")(0)), CByte(Split(strTmp, ";")(1)), CStr(Split(strTmp, ";")(2)), 0, "", "", "", "", "", "", "_" & i
        Next
    End If
    
    '分析对象
    strObject = SQLObject(strSQL)
    If strObject = "" Then Exit Function
    strObject = ObjectOwner(strObject)
    If strObject = "取消" Then Exit Function '取消操作
    
    '分析字段
    strTmp = SQLOwner(RemoveNote(strSQL), strObject)
    If GetParCount(strTmp) = 0 Then
        strField = CheckSQL(strTmp, strR)
    Else
        strField = CheckSQL(strTmp, strR, ReplaceParSysNo(objPars, lngSys))
    End If
    If strField = "" Then Exit Function
    
    objReport.Datas.Add "报表数据", 0, strSQL, strField, strObject, 0, "", objPars, "_报表数据"
    
    '格式
    objReport.Fmts.Add 1, strName, INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    
    '元素
    '清册表格
    j = 1
    objReport.Items.Add j, 1, "清册表", 0, 4, 0, "", 0, "报表数据", "", Twip_mm * 10, Twip_mm * 30 _
        , INIT_WIDTH - Twip_mm * 20, INIT_HEIGHT - Twip_mm * 40, 285, 0, False, "宋体", 9, False _
        , False, False, 0, 0, vbWhite, False, 1, "", "", "", False, False, , , , , , "_" & j
    For i = 0 To UBound(Split(strField, "|"))
        j = j + 1
        Select Case CLng(Split(Split(strField, "|")(i), ",")(1))
            Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger _
                , adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                k = 2
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                k = 0
            Case adDBTimeStamp, adDBTime, adDBDate, adDate
                k = 0
            Case Else
                k = 0
        End Select
        objReport.Items.Add j, 1, "", 1, 6, i, "", 0, "[报表数据." & Split(Split(strField, "|")(i), ",")(0) & "]" _
            , "4^285^" & Split(Split(strField, "|")(i), ",")(0), 0, 0, 1000, 0, 0, CByte(k), False, "", 0, False _
            , False, False, 0, 0, 0, False, 0, "", "", "", False, False, , , , , , "_" & j
        objReport.Items("_1").SubIDs.Add j, "_" & j
    Next
    '标题
    j = j + 1
    objReport.Items.Add j, 1, "标题", 0, 2, 0, "清册表", 12, "[单位名称]" & strName, "", Twip_mm * 10 _
        , Twip_mm * 10, Twip_mm * 10, Twip_mm * 10, 0, 1, True, "宋体", 20, False, False, False, 0, 0 _
        , vbWhite, False, 0, 0, 0, 0, False, False, , True, , , , "_" & j
    
    lngNewID = GetNextID("zlReports")
    
    '插入报表
    blnTran = True
    gcnOracle.BeginTrans
    gcnOracle.Execute "Insert Into zlReports(ID,编号,名称,系统,密码) Values(" & lngNewID & ",'" & strNum & "','" & strName & "'," & lngSys & "," & AdjustStr(GetPass(strNum, strName)) & ")"
    gcnOracle.CommitTrans
    blnTran = False
    
    '保存存报表内容
    ReportNew = SaveReport(lngNewID, objReport)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Public Function ReportDel(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, ByVal strNum As String) As Boolean
'功能：删除一张报表
'参数：cnMain=数据库链接,strNum=报表编号
    Dim strSQL As String
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
        
    strSQL = "Delete From zlReports Where Upper(编号)='" & UCase(strNum) & "' And Nvl(系统,0)=" & lngSys
    cnMain.BeginTrans
    cnMain.Execute strSQL
    cnMain.CommitTrans
    
    ReportDel = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    cnMain.RollbackTrans
    Call SaveErrLog
End Function

Public Sub ReportAbout(Optional frmParent As Object)
'功能：显示自定义报表关于窗口
'说明：调用该接口前一定注意先调用InitOracle
    Call ShowAbout(frmParent)
End Sub

Public Function GenReportPass(ByVal strCode As String, ByVal strName As String) As String
'功能：根据报表的编号及名称返回合法的密码
'说明：调用该接口前一定注意先调用InitOracle
    GenReportPass = GetPass(strCode, strName)
End Function

Public Sub InitOracle(ByVal cnMain As ADODB.Connection)
'功能：对自定义报表的数据库连接进行初始化
    Set gcnOracle = cnMain
    If InitRegister = False Then Exit Sub
End Sub

Private Sub Class_Initialize()
    mblnDataIsEmpty = False
    
    gstrFind = MakeFile("Find", "AVI")
    gstrErrorContent = ""
    Call RegReportFile
    
    Call InitEnv    '初始化加密串
    Call GetCompare '还原加密串
    
    '初始化报表显示窗体
    arrForms = Array()
    
    Set gclsCNs = New RPTDBCNs
    Set grsConnect = New ADODB.Recordset
    Set gfrmDBConnect = Nothing
    gblnManagementTool = False
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    
    If Not mrsFormatInfo Is Nothing Then Set mrsFormatInfo = Nothing
    
    '自动卸载报表显示窗体
    For i = 0 To UBound(arrForms)
        If Not arrForms(i) Is Nothing Then
            Unload arrForms(i)
            Set arrForms(i) = Nothing
        End If
    Next
    Unload frmFlash
    
    Set gcnOLEDB = Nothing
    Set gcolOLEDBConnect = Nothing
    Set grsConnect = Nothing
    Set gclsCNs = Nothing
    'Set gfrmDBConnect = Nothing
    
    'Set gobjRegister = Nothing'(此部件中的对象只是引用，所以，由原部件来卸载)
    'Set gcnOracle = Nothing
    'Set grsObject = Nothing
    
    On Error Resume Next
    Kill gstrFind
    Err.Clear
    On Error GoTo 0
    
    mblnDataIsEmpty = False
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.count = 0)
End Function

Public Function SetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, _
    ByVal lngSysNO As Long, _
    ByVal strReportCode As String, _
    ByVal strKey As String, _
    ByVal strValue As String, _
    Optional ByVal bytType As Byte = 1, _
    Optional ByVal intFormat As Integer = 0 _
) As Boolean
'功能：保存自定义报表工具的打印设置信息
'参数：
'  cnOracle：数据库连接对象
'  lngSysNo：系统号
'  strReportCode：报表编号
'  strKey：项目名
'  strValue：项目值
'  bytType：类型
'    =1："报表类"的设置信息
'    =2："格式类"的设置信息
'  intFormat：指定格式
'    =0：表示默认的格式
'    >0：表示指定的格式

    Dim strSection As String, strFormat As String
    
    '检查
    If InStr("|format|allformat|paperbin|papercopy|paperform|paperformname|printer|", "|" & LCase(strKey) & "|") <= 0 Then
        Exit Function
    Else
        If InStr("|paperbin|papercopy|paperform|paperformname|printer|", "|" & LCase(strKey) & "|") <= 0 Then
            bytType = Val("1-报表类")
        Else
            bytType = Val("2-格式类")
        End If
    End If
    
    On Error GoTo hErr
    
    If bytType <= Val("1-报表类") Then
        strSection = "私有模块\" & App.ProductName & "\LocalSet\" & strReportCode
        GoSub makSub
    ElseIf intFormat = 0 Then
        strSection = "私有模块\" & App.ProductName & "\LocalSet\" & strReportCode & "\所有格式"
        GoSub makSub
    Else
        '初始化
        If gobjRegister Is Nothing Then
            On Error Resume Next
            Set gobjRegister = GetObject("", "zlRegister.clsRegister")
            Err.Clear: On Error GoTo hErr
        End If
        
        If gcnOracle Is Nothing Then
            Set gcnOracle = cnOracle
        End If
        If gcnOracle Is Nothing Then Exit Function
        
        If mrsFormatInfo Is Nothing Then
makGetInfo:
            Set mrsFormatInfo = GetReportFormat(strReportCode)
        Else
            mrsFormatInfo.Filter = ""
        End If
        If mrsFormatInfo.State <> adStateOpen And bytType <> 1 Then Exit Function
        If mrsFormatInfo.RecordCount > 0 Then
            If mrsFormatInfo!编号 <> strReportCode Then
                GoTo makGetInfo
            End If
        End If
        
        mrsFormatInfo.Filter = "序号=" & intFormat & " and 系统=" & lngSysNO
        If mrsFormatInfo.RecordCount = 1 Then
            strFormat = mrsFormatInfo!说明
            strSection = "私有模块\" & App.ProductName & "\LocalSet\" & strReportCode & "\" & strFormat
            GoSub makSub
        End If
    End If
    
    SetReportPrintSet = True
    Exit Function
    
makSub:
    Call SaveSetting("ZLSOFT", _
                     strSection, _
                     strKey, _
                     strValue)
    Return
    
hErr:
    Call ErrCenter
End Function

Public Function GetReportPrintSet( _
    ByVal cnOracle As ADODB.Connection, _
    ByVal lngSysNO As Long, _
    ByVal strReportCode As String, _
    ByVal strUser As String, _
    ByVal bytType As Byte, _
    Optional ByVal intFormat As Integer = 0, _
    Optional ByVal strKey As String = "" _
) As String
'功能：获取自定义报表工具的打印设置信息
'参数：
'  cnOracle：数据库连接对象
'  lngSysNo：系统号
'  strReportCode：报表编号
'  strUser：用户名
'  bytType：类型
'    =0：返回报表的所有设置信息
'    =1：返回"报表类"的设置信息
'    =2：返回"格式类"的设置信息
'  intFormat：格式号
'    =0：表示默认的格式
'    >0：表示指定的格式，如果没有找到指定格式的信息，则返回默认格式的信息；
'  strKey：项目名称
'返回:
'  bytType=0：报表类\项目名1=项目值1[,项目名2=项目值2]…[\\格式号1..N\项目名1=项目值1[,项目名2=项目值2]]...
'    注1： 报表类=-1值；格式号=0表示默认格式；格式号>0表示指定格式
'    注2： 信息分为三级； "\\"为一级分隔符； "\"为二级分隔符； ","为三级分隔符。
'  bytType=1：
'    strKey=""：项目名1=项目值1[,项目名2=项目值2]…
'    strKey#""：项目值
'  bytType=2：
'    intFormat=0
'      strKey=""：0\项目名1=项目值1[,项目名2=项目值2]]…[\\格式号1..N\项目名1=项目值1[,项目名2=项目值2]]...
'      strKey#""：项目值
'    intFormat=1..N
'      strKey=""：格式号1..N\项目名1=项目值1[,项目名2=项目值2]]…
'      strKey#""：项目值（如果没有找到项目，就找默认格式的项目）

    Dim strValue, strResult As String, strFormat As String, strTmp As String
    Dim strReportInfo As String
    Dim varTmp As Variant
    Dim i As Integer, j As Integer
    Dim bytVal As Byte
    
    On Error GoTo hErr
    
    '初始化
    If gobjRegister Is Nothing Then
        On Error Resume Next
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear: On Error GoTo hErr
    End If

    If gcnOracle Is Nothing Then
        Set gcnOracle = cnOracle
    End If
    If gcnOracle Is Nothing Then Exit Function
    
    If mrsFormatInfo Is Nothing Then
makGetInfo:
        Set mrsFormatInfo = GetReportFormat(strReportCode)
    Else
        mrsFormatInfo.Filter = ""
    End If
    If mrsFormatInfo.State <> adStateOpen And bytType <> 1 Then Exit Function
    If mrsFormatInfo.RecordCount > 0 Then
        If mrsFormatInfo!编号 <> strReportCode Then
            GoTo makGetInfo
        End If
    End If
    
    '处理
    Select Case bytType
    Case Val("1-报表类")
        bytVal = 1
        GoSub makReport
        strReportInfo = strValue
    Case Val("2-格式类")
        bytVal = 2
        GoSub makFormat
        strReportInfo = strValue
    Case Else
        '所有打印设置信息
        bytVal = 1
        GoSub makReport
        strReportInfo = strValue
        
        bytVal = 2
        GoSub makFormat
        strReportInfo = "-1\" & strReportInfo & "\\" & strValue
    End Select
    
    GetReportPrintSet = strReportInfo
    Exit Function

makReport:
    If Trim(strKey) = "" Then
        varTmp = Split("Format|AllFormat", "|")
        GoSub makSub
        strValue = strResult
    Else
        strValue = mdlPublic.GetRegPrinterInfo(strKey, strReportCode, "", , Val("99-报表类"), strUser)
    End If
    Return

makFormat:
    strValue = ""
    If intFormat > 0 Then
        '指定格式
        mrsFormatInfo.Filter = "序号=" & intFormat & " and 系统=" & lngSysNO
        If mrsFormatInfo.RecordCount = 1 Then
            If strKey = "" Then
                varTmp = Split("Printer|PaperBin|PaperCopy|PaperForm|PaperFormName", "|")
                strFormat = mrsFormatInfo!说明
                GoSub makSub
                strValue = strResult
            Else
                strValue = mdlPublic.GetRegPrinterInfo(strKey, strReportCode, mrsFormatInfo!说明, Nothing, Val("1-格式类"), strUser)
            End If
        End If
    Else
        If strKey = "" Then
            '所有格式
            varTmp = Split("Printer|PaperBin|PaperCopy|PaperForm|PaperFormName", "|")
            With mrsFormatInfo
                '默认格式
                strFormat = "默认格式"
                GoSub makSub
                strValue = strValue & "\\0\" & strResult
                
                '具体格式
                .Filter = "系统=" & lngSysNO
                Do While .EOF = False
                    strFormat = !说明
                    GoSub makSub
                    strValue = strValue & "\\" & CStr(!序号) & "\" & strResult
                    .MoveNext
                Loop
                If Left(strValue, 2) = "\\" Then strValue = Mid(strValue, 3)
            End With
        Else
            '指定项目名
            strValue = mdlPublic.GetRegPrinterInfo(strKey, strReportCode, strKey, Nothing, Val("1-格式类"), strUser)
        End If
    End If
    Return

makSub:
    strResult = ""
    For i = LBound(varTmp) To UBound(varTmp)
        If bytVal = 1 Then
            strTmp = mdlPublic.GetRegPrinterInfo(varTmp(i), strReportCode, "", , 99, strUser)
            strResult = strResult & "," & varTmp(i) & "=" & strTmp
        ElseIf bytVal = 2 Then
            strTmp = mdlPublic.GetRegPrinterInfo(varTmp(i), strReportCode, strFormat, , 1, strUser)
            strResult = strResult & "," & varTmp(i) & "=" & strTmp
        End If
    Next
    If strResult <> "" Then strResult = Mid(strResult, 2)
    Return
    
hErr:
    Call ErrCenter
End Function
    
Private Function GetReportFormat(ByVal strReportCode As String) As ADODB.Recordset
'功能：通过数据库获取报表格式信息
'参数：
'  strReportCode：报表编号
'返回：格式名称
    
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    
    strSQL = "" & _
            "Select a.序号, a.说明, b.系统, b.编号 " & vbNewLine & _
            "From zlRPTFMTs A, zlReports B " & vbNewLine & _
            "Where a.报表id = b.Id And b.编号 = [1] " & vbNewLine & _
            "Order By a.序号 "
    Set rsTemp = OpenSQLRecord(strSQL, "获取报表格式信息", UCase(strReportCode))
    
    Set GetReportFormat = rsTemp.Clone
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Public Function LoadReport(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, objForm As Object, LibDatas As Object, ParamArray arrPar() As Variant) As Boolean
'功能：按报表发布的程序号或报表编号加载报表。
'参数：
'     lngSys=报表所属的系统,当为共享系统报表时,传入为0
'     varReport:需要执行的报表编号或对应的程序ID(当报表为独立菜单程序时才使用程序ID)
'     cnMain:主程序的数据库连接
'     frmParent:主窗体
'     objForm：加载的报表预览窗体，用于嵌入式预览
'     arrPar():1.报表参数,格式为"参数名1=参数值1"、"参数数据源1"...,用是否有等号区分是报表参数还是附加报表数据。
'              2.对于选择器或下接框类型的参数,如果要想直接指定显示值/绑定值,可以以"参数名=显示值|绑定值"的格式传入
'                也可只传入"参数名=绑定值"的方式，程序会自动尝试去寻找对应的显示值。
'              3.对于有多个格式的报表,可以在参数中传入"ReportFormat=x",以确定要打开的报表格式,对报表组无效
'              4.在直接打印时,可以在参数中传入"PrintEmpty=1/0",表明当前格式中的所有表格无数据时,是否可以打印
'              5.如果明确打印或预览时需要进行套打，则在参数中传入"PressWork=1"，或"PressWorkFirst=1".
'                表示全部或首页套打，这样就只会对表格中的数据进行输出，对表头、表格线及其他元素不作输出。
'              6.如果打印页号时要确定起始计算基准页，可传入"StartPageNum=x"(x>=1)参数，[页号]标签会在此基础上计算。
'              7.最后附加参数:0=缺省值,可不传,表示正常(含报表及预览),1=直接到预览,2=直接打印,3-输出到Excel,4-输出到PDF
'              8.输出到Excel时，可以通过"ExcelFile=***.xls"参数来指定输出文件，指定了输出文件时不会自动打开
'              9.当打开的是报表组时,可以将所有报表的参数依次传入
'返回：调用是否成功
'说明：1.当使用程序号打开报表时,报表必须已经发布。当打开未发布的报表时,必须手工在数据库上增加权限。
'      2.参数其个数由报表定义时决定,每个形式为"参数名=参数值",如:"部门名称='内一科','妇产科'","开始时间=2000-01-02","姓名=张三","部门ID=23"。
'      3.这种方式允许主程序自已设计参数输入窗体,以加强使用者输入参数的灵活性及适用性(如输入操作员姓名时,可以加一个人员选择器)。
'      4.因为报表有多个数据源,每个数据源中有多个参数,这些参数名称唯一,序号却是乱的,所以传参数时要以名称为键。
'      5.参数如没有传或没有传完(包括报表格式参数),则系统会自动弹出参数输入窗体对未传参数进行输入。
    Dim bytStyle As Byte
    Dim blnCache As Boolean, blnDo As Boolean
    Dim intTmp As Integer, intState As Integer
    Dim strInfo As String, strGroup As String
    Dim objfrmShow   As Object
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then Exit Function
    
    Set gfrmMain = frmParent
    glngSys = lngSys
    
    mblnDataIsEmpty = False
    
    If Not objForm Is Nothing Then Unload objForm: Set objForm = Nothing
    
    Call InitPar
    
    '参数及显示方式
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    If bytStyle = 3 Then
        If zlRegInfo("授权性质") <> "1" Then
            MsgBox "试用或测试版本不能使用该功能。", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '初始化数据连接的记录集对象
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup, strInfo) Then
        '打开报表组
        If glngGroup = 0 Then Exit Function
        
        '报表状态
        strGroup = "组"
        GoSub proState
        If blnDo = False Then Exit Function
        
        '执行报表组
        Set gobjReport = Nothing
        LoadReport = ShowReport(frmParent, Me, bytStyle)
    Else
        '合法性检查
        If Not CheckPass(glngGroup) Then
            MsgBox "报表:[" & grsReport!编号 & "]" & grsReport!名称 & " 数据错误，不能执行该报表！", vbInformation, App.Title
            Exit Function
        End If
        '权限判断
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then
                '报表类
                MsgBox "你没有权限执行报表:[" & grsReport!编号 & "]" & grsReport!名称 & "，请确定系统是否正确授权！", vbInformation, App.Title
            Else
                '单据类
                MsgBox "你没有权限执行单据:[" & grsReport!编号 & "]" & grsReport!名称 & "，请通知管理员检查单据信息是否被改动！", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        strGroup = ""
        GoSub proState
        If blnDo = False Then Exit Function
        
        '读取报表
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.编号 = grsReport!编号 And gobjReport.系统 = Nvl(grsReport!系统, 0) _
                And Format(gobjReport.修改时间, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!修改时间, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '打开报表
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        LoadReport = GetReportFrom(frmParent, Me, bytStyle, objForm, LibDatas, mobjfrmReport)
    End If
    
    LoadReport = LoadReport And Not gblnError
    Exit Function
    
proState:
    intState = mdlPublic.ReportStateSwitch(lngSys, varReport, strGroup <> "")
    
    blnDo = False
    Select Case intState
    Case Val("0-报表不存在或未发布")
        If strInfo = "" Then
            MsgBox mdlPublic.FormatString("报表[1]不存在，请联系管理员！", strGroup), vbInformation, App.Title
        Else
'            MsgBox mdlPublic.FormatString("“[1]”报表[2]未发布，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
            blnDo = True
        End If
    Case Val("1-报表启用中")
        blnDo = True
    Case Val("2-报表停用中")
        MsgBox mdlPublic.FormatString("“[1]”报表[2]停用中，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
    End Select
    Return
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetReportPageRow(ByVal intPage As Integer, ByRef lngBeginRow As Long, ByRef lngEndRow As Long) As Boolean
    On Error GoTo errH
    If mobjfrmReport Is Nothing Then Exit Function
    If UBound(mobjfrmReport.marrPage) < 0 Then Exit Function
    If mobjfrmReport.marrPage(intPage - 1).count <= 0 Then Exit Function
    lngBeginRow = mobjfrmReport.marrPage(intPage - 1).Item(1).RowB
    lngEndRow = mobjfrmReport.marrPage(intPage - 1).Item(1).RowE
    GetReportPageRow = True
    Exit Function
errH:
    Err.Clear
End Function

Public Function ReportOpenForRec(ByVal cnMain As ADODB.Connection, ByVal lngSys As Long, _
    ByVal varReport As Variant, frmParent As Object, LibDatas As Object, ParamArray arrPar() As Variant) As Boolean
'功能：按报表发布的程序号或报表编号打开报表。
'参数：
'     lngSys=报表所属的系统,当为共享系统报表时,传入为0
'     varReport:需要执行的报表编号或对应的程序ID(当报表为独立菜单程序时才使用程序ID)
'     cnMain:主程序的数据库连接
'     frmParent:主窗体
'     LibDatas：传入的记录集，用于打印
'返回：调用是否成功
'说明：1.当使用程序号打开报表时,报表必须已经发布。当打开未发布的报表时,必须手工在数据库上增加权限。
'      2.参数其个数由报表定义时决定,每个形式为"参数名=参数值",如:"部门名称='内一科','妇产科'","开始时间=2000-01-02","姓名=张三","部门ID=23"。
'      3.这种方式允许主程序自已设计参数输入窗体,以加强使用者输入参数的灵活性及适用性(如输入操作员姓名时,可以加一个人员选择器)。
'      4.因为报表有多个数据源,每个数据源中有多个参数,这些参数名称唯一,序号却是乱的,所以传参数时要以名称为键。
'      5.参数如没有传或没有传完(包括报表格式参数),则系统会自动弹出参数输入窗体对未传参数进行输入。
    Dim blnCache As Boolean, blnDo As Boolean
    Dim intTmp As Integer, intState As Integer
    Dim strInfo As String, strGroup As String
    Dim bytStyle As Byte
    
    On Error GoTo errH
    
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    If Not ChartInstall Then Exit Function
    
    Set gcnOracle = cnMain
    gstrDBUser = GetDBUser
    If InitRegister = False Then Exit Function
    
    Set gfrmMain = frmParent
    glngSys = lngSys
    
    mblnDataIsEmpty = False
    
    Call InitPar
    '参数及显示方式
    garrPars = arrPar
    If UBound(arrPar) <> -1 Then
        If IsNumeric(arrPar(UBound(arrPar))) Then
            bytStyle = CByte(arrPar(UBound(arrPar)))
        End If
    End If
    
    If bytStyle = 3 Then
        If zlRegInfo("授权性质") <> "1" Then
            MsgBox "试用或测试版本不能使用该功能。", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    '初始化数据连接的记录集对象
    If grsConnect Is Nothing Then
        Call mdlPublic.SetControlDBConnect(grsConnect)
    Else
        If grsConnect.State <> adStateOpen Then
            Call mdlPublic.SetControlDBConnect(grsConnect)
        End If
    End If
    
    gblnError = False
    
    If isGroup(lngSys, varReport, glngGroup, strInfo) Then
        '打开报表组
        If glngGroup = 0 Then Exit Function
        
        '报表状态
        strGroup = "组"
        GoSub proState
        If blnDo = False Then Exit Function
        
        '执行报表组
        Set gobjReport = Nothing
        ReportOpenForRec = ShowReport(frmParent, Me, bytStyle)
    Else
        '合法性检查
        If Not CheckPass(glngGroup) Then
            MsgBox "报表:[" & grsReport!编号 & "]" & grsReport!名称 & " 数据错误，不能执行该报表！", vbInformation, App.Title
            Exit Function
        End If
        '权限判断
        intTmp = ReportCanQuery(glngGroup)
        If intTmp <> 0 Then
            If intTmp = 1 Then
                '报表类
                MsgBox "你没有权限执行报表:[" & grsReport!编号 & "]" & grsReport!名称 & "，请确定系统是否正确授权！", vbInformation, App.Title
            Else
                '单据类
                MsgBox "你没有权限执行单据:[" & grsReport!编号 & "]" & grsReport!名称 & "，请通知管理员检查单据信息是否被改动！", vbInformation, App.Title
            End If
            Exit Function
        End If
        
        strGroup = ""
        GoSub proState
        If blnDo = False Then Exit Function
        
        '读取报表
        If Not gobjReport Is Nothing And ReportReaded(glngGroup) Then
            If gobjReport.编号 = grsReport!编号 And gobjReport.系统 = Nvl(grsReport!系统, 0) _
                And Format(gobjReport.修改时间, "yyyy-MM-dd HH:mm:ss") = Format(grsReport!修改时间, "yyyy-MM-dd HH:mm:ss") Then
                blnCache = True
            End If
        End If
        If Not blnCache Then
            Set gobjReport = Nothing
            Set gobjReport = ReadReport(glngGroup)
        End If
        
        '打开报表
        glngGroup = 0
        If gobjReport Is Nothing Then Exit Function
        ReportOpenForRec = ShowReportForRec(frmParent, Me, LibDatas, bytStyle)
    End If
    
    ReportOpenForRec = ReportOpenForRec And Not gblnError
    Exit Function
    
proState:
    intState = mdlPublic.ReportStateSwitch(lngSys, varReport, strGroup <> "")
    
    blnDo = False
    Select Case intState
    Case Val("0-报表不存在或未发布")
        If strInfo = "" Then
            MsgBox mdlPublic.FormatString("报表[1]不存在，请联系管理员！", strGroup), vbInformation, App.Title
        Else
'            MsgBox mdlPublic.FormatString("“[1]”报表[2]未发布，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
            blnDo = True
        End If
    Case Val("1-报表启用中")
        blnDo = True
    Case Val("2-报表停用中")
        MsgBox mdlPublic.FormatString("“[1]”报表[2]停用中，请联系管理员！", strInfo, strGroup), vbInformation, App.Title
    End Select
    Return
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetReportPage() As Integer
    On Error GoTo errH
    If mobjfrmReport Is Nothing Then Exit Function
    GetReportPage = UBound(mobjfrmReport.marrPage) + 1
    Exit Function
errH:
    Err.Clear
End Function

Public Function ErrorSilentMode() As String
    If gblnSilentMode = False Then
        gstrErrorContent = ""
    End If
    ErrorSilentMode = gstrErrorContent
End Function

Private Function GetSilentMode(ByVal arrPar As Variant) As Boolean
    Dim i As Integer
    Dim arrTmp As Variant
    
    GetSilentMode = False
    For i = 0 To UBound(arrPar)
        If InStr(CStr(arrPar(i)), "=") > 0 Then
            arrTmp = Split(arrPar(i), "=")
            If UCase(Trim(arrTmp(0))) = UCase("SilentMode") Then
                GetSilentMode = Val(arrTmp(1)) = 1
                Exit For
            End If
        End If
    Next
End Function
