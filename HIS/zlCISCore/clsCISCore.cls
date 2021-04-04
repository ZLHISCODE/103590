VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISCore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Set gobjCISCore = Me
End Sub

Public Sub InitCISCore(cnMain As ADODB.Connection)
'------------------------------------------------
'功能： 部件的初始化，将主程序的必要内容传递下来
'参数：
'   cnMain:主程序的数据库连接
'返回：
'------------------------------------------------
    Set gcnOracle = cnMain
    
    gstrAviPath = GetSetting("ZLSOFT", "注册信息", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrHelpPath = gstrAviPath & "\help"
    glngSys = 100
    Call GetUserInfo
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
End Sub

Public Function ShowSpecChar(frmParent As Object) As String
'功能：以模态窗体运行特殊字符程序
'参数：frmParent=调用父窗体
'返回：选择的特殊字符串；取消操作返回空
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Function EditMap(frmParent As Object, var元素 As Variant, Optional lng病历ID As Long, Optional blnViewOnly As Boolean) As Variant
'功能：在单独的模态窗体中编辑或查看指定的病历标记图
'参数：frmParnet=调用父窗体
'      var元素=标记图元素的编码(字符型)或ID(数字型)
'      lng病历ID=要修改的"病人病历内容"中标记图元素对应的ID；如果不传，则表示新增标注
'      blnViewOnly=是否只查看，不能编辑
'返回：返回字符串数组，可以自行组合到SQL语句中；未作任何标记时,该数组没有元素(Ubound=-1)
'      每个字符串的格式如："类型,'内容','字体','点集',X1,Y1,X2,Y2,填充色,填充方式,线条色,线型,线宽"
'      取消编辑或查看模式返回Empty(Not isArray)。
    Dim frmNew As frmMapEdit
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL() As Variant, strSQL As String
    
    Dim objCaseMap As StdPicture, i As Long
    Dim objMapItems As New MapItems, objMapItem As MapItem
    Dim lngMapID As Long, strMapName As String
    
    On Error GoTo errH
        
    '读取标记图元素的内容
    If TypeName(var元素) = "String" Then
        strSQL = "Select * From 病历元素目录 Where 编码=[1]"
    Else
        strSQL = "Select * From 病历元素目录 Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(var元素))
    If rsTmp.EOF Then Exit Function '必须要有图形背景
    
    lngMapID = rsTmp!ID
    strMapName = rsTmp!名称 & IIf(IsNull(rsTmp!说明), "", "(" & rsTmp!说明 & ")")
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Function '必须要有图形背景
    
    '读取标记图的标注内容
    If lng病历ID <> 0 Then
        strSQL = "Select * From 病人病历标记图 Where 病历ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lng病历ID)
        Do While Not rsTmp.EOF
            With rsTmp
                objMapItems.Add !类型, zlCommFun.NVL(!内容), _
                    IIf(IsNull(!字体), IIf(!类型 = 0, "宋体,9,0,0000", ""), !字体), _
                    zlCommFun.NVL(!点集), zlCommFun.NVL(!X1, 0), _
                    zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                    zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!填充色, &HFFFFFF), _
                    zlCommFun.NVL(!填充方式, -1), zlCommFun.NVL(!线条色, 0), _
                    zlCommFun.NVL(!线型, 0), zlCommFun.NVL(!线宽, 1), "_" & objMapItems.Count + 1
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    On Error GoTo 0
    
    Set frmNew = New frmMapEdit
    frmNew.mblnModi = Not blnViewOnly
    frmNew.mlngMapID = lngMapID
    frmNew.mstrMapName = strMapName
    Set frmNew.mobjCaseMap = objCaseMap
    Set frmNew.mobjMapItems = objMapItems
    frmNew.Show 1, frmParent
    
    If gblnOK Then
        arrSQL = Array()
        If frmNew.mobjMapItems.Count > 0 Then
            ReDim arrSQL(frmNew.mobjMapItems.Count - 1)
            For i = 1 To frmNew.mobjMapItems.Count
                Set objMapItem = frmNew.mobjMapItems(i)
                '"类型,'内容','字体','点集',X1,Y1,X2,Y2,填充色,填充方式,线条色,线型,线宽"
                With objMapItem
                    arrSQL(i - 1) = .类型 & ",'" & .内容 & "','" & .字体 & "','" & .点集 & "'," & _
                        .X1 & "," & .Y1 & "," & .X2 & "," & .Y2 & "," & .填充色 & "," & .填充方式 & "," & _
                        .线条色 & "," & .线型 & "," & .线宽
                End With
            Next
        End If
        EditMap = arrSQL
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMapInOjbect(objDraw As Object, var元素 As Variant, Optional lng病历ID As Long, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'功能：在指定的对象(PictureBox或Form)上显示标记图
'参数：objDraw=PictureBox或窗体对象,它的ScaleMode必须为Pixel
'      var元素=标记图元素的编码(字符型)或ID(数字型)
'      lng病历ID="病人病历内容"中标记图元素对应的ID,如果不传,仅显示标记图背景
'      X,Y,W,H=显示的目标客户端范围,可以不指定,单位为Pixel
'说明：可以利用该函数进行打印输出(因为是API作图,因此不能直接将objDraw指定为打印机,而是在PictureBox上按一定比例输出后,取PictureBox.Image输出到打印机)
    Call ShowMapInOjbect_1(objDraw, var元素, lng病历ID, x, y, W, H, blnMoved)
End Sub

Public Function EditPatientFile(ByVal FileID As String, Optional ByVal sPatientID As String = "", _
    Optional ByVal sPageID As String = "", Optional ByVal iPatientType As Integer = 0, _
    Optional ByVal sTemplateID As String = "", Optional ByVal bSample As Boolean = False, _
    Optional ParentForm As Object, Optional frmNewEdit As Object, Optional AllowEdit As Boolean = True, Optional FileType As Integer = 0, _
    Optional ByVal btModal As Byte = 0, Optional ByVal lngAdviceID As Long = 0) As Long
'功能：编辑病历文件或病历示范
'参数：FileID 病历记录ID或病历示范ID
'      sPatientID 病人ID
'      sPageID 主页ID或挂号单ID
'      iPatientType 0=门诊病人 1=住院病人
'      sTemplateID 病历模板文件ID
'      bSample 是否病历示范
'      FileType 病历种类：1-门诊病历;2-住院病历;3-护理记录;4-诊断文书;5-诊疗单据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If Len(FileID) = 0 Then
        strSQL = "Select Count(*)" + _
            " From 病历文件组成 Where 病历文件ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", sTemplateID)
        If rsTmp(0) = 0 Then
            MsgBox "未定义病历文件组成，不能编辑", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From 病人病历内容 Where 病历记录ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", FileID)
        If rsTmp(0) = 0 Then
            If Len(sTemplateID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From 病历文件组成 Where 病历文件ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", sTemplateID)
                If rsTmp(0) = 0 Then
                    MsgBox "未定义病历文件组成，不能编辑", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "没有病历内容，不能编辑", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    Set frmNewEdit = New frmFileEdit
    With frmNewEdit
        EditPatientFile = .ShowMe(FileID, sPatientID, sPageID, iPatientType, sTemplateID, bSample, ParentForm, AllowEdit, FileType, btModal, lngAdviceID)
    End With
End Function

Public Sub PreviewPatientFile(objParent As Object, ByVal lng病历种类 As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng病人ID As Long, _
                        ByVal var主页或单据 As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lng页码 As Long = 0, Optional ByVal lng开始页 As Long = 0, Optional ByVal lng结束页 As Long = 0)
'功能：打印所有病历
'参数：ObjParent        所有者对象
'       blnCurCase      是否为只打印输出当前这页
'       lngCurCase      指定当前打印输出的那份病历，打印输出时就从那份往后打印输出
'                       负数时表示病历记录ID
'       lng病人id
'       var主页或单据   如果是住院病人就记录主页ID，如果是门诊病人就记录挂号单，通过参数类型判断是住院还是门诊
'       blnPatiInfo     是否打印病人信息
'       lngY            打印开始的Y坐标
'       lng页码         设置起始的页码,为0时表示不打印页码
    Dim frmPreview As New frmCasePrint
    
    PrintOutCase objParent, frmPreview, lng病历种类, blnCurCase, lngCurCase, lng病人ID, var主页或单据, blnPatiInfo, lngY, lng页码, lng开始页, lng结束页
    frmPreview.Preview objParent, lng病历种类, blnCurCase, lngCurCase, lng病人ID, var主页或单据, blnPatiInfo, lngY, lng页码, lng开始页, lng结束页
End Sub

Public Sub PrintPatientFile(objParent As Object, ByVal lng病历种类 As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng病人ID As Long, _
                        ByVal var主页或单据 As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lng页码 As Long = 0, Optional ByVal lng开始页 As Long = 0, Optional ByVal lng结束页 As Long = 0)
'功能：打印所有病历
'参数：ObjParent        所有者对象
'       blnCurCase      是否为只打印输出当前这页
'       lngCurCase      指定当前打印输出的那份病历，打印输出时就从那份往后打印输出
'                       负数时表示病历记录ID
'       lng病人id
'       var主页或单据   如果是住院病人就记录主页ID，如果是门诊病人就记录挂号单，通过参数类型判断是住院还是门诊
'       blnPatiInfo     是否打印病人信息
'       lngY            打印开始的Y坐标
'       lng页码         设置起始的页码,为0时表示不打印页码
    Dim intPage As Integer
    
    intPage = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\打印设置", "纸张", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(objParent) Then
        MsgBox "打印机初始化失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    PrintOutCase objParent, Printer, lng病历种类, blnCurCase, lngCurCase, lng病人ID, var主页或单据, blnPatiInfo, lngY, lng页码, lng开始页, lng结束页
    'WinNT自定义纸张处理
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(objParent)
End Sub

Public Sub EditRequest(frmParent As Object, ByVal lng医嘱ID As Long, ByVal lng单据ID As Long, ByVal lng病历ID As Long, ByVal str医嘱内容 As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional frmEditForm As Object, Optional ByVal DataMoved As Boolean = False)
'功能：编辑申请单据
'参数：ModalWindow：编辑窗口是否为模式窗体
'      frmEditForm：返回编辑窗口对象
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe frmParent, lng医嘱ID, lng单据ID, lng病历ID, str医嘱内容, ReadOnly, ModalWindow, DataMoved
    End With
End Sub

Public Function getRequestForm() As Object
'功能：返回显示申请单据的窗体
    Set getRequestForm = New frmBillView
End Function

Public Sub AddRequest(frmParent As Object, ByVal lng病人ID As Long, ByVal var主页或挂号 As Variant, ByVal lng单据ID As Long, _
    ByVal bln护士站 As Boolean, Optional ByVal ModalWindow As Boolean = True, Optional frmEditForm As Object, Optional ByVal lng前提ID As Long = 0)
'功能：增加申请医嘱并编辑申请单据
'参数：ModalWindow：编辑窗口是否为模式窗体
'      frmEditForm：返回编辑窗口对象
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe_Request frmParent, lng病人ID, var主页或挂号, lng单据ID, bln护士站, ModalWindow, lng前提ID
    End With
End Sub

Public Sub EditReport(frmParent As Object, _
                        ByVal strNO As String, _
                        ByVal int记录性质 As Integer, _
                        ByVal lng单据ID As Long, _
                        ByVal lng病历ID As Long, _
                        ByVal str医嘱内容 As String, _
                        Optional ByVal ReadOnly As Boolean = False, _
                        Optional ByVal ModalWindow As Boolean = True, _
                        Optional frmEditForm As Object, _
                        Optional ByVal lng前提ID As Long = 0, _
                        Optional ByVal If医技执行 As Boolean = False, _
                        Optional ByVal blnShow申请 As Boolean = True, _
                        Optional ByVal lng医嘱ID As Long = 0, _
                        Optional ByVal DataMoved As Boolean = False, _
                        Optional ByVal strPrivs As String = "000")
'功能：编辑报告
'参数：ModalWindow：编辑窗口是否为模式窗体
'      frmEditForm：返回编辑窗口对象
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe_Report frmParent, strNO, int记录性质, lng单据ID, lng病历ID, str医嘱内容, ReadOnly, ModalWindow, lng前提ID, If医技执行, blnShow申请, lng医嘱ID, DataMoved, strPrivs
    End With
End Sub

Public Sub PrintSetup_Patient(objParent As Object)
'功能：病历打印设置
    frmPrintSet.Show vbModal, objParent
End Sub

Public Function PrintOptionSetup_Patient(objParent As Object, ByVal blnFirst As Boolean, ByRef blnCurCase As Boolean, _
    ByRef blnPatiInfo As Boolean, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, _
    Optional ByVal lng病历记录ID As Long, Optional ByRef lng开始页 As Long = 0, Optional ByRef lng结束页 As Long = 0, Optional ByRef lngPatientID As Long, _
    Optional ByRef lngPageID As Long) As Byte
'功能：病历打印选项
    PrintOptionSetup_Patient = frmCasePrintSet.PrintSet(objParent, blnFirst, blnCurCase, _
        blnPatiInfo, lngBeginY, intBeginPage, lng病历记录ID, lng开始页, lng结束页, lngPatientID, lngPageID)
End Function

'调用历史病案查询
Public Function QryPatientFile(frmParent As Object, Optional ByVal ModalWindow As Boolean = True) As Object
    Set QryPatientFile = frmPatiFileQry
    Call frmPatiFileQry.ShowMe(frmParent, ModalWindow)
End Function

Public Sub ShowDiagHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示疾病参考
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngItemId-要显示的疾病ID，不为0时，缺省不显示目录区；
    '---------------------------------------------
    Call frmDiagHelp.ShowMe(bytModal, frmParent, lngItemID)
End Sub

Public Sub ShowClinicHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示诊疗参考
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngItemId-要显示的诊疗项目ID，不为0时，缺省不显示目录区；
    '---------------------------------------------
    Call frmClinicHelp.ShowMe(bytModal, frmParent, lngItemID)
End Sub

Public Sub ShowRcdAnalyse(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngPatiId As Long)
    '---------------------------------------------
    '功能：根据上级程序要求，以模态或非模态显示病人病史分析
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngItemId-要显示的病人ID
    '---------------------------------------------
    Call frmRcdAnalyse.ShowMe(bytModal, frmParent, lngPatiId)
End Sub

Public Sub ShowDiagnotor(ByVal bytModal As Byte, ByVal frmParent As Object, _
    ByVal lngPatiId As Long, ByVal blnInpati As Boolean, _
    Optional ByVal lngPageID As Long = 1, Optional ByVal strRegistNo As String)
    '---------------------------------------------
    '功能：根据上级程序要求，调用病人疾病筛查程序，供医生参考
    '入参：frmParent-父窗体；
    '      blnModal-是否模态显示（通常和上级窗体一致）；
    '      lngPatiId-要显示的病人ID；
    '      blnInpati-是否住院病人，否则为门诊病人；
    '      lngPageId-要显示的住院病人主页ID；
    '      strRegist-要显示的门诊病人挂号单号；
    '---------------------------------------------
    Call frmDiagnotor.ShowMe(bytModal, frmParent, lngPatiId, blnInpati, lngPageID, strRegistNo)
End Sub

Public Sub EditElementDemo(ByVal frmParent As Object, ByVal blnAdd As Boolean, ByVal lng元素Id As Long, Optional ByVal lngDemoID As Long = 0)
    '---------------------------------------------
    '功能：编辑元素示范
    '入参：frmParent-父窗体；
    '      blnAdd-是否增加；
    '      lng元素Id-元素ID；
    '      lngDemoID-示范对应的病历内容ID
    '---------------------------------------------
    frmElementDemo.ShowMe frmParent, blnAdd, lng元素Id, lngDemoID
End Sub

Public Function RequestRegister(objParent As Object, ByVal lngDeptID As Long, Optional ByVal iItemType As Integer = 1) As Boolean
'功能：预约申请
'返回：确认/取消
    RequestRegister = frmReqEdit.ShowMe_Request(objParent, lngDeptID, iItemType)
End Function

Public Function ShowFileObject(ByVal frmMain As Object, _
                                    ByVal objContainer As Object, _
                                    ByVal lng医嘱ID As Long, _
                                    ByVal lng发送id As Long, _
                                    ByVal cnMain As ADODB.Connection, _
                                    ByVal strDBUser As String, _
                                    ByVal lngSys As Long, _
                                    Optional ByVal strPrivs As String = "", _
                                    Optional ByVal SysName As String = "检验") As Object
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    glngSys = lngSys
'    Set gfrmMain = frmMain
'    gstrUnitName = GetUnitName
    
    Call InitCISCore(gcnOracle)
        
    Set ShowFileObject = New frmFileRequest
    
    Load ShowFileObject
    
    lngStyle = GetWindowLong(ShowFileObject.hWnd, GWL_STYLE)
    
    Call SetWindowLong(ShowFileObject.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(ShowFileObject.hWnd, objContainer.hWnd)
    Call MoveWindow(ShowFileObject.hWnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
    '调用功能窗口数据刷新
    
'    Call ShowFileObject.zlRefresh(frmMain, lng医嘱ID, lng发送id, strPrivs, SysName)
    
    ShowFileObject.Show
End Function
