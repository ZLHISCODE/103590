Attribute VB_Name = "mdlPublic"
Option Explicit

Public gcnOracle As ADODB.Connection
Public gstrSysName As String                '系统名称
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String    '产品名称

Public gobjComlib As Object
Public gobjCommFun As Object
Public gobjControl As Object
Public gobjGrid As Object
Public gobjDatabase As Object
Public gobjXWHIS As Object     'RIS接口部件zl9XWInterface.clsHISInner
Public gblnXW As Boolean      '系统参数：“启用医学影像信息系统专业版接口”
Public glngSys As Long
Public glngModule As Long
Public gMainPrivs As String
Public gstrDBUser As String
Public gstrNodeNo As String          '当前站点编号；如果未设置启用站点，则为"-"
Public gcolPrivs As Collection              '记录内部模块的权限
Public gobjLIS As Object     'Lis部件
Public gobjPlugIn As Object    '插件对象

Public gstrLike As String  '项目匹配方法,%或空
Public gblnMyStyle As Boolean '使用个性化风格
Public gstrIme As String '自动的开启输入法
Public gbytCode As Byte '简码生成方式，0-拼音,1-五笔,2-两者
Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    性质 As String
    部门ID As Long
    部门码 As String
    部门名 As String
    专业技术职务 As String
    用药级别 As Long
End Type
Public UserInfo As TYPE_USER_INFO

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public glngTXTProc As Long '保存默认的消息函数的地址

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Function CaptionHeight() As Long
'功能:返回系统窗体标题栏高度(以象素为单位)
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
End Function

Public Sub SetItemInfo(lvw As Object, pan As Object)
'功能：根据Listview当前选中行，显示在状态条上
    Dim i As Integer, strInfo As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If lvw.SelectedItem.Text <> "" Then
        strInfo = "/" & lvw.ColumnHeaders(1).Text & ":" & lvw.SelectedItem.Text
    End If
    
    For i = 2 To lvw.ColumnHeaders.Count
        If lvw.SelectedItem.SubItems(i - 1) <> "" Then
            strInfo = strInfo & "/" & lvw.ColumnHeaders(i).Text & ":" & lvw.SelectedItem.SubItems(i - 1)
        End If
    Next
    If strInfo <> "" Then pan.Text = Mid(strInfo, 2)
End Sub

Public Function CheckFormInput(objForm As Object, Optional ByVal strToNumText As String = "") As Boolean
'参数:strToNumText--需要进行将千分位格式的金额转成正常金额格式的文本控件名称,允许有多个,可用,号等分隔
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                    If InStr(1, "," & UCase(strToNumText) & ",", "," & UCase(obj.Name) & ",") > 0 Then
                        strText = StrToNum(strText)
                    End If
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 _
                    Or InStr(strText, ",") > 0 _
                    Or InStr(strText, ";") > 0 _
                    Or InStr(strText, "|") > 0 _
                    Or InStr(strText, "~") > 0 _
                    Or InStr(strText, "^") > 0 Then
                    MsgBox "输入数据中包含非法字符！", vbInformation, gstrSysName
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function StrToNum(ByVal strNumber As String) As Double
    '功能:将字符串转换成数据
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function

Public Function GetIDDate(ID As String) As String
'功能：根据身份证号返回出生日期,格式"yyyy-MM-dd"
'参数：ID=身份证号,应该为15位或18位
    Dim strTmp As String
    
    If Len(ID) = 15 Then
        strTmp = Mid(ID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(ID) = 18 Then
        strTmp = Mid(ID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDDate = strTmp
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = gobjComlib.Nvl(varValue, DefaultValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = gobjComlib.ZVal(varValue, blnForceNum)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Set rsTmp = gobjDatabase.GetUserInfo
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            UserInfo.ID = rsTmp!ID
            UserInfo.用户名 = rsTmp!User
            UserInfo.编号 = rsTmp!编号
            UserInfo.简码 = Nvl(rsTmp!简码)
            UserInfo.姓名 = Nvl(rsTmp!姓名)
            UserInfo.部门ID = Nvl(rsTmp!部门ID, 0)
            UserInfo.部门码 = Nvl(rsTmp!部门码)
            UserInfo.部门名 = Nvl(rsTmp!部门名)
            UserInfo.性质 = Get人员性质
            UserInfo.专业技术职务 = Nvl(rsTmp!专业技术职务)
            GetUserInfo = True
        End If
    End If
    
    gstrDBUser = UserInfo.用户名
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function Get人员性质(Optional ByVal str姓名 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取当前登录人员或指定人员的人员性质
    '返回:返回人员性质,多个用逗号分离
    '编制:刘兴洪
    '日期:2014-04-09 13:46:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If str姓名 <> "" Then
        strSQL = "Select B.人员性质 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID And A.姓名=[1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取人员性质", str姓名)
    Else
        strSQL = "Select 人员性质 From 人员性质说明 Where 人员ID = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "获取人员性质", UserInfo.ID)
    End If
    Do While Not rsTmp.EOF
        Get人员性质 = Get人员性质 & "," & rsTmp!人员性质
        rsTmp.MoveNext
    Loop
    Get人员性质 = Mid(Get人员性质, 2)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function zlGetComLib() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取公共部件相关对象
    '返回:获取成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    
    Err = 0: On Error Resume Next
    Set gobjComlib = GetObject("", "zl9Comlib.clsComlib")
    Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
    Set gobjControl = GetObject("", "zl9Comlib.clsControl")
    Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
    Set gobjGrid = GetObject("", "zl9Comlib.clsGrid")
    gstrNodeNo = ""
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
    If Not gobjComlib Is Nothing Then zlGetComLib = True: Exit Function
    Err = 0: On Error Resume Next
    Set gobjComlib = CreateObject("zl9Comlib.clsComlib")
    'Call gobjComlib.InitCommon(gcnOracle)
    Set gobjCommFun = gobjComlib.zlCommFun
    Set gobjControl = gobjComlib.zlControl
    Set gobjDatabase = gobjComlib.zlDatabase
    If Not gobjComlib Is Nothing Then gstrNodeNo = gobjComlib.gstrNodeNo
    Err = 0: On Error GoTo 0
End Function

Public Sub InitLocPar()
    Dim strValue As String
    
    On Error Resume Next
    gstrLike = IIf(gobjDatabase.GetPara("输入匹配") = 0, "%", "")
    strValue = gobjDatabase.GetPara("输入法")
    gstrIme = IIf(strValue = "", "不自动开启", strValue)
    gbytCode = Val(gobjDatabase.GetPara("简码方式"))
    gblnMyStyle = gobjDatabase.GetPara("使用个性化风格") = "1"
    gblnXW = Val(gobjDatabase.GetPara(255, glngSys)) = 1
    If Err <> 0 Then Err.Clear
End Sub

Public Function Between(X, a, B) As Boolean
'功能：判断x是否在a和b之间
    Between = gobjComlib.Between(X, a, B)
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = gobjComlib.IntEx(vNumber)
End Function

Public Function GetInsidePrivs(ByVal lngProg As Long, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If Err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        Err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = gobjComlib.GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function

Public Function RecalcBirth(ByVal strAge As String, ByRef strDateOfBirth As String, Optional ByVal strCalcDate As String, Optional ByRef strMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人年龄获取病人出生日期
    '入参:strAge:病人年龄,如：23岁、1岁2月
    'strCalcDate-传入计算日期
    '返回:传入的病人年龄格式正确则计算返回出生日期,否则返回空
    'strMsg-返回警示信息
    '正确年龄格式:X岁[X月]、X月[X天]、X天、X小时[X分钟]
    '    X岁:X不能大于200,X月:X不能大于12,X天:X不能大于31,X小时:X不能大于24,X分钟:X不能大于59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBirthday As String, strSQL As String
    Dim strCurDate As String
    Dim intAge As Integer
    Dim rsTemp As New ADODB.Recordset
    
    '检查病人的年龄格式是否正确
    strSQL = "Select Zl_Age_Check([1]) From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "Zl_Age_Check", strAge)
    strMsg = Trim(Nvl(rsTemp.Fields(0).Value))
    If strMsg <> "" Then Exit Function
    
    '根据年龄计算出生日期
    strBirthday = ""
    If strCalcDate = "" Then
        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    Else
        strCurDate = strCalcDate
    End If
    
    If Not strAge Like "约*" And Not strAge Like "不详" Then
        If strAge Like "*岁" Or strAge Like "*岁*月" Then
            intAge = Mid(strAge, 1, InStr(1, strAge, "岁") - 1)
            strBirthday = Format(DateAdd("yyyy", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
            If Right(strAge, 1) = "月" Then
                intAge = Mid(strAge, InStr(1, strAge, "岁") + 1, Len(strAge) - InStr(1, strAge, "岁") - 1)
                strBirthday = Format(DateAdd("m", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
            End If
            strBirthday = Format(strBirthday, "YYYY-MM-DD")
        ElseIf strAge Like "*月" Or strAge Like "*月*天" Then
            intAge = Mid(strAge, 1, InStr(1, strAge, "月") - 1)
            strBirthday = Format(DateAdd("m", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
            If Right(strAge, 1) = "天" Then
                intAge = Mid(strAge, InStr(1, strAge, "月") + 1, Len(strAge) - InStr(1, strAge, "月") - 1)
                strBirthday = Format(DateAdd("d", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
            End If
            strBirthday = Format(strBirthday, "YYYY-MM-DD")
        ElseIf strAge Like "*天" Or strAge Like "*天*小时" Then
            intAge = Mid(strAge, 1, InStr(1, strAge, "天") - 1)
            strBirthday = Format(DateAdd("d", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
            strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
            If Right(strAge, 2) = "小时" Then
                intAge = Mid(strAge, InStr(1, strAge, "天") + 1, Len(strAge) - InStr(1, strAge, "天") - 2)
                strBirthday = Format(DateAdd("h", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
                strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
            End If
        ElseIf strAge Like "*小时" Or strAge Like "*小时*分钟" Then
            intAge = Mid(strAge, 1, InStr(1, strAge, "小时") - 1)
            strBirthday = Format(DateAdd("h", -1 * intAge, CDate(strCurDate)), "YYYY-MM-DD HH:mm")
            If Right(strAge, 2) = "分钟" Then
                intAge = Mid(strAge, InStr(1, strAge, "小时") + 2, Len(strAge) - InStr(1, strAge, "小时") - 3)
                strBirthday = Format(DateAdd("n", -1 * intAge, CDate(strBirthday)), "YYYY-MM-DD HH:mm")
            End If
            strBirthday = Format(strBirthday, "YYYY-MM-DD HH:mm")
        End If
    Else
        Exit Function
    End If
    strDateOfBirth = strBirthday
    RecalcBirth = True
End Function

Public Function CheckOldData(ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox) As Boolean
'功能：检查年龄输入值的有效性
'返回：
    If Not IsNumeric(txt年龄.Text) Then CheckOldData = True: Exit Function
    
    Select Case cbo年龄单位.Text
        Case "岁"
            If Val(txt年龄.Text) > 200 Then
                MsgBox "年龄不能大于200岁!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "月"
            If Val(txt年龄.Text) > 2400 Then
                MsgBox "年龄不能大于2400月!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
        Case "天"
            If Val(txt年龄.Text) > 73000 Then
                MsgBox "年龄不能大于73000天!", vbInformation, gstrSysName
                If txt年龄.Enabled And txt年龄.Visible Then txt年龄.SetFocus
                CheckOldData = False: Exit Function
            End If
    End Select
    CheckOldData = True
End Function

Public Function GetOldAcademic(ByVal DateBir As Date, ByVal str年龄单位 As String) As Long
'功能：根据当前的出生日期和年龄单位，计算理论上的年龄值
'返回：年龄
    Dim DatCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" 岁月天", str年龄单位) < 2 Then Exit Function
    
    DatCur = gobjDatabase.Currentdate
    
    strInterval = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
    lngOld = DateDiff(strInterval, DateBir, DatCur)
    If DateAdd(strInterval, lngOld, DateBir) > DatCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function

Public Function ReCalcOld(ByVal DateBir As Date, ByRef cbo年龄单位 As ComboBox, Optional ByVal lng病人ID As Long, Optional ByVal blnSetControl As Boolean = True, _
    Optional ByVal datCalc As Date) As String
'功能:根据出生日期重新计算病人的年龄,重设年龄单位
'参数:blnSetControl是否设置年龄单位控件
'     datCalc-指定计算日期,未指定时按系统时间计算
'返回:年龄,年龄单位
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    If datCalc = CDate(0) Then
        strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    Else
        strSQL = "Select Zl_Age_Calc([1],[2],[3]) old From Dual"
    End If
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, DateBir, datCalc)
    If blnSetControl = False Then
        ReCalcOld = Trim(Nvl(rsTmp!old))
        Exit Function
    End If
    
    If Not IsNull(rsTmp!old) Then
        If rsTmp!old Like "*岁" Or rsTmp!old Like "*月" Or rsTmp!old Like "*天" Then
            strTmp = Mid(rsTmp!old, 1, Len(rsTmp!old) - 1)
            If IsNumeric(strTmp) Then
                Call gobjControl.cboLocate(cbo年龄单位, Mid(rsTmp!old, Len(rsTmp!old), 1))
            Else
                strTmp = rsTmp!old
                cbo年龄单位.ListIndex = -1
            End If
        Else
            strTmp = rsTmp!old
            If IsNumeric(strTmp) Then
                cbo年龄单位.ListIndex = 0
            Else
                cbo年龄单位.ListIndex = -1
            End If
        End If
    End If
    If cbo年龄单位.ListIndex = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
    
    ReCalcOld = strTmp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function InitRS(ByVal strFields As String) As ADODB.Recordset
'功能:构造医嘱记录
'参数:strFields = "家属ID|adBigInt|20,关系|adVarChar|30,姓名|adVarChar|100,性别|adVarChar|4,年龄||20"
    Dim rs As ADODB.Recordset
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '字段名称|字段类型|字段长度 缺省字段类型 为adVarChar

    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            Select Case UCase(arrSubFeld(1) & "")
            Case UCase("adVarChar")
                FieldType = adVarChar
            Case UCase("adBigInt")
                FieldType = adBigInt
            Case UCase("adInteger")
                FieldType = adInteger
            Case UCase("adSingle")
                FieldType = adSingle
            Case Else
                FieldType = adVarChar
            End Select
            
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

