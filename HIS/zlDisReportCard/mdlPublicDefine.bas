Attribute VB_Name = "mdlPublicDefine"
Option Explicit

Public gstrSysName As String
Public gbytDiseaseType As Byte  '0表示疑似病例 1临床诊断病例 2确诊病例 3病原携带者 4阳性检测结果（献血员）5未选择
Public gbytAcute As Byte        '0表示急性 1表示慢性 2表示未分型 3表示未选择
Public gstrKey As String        '表示类集合关键字
Public gstrSql As String
Public glngSys As Long          '系统号
Public glngVersion As VersionLevel
Public gblnLock As Boolean  '是否只能修改病人信息

Public gLngTwo As Long     '打印定位
Public gLngThree As Long     '打印定位
Public gLngFour As Long     '打印定位

Public gobjCardEx As Object 'zlDisReportCardEx部件对象

'当前窗体格式已有的复选框控件字符串
Public Const GSTR_Controls = ",ucCheckType,ucAge,ucSex,ucFrom,ucCheckJob,ucCaseType1,ucCaseType2,ucPTB,ucAIDS,ucTyphia,ucHepatitis" & _
                            ",ucAnthrax,ucDysentery,ucSyphilis,ucMalaria,ucInfectiousB,ucInfectiousC,ucInfectiousA"


'需要用UCheckNorm表示的元素有：对象号1,对象号2,...
Public Const GSTR_OBJNO_2014 = ",2,6,9,12,14,15,16,20,21,22,23,24,25,26,27,28,29,30,32,33,34,35,36,37,45,"

'要素名称
Public Const GSTR_ELENAME_2014 = "卡片编号$报卡类别$姓名$家长姓名$身份证号$性别$出生日期$年龄$年龄单位" & _
                    "$工作单位$联系电话$病人属于$住址$患者职业$病例分类1$病例分类2$发病日期$诊断日期" & _
                    "$死亡日期$甲类传染病$乙类传染病$艾滋病$病毒性肝炎$炭疽$痢疾$肺结核$伤寒" & _
                    "$梅毒$疟疾$丙类传染病$其它传染病$监测性病$婚姻状况$学历$感染途径$异性传播" & _
                    "$血液传播$订正病名$退卡原因$报告单位$联系电话$填卡医生$填卡日期$备注$附加传染病"
'替换域
Public Const GSTR_REPLACE_2014 = "0$0$1$0$1$1$1$1$0$1$" & _
                                 "1$0$0$0$0$0$0$0$0$0$" & _
                                 "0$0$0$0$0$0$0$0$0$0$" & _
                                 "0$0$1$1$0$0$0$0$0$1$" & _
                                 "0$0$1$0$0"
'要素类型
Public Const GSTR_ELETYPE_2014 = "1$1$1$1$1$1$2$1$1$1$" & _
                                 "1$1$1$1$1$1$2$2$2$1$" & _
                                 "1$1$1$1$1$1$1$1$1$1$" & _
                                 "1$1$1$1$1$1$1$1$1$1$" & _
                                 "1$1$2$1$1"

'要素表示
Public Const GSTR_ELEIDT_2014 = "0$2$0$0$0$0$0$0$2$0$" & _
                                "0$2$0$2$2$2$0$0$0$3$" & _
                                "3$2$2$2$2$2$2$2$2$3$" & _
                                "0$3$2$2$2$2$2$0$0$0$" & _
                                "0$0$0$0$3"


'需要用UCheckNorm表示的元素有：对象号1,对象号2,...
Public Const GSTR_OBJNO_2016 = ",2,6,9,12,14,15,16,20,21,22,23,24,25,26,27,28,29,30,39,"

'要素名称
Public Const GSTR_ELENAME_2016 = "卡片编号$报卡类别$姓名$家长姓名$身份证号$性别$出生日期$年龄$年龄单位" & _
                    "$工作单位$联系电话$病人属于$住址$患者职业$病例分类1$病例分类2$发病日期$诊断日期" & _
                    "$死亡日期$甲类传染病$乙类传染病$艾滋病$病毒性肝炎$炭疽$痢疾$肺结核$伤寒" & _
                    "$梅毒$疟疾$丙类传染病$其它传染病$订正病名$退卡原因$报告单位$联系电话$填卡医生$填卡日期$备注$附加传染病"
'替换域
Public Const GSTR_REPLACE_2016 = "0$0$1$0$1$1$1$1$0$1$" & _
                                 "1$0$0$0$0$0$0$0$0$0$" & _
                                 "0$0$0$0$0$0$0$0$0$0$" & _
                                 "0$0$0$1$0$0$1$0$0"
'要素类型
Public Const GSTR_ELETYPE_2016 = "1$1$1$1$1$1$2$1$1$1$" & _
                                 "1$1$1$1$1$1$2$2$2$1$" & _
                                 "1$1$1$1$1$1$1$1$1$1$" & _
                                 "1$1$1$1$1$1$2$1$1"

'要素表示
Public Const GSTR_ELEIDT_2016 = "0$2$0$0$0$0$0$0$2$0$" & _
                                "0$2$0$2$2$2$0$0$0$3$" & _
                                "3$2$2$2$2$2$2$2$2$3$" & _
                                "0$0$0$0$0$0$0$0$3"
                           
Public gcnOracle As ADODB.Connection
Public Const conMenu_Manage_Save = 2601     '暂存
Public Const conMenu_Manage_Finish = 2602   '完成
Public Const conMenu_Manage_Cancel = 2603   '取消完成
Public Const conMenu_Manage_Exit = 2604     '退出
Public Const M_STR_MODULE_MENU_TAG = 26     '系统号
Public Const FCONTROL = 8
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    用药级别 As Long
End Type

Public UserInfo As TYPE_USER_INFO   '用户信息

Public Enum SignLevel
    cprSL_空白 = 0              '未签名
    cprSL_经治 = 1              '经治医师签名
    cprSL_主治 = 2              '主治医师签名
    cprSL_主任 = 3              '主任医师签名
    cprSL_正高 = 4              '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师
End Enum

Public Enum VersionLevel
    VL_2014 = 0
    VL_2016 = 1
End Enum

Public Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Public Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const WM_MOUSEWHEEL = &H20A          '鼠标滚动

Public glngOffsetX As Long, glngOffsetY As Long

'*************************************************************************
'**函 数 名：HIWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的高16位
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' 取出32位值的高16位
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**函 数 名：LOWORD
'**输    入：LongIn(Long) - 32位值
'**输    出：(Integer) - 32位值的低16位
'**功能描述：取出32位值的低16位
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' 取出32位值的低16位
     LOWORD = LongIn And &HFFFF&
End Function

Public Sub ClearInfo(objCtl As Control)
    On Error GoTo errHand
    
    Select Case TypeName(objCtl)
        Case "uCheckNorm"
            objCtl.Checked = False
        Case "TextBox"
            objCtl.Text = ""
    End Select
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub PrintInfo(ByVal objCtl As Control)

    Dim X As Integer
    Dim Y As Integer
    Dim strXY() As String
    Dim intOffset As Integer
    
    On Error GoTo errHand
    intOffset = 0   '保留，设置误差偏移量
    If objCtl.Tag <> "" Then
        strXY = Split(objCtl.Tag, ",")
        X = strXY(0) - intOffset
        Y = strXY(1) - intOffset
    Else
        Exit Sub
    End If
    
    Select Case TypeName(objCtl)
        Case "uCheckNorm"
            If objCtl.BoxVisible = True Then
                Printer.Line (glngOffsetX + PScaleX(X), glngOffsetY + PScaleY(Y + 2))-(glngOffsetX + PScaleX(X + 13), glngOffsetY + PScaleY(Y + 16)), &H0&, B
            End If
            
            If objCtl.Checked = True Then
                Printer.CurrentX = glngOffsetX + PScaleX(X + 1): Printer.CurrentY = glngOffsetY + PScaleY(Y + 4)
                Printer.FontName = "宋体": Printer.FontSize = 8
                Printer.Print "√"
            End If
            
            Printer.FontName = "仿宋_GB2312": Printer.FontSize = 9 '小五号

            If objCtl.BoxVisible = True Or objCtl.Name = "ucCheckType" Then
                Printer.CurrentX = glngOffsetX + PScaleX(X + 14)
                Printer.CurrentY = glngOffsetY + PScaleY(Y + 3)
            Else
                Printer.CurrentX = glngOffsetX + PScaleX(X)
                Printer.CurrentY = glngOffsetY + PScaleY(Y + 3)
            End If

            Printer.Print Trim(objCtl.Caption)
        Case "Label"
            Printer.FontName = "仿宋_GB2312": Printer.FontSize = IIf(objCtl.Name = "lblTitle", 18, 9)  '小五号
            If objCtl.Name = "Label1" Then Printer.FontSize = 8
            Printer.FontBold = IIf(objCtl.Name = "lblTitle", True, False)
            Printer.CurrentX = glngOffsetX + PScaleX(X)
            Printer.CurrentY = glngOffsetY + PScaleY(Y)
            Printer.Print Trim(objCtl.Caption)
            Printer.FontBold = False
        Case "TextBox"
            If objCtl.Name = "txtIDCard" Then
                Printer.Line (glngOffsetX + PScaleX(X), glngOffsetY + PScaleY(Y + 2))-(glngOffsetX + PScaleX(X + 14), glngOffsetY + PScaleY(Y + 17)), &H0&, B
                Printer.FontName = "仿宋_GB2312": Printer.FontSize = 9 '小五号
                Printer.CurrentX = glngOffsetX + PScaleX(X + 3)
                Printer.CurrentY = glngOffsetY + PScaleY(Y + 3)
                Printer.Print Trim(objCtl.Text)
                Exit Sub
            End If
            Printer.FontName = "仿宋_GB2312": Printer.FontSize = 9  '小五号
            Printer.CurrentX = glngOffsetX + PScaleX(X + 2)
            Printer.CurrentY = glngOffsetY + PScaleY(Y)
            Printer.Print Trim(objCtl.Text)
        Case "Line"
            Printer.Line (glngOffsetX + PScaleX(X), glngOffsetY + PScaleY(Y + 2))-(glngOffsetX + PScaleX(strXY(2)), glngOffsetY + PScaleY(Y + 2)), &H0&, B
    End Select
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function PScaleX(ByVal X As Single) As Single
'打印机的像素与屏幕的像素不一至，同样是210毫米，打印机像素是4960.625,屏幕是793.7
    PScaleX = Printer.ScaleX(Screen.TwipsPerPixelX * X, vbTwips, vbPixels)
End Function

Public Function PScaleY(ByVal Y As Single) As Single
    PScaleY = Printer.ScaleY(Screen.TwipsPerPixelY * Y, vbTwips, vbPixels)
End Function

Public Sub GetUserInfo()
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
        
    Set rsTemp = zlDatabase.GetUserInfo
    With rsTemp
        If .RecordCount <> 0 Then
            UserInfo.用户名 = .Fields("用户名").Value
            UserInfo.ID = .Fields("ID").Value                 '当前用户id
            UserInfo.编号 = .Fields("编号").Value             '当前用户编码
            UserInfo.姓名 = .Fields("姓名").Value             '当前用户姓名
            UserInfo.简码 = Nvl(.Fields("简码").Value, "")   '当前用户简码
            UserInfo.部门ID = .Fields("部门id").Value             '当前用户部门id
        Else
            UserInfo.用户名 = ""
            UserInfo.ID = 0
            UserInfo.编号 = ""
            UserInfo.姓名 = ""
            UserInfo.简码 = ""
            UserInfo.部门ID = 0
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Function AddStrKey(ByVal strKey As String) As Boolean
'功能：添加关键字
'返回：TRUE表示添加成功，False表示添加失败
    On Error GoTo errHand
    
    If InStr(gstrKey, strKey) = 0 Then
        gstrKey = gstrKey & "," & Trim(strKey)
        AddStrKey = True
    Else
        AddStrKey = False
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function CheckVal(ByRef intVal As Integer) As Boolean
    On Error GoTo errHand
    
    If InStr("0,1,2,3,4,5,6,7,8,9", Chr(intVal)) = 0 And intVal <> 8 Then
        intVal = 0
        CheckVal = False
    Else
        CheckVal = True
    End If
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub ShowMsg(ByVal strMsg As String)
    MsgBox strMsg, vbOKOnly + vbInformation, gstrSysName
End Sub

Public Function GetSaveSql(arrSql() As Variant, colCls As Collection, ByVal strFileId As String, strReportInfo) As Boolean
'功能：组织保存的Sql语句
'参数：arrSql:过程Sql数组
'      colcls:对象集合
'      strFile:文件ID
'      strReport:报告信息
    Dim objCls As clsReport
    Dim strAllInfo() As String  '所有报告信息格式：对象序号|内容文本
    Dim strObjNo() As String    '对象序号信息格式：对象序号1$对象序号2$对象序号3.......
    Dim strContent() As String
    Dim strReplace() As String  '替换域信息格式：替换域1$替换域2$替换域3.......
    Dim strEleName() As String  '要素名称信息格式：要素名称1$要素名称2$要素名称3.......
    Dim strEleType() As String  '要素类型信息格式：要素类型1$要素类型2$要素类型3.......
    Dim strEleIdt() As String   '要素表示信息格式：要素表示1$要素表示2$要素表示3.......
    Dim blnAddCol As Boolean    '是否需要增加新的对象到集合
    Dim strKey As String        '对象集合的关键字
    Dim i As Integer
    Dim intNo As Integer
    Dim strTmp As String
    On Error GoTo errHand
    
    GetSaveSql = False
    strAllInfo = Split(strReportInfo, "|")
    
    strObjNo = Split(strAllInfo(0), "$")   '对象序号信息格式
    strContent = Split(strAllInfo(1), "$") '内容文本
    
    If glngVersion = VL_2014 Then
        strReplace = Split(GSTR_REPLACE_2014, "$")  '替换域信息格式
        strEleName = Split(GSTR_ELENAME_2014, "$")  '要素名称信息格式
        strEleType = Split(GSTR_ELETYPE_2014, "$")  '要素类型信息格式
        strEleIdt = Split(GSTR_ELEIDT_2014, "$")    '要素表示信息格式
    Else
        strReplace = Split(GSTR_REPLACE_2016, "$")  '替换域信息格式
        strEleName = Split(GSTR_ELENAME_2016, "$")  '要素名称信息格式
        strEleType = Split(GSTR_ELETYPE_2016, "$")  '要素类型信息格式
        strEleIdt = Split(GSTR_ELEIDT_2016, "$")    '要素表示信息格式
    End If
    
    For i = 0 To UBound(strContent) - 1
        strKey = "K" & Trim(strObjNo(i))
        intNo = Val(Trim(strObjNo(i))) - 1
        blnAddCol = AddStrKey(strKey)
        
        Set objCls = colCls(strKey)
        objCls.FileID = Trim(strFileId)
        objCls.StartR = 1
        objCls.StopR = 0
        objCls.ObjNo = Trim(strObjNo(i))
        objCls.ObjType = IIf(Val(objCls.ObjNo) = 42, 8, 4)
        
        strTmp = Replace(Trim(strContent(i)), "、", "")
        strTmp = Replace(strTmp, "(", "")
        strTmp = Replace(strTmp, ")", "")
        
        objCls.Txt = strTmp
        objCls.Replace = Trim(strReplace(intNo))
        objCls.EleName = Trim(strEleName(intNo))
        objCls.EleType = Trim(strEleType(intNo))
        objCls.EleIdt = Trim(strEleIdt(intNo))
        objCls.EleRange = ""
        Call objCls.GetSaveSql(arrSql, blnAddCol)
    Next
    
    GetSaveSql = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetUserSignLevel(ByVal lngUserID As Long, ByVal lngPatiID As Long, lngPatiPageID As Long) As SignLevel
'## 说明：  根据“人员表”中的“聘任技术职务”字段确定医生技术级别（住院医师、主治医师、主任医师）
    Dim rs As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    Err = 0: On Error GoTo errHand
    
    gstrSql = "select 聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "mRichEPR", lngUserID)
    If Not rs.EOF Then
        lngR = Nvl(rs("聘任技术职务"), 0)
    End If
    Select Case lngR    '1 正高  2 副高  3 中级  4 助理/师级  5 员/士  9 待聘
    Case 1: lngLevel1 = cprSL_正高
    Case 2: lngLevel1 = cprSL_主任
    Case 3: lngLevel1 = cprSL_主治
    Case Else: lngLevel1 = cprSL_经治
    End Select
    If lngLevel1 = cprSL_正高 Then lngLevel1 = cprSL_主任 '正高：签名级别不包含，只表示人员居右正高职称，以便区别副主任医师;在本部件中不使用 正高
    rs.Close
    
    If lngPatiID > 0 Then
        gstrSql = "Select 经治医师, 主治医师, 主任医师 " & _
            " From 病人变动记录 " & _
            " Where 病人ID = [1] And 主页ID = [2] And (终止时间 Is Null Or 终止原因 = 1) " & _
            "       And 开始时间 Is Not Null And Nvl(附加床位, 0) = 0"
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, "cEPRDocument", lngPatiID, lngPatiPageID)
        If rs.EOF Then
            lngLevel2 = cprSL_经治
        Else
            If rs.Fields("主任医师") = UserInfo.姓名 Then
                lngLevel2 = cprSL_主任
            ElseIf rs.Fields("主治医师") = UserInfo.姓名 Then
                lngLevel2 = cprSL_主治
            Else
                lngLevel2 = cprSL_经治
            End If
        End If
    End If
    GetUserSignLevel = IIf(lngLevel1 >= lngLevel2, lngLevel1, lngLevel2)
    Exit Function
errHand:
    GetUserSignLevel = cprSL_空白
End Function

Public Function GetNextDoubleId(strTable As String) As Double
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strtab As String
    
    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!
    '31730
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "门诊费用记录" Or strtab = "住院费用记录" Then strtab = "病人费用记录"
    
    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetNextDoubleId")
    GetNextDoubleId = rsTmp.Fields(0).Value
'    Exit Function
'errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


Public Function CreateCardExOK(ByVal lngMod As Long) As Boolean
'功能：zlDisReportCardEx创建与检查
    If gobjCardEx Is Nothing Then
        On Error Resume Next
        Set gobjCardEx = CreateObject("zlDisReportCardEx.clsCardEx")
        If Not gobjCardEx Is Nothing Then
            Call gobjCardEx.Initialize(gcnOracle, glngSys, lngMod)
            Call CardExErrH(Err, "Initialize")
        End If
        Err.Clear: On Error GoTo 0
    Else
        On Error Resume Next
        If (Not gcnOracle Is Nothing) And gobjCardEx.gcnOracle Is Nothing Then
            Call gobjCardEx.Initialize(gcnOracle, glngSys, lngMod)
            Call CardExErrH(Err, "Initialize")
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjCardEx Is Nothing Then CreateCardExOK = True
End Function

Public Sub CardExErrH(ByVal objErr As Object, ByVal strFunName As String)
'功能：zlDisReportCardEx部件出错处理，
'参数：objErr 错误对象， strFunName 接口方法名称
'说明：当方法不存在（错误号438）时不提示，其它错误弹出提示框
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlDisReportCardEx部件执行 " & strFunName & " 时出错：" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Public Function GetReport() As Object
    Dim objCls As Object
    '检查是否开启外挂部件加载控件信息
    If Not gobjCardEx Is Nothing Then
        Err.Clear: On Error Resume Next
        Set objCls = New clsReportEx
        Set GetReport = gobjCardEx.GetReportEx(objCls)
        Call CardExErrH(Err, "getReportEx")
        Err.Clear: On Error GoTo 0
    End If
End Function


