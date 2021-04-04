Attribute VB_Name = "mReport"
Option Explicit

Public pReport_CheckViewName As String
Public pReport_ResultName As String
Public pReport_AdviceName As String

Public preWinProc As Long
Public fReport As frmReportWord

Public Const ReportViewType_检查所见 = "检查所见"
Public Const ReportViewType_诊断意见 = "诊断意见"
Public Const ReportViewType_建议 = "建议"
Public Const ReportViewType_病理诊断 = "病理诊断"
Public Const ReportViewType_活检部位 = "活检部位"

'################################################################################################################
'## 功能：  判断指定用户是否是主任医师
'##
'## 参数：  lngUserID       ：用户ID
'##         strUserName     ：用户名
'##         lngPatiID       ：病人ID
'##         lngPatiPageID   ：主页ID
'##
'## 说明：  根据“人员表”中的“聘任技术职务”字段确定医生技术级别（住院医师、主治医师、主任医师）
'##         ＋病人变动记录中的医生级别，从而确定审核级别
'################################################################################################################
Public Function GetUserSignLevel(lngUserID As Long, Optional strUserName As String, _
    Optional lngPatiID As Long, Optional lngPatiPageID As Long) As EPRSignLevelEnum
    Dim RS As New ADODB.Recordset, lngR As Long, lngLevel1 As Long, lngLevel2 As Long
    
    err = 0: On Error GoTo errHand
    gstrSQL = "Select g.功能" & vbNewLine & _
            "From zlRoleGrant g, Sys.Dba_Role_Privs r, 上机人员表 p" & vbNewLine & _
            "Where r.Grantee = p.用户名 And g.角色 = r.Granted_Role And g.系统 = [2] And g.序号 = [3] And g.功能 = [4] And" & vbNewLine & _
            "      p.人员id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select [4] As 功能 From 上机人员表 p Where 用户名 = '" & UCase(UserInfo.用户名) & "' And p.人员id = [1]"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "mReport", lngUserID, glngSys, 1070, "签名权")
    If RS.RecordCount <= 0 Then GetUserSignLevel = cprSL_空白: Exit Function
    
    gstrSQL = "select 聘任技术职务 from 人员表 p where ID=[1]"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", lngUserID)
    If Not RS.EOF Then
        lngR = Nvl(RS("聘任技术职务"), 0)
    End If
    Select Case lngR    '1 正高  2 副高  3 中级  4 助理/师级  5 员/士  9 待聘
    Case 1: lngLevel1 = cprSL_正高
    Case 2: lngLevel1 = cprSL_主任
    Case 3: lngLevel1 = cprSL_主治
    Case Else: lngLevel1 = cprSL_经治
    End Select
    RS.Close
    
    If lngPatiID > 0 Then
        gstrSQL = "Select 经治医师, 主治医师, 主任医师 " & _
            " From 病人变动记录 " & _
            " Where 病人ID = [1] And 主页ID = [2] And (终止时间 Is Null Or 终止原因 = 1) " & _
            "       And 开始时间 Is Not Null And Nvl(附加床位, 0) = 0"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRDocument", lngPatiID, lngPatiPageID)
        If RS.EOF Then
            lngLevel2 = cprSL_经治
        Else
            If RS.Fields("主任医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
                lngLevel2 = cprSL_主任
            ElseIf RS.Fields("主治医师") = IIf(strUserName = "", UserInfo.姓名, strUserName) Then
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

'################################################################################################################
'## 功能：  搜索整个文本给出指定关键字区域的定位信息
'##
'## 参数：  edtThis         :   IN  ，编辑控件
'##         strKeyType      :   IN  ，给定关键字名称。取值为："O"、"P"、"T"、"E"、"U"
'##         lngKey           :   IN  ，给定欲查找的关键字ID号。
'##         lngKSS、lngKSE  :   OUT ，分别表示起始关键字的开始位置和结束位置；
'##         lngKES、lngKEE  :   OUT ，分别表示终止关键字的开始位置和结束位置；
'##         blnNeeded:      :   OUT ，是否是保留对象
'##
'## 返回：  如果找到该关键字具体位置，则返回True，否则返回False
'################################################################################################################
Public Function FindKey(ByRef edtThis As Object, _
        ByRef strKeyType As String, _
        ByRef lngKey As Long, _
        ByRef lngKSS As Long, _
        ByRef lngKSE As Long, _
        ByRef lngKES As Long, _
        ByRef lngKEE As Long, _
        ByRef blnNeeded As Boolean) As Boolean
        
    Dim i As Long, j As Long
    Dim sTMP As String
    Dim sText As String     '尽量少用.Text属性，因此用一个字符串变量来减少时间开支！
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    With edtThis
        sText = .Text   '只读取.Text属性1次！！！
        i = 1
LL1:
        i = InStr(i, sText, sTMP)
        If i <> 0 Then
            '看是否是关键字
            If .TOM.TextDocument.range(i - 1, i).Font.Hidden = False Then   '若为关键字，必须是隐藏且受保护的。
                i = i + 1
                GoTo LL1
            End If
            '已找到起始关键字
            
            '查找结束关键字
            j = i + 16
LL2:
            sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
            j = InStr(j, sText, sTMP)
            If j <> 0 Then
                '看是否是关键字
                If .TOM.TextDocument.range(j - 1, j).Font.Hidden = False Then
                    j = j + 1
                    GoTo LL2
                End If
                '找到结束关键字
                strKeyType = strKeyType
                lngKSS = i - 1 '转换为0开始的坐标位置。
                lngKSE = i + 15
                lngKES = j - 1
                lngKEE = j + 15
                blnNeeded = -Val(.TOM.TextDocument.range(i + 11, i + 12))
                FindKey = True
            End If
        End If
    End With
End Function


Public Sub richTextBoxShowElements(rText As RichTextBox)
    Dim strSel As String
    Dim miESingleS As Integer
    Dim miESingleE As Integer
    Dim miEMultiS As Integer
    Dim miEMultiE As Integer
    
    
    '判断当前选中内容是否要素
    If rText.SelColor = vbBlue Then
        miESingleS = InStrRev(rText.Text, "{{", rText.SelStart, vbTextCompare)
        miEMultiS = InStrRev(rText.Text, "{<", rText.SelStart, vbTextCompare)
        If miESingleS > miEMultiS Then  '当前最接近光标的是单选要素
            miESingleE = InStr(rText.SelStart, rText.Text, "}}", vbTextCompare)
            miESingleE = miESingleE + 1
            If miESingleE > miESingleS Then
                '是单选要素
                strSel = Left(rText.Text, miESingleE)
                strSel = Right(strSel, miESingleE - miESingleS + 1)
                frmReportElement.ShowElement strSel, 0
                rText.SelStart = miESingleS - 1
                rText.SelLength = miESingleE - miESingleS + 1
                rText.SelText = frmReportElement.strReturnElement
            End If
        ElseIf miEMultiS > miESingleS Then  '当前最接近的是多选要素
            miEMultiE = InStr(rText.SelStart, rText.Text, ">}", vbTextCompare)
            miEMultiE = miEMultiE + 1
            If miEMultiE > miEMultiS Then
                '是多选要素
                strSel = Left(rText.Text, miEMultiE)
                strSel = Right(strSel, miEMultiE - miEMultiS + 1)
                frmReportElement.ShowElement strSel, 1
                rText.SelStart = miEMultiS - 1
                rText.SelLength = miEMultiE - miEMultiS + 1
                rText.SelText = frmReportElement.strReturnElement
            End If
        Else    '两个要素的位置相等，说明都等于0，当前什么要素都没有
        
        End If
    End If
End Sub

Public Function Wndproc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer
    On Error Resume Next
    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)
    Select Case Msg
        Case WM_MOUSEWHEEL
            If fReport.picWordShow.Visible = False Or fReport.vscroWordH.Enabled = False Then Exit Function
            
            If Sgn(wzDelta) = 1 Then
                If fReport.vscroWordH.value - 1 < 0 Then
                    fReport.vscroWordH.value = 0
                Else
                    fReport.vscroWordH.value = fReport.vscroWordH.value - 1
                End If
            Else
                If fReport.vscroWordH.value + 1 > fReport.vscroWordH.Max Then
                    fReport.vscroWordH.value = fReport.vscroWordH.Max
                Else
                    fReport.vscroWordH.value = fReport.vscroWordH.value + 1
                End If
            End If
    End Select
    Wndproc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam)
End Function

Public Function zlGetWordPower() As Integer
'******************************************************************************************************************
'功能：获得当前用户的词句管理的权限
'返回：词句管理权限数值
'******************************************************************************************************************
    Dim intWordPower As Integer
    Dim strPrivs As String
    
    strPrivs = GetPrivFunc(glngSys, 1070)
    If InStr(1, strPrivs, "全院病历词句") <> 0 Then
        intWordPower = 0
    ElseIf InStr(1, strPrivs, "科室病历词句") <> 0 Then
        intWordPower = 1
    ElseIf InStr(1, strPrivs, "个人病历词句") <> 0 Then
        intWordPower = 2
    Else
        intWordPower = -1
    End If
    zlGetWordPower = intWordPower
End Function

Public Function zlDefaultWordCode(lngClassID As Long) As String
'功能：设置词句示范的默认编号
'参数： lngClassID --- 词句分类ID

    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    strSql = "Select LPad(Nvl(To_Number(Max(编号)), 0) + 1, Nvl(Max(Length(编号)), 5), '0') As 编码" & vbNewLine & _
            "From 病历词句示范" & vbNewLine & _
            "Where 分类id = [1]"
    err = 0: On Error Resume Next
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取词句编号", lngClassID)
    zlDefaultWordCode = rsTemp.Fields(0).value
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetSignSourceString(int提取类型 As Integer, lngReportID As Long, int签名版本 As Integer, blnMoved As Boolean, _
    thisSign As cEPRSign, strSourceOut As String) As Integer
'------------------------------------------------
'功能：获取用于电子签名，签名验证的报告源文内容
'参数： int提取类型 -- 1、签名时提取源文；2、签名验证时提取源文
'       lngReportID -- 报告ID，电子病历记录ID
'       int签名版本 -- 本次签名/验证签名提取源文的版本号
'       blnMoved --- 报告数据是否已经转储
'       thisSign --- 签名对象，签名的时候传入此对象，验证签名的时候传入nothing
'       strSourceOut -- 【返回】签名源文
'返回： 签名/验证签名的源文生成规则
'-----------------------------------------------
    Dim intRule As Integer
    Dim lng签名ID  As Long                  '签名所在的行的ID
    Dim strSql As String
    Dim rs病历记录 As ADODB.Recordset
    Dim rs病历内容 As ADODB.Recordset
    Dim rs签名记录 As ADODB.Recordset
    Dim str签名时间 As String
    Dim arr对象属性() As String
    
    '源文提取规则：
    'intRule = 1时，提取 ID，病人ID，婴儿，创建人，创建时间，医生姓名，签名级别，签名时间,检查所见，诊断意见，建议
    '验证签名的时候，医生姓名，签名级别，签名时间从签名记录中获取，分别是医生姓名= “内容文本”，签名级别=“要素表示”，签名时间 =“对象属性（5）”
    '签名的时候，医生姓名，签名级别，签名时间 从签名对象中获取
    On Error GoTo err
    
    If lngReportID = 0 Or int签名版本 = 0 Then Exit Function
    
    
    '初始化默认值
    intRule = 1
    strSourceOut = ""
    
    '根据int提取类型 来判断是签名还是验证签名，分别从对应的地方提取数据
    '从电子病历记录中提取报告源文的基本信息
    strSql = "Select ID,病人ID,婴儿,创建人,创建时间 From 电子病历记录 Where Id = [1]"
    Set rs病历记录 = zlDatabase.OpenSQLRecord(strSql, "提取报告源文基本信息", lngReportID)
    If rs病历记录.RecordCount = 0 Then
        Exit Function
    End If
    
    '从电子病历内容中提取报告源文的内容信息
    strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.开始版 = [2]  "
    Set rs病历内容 = zlDatabase.OpenSQLRecord(strSql, "提取报告源文内容信息", lngReportID, int签名版本)
    If rs病历内容.RecordCount = 0 Then
        Exit Function
    End If
    
    If int提取类型 = 1 Then
        '签名，检查签名对象是否存在
        If thisSign Is Nothing Then
            Exit Function
        End If
    Else
        '验证签名，从签名记录中提取医生姓名，签名级别，签名时间信息,签名规则
        strSql = "Select 内容文本 as 医生姓名 ,要素表示  as 签名级别 ,对象属性 From 电子病历内容 Where 文件ID = [1] And 对象类型 = 8 and 开始版 =[2] "
        Set rs签名记录 = zlDatabase.OpenSQLRecord(strSql, "提取最后报告源文签名信息", lngReportID, int签名版本)
        If rs签名记录.RecordCount = 0 Then
            Exit Function
        End If
        
        '提取格式化的签名时间，签名规则
        arr对象属性 = Split(rs签名记录!对象属性, ";")
        If UBound(arr对象属性) >= 5 Then
            intRule = Val(arr对象属性(1))
            str签名时间 = Format(arr对象属性(4), "yyyy-MM-dd HH:mm:ss")
        End If
        If intRule = 0 Then Exit Function
    End If
    
    '根据规则组织报告源文： ID，病人ID，婴儿，创建人，创建时间，医生姓名，签名级别，签名时间,检查所见，诊断意见，建议
    If intRule = 1 Then
        '源文基本信息
        strSourceOut = rs病历记录!ID
        strSourceOut = strSourceOut & vbTab & Nvl(rs病历记录!病人ID)
        strSourceOut = strSourceOut & vbTab & Nvl(rs病历记录!婴儿)
        strSourceOut = strSourceOut & vbTab & Nvl(rs病历记录!创建人)
        strSourceOut = strSourceOut & vbTab & Nvl(rs病历记录!创建时间)
        
        '源文签名信息
        If int提取类型 = 1 Then
            '签名，从签名对象提取
            strSourceOut = strSourceOut & vbTab & thisSign.姓名
            strSourceOut = strSourceOut & vbTab & thisSign.签名级别
            strSourceOut = strSourceOut & vbTab & Format(thisSign.签名时间, "yyyy-MM-dd HH:mm:ss")
        Else
            '验证签名，从数据库签名记录提取
            strSourceOut = strSourceOut & vbTab & Nvl(rs签名记录!医生姓名)
            strSourceOut = strSourceOut & vbTab & Nvl(rs签名记录!签名级别)
            strSourceOut = strSourceOut & vbTab & str签名时间
        End If
        
        '源文报告内容
        rs病历内容.Filter = "标题 ='" & ReportViewType_检查所见 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs病历内容!正文)
        End If
        
        rs病历内容.Filter = "标题 ='" & ReportViewType_诊断意见 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs病历内容!正文)
        End If
        
        rs病历内容.Filter = "标题 ='" & ReportViewType_建议 & "'"
        If rs病历内容.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & Nvl(rs病历内容!正文)
        End If
    End If
    
    GetSignSourceString = intRule
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
