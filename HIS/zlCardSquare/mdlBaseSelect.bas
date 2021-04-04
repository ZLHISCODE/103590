Attribute VB_Name = "mdlBaseSelect"
Option Explicit
Public gbln存在站点控制 As Boolean

Public Function zl_GetFieldValue(ByVal rsTemp As ADODB.Recordset, _
    Optional ByVal strShowFields As String = "编码,名称", _
    Optional ByVal strShowSplit As String = "-") As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:返回显示字段的相关值
    '入参:rsTemp-记录集
    '     strShowFields-显示的字段
    '     strShowSplit-显示的分离符
    '出参:
    '返回:成功,返回相关的字段值
    '编制:刘兴洪
    '日期:2009-03-06 11:59:19
    '-----------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strValue As String, strLeft As String, strRight As String
    varData = Split(strShowFields, ",")
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.RecordCount = 0 Then Exit Function
    
    Select Case strShowSplit
    Case "[", "[]", "]"
        strLeft = "[": strRight = "]"
    Case "〖〗", "〖", "〗"
        strLeft = "〖": strRight = "〗"
    Case "【】", "【", "】"
        strLeft = "【": strRight = "】"
    Case "（）", "（", "）"
        strLeft = "（": strRight = "）"
    Case "〔〕", "〔", "〕"
        strLeft = "〔": strRight = "〕"
    Case "〈〉", "〈", "〉"
        strLeft = "〈": strRight = "〉"
    Case "［］", "［", "］"
        strLeft = "［": strRight = "］"
    Case "[]", "[", "]"
        strLeft = "[": strRight = "]"
    Case "｛｝", "｛", "｝"
        strLeft = "｛": strRight = "｝"
    Case "{}", "{", "}"
        strLeft = "{": strRight = "}"
    Case "「」", "「", "」"
        strLeft = "「": strRight = "」"
    Case "『』", "『", "』"
        strLeft = "『": strRight = "』"
    Case Else
        strLeft = "": strRight = strShowSplit
    End Select
    
    strValue = ""
    With rsTemp
        For i = 0 To UBound(varData) - 1
            strValue = strValue & strLeft & nvl(.Fields(varData(i))) & strRight
        Next
        strValue = strValue & nvl(.Fields(varData(UBound(varData))))
    End With
    zl_GetFieldValue = strValue
End Function
Public Function zl_获取部门信息(ByVal strCaption As String, Optional bln人员所属部门 As Boolean = False, _
    Optional str工作性质 As String = "", Optional strAddRoomCaption As String = "") As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取部门信息
    '     str工作性质-工作性质(以豆号分离,R,Z等)
    '     strAddRoomCaption-是否加入根项(只有再显示全部数据才显示时)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-03-03 14:47:10
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If str工作性质 <> "" Then str工作性质 = "," & str工作性质 & ","
    
    If str工作性质 <> "" Then
        strSQL = "" & _
        "   SELECT DISTINCT a.id,NULL as 上级ID,a.编码, a.名称,A.简码,A.末级" & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 " & _
        "       AND instr([2],','||b.编码||',')>0 " & _
        "       AND a.id = c.部门id " & zl_获取站点限制(True, "a") & _
        "       AND  a.撤档时间 >=to_date('3000-01-01','yyyy-mm-dd')" & _
                IIf(bln人员所属部门 = False, "", " and ID in (Select 部门id From 部门人员 where 人员id=[1]) ") & _
        "   order by a.编码"
    Else
        If bln人员所属部门 Or gbln存在站点控制 Then
            strSQL = "" & _
            "   Select ID, NULL as 上级ID,编码,名称,简码, 末级 " & _
            "   From 部门表" & _
            "   Where (撤档时间 Is NULL Or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01')  " & zl_获取站点限制 & _
                    IIf(bln人员所属部门 = False, "", " and ID in (Select 部门id From 部门人员 where 人员id=[1]) ") & _
            "  order by 编码"
        Else
            gstrSQL = ""
            If strAddRoomCaption <> "" Then
                gstrSQL = "" & _
                "  Select  -1 id, -NULL 上级id,'' 编码, '" & strAddRoomCaption & "' 名称,'" & zlCommFun.zlGetSymbol(strAddRoomCaption, 0) & "' 简码,0  as 末级" & _
                "   From dual  UNION ALL " & vbCrLf
            End If
            strSQL = gstrSQL & _
            "   Select ID," & IIf(strAddRoomCaption <> "", "nvl(上级id,-1)", "上级ID") & " as 上级ID,编码,名称,简码,末级 " & _
            "   From 部门表" & _
            "   Where (撤档时间 Is NULL Or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01')  " & _
            "   Start With 上级ID Is NULL Connect By Prior ID=上级ID"
        End If
    End If
    Set zl_获取部门信息 = zlDatabase.OpenSQLRecord(strSQL, strCaption, UserInfo.id, str工作性质)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zl_SetCtlBackColor(varCtls As Variant, Optional ByVal frmMain As Form = Nothing)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置控件的背景色
    '入参:objCtl-指定的控件
    '     frmMain-父窗口
    '出参:
    '编制:刘兴洪
    '日期:2009-03-03 11:55:10
    '-----------------------------------------------------------------------------------------------------------
    Dim lngNotEnabledBackColor As Long
    Dim objCtl As Object, i As Integer
    
    If IsArray(varCtls) = False Then
        varCtls = Array(varCtls)
    End If
        
    For i = 0 To UBound(varCtls)
        Set objCtl = varCtls(i)
        
        Select Case UCase(TypeName(objCtl))
        Case "TEXTBOX", "COMBOBOX"
        Case "CHECKBOX": Exit Sub
        Case "LABEL": Exit Sub
        Case "DTPICKER": Exit Sub
        Case UCase("CommandButton"): Exit Sub
        Case UCase("ListView"): Exit Sub
        End Select
        
        '禁止状态为窗体背景色,或者其为灰色
        lngNotEnabledBackColor = &H8000000A
        If Not frmMain Is Nothing Then lngNotEnabledBackColor = frmMain.BackColor
        If objCtl.Enabled Then
            objCtl.BackColor = &H80000005
        Else
            objCtl.BackColor = lngNotEnabledBackColor
        End If
    Next
End Sub

Public Function Select部门选择器(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln指定人员所属部门 As Boolean = False, _
    Optional strSQL As String = "", _
    Optional lngID As Long = 0, _
    Optional strTittle As String = "部门选择器", _
    Optional strNote As String = "", _
    Optional strNotFindMsg As String = "没有满足条件的部门,请检查!", _
    Optional strShowField As String = "编码,名称", _
    Optional strShowSplit As String = "-", _
    Optional lng人员ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:公共部门选择器
    '入参:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln指定人员所属部门-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '     strShowField-显示的相关字段
    '     strShowSplit-显示分离符号(strShowField为两字段以上才存在)
    '出参:lngID-返回部门ID(主要用于其他)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-03-06 12:18:31
    '-----------------------------------------------------------------------------------------------------------
 
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, lngH As Long, strFind As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim blnTree As Boolean
    
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strKey = GetMatchingSting(strSearch, False)
    
    blnTree = (strSearch = "" And str工作性质 = "" And bln指定人员所属部门 = False And strSQL = "" And Not (gstrNodeNo <> "-"))
    
    
    If strSQL <> "" Then
        gstrSQL = strSQL
    Else
    
        If str工作性质 = "" And bln指定人员所属部门 = False Then
            gstrSQL = "" & _
            "   Select   /*+ rule */ a.Id," & IIf(blnTree, "a.上级id", "-1*NULL ") & " as 上级ID,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
            "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间 " & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSQL = "" & _
            "   Select  /*+ rule */  distinct a.Id," & IIf(blnTree, "a.上级id", "-1*NULL ") & " as 上级ID,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
            "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间" & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c " & _
            IIf(str工作性质 = "", "", "       ,(Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.编码=J.column_value ") & _
            "         AND a.id = c.部门id " & _
            IIf(bln指定人员所属部门 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
        "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) " & zl_获取站点限制(True, "a") & ""
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
           strFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            strFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [3] "
        End If
    End If
    
    If blnTree Then
        gstrSQL = gstrSQL & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.编码"
    End If
   ' MsgBox GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX
    '坐标定位
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    If strSearch = "" And str工作性质 = "" And bln指定人员所属部门 = False And strSQL = "" Then
        '分上下级
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 1, strTittle, False, "", strNote, False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", strNote, False, False, True, sngX, sngY, lngH, blnCancel, False, False, IIf(lng人员ID = 0, UserInfo.id, lng人员ID), str工作性质, strKey)
    End If
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    Call zlControl.ControlSetFocus(objCtl, True)
    
    '设置相关的值
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp)
        objCtl.Tag = Val(rsTemp!id)
        zlCommFun.PressKey vbKeyTab
    End Select
    lngID = Val(nvl(rsTemp!id))
    Select部门选择器 = True
End Function


Public Function Select人员选择器(ByVal frmMain As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng部门ID As Long = 0, _
    Optional lng人员ID As Long = 0, _
    Optional bln按部门人员显示 As Boolean = False, _
    Optional strSearchKey As String = "", _
    Optional str人员性质 As String = "", _
    Optional str管理职务 As String = "", _
    Optional str专业技术职务 As String = "", _
    Optional strTittle As String = "人员选择器", _
    Optional strNote As String = "请选择相关的人员", _
    Optional strNotFindMsg As String = "未找到指定的人员,请检查!", _
    Optional strShowField As String = "姓名", _
    Optional strShowSplit As String = "-") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的人员
    '入参:frmMain-调用的父窗口
    '     objCtl-控件(目前只支持文本框)
    '     strKey-输入的建值
    '     lng部门ID-如果不为零,找所有人员,否则, 找指定部门下的人员
    '     str人员性质: 以医生,医生1... 格式
    '     str管理职务及str专业技术职务: 以职务1,职务21... 格式
    '出参:lng人员id-返回人员ID
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, bytType As Byte, str人员性质Table As String, strWhere As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '功能：多功能选择器
    '参数：
    '     frmMain=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    Err = 0: On Error GoTo Errhand:
    bytType = 0: strWhere = ""
    If str人员性质 <> "" Then
        str人员性质Table = "人员性质说明 Q1,(Select Column_Value From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) Q2" & vbCrLf
        strWhere = strWhere & " And ( A.ID=Q1.人员ID and Q1.人员性质 = Q2.Column_Value ) " & vbCrLf
    End If
    If str管理职务 <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))  Where a.管理职务=Column_Value) " & vbCrLf
    If str专业技术职务 <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))  Where a.专业技术职务=Column_Value) " & vbCrLf
    
    
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
        If lng部门ID = 0 Then
            gstrSQL = "" & _
                "   Select /*+ Rule */  A.ID,A.编号,A.姓名,A.别名,A.简码,A.性别,A.民族,A.出生日期,A.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 A " & str人员性质Table & _
                "   Where (A.姓名 like [1] or A.编号 like [1] or A.简码 like Upper([1]) or A.别名 like [1]) " & strWhere & zl_获取站点限制(True, "A") & "" & _
                "       and (A.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                "   order by A.编号"
        Else
            gstrSQL = "" & _
                "   Select   /*+ rule */ distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 a,部门人员 C " & str人员性质Table & _
                "   Where a.id=c.人员id and c.部门Id=[2]   " & strWhere & zl_获取站点限制(True, "a") & _
                "       and (a.姓名 like [1] or a.编号 like [1] or a.简码 like Upper([1]) or a.别名 like [1]) " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & _
                "   order by 编号"
        End If
     Else
        If lng部门ID = 0 Then
            If bln按部门人员显示 Then
                gstrSQL = "" & _
                "   Select   /*+ rule */ id," & IIf(gstrNodeNo <> "-", "1 as 级数ID,-1*NULL as 上级ID", "Level as 级数ID,上级id") & " ,编码,名称,0 末级,'' as 别名,'' as 简码,''as 性别,''as 民族, to_date(Null,'yyyy-mm-dd')  as 出生日期, '' as  办公室电话 ,'' 执业类别, '' 管理职务,'' 专业技术职务" & _
                "   From 部门表 " & _
                "   where 撤档时间 is null or 撤档时间>=to_date('3000-01-01','yyyy-mm-dd') " & zl_获取站点限制() & _
                    IIf(gstrNodeNo <> "-", "", "   Start with 上级id is null connect by prior id=上级id ") & _
                "   union all " & _
                "   Select a.ID,999999 AS 级数ID,b.部门id as 上级ID,a.编号,a.姓名,1 as 末级,别名,简码,性别,民族,出生日期,办公室电话,A.执业类别,A.管理职务,A.专业技术职务 " & _
                "   From 人员表 a,部门人员 b  " & str人员性质Table & _
                "   Where a.id=b.人员id and b.缺省=1  " & strWhere & zl_获取站点限制(True, "a") & _
                "         And (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
                "   Order by 级数ID,编码"
                bytType = 2
            Else
                gstrSQL = "" & _
                    "   Select  /*+ rule */ A.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                    "   From 人员表 A " & str人员性质Table & _
                    "   Where (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & strWhere & zl_获取站点限制(True, "a") & _
                    "   order by a.编号"
            End If
        Else
            gstrSQL = "" & _
                "   Select   /*+ rule */ distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话,A.执业类别,A.管理职务,A.专业技术职务" & _
                "   From 人员表 a,部门人员 C " & str人员性质Table & _
                "   Where a.id=c.人员id and c.部门Id=[2] " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & strWhere & zl_获取站点限制(True, "a") & _
                "   order by a.编号"
        End If
    End If
   
   
   '坐标定位
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, strTittle, bytType = 2, strSearchKey, strNote, bytType = 2, False, Not (bytType = 2), sngX, sngY, lngH, blnCancel, False, False, strKey, lng部门ID, str人员性质, str管理职务, str专业技术职务)
    
    lng人员ID = 0
    If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    If bytType = 2 Then
        strShowField = "," & strShowField & ",M_刘,"
        strShowField = Replace(strShowField, ",编号,", ",编码,")
        strShowField = Replace(strShowField, ",姓名,", ",名称,")
        strShowField = Mid(strShowField, 2)
        strShowField = Replace(strShowField, ",M_刘,", "")
    End If
    
    '设置相关的值
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
        objCtl.Tag = Val(rsTemp!id)
        zlCommFun.PressKey vbKeyTab
    End Select
    lng人员ID = Val(nvl(rsTemp!id))
    rsTemp.Close
    Select人员选择器 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Select供应商(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    Optional ByVal blnCheck执照效期 As Boolean = False, _
    Optional ByVal bln药品 As Boolean, Optional ByVal bln物资 As Boolean = False, _
    Optional ByVal bln设备 As Boolean = False, Optional ByVal bln其他 As Boolean = False, _
    Optional ByVal bln卫材 As Boolean = False, Optional blnOnlyName As Boolean = False, _
    Optional lng供应商ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:选择供应商
    '入参:objCtl-传入控件
    '    strKey-选择供应商
    '
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-25 11:26:43
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean, str类型 As String, bytType As Byte
    strKey = Replace(strKey, "'", "")
    '处理缺省值,以操作系统来确定相关的供应商
    If Not (bln其他 Or bln设备 Or bln卫材 Or bln药品 Or bln物资) Then
        bln药品 = Int(glngSys / 100) = 1
        bln卫材 = Int(glngSys / 100) = 1
        bln物资 = Int(glngSys / 100) = 4
        bln设备 = Int(glngSys / 100) = 6
    End If
    str类型 = IIf(bln药品, "1", "_")
    str类型 = str类型 & IIf(bln物资, "1", "_")
    str类型 = str类型 & IIf(bln设备, "1", "_")
    str类型 = str类型 & IIf(bln其他, "1", "_")
    str类型 = str类型 & IIf(bln卫材, "1", "_")
    
    Err = 0: On Error GoTo Errhand:
    
    '类型:以四位数表示,1位--药品供应商 2位--物资供应商　　3位--设备供应商　　4位--其他,   5-卫生材料 每位以1或零表示,1表示为true,0为false,以后扩充系统从第6位开始
    If strKey = "" Then
        '需要双列表形式
        gstrSQL = "" & _
            "   Select id,上级ID,编码,名称,末级,简码,许可证号,许可证效期,执照号,to_char(执照效期,'yyyy-mm-dd') as 执照效期 ,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
            "   From 供应商 " & _
            "   Where  (撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or 撤档时间 is null)  " & _
            "           and (( 类型 like [2]  And nvl(末级,0)=1 " & zl_获取站点限制 & ") or nvl(末级,0)=0 ) " & _
            "   Start with 上级ID is null connect by prior ID =上级ID " & _
            "   Order by level,ID"
        bytType = 2
    Else
        gstrSQL = "" & _
            "   Select id, 编码,名称,末级,简码,许可证号,许可证效期,执照号,to_char(执照效期,'yyyy-mm-dd') as 执照效期 ,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
            "   From 供应商 " & _
            "   where   末级=1 and 类型 like [2]  " & zl_获取站点限制 & "  " & _
            "           and  (撤档时间 is null or 撤档时间>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
            "           and (编码 like [1] or 名称 like [1] or 简码 like upper([1]))  "
            bytType = 0
    End If
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
    End If
 
    'ShowSelect:
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim sngX As Single, sngY As Single, lngH As Long
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, "供应商选择", IIf(bytType = 2, True, False), "", "请选择符合设备的供应商", IIf(bytType = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey, str类型)
    If blnCancel Then
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "不存在符何条件的供应商,请检查!"
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        Call zlControl.ControlSetFocus(objCtl)
        Exit Function
    End If
    If blnCheck执照效期 Then
        If Not IsNull(rsTemp!执照效期) Then
            If Format(zlDatabase.Currentdate, "yyyy-mm-dd") > nvl(rsTemp!执照效期) Then
                If MsgBox("供应商的执照效期已经过期,是否继续操作?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                    Call zlControl.ControlSetFocus(objCtl)
                    Exit Function
                End If
            End If
        End If
    End If
    Call zlControl.ControlSetFocus(objCtl)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, nvl(rsTemp!名称), "[" & nvl(rsTemp!编码) & "] " & nvl(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!id)
            Else
                .Text = IIf(blnOnlyName, nvl(rsTemp!名称), "[" & nvl(rsTemp!编码) & "] " & nvl(rsTemp!名称))
            End If
        End With
    Else
        With rsTemp
            objCtl.Text = "[" & nvl(!编码) & "] " & nvl(!名称)
            objCtl.Tag = nvl(!id)
        End With
        zlCommFun.PressKey vbKeyTab
    End If
    lng供应商ID = Val(nvl(rsTemp!id))
    Select供应商 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional bln未找到增加 As Boolean = False, Optional strOra过程 As String, Optional strWhere As String, _
    Optional bln站点 As Boolean = False, Optional blnNotMsg As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '     bln站点-是否进行站点限制
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str编码 As String, str名称 As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    str名称 = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  "
    End If
    gstrSQL = gstrSQL & strWhere & IIf(bln站点, zl_获取站点限制, "") & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If bln未找到增加 Then
            If zlCommFun.IsCharChinese(str名称) = False Then GoTo NOAdd::
            If MsgBox("注意:" & vbCrLf & _
                   "     未找到相关的" & strTable & ",是否增加“" & str名称 & "”？", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If zl_AutoAddBaseItem(strTable, str编码, str名称, strTable & "增加", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = str名称
                    End If
                End With
            Else
                If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                objCtl.Tag = str名称
                zlCommFun.PressKey vbKeyTab
            End If
            zl_SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
           If blnNotMsg = False Then ShowMsgbox "没有找到满足条件的" & strTable & ",请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, nvl(rsTemp!名称), nvl(rsTemp!编码) & "-" & nvl(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .Cell(flexcpData, .Row, .Col) = nvl(rsTemp!名称)
            Else
                .Text = IIf(blnOnlyName, nvl(rsTemp!名称), nvl(rsTemp!编码) & "-" & nvl(rsTemp!名称))
            End If
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = nvl(rsTemp!名称)
        objCtl.Tag = nvl(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    zl_SelectAndNotAddItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zl_AutoAddBaseItem(ByVal strTable As String, str编码 As String, str名称 As String, _
    Optional strTittle As String = "增加项目", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加项目信息(只针对有编码,名称的信息增加(只增加：编码和名称,简码)
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int编码 As Integer, strCode As String, strSpecify As String
    zl_AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的" & strTable & "，你要把它加入" & strTable & "中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int编码 = rsTemp!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str名称)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str编码 = strCode
    zl_AutoAddBaseItem = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zl_From人员获取缺省部门(ByVal lng人员ID As Long, ByRef str编码 As String, ByRef str名称 As String, ByRef lng部门ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-16 10:56:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    gstrSQL = "Select b.Id,b.编码,b.名称 From 部门人员 A,部门表  b Where a.部门id=b.Id And a.缺省=1 And a.人员id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取缺省部门信息", lng人员ID)
    If rsTemp.EOF Then Exit Function
    lng部门ID = nvl(rsTemp!id): str编码 = nvl(rsTemp!编码): str名称 = nvl(rsTemp!名称)
    zl_From人员获取缺省部门 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
