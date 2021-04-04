VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOneCardDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'功能:一卡通涉及的相关数据对象(包含获取一卡通的方法，数据集等)
'涉及的方法或函数:
'  1.公共方法
'  1.0  InitCommon-初始化对象(第一步必须先初始化，否则可能使用中会发生异常)
'  1.1  zlInitPatiCards-初始化对象
'  1.2  ErrCenter-错误处理中心
'  1.3  SaveErrLog-保存错误日志
'  1.3  WritDebugLog-写调试日志
'  1.4  zlGetUserUnits:根据操作员获取操作员所涉及的病区ID或部门IDs
'  1.5  zlGetUnitRecordFromDepdIDs:根据部门IDs，获取相关的部门集
'  2.病人相关数据
'  2.1 zlGetPatiIDFromBedNumber:根据床位号，获取病人ID
'  2.2 zlGetPatiInforFromPatiID:根据病人ID获取病人信息对象
'  2.3 zlGetPatiInforFromXML:根据XML串，获取病人信息对象
'编制:刘兴洪
'日期:2018-12-03 11:06:14
'说明:该部件不区分版本
'*********************************************************************************************************************************************
Private mrs医疗卡类别  As ADODB.Recordset
Private mrs消费卡接口 As ADODB.Recordset
Private mrs收费类别 As ADODB.Recordset
Private mrs手机号 As ADODB.Recordset

Private mobjYLCards As Cards
Private mobjYLCardObjs As Cards   '当前启用有效的医疗卡
Private mobjXFCards As Cards
Private mobjXFCardObjs As Cards

Private gfrmCardMgr As Object
Private mrsSystem As ADODB.Recordset
Private mstrCardTypeColumns As String   '卡类别所涉及的列数
Private mobjDataBase As clsDataBase
Private mcnOracle As ADODB.Connection
Private mstrXFCardColumns As String '多个用逗号分离
Private mstrYLCardColumns As String '多个用逗号分离
Private mstrPatiCardCumns As String '病人医疗卡信息列集
Private mobjXML As clsXML
Private mstrLike As String

Public Sub InitCommon(cnOracle As ADODB.Connection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化连接
    '入参:cnOracle:主程序的数据库连接
    '编制:刘兴洪
    '日期:2018-12-03 11:10:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcnOracle = cnOracle
    If gcnOracle Is Nothing Then Set gcnOracle = cnOracle
    Set mrs医疗卡类别 = Nothing: Set mrs消费卡接口 = Nothing
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Sub
    mstrLike = IIf(Val(mobjDataBase.GetPara("输入匹配")) = 0, "%", "")
End Sub
Public Function ErrCenter() As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:错误处理中心
    '编制:刘兴洪
    '日期:2018-12-05 11:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Function
    End If
    
   ErrCenter = ErrCenter
End Function
Public Sub SaveErrLog()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:写日志
    '编制:刘兴洪
    '日期:2018-12-05 11:33:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.SaveErrLog
End Sub
Public Sub WritDebugLog(ByVal strDev As String, strInput As String, strOutPut As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:记录日志
    '编制:刘兴洪
    '日期:2018-12-05 11:35:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDataBase Is Nothing Then
        If zlGetOneDataBase(mcnOracle, mobjDataBase) = False Then Exit Sub
    End If
    Call mobjDataBase.WritLog(strDev, strInput, strOutPut)
End Sub


Public Function zlClearCache()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除相关缓存
    '编制:刘兴洪
    '日期:2018-12-04 15:17:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If Not mobjYLCards Is Nothing Then Set mobjYLCards = Nothing
    If Not mobjYLCardObjs Is Nothing Then Set mobjYLCardObjs = Nothing
    If Not mrs医疗卡类别 Is Nothing Then Set mrs医疗卡类别 = Nothing
    If Not mrs消费卡接口 Is Nothing Then Set mrs消费卡接口 = Nothing
    If Not mrs收费类别 Is Nothing Then Set mrs收费类别 = Nothing
    If Not mrsSystem Is Nothing Then Set mrsSystem = Nothing
    If Not mrs手机号 Is Nothing Then Set mrs手机号 = Nothing
    If Not mobjXFCards Is Nothing Then Set mobjXFCards = Nothing
    If Not mobjXFCardObjs Is Nothing Then Set mobjXFCardObjs = Nothing
End Function
Public Function zlGet医疗卡类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '返回:返回医疗卡类别的记录集
    '编制:刘兴洪
    '日期:2011-05-23 17:25:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, I As Long
    
    On Error GoTo errHandle
    
    If Not mrs医疗卡类别 Is Nothing Then
        If mrs医疗卡类别.State = 1 Then
            mrs医疗卡类别.Filter = 0
            Set zlGet医疗卡类别 = mrs医疗卡类别
            Exit Function
        End If
    End If
    
    
    '刘兴洪:主要是涉及兼容处理，所以使用了A.*(不分版本)动态创建
    strSQL = "" & _
    "   Select A.*, B.性质 as 结算性质 " & _
    "    From 医疗卡类别 A,结算方式 B" & _
    "    Where A.结算方式=B.名称(+)" & _
    "    Order by A.编码"
    
    Set mrs医疗卡类别 = mobjDataBase.OpenSQLRecord(strSQL, "获取消费卡接口 ")
    
    mstrYLCardColumns = ""
    For I = 0 To mrs医疗卡类别.fields.count - 1
        mstrYLCardColumns = mstrYLCardColumns & "," & UCase(mrs医疗卡类别.fields(I).Name)
    Next
    If mstrYLCardColumns <> "" Then mstrYLCardColumns = Mid(mstrYLCardColumns, 2)
    Set zlGet医疗卡类别 = mrs医疗卡类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetPatiIDFromBedNumber(ByVal lng病区ID As Long, ByVal str床号 As String, Optional ByRef lng主页ID As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    lng主页ID = 0
    strSQL = _
    "   Select  A.病人ID,A.主页ID" & _
    "   From 病人信息 A,床位状况记录 C" & _
    "   Where  A.病人ID=C.病人ID And A.停用时间 is NULL " & _
    "           And C.病区ID=[1] And C.床号=[2] "
   
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取病人信息", lng病区ID, str床号)
    If rsTemp.EOF Then zlGetPatiIDFromBedNumber = 0: Exit Function
    
    lng主页ID = Val(NVL(rsTemp!主页ID))
    zlGetPatiIDFromBedNumber = Val(NVL(rsTemp!病人ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlGetArea(frmParent As Object, txtInput As Object, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    If gobjDatabase Is Nothing Then zlInitCommLib
    If gobjDatabase Is Nothing Then Exit Function
    
    vRect = GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select 编码 as ID,编码,名称,简码 From 区域" & _
                 " Where Nvl(级数,0)<3 And (编码 Like [1] Or upper(简码) Like '" & mstrLike & "'||[1]||'%' Or 名称 Like '" & mstrLike & "'||[1]||'%')"
        Set zlGetArea = gobjDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, mstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select 编码 as ID,编码,名称,简码 From 区域 Where Nvl(级数,0)<3"
        Set zlGetArea = gobjDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, mstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetCardFindPati(ByVal strCardNo As String, Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号，模糊查找病人
    '        strCardNo-卡号
    '        blnNotShowErrMsg-不显示错误的提示信息
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        strCardPass-返回卡号的密码
    '        lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡类别ID As String, str病人ID As String, lngTemp As Long
    
    On Error GoTo errHandle
    
    '模糊查找
    '76020,冉俊明,2014-7-30,仍正常提取出了卡类别是支持模糊查找、已停用的持卡人信息
    '114161:李南春,2017/11/7,挂失有效天数为Null或0时，表示挂失一直生效
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    If mstrPatiCardCumns = "" Then
        If zlGetCojumnsFromTableName("病人医疗卡信息", mstrPatiCardCumns, "A", "状态,当前时间") = False Then Exit Function
        If mstrPatiCardCumns = "" Then mstrPatiCardCumns = "a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人,To_Date('3000-01-01','yyyy-mm-dd') as 终止使用时间"
    End If
    
    strSQL = "" & _
    " Select " & mstrPatiCardCumns & "," & _
    "       Case  When Nvl(a.状态, 0) = 1  And (Nvl(b.有效天数, 0)  = 0 Or Nvl(a.挂失时间, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(b.有效天数, 0) > Sysdate) Then 1" & _
    "             When Nvl(a.状态, 0) = 2 Then 2 Else 0 End As 状态,sysdate as 当前时间 " & _
    " From 病人医疗卡信息 A, 医疗卡挂失方式 B, 医疗卡类别 C" & _
    " Where a.卡类别id = c.Id And Nvl(c.是否模糊查找, 0) = 1" & _
    "      And a.卡号 = [2] And a.挂失方式 = b.名称(+)" & _
    "      And Nvl(c.是否启用, 0) = 1" & _
    " Order By 状态"
    
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取病人ID", lngCardTypeID, strCardNo)
    If rsTemp.EOF Then
        strErrMsg = "未找到卡号为" & strCardNo & "的病人，请检查该卡是否有效卡！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rsTemp.Filter = "状态=0"
    '0-正常有效卡;1-已挂失; 2-补卡停用
    If rsTemp.RecordCount = 1 Then
        '有一条，直接返回
        lng病人ID = Val(NVL(rsTemp!病人ID))
        strCardPassWord = NVL(rsTemp!密码)
        rsTemp.Close: Set rsTemp = Nothing
        GetCardFindPati = True: Exit Function
    End If
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '多条
    rsTemp.MoveFirst
    With rsTemp
        str病人ID = ""
        Do While Not .EOF
            lngTemp = Val(NVL(!卡类别id))
            If lngTemp <> 0 Then
                If InStr(1, str卡类别ID & ",", "," & lngTemp & ",") = 0 Then str卡类别ID = str卡类别ID & "," & lngTemp
                If InStr(1, str病人ID & ",", "," & Val(NVL(!病人ID)) & ",") = 0 Then str病人ID = str病人ID & "," & Val(NVL(!病人ID))
            End If
            .MoveNext
        Loop
        If str病人ID <> "" Then str病人ID = Mid(str病人ID, 2)
        If str卡类别ID <> "" Then str卡类别ID = Mid(str卡类别ID, 2)
        If InStr(1, str病人ID, ",") = 0 Then
            .MoveFirst
            If NVL(rsTemp!终止使用时间) <> "" Then
                If CDate(NVL(rsTemp!终止使用时间)) < CDate(rsTemp!当前时间) Then
                    strErrMsg = "卡号为" & strCardNo & "已失效！"
                    If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    rsTemp.Close: Set rsTemp = Nothing
                    Exit Function
                End If
            End If
            lng病人ID = Val(NVL(rsTemp!病人ID)): lngCardTypeID = Val(NVL(rsTemp!卡类别id))
            strCardPassWord = NVL(rsTemp!密码)
            rsTemp.Close: Set rsTemp = Nothing
            GetCardFindPati = True: Exit Function
        End If
        If frmSelectType.zlSelect(Nothing, str卡类别ID, lngCardTypeID) = False Then lngCardTypeID = 0: Exit Function
        rsTemp.Filter = "卡类别ID=" & lngCardTypeID & " And 状态=0"
        If rsTemp.EOF Then lngCardTypeID = 0: Exit Function
        
        If InStr("," & mstrPatiCardCumns & ",", ",A.终止使用时间,") > 0 Then
            If NVL(rsTemp!终止使用时间) <> "" Then
                If CDate(NVL(rsTemp!终止使用时间)) < CDate(rsTemp!当前时间) Then
                    strErrMsg = "卡号为" & strCardNo & "已失效！"
                    If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    rsTemp.Close: Set rsTemp = Nothing
                    Exit Function
                End If
            End If
        End If
        lng病人ID = Val(NVL(rsTemp!病人ID))
        strCardPassWord = NVL(rsTemp!密码)
        rsTemp.Close: Set rsTemp = Nothing: GetCardFindPati = True: Exit Function
    End With
    
    '肯定有误，按第一条提示
     rsTemp.Filter = 0: rsTemp.MoveFirst
    If InStr("," & mstrPatiCardCumns & ",", ",A.终止使用时间,") > 0 Then    '主要是做兼容处理
        If NVL(rsTemp!终止使用时间) <> "" Then
            If CDate(NVL(rsTemp!终止使用时间)) < CDate(rsTemp!当前时间) Then
                strErrMsg = "卡号为" & strCardNo & "已失效！"
                If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                rsTemp.Close: Set rsTemp = Nothing
                Exit Function
            End If
        End If
    End If
    
    If Val(NVL(rsTemp!状态)) = 1 Then
       strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
       If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If
    If Val(NVL(rsTemp!状态)) = 2 Then
       strErrMsg = "卡号为" & strCardNo & "已经被停用!"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
       Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardFromCardTypeID(ByVal strCardType As String, ByVal bln消费卡 As Boolean, ByRef objCard_Out As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡类别序号,获取指定的卡对象
    '入参:lngCardTypeID-指定的卡类别ID
    '     bln消费卡-是否消费卡
    '出参:objCard_Out-获取成功，返回Card对象,否则返回Nothing
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:09:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim blnName As String
    On Error GoTo errHandle
    If mobjYLCards Is Nothing Then zlInitPatiCards
    blnName = Not IsNumeric(strCardType)
    For Each objCard In mobjYLCards
        If blnName Then
            If objCard_Out.名称 = Trim(strCardType) And objCard.消费卡 = bln消费卡 Then
                 Set objCard_Out = objCard: zlGetCardFromCardTypeID = True: Exit Function
            End If
        Else
            If objCard_Out.接口序号 = Val(strCardType) And objCard.消费卡 = bln消费卡 Then
                Set objCard_Out = objCard: zlGetCardFromCardTypeID = True: Exit Function
            End If
        End If
    Next
    zlGetCardFromCardTypeID = False
    Set objCard_Out = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Function zlGetInputType(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : zlGetInputType
    ' 参数 : lngCardTypeID-卡类别ID
    ' 日期 : 2018-11-05
    ' 说明 : 返回true-转换为大写，false-不转换
    '---------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errH
    If zlGetCardFromCardTypeID(lngCardTypeID, False, objCard) = False Then Exit Function
    If objCard.启用 = False Then Exit Function
    zlGetInputType = objCard.卡号识别规则 = 1
    Exit Function
errH:
    If ErrCenter() Then Resume
End Function


Public Function zlGetPatiID(ByVal strCardType As String, _
    ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, _
    Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, _
    Optional objCtl As Object = Nothing, _
    Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, _
    Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, _
    Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, _
    Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '       blnNotCheckValidDate-是否对卡终止使用时间进行检查,true-不检查终止使用时间,false-检查
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str卡类别ID As String, str病人ID As Long, lngTemp As Long
    Dim strWhere As String, blnCard As Boolean '
    Dim str类别名称 As String, str标识号 As String
 
    
    On Error GoTo errHandle
    
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    strCardPassWord = "": strErrMsg = ""
    lng病人ID = 0
    If strCardType = "" Then Exit Function
    If Val(strCardType) = -1 Then
       zlGetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID)
       Exit Function
    End If
    
    str类别名称 = ""
    If strCardType Like "*身份证*" Or strCardType Like "*IC卡*" Then
        Set rsTemp = zlGet医疗卡类别
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Val(NVL(rsTemp!是否固定)) = 1 And Val(NVL(rsTemp!是否启用)) = 1 Then
                If strCardType Like "*身份证*" And NVL(rsTemp!名称) Like "*身份证*" And strCardType <> "联系人身份证" Then
                    str类别名称 = NVL(rsTemp!名称)
                     strCardType = Val(NVL(rsTemp!id)): Exit Do
                ElseIf strCardType Like "*IC卡*" And NVL(rsTemp!名称) Like "*IC卡*" Then
                     str类别名称 = NVL(rsTemp!名称)
                     strCardType = Val(NVL(rsTemp!id))
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If IsNumeric(strCardType) Then  '以卡类别ID为准
        
        If zlGetInputType(Val(strCardType)) Then
            strCardNo = UCase(strCardNo)
        End If
        If mstrPatiCardCumns = "" Then
            If zlGetCojumnsFromTableName("病人医疗卡信息", mstrPatiCardCumns, "A", "状态,当前时间") = False Then Exit Function
            If mstrPatiCardCumns = "" Then mstrPatiCardCumns = "a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人,To_Date('3000-01-01','yyyy-mm-dd') as 终止使用时间"
        End If
    
        strSQL = "" & _
        "   Select   " & mstrPatiCardCumns & "," & _
        "               nvl(挂失时间,to_date('3000-01-01','yyyy-mm-dd'))+nvl(B.有效天数,0) as 挂失时间,A.状态," & _
        "               sysdate as 当前时间" & _
        "   From 病人医疗卡信息 A,医疗卡挂失方式 B" & _
        "   Where  A.卡类别ID=[1] and A.卡号=[2] And A.挂失方式=B.名称(+)"
        Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取病人ID", Val(strCardType), strCardNo)
        If Not rsTemp.EOF Then
            If Not blnNotCheckValidDate And InStr("," & mstrPatiCardCumns & ",", ",A.终止使用时间,") > 0 Then '主要是做兼容处理
                If NVL(rsTemp!终止使用时间) <> "" Then
                    If CDate(NVL(rsTemp!终止使用时间)) < CDate(rsTemp!当前时间) Then
                        strErrMsg = "卡号为" & strCardNo & "已失效！"
                        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                        rsTemp.Close: Set rsTemp = Nothing
                        Exit Function
                    End If
                End If
            End If
            lng病人ID = Val(NVL(rsTemp!病人ID))
            strCardPassWord = NVL(rsTemp!密码)
            If Val(NVL(rsTemp!状态)) = 1 Then
                If Format(rsTemp!挂失时间, "yyyy-mm-dd") <= Format(rsTemp!当前时间, "yyyy-mm-dd") Then
                    strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
                    If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                    rsTemp.Close: Set rsTemp = Nothing
                    Exit Function
                End If
            End If
            If Val(NVL(rsTemp!状态)) = 2 Then
                strErrMsg = "卡号为" & strCardNo & "已经被停用!"
                 If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
                 rsTemp.Close: Set rsTemp = Nothing
                Exit Function
            End If
            zlGetPatiID = True
            Exit Function
        End If
        If blnOnlyContractPati Then Exit Function
        
        If str类别名称 = "" Then
              Set rsTemp = zlGet医疗卡类别()
              rsTemp.Filter = "ID=" & Val(strCardType) & " And 是否启用=1"
              If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
              If Not rsTemp.EOF Then
                  If Val(NVL(rsTemp!是否固定)) = 1 Then
                      str类别名称 = rsTemp!名称
                  End If
              End If
              rsTemp.Filter = 0
          End If
          
        If str类别名称 Like "*身份证*" Then
            strCardType = "身份证"
        ElseIf UCase(str类别名称) Like "*IC卡*" Then
            strCardType = "IC卡"
        Else
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
    End If
    
    '90875:李南春,2015/12/16,医疗卡证件类型
    If blnCertificate Then
        If mstrPatiCardCumns = "" Then
            If zlGetCojumnsFromTableName("病人医疗卡信息", mstrPatiCardCumns, "A", "状态,当前时间") = False Then Exit Function
            If mstrPatiCardCumns = "" Then mstrPatiCardCumns = "a.病人id, a.卡类别id, a.卡号, a.密码, a.状态, a.挂失时间, a.挂失方式, a.挂失人, a.发卡日期, a.发卡人,To_Date('3000-01-01','yyyy-mm-dd') as 终止使用时间"
        End If
        
        strSQL = "" & _
        "   Select   " & mstrPatiCardCumns & ",A.状态" & _
        "   From 病人医疗卡信息 A,医疗卡类别 B" & _
        "   Where A.卡类别ID=B.ID And A.状态=0 And B.是否启用=1 And B.名称=[1] and A.卡号=[2] And Nvl(B.是否证件,0)=1"
        Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取病人ID", strCardType, strCardNo)
        
        If rsTemp.EOF Then Exit Function
        lng病人ID = Val(NVL(rsTemp!病人ID))
        strCardPassWord = NVL(rsTemp!密码)
        If Val(NVL(rsTemp!状态)) = 1 Then
            strErrMsg = "卡号为" & strCardNo & "已经被挂失!"
            If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
            rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        If Val(NVL(rsTemp!状态)) = 2 Then
            strErrMsg = "卡号为" & strCardNo & "已经被停用!"
             If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
             rsTemp.Close: Set rsTemp = Nothing
            Exit Function
        End If
        zlGetPatiID = True
        Exit Function
    End If
    
     blnCard = True
    '问题:47939
    Select Case UCase(strCardType)
    Case "IC卡", "IC卡号"
        strWhere = "IC卡号=[2] "
    Case "身份证", "身份证号"
        strWhere = "身份证号=[2] "
    Case "联系人身份证" '问题号:51071
        strWhere = "联系人身份证号=[2]"
    Case "医保号", "医保证号"
        strWhere = "医保号=[2] "
    Case "手机号"
        strWhere = "手机号=[2] "
    Case "门诊号"
        strWhere = "门诊号=[3] "
        str标识号 = strCardNo
    '84247:李南春,2015/4/24,住院号查找病人
    Case "住院号"
        strWhere = "a.病人ID = (Select Nvl(Max(病人ID),0) As 病人ID From 病案主页 Where 住院号 = [3]) "
        str标识号 = strCardNo
    Case Else
        strWhere = "" & strCardType & "=[2] "
        blnCard = False
    End Select
    
    
    strSQL = "" & _
    "Select Rownum As ID, a.*" & vbNewLine & _
    "From (Select Decode(Nvl(max(a.在院), 0), 1, '√', '') As 在院, a.病人id, max(a.姓名) As 名称, max(a.性别)as 性别, max(a.年龄) as 年龄, max(a.身份证号) as 身份证号," & vbNewLine & _
    "             max(a.Ic卡号) as IC卡号,max( a.门诊号)as 门诊号, max(a.住院号)as 住院号,max(a.手机号)as 手机号,max( a.出生日期) as 出生日期, max(a.出生地点) as 出生地点," & vbNewLine & _
    "             max(a.费别) as 费别, max(a.医疗付款方式)as 医疗付款方式, max(a.民族) as 民族,max(a.家庭地址) as 家庭地址, max(a.家庭电话) as 家庭电话," & vbNewLine & _
    "             max(a.联系人姓名) as 联系人姓名, max(a.联系人关系)as 联系人关系, max(a.联系人电话) as 联系人电话,max(a.联系人身份证号) as 联系人身份证号," & vbNewLine & _
    "             max(a.住院次数)as 住院次数,max(a.卡验证码) As 密码id," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(B.类型, 1, Nvl(b.预交余额, 0), 0)), '99999999990.00')) As 门诊预交余额," & vbNewLine & _
    "             LTrim(To_Char(max(Decode(B.类型, 1, 0, Nvl(b.预交余额, 0))), '99999999990.00')) As 住院预交余额" & IIf(lngShowCardNoTypeID <> 0, ",max(c.卡号) as 卡号", "") & vbNewLine & _
    "       From 病人信息 A, 病人余额 B" & IIf(lngShowCardNoTypeID <> 0, ",病人医疗卡信息 C", "") & vbNewLine & _
    "       Where a.停用时间 Is Null And a.病人id = b.病人id(+) And b.性质(+) = 1 And " & IIf(lngShowCardNoTypeID <> 0, "a.病人ID=c.病人ID(+) And c.卡类别ID(+)=[5] And ", "") & strWhere & vbNewLine & _
    "       group by Nvl(a.就诊时间, Nvl(a.入院时间, a.登记时间)),a.病人ID" & vbNewLine & _
    "       Order By a.病人id,Nvl(a.就诊时间, Nvl(a.入院时间, a.登记时间))Desc " & IIf(lngShowCardNoTypeID <> 0, ", 卡号", "") & " ) A"
    Dim frmSel As New frmPatiSelect
    
    '80886,冉俊明,2014-12-18,当卡号为"310D664700068D9E",使用Val(strCardNo)会报"溢出"错误
    If Not frmSel.ShowSelect(frmMain, mcnOracle, glngSys, glngModul, objCtl, strSQL, "病人选择", "当前搜索出多个病人信息,请选择指定的病人", True, blnShowMergePati, _
                             IIf(objCtl Is Nothing, False, True), "", "密码ID,ID", rsTemp, blnUserCancel, Val(strCardType), strCardNo, Val(str标识号), lng病人ID, lngShowCardNoTypeID) Then
        If Not frmSel Is Nothing Then Unload frmSel
        Set frmSel = Nothing: Exit Function
    End If
    
    If Not frmSel Is Nothing Then Unload frmSel
    Set frmSel = Nothing
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    
    If rsTemp.EOF Then
        If blnCard Then
            'IC卡,可能是模糊查找
            zlGetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID)
        End If
        rsTemp.Close
        Exit Function
    End If
    lng病人ID = Val(NVL(rsTemp!病人ID))
    strCardPassWord = NVL(rsTemp!密码ID)
    zlGetPatiID = True: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Set rsTemp = Nothing
End Function


Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPati As clsPatiInfor, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String
    Dim objDatabase As Object
    
    On Error GoTo errHandle
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
     
    Set objPati = New clsPatiInfor
    
    '检查该身份证下有此病人没有
    If strOtherName = "" Then
        strWhere = " And A.病人ID=[1]"
    ElseIf strOtherName = "门诊号" Then
        strWhere = " And A.门诊号=[2]"
    ElseIf strOtherName = "住院号" Then
        strWhere = " And A.病人ID=(Select Max(病人ID) From 病案主页 Where 住院号 = [2])"
    Else
        strWhere = " And A." & strOtherName & "=[3]"
    End If
    strSQL = "" & _
    "   Select a.病人id, a. 门诊号, a.住院号, a.就诊卡号, a.卡验证码, a.费别, a.医疗付款方式,p.编码 as 医疗付款方式编码, a. 姓名, a.性别, a. 年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件, a.身份, " & _
    "        a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, " & _
    "        a.合同单位id, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人, a.担保额, a.担保性质, a.就诊时间, a.就诊状态, a.就诊诊室, a.在院, a.Ic卡号, " & _
    "        a.健康号, a.医保号, a.登记时间, a.停用时间, a.锁定, a.户口地址, a.户口地址邮编, a.籍贯, '' as 卡号, 0As 卡状态,'' as 密码, '' as 挂失方式, " & _
    "        a.病人类型 as 病人类型,sysdate as 挂失时间, 0  as 挂失有效天数,sysdate as 当前时间,a.手机号,a.险类,B.名称 险类名称" & _
    "   From 病人信息 A,保险类别 B,医疗付款方式 P" & _
    "   Where A.险类 = B.序号(+) And a.医疗付款方式=P.名称(+)  " & strWhere
    
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取病人信息", lng病人ID, Val(strOtherValue), strOtherValue)
 
    If rsTemp.EOF Then Exit Function
    
    objPati.病人ID = rsTemp!病人ID
    objPati.门诊号 = IIf(Val(NVL(rsTemp!门诊号)) = 0, "", NVL(rsTemp!门诊号))
    objPati.姓名 = NVL(rsTemp!姓名)
    objPati.性别 = NVL(rsTemp!性别)
    objPati.年龄 = NVL(rsTemp!年龄)
    objPati.出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
    objPati.出生地址 = NVL(rsTemp!出生地点)
    objPati.身份证号 = NVL(rsTemp!身份证号)
    objPati.其他证件 = NVL(rsTemp!其他证件)
    objPati.职业 = NVL(rsTemp!职业)
    objPati.费别 = NVL(rsTemp!费别)
    objPati.民族 = NVL(rsTemp!民族)
    objPati.国籍 = NVL(rsTemp!国籍)
    objPati.学历 = NVL(rsTemp!学历)
    objPati.婚姻状况 = NVL(rsTemp!婚姻状况)
    objPati.区域 = NVL(rsTemp!婚姻状况)
    objPati.家庭地址 = NVL(rsTemp!家庭地址)
    objPati.家庭电话 = NVL(rsTemp!家庭电话)
    objPati.家庭邮编 = NVL(rsTemp!家庭地址邮编)
    objPati.监护人 = NVL(rsTemp!监护人)
    objPati.联系人 = NVL(rsTemp!联系人姓名)
    objPati.联系人关系 = NVL(rsTemp!联系人关系)
    objPati.联系人地址 = NVL(rsTemp!联系人地址)
    objPati.联系人电话 = NVL(rsTemp!联系人电话)
    objPati.工作单位 = NVL(rsTemp!工作单位)
    objPati.工作单位电话 = NVL(rsTemp!单位电话)
    objPati.工作单位邮编 = NVL(rsTemp!单位邮编)
    objPati.工作单位开户行 = NVL(rsTemp!单位开户行)
    objPati.工作单位开户行帐户 = NVL(rsTemp!单位帐号)
    objPati.户口地址 = NVL(rsTemp!户口地址)
    objPati.户口地址邮编 = NVL(rsTemp!户口地址邮编)
    objPati.籍贯 = NVL(rsTemp!籍贯)
    objPati.密码 = NVL(rsTemp!密码)
    objPati.医疗付款方式编码 = NVL(rsTemp!医疗付款方式编码)
    objPati.医疗付款方式 = NVL(rsTemp!医疗付款方式)
    objPati.病人类型 = NVL(rsTemp!病人类型)
    objPati.就诊卡号 = NVL(rsTemp!就诊卡号)
    objPati.手机号 = NVL(rsTemp!手机号)
    objPati.险类 = Val(NVL(rsTemp!险类))
    objPati.险类名称 = Trim(NVL(rsTemp!险类名称))
    zlGetPatiInforFromPatiID = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlIsMobileNo(ByVal strInput As String, Optional ByRef strRutType As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:判断传入的是否为手机号
    '入参:strInput-手机号
    '出参:strRutType-查询结果:0-成功;1-不是有效号段;2-号码长度不对
    '返回:True-传入号码为手机号;False-传入号码不为手机号
    '编制:刘尔旋
    '日期:2017-1-25
    '---------------------------------------------------------------------------------------------
    Dim strMobileRange As String, strSQL As String
    
    On Error GoTo errHand
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    strRutType = 1
    If mrs手机号 Is Nothing Then
        strSQL = "Select 号段, 号码长度 From 手机号常用号段表"
        Set mrs手机号 = mobjDataBase.OpenSQLRecord(strSQL, "获取手机号段")
        If mrs手机号 Is Nothing Then Exit Function
        
        If mrs手机号.RecordCount = 0 Then Set mrs手机号 = Nothing:  Exit Function
    End If
    With mrs手机号
        Do While Not .EOF
            If NVL(!号段) = Mid(strInput, 1, Len(NVL(!号段))) Then
                If Len(strInput) = Val(NVL(!号码长度)) Then
                    strRutType = 0: zlIsMobileNo = True: Exit Do
                Else
                    strRutType = 2
                End If
            End If
            .MoveNext
        Loop
    End With
    If mrs手机号.RecordCount > 0 Then mrs手机号.MoveFirst
    Exit Function
errHand:
    strRutType = 1
    Set mrs手机号 = Nothing
End Function
 
 
Private Function zlGetSingleNodeValue(ByRef objXML As clsXML, ByRef strNodeName As String, Optional strNodePath As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定节点的值
    '入参:strNodeName-接点名
    '     strNodePath-节点路径
    '返回:接点值
    '编制:刘兴洪
    '日期:2018-12-05 14:10:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNodeTemp As String, strValue As String
    
    On Error GoTo errHandle
    
    If objXML Is Nothing Then Exit Function
    strNodeTemp = strNodePath & IIf(strNodeName = "", "", "/") & strNodeName
    If Not mobjXML.GetSingleNodeValue(strNodeTemp, strValue) Then Exit Function
    zlGetSingleNodeValue = strValue
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetPatiInforFromXML(ByVal strPatiXml As String, ByRef strErrMsg_Out As String, ByRef objPatiInfor_Out As clsPatiInfor) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从XML中获取病人信息
    '入参:strPatiXml-病人信息XML
    '
    '出参:objPatiInfor_Out-返回病人信息对象集
    '      strErrMsg_Out-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 14:29:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNodePath As String, strValue As String, strTmp As String
    
    On Error GoTo errHandle
    Set objPatiInfor_Out = New clsPatiInfor
    If strPatiXml = "" Then
        strErrMsg_Out = "接口返回（zlGetPatiInforFromXML）的XML无效，请与接口开发商或系统管理员联系!"
        Exit Function
    End If
    '打开XML文件
    If mobjXML.OpenXMLDocument(strPatiXml) = False Then
        strErrMsg_Out = "接口返回（zlGetPatiInforFromXML）的XML无效，请与接口开发商或系统管理员联系!"
        Exit Function
    End If
    strNodePath = "DATA"
    '    标识    数据类型    长度    精度    说明
    strValue = zlGetSingleNodeValue(mobjXML, "PatiXML", strNodePath)
    
    objPatiInfor_Out.PatiXML = strValue
    '    卡号    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "卡号", strNodePath)
    objPatiInfor_Out.卡号 = strValue
    '    姓名    Varchar2    64
    strValue = zlGetSingleNodeValue(mobjXML, "姓名", strNodePath)
    objPatiInfor_Out.姓名 = strValue
    '    性别    Varchar2    4
    strValue = zlGetSingleNodeValue(mobjXML, "性别", strNodePath)
    objPatiInfor_Out.性别 = strValue
    '    年龄    Varchar2    10
    strValue = zlGetSingleNodeValue(mobjXML, "年龄", strNodePath)
    objPatiInfor_Out.年龄 = strValue
    '    出生日期    Varchar2    20      yyyy-mm-dd hh24:mi:ss
    strValue = zlGetSingleNodeValue(mobjXML, "出生日期", strNodePath)
    objPatiInfor_Out.出生日期 = strValue
    '    出生地点    Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "出生地点", strNodePath)
    objPatiInfor_Out.出生地址 = strValue
    '    身份证号    VARCHAR2    18
    strValue = zlGetSingleNodeValue(mobjXML, "身份证号", strNodePath)
    objPatiInfor_Out.身份证号 = strValue
    
    If objPatiInfor_Out.出生日期 = "" And strValue <> "" Then
        strTmp = GetIDCardDate(strValue)
        If IsDate(strTmp) Then objPatiInfor_Out.出生日期 = strTmp
    End If
    
    '    其他证件    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "其他证件", strNodePath)
    objPatiInfor_Out.其他证件 = strValue
    '    职业    Varchar2    80
    strValue = zlGetSingleNodeValue(mobjXML, "职业", strNodePath)
    objPatiInfor_Out.职业 = strValue
    '    民族    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "民族", strNodePath)
    objPatiInfor_Out.民族 = strValue
    '    国籍    Varchar2    30
    strValue = zlGetSingleNodeValue(mobjXML, "国籍", strNodePath)
    objPatiInfor_Out.国籍 = strValue
    '    学历    Varchar2    10
    strValue = zlGetSingleNodeValue(mobjXML, "学历", strNodePath)
    objPatiInfor_Out.学历 = strValue
    '    婚姻状况    Varchar2    4
    strValue = zlGetSingleNodeValue(mobjXML, "婚姻状况", strNodePath)
    objPatiInfor_Out.婚姻状况 = strValue
    
    '    区域    Varchar2    30
    strValue = zlGetSingleNodeValue(mobjXML, "区域", strNodePath)
    objPatiInfor_Out.区域 = strValue
    '    家庭地址    Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "家庭地址", strNodePath)
    objPatiInfor_Out.家庭地址 = strValue
     '    户口地址    Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "户口地址", strNodePath)
    objPatiInfor_Out.户口地址 = strValue
    '    家庭电话    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "家庭电话", strNodePath)
    objPatiInfor_Out.家庭电话 = strValue
    '    家庭地址邮编    Varchar2    6
    strValue = zlGetSingleNodeValue(mobjXML, "家庭地址邮编", strNodePath)
    objPatiInfor_Out.家庭邮编 = strValue
    '    家庭地址    Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "户口地址", strNodePath)
    objPatiInfor_Out.户口地址 = strValue
    '    监护人  Varchar2    64
    strValue = zlGetSingleNodeValue(mobjXML, "监护人", strNodePath)
    objPatiInfor_Out.监护人 = strValue
    
    '    联系人姓名  Varchar2    64
    strValue = zlGetSingleNodeValue(mobjXML, "联系人姓名", strNodePath)
    objPatiInfor_Out.联系人 = strValue
    '    联系人关系  Varchar2    30
    strValue = zlGetSingleNodeValue(mobjXML, "联系人关系", strNodePath)
    objPatiInfor_Out.联系人关系 = strValue
    '    联系人地址  Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "联系人地址", strNodePath)
    objPatiInfor_Out.联系人地址 = strValue
    '    联系人电话  Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "联系人电话", strNodePath)
    objPatiInfor_Out.联系人电话 = strValue
    '    工作单位    Varchar2    100
    strValue = zlGetSingleNodeValue(mobjXML, "工作单位", strNodePath)
    objPatiInfor_Out.工作单位 = strValue
    '    单位电话    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "单位电话", strNodePath)
    objPatiInfor_Out.工作单位电话 = strValue
    '    单位邮编    Varchar2    6
    strValue = zlGetSingleNodeValue(mobjXML, "单位邮编", strNodePath)
    objPatiInfor_Out.工作单位邮编 = strValue
    '    单位开户行  Varchar2    50
    strValue = zlGetSingleNodeValue(mobjXML, "单位开户行", strNodePath)
    objPatiInfor_Out.工作单位开户行 = strValue
    '    单位帐号    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "单位帐号", strNodePath)
   'txt单位帐号.Text = strValue
    objPatiInfor_Out.工作单位开户行帐户 = strValue
    '    手机号    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "手机号", strNodePath)
    objPatiInfor_Out.手机号 = strValue
    '    照片文件    Varchar2    20
    strValue = zlGetSingleNodeValue(mobjXML, "照片文件", strNodePath)
    objPatiInfor_Out.照片文件 = strValue
    '    照片
    If Trim(strValue) <> "" Then
        Err = 0: On Error Resume Next
        objPatiInfor_Out.照片 = LoadPicture(strValue)
        If objPatiInfor_Out.照片 = 0 Then objPatiInfor_Out.照片 = Nothing
        Err = 0: On Error GoTo errHandle
    End If
    mobjXML.CloseXMLDocument
    zlGetPatiInforFromXML = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCards_XF(ByRef objXFCards_Out As Cards) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡对象
    '出参:objXFCards_Out-返回消费卡对象
    '返回:获取 成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 18:01:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjXFCards Is Nothing Then Set objXFCards_Out = mobjXFCards: zlGetCards_XF = True: Exit Function
    Call zlInitXFCards  '重新初始化
    If Not mobjXFCards Is Nothing Then Set objXFCards_Out = mobjXFCards: zlGetCards_XF = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetCardObjs_XF(ByRef objXFCards_Out As Cards) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡对象
    '入参:
    '
    '出参:objXFCards_Out-返回消费卡对象
    '返回:获取 成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 18:01:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjYLCardObjs Is Nothing Then Set objXFCards_Out = mobjYLCardObjs: zlGetCardObjs_XF = True: Exit Function
    Call zlInitXFCards  '重新初始化
    If Not mobjYLCardObjs Is Nothing Then Set objXFCards_Out = mobjYLCardObjs: zlGetCardObjs_XF = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInitXFCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡集
    '返回:成功!返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-15 14:31:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, bln启用 As Boolean, str部件 As String, objCard As Card
    Dim objBrushCardObject As Object, int自动间隔 As Integer
    Dim str读卡性质 As String
    
    Err = 0: On Error GoTo errHand:
    Set mobjXFCards = New Cards
    Set mobjYLCardObjs = New Cards
    
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    Set rsTemp = zlGet消费卡接口
    '自制卡(即消费卡)
    rsTemp.Filter = "自制卡=1"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
        int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编号), "自动读取", "0"))
        bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编号), "启用", "1")) = 1
        int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编号), "自动读取间隔", "1"))
        str部件 = NVL(rsTemp!部件)
            
        Set objCard = New Card
        With objCard
            .卡种类 = EM_CardType_Consume
            .接口序号 = NVL(rsTemp!编号)
            .接口编码 = NVL(rsTemp!编号)
            .短名 = Left(NVL(rsTemp!名称), 1)   '默认取第一个
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度))
            .系统 = Val(NVL(rsTemp!系统)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动间隔
            .自制卡 = Val(GetFieldValueFromRecord("自制卡", rsTemp, True)) = 1
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(GetFieldValueFromRecord("是否全退", rsTemp, True)) = 1
            .结算方式 = Trim(GetFieldValueFromRecord("结算方式", rsTemp, True))
            .接口程序名 = Trim(GetFieldValueFromRecord("部件", rsTemp, True))
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = Trim(GetFieldValueFromRecord("是否密文", rsTemp, True))
            .消费卡 = True
            .是否退现 = Val(GetFieldValueFromRecord("是否退现", rsTemp, True)) = 1
            .密码长度 = Val(GetFieldValueFromRecord("密码长度", rsTemp, True))
            .密码长度限制 = Val(GetFieldValueFromRecord("密码长度限制", rsTemp, True))
            .密码规则 = Val(GetFieldValueFromRecord("密码规则", rsTemp, True))
            .密码输入限制 = Val(GetFieldValueFromRecord("密码输入限制", rsTemp, True))
            .是否缺省密码 = Val(GetFieldValueFromRecord("是否缺省密码", rsTemp, True)) = 1
            .是否制卡 = Val(GetFieldValueFromRecord("是否制卡", rsTemp, True)) = 1
            .是否发卡 = Val(GetFieldValueFromRecord("是否发卡", rsTemp, True)) = 1 Or .自制卡
            .是否写卡 = Val(GetFieldValueFromRecord("是否写卡", rsTemp, True)) = 1
            
            str读卡性质 = GetFieldValueFromRecord("读卡性质", rsTemp, True)
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .结算性质 = Val(GetFieldValueFromRecord("结算性质", rsTemp, True))
        End With
        Set objCard = mobjXFCards.Add(objCard, "K" & NVL(rsTemp!编号))
        If zlCreateBrushObjects(objCard, objBrushCardObject) Then
            Set objCard.CardObject = objBrushCardObject
            mobjXFCardObjs.Add objCard, "K" & NVL(rsTemp!编号)
        End If
        rsTemp.MoveNext
    Loop
    
    '银联卡
    rsTemp.Filter = "自制卡<>1"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编号), "自动读取", 0))
        int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编号), "自动读取间隔", "1"))
        bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编号), "启用", "1")) = 1
        str部件 = NVL(rsTemp!部件)
        Set objCard = New Card
        With objCard
            .卡种类 = EM_CardType_Consume
            .接口序号 = NVL(rsTemp!编号)
            .接口编码 = NVL(rsTemp!编号)
            .短名 = Left(NVL(rsTemp!名称), 1)   '默认取第一个
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度))
            .系统 = Val(NVL(rsTemp!系统)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动间隔
            .自制卡 = Val(GetFieldValueFromRecord("自制卡", rsTemp, True)) = 1
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(GetFieldValueFromRecord("是否全退", rsTemp, True)) = 1
            .结算方式 = Trim(GetFieldValueFromRecord("结算方式", rsTemp, True))
            .接口程序名 = Trim(GetFieldValueFromRecord("部件", rsTemp, True))
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = Trim(GetFieldValueFromRecord("是否密文", rsTemp, True))
            .消费卡 = True
            .是否退现 = Val(GetFieldValueFromRecord("是否退现", rsTemp, True)) = 1
            .密码长度 = Val(GetFieldValueFromRecord("密码长度", rsTemp, True))
            .密码长度限制 = Val(GetFieldValueFromRecord("密码长度限制", rsTemp, True))
            .密码规则 = Val(GetFieldValueFromRecord("密码规则", rsTemp, True))
            .密码输入限制 = Val(GetFieldValueFromRecord("密码输入限制", rsTemp, True))
            .是否缺省密码 = Val(GetFieldValueFromRecord("是否缺省密码", rsTemp, True)) = 1
            .是否制卡 = Val(GetFieldValueFromRecord("是否制卡", rsTemp, True)) = 1
            .是否发卡 = Val(GetFieldValueFromRecord("是否发卡", rsTemp, True)) = 1 Or .自制卡
            .是否写卡 = Val(GetFieldValueFromRecord("是否写卡", rsTemp, True)) = 1
            
            str读卡性质 = GetFieldValueFromRecord("读卡性质", rsTemp, True)
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .结算性质 = Val(GetFieldValueFromRecord("结算性质", rsTemp, True))
        End With
        If zlCreateBrushObjects(objCard, objBrushCardObject) Then
            Set objCard.CardObject = objBrushCardObject
            mobjXFCardObjs.Add objCard, "K" & objCard.接口编码
        End If
        rsTemp.MoveNext
    Loop
 
    zlInitXFCards = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlCreateBrushObjects(ByVal objCard As Card, ByRef objBrhushCardObject As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建刷卡对象
    '入参:objCard-卡对象
    '出参:objBrhushCardObject-返回创建的对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-31 14:46:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String
    If Not objCard.启用 Then Set objBrhushCardObject = Nothing: Exit Function
    '检查设备是否启用
    If objCard.接口程序名 = "" Then
        '消费卡
        Set objBrhushCardObject = New clsSimulateSquareCard: zlCreateBrushObjects = True: Exit Function
    End If
    
    strCommpentName = objCard.接口程序名 & "." & "cls" & Replace(Replace(UCase(objCard.接口程序名), "ZL9", ""), "ZL", "")
    Err = 0: On Error Resume Next
    Set objBrhushCardObject = CreateObject(strCommpentName)
    If Err <> 0 Then
        ShowMsgbox "部件:" & objCard.接口编码 & "-" & objCard.名称 & "( " & strCommpentName & ")创建失败!" & vbCrLf & "详细的信息为:" & Err.Description
        Call mobjDataBase.WritLog("mdlCardSquare.zlCreateBrushObjects", "", "部件:" & objCard.接口编码 & "-" & objCard.名称 & "创建失败!详细的信息为:" & Err.Description)
        Exit Function
    End If
    zlCreateBrushObjects = True
End Function


Public Function zlGet消费卡接口(Optional ByVal blnOnlyStart As Boolean = True) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡接口
    '入参:blnOnlyStart-是否仅读取启用的消费卡
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim I As Long
    On Error GoTo errHandle
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    If Not mrs消费卡接口 Is Nothing Then
        If mrs消费卡接口.State = 1 Then
            mrs消费卡接口.Filter = 0
            Set zlGet消费卡接口 = mrs消费卡接口
            Exit Function
        End If
    End If
    strWhere = ""
    If blnOnlyStart Then strWhere = " And nvl(A.启用,0)=1"
    
    '用A.*主要是应用于兼容处理.
    If Not zlCheckTableIsExsit("消费卡类别目录", mcnOracle) Then
        strSQL = "" & _
        "   Select  A.* " & _
        "   From 卡消费接口目录 A  " & vbCrLf & _
        "   Where 1=1 " & strWhere & _
        "   Order by 编号 "
    Else
        strSQL = _
        " Select A.*,nvl(b.性质,0) as 结算性质" & vbNewLine & _
        " From 消费卡类别目录 A,结算方式 B" & vbNewLine & _
        " Where A.结算方式=b.名称(+) " & strWhere & _
        " Order By 编号"
    '        strSQL = _
    '        " Select A.编号, A.名称, A.结算方式, Nvl(A.自制卡, 0) As 自制卡, A.前缀文本, A.卡号长度, A.启用," & vbNewLine & _
    '        "        Nvl(A.是否退现, 0) As 是否退现, Nvl(A.是否全退, 0) As 是否全退," & vbNewLine & _
    '        "        Nvl(A.密码长度, 10) As 密码长度, Nvl(A.密码长度限制, 0) As 密码长度限制, Nvl(A.密码规则, 0) As 密码规则," & vbNewLine & _
    '        "        A.部件, A.系统, A.是否密文, 0 As 密码输入限制, 0 As 是否缺省密码," & vbNewLine & _
    '        "        0 As 是否模糊查找, 0 As 是否制卡, 1 As 是否发卡, 0 As 是否写卡, Nvl(A.读卡性质, '1000') As 读卡性质, nvl(A.键盘控制方式,0) As 键盘控制方式," & vbNewLine & _
    '        "        Nvl(A.是否严格控制, 0) As 是否严格控制, A.限制类别, A.应用场合, Nvl(A.是否特定病人, 0) As 是否特定病人," & vbNewLine & _
    '        "        Nvl(A.是否允许换卡, 0) As 是否允许换卡, Nvl(A.是否允许补卡, 0) As 是否允许补卡," & vbNewLine & _
    '        "        Nvl(A.是否允许余额退款, 0) As 是否允许余额退款,nvl(b.性质,0) as 结算性质" & vbNewLine & _
    '        " From 消费卡类别目录 A,结算方式 B" & vbNewLine & _
    '        " Where A.结算方式=b.名称(+) " & strWhere & _
    '        " Order By 编号"
    End If
    Set mrs消费卡接口 = mobjDataBase.OpenSQLRecord(strSQL, "获取消费卡接口 ")
    
    mstrXFCardColumns = ""
    For I = 0 To mrs消费卡接口.fields.count - 1
        mstrXFCardColumns = mstrXFCardColumns & "," & UCase(mrs消费卡接口.fields(I).Name)
    Next
    If mstrXFCardColumns <> "" Then mstrXFCardColumns = Mid(mstrXFCardColumns, 2)
    
    
    Set zlGet消费卡接口 = mrs消费卡接口
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet_SquareRecordFromCardNo(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByRef rsRecord_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否正确
    '入参:lngCardTypeID-接口编号
    '     strCardNo-消费卡卡号
    '出参;rsRecord_Out-返回记录集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-04 14:20:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
     
    If Not zlCheckTableIsExsit("消费卡类别目录", mcnOracle) Then
        strSQL = "" & _
        "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期,  a.密码," & _
        "          to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
        "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态, " & _
        "          to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额 ," & _
        "          to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额 ," & _
        "          to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率 ," & _
        "          to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额 ," & _
        "          to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
        "          a.限制类别 ,A.密码" & _
        "   From 消费卡目录 A  " & _
        "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)  " & _
        "   Order by a.序号"
    Else
        strSQL = _
        "Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期," & vbNewLine & _
        "       to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间," & vbNewLine & _
        "       decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态," & vbNewLine & _
        "       to_char(a.卡面金额," & gOraFmtString.FM_金额 & ") as 卡面金额," & vbNewLine & _
        "       to_char(a.销售金额," & gOraFmtString.FM_金额 & ") as 销售金额," & vbNewLine & _
        "       to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ") as 充值折扣率," & vbNewLine & _
        "       to_char(a.余额," & gOraFmtString.FM_金额 & ") as 余额," & vbNewLine & _
        "       to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & vbNewLine & _
        "       a.限制类别 ,A.密码, b.应用场合, b.是否特定病人, a.病人ID" & vbNewLine & _
        "From 消费卡信息 A, 消费卡类别目录 B" & vbNewLine & _
        "Where a.接口编号 = b.编号 And A.卡号 = [1] and A.接口编号=[2]" & vbNewLine & _
        "      And 序号 = (Select Max(序号) From 消费卡信息 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)" & vbNewLine & _
        "Order by a.序号"
    End If
    Err = 0: On Error GoTo errHandle:
    Set rsRecord_Out = mobjDataBase.OpenSQLRecord(strSQL, "消费卡检查", strCardNo, lngCardTypeID)
    zlGet_SquareRecordFromCardNo = rsRecord_Out.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function zlGet失效面额(ByVal lng消费卡ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取失效面额
    '返回:失效面额
    '编制:刘兴洪
    '日期:2009-12-23 15:08:04
    '说明：只有在当前时间大于了有效期时才调用该函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblTemp As Double
    
    Err = 0: On Error GoTo errHand:
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    If Not zlCheckTableIsExsit("消费卡类别目录", mcnOracle) Then
    
        strSQL = " " & _
        "  Select Sum(Nvl(失效金额, 0)) As 失效金额 " & _
        "  From ( Select 卡面金额 As 失效金额 " & _
        "         From 消费卡目录 A " & _
        "         Where ID =  [1]  " & _
        "         Union All " & _
        "         Select -1 * Sum(Nvl(A.结算金额, 0)) As 失效金额 " & _
        "         From 病人卡结算记录 A, 消费卡目录 B " & _
        "         Where A.消费卡id = B.ID And A.消费卡id =  [1] And   A.交易时间 <= Nvl(B.有效期, To_Date('3000-01-01', 'yyyy-mm-dd')) " & _
        "       )"
        
        Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取失效额", lng消费卡ID)
        dblTemp = Val(NVL(rsTemp!失效金额))
        If dblTemp < 0 Then dblTemp = 0
        zlGet失效面额 = dblTemp
        Exit Function
    End If
    
    strSQL = _
    " Select b.交易序号, Nvl(b.余额, 0) As 失效金额" & vbNewLine & _
    " From 病人卡结算记录 A, 帐户缴款余额 B" & vbNewLine & _
    " Where a.交易序号 = b.交易序号 And a.消费卡id = b.消费卡id And a.记录性质 = 1 And a.消费卡id = [1]"
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取失效额", lng消费卡ID)
    
    '没有记录说明该卡金额已全部使用完
    If rsTemp.EOF Then zlGet失效面额 = 0:        Exit Function
    
    If Val(NVL(rsTemp!交易序号)) > 0 Then
        '升级后的发卡记录，直接取失效金额
        dblTemp = Val(NVL(rsTemp!失效金额))
    Else
        '升级前的发卡记录，需要统计失效金额
        strSQL = _
            "Select Sum(Nvl(失效金额, 0)) As 失效金额" & vbNewLine & _
            "From (" & vbNewLine & _
            "    Select 卡面金额 As 失效金额 From 消费卡信息 A Where ID = [1] And 有效期 < Sysdate" & vbNewLine & _
            "    Union All" & vbNewLine & _
            "    Select Nvl(Sum(a.应收金额), 0) As 失效金额" & vbNewLine & _
            "    From 病人卡结算记录 A, 消费卡信息 B" & vbNewLine & _
            "    Where a.消费卡id = b.Id And a.记录性质 = 4 And a.消费卡id = [1]" & vbNewLine & _
            "          And a.交易时间 <= Nvl(b.有效期, To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
            "     )"
        Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取失效额", lng消费卡ID)
        dblTemp = Val(NVL(rsTemp!失效金额))
    End If
    
    If dblTemp < 0 Then dblTemp = 0
    zlGet失效面额 = dblTemp
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function zlGetPatiCardFromCards(ByVal strCardTypeIDs As String, ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从给定卡类别中检索指定病人持有有效卡的卡类别
    '入参: strCardTypeIDs 给定卡类别，多个用逗号分隔
    '返回:返回病人持有有效卡的卡类别，多个用逗号分隔
    '编制:刘兴洪
    '日期:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '问题号：113121
    Dim strCards As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If strCardTypeIDs = "" Or lng病人ID = 0 Then Exit Function
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    strSQL = _
        "Select Distinct a.卡类别id" & vbNewLine & _
        "From 病人医疗卡信息 A" & vbNewLine & _
        "Where a.病人id = [1] And Nvl(a.状态, 0) = 0" & vbNewLine & _
        "      And 卡类别id In (Select /*+cardinality(j,10)*/  j.Column_Value From Table(f_Num2list([2])) J)"
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "zlpublicOneCard.clsOneCardObject", lng病人ID, strCardTypeIDs)
    Do While Not rsTemp.EOF
        strCards = strCards & "," & Val(NVL(rsTemp!卡类别id))
        rsTemp.MoveNext
    Loop
    If strCards <> "" Then strCards = Mid(strCards, 2)
    zlGetPatiCardFromCards = strCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet收费类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    Dim strSQL As String
    On Error GoTo errHandle
    
    If Not mrs收费类别 Is Nothing Then
        If mrs收费类别.State = 1 Then
            mrs收费类别.Filter = 0
            Set zlGet收费类别 = mrs收费类别
            Exit Function
        End If
    End If
    strSQL = "Select 编码,名称 From 收费项目类别"
    Set mrs收费类别 = mobjDataBase.OpenSQLRecord(strSQL, "获取收费类别")
    Set zlGet收费类别 = mrs收费类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetFeeTypeFromNameToCode(ByVal str限制类别 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据限制类别名称取相关的编码
    '入参:str限制类别-收费类别名称,多个用逗号分离:如：西药，草药，检查...
    '返回:转换为编码集,多个用逗号,如:7,8,4...
    '编制:刘兴洪
    '日期:2009-12-23 16:31:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTemp = zlGet收费类别
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    If str限制类别 = "" Then zlGetFeeTypeFromNameToCode = "": Exit Function
    str限制类别 = "," & str限制类别 & ","
    With rsTemp
        Do While Not .EOF
            str限制类别 = Replace(str限制类别, "," & NVL(rsTemp!名称) & ",", "," & NVL(rsTemp!编码) & ",")
            .MoveNext
        Loop
    End With
    zlGetFeeTypeFromNameToCode = str限制类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetCojumnsFromTableName(ByVal strTableName As String, ByRef strColumns_out As String, _
    Optional strAliaName As String, Optional ByVal strNotColumns As String, Optional ByVal strOlnyHaveColumns As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据表名获取其对应的列
    '入参:strTableName-获取的表名
    '     strAliaName-别名
    '     strNotColumns-不包含的列，多个用逗号
    '     strOlnyHaveColumns-仅有的例,没有时，返回NULL值作为列
    '出参:strColumns_Out-返回列数，多个用逗号分离
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 15:39:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strColumn As String
    Dim rsTemp As ADODB.Recordset, blnAdd As Boolean
    Dim blnTemp As Boolean, varTemp As Variant
    Dim I As Long
    
    On Error GoTo errHandle
    
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
  
    strSQL = "" & _
    " Select Column_name,Column_ID  " & _
    " From all_tab_columns " & _
    " Where Table_name=[1] And Owner=(Select Max(Owner) From all_tab_columns Where  Table_name=[1]) " & _
    " Order by Column_ID "
    
    Set rsTemp = mobjDataBase.OpenSQLRecord(strSQL, "获取指定表的列信息", strTableName)
    If rsTemp.EOF Then Set rsTemp = Nothing: Exit Function
    
    strNotColumns = UCase(strNotColumns)
    strOlnyHaveColumns = UCase(strOlnyHaveColumns)
    blnAdd = strOlnyHaveColumns = ""
    With rsTemp
        strColumns_out = ""
        Do While Not .EOF
            
            strColumn = UCase(NVL(rsTemp!Column_name))
            If InStr("," & strNotColumns & ",", "," & strColumn & ",") = 0 Then
                blnTemp = blnAdd
                If Not blnTemp Then blnTemp = InStr("," & strOlnyHaveColumns & ",", "," & strColumn & ",") > 0
                If blnTemp Then
                    strColumns_out = strColumns_out & "," & IIf(strAliaName = "", "", strAliaName & ".") & strColumn
                End If
            End If
            .MoveNext
        Loop
    End With
    If strOlnyHaveColumns <> "" Then
        varTemp = Split(strOlnyHaveColumns & ",", ",")
        For I = 0 To UBound(varTemp)
            If Trim(varTemp(I)) <> "" Then
                If InStr(strColumns_out & ",", "," & strAliaName & "." & varTemp(I) & ",") = 0 Then
                    strColumns_out = strColumns_out & ",NULL as " & varTemp(I)
                End If
            End If
        Next
    End If
    If strColumns_out <> "" Then strColumns_out = Mid(strColumns_out, 2)
    zlGetCojumnsFromTableName = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetCardObj(ByVal frmMain As Object, ByVal lngCardTypeID As Long, _
    Optional bln消费卡 As Boolean = False, _
    Optional ByRef objPatiCardObj As Card, _
    Optional ByRef blnNotParaCreateObject As Boolean = False, _
    Optional ByVal blnNotStartCreateObject As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡结算对象
    '入参:lngCardTypeID-指定卡类别
    '       bln消费卡-是否消费卡
    '       blnNotParaCreateObject-不根据参数创建对象
    '       blnNotStartCreateObject-为true时，未设置启用的也要创建接口对象, _
    '                                                  为False时, 只有设置了启用才创建对象
    '返回:函数返回True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2011-05-23 15:29:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCardObject As Object, strKey As String, strExpand As String
    Dim blnOnlyNotObject As Boolean
    Dim objYLCards As Cards
    Dim objYlCardObjs As Cards
    
    If Not objPatiCardObj Is Nothing Then
        If objPatiCardObj.接口序号 = lngCardTypeID And objPatiCardObj.消费卡 = bln消费卡 Then
            If Not objPatiCardObj.InitCompents Then
                If objPatiCardObj.CardObject Is Nothing Then
                    blnOnlyNotObject = True: GoTo GoCreateObject:
                End If
                If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then                    '初始化部件
                    Exit Function
                End If
                objPatiCardObj.InitCompents = True
            End If
            zlGetCardObj = True
            Exit Function
        End If
    End If
    Err = 0: On Error Resume Next
GoCreateObject:
    strKey = IIf(bln消费卡, "X", "K") & lngCardTypeID
    '59760
    '检查设备是否启用
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Set objPatiCardObj = objYlCardObjs(strKey)
    If Err <> 0 Or _
        blnNotParaCreateObject And objPatiCardObj.CardObject Is Nothing _
        Or blnNotStartCreateObject Then
          '错误或不根据参数创建对象时,才需要重新创建对象
            Err = 0: On Error Resume Next
            Set objCard = objYLCards.Item(strKey)
            If Err <> 0 Then
                ShowMsgbox "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!"
                Call mobjDataBase.WritLog("zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到或该" & IIf(bln消费卡, "结算卡", "医疗卡类别") & "不存在,请检查!")
                Exit Function
            End If
            '未启用也要创建对象
            If blnNotStartCreateObject Then objCard.启用 = True
            '重新创建
            If zlCreatePatiCardObject(objCard, objCardObject) = False Then Exit Function
            '增加对应
           Set objPatiCardObj = objYlCardObjs.Add(objCardObject, strKey)
    End If
    
    If objPatiCardObj Is Nothing Then
        If Not objCard Is Nothing Then
                MsgBox "注意:" & vbCrLf & "调用接口(" & objCard.接口编码 & "-" & objCard.名称 & ")调用失败,请检查!", vbInformation, gstrSysName
        Else
                MsgBox "注意:" & vbCrLf & "调用接口(" & lngCardTypeID & ")调用失败,请检查!", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    Err = 0: On Error Resume Next
    Set objCard = objPatiCardObj
    If Err <> 0 Then
        ShowMsgbox "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description
        Call mobjDataBase.WritLog("clsPatiCard.zlInitInterFacel", "", "部件:" & lngCardTypeID & "未找到,请检查!" & vbCrLf & " 详细的错误信息:" & Err.Description)
        Exit Function
    End If
    If Not objPatiCardObj.InitCompents Then
        If Not objPatiCardObj.CardObject.zlInitComponents(frmMain, glngModul, glngSys, gstrDBUser, gcnOracle, False, strExpand) Then Exit Function
         objPatiCardObj.InitCompents = True
    End If
    zlGetCardObj = True
End Function

Private Function GetFieldValueFromRecord(ByVal strFieldName As String, ByRef rsTemp As ADODB.Recordset, Optional bln消费卡 As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据记录集取字段的值
    '入参:strFieldName-字段名
    '     rsTemp-取值的记录集
    '    bln消费卡-是否消费卡
    '返回:返回具体值
    '编制:刘兴洪
    '日期:2018-12-04 11:08:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.EOF Then Exit Function
    If bln消费卡 = False Then
        If InStr("," & mstrYLCardColumns & ",", "," & strFieldName & ",") = 0 Then Exit Function
        GetFieldValueFromRecord = NVL(rsTemp.fields(strFieldName))
        Exit Function
    End If
    If InStr("," & mstrXFCardColumns & ",", "," & strFieldName & ",") = 0 Then Exit Function
    GetFieldValueFromRecord = NVL(rsTemp.fields(strFieldName))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlInitPatiCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡集
    '编制:刘兴洪
    '日期:2011-05-23 17:54:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, bln启用 As Boolean, str部件 As String, objCard As Card
    Dim int自动读取间隔 As Integer, str读卡性质 As String
    Dim objCardObject As Object
    
    Err = 0: On Error GoTo errHand:
    
    Set mobjYLCards = New Cards: Set mobjYLCardObjs = New Cards
    
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    Set rsTemp = zlGet医疗卡类别()
    If rsTemp Is Nothing Then Exit Function
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
    
        ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
        int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "自动读取", "0"))
        int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "自动读取间隔", "300"))
        bln启用 = Val(NVL(rsTemp!是否启用)) = 1
        
        '90875,李南春,2016/1/22:证件类型都启用
        If bln启用 Then
            If Val(GetFieldValueFromRecord("是否自制", rsTemp)) = 1 Or Val(GetFieldValueFromRecord("是否证件", rsTemp)) = 1 Then        '自制卡,都启用
                bln启用 = True
            Else
                '问题号:54098
                If (NVL(rsTemp!名称) Like "*身份证*" Or NVL(rsTemp!名称) Like "*IC卡*") And Val(NVL(rsTemp!是否固定)) = 1 And NVL(rsTemp!部件) = "" Then
                    bln启用 = True
                Else
                    bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\医疗卡\" & NVL(rsTemp!编码), "启用", "0")) = 1
                End If
            End If
        End If
        
        
        str部件 = Trim(NVL(rsTemp!部件))
        'ID,编码,名称,短名,前缀文本,卡号长度,缺省标志,是否固定,是否严格控制,是否刷卡,是否自制,是否存在帐户,是否全退,部件,备注,特定项目,结算方式,是否启用
        '77872,李南春,2014/10/28:是否支持转帐及代扣
        '85565,李南春,2015/7/10:读卡性质
        '90875,李南春,2016/1/22:是否证件
        '103310,李南春,2016/12/7:卡号后增加回车符位
        Set objCard = New Card
        With objCard
            .卡种类 = EM_CardType_Square
            .接口序号 = NVL(rsTemp!id)
            .接口编码 = NVL(rsTemp!编码)
            .短名 = NVL(rsTemp!短名)
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度)) + Val(GetFieldValueFromRecord("设备是否启用回车", rsTemp))
            .缺省标志 = Val(NVL(rsTemp!缺省标志)) = 1
            .系统 = Val(NVL(rsTemp!是否固定)) = 1
            .是否严格控制 = Val(NVL(rsTemp!是否严格控制)) = 1
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(GetFieldValueFromRecord("是否自制", rsTemp)) = 1
            .是否存在帐户 = Val(GetFieldValueFromRecord("是否存在帐户", rsTemp)) = 1
            .是否全退 = Val(NVL(rsTemp!是否全退)) = 1
            .卡号重复使用 = Val(NVL(rsTemp!是否重复使用)) = 1
            .结算方式 = NVL(rsTemp!结算方式)
            .接口程序名 = NVL(rsTemp!部件)
            .特定项目 = NVL(rsTemp!特定项目)
            .启用 = bln启用
            .备注 = NVL(rsTemp!备注)
            .卡号密文规则 = NVL(rsTemp!卡号密文)
            .是否退现 = Val(GetFieldValueFromRecord("是否退现", rsTemp)) = 1
            .密码长度 = Val(GetFieldValueFromRecord("密码长度", rsTemp))
            .密码长度限制 = Val(GetFieldValueFromRecord("密码长度限制", rsTemp))
            .密码规则 = Val(GetFieldValueFromRecord("密码规则", rsTemp))
            .密码输入限制 = Val(GetFieldValueFromRecord("密码输入限制", rsTemp))
            .是否缺省密码 = Val(GetFieldValueFromRecord("是否缺省密码", rsTemp))
            .是否制卡 = Val(GetFieldValueFromRecord("是否制卡", rsTemp)) = 1     '56615
            .是否发卡 = Val(GetFieldValueFromRecord("是否发卡", rsTemp)) = 1 Or .自制卡
            .是否写卡 = Val(GetFieldValueFromRecord("是否写卡", rsTemp)) = 1
            .是否模糊查找 = Val(GetFieldValueFromRecord("是否模糊查找", rsTemp)) = 1
            .是否转帐及代扣 = Val(GetFieldValueFromRecord("是否转帐及代扣", rsTemp)) = 1
            str读卡性质 = Trim(GetFieldValueFromRecord("读卡性质", rsTemp))
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            
            .是否证件 = Val(GetFieldValueFromRecord("是否证件", rsTemp)) = 1
            .是否持卡消费 = Val(GetFieldValueFromRecord("是否持卡消费", rsTemp)) = 1
            .是否发送调用接口 = Val(GetFieldValueFromRecord("发送调用接口", rsTemp)) = 1
            .是否退款验卡 = Val(GetFieldValueFromRecord("是否退款验卡", rsTemp)) = 1
            .设备是否启用回车 = Val(GetFieldValueFromRecord("设备是否启用回车", rsTemp)) = 1
            .是否缺省退现 = Val(GetFieldValueFromRecord("是否缺省退现", rsTemp)) = 1
            .结算性质 = Val(GetFieldValueFromRecord("结算性质", rsTemp))
            .是否独立结算 = Val(GetFieldValueFromRecord("是否独立结算", rsTemp)) = 1
            .是否支持扫码付 = Val(GetFieldValueFromRecord("是否支持扫码付", rsTemp)) = 1
            .卡号识别规则 = Val(GetFieldValueFromRecord("卡号识别规则", rsTemp))
        End With
        
        mobjYLCards.Add objCard, "K" & objCard.接口序号
        If zlCreatePatiCardObject(objCard, objCardObject) Then
            Set objCard.CardObject = objCardObject
            mobjYLCardObjs.Add objCard, "K" & objCard.接口序号
        End If
        rsTemp.MoveNext
    Loop
 
    
    Set rsTemp = zlGet消费卡接口()

    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
        int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编号), "自动读取", "0"))
        int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编号), "自动读取间隔", "300"))
        bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编号), "启用", "0")) = 1
        
        '编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度,部件,系统,是否密文
        str部件 = Trim(NVL(rsTemp!部件))
        Set objCard = New Card
        With objCard
            .卡种类 = EM_CardType_Consume
            .接口序号 = NVL(rsTemp!编号)
            .接口编码 = NVL(rsTemp!编号)
            .短名 = Left(NVL(rsTemp!名称), 1)   '默认取第一个
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度))
            .系统 = Val(NVL(rsTemp!系统)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(GetFieldValueFromRecord("自制卡", rsTemp, True)) = 1
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(GetFieldValueFromRecord("是否全退", rsTemp, True)) = 1
            .结算方式 = Trim(GetFieldValueFromRecord("结算方式", rsTemp, True))
            .接口程序名 = Trim(GetFieldValueFromRecord("部件", rsTemp, True))
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = Trim(GetFieldValueFromRecord("是否密文", rsTemp, True))
            .消费卡 = True
            .是否退现 = Val(GetFieldValueFromRecord("是否退现", rsTemp, True)) = 1
            .密码长度 = Val(GetFieldValueFromRecord("密码长度", rsTemp, True))
            .密码长度限制 = Val(GetFieldValueFromRecord("密码长度限制", rsTemp, True))
            .密码规则 = Val(GetFieldValueFromRecord("密码规则", rsTemp, True))
            .密码输入限制 = Val(GetFieldValueFromRecord("密码输入限制", rsTemp, True))
            .是否缺省密码 = Val(GetFieldValueFromRecord("是否缺省密码", rsTemp, True)) = 1
            .是否制卡 = Val(GetFieldValueFromRecord("是否制卡", rsTemp, True)) = 1
            .是否发卡 = Val(GetFieldValueFromRecord("是否发卡", rsTemp, True)) = 1 Or .自制卡
            .是否写卡 = Val(GetFieldValueFromRecord("是否写卡", rsTemp, True)) = 1
            
            str读卡性质 = GetFieldValueFromRecord("读卡性质", rsTemp, True)
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .结算性质 = Val(GetFieldValueFromRecord("结算性质", rsTemp, True))
        End With
        mobjYLCards.Add objCard, "X" & objCard.接口序号
        If zlCreatePatiCardObject(objCard, objCardObject) Then
            Set objCard.CardObject = objCardObject
            mobjYLCardObjs.Add objCard, "X" & objCard.接口序号
        End If
        rsTemp.MoveNext
    Loop
    zlInitPatiCards = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlCreatePatiCardObject(ByVal objCard As Card, ByRef objCardObject As Object, Optional blnAdviceSend As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定卡的对象
    '出参:objCardObject-被创建的对象
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-25 10:47:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String, strHead As String
    If Not objCard.启用 And Not blnAdviceSend Then
        Set objCardObject = Nothing: Exit Function
    End If
    
    '检查设备是否启用
    strHead = IIf(objCard.消费卡, "", "zl9Card_")
    If objCard.接口程序名 = "" Then
        '99858:李南春,2016/9/2,三方账户医疗卡、非自制消费卡必须有接口部件
        If objCard.消费卡 And objCard.自制卡 Then
            Set objCardObject = New clsSimulateSquareCard: zlCreatePatiCardObject = True: Exit Function
        ElseIf Not objCard.是否存在帐户 Then
            Set objCardObject = New clsOwnerCardObject: zlCreatePatiCardObject = True: Exit Function
        End If
        MsgBox objCard.名称 & "未设置接口部件，请在" & IIf(objCard.消费卡, "【消费卡管理】", "【医疗卡类别管理】") & "中设置部件名!"
        Exit Function
    End If
    strCommpentName = zlGetFullComponentNameFromCard(objCard, strHead)
    Err = 0: On Error Resume Next
    Set objCardObject = CreateObject(strCommpentName)
    If Err <> 0 Then
        ShowMsgbox "部件:" & objCard.接口编码 & "-" & objCard.名称 & "( " & strCommpentName & ")创建失败,请与系统管理员联系!" & vbCrLf & "详细的信息为:" & Err.Description
        Call mobjDataBase.WritLog("mdlCardSquare.zlCreatePatiCardObject", "", "部件:" & objCard.接口编码 & "-" & objCard.名称 & "创建失败!详细的信息为:" & Err.Description)
        Err = 0: On Error GoTo 0
        Exit Function
    End If
    zlCreatePatiCardObject = True
End Function

Public Function zlGetFullComponentNameFromCard(ByVal objCard As Card, Optional strHead As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡对象，获取完成的部件名称
    '入参:objCard- 当前卡对象
    '     strHead-部件头
    '返回:返回完整的部件名称
    '编制:刘兴洪
    '日期:2011-06-22 13:57:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str部件 As String
    On Error GoTo errHandle
    
    str部件 = objCard.接口程序名
    If str部件 = "" Then zlGetFullComponentNameFromCard = "": Exit Function
    If strHead <> "" Then
        '有开头部分,需要检查传入是否包含这部分
        If str部件 Like strHead & "*" Then
            str部件 = str部件 & "." & "cls" & Mid(str部件, Len(strHead) + 1)
        Else
            str部件 = strHead & str部件 & "." & "cls" & Replace(Replace(UCase(str部件), "ZL9", ""), "ZL", "")
        End If
    Else
        str部件 = str部件 & "." & "cls" & Replace(Replace(UCase(str部件), "ZL9", ""), "ZL", "")
    End If
    zlGetFullComponentNameFromCard = str部件
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetComponentObject(ByVal lng卡类别ID As Long, _
     Optional bln消费卡 As Boolean = False) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的卡对象
    '入参:lng卡类别ID-卡类别ID
    '        bln消费卡
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-06-25 23:52:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Dim objYlCardObjs As Cards
    strKey = IIf(bln消费卡, "X", "K") & lng卡类别ID
    '59760
    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    
    Err = 0: On Error Resume Next
    Set zlGetComponentObject = objYlCardObjs(strKey).CardObject
    If Err <> 0 Then
        Err.Clear: On Error GoTo 0
    End If
End Function
Public Function zlGetYLCardObjs(ByRef objYlCardObjects As Cards) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡对象
    '出参:objYlCardObjects-返回卡对象
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 13:59:24
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjYLCardObjs Is Nothing Then
        Set objYlCardObjects = mobjYLCardObjs
        zlGetYLCardObjs = True
        Exit Function
    End If
    If zlInitPatiCards = False Then Exit Function
    Set objYlCardObjects = mobjYLCardObjs
    zlGetYLCardObjs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsExistFeeInsurePatient(lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医保病人是否存在未结费用
    '入参:lng病人ID
    '返回:存在未结返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    strSQL = "Select Nvl(sum(B.费用余额),0) 费用余额 From 病人信息 A,病人余额 B Where A.病人ID=B.病人ID And Nvl(A.险类,0)<>0 And A.病人ID=[1]"

    Set rsTmp = mobjDataBase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If Not rsTmp.EOF Then zlIsExistFeeInsurePatient = (rsTmp!费用余额 <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetUserUnits(Optional ByVal blnDept As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：获取当前用户的所有部门ID或病区ID
    '      如果操作员属于科室,则返回科室ID及科室所属病区ID
    '      blnDept:True表示获取操作员所属科室,以及所属病区下的所有科室,否则返回操作员所属病区,以及所在科室所属的病区
    '返回:成功返回部门ID或病区ID,多个用逗号分离
    '编制:刘兴洪
    '日期:2018-12-05 16:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, I As Integer
    Dim strUnitIDs As String
    On Error GoTo errH
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    'union会去掉重复的
    If blnDept Then
        strSQL = "Select A.科室ID 部门ID From 病区科室对应 A,部门人员 B Where A.病区ID=B.部门ID And B.人员ID=[1]" & _
            " Union Select 部门ID as 部门ID From 部门人员 Where 人员ID=[1]"
    Else
        strSQL = "Select A.病区ID 部门ID From 病区科室对应 A,部门人员 B Where A.科室ID=B.部门ID And B.人员ID=[1]" & _
            " Union Select 部门ID as 部门ID From 部门人员 Where 人员ID=[1]"
    End If
    
    Set rsTmp = mobjDataBase.OpenSQLRecord(strSQL, App.ProductName, UserInfo.id)
    
    For I = 1 To rsTmp.RecordCount
        strUnitIDs = strUnitIDs & "," & rsTmp!部门ID
        rsTmp.MoveNext
    Next
    
    If strUnitIDs = "" Then
        strUnitIDs = "0"
    Else
        strUnitIDs = Mid(strUnitIDs, 2)
    End If
    zlGetUserUnits = strUnitIDs
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetUnitRecordFromDepdIDs(ByVal blnLimitUnit As Boolean, _
    ByVal strServiceRange As String, ByVal strType As String, _
    Optional bln简码 As Boolean = False, _
    Optional blnNotNode As Boolean = False, _
    Optional blnShowNodeCode As Boolean = False) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有病区或科室列表
    '入参:blnLimitUnit=是否有所有病区权限，没有时，只获取操作员所属的科室或病区
    '       blnNotNode-是否区分站点:true,不区分站点,区分站点
    '       blnShowNodeCode:显示站点编号
    '出参:
    '返回:病区或科室数据集
    '编制:刘兴洪
    '日期:2011-02-28 17:21:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strUnitIDs As String
    Dim strWhere As String
    
    On Error GoTo errH
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    If blnLimitUnit Then strUnitIDs = zlGetUserUnits
    strWhere = ""
    If blnNotNode = False Then strWhere = " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null) "
    strSQL = _
         " Select A.ID,A.编码,A.名称 " & IIf(bln简码, ",A.简码", "") & IIf(blnShowNodeCode, ",A.站点", "") & _
         " From 部门表 A,部门性质说明 B" & _
         " Where B.部门ID = A.ID And B.服务对象 IN(" & strServiceRange & ") And B.工作性质 = [2]" & _
         " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            strWhere & vbNewLine & _
         IIf(blnLimitUnit, " And Instr(','||[1]||',',','||A.ID||',')>0", "") & _
         " Order by A.编码"
    Set zlGetUnitRecordFromDepdIDs = mobjDataBase.OpenSQLRecord(strSQL, App.ProductName, strUnitIDs, strType)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlGetCards_YL(ByRef objCards As Cards) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医疗卡类别
    '出参:objCards-医疗卡类别对象
    '返回:返回成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-23 12:03:26
    '说明:59760
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjYLCards Is Nothing Then
        Set objCards = mobjYLCards: zlGetCards_YL = True: Exit Function
    End If
    If zlInitPatiCards = False Then Exit Function
    Set objCards = mobjYLCards
    zlGetCards_YL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set mobjXML = New clsXML
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not gfrmCardMgr Is Nothing Then Set gfrmCardMgr = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjXML Is Nothing Then Set mobjXML = Nothing
    Call zlClearCache '清除缓存
End Sub
