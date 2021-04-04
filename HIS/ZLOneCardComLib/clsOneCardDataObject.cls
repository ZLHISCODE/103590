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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'  1.6.  zlGetAvailabilityWriteCardType-获取有效的写卡类别
'  2.病人相关数据
'  2.1 zlGetPatiIDFromBedNumber:根据床位号，获取病人ID
'  2.2 zlGetPatiInforFromPatiID:根据病人ID获取病人信息对象
'  2.3 zlGetPatiInfoFromPatiName-根据病人姓名，获取病人的详细信息
'  2.4 zlGetPatiInforFromXML:根据XML串，获取病人信息对象
'  2.4 zlGetPatiIDFromCardNo-根据卡号来获取病人ID
'编制:刘兴洪
'日期:2018-12-03 11:06:14
'说明:该部件不区分版本
'*********************************************************************************************************************************************
Private mrs收费类别 As ADODB.Recordset
Private mrs手机号 As ADODB.Recordset

Private mobjYLCards As Cards
Private mobjYLCardObjs As Cards   '当前启用有效的医疗卡
Private mobjXFCards As Cards
Private mobjXFCardObjs As Cards

Private gfrmCardMgr As Object
Private mrsSystem As ADODB.Recordset
Private mobjDataBase As clsDataBase
Private mcnOracle As ADODB.Connection
Private mstrXFCardColumns As String '多个用逗号分离
Private mstrYLCardColumns As String '多个用逗号分离
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
   ErrCenter = mobjDataBase.ErrCenter
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
    Set mobjYLCards = Nothing: Set mobjYLCardObjs = Nothing
    Set mrs收费类别 = Nothing: Set mrsSystem = Nothing
    Set mrs手机号 = Nothing: Set mobjXFCards = Nothing
    Set mobjXFCardObjs = Nothing
End Function

Public Function zlGetPatiIDFromBedNumber(ByVal lng病区ID As Long, ByVal str床号 As String, _
     Optional ByVal str住院号 As String, Optional ByRef lng主页Id As Long) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据床号获取病人ID
    '出参:lng主页ID-返回当前床号的主页ID
    '返回:成功返回病人ID,否则返回False
    '编制:刘兴洪
    '日期:2012-09-19 15:50:18
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    On Error GoTo errHandle
    Set cllFilter = New Collection
    If str住院号 <> "" Then
        cllFilter.Add Array("住院号", str住院号)
    Else
        cllFilter.Add Array("病区ID", lng病区ID)
        cllFilter.Add Array("床号", str床号)
    End If
    zlGetPatiIDFromBedNumber = zl_CisSvr_GetPatiID(cllFilter, lng主页Id)
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


Public Function zlGetArea(frmParent As Object, txtInput As Object, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo ErrH
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
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
   Call gobjComLib.SaveErrLog
End Function
Private Function GetCardFindPati(ByVal strCardNo As String, Optional ByVal blnNotShowErrMsg As Boolean = False, _
    Optional ByRef lng病人ID As Long, Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional ByVal strRQCode As String = "", _
    Optional ByRef intCardStatus As Integer, _
    Optional ByVal bln检查使用时间 As Boolean = True, Optional ByVal bln检查停用或挂失 As Boolean = True, _
    Optional ByRef strValidTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号，模糊查找病人
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       strRQCode-二维码
    '出参:strErrMsg-返回的错误信息
    '        lng病人ID-返回的病人ID
    '        strCardPass-返回卡号的密码
    '        lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 09:36:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim cllData As Collection, cllTemp    As Collection
    Dim lng病人ID1 As Long, lngTemp As Long
    Dim str卡类别ID As String, str病人ID As String
    Dim cllFindCons As Collection
    
    On Error GoTo errHandle
    
    Set cllFindCons = New Collection
    
    If strCardNo <> "" Then cllFindCons.Add Array("卡号", strCardNo)
    If strRQCode <> "" Then cllFindCons.Add Array("二维码", strCardNo)
    
    If zl_PatiSvr_GetPatiID(cllFindCons, Nothing, cllData, blnNotShowErrMsg, strErrMsg, bln检查使用时间, bln检查停用或挂失, intCardStatus) = False Then Exit Function
    'output
    '    code    N   1   应答码：0-失败；1-成功
    '    message C   1   应答消息： 失败时返回具体的错误信息
    '    pati_list[] C   1   病人列表,模糊查找时，可能存在多个
    '        cardtype_id N   1   卡类别ID
    '        pati_id N   1   病人ID:未找到时也成功，返回0
    '        card_pwd    C   1   密码
    '        pati_pageid N       主页ID
    '        enduse_time C   1   终止使用时间:yyyy-mm-dd hh24mi:ss

    If cllData.count = 0 Then
        If strRQCode <> "" Then
            strErrMsg = "未找到符合条件的病人信息，请检查！"
        Else
            strErrMsg = "未找到卡号为" & strCardNo & "的病人，请检查该卡是否有效卡！"
        End If
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If cllData.count = 1 Then
        Set cllTemp = cllData(1)
        '找到一个，直接返回
        lng病人ID = Val(cllTemp("_pati_id"))
        strCardPassWord = NVL(cllTemp("_card_pwd"))
        lngCardTypeID = Val(NVL(cllTemp("_cardtype_id")))
        intCardStatus = Val(NVL(cllTemp("_card_status")))
        strValidTime = NVL(cllTemp("_enduse_time"))
        GetCardFindPati = True: Exit Function
    End If
    
    '多个病人ID
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        lngTemp = Val(cllTemp("_cardtype_id"))
        If lngTemp <> 0 Then
            lng病人ID1 = Val(cllTemp("_pati_id"))
            If InStr(1, str卡类别ID & ",", "," & lngTemp & ",") = 0 Then str卡类别ID = str卡类别ID & "," & lngTemp
            If InStr(1, str病人ID & ",", "," & lng病人ID1 & ",") = 0 Then str病人ID = str病人ID & "," & lng病人ID1
        End If
    Next
 
    If frmSelectType.zlSelect(Nothing, str卡类别ID, lngCardTypeID) = False Then lngCardTypeID = 0: Exit Function
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        lngTemp = Val(cllTemp("_cardtype_id"))
        If lngTemp = lngCardTypeID Then
            lng病人ID = Val(NVL(cllTemp("_pati_id")))
            strCardPassWord = NVL(cllTemp("_card_pwd"))
            lngCardTypeID = Val(NVL(cllTemp("_cardtype_id")))
            intCardStatus = Val(NVL(cllTemp("_card_status")))
            strValidTime = NVL(cllTemp("_enduse_time"))
            GetCardFindPati = True: Exit Function
        End If
    Next
    lngCardTypeID = 0
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
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
            If objCard.名称 = Trim(strCardType) And objCard.消费卡 = bln消费卡 Then
                 Set objCard_Out = objCard: zlGetCardFromCardTypeID = True: Exit Function
            End If
        Else
            If objCard.接口序号 = Val(strCardType) And objCard.消费卡 = bln消费卡 Then
                Set objCard_Out = objCard: zlGetCardFromCardTypeID = True: Exit Function
            End If
        End If
    Next
    zlGetCardFromCardTypeID = False
    Set objCard_Out = Nothing
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function
Public Function zlGetInputType(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : zlGetInputType
    ' 参数 : lngCardTypeID-卡类别ID
    ' 日期 : 2018-11-05
    ' 说明 : 返回true-转换为大写，false-不转换
    '---------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo ErrH
    If zlGetCardFromCardTypeID(lngCardTypeID, False, objCard) = False Then Exit Function
    If objCard.启用 = False Then Exit Function
    zlGetInputType = objCard.卡号识别规则 = 1
    Exit Function
ErrH:
    If ErrCenter() Then Resume
End Function
Public Function zlGetPatiIDFromInpatientNum(ByVal str住院号 As String, ByRef lng病人ID_Out As Long, _
        Optional ByVal blnNotShowMsg As Boolean, Optional ByRef strErrMsg_Out As String, _
        Optional lng主页ID_out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据住院号获取病人信息
    '入参:str住院号- 住院号
    '     blnNotShowMsg-是否不显示消息框
    '出参:strErrMsg_Out-返回的错误信息
    '     lng病人Id_out-返回的病人ID
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-29 21:12:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    
    On Error GoTo errHandle
    Set cllFilter = New Collection
    cllFilter.Add Array("住院号", str住院号)
    lng病人ID_Out = zl_CisSvr_GetPatiID(cllFilter, lng主页ID_out)
    If lng病人ID_Out = 0 Then
        strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        If Not blnNotShowMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlGetPatiIDFromInpatientNum = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
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
    Optional ByVal blnNotCheckValidDate As Boolean = False, _
    Optional ByVal strRQCode As String, _
    Optional ByRef intCardStatus As Integer, _
    Optional ByRef strValidTime As String, _
    Optional ByVal bytPatiInfoShowType As Byte) As Boolean
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
    '       bytPatiInfoShowType-病区选择器列显示方式：0-所有信息，1-简略信息
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '       strValidTime-当前刷卡的有效时间：如果Isdate(strValidTime)=false,则代表是长期有效卡
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, i As Long, J As Long
    Dim cllData  As Collection, cllTemp  As Collection, cllItems As Collection
    Dim rsTemp As ADODB.Recordset
    Dim str病人Ids As String, str类别名称 As String
    Dim frmSel As frmPatiSelect
    Dim blnCard As Boolean
    Dim cllFind As Collection, cllFilter As Collection
    Dim rsSeled As ADODB.Recordset
    
    On Error GoTo errHandle
    strCardPassWord = "": strErrMsg = ""
    lng病人ID = 0
    If strCardType = "" Then Exit Function
    If Val(strCardType) = -1 Then
       zlGetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID, strRQCode, intCardStatus, , , strValidTime)
       Exit Function
    End If
    
    str类别名称 = ""
    strCardType = Replace(strCardType, " ", "")
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
                     Exit Do
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If Not IsNumeric(strCardType) Then
        Set rsTemp = zlGet医疗卡类别
        rsTemp.Filter = "名称 ='" & strCardType & "'"
        If Not rsTemp.EOF Then
            If Val(NVL(rsTemp!是否启用)) = 0 Then
                If blnNotShowErrMsg Then
                    strErrMsg = "卡类别为" & strCardType & "的医疗卡未启用，无法获取有效的病人信息!"
                Else
                     MsgBox "卡类别为" & strCardType & "的医疗卡未启用，无法获取有效的病人信息!", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strCardType = Val(NVL(rsTemp!id))
            
        End If
        rsTemp.Filter = 0
    End If
    
    If IsNumeric(strCardType) Then  '以卡类别ID为准
        
        If zlGetInputType(Val(strCardType)) Then
            strCardNo = UCase(strCardNo)
        End If
        
        Set cllFind = New Collection
        cllFind.Add Array("卡类别ID", Val(strCardType))
        cllFind.Add Array("卡号", strCardNo)
        If strRQCode <> "" Then cllFind.Add Array("二维码", strRQCode)
        If zl_PatiSvr_GetPatiID(cllFind, Nothing, cllData, blnNotShowErrMsg, strErrMsg, Not blnNotCheckValidDate, True, _
            intCardStatus, True) = False Then Exit Function
        
        '    pati_list[] C   1   病人列表,模糊查找时，可能存在多个
        '        cardtype_id N   1   卡类别ID
        '        pati_id N   1   病人ID:未找到时也成功，返回0
        '        card_pwd    C   1   密码
        '        pati_pageid N       主页ID
        '
        If cllData.count = 0 Then
          If blnOnlyContractPati Then Exit Function
        Else
            Set cllTemp = cllData(1)
            '找到一个，直接返回
            lng病人ID = Val(NVL(cllTemp("_pati_id")))
            strCardPassWord = NVL(cllTemp("_card_pwd"))
            lngCardTypeID = Val(NVL(cllTemp("_cardtype_id")))
            intCardStatus = Val(NVL(cllTemp("_card_status")))
            strValidTime = NVL(cllTemp("_enduse_time"))
            If lng病人ID <> 0 Or blnOnlyContractPati Then
                zlGetPatiID = True: Exit Function
            End If
        End If
    
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
    If blnCertificate Then Exit Function
    
    If strCardType = "住院号" Or strCardType = "留观号" Then
        Set cllFilter = New Collection
        cllFilter.Add Array(strCardType, strCardNo)
        lng病人ID = zl_CisSvr_GetPatiID(cllFilter)
        zlGetPatiID = lng病人ID <> 0
        Exit Function
    End If
    Set cllFind = New Collection
    cllFind.Add Array(strCardType, strCardNo)
    cllFind.Add Array("住院状态", 2) '0-仅门诊;1-在院 ;2-门诊及在院
    If zl_PatiSvr_GetPatiInfo(0, cllFind, cllData, 1, , , , True, blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    '只有一条，直接返回
    If cllData.count = 1 Then
        Set cllTemp = cllData(1)
        lng病人ID = Val(NVL(cllTemp("_pati_id")))
        zlGetPatiID = True
        Exit Function
    End If
    
    
    str病人Ids = ""
    '构建查询的本地数据集
    Call zlGetQueryPatiInforStru(rsTemp, bytPatiInfoShowType)
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        
        If InStr(str病人Ids & ",", "," & Val(NVL(cllTemp("_pati_id"))) & ",") = 0 Then
            str病人Ids = str病人Ids & "," & Val(NVL(cllTemp("_pati_id")))
        End If
                
        With rsTemp
            rsTemp.AddNew
            !id = i
            !在院 = IIf(Val(NVL(cllTemp("_is_inhspt"))) = 1, "√", "")
            !病人ID = Val(NVL(cllTemp("_pati_id")))
            !名称 = cllTemp("_pati_name")
            
            !性别 = NVL(cllTemp("_pati_sex"), "")
            !年龄 = cllTemp("_pati_age")
            !身份证号 = cllTemp("_pati_idcard")
            !Ic卡号 = cllTemp("_iccard_no")
            !门诊号 = cllTemp("_outpatient_num")
            !住院号 = NVL(cllTemp("_inpatient_num"))
            !手机号 = cllTemp("_phone_number")
                
            If bytPatiInfoShowType = 0 Then
                !出生日期 = cllTemp("_pati_birthdate")
                !出生地点 = cllTemp("_pati_birthplace")
                !费别 = cllTemp("_fee_category")
                !医疗付款方式 = cllTemp("_mdlpay_mode_name")
                !民族 = cllTemp("_pati_nation")
                !家庭地址 = cllTemp("_pat_home_addr")
                !家庭电话 = cllTemp("_pat_home_phno")
                !联系人姓名 = cllTemp("_contacts_name")
                !联系人关系 = cllTemp("_contacts_relation")
                !联系人电话 = cllTemp("_contacts_phno")
                 
                Set cllItems = Nothing
                If mdlPubJson.CollectionExitsValue(cllTemp, "_card_list") Then
                     If UCase(TypeName(cllTemp("_card_list"))) = UCase("Collection") Then
                         Set cllItems = cllTemp("_card_list")
                     End If
                End If
                
                strCardNo = ""
                If Not cllItems Is Nothing Then
                    For J = 1 To cllItems.count
                        Set cllTemp = cllItems(J)
                        If Len(strCardNo & "," & cllTemp("_card_no")) > 200 Then Exit For
                        strCardNo = strCardNo & "," & cllTemp("_card_no")
                    Next
                    If strCardNo <> "" Then
                        strCardNo = Mid(strCardNo, 2)
                        !卡号 = strCardNo
                    End If
                End If
            End If
            rsTemp.Update
        End With
    Next
    
    If bytPatiInfoShowType = 0 And str病人Ids <> "" Then
        str病人Ids = Mid(str病人Ids, 2)
        If zl_ExseSvr_GetPatiSurplusInfo(str病人Ids, cllData, blnNotShowErrMsg, strErrMsg) Then
            For i = 1 To cllData.count
                Set cllTemp = cllData(i)
                rsTemp.Filter = "病人ID=" & Val(NVL(cllTemp("_pati_id")))
                If Not rsTemp.EOF Then
                    rsTemp!门诊预交余额 = Val(cllTemp("_outdpst_surplus"))
                    rsTemp!住院预交余额 = Val(cllTemp("_indpst_surplus"))
                    rsTemp.Update
                End If
            Next
        End If
    End If
    rsTemp.Filter = 0
    
    Set frmSel = New frmPatiSelect
    If Not frmSel.ShowSelect(frmMain, mcnOracle, glngSys, glngModul, objCtl, rsTemp, "病人选择", "", True, blnShowMergePati, _
                             IIf(objCtl Is Nothing, False, True), "", "ID", rsSeled, blnUserCancel) Then
        If Not frmSel Is Nothing Then Unload frmSel
        Set frmSel = Nothing: Exit Function
    End If
    
    If Not frmSel Is Nothing Then Unload frmSel
    Set frmSel = Nothing
    
    If rsSeled Is Nothing Then Exit Function
    If rsSeled.State <> 1 Then Exit Function
    
    If rsSeled.EOF Then
        If blnCard Then
            'IC卡,可能是模糊查找
            zlGetPatiID = GetCardFindPati(strCardNo, blnNotShowErrMsg, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID, , intCardStatus, , , strValidTime)
        End If
        rsSeled.Close
        Exit Function
    End If
    lng病人ID = Val(NVL(rsSeled!病人ID))
    zlGetPatiID = True: Set rsSeled = Nothing
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Set rsTemp = Nothing
End Function


Public Function zlGetPatiInfoRecordsFromCons(ByVal lng病人ID As Long, cllFilter As Collection, _
    ByRef rsPatiInfo_Out As ADODB.Recordset, Optional blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据条件获取病人信息数据给记录集
    '入参:cllFilter-病人信息集(array(查询名称,查询值)),如:array("卡号","A00001")
    '
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-21 19:12:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllData  As Collection, cllTemp  As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    If cllFilter Is Nothing And lng病人ID = 0 Then Exit Function
    
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        .fields.Append "病人id", adBigInt, , adFldIsNullable
        .fields.Append "主页id", adBigInt, , adFldIsNullable
        .fields.Append "姓名", adLongVarChar, 100, adFldIsNullable
        .fields.Append "性别", adLongVarChar, 100, adFldIsNullable
        .fields.Append "年龄", adLongVarChar, 100, adFldIsNullable
        .fields.Append "出生日期", adLongVarChar, 100, adFldIsNullable
        .fields.Append "费别", adLongVarChar, 100, adFldIsNullable
        .fields.Append "门诊号", adLongVarChar, 20, adFldIsNullable
        .fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .fields.Append "医疗付款方式名称", adLongVarChar, 100, adFldIsNullable
        .fields.Append "医疗付款方式编码", adLongVarChar, 30, adFldIsNullable
        .fields.Append "民族", adLongVarChar, 50, adFldIsNullable
        .fields.Append "医保号", adLongVarChar, 100, adFldIsNullable
        .fields.Append "身份证号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "就诊卡号", adLongVarChar, 100, adFldIsNullable
        .fields.Append "Ic卡号", adLongVarChar, 100, adFldIsNullable
        .fields.Append "住院次数", adBigInt, , adFldIsNullable
        .fields.Append "学历", adLongVarChar, 30, adFldIsNullable
        .fields.Append "职业", adLongVarChar, 50, adFldIsNullable
        .fields.Append "婚姻状况", adLongVarChar, 50, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 100, adFldIsNullable
        .fields.Append "家庭电话", adLongVarChar, 50, adFldIsNullable
        .fields.Append "出生地点", adLongVarChar, 100, adFldIsNullable
        .fields.Append "手机号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "当前床号", adLongVarChar, 30, adFldIsNullable
        .fields.Append "病人类型", adLongVarChar, 30, adFldIsNullable
        .fields.Append "险类", adLongVarChar, 20, adFldIsNullable
        .fields.Append "险类名称", adLongVarChar, 50, adFldIsNullable
        .fields.Append "是否在院", adBigInt, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    If zl_PatiSvr_GetPatiInfo(lng病人ID, cllFilter, cllData, 0, , , , , blnNotShowErrMsg, strErrMsg_Out) = False Then Exit Function   '只取基本信息
    If cllData.count = 0 Then
        strErrMsg_Out = "未找到符合条件的病人信息，请检查！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsPatiInfo_Out
            .AddNew
            !病人ID = Val(NVL(cllTemp("_pati_id")))
            !主页ID = Val(NVL(cllTemp("_pati_pageid")))
            !姓名 = NVL(cllTemp("_pati_name"))
            !性别 = NVL(cllTemp("_pati_sex"))
            !年龄 = NVL(cllTemp("_pati_age"))
            !出生日期 = NVL(cllTemp("_pati_birthdate"))
            !费别 = NVL(cllTemp("_fee_category"))
            !门诊号 = NVL(cllTemp("_outpatient_num"))
            !住院号 = NVL(cllTemp("_inpatient_num"))
            !医疗付款方式名称 = NVL(cllTemp("_mdlpay_mode_name"))
            !医疗付款方式编码 = NVL(cllTemp("_mdlpay_mode_code"))
            !民族 = NVL(cllTemp("_pati_nation"))
            !医保号 = NVL(cllTemp("_insurance_num"))
            !身份证号 = NVL(cllTemp("_pati_idcard"))
            !就诊卡号 = NVL(cllTemp("_vcard_no"))
            !Ic卡号 = NVL(cllTemp("_iccard_no"))
            !住院次数 = NVL(cllTemp("_inp_times"))
            !学历 = NVL(cllTemp("_pati_education"))
            !职业 = NVL(cllTemp("_ocpt_name"))
            !婚姻状况 = NVL(cllTemp("_pati_marital_cstatus"))
            !家庭地址 = NVL(cllTemp("_pat_home_addr"))
            !家庭电话 = NVL(cllTemp("_pat_home_phno"))
            !出生地点 = NVL(cllTemp("_pati_birthplace"))
            !手机号 = NVL(cllTemp("_phone_number"))
            !当前床号 = NVL(cllTemp("_pati_bed"))
            !病人类型 = NVL(cllTemp("_pati_type"))
            !险类 = Val(NVL(cllTemp("_insurance_type")))
            !险类名称 = NVL(cllTemp("_insurance_name"))
            !是否在院 = Val(NVL(cllTemp("_is_inhspt")))
            .Update
        End With
    Next
    zlGetPatiInfoRecordsFromCons = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPati As clsPatientInfo, _
    ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String, _
    Optional ByVal blnLoadFamily As Boolean, Optional ByVal blnLoadDrug As Boolean, _
    Optional ByVal blnLoadImmune As Boolean, Optional ByVal blnLoadCard As Boolean, _
    Optional cllFamily As Collection, Optional cllDrug As Collection, _
    Optional cllImmune As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
    Optional ByRef strExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '入参:blnLoadFamily-是否获取病人家属信息
    '     blnLoadDrug-是否获取病人过敏药物信息
    '     blnLoadImmune-是否获取病人家属信息
    '     blnLoadCard-仅占位，目前不用
    '     blnNotShowErrMsg - 是否提示错误信息
    '     strExpend -带扩展
    '出参
    '     strExpend -带扩展
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection
    Dim cllData  As Collection, cllTemp  As Collection, cllItems As Collection
    Dim i As Long
    
    On Error GoTo errHandle
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
     
    Set objPati = New clsPatientInfo
    If UCase(strOtherName) = "住院号" Then
         If zlGetPatiIDFromInpatientNum(strOtherValue, lng病人ID, blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    ElseIf strOtherName <> "" Then
        Set cllFilter = New Collection
        cllFilter.Add Array(UCase(strOtherName), strOtherValue)
    End If
    
    'int查询类型-0-基本;1-基本+联系人;2-所有
    If zl_PatiSvr_GetPatiInfo(lng病人ID, cllFilter, cllData, 2, blnLoadFamily, blnLoadDrug, blnLoadImmune, , blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    If cllData.count = 0 Then
          strErrMsg = "未找到符合条件的病人信息，请检查！"
          If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
    
    Set cllTemp = cllData(1)
    objPati.病人ID = cllTemp("_pati_id")
    objPati.主页ID = Val(NVL(cllTemp("_pati_pageid")))
    objPati.门诊号 = IIf(Val(NVL(cllTemp("_outpatient_num"))) = 0, "", cllTemp("_outpatient_num"))
    objPati.住院号 = IIf(Val(NVL(cllTemp("_inpatient_num"))) = 0, "", cllTemp("_inpatient_num"))
    objPati.姓名 = cllTemp("_pati_name")
    objPati.性别 = cllTemp("_pati_sex")
    objPati.年龄 = cllTemp("_pati_age")
    
    If NVL(cllTemp("_pati_birthdate")) <> "" Then
        objPati.出生日期 = Format(CDate(cllTemp("_pati_birthdate")), "yyyy-mm-dd")
        objPati.出生时间 = Format(CDate(cllTemp("_pati_birthdate")), "hh:mm")
    End If
    
    objPati.出生地址 = NVL(cllTemp("_pati_birthplace"))
    objPati.身份证号 = NVL(cllTemp("_pati_idcard"))
    objPati.其他证件 = NVL(cllTemp("_cert_no_other"))
    objPati.职业 = NVL(cllTemp("_ocpt_name"))
    objPati.费别 = NVL(cllTemp("_fee_category"))
    objPati.民族 = NVL(cllTemp("_pati_nation"))
    objPati.国籍 = NVL(cllTemp("_country_name"))
    objPati.学历 = NVL(cllTemp("_pati_education"))
    objPati.婚姻状况 = NVL(cllTemp("_pati_marital_cstatus"))
    objPati.区域 = NVL(cllTemp("_pati_area"))
    objPati.家庭地址 = NVL(cllTemp("_pat_home_addr"))
    objPati.家庭电话 = NVL(cllTemp("_pat_home_phno"))
    objPati.家庭邮编 = NVL(cllTemp("_pat_home_postcode"))
    objPati.监护人 = NVL(cllTemp("_pat_grdn_name"))
    objPati.联系人 = NVL(cllTemp("_contacts_name"))
    objPati.联系人关系 = NVL(cllTemp("_contacts_relation"))
    objPati.联系人地址 = NVL(cllTemp("_contacts_addr"))
    objPati.联系人电话 = NVL(cllTemp("_contacts_phno"))
    objPati.联系人身份证号 = NVL(cllTemp("_contacts_idcard"))
    objPati.工作单位 = NVL(cllTemp("_emp_name"))
    objPati.工作单位电话 = NVL(cllTemp("_emp_phno"))
    objPati.工作单位邮编 = NVL(cllTemp("_emp_postcode"))
    objPati.工作单位开户行 = NVL(cllTemp("_emp_bank_name"))
    objPati.工作单位开户行帐户 = NVL(cllTemp("_emp_bank_accnum"))
    objPati.户口地址 = NVL(cllTemp("_pat_hous_addr"))
    objPati.户口地址邮编 = NVL(cllTemp("_pat_hous_postcode"))
    objPati.籍贯 = NVL(cllTemp("_ntvplc_name"))
    objPati.密码 = ""
    objPati.医疗付款方式编码 = NVL(cllTemp("_mdlpay_mode_code"))
    objPati.医疗付款方式 = NVL(cllTemp("_mdlpay_mode_name"))
    objPati.病人类型 = NVL(cllTemp("_pati_type"))
    objPati.就诊卡号 = ""
    objPati.手机号 = NVL(cllTemp("_phone_number"))
    objPati.险类 = Val(NVL(cllTemp("_insurance_type")))
    objPati.险类名称 = NVL(cllTemp("_insurance_name"))
    objPati.医保号 = NVL(cllTemp("_insurance_num"))
    objPati.身份 = NVL(cllTemp("_pati_identity"))
    objPati.合同单位ID = Val(NVL(cllTemp("_ctt_unit_id")))
    objPati.登记时间 = NVL(cllTemp("_create_time"))
    objPati.在院 = Val(NVL(cllTemp("_is_inhspt"))) = 1
    objPati.入院日期 = NVL(cllTemp("_adta_time"))
    objPati.出院日期 = NVL(cllTemp("_adtd_time"))
    objPati.卡验证码 = NVL(cllTemp("_card_captcha"))
    objPati.当前科室ID = Val(NVL(cllTemp("_pati_dept_id")))
    objPati.当前科室名称 = NVL(cllTemp("_pati_dept_name"))
    objPati.当前病区ID = Val(NVL(cllTemp("_pati_wardarea_id")))
    objPati.当前病区名称 = NVL(cllTemp("_pati_wardarea_name"))
    objPati.最后主页ID = objPati.主页ID
    objPati.就诊时间 = NVL(cllTemp("_visit_time"))
    objPati.医保密码 = NVL(cllTemp("_insurance_pwd"))
    If NVL(cllTemp("_visit_time")) = "" Then
        objPati.初诊 = True
    ElseIf NVL(cllTemp("_create_time")) <> "" Then '就诊时间等于登记时间为初诊
        objPati.初诊 = CDate(NVL(cllTemp("_visit_time"))) = CDate(NVL(cllTemp("_create_time")))
    End If
    
    If blnLoadFamily Then
        If CollectionExitsValue(cllData(1), "_family_list") Then
            Set cllItems = cllData(1)("_family_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_family_id"), "家属ID"
                cllTemp.Add cllItems(i)("_family_relation"), "关系"
                cllFamily.Add cllTemp, "_" & cllTemp("家属id")
            Next
        End If
    End If
    
    If blnLoadDrug Then
        If CollectionExitsValue(cllData(1), "_drug_list") Then
            Set cllItems = cllData(1)("_drug_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_pat_algc_cadn_id"), "过敏药品ID"
                cllTemp.Add cllItems(i)("_pat_algc_cadn"), "过敏药物名称"
                cllTemp.Add cllItems(i)("_allergy_info"), "过敏药物反应"
                cllDrug.Add cllTemp, "_" & cllTemp("过敏药品ID")
            Next
        End If
    End If
    
    If blnLoadImmune Then
        If CollectionExitsValue(cllData(1), "_immune_list") Then
            Set cllItems = cllData(1)("_immune_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_vaccinate_time"), "接种时间"
                cllTemp.Add cllItems(i)("_vaccinate_name"), "接种名称"
                cllImmune.Add cllTemp
            Next
        End If
    End If
    zlGetPatiInforFromPatiID = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlGetMultiPatiInforFromPatiID(ByVal str病人Ids As String, ByRef cllPatis_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID一次获取多个病人信息集
    '入参:
    '   str病人IDs=多个病人ID用英文逗号分隔
    '出参:
    '   cllPatis_Out-返回病人信息集,成员：clsPatinetInfo=cllPatis_Out(_病人ID)
    '返回:成功返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, blnNotShowErrMsg As Boolean, strErrMsg As String
    Dim strJson As String, i As Long, objPati As clsPatientInfo
    Dim cllData  As Collection, cllTemp  As Collection
    Dim lng病人ID1 As Long, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Set cllPatis_Out = New Collection
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    Set cllFilter = New Collection
    cllFilter.Add Array("病人IDS", str病人Ids)
    cllFilter.Add Array("住院状态", 2)
    If zl_PatiSvr_GetPatiInfo(0, cllFilter, cllData, 2, , , , , True, strErrMsg, True) = False Then Exit Function
    If cllData.count = 0 Then
        strErrMsg = "未找到符合条件的病人信息，请检查！"
        If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        Set objPati = New clsPatientInfo
        objPati.病人ID = cllTemp("_pati_id")
        objPati.主页ID = Val(NVL(cllTemp("_pati_pageid")))
        objPati.门诊号 = IIf(Val(NVL(cllTemp("_outpatient_num"))) = 0, "", cllTemp("_outpatient_num"))
        objPati.住院号 = IIf(Val(NVL(cllTemp("_inpatient_num"))) = 0, "", cllTemp("_inpatient_num"))
        objPati.姓名 = cllTemp("_pati_name")
        objPati.性别 = cllTemp("_pati_sex")
        objPati.年龄 = cllTemp("_pati_age")
        
        If NVL(cllTemp("_pati_birthdate")) <> "" Then
            objPati.出生日期 = Format(CDate(cllTemp("_pati_birthdate")), "yyyy-mm-dd")
            objPati.出生时间 = Format(CDate(cllTemp("_pati_birthdate")), "hh:mm")
        End If
        
        objPati.出生地址 = NVL(cllTemp("_pati_birthplace"))
        objPati.身份证号 = NVL(cllTemp("_pati_idcard"))
        objPati.其他证件 = NVL(cllTemp("_cert_no_other"))
        objPati.职业 = NVL(cllTemp("_ocpt_name"))
        objPati.费别 = NVL(cllTemp("_fee_category"))
        objPati.民族 = NVL(cllTemp("_pati_nation"))
        objPati.国籍 = NVL(cllTemp("_country_name"))
        objPati.学历 = NVL(cllTemp("_pati_education"))
        objPati.婚姻状况 = NVL(cllTemp("_pati_marital_cstatus"))
        objPati.区域 = NVL(cllTemp("_pati_area"))
        objPati.家庭地址 = NVL(cllTemp("_pat_home_addr"))
        objPati.家庭电话 = NVL(cllTemp("_pat_home_phno"))
        objPati.家庭邮编 = NVL(cllTemp("_pat_home_postcode"))
        objPati.监护人 = NVL(cllTemp("_pat_grdn_name"))
        objPati.联系人 = NVL(cllTemp("_contacts_name"))
        objPati.联系人关系 = NVL(cllTemp("_contacts_relation"))
        objPati.联系人地址 = NVL(cllTemp("_contacts_addr"))
        objPati.联系人电话 = NVL(cllTemp("_contacts_phno"))
        objPati.工作单位 = NVL(cllTemp("_emp_name"))
        objPati.工作单位电话 = NVL(cllTemp("_emp_phno"))
        objPati.工作单位邮编 = NVL(cllTemp("_emp_postcode"))
        objPati.工作单位开户行 = NVL(cllTemp("_emp_bank_name"))
        objPati.工作单位开户行帐户 = NVL(cllTemp("_emp_bank_accnum"))
        objPati.户口地址 = NVL(cllTemp("_pat_hous_addr"))
        objPati.户口地址邮编 = NVL(cllTemp("_pat_hous_postcode"))
        objPati.籍贯 = NVL(cllTemp("_ntvplc_name"))
        objPati.密码 = ""
        objPati.医疗付款方式编码 = NVL(cllTemp("_mdlpay_mode_code"))
        objPati.医疗付款方式 = NVL(cllTemp("_mdlpay_mode_name"))
        objPati.病人类型 = NVL(cllTemp("_pati_type"))
        objPati.就诊卡号 = ""
        objPati.手机号 = NVL(cllTemp("_phone_number"))
        objPati.险类 = Val(NVL(cllTemp("_insurance_type")))
        objPati.险类名称 = NVL(cllTemp("_insurance_name"))
        objPati.医保号 = NVL(cllTemp("_insurance_num"))
        objPati.身份 = NVL(cllTemp("_pati_identity"))
        objPati.合同单位ID = Val(NVL(cllTemp("_ctt_unit_id")))
        objPati.登记时间 = NVL(cllTemp("_create_time"))
        objPati.在院 = Val(NVL(cllTemp("_is_inhspt"))) = 1
        objPati.入院日期 = NVL(cllTemp("_adta_time"))
        objPati.出院日期 = NVL(cllTemp("_adtd_time"))
        objPati.最后主页ID = objPati.主页ID
        objPati.就诊时间 = NVL(cllTemp("_visit_time"))
        objPati.医保密码 = NVL(cllTemp("_insurance_pwd"))
        If NVL(cllTemp("_visit_time")) = "" Then
            objPati.初诊 = True
        ElseIf NVL(cllTemp("_create_time")) <> "" Then  '就诊时间等于登记时间为初诊
            objPati.初诊 = CDate(NVL(cllTemp("_visit_time"))) = CDate(NVL(cllTemp("_create_time")))
        End If
        cllPatis_Out.Add objPati, "_" & objPati.病人ID
    Next
    zlGetMultiPatiInforFromPatiID = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlGetPatiOtherInforFromPatiID(ByVal lng病人ID As Long, _
    Optional ByVal blnNotShowErrMsg As Boolean, Optional ByRef strErrMsg As String, _
    Optional ByVal blnLoadFamily As Boolean, Optional ByVal blnLoadDrug As Boolean, _
    Optional ByVal blnLoadImmune As Boolean, Optional ByVal blnLoadCert As Boolean, _
    Optional ByVal blnLoadOther As Boolean, _
    Optional cllFamily As Collection, Optional cllDrug As Collection, _
    Optional cllImmune As Collection, Optional cllCert As Collection, _
    Optional cllOther As Collection, Optional cllContact As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-04-06 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, i As Long
    Dim cllData  As Collection, cllTemp  As Collection, cllItems As Collection
    Dim strName As String, strValue As String
    Dim intReturn As Integer
    
    On Error GoTo errHandle
    Set cllFamily = New Collection
    Set cllDrug = New Collection
    Set cllImmune = New Collection
    Set cllCert = New Collection
    Set cllOther = New Collection
    Set cllContact = New Collection
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    'int查询类型-0-基本;1-基本+联系人;2-所有
    If blnLoadFamily Or blnLoadDrug Or blnLoadImmune Then
        If zl_PatiSvr_GetPatiInfo(lng病人ID, cllFilter, cllData, 0, blnLoadFamily, blnLoadDrug, blnLoadImmune, , blnNotShowErrMsg, strErrMsg) = False Then Exit Function
        If cllData Is Nothing Then
            If strErrMsg = "" Then strErrMsg = "未找到符合条件的病人信息，请检查！"
              If Not blnNotShowErrMsg Then MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
              Exit Function
        End If
        If cllData.count = 0 Then zlGetPatiOtherInforFromPatiID = True: Exit Function
    End If
    
    If blnLoadFamily Then
        If CollectionExitsValue(cllData(1), "_family_list") Then
            Set cllItems = cllData(1)("_family_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_family_id"), "家属ID"
                cllTemp.Add cllItems(i)("_family_relation"), "关系"
                cllFamily.Add cllTemp, "_" & cllTemp("家属id")
            Next
        End If
    End If
    
    If blnLoadDrug Then
        If CollectionExitsValue(cllData(1), "_drug_list") Then
            Set cllItems = cllData(1)("_drug_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_pat_algc_cadn_id"), "过敏药品ID"
                cllTemp.Add cllItems(i)("_pat_algc_cadn"), "过敏药物名称"
                cllTemp.Add cllItems(i)("_allergy_info"), "过敏药物反应"
                cllDrug.Add cllTemp, "_" & cllTemp("过敏药品ID")
            Next
        End If
    End If
    
    If blnLoadImmune Then
        If CollectionExitsValue(cllData(1), "_immune_list") Then
            Set cllItems = cllData(1)("_immune_list")
            For i = 1 To cllItems.count
                Set cllTemp = New Collection
                cllTemp.Add cllItems(i)("_vaccinate_time"), "接种时间"
                cllTemp.Add cllItems(i)("_vaccinate_name"), "接种名称"
                cllImmune.Add cllTemp
            Next
        End If
    End If
    
    If blnLoadCert Then
        Call zlGetPatiCertInfo(lng病人ID, cllCert)
    End If
    
    If Not blnLoadOther Then zlGetPatiOtherInforFromPatiID = True: Exit Function
    Set cllData = New Collection
    If zl_PatiSvr_GetPatiExtendInfo(lng病人ID, "", cllData, blnNotShowErrMsg, strErrMsg) = False Then Exit Function
    If cllData Is Nothing Then zlGetPatiOtherInforFromPatiID = True: Exit Function
    If cllData.count = 0 Then zlGetPatiOtherInforFromPatiID = True: Exit Function
    
    For i = 1 To cllData.count
        Set cllTemp = New Collection
        Set cllItems = cllData(i)
        If cllItems("_info_name") Like "联系人*" And cllItems("_info_name") <> "联系人附加信息" Then
            strName = cllItems("_info_name"): strValue = ""
            If strValue Like "联系人姓名*" Then
                strValue = "联系人姓名"
            ElseIf strValue Like "联系人关系*" Then
                strValue = "联系人关系"
            ElseIf strValue Like "联系人电话*" Then
                strValue = "联系人电话"
            ElseIf strValue Like "联系人身份证号*" Then
                strValue = "联系人身份证号"
            ElseIf strValue Like "联系人附加信息*" Then
                strValue = "联系人附加信息"
            End If
            If strValue <> "" Then
                intReturn = Mid(strValue, Len(strValue) + 1)
                If CollectionExitsValue(cllContact, "Key" & intReturn) Then
                    Set cllTemp = cllContact("Key" & intReturn)
                    cllTemp.Add cllItems("_upd_info_value"), strValue
                Else
                    Set cllTemp = New Collection
                    cllTemp.Add cllItems("_upd_info_value"), strValue
                    cllContact.Add cllTemp, "Key" & intReturn
                End If
            End If
        Else
            cllTemp.Add cllItems("_info_name"), "信息名"
            cllTemp.Add cllItems("_info_value"), "信息值"
            cllTemp.Add NVL(cllItems("_visit_id")), "就诊ID"
            cllOther.Add cllTemp, cllTemp("信息名") & IIf(Val(cllTemp("就诊ID")) <> 0, "_" & cllTemp("就诊ID"), "")
        End If
    Next
    
    zlGetPatiOtherInforFromPatiID = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function zlGetPatiCertInfo(ByVal lng病人ID As Long, ByRef cllCert As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人有效的证件卡信息
    ' 入参 :
    ' 出参 : cllCert-证件列表
    ' 返回 :
    ' 编制 : 李南春
    ' 日期 : 2019/11/8 19:55
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection
    Dim cllItems As Collection
    Dim cllTemp As Collection
    Dim i As Long
    On Error GoTo ErrHand

    Set cllCert = New Collection
    If Zl_Patisvr_GetPatiCardInfo("", lng病人ID, 2, , , cllData, True) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    For i = 1 To cllData.count
        Set cllItems = cllData(i)
        Set cllTemp = New Collection
        cllTemp.Add cllItems("_cardtype_name"), "卡类别名称"
        cllTemp.Add cllItems("_card_no"), "卡号"
        cllCert.Add cllTemp, cllTemp("卡类别名称")
    Next
    zlGetPatiCertInfo = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
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
    Dim strSQL As String
    
    On Error GoTo ErrHand
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
ErrHand:
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlGetPatiInforFromXML(ByVal strPatiXml As String, ByRef strErrMsg_Out As String, ByRef objPatiInfor_Out As clsPatientInfo) As Boolean
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
    Set objPatiInfor_Out = New clsPatientInfo
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
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    
    Err = 0: On Error GoTo ErrHand:
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
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
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

Public Function zlGetPatiCardFromCards(ByVal strCardTypeIDs As String, ByVal lng病人ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从给定卡类别中检索指定病人持有有效卡的卡类别
    '入参: strCardTypeIDs 给定卡类别，多个用逗号分隔
    '返回:返回病人持有有效卡的卡类别，多个用逗号分隔
    '编制:刘兴洪
    '日期:2018-12-03 15:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardTypes As String
    Err = 0: On Error GoTo errHandle:
    If Zl_Patisvr_GetPatiCardInfo(strCardTypeIDs, lng病人ID, , True, strCardTypes) = False Then Exit Function
    zlGetPatiCardFromCards = strCardTypes
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    Dim blnTemp As Boolean, varTemp  As Variant
    Dim i As Long
    
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
        For i = 0 To UBound(varTemp)
            If Trim(varTemp(i)) <> "" Then
                If InStr(strColumns_out & ",", "," & strAliaName & "." & varTemp(i) & ",") = 0 Then
                    strColumns_out = strColumns_out & ",NULL as " & varTemp(i)
                End If
            End If
        Next
    End If
    If strColumns_out <> "" Then strColumns_out = Mid(strColumns_out, 2)
    zlGetCojumnsFromTableName = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    
    Err = 0: On Error GoTo ErrHand:
    
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
            If Val(NVL(rsTemp!是否自制)) = 1 Or Val(NVL(rsTemp!是否证件)) = 1 Then        '自制卡,都启用
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
            .卡号长度 = Val(NVL(rsTemp!卡号长度)) + Val(NVL(rsTemp!是否启用回车))
            .缺省标志 = Val(NVL(rsTemp!缺省标志)) = 1
            .系统 = Val(NVL(rsTemp!是否固定)) = 1
            .是否严格控制 = Val(NVL(rsTemp!是否严格控制)) = 1
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(NVL(rsTemp!是否自制))
            .是否存在帐户 = Val(NVL(rsTemp!是否存在帐户)) = 1
            .是否全退 = Val(NVL(rsTemp!是否全退)) = 1
            .卡号重复使用 = Val(NVL(rsTemp!是否重复使用)) = 1
            .结算方式 = NVL(rsTemp!结算方式)
            .接口程序名 = NVL(rsTemp!部件)
            .特定项目 = NVL(rsTemp!特定项目)
            .启用 = bln启用
            .备注 = NVL(rsTemp!备注)
            .卡号密文规则 = NVL(rsTemp!卡号密文)
            .是否退现 = Val(NVL(rsTemp!是否退现)) = 1
            .密码长度 = Val(NVL(rsTemp!密码长度))
            .密码长度限制 = Val(NVL(rsTemp!密码长度限制))
            .密码规则 = Val(NVL(rsTemp!密码规则))
            .密码输入限制 = Val(NVL(rsTemp!密码输入限制))
            .是否缺省密码 = Val(NVL(rsTemp!是否缺省密码))
            .是否制卡 = Val(NVL(rsTemp!是否制卡)) = 1     '56615
            .是否发卡 = Val(NVL(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(NVL(rsTemp!是否写卡)) = 1
            .是否模糊查找 = Val(NVL(rsTemp!是否模糊查找)) = 1
            .是否转帐及代扣 = Val(NVL(rsTemp!是否转帐及代扣)) = 1
            .发卡性质 = Val(NVL(rsTemp!发卡性质))
            .发卡控制 = Val(NVL(rsTemp!发卡控制))
            
            str读卡性质 = Trim(NVL(rsTemp!读卡性质))
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            
            .是否证件 = Val(NVL(rsTemp!是否证件)) = 1
            .是否持卡消费 = Val(NVL(rsTemp!是否持卡消费)) = 1
            .是否发送调用接口 = Val(NVL(rsTemp!发送调用接口)) = 1
            .是否退款验卡 = Val(NVL(rsTemp!是否退款验卡)) = 1
            .设备是否启用回车 = Val(NVL(rsTemp!是否启用回车)) = 1
            .是否缺省退现 = Val(NVL(rsTemp!是否缺省退现)) = 1
            .结算性质 = Val(NVL(rsTemp!结算性质))
            .是否独立结算 = Val(NVL(rsTemp!是否独立结算)) = 1
            .是否支持扫码付 = Val(NVL(rsTemp!是否支持扫码付)) = 1
            .卡号识别规则 = Val(NVL(rsTemp!卡号识别规则))
            .缺省有效时间 = NVL(rsTemp!缺省有效时间)
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
        int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编码), "自动读取", "0"))
        int自动读取间隔 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & NVL(rsTemp!编码), "自动读取间隔", "300"))
        bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & NVL(rsTemp!编码), "启用", "0")) = 1
        
        '编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度,部件,系统,是否密文
        str部件 = Trim(NVL(rsTemp!部件))
        Set objCard = New Card
        With objCard
            .卡种类 = EM_CardType_Consume
            .接口序号 = NVL(rsTemp!id)
            .接口编码 = NVL(rsTemp!编码)
            .短名 = NVL(rsTemp!短名)     '默认取第一个
            .名称 = NVL(rsTemp!名称)
            .前缀文本 = NVL(rsTemp!前缀文本)
            .卡号长度 = Val(NVL(rsTemp!卡号长度))
            .系统 = Val(NVL(rsTemp!是否固定)) = 1
            .是否严格控制 = False
            .是否自动读取 = int自动读取
            .自动读取间隔 = int自动读取间隔
            .自制卡 = Val(NVL(rsTemp!是否自制))
            .是否存在帐户 = True 'Not (Val(Nvl(rsTemp!自制卡)) = 1)
            .是否全退 = Val(NVL(rsTemp!是否全退)) = 1
            .结算方式 = Trim(NVL(rsTemp!结算方式))
            .接口程序名 = Trim(NVL(rsTemp!部件))
            .特定项目 = ""
            .启用 = bln启用
            .卡号重复使用 = True
            .备注 = ""
            .卡号密文规则 = NVL(rsTemp!卡号密文)
            .消费卡 = True
            .是否退现 = Val(NVL(rsTemp!是否退现)) = 1
            .密码长度 = Val(NVL(rsTemp!密码长度))
            .密码长度限制 = Val(NVL(rsTemp!密码长度限制))
            .密码规则 = Val(NVL(rsTemp!密码规则))
            .密码输入限制 = Val(NVL(rsTemp!密码输入限制))
            .是否缺省密码 = Val(NVL(rsTemp!是否缺省密码)) = 1
            .是否制卡 = Val(NVL(rsTemp!是否制卡)) = 1
            .是否发卡 = Val(NVL(rsTemp!是否发卡)) = 1 Or .自制卡
            .是否写卡 = Val(NVL(rsTemp!是否写卡)) = 1
            
            str读卡性质 = NVL(rsTemp!读卡性质)
            str读卡性质 = IIf(str读卡性质 = "", "1000", str读卡性质)
            
            .是否刷卡 = Mid(str读卡性质, 1, 1) = 1
            .是否扫描 = Mid(str读卡性质, 2, 1) = 1
            .是否接触式读卡 = Mid(str读卡性质, 3, 1) = 1
            .是否非接触式读卡 = Mid(str读卡性质, 4, 1) = 1
            .结算性质 = Val(NVL(rsTemp!结算性质))
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
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlIsExistFeeInsurePatient(lng病人ID As Long, Optional blnExist_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断医保病人是否存在未结费用
    '入参:lng病人ID
    '出参:blnExist_out-存在未结费用，返回true,否则返回False
    '返回:获取成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 16:40:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim cllData  As Collection, cllTemp  As Collection
    Dim intReturn As Integer
    Dim dblMoney As Double
    
    On Error GoTo ErrH
    
    blnExist_Out = False
    
    If zl_PatiSvr_GetInsureByPatiID(lng病人ID, intReturn) = False Then Exit Function
    If intReturn = 0 Then zlIsExistFeeInsurePatient = True: Exit Function
    
    If zl_ExseSvr_GetPatiSurplusInfo(lng病人ID, cllData) = False Then Exit Function
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        dblMoney = dblMoney + Val(cllTemp("_outfee_surplus")) + Val(cllTemp("_infee_surplus"))
    Next
    
    blnExist_Out = dblMoney <> 0
    zlIsExistFeeInsurePatient = True
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
   Call gobjComLib.SaveErrLog
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
    Dim strSQL As String, i As Integer
    Dim strUnitIDs As String
    On Error GoTo ErrH
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
    For i = 1 To rsTmp.RecordCount
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
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
   Call gobjComLib.SaveErrLog
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
    
    On Error GoTo ErrH
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
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
   Call gobjComLib.SaveErrLog
End Function

Public Function zlGetPatiRecordFromPatiName(ByVal strName As String, ByRef rsPati As ADODB.Recordset, _
    Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal blnOnlyUnitPatient As Boolean, _
    Optional ByVal blnLikeByName As Boolean = True, Optional ByVal byt场合 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息记录集
    '入参:
    '   objControl-调用的控件
    '   strName-病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   blnOnlyUnitPatient-只查合约单位病人
    '   blnLikeByName-是否按姓名做匹配模糊查找
    '出参:
    '   rsPati-病人信息集,字段：排序ID,ID,病人ID,主页ID,姓名,性别,年龄,病人类型,险类,门诊号,住院号,科室,
    '                           床号,出生日期,身份证号,家庭地址,工作单位,在院标志,手机号,是否有医疗卡,就诊时间,入院日期,出院日期
    '返回:成功返回true,否则返回False
    '编制:李南春
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllData As Collection, cllTemp As Collection
    Dim datNow As Date, str开始时间  As String, str结束时间  As String
    Dim cllOthersFind As Collection
    Dim i As Long, str病人Ids As String, lng病人ID As Long
    
    On Error GoTo ErrH
    Set rsPati = New ADODB.Recordset
    With rsPati
        .fields.Append "排序ID", adInteger, 2, adFldIsNullable
        .fields.Append "ID", adBigInt, 18, adFldIsNullable
        .fields.Append "病人ID", adBigInt, 18, adFldIsNullable
        .fields.Append "主页ID", adBigInt, 18, adFldIsNullable
        .fields.Append "姓名", adVarChar, 100, adFldIsNullable
        .fields.Append "性别", adVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adVarChar, 20, adFldIsNullable
        
        .fields.Append "病人类型", adVarChar, 100, adFldIsNullable
        .fields.Append "险类", adVarChar, 18, adFldIsNullable
        
        .fields.Append "住院号", adVarChar, 18, adFldIsNullable
        .fields.Append "门诊号", adVarChar, 18, adFldIsNullable
        .fields.Append "科室", adVarChar, 50, adFldIsNullable
        .fields.Append "床号", adVarChar, 100, adFldIsNullable
        
        .fields.Append "出生日期", adVarChar, 30, adFldIsNullable
        .fields.Append "身份证号", adVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "在院标志", adVarChar, 4, adFldIsNullable
        .fields.Append "手机号", adVarChar, 50, adFldIsNullable
        
        .fields.Append "入院日期", adVarChar, 30, adFldIsNullable
        .fields.Append "出院日期", adVarChar, 30, adFldIsNullable
        
        .fields.Append "是否有医疗卡", adVarChar, 4, adFldIsNullable
        .fields.Append "就诊时间", adVarChar, 20, adFldIsNullable
                
        If bln显示门诊预交 Then .fields.Append "门诊预交余额", adLongVarChar, 100, adFldIsNullable
        If bln显示住院预交 Then .fields.Append "住院预交余额", adLongVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    'If LenB(strName) < 4 Then Exit Function
    
    '查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
    Set cllOthersFind = New Collection
    cllOthersFind.Add Array("姓名", strName & IIf(blnLikeByName, "%", "")), "姓名"
    cllOthersFind.Add Array("最大记录数", 101), "最大记录数"
    If intDefaultDays <> 0 Then
        datNow = gobjDatabase.Currentdate
        str开始时间 = Format(datNow - intDefaultDays, "YYYY-MM-DD")
        str结束时间 = Format(datNow, "YYYY-MM-DD hh:mm:ss")
        cllOthersFind.Add Array("登记时间", str开始时间, str结束时间), "登记时间"
        cllOthersFind.Add Array("就诊时间", str开始时间, str结束时间), "就诊时间"
    End If
    If lngDefaultCardTypeID <> 0 Then
        cllOthersFind.Add Array("缺省卡类别ID", lngDefaultCardTypeID), "缺省卡类别ID"
    End If
    If blnOnlyUnitPatient Then
        cllOthersFind.Add Array("仅合约单位病人", 1), "仅合约单位病人"
    End If
    
    If byt场合 <> 0 Then
        cllOthersFind.Add Array("场合", byt场合), "场合"
    End If
    
    If zl_PatiSvr_GetPatiInfsByRange(Decode(byt病人来源, 1, 0, 2, 1, 2), cllOthersFind, cllData) = False Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    str病人Ids = ""
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsPati
            .AddNew
            !排序ID = 1
            !id = cllTemp("_pati_id")
            !病人ID = cllTemp("_pati_id")
            !主页ID = cllTemp("_pati_pageid")
            !姓名 = cllTemp("_pati_name")
            !性别 = cllTemp("_pati_sex")
            !年龄 = cllTemp("_pati_age")
            !病人类型 = cllTemp("_pati_type")
            !险类 = cllTemp("_insurance_type")
            
            !门诊号 = cllTemp("_outpatient_num")
            !住院号 = cllTemp("_inpatient_num")
            !科室 = cllTemp("_pati_dept_name")
            !床号 = cllTemp("_pati_bed")
            
            !工作单位 = cllTemp("_emp_name")
            !出生日期 = cllTemp("_pati_birthdate")
            !身份证号 = cllTemp("_pati_idcard")
            !家庭地址 = cllTemp("_pati_identity")
            !在院标志 = IIf(Val(NVL(cllTemp("_is_inhspt"))) = 1, "√", "")
            !手机号 = cllTemp("_phone_number")
            
            !是否有医疗卡 = IIf(NVL(cllTemp("_medc_card_no")) <> "", "√", "")
            !就诊时间 = cllTemp("_visit_time")
            
            .Update
        End With
        str病人Ids = str病人Ids & "," & Val(NVL(cllTemp("_pati_id")))
    Next
    If str病人Ids <> "" Then str病人Ids = Mid(str病人Ids, 2)
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    If (bln显示门诊预交 Or bln显示住院预交) And str病人Ids <> "" Then
        Dim cllSurplus As Collection
        If zl_ExseSvr_GetPatiSurplusInfo(str病人Ids, cllSurplus) Then
            With rsPati
                Do While Not .EOF
                    lng病人ID = Val(NVL(!病人ID))
                    Set cllTemp = mdlPubJson.zlGetNodeObjectFromCollect(cllData, "_" & lng病人ID)
                    If Not cllTemp Is Nothing Then  '
                        If bln显示门诊预交 Then !门诊预交余额 = Format(Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outdpst_surplus", "C"))), "####0.00;-####0.00; ;")
                        If bln显示住院预交 Then !住院预交余额 = Format(Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_indpst_surplus", "C"))), "####0.00;-####0.00; ;")
                        .Update
                    End If
                    .MoveNext
                Loop
                .MoveFirst
            End With
        End If
    End If
    zlGetPatiRecordFromPatiName = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiIdFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef lngPatiId As Long, _
    Optional frmMain As Object, Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, Optional ByVal blnAddNewPati As Boolean, _
    Optional ByVal blnOnlyUnitPatient As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院;3-分诊;4-挂号过滤窗体;5-医疗卡;6-入院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsSel As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strHideCols As String
    Dim byt调用场合 As Byte
    
    If byt场合 = 0 Then
        byt调用场合 = 1
    ElseIf byt场合 = 3 Then
        byt调用场合 = 2
    ElseIf byt场合 = 4 Then
        byt调用场合 = 3
    End If
    
    If zlGetPatiRecordFromPatiName(strName, rsTmp, bln显示门诊预交, bln显示住院预交, _
        intDefaultDays, lngDefaultCardTypeID, byt病人来源, blnOnlyUnitPatient, , byt调用场合) = False Then Exit Function
    
    If blnAddNewPati Then '添加新病人
        rsTmp.AddNew
        rsTmp!排序ID = 0
        rsTmp!id = 0
        rsTmp!姓名 = "[新病人]"
    End If
    If rsTmp.RecordCount = 0 Then Exit Function
    
    strHideCols = "排序ID,ID"
    Select Case byt场合
    Case 0, 3, 4 '门诊挂号
        strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志,入院日期,出院日期"
    Case 1 '门诊收费
        If byt病人来源 = 2 Then
            strHideCols = strHideCols & ",病人类型,险类,门诊号,在院标志,手机号,是否有医疗卡,就诊时间"
        Else
            strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志,手机号,是否有医疗卡,就诊时间"
        End If
    Case 2 '住院
        strHideCols = strHideCols & ",病人类型,险类,在院标志,手机号,是否有医疗卡,就诊时间"
    End Select
    
    rsTmp.Sort = "排序ID,姓名"
    If frmPatiSelect.ShowSelect(frmMain, mcnOracle, glngSys, glngModul, objControl, _
        rsTmp, "病人信息查询", "", True, False, True, "", strHideCols, rsSel, blnCancel) = False Then Exit Function
    If rsSel Is Nothing Then Exit Function
    lngPatiId = NVL(rsSel!id)
    zlGetPatiIdFromPatiName = lngPatiId <> 0
End Function

Public Function zlGetPatiRecordFromPatiIDCard(ByVal strIDCard As String, ByRef rsPati As ADODB.Recordset, _
    Optional ByVal strName As String, Optional ByVal strSex As String, _
    Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal blnOnlyUnitPatient As Boolean, _
    Optional ByVal byt场合 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息记录集
    '入参:
    '   objControl-调用的控件
    '   strName-病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   blnOnlyUnitPatient-只查合约单位病人
    '   blnLikeByName-是否按姓名做匹配模糊查找
    '出参:
    '   rsPati-病人信息集,字段：排序ID,ID,病人ID,主页ID,姓名,性别,年龄,病人类型,险类,门诊号,住院号,科室,
    '                           床号,出生日期,身份证号,家庭地址,工作单位,在院标志,手机号,是否有医疗卡,就诊时间,入院日期,出院日期
    '返回:成功返回true,否则返回False
    '编制:李南春
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllData As Collection, cllTemp As Collection
    Dim datNow As Date, str开始时间  As String, str结束时间  As String
    Dim cllOthersFind As Collection
    Dim i As Long, str病人Ids As String, lng病人ID As Long
    
    On Error GoTo ErrH
    Set rsPati = New ADODB.Recordset
    With rsPati
        .fields.Append "排序ID", adInteger, 2, adFldIsNullable
        .fields.Append "ID", adBigInt, 18, adFldIsNullable
        .fields.Append "病人ID", adBigInt, 18, adFldIsNullable
        .fields.Append "主页ID", adBigInt, 18, adFldIsNullable
        .fields.Append "姓名", adVarChar, 100, adFldIsNullable
        .fields.Append "性别", adVarChar, 20, adFldIsNullable
        .fields.Append "年龄", adVarChar, 20, adFldIsNullable
        
        .fields.Append "病人类型", adVarChar, 100, adFldIsNullable
        .fields.Append "险类", adVarChar, 18, adFldIsNullable
        
        .fields.Append "住院号", adVarChar, 18, adFldIsNullable
        .fields.Append "门诊号", adVarChar, 18, adFldIsNullable
        .fields.Append "科室", adVarChar, 50, adFldIsNullable
        .fields.Append "床号", adVarChar, 100, adFldIsNullable
        
        .fields.Append "出生日期", adVarChar, 30, adFldIsNullable
        .fields.Append "身份证号", adVarChar, 30, adFldIsNullable
        .fields.Append "家庭地址", adLongVarChar, 200, adFldIsNullable
        .fields.Append "工作单位", adLongVarChar, 200, adFldIsNullable
        .fields.Append "在院标志", adVarChar, 4, adFldIsNullable
        .fields.Append "手机号", adVarChar, 50, adFldIsNullable
        
        .fields.Append "入院日期", adVarChar, 30, adFldIsNullable
        .fields.Append "出院日期", adVarChar, 30, adFldIsNullable
        
        .fields.Append "是否有医疗卡", adVarChar, 4, adFldIsNullable
        .fields.Append "就诊时间", adVarChar, 20, adFldIsNullable
                
        If bln显示门诊预交 Then .fields.Append "门诊预交余额", adLongVarChar, 100, adFldIsNullable
        If bln显示住院预交 Then .fields.Append "住院预交余额", adLongVarChar, 100, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    'If LenB(strName) < 4 Then Exit Function
    
    '查询住院状态:0-仅门诊;1-在院 ;2-门诊及在院
    Set cllOthersFind = New Collection
    
    cllOthersFind.Add Array("身份证号", strIDCard), "身份证号"
    If strName <> "" Then
        cllOthersFind.Add Array("姓名", strName), "姓名"
    End If
    If strSex <> "" Then
        cllOthersFind.Add Array("性别", strSex), "性别"
    End If
    cllOthersFind.Add Array("最大记录数", 101), "最大记录数"
   
    If blnOnlyUnitPatient Then
        cllOthersFind.Add Array("仅合约单位病人", 1), "仅合约单位病人"
    End If
    
    If byt场合 <> 0 Then
        cllOthersFind.Add Array("场合", byt场合), "场合"
    End If
    
    If zl_PatiSvr_GetPatiInfsByRange(Decode(byt病人来源, 1, 0, 2, 1, 2), cllOthersFind, cllData) = False Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    str病人Ids = ""
    For i = 1 To cllData.count
        Set cllTemp = cllData(i)
        With rsPati
            .AddNew
            !排序ID = 1
            !id = cllTemp("_pati_id")
            !病人ID = cllTemp("_pati_id")
            !主页ID = cllTemp("_pati_pageid")
            !姓名 = cllTemp("_pati_name")
            !性别 = cllTemp("_pati_sex")
            !年龄 = cllTemp("_pati_age")
            !病人类型 = cllTemp("_pati_type")
            !险类 = cllTemp("_insurance_type")
            
            !门诊号 = cllTemp("_outpatient_num")
            !住院号 = cllTemp("_inpatient_num")
            !科室 = cllTemp("_pati_dept_name")
            !床号 = cllTemp("_pati_bed")
            
            !工作单位 = cllTemp("_emp_name")
            !出生日期 = cllTemp("_pati_birthdate")
            !身份证号 = cllTemp("_pati_idcard")
            !家庭地址 = cllTemp("_pati_identity")
            !在院标志 = IIf(Val(NVL(cllTemp("_is_inhspt"))) = 1, "√", "")
            !手机号 = cllTemp("_phone_number")
            
            !是否有医疗卡 = IIf(NVL(cllTemp("_medc_card_no")) <> "", "√", "")
            !就诊时间 = cllTemp("_visit_time")
            
            .Update
        End With
        str病人Ids = str病人Ids & "," & Val(NVL(cllTemp("_pati_id")))
    Next
    If str病人Ids <> "" Then str病人Ids = Mid(str病人Ids, 2)
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    If (bln显示门诊预交 Or bln显示住院预交) And str病人Ids <> "" Then
        Dim cllSurplus As Collection
        If zl_ExseSvr_GetPatiSurplusInfo(str病人Ids, cllSurplus) Then
            With rsPati
                Do While Not .EOF
                    lng病人ID = Val(NVL(!病人ID))
                    Set cllTemp = mdlPubJson.zlGetNodeObjectFromCollect(cllData, "_" & lng病人ID)
                    If Not cllTemp Is Nothing Then  '
                        If bln显示门诊预交 Then !门诊预交余额 = Format(Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_outdpst_surplus", "C"))), "####0.00;-####0.00; ;")
                        If bln显示住院预交 Then !住院预交余额 = Format(Val(NVL(mdlPubJson.zlGetNodeValueFromCollect(cllTemp, "_indpst_surplus", "C"))), "####0.00;-####0.00; ;")
                        .Update
                    End If
                    .MoveNext
                Loop
                .MoveFirst
            End With
        End If
    End If
    zlGetPatiRecordFromPatiIDCard = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGetPatiIdFromPatiIDCard(ByVal objControl As Object, ByVal strIDCard As String, ByVal strName As String, ByVal strSex As String, ByRef lngPatiId As Long, _
    Optional frmMain As Object, Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, _
    Optional ByVal blnOnlyUnitPatient As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院;3-分诊;4-挂号过滤窗体;5-医疗卡;6-入院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsSel As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strHideCols As String
    Dim byt调用场合 As Byte
    
    If byt场合 = 0 Then
        byt调用场合 = 1
    ElseIf byt场合 = 3 Then
        byt调用场合 = 2
    ElseIf byt场合 = 4 Then
        byt调用场合 = 3
    End If
    
    If zlGetPatiRecordFromPatiIDCard(strIDCard, rsTmp, strName, strSex, bln显示门诊预交, bln显示住院预交, _
        byt病人来源, blnOnlyUnitPatient, byt调用场合) = False Then Exit Function
    
    If rsTmp.RecordCount = 0 Then Exit Function
    
    strHideCols = "排序ID,ID"
    Select Case byt场合
    Case 0, 3, 4 '门诊挂号
        strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志"
    Case 1 '门诊收费
        If byt病人来源 = 2 Then
            strHideCols = strHideCols & ",病人类型,险类,门诊号,在院标志,手机号,是否有医疗卡,就诊时间"
        Else
            strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志,手机号,是否有医疗卡,就诊时间"
        End If
    Case 2 '住院
        strHideCols = strHideCols & ",病人类型,险类,在院标志,手机号,是否有医疗卡,就诊时间"
    End Select
    
    rsTmp.Sort = "排序ID,姓名"
    If frmPatiSelect.ShowSelect(frmMain, mcnOracle, glngSys, glngModul, objControl, _
        rsTmp, "病人信息查询", "", True, False, True, "", strHideCols, rsSel, blnCancel) = False Then Exit Function
    If rsSel Is Nothing Then Exit Function
    lngPatiId = NVL(rsSel!id)
    zlGetPatiIdFromPatiIDCard = lngPatiId <> 0
End Function

Public Function zlGetPatiInfoFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef rsPatiInfo_Out As ADODB.Recordset, _
    Optional frmMain As Object, Optional ByVal bln显示门诊预交 As Boolean, Optional ByVal bln显示住院预交 As Boolean, _
    Optional ByVal intDefaultDays As Integer, Optional ByVal lngDefaultCardTypeID As Long, _
    Optional ByVal byt病人来源 As Byte, Optional ByVal byt场合 As Byte, _
    Optional ByRef blnCancel As Boolean, Optional ByVal blnAddNewPati As Boolean, _
    Optional ByVal blnOnlyUnitPatient As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人姓名，获取病人信息详细情况
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   byt病人来源-0-门诊和住院,1-门诊,2-住院
    '   byt场合-调用场合,用于控制显示列：0-门诊挂号,1-门诊收费,2-住院;3-分诊;4-医疗卡;5-入院
    '   blnAddNewPati=在按病人模糊查找时是否增加新病人
    '   blnOnlyUnitPatient-只查合约单位病人
    '出参:rsPatiInfo_Out-返回病人信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-01 11:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsSel As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strHideCols As String
    Dim byt调用场合 As Byte
    
    If byt场合 = 0 Then
        byt调用场合 = 1
    ElseIf byt场合 = 3 Then
        byt调用场合 = 2
    ElseIf byt场合 = 4 Then
        byt调用场合 = 3
    End If
    
    If zlGetPatiRecordFromPatiName(strName, rsTmp, bln显示门诊预交, bln显示住院预交, _
        intDefaultDays, lngDefaultCardTypeID, byt病人来源, blnOnlyUnitPatient, , byt调用场合) = False Then Exit Function
    
    strHideCols = "排序ID,ID"
    Select Case byt场合
    Case 0 '门诊挂号
        strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志,主页ID,入院日期,出院日期"
    Case 1 '门诊收费
        If byt病人来源 = 2 Then
            strHideCols = strHideCols & ",病人类型,险类,门诊号,在院标志,手机号,是否有医疗卡,就诊时间,主页ID,入院日期,出院日期"
        Else
            strHideCols = strHideCols & ",病人类型,险类,住院号,科室,床号,在院标志,手机号,是否有医疗卡,就诊时间,主页ID,入院日期,出院日期"
        End If
    Case 2 '住院
        strHideCols = strHideCols & ",病人类型,险类,门诊号,在院标志,手机号,是否有医疗卡,就诊时间"
    End Select
    
    rsTmp.Sort = "排序ID,姓名"
    If frmPatiSelect.ShowSelect(frmMain, mcnOracle, glngSys, glngModul, objControl, _
        rsTmp, "病人信息查询", "", True, False, True, "", strHideCols, rsSel, blnCancel) = False Then Exit Function
    If rsSel Is Nothing Then Exit Function
    Set rsPatiInfo_Out = rsSel
    zlGetPatiInfoFromPatiName = True
End Function


Public Function zlPatiSvr_GetPatiExtendInfo(ByVal lng病人ID As Long, ByVal str信息名集 As String, ByRef cllExtends_Out As Collection, Optional ByVal blnNotShowErrMsg As Boolean, _
                Optional ByRef strErrMsg As String, Optional ByVal lng就诊ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息从表信息服务接口
    '入参:str信息名集-多个用逗号分离,如：医学警示,联系人2,联系人3等
    '
    '出参:cllPatiData_Out-返回病人从表信息数据集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-10-31 20:10:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiData_Out As Collection, cllSubData As Collection, cllTemp As Collection
    Dim i As Long
    
    Set cllExtends_Out = New Collection
    If zl_PatiSvr_GetPatiExtendInfo(lng病人ID, str信息名集, cllPatiData_Out, blnNotShowErrMsg, strErrMsg, lng就诊ID) = False Then Exit Function
    If cllPatiData_Out Is Nothing Then Exit Function
    If cllPatiData_Out.count = 0 Then zlPatiSvr_GetPatiExtendInfo = True: Exit Function
    For i = 1 To cllPatiData_Out.count
        Set cllSubData = cllPatiData_Out(i)
        Set cllTemp = New Collection
        cllTemp.Add cllSubData("_info_name"), "信息名"
        cllTemp.Add cllSubData("_info_value"), "信息值"
        cllTemp.Add cllSubData("_visit_id"), "就诊id"
        
        cllExtends_Out.Add cllTemp, cllTemp("信息名")
    Next
    
    zlPatiSvr_GetPatiExtendInfo = True
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
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub Class_Initialize()
    Err = 0: On Error Resume Next
    Set mobjXML = New clsXML
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjDataBase = Nothing:  Set gfrmCardMgr = Nothing
    Set mcnOracle = Nothing: Set mobjXML = Nothing
    Set mrs收费类别 = Nothing: Set mrs手机号 = Nothing
    Set mobjYLCards = Nothing: Set mobjYLCardObjs = Nothing
    Set mobjXFCards = Nothing: Set mobjXFCardObjs = Nothing
    Set mrsSystem = Nothing
    Call zlClearCache '清除缓存
End Sub

Public Function zlGetContractUnitPatiInfor(ByRef cllPati As Collection, Optional ByVal lng合同单位ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取所有合同单位病人
    '返回:获取成功返回True,否则返回False
    '出参:
    '   cllPati=病人信息，成员：clsPatientInfo对象=cllPati(_病人ID)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, cllData As Collection
    Dim objPati As clsPatientInfo, i As Long, cllTemp As Collection
    
    On Error GoTo ErrHandler
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
    
    Set cllPati = New Collection
    
    Set cllFilter = New Collection
    cllFilter.Add Array("仅合约单位病人", 1)
    If lng合同单位ID <> 0 Then cllFilter.Add Array("合同单位ID", lng合同单位ID)
    If zl_PatiSvr_GetPatiInfsByRange(2, cllFilter, cllData) = False Then Exit Function
    
    For i = 1 To cllData.count
        Set objPati = New clsPatientInfo
        Set cllTemp = cllData(i)
        objPati.病人ID = cllTemp("_pati_id")
        objPati.主页ID = Val(NVL(cllTemp("_pati_pageid")))
        objPati.门诊号 = IIf(Val(NVL(cllTemp("_outpatient_num"))) = 0, "", cllTemp("_outpatient_num"))
        objPati.住院号 = IIf(Val(NVL(cllTemp("_inpatient_num"))) = 0, "", cllTemp("_inpatient_num"))
        objPati.姓名 = cllTemp("_pati_name")
        objPati.性别 = cllTemp("_pati_sex")
        objPati.年龄 = cllTemp("_pati_age")
        
        If NVL(cllTemp("_pati_birthdate")) <> "" Then
            objPati.出生日期 = Format(CDate(cllTemp("_pati_birthdate")), "yyyy-mm-dd")
            objPati.出生时间 = Format(CDate(cllTemp("_pati_birthdate")), "hh:mm")
        End If
        
'        objPati.出生地址 = NVL(cllTemp("_pati_birthplace"))
        objPati.身份证号 = NVL(cllTemp("_pati_idcard"))
'        objPati.其他证件 = NVL(cllTemp("_cert_no_other"))
        objPati.职业 = NVL(cllTemp("_ocpt_name"))
        objPati.费别 = NVL(cllTemp("_fee_category"))
        objPati.民族 = NVL(cllTemp("_pati_nation"))
        objPati.国籍 = NVL(cllTemp("_country_name"))
        objPati.学历 = NVL(cllTemp("_pati_education"))
'        objPati.婚姻状况 = NVL(cllTemp("_pati_marital_cstatus"))
        objPati.区域 = NVL(cllTemp("_pati_area"))
        objPati.家庭地址 = NVL(cllTemp("_pat_home_addr"))
'        objPati.家庭电话 = NVL(cllTemp("_pat_home_phno"))
'        objPati.家庭邮编 = NVL(cllTemp("_pat_home_postcode"))
'        objPati.监护人 = NVL(cllTemp("_pat_grdn_name"))
'        objPati.联系人 = NVL(cllTemp("_contacts_name"))
'        objPati.联系人关系 = NVL(cllTemp("_contacts_relation"))
'        objPati.联系人地址 = NVL(cllTemp("_contacts_addr"))
'        objPati.联系人电话 = NVL(cllTemp("_contacts_phno"))
        objPati.工作单位 = NVL(cllTemp("_emp_name"))
'        objPati.工作单位电话 = NVL(cllTemp("_emp_phno"))
'        objPati.工作单位邮编 = NVL(cllTemp("_emp_postcode"))
'        objPati.工作单位开户行 = NVL(cllTemp("_emp_bank_name"))
'        objPati.工作单位开户行帐户 = NVL(cllTemp("_emp_bank_accnum"))
'        objPati.户口地址 = NVL(cllTemp("_pat_hous_addr"))
'        objPati.户口地址邮编 = NVL(cllTemp("_pat_hous_postcode"))
'        objPati.籍贯 = NVL(cllTemp("_ntvplc_name"))
        objPati.密码 = ""
        objPati.医疗付款方式编码 = NVL(cllTemp("_mdlpay_mode_code"))
        objPati.医疗付款方式 = NVL(cllTemp("_mdlpay_mode_name"))
        objPati.病人类型 = NVL(cllTemp("_pati_type"))
        objPati.就诊卡号 = ""
        objPati.手机号 = NVL(cllTemp("_phone_number"))
        objPati.险类 = Val(NVL(cllTemp("_insurance_type")))
'        objPati.险类名称 = NVL(cllTemp("_insurance_name"))
'        objPati.医保号 = NVL(cllTemp("_insurance_num"))
        objPati.身份 = NVL(cllTemp("_pati_identity"))
        objPati.合同单位ID = Val(NVL(cllTemp("_ctt_unit_id")))
        objPati.登记时间 = NVL(cllTemp("_create_time"))
        objPati.在院 = Val(NVL(cllTemp("_is_inhspt"))) = 1
        objPati.最后主页ID = objPati.主页ID
        objPati.就诊时间 = NVL(cllTemp("_visit_time"))
        If NVL(cllTemp("_visit_time")) = "" Then
            objPati.初诊 = True
        ElseIf NVL(cllTemp("_create_time")) <> "" Then  '就诊时间等于登记时间为初诊
            objPati.初诊 = CDate(NVL(cllTemp("_visit_time"))) = CDate(NVL(cllTemp("_create_time")))
        End If
        cllPati.Add objPati, "_" & objPati.病人ID
    Next
    zlGetContractUnitPatiInfor = True
    Exit Function
ErrHandler:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zlGetPatiPageInforByID(ByVal str病人ID As String, Optional ByRef objPati As clsPatientInfo, _
    Optional ByRef bln仅取最后住院 As Boolean, Optional cllData As Collection, Optional strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID获取病人病案主页信息
    '返回:合法,返回true,否则返回False
    '编制:焦博
    '日期:2019-11-20 18:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnNotShowErrMsg As Boolean
    Dim cllTemp  As Collection

    On Error GoTo errHandle
    If mobjDataBase.zlIsExistOraConnect = False Then Exit Function
     
    Set objPati = New clsPatientInfo
    
    'int查询类型-0-基本;1-基本+联系人;2-所有
    If zl_CIsSvr_GetPatiPageInfo(1, str病人ID, cllData, bln仅取最后住院, False, False, True) = False Then Exit Function
    If cllData.count = 0 Then
          strErrMsg = "未找到符合条件的病人信息，请检查！"
          Exit Function
    End If
    
    Set cllTemp = cllData(1)
    objPati.病人ID = cllTemp("_pati_id")
    objPati.主页ID = Val(NVL(cllTemp("_pati_pageid")))
    objPati.住院号 = IIf(Val(NVL(cllTemp("_inpatient_num"))) = 0, "", cllTemp("_inpatient_num"))
    objPati.姓名 = cllTemp("_pati_name")
    objPati.性别 = cllTemp("_pati_sex")
    objPati.年龄 = cllTemp("_pati_age")
    objPati.费别 = NVL(cllTemp("_fee_category"))
    objPati.医疗付款方式编码 = NVL(cllTemp("_mdlpay_mode_code"))
    objPati.医疗付款方式 = NVL(cllTemp("_mdlpay_mode_name"))
    objPati.床号 = NVL(cllTemp("_pati_bed"))
    objPati.病人类型 = NVL(cllTemp("_pati_type"))
    objPati.学历 = NVL(cllTemp("_pati_education"))
    objPati.职业 = NVL(cllTemp("_ocpt_name"))
    objPati.国籍 = NVL(cllTemp("_country_name"))
    objPati.婚姻状况 = NVL(cllTemp("_pati_marital_cstatus"))
    objPati.病人性质 = Val(NVL(cllTemp("_pati_nature")))
    objPati.当前病区ID = Val(NVL(cllTemp("_pati_wardarea_id")))
    objPati.出院科室ID = Val(NVL(cllTemp("_pati_dept_id")))
    objPati.入院日期 = NVL(cllTemp("_adta_time"))
    objPati.出院日期 = NVL(cllTemp("_adtd_time"))
    objPati.险类 = Val(NVL(cllTemp("_insurance_type")))
    objPati.挂号ID = Val(NVL(cllTemp("_rgst_id")))
    objPati.已入科 = Val(NVL(cllTemp("_in_dept"))) = 1
    objPati.最后主页ID = objPati.主页ID
    zlGetPatiPageInforByID = True
    Exit Function
errHandle:
    strErrMsg = Err.Description
End Function

Public Function ZlGetPatiInsureInfo(ByVal lng病人ID As Long, ByRef insure As Integer, Optional insureName As String, _
                                                       Optional str医保号 As String, Optional str医保密码 As String, Optional str卡号 As String, _
                                                       Optional str登记时间 As String, Optional lng病种id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人id和险类，获取病人的保险信息
    '入参: insure :险类
    '返回:返回病人医保信息
    '编制:焦博
    '日期:2019-11-20 19:43:55
    '---------------------------------------------------------------------------------------------------------------------------------------------

    ZlGetPatiInsureInfo = Zl_Patisvr_Getinsureinfo(lng病人ID, insure, insureName, str医保号, str医保密码, str卡号, str登记时间, lng病种id)
End Function

Public Function ZlPatisvrGetdeptfrombad(ByRef str科室IDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取床位状况记录里的科室id
    '出参: str科室IDs :科室id字符串,用逗号分隔
    '返回:获取科室ids成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------

    ZlPatisvrGetdeptfrombad = Zl_Patisvr_Getdeptfrombad(str科室IDs)
End Function

Public Function zlGetInputItemLength(ByVal strTableItems As String, ByRef cllColumn As Collection) As Boolean
    '---------------------------------------------------------------------------------------
    '功能 : 获取病人信息中指定字段的最大长度
    '入参 :
    '   strTableItems：表1:列1,列2|表2:列1,列2,列3|..
    '出参 :
    '   cllColumn(Collect):成员(表名,列名,列长度)=cllcolumn(表名_列名)
    '返回 :
    '编制 : 李南春
    '日期 : 2019/11/21 20:16
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllSubData As Collection, cllTmp As Collection
    Dim i As Long
    
    Set cllColumn = New Collection
    If zl_PatiSvr_GetInputItemLength(strTableItems, cllData) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
        cllTmp.Add cllSubData("_table_name"), "表名"
        cllTmp.Add cllSubData("_column_name"), "列名"
        cllTmp.Add cllSubData("_column_size"), "列长度"
        
        cllColumn.Add cllTmp, cllTmp("表名") & "_" & cllTmp("列名")
    Next
    zlGetInputItemLength = True
End Function

Public Function zlGetPatiIDFromCardNo(ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByRef lng病人ID_Out As Long, _
    Optional ByVal bln检查使用时间 As Boolean = False, Optional ByVal bln检查停用或挂失 As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据卡号来获取指定的病人ID
    '入参:lngCardTypeID-卡类别ID
    '     strCardNo-卡号
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-25 16:29:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllFind As Collection, cllPatiData As Collection
    Dim cllTemp As Collection
    On Error GoTo errHandle
    
    Set cllFind = New Collection
    cllFind.Add Array("卡类别ID", lngCardTypeID)
    cllFind.Add Array("卡号", strCardNo)
    
    
    If zl_PatiSvr_GetPatiID(cllFind, Nothing, cllPatiData, True, , bln检查使用时间, bln检查停用或挂失) = False Then Exit Function
    
    If cllPatiData Is Nothing Then Exit Function
    If cllPatiData.count = 0 Then Exit Function
    Set cllTemp = cllPatiData(1)
    
    '    pati_list[] C   1   病人列表,模糊查找时，可能存在多个
    '        cardtype_id N   1   卡类别ID
    '        pati_id N   1   病人ID:未找到时也成功，返回0
    '        card_pwd    C   1   密码
    '        pati_pageid N       主页ID
    '
    lng病人ID_Out = Val(NVL(cllTemp("_pati_id")))
    zlGetPatiIDFromCardNo = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Public Function zlGetPatiVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号ID As String, _
                ByRef cllVital As Collection, Optional ByVal blnOutPati As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------
    ' 功能 : 获取病人生命体征信息
    ' 入参 : blnOutPati-门诊病人
    ' 出参 : cllVital:体征信息(Collect)(项目,值,单位)
    ' 返回 : 返回病人的体征信息，包括项目，数值，单位
    ' 编制 : 李南春
    ' 日期 : 2019/11/18 19:48
    '---------------------------------------------------------------------------------------
    Dim cllData As Collection, cllTmp As Collection, cllSubData As Collection
    Dim i As Long
    On Error GoTo ErrHand
    Set cllVital = New Collection
    If Zl_Cissvr_GetPatiVitalSigns(lng病人ID, lng挂号ID, cllData, blnOutPati) = False Then Exit Function
    If cllData Is Nothing Then Exit Function
    If cllData.count = 0 Then Exit Function
    For i = 1 To cllData.count
        Set cllSubData = cllData(i): Set cllTmp = New Collection
        '项目单位, b.项目名称, b.记录内容
        cllTmp.Add NVL(cllSubData("_pativital_item")), "项目名称"
        cllTmp.Add NVL(cllSubData("_pativital_value")), "项目值"
        cllTmp.Add NVL(cllSubData("_pativital_unit")), "项目单位"
        cllVital.Add cllTmp, cllTmp("项目名称")
    Next
    zlGetPatiVitalSigns = True
    Exit Function
ErrHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function ZlPatisvrCheckDepositNo(ByVal lng病人ID As Long, ByRef strNo As String, Optional intOcc As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断预交NO是否存在"病人结算异常记录"中
    '出参:strNo-预交单据号
    '       intOcc-病人结算异常记录.操作情景
    '返回:传入的NO号存在"病人结算异常记录"中返回true,否则返回False
    '编制:焦博
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlPatisvrCheckDepositNo = Zl_Patisvr_Checkdepositno(lng病人ID, strNo, intOcc)
End Function

Public Function ZlCissvrCheckDepositNo(ByVal lng病人ID As Long, ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断预交NO是否存在"病人结算异常记录"中
    '入参:strNo-预交单据号
    '返回:传入的NO号存在"病人结算异常记录"中返回true,否则返回False
    '编制:焦博
    '日期:2019-12-2 13:50:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ZlCissvrCheckDepositNo = Zl_Cissvr_Checkdepositno(lng病人ID, strNo)
End Function
