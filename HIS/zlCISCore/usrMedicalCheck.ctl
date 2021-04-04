VERSION 5.00
Begin VB.UserControl usrMedicalCheck 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ScaleHeight     =   4710
   ScaleWidth      =   7350
   Begin zl9CISCore.VsfGrid vsf 
      Height          =   1530
      Left            =   1740
      TabIndex        =   0
      Top             =   345
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2699
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      Height          =   435
      Left            =   270
      Top             =   495
      Width           =   585
   End
End
Attribute VB_Name = "usrMedicalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strSQL As String
Private rsTmp As New ADODB.Recordset
Private mlng病历id As Long                      '外界传入
Private mlng医嘱id As Long                      '外界传入
Private mblnMode As Boolean '为真是表示是用户进行的编辑，这时才赋值
Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mblnMoved As Boolean '数据是否已转出
Private mblnModify As Boolean
Private mblnCommon As Boolean
Private mstrSQL As String
Private mlngLoop As Long
Private mblnLoaded As Boolean

Private Enum mCol

    项目 = 1
    结果
    单位
    类型
    数值域
    正常域
    初始值
    表示法
    空值文字
End Enum

Private Enum COLOR
    
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    
End Enum

Private mobjParentObject As Object

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:补充表格控件的空行
    '参数:objVsf 要补充行的表格控件对象
    '返回:若成功返回True,否则返回 False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo errHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.隐藏所有的线
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.重新计算需要的纵线
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.重新计算需要的横线
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
errHand:
    
End Function

Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '正小数
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

'Private Function ShowOpenList(Optional strText As String, Optional blnWhere As Boolean = False) As Byte
'    '-----------------------------------------------------------------------------------------
'    '功能:打开列表结构的诊疗检验标本数据
'    '返回:出错返回2;成功返回1;取消返回0
'    '-----------------------------------------------------------------------------------------
'    Dim strLvw As String
'    Dim sglX As Single
'    Dim sglY As Single
'    Dim rs As New ADODB.Recordset
'    Dim strSQL As String
'
'    On Error GoTo errHand
'
'    strLvw = "编码,900,0,1;取值,1800,0,0;结果标志,900,0,0"
'
'    ShowOpenList = 2
'
'    strSQL = "SELECT ROWNUM AS ID,编码,取值,DECODE(结果标志,1,'1-正常',2,'2-偏低',3,'3-偏高',4,'4-阳性',5,'5-阴性','') AS 结果标志 FROM 检验项目取值 A WHERE 项目id=[1]"
'    If blnWhere Then
'        strSQL = strSQL & " AND (A.编码 Like [2] OR A.取值 Like [2])"
'    End If
'
'    Set rs = OpenSQLRecord(strSQL, "检验结果", CLng(Val(vsf.RowData(vsf.Row))), "%" & strText & "%")
'
'    If rs.BOF Then
'
'        ShowOpenList = 0
'
'        Exit Function
'    End If
'
'    If rs.RecordCount = 1 And blnWhere Then GoTo Over
'
'    Call CalcPosition(sglX, sglY, vsf)
'
'    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 3600, 4500, "检验取值选择", "请从下表中选择一个取值") Then
'        GoTo Over
'    End If
'
'    Exit Function
'
'Over:
''    vsf.EditText = zlCommFun.Nvl(rs("取值").Value)
''    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("取值").Value)
''    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("取值").Value)
''    vsf.TextMatrix(vsf.Row, mCol.结果标志) = zlCommFun.Nvl(rs("结果标志").Value)
'
'    ShowOpenList = 1
'
'    Exit Function
'
'errHand:
'    If ErrCenter = 1 Then Resume
'End Function

'公共方法、属性
Public Property Get DispMode() As Boolean
    '是否为显示模式
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    ShowUsrControl mlng医嘱id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
    End If
    
End Property

Public Property Get ID病人病历() As Long
    '返回病人病历ID
    
    ID病人病历 = mlng病历id
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
    '设置病人病历ID,并检查该病历是不是存在
    
    mlng病历id = New_ID病人病历
    ShowUsrControl mlng医嘱id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_医嘱ID As Long, ByVal New_发送号)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlng医嘱id = New_医嘱ID
    strSQL = "SELECT DECODE(相关id,NULL,ID,相关id) AS ID FROM 病人医嘱记录 WHERE ID=[1]"
    
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rs = OpenSQLRecord(strSQL, "检验报告专用纸", mlng医嘱id)
    If rs.BOF = False Then mlng医嘱id = rs("ID").Value
        
        
    mblnModify = True
        
    strSQL = "SELECT DISTINCT A.执行状态,D.项目类别 FROM 病人医嘱发送 A,病人医嘱记录 B,检验报告项目 C,检验项目 D WHERE B.诊疗项目id=C.诊疗项目id AND C.报告项目id=D.诊治项目id AND A.医嘱id=B.ID AND B.相关id=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rs = OpenSQLRecord(strSQL, "检验报告专用纸", mlng医嘱id)
    If rs.BOF = False Then mblnModify = (rs("执行状态").Value <> 1)
                    
    vsf.Visible = False
        
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "项目", 1500, 1
        .NewColumn "结果", 3300, 1, , 1
        .NewColumn "单位", 600, 1
        .NewColumn "类型", 0, 1
        .NewColumn "数值域", 0, 1
        .NewColumn "正常域", 0, 1
        .NewColumn "初始值", 0, 1
        .NewColumn "表示法", 0, 1
        .NewColumn "空值文字", 0, 1
    
        .NewColumn "", 255, 4
        
        .ExtendLastCol = True
        .Body.Appearance = flexFlat
        
        .Body.BorderStyle = flexBorderNone
        .Body.BackColorFixed = .Body.BackColor
        .Body.ColHidden(mCol.类型) = True
        .Body.ColHidden(mCol.数值域) = True
        .Body.ColHidden(mCol.正常域) = True
        .Body.ColHidden(mCol.初始值) = True
        .Body.ColHidden(mCol.表示法) = True
        .Body.ColHidden(mCol.空值文字) = True
        
        .FixedCols = 1
        
        .Cell(flexcpFontBold, 0, 0, 0, vsf.Cols - 1) = True
        
        .SelectMode = True
        
        .Visible = True
    End With
    
    If mblnModify = False Then
        vsf.EditMode(mCol.项目) = 0
        vsf.EditMode(mCol.结果) = 0
        vsf.EditMode(mCol.单位) = 0
        vsf.ComboList(mCol.结果) = ""

    End If
End Sub

Public Property Get Get医嘱id() As Long
        
    Get医嘱id = mlng医嘱id
        
End Property

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '设置错误描述及错误号
    '如果lngErrNum=-1 表示 控件自己定义的错误
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
    '返回最后一次的错误号
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
    '返回最后一次错误描述字符串
    ReturnErrDescription = mReturnErrDescription
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function CheckStrValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByRef strError As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
        
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        strError = "所输入内容含有非法字符。"
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            strError = "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。"
            Exit Function
        End If
    End If
    
    CheckStrValid = True
End Function

'------------------------------------------------------------------------------------------------------------

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------
    '功能：外部调用显示手术概要的过程
    '------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
        
    mDispMode = Not blnEditMode
    
    '按逻辑应先初始控件
    Call InitData
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub

    Call ReadData
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub SetgcnOracle()
    '-------------------------------------------------------------------------------------------------
    '接口
    '-------------------------------------------------------------------------------------------------
    
    Call InitCommon(gcnOracle)
End Sub

Private Sub InitData()
    '初始化窗体
    
    Dim strTmp As String
    
    On Error GoTo ErrHandle
        
    If Not gcnOracle Is Nothing Then
        If Not gcnOracle.State <> adStateOpen Then
            If Ambient.UserMode = True Then

            End If
        End If
    
    End If
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：读出数据库里的数据
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strDec As String
    
    On Error GoTo ErrHandle
    
    mblnLoaded = True
    
    '连接检查
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Function
                                        
    
    '清除数据
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    '读取并装载数据
    
    mstrSQL = "Select 1 From 病人病历记录 a,病人病历内容 b Where a.书写日期 Is Not Null And a.id=b.病历记录id and b.ID=[1]"
    Set rs = OpenSQLRecord(mstrSQL, "体检检查报告", mlng病历id)
    If rs.BOF Then mlng病历id = 0
    
    If mlng病历id > 0 Then

        mstrSQL = "Select A.ID,a.中文名 As 项目,b.所见内容 As 结果,a.单位,a.类型,a.数值域,a.正常域,a.初始值,a.表示法,a.空值文字 " & _
                    "From 诊治所见项目 a,病人病历所见单 b " & _
                    "Where a.ID=b.所见项id and b.病历id=[1] "
                                
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "病人病历所见单", "H病人病历所见单")
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "体检检查报告", mlng病历id)
    Else
        
        mstrSQL = "Select a.ID,a.中文名 As 项目,Decode(a.初始值,Null,a.空值文字,a.初始值) As 结果,a.单位,a.类型,a.数值域,a.正常域,a.初始值,a.表示法,a.空值文字 " & _
                    "From 诊治所见项目 a,病历所见单 b,病历元素目录 c,病人医嘱记录 d " & _
                    "Where c.类型=-1 and c.id=b.元素id and b.行=d.诊疗项目id and a.id=b.所见项id and d.ID=[1] " & _
                    "Order By b.控件号 "
        
        If mblnMoved Then
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rs = OpenSQLRecord(mstrSQL, "体检检查报告", mlng医嘱id)
        
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
                    
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.项目) = zlCommFun.NVL(rs("项目").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.结果) = zlCommFun.NVL(rs("结果").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.类型) = zlCommFun.NVL(rs("类型").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.数值域) = zlCommFun.NVL(rs("数值域").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.正常域) = zlCommFun.NVL(rs("正常域").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.初始值) = zlCommFun.NVL(rs("初始值").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.表示法) = zlCommFun.NVL(rs("表示法").Value)
            vsf.TextMatrix(vsf.Rows - 1, mCol.空值文字) = zlCommFun.NVL(rs("空值文字").Value)
                                            
            rs.MoveNext
        Loop
    End If
    
    '自动设置高度
    If (vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30) < UserControl.Height Then
        UserControl.Height = vsf.Rows * (vsf.Body.RowHeight(0) + 15) + 30
    End If
    
    vsf.Cell(flexcpForeColor, 1, mCol.结果, vsf.Rows - 1, mCol.结果) = COLOR.兰色
    
    Call vsf.Body.AutoSize(mCol.结果, mCol.结果)
    
    mblnLoaded = False
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    
End Function

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpText, 1, mCol.结果, vsf.Rows - 1, mCol.结果) = ""
    
End Sub

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim strValue As String
    
    ReDim Preserve strSQL(0 To vsf.Rows)
    
'    strSQL(0) = "ZL_病人病历内容_DELETE(" & lng病历ID & ")"
    
    If mblnModify Then
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                
                strValue = vsf.TextMatrix(lngLoop, mCol.结果)
                
                strSQL(lngLoop) = "ZL_病人病历所见单_SAVE(" & lng病历ID & "," & _
                                                            lngLoop & "," & _
                                                            "2,'" & _
                                                            vsf.TextMatrix(lngLoop, mCol.项目) & "'," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            "NULL," & _
                                                            Val(vsf.RowData(lngLoop)) & "," & _
                                                            "0,'" & _
                                                            vsf.TextMatrix(lngLoop, mCol.单位) & "','" & _
                                                            strValue & "'" & _
                                                            ")"
            End If
        Next
            
        strTmp = ""
        For lngLoop = 0 To UBound(strSQL)
            If strSQL(lngLoop) <> "" Then
                If strTmp = "" Then
                    strTmp = strSQL(lngLoop)
                Else
                    strTmp = strTmp & Chr(9) & strSQL(lngLoop)
                End If
            End If
        Next
        
        '返回SQL语句
        strReturnSQL = strTmp
    End If
    
    SaveData = True
    
End Function

Private Sub UserControl_Initialize()

    '初始化控件属性
    
    On Error GoTo ErrHandle
    
    vsf.ComboEdit = True
    vsf.SelEdit = True
    vsf.Body.WordWrap = True
    
    Exit Sub
    
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
    
    '功能：判断当前运行程序是否在VB的工程环境中
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub UserControl_InitProperties()
    '初始病人病历为0
    mlng病历id = 0
    mDispMode = True
    mblnMoved = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", True)
    mblnMoved = PropBag.ReadProperty("DataMoved", False)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
End Sub

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Let Locked(ByVal vData As Boolean)
    MsgBox "a"
End Property

Private Sub UserControl_Resize()
    
    On Error Resume Next
    
    With shp
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    
    With vsf
        .Left = 15
        .Top = 15
        .Width = UserControl.Width - .Left - 15
        .Height = UserControl.Height - .Top - 15
    End With

End Sub

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
    
    On Error Resume Next
    
    Set mobjParentObject = Nothing
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, True)
    Call PropBag.WriteProperty("DataMoved", mblnMoved, False)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
End Sub

Private Sub UserControl_Show()
    Dim objCtl As Control
         
    '只在运行时显示
    
    On Error Resume Next
    
    If Ambient.UserMode = True And InDesign = False Then
        If mDispMode Then
            For Each objCtl In Controls
                If UCase(TypeName(objCtl)) <> UCase("ImageList") Then
                    objCtl.Enabled = False
                End If
            Next
        End If
    End If
    
    If mblnLoaded = False Then InitData
        
    
End Sub

Public Property Get Text() As String
    '为每一个控件加上文本转储属性
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSvrKey As Long
'
'    '通过用户输入的内容得到转储文本
'    strTmp = "检验报告：" & vbCrLf
'    If mblnCommon Then
'        For lngLoop = 0 To vsf.Rows - 1
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.检验项目) & Space(50), 1, 50)
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.检验结果) & Space(20), 1, 20)
'            strTmp = strTmp & MidB(vsf.TextMatrix(lngLoop, mCol.结果标志) & Space(20), 1, 20)
'            strTmp = strTmp & vsf.TextMatrix(lngLoop, mCol.结果参考) & vbCrLf
'            If lngLoop = 0 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
'            If lngLoop = vsf.Rows - 1 Then strTmp = strTmp & "------------------------------------------------------------------------------------------" & vbCrLf
'        Next
'    Else
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.细菌) & Space(40), 1, 40)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.计数) & Space(20), 1, 20)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.抗生素) & Space(40), 1, 40)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.结果) & Space(20), 1, 20)
'        strTmp = strTmp & MidB(vsf2.TextMatrix(1, mCol.类型) & Space(20), 1, 20)
'        strTmp = strTmp & vsf2.TextMatrix(1, mCol.培养描述) & vbCrLf
'        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
'
'        For lngLoop = 2 To vsf2.Rows - 1
'
'            If strSvrKey <> vsf2.RowData(lngLoop) Then
'
'                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.细菌) & Space(40), 1, 40)
'                strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.计数) & Space(20), 1, 20)
'
'            End If
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.抗生素) & Space(40), 1, 40)
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.结果) & Space(20), 1, 20)
'            strTmp = strTmp & MidB(vsf2.TextMatrix(lngLoop, mCol.类型) & Space(20), 1, 20)
'
'            If strSvrKey <> vsf2.RowData(lngLoop) Then
'
'                strSvrKey = vsf2.RowData(lngLoop)
'                strTmp = strTmp & vsf2.TextMatrix(lngLoop, mCol.培养描述)
'
'            End If
'            strTmp = strTmp & vbCrLf
'        Next
'
'        strTmp = strTmp & "-----------------------------------------------------------------------------------------------------" & vbCrLf
'    End If
'
    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Trim(vsf.TextMatrix(Row, mCol.结果)) = "" Then
        vsf.TextMatrix(Row, mCol.结果) = vsf.TextMatrix(Row, mCol.空值文字)
    End If
    
    Call vsf.Body.AutoSize(mCol.结果, mCol.结果)
    
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.结果
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    Dim strTmp As String
    Dim aryTmp As Variant
    
    On Error Resume Next
    
    If mblnModify = False Then Exit Sub

    If NewCol = mCol.结果 Then
        
        '0-文本,1-上下,2-下拉,3-复选,4-单选;5-指针(该项目的描述值由某个数据表或视图来描述，本功能暂不提供)
        If Val(vsf.TextMatrix(NewRow, mCol.类型)) = 1 Then
            Select Case Val(vsf.TextMatrix(NewRow, mCol.表示法))
'            Case 2, 4
'                strTmp = vsf.TextMatrix(NewRow, mCol.数值域)
'
'                aryTmp = Split(strTmp, ";")
'                strTmp = " |" & Join(aryTmp, "|")
'
'                vsf.ComboList(mCol.结果) = strTmp
            Case 2, 4, 3    '复选
                vsf.ComboList(mCol.结果) = "..."
            Case Else
                vsf.ComboList(mCol.结果) = ""
            End Select
        Else
            vsf.ComboList(mCol.结果) = ""
        End If
        
        vsf.VsfComboList = vsf.ComboList(mCol.结果)
    End If

End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim sglX As Single
    Dim sglY As Single
    Dim strText As String
    Dim strTmp As String
    Dim aryTmp As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    rs.Fields.Append "ID", adBigInt
    rs.Fields.Append "末级", adBigInt
    rs.Fields.Append "选择", adBigInt
    rs.Fields.Append "名称", adVarChar, 100
    rs.Open
    
    strTmp = vsf.TextMatrix(Row, mCol.数值域)
    aryTmp = Split(strTmp, ";")
    
    strText = ";" & vsf.TextMatrix(Row, mCol.结果) & ";"
    
    For lngLoop = LBound(aryTmp) To UBound(aryTmp)
        
        If InStr(strText, ";" & aryTmp(lngLoop) & ";") > 0 Then
            rs.AddNew
            rs("ID").Value = lngLoop
            rs("末级").Value = 1
            rs("选择").Value = 1
            rs("名称").Value = CStr(aryTmp(lngLoop))
                        
        Else
            
            rs.AddNew
            rs("ID").Value = lngLoop
            rs("末级").Value = 1
            rs("选择").Value = 0
            rs("名称").Value = CStr(aryTmp(lngLoop))
            
        End If
        
    Next
        
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    
    Call CalcPosition(sglX, sglY, vsf)
    
    Dim blnMuli As Boolean
    
    If Val(vsf.TextMatrix(Row, mCol.表示法)) = 3 Then
        '多选
        blnMuli = True
    End If
    
    If frmSelectDialog.ShowSelect(Nothing, 2, rs, "名称,3300,0,0", "请从下面选择多个项目,然后回车或双击退出", sglX + 60, sglY + 30, 6000, vsf.Body.ColWidth(Col), 300, , "检查结果选择", , False, blnMuli) Then
        
        vsf.TextMatrix(Row, Col) = ""
        
        If blnMuli Then
            rs.Filter = ""
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("名称").Value) & ";"
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(Row, Col) <> "" Then vsf.TextMatrix(Row, Col) = Mid(vsf.TextMatrix(Row, Col), 1, Len(vsf.TextMatrix(Row, Col)) - 1)
                            
            End If
        Else
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        End If
        
        Call vsf.Body.AutoSize(mCol.结果, mCol.结果)
    End If
End Sub

Private Sub vsf_ChangeEdit(ByVal Row As Long, ByVal Col As Long)
    vsf.TextMatrix(Row, Col) = vsf.EditText
    Call vsf.Body.AutoSize(mCol.结果, mCol.结果)
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If mblnModify = False Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    
    '0-数值；1-文字；2-日期；(3-逻辑)
    
    Select Case Val(vsf.TextMatrix(vsf.Row, mCol.类型))
    Case 0
        KeyAscii = FilterKeyAscii(KeyAscii, 2)
    Case Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End Select
    
    
End Sub

'数据是否转出
Public Property Get DataMoved() As Boolean
    DataMoved = mblnMoved
End Property

Public Property Let DataMoved(ByVal vNewValue As Boolean)
    mblnMoved = vNewValue
End Property

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim aryTmp As Variant
    
    If Col <> mCol.结果 Then Exit Sub
    If Trim(vsf.EditText) = "" Then Exit Sub
    
    '0-数值；1-文字；2-日期；(3-逻辑)
    If Val(vsf.TextMatrix(Row, mCol.类型)) = 0 Then
        If vsf.TextMatrix(Row, mCol.数值域) <> "" Then
            
            aryTmp = Split(Trim(vsf.TextMatrix(Row, mCol.数值域)), ";")
            If Val(vsf.EditText) < aryTmp(0) Then
                vsf.EditText = ""
                Cancel = True
            End If
            
            If UBound(aryTmp) > 0 Then
                If Val(vsf.EditText) > aryTmp(1) Then
                    vsf.EditText = ""
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub
