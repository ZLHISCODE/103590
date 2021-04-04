Attribute VB_Name = "mdlBaseItem"
Option Explicit
Public gbln使用中医 As Boolean
Public gbln购买中医 As Boolean
Public gstr医价接口编号 As String
Public gbln允许医价收费项目 As Boolean
Public gbln从项汇总折扣  As Boolean
'外挂功能
Public gobjPlugIn As Object
Public gblnMyStyle As Boolean
Public gstrMatchMode As String
Public gbytCode As Byte
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--定义系统参数
'问题:27990
Private Type Ty_System_Para
     byt药品名称显示 As Byte   '药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
     byt输入药品显示 As Byte  '输入药品显示（通过输入简码方式进入选择器时药品名称的显示）：0-按输入匹配显示，1-固定显示通用名和商品名
End Type
Public gTy_System_Para As Ty_System_Para
Public gblnFeeKindCode As Boolean
Public gstr药品价格等级 As String
Public gstr卫材价格等级 As String
Public gstr普通价格等级 As String
'Windows风格----------------------------------
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
  
Public Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public gstrLike As String
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'设定一个窗体捕获鼠标，即所有鼠标输入消息都发往该窗体
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'取消鼠标捕获
Public Declare Function ReleaseCapture Lib "user32" () As Long

'IP地址格式检查
Public Declare Function inet_addr Lib "ws2_32" (ByVal lpszAddress As String) As Long
Public Const INADDR_NONE = &HFFFFFFFF

'系统方案设置----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const GCST_INVALIDCHAR = " '"    '对于输入的无效字符
Public gobjCustAcc As Object

Public Enum EditMode 'medit方式  取值为：0、新增；1、修改；2、调价；3、执行科室、4、从属项目、5、批量修改执行科室
    EditNew = 0
    EditModify = 1
    EditRaise = 2
    EditDept = 3
    EditSlave = 4
    EditCopy = 5
End Enum
Public gobjNurseIntegrate As Object  '整体护理接口对象
Public gobjRIS As Object                    '新网RIS接口对象
Public Enum RISBaseItemOper                 '新网RIS基础数据操作类型：1-新增；2-修改；3-删除
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '新网RIS基础数据类型：3：用户(人员）
    Personnel = 3
End Enum

'本地日志模块
Private mobjFso As New FileSystemObject '文件对象

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub SetFormVisible(ByVal new_Hwnd As Long)
    '隐藏窗体最大最小按钮
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 Or WS_SYSMENU Or &H20000
End Sub

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'功能：初始化新网接口部件
'参数：blnMsg－创建失败时是否提示
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'功能：判断当前屏幕鼠标是否在指定窗口的显示区域内
    Dim vRect As RECT, vPos As POINTAPI
    
    vPos = zlControl.GetCursorPosition
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Function MoveSpecialChar(ByVal strInputString As String) As String
    '1 去除一般字符: " '_%?"，把_%?转换为对应的全角字符
    '2 去除特殊字符:退格、制表、换行、回车
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intASC As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '允许转换的字符
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "？"
                Case "%"
                    strTmp = strTmp & "％"
                Case "_"
                    strTmp = strTmp & "＿"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intASC = Asc(Mid(strText, n, 1))
        Select Case intASC
            Case 8, 9, 10, 13, 32
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

Public Sub 改变编码(nodParent As Node, int舍去长度 As Integer, str新增长度 As String)
'功能:改变树形列表各节点的标题中编码的值
'参数:nodParent         要改变编码的起始节点
'     int舍去长度       编码中舍去长度
'     str新增长度       编码中新增部分

    Dim nod As Node
    '它是下级也要改变编码
    If nodParent.Children > 0 Then
        Set nod = nodParent.Child
        Do While Not (nod Is Nothing)
            nod.Text = "【" & str新增长度 & Mid(nod.Text, int舍去长度 + 2)
            改变编码 nod, int舍去长度, str新增长度
            Set nod = nod.Next
        Loop
    End If
End Sub

Public Function GetRoot(ByVal nod As Node) As Node
'功能：读出任意节点的根节点
    Dim nodTemp As Node
    
    If nod Is Nothing Then Exit Function
    Set nodTemp = nod
    Do Until nodTemp.Parent Is Nothing
        Set nodTemp = nodTemp.Parent
    Loop
    Set GetRoot = nodTemp
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = "-") As String
'参数：cmbTemp  准备获取数据的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        '直接返回整个字符串
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            '圆点之前
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = "-")
'参数：cmbTemp  准备设置的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = cmbTemp.Text Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '已经找到
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Sub 调查报盘(frmParent As Form)
    MsgBox "请运行病案系统的人员管理。", vbInformation, gstrSysName
End Sub


Public Function GetPictureInfo(picTemp As StdPicture, Optional strBitmap As String = "") As String
'获得一张图片的信息
    Dim hFile As Integer
    Dim FileHeader As BITMAPFILEHEADER
    Dim InfoHeader As BITMAPINFOHEADER
    
    If picTemp.Handle = 0 Then
        GetPictureInfo = "无照片"
        Exit Function
    End If
    
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    
    If strBitmap = "" Then
        '产生临时文件
        strPath = Space(256): strFile = Space(256)
        GetTempPath 256, strPath
        strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
        
        GetTempFileName strPath, "pic", 0, strFile
        strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    
        SavePicture picTemp, strFile
    Else
        '直接使用现在文件
        strFile = strBitmap
    End If
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
      Get #hFile, , FileHeader
      Get #hFile, , InfoHeader
    Close #hFile
    
    If strBitmap = "" Then
        '删除临时文件
        Kill strFile
    End If
    
    If InfoHeader.biBitCount > 8 Then
         GetPictureInfo = InfoHeader.biWidth & "×" & InfoHeader.biHeight & " " & InfoHeader.biBitCount & "位色"
    Else
         GetPictureInfo = InfoHeader.biWidth & "×" & InfoHeader.biHeight & " " & 2 ^ InfoHeader.biBitCount & "色"
    End If
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

Public Sub InitSystemPara()
    '个人全局参数
    '-------------------------------------------------------------------------------------------------
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    '当不输类别时,输入费用项目时,首位当作类别简码
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1"
    gstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gbln从项汇总折扣 = zlDatabase.GetPara(93, glngSys) = "1"
    '问题:27990
    With gTy_System_Para
        .byt输入药品显示 = Val(zlDatabase.GetPara("输入药品显示")) '0-按输入匹配显示，1-固定显示通用名和商品名
        .byt药品名称显示 = Val(zlDatabase.GetPara("药品名称显示"))  '：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    End With
End Sub
Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 编码, 名称, 简码 From 收费项目类别"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "获取收费类别")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:提示消息框
    '入参:strMsgInfor-提示信息
    '        blnYesNo-是否提供YES或NO按钮
    '出参:
    '返回:blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '编制:刘兴洪
    '日期:2010-08-27 16:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function MedicalTeamPatients(ByVal lngTeamID As Long, ByVal lngMemberID As Long) As String
'----------------------------------------------------------------------
'功能： 列出医疗小组医生的病人信息
'参数： lngTeamID: 医疗小组ID
'       lngMemberID: 医生ID
'返回： 病人信息字符串
'----------------------------------------------------------------------
    Dim strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Select a.病人id, a.住院号, a.出院病床, b.姓名" & vbNewLine & _
              "From 病案主页 a, 病人信息 b " & vbNewLine & _
              "Where a.住院医师 = (Select 姓名" & vbNewLine & _
              "              From 人员表" & vbNewLine & _
              "              Where ID = [2] And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)) And" & vbNewLine & _
              "      a.医疗小组id = [1] and a.病人id=b.病人id and a.主页id=b.主页id and b.在院=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "医疗小组医生病人信息", lngTeamID, lngMemberID)
    With rsTmp
        For i = 1 To .RecordCount
            strMess = strMess & "姓名：" & !姓名 & "；" & vbTab & _
                      "住院号：" & IIF(IsNull(!住院号), "", !住院号) & "；" & vbTab & _
                      "床号：" & IIF(IsNull(!出院病床), "", !出院病床) & vbTab & vbNewLine
            .MoveNext
        Next
    End With
    MedicalTeamPatients = strMess
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDeptPermission(ByVal lngOperationID As Long, Optional ByVal lngDeptID As Long) As Boolean
'功能: 检查部门权限
'lngOperationID: 要操作的人员ID
'lngDeptID: 要操作人员的部门ID
'返回: True有权限, False无权限
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    If lngDeptID = 0 Then
        gstrSQL = "Select Count(*) Rec From 部门人员 " & _
                  "Where 人员id = [2] And [3] In (Select 部门id From 部门人员 Where 人员id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查人员的部门权限", glngUserId, lngOperationID, lngDeptID)
    Else
        gstrSQL = "Select ID " & _
                  "From 部门表 " & _
                  "  Start With ID In (Select 部门id From 部门人员 Where 人员id = [1]) " & _
                  "  Connect By Prior ID = 上级id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查人员的部门权限", glngUserId)
        Do While Not rsTmp.EOF
            If rsTmp!ID = lngDeptID Then
                CheckDeptPermission = True
                Exit Function
            End If
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End If
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlGetBillFormatRec(ByVal strReportCode As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定报表的打印格式
    '入参:strReportCode-报表名称
    '返回:报表打印格式的记录集
    '编制:刘兴洪
    '日期:2015-06-10 11:43:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  " & _
    "   From Dual " & _
    "   Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1] " & _
    "   Order by 序号"
    Set zlGetBillFormatRec = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strReportCode)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function GetPriceGrade(ByRef str药品价格等级 As String, _
    ByRef str卫材价格等级 As String, ByRef str普通价格等级 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前站点价格等级
    '入参:
    '返回:价格等级获取成功返回True，否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    str药品价格等级 = "": str卫材价格等级 = "": str普通价格等级 = ""
    strSQL = "" & _
        "Select Max(Decode(b.是否适用药品, 1, 价格等级, Null)) As 药品等级," & vbNewLine & _
        "       Max(Decode(b.是否适用卫材, 1, 价格等级, Null)) As 卫材等级," & vbNewLine & _
        "       Max(Decode(b.是否适用普通项目, 1, 价格等级, Null)) As 普通等级" & vbNewLine & _
        "From 收费价格等级应用 A, 收费价格等级 B" & vbNewLine & _
        "Where a.价格等级 = b.名称 And a.性质 = 0 And a.站点 = [1]" & vbNewLine & _
        "      And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取价格等级", gstrNodeNo)
    If Not rsTemp.EOF Then
        str药品价格等级 = Nvl(rsTemp!药品等级)
        str卫材价格等级 = Nvl(rsTemp!卫材等级)
        str普通价格等级 = Nvl(rsTemp!普通等级)
    End If
    GetPriceGrade = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    '写一行日志，如果内容中有回车,换行符，替换为<CR><LF>
    '日志保存在当前目录下的[应用程序名称]Log目录下，文件名为日期.txt,默认保存7天的日志。

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '日志路径，文件名，配置文件名
    Dim strLogSaveDays As String '日志保留天数
    Dim dblFreeSpace As Double   '剩余空间
    Dim strDelOldFile As String  '过期文件
    Dim objFile As File

    If Val(OS.IniRead("LOG", "OPENLOG", App.Path & "\CONFIG.INI")) = 0 Then Exit Sub
    '始终保存日志
    '2、清除过期日志
    strLogSaveDays = "7"  '保留7天的日志
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\日志*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    '3、空间是否足够
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '空间不足，不写日志,产生一个警告文件
        If Not mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\空间不足.txt", True)
        Exit Sub
    Else
        '清除警告文件
        If mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.DeleteFile(strLogPath & "\空间不足.txt", True)
    End If
    '4、写入日志行
    strLogFile = strLogPath & "\日志" & Format(Now, "yyyyMMdd") & ".log"

    Call SaveLog(strLogFile, strLogTxt)

End Sub

Public Sub SaveLog(ByVal strFileName As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFileName) Then Call mobjFso.CreateTextFile(strFileName)
        Set objStream = mobjFso.OpenTextFile(strFileName, ForAppending)
        If strDate = "" Then
            strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
            objStream.WriteLine (strDate & Chr(&H9) & strInput)
        Else
            objStream.WriteLine (strInput)
        End If
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '获取剩余空间
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Function FuncGetStr(ByVal strVal As String) As String
    strVal = Replace(strVal, vbTab, "")
    strVal = Replace(strVal, vbCrLf, "")
    strVal = Replace(strVal, Chr(10), "")
    strVal = Replace(strVal, "'", "''")
    strVal = Replace(strVal, " ", "")
    FuncGetStr = Trim(strVal)
End Function

Public Function IsPriceGradeEnabled() As Boolean
    '是否启用了价格等级
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = "Select 1 From 收费价格等级应用 Where Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断是否启用了价格等级")
    IsPriceGradeEnabled = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsIPAddress(ByVal strAddress As String) As Boolean
'功能：判断输入的Ip地址格式是否合法
    If inet_addr(strAddress) <> INADDR_NONE Then
        IsIPAddress = True
    Else
        IsIPAddress = False
    End If
End Function

Public Function InitNurseIntegrate(Optional blnMsg As Boolean = False) As Boolean
'判断如果整体护理部件为空就初始化
    If gobjNurseIntegrate Is Nothing Then
        On Error Resume Next
        Set gobjNurseIntegrate = CreateObject("zlNurseIntegrate.clsNurseIntegrate")
        If Not gobjNurseIntegrate Is Nothing Then
            If gobjNurseIntegrate.zlInitCommon(gcnOracle, gstrDbUser) = False Then
                Set gobjNurseIntegrate = Nothing
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    If blnMsg = True And gobjNurseIntegrate Is Nothing Then
        MsgBox "整体护理接口部件：zlNurseIntegrate  创建失败！", vbInformation, gstrSysName
    End If
    InitNurseIntegrate = Not gobjNurseIntegrate Is Nothing
End Function
