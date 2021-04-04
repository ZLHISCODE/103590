Attribute VB_Name = "mdlDue"
Option Explicit

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long
Public gstrAviPath As String
Public gstrVersion As String
Public gstrMatchMethod As String

Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public gstrIme As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum mAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Private Type POINTAPI
        X As Long
        Y As Long
End Type
 
'切换到指定的输入法。
'Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
'Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
'Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
'Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'获取指定输入法所在Layout,参数为0时表示当前输入法。
'Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'获取当前输入法所在Layout名
'Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'根据输入法Layout名将该输入法切换到输入法切换顺序的最前头(重新启动后无效),flags参数=KLF_REORDER
'Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type SystemParameter
    int简码方式 As Integer
    bln个性化风格 As Boolean               '使用个性化风格
    Para_输入方式 As String             ''第1位1-全数字只查编码,第2位1-全字母只查简码,在HIS基础参数中设置
    bln存在站点 As Boolean      '是否存在站点管理
End Type
Public Enum g小数类型
    g_数量 = 0
    g_成本价
    g_售价
    g_金额
End Enum

Private Type m_小数位
    数量小数 As Integer
    成本价小数 As Integer
    零售价小数 As Integer
    金额小数 As Integer
End Type

Public g_小数位数 As m_小数位

'小数格式化串
Public Type g_FmtString
    FM_数量 As String
    FM_成本价 As String
    FM_零售价 As String
    FM_金额 As String
End Type

Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
Public gSystemPara As SystemParameter


'系统方案设置----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48

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

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

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


Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo ErrHand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
ErrHand:
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = "" & rsTmp!简码
        UserInfo.姓名 = "" & rsTmp!姓名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'以下函数没有使用 by lesfeng 2009-12-2 性能优化
'Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
'    '功能描述：读取指定表的本级编码的最大长度
'    '输入参数：本级ID，表名
'    '输出参数：成功返回 下级最大编码; 否者返回 0
'    Dim strSQL As String
'    Dim rsTemp As New ADODB.Recordset
'
'    Err = 0
'    On Error GoTo Error_Handle
'    If strID = "" Then
'        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID is null " & strWhere & " connect by prior id=上级id"
'        zldatabase.OpenRecordset rsTemp, strSQL, "读取指定表的本级编码的最大长度"
'    Else
'        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with ID=[1] " & strWhere & " connect by prior id=上级id"
'        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "读取指定表的本级编码的最大长度", CLng(strID))
'    End If
'
'    If rsTemp.EOF Then
'        GetDownCodeLength = 0
'    Else
'        GetDownCodeLength = rsTemp.Fields("LenCode").Value
'    End If
'    Exit Function
'Error_Handle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'    GetDownCodeLength = 0
'End Function

Public Function GetLocalCodeLength(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：上级ID，表名
    '输出参数：成功返回 最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID is null" & strWhere
        zlDatabase.OpenRecordset rsTemp, strSQL, "读取指定表的本级编码的最大长度"
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID=[1]" & strWhere
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取指定表的本级编码的最大长度", CLng(str上级ID))
    End If
    
    
    If rsTemp.EOF Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str上级ID As String, ByVal strTableName As String) As String
    '功能描述：读取上级编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 上级编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select 编码 from " & strTableName & " where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取上级编码", CLng(str上级ID))
    If rsTemp.EOF Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("编码").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

'Public Function GetMaxLocalCode(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
'    '功能描述：根据指定表的上级ID 读取本级的最大编码
'    '输入参数：上级ID,表名
'    '输出参数：成功返回 最大编码; 否者返回 空
'    Dim strSQL As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim intCode As Integer, StrCode As String, strAllCode As String
'    Dim intLength   As Integer
'    Err = 0
'    On Error GoTo Error_Handle
'    If str上级ID = "" Then
'        strSQL = "select max(to_number(编码))+1 as MaxCode from " & strTableName & " where 上级ID is null" & strWhere
'        zldatabase.OpenRecordset rsTemp, strSQL, "根据指定表的上级ID 读取本级的最大编码"
'    Else
'        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID=[1]" & strWhere
'        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "根据指定表的上级ID 读取本级的最大编码", CLng(str上级ID))
'    End If
'    intCode = GetLocalCodeLength(str上级ID, strTableName, strWhere)
'
'    If rsTemp.EOF Then
'        GetMaxLocalCode = ""
'        Exit Function
'    End If
'    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
'    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
'    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
'    Exit Function
'Error_Handle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'    GetMaxLocalCode = ""
'End Function

Public Function NextNo(intBillId As Integer) As Variant
    '------------------------------------------------------------------------------------
    '功能：根据特定规则产生新的入库单号码,规则如下：
    '       年度位确定原则:
    '       以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsCtrl As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim vntNo As Variant        '获取的号码的中间变量
    Dim intYear, strYear As String      '年度标志位

RESTART:
    Err = 0
    On Error GoTo ErrHand
    
    With rsCtrl
        If .State = adStateOpen Then .Close
        .Open "Select C.项目序号,C.项目名称,C.最大号码,C.自动补缺,C.编号规则,sysdate as Today From 号码控制表 C Where C.项目序号=" & intBillId, gcnOracle, adOpenKeyset, adLockOptimistic
        If .EOF Or .BOF Then
            NextNo = Null
            Exit Function
        End If
        intYear = Format(!Today, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        vntNo = IIf(IsNull(!最大号码), "", !最大号码)
        If Left(vntNo, 1) < strYear Then
            vntNo = strYear & "0000000"
        End If
        vntNo = Left(vntNo, 1) & Right(String(7, "0") & CStr(Val(Mid(vntNo, 2)) + 1), 7)
        
        On Error Resume Next
        .Update "最大号码", vntNo
        If Err <> 0 Then
            .CancelUpdate
            GoTo RESTART
        End If
        NextNo = vntNo
    End With
    Exit Function

ErrHand:
    Call ErrCenter
    Call SaveErrLog
    NextNo = Null
End Function

Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function

'Public Function BinTOHex(sString As String) As String
'    Dim lngLoop As Integer, lngTemp As Long, lngJLoop As Integer, lngTmp As Long
'    lngTemp = 0
'    For lngLoop = 1 To Len(sString)
'        If Mid(sString, lngLoop, 1) = "1" Then
'            lngTmp = 1
'            For lngJLoop = 0 To lngLoop - 2
'                lngTmp = lngTmp * 2
'            Next
'        Else
'            lngTmp = 0
'        End If
'        lngTemp = lngTemp + lngTmp
'    Next
'    BinTOHex = CStr(lngTemp)
'End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function CheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:20070101或2007-01-01
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgbox strTittle & "必须为日期型,请检查！"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgbox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！"
        Exit Function
    End If
    CheckIsDate = strKey
End Function

Public Sub zlChangeCode(ByVal strTableName As String, _
    ByVal lng上级id As Long, _
    ByVal txtUpCode As TextBox, _
    ByVal txtCode As TextBox, _
    Optional ByVal chkChangeCode As CheckBox = Nothing, _
    Optional ByVal strCaption As String = "")
    '------------------------------------------------------------------------------------
    '功能：根据选择的上级确定当前的编码，并在上级及本级中显示出来
    '参数：strTableName-存在分类的表名
    '      lng上级ID-选择的上级
    '      TxtUpCode-显示的上级文本框
    '      TxtUpCode-显示的本级文本框
    '      chkChangeCode-设置是否改变原有数据库中的历史编码选择控件
    '      strCaption-调用窗体的Capiton
    '注意：表中必需有ID,上级id,编码
    '------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intMaxCodeLen As Integer  '确定编码的实际长度
    Err = 0: On Error GoTo ErrHand
    
   chkChangeCode.Value = 0
   chkChangeCode.Enabled = True
   
    If lng上级id = 0 Then
        txtUpCode.Text = ""
        gstrSQL = "select max(编码) as 编码 From " & strTableName & " Where 上级ID is null "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
            
        With rsTemp
            intMaxCodeLen = .Fields("编码").DefinedSize
            If IsNull(!编码) Then
                txtCode.Text = "01"
                txtCode.MaxLength = intMaxCodeLen
                txtCode.Tag = txtCode.MaxLength
                chkChangeCode.Value = 1
                chkChangeCode.Enabled = False
            Else
                txtCode.MaxLength = Len(Trim(!编码))
                txtCode.Tag = txtCode.MaxLength
                If !编码 = String(txtCode.MaxLength, "9") Then
                    If txtCode.MaxLength >= intMaxCodeLen Then
                        ShowMsgbox "最大编码和编码长度已经达到最大限制，无法递增编码"
                        txtCode.Text = Space(txtCode.MaxLength)
                       chkChangeCode.Value = 0
                       chkChangeCode.Enabled = False
                    Else
                        ShowMsgbox "最大编码已经达到本级限制，你可以扩充编码长度以满足需要"
                        txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                        txtCode.MaxLength = txtCode.MaxLength + 1
                        txtCode.Tag = txtCode.MaxLength
                       chkChangeCode.Value = 1
                    End If
                Else
                    txtCode.Text = Format(Mid(!编码, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
                End If
            End If
        End With
        Exit Sub
   End If
   '确定上级编码
   
    gstrSQL = "Select 编码 From " & strTableName & " where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng上级id)
    
    If Not rsTemp.EOF Then
        txtUpCode.Text = zlCommFun.Nvl(rsTemp!编码)
    End If
    
    '先确定是否有下级
    gstrSQL = "select nvl(max(编码),'') as 编码  From " & strTableName & " Where  上级ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng上级id)
    intMaxCodeLen = rsTemp.Fields("编码").DefinedSize

    If zlCommFun.Nvl(rsTemp!编码) = "" Then
        '不存在下级
        '根据上级ID取上级编码
'        gstrSQL = "Select 编码 From " & strTableName & " where id=" & lng上级id
'        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
'        txtUpCode.Text = zlCommFun.Nvl(rsTemp!编码)
        txtCode.MaxLength = intMaxCodeLen - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If txtCode.MaxLength > 1 Then
            txtCode.Text = "01"
        Else
            txtCode.Text = "1"
        End If
        chkChangeCode.Value = 1
        chkChangeCode.Enabled = False
        Exit Sub
    End If
    
    With rsTemp
        txtCode.MaxLength = Len(!编码) - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If Mid(!编码, Len(txtUpCode.Text) + 1) = String(txtCode.MaxLength, "9") Then
            If Len(txtUpCode.Text) + txtCode.MaxLength >= intMaxCodeLen Then
                ShowMsgbox "该分类下级最大编码和编码长度已经达到最大限制，无法递增编码"
                txtCode.Text = Space(txtCode.MaxLength)
               chkChangeCode.Value = 0
               chkChangeCode.Enabled = False
            Else
                ShowMsgbox "该分类下级最大编码已经达到本级限制，你可以扩充编码长度以满足需要"
                txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                txtCode.MaxLength = txtCode.MaxLength + 1
                txtCode.Tag = txtCode.MaxLength
               chkChangeCode.Value = 1
            End If
        Else
            txtCode.Text = Format(Mid(!编码, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ImeLanguage(ByVal blnOpen As Boolean)
    '-----------------------------------------------------------------------------------
    '功能: 打开/关闭输入法
    '参数: blnOpen-是打开还是关闭(true为打开,false为关闭)
    '返回：
    '-----------------------------------------------------------------------------------
    If blnOpen Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '功能：对文本框的的文本选中或进入进打开输入法
    '参数:blnOpenIme-是否打开输入法
    '返回:
    '--------------------------------------------------------------------------------------------------------
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text) ' Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
    If blnOpenIme Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

'Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
'    '-----------------------------------------------------------------------------------
'    '功能:取某字段的值
'    '参数:rsObj          被检查的字段
'    '     varValue       当rsObj为NULL值时的取新值
'    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
'    '-----------------------------------------------------------------------------------
'    If IsNull(rsObj) Then
'        Nvl = varValue
'    Else
'        Nvl = rsObj
'    End If
'End Function

'Public Function Dec2Bin(bDec As Byte) As String
'    '功能：十进制转为二进制函数
'    '用法：String  Dec2Bin(Bdec as Byte)
'    '返回：  十进制的二进制 字符串(String)
'    '错误：  返回"0"
'    Dim strBin As String
'
'    On Error GoTo Err
'    If bDec > 255 Then
'        Dec2Bin = "-1"
'        Exit Function
'    End If
'    strBin = ""
'    '转为字符串
'    While bDec > 0
'        strBin = bDec Mod 2 & strBin
'        bDec = Fix(bDec / 2)
'    Wend
'    '补零足8位
'    If Len(strBin) < 9 Then
'        While Len(strBin) < 8
'            strBin = "0" & strBin
'        Wend
'    End If
'    Dec2Bin = strBin
'    Exit Function
'Err:
'   Dec2Bin = "0"
'End Function
'
'Public Function Bin2Dec(strBin As String) As Long
'    '功能：二进制转为十进制函数
'    '用法：Long  bin2dec(strBin as String)
'    '返回：  二进制的十进制 长整数（Long）
'    '错误：  返回-1
'    Dim lDec As Long
'    Dim lCount As Long
'    Dim i As Long
'
'    On Error GoTo ErrHand
'    lDec = 0
'    If strBin = "" Then strBin = "0"
'    lCount = Len(strBin)
'    For i = 1 To lCount
'        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
'        strBin = Right(strBin, Len(strBin) - 1)
'        DoEvents
'    Next
'    Bin2Dec = lDec
'    Exit Function
'ErrHand:
'    Bin2Dec = -1
'End Function

Public Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer, Optional blnNum As Boolean = False)
    '----------------------------------------------------------------------------------------------------------------
    '功能描述：对指定的列进行排序
    '输入参数：mshFilter-指定的网格
    '          intPreCol-上次列
    '           intPreSort-上次排序
    '           blnNum-是否为数字列
    '输出参数：
    '返回：
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            strTemp = .TextMatrix(.Row, 0)
            If blnNum Then
                If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                   .Sort = flexSortNumericAscending
                   intPreSort = flexSortNumericAscending
                Else
                   .Sort = flexSortNumericDescending
                   intPreSort = flexSortNumericDescending
                End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            .Row = FindRow(mshFilter, strTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Public Function FindRow(ByVal mshgrd As MSHFlexGrid, ByVal varTemp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------------------------------------------
    '功能描述：查找符合条件的行
    '输入参数：varTemp-指定的值
    '           mshGrd-指定网络
    '           intCol-指定的列
    '输出参数：
    '返回：成功返回找到的行
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intTmp As Integer
    
    With mshgrd
        For intTmp = 1 To .Rows - 1
            If IsDate(varTemp) Then
               If Format(.TextMatrix(intTmp, intCol), "yyyy-mm-dd") = Format(varTemp, "yyyy-mm-dd") Then
                  FindRow = intTmp
                  Exit Function
               End If
            Else
                If .TextMatrix(intTmp, intCol) = varTemp Then
                  FindRow = intTmp
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Public Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    Err = 0
    On Error GoTo ErrHand:
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    Exit Function
ErrHand:
    TranNumToDate = ""
End Function

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
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
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
    Err = 0
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
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Function Check相关权限(ByVal strPrivs As String, ByVal strPrv As String) As Boolean
    '功能:检查权限是否存在
    Dim strTmp As String
    strTmp = strPrv
    If IsNumeric(strPrv) Then
        '1位--药品供应商　　2位--物资供应商　　3位--设备供应商　　4位--其他,   5-卫生材料 每位以1或零表示,1表示为true,0为false,以后扩充系统从第6位开始
        strTmp = Decode(strPrv, 1, "药品", 2, "物资", 3, "设备", 4, "其他", 5, "卫材", "无权限")
    End If
    Check相关权限 = InStr(1, ";" & strPrivs & ";", ";" & strPrv & ";") <> 0
End Function

Public Function Get分类权限(ByVal strPrivs As String, Optional aliasName As String = "", Optional bln供应商 As Boolean = True) As String
    '功能:检查权限是否存在
    Dim strTmp As String
    
    '应付记录中的付款标识:1——药品应付款   2——物资应付款   3——设备应付款   4——其他,5--卫生材料
    strTmp = ""
    If InStr(1, ";" & strPrivs & ";", ";药品;") <> 0 Then
        If bln供应商 Then
            strTmp = strTmp & " or substr(" & aliasName & "类型,1,1)=1"
        Else
            strTmp = strTmp & " ,1"
        End If
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";物资;") <> 0 Then
        If bln供应商 Then
            strTmp = strTmp & " or substr(" & aliasName & "类型,2,1)=1"
        Else
            strTmp = strTmp & " ,2"
        End If
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";设备;") <> 0 Then
        If bln供应商 Then
            strTmp = strTmp & " or substr(" & aliasName & "类型,3,1)=1"
        Else
            strTmp = strTmp & " ,3"
        End If
        
    End If
    If InStr(1, ";" & strPrivs & ";", ";其他;") <> 0 Then
        If bln供应商 Then
            strTmp = strTmp & " or substr(" & aliasName & "类型,4,1)=1"
        Else
            strTmp = strTmp & " ,4"
        End If
        
    End If
    
    If InStr(1, ";" & strPrivs & ";", ";卫材;") <> 0 Then
        If bln供应商 Then
            strTmp = strTmp & " or substr(" & aliasName & "类型,5,1)=1"
        Else
            strTmp = strTmp & " ,5"
        End If
        
    End If
    If strTmp <> "" Then
        If bln供应商 Then
            strTmp = "  (" & Mid(strTmp, 4) & ") "
        Else
            strTmp = " NVL(" & aliasName & "系统标识,4)  in (" & Mid(strTmp, 3) & ") "
        End If
    Else
        strTmp = " 1=2 "
    End If
    
    Get分类权限 = strTmp
End Function

'Public Function Decode(ParamArray arrPar() As Variant) As Variant
''功能：模拟Oracle的Decode函数
'    Dim varValue As Variant, i As Integer
'
'    i = 1
'    varValue = arrPar(0)
'    Do While i <= UBound(arrPar)
'        If i = UBound(arrPar) Then
'            Decode = arrPar(i): Exit Function
'        ElseIf varValue = arrPar(i) Then
'            Decode = arrPar(i + 1): Exit Function
'        Else
'            i = i + 2
'        End If
'    Loop
'End Function

Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "", Optional TxtAlignment As mAlignment = 1, Optional blnFontBold As Boolean = False)
    '功能：将PictureBox模拟成3D平面按钮
    '参数：intStyle:0=平面,-1=凹下,1=凸起,2-深凸起
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            Select Case IntStyle
            Case 1
                DrawEdge .hDC, PicRect, BDR_RAISEDINNER Or BF_SOFT, BF_RECT
            Case 2
                DrawEdge .hDC, PicRect, EDGE_RAISED, BF_RECT
            Case -1
                DrawEdge .hDC, PicRect, BDR_SUNKENOUTER Or BF_SOFT, BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) - 10
            End If
            .FontBold = blnFontBold
            picBox.Print strName
        End If
    End With
End Sub

Public Function Check付款与应付明细(ByVal lng付款序号 As Long) As Boolean
    '-------------------------------------------------------------------------------------
    '功能:检查付款明细总金额与应付明细总金额之和是否相等!
    '参数:lng付款序号-付款序号
    '返回:相等,返回true，否则返回false
    '-------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim dbl付款金额 As Double
    
    On Error GoTo errHandle
    strSQL = "Select sum(nvl(a.金额,0)) AS 付款金额 from 付款记录 a where 付款序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取付款金额", lng付款序号)
    If rsTemp.EOF Then
        ShowMsgbox "不存在相关的付款单，请检查(付款序号:" & lng付款序号 & ")!"
        Exit Function
    End If
    dbl付款金额 = Val(Nvl(rsTemp!付款金额))
     
    strSQL = "Select Sum(Case When 记录性质 = 2 Then 计划金额 " & _
             "                When not 记录性质 In (-1, 2) And Nvl(计划金额, 0) <> nvl(发票金额,0) and 计划金额 is null then 发票金额 " & _
             "                When not 记录性质 In (-1, 2) And Nvl(计划金额, 0) <> nvl(发票金额,0) and 计划金额 is not null then 计划金额 " & _
             "                Else 0 End) 发票金额 " & _
             "From 应付记录 " & _
             "Where 付款序号 = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取付款金额", lng付款序号)
    If rsTemp.EOF Then
        ShowMsgbox "不存在相关的应付明细，请检查(付款序号:" & lng付款序号 & ")!"
        Exit Function
    End If
    
    If Round(dbl付款金额, 2) <> Round(Val(Nvl(rsTemp!发票金额)), 2) Then
        Call ShowMsgbox("本次付款(" & Format(dbl付款金额, "###0.00;-###0.00;0;0") & ")与本次付款的明细总额(" & Format(Round(Val(Nvl(rsTemp!发票金额)), 2), "####0.00;-###0.00;0;0") & ")不等，请检查(付款序号:" & lng付款序号 & ")!")
        Exit Function
    End If
    
    Check付款与应付明细 = True
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

'Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
''功能：获取指定控件在屏幕中的位置(Twip)
'    Dim vRect As RECT
'    Call GetWindowRect(lngHwnd, vRect)
'    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
'    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
'    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
'    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
'    GetControlRect = vRect
'End Function

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

''取数据库中发票号的长度，这样，程序中的批号长度与数据库中保持一致了
'Public Function Get发票号Len() As Integer
'    Dim rsTemp As New Recordset
'
'    On Error GoTo errHandle
'    gstrSQL = "select 发票号 from 应付记录 where rownum<1 "
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "取字段长度"
'    Get发票号Len = rsTemp.Fields(0).DefinedSize
'    rsTemp.Close
'    Exit Function
'
'errHandle:
'    If ErrCenter = 1 Then Resume
'End Function

Public Sub zlInitSystemPara()
    '------------------------------------------------------------------------------
    '功能:初始化相关的系统参数
    '返回:填充成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    With gSystemPara
        '0-拼音码,1-五笔码,2-两者
        .int简码方式 = Val(zlDatabase.GetPara("简码方式"))
        .bln个性化风格 = zlDatabase.GetPara("使用个性化风格") = "1"
        '第1位1-全数字只查编码,第2位1-全字母只查简码,在HIS基础参数中设置
        .Para_输入方式 = zlDatabase.GetPara(44, glngSys, 0, "11")
        '.Para_输入方式 = IIf(.Para_输入方式 = "", "11", .Para_输入方式)
     End With
     
     '初如化站点信息
     Call Init站点信息
End Sub

Public Sub 初始小数位数()
    '------------------------------------------------------------------------------------------------------
    '功能:初始小数位数
    '入参:
    '出参:
    '返回:
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_小数位数
        .成本价小数 = 7
        .零售价小数 = 7
        .金额小数 = 4
        .数量小数 = 3
    End With
    With gVbFmtString
        .FM_成本价 = GetFmtString(g_成本价, False)
        .FM_金额 = GetFmtString(g_金额, False)
        .FM_零售价 = GetFmtString(g_售价, False)
        .FM_数量 = GetFmtString(g_数量, False)
    End With
    With gOraFmtString
        .FM_成本价 = GetFmtString(g_成本价, True)
        .FM_金额 = GetFmtString(g_金额, True)
        .FM_零售价 = GetFmtString(g_售价, True)
        .FM_数量 = GetFmtString(g_数量, True)
    End With
End Sub

Public Function GetFmtString(ByVal 小数类型 As g小数类型, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参: lng小数位数-小数位数
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim int位数 As Integer
    Select Case 小数类型
    Case g_数量
         int位数 = g_小数位数.数量小数
    Case g_金额
         int位数 = g_小数位数.金额小数
    Case g_成本价
         int位数 = g_小数位数.成本价小数
    Case g_售价
         int位数 = g_小数位数.零售价小数
    Case Else
        int位数 = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(int位数, "9") & "'"
    Else
       GetFmtString = "#0." & String(int位数, "0") & ";-#0." & String(int位数, "0") & "; ;"
    End If
End Function

'Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
'    '------------------------------------------------------------------------------
'    '功能:判断控件是否可
'    '返回:初如成功,返回true,否则返回False
'    '编制:刘兴宏
'    '日期:2008/01/24
'    '------------------------------------------------------------------------------
'    Dim rsTemp As New ADODB.Recordset
'    Err = 0: On Error GoTo ErrHand:
'    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'End Function

'Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
'    '功能:将集点移动控件中:2008-07-08 16:48:35
'    Err = 0: On Error Resume Next
'    If blnDoEvnts Then DoEvents
'    If zlControl.IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
'End Sub

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Sub Init站点信息()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化站点的相关信息
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-09-01 11:32:00
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gSystemPara.bln存在站点 = gstrNodeNo <> "-"
 End Sub
 
Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function

Public Function zlSelectDept(ByVal FrmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot优先级 As Boolean = False, Optional str所有部门 As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参:cboDept-指定的部门部件
    '     rsDept-指定的部门
    '     strSearch-要搜索的串
    '     blnNot优先级-是否存在优先级字段
    '     str所有部门-所有部门名称
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-26 10:20:11
    '问题:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strIDs As String, str简码 As String
    
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(GetMatchingSting(strSearch, False), "%", "*")
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str所有部门 <> "" Then
        str简码 = zlCommFun.SpellCode(str所有部门)
        If intInputType = 1 Then
            If Trim(str简码) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str简码) Like strCompents Or UCase(str所有部门) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!编码 = "-"
                rsTemp!名称 = str所有部门
                rsTemp!简码 = str简码
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编码) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编码)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Nvl(!编码) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有LXH01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编码) = strSearch Or Trim(!简码) = strSearch Or UCase(Trim(!名称)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If UCase(Trim(!编码)) Like strSearch & "*" Or Trim(Nvl(!简码)) Like strCompents Or UCase(Trim(Nvl(!名称))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!ID)
        
    '刘兴洪:直接定位
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '需要检查是否有多条满足条件的记录
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    Case 1 '输入全拼音
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "简码"
    Case Else
        rsTemp.Sort = IIf(blnNot优先级, "", "优先级,") & "编码"
    End Select
    
    '弹出选择器
    If zlDatabase.zlShowListSelect(FrmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "缺省," & IIf(blnNot优先级, "", ",优先级") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    '未找到
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function

Public Function GetStoreInfo(ByVal strClass As String) As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select Distinct a.Id, a.编码, a.名称, a.简码 " & _
             "From 部门表 A, 部门性质说明 B, 部门性质分类 C " & _
             "Where a.Id = b.部门id And c.名称 = b.工作性质 And c.编码 In (" & strClass & ") " & zl_获取站点限制(True, "A") & " " & _
             "Order By a.名称 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取所有库房信息")
    
   Set GetStoreInfo = rsTmp.Clone
    
    Exit Function

errHandle:
    If ErrCenter = 1 Then Resume
End Function
