Attribute VB_Name = "mdlPublic"
Option Explicit


Public Enum TMediaType
    imgTag = 0   '图像标记
    MULFRAMETAG = 1 '多侦图
    VIDEOTAG = 2 '视频标记
    AUDIOTAG = 3 '音频标记
End Enum


Public lngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long型最大值


Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    LL As Long
End Type



Public Type TFtpDeviceInf
    strDeviceId As String
    strFTPIP As String
    strFTPUser As String
    strFTPPwd As String
    strFtpDir As String
    strSDDir As String
    strSDUser As String
    strSDPswd As String
End Type


Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gobjGetImage() As Object       'zlPacsGetImage.clsPacsGetIamge
Public gblnUseActivexLoad As Boolean
 
Private grsDeptParas As ADODB.Recordset '本科参数表缓存

Public gblnUseDebugLog As Boolean

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, G_STR_HINT_TITLE
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

'获取虚拟键对应的名称
Public Function GetKeyAliasEx(ByVal lngVirtualKey As Long) As String
    GetKeyAliasEx = ""
    
    If lngVirtualKey >= 59 And lngVirtualKey <= 68 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 58)
    End If
    
    If lngVirtualKey >= 87 And lngVirtualKey <= 88 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 76)
    End If
End Function

'取得组合键别名
Public Function GetKeyAlias(ByVal KeyCode As Integer, ByVal Shift As Integer) As String

    Dim strShift As String
    Dim strTemp As String
    
    
    strShift = IIf((Shift And vbCtrlMask) <> 0, "CTRL", "")
    
    strTemp = IIf((Shift And vbShiftMask) <> 0, "SHIFT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    strTemp = IIf((Shift And vbAltMask) <> 0, "ALT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
     
    
             
    strTemp = ""
    If KeyCode >= 48 And KeyCode <= 90 Then
        strTemp = Chr(KeyCode)
        
        If strShift = "" Then strShift = "MENU"
    End If
    
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        strTemp = "F" & (KeyCode - 111)
    End If
    
    Select Case KeyCode
        Case vbKeySpace
            strTemp = "SPACE"
    End Select
    
    
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    GetKeyAlias = strShift
                
End Function


Public Function GetTopHwnd(ByVal lngHwnd As Long) As Long
'获取顶层窗口句柄
    Dim lngResult As Long
    
    lngResult = GetAncestor(lngHwnd, GA_ROOT)
    
    If lngResult = 0 Then
        GetTopHwnd = lngHwnd
    Else
        GetTopHwnd = lngResult
        'GetTopHwnd = GetTopHwnd(lngResult)
    End If
End Function


Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
    Dim vRect As RECT, vPos As PointAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function


Public Sub MkLocalDir(ByVal strDir As String)
'------------------------------------------------
'功能：创建本地目录
'参数： strDir－－本地目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next

    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir

    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub


Public Sub ClearCacheFolder(ByVal strCacheFolder As String, ObjFrm As Object)
'------------------------------------------------
'功能：当指定目录的大小达到一定百分比时，清空该目录
'参数： strCacheFolder--需要检查是否清空的目录
'返回：无
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String

    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        Call zlCL_ShowFlash("正清空图像缓冲目录，请等待！", ObjFrm)
        
        objCurFolder.Delete True
        Call zlCL_StopFlash
    End If
End Sub


Public Function SetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
'功能：设置指定的参数值
'参数：lngDept=科室ID
'      varPara=参数名
'      strValue=参数名值
'返回：设置是否成功
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "ZL_影像流程参数_UPDATE(" & lngDeptId & ",'" & varPara & "','" & strValue & "')"
    Call zlCL_ExecuteProcedure(strSQL, "SetPara")
    
    '设置成功后清除缓存
    Set grsDeptParas = Nothing
    
    SetDeptPara = True
    Exit Function
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Public Function GetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
'功能：读取指定的参数值
'参数：lngDept=科室ID
'      varPara=参数名
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'      blnNotCache=是否不从缓存中读取
'返回：参数值，字符串形式
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    
    If blnNotCache Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select 参数值 from 影像流程参数 where 科室ID = [1] and 参数名=[2]"
        Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "读取参数", lngDeptId, varPara)
        
        If Not rsTmp.EOF Then
            GetDeptPara = Nvl(rsTmp!参数值)
        Else
            GetDeptPara = strDefault
        End If
    Else
        '第一次加载参数缓存
        If grsDeptParas Is Nothing Then
            blnNew = True
        ElseIf grsDeptParas.State = 0 Then
            blnNew = True
        End If
        If blnNew Then
            strSQL = "Select 参数值,参数名,科室ID from 影像流程参数"
            Set grsDeptParas = New ADODB.Recordset
            Set grsDeptParas = zlCL_GetDBObj.OpenSQLRecord(strSQL, "读取参数")
        End If
        
        '根据缓存读取参数值
        grsDeptParas.Filter = "参数名='" & CStr(varPara) & "' AND 科室ID=" & lngDeptId
        If Not grsDeptParas.EOF Then
            GetDeptPara = Nvl(grsDeptParas!参数值)
        Else
            GetDeptPara = strDefault
        End If
    End If
    Exit Function
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer)
'-----------------------------------------------------------------------------
'功能：根据输入的图像数量，图像区域的宽度和高度，返回最佳的图像排列行数和列数
'参数： ImageCount－－图像数量
'       RegionWidth--图像显示区域的宽度
'       RegionHeight--图像显示区域的高度
'       Rows－－[返回]最佳行数
'       Cols－－[返回]最佳列数
'返回：返回最佳行数Rows，最佳列数Cols
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    Dim iBase As Integer, blnDoLoop As Integer
    Dim lngFreeCount As Long
    
    If RegionHeight = 0 Then RegionHeight = 1
    If RegionWidth = 0 Then RegionWidth = 1
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    '当图像格式为如下等形式时，需要对行列进行修正
    
    '格式1：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '空1  空2  空3  空4
    
    '格式2：
    '图1  图2  图3  图4
    '图5  图6  图7  图8
    '图9  空1  空2  空3
    
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / iCols > RegionHeight > iRows Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '再次修正行列数
    lngFreeCount = iRows * iCols - ImageCount
    Do While lngFreeCount >= iCols Or lngFreeCount >= iRows
        If lngFreeCount >= iCols Then
            iRows = iRows - 1
        Else
            iCols = iCols - 1
        End If
        
        lngFreeCount = iRows * iCols - ImageCount
    Loop
    
    Rows = iRows: Cols = iCols
    
'    If ImageCount <> 0 Then
'        If Rows * Cols > ImageCount Then
'            iBase = 6
'            blnDoLoop = True
'
'            While blnDoLoop
'                iBase = iBase - 1
'
'                If ImageCount Mod iBase = 0 Then
'                    blnDoLoop = False
'                End If
'            Wend
'
'
'            If RegionWidth > RegionHeight Then
'                If ImageCount / iBase > iBase Then
'                    Cols = ImageCount / iBase
'                    Rows = iBase
'                Else
'                    Rows = ImageCount / iBase
'                    Cols = iBase
'                End If
'            Else
'                If ImageCount / iBase > iBase Then
'                    Cols = iBase
'                    Rows = ImageCount / iBase
'                Else
'                    Rows = iBase
'                    Cols = ImageCount / iBase
'                End If
'            End If
'        End If
'    End If
err:
End Sub


Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hWnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
'Public Function GetColNum(listTemp As Object, strHead As String) As Integer
'    Dim i As Integer
'    Select Case UCase(TypeName(listTemp))
'        Case UCase("ReportControl")
'            For i = 0 To listTemp.Columns.Count - 1
'                If listTemp.Columns.Column(i).Caption = strHead Then GetColNum = listTemp.Columns.Column(i).ItemIndex: Exit Function
'            Next
'        Case UCase("ListView")
'            For i = 1 To listTemp.ColumnHeaders.Count
'                If listTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
'            Next
'        Case UCase("MSHFlexGrid") '以下类型待增，尚未用到
'        Case UCase("BillEdit")
'        Case UCase("VSFlexGrid")
'            For i = 0 To listTemp.Cols - 1
'                If listTemp.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
'            Next
'        Case UCase("BillEdit")
'        Case UCase("DataGrid")
'    End Select
'End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As PointAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
''功能：将文本按Varchar2的长度计算方法进行截断
'    Dim strText As String
'
'    strText = IIf(IsNull(varText), "", varText)
'    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
'    '去掉可能出现的半个字符
'    ToVarchar = Replace(ToVarchar, Chr(0), "")
'End Function
Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
'Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
''功能：将0零转换为"NULL"串,在生成SQL语句时用
''参数：blnForceNum=当为Null时，是否强制表示为数字型
'    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
'End Function

'Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
''功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
''参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
'    Dim strNumber As String
'
'    If TypeName(vNumber) = "String" Then
'        If vNumber = "" Then Exit Function
'        If Not IsNumeric(vNumber) Then Exit Function
'        vNumber = Val(vNumber)
'    End If
'
'    If vNumber = 0 Then
'        strNumber = 0
'    ElseIf Int(vNumber) = vNumber Then
'        strNumber = vNumber
'    Else
'        strNumber = Format(vNumber, "0." & String(intBit, "0"))
'        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
'        If InStr(strNumber, ".") > 0 Then
'            Do While Right(strNumber, 1) = "0"
'                strNumber = Left(strNumber, Len(strNumber) - 1)
'            Loop
'            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
'        End If
'    End If
'    FormatEx = strNumber
'End Function

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

Public Function GetFullDate(ByVal strText As String) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd HH:mm)
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlCL_Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    GetFullDate = strTmp
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function
Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional blnIsSearchNo As Boolean = False)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
      'blnPreserve--如果找不到匹配项目，则保持原有项目
      'blnIsSearchNo --是否是通过编码定位
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long

    For i = 0 To objCbo.ListCount - 1
        If IIf(blnIsSearchNo, NeedNo(objCbo.List(i)), NeedName(objCbo.List(i))) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlCL_CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    
    If blnPreserve = True Then
        If blnEvent = False Then
            Call zlCL_CboSetIndex(objCbo.hWnd, objCbo.ListIndex)
        End If
    Else
        If blnEvent Then
            objCbo.ListIndex = -1
        Else
            Call zlCL_CboSetIndex(objCbo.hWnd, -1)
        End If
    End If
    
End Sub
Public Sub SeekIndexWithNo(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedNo(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlCL_CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlCL_CboSetIndex(objCbo.hWnd, -1)
    End If
End Sub
Public Function NeedNo(strList As String) As String
    If InStr(strList, "[") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "[") - 1))
    ElseIf InStr(strList, "(") > 0 And InStr(strList, "-") = 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "(") - 1))
    ElseIf InStr(strList, "-") > 0 Then
        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "-") - 1))
    Else
        NeedNo = LTrim(strList)
    End If
End Function
Public Function Get年龄(str出生日期 As String) As Integer
'功能:根据出生日期取得年龄
    If IsDate(str出生日期) Then
        Get年龄 = DateDiff("yyyy", CDate(str出生日期), Format(zlCL_Currentdate, "YYYY-MM-DD"))
    End If
End Function


Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function Between(X, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = Substr(strCode, 1, lngLen)
    End If
    '取掉最后半个字符
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    err = 0
    On Error GoTo errhand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
errhand:
    Substr = ""
End Function

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'功能：由字符串在ComboBox中查找索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    '先精确查找
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '最后模糊查找
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'功能：由项目值查找ComboBox的项目索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'功能：由ItemData查找ComboBox的索引值
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function
Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlCL_Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If err.Number <> 0 Then err.Clear: InDesign = True
End Function


Public Function HIWORD(LongIn As Long) As Integer
    ' 取出32位值的高16位
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' 取出32位值的低16位
    If (LongIn And &HFFFF&) > &H7FFF Then
        LOWORD = (LongIn And &HFFFF&) - &H10000
    Else
        LOWORD = LongIn And &HFFFF&
    End If
End Function


Public Function HasMenu(objMenuBar As Object, ByVal lngMenuId As Long) As Boolean
'是否存在指定菜单
    Dim cbrParentMenu As CommandBarControl
    
    Set cbrParentMenu = objMenuBar.FindControl(, lngMenuId)
    
    HasMenu = IIf(cbrParentMenu Is Nothing, False, True)
End Function



Public Function CreateStudyUid(ByVal strUID As String) As String
'创建检查UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewStudyUID As String
    
    strNewStudyUID = strUID 'M_STR_STUDY_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)

    strSQL = "select 检查UID from 影像检查记录 where 检查UID = [1]" & _
              " Union All Select 检查UID from 影像临时记录 where 检查UID = [1]"
              
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS图像保存", strNewStudyUID)
    
    If rsData.RecordCount > 0 Then
        '创建一个新的检查UID
        strSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS图像保存")
        
        If Len(strNewStudyUID) <= 55 Then
            strNewStudyUID = strNewStudyUID & ".A" & rsData(0)
        Else
            strNewStudyUID = Left(strNewStudyUID, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateStudyUid = strNewStudyUID
End Function


Public Function CreateSeriesUid(ByVal strUID As String) As String
'创建序列UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSQL = "select 序列UID from 影像检查序列 where 序列UID = [1]" & _
              " Union All Select 序列UID from 影像临时序列 where 序列UID = [1]"
              
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS图像保存", strNewSeriesUid)
    
    If rsData.RecordCount > 0 Then
        '创建一个新的检查UID
        strSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS图像保存")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Public Function DeleteImages(frmParent As Form, intType As Integer, strImageUID As String, _
    strSeriesUID As String) As Boolean
'------------------------------------------------
'功能：删除FTP中的一个图像或者一个序列
'参数： frmParent -- 主窗体
'       intType -- 删除图像的类型，1-删除图像；2-删除序列
'       strImageUID -- 要删除图像的UID，intType=1时，需要有值
'       strSeriesUID - 要删除序列UID，intType=2时，需要有值
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    '如果是删除一个图像，同时删除同名报告图，调用过程 ZL_影像图象_DELETE
    '如果是删除一个序列的图像，同时删除同名的报告图
    
    Dim iNet As New clsFtp             'FTP类
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFTPIP As String
    Dim strFTPUser As String
    Dim strFtpPass As String
    Dim arrTmp() As String
    Dim strReportImage As String
    Dim intDeviceUsed As Integer
    Dim i As Integer
    Dim strRoot As String
    Dim strImagePath As String
    
    On Error GoTo err
    If intType = 1 And strImageUID = "" Then Exit Function
    If intType = 2 And strSeriesUID = "" Then Exit Function
    
    If intType = 1 Then         '删除图像
        strSQL = "Select /*+RULE*/ a.医嘱ID,a.发送号,c.图像UID,a.报告图象, " & _
            " Decode(a.接收日期,Null,'',to_Char(a.接收日期,'YYYYMMDD')||'/')||a.检查UID As 图像目录, " & _
            "D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,d.设备号 as 设备号1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,e.设备号 as 设备号2 " & _
            "From 影像检查记录 a,影像检查序列 b,影像检查图象 c,影像设备目录 D,影像设备目录 E " & _
            "Where a.检查UID=b.检查UID And b.序列UID=c.序列UID And c.图像UID = [1] " & _
            "And a.位置一=D.设备号(+) And a.位置二=E.设备号(+)"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS删除图像", strImageUID)
        
    ElseIf intType = 2 Then
        strSQL = "Select /*+RULE*/ a.医嘱ID,a.发送号,c.图像UID, " & _
            " Decode(a.接收日期,Null,'',to_Char(a.接收日期,'YYYYMMDD')||'/')||a.检查UID As 图像目录, " & _
            "D.FTP用户名 As User1,D.FTP密码 As Pwd1,D.IP地址 As Host1,'/'||D.Ftp目录||'/' As Root1,d.设备号 as 设备号1," & _
            "E.FTP用户名 As User2,E.FTP密码 As Pwd2,E.IP地址 As Host2,'/'||E.Ftp目录||'/' As Root2,e.设备号 as 设备号2 " & _
            "From 影像检查记录 a,影像检查序列 b,影像检查图象 c,影像设备目录 D,影像设备目录 E " & _
            "Where a.检查UID=b.检查UID And b.序列UID=c.序列UID And b.序列UID = [1] " & _
            "And a.位置一=D.设备号(+) And a.位置二=E.设备号(+)"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACS删除序列", strSeriesUID)
        
    End If
    
    If rsTemp.EOF = True Then
        MsgboxCus "没有找到可以删除的图像!", vbInformation, G_STR_HINT_TITLE
        DeleteImages = False
        Exit Function
    End If
    
    '先查找设备一，在查找设备二
    If Not IsNull(rsTemp!设备号1) Then
        strFTPIP = Nvl(rsTemp!Host1)
        strFTPUser = Nvl(rsTemp!User1)
        strFtpPass = Nvl(rsTemp!Pwd1)
        
        intDeviceUsed = 1
        lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)
        
        If lngResult = 0 Then
            If Not IsNull(rsTemp!设备号2) Then
                strFTPIP = Nvl(rsTemp!Host2)
                strFTPUser = Nvl(rsTemp!User2)
                strFtpPass = Nvl(rsTemp!Pwd2)
                
                intDeviceUsed = 2
                lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)
                
                If lngResult = 0 Then
                    If MsgboxCus("连接FTP失败，是否继续删除图像？" & vbCrLf & "此时继续删除，则只能删除数据库内容，无法删除图像文件。" & vbCrLf & "‘是’则继续删除，‘否’则不删除。", vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then
                        DeleteImages = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    strRoot = IIf(intDeviceUsed = 1, Nvl(rsTemp!Root1), Nvl(rsTemp!Root2))
    strImagePath = rsTemp!图像目录
    
    If intType = 1 Then
        '如果是删除单个图像，则删除同名报告图
        If Not IsNull(rsTemp("报告图象")) Then
            arrTmp = Split(rsTemp("报告图象"), ";")
            
            For i = 0 To UBound(arrTmp)
                If Trim(arrTmp(i)) <> strImageUID & ".jpg" Then
                    strReportImage = strReportImage & ";" & arrTmp(i)
                End If
            Next
            
            strReportImage = Mid(strReportImage, 2)
        End If
        
        strSQL = "ZL_影像图象_DELETE(" & rsTemp("医嘱ID") & "," & rsTemp("发送号") & ",'" & strImageUID & "','" & strReportImage & "')"
        zlCL_ExecuteProcedure strSQL, "影像图像删除"
        
        '删除图像文件
        Call iNet.FuncDelFile(strRoot & strImagePath, strImageUID)
        Call iNet.FuncDelFile(strRoot & strImagePath, strImageUID & ".jpg")
    ElseIf intType = 2 Then
        '先删除图像文件,同时删除同名的报告图
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            Call iNet.FuncDelFile(strRoot & strImagePath, rsTemp!图像UID)
            Call iNet.FuncDelFile(strRoot & strImagePath, rsTemp!图像UID & ".jpg")
            rsTemp.MoveNext
        Wend
        
        '如果是删除序列，则直接删除序列中的图像
        rsTemp.MoveFirst
        strSQL = "Zl_影像序列_Delete(" & rsTemp("医嘱ID") & ",'" & strSeriesUID & "')"
        zlCL_ExecuteProcedure strSQL, "影像序列删除"
        
        '如果删除序列之后，本次检查没有图像，则删除FTP目录
        strSQL = "Select 检查UID from 影像检查记录 where 医嘱ID =[1]"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "检查是否还有图像", CStr(rsTemp!医嘱id))
        If IsNull(rsTemp!检查uid) Then
            '删除目录
            Call iNet.FuncFtpDelDir(strRoot, strImagePath)
        End If
    End If
    
    '关闭FTP连接
    iNet.FuncFtpDisConnect
    
    DeleteImages = True
    Exit Function
err:
    iNet.FuncFtpDisConnect
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Private Sub ImportImgToDicom(objDcmImage As DicomImage, ByVal strImgFile As String)
On Error GoTo errHandle
    Dim objTmp As StdPicture
    Dim objFs As New FileSystemObject
    
    Set objTmp = LoadPicture(strImgFile)
    
    Call objDcmImage.FileImport(strImgFile, "JPG")
Exit Sub
errHandle:
    Call objFs.DeleteFile(strImgFile, True)
End Sub


Public Function funGetFtpDeviceInf(frmParent As Form, objFtp As TFtpDeviceInf) As Boolean
'------------------------------------------------
'功能：从数据库中读取制定存储设备ID的FTP访问参数
'参数： frmParent  -- 父窗体
'       strSaveDeviceID －－存储设备ID
'       strDirURL－－[OUT] FTP目录
'       strIp －－[OUT] IP地址
'       strUser －－ [OUT]用户名
'       strPwd －－[OUT]用户名
'返回：True－－获取成功，False－－获取失败
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    objFtp.strFtpDir = ""
    objFtp.strFTPIP = ""
    objFtp.strFTPUser = ""
    objFtp.strFTPPwd = ""

    '检查存储设备是否存在
    strSQL = "Select '/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As URL,FTP用户名,FTP密码,IP地址,共享目录,共享目录用户名,共享目录密码 From 影像设备目录 Where 设备号=[1]"
    Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "判断存储设备是否存在", objFtp.strDeviceId)
    
     '没有存储设备时退出
    If rsTemp.EOF = True Then
        MsgboxCus "没有找到存储设备,请重新选择存储设备!", vbInformation, G_STR_HINT_TITLE
        funGetFtpDeviceInf = False
        
        Exit Function
    End If
    
    objFtp.strFtpDir = Nvl(rsTemp("URL"))
    objFtp.strFTPIP = Nvl(rsTemp("IP地址"))
    objFtp.strFTPUser = Nvl(rsTemp("FTP用户名"))
    objFtp.strFTPPwd = Nvl(rsTemp("FTP密码"))
    objFtp.strSDDir = Nvl(rsTemp("共享目录"))
    objFtp.strSDUser = Nvl(rsTemp("共享目录用户名"))
    objFtp.strSDPswd = Nvl(rsTemp("共享目录密码"))
    
    funGetFtpDeviceInf = True
End Function



Public Sub AddVideoLabelToDicomImage(dcmImage As DicomImage, ByVal strCaptureTimeText As String, _
    ByVal strTimeLenText As String, ByVal strEncoderName As String)
    '功能:添加label
    '参数:dcmImage：dicom图像
    '     strCaption： label文本
    Dim labCaption As New DicomLabel
    
    labCaption.LabelType = doLabelText
    '不显示编码器的名称
    labCaption.Text = strCaptureTimeText & vbCrLf & strTimeLenText '& vbCrLf & strEncoderName
    labCaption.Font.Bold = True
    labCaption.Font.Name = "宋体"
    labCaption.Font.Size = 10
    labCaption.ForeColour = vbYellow
    labCaption.AutoSize = False

    
    labCaption.Left = 0
    labCaption.Top = 0
    
    Call dcmImage.Labels.Add(labCaption)
End Sub


Public Function GetSingleImage(lngImageUID As String, lngSerialUID As String, ObjFrm As Object, Optional blnMoved As Boolean = False) As Boolean
    '功能:从FTP下载文件
    '传入:序列UID
    '返回下载成功后的文件路径
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    
    On Error GoTo WriteFileErr
    
    GetSingleImage = True
    
    strSQL = "Select A.图像号, A.动态图, D.FTP用户名 As User1,D.FTP密码 As Pwd1,a.图像UID, " & _
        "D.IP地址 As Host1," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL1,d.设备号 as 设备号1, " & _
        "E.FTP用户名 As User2,E.FTP密码 As Pwd2," & _
        "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/'||A.图像UID As URL2 , e.设备号 as 设备号2, A.动态图,A.编码名称 " & _
        "From 影像检查图象 A,影像检查序列 B,影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where A.序列UID=B.序列UID And B.检查UID=C.检查UID And C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And A.图像UID= [1]  and a.序列UID = [2]  Order By A.图像号"
        
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查图象", "H影像检查图象")
        strSQL = Replace(strSQL, "影像检查序列", "H影像检查序列")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
            
    Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "下载文件", lngImageUID, lngSerialUID)
    
    strCachePath = zlCL_GetCacheDir
    ClearCacheFolder strCachePath, ObjFrm
    
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    
    Do While Not rsTmp.EOF
        If strDeviceNO1 <> rsTmp("设备号1") Then
            strDeviceNO1 = rsTmp("设备号1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("设备号2") Then
            strDeviceNO2 = rsTmp("设备号2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        If rsTmp("动态图") = VIDEOTAG Then
            strTmpFile = strCachePath & Nvl(rsTmp("URL1")) & ".avi"
        ElseIf rsTmp("动态图") = AUDIOTAG Then
            strTmpFile = strCachePath & Nvl(rsTmp("URL1")) & ".wav"
        Else
            strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
        End If
        
        If Dir(strTmpFile) = "" Then
            If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                strTmpFile = strCachePath & Nvl(rsTmp("URL2"))

                Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
            End If
        End If

        rsTmp.MoveNext
    Loop
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    Exit Function
WriteFileErr:
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE

End Function


Public Function GetIsValidOfStorageDevice(ByVal lngDeptId As Long) As Boolean
'初始化科室级参数
    Dim rsTmp As New ADODB.Recordset
    Dim strSaveDeviceID As String
    Dim strSQL As String
    
    On Error GoTo DBError
    
    '读取并检测存储设备号
    strSaveDeviceID = GetDeptPara(lngDeptId, "存储设备号")
    
    strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "获取存储设备信息", strSaveDeviceID)
    
    
    GetIsValidOfStorageDevice = Not rsTmp.EOF
    
    Exit Function
DBError:
    GetIsValidOfStorageDevice = False
    
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Public Sub SaveVideoAreaCfg(ByVal strAreaName As String, ByVal lngHeight As Long)
'保存视频采集区域配置
  Dim strRegPath As String
  
  If lngHeight <= 2500 Then Exit Sub
  
  '保存注册表参数
  strRegPath = G_STR_REG_PATH_PUBLIC & "\" & strAreaName
  
BUGEX "SaveVideoAreaCfg RegPath:" & strRegPath & " Value:" & lngHeight

  SaveSetting "ZLSOFT", strRegPath, "CY1", lngHeight
End Sub


Public Function LoadVideoAreaCfg(ByVal strAreaName As String) As Long
'载入视频采集区域配置
    Dim strRegPath As String
     
    strRegPath = G_STR_REG_PATH_PUBLIC & "\" & strAreaName
    
BUGEX "LoadVideoAreaCfg RegPath:" & strRegPath

    LoadVideoAreaCfg = Val(GetSetting("ZLSOFT", strRegPath, "CY1", 4000))
End Function


Public Function GetInsidePrivs(ByVal lngProg As Long) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
On Error Resume Next

    Dim strPrivs As String
    
    strPrivs = zlCL_GetPrivFunc(glngSys, lngProg)

    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
End Function


Public Function MsgboxCus(sPrompt As String, Optional dwStyle As Long, Optional sTitle As String) As Long
    Dim lngHwnd As Long
    
BUGEX "MsgboxCus 1"
    
    If gobjOwner Is Nothing Then
        lngHwnd = GetActiveWindow
    Else
        lngHwnd = gobjOwner.hWnd
    End If
    
    If lngHwnd = GetDesktopWindow Or lngHwnd = 0 Then
BUGEX "MsgboxCus 2 GetForegroundWindow" & " DesktopWindowHwnd:" & lngHwnd
        lngHwnd = GetForegroundWindow
    End If
    
BUGEX "MsgboxCus 3 Hwnd:" & lngHwnd
    
    MsgboxCus = mdlMsgBox.MsgboxEx(lngHwnd, sPrompt, dwStyle, sTitle)
    
    '当打开调试状态后，如果有错误信息，则自动提示
    If err.Number <> 0 And gblnOpenDebug Then
        Call mdlMsgBox.MsgboxEx(lngHwnd, "errSource:" & err.Source & "  errDescription:" & err.Description, vbOKOnly, G_STR_HINT_TITLE)
    End If
    
BUGEX "MsgboxCus End"
End Function


Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function
