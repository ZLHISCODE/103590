Attribute VB_Name = "mdlPublic"
Option Explicit

Public lngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long型最大值

Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    LL As Long
End Type

Public Type TYPE_USER_INFO
    id As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO
 
Public gblnUseDebugLog As Boolean

Private gstrDebugPath As String

Private grsDeptParas As ADODB.Recordset '本科参数表缓存


Public Function GetAppPath() As String
    If gstrDebugPath = "" Then
        If App.LogMode = 0 Then
            gstrDebugPath = "C:\Appsoft\Apply"
        Else
            gstrDebugPath = Replace(App.Path & "\", "\\", "")
        End If
    End If
    
    GetAppPath = gstrDebugPath
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
    End If
End Function

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function
'
'Public Sub MkLocalDir(ByVal strDir As String)
''------------------------------------------------
''功能：创建本地目录
''参数： strDir－－本地目录
''返回：无
''------------------------------------------------
'    Dim objFile As New Scripting.FileSystemObject
'    Dim aNestDirs() As String, i As Integer
'    Dim strPath As String
'    On Error Resume Next
'
'    '读取全部需要创建的目录信息
'    ReDim Preserve aNestDirs(0)
'    aNestDirs(0) = strDir
'
'    strPath = objFile.GetParentFolderName(strDir)
'    Do While Len(strPath) > 0
'        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
'        aNestDirs(UBound(aNestDirs)) = strPath
'        strPath = objFile.GetParentFolderName(strPath)
'    Loop
'    '创建全部目录
'    For i = UBound(aNestDirs) To 0 Step -1
'        MkDir aNestDirs(i)
'    Next
'End Sub
'
'Public Sub ClearCacheFolder(ByVal strCacheFolder As String, ObjFrm As Object)
''------------------------------------------------
''功能：当指定目录的大小达到一定百分比时，清空该目录
''参数： strCacheFolder--需要检查是否清空的目录
''返回：无
''------------------------------------------------
'    Dim objFile As New Scripting.FileSystemObject
'    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
'    Dim strDriver As String
'
'    On Error Resume Next
'    strDriver = objFile.GetDriveName(strCacheFolder)
'    Set objCurFolder = objFile.GetFolder(strCacheFolder)
'    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
'        Call zlCommFun.ShowFlash("正清空图像缓冲目录，请等待！", ObjFrm)
'
'        objCurFolder.Delete True
'        Call zlCommFun.StopFlash
'    End If
'End Sub
'
'Public Function SetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
''功能：设置指定的参数值
''参数：lngDept=科室ID
''      varPara=参数名
''      strValue=参数名值
''返回：设置是否成功
'    Dim strSQL As String
'
'    On Error GoTo errH
'
'    strSQL = "ZL_影像流程参数_UPDATE(" & lngDeptId & ",'" & varPara & "','" & strValue & "')"
'    Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
'
'    '设置成功后清除缓存
'    Set grsDeptParas = Nothing
'
'    SetDeptPara = True
'    Exit Function
'errH:
'    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
'End Function
'
'Public Function GetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
''功能：读取指定的参数值
''参数：lngDept=科室ID
''      varPara=参数名
''      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
''      blnNotCache=是否不从缓存中读取
''返回：参数值，字符串形式
'    Dim rsTmp As ADODB.Recordset
'    Dim strSQL As String, blnNew As Boolean
'
'    On Error GoTo errH
'
'    If blnNotCache Then
'        Set rsTmp = New ADODB.Recordset
'        strSQL = "Select 参数值 from 影像流程参数 where 科室ID = [1] and 参数名=[2]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取参数", lngDeptId, varPara)
'
'        If Not rsTmp.EOF Then
'            GetDeptPara = Nvl(rsTmp!参数值)
'        Else
'            GetDeptPara = strDefault
'        End If
'    Else
'        '第一次加载参数缓存
'        If grsDeptParas Is Nothing Then
'            blnNew = True
'        ElseIf grsDeptParas.State = 0 Then
'            blnNew = True
'        End If
'        If blnNew Then
'            strSQL = "Select 参数值,参数名,科室ID from 影像流程参数"
'            Set grsDeptParas = New ADODB.Recordset
'            Set grsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "读取参数")
'        End If
'
'        '根据缓存读取参数值
'        grsDeptParas.Filter = "参数名='" & CStr(varPara) & "' AND 科室ID=" & lngDeptId
'        If Not grsDeptParas.EOF Then
'            GetDeptPara = Nvl(grsDeptParas!参数值)
'        Else
'            GetDeptPara = strDefault
'        End If
'    End If
'    Exit Function
'errH:
'    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
'End Function

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
err:
End Sub

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    lngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
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
    SetWindowLong objForm.hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
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
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
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

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
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
'
'Public Function GetFullDate(ByVal strText As String) As String
''功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd HH:mm)
'    Dim curDate As Date, strTmp As String
'
'    If strText = "" Then Exit Function
'    curDate = zlDatabase.Currentdate
'    strTmp = strText
'
'    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
'        '输入串中包含日期分隔符
'        If IsDate(strTmp) Then
'            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
'            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
'                '只输入了日期部份
'                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
'            ElseIf Left(strTmp, 10) = "1899-12-30" Then
'                '只输入了时间部份
'                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
'            End If
'        Else
'            '输入非法日期,返回原内容
'            strTmp = strText
'        End If
'    Else
'        '不包含日期分隔符
'        If Len(strTmp) <= 2 Then
'            '当作输入dd
'            strTmp = Format(strTmp, "00")
'            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 4 Then
'            '当作输入MMdd
'            strTmp = Format(strTmp, "0000")
'            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 6 Then
'            '当作输入yyMMdd
'            strTmp = Format(strTmp, "000000")
'            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 8 Then
'            '当作输入MMddHHmm
'            strTmp = Format(strTmp, "00000000")
'            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
'            If Not IsDate(strTmp) Then
'                '当作输入yyyyMMdd
'                strTmp = Format(strText, "00000000")
'                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
'            End If
'        Else
'            '当作输入yyyyMMddHHmm
'            strTmp = Format(strTmp, "000000000000")
'            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
'        End If
'    End If
'    GetFullDate = strTmp
'End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function
'
'Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional blnIsSearchNo As Boolean = False)
''功能：在ComboBox中查找并定位
''参数：blnEvent=定位时是否触发Click事件
'      'blnPreserve--如果找不到匹配项目，则保持原有项目
'      'blnIsSearchNo --是否是通过编码定位
''说明：未能定位时,设置ListIndex=-1
'    Dim i As Long
'
'    For i = 0 To objCbo.ListCount - 1
'        If IIf(blnIsSearchNo, NeedNo(objCbo.List(i)), NeedName(objCbo.List(i))) = strText Then
'            If blnEvent Then
'                objCbo.ListIndex = i
'            Else
'                Call zlControl.CboSetIndex(objCbo.hwnd, i)
'            End If
'            Exit Sub
'        End If
'    Next
'
'    If blnPreserve = True Then
'        If blnEvent = False Then
'            Call zlControl.CboSetIndex(objCbo.hwnd, objCbo.ListIndex)
'        End If
'    Else
'        If blnEvent Then
'            objCbo.ListIndex = -1
'        Else
'            Call zlControl.CboSetIndex(objCbo.hwnd, -1)
'        End If
'    End If
'
'End Sub
'
'Public Sub SeekIndexWithNo(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
''功能：在ComboBox中查找并定位
''参数：blnEvent=定位时是否触发Click事件
''说明：未能定位时,设置ListIndex=-1
'    Dim i As Long
'
'    For i = 0 To objCbo.ListCount - 1
'        If NeedNo(objCbo.List(i)) = strText Then
'            If blnEvent Then
'                objCbo.ListIndex = i
'            Else
'                Call zlControl.CboSetIndex(objCbo.hwnd, i)
'            End If
'            Exit Sub
'        End If
'    Next
'    If blnEvent Then
'        objCbo.ListIndex = -1
'    Else
'        Call zlControl.CboSetIndex(objCbo.hwnd, -1)
'    End If
'End Sub
'
'Public Function NeedNo(strList As String) As String
'    If InStr(strList, "[") > 0 And InStr(strList, "-") = 0 Then
'        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "[") - 1))
'    ElseIf InStr(strList, "(") > 0 And InStr(strList, "-") = 0 Then
'        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "(") - 1))
'    ElseIf InStr(strList, "-") > 0 Then
'        NeedNo = LTrim(Mid(strList, 1, InStr(strList, "-") - 1))
'    Else
'        NeedNo = LTrim(strList)
'    End If
'End Function
'
'Public Function Get年龄(str出生日期 As String) As Integer
''功能:根据出生日期取得年龄
'    If IsDate(str出生日期) Then
'        Get年龄 = DateDiff("yyyy", CDate(str出生日期), Format(zlDatabase.Currentdate, "YYYY-MM-DD"))
'    End If
'End Function
'
'Public Function IntEx(vNumber As Variant) As Variant
''功能：取大于指定数值的最小整数
'    IntEx = -1 * Int(-1 * vNumber)
'End Function
'
'Public Function Between(X, a, b) As Boolean
''功能：判断x是否在a和b之间
'    If a < b Then
'        Between = X >= a And X <= b
'    Else
'        Between = X >= b And X <= a
'    End If
'End Function

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


Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Public Function GetCacheDir() As String
'获取缓存目录
    GetCacheDir = GetAppPath & "\TmpImage\"
End Function

Public Function GetResourceDir() As String
'获取资源目录
    GetResourceDir = GetAppPath & "\..\附加文件\"
End Function
