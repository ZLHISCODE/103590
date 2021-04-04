Attribute VB_Name = "mdlComlib"

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Sub TxtSelAll(objTxt As Object)
'功能：将编辑框的的文本全部选中
'参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    If TypeName(objTxt) = "TextBox" Then
        If objTxt.MultiLine Then
            SendMessage objTxt.hWnd, WM_VSCROLL, SB_TOP, 0
        End If
    End If
End Sub

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    rs.Open "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", gcnOracle
    GetMaxLength = rs.Fields(0).DefinedSize
    
End Function

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub SelectRow(objVsf As Object, ByVal OldRow As Long, ByVal NewRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    If OldRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, OldRow, objVsf.FixedCols, OldRow, objVsf.Cols - 1) = objVsf.BackColor
    End If
    
    If NewRow + 1 > objVsf.FixedRows Then
        objVsf.Cell(flexcpBackColor, NewRow, objVsf.FixedCols, NewRow, objVsf.Cols - 1) = objVsf.BackColorSel
    End If
    
End Sub

Public Function GetNextId(strTable As String) As Long
    '------------------------------------------------------------------------------------
    '功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
    '参数：
    '   strTable：表名称
    '返回：
    '------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select " & Trim(strTable) & "_ID.Nextval From Dual"
    
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    
    GetNextId = rsTmp.Fields(0).Value
    Exit Function
errH:
    
End Function

Public Function CboLocate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
    'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If blnItem Then
            If cboObj.ItemData(lngLocate) = Val(strValue) Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboObj.List(lngLocate), InStr(1, cboObj.List(lngLocate), "-") + 1) = strValue Then
                cboObj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Public Sub PressKey(bytKey As Byte)
'功能：向键盘发送一个键,类似SendKey
'参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
    
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
'显示帮助窗体
'ChmName:CHM格式文件
'SHwnd:传入窗口句柄(作为宿主窗口)
'htmName:射映在CHM中的htm文件名称

    Dim Path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\zlPiesFlat" & ".chm"
    If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, Path, &H0, htmName & ".htm")
    
    ShowHelp = True
    Exit Function

ShowHelpErr:
    Err.Clear
End Function

Public Function SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：保存窗体及其中各种控件的状态
'参数：objForm:要保存的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
    
    Dim objThis As Object
    Dim strTmp As String
    Dim strIndex As String
    Dim i As Integer, strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '保存窗体状态、位置、大小
    With objForm
        Select Case .WindowState
            Case 0
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState & "," & .Left & "," & .Top & "," & .Width & "," & .Height
            Case 1
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", 0
            Case 2
                SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", objForm.WindowState
        End Select
    End With
   
    SaveWinState = True
End Function

Public Function RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'功能：恢复窗体的状态，当左顶边界超出时，则自动设置为0
'参数：objForm:要恢复的窗体
'      strProjectName：当前工程名，通常可用app.ProductName传递，用以区分不同工程中的同名窗体，保证恢复的正确性；
'      strUserDef：主要适用于工程中，一个窗体多个程序使用(程序使用 set frmxxx=new frm设计窗体形式)，为了按不同应用保存恢复各自的个性化状态，需要直接确定命名。
   
    Dim aryInfo() As String
    Dim strTmp As String, i As Integer
    Dim objThis As Object
    Dim strIndex As String
    Dim strSQL As String
    Dim strOEM As String
    
    On Error Resume Next
    
    If strProjectName <> "" Then strProjectName = strProjectName & "\"
    
    '恢复窗体的状态、位置、大小
    strTmp = "0," & (Screen.Width - objForm.Width) / 2 & "," & (Screen.Height - objForm.Height) / 2 & "," & objForm.Width & "," & objForm.Height
    
    aryInfo = Split(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & strProjectName & objForm.Name & strUserDef & "\Form", "状态", strTmp), ",")
    
    With objForm
        .WindowState = aryInfo(0)
        If UBound(aryInfo) = 4 Then
            .Left = IIf(aryInfo(1) < 0, 0, aryInfo(1))
            .Top = IIf(aryInfo(2) < 0, 0, aryInfo(2))
            .Width = IIf(aryInfo(3) > Screen.Width, Screen.Width, aryInfo(3))
            .Height = IIf(aryInfo(4) > Screen.Height, Screen.Height, aryInfo(4))
        Else
            .Left = (Screen.Width - objForm.Width) / 2
            .Top = (Screen.Height - objForm.Height) / 2
        End If
    End With

    RestoreWinState = True
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPrivs As String
    Dim strWhere As String
    
    On Error GoTo errH
    
    'strWhere = zlRegFunctions(GetUnitInfo("注册码"))
    
    strWhere = "1=1"
    
    If strWhere = "" Or strWhere = "-" Then Exit Function
    
        strSQL = _
            "Select Distinct 功能 From (" & _
            " Select A.系统,A.序号,A.功能" & _
            " From zlRoleGrant A,Session_Roles B" & _
            " Where A.角色 = B.Role And A.序号=" & lngProgID & " And A.系统=" & lngSys & _
            " Union All" & _
            " Select A.系统,B.序号,B.功能" & _
            " From zlPrograms A,zlProgFuncs B" & _
            " Where A.序号=B.序号 And A.系统=B.系统 And A.序号=" & lngProgID & " And A.系统=" & lngSys & _
            " And (Exists(Select 1 From Session_Roles Where Role='DBA')" & _
            " Or A.系统 in (Select 编号 From zlSystems Where Upper(所有者)=USER)" & _
            ")) Where " & strWhere
    
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSQL, gcnOracle
    Do While Not rsTmp.EOF
        strPrivs = strPrivs & ";" & rsTmp!功能
        rsTmp.MoveNext
    Loop
    GetPrivFunc = Mid(strPrivs, 2)
    Exit Function
errH:
    
    ShowSimpleMsg Err.Description
    
End Function
