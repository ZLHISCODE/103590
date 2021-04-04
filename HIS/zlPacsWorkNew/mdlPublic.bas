Attribute VB_Name = "mdlPublic"
Option Explicit

Public lngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long

Public Enum TNeedType
    tNeedName = 0
    tNeedNo = 1
    tNeedAll = 2
End Enum

Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Public Enum Enum_Inside_Program
    p门诊病历管理 = 1250
    p住院病历管理 = 1251
    p门诊医嘱下达 = 1252
    p住院医嘱下达 = 1253
    p住院医嘱发送 = 1254
    p护理记录管理 = 1255
    p辅诊记录管理 = 1256
    p医嘱附费管理 = 1257
    p诊疗报告管理 = 1258
    p疾病诊断参考 = 1270
    p药品诊疗参考 = 1271
    p病人病历检索 = 1273
    p门诊电子病历 = 2251
    p住院电子病历 = 2252
End Enum

Public gcolPrivs As Collection              '记录内部模块的权限
'DICOM图象参数
Public Const ATTR_检查日期 As String = "8:20"
Public Const ATTR_检查时间 As String = "8:30"
Public Const ATTR_影像类别 As String = "8:60"
Public Const ATTR_检查设备 As String = "8:1090"

'报告内容分隔符
Public Const SPLITER_REPORT = "[[@]]"
Public Const SPLITER_ELEMENT = "[[;]]"
'报告窗体
Public Const Report_Form_frmReportES  As String = "内镜基本信息"
Public Const Report_Form_frmReportPathology As String = "病理妇科液基薄层信息"
Public Const Report_Form_frmReportUS As String = "B超心脏测量信息"
Public Const Report_Form_frmReportCustom As String = "自定义专科报告"

'Public gobjLogFile As clsLogFile
Public gblnUseDebugLog As Boolean


'控制文本框输入值
Public Sub TxtInputControl(ByRef TxtBox As TextBox, ByRef KeyAscii As Integer, ByVal intDecimalPointNum As Integer)
'txtBox：文本框控件
'intDecimalPointNum：小数点位数
'KeyAscii:输入的ASC

    If Chr(KeyAscii) = "." Then
        If InStr(TxtBox.Text, ".") > 0 Then KeyAscii = 0
    End If
    
    If InStr(TxtBox.Text, ".") > 0 And KeyAscii <> 8 Then
        If Len(Mid(TxtBox.Text, InStr(TxtBox.Text, ".") + 1)) >= intDecimalPointNum Then KeyAscii = 0
    End If
End Sub


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

Public Sub WriteLog(ByVal strLog As String)
'写入日志
'    '如果未启用调试日志，则直接退出
'    If Not gblnUseDebugLog Then Exit Sub
'
'    '初始化日志对象
'    If gobjLogFile Is Nothing Then
'        Set gobjLogFile = New clsLogFile
'    End If
'
'    If gobjLogFile.OpenLog() Then
'        Call gobjLogFile.WriteLog(strLog)
'        Call gobjLogFile.CloseLog
'    End If
    LogWrite "PACS主要功能调试日志", glngModul, "模块加载流程跟踪", strLog
End Sub

Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    OutputDebugString "[" & App.ProductName & "]" & strMethob & "：" & objErr.Description
End Sub


Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub

Public Function GetColNum(listTemp As Object, strHead As String) As Integer
    Dim i As Integer
    Select Case UCase(TypeName(listTemp))
        Case UCase("ReportControl")
            For i = 0 To listTemp.Columns.Count - 1
                If listTemp.Columns.Column(i).Caption = strHead Then GetColNum = listTemp.Columns.Column(i).ItemIndex: Exit Function
            Next
        Case UCase("ListView")
            For i = 1 To listTemp.ColumnHeaders.Count
                If listTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("MSHFlexGrid") '以下类型待增，尚未用到
        Case UCase("BillEdit")
        Case UCase("VSFlexGrid")
            For i = 0 To listTemp.Cols - 1
                If listTemp.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("BillEdit")
        Case UCase("DataGrid")
    End Select
End Function

Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean, Optional blnPreserve As Boolean = False, Optional intIsSearchNo As TNeedType = tNeedName)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
      'blnPreserve--如果找不到匹配项目，则保持原有项目
      'intIsSearchNo -- 0:通过编码定位,1:通过名字定位,2:用过编码加名字定位
'说明：未能定位时,设置ListIndex=-1
'       Cbo.SeekIndex功能比较简单，设置index后会触发事件，不适合使用
    Dim i As Long

    For i = 0 To objCbo.ListCount - 1
        If IIf(Abs(intIsSearchNo) = tNeedAll, objCbo.list(i), IIf(Abs(intIsSearchNo) = tNeedNo, zlStr.NeedCode(objCbo.list(i)), zlStr.NeedName(objCbo.list(i)))) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hWnd, i)
            End If
            Exit Sub
        End If
    Next
    
    If blnPreserve = True Then
        If blnEvent = False Then
            Call zlControl.CboSetIndex(objCbo.hWnd, objCbo.ListIndex)
        End If
    Else
        If blnEvent Then
            objCbo.ListIndex = -1
        Else
            Call zlControl.CboSetIndex(objCbo.hWnd, -1)
        End If
    End If
    
End Sub

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetInsidePrivs(ByVal lngProg As Enum_Inside_Program, Optional ByVal blnLoad As Boolean) As String
'功能：获取指定内部模块编号所具有的权限
'参数：blnLoad=是否固定重新读取权限(用于公共模块初始化时,可能用户通过注销的方式切换了)
    Dim strPrivs As String
    
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
    End If
    
    On Error Resume Next
    strPrivs = gcolPrivs("_" & lngProg)
    If err.Number = 0 Then
        If blnLoad Then
            gcolPrivs.Remove "_" & lngProg
        End If
    Else
        err.Clear: On Error GoTo 0
        blnLoad = True
    End If
    
    If blnLoad Then
        strPrivs = GetPrivFunc(glngSys, lngProg)
        gcolPrivs.Add strPrivs, "_" & lngProg
    End If
    GetInsidePrivs = IIf(strPrivs <> "", ";" & strPrivs & ";", "")
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

'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBoxD(ByRef frmParent As Form, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, Optional Title As String = "") As Long

    MsgBoxD = MessageBox(frmParent.hWnd, Prompt, Title, Buttons)

End Function

Public Sub SetWindowStyle(ByVal lngHandle As Long)
'不能使用zlControl.FormSetCaption代替，zlControl.FormSetCaption中还设置了窗体位置，会导致嵌入式窗体的位置改变
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(lngHandle, GWL_STYLE)
    
'    If (lngWindowStyle And WS_CHILD) = WS_CHILD Then Exit Sub
    
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)

    Call SetWindowLong(lngHandle, GWL_STYLE, lngWindowStyle Or WS_CHILD)
End Sub
