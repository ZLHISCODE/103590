Attribute VB_Name = "mdlPublic"
Option Explicit

Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long�����ֵ

Public Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong
    LL As Long
End Type

Public Type TYPE_USER_INFO
    id As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO
 
Public gblnUseDebugLog As Boolean

Private gstrDebugPath As String

Private grsDeptParas As ADODB.Recordset '���Ʋ�������


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



'��ȡ�������Ӧ������
Public Function GetKeyAliasEx(ByVal lngVirtualKey As Long) As String
    GetKeyAliasEx = ""
    
    If lngVirtualKey >= 59 And lngVirtualKey <= 68 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 58)
    End If
    
    If lngVirtualKey >= 87 And lngVirtualKey <= 88 Then
        GetKeyAliasEx = "F" & (lngVirtualKey - 76)
    End If
End Function

'ȡ����ϼ�����
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
'��ȡ���㴰�ھ��
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
''���ܣ���������Ŀ¼
''������ strDir��������Ŀ¼
''���أ���
''------------------------------------------------
'    Dim objFile As New Scripting.FileSystemObject
'    Dim aNestDirs() As String, i As Integer
'    Dim strPath As String
'    On Error Resume Next
'
'    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
'    ReDim Preserve aNestDirs(0)
'    aNestDirs(0) = strDir
'
'    strPath = objFile.GetParentFolderName(strDir)
'    Do While Len(strPath) > 0
'        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
'        aNestDirs(UBound(aNestDirs)) = strPath
'        strPath = objFile.GetParentFolderName(strPath)
'    Loop
'    '����ȫ��Ŀ¼
'    For i = UBound(aNestDirs) To 0 Step -1
'        MkDir aNestDirs(i)
'    Next
'End Sub
'
'Public Sub ClearCacheFolder(ByVal strCacheFolder As String, ObjFrm As Object)
''------------------------------------------------
''���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
''������ strCacheFolder--��Ҫ����Ƿ���յ�Ŀ¼
''���أ���
''------------------------------------------------
'    Dim objFile As New Scripting.FileSystemObject
'    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
'    Dim strDriver As String
'
'    On Error Resume Next
'    strDriver = objFile.GetDriveName(strCacheFolder)
'    Set objCurFolder = objFile.GetFolder(strCacheFolder)
'    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
'        Call zlCommFun.ShowFlash("�����ͼ�񻺳�Ŀ¼����ȴ���", ObjFrm)
'
'        objCurFolder.Delete True
'        Call zlCommFun.StopFlash
'    End If
'End Sub
'
'Public Function SetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
''���ܣ�����ָ���Ĳ���ֵ
''������lngDept=����ID
''      varPara=������
''      strValue=������ֵ
''���أ������Ƿ�ɹ�
'    Dim strSQL As String
'
'    On Error GoTo errH
'
'    strSQL = "ZL_Ӱ�����̲���_UPDATE(" & lngDeptId & ",'" & varPara & "','" & strValue & "')"
'    Call zlDatabase.ExecuteProcedure(strSQL, "SetPara")
'
'    '���óɹ����������
'    Set grsDeptParas = Nothing
'
'    SetDeptPara = True
'    Exit Function
'errH:
'    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
'End Function
'
'Public Function GetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
''���ܣ���ȡָ���Ĳ���ֵ
''������lngDept=����ID
''      varPara=������
''      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
''      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
''���أ�����ֵ���ַ�����ʽ
'    Dim rsTmp As ADODB.Recordset
'    Dim strSQL As String, blnNew As Boolean
'
'    On Error GoTo errH
'
'    If blnNotCache Then
'        Set rsTmp = New ADODB.Recordset
'        strSQL = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", lngDeptId, varPara)
'
'        If Not rsTmp.EOF Then
'            GetDeptPara = Nvl(rsTmp!����ֵ)
'        Else
'            GetDeptPara = strDefault
'        End If
'    Else
'        '��һ�μ��ز�������
'        If grsDeptParas Is Nothing Then
'            blnNew = True
'        ElseIf grsDeptParas.State = 0 Then
'            blnNew = True
'        End If
'        If blnNew Then
'            strSQL = "Select ����ֵ,������,����ID from Ӱ�����̲���"
'            Set grsDeptParas = New ADODB.Recordset
'            Set grsDeptParas = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
'        End If
'
'        '���ݻ����ȡ����ֵ
'        grsDeptParas.Filter = "������='" & CStr(varPara) & "' AND ����ID=" & lngDeptId
'        If Not grsDeptParas.EOF Then
'            GetDeptPara = Nvl(grsDeptParas!����ֵ)
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
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'���أ������������Rows���������Cols
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
    
    '��ͼ���ʽΪ���µ���ʽʱ����Ҫ�����н�������
    
    '��ʽ1��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    '��1  ��2  ��3  ��4
    
    '��ʽ2��
    'ͼ1  ͼ2  ͼ3  ͼ4
    'ͼ5  ͼ6  ͼ7  ͼ8
    'ͼ9  ��1  ��2  ��3
    
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
    
    '�ٴ�����������
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
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
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
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'���ܣ�������ʽ���߰�ť�е���һ���˵�
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
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal LngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
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
''���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd HH:mm)
'    Dim curDate As Date, strTmp As String
'
'    If strText = "" Then Exit Function
'    curDate = zlDatabase.Currentdate
'    strTmp = strText
'
'    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
'        '���봮�а������ڷָ���
'        If IsDate(strTmp) Then
'            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
'            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
'                'ֻ���������ڲ���
'                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
'            ElseIf Left(strTmp, 10) = "1899-12-30" Then
'                'ֻ������ʱ�䲿��
'                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
'            End If
'        Else
'            '����Ƿ�����,����ԭ����
'            strTmp = strText
'        End If
'    Else
'        '���������ڷָ���
'        If Len(strTmp) <= 2 Then
'            '��������dd
'            strTmp = Format(strTmp, "00")
'            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 4 Then
'            '��������MMdd
'            strTmp = Format(strTmp, "0000")
'            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 6 Then
'            '��������yyMMdd
'            strTmp = Format(strTmp, "000000")
'            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
'        ElseIf Len(strTmp) <= 8 Then
'            '��������MMddHHmm
'            strTmp = Format(strTmp, "00000000")
'            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
'            If Not IsDate(strTmp) Then
'                '��������yyyyMMdd
'                strTmp = Format(strText, "00000000")
'                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
'            End If
'        Else
'            '��������yyyyMMddHHmm
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
''���ܣ���ComboBox�в��Ҳ���λ
''������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
'      'blnPreserve--����Ҳ���ƥ����Ŀ���򱣳�ԭ����Ŀ
'      'blnIsSearchNo --�Ƿ���ͨ�����붨λ
''˵����δ�ܶ�λʱ,����ListIndex=-1
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
''���ܣ���ComboBox�в��Ҳ���λ
''������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
''˵����δ�ܶ�λʱ,����ListIndex=-1
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
'Public Function Get����(str�������� As String) As Integer
''����:���ݳ�������ȡ������
'    If IsDate(str��������) Then
'        Get���� = DateDiff("yyyy", CDate(str��������), Format(zlDatabase.Currentdate, "YYYY-MM-DD"))
'    End If
'End Function
'
'Public Function IntEx(vNumber As Variant) As Variant
''���ܣ�ȡ����ָ����ֵ����С����
'    IntEx = -1 * Int(-1 * vNumber)
'End Function
'
'Public Function Between(X, a, b) As Boolean
''���ܣ��ж�x�Ƿ���a��b֮��
'    If a < b Then
'        Between = X >= a And X <= b
'    Else
'        Between = X >= b And X <= a
'    End If
'End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '���ڳ���ʱ,�Զ��ض�
        strTmp = Substr(strCode, 1, lngLen)
    End If
    
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ָ���������ƿո�
    '--�����:
    '--������:
    '--��  ��:�����ִ�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '��Ҫ�пո������
        strTmp = Substr(strCode, 1, lngLen)
    End If
    
    'ȡ��������ַ�
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function

Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡָ���ִ���ֵ,�ִ��п��԰�������
    '--�����:strInfor-ԭ��
    '         lngStart-ֱʼλ��
    '         lngLen-����
    '--������:
    '--��  ��:�Ӵ�
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
'��ȡ����Ŀ¼
    GetCacheDir = GetAppPath & "\TmpImage\"
End Function

Public Function GetResourceDir() As String
'��ȡ��ԴĿ¼
    GetResourceDir = GetAppPath & "\..\�����ļ�\"
End Function
