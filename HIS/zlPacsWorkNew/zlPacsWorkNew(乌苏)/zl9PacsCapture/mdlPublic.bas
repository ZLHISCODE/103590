Attribute VB_Name = "mdlPublic"
Option Explicit


Public Enum TMediaType
    imgTag = 0   'ͼ����
    MULFRAMETAG = 1 '����ͼ
    VIDEOTAG = 2 '��Ƶ���
    AUDIOTAG = 3 '��Ƶ���
End Enum


Public lngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long�����ֵ


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
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

Public gobjGetImage() As Object       'zlPacsGetImage.clsPacsGetIamge
Public gblnUseActivexLoad As Boolean
 
Private grsDeptParas As ADODB.Recordset '���Ʋ�������

Public gblnUseDebugLog As Boolean

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'��̬��������
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "�������ʧ�ܣ�����ϵ����Ա����Ƿ���ȷ��װ!", vbInformation, G_STR_HINT_TITLE
        Set DynamicCreate = Nothing
    End If
    err.Clear
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
'���ܣ���������Ŀ¼
'������ strDir��������Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next

    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir

    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub


Public Sub ClearCacheFolder(ByVal strCacheFolder As String, ObjFrm As Object)
'------------------------------------------------
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
'������ strCacheFolder--��Ҫ����Ƿ���յ�Ŀ¼
'���أ���
'------------------------------------------------
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String

    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        Call zlCL_ShowFlash("�����ͼ�񻺳�Ŀ¼����ȴ���", ObjFrm)
        
        objCurFolder.Delete True
        Call zlCL_StopFlash
    End If
End Sub


Public Function SetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, ByVal strValue As String) As Boolean
'���ܣ�����ָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strValue=������ֵ
'���أ������Ƿ�ɹ�
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "ZL_Ӱ�����̲���_UPDATE(" & lngDeptId & ",'" & varPara & "','" & strValue & "')"
    Call zlCL_ExecuteProcedure(strSQL, "SetPara")
    
    '���óɹ����������
    Set grsDeptParas = Nothing
    
    SetDeptPara = True
    Exit Function
errH:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Public Function GetDeptPara(ByVal lngDeptId As Long, ByVal varPara As String, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
'���ܣ���ȡָ���Ĳ���ֵ
'������lngDept=����ID
'      varPara=������
'      strDefault=�����ݿ���û�иò���ʱʹ�õ�ȱʡֵ(ע�ⲻ��Ϊ��ʱ)
'      blnNotCache=�Ƿ񲻴ӻ����ж�ȡ
'���أ�����ֵ���ַ�����ʽ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnNew As Boolean
    
    On Error GoTo errH
    
    If blnNotCache Then
        Set rsTmp = New ADODB.Recordset
        strSQL = "Select ����ֵ from Ӱ�����̲��� where ����ID = [1] and ������=[2]"
        Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ȡ����", lngDeptId, varPara)
        
        If Not rsTmp.EOF Then
            GetDeptPara = Nvl(rsTmp!����ֵ)
        Else
            GetDeptPara = strDefault
        End If
    Else
        '��һ�μ��ز�������
        If grsDeptParas Is Nothing Then
            blnNew = True
        ElseIf grsDeptParas.State = 0 Then
            blnNew = True
        End If
        If blnNew Then
            strSQL = "Select ����ֵ,������,����ID from Ӱ�����̲���"
            Set grsDeptParas = New ADODB.Recordset
            Set grsDeptParas = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ȡ����")
        End If
        
        '���ݻ����ȡ����ֵ
        grsDeptParas.Filter = "������='" & CStr(varPara) & "' AND ����ID=" & lngDeptId
        If Not grsDeptParas.EOF Then
            GetDeptPara = Nvl(grsDeptParas!����ֵ)
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
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
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
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
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
'        Case UCase("MSHFlexGrid") '�������ʹ�������δ�õ�
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
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = LngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
''���ܣ����ı���Varchar2�ĳ��ȼ��㷽�����нض�
'    Dim strText As String
'
'    strText = IIf(IsNull(varText), "", varText)
'    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
'    'ȥ�����ܳ��ֵİ���ַ�
'    ToVarchar = Replace(ToVarchar, Chr(0), "")
'End Function
Public Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function
'Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
''���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
''������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
'    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
'End Function

'Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
''���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
''������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
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

Public Function GetFullDate(ByVal strText As String) As String
'���ܣ�������������ڼ�,�������������ڴ�(yyyy-MM-dd HH:mm)
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlCL_Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '���봮�а������ڷָ���
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                'ֻ���������ڲ���
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                'ֻ������ʱ�䲿��
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '����Ƿ�����,����ԭ����
            strTmp = strText
        End If
    Else
        '���������ڷָ���
        If Len(strTmp) <= 2 Then
            '��������dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '��������MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '��������yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '��������MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '��������yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '��������yyyyMMddHHmm
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
'���ܣ���ComboBox�в��Ҳ���λ
'������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
      'blnPreserve--����Ҳ���ƥ����Ŀ���򱣳�ԭ����Ŀ
      'blnIsSearchNo --�Ƿ���ͨ�����붨λ
'˵����δ�ܶ�λʱ,����ListIndex=-1
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
'���ܣ���ComboBox�в��Ҳ���λ
'������blnEvent=��λʱ�Ƿ񴥷�Click�¼�
'˵����δ�ܶ�λʱ,����ListIndex=-1
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
Public Function Get����(str�������� As String) As Integer
'����:���ݳ�������ȡ������
    If IsDate(str��������) Then
        Get���� = DateDiff("yyyy", CDate(str��������), Format(zlCL_Currentdate, "YYYY-MM-DD"))
    End If
End Function


Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * vNumber)
End Function

Public Function Between(X, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < b Then
        Between = X >= a And X <= b
    Else
        Between = X >= b And X <= a
    End If
End Function

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

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'���ܣ����ַ�����ComboBox�в�������
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    '�Ⱦ�ȷ����
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '���ģ������
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
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
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
'���ܣ���ItemData����ComboBox������ֵ
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
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
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
    ' ȡ��32λֵ�ĸ�16λ
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' ȡ��32λֵ�ĵ�16λ
    If (LongIn And &HFFFF&) > &H7FFF Then
        LOWORD = (LongIn And &HFFFF&) - &H10000
    Else
        LOWORD = LongIn And &HFFFF&
    End If
End Function


Public Function HasMenu(objMenuBar As Object, ByVal lngMenuId As Long) As Boolean
'�Ƿ����ָ���˵�
    Dim cbrParentMenu As CommandBarControl
    
    Set cbrParentMenu = objMenuBar.FindControl(, lngMenuId)
    
    HasMenu = IIf(cbrParentMenu Is Nothing, False, True)
End Function



Public Function CreateStudyUid(ByVal strUID As String) As String
'�������UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewStudyUID As String
    
    strNewStudyUID = strUID 'M_STR_STUDY_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)

    strSQL = "select ���UID from Ӱ�����¼ where ���UID = [1]" & _
              " Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = [1]"
              
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", strNewStudyUID)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSͼ�񱣴�")
        
        If Len(strNewStudyUID) <= 55 Then
            strNewStudyUID = strNewStudyUID & ".A" & rsData(0)
        Else
            strNewStudyUID = Left(strNewStudyUID, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateStudyUid = strNewStudyUID
End Function


Public Function CreateSeriesUid(ByVal strUID As String) As String
'��������UID
    Dim rsData As New ADODB.Recordset
    Dim strSQL As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSQL = "select ����UID from Ӱ�������� where ����UID = [1]" & _
              " Union All Select ����UID from Ӱ����ʱ���� where ����UID = [1]"
              
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSͼ�񱣴�", strNewSeriesUid)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSQL = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSͼ�񱣴�")
        
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
'���ܣ�ɾ��FTP�е�һ��ͼ�����һ������
'������ frmParent -- ������
'       intType -- ɾ��ͼ������ͣ�1-ɾ��ͼ��2-ɾ������
'       strImageUID -- Ҫɾ��ͼ���UID��intType=1ʱ����Ҫ��ֵ
'       strSeriesUID - Ҫɾ������UID��intType=2ʱ����Ҫ��ֵ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    '�����ɾ��һ��ͼ��ͬʱɾ��ͬ������ͼ�����ù��� ZL_Ӱ��ͼ��_DELETE
    '�����ɾ��һ�����е�ͼ��ͬʱɾ��ͬ���ı���ͼ
    
    Dim iNet As New clsFtp             'FTP��
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
    
    If intType = 1 Then         'ɾ��ͼ��
        strSQL = "Select /*+RULE*/ a.ҽ��ID,a.���ͺ�,c.ͼ��UID,a.����ͼ��, " & _
            " Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'/')||a.���UID As ͼ��Ŀ¼, " & _
            "D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And c.ͼ��UID = [1] " & _
            "And a.λ��һ=D.�豸��(+) And a.λ�ö�=E.�豸��(+)"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSɾ��ͼ��", strImageUID)
        
    ElseIf intType = 2 Then
        strSQL = "Select /*+RULE*/ a.ҽ��ID,a.���ͺ�,c.ͼ��UID, " & _
            " Decode(a.��������,Null,'',to_Char(a.��������,'YYYYMMDD')||'/')||a.���UID As ͼ��Ŀ¼, " & _
            "D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,d.�豸�� as �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,e.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ a,Ӱ�������� b,Ӱ����ͼ�� c,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where a.���UID=b.���UID And b.����UID=c.����UID And b.����UID = [1] " & _
            "And a.λ��һ=D.�豸��(+) And a.λ�ö�=E.�豸��(+)"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "PACSɾ������", strSeriesUID)
        
    End If
    
    If rsTemp.EOF = True Then
        MsgboxCus "û���ҵ�����ɾ����ͼ��!", vbInformation, G_STR_HINT_TITLE
        DeleteImages = False
        Exit Function
    End If
    
    '�Ȳ����豸һ���ڲ����豸��
    If Not IsNull(rsTemp!�豸��1) Then
        strFTPIP = Nvl(rsTemp!Host1)
        strFTPUser = Nvl(rsTemp!User1)
        strFtpPass = Nvl(rsTemp!Pwd1)
        
        intDeviceUsed = 1
        lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)
        
        If lngResult = 0 Then
            If Not IsNull(rsTemp!�豸��2) Then
                strFTPIP = Nvl(rsTemp!Host2)
                strFTPUser = Nvl(rsTemp!User2)
                strFtpPass = Nvl(rsTemp!Pwd2)
                
                intDeviceUsed = 2
                lngResult = iNet.FuncFtpConnect(strFTPIP, strFTPUser, strFtpPass)
                
                If lngResult = 0 Then
                    If MsgboxCus("����FTPʧ�ܣ��Ƿ����ɾ��ͼ��" & vbCrLf & "��ʱ����ɾ������ֻ��ɾ�����ݿ����ݣ��޷�ɾ��ͼ���ļ���" & vbCrLf & "���ǡ������ɾ����������ɾ����", vbQuestion + vbYesNo, G_STR_HINT_TITLE) = vbNo Then
                        DeleteImages = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    strRoot = IIf(intDeviceUsed = 1, Nvl(rsTemp!Root1), Nvl(rsTemp!Root2))
    strImagePath = rsTemp!ͼ��Ŀ¼
    
    If intType = 1 Then
        '�����ɾ������ͼ����ɾ��ͬ������ͼ
        If Not IsNull(rsTemp("����ͼ��")) Then
            arrTmp = Split(rsTemp("����ͼ��"), ";")
            
            For i = 0 To UBound(arrTmp)
                If Trim(arrTmp(i)) <> strImageUID & ".jpg" Then
                    strReportImage = strReportImage & ";" & arrTmp(i)
                End If
            Next
            
            strReportImage = Mid(strReportImage, 2)
        End If
        
        strSQL = "ZL_Ӱ��ͼ��_DELETE(" & rsTemp("ҽ��ID") & "," & rsTemp("���ͺ�") & ",'" & strImageUID & "','" & strReportImage & "')"
        zlCL_ExecuteProcedure strSQL, "Ӱ��ͼ��ɾ��"
        
        'ɾ��ͼ���ļ�
        Call iNet.FuncDelFile(strRoot & strImagePath, strImageUID)
        Call iNet.FuncDelFile(strRoot & strImagePath, strImageUID & ".jpg")
    ElseIf intType = 2 Then
        '��ɾ��ͼ���ļ�,ͬʱɾ��ͬ���ı���ͼ
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            Call iNet.FuncDelFile(strRoot & strImagePath, rsTemp!ͼ��UID)
            Call iNet.FuncDelFile(strRoot & strImagePath, rsTemp!ͼ��UID & ".jpg")
            rsTemp.MoveNext
        Wend
        
        '�����ɾ�����У���ֱ��ɾ�������е�ͼ��
        rsTemp.MoveFirst
        strSQL = "Zl_Ӱ������_Delete(" & rsTemp("ҽ��ID") & ",'" & strSeriesUID & "')"
        zlCL_ExecuteProcedure strSQL, "Ӱ������ɾ��"
        
        '���ɾ������֮�󣬱��μ��û��ͼ����ɾ��FTPĿ¼
        strSQL = "Select ���UID from Ӱ�����¼ where ҽ��ID =[1]"
        Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "����Ƿ���ͼ��", CStr(rsTemp!ҽ��id))
        If IsNull(rsTemp!���uid) Then
            'ɾ��Ŀ¼
            Call iNet.FuncFtpDelDir(strRoot, strImagePath)
        End If
    End If
    
    '�ر�FTP����
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
'���ܣ������ݿ��ж�ȡ�ƶ��洢�豸ID��FTP���ʲ���
'������ frmParent  -- ������
'       strSaveDeviceID �����洢�豸ID
'       strDirURL����[OUT] FTPĿ¼
'       strIp ����[OUT] IP��ַ
'       strUser ���� [OUT]�û���
'       strPwd ����[OUT]�û���
'���أ�True������ȡ�ɹ���False������ȡʧ��
'-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    objFtp.strFtpDir = ""
    objFtp.strFTPIP = ""
    objFtp.strFTPUser = ""
    objFtp.strFTPPwd = ""

    '���洢�豸�Ƿ����
    strSQL = "Select '/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ,����Ŀ¼,����Ŀ¼�û���,����Ŀ¼���� From Ӱ���豸Ŀ¼ Where �豸��=[1]"
    Set rsTemp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "�жϴ洢�豸�Ƿ����", objFtp.strDeviceId)
    
     'û�д洢�豸ʱ�˳�
    If rsTemp.EOF = True Then
        MsgboxCus "û���ҵ��洢�豸,������ѡ��洢�豸!", vbInformation, G_STR_HINT_TITLE
        funGetFtpDeviceInf = False
        
        Exit Function
    End If
    
    objFtp.strFtpDir = Nvl(rsTemp("URL"))
    objFtp.strFTPIP = Nvl(rsTemp("IP��ַ"))
    objFtp.strFTPUser = Nvl(rsTemp("FTP�û���"))
    objFtp.strFTPPwd = Nvl(rsTemp("FTP����"))
    objFtp.strSDDir = Nvl(rsTemp("����Ŀ¼"))
    objFtp.strSDUser = Nvl(rsTemp("����Ŀ¼�û���"))
    objFtp.strSDPswd = Nvl(rsTemp("����Ŀ¼����"))
    
    funGetFtpDeviceInf = True
End Function



Public Sub AddVideoLabelToDicomImage(dcmImage As DicomImage, ByVal strCaptureTimeText As String, _
    ByVal strTimeLenText As String, ByVal strEncoderName As String)
    '����:���label
    '����:dcmImage��dicomͼ��
    '     strCaption�� label�ı�
    Dim labCaption As New DicomLabel
    
    labCaption.LabelType = doLabelText
    '����ʾ������������
    labCaption.Text = strCaptureTimeText & vbCrLf & strTimeLenText '& vbCrLf & strEncoderName
    labCaption.Font.Bold = True
    labCaption.Font.Name = "����"
    labCaption.Font.Size = 10
    labCaption.ForeColour = vbYellow
    labCaption.AutoSize = False

    
    labCaption.Left = 0
    labCaption.Top = 0
    
    Call dcmImage.Labels.Add(labCaption)
End Sub


Public Function GetSingleImage(lngImageUID As String, lngSerialUID As String, ObjFrm As Object, Optional blnMoved As Boolean = False) As Boolean
    '����:��FTP�����ļ�
    '����:����UID
    '�������سɹ�����ļ�·��
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
    
    strSQL = "Select A.ͼ���, A.��̬ͼ, D.FTP�û��� As User1,D.FTP���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2 , e.�豸�� as �豸��2, A.��̬ͼ,A.�������� " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.ͼ��UID= [1]  and a.����UID = [2]  Order By A.ͼ���"
        
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
            
    Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "�����ļ�", lngImageUID, lngSerialUID)
    
    strCachePath = zlCL_GetCacheDir
    ClearCacheFolder strCachePath, ObjFrm
    
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    
    Do While Not rsTmp.EOF
        If strDeviceNO1 <> rsTmp("�豸��1") Then
            strDeviceNO1 = rsTmp("�豸��1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("�豸��2") Then
            strDeviceNO2 = rsTmp("�豸��2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        If rsTmp("��̬ͼ") = VIDEOTAG Then
            strTmpFile = strCachePath & Nvl(rsTmp("URL1")) & ".avi"
        ElseIf rsTmp("��̬ͼ") = AUDIOTAG Then
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
'��ʼ�����Ҽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSaveDeviceID As String
    Dim strSQL As String
    
    On Error GoTo DBError
    
    '��ȡ�����洢�豸��
    strSaveDeviceID = GetDeptPara(lngDeptId, "�洢�豸��")
    
    strSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
    Set rsTmp = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ȡ�洢�豸��Ϣ", strSaveDeviceID)
    
    
    GetIsValidOfStorageDevice = Not rsTmp.EOF
    
    Exit Function
DBError:
    GetIsValidOfStorageDevice = False
    
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Function


Public Sub SaveVideoAreaCfg(ByVal strAreaName As String, ByVal lngHeight As Long)
'������Ƶ�ɼ���������
  Dim strRegPath As String
  
  If lngHeight <= 2500 Then Exit Sub
  
  '����ע������
  strRegPath = G_STR_REG_PATH_PUBLIC & "\" & strAreaName
  
BUGEX "SaveVideoAreaCfg RegPath:" & strRegPath & " Value:" & lngHeight

  SaveSetting "ZLSOFT", strRegPath, "CY1", lngHeight
End Sub


Public Function LoadVideoAreaCfg(ByVal strAreaName As String) As Long
'������Ƶ�ɼ���������
    Dim strRegPath As String
     
    strRegPath = G_STR_REG_PATH_PUBLIC & "\" & strAreaName
    
BUGEX "LoadVideoAreaCfg RegPath:" & strRegPath

    LoadVideoAreaCfg = Val(GetSetting("ZLSOFT", strRegPath, "CY1", 4000))
End Function


Public Function GetInsidePrivs(ByVal lngProg As Long) As String
'���ܣ���ȡָ���ڲ�ģ���������е�Ȩ��
'������blnLoad=�Ƿ�̶����¶�ȡȨ��(���ڹ���ģ���ʼ��ʱ,�����û�ͨ��ע���ķ�ʽ�л���)
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
    
    '���򿪵���״̬������д�����Ϣ�����Զ���ʾ
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
