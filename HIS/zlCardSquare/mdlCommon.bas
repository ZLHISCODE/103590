Attribute VB_Name = "mdlCommon"
Option Explicit
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrMatchMethod As String
Public gstrProductName As String
Public gstrDBUser As String   '��ǰ���ݿ��û�
Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object
Public gstrSQL As String
Public gblnTestCardNo As Boolean  '����
Public gintDebug As Integer
Private Type gPrecision
      ty_С�� As Integer
      ty_Fmt_Vb As String
      ty_Fmt_Ora As String
End Type
Private Type FeePrecision   '������ؾ���
        ty_���� As gPrecision
        ty_��� As gPrecision
End Type
Public glngOld As Long
Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect

Private Type SystemParameter
    int���뷽ʽ As Integer
    bln���Ի���� As Boolean               'ʹ�ø��Ի����
    blnȫ���ְ������ As Boolean
    blnȫ��ĸ������� As Boolean
    bln����վ�� As Boolean      '�Ƿ����վ�����
    ty_���þ��� As FeePrecision    '���þ���
    bln��Һ�ģʽ As Boolean '�Ƿ����ģʽ,���̣�ֱ���ڷ���̨ȡ�ţ�Ȼ���ڽ���ʱ���������۵�
End Type
Public gSystemPara As SystemParameter
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

Public Type Ty_UserInfor
    id As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    �������� As String
    
End Type
Public UserInfo As Ty_UserInfor
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public Type Ty_Color
     lngGridColorSel As OLE_COLOR     'ѡ����ɫ
     lngGridColorLost As OLE_COLOR   '�뿪��ɫ
End Type
Public gSysColor As Ty_Color
Public glngHook As Long
Public gdtBegin As Date

'����Ϊ������
Public gstrComputerName As String '���������
Public glngInstanceCount As Long '��ǰʵ������
Public gcolPrivs As Collection 'Ȩ�޶���

Public Sub UnHookKBD()
    If glngHook <> 0 Then
    UnhookWindowsHookEx glngHook
    glngHook = 0
    End If
End Sub

Public Function EnableKBDHook()
    If glngHook <> 0 Then
        gdtBegin = Time
        Exit Function
    End If
    gdtBegin = Time
    glngHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHFunc, App.hInstance, App.ThreadID)
End Function

Public Function MyKBHFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Time - gdtBegin) * 60 * 60 * 24 < 0.3 Then
        MyKBHFunc = 1 '��ʾҪ�������ѶϢIf wParam = vbKeySnapshot Then '��� ��û�а���PrintScreen��MyKBHFunc = 1 '�����Hook��Ե����ѶϢEnd If
    Else
        MyKBHFunc = 0
    End If
    Call CallNextHookEx(glngHook, iCode, wParam, lParam) '������һ��HookEnd Function
End Function


Public Function SetWindowResizeWndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ����������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = gWinRect.MinW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = gWinRect.MinH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = gWinRect.MaxW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = gWinRect.MaxH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SetWindowResizeWndMessage = 1
        Exit Function
    End If
    SetWindowResizeWndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function


'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '��ȡָ������ĸ�����
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function


Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    On Error Resume Next
    '��ȡָ������ı���
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function


Public Sub zlSetWindowsBroldStyle(ByVal frmMain As Form)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ı�ɵ����岻���ɵ����壨������ֻ�йرհ�ť����,������屾��ֻ�йرգ�ֻ���Զ�������󻯡���С���Ȱ�ť)
    '���:frmMain.hwnd-����ľ��
    '����:
    '����:
    '����:���˺�
    '����:2009-12-10 14:58:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim pt_SavePoint As POINTAPI, pt_MovePoint As POINTAPI
    Err = 0: On Error GoTo Errhand:
    With pt_MovePoint
      .X = (-1): .Y = 10
    End With
    '���ô����broldStyle
    Call SetWindowLong(frmMain.hWnd, GWL_STYLE, GetWindowLong(frmMain.hWnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call GetSystemMenu(frmMain.hWnd, 1&)
    '�����ػ�����
    With frmMain
        .Move .Left, .Top, .Width - 15, .Height - 15
        .Move .Left, .Top, .Width + 15, .Height + 15
    End With
    Call GetCursorPos(pt_SavePoint)
    Call ClientToScreen(frmMain.hWnd, pt_MovePoint)
    Call SetCursorPos(pt_MovePoint.X, pt_MovePoint.Y)
    Call SetCursorPos(pt_SavePoint.X, pt_SavePoint.Y)
Errhand:
End Sub

Public Sub zlInitColorSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼϵͳ��ɫ
    '����:���˺�
    '����:2009-11-27 17:12:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '        Public Const G_Row_COLORSEL = &H8000000D
    '        Public Const G_Row_COLORLost = &HE0E0E0
    With gSysColor
        .lngGridColorLost = &HE0E0E0   '�뿪��ɫ
        .lngGridColorSel = &HFFEBD7       'ѡ����ɫ
    End With
End Sub

Public Function zl_GetUserInfo(Optional cnOracle As ADODB.Connection) As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim objDatabase As clsDataBase
    If Not cnOracle Is Nothing Then
        Set objDatabase = New clsDataBase
        Call objDatabase.InitCommon(cnOracle)
        Set rsTmp = objDatabase.GetUserInfo
        Set objDatabase = Nothing
    Else
        Set rsTmp = zlDatabase.GetUserInfo
    End If
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.id = rsTmp!id
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.�������� = "" & rsTmp!������
        UserInfo.���� = "" & rsTmp!����
        UserInfo.���� = "" & rsTmp!����
        zl_GetUserInfo = True
    End If
    Exit Function
Errhand:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter() = 1 Then Resume
        Call objDatabase.SaveErrLog
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub
Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    GetTaskbarHeight = os.TaskbarHeight
End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '����:���ؼ�ƥ�䴮%dd%,�����Ǵ�д
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function CheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:20070101��2007-01-01
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgbox strTittle & "����Ϊ������,���飡"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgbox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡"
        Exit Function
    End If
    CheckIsDate = strKey
End Function


Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '���ܣ����ı���ĵ��ı�ѡ�л����������뷨
    '����:blnOpenIme-�Ƿ�����뷨
    '����:
    '--------------------------------------------------------------------------------------------------------
    zlControl.TxtSelAll (objTxt)
    
    If blnOpenIme Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

Public Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    Err = 0
    On Error GoTo Errhand:
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
Errhand:
    TranNumToDate = ""
End Function

Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub

Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
End Sub

Public Function zl_GetFieldLens(ByVal strTableName As String, ByVal strFields As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ֶε�ʵ�ʳ���
    '���:strTableName-������
    '     strFields-�ֶ���(�ֶ���ҪΨһ�����򱨴�),��:����,����,����
    '����:
    '����:
    '����:���˺�
    '����:2009-11-17 16:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, cllFields As New Collection
    Dim varFields As Variant, i As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "Select " & strFields & " From " & strTableName & " where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "ȡ�ֶγ���"
    
    varFields = Split(strFields, ",")
    With rsTemp
        For i = 0 To UBound(varFields)
            Select Case .Fields(varFields(i)).type
            Case 222
            Case Else
                cllFields.Add .Fields(varFields(i)).DefinedSize, varFields(i)
            End Select
        Next
    End With
    Set zl_GetFieldLens = cllFields
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub Initվ����Ϣ()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��վ��������Ϣ
    '����:���˺�
    '����:2009-03-02 17:23:24
    '-----------------------------------------------------------------------------------------------------------
    gbln����վ����� = gstrNodeNo <> "-"
 End Sub
Public Sub zl_����վ����Ϣ(ByVal objcbo As ComboBox)
    '-----------------------------------------------------------------------------------------------------------
    '����:����վ����Ϣֵ
    '����:���˺�
    '����:2009-03-03 12:09:01
    '-----------------------------------------------------------------------------------------------------------
    With objcbo
        .Clear
        .AddItem ""
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .ListIndex = 0
    End With
End Sub
 
Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function


Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '����:�жϿؼ��Ƿ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub


'*********************************************************************************************************************
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlComboxLoadFromRecodeset(ByVal strFromCaption As String, ByVal rsSource As ADODB.Recordset, cboControls As Variant, Optional ByVal blnID As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ĺ����Ǵӱ��ؼ�¼��ʱ��װ����������
    '���:cboControls-�ؼ�����
    '     rsSource:Դ��¼(����,����,ȱʡ��־)
    '����:
    '����:
    '����:���˺�
    '����:2009-12-09 14:54:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cboArrays As Variant
    On Error GoTo errHandle
    
    Set rsTemp = rsSource
    '����������
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        'ǿ�����һ������
        cboArrays = Array(cboControls)
    End If
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("����")) Then
                cboArrays(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("����")
            Else
                cboArrays(intCount).AddItem rsTemp("����") & "." & rsTemp("����")
            End If
            If blnID = True Then cboArrays(intCount).ItemData(cboArrays(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("ȱʡ��־") = 1 Then
                cboArrays(intCount).ListIndex = cboArrays(intCount).NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True And cboArrays(intCount).ListIndex < 0 Then cboArrays(intCount).ListIndex = 0
    Next
    zlComboxLoadFromRecodeset = True
    Exit Function
errHandle:
    zlComboxLoadFromRecodeset = False
End Function

Public Function zlComboxLoadFromArray(ByVal varArray As Variant, cboControls As Variant, Optional blnSaveItemData As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ĺ����������ж����б�ֵװ����������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-12-09 14:53:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cboArrays As Variant
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo errHandle
    
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        'ǿ�����һ������
        cboArrays = Array(cboControls)
    End If
    
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        For intArray = LBound(varArray) To UBound(varArray)
            cboArrays(intCount).AddItem varArray(intArray)
        Next
        cboArrays(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromArray = True
    Exit Function
errHandle:
    zlComboxLoadFromArray = False
End Function

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     blnNegative     �Ƿ���и������
    '     blnZero         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    Dim dblValue As Double
    If blnZero = True Then
        If strInput = "" Then
            ShowMsgbox str��Ŀ & "δ���룬����!"
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlDblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '���ý���
        Exit Function
    End If
    zlDblIsValid = True
End Function
Public Function zl_FromComboxGetData(cboControl As ComboBox, Optional ByVal blnID As Boolean = False, Optional strSplit As String = ".") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Combox�л�ȡ����
    '���:blnID-�Ƿ��ȡComboxData����
    '����:
    '����:
    '����:���˺�
    '����:2009-12-11 15:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------

    If cboControl.ListIndex < 0 Then zl_FromComboxGetData = "NULL"
    If blnID = False Then
        If cboControl.Text = "" Or cboControl.Enabled = False Then
            zl_FromComboxGetData = "NULL"
        Else
            zl_FromComboxGetData = "'" & Mid(cboControl.Text, InStr(cboControl.Text, strSplit) + 1) & "'"
        End If
    Else
        zl_FromComboxGetData = cboControl.ItemData(cboControl.ListIndex)
    End If
End Function
 Public Function IsDesinMode() As Boolean
      '���˺� ȷ����ǰģʽΪ���ģʽ
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
  

Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
Errhand:
End Function

Public Function zlGetReDawImge(ByVal frmMain As Form, ByVal lngColor As Long, _
    ByVal strCaption As String, sngWidth As Single, sngHeight As Single, _
    Optional sngFontSize As Single = 9, _
    Optional blnFontBold As Boolean = True) As StdPicture
    Dim objPicture As PictureBox
    Set objPicture = frmMain.Controls.Add("VB.PictureBox", "objPictemp")
    With objPicture
        .Cls
        .AutoRedraw = True
        .FontSize = 9
        .Width = sngWidth: .Height = sngHeight
        objPicture.Line (20, 20)-(sngWidth, sngHeight), lngColor, BF              'һ������(���)
        .ForeColor = &H80000016
        .CurrentY = 20
        .FontBold = blnFontBold
        .FontSize = sngFontSize
        If strCaption <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight("��")) \ 2
            .CurrentX = (.ScaleWidth - .TextWidth(strCaption)) \ 2
            objPicture.Print strCaption
        End If
    End With
    Set zlGetReDawImge = objPicture.Image
    frmMain.Controls.Remove ("objPictemp")
    Set objPicture = Nothing
End Function
Public Sub zlSetStatusPanelCololor(ByVal frmMain As Form, ByVal objStatus As Object, _
    ByVal intPancelIdex As Integer, strCaption As String, _
    ByVal lngColor As Long, Optional blnTextCenter As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����õ�Ԫ�����ɫ
    '��Σ�blnTextCenter-�ı�����
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-23 15:22:18
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    With objStatus
        sngWidth = frmMain.TextWidth(strCaption) + 60
        sngHeight = frmMain.TextHeight("��") + 60
        .Panels(intPancelIdex).Width = sngWidth
        If blnTextCenter = False Then
            .Panels(intPancelIdex).Width = sngWidth + 300
            .Panels(intPancelIdex).Text = strCaption
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, "", 300, sngHeight, 7, True)
        Else
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, strCaption, sngWidth, sngHeight, 7, True)
        End If
    End With
End Sub
 

Public Sub zlDebugTool(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ٵ�����Ϣ
    '���:strInfo-������Ϣ
    '����:���˺�
    '����:2011-05-27 11:36:33
    '˵��:
    '     gintDebug:1-��ʾ��δ������Ϣ,2-����ʽ��Ϣд���ı���������������������Ϣ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFile As FileSystemObject, objText As TextStream, strFile As String
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "�����㲿��", "����", 0))
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    If gintDebug <= 0 Or gintDebug > 2 Then Exit Sub
    If gintDebug = 2 Then
        'д���ļ���
        Set objFile = New FileSystemObject
        strFile = App.Path & "\Square" & Format(Now, "yyyy_MM_DD") & ".Log"
        If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfo: objText.Close
    End If
    MsgBox strInfo, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
End Sub
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

Public Function zlAuditingWarn(ByVal strPrivs As String, _
    ByVal strNos As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��˻��۵�ʱ���Է��ý��б���
    '���:str���=ָ��������Ҫ��˵��к�,Ϊ�ձ�ʾ������
    '����:
    '����:���˺�
    '����:2011-06-23 10:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWarn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, j As Long, str���s As String
    Dim cur���ն� As Currency, cur��� As Currency, cur��� As Currency
    Dim strWarn As String, intWarn As Integer
    Dim bln�����������۷���  As Boolean
    '���ʱ����������۷���
    bln�����������۷��� = zlDatabase.GetPara(98, glngSys) = "1"
    
    strSQL = "" & _
    " Select /*+ rule */ A.�����־, A.����, A.����id , E.Ԥ����� - E.������� As ���, B.������, C.���� As ������," & vbNewLine & _
    "        A.�շ����, D.���� As �������, Sum(A.ʵ�ս��) As ���, Zl_Patiwarnscheme(A.����id) As ���ò���" & vbNewLine & _
    " From ������ü�¼ A, ������Ϣ B, Table(f_Str2list([1])) J," & _
    "           ҽ�Ƹ��ʽ C, �շ���Ŀ��� D," & _
    "           (   Select ����ID,Sum(Nvl(Ԥ�����,0)) as Ԥ�����,Sum(nvl(�������,0))  �������" & _
    "               From  ������� " & vbNewLine & _
    "               Where   ����ID=[2]  and ����=1 And nvl(����,2)=1 Group by ����ID)  E" & vbNewLine & _
    " Where A.��¼���� = 2 And A.����ID+0=[2] And A.��¼״̬ = 0 " & _
    "           And A.NO = J.Column_value " & vbNewLine & _
    "           And A.�շ���� = D.���� And A.����id = E.����id(+) " & vbNewLine & _
    "           And A.����id = B.����id And B.ҽ�Ƹ��ʽ = C.����(+)" & vbNewLine & _
    " Group By Nvl(A.�۸񸸺�, A.���), A.�����־, A.����, A.����id,  B.������, E.Ԥ�����, E.�������, C.����," & vbNewLine & _
    "         A.�շ����, D.����"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNos, lng����ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If InStr(str���s, rsTmp!�շ���� & rsTmp!�������) = 0 Then
                str���s = str���s & "," & rsTmp!�շ���� & rsTmp!�������
            End If
            cur��� = cur��� + rsTmp!���
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
        str���s = Mid(str���s, 2)
        If cur��� > 0 Then
            Set rsWarn = zlGetUnitWarn(rsTmp!���ò���, "0")
            cur���ն� = GetPatiDayMoney(rsTmp!����ID)
            cur��� = Nvl(rsTmp!���, 0)
            If bln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, rsTmp!����ID) + cur���
            '���౨��
            For j = 0 To UBound(Split(str���s, ","))
                intWarn = zlBillingWarn(strPrivs, rsTmp!����, rsTmp!���ò���, rsWarn, _
                    cur���, cur���ն�, cur���, Nvl(rsTmp!������, 0), _
                    Left(Split(str���s, ",")(j), 1), Mid(Split(str���s, ",")(j), 2), strWarn)
                If intWarn = 2 Or intWarn = 3 Then Exit Function
            Next
        End If
    End If
    zlAuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlGetUnitWarn(Optional ByVal str���ò��� As String, Optional ByVal str����ID As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز������ʱ�����¼��
    '���:str���ò���-���õĲ���
    '        str����I�ģ�����ID��
    '����:
    '����:����������
    '����:���˺�
    '����:2011-06-24 14:59:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select Nvl(����ID,0) ����ID,���ò���,Nvl(��������,1) as ��������," & _
            " ����ֵ,������־1,������־2,������־3" & _
            " From ���ʱ����� Where 1=1" & _
            IIf(str���ò��� = "", "", " And ���ò��� = [1]") & _
            IIf(str����ID = "", "", " And Nvl(����ID,0) = [2]")
    Set zlGetUnitWarn = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str���ò���, str����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlBillingWarn(strPrivs As String, str���� As String, str���ò��� As String, _
    rsWarn As ADODB.Recordset, cur��� As Currency, cur���ն� As Currency, _
    cur���ݽ�� As Currency, cur���� As Currency, str��� As String, _
    ByVal str����� As String, ByRef str�ѱ���� As String, Optional bln�ಡ�� As Boolean, Optional strMoneyFMT As String = "") As Integer
'����:�Բ��˼��ʽ��б�����ʾ
'����:
'     str����=��������,������ʾ
'     str���ò���=���ݲ�����ݷ��صļ��ʱ������÷���
'     rsWarn=��ǰ�������ʱ������ü�¼
'     cur���=�������,�����ۼƱ���
'     cur���ն�=���˵��շ����ķ��ö�,����ÿ�ձ���
'     cur���ݽ��=���˵���������ķ���
'     cur����=���˵������ö�,�����ۼƱ���
'     str���=��ǰҪ�������,���ڷ��౨��
'     str�����=�������,������ʾ
'     strMoneyFMT-��ʽ����
'����:0;û�б���,����
'     1:������ʾ���û�ѡ�����
'     2:������ʾ���û�ѡ���ж�
'     3:������ʾ�����ж�
'     4:ǿ�Ƽ��ʱ���,����
'     str�������="CDE":�����ڱ��α�����һ�����,"-"Ϊ������𡣸÷������ڴ����ظ�����
    Dim i As Integer, byt��־ As Byte
    Dim bln�ѱ��� As Boolean
    Dim byt��ʽ As Byte, byt�ѱ���ʽ As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    If strMoneyFMT = "" Then
        strMoneyFMT = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    End If
    '�����������
    rsWarn.Filter = "����ID=0 And ���ò���='" & str���ò��� & "'"
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!����ֵ) Then Exit Function
    
    '��Ӧ���λ��Ч��������
    If Not IsNull(rsWarn!������־1) Then
        If rsWarn!������־1 = "-" Or InStr(rsWarn!������־1, str���) > 0 Then byt��־ = 1
        If rsWarn!������־1 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־2) Then
        If rsWarn!������־2 = "-" Or InStr(rsWarn!������־2, str���) > 0 Then byt��־ = 2
        If rsWarn!������־2 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 And Not IsNull(rsWarn!������־3) Then
        If rsWarn!������־3 = "-" Or InStr(rsWarn!������־3, str���) > 0 Then byt��־ = 3
        If rsWarn!������־3 = "-" Then str����� = "" '�������ʱ,������ʾ��������
    End If
    If byt��־ = 0 Then Exit Function '����Ч����
    
    '������־2ʵ�����������жϢ٢�,����ֻ��һ���жϢ�
    '���ִ����ǰ����һ�����ֻ������һ�ֱ�����ʽ(������������ʱ)
    If bln�ಡ�� Then
        'ʾ����",��:-,��:DEF,��:567,��567"
        '������־2ʾ����",��:-��,��:DEF��,��:567��,��567��"
        bln�ѱ��� = str�ѱ���� & "," Like "*," & str���� & ":-*,*" _
            Or str�ѱ���� & "," Like "*," & str���� & ":*" & str��� & "*,*"
    Else
        'ʾ����"-" �� ",ABC,567,DEF"
        '������־2ʾ����"-��" �� ",ABC��,567��,DEF��"
        bln�ѱ��� = InStr(str�ѱ����, str���) > 0 Or str�ѱ���� Like "-*"
    End If
    
    If bln�ѱ��� Then
        If byt��־ = 2 Then
            If bln�ಡ�� Then
                arrTmp = Split(str�ѱ����, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str���� & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str���� & ":*" & str��� & "*,*" Then
                        byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                        'Exit For  '˵����סԺģ��
                    End If
                Next
            Else
                If str�ѱ���� Like "-*" Then
                    byt�ѱ���ʽ = IIf(Right(str�ѱ����, 1) = "��", 2, 1)
                Else
                    arrTmp = Split(str�ѱ����, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str���) > 0 Then
                            byt�ѱ���ʽ = IIf(Right(arrTmp(i), 1) = "��", 2, 1)
                            'Exit For '˵����סԺģ��
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str����� <> "" Then str����� = """" & str����� & """����"
        
    '---------------------------------------------------------------------
    If rsWarn!�������� = 1 Then  '�ۼƷ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ(����Ԥ����ľ�)��ʾѯ�ʼ���
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        '��ֻ������:1.ǿ�Ƽ���,��Ȩ��ʱ,��ֹ����
                        Call MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ����", vbInformation + vbOKOnly, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & " ����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 2 '���ڱ���ֵ��ʾѯ�ʼ���,Ԥ����ľ�ʱ��ֹ����
                If Not bln�ѱ��� Then
                    If cur��� + cur���� - cur���ݽ�� < 0 Then
                        byt��ʽ = 2
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                            zlBillingWarn = 3
                        Else
                            MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    ElseIf cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                        byt��ʽ = 1
                        If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                            If MsgBox(str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                zlBillingWarn = 2
                            Else
                                zlBillingWarn = 1
                            End If
                        Else
                            MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    End If
                Else
                    '�ϴ��ѱ�����ѡ�������ǿ�Ƽ���
                    If byt�ѱ���ʽ = 1 Then
                        '�ϴε��ڱ���ֵ��ѡ�������ǿ�Ƽ���,���ٴ�����ڵ����,������Ҫ�ж�Ԥ�����Ƿ�ľ�
                        If cur��� + cur���� - cur���ݽ�� < 0 Then
                            byt��ʽ = 2
                            If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                                MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ�," & str����� & "��ֹ���ʡ�", vbInformation, gstrSysName
                                zlBillingWarn = 3
                            Else
                                MsgBox str����� & "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & ")�Ѿ��ľ���", vbInformation, gstrSysName
                                zlBillingWarn = 4
                            End If
                        End If
                    ElseIf byt�ѱ���ʽ = 2 Then
                        '�ϴ�Ԥ�����Ѿ��ľ���ǿ�Ƽ���,���ٴ���
                        Exit Function
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur��� + cur���� - cur���ݽ�� < rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ��ǰʣ���(��������:" & Format(cur����, "0.00") & "):" & Format(cur��� + cur���� - cur���ݽ��, "0.00") & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!�������� = 2 Then  'ÿ�շ��ñ���(����)
        Select Case byt��־
            Case 1 '���ڱ���ֵ��ʾѯ�ʼ���
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        Call MsgBox(str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, strMoneyFMT) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ����.", vbOKOnly + vbInformation, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, strMoneyFMT) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 3 '���ڱ���ֵ��ֹ����
                If cur���ն� + cur���ݽ�� > rsWarn!����ֵ Then
                    If InStr(";" & strPrivs & ";", ";Ƿ��ǿ�Ƽ���;") = 0 Then
                        MsgBox str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, strMoneyFMT) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & ",��ֹ���ʡ�", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "ǿ�Ƽ�������:" & vbCrLf & vbCrLf & str���� & " ���շ���:" & Format(cur���ն� + cur���ݽ��, strMoneyFMT) & ",����" & str����� & "����ֵ:" & Format(rsWarn!����ֵ, "0.00") & "��", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '���ڼ�����Ĳ���,�����ѱ������
    If zlBillingWarn = 1 Or zlBillingWarn = 4 Then
        If byt��־ = 1 Then
            If rsWarn!������־1 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־1
            End If
        ElseIf byt��־ = 2 Then
            If rsWarn!������־2 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־2
            End If
            '���ӱ�ע���ж��ѱ����ľ��巽ʽ
            str�ѱ���� = str�ѱ���� & IIf(byt��ʽ = 2, "��", "��")
        ElseIf byt��־ = 3 Then
            If rsWarn!������־3 = "-" Then
                str�ѱ���� = IIf(bln�ಡ��, str�ѱ���� & "," & str���� & ":", "") & "-"
            Else
                str�ѱ���� = str�ѱ���� & IIf(bln�ಡ��, "," & str���� & ":", ",") & rsWarn!������־3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUnitID(bytFlag As Byte, lngID As Long) As Long
'���ܣ������շ��ض���Ŀ��ִ�п���
'������bytFlag=ִ�п��ұ�־,lngID=�շ�ϸĿID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '����ȷ����
            GetUnitID = UserInfo.����ID 'ȡ����Ա���ڿ���
        Case 4 'ָ������
            strSQL = "Select B.ִ�п���ID From �շ���ĿĿ¼ A,�շ�ִ�п��� B Where B.�շ�ϸĿID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                GetUnitID = rsTmp!ִ�п���ID 'Ĭ��ȡ��һ��(���ж��)
            Else
                GetUnitID = UserInfo.����ID '��û��ָ������ȡ����Ա���ڿ���
            End If
        Case 1, 2, 3 '���˿���,����Ա����
            GetUnitID = UserInfo.����ID '��ȡ����Ա����
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng����ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����˵��췢���ķ����ܶ�
    '����:��ȡ���˵��췢���ķ����ܶ�
    '����:���˺�
    '����:2011-06-23 10:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as ��� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!���)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(intTYPE As Integer, lng����ID As Long) As Double
'����:��ȡָ�����˵Ļ��۵����ϼ�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnAllFee As Boolean, strWhere As String
        
    On Error GoTo errH
    
    '���ʱ�����������סԺ���۷���
    If intTYPE = 1 Then
        blnAllFee = Val(zlDatabase.GetPara(42, glngSys, 1150)) = 1
        If blnAllFee Then
            strWhere = ""
        Else
            strWhere = " And Nvl(��ҳID,0) = (Select Nvl(��ҳID,0) From ������Ϣ Where ����ID = [1])"
        End If
    Else
        strWhere = ""
    End If
    
    If intTYPE = 1 Then
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־=2" & strWhere
    Else
        strSQL = "" & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ� " & _
        "   From ������ü�¼  " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1]  and �����־<>2" & _
        "   Union ALL   " & _
        "   Select Nvl(Sum(ʵ�ս��),0) As ���۷��úϼ�  " & _
        "   From סԺ���ü�¼ " & _
        "   Where ��¼״̬=0 And ���ʷ���=1 And ����ID=[1] and �����־<>2 "
        strSQL = "" & _
        "   Select Sum(nvl(���۷��úϼ�,0)) as ���۷��úϼ�  " & _
        "   From ( " & strSQL & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡָ�����˵Ļ����ܶ�", lng����ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = rsTmp!���۷��úϼ�
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
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
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function
Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
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
Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function
Public Function StrToNum(ByVal strNumber As String) As Double
    '����:���ַ���ת��������
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function


Public Function ExistFeeInsurePatient(lng����ID As Long) As Boolean
'���ܣ��ж�ҽ�������Ƿ����δ�����
'���أ�
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSQL = "Select Nvl(sum(B.�������),0) ������� From ������Ϣ A,������� B Where A.����ID=B.����ID And Nvl(A.����,0)<>0 And A.����ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng����ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!������� <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'���ܣ���ȡ�����б��ѡ��ĵ���
'������
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select ���� as ID,����,����,���� From ����" & _
                 " Where Nvl(����,0)<3 And (���� Like [1] Or upper(����) Like '" & gstrLike & "'||[1]||'%' Or ���� Like '" & gstrLike & "'||[1]||'%')"
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select ���� as ID,����,����,���� From ���� Where Nvl(����,0)<3"
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "����", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str�������� As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ�������,���ò�ͬ�������͵���ʾ��ɫ
    '���:objPatiControl-���˿ؼ�(�ı���,��ǩ)
    '    str��������-��������
    '    lngDefaultColor-ȱʡ���˵���ʾ��ɫ
    '����:True-������ɫ�ɹ���False-ʧ��
    '����:���ϴ�
    '����:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str�������� <> "" Then
        lngColor = zlDatabase.GetPatiColor(str��������)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'���ܣ��������뷽ʽ��ʽ������
'������intBit=���С��λ��
'����ţ�94552
'˵����VB�Դ���Round�����м����뷨,��ʵ�ʲ�һ�¡���Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function

Public Function CentMoney(ByVal curMoney As Currency, ByVal bytMoney As Byte) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      bytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6.��������:eg:0.15=0.10:0.16=0.2:   ���˺� ����:34519  ����:2010-12-06 09:58:02
'91385,������5.�������塢������롱�����ȶԷֱҽ����������룬��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74������Ϊ0.5
'       �ֱ����������룬��ô0.00��0.24=0��0.25��0.5=0.50, 0.50��0.74=0.50��0.75��1.00=1������������ռ50%�ı���
    
    Dim intSign As Integer, curTmp As Currency

    If bytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf bytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf bytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf bytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf bytMoney = 4 Then
        CentMoney = Format(FormatEx(curMoney, 1), "0.00")
    ElseIf bytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf bytMoney = 6 Then
        '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
         CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

'============================================================================================================
'��Ԫ���Զ��������ݵ�����������ʹ�䲻���ֺ��������
Public Sub ZL_vsGrid_AutoSetGridRowAndCol(vsfGrid As VSFlexGrid)
    '�Զ������������У�ʹ�䲻���ֺ��������
    Dim sngWidth As Single
    
    With vsfGrid
        If .Cols <= 0 Or .Rows <= 0 Then Exit Sub
        .AutoSize 0, .Cols - 1
        sngWidth = GetAllCellWidth(vsfGrid)
        If sngWidth <= .Width - 300 Or .Cols = 1 Then Exit Sub
        
        Call MoveGridCell(vsfGrid)
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfGrid)
    End With
End Sub

Private Sub MoveGridCell(vsfGrid As VSFlexGrid)
    '�ƶ���Ԫ��
    Dim lngNewRows As Long, lngNewCols As Long
    Dim i As Long, j As Long
    Dim lngMoveCells As Long, lngMoveRows As Long
    Dim lngAfterCurRowCells As Long, lngCurRowAfterCells As Long
    Dim lngNewCol As Long
    Dim blnExitFor As Boolean
    
    With vsfGrid
        '1.�µ�����
        lngNewCols = .Cols - 1
        
        '2.�µ�����
        lngAfterCurRowCells = 0 'ȷ�����һ��ʣ��Ŀհ׵�Ԫ��
        For i = .Cols - 1 To 0 Step -1
            If Trim(.TextMatrix(.Rows - 1, i)) <> "" Then Exit For
            lngAfterCurRowCells = lngAfterCurRowCells + 1
        Next
        If .Rows - lngAfterCurRowCells > 0 Then
            lngNewRows = .Rows + (Ceil((.Rows - lngAfterCurRowCells) / lngNewCols))
        Else
            lngNewRows = .Rows
        End If
        
        '3.��ʼ�ƶ���Ԫ��
        .Rows = lngNewRows
        blnExitFor = False
        For i = .Rows - 1 To 0 Step -1
            For j = .Cols - 1 To 0 Step -1
                If Trim(.TextMatrix(i, j)) <> "" Then
                    'ȷ��Ŀ�굥Ԫ���λ��
                    If j = .Cols - 1 Then
                        lngMoveCells = i + 1 '�ƶ���Ԫ����
                        lngAfterCurRowCells = 0 '��ǰ�к���ĵ�Ԫ����
                    Else
                        lngMoveCells = i  '�ƶ���Ԫ����
                        lngAfterCurRowCells = lngNewCols - 1 - j '��ǰ�к���ĵ�Ԫ����
                    End If
                    
                    lngCurRowAfterCells = lngMoveCells - lngAfterCurRowCells
                    lngMoveRows = Ceil(lngCurRowAfterCells / lngNewCols) '�ƶ�������
                    If lngMoveRows = 0 Then
                        lngNewCol = j + lngMoveCells
                    Else
                        If lngCurRowAfterCells Mod lngNewCols = 0 Then
                            lngNewCol = lngNewCols - 1
                        Else
                            lngNewCol = lngMoveCells - (Floor(lngCurRowAfterCells / lngNewCols)) * lngNewCols - lngAfterCurRowCells - 1
                        End If
                    End If
                    
                    '�ƶ�����
                    .TextMatrix(i + lngMoveRows, lngNewCol) = .TextMatrix(i, j)
                    .Cell(flexcpData, i + lngMoveRows, lngNewCol) = .Cell(flexcpData, i, j)
                    .Cell(flexcpChecked, i + lngMoveRows, lngNewCol) = .Cell(flexcpChecked, i, j)
                End If
                If i = 0 And j = .Cols - 1 Then blnExitFor = True: Exit For
            Next
            If blnExitFor Then Exit For
        Next
        .Cols = lngNewCols
    End With
End Sub

Private Function GetAllCellWidth(vsfGrid As VSFlexGrid) As Single
    '��ȡ���е�Ԫ���ܿ��
    Dim i As Long
    Dim sngWith As Single
    
    sngWith = 0
    With vsfGrid
        For i = 0 To .Cols - 1
            sngWith = sngWith + .ColWidth(i) + 10
        Next
    End With
    GetAllCellWidth = sngWith
End Function

Public Function ZL_vsGrid_CurrCellHaveData(ByVal vsGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '��鵥Ԫ���Ƿ�������
    On Error GoTo ErrHandler
    With vsGrid
        If lngRow = -1 Then lngRow = .Row
        If lngCol = -1 Then lngCol = .Col
        If lngRow < 0 Or lngCol < 0 Then Exit Function
        If lngRow > .Rows - 1 Or lngCol > .Cols - 1 Then Exit Function
        If Trim(.TextMatrix(lngRow, lngCol)) = "" Then Exit Function
    End With
    ZL_vsGrid_CurrCellHaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZL_vsGrid_RemoveCell(ByVal vsGrid As VSFlexGrid, _
    Optional ByVal lngRow As Long = -1, Optional ByVal lngCol As Long = -1) As Boolean
    '�ӱ�����Ƴ���ǰѡ�񿨺�/����Χ
    '��Σ�
    '   lngRow - ��ǰ��
    '   lngCol - ��ǰ��
    Dim i As Long, j As Long
    
    On Error GoTo ErrHandler
    With vsGrid
        .Redraw = flexRDNone
        If lngRow = -1 Then lngRow = .Row
        If lngCol = -1 Then lngCol = .Col
        
        For i = lngRow To .Rows - 1
            For j = 0 To .Cols - 1
                If (i = lngRow And j >= lngCol) Or i > lngRow Then
                    If (i < .Rows - 1 And j = .Cols - 1) Then
                        '�����һ�����һ��
                        If Trim(.TextMatrix(i + 1, 0)) = "" Then
                            .Redraw = flexRDBuffered
                            ZL_vsGrid_RemoveCell = True
                            Exit Function
                        Else
                            .TextMatrix(i, j) = .TextMatrix(i + 1, 0)
                            .Cell(flexcpData, i, j) = .Cell(flexcpData, i + 1, 0)
                            .Cell(flexcpChecked, i, j) = .Cell(flexcpChecked, i + 1, 0)
                        End If
                    ElseIf i = .Rows - 1 And j = .Cols - 1 Then
                        '���һ�����һ��
                        .TextMatrix(i, j) = ""
                        .Cell(flexcpData, i, j) = ""
                        .Cell(flexcpChecked, i, j) = 0
                        
                        If i = 0 Then
                            .Cols = .Cols - 1
                        Else
                            .Col = j - 1
                        End If
                        .Redraw = flexRDBuffered
                        ZL_vsGrid_RemoveCell = True
                        Exit Function
                    Else
                        If .TextMatrix(i, j + 1) = "" Then
                            .TextMatrix(i, j) = ""
                            .Cell(flexcpData, i, j) = ""
                            .Cell(flexcpChecked, i, j) = 0
                            
                            If j = 0 Then
                                .Rows = .Rows - 1
                                .Row = .Rows - 1: .Col = .Cols - 1
                            Else
                                If .Col = j Then .Col = j - 1
                            End If
                            .Redraw = flexRDBuffered
                            ZL_vsGrid_RemoveCell = True
                            Exit Function
                        Else
                            .TextMatrix(i, j) = .TextMatrix(i, j + 1)
                            .Cell(flexcpData, i, j) = .Cell(flexcpData, i, j + 1)
                            .Cell(flexcpChecked, i, j) = .Cell(flexcpChecked, i, j + 1)
                        End If
                    End If
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ZL_vsGrid_AddCell(ByVal vsGrid As VSFlexGrid, _
    ByVal strText As String, ByVal varData As Variant, Optional ByVal blnCheck As Boolean) As Boolean
    '�������������
    '��Σ�
    '   strText - �����ʾ�ı�
    '   varData - ��Ԫ��洢����
    Dim i As Long, j As Long
    Dim blnNoData As Boolean
    
    On Error GoTo ErrHandler
    With vsGrid
        .Redraw = flexRDNone
        If .Rows = 0 Then .Rows = 1: blnNoData = True
        If .Cols = 0 Then .Cols = 1: blnNoData = True
        If blnNoData Then
            .TextMatrix(0, 0) = strText
            .Cell(flexcpData, 0, 0) = varData
            If blnCheck Then .Cell(flexcpChecked, 0, 0) = 2
            .Row = 0: .Col = 0
            .Redraw = flexRDBuffered
            ZL_vsGrid_AddCell = True
            Exit Function
        End If
        
        For i = .Rows - 1 To 0 Step -1
            For j = .Cols - 1 To 0 Step -1
                If Trim(.TextMatrix(i, j)) <> "" Then
                    If j = .Cols - 1 Then
                        If i = 0 Then
                            .Cols = .Cols + 1
                            .TextMatrix(0, .Cols - 1) = strText
                            .Cell(flexcpData, 0, .Cols - 1) = varData
                            If blnCheck Then .Cell(flexcpChecked, 0, .Cols - 1) = 2
                        Else
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = strText
                            .Cell(flexcpData, .Rows - 1, 0) = varData
                            If blnCheck Then .Cell(flexcpChecked, .Rows - 1, 0) = 2
                        End If
                    Else
                        .TextMatrix(i, j + 1) = strText
                        .Cell(flexcpData, i, j + 1) = varData
                        If blnCheck Then .Cell(flexcpChecked, i, j + 1) = 2
                    End If
                    .Redraw = flexRDBuffered
                    ZL_vsGrid_AddCell = True
                    Exit Function
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


