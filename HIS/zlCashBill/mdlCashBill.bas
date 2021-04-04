Attribute VB_Name = "mdlCashBill"
Option Explicit

Public gcnOracle As New ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���

Public gstrSysName As String                'ϵͳ����
Public gstrVersion As String                'ϵͳ�汾
Public gstrAviPath As String                'AVI�ļ��Ĵ��Ŀ¼
Public gstrProductName As String            '��Ʒ����
Public gstrMatchMethod As String
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Private mrsPayMode As ADODB.Recordset
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    �������� As String
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public UserInfo As TYPE_USER_INFO

'Ʊ�ݿ���
Public gobjBillPrint As Object '������Ʊ�ݴ�ӡ����
Public gblnBillPrint As Boolean '������Ʊ�ݴ�ӡ�����Ƿ����

Public gstrSQL As String
Public gstr��λ���� As String
Public glngSys  As Long
Public glngModul As Long

Public Enum gBillType 'Ʊ������
    �շ��վ� = 1
    Ԥ���վ� = 2
    �����վ� = 3
    �Һ��վ� = 4
    ���￨ = 5
    ���ѿ� = 6
    ��Ա�� = 5
End Enum

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Enum gAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum
Public Enum EM_DrawStyle
    DW_Flat = 0  '= ƽ��
    Dw_SubKen = -1 '= ����
    Dw_Heave = 1  '= ͹��
    Dw_Deepen_Subken = -2 '= ���,
    Dw_Deepen_Heave = 2 ' = ��͹��
End Enum

'�ؼ���λ
Public Type ty_ctlObject_Locale
    '�ؼ���λ��
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '�����б����С�߶ȺͿ��
    minWidth As Single
    minHeight As Single
    
    '�½��б��ʵ��λ��
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '��ģ���
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type

Public Enum Em_Appearance
    Show_3D = 1     '3D��ʾ
    Show_Flat = 0   'ƽ��
End Enum
Public Enum Em_BorderStyle
    Show_Fixed_Single = 1
    Show_None = 0   '�ޱ߿���
End Enum

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const SPI_GETWORKAREA = 48
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CYFULLSCREEN = 17
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private mlng���ű���ƽ������ As Long
Public gstrLike  As String
Public gstrDec As String

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MshGetColNum(msh As MSHFlexGrid, strColName As String) As Long
'����:������������MSHFlexGrid�ؼ��е������,û���ҵ�ʱ����-1
'����:strColName-����
    Dim i As Long
    
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strColName Then MshGetColNum = i: Exit Function
    Next
    MshGetColNum = -1
End Function


Public Sub zlRaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment = 1)
    '���ܣ���PictureBoxģ���3Dƽ�水ť
    'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    Dim PicRect As RECT
    Dim lngTmp As Long
    If picBox Is Nothing Then Exit Sub
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            If intStyle = 2 Then
                    DrawEdge .hDC, PicRect, EDGE_RAISED Or BF_SOFT, BF_RECT
            ElseIf intStyle = -2 Then
                    DrawEdge .hDC, PicRect, EDGE_SUNKEN Or BF_SOFT, BF_RECT
            Else
                DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
            End If
        End If
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) '-10
            End If
            picBox.Print strName
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub

Public Function GetPersonnelDept(ByVal lngID As Long) As ADODB.Recordset
'���ܣ���ȡָ����Ա�����в���
    Dim strSQL As String
 
    strSQL = "Select B.����,B.ID From ������Ա A, ���ű� B Where A.����id = B.ID And A.��Աid = [1] Order by ȱʡ Desc"
    On Error GoTo errH
    Set GetPersonnelDept = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function TruncateDate(ByVal datFull As Date) As Date
'ȥ�������е�ʱ���֡���
    TruncateDate = CDate(Format(datFull, "yyyy-MM-dd"))
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

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
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
Public Function ReturnMovedExes(ByVal strNO As String, Optional ByVal bytType As Byte = 2, Optional ByVal strFormCaption As String) As Boolean
'����:�����û�ѡ���ѡ�����ݱ��е����ݵ���ǰ���ݱ���
'����:bytType��ʾ��������,ֵ::1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-���ʣ�
'����:�û�ѡ��ȡ������,���߳�ѡ����ת��ʧ��,�򷵻�False
    
    MsgBox "��ǰ�����ĵ���" & strNO & "�ں����ݱ���!" & vbCrLf _
        & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
    ReturnMovedExes = False
    
'�����ǳ�ѡ�������ݵĹ��̣��ݴ棬���ڽ���͸������ʱ����
'    If MsgBox("��ǰ��������" & strNO & "�ں����ݱ���,ϵͳ��Ҫ�Ȱ���˵�����ص�����ת�뵽�������ݱ���ܼ���!" & vbCrLf & _
'                             "ȷ��Ҫ���д˲�����?", vbInformation + vbYesNo, gstrSysName) = vbNo Then
'        ReturnMovedExes = False     '�˾��ʡ
'        Exit Function
'    End If
'
'    If zlDatabase.ReturnMovedExes(strNO, bytType, strFormCaption) Then
'        ReturnMovedExes = True
'    Else
'        '��ϸ������֮ǰ��ִ�й��̳���ʱ����
'        MsgBox "��ϵͳ����,��õ�����ص�����δ��ת�뵽�������ݱ�." & vbCrLf & "����δ�ɹ�,����ϵͳ����Ա��ϵ!", vbInformation, gstrSysName
'        ReturnMovedExes = False
'    End If
End Function

Public Sub zlSetCrlEnbled(ByVal objCrl As Object, blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ؼ���Nabled����,���ΪFalse,ͬʱ��Ҫ������صı���ɫ
    '���:objCrl-ת���ָ���ؼ�
    '     blnEnabled-�������
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:44:25
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Select Case UCase(TypeName(objCrl))
    Case UCase("TextBox"), UCase("COMBOBOX")
        objCrl.Enabled = blnEnabled
        zlSetCtrolBackColor objCrl
    Case UCase("dtpicker"), UCase("frame"), UCase("CHECKBOX"), UCase("LABEL"), UCase("COMMANDBUTTON")
        objCrl.Enabled = blnEnabled
    Case Else
       ' objCrl.Enabled = blnEnabled
    End Select
End Sub
Public Sub zlSetCtrolBackColor(ByVal objCtl As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ɫ����ɫ
    '���:objCtl-ת��Ŀؼ�
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:43:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If objCtl.Enabled = False Then
        objCtl.BackColor = &H8000000F
    Else
        objCtl.BackColor = vbWhite
    End If
End Sub

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
    Err = 0: On Error GoTo ErrHand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
ErrHand:
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
    Err = 0: On Error GoTo ErrHand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
ErrHand:
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
    On Error GoTo ErrHand:
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
ErrHand:
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
    On Error GoTo ErrHand:
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
ErrHand:
End Sub

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
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
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ�������߶�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Call OS.TaskbarHeight
End Function


Public Function zlPersonSelect(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboSel As ComboBox, ByVal rsPerson As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Աѡ��ѡ����
    '���:cboSel-ָ���Ĳ���ѡ�񲿼�
    '     rsPerson-ָ������Ա��Ϣ(ID,���,����,����)
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str����-��������(������,���в���Ա��)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String, strLike As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsPerson)
    
    strSearch = UCase(strSearch)
        
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���� <> "" Then
        str���� = zlCommFun.SpellCode(str����)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!��� = "-"
                rsTemp!���� = str����
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strSearch Then lngID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strSearch) Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!���) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!���)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsPerson, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngID = 0
    If lngID <> 0 And rsTemp.RecordCount = 1 Then lngID = Nvl(rsTemp!ID)
        
    '���˺�:ֱ�Ӷ�λ
    If lngID <> 0 Then GoTo GoOver:
    If lngID < 0 Then lngID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "���"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboSel, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngID = Val(Nvl(rsReturn!ID))
    If lngID < 0 Then lngID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.CboLocate cboSel, lngID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlPersonSelect = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.TxtSelAll cboSel
End Function


Public Function zlIsShowDeptCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲿����Ϣ�Ƿ���ر���
    '����:��ʾ����,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 13:11:01
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If mlng���ű���ƽ������ = 0 Then
        strSQL = "Select Avg(length(����)) As ���� From ���ű�"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ���ű����ƽ������")
        mlng���ű���ƽ������ = Val(Nvl(rsTemp!����))
    End If
    '���ڱ��볤�ȿ��ܹ���,�޷���ʾ���ŵ�����,����Զ���ʾ�Ͳ���ʾ����,������5ʱ,����ʾ.С��5ʱ,��ʾ
   zlIsShowDeptCode = mlng���ű���ƽ������ <= 5
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
 
      
    
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!ID = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!ID): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!ID)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!ID)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!ID))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
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
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!ID))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlcontrol.TxtSelAll cboDept
End Function

Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo ErrHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ���֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    Case "Ʊ�����ü�¼"
        zlGetFeeFields = "ID,Ʊ��,ʹ�����,������,ǰ׺�ı�,��ʼ����,��ֹ����,ʹ�÷�ʽ,�Ǽ�ʱ��,ʹ��ʱ��," & _
        "�Ǽ���,��ǰ����,ʣ������,����,�˶���,�˶�ʱ��,�˶Խ��,�˶�ģʽ,��ע,ǩ����,ǩ��ʱ��"
        Exit Function
    Case "Ʊ��ʹ����ϸ"
        zlGetFeeFields = "ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,���մ���,ʹ��ʱ��,ʹ����,�˶���,�˶�ʱ��,�˶Խ��,��ע"
        Exit Function
    Case "��Ա�ɿ��¼"
        zlGetFeeFields = "ID,����ID,�տ�Ա,�տ��ID,���㷽ʽ,�����,���,ժҪ,��ֹʱ��,�Ǽ�ʱ��,�Ǽ���"
        Exit Function
    Case "���ѿ����ü�¼"
        zlGetFeeFields = "ID,�ӿڱ��,������,ǰ׺�ı�,��ʼ����,��ֹ����,ʹ�÷�ʽ,�Ǽ�ʱ��,ʹ��ʱ��," & _
        "�Ǽ���,��ǰ����,ʣ������,����,�˶���,�˶�ʱ��,�˶Խ��,�˶�ģʽ,��ע,ǩ����,ǩ��ʱ��"
        Exit Function
    Case "���ѿ�ʹ�ü�¼"
        zlGetFeeFields = "ID,�ӿڱ��,����,����,ԭ��,����ID,���մ���,ʹ��ʱ��,ʹ����,�˶���,�˶�ʱ��,�˶Խ��,��ע"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo ErrHand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!column_name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
ErrHand:
  zlGetFeeFields = "*"
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function



Public Function Select��Աѡ����(ByVal frmMain As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng����ID As Long = 0, _
    Optional lng��ԱID As Long = 0, _
    Optional bln��������Ա��ʾ As Boolean = False, _
    Optional strSearchKey As String = "", _
    Optional str��Ա���� As String = "", _
    Optional str����ְ�� As String = "", _
    Optional strרҵ����ְ�� As String = "", _
    Optional strTittle As String = "��Աѡ����", _
    Optional strNote As String = "��ѡ����ص���Ա", _
    Optional strNotFindMsg As String = "δ�ҵ�ָ������Ա,����!", _
    Optional strShowField As String = "����", _
    Optional strShowSplit As String = "-") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������Ա
    '���:frmMain-���õĸ�����
    '     objCtl-�ؼ�(Ŀǰֻ֧���ı���)
    '     strKey-����Ľ�ֵ
    '     lng����ID-�����Ϊ��,��������Ա,����, ��ָ�������µ���Ա
    '     str��Ա����: ��ҽ��,ҽ��1... ��ʽ
    '     str����ְ��strרҵ����ְ��: ��ְ��1,ְ��21... ��ʽ
    '����:lng��Աid-������ԱID
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, bytType As Byte, str��Ա����Table As String, strWhere As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmMain=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Err = 0: On Error GoTo ErrHand:
    bytType = 0: strWhere = ""
    If str��Ա���� <> "" Then
        str��Ա����Table = ",��Ա����˵�� Q1,(Select Column_Value From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) Q2" & vbCrLf
        strWhere = strWhere & " And ( A.ID=Q1.��ԱID and Q1.��Ա���� = Q2.Column_Value ) " & vbCrLf
    End If
    If str����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))  Where a.����ְ��=Column_Value) " & vbCrLf
    If strרҵ����ְ�� <> "" Then strWhere = strWhere & "  And Exists(Select 1 From (Select Column_Value From Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))  Where a.רҵ����ְ��=Column_Value) " & vbCrLf
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey, False)
        If lng����ID = 0 Then
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct A.ID,A.���,A.����,A.����,A.����,A.�Ա�,A.����,A.��������,A.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� A " & str��Ա����Table & _
                "   Where (A.���� like [1] or A.��� like [1] or A.���� like Upper([1]) or A.���� like [1]) " & strWhere & zl_��ȡվ������(True, "A") & "" & _
                "       and (A.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                "   order by A.���"
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2]   " & strWhere & zl_��ȡվ������(True, "a") & _
                "       and (a.���� like [1] or a.��� like [1] or a.���� like Upper([1]) or a.���� like [1]) " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)" & _
                "   order by ���"
        End If
     Else
        If lng����ID = 0 Then
            If bln��������Ա��ʾ Then
                gstrSQL = "" & _
                "   Select /*+ rule */  id," & IIf(gstrNodeNo <> "-", "1 as ����ID,-1*NULL as �ϼ�ID", "Level as ����ID,�ϼ�id") & " ,����,����,0 ĩ��,'' as ����,'' as ����,''as �Ա�,''as ����, to_date(Null,'yyyy-mm-dd')  as ��������, '' as  �칫�ҵ绰 ,'' ִҵ���, '' ����ְ��,'' רҵ����ְ��" & _
                "   From ���ű� " & _
                "   where ����ʱ�� is null or ����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') " & zl_��ȡվ������() & _
                    IIf(gstrNodeNo <> "-", "", "   Start with �ϼ�id is null connect by prior id=�ϼ�id ") & _
                "   union all " & _
                "   Select  distinct a.ID,999999 AS ����ID,b.����id as �ϼ�ID,a.���,a.����,1 as ĩ��,����,����,�Ա�,����,��������,�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ�� " & _
                "   From ��Ա�� a,������Ա b  " & str��Ա����Table & _
                "   Where a.id=b.��Աid and b.ȱʡ=1  " & strWhere & zl_��ȡվ������(True, "a") & _
                "         And (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
                "   Order by ����ID,����"
                bytType = 2
            Else
                gstrSQL = "" & _
                    "   Select  /*+ rule */  distinct A.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                    "   From ��Ա�� A " & str��Ա����Table & _
                    "   Where (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & strWhere & zl_��ȡվ������(True, "a") & _
                    "   order by a.���"
            End If
        Else
            gstrSQL = "" & _
                "   Select /*+ rule */ distinct a.ID,a.���,a.����,a.����,a.����,a.�Ա�,a.����,a.��������,a.�칫�ҵ绰,A.ִҵ���,A.����ְ��,A.רҵ����ְ��" & _
                "   From ��Ա�� a,������Ա C " & str��Ա����Table & _
                "   Where a.id=c.��Աid and c.����Id=[2] " & _
                "       and (a.����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & strWhere & zl_��ȡվ������(True, "a") & _
                "   order by a.���"
        End If
    End If
   
   
   '���궨λ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        lngH = objCtl.MsfObj.CellHeight
    Case Else
        vRect = zlcontrol.GetControlRect(objCtl.hWnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = objCtl.Height
    End Select
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytType, strTittle, bytType = 2, strSearchKey, strNote, bytType = 2, False, Not (bytType = 2), sngX, sngY, lngH, blnCancel, False, False, strKey, lng����ID, str��Ա����, str����ְ��, strרҵ����ְ��)
    
    lng��ԱID = 0
    If blnCancel = True Then
        Call zlcontrol.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        If strNotFindMsg <> "" Then ShowMsgbox strNotFindMsg
        Call zlcontrol.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlcontrol.TxtSelAll objCtl
        Exit Function
    End If
    Call zlcontrol.ControlSetFocus(objCtl, True)
    If bytType = 2 Then
        strShowField = "," & strShowField & ",M_��,"
        strShowField = Replace(strShowField, ",���,", ",����,")
        strShowField = Replace(strShowField, ",����,", ",����,")
        strShowField = Mid(strShowField, 2)
        strShowField = Replace(strShowField, ",M_��,", "")
    End If
    
    '������ص�ֵ
    Select Case UCase(TypeName(objCtl))
    Case UCase("VSFlexGrid")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .EditText = .TextMatrix(.Row, .Col)
            .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!ID)
        End With
    Case UCase("BILLEDIT")
        With objCtl
            .TextMatrix(.Row, .Col) = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
            .Text = .TextMatrix(.Row, .Col)
        End With
    Case UCase("ComboBox")
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!ID) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgbox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        objCtl.Text = zl_GetFieldValue(rsTemp, strShowField, strShowSplit)
        objCtl.Tag = Val(rsTemp!ID)
        zlCommFun.PressKey vbKeyTab
    End Select
    lng��ԱID = Val(Nvl(rsTemp!ID))
    rsTemp.Close
    Select��Աѡ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_��ȡվ������ = strWhere
End Function
Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_GetFieldValue(ByVal rsTemp As ADODB.Recordset, _
    Optional ByVal strShowFields As String = "����,����", _
    Optional ByVal strShowSplit As String = "-") As String
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ�ֶε����ֵ
    '���:rsTemp-��¼��
    '     strShowFields-��ʾ���ֶ�
    '     strShowSplit-��ʾ�ķ����
    '����:
    '����:�ɹ�,������ص��ֶ�ֵ
    '����:���˺�
    '����:2009-03-06 11:59:19
    '-----------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, strValue As String, strLeft As String, strRight As String
    varData = Split(strShowFields, ",")
    
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.RecordCount = 0 Then Exit Function
    
    Select Case strShowSplit
    Case "[", "[]", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "�ۣ�", "��", "��"
        strLeft = "��": strRight = "��"
    Case "[]", "[", "]"
        strLeft = "[": strRight = "]"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "{}", "{", "}"
        strLeft = "{": strRight = "}"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case "����", "��", "��"
        strLeft = "��": strRight = "��"
    Case Else
        strLeft = "": strRight = strShowSplit
    End Select
    
    strValue = ""
    With rsTemp
        For i = 0 To UBound(varData) - 1
            strValue = strValue & strLeft & Nvl(.Fields(varData(i))) & strRight
        Next
        strValue = strValue & Nvl(.Fields(varData(UBound(varData))))
    End With
    zl_GetFieldValue = strValue
End Function

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
    i = cllData.Count + 1
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
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function zlIsOnlyNum(ByVal strAsk As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж�ָ���ַ����Ƿ�ȫ�������ֹ���
    '���:strAsk-��Ҫ�жϵ��ַ�
    '����:
    '����:���ȫ�����ֹ��ɣ�����true,���򷵻�False
    '����:���˺�
    '����:2010-11-17 11:19:15
    '˵��:
    '     isnumberic���ܼ����Щ:-099.22,22d2��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            strTemp = Mid(Trim(strAsk), i, 1)
            If InStr("0123456789", strTemp) = 0 Then Exit Function
        Next
        zlIsOnlyNum = True
    End If
End Function

Public Function Get���㷽ʽ() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㷽ʽ
    '����:���㷽ʽ��
    '����:���˺�
    '����:2013-09-04 17:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "" & _
    "   Select ����,����,����,nvl(Ӧ�տ�,0) as Ӧ�տ�,nvl(Ӧ����,0) as Ӧ����," & _
    "               nvl(ȱʡ��־,0) as ȱʡ��־  " & _
    "   From ���㷽ʽ"
    If mrsPayMode Is Nothing Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���㷽ʽ")
    ElseIf mrsPayMode.State <> 1 Then
        Set mrsPayMode = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���㷽ʽ")
    End If
    Set Get���㷽ʽ = mrsPayMode
End Function
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
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub



Public Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
    '���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
    '������intNum=��Ŀ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intTYPE As Integer
    Dim dtCurDate As Date, strMaxNo As String
    Dim strYearStr As String
    Err = 0: On Error GoTo errH:
    
    strSQL = "Select ��Ź���,Sysdate as ����,������ From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", intNum)
    If rsTmp.EOF Then GetFullNO = strNO: Exit Function
    Select Case Val(Nvl(rsTmp!��Ź���))
    Case 0, 1 '0-����˳����,1-����˳����
        If Len(strNO) >= 8 Then
            GetFullNO = Right(strNO, 8)
            Exit Function
        ElseIf Len(strNO) = 7 Then
            GetFullNO = zlStr.PrefixNO & strNO
            Exit Function
        End If
        GetFullNO = strNO
        dtCurDate = Date
        If Not rsTmp.EOF Then
            intTYPE = Val("" & rsTmp!��Ź���)
            dtCurDate = rsTmp!����
            strMaxNo = Nvl(rsTmp!������)
        End If
        strYearStr = zlStr.PrefixNO
        If strMaxNo = "" Then strMaxNo = strYearStr & "000001"
        If intTYPE = 1 Then
            '���ձ��
            strSQL = Format(CDate(Format(dtCurDate, "YYYY-MM-dd")) - CDate(Format(dtCurDate, "YYYY") & "-01-01") + 1, "000")
            GetFullNO = zlStr.PrefixNO & strSQL & Format(Right(strNO, 4), "0000")
            Exit Function
        End If
        '������
        If Len(strNO) = 6 Then
            GetFullNO = Left(strMaxNo, 2) & strNO: Exit Function
        End If
        GetFullNO = Left(strMaxNo, 2) & zlStr.Lpad(Right(strNO, 6), 6, "0", True)
    Case 2  '2-�����ҷ��»���ձ����Ҫ��ȡ���Һ����,
    Case 3   '3-��������+˳���(��ȡ��λ,˳���ȡ4λ)
        If Len(strNO) <= 6 Then
            GetFullNO = Format(rsTmp!����, "YYMMDD") & zlStr.Lpad(strNO, 6, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 8 Then
            GetFullNO = Format(rsTmp!����, "YYMM") & zlStr.Lpad(strNO, 8, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 10 Then
            GetFullNO = Format(rsTmp!����, "YY") & zlStr.Lpad(strNO, 10, "0", True)
            Exit Function
        End If
        If Len(strNO) <= 12 Then
            GetFullNO = zlStr.Lpad(strNO, 12, "0", True)
            Exit Function
        End If
    Case 4    '4-��ִ�п��ҷ��ڼ���(��(�ڼ���е���)+ִ�п��ұ��+�·�(�ڼ���е���)+˳���)
    Case 5    '5-�����½��б��(yyyyMM000000)
    Case Else
    End Select
    GetFullNO = strNO
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get���ʽ�������(ByVal strRollingType As String, _
    ByRef strOut�������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ʵĽ�������
    '���:strRollingType-�������(0-�������(��ȫ������),1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    '����:strOut��������-���ر��εĽ�������,����ö��ŷָ�,����:,2,...
    '     �������������Ԥ�������ѿ�,�򷵻ؿ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-03-05 15:04:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strOut�������� = ""
    On Error GoTo errHandle

    'Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ����
    'strRollingType:1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨,6-���ѿ�)
    If InStr("," & strRollingType & ",", ",1,") > 0 Then  '�շ�
        strOut�������� = "3,6,"
    End If
    If InStr("," & strRollingType & ",", ",3,") > 0 Then  '����
        strOut�������� = strOut�������� & "2,"
    End If
    If InStr("," & strRollingType & ",", ",4,") > 0 Then  '�Һ�
        strOut�������� = strOut�������� & "4,"
    End If
    If InStr("," & strRollingType & ",", ",5,") > 0 Then  '���￨
        strOut�������� = strOut�������� & "5,"
    End If
    If strOut�������� <> "" Then strOut�������� = "," & strOut��������
    Get���ʽ������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CurrentIsBill(ByVal intƱ�� As gBillType) As Boolean
    '����Ʊ���ж��Ƿ�ΪƱ��
    '���أ������Ʊ�ݣ�����TRUE�����򷵻�False
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.Ԥ���վ�, gBillType.�����վ�, gBillType.�Һ��վ�
        CurrentIsBill = True
    Case Else
        CurrentIsBill = False
    End Select
End Function
