Attribute VB_Name = "mdlPublic"
Option Explicit
'------------------------------------------------------------
'���Ͷ���
'------------------------------------------------------------

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

'ע����Ϣ����
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum
'gclsPros.SecdInfoRec!�ı�״̬
'ɾ���У�������,δ�ı�,�滻�о���ָ����ʼ���ݼ���Ϊ�����༭����
'�����У����е���Ҫ����δ�ı䣬��Ҫ��Ϣ�ı���
'�滻��: ���е���Ҫ��Ϣ�����ı�
Public Enum Change_State
    CS_ɾ���� = -1
    CS_δ�ı� = 0
    CS_������ = 1
    CS_�滻�� = 2
    CS_������ = 3
End Enum
'gclsPros.MainInfoRec!�Ƿ�ı�
Public Enum Main_Change_State
    MS_���ж� = -1
    MS_δ�ı� = 0
    MS_�ı��� = 1
End Enum

'gclsPros.MainInfoRec!ExpState
'��չ���Ƿ����Ϣ�дμ���Ϣ��¼����¼
'��ʼ��չ:�ڴμ���Ϣ��¼����ʼ��ʱ��չ
'������չ:�����ݼ���ʱ��չ
Public Enum Expan_State
    ES_������չ = 0 '����չ���ӡ�Ӵ��Ϊ�˴�������
    ES_��ʼ��չ = 1 '��ʼ��չ
    ES_������չ = 2 '������չ
End Enum

Public Enum DiagMsgPos
    DMP_��ϴ��� = 0
    DMP_������� = 1
    DMP_�������� = 2
    DMP_��ϱ��� = 3
    DMP_�������� = 4
    DMP_������� = 5
    DMP_֤����� = 6
    DMP_֤������ = 7
    DMP_�Ƿ����� = 8
End Enum
'��ҳ����
Public Enum MedRec_Operate
    MOP_���� = 0
    MOP_Ԥ�� = 1
    MOP_��ӡ = 2
    MOP_ȷ�� = 3
End Enum
'�û���Ϣ
Public Type TYPE_USER_INFO
    ID As Long
    ��� As String
    ���� As String '��Ա����
    ���� As String
    DeptID As Long '����ID
    DeptNo As String '���ű��
    DeptName As String '��������
    DBUser As String '���ݿ��û�
End Type
'��������
Public Enum Code_Type
    ' intType����
    'GetNextNo int��� ����;IsHavePageNos,IsPageNosCodeRule:intType����
    CT_����ID = 1
    CT_סԺ�� = 2
    CT_סԺ��ex = 3
    CT_������ = 4
    CT_������ = 5
End Enum

'-----------------------------------------------------------
'����
'------------------------------------------------------------
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
'GetWindowLong,SetWindowLong
Public Const GWL_WNDPROC = -4&
'CallWindowProc
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_MOUSEWHEEL = &H20A '��������Ϣ
Public Const GRD_UNEDITCELL_COLOR = &H8000000B  'δ�༭�ĵ�Ԫ����ɫ������ɫ
Public Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '�뿪����ʱ,ѡ�����ʾ��ɫ
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '����ؼ�ʱ,ѡ����ʾ��ɫ

Public Const CB_GETDROPPEDSTATE = &H157 '��ȡ�����б�״̬
Public Const CB_SHOWDROPDOWN = &H14F '�رջ�������б�

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Public Const GPAGECOLOR = vbWindowBackground
Public Const SW_RESTORE = 9
Public Const SM_CYFULLSCREEN = 17

'----------------------------------------------------------
'API����
'----------------------------------------------------------
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private PrevWndProc     As Long

Public Function DecodeEx(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����,�������仯
'           ǰһλΪBoolean���ͣ���һλΪ����ֵ��������һ��ΪTrue��ֵ�ͷ��أ����ټ����ж�,���һλΪTrue��Ĭ��ֵ
'          �磺ture,3,true,4,�򷵻�3
    Dim i As Integer, blnObjReturn As Boolean
    i = 0
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            If IsObject(arrPar(i)) Then
                Set DecodeEx = arrPar(i): Exit Function
            Else
                DecodeEx = arrPar(i): Exit Function
            End If
        Else
            If arrPar(i) Then
                If IsObject(arrPar(i + 1)) Then
                    Set DecodeEx = arrPar(i + 1): Exit Function
                Else
                    DecodeEx = arrPar(i + 1): Exit Function
                End If
            ElseIf Not blnObjReturn Then
                blnObjReturn = IsObject(arrPar(i + 1))
            End If
            i = i + 2
        End If
    Loop
    If blnObjReturn Then Set DecodeEx = Nothing
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(gclsPros.TXTProc, hwnd, msg, wp, lp)
End Function

Public Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboInfo As Variant, Optional ByVal intDefault As Integer = -1)

'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String����
'      arrCboIdx=ComboBox����,���ComboBoxʱ,װ��������ͬ
'      intDefaut=ȱʡ����
    Dim i As Long, j As Long

    For i = 0 To UBound(arrCboInfo)
        arrCboInfo(i).Clear
        For j = 0 To UBound(arrList)
            arrCboInfo(i).AddItem arrList(j)
        Next
        arrCboInfo(i).ListIndex = intDefault 'ȱʡΪδѡ��
        arrCboInfo(i).Tag = intDefault '����Ĭ��ѡ�������ս����ȷ��Ĭ��ֵ
    Next
End Sub

Public Sub SetCboDefault(objCbo As Object, Optional ByVal intDefault As Integer = -1)
'���ܣ�����Cbo�ؼ���ȱʡֵ
    objCbo.ListIndex = intDefault 'ȱʡΪδѡ��
    objCbo.Tag = intDefault '����Ĭ��ѡ�������ս����ȷ��Ĭ��ֵ
End Sub

Public Sub SetCboDefaultByRec(ByVal arrIndex As Variant)
'���ܣ�ͨ�������ֵ�����Cbo�ؼ���ȱʡֵ
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If TypeName(arrIndex) <> "Variant()" Then
        arrIndex = Array(arrIndex)
    End If
    If TypeName(arrIndex) = "Variant()" Then
        For i = LBound(arrIndex) To UBound(arrIndex)
            Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrIndex(i))
            Set rsTmp = GetBaseCode(arrIndex(i))
            rsTmp.Filter = rsTmp.Filter & " And ȱʡ=1"
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(objCboTmp, Val(rsTmp!ID & ""))
            End If
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function Identity(ByRef lngCount As Long) As Long
'���ܣ�ģ����������
'������lngCount=��������
    lngCount = lngCount + 1
    Identity = lngCount
End Function

Public Function GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, ByVal strKEY As String) As String
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       ���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strKeyValue As String

    On Error GoTo Errhand:

    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKEY, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKEY, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKEY, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKEY, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & UserInfo.DBUser & "\" & strSection, strKEY, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.DBUser & "\" & App.ProductName & "\" & strSection, strKEY, "")
    End Select
    GetRegInFor = strKeyValue
    Exit Function
Errhand:

End Function

Public Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean, Optional tbsInfo As TabStrip) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
    Dim lngColor As Long
    On Error GoTo errH
    
    If gclsPros.FuncType <> f���ѡ�� Then
        Call LocateObjectPage(objTmp)
    Else
        gclsPros.CurrentForm.tabFunc.Tabs(IIf(objTmp.Name = "vsDiagXY", "��ҽ���", "��ҽ���")).Selected = True
    End If
    
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    gclsPros.CurrentForm.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function AddErrInfo(ByVal strMsg As String, ByVal intErr As Integer, ParamArray objErr() As Variant) As Boolean
    Dim i As Long
    Dim clsErrTmp As clsErrInfo
    Dim objTmp As Object
    
    On Error GoTo errH
    
    If gclsPros.FuncType <> fҽ����ҳ And gclsPros.FuncType <> f������ҳ Then
        Exit Function
    End If
    
    Set clsErrTmp = New clsErrInfo

    With clsErrTmp
        .IntErrType = intErr
        .StrErrInfo = strMsg
    End With

    For i = LBound(objErr) To UBound(objErr)
        Set objTmp = objErr(i)
        Call clsErrTmp.AddErrObj(objTmp)
    Next
    If intErr = 0 Then
        clsErrTmp.strErrID = "Error-" & CStr(gColErr.Count + 1)
        gColErr.Add clsErrTmp, clsErrTmp.strErrID
    ElseIf intErr = 1 Then
        clsErrTmp.strErrID = "Warn-" & CStr(gColWarn.Count + 1)
        gColWarn.Add clsErrTmp, clsErrTmp.strErrID
    End If
    
    Set clsErrTmp = Nothing
    AddErrInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
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

Public Function Calc���ڷֽ�ʱ��(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal strִ��ʱ�� As String, ByVal intƵ�ʴ��� As Integer, ByVal intƵ�ʼ�� As Integer, ByVal str�����λ As String, _
    Optional ByVal dat�������� As Date) As String
'���ܣ���ʱ��μ�����εķֽ�ִ��ʱ�估����
'������datBegin-datEnd=Ҫ�����ʱ���,����datBeginӦΪÿ�����ڵĿ�ʼ��׼ʱ��
'      strPause=��ͣ��ʱ���
'      dat��������=��������ʱ��������
'���أ�"ʱ��1,ʱ��2,...."(yyyy-MM-dd HH:mm:ss),ʱ�������Ϊ����
'˵����1.ʱ�����Ҫ�ų���ͣ��ʱ���,����������˶�����
'      2.�������Ǽٶ���ִ��ʱ�估Ƶ��������ȫ��ȷ������¼��㡣
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer

    If InStr(strִ��ʱ��, ",") > 0 Then
        arrNormal = Split(Split(strִ��ʱ��, ",")(1), "-")
        arrFirst = Split(Split(strִ��ʱ��, ",")(0), "-")
    Else
        arrNormal = Split(strִ��ʱ��, "-")
        arrFirst = Array()
    End If

    vCurTime = datBegin

    If str�����λ = "��" Then
        vCurTime = zlCommFun.GetWeekBase(datBegin)
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = zlCommFun.GetWeekBase(dat��������))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False

            '1/8:00-3/15:00-5/9:00
            For i = 1 To intƵ�ʴ���
                If i - 1 <= UBound(arrTime) Then '���ܿ��ܴ�������
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "��" Then
        If dat�������� <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat��������))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False

            If intƵ�ʼ�� = 1 Then
                '8:00-12:00-14:00��8-12-14
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To intƵ�ʴ���
                    If i - 1 <= UBound(arrTime) Then '���տ��ܴ�������
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + intƵ�ʼ��, "yyyy-MM-dd") '������
        Loop
    ElseIf str�����λ = "Сʱ" Then
        '10:00-20:00-40:00��10-20-40��02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To intƵ�ʴ���
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + intƵ�ʼ�� / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str�����λ = "����" Then
        '��ִ��ʱ��
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime

            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + intƵ�ʼ�� / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    Calc���ڷֽ�ʱ�� = Mid(strDetailTime, 2)
End Function


Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'���ܣ��ж�һ��ʱ���Ƿ�����ͣ��ʱ�����
'������strPause="��ͣʱ��,��ʼʱ��;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String

    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '������δ���û���ͣ��ʱ��ֹͣ
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function GetTextByDot(ByVal strText As String, Optional ByVal blnBefore As Boolean, Optional ByVal strSpliter As String = ".") As String
'����: �õ�Բ��֮���֮ǰ���ı�
    If blnBefore Then
        If InStr(strText, strSpliter) > 0 Then
            GetTextByDot = Mid(strText, 1, InStr(strText, strSpliter) - 1)
        End If
    Else
        GetTextByDot = Mid(strText, InStr(strText, strSpliter) + Len(strSpliter))
    End If
End Function

Public Function PrintVsCol(ByRef vsTmp As VSFlexGrid, Optional ByVal strName As String) As String
'���ܣ���ӡ�����,���������п�
    Dim i As Long
    Dim strTmp As String
    With vsTmp
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                strTmp = strTmp & ";" & .TextMatrix(0, i) & "," & .ColWidth(i)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        PrintVsCol = strName & "==" & strTmp
    End With
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'���ܣ���Ҵ�������
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, gclsPros.SysNo, lngMod)
            Call zlPlugInErrH(Err, "Initialize")
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True
End Function

Public Sub SetCtrlLocked(ByRef objInput As Object, ByVal blnLocked As Boolean, Optional ByVal blnClear As Boolean, Optional ByVal blnSetForeColor As Boolean)
'���ܣ�����������ؼ�
'������objInput=�ؼ�����
'         blnLocked=�Ƿ�����
'         blnClear=�Ƿ�ؼ���������
    Dim strType  As String
    Dim objCmd As CommandButton
    Dim strTmp As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            If strType = "TextBox" Then
                'Ѱ�Ҷ�Ӧ��ť
                On Error Resume Next
                If objInput.Name = "txtSpecificInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdSpecificInfo(objInput.Index)
                    strTmp = objCmd.Name
                ElseIf objInput.Name = "txtInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdInfo(objInput.Index)
                    strTmp = objCmd.Name
                ElseIf objInput.Name = "txtAdressInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdAdressInfo(objInput.Index)
                    strTmp = objCmd.Name
                End If
                If Err.Number = 0 Then
                    Call SetCtrlLocked(objCmd, blnLocked)
                    On Error GoTo errH
                Else
                    Err.Clear
                    On Error GoTo errH
                End If
            End If
            If blnClear And blnLocked Then
                If strType = "ComboBox" Then
                    Call zlControl.CboSetIndex(objInput.hwnd, -1)
                Else
                    objInput.Text = ""
                End If
            End If
            objInput.Locked = blnLocked
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.TabStop = Not blnLocked
            If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
        Case "CheckBox"
            If blnClear And blnLocked Then
                objInput.Value = 0
            End If
            objInput.Enabled = Not blnLocked
'            objInput.BackColor = IIf(blnLocked, vbButtonFace, &H8000000F)
            objInput.TabStop = Not blnLocked
            If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
        Case "CommandButton"
            objInput.Enabled = Not blnLocked
        Case "MaskEdBox", "MonthView", "ListBox"
            objInput.Enabled = Not blnLocked
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.TabStop = Not blnLocked
            If blnClear And blnLocked And strType = "MaskEdBox" Then
                objInput.Text = Replace(objInput.Mask, "#", "_")
            End If
            If objInput.Name = "mskDateInfo" Then
                Call SetCtrlLocked(gclsPros.CurrentForm.txtDateInfo(objInput.Index), blnLocked, blnClear, blnSetForeColor)
                gclsPros.CurrentForm.txtDateInfo(objInput.Index).Text = objInput.Text
                gclsPros.CurrentForm.txtDateInfo(objInput.Index).Visible = blnLocked
                objInput.Visible = Not blnLocked
            Else
                If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
            End If
        Case "VSFlexGrid"
            'ͬʱע��Ҫ�ڼ�������¼��н���һЩ����
            objInput.Editable = IIf(blnLocked, flexEDNone, flexEDKbdMouse)
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.BackColorBkg = IIf(blnLocked, vbButtonFace, vbWindowBackground)
        Case "PatiAddress"
            objInput.ControlLock = blnLocked
        Case "OptionButton"
            objInput.Enabled = Not blnLocked
'            objInput.BackColor = IIf(blnLocked, vbButtonFace, &H8000000F)
        Case "Label"
            objInput.ForeColor = IIf(Not blnLocked, gclsPros.CurrentForm.ForeColor, &H808080)
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ControlHaveValue(ByRef objInput As Object) As Boolean
'���ܣ�����������ؼ�
'������objInput=�ؼ�����
'         blnLocked=�Ƿ�����
'         blnClear=�Ƿ�ؼ���������
    Dim strType  As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            ControlHaveValue = objInput.Text <> ""
        Case "CheckBox"
            ControlHaveValue = objInput.Value <> 0
        Case "MaskEdBox"
            ControlHaveValue = IsDate(objInput.Text)
        Case "PatiAddress"
            ControlHaveValue = objInput.Value <> ""
    End Select

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ControlIsLocked(ByRef objInput As Object) As Boolean
'���ܣ�����������ؼ�
'������objInput=�ؼ�����
'         blnLocked=�Ƿ�����
'         blnClear=�Ƿ�ؼ���������
    Dim strType  As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            ControlIsLocked = objInput.Locked
        Case "VSFlexGrid"
            ControlIsLocked = objInput.Editable = flexEDNone
        Case "PatiAddress"
            ControlIsLocked = objInput.ControlLock
        Case Else
            ControlIsLocked = Not objInput.Enabled
    End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub zlVsGridLostFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '���ܣ��뿪����ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Ƿ����Զ�����ɫ������(BackColor)�ķ�ʽ������)
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:03:05
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
             If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
    End With
End Sub

Public Function GetFormat(ByVal strMask As String) As String
    GetFormat = Decode(strMask, "####-##-##", "yyyy-mm-dd", "####-##-## ##:##", "yyyy-mm-dd hh:mm", "####-##-## ##:##:##", "yyyy-mm-dd hh:mm:ss", "##;##", "hh:mm", "")
End Function

Public Function SubWndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'���ε��ؼ����������¼�
    Select Case msg    '��������й���.���֪����������Ϣ,Ҳ�������������.
        Case WM_MOUSEWHEEL
            SubWndProc = 1 '���ε�
            Exit Function
    End Select
    SubWndProc = CallWindowProc(PrevWndProc, hwnd, msg, wParam, lParam)             '������Ϣ����
End Function

Public Sub CallHook(ByVal hwnd As Long)
    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub

Public Sub CallUnhook(ByVal hwnd As Long)
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
End Sub

