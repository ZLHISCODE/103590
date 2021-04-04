Attribute VB_Name = "mdlCustomCommon"
Option Explicit
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
    g��������ģ�� = 5
    g����˽��ģ�� = 6
End Enum
Public gstrMatchMethod As String
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Public Const GCTRL_SELBACK_COLOR = &H8000000D
Private Const GWL_STYLE As Long = (-16&)
Private Const WS_THICKFRAME = &H40000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function SetCursorPos& Lib "user32" (ByVal X&, ByVal Y&)
Private Declare Function ClientToScreen& Lib "user32" (ByVal Hwnd&, lpPoint As POINTAPI)
Private Declare Function GetSystemMenu& Lib "user32" (ByVal Hwnd&, ByVal bRevert&)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd&, ByVal nIndex&)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd&, ByVal nIndex&, ByVal dwNewLong&)

Public Sub zlSetWindowsBroldStyle(ByVal frmMain As Form)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ı�ɵ����岻���ɵ����壨������ֻ�йرհ�ť����,������屾��ֻ�йرգ�ֻ���Զ�������󻯡���С���Ȱ�ť)
    '���:frmMain.hwnd-����ľ��
    '����:
    '����:
    '����:���˺�
    '����:2009-09-29 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim pt_SavePoint As POINTAPI, pt_MovePoint As POINTAPI
    Err = 0: On Error GoTo Errhand:
    With pt_MovePoint
      .X = (-1): .Y = 10
    End With
    '���ô����broldStyle
    Call SetWindowLong(frmMain.Hwnd, GWL_STYLE, GetWindowLong(frmMain.Hwnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call GetSystemMenu(frmMain.Hwnd, 1&)
    '�����ػ�����
    With frmMain
        .Move .Left, .Top, .Width - 15, .Height - 15
        .Move .Left, .Top, .Width + 15, .Height + 15
    End With
    Call GetCursorPos(pt_SavePoint)
    Call ClientToScreen(frmMain.Hwnd, pt_MovePoint)
    Call SetCursorPos(pt_MovePoint.X, pt_MovePoint.Y)
    Call SetCursorPos(pt_SavePoint.X, pt_SavePoint.Y)
Errhand:
End Sub
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

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    If gstrMatchMethod = "" Then
        gstrMatchMethod = Val(zlDatabase.GetPara("����ƥ��"))
    End If
    
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
    GetTaskbarHeight = OS.TaskbarHeight
Errhand:
End Function
Public Function zlGetNextWeekDate(Optional strDate As String = "") As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����һ����
    '����:
    '����:���˺�
    '����:2009-09-21 11:19:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date
    If strDate <> "" Then
        dtDate = DateAdd("d", 7, CDate(strDate))
    Else
        dtDate = DateAdd("d", 7, zlDatabase.Currentdate)
    End If
    dtDate = DateAdd("d", -1 * (Weekday(dtDate, vbMonday) - 1), dtDate)
    zlGetNextWeekDate = CDate(Format(dtDate, "yyyy-mm-dd"))
End Function


