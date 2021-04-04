Attribute VB_Name = "mdlLISComm"
Option Explicit

'Public gcnOracle As ADODB.Connection    '�������ݿ�����
Public gstrSQL As String

'Public gstrSysName As String                'ϵͳ����

Public glngExeDeptID As Long 'ִ�п���
Public ParentWnd As Object
Public blnDataReceived As Boolean
'------������ͼ�괦��
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_ACTIVATE = &H6
Public Const WM_KEYDOWN = &H100
Public Const WM_PAINT = &HF

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

'Public Const GWL_EXSTYLE = (-20)
'Public Const WinStyle = &H40000
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4

'ø���ǲ���
Public glngMBDeviceID As Long, gstrMBChannel As String, glngMBNo As Long, gstrMBPosition As String

Private mItem() As Variant

Public Const LOG_������־ = 0
Public Const LOG_ͨѶ��־ = 1
Public Const LOG_δ֪�� = 2

Public pLast������־ As String '�ϴδ�����Ϣ,���ڱ�������ظ�����־
Public pLastͨѶ��־ As String
Public mMakeNoRule As String    '�걾�������ʱ�����

Public gblnFromDB As Boolean ' �Ƿ��Ǵ����ݿ��ȡ����.

Public gobjFSO As New Scripting.FileSystemObject    'FSO����
Public mclsUnzip As New cUnzip
Public mclsZip As New cZip

Public Sub SavePortsSetting()
'���ܣ��������Ӽ��������Ĵ�������
    Dim i As Integer
    Dim strSet As String
    Dim aPorts As Variant
    On Error GoTo errH
    
    strSet = ""
    If gblnFromDB Then
        '���ԭ��������
        Call gobjDatabase.SetPara("������������", "", glngSys, 1208)
        For i = LBound(g����) To UBound(g����)
            '����id , ����, COM��, ������, ����λ, У��λ, ֹͣλ, ����, TCPIP�˿�, IP��ַ, �ַ�ģʽ, ���Ϊ������ID, ����,�Զ�Ӧ��,�ɷ��Ѻ˱걾,ͨѶĿ¼,�Զ������,�Զ������ʿ�,���Ϊͨ����
            If g����(i).ID > 0 Then
                strSet = strSet & ";" & g����(i).ID & "," & g����(i).���� & "," & g����(i).COM�� & "," & g����(i).������ & _
                   "," & g����(i).����λ & "," & g����(i).У��λ & "," & g����(i).ֹͣλ & "," & g����(i).���� & _
                   "," & g����(i).IP�˿� & "," & g����(i).IP & "," & g����(i).�ַ�ģʽ & "," & g����(i).SaveAsID & "," & g����(i).���� & _
                   "," & g����(i).�Զ�Ӧ�� & "," & g����(i).�ɷ��Ѻ˱걾 & "," & g����(i).ͨѶĿ¼ & "," & g����(i).�Զ������ & "," & g����(i).�Զ������ʿ� & "," & g����(i).���Ϊͨ����
            
            
                If Dir(g����(i).ͨѶĿ¼ & "\ReceiveSend.ini") <> "" Then Kill g����(i).ͨѶĿ¼ & "\ReceiveSend.ini"
            End If
        Next
        If strSet <> "" Then
            Call gobjDatabase.SetPara("������������", strSet, glngSys, 1208)
        End If
    Else
        'DeleteSetting "ZLSOFT", "����ģ��", "ZlLISSrv"
        Err = 0: On Error Resume Next
        aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
        On Error GoTo errH
        If IsEmpty(aPorts) Then
            ReDim aPorts(8, 0)
            For i = 0 To 7
                aPorts(i, 0) = "COM" & i + 1
            Next
        End If
        Err = 0: On Error Resume Next
        For i = LBound(aPorts) To UBound(aPorts)
            DeleteSetting "ZLSOFT", "����ģ��\ZLLISSrv", aPorts(i, 0)
            DeleteSetting "ZLSOFT", "����ģ��\ZLLISSrv\" & aPorts(i, 0)
        Next
        On Error GoTo errH
        For i = LBound(g����) To UBound(g����)
            If g����(i).���� = 1 Then
                'TCP
                If g����(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "IP" & g����(i).ID, "")
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Device", g����(i).ID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Enabled", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Host", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "InMode", g����(i).�ַ�ģʽ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "IP", g����(i).IP)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Port", g����(i).IP�˿�)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "SaveAs", g����(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "Auto", g����(i).�Զ�Ӧ��)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "blnSend", g����(i).�ɷ��Ѻ˱걾)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "ReceiveDir", g����(i).ͨѶĿ¼)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "AutoCheckMan", g����(i).�Զ������)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "AutoQCCalc", g����(i).�Զ������ʿ�)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\IP" & g����(i).ID, "SaveAsTonDao", g����(i).���Ϊͨ����)
                End If
            Else
                If g����(i).COM�� > 0 And g����(i).ID > 0 Then
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "COM" & g����(i).COM��, "")
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Device", g����(i).ID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Speed", g����(i).������)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "DataBit", g����(i).����λ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Parity", g����(i).У��λ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "StopBit", g����(i).ֹͣλ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "HandShaking", g����(i).����)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "InputMode", g����(i).�ַ�ģʽ)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "SaveAs", g����(i).SaveAsID)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "Auto", g����(i).�Զ�Ӧ��)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "blnSend", g����(i).�ɷ��Ѻ˱걾)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "ReceiveDir", g����(i).ͨѶĿ¼)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "AutoCheckMan", g����(i).�Զ������)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "AutoQCCalc", g����(i).�Զ������ʿ�)
                    Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & g����(i).COM��, "SaveAsTonDao", g����(i).���Ϊͨ����)
                    
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    MsgBox Err.Description

End Sub

Public Function GetConnectDevs() As Variant
'���ܣ���ȡϵͳ���ӵļ�������
    Dim aSettings() As Variant
    Dim aPorts As Variant, i As Integer, PortIndex As Integer
    Dim lngDeviceID As Long, rsTmp As New adodb.Recordset, rsTmp1 As New adodb.Recordset
    Dim strConnType As String  '��������
    Dim strIP As String, strPort As String 'ip �� Port
    Dim varIPSet As Variant 'IP������
    Dim lngSaveAsID As Long '���Ϊ������ID
    Dim strSaveAsName As String
    
    aSettings = Array()
    
    Err = 0: On Error Resume Next
    aPorts = GetAllSettings("ZLSOFT", "����ģ��\ZlLISSrv")
    On Error GoTo errH
    If IsEmpty(aPorts) Then
        ReDim aPorts(8, 0)
        For i = 0 To 7
            aPorts(i, 0) = "COM" & i + 1
        Next
    End If
   
    If Not IsEmpty(aPorts) Then
        
        ReDim g����(UBound(aPorts))
        
        For i = LBound(g����) To UBound(g����)
            g����(i).ID = 0
            g����(i).IP = "127.0.0.1"
            g����(i).IP�˿� = 6666
            g����(i).SaveAsID = 0
            g����(i).������ = 9600
            g����(i).���� = 1
            g����(i).COM�� = 0
            g����(i).����λ = 8
            g����(i).ֹͣλ = 1
            g����(i).���� = 0
            g����(i).У��λ = "N"
            g����(i).�ַ�ģʽ = 0
            g����(i).���� = 0
            g����(i).�Զ�Ӧ�� = "0"
            g����(i).�ɷ��Ѻ˱걾 = 1
        Next
        
        For i = LBound(aPorts) To UBound(aPorts)
            
            strIP = "": strPort = ""
            lngSaveAsID = 0
            strSaveAsName = ""
            
            lngSaveAsID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "SaveAs", 0))
            If lngSaveAsID > 0 Then
                Set rsTmp1 = gobjDatabase.OpenSqlRecord("Select ���� From �������� where ID=[1]", "ȡ������������", lngSaveAsID)
                Do Until rsTmp1.EOF
                    strSaveAsName = "" & rsTmp1!����
                    rsTmp1.MoveNext
                Loop
            End If
            
            strConnType = aPorts(i, 0)

            If strConnType Like "IP*" Then
                'TCPIP����
                g����(i).���� = 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                
                If lngDeviceID > 0 Then

                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select ���� From �������� Where ID=[1]"
                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
                    If Not rsTmp.EOF Then

                        If Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Enabled", 0)) = 1 Then
                            '������IP��ʽ,���IP�Ͷ˿��Ƿ�Ϸ�
                            strIP = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "IP", "127.0.0.1")
                            strPort = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Port", 6666)
                            g����(i).IP = strIP
                            g����(i).IP�˿� = Val(strPort)
                            g����(i).���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Host", 0))
                            
                            g����(i).�Զ�Ӧ�� = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0")
                            g����(i).�ɷ��Ѻ˱걾 = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                            If Not ValidateIP(strIP) And Not ValidatePort(strPort) Then

                                If UBound(aSettings) = -1 Then
                                    ReDim aSettings(2, 0) As Variant
                                Else
                                    ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                                End If

                                aSettings(0, UBound(aSettings, 2)) = strIP & ":" & strPort
                                aSettings(1, UBound(aSettings, 2)) = "IP " & strIP & " " & rsTmp("����") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                                aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                            End If

                        End If
                    End If
                End If
            ElseIf strConnType Like "COM*" Then
                'COM����
                PortIndex = Val(Mid(aPorts(i, 0), 4)) - 1
                lngDeviceID = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Device", 0))
                g����(i).���� = 0
                g����(i).COM�� = Val(PortIndex + 1)
                If lngDeviceID > 0 Then
                    If rsTmp.State <> adStateClosed Then rsTmp.Close
                    gstrSQL = "Select ���� From �������� Where ID=[1] "
                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
                    If Not rsTmp.EOF Then
                        If UBound(aSettings) = -1 Then
                            ReDim aSettings(2, 0) As Variant
                        Else
                            ReDim Preserve aSettings(2, UBound(aSettings, 2) + 1) As Variant
                        End If
                        aSettings(0, UBound(aSettings, 2)) = PortIndex
                        aSettings(1, UBound(aSettings, 2)) = "COM" & PortIndex + 1 & " " & rsTmp("����") & IIf(strSaveAsName = "", "", " -> " & strSaveAsName)
                        aSettings(2, UBound(aSettings, 2)) = lngDeviceID
                    End If
                
                    With g����(i)
                        .ID = lngDeviceID
                        .������ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Speed", "9600"))
                        .����λ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "DataBit", "8"))
                        .ֹͣλ = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "StopBit", "1"))
                        .У��λ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Parity", "n")
                        .���� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\COM" & aPorts(i, 0), "HandShaking", "0"))
                        .�ַ�ģʽ = GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "InputMode", "0")
                        .SaveAsID = lngSaveAsID
                        .�Զ�Ӧ�� = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "Auto", "0"))
                        .�ɷ��Ѻ˱걾 = Val(GetSetting("ZLSOFT", "����ģ��\ZlLISSrv\" & aPorts(i, 0), "blnSend", "1"))
                    End With
                End If
            End If
        Next
    End If
    
    If UBound(aSettings) > -1 Then GetConnectDevs = aSettings
    Exit Function
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Public Function GetDevices() As adodb.Recordset
'���ܣ���ȡ���м�������
    On Error GoTo DBError
    Set GetDevices = Nothing
    If gstr�������� = "" Then
        gstrSQL = "Select ID,����,����,ͨѶ������ From �������� Order by ID"
    Else
        gstrSQL = "Select * From (Select ID,����,����,ͨѶ������ From �������� Order by ID) where Rownum<=[1]"
    End If
    Set GetDevices = gobjDatabase.OpenSqlRecord(gstrSQL, "�������ݽ���", Val(gstr��������))
    Exit Function
DBError:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Public Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub WriteLog(ByVal ModuleName As String, ByVal ErrorType As Integer, ByVal ErrorNum As Long, ByVal ErrorDesc As String)
    'Module:ģ���������
    'ErrorType:��־����
    'errorNum:����Ż���־���
    'errorDesc:������Ϣ����־��Ϣ
    Dim strSQL As String
    
    Call WriteTxtLog(ErrorType, ModuleName, IIf(ErrorNum = 0, "", " ") & ErrorDesc)
    
End Sub

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "�������ݽ���", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "", Optional ByVal blnMessage As Boolean = True)
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = IIf(blnMessage, WM_MOUSEMOVE, 0)
    t.hIcon = stdIcon
    t.szTip = IIf(Len(strTip) = 0, "�������ݽ���", strTip) & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Public Sub ResultFromFile(ByVal strFile As String, ByVal lngDeviceID As Long, ByVal strSampleNO As String, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31"))

        Dim rsTmp As New adodb.Recordset
        Dim strDevice As String
        Dim objDevice As Object
        Dim aRecord() As String
        Dim i As Integer
        Dim intMicrobe As Integer   '΢���� =1 ��ʾ΢����
        Dim lngExeDeptID As Long
    
100     If Len(Trim(strFile)) = 0 Then Exit Sub
    
102     gstrSQL = "Select ͨѶ������,nvl(΢����,0) as ΢����,ʹ��С��ID From �������� Where ID=[1]"
104     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
106     If Not rsTmp.EOF Then
108     strDevice = rsTmp(0)
110         intMicrobe = Nvl(rsTmp(1), 0)
112         lngExeDeptID = Nvl(rsTmp(2), 0)
        End If
114     If intMicrobe = 1 Then
116         gstrSQL = "Select ͨ������,������ID As ��ĿID, 2 as С��λ��,b.����||nvl(b.����,b.������) as ���� From ����ϸ������ A, �����ÿ����� B Where a.������id = b.Id And a.����id = [1] "
        Else
118         gstrSQL = "Select a.ͨ������, a.��Ŀid, Nvl(a.С��λ��, 2) As С��λ��, b.���� || '-' || Nvl(b.Ӣ����, b.������) As ����," & vbNewLine & _
                        "       LPad(Decode(c.�������, Null, b.����, c.�������), 10, '0') As ����" & vbNewLine & _
                        "From ������Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
                        "Where a.��Ŀid = b.Id And a.��Ŀid = c.������Ŀid And a.����id = [1] " & vbNewLine & _
                        "Order By LPad(Decode(c.�������, Null, b.����, c.�������), 10, '0')"

            '2011-12-07 ���������޸ģ�4/5 - ָ������
        End If
120     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, App.ProductName, lngDeviceID)
    
122     If rsTmp.EOF Then
124         ReDim mItem(1, 0) As Variant
126         mItem(1, 0) = -1
        Else
128         mItem = rsTmp.GetRows
        End If
    
        On Error Resume Next
130     Set objDevice = CreateObject(strDevice)
132     If objDevice Is Nothing Then Call WriteLog("ResultFromFile", LOG_������־, Err.Number, "��������:" & strDevice & "����ʧ��!" & vbNewLine & Err.Description)
    
134     Call WriteLog(strDevice & ".ResultFromFile", LOG_ͨѶ��־, 0, "strFile:" & strFile & vbNewLine & "strSampleNO:" & strSampleNO & vbNewLine & "dtStart:" & CStr(dtStart) & vbNewLine & "dtEnd:" & CStr(dtEnd))
136     aRecord = objDevice.ResultFromFile(strFile, strSampleNO, dtStart, dtEnd)
    
        On Error GoTo errH
        'aRecord�����صļ���������(������������밴���±�׼��֯���)
        '   Ԫ��֮����|�ָ�
        '   ��0��Ԫ�أ�����ʱ��
        '   ��1��Ԫ�أ��������
        '   ��2��Ԫ�أ�������
        '   ��3��Ԫ�أ��걾
        '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
        '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
        '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    
        '�з��ؽ��

138     If UBound(aRecord) > -1 Then
        
            Dim StrUnknow As String, strCaclInfo As String, lngErr As Long, strErr As String
140         For i = 0 To UBound(aRecord)
142             Call WriteLog("mdlLISComm.ResultFromFile", LOG_ͨѶ��־, 0, "��¼" & i & ":" & aRecord(i))
144             If InStr(aRecord(i), "|") > 0 Then
                    '�ļ����ط�ʽ�����Զ������ʿؼ��㣬���Զ����
146                 Call SaveToDataBase(lngDeviceID, lngDeviceID, lngExeDeptID, intMicrobe, 0, "", aRecord(i), mItem, StrUnknow, strCaclInfo, lngErr, strErr)
                
148                 If lngErr <> 0 Then
150                     Call WriteLog("ResultFromFile", LOG_������־, lngErr, strErr & vbCrLf & gstrSQL)
                    End If
                End If
            Next
        End If


        Exit Sub
errH:
    If CStr(Erl()) = 138 And Err.Number = 9 Then
        Call WriteLog("ResultFromFile", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  û�з��ؽ�����")
    Else
152     Call WriteLog("ResultFromFile", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
    End If
End Sub

Private Function GetItemID(ByVal strChannel As String, ByVal vItems As Variant, Optional ByRef iDec As Integer, Optional ByRef strItemName As String) As Long
    'iDec:С��λ��,strItemNmae : ��Ŀ��д�����Ϊ����Ϊ������
    Dim i As Integer
    For i = 0 To UBound(vItems, 2)
        If Trim(Replace(Replace(UCase(strChannel), Chr(10), ""), Chr(13), "")) = _
           Replace(Replace((UCase(vItems(0, i))), Chr(10), ""), Chr(13), "") Then Exit For
    Next
    If i > UBound(vItems, 2) Then
        GetItemID = -1
        iDec = 2
        strItemName = ""
    Else
        GetItemID = CLng(vItems(1, i))
        iDec = Val(vItems(2, i))
        strItemName = vItems(3, i)
    End If
End Function

Public Function ValidateIP(ByVal strIP As String, Optional strErrInfo As String) As Boolean
    '���IP��ַ����ȷ�ԡ�
    
    Dim varIP As Variant
    Dim IPError As Integer
    Dim IPd As Integer
    Dim i As Integer
    
    varIP = Split(strIP, ".")
    If UBound(varIP) <> 3 Then
        IPError = 0
    Else
        For i = 0 To 3
            If Not IsNumeric(varIP(i)) Then
                IPError = 1
                Exit For
            Else
                IPd = CInt(varIP(i))
                If IPd < 0 Or IPd > 255 Then
                    IPError = 2
                    Exit For
                Else
                    IPError = -1
                End If
            End If
        Next i
    End If
    
    ValidateIP = True
    Select Case IPError
        Case -1
            If strIP <> "0.0.0.0" Then
                ValidateIP = False
                strErrInfo = ""
            Else
                strErrInfo = "IP������Ϊ0.0.0.0��"
            End If
        Case 0
            strErrInfo = "IP��ʽ���ԣ�ӦΪXXX.XXX.XXX.XXX������XXXΪ0-255�����֡�"
        Case 1
            strErrInfo = "IP��ַֻ��Ϊ0-255�����֡�"
        Case 2
            strErrInfo = "IP��ַ�ķ�Χֻ��Ϊ0-255֮�䡣"
    End Select
End Function

Public Function ValidatePort(ByVal strPort As String, Optional strErrInfo As String) As Boolean
    '���˿ںŵ���ȷ�ԡ�
    ValidatePort = True
    If Not IsNumeric(Trim(strPort)) Then
        strErrInfo = "�˿ں�ֻ��Ϊ1-65535�����֡�"
    Else
        If Val(Trim(strPort)) > 0 And Val(Trim(strPort)) <= 65535 Then
            ValidatePort = False
            strErrInfo = ""
        Else
            strErrInfo = "�˿ںŵķ�Χֻ����1-65535֮�䡣"
        End If
    End If
End Function

Private Sub WriteTxtLog(ByVal lng���� As String, ByVal str��Ŀ As String, ByVal str���� As String)
    '���±������ڼ�¼���ýӿڵ����
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    Dim objFileSystem As New FileSystemObject
    Dim blnClearData As Boolean
    
    '���ж��Ƿ���ڸ��ļ����������򴴽�������=0��ֱ���˳���������������������Ϣ��
    'If Val(GetSetting("ZLSOFT", "zlLisLog", "Test", 0)) = 0 Then Exit Sub
    
    blnClearData = gblnClearData
    
    '������־(����ʱ��,��������,�����,������Ϣ
    If str��Ŀ <> "" Or str���� <> "" Then
        
        If lng���� = LOG_������־ Then
            '������־
            strFileName = App.Path & "\zlLis������־_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLast������־ = str��Ŀ & "|" & str���� Then
                Exit Sub
            Else
                pLast������־ = str��Ŀ & "|" & str����
            End If
        ElseIf lng���� = LOG_ͨѶ��־ Then
            'ͨѶ��־
            
            If blnClearData Then Exit Sub '���������־ѡ���д��־
            strFileName = App.Path & "\zlLisͨѶ��־_" & Format(date, "yyyyMMdd") & ".LOG"
            If pLastͨѶ��־ = str��Ŀ & "|" & str���� Then
                Exit Sub
            Else
                pLastͨѶ��־ = str��Ŀ & "|" & str����
            End If
        ElseIf lng���� = LOG_δ֪�� Then
            'δ֪��
            If blnClearData Then Exit Sub '���������־ѡ���д��־
            strFileName = App.Path & "\zlLisδ֪��Ŀ_" & Format(date, "yyyyMMdd") & ".LOG"
        End If
        
        If Not objFileSystem.FileExists(strFileName) Then Call objFileSystem.CreateTextFile(strFileName)
        Set objStream = objFileSystem.OpenTextFile(strFileName, ForAppending)
        
        
        strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine ("ʱ��:" & strDate & " �汾:" & App.major & "." & App.minor & "." & App.Revision)
        
        objStream.WriteLine (str��Ŀ)
        objStream.WriteLine (str����)
        
        'objStream.WriteLine (String(50, "-"))
        objStream.Close
        Set objStream = Nothing
    End If
End Sub

Public Sub SaveImg(ByVal lngDevID As Long, ByVal lngID As Long, ByVal strImg As String)
        '����ͼ�����ݵ����ݿ���
        'lngDevID   ����ID
        'lngID      �걾ID
        'strImg     ͼ������
    
        Dim aGraphItem() As String
        Dim strImageVal As String
        Dim strImageType As String
        Dim strImageData As String
        Dim intLoop As Integer
        Dim IntCount As Integer
        Dim blnDeleImg As Boolean '������Ƿ�ɾ��ԭ����ͼƬ
        Dim strPicPath As String, strSQL() As String
        Dim intLayOut As Integer 'ͼƬ����ʾ��ʽ
        Dim strBMPFile As String
        Dim blnFtp As Boolean       'FTP�Ƿ����
        Static strFtpPara As String       '����FTP����
        Dim strFTPuser As String, strFTPpass As String, strFTPIP As String, strFPTPath As String
        Dim strUploadOk As String, strFTPDir As String, strNewName As String
        Dim objStream As TextStream
    
        On Error GoTo ErrHandle
    
        'FTP���Ӽ�飬��Ч����԰�FTP��ʽ����ͼƬ
100     blnFtp = False
102     If strFtpPara = "" Then
104         strFtpPara = gobjDatabase.GetPara("FTP����", glngSys, 1208, "")
        End If
106     If UBound(Split(strFtpPara, ";")) >= 3 Then
108        strFTPuser = Split(strFtpPara, ";")(0)
110        strFTPpass = Split(strFtpPara, ";")(1)
112        strFTPIP = Split(strFtpPara, ";")(2)
114        strFPTPath = Split(strFtpPara, ";")(3)
116        If TestFTP(strFTPuser, strFTPpass, strFTPIP, strFPTPath) = "" Then
118             blnFtp = True
           End If
        End If
    
120     aGraphItem = Split(strImg, "^")
    
    
122     For intLoop = 0 To UBound(aGraphItem)
124         strImageVal = Replace(aGraphItem(intLoop), vbCrLf, "")
126         strImageType = Mid(strImageVal, 1, InStr(strImageVal, ";") - 1)
128         strImageData = Mid(strImageVal, InStr(strImageVal, ";") + 1)
        
130         If Mid(strImageData, 1, InStr(strImageData, ";") - 1) >= 100 And Mid(strImageData, 1, InStr(strImageData, ";") - 1) <= 227 Then
                '��֯ͼƬ����
            
132             intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
134             strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
            
136             If InStr(strPicPath, ";") > 0 Then
138                 If Left(strPicPath, 2) = "1;" Then
140                     blnDeleImg = True
                    End If
142                 strPicPath = Mid(strPicPath, InStr(strPicPath, ";") + 1)
                End If
            
144             If Dir(strPicPath) <> "" Then
146                 If UCase(Right(strPicPath, 4)) = ".BMP" And intLayOut >= 100 And intLayOut <= 107 Then
148                     strBMPFile = strPicPath
150                 ElseIf (UCase(Right(strPicPath, 4)) = ".JPG" Or UCase(Right(strPicPath, 4)) = ".GIF") And intLayOut >= 110 And intLayOut <= 127 Then
152                     strBMPFile = strPicPath
154                 ElseIf intLayOut >= 200 And intLayOut <= 227 Then
156                     strPicPath = UCase$(strPicPath)
158                     strBMPFile = zlFileZip(strPicPath)
                    Else
160                     frmLISSrv.picTmp.Picture = LoadPicture(strPicPath)
162                     If Dir(App.Path & "\zlLisIn.bmp") <> "" Then Kill App.Path & "\zlLisIn.bmp"
164                     SavePicture frmLISSrv.picTmp.Picture, App.Path & "\zlLisIn.bmp"
166                     strBMPFile = App.Path & "\zlLisIn.bmp"
                    End If
                
168                 If Not blnFtp Then
                        '���浽���ݿ�
170                     If zlLisBlobSql(lngID, strImageType, strBMPFile, intLayOut, strSQL) Then
172                         WriteLog "ִ�� SaveImg", LOG_ͨѶ��־, 0, "��ʼʱ��"
174                         For IntCount = LBound(strSQL) To UBound(strSQL)
176                             If strSQL(IntCount) <> "" Then
178                                 gstrSQL = strSQL(IntCount)
180                                 gobjDatabase.ExecuteProcedure Replace(strSQL(IntCount), "Call", ""), "����ͼ������"
                                End If
                            Next
182                         WriteLog "ִ�� SaveImg", LOG_ͨѶ��־, 0, "����ʱ��"
                        End If
                    Else
                        '���浽FTP
                        'ͼ��λ�ñ�������ݸ�ʽΪ��ͼ���ʽ;FTP�ļ�·��
                        'ͼ���ʽΪ100-227 ��
184                     strFTPDir = strFPTPath & IIf(Right(strFPTPath, 1) = "/", "", "/") & "Dev_" & lngDevID & "/" & Format(gobjDatabase.Currentdate, "yyyyMM")
186                     strNewName = lngID & "_" & strImageType & Right(strPicPath, 4)
188                     strUploadOk = UploadFile(strFTPuser, strFTPpass, strFTPIP, strFTPDir, strBMPFile, strNewName)
190                     If strUploadOk = "" Then
192                         gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "',Null,0,1,'" & _
                             intLayOut & ";" & strFTPDir & "/" & strNewName & "')"
194                         gobjDatabase.ExecuteProcedure gstrSQL, "�������ͼ������"
                        Else
196                         WriteLog "�ϴ�ͼƬ�ļ���FTP", LOG_ͨѶ��־, 0, strUploadOk
                        End If
                    End If
                
198                 If blnDeleImg Then
200                     IntCount = 0
202                     Do While Dir(strPicPath) <> "" And IntCount < 100
204                         IntCount = IntCount + 1
206                         gobjFSO.DeleteFile strPicPath, True
                        Loop
208                     If Dir(strPicPath) <> "" Then
210                         Call WriteLog("SaveImg", LOG_������־, 0, "�ļ�" & strPicPath & "���ƻ��������������ã�δ��ɾ�������º��ֹ�ɾ����")
                        End If
                    End If
                End If
            Else
                'ͼ������
212             If Not blnFtp Then
214                 If Len(strImageData) > 2000 Then
                        '�������2000��������
216                     For IntCount = 1 To CInt(Len(strImageData) / 1000) + 1
218                         If Len(strImageData) > 0 Then
                            
220                             gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "','" & _
                                                        Mid(strImageData, IntCount * 1000 - 999, 1000) & "'," & _
                                                        "1," & IntCount & ")"
222                             gobjDatabase.ExecuteProcedure gstrSQL, "����ͼ�񱣴�"
                            End If
                        Next
                    Else
224                     gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "','" & strImageData & "',0,1)"
226                     gobjDatabase.ExecuteProcedure gstrSQL, "����ͼ�񱣴�"
                    End If
                Else
                    '��ΪTXT�ļ�Ȼ���ϴ�
228                 intLayOut = Mid(strImageData, 1, InStr(strImageData, ";") - 1)
230                 strPicPath = Mid(strImageData, InStr(strImageData, ";") + 1)
                
232                 strBMPFile = App.Path & "\" & lngID & "_" & strImageType & ".txt"
234                 If gobjFSO.FileExists(strBMPFile) Then gobjFSO.DeleteFile strBMPFile
236                 Set objStream = gobjFSO.CreateTextFile(strBMPFile)
238                 objStream.Write strPicPath
240                 objStream.Close
242                 Set objStream = Nothing
                
                    '���浽FTP
                    'ͼ��λ�ñ�������ݸ�ʽΪ��ͼ���ʽ;FTP�ļ�·��
                    'ͼ���ʽΪ0-6
                
244                 strFTPDir = strFPTPath & IIf(Right(strFPTPath, 1) = "/", "", "/") & "Dev_" & lngDevID & "/" & Format(gobjDatabase.Currentdate, "yyyyMM")
246                 strNewName = lngID & "_" & strImageType & ".txt"
248                 strUploadOk = UploadFile(strFTPuser, strFTPpass, strFTPIP, strFTPDir, strBMPFile, strNewName)
250                 If strUploadOk = "" Then
252                     gstrSQL = "Zl_����ͼ����_Update(" & lngID & ",'" & strImageType & "',Null,0,1,'" & _
                         intLayOut & ";" & strFTPDir & "/" & strNewName & "')"
254                     gobjDatabase.ExecuteProcedure gstrSQL, "�������ͼ������"
                    Else
256                     WriteLog "�ϴ�ͼ�������ļ���FTP", LOG_ͨѶ��־, 0, strUploadOk
                    End If
258                 If gobjFSO.FileExists(strBMPFile) Then gobjFSO.DeleteFile strBMPFile
                End If
            End If
        Next

        Exit Sub
ErrHandle:
260     Call WriteLog("SaveImg-" & CStr(Erl()), LOG_������־, Err.Number, Err.Description)

End Sub


Private Function zlLisBlobSql(ByVal Action As Long, ByVal KeyWord As String, ByVal strFile As String, ByVal layOut As Integer, ByRef arySql() As String) As Boolean
    '���ɱ���ͼƬ��SQL
    'Action ����ID
    'KeyWord ����
    'strFile ͼƬ�ļ�
    'arySql ���ɵ�SQL����ڴ�������
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    Dim lngLBound As Long, lngUBound As Long    '�����������С����±�
    Err = 0: On Error Resume Next
    lngLBound = LBound(arySql): lngUBound = UBound(arySql)
    If Err <> 0 Then lngLBound = 0: lngUBound = -1
    Err = 0: On Error GoTo 0
    
    lngFileNum = FreeFile
    WriteLog "����BlobSQL", LOG_ͨѶ��־, 0, "��ʼʱ��"
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    Err = 0: On Error GoTo errHand
    conChunkSize = 512
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    
    ReDim Preserve arySql(lngLBound To lngUBound + lngBlocks + 1)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        strText = Join(aryHex, "")
        If strText <> "" Then
            If lngCount = 0 Then strText = layOut & ";" & strText
            arySql(lngUBound + lngCount + 1) = "Zl_����ͼ����_Update(" & Action & ",'" & KeyWord & "','" & strText & "',1," & IIf(lngCount = 0, 1, 0) & ")"
        End If
    Next
    Close lngFileNum
    WriteLog "����BlobSQL", LOG_ͨѶ��־, 0, "����ʱ��"
    zlLisBlobSql = True
    Exit Function

errHand:
    Close lngFileNum
    zlLisBlobSql = False
End Function
Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1, Optional ByVal BeginDate As String) As String
    '-----------------------------------------------------------------------------------------
    '����:��ȡ����ʱ��
    '����:
    '-----------------------------------------------------------------------------------------
    Dim intDay As Integer
    Dim dateNow As Date
    
    If BeginDate = "" Then
        dateNow = gobjDatabase.Currentdate
    Else
        dateNow = BeginDate
    End If
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(dateNow, "YYYY-MM-DD")))
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 2, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 8 - intDay, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dateNow, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(dateNow, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(dateNow, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(dateNow, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(dateNow, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(dateNow, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(dateNow, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "���ظ�"
        If bytFlag = 1 Then
            GetDateTime = "2000-01-01 00:00:00"
        Else
            GetDateTime = "3000-12-31 23:59:59"
        End If
    End Select
    
End Function

Public Function CreateSample(ByVal lngDeviceID As Long, ByVal strBarcode As String, _
    ByRef strSampleNO As String, ByVal dtSampleDate As Date, ByVal intType As Integer) As Boolean
        'inttype=0
        Dim strSQL As String, rsTmp As adodb.Recordset, rs As New adodb.Recordset
        Dim lngKey As Long, strItemRecords As String
        Dim lngDeptID As Long '��ǰ��������
        Dim rsItem As New adodb.Recordset
        Dim strItem As String                           '������Ŀ
        Dim str���� As String, str�Ա� As String, str���� As String
        On Error GoTo DBErr
    
100     CreateSample = False
    
        '������������
102     strSQL = "Select ʹ��С��id From �������� Where ID = [1]"
104     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "��������걾", lngDeviceID)
106     lngDeptID = glngExeDeptID
108     If Not rsTmp.EOF Then
110         lngDeptID = Nvl(rsTmp("ʹ��С��id"), glngExeDeptID)
        End If
    
112     If Val(strSampleNO) <= 0 Then
114         strSampleNO = Val(CalcNextCode(lngDeviceID, 0, intType))
        End If

        '���ҷ����������Ŀָ��
    '    strSql = "Select A.���ID AS ID," & _
            "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����,A.�Ա�,A.����,F.No," & _
            "I.������ĿID As ��ĿID,Decode(I.�������,3,Nvl(I.Ĭ��ֵ,'-'),2,I.Ĭ��ֵ,'') As ���,'' As ��־," & _
            "Trim(REPLACE(REPLACE(' '||zlGetReference(I.������ĿID,A.�걾��λ,DECODE(A.�Ա�,'��',1,'Ů',2,0),C.��������,Y.����ID,A.����),' .','0.'),'��.','��0.')) AS ����ο�," & _
            "NVL(A.������־,0) AS ����,F.����ʱ��,F.������ " & _
            "FROM ����ҽ����¼ A," & _
            "������Ϣ C,����ҽ������ F,���鱨����Ŀ G,������Ŀ I,����������Ŀ Y " & _
            "WHERE A.������� = 'C' " & _
            "AND A.����ID=C.����ID " & _
            "AND A.���id IS NOT NULL " & _
            "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
            "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null AND G.������Ŀid=Y.��Ŀid(+) " & _
            "AND G.������ĿID=I.������ĿID " & _
            "AND (Y.����ID+0=[1] Or (Y.����ID Is Null And F.ִ�в���ID=[3])) " & _
            "And F.��������=[2] "
    '        "AND F.ִ��״̬=0 "
    
116     strSQL = "Select ID, ����, �Ա�, ����, NO, ��Ŀid, ���, ��־, ����ο�, ����, ����ʱ��, ������, Rownum As �������, ������Ŀid," & vbNewLine & _
                "       ����,�걾��λ,��������ID,����ҽ��,��ʶ��,��ǰ����,���˿��� " & vbNewLine & _
                "From (Select A.���id As ID, C.���� || Decode(A.Ӥ��, 0, '', Null, '', '(Ӥ��)') As ����, A.�Ա�, A.����, F.NO," & vbNewLine & _
                "              I.������Ŀid As ��Ŀid, Decode(I.�������, 3, Nvl(I.Ĭ��ֵ, '-'), 2, I.Ĭ��ֵ, '') As ���, '' As ��־," & vbNewLine & _
                "              Trim(Replace(Replace(' ' || Zlgetreference(I.������Ŀid, A.�걾��λ, Decode(A.�Ա�, '��', 1, 'Ů', 2, 0)," & vbNewLine & _
                "                                                          C.��������, Y.����id, A.����), ' .', '0.'), '��.', '��0.')) As ����ο�," & vbNewLine & _
                "              Nvl(A.������־, 0) As ����, F.����ʱ��, F.������, G.�������, A.������Ŀid, M.����, " & vbNewLine & _
                "              a.�걾��λ,��������ID,����ҽ��,decode(a.������Դ,2,c.סԺ��,c.�����) as ��ʶ��,c.��ǰ����,l.���� as ���˿��� " & vbNewLine & _
                "       From ����ҽ����¼ A, ������Ϣ C, ����ҽ������ F, ���鱨����Ŀ G, ������Ŀ I, ����������Ŀ Y, ������ĿĿ¼ M ,���ű� L " & vbNewLine & _
                "       Where A.������� = 'C' And A.����id = C.����id And A.���id Is Not Null And A.ҽ��״̬ = 8 And A.ID = F.ҽ��id And" & vbNewLine & _
                "             A.������Ŀid = G.������Ŀid And G.ϸ��id Is Null And G.������Ŀid = Y.��Ŀid(+) And" & vbNewLine & _
                "             G.������Ŀid = I.������Ŀid And A.������Ŀid = M.ID(+) And a.���˿���ID = l.ID" & vbNewLine & _
                "             and (Y.����id + 0 = [1] Or (Y.����id Is Null And F.ִ�в���id = [3])) And nvl(F.ִ��״̬,0) = 0  And F.�������� = [2]" & vbNewLine & _
                "       Order By M.����, G.�������)"

118     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "��������걾", lngDeviceID, strBarcode, lngDeptID)
120     If rsTmp.EOF Then Exit Function
    
122     gstrSQL = "Select B.����id, B.��ҳid, B.���, B.Ӥ������, B.Ӥ���Ա�" & vbNewLine & _
                        "From ����ҽ����¼ A, ������������¼ B" & vbNewLine & _
                        "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.Ӥ�� = B.��� And A.���id = [1] And Rownum = 1"
124     Set rs = gobjDatabase.OpenSqlRecord(gstrSQL, "CreateSample", CLng(rsTmp("ID")))
126     If rs.EOF = False Then
128         str���� = Nvl(rs("Ӥ������"))
130         str�Ա� = Nvl(rs("Ӥ���Ա�"))
132         str���� = "Ӥ��"
        Else
134         str���� = Nvl(rsTmp("����"))
136         str�Ա� = Nvl(rsTmp("�Ա�"))
138         str���� = Nvl(rsTmp("����"))
        End If
    
        '����������Ŀ
140     gstrSQL = "select distinct ҽ������ from ����ҽ����¼ a , ����ҽ������ b, ���鱨����Ŀ c , ����������Ŀ d " & vbNewLine & _
                  "  where a.id = b.ҽ��ID and a.���id is not null and a.������ĿID = c.������ĿID and " & vbNewLine & _
                  "  c.������ĿID = d.��ĿID(+) and  (d.����id + 0 = [1] Or (d.����id Is Null And b.ִ�в���id = [3])) and b.�������� = [2] "
142     Set rsItem = gobjDatabase.OpenSqlRecord(gstrSQL, "��������걾_1", lngDeviceID, strBarcode, lngDeptID)
144     Do Until rsItem.EOF
146         strItem = strItem & " " & Nvl(rsItem("ҽ������"))
148         rsItem.MoveNext
        Loop
150     strItem = Trim(strItem) & "(" & Nvl(rsTmp("�걾��λ")) & ")"
        
        '�����걾��¼
152     lngKey = gobjDatabase.GetNextId("����걾��¼")
154     gstrSQL = "ZL_����걾��¼_�걾����(" & lngKey & "," & _
            rsTmp("ID") & ",'" & rsTmp("ID") & "',0,'" & _
            strSampleNO & "'," & _
            IIf(IsNull(rsTmp("����ʱ��")), "Null", "TO_DATE('" & Format(rsTmp("����ʱ��"), "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & "," & _
            IIf(IsNull(rsTmp("������")), "Null", "'" & rsTmp("������") & "'") & "," & _
            lngDeviceID & "," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),Null," & _
            "'" & _
            gstrUserName & "'," & _
            "TO_DATE('" & Format(dtSampleDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),0," & _
            intType & ",NULL,'" & _
            str���� & "','" & str�Ա� & "','" & str���� & "','" & Nvl(rsTmp("No")) & "','" & _
            Nvl(rsTmp("�걾��λ")) & "'," & Nvl(rsTmp("��������ID")) & ",'" & Nvl(rsTmp("����ҽ��")) & "'," & _
            Nvl(rsTmp("��ʶ��")) & ",'" & Nvl(rsTmp("��ǰ����")) & "','" & Nvl(rsTmp("���˿���")) & "','" & _
            strItem & "',Null,Null," & lngDeptID & ",'" & gstrUserCode & "','" & gstrUserName & "')"
156     gobjDatabase.ExecuteProcedure gstrSQL, "��������걾"
                                                                
        '��дָ��
158     strItemRecords = ""
160     Do While Not rsTmp.EOF
162         strItemRecords = strItemRecords & "|" & rsTmp("ID") & "^" & rsTmp("��ĿID") & "^" & _
                Nvl(rsTmp("���")) & "^" & Nvl(rsTmp("��־"), 0) & "^" & Nvl(rsTmp("����ο�")) & "^" & _
                Nvl(rsTmp("������ĿID")) & "^" & Nvl(rsTmp("�������"))
            
164         rsTmp.MoveNext
        Loop
    
166     If Len(strItemRecords) > 0 Then
168         strItemRecords = Mid(strItemRecords, 2)
            
170         gstrSQL = "Zl_������ͨ���_Write(" & lngKey & "," & _
                lngDeviceID & ",'" & strItemRecords & "',0,0)"
172         gobjDatabase.ExecuteProcedure gstrSQL, "��������걾"
        End If
        Exit Function
DBErr:
174     Call WriteLog("clsLISComm.CreateSample", LOG_������־, Err.Number, CStr(Erl()) & "��," & Err.Description)
End Function

Private Function CalcNextCode(ByVal lngKey As Long, ByVal intRow As Integer, ByVal iType As Integer) As String
    '--------------------------------------------------------------------------------------------------------
    '����:����ָ�������ڵ����ڵ���һ��ȱʡ�걾��
    '����:lngKey                ��������ID
    '     iType                 �걾���0=��ͨ��1=����
    '����:ȱʡ�걾����
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New adodb.Recordset
    Dim strToday As String
    Dim strTmp As String
    Dim lng���� As Long
    Dim strLabNo As String, strLabQCNo As String '����걾���ʿر걾
    Dim mstrSQL As String, mlngLoop As Long
    Dim mlngDefaultItemID As Long
    
    'ʱ��,����,�걾��
    On Error GoTo errHand
    mlngDefaultItemID = 0
    strToday = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
    
    On Error GoTo point1
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                        IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                           CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("������"))
    
    On Error GoTo errHand
    GoTo point2
    
point1:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾id(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & " And ҽ��ID Is Not Null" & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabNo = gobjCommFun.Nvl(rs("������"))
    
point2:
    On Error GoTo point3
    
    mstrSQL = "SELECT NVL(MAX(TO_NUMBER(�걾���)),0) AS ������ FROM ����걾��¼ a,����������Ŀ b " & _
                "WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id= [1] ") & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("������"))
    
    On Error GoTo errHand
    GoTo point4
    
point3:
    On Error GoTo errHand
    
    mstrSQL = "SELECT NVL(MAX(�걾���),'') AS ������ FROM ����걾��¼ a,����������Ŀ b" & _
                " WHERE ����ʱ�� BETWEEN [2] and [3] And a.id = b.�걾ID(+) And nvl(a.�Ƿ��ʿ�Ʒ,0) = 0 " & _
                    IIf(lngKey = -1, " AND ����id IS NULL " & _
                    IIf(mlngDefaultItemID > 0, " And b.������Ŀid = [4] ", ""), "AND ����id=[1] ") & _
                    IIf(iType = 1, " And �걾���=1", " And Nvl(�걾���,0)<>1")
    Set rs = gobjDatabase.OpenSqlRecord(mstrSQL, "����", lngKey, CDate(Format(strToday & " 00:00:00", "yyyy-MM-dd hh:mm:ss")), _
                            CDate(Format(strToday & " 23:59:59", "yyyy-MM-dd hh:mm:ss")), mlngDefaultItemID)
    
    If Not rs.EOF Then strLabQCNo = gobjCommFun.Nvl(rs("������"))
    
point4:
    If strLabNo >= strLabQCNo Then
        CalcNextCode = strLabNo
    Else
        CalcNextCode = strLabQCNo
    End If
'    If Val(strLabQCNo) > Val(strLabNo) + 100 Then CalcNextCode = strLabNo

'    For mlngLoop = 1 To vsf2.Rows - 1
'        If mlngLoop <> intRow Then
'            If Val(vsf2.RowData(mlngLoop)) = lngKey Then
'                If Val(CalcNextCode) < Val(vsf2.TextMatrix(mlngLoop, 2)) Then
'                    CalcNextCode = Val(vsf2.TextMatrix(mlngLoop, 2))
'                End If
'            End If
'        End If
'    Next
'
    If Val(CalcNextCode) <= 0 Then
        CalcNextCode = "1"
        Exit Function
    End If
'
    CalcNextCode = Val(CalcNextCode) + 1
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String, Optional ByVal strUnZipFile As String) As String
    Dim strZipPath As String
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If gobjFSO.FileExists(strUnZipFile) Then gobjFSO.DeleteFile strUnZipFile
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strUnZipFile) <> "" Then
        zlFileUnzip = strUnZipFile
    Else
        zlFileUnzip = ""
    End If
End Function
'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLLIS" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function

Public Function SaveToDataBase(ByVal lngDeviceID As Long, ByVal lngMainID As Long, ByVal lngExeDeptID As Long, ByVal intMicrobe As Integer, ByVal intAutoQCCalc As Integer, ByVal strAutoCheckMan As String, ByVal strResult As String, ByVal vItems As Variant, ByRef strUnknown As String, ByRef strAutoCaleInfo As String, ByRef lngErr As Long, ByRef strErr As String, Optional ByRef strIDs As String, Optional ByRef strlogs As String) As Boolean
        '�������ݵ����ݿ�
        'lngDeviceID :����ID
        'lngExeDeptID : ����С��ID
        ' intMicrobe: �Ƿ�΢����
        ' intAutoQCCalc  :�Ƿ�Ҫ�Զ������ʿر걾
        ' strAutoCheckMan: �Զ������
        ' strResult : ��������
        ' strUnknown�� ���� δ֪��
        ' strAutoCaleInfo : �Զ�����Ľ����Ϣ
        ' lngErr :�����
        ' strErr :����
        ' strIDs As String ԭʼ���ݶ�Ӧ�ļ����¼ID�����ܶ����,���ڴ���ͨѶ�з���ǰ̨����ˢ��
      '�������ݵ����ݿ�
      Dim aRecord() As String, aItem() As String
      Dim aTmp() As String
      Dim strDate As String, strSampleID As String, strBarcode As String
      Dim strName As String, strSample As String, strSex As String, strBirth As String
      Dim i As Long, j As Long
      Dim rsTmp As New adodb.Recordset, strSQL As String

      Dim lngID As Long

      Dim blnAuditing As Boolean '�Ƿ����
      Dim intItemAuditing As Integer 'ָ����ˣ�1-����0-δ��
      Dim strItemAuditing As String 'ָ���������
      Dim strItemCode As String     'ָ��ͨ����
      Dim lngItemID As Long '��ĿID
      Dim strItemRecords As String
      Dim aNos() As String, iType As Integer '�걾������
      Dim aQC() As String                    '�ʿ�����
  
      Dim iDec As Integer 'С��λ��
      Dim blnQryWithSampleNO As Boolean
      Dim strδ֪�� As String
      Dim strStartDate As String
      Dim strEndDate As String
    
      Dim strQCList() As String '������Ҫ���������
      Dim strAutoCheck() As String  '����Ҫ�Զ���˵�����
      Dim int2Verify As Integer
    
      Dim strBatchSQL() As String '����Ҫִ�е�SQL
      Dim str���鱸ע As String
      Dim bln�����걾 As Boolean  '�����걾�����Զ����
      Dim strLog  As String '���������־��д��ָ��Ŀ¼
      Dim strItemName As String 'ָ������
      Dim strItemsInfo As String '����ָ����Ϣ
      
      On Error GoTo DBError
100   ReDim strQCList(0) As String
102   ReDim strAutoCheck(0) As String
      Dim intMicrobeDay As Integer   '΢����������ѯ


    
104    SaveToDataBase = False
       intMicrobeDay = gobjDatabase.GetPara("΢�����ѯʱ��", 100, 1208, 0)
106    int2Verify = gobjDatabase.GetPara("ʹ�ö����������", 100, 1208, 0)
    
108    strLog = Format(Now, "yyyy-MM-dd HH:mm:ss") & " ��ʼ���"
110    If Len(strResult) > 0 Then
       
112        aRecord = Split(strResult, "||")
114        For i = 0 To UBound(aRecord)
116            ReDim strBatchSQL(0) As String
118            blnAuditing = False
            
    '118            Call Return_Decode(aRecord(i))   '���ؽ��������ͨѶ���
120            aTmp = Split(aRecord(i), vbCrLf)
            
122            aItem = Split(aTmp(0), "|")
124            aQC = Split(aItem(4), "^")              '����ʿ�
126            If UBound(aItem) >= 4 Then
                  '��Ч�ı�����
128                aNos = Split(aItem(1), "^") '�걾�Ÿ�ʽ���걾��^�걾���^SampleID��0�����棬1�����
130                If UBound(aNos) = 0 Then
                      'û�б걾����򰴳���걾����
132                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): iType = 0: strBarcode = ""
                   Else
134                    If gblnEmerge = True Then
136                         iType = Val(aNos(1))
                       Else
                            '2011-02-21 �������뱣��ʱ,"�Ƿ����ּ���"����ΪFalseʱ,������걾����.
138                         iType = 0
                       End If
140                    strDate = Trim(aItem(0)): strSampleID = IIf(aQC(0) = "1", aNos(0), Val(aNos(0))): strBarcode = ""
142                    If UBound(aNos) > 1 Then
144                        strBarcode = Trim(aNos(2))
                       End If
                   End If
                  '��������걾���ɹ��򣨰�ʱ�䣩
146                strStartDate = GetDateTime(mMakeNoRule, 1, strDate)
148                strEndDate = GetDateTime(mMakeNoRule, 2, strDate)
                
150                strName = Trim(aItem(2)): strSample = Trim(aItem(3))
152                 If Trim(strName) = "" Then strName = gstrUserName
                  '�ж��Ƿ������걾
154                If Len(Trim(strBarcode)) = 0 Then
                      '���걾�Ų�
156                    blnQryWithSampleNO = True
                   Else
                      '�������ѯ
158                    gstrSQL = "Select  a.id,a.ҽ��ID,a.��������,a.�����,a.�걾����, a.������,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As ��������A From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                          " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                          " And a.����ʱ�� Between [1] And [2]" & _
                          " And a.����ID=[3] And a.��������=[6]"
160                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "��ѯ�걾��¼", CDate(strStartDate), _
                          CDate(strEndDate), lngDeviceID, strSampleID, iType, strBarcode)
162                    If Not rsTmp.EOF Then
164                        blnQryWithSampleNO = False
                       Else
                          '�����Ƿ����б걾
166                        gstrSQL = "Select a.id,a.ҽ��ID,a.��������,a.�����,a.�걾����, a.������,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As ��������A From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                          " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                          " And a.����ʱ�� Between [1] And [2]" & _
                          " And a.����ID=[3] And a.�걾���=[4] " & IIf(gblnEmerge, " And Nvl(a.�걾���,0)=[5]", "")
168                        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "��ѯ�걾��¼", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                              CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, strSampleID, iType)
170                        If rsTmp.EOF = True Then
                              '�����������ɱ걾
172                            Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
174                            blnQryWithSampleNO = True
                           Else
176                            If Val(Nvl(rsTmp("ҽ��id"))) = 0 Then
                                  '�걾Ϊ����ʱҲ����
178                                Call CreateSample(lngDeviceID, strBarcode, strSampleID, CDate(strDate), iType)
180                                blnQryWithSampleNO = True
                               End If
                           End If
                       End If
                   End If
182                If blnQryWithSampleNO Then
184                    gstrSQL = "Select a.id,a.ҽ��ID,a.��������,a.�����,a.�걾����, a.������,Decode(A.�Ա�,Null,0,'��',1,'Ů',2,0) As �Ա�A,to_char(c.��������,'yyyy-mm-dd') As ��������A From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                          " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                          " And a.����ʱ�� Between [1] And [2]" & _
                          " And a.����ID=[3] And a.�걾���=[4] " & IIf(gblnEmerge, " And Nvl(a.�걾���,0)=[5]", "")
                        '--- 2012-11-21 ΢����걾������ʱ��������
186                     If intMicrobe = 1 Then
                            If intMicrobeDay = 0 Then
                                gstrSQL = Replace(gstrSQL, "And a.����ʱ�� Between [1] And [2]", " ")
                            End If
                            strStartDate = Format(gobjDatabase.Currentdate - intMicrobeDay, "yyyy-mm-dd 00:00:00")
                        End If
                        Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "��ѯ�걾��¼", CDate(strStartDate), _
                        CDate(strEndDate), lngDeviceID, strSampleID, iType)
                   End If
188                bln�����걾 = False
190                If rsTmp.EOF Then
                      '�����걾������ʱ�걾��¼
192                    bln�����걾 = True
194                    strSex = 0
196                    strBirth = ""
                    
198                    lngID = gobjDatabase.GetNextId("����걾��¼")
                       
200                    gstrSQL = "ZL_����걾��¼_INSERT(" & lngID & ",NULL,'" & _
                          strSampleID & "',NULL,NULL," & lngDeviceID & ",'" & strName & "'," & _
                          "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                          "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSample & "'," & _
                          "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strName & "','0'," & lngExeDeptID & "," & iType & "," & intMicrobe & ")"

202                    gobjDatabase.ExecuteProcedure gstrSQL, "���������ʱ��¼"
204                    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ���������걾" & vbNewLine & "�걾ID=" & lngID & ",����=" & strDate & ",�걾��=" & strSampleID & ",����=" & strBarcode & ".����id=" & lngDeviceID & ",����Ա=" & strName & ",΢����걾=" & intMicrobe
                   Else
206                    If Val("" & rsTmp!ҽ��ID) = 0 Then bln�����걾 = True
208                    strSex = Nvl(rsTmp("�Ա�A"), 0)
210                    strBirth = Nvl(rsTmp("��������A"))
212                    If intMicrobe = 0 Then
214                        strSample = Nvl(rsTmp("�걾����"))
                       End If
216                    lngID = rsTmp("ID")
218                    blnAuditing = Not IsNull(rsTmp("������"))
220                    If blnAuditing = False Then
222                        blnAuditing = Not IsNull(rsTmp("�����"))
                       End If
224                    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �ҵ��걾" & vbNewLine & "�걾ID=" & lngID & ",����=" & strDate & ",�걾��=" & strSampleID & ",����=" & strBarcode & ",����id=" & lngDeviceID & ",����Ա=" & strName & ",΢����걾=" & intMicrobe
                   End If
                

226                If Not blnAuditing Then
228                    If InStr(strIDs, "," & lngID) = 0 Then strIDs = strIDs & "," & lngID
                      '���������Ŀ
230                    strItemRecords = ""
232                    strδ֪�� = ""
234                    str���鱸ע = ""
236                    strItemsInfo = ""
238                    For j = 5 To UBound(aItem) Step 2
                          '����ͨ�����޸���Ӧ��Ŀ�����δ�ҵ�����ֱ�����ӣ�����ͨ�����Ҳ�����Ŀ���ݲ�����
                          '����ͨ��������Ŀ
                            strItemAuditing = ""
                            If InStr(aItem(j), "^") > 0 Then
                                strItemCode = Split(aItem(j), "^")(0)
                                intItemAuditing = Val(Split(aItem(j), "^")(1))
                                If UBound(Split(aItem(j), "^")) = 2 Then
                                    strItemAuditing = Split(aItem(j), "^")(2)
                                End If
                            Else
                                strItemCode = aItem(j)
                                intItemAuditing = 0
                            End If
240                         lngItemID = GetItemID(strItemCode, vItems, iDec, strItemName)
242                         If lngItemID > 0 Then
                            
244                            gstrSQL = "select ��Ŀid from ����������Ŀ where ����id = [1] and ��������Ŀ = -1 and ��Ŀid = [2] "
246                            Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "������", lngMainID, lngItemID)
248                            If rsTmp.EOF = False Then
                                  '��������������Ŀʱ�Ĵ���
250                                If strBarcode <> "" Then
                                      '������ʱ�Ĵ��� ,����ͨ����
252                                    gstrSQL = "Select d.��Ŀid" & vbNewLine & _
                                              "From ����ҽ����¼ A, ����ҽ������ B, ���鱨����Ŀ C, ����������Ŀ D" & vbNewLine & _
                                              "Where A.ID = B.ҽ��id And B.�������� = [2] And A.������Ŀid = C.������Ŀid And C.������Ŀid = D.��Ŀid" & vbNewLine & _
                                              "      And D.����id = [1] And D.ͨ������ =[3] And D.��������Ŀ = -1"
254                                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "����������", lngMainID, strBarcode, CStr(strItemCode))
256                                    If rsTmp.EOF = False Then
258                                         strItemRecords = strItemRecords & "|" & Nvl(rsTmp("��ĿID")) & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
260                                         strItemsInfo = strItemsInfo & "," & strItemName & "(��������Ŀ)" & "=" & aItem(j + 1)
                                       Else
262                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
264                                        strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                                       End If
                                   Else
                                      'û������ʱ�Ĵ���
266                                    gstrSQL = "Select B.��Ŀid" & vbNewLine & _
                                              " From ������ͨ��� A, ����������Ŀ B" & vbNewLine & _
                                              " Where A.������Ŀid = B.��Ŀid And B.����id = [1] And B.��������Ŀ = -1 And B.ͨ������=[3]  And A.����걾id = [2] "
268                                    Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "����������", lngMainID, lngID, CStr(strItemCode))
270                                    If rsTmp.EOF = False And rsTmp.RecordCount = 1 Then
272                                        strItemRecords = strItemRecords & "|" & Nvl(rsTmp("��ĿID")) & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
274                                        strItemsInfo = strItemsInfo & "," & strItemName & "(��������Ŀ)" & "=" & aItem(j + 1)
                                       Else
276                                        strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
278                                        strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                                       End If
                                       
                                   End If
                               Else
                                  '����û����������Ŀʱ�Ĵ���
280                                strItemRecords = strItemRecords & "|" & lngItemID & "^" & aItem(j + 1) & "<Split>" & intItemAuditing
282                                strItemsInfo = strItemsInfo & "," & strItemName & "=" & aItem(j + 1)
                               End If
                           Else
   
                            
284                            If strItemCode = "���鱸ע" Then
                            
286                                str���鱸ע = str���鱸ע & IIf(str���鱸ע <> "", vbNewLine, "") & aItem(j + 1)
288                                If InStr(UCase(str���鱸ע), "VBNEWLINE") > 0 Then
290                                    str���鱸ע = Replace(str���鱸ע, "vbnewline", vbNewLine, , , vbTextCompare)
                                   End If
                               Else
292                                If strδ֪�� = "" Then strδ֪�� = "�걾��     ��Ŀ��ʶ     ��Ŀֵ" & vbNewLine
294                                strδ֪�� = strδ֪�� & strSampleID & Space(30 - Len(strSampleID)) & _
                                  strItemCode & Space(30 - Len(strItemCode)) & _
                                  aItem(j + 1) & vbNewLine
296                               strItemsInfo = strItemsInfo & "," & strItemCode & "(δ����)=" & aItem(j + 1)
                               End If
    '                            mcnAccess.Execute strSql
                           End If
                           If strItemAuditing <> "" Then
                                strItemRecords = strItemRecords & "^" & strItemAuditing
                           End If
                       Next
298                    If strδ֪�� <> "" Then Call WriteLog("SaveToDataBase", LOG_δ֪��, 0, strδ֪��)
300                    strUnknown = strδ֪��
                       
302                    If Len(strItemRecords) > 0 Then
304                        strItemRecords = Mid(strItemRecords, 2)
                           strItemRecords = ItemLocalSort(strItemRecords, vItems) '2011-12-07 ���������޸ģ�1/5 - ָ������
                            
306                        gstrSQL = "ZL_������ͨ���_BATCHUPDATE(" & lngID & "," & _
                              lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                              IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                              strItemRecords & "'," & intMicrobe & ")"
308                        gobjDatabase.ExecuteProcedure gstrSQL, "����������"
                            
                            gstrSQL = "Zl_���¼�����_Cale(" & lngID & ")" '2011-12-07 ���������޸ģ�2/5 - �ӵ����¼������
                            gobjDatabase.ExecuteProcedure gstrSQL, "����������"
                            
310                        strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ���������" & vbNewLine & Mid$(strItemsInfo, 2)
312                        If str���鱸ע <> "" Then
314                            str���鱸ע = Replace(str���鱸ע, "'", "")
316                            gstrSQL = "Zl_����걾��¼_���±�ע(" & lngID & ",'" & str���鱸ע & "',1)"
318                            gobjDatabase.ExecuteProcedure gstrSQL, "��������ע"
320                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ������鱸ע" & vbNewLine & str���鱸ע
                           End If

                        
                          '����Ϊ�ʿ�
322                        If aQC(0) = 1 Then
                               Dim date��ǰ���� As Date, lngQCID As Long, str�걾�� As String
                               Dim var�걾�� As Variant, iCoutn As Integer
324                            lngQCID = 0
326                            date��ǰ���� = gobjDatabase.Currentdate
328                            gstrSQL = "Select ID,�걾�� From �����ʿ�Ʒ Where [2] between ��ʼ���� and �������� And ����id = [1] "
330                            Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, gstrSysName, lngDeviceID, date��ǰ����)
                            
332                            Do Until rsTmp.EOF Or lngQCID <> 0
334                                str�걾�� = "" & rsTmp.Fields("�걾��")
336                                If InStr(str�걾��, ",") > 0 Then
338                                    var�걾�� = Split(str�걾��, ",")
340                                    For iCoutn = 0 To UBound(var�걾��)
342                                        If var�걾��(iCoutn) Like "*-*" Then
344                                            If strSampleID >= Val(Split(var�걾��(iCoutn), "-")(0)) And strSampleID <= Val(Split(var�걾��(iCoutn), "-")(1)) Then
346                                                lngQCID = rsTmp.Fields("ID")
                                               End If
                                           Else
348                                            If var�걾��(iCoutn) = strSampleID Then
350                                                lngQCID = rsTmp.Fields("ID")
                                               End If
                                           End If
                                       Next
352                                ElseIf str�걾�� Like "*-*" Then
354                                    If strSampleID >= Val(Split(str�걾��, "-")(0)) And strSampleID <= Val(Split(str�걾��, "-")(1)) Then
356                                        lngQCID = rsTmp.Fields("ID")
                                       End If
                                   Else
358                                    If strSampleID = str�걾�� Then
360                                        lngQCID = rsTmp.Fields("ID")
                                       End If
                                   End If
                                
362                                rsTmp.MoveNext
                               Loop
                            
364                            If lngQCID > 0 Then
366                                gstrSQL = "ZL_�����ʿؼ�¼_EDIT(1," & lngID & "," & lngQCID & ")"
368                                gobjDatabase.ExecuteProcedure gstrSQL, "����Ϊ�ʿ�Ʒ"
370                                strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����Ϊ�ʿ�Ʒ:�ɹ�!"
                                      'Ҫ�Զ�����ʧ�ؼ���
372                                   If intAutoQCCalc = 1 Then
374                                     If strQCList(UBound(strQCList)) <> "" Then ReDim Preserve strQCList(UBound(strQCList) + 1)
376                                       strQCList(UBound(strQCList)) = Format(CDate(strDate), "yyyy-MM-dd") & "," & CStr(lngQCID)
                                      End If
                                End If
378                        ElseIf strAutoCheckMan <> "" Then
                              '�Զ����
380                            If InStr(1, gstrPrivs, "��˱걾") > 0 And bln�����걾 = False Then
                                   
382                                If strAutoCheck(UBound(strAutoCheck)) <> "" Then ReDim Preserve strAutoCheck(UBound(strAutoCheck) + 1)
384                                If int2Verify = 1 Then
386                                    strAutoCheck(UBound(strAutoCheck)) = lngID & "|Zl_����걾��¼_���󱨸�(" & lngID & ",1,'" & gstrUserName & "')"
                                   Else
388                                    strAutoCheck(UBound(strAutoCheck)) = lngID & "|ZL_����걾��¼_�������(" & lngID & ",'" & strAutoCheckMan & "','" & gstrUserCode & "','" & gstrUserName & "')"
                                   End If
                               End If
                           End If
390                    ElseIf intMicrobe = 1 Then
                              '����΢����ֻ��ϸ�����ص����
392                        gstrSQL = "ZL_������ͨ���_BATCHUPDATE(" & lngID & "," & _
                              lngDeviceID & ",'" & strSample & "'," & strSex & "," & _
                              IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                              "0^^^" & "'," & intMicrobe & ")"
394                        gobjDatabase.ExecuteProcedure gstrSQL, "����������"
396                        strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ���������:ֻ��ϸ������,����ɹ�!"
                        End If
                        
                        
398                     If UBound(aTmp) > 0 Then
400                        If Trim(aTmp(1)) <> "" Then
                              '����ͼ������
                               'Call WriteLog("SaveImg", LOG_ͨѶ��־, 0, "��ʼʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
402                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ��ʼ����ͼ������" & vbNewLine & aTmp(1)
404                            Call SaveImg(lngDeviceID, lngID, aTmp(1))
                               'Call WriteLog("SaveImg", LOG_ͨѶ��־, 0, "����ʱ��:" & Format(Now(), "yyyy-MM-dd HH:mm:ss"))
406                            strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ��������ͼ������"
                           End If
                        Else
                            strLog = strLog & vbNewLine & "û��ͼ������"
                        End If 'End Ubound(atmp)>0
                        
                   Else
                        strLog = strLog & vbNewLine & "�걾����ˣ�������!"
                   End If 'blnAuditing
        
               Else
                  strLog = strLog & vbNewLine & "����ĸ�ʽ����ȷ����Ҫ�����ĸ�Ԫ��"
               End If 'If UBound(aItem) >= 4 Then
           Next
       End If
   
      '�����ʿ�

408    SaveToDataBase = True

410    For i = LBound(strQCList) To UBound(strQCList)
412        If InStr(strQCList(i), ",") > 0 Then
414            Call AutoQCCompute(lngDeviceID, CDate(Split(strQCList(i), ",")(0)), Split(strQCList(i), ",")(1), strAutoCaleInfo)
           End If
       Next
    
      '�Զ����
       Dim strInfo As String

416    For i = LBound(strAutoCheck) To UBound(strAutoCheck)
418        If InStr(strAutoCheck(i), "|") > 0 Then
420            lngID = Val(Split(strAutoCheck(i), "|")(0))
422             If CheckSample(lngID) Then
                '����Ƿ�������
424            If VerifyAuditingRule(lngID, strInfo) <> 1 Then

426                strSQL = Split(strAutoCheck(i), "|")(1)
428                gobjDatabase.ExecuteProcedure strSQL, "�Զ����" & lngID
430             strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �Զ���� " & lngID
               End If
                End If
           End If
       Next
432    strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �������"
434    strlogs = strLog
       Exit Function
DBError:
436   Call WriteLog("SaveToDataBase", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
438 strLog = strLog & vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ���" & CStr(Erl()) & "�г��ִ���" & Err.Description
End Function



Private Sub AutoQCCompute(ByVal lngDeviceID As Long, ByVal date���� As Date, ByVal str�ʿ�Ʒ As String, ByRef strRetuInfo As String)

        '�Զ������ʿر걾
        ' date���� :�ʿؼ�������
        ' str�ʿ�Ʒ :�ʿ�Ʒ
        Dim rsTemp As adodb.Recordset, rsTmp As adodb.Recordset, strReturn As String
        On Error GoTo errH
100     gstrSQL = "Select Distinct B.��Ŀid, C.����, C.������, C.Ӣ����" & vbNewLine & _
                  " From �����ʿ�Ʒ A, �����ʿ�Ʒ��Ŀ B, ����������Ŀ C" & vbNewLine & _
                  " Where A.ID = B.�ʿ�Ʒid And B.��Ŀid = C.ID And A.����id = [1] "
        
102     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "LisComm�Զ�����", lngDeviceID)
104     Do Until rsTmp.EOF
            '����һ��ʱ��
106             gstrSQL = "Select Zl_�����ʿؼ�¼_Compute(" & lngDeviceID & ", " & rsTmp("��ĿID") & ", To_Date('" & Format(date����, "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & str�ʿ�Ʒ & "') From Dual"
108             Set rsTemp = gobjDatabase.OpenSqlRecord(gstrSQL, "LisComm�Զ�����")

110             If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")  ������̵��ô���" & vbCrLf
112             If InStr(rsTemp.Fields(0).Value, "����ʧ�أ�") > 0 Then
114                 strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf

116             ElseIf InStr(rsTemp.Fields(0).Value, "������ɣ�") <= 0 Then
118                 If InStr(rsTemp.Fields(0).Value, "������δ���־����ʧ�أ�") <= 0 Then
120                 strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                    End If
                End If
122         rsTmp.MoveNext
        Loop
124     If Trim(strReturn) <> "" Then
126        strRetuInfo = strReturn
        End If
        Exit Sub
errH:
128    WriteLog "AutoQCCompute", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description
End Sub


Public Function VerifyAuditingRule(lngSampleID As Long, Optional strErrMessage As String) As Integer
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '����                       ���ʱ������˹���
        '����                       lngSampleID �걾ID; strErrMessage ����1ʱ�Ĵ�����ʾ��
        '����                       0 ���� 1 �н��������ʾֵ
        '
        '�����־ 3-����2-����1-������4-�쳣��5-������6-����
        '
        Dim strSQL As String
        Dim rsTmp As New adodb.Recordset
        Dim int����id As Integer '
        On Error GoTo errH
        '��������ʾֵ�Ľ��
100     strSQL = " select �����־ from ����걾��¼ a , ������ͨ��� b " & _
                 " Where a.ID = b.����걾id and a.id = [1] and (b.�����־ = 5 Or b.�����־ = 6)"
102     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
104     If rsTmp.EOF = False Then
106         VerifyAuditingRule = 1: strErrMessage = "  ���������ʾֵ��"
        End If
       '��������ʾֵ�Ľ��
108     strSQL = " select �����־ from ����걾��¼ a , ������ͨ��� b " & _
                 " Where a.ID = b.����걾id and a.id = [1] and (b.�����־ = 5 Or b.�����־ = 6)"
110     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
112     If rsTmp.EOF = False Then
114         VerifyAuditingRule = 1: strErrMessage = "  ���������ʾֵ��"
        End If
        '-- �����޸ģ�������ȫ��Ϊ�գ�����ʾ��
116     strSQL = "Select Count(B.ID) - Sum(Decode(Trim(b.������), Null, 1, 0)) As ���" & vbNewLine & _
                 "From ����걾��¼ a , ������ͨ��� B Where a.id = b.����걾ID and  a.id = [1] and nvl(a.΢����걾,0) = 0 "
118     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName, lngSampleID)
120     Do Until rsTmp.EOF
122         If Nvl(rsTmp("���")) <> "" Then
124             If Val("" & rsTmp!���) <= 0 Then
126                VerifyAuditingRule = 1: strErrMessage = "  ���ȫ��Ϊ�գ�"
                End If
            End If
128         rsTmp.MoveNext
        Loop
    
    
130     int����id = gobjDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
    
132     If VerifyAuditingRule <> 1 And strErrMessage = "" Then
134         strSQL = "Select Zl_������˹���_Check(" & lngSampleID & "," & int����id & ") as ��˽�� From Dual"
136         Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, gstrSysName)
138         If rsTmp.RecordCount <= 0 Then
140             VerifyAuditingRule = 1
142             strErrMessage = "  ������̵��ô���! "
                Exit Function
            End If
144         strErrMessage = "" & rsTmp.Fields(0).Value
146         If strErrMessage <> "" Then VerifyAuditingRule = 1
        End If
        Exit Function
errH:
148     WriteLog "VerifyAuditingRule", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description
End Function

Private Function CheckSample(ByVal lngID As Long) As Boolean
        '���ǰ���
        Dim rsTmp As adodb.Recordset, strSQL As String
        On Error GoTo errH
        '11210 Ȩ�ޡ�δ�շ���ˡ�������˵�������ʱ��δ��Ч��
100     If InStr(gstrPrivs, "δ�շ����") <= 0 Then
102         If CheckChargeState(lngID, False) = False Then
104             WriteLog "InDataBase", LOG_������־, lngID, "����δ�շѣ����ܽ�����ˣ�"
                Exit Function
            End If
        End If
    
        '21137 �ѹ鵵���治�����
106     strSQL = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
108     Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "CheckSample", lngID)
110     If rsTmp.EOF = False Then
112         WriteLog "InDataBase", LOG_������־, lngID, "���˱���סԺ�Ĳ������ύ��飬���ܽ�����ˣ�"
            Exit Function
        End If
114     If CheckExesState(lngID) = False Then
116         WriteLog "InDataBase", LOG_������־, lngID, "��ǰסԺ���˻��л��۵�δ��ˣ����ѳ�Ժ��Ԥ��Ժ��"
            Exit Function
        End If
    
        '��������
118     Call CheckPatientInfo(lngID)
120     CheckSample = True
        Exit Function
errH:
122     Call WriteLog("CheckSample", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function


Private Function CheckChargeState(ByVal lngKey As Long, Optional ByVal blnOrder As Boolean = True, Optional ByVal DataMoved As Boolean = False) As Boolean
        '�����շ�״̬
        Dim strSQL As String
        Dim rs As New adodb.Recordset
        Dim strSQLbak As String
        Dim intPatientType As Integer               '������Դ
        On Error GoTo errH
    
100     CheckChargeState = False
    
102     strSQL = "select ������Դ from ����ҽ����¼ where id = [1]"
104     Set rs = gobjDatabase.OpenSqlRecord(strSQL, "��������", lngKey)
106     If rs.EOF = True Then Exit Function
108     intPatientType = rs("������Դ")
    
110     If blnOrder Then
112         strSQL = _
                "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                      "from סԺ���ü�¼ A, " & _
                      "( " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id))  " & _
                           "Union " & _
                           "select No from ����ҽ������ where ҽ��id IN (SELECT ID FROM ����ҽ����¼ WHERE [1] In (ID,���id)) " & _
                      ") B " & _
                    "Where A.NO = B.NO "
114         If intPatientType <> 2 Then
116             strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            End If
        Else
118         strSQL = _
                "select NVL(A.��¼״̬,0) As ��¼״̬ " & _
                      "from סԺ���ü�¼ A, " & _
                      "( " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                           "Union " & _
                           "select No,��¼���� from ����ҽ������ where ҽ��id IN (Select ID From ����ҽ����¼ A,(Select ҽ��id From ����걾��¼ Where ID= [1] Union Select ҽ��id From ������Ŀ�ֲ� Where �걾id= [1]) B where B.ҽ��id In (A.ID,A.���id) and A.������� = 'C' ) " & _
                      ") B " & _
                    "Where A.NO = B.NO and a.��¼���� = b.��¼���� "
120         If intPatientType <> 2 Then
122             strSQL = Replace(strSQL, "סԺ���ü�¼", "������ü�¼")
            End If
        End If
    
124     strSQL = strSQL & " Order by ��¼״̬ "
126     If DataMoved Then
128         strSQL = Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼")
130         strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
132         strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
134         strSQL = Replace(strSQL, "����걾��¼", "H����걾��¼")
136         strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
138         strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
    
140     Set rs = gobjDatabase.OpenSqlRecord(strSQL, "mdlLisWork", lngKey)

142     If rs.BOF Then Exit Function
144     If rs("��¼״̬").Value = 0 Then Exit Function
    
146     CheckChargeState = True
    
        Exit Function
errH:
148     Call WriteLog("CheckChargeState", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function

Private Function CheckExesState(lngKey As Long) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '����:      ���סԺ���˳�Ժ���Ƿ��л��۵���Ҫ�������
        '����       �걾ID
        '����       �л��۵�δ��� = Fasle û���� = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim rsTmp As New adodb.Recordset
        On Error GoTo errH
100     CheckExesState = True
    
        '81��ϵͳ����Чʱ�����
102     If gobjDatabase.GetPara(81, 100) <> 1 Then Exit Function
        
        '��ǰ�����Ƿ��ѳ�Ժ��Ԥ��Ժ
104     gstrSQL = "select d.no" & vbNewLine & _
                "from (select distinct d.ҽ��id" & vbNewLine & _
                "       from ����걾��¼ a, ������Ϣ b, ������ҳ c, ������Ŀ�ֲ� d" & vbNewLine & _
                "       where a.����id = b.����id and a.����id = c.����id and a.��ҳid = c.��ҳid and" & vbNewLine & _
                "             a.id = [1] and a.������Դ = 2 and (b.��Ժʱ�� is not null or c.״̬ = 3) and" & vbNewLine & _
                "             a.id = d.�걾id) a, ����ҽ����¼ b, ����ҽ������ c, סԺ���ü�¼ d" & vbNewLine & _
                "where a.ҽ��id in (b.���id, b.id) and b.id = c.ҽ��id and c.��¼���� = d.��¼���� and" & vbNewLine & _
                "      c.no = d.no and d.��¼���� = 2 and d.��¼״̬ = 0 "
106     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "���鼼ʦ����վ-����״̬���", lngKey)
    
108     CheckExesState = rsTmp.EOF
        Exit Function
errH:
110     Call WriteLog("CheckExesState", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function

Private Function CheckPatientInfo(lngSampleID As Long) As Boolean
        Dim rsTmp As New adodb.Recordset
        Dim int��ʾ���� As Integer '1-��ʾ������2-����ʾ������3-������

        On Error GoTo errH
    
100     gstrSQL = "Select A.������Դ,A.����id, A.�Ա� As �Ա�1, B.�Ա� As �Ա�2, A.���� As ����1, B.���� As ����2, A.���� As ����1, B.���� As ����2,nvl(a.Ӥ��,0) as Ӥ�� " & vbNewLine & _
                            "From ����걾��¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = B.����id And A.ID = [1]"
102     Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "CheCkPatientInfo", lngSampleID)
    
        '��Ӥ��ʱ�����жԱ�
104     If rsTmp("Ӥ��") > 0 Then
            Exit Function
        End If
    
    
106     If Nvl(rsTmp("����1")) <> Nvl(rsTmp("����2")) Or Nvl(rsTmp("�Ա�1")) <> Nvl(rsTmp("�Ա�2")) Or _
            Nvl(rsTmp("����1")) <> Nvl(rsTmp("����2")) Then
        
108         int��ʾ���� = 1
        
110         If rsTmp("������Դ") = 4 Then
112             int��ʾ���� = Val(gobjDatabase.GetPara("��첡����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
114         ElseIf rsTmp("������Դ") = 3 Then
116             int��ʾ���� = Val(gobjDatabase.GetPara("Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
118         ElseIf rsTmp("������Դ") = 2 Then
120             int��ʾ���� = Val(gobjDatabase.GetPara("סԺ������Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
122         ElseIf rsTmp("������Դ") = 1 Then
124             int��ʾ���� = Val(gobjDatabase.GetPara("���ﲡ����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
            End If
        
126         If int��ʾ���� = 1 Then
128             WriteLog "InDataBase", LOG_������־, lngSampleID, "���ּ�����Ϣ�еĲ�����Ϣ�Ͳ�����Ϣ�в�����Ϣ��һ��!"
130         ElseIf int��ʾ���� = 2 Then
132             gstrSQL = "zl_����걾��¼_Update(" & lngSampleID & ",'" & Nvl(rsTmp("����2")) & "','" & Nvl(rsTmp("�Ա�2")) & _
                                             "','" & Nvl(rsTmp("����2")) & "')"
134             gobjDatabase.ExecuteProcedure gstrSQL, "CheckPatientInfo"
            End If
136         CheckPatientInfo = True
            Exit Function
        End If
138     CheckPatientInfo = False
    
        Exit Function
errH:
140     Call WriteLog("CheckPatientInfo", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function


Public Function TestFTP(ByVal strUser As String, ByVal strPassWord As String, _
                            ByVal strDevAdress As String, ByVal strFtpPath As String) As String
                            
    Dim FtpNet As New clsFtp, strPath As String, strTmpPath As String           'FTP��
    Dim lngFileNo As Long
    strPath = Format(Now, "yyyymmddHHMMSS")
    strTmpPath = IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "temp.txt"
    lngFileNo = FreeFile
    Open strTmpPath For Output As lngFileNo
    Close lngFileNo
    If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) > 0 Then
        If FtpNet.FuncFtpMkDir(strFtpPath, "FTP����" & strPath) > 0 Then
            TestFTP = "��FTP�ϲ��ܴ���Ŀ¼��"
        Else
            If FtpNet.FuncUploadFile(strFtpPath, strTmpPath, "temp.txt") > 0 Then
                TestFTP = "�ϴ��ļ�ʧ��"
            Else
                FtpNet.FuncFtpDisConnect '�ȶϿ�����ɾ������Ȼɾ����
                If FtpNet.FuncFtpConnect(strDevAdress, strUser, strPassWord) <= 0 Then
                     TestFTP = "FTP�������ӣ�"
                ElseIf FtpNet.FuncFtpDelDir(strFtpPath, "FTP����" & strPath) > 0 Then
                    TestFTP = "��FTP�ϲ���ɾ��Ŀ¼"
                Else
                    TestFTP = ""
                End If
            End If
        End If
    Else
        TestFTP = "��������FTP��"
    End If
    FtpNet.FuncFtpDisConnect
    Set FtpNet = Nothing
    Kill strTmpPath
End Function

Private Function DownFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                          ByVal strFtpFile As String, ByVal strFile As String) As String
        '��FTP�����������ļ���
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpFile :FTP�ϵ��ļ���
        'strFile    :�����ļ�ȫ·����
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
        Dim objFtp As New clsFtp, lngReturn As Long, strFtpFileName As String, strLocaFile As String
        Dim strFTPDir As String
        On Error GoTo errH
100     If strFtpFile = "" Then
102         DownFile = "��ָ��Ҫ���ص��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
104     strFtpFileName = Split(strFtpFile, "/")(UBound(Split(strFtpFile, "/")))
106     strFTPDir = Replace(strFtpFile, "/" & strFtpFileName, "")
108     strLocaFile = strFile
110     If strLocaFile = "" Then
112         DownFile = "��ָ�����ص��ļ����浽�δ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
114     If Dir(strLocaFile) <> "" Then
116         DownFile = "Ҫ���ص��ļ��Ѵ��ڣ�"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
    
118     If strServer = "" Then
120         DownFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
122     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
124     If lngReturn = 0 Then
126         DownFile = "�������ӷ�������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
128     lngReturn = objFtp.FuncChangeDir(strFTPDir)
130     If lngReturn <> 0 Then
132         DownFile = "���ܽ���ָ����Ŀ¼��������Ȩ�޲������������޴�Ŀ¼��"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
134     lngReturn = objFtp.FuncDownloadFile(strFTPDir, strLocaFile, strFtpFileName)
136     If lngReturn <> 0 Then
138         DownFile = "����ʧ�ܣ�������Ȩ�޲������������޴��ļ���"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        objFtp.FuncFtpDisConnect
140     Set objFtp = Nothing
        Exit Function
errH:
142     DownFile = CStr(Erl()) & "�У�" & Err.Description
End Function

Private Function UploadFile(ByVal strUser As String, ByVal strPass As String, ByVal strServer As String, _
                            ByVal strFtpPath As String, ByVal strFile As String, Optional strNewFileName As String) As String
        '�������ļ����ϴ��ļ���FTP��������
        'strUser    :�û���
        'strPass    :����
        'strServer  :������
        'strFtpPath :FTP�ϵ�Ŀ¼����Ŀ¼���Զ�������
        'strFile    :�����ļ�ȫ·����
        'strNewFileName: ����FTP�Ϻ���ļ�����Ϊ���򰴱����ļ�������
        '���أ��մ���ʾ�ɹ�������Ϊ������ʾ��
    
        Dim objFtp As New clsFtp, lngReturn As Long, strFileName As String, strLocaFile As String
        On Error GoTo errH
    
    
100     If Left(strFtpPath, 1) = "/" Then strFtpPath = Mid$(strFtpPath, 2)
    
102     If strServer = "" Then
104         UploadFile = "��ָ��FTP������"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
106     strLocaFile = strFile
108     If Dir(strLocaFile) = "" Then
110         UploadFile = "�ļ�" & strLocaFile & "������!"
            objFtp.FuncFtpDisConnect
            Set objFtp = Nothing
            Exit Function
        End If
        If strNewFileName = "" Then
112         strFileName = Split(strLocaFile, "\")(UBound(Split(strLocaFile, "\")))
        Else
            strFileName = strNewFileName
        End If
114     lngReturn = objFtp.FuncFtpConnect(strServer, strUser, strPass)
116     If lngReturn <> 0 Then
            '���Ŀ¼�Ƿ����
118         lngReturn = objFtp.FuncChangeDir(strFtpPath)
120         If lngReturn <> 0 Then
122             lngReturn = objFtp.FuncFtpMkDir("/", strFtpPath)
124             If lngReturn <> 0 Then
126                 UploadFile = "����Ŀ¼ʧ�ܣ�������Ȩ�޲��㣡"
                    objFtp.FuncFtpDisConnect
                    Set objFtp = Nothing
                    Exit Function
                End If
            End If
        
128         lngReturn = objFtp.FuncUploadFile("/" & strFtpPath, strLocaFile, strFileName)
130         If lngReturn <> 0 Then
132             UploadFile = "�ϴ��ļ�ʧ�ܣ�������Ȩ�޲��㣡"
                objFtp.FuncFtpDisConnect
                Set objFtp = Nothing
                Exit Function

            Else
134             UploadFile = ""
            End If
        Else
136         UploadFile = "�������ӷ�������"
        End If
        objFtp.FuncFtpDisConnect
        Set objFtp = Nothing
        Exit Function
errH:
138     UploadFile = CStr(Erl()) & "�У�" & Err.Description
End Function

Private Function ItemLocalSort(ByVal strItems As String, ByRef varItems As Variant) As String
    'ָ������2011-12-07 ���������޸ģ�5/5 - ָ������
    Dim strReturn As String, strTmp As String
    
    Dim i As Integer, varTmp As Variant
    Dim x As Integer
    'ֻ��һ���������������
    If InStr(strItems, "|") <= 0 Then
        ItemLocalSort = strItems
        Exit Function
    End If
    
    varTmp = Split(strItems, "|")
    For i = 0 To UBound(varItems, 2)
        For x = LBound(varTmp) To UBound(varTmp)
            strTmp = Split(varTmp(x), "^")(0)
            If CStr(varItems(1, i)) = strTmp Then
                strReturn = strReturn & "|" & varTmp(x)
                Exit For
            End If
        Next
    Next
    
    If strReturn <> "" Then strReturn = Mid(strReturn, 2)
    ItemLocalSort = strReturn
End Function
