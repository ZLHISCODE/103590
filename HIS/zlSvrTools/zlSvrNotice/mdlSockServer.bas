Attribute VB_Name = "mdlSockServer"
Option Explicit

'**************************
'       OEM����
'
'����    B0AEC9FA
'ҽҵ    D2BDD2B5
'����    CDD0C6D5
'����    D6D0C8ED
'��̩  BDF0BFB5CCA9
'ҽԺ    D2BDD4BA
'**************************

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public gstrProductName As String
Public gstrSysName As String                'ϵͳ����
Public gstrUserName As String               '�û���
Public gstrServer As String                 '��������
Public gstrSQL    As String                 'ͨ�õ�SQL������

Public gcnOracle As ADODB.Connection     '�������ݿ�����

Public Sub Main()
    Dim objLogin As Object
    'Ϊʵ��XP�������ʾ����ǰ����ִ�иú���
    
    If App.PrevInstance Then
        MsgBox " �Զ����ѷ����Ѿ������� ", vbOKOnly, "�Զ�����"
        Exit Sub
    End If
    On Error Resume Next
    If objLogin Is Nothing Then
        Set objLogin = CreateObject("ZLLogin.clsLogin")
    End If
    If objLogin Is Nothing Then
        MsgBox "����ZLLogin��������ʧ��,�����ļ��Ƿ���ڲ�����ȷע�ᡣ"
        Exit Sub
    Else
        Set gcnOracle = objLogin.Login(2, CStr(Command()), , True)
        If gcnOracle Is Nothing Then
            Exit Sub
        ElseIf gcnOracle.State <> adStateOpen Then
            Exit Sub
        End If
    End If
    gstrServer = objLogin.ServerName
    gstrUserName = objLogin.InputUser
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "") & "���"
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    frmMain.Show
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub ApplyOEM_Picture(objPicture As Object, ByVal str���� As String)
'��Ը���ͼ��Ӧ��OEM����
    Dim strOEM As String
    Dim blnCorp As Boolean
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        '����״̬��ͼ���OEM����
        If gstrProductName <> "����" Then
            If Right(str����, 1) = "B" Then
                '��ʾ��ƷͼƬ
                blnCorp = False
                str���� = Mid(str����, 1, Len(str����) - 1)
            Else
                '��ʾ��˾�ձ�
                blnCorp = True
            End If
            
            strOEM = GetOEM(gstrProductName, blnCorp)
            If str���� = "Picture" Then
                Set objPicture.Picture = LoadCustomPicture(strOEM)
            ElseIf str���� = "Icon" Then
                Set objPicture.Icon = LoadCustomPicture(strOEM)
            End If
            
            If Err <> 0 Then
                Err.Clear
            End If
        End If
    End If
End Sub

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEMͼƬ���������� ��һ��ָ��˾�ձ꣬��һ���ǲ�Ʒ��ʶ
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Sub ApplyOEM(objStatus As Object)
'���״̬��Ӧ��OEM����
    Dim strOEM As String
    On Error Resume Next
    
    If gstrProductName <> "-" Then
        objStatus.Panels(1).Text = gstrProductName & "���"
        '����״̬��ͼ���OEM����
        If gstrProductName = "����" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(gstrProductName)
            Set objStatus.Panels(1).Picture = LoadCustomPicture(strOEM)
            If Err <> 0 Then
                Err.Clear
                Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
            End If
        End If
        objStatus.Panels(1).ToolTipText = ""
        objStatus.Height = 360
    End If
End Sub

Public Function LoadCustomPicture(strID As String) As StdPicture
'����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
'����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
'����:�����ļ���
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
    
    '���ܣ���������������һ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼����������壬Ϊ�˲�����������¼����ͻ�����Ե�����һ���ؼ�
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '���ܣ�����������ɾ��ͼ��
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hWnd = lngHwnd   '�¼�����������
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    
    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)
    
    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function


