Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrSysName As String
Public gstrUser As String
Public gstrChatURL As String        '�������۵�ַ
Public gstrMyChatUrl As String        '�Ҳ�������۵�ַ
Public gobjMain As Object           '����̨����  ͨ���˶���������ݿ�
 
Public gfrmMain As frmMain

Public grsList  As ADODB.Recordset  '������Ϣ
Public gcolChat As Collection       '��¼�򿪵�����
Public gblnLog  As Boolean              '
Public gblnShow As Boolean          '���ڼ�¼�ȴ������Ƿ��

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum E_Notify_Type   '0-��ʼ��  1-��Ϣ 2-��˸ 3-��ԭ
    E_��ʼ�� = 0
    E_��Ϣ = 1
    E_��˸ = 2
    E_��ԭ = 3
End Enum
'----------------------------------------------------------------------------------------------------
'-----ϵͳ�����������
'----------------------------------------------------------------------------------------------------
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEWHEEL = &H20A          '������
Public Const SW_RESTORE = 9
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOACTIVATE = &H10 '�������
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE  As Long = (-20)
Public Const conCOLOR_BULELIGHT As Long = &HE4B440
Public Const conCOLOR_BULE As Long = &HD48A00
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
'----------------------------------------------------------
'-------��ɫ-����
'---------------------------------------------------------------
Public Const conCOLOR_TITLE_BAR As Long = 16298544 '16298544 rgb(48,178,248); 14392064 'RGB(0, 155, 219)


Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'������ָ������Ļ�����ϵ�λ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'��ô�������Ļ�����е�λ��
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'�ж�ָ���ĵ��Ƿ���ָ���ľ����ڲ�
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long
'����ʹ����ʼ������ǰ��
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter _
    As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'��ȡ����״̬
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'��
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'д
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
'����ֵ:�����ʾ�ɹ������ʾʧ�ܡ�������GetLastError

Private Const CP_UTF8 = 65001
'Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
    
    
Private Const CON_SPLIT As String = ";"
Private mobjFso As New FileSystemObject         '�ļ�����
    
Public Sub InitRsList()
    Set grsList = New ADODB.Recordset
    With grsList
        .Fields.Append "ID", adBigInt
        .Fields.Append "Url", adVarChar, 500
        .Fields.Append "Sys_Code", adVarChar, 20
        .Fields.Append "Main_Code", adVarChar, 20
        .Fields.Append "Main_ID", adVarChar, 36
        .Fields.Append "Subject", adVarChar, 50
        
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub
 
Public Function StringToUTF8Bytes(strInput As String) As Byte()
    Dim bytUTF8Bytes() As Byte
    Dim lngBytesRequired As Long
    
    '�ȼ��������ֽ���
    lngBytesRequired = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), ByVal 0, 0, ByVal 0, ByVal 0)
     
    'Ȼ��ת��
    ReDim bytUTF8Bytes(lngBytesRequired - 1)
    WideCharToMultiByte CP_UTF8, 0, ByVal StrPtr(strInput), Len(strInput), bytUTF8Bytes(0), lngBytesRequired, ByVal 0, ByVal 0
    
    StringToUTF8Bytes = bytUTF8Bytes
End Function

Public Function URLEncode(ByVal strParameter As String, Optional strEncodeType As String = "utf8") As String
          Dim strTemp As String
          Dim strRet As String
          Dim strInput As String
          
          Dim i As Long
          Dim lngValue As Long
          Dim lngLen As Long
          Dim lngMax As Long
          
          Dim bytData() As Byte

10        On Error GoTo ErrH
20        lngLen = 32767
30        Do While Len(strParameter) > 0
40            lngMax = Len(strParameter)
50            If lngMax > lngLen Then
60                strInput = Mid(strParameter, 1, lngLen)
70                strParameter = Mid(strParameter, lngLen + 1, lngMax - lngLen)
80            Else
90                strInput = strParameter
100               strParameter = ""
110           End If
120           strTemp = ""
130           If "UTF8" = UCase(strEncodeType) Then
140               bytData = StringToUTF8Bytes(strInput)
150           Else
160               bytData = StrConv(strInput, vbFromUnicode)
170           End If
              
180           For i = 0 To UBound(bytData)
190               lngValue = bytData(i)
200               If (lngValue >= 48 And lngValue <= 57) Or _
                      (lngValue >= 65 And lngValue <= 90) Or _
                      (lngValue >= 97 And lngValue <= 122) Or _
                       InStr("$-_.+*'()", Chr(lngValue)) > 0 Then
                       '�����ַ���ת"$-_.+*'()"
210                   strTemp = strTemp & Chr(lngValue)
220               ElseIf lngValue = 32 Then
                      '�ո�
230                   strTemp = strTemp & "+"
240               Else
250                   If lngValue <= 15 Then
260                       strTemp = strTemp & "%0" & UCase(Hex(lngValue))
270                   Else
280                       strTemp = strTemp & "%" & UCase(Hex(lngValue))
290                   End If
300               End If
310           Next
320           strRet = strRet & strTemp
330       Loop
340       URLEncode = strRet
350       Exit Function
ErrH:
360
     WriteLog "��zlChatNotify.mdlPublic.URLEncode�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description & vbNewLine
End Function

Public Function OpenChatRoom(ByVal strUrl As String, ByVal strSubject As String, Optional ByVal strSysCode As String, _
    Optional ByVal strMainCode As String, Optional ByVal dblMainId As Double, Optional ByVal strSender As String, _
    Optional ByVal strReceivers As String, Optional ByRef strMsg As String) As Boolean
    '����:
    'strSubject         -���۱���
    'strSysCode         -ϵͳ����
    'strMainCode        -�������
    'dblMainId          -����ID
    'strSender          -������
    'strReceivers       -������(����������÷ָ���";"�ֿ�)
    'strMsg             -���ش�����Ϣ(���ⵯ��ģ̬��ʾ���������̹���,����ʾ��Ϣ���ظ������̴���)
    '                    ���ظ�ʽ:��ʾ����[,]��ʾ���
          Dim strKey As String
          Dim objChat As frmChat
          
1         On Error GoTo ErrH
2         WriteLog "������OpenChatRoom ��ʼ" & vbNewLine & _
                         "��Σ�url=" & strUrl & vbNewLine & _
                         "���۱���:" & strSubject & vbNewLine & _
                         "ϵͳ����:" & strSysCode & vbNewLine & _
                         "�������:" & strMainCode & vbNewLine & _
                         "����ID:" & dblMainId & vbNewLine & _
                         "������:" & strSender & vbNewLine & _
                         "������:" & strReceivers & vbNewLine
3         strKey = strSysCode & "_" & strMainCode & "_" & dblMainId
4         On Error Resume Next
5         Set objChat = gcolChat(strKey)
6         On Error GoTo ErrH
7         If objChat Is Nothing Then
8             Set objChat = New frmChat
9             gcolChat.Add objChat, strKey
10        End If
11        OpenChatRoom = objChat.OpenChatRoom(strUrl, strSubject, strSysCode, strMainCode, dblMainId, strSender, strReceivers, strMsg)
12        WriteLog "������OpenChatRoom ����"
13        Exit Function
ErrH:
14        strMsg = vbExclamation & "[,]" & "��zlChatNotify.mdlPublic.OpenChatRoom�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description
15        WriteLog strMsg & vbNewLine
End Function

Public Function ReadIni(ByVal strNodeName As String, ByVal strKeyName As String, strFilePath As String) As String
    Dim strBuff As String
    Dim strReadStr As String
    Dim lngPos As Long
    
    On Error GoTo ErrH

    strBuff = VBA.String(255, 0)
    GetPrivateProfileString strNodeName, strKeyName, "", strBuff, 256, strFilePath
    strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
    
    lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)     '�ҵ� ;��λ��(������־)
    If lngPos >= 1 Then
        ReadIni = Trim(Left(strReadStr, lngPos - 1))
    Else
       '���û���ҵ� ��ע�͵ı�־
       ReadIni = strReadStr
    End If
    
    Exit Function
ErrH:
    Err.Clear
    ReadIni = ""
End Function

Public Function WriteIni(ByVal strNodeName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFilePath As String) As Long
    Dim strBuff As String
    Dim strComment As String
    Dim strReadStr As String
    
    Dim lngRet As Long
    Dim lngPos As Long
    On Error GoTo ErrH
   strBuff = String(255, 0)
   lngRet = GetPrivateProfileString(strNodeName, strKeyName, "", strBuff, 256, strFilePath)
   strReadStr = VBA.Replace(strBuff, VBA.Chr(0), "")
   lngPos = InStr(1, strReadStr, CON_SPLIT, vbTextCompare)    '�ҵ� ;��λ��(������־)
   '�����;ȡ������ע��
   If lngPos >= 1 Then
      strComment = Trim(Right(strReadStr, lngRet - lngPos))
      strValue = strValue & strComment
   End If
   
    WriteIni = WritePrivateProfileString(strNodeName, strKeyName, strValue, strFilePath)
        
    Exit Function
ErrH:
    Err.Clear
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��

    Dim strLogPath As String, strLogFile  As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File
    
    '�Ƿ�����־
    If Not gblnLog Then Exit Sub
     
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\" & App.EXEName & "*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    
    '3���ռ��Ƿ��㹻
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '�ռ䲻�㣬��д��־,����һ�������ļ�
        If Not mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\�ռ䲻��.txt", True)
        Exit Sub
    Else
        '��������ļ�
        If mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.DeleteFile(strLogPath & "\�ռ䲻��.txt", True)
    End If
    '4��д����־��
    strLogFile = strLogPath & "\" & App.EXEName & Format(Now, "yyyyMMdd") & ".log"
    Call SaveLog(strLogFile, strLogTxt)

End Sub

Private Sub SaveLog(ByVal strFilename As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFilename) Then Call mobjFso.CreateTextFile(strFilename)
        Set objStream = mobjFso.OpenTextFile(strFilename, ForAppending)
        If strDate = "" Then strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        objStream.WriteLine (strDate & Chr(&H9) & strInput)
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '��ȡʣ��ռ�
    Dim strDriv As String, Drv As Drive
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Sub SetFormTranslucency(hWnd As Long, crKey As Long, bAlpha As Byte, dwFlags As Long) 'ʵ�ְ�͸������
'����:���ô���͸����
'hwnd,  ���ھ��
'crKey:ָ����Ҫ͸���ı�����ɫֵ������RGB()��
'bAlpha:����͸���ȣ�0��ʾ��ȫ͸����255��ʾ��͸��
'dwFlags: ͸����ʽdwFlags������ȡ����ֵ��
'       LWA_ALPHA=&H2ʱ��crKey������Ч��bAlpha������Ч��
'       LWA_COLORKEY=&H1�������е�������ɫΪcrKey�ĵط�����Ϊ͸����bAlpha������Ч���䳣��ֵΪ1��
'       LWA_ALPHA | LWA_COLORKEY��crKey�ĵط�����Ϊȫ͸�����������ط�����bAlpha����ȷ��͸���ȡ�
   Dim lngRet As Long
   
    lngRet = GetWindowLong(hWnd, GWL_EXSTYLE)
    lngRet = lngRet Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, lngRet
    SetLayeredWindowAttributes hWnd, crKey, bAlpha, dwFlags
 End Sub

Public Sub SetWindowsInTaskBar(ByVal lnghwnd As Long, ByVal blnShow As Boolean)
'���ܣ����ô����Ƿ�������������ʾ
    Dim lngStyle As Long
    
    lngStyle = GetWindowLong(lnghwnd, GWL_EXSTYLE)
    If blnShow Then
        lngStyle = lngStyle Or &H40000
    Else
        lngStyle = lngStyle And Not &H40000
    End If
    Call SetWindowLong(lnghwnd, GWL_EXSTYLE, lngStyle)
End Sub

