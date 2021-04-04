Attribute VB_Name = "mdlCommon"
Option Explicit
'**************************
'       OEM����
'
'ҽҵ  D2BDD2B5
'**************************
Public gobjFSO As New FileSystemObject
Public gcnOracle As ADODB.Connection
Public gstrSysName As String
Public gstrDBUser As String '�û���
Public gstrPrivs As String

'������־������ر���
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mstrRecentSQL As String  '���ִ�е�SQL���

'SQLLog����
Private msngTime As Single
Private mobjLogText As TextStream

Public gobjFile As New FileSystemObject

Global Const gintTends% = 1                       '��ӡ������zlPrintTends
Global gintObjType As Integer                    '��ӡ������ʲô����

Global Const gintMAX_SIZE% = 255                        '�����ַ�������
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_Sub_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const KEY_SET_VALUE = &H2
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_Sub_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const EM_GETLINECOUNT = &HBA&            '��ȡ������
Public Const EM_GETLINE = &HC4&                '����һ���ı���ָ����������

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
' ����ָ����Ϣ�����壬�ȴ�������ŷ��أ��� PostMessage() ����������Ϣ���������أ�HWND hWnd Ŀ�괰��ľ����wMsg�����͵���Ϣ��wParam��Ϣ��һ������lParam��Ϣ�ڶ�������
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Global gstrUserName As String     '��װWindowsʱ��д���û���
Global gstrUnitName As String     '��װWindowsʱ��д�ĵ�λ��

Global gobjOutTo As Object        '��ӡ�����Ŀ�����,������printer��PictureBox
Global gobjSend As Object         'Ҫ��ӡ�Ķ���
Public arrFormat As Variant       '���������Excel���и�ʽ����

Global gintRowTotal As Long    '��ҳ��
Global gintColTotal As Long    '��ҳ����
'ÿҳ�ĵ�һ�е��к������һ�е��кţ���һ�е��к������һ�е��к�
Global gintRow() As Long
Global gintCol() As Long

Global gintPage As Integer        '��ǰ��ʾ��ҳ��
Global gintCopies As Integer      '��ӡ�ķ���

Global gintBegin As Integer       '��ʼҳ��
Global gintShow As Integer        'Ԥ����ҳ��

Global gsngTotalWidth As Single   '����ҳ����ܿ��


Global gsngTitle As Single        '����ĸ߶�
Global gsngUpAppRow As Single     '������Ŀ�ĸ߶�
Global gsngDownAppRow As Single   '������Ŀ�ĸ߶�
Global gsngFixRow As Single       '�̶��еĸ߶�
Global gsngFixCol As Single       '�̶��еĿ��

Global gsngScaleWidth As Single   '���������������������ֽ��ʵ�ʴ�ӡ�Ŀ��
Global gsngScaleHeight As Single  '���������������������ֽ��ʵ�ʴ�ӡ�ĸ߶�
Global gsngHeight As Single       'ҳ�����Ч�߶�
Global gsngWidth As Single        'ҳ�����Ч���
Global gsngPrintedWidth() As Single 'ÿһҳ��ʵ�ʴ�ӡ�˵Ŀ��

Global gsngScale As Single        '���ű���
Global gcolGrid As New Collection '�Ѵ�ӡ��Ԫ��ļ���

Global gfrmTemp  As New frmSample  '�Ѵ�ӡ��Ԫ��ļ���
'-------------------------------------------------------------
Global gstrHeader As String           'ҳü����
Global gsngHeader As Single           'ҳüλ��   '�Ժ���Ϊ��λ
Global gstrFooter As String           'ҳ������
Global gsngFooter As Single           'ҳ��λ��   '�Ժ���Ϊ��λ
Global gsngPageWidth As Single        'ֽ�ſ��   ���Ϊ��λ
Global gsngPageHeight As Single       'ֽ�Ÿ߶�   ���Ϊ��λ
Global gsngPageScaleWidth As Single   'ֽ��ʵ�ʴ�ӡ�Ŀ��   ���Ϊ��λ
Global gsngPageScaleHeight As Single  'ֽ��ʵ�ʴ�ӡ�ĸ߶�   ���Ϊ��λ
Global gintSize As Integer            'ֽ�ŵĳߴ�,�Զ���Ϊ256
Global gintOri As Integer             'ֽ�ŵĽ�ֽ����.2��ʾ����1��ʾ����

Global gsngUp As Single               '�ϱ߾�   '�Ժ���Ϊ��λ
Global gsngDown As Single             '�±߾�   '�Ժ���Ϊ��λ
Global gsngLeft As Single             '��߾�   '�Ժ���Ϊ��λ
Global gsngRight As Single            '�ұ߾�   '�Ժ���Ϊ��λ
Global gstrTabTitle As String         '��������
Global gstrTitleFName As String       '�����������
Global gintTitleFSize As Integer      '����������С
Global gblnTitleFBold As Boolean      '�����Ƿ����
Global gblnTitleFItalic As Boolean    '�����Ƿ�б��
Global glngTitleColor As Long         '�������ɫ
Global gstrAppRowFName As String      '����Ŀ��������
Global gintAppRowFSize As Integer     '����Ŀ�������С
Global gblnAppRowFBold As Boolean     '����Ŀ�Ƿ����
Global gblnAppRowFItalic As Boolean   '����Ŀ�Ƿ�б��
Global glngAppRowColor As Long        '����Ŀ����ɫ
Global gintUpAppRow As Long           '������Ŀ������
Global gintDownAppRow As Long         '������Ŀ������
Global gintTotalRow As Long           '������
Global gintTotalCol As Long           '������
Global gintFixRow As Integer          '�̶��к�
Global gintFixCol As Integer          '�̶��к�

Global gintGroups As Long             '����

Global gstrGrant As String           '"","��ʽ","����","����"

Public glng�ļ�ID As Long
Public glng����ID As Long
Public glng��ҳID As Long
Public gintӤ�� As Integer
Public gstrSQL As String
Public glngPage As Long                 '��ǰҳ��,��һҳ�Ŵ��ڴ�ӡ�������ݲ���ӡ������
Public gblnMoved_HL As Boolean
Public frmAsk As frmTendPrintAsk        'ѯ�ʴ���
Public gstr�Խ��� As String             '�����к����
Public glngPrintRow As Long             '�Ӵ��п�ʼ��ӡ
Public gblnPrintMode As Boolean         '��ӡģʽΪTRUE
Public gintPrintState As Integer        '��ӡģʽ��1-����2-�ش�3-�����ش�

Public Enum enuPage
    ����
    �ش�
    �����ش�
End Enum

'WinNT�Զ���ֽ�ſ���================================================================
'ע����dmFields��Long��,as Long��β����&��
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
'Constants for DocumentProperties() call
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY
Public Const DM_PROMPT = 4
Public Const DM_IN_PROMPT = DM_PROMPT
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
'Constants for DocumentProperties() return
Public Const IDOK = 1
Public Const IDCANCEL = 2
'Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal Hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)



'--------------------------------------------------------------
'ReadVar                ����ӡ��������ݶ���������
'IsWindows95            �ж��Ƿ���Windows95�¹���
'GetWinPlatform         ���ص�ǰ��ϵͳ�汾����
'StripTerminator        ȥ���ַ��������е� Chr$(0)�ַ�
'CalculateRC            Ϊÿһ����Ԫ���������λ��
'CalculateHeight        ��������⡢������Ŀ�ͱ�����Ŀ�ĸ߶�,�̶��еĸ߶ȡ��̶��еĿ��
'PrintPage              ��ָ���豸�ϴ�ӡָ��ҳ
'PrintHeadFoot          ��ӡҳüҳ��
'zlOutTabAppRow         ���listview���ϻ������Ŀ
'zlOutTabAppSet         �������ı��ϻ������Ŀ
'zlOutTitle             �������
'OutRow                 ���һ������
'ConvHF                 ��ҳü��ҳ��ת����ʵ�ʴ�ӡ������
'RealPrint              �����ӡ����
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function ReadVar() As Boolean
'��    �ܣ�����ӡ��������ݶ���������
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��12��
'��    ������
'��    �أ����������Ч�򷵻���
    ReadVar = True
    On Error GoTo errHandle
    gsngPageWidth = Printer.Width
    gsngPageHeight = Printer.Height
    gsngPageScaleWidth = Printer.ScaleWidth
    gsngPageScaleHeight = Printer.ScaleHeight
    
    gintSize = Printer.PaperSize
    gintOri = Printer.Orientation
'    If gintOri = 1 Then '����
'        gsngScaleWidth = IIf(gsngPageScaleWidth < gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight) '�ĵ���ӡ��ֽ��խ����������
'        gsngScaleHeight = IIf(gsngPageScaleWidth > gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight)
'    Else
'        gsngScaleWidth = IIf(gsngPageScaleWidth > gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight) '�ĵ���ӡ��ֽ�Ŀ��������
'        gsngScaleHeight = IIf(gsngPageScaleWidth < gsngPageScaleHeight, gsngPageScaleWidth, gsngPageScaleHeight)
'    End If
    gsngScaleWidth = gsngPageScaleWidth
    gsngScaleHeight = gsngPageScaleHeight
    With gobjSend
        '������������
        gstrTabTitle = .Title.Text
        gstrTitleFName = .Title.Font.Name
        gintTitleFSize = .Title.Font.Size
        gblnTitleFItalic = .Title.Font.Italic
        gblnTitleFBold = .Title.Font.Bold
        glngTitleColor = .Title.Color
        '���������Ŀ�������Ŀ������
        gstrAppRowFName = .AppFont.Name
        gintAppRowFSize = .AppFont.Size
        gblnAppRowFItalic = .AppFont.Italic
        gblnAppRowFBold = .AppFont.Bold
        glngAppRowColor = .AppColor
        gintUpAppRow = .UnderAppRows.Count
        gintDownAppRow = .BelowAppRows.Count
        
        If .FixRow = 0 Then .FixRow = .Body.FixedRows
        gintFixRow = .FixRow
        gintFixCol = .FixCol
        gintGroups = 1
        
        gsngDown = .EmptyDown
        gsngLeft = .EmptyLeft
        gsngRight = .EmptyRight
        gsngUp = .EmptyUp
        gsngHeader = .PageHeader
        gsngFooter = .PageFooter
        
        gstrHeader = .Header
        gstrHeader = IIf(gstrHeader = "", ";;", gstrHeader)
        gstrFooter = .Footer
        gstrFooter = IIf(gstrFooter = "", ";;", gstrFooter)
    End With
    If gsngDown < 0 Or gsngUp < 0 Or gsngLeft < 0 Or gsngRight < 0 Or gsngHeader < 0 Or gsngFooter < 0 Then
        MsgBox "ҳ�߾಻����Ϊ��ֵ��", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngDown + gsngUp) * conRatemmToTwip > gsngScaleHeight Then
        MsgBox "ҳ�ϱ߾��ҳ�±߾��ֵ̫���ˡ�", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngLeft + gsngRight) * conRatemmToTwip > gsngScaleWidth Then
        MsgBox "ҳ��߾��ҳ�ұ߾��ֵ̫���ˡ�", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    If (gsngHeader + gsngFooter) * conRatemmToTwip > gsngScaleHeight Then
        MsgBox "ҳü���ҳ�ž��ֵ̫���ˡ�", vbCritical, gstrSysName
        ReadVar = False
        Exit Function
    End If
    
    Dim strKeyValue As String       '��ֵ
    Dim lngKey As Long
    Dim lngKeySize As Long
    Dim strRegPath As String
    If IsWindows95 Then
        strRegPath = "Software\MicroSoft\Windows\CurrentVersion"
    Else
        strRegPath = "Software\MicroSoft\Windows NT\CurrentVersion"
    End If
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRegPath, 0, KEY_READ, lngKey) = 0 Then
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOrganization", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUnitName = StripTerminator(strKeyValue)
        End If
        strKeyValue = Space(256)
        lngKeySize = 256
        If RegQueryValueEx(lngKey, "RegisteredOwner", 0, 1, strKeyValue, lngKeySize) = 0 Then
            gstrUserName = StripTerminator(strKeyValue)
        End If
    End If
    RegCloseKey lngKey

    gintRowTotal = 0
    gintColTotal = 0
    gintPage = 0
    gsngTotalWidth = 0
    gintCopies = 1
    gintBegin = 1
    gintShow = 1
    Exit Function
errHandle:
    ReadVar = False
End Function

Public Function IsWindowsNT() As Boolean
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Public Function IsWindows95() As Boolean
'��    �ܣ��ж��Ƿ���Windows95�¹���
'��    ������
'��    �أ��Ƿ���True
    Const dwMask95 = &H1&
    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function

Private Function GetWinPlatform() As Long
'��    �ܣ����ص�ǰ��ϵͳ�汾����
'��    ������
'��    �أ�
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

Function StripTerminator(ByVal strString As String) As String
'��    �ܣ�ȥ���ַ��������е� Chr$(0)�ַ�
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��2��
'��    ������
'��    �أ���
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub CalculateRC()
'��    �ܣ�Ϊÿһ����Ԫ���������λ�ã�ҳ�š�ҳ��ţ�
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��2��
'��    ������
'��    �أ���
    Dim intPageRow As Long, intPageCol As Long '��ʱ�õ���ҳ�����ҳ��
    Dim sngPageWidth As Single, sngPageHeight As Single    '��ʱ�õ���ҳ�����ҳ�߶�
    Dim sngRowHeight As Single '�ó�һ���ֵĸ߶�
    Dim intCol As Long      'ʵ�ʵ�����
    Dim i As Long

    Dim iTemp As Long
    Dim sngTemp As Single

    intPageCol = 1
    intPageRow = 1
    gsngTotalWidth = 0
    ReDim gsngPrintedWidth(1 To gintTotalCol)
    ReDim gintRow(1 To 2, 1 To gintTotalRow) '��һά���ڸ�ҳ�ĵ�һ�е��кţ��ڶ�ά���ڸ�ҳ�����һ�е��к�
    ReDim gintCol(1 To 2, 1 To gintTotalCol) '��һά���ڸ�ҳ�ĵ�һ�е��кţ��ڶ�ά���ڸ�ҳ�����һ�е��к�

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo errHand
    '�ӵ�һ���ǹ̶��п�ʼ�����ÿ�е�ҳ���
    gintCol(1, 1) = gintFixCol + 1
    For iTemp = gintFixCol + 1 To gintTotalCol

        '���еĿ��
        If gobjSend.Body.ColHidden(iTemp - 1) Then
            sngTemp = 0
        Else
            sngTemp = gobjSend.Body.ColWidth(iTemp - 1)
        End If
        If sngPageWidth + sngTemp > gsngWidth Then

            '������
            If sngPageWidth = 0 Then

                '��û��һ���ǹ̶���,��ǿ����������
                gsngPrintedWidth(intPageCol) = gsngFixCol + sngTemp
                gintCol(2, intPageCol) = iTemp '��ҳ�����һ�о��Ǳ���
                If iTemp <> gintTotalCol Then  '����Ҫ��ӡ�����һ��
                    intPageCol = intPageCol + 1
                    If intPageCol <= gintTotalCol Then gintCol(1, intPageCol) = iTemp + 1 '��ҳ�ĵ�һ�о��Ǳ���
                End If
                gsngTotalWidth = gsngTotalWidth + sngTemp
            Else

                gsngPrintedWidth(intPageCol) = gsngFixCol + sngPageWidth
                sngPageWidth = 0
                gintCol(2, intPageCol) = iTemp - 1 '��һҳ�����һ�о�����һ��
                intPageCol = intPageCol + 1
                '��һ�з�����һҳ����м���
                'ֻ������ѭ��һ��,����������һ�б�����ֽ��������
                gintCol(1, intPageCol) = iTemp      '��ҳ�ĵ�һ�о��Ǳ���
                iTemp = iTemp - 1
            End If
        Else
            'gintCol(iTemp) = intPageCol
            sngPageWidth = sngPageWidth + sngTemp
            gsngTotalWidth = gsngTotalWidth + sngTemp
        End If
    Next
    If sngPageWidth <> 0 Then 'ͳ�����һҳ�Ŀ��
          gintCol(2, intPageCol) = iTemp - 1 '��һҳ�����һ�о�����һ��
          gsngPrintedWidth(intPageCol) = gsngFixCol + sngPageWidth
    End If

    '�ӵ�һ���ǹ̶��п�ʼ�����ÿ�е�ҳ��
    gintRow(1, 1) = gintFixRow + 1
    For iTemp = gintFixRow + 1 To gintTotalRow
        '���еĸ߶�
        If gobjSend.Body.RowHidden(iTemp - 1) Then
            sngTemp = 0
        Else
            sngTemp = gobjSend.Body.RowHeightMin
        End If
        If sngPageHeight + sngTemp > gsngHeight Then
            '������
            If sngPageHeight = 0 Then
                '��û��һ���ǹ̶���,��ǿ����������
                gintCol(2, intPageRow) = iTemp '��ҳ�����һ�о��Ǳ���
                intPageRow = intPageRow + 1
                If intPageRow <= gintTotalRow Then gintRow(1, intPageRow) = iTemp + 1   '��ҳ�ĵ�һ�о��Ǳ���

            Else
                sngPageHeight = 0
                gintRow(2, intPageRow) = iTemp - 1 '��һҳ�����һ�о�����һ��
                intPageRow = intPageRow + 1
                'ֻ������ѭ��һ��,����������һ�б�����ֽ���ߵ����
                gintRow(1, intPageRow) = iTemp      '��ҳ�ĵ�һ�о��Ǳ���
                iTemp = iTemp - 1
            End If
        Else
            'gintRow(iTemp) = intPageRow
            sngPageHeight = sngPageHeight + sngTemp
        End If
    Next
    If sngPageHeight <> 0 Then gintRow(2, intPageRow) = iTemp - 1 '��һҳ�����һ�о�����һ��

    gintColTotal = intPageCol
    gintRowTotal = intPageRow
    gsngTotalWidth = gsngTotalWidth + gsngFixCol * gintColTotal
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Sub CalculateHeight()
'��    �ܣ���������⡢������Ŀ�ͱ�����Ŀ�ĸ߶�,�̶��еĸ߶ȡ��̶��еĿ��
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��2��
'��    ������
'��    �أ���
    Dim intCol As Long, intRow As Long '��ʱ������ָ�����
    
    '���������ĸ߶�
    gfrmTemp.Font.Name = gstrTitleFName
    gfrmTemp.Font.Size = gintTitleFSize
    gfrmTemp.Font.Bold = gblnTitleFBold
    gfrmTemp.Font.Italic = gblnTitleFItalic
    gsngTitle = gfrmTemp.TextHeight(gstrTabTitle) + 2 * conLineHigh
    '�����������Ŀ�ͱ�����Ŀ�ĸ߶�
    gfrmTemp.Font.Name = gstrAppRowFName
    gfrmTemp.Font.Size = gintAppRowFSize
    gfrmTemp.Font.Bold = gblnAppRowFBold
    gfrmTemp.Font.Italic = gblnAppRowFItalic
    gsngDownAppRow = (gfrmTemp.TextHeight("jg") + conLineHigh) * gintDownAppRow + conLineHigh
    gsngUpAppRow = (gfrmTemp.TextHeight("jg") + conLineHigh) * gintUpAppRow + conLineHigh
    gsngFixRow = 0
    gsngFixCol = 0
            
    gintTotalCol = gobjSend.Body.Cols
    gintTotalRow = gobjSend.Body.Rows
    '������̶��еĸ߶�
    For intRow = 1 To gintFixRow
        gsngFixRow = gsngFixRow + gobjSend.Body.RowHeightMin
    Next
    '������̶��еĿ��(��ӡʱ������̶���)
    For intCol = 1 To gintFixCol
        gsngFixCol = gsngFixCol + gobjSend.Body.ColWidth(intCol - 1)
    Next
    
'    If gintGroups = 1 Then
'        '������̶��еĸ߶�
'        grsGrid.Filter = "�к�=1 and �к�<=" & CStr(gintFixRow)
'        Do Until grsGrid.EOF
'            gsngFixRow = gsngFixRow + grsGrid("�߶�")
'            grsGrid.MoveNext
'        Loop
'        '������̶��еĿ��
'        grsGrid.Filter = "�к�=1 and �к�<=" & CStr(gintFixCol)
'        Do Until grsGrid.EOF
'            gsngFixCol = gsngFixCol + grsGrid("���")
'            grsGrid.MoveNext
'        Loop
'        grsGrid.Filter = ""
'    End If
    gsngHeight = gsngScaleHeight - (gsngUp + gsngDown) * conRatemmToTwip - gsngTitle - gsngDownAppRow - gsngUpAppRow - gsngFixRow - 2 * conLineHigh
    gsngWidth = gsngScaleWidth - (gsngLeft + gsngRight) * conRatemmToTwip - gsngFixCol - 2 * conLineWide
End Sub

Public Sub PrintPage(ByVal intPage As Integer)
'��    �ܣ���ָ���豸�ϴ�ӡָ��ҳ
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��5��
'��    ����intPage  ��ӡ��ҳ��
'��    �أ���
    '��ҳ���ڵ�ҳ����ҳ���
    Dim intPageRow As Long, intPageCol As Long
    Dim sngOriY As Single
    '���Ϊ���ʾ���������ӡ��������ʾfrmBusy����
    
    If intPage = 0 Then Exit Sub
    glngPage = intPage
    intPageRow = 1
    intPageCol = 1
    If intPageCol = 0 Then intPageCol = gintColTotal
    Set gcolGrid = Nothing
    
    Dim sngLeft As Single, sngWidth As Single
    sngLeft = gsngLeft * conRatemmToTwip
    'sngWidth = gsngWidth
    sngWidth = gsngPrintedWidth(intPageCol)
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or glngPage > 1 Or gintPrintState > 1 Then
        zlOutTitle sngLeft, gsngUp * conRatemmToTwip - conLineHigh, sngWidth
    End If
    
    gobjOutTo.CurrentY = gsngUp * conRatemmToTwip + gsngTitle + gsngUpAppRow
            
    PrintTends intPageRow, intPageCol
    sngOriY = gobjOutTo.CurrentY
    
    If glngPrintRow = 0 Or Not gblnPrintMode Or glngPage > 1 Or gintPrintState > 1 Then
        zlOutTabAppSet gobjSend.UnderAppRows, sngLeft, gsngTitle + gsngUp * conRatemmToTwip - conLineHigh, sngWidth
        'PrintHeadFoot
        Call frmTendFileReader.PrintHead
        Call frmTendFileReader.PrintFoot
    End If
    
    gobjOutTo.CurrentY = sngOriY
    zlOutTabAppSet gobjSend.BelowAppRows, sngLeft, gobjOutTo.CurrentY + 100, sngWidth
    
    If gstrGrant <> "" Then
        PrintCell gstrGrant & "����", sngLeft, gsngUp * conRatemmToTwip, sngWidth, sngOriY - gsngUp * conRatemmToTwip, 2, RGB(255, 0, 0), , , "0000", "����", 48 * gsngScale
    End If
End Sub

Public Sub PrintHeadFoot()
'��    �ܣ���ӡҳüҳ��
'�� �� �ˣ���ǿ
'�������ڣ�1999��7��10��
'��    ������
'��    �أ���
    Dim strLeft As String, strMiddle As String, strRight As String
    Dim intPos As Long
    Dim intPos1 As Long
    Dim strHeader As String, strFooter As String
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
    End With
    On Error Resume Next
    strHeader = ConvHF(gstrHeader)
    intPos = InStr(strHeader, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strHeader, 1, intPos - 1)
    intPos = InStr(intPos1, strHeader, ";")
    strMiddle = Mid(strHeader, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strHeader, intPos1)

    PrintCell strLeft, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 0, _
        , , , "0000"
    PrintCell strMiddle, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 2, _
        , , , "0000"
    PrintCell strRight, gsngLeft * conRatemmToTwip, gsngHeader * conRatemmToTwip, gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 1, _
        , , , "0000"
    
    strFooter = ConvHF(gstrFooter)
    intPos = InStr(strFooter, ";")
    intPos1 = intPos + 1
    strLeft = Mid(strFooter, 1, intPos - 1)
    intPos = InStr(intPos1, strFooter, ";")
    strMiddle = Mid(strFooter, intPos1, intPos - intPos1)
    intPos1 = intPos + 1
    strRight = Mid(strFooter, intPos1)

    PrintCell strLeft, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("��"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 0, _
        , , , "0000"
    PrintCell strMiddle, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("��"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 2, _
        , , , "0000"
    PrintCell strRight, gsngLeft * conRatemmToTwip, gsngScaleHeight - gsngFooter * conRatemmToTwip - gobjOutTo.TextHeight("��"), gsngWidth + gsngFixCol, gobjOutTo.TextHeight("��"), 1, _
        , , , "0000"
'    On Error GoTo 0
End Sub

Public Function zlOutTabAppRow(colItem As zlTabAppRow, ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '���ܣ� ������ϻ������Ŀ
    '������
    '   colItem:��Ҫ�����zlPrintLvw����ı��ϻ������Ŀ
    '   X�����ܿ�ȵ�Left ΪX����ʼ��ӡ������������Left
    '   Y:��������Y����
    '   Width: ��ӡ��ʵ�ʿ��
    '���أ�
    '------------------------------------------------
    Dim objApp As zlTabAppItem            '���ϱ�����Ŀ����
    Dim sngXStep As Single               'X����ƽ�Ʋ���
    Dim iCount As Long
    Dim sngCurrentY As Single
    Dim sngCurrentX As Single
    If colItem.Count = 0 Then Exit Function
    
    sngCurrentY = Y
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
        
        iCount = 0
        If colItem.Count = 1 Then
            sngXStep = Width
        Else
            sngXStep = Width / (colItem.Count - 1)
        End If
        For Each objApp In colItem
            iCount = iCount + 1
            .CurrentY = Y
            Select Case iCount
            Case Is = 1                             '������Ŀ
                sngCurrentX = 0
            Case Is = colItem.Count   '������Ŀ
                sngCurrentX = Width - .TextWidth(objApp.Text)
            Case Else                               '������Ŀ
                sngCurrentX = sngXStep * (iCount - 1) - .TextWidth(objApp.Text) / 2
            End Select
            PrintCell objApp.Text, x + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("��"), , _
                , , , "0000"
            
'            OutRow objApp.Text, X, sngCurrentX, Width
        Next

    End With
    zlOutTabAppRow = True
    
End Function

Public Function zlOutTabAppSet(TabAppRows As zlTabAppRows, ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '���ܣ� �������ı��ϻ������Ŀ
    '������
    '   TabAppRows:���ϻ��Ǳ�����Ŀ
    '   X�����ܿ�ȵ�Left ΪX����ʼ��ӡ������������Left
    '   Y:��������Y����
    '   Width: ��ӡ��ʵ�ʿ��
    '���أ�
    '------------------------------------------------
    
    Dim sngXStep As Single             'X����ƽ�Ʋ���
    Dim iCount As Long
    Dim sngCurrentY As Single
    Dim sngCurrentX As Single
    Dim objApp As zlTabAppItem          '���ϱ�����Ŀ����
    Dim colItem As zlTabAppRow          '���ϻ������Ŀ��
    
    Dim strTemp As String
    
    If TabAppRows.Count = 0 Then Exit Function
    sngCurrentY = Y
    With gobjOutTo
        .FontName = gstrAppRowFName
        .FontSize = gintAppRowFSize * gsngScale
        .FontBold = gblnAppRowFBold
        .FontItalic = gblnAppRowFItalic
        .ForeColor = glngAppRowColor
        
        For Each colItem In TabAppRows
            If colItem.Count = 1 Then
                sngXStep = Width
            Else
                sngXStep = Width / (colItem.Count - 1)
            End If
            iCount = 0
            For Each objApp In colItem
                iCount = iCount + 1
                .CurrentY = sngCurrentY
                strTemp = objApp.Text
                Select Case iCount
                Case Is = 1                             '������Ŀ
                    sngCurrentX = 0
                Case Is = colItem.Count                 '������Ŀ
                    sngCurrentX = Width - .TextWidth(strTemp)
                Case Else                               '������Ŀ
                    sngCurrentX = sngXStep * (iCount - 1) - .TextWidth(strTemp) / 2
                End Select
               PrintCell objApp.Text, x + sngCurrentX, .CurrentY, , gobjOutTo.TextHeight("��"), , _
                     , , , "0000"
'                OutRow strTemp, X, sngCurrentX, Width
            Next
            sngCurrentY = sngCurrentY + .TextHeight("ZL")
        Next
    End With
    
    zlOutTabAppSet = True
        
End Function

Public Function zlOutTitle(ByVal x As Single, ByVal Y As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '���ܣ� �������
    '������X�����ܿ�ȵ�Left ΪX����ʼ��ӡ������������Left
    '      Y:��������Y����
    '      Width: ��ӡ��ʵ�ʿ��
    '���أ���
    '------------------------------------------------
    Dim sinLeft As Single
    
    If gstrTabTitle = "" Then Exit Function
    
    With gobjOutTo
        .ForeColor = glngTitleColor
        .FontName = gstrTitleFName
        .FontSize = gintTitleFSize * gsngScale
        .FontBold = gblnTitleFBold
        .FontItalic = gblnTitleFItalic
        .CurrentY = Y
        '����������ʼ��ӡ��λ��
'        sinLeft = (gsngTotalWidth - .TextWidth(gstrTabTitle)) / 2
        PrintCell gstrTabTitle, x, .CurrentY, Width, gobjOutTo.TextHeight("��"), 2, _
            , , , "0000"

'        OutRow gstrTabTitle, X, sinLeft, Width
    End With
    zlOutTitle = True
End Function

Public Function ErrCenter() As Byte
'------------------------------------------------
'���ܣ� �����������������
'������
'���أ� cancel      ���� 0
'       resume      ���� 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    
    bytReturnType = 1
    If gcnOracle.Errors.Count <> 0 Then
        'PL/SQL�洢���̴���
        If gcnOracle.Errors(0).NativeError >= 20000 And gcnOracle.Errors(0).NativeError <= 20200 Then
            '��־����
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            
            strNote = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If
        'ORACLE��������
        '��־����
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "�Ѿ�������ͬ���ݵ����ݣ�Ҫ��Ψһ������[���š����Ƶ�]���ظ�����"
            bytReturnType = 0
        Case 903
            strNote = "�����ƴ���"
        Case 904
            strNote = "�����ƴ���"
        Case 942
            strNote = "�����ͼ�����ڣ��ܿ������㲻�߱�ʹ�øò������ݵ�Ȩ�ޡ�"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "������ж�����м�飺" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "����SQL���Ϊ��" & vbCrLf & vbCrLf & mstrRecentSQL
            End If
        Case 1000
            strNote = "�򿪵����ݱ�̫�࣬��Ҫʱ��ϵͳ����Ա�޸����ݿ��Open_Cursors���á�"
        Case 1005
            strNote = "������û��������롣"
        Case 1017
            strNote = "������û��������롣"
            bytReturnType = 0
        Case 1031
            strNote = "û���㹻��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1045
            strNote = "û���������ݿ��Ȩ�ޡ�"
            bytReturnType = 0
        Case 1400
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ����������ʧ�ܡ�"
            bytReturnType = 0
        Case 1401
            strNote = "���ڸ����ֵ�������п����ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1402
            strNote = "���ڸ����ֵ��������ͼ���������ƣ��������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 1403
            strNote = "����δ���������ݣ����º�������ʧ�ܡ�"
        Case 1404
            strNote = "�޸��в�����������ص�����̫��"
        Case 1405
            strNote = "ȡ�õ���ֵΪ�ա�"
        Case 1406
            strNote = "ȡ�õ���ֵ���ж϶������ˡ�"
        Case 1407
            strNote = "���ڸ�������Ҫ��ǿ��и����˿�ֵ�����¸���ʧ�ܡ�"
            bytReturnType = 0
        Case 1408
            strNote = "ָ�������Ѿ�������������"
        Case 1409
            strNote = "���ܽ�����˳�����(NoSort)����Ϊ�����û����"
        Case 1410
            strNote = "�������ID(ROWID)����ID���������ֺ��ַ���ɵ�16���Ƹ�ʽ��"
        Case 1411
            strNote = "��ǰ�в��ܴ洢����64K�����ݡ�"
            bytReturnType = 0
        Case 1412
            strNote = "��ǰ���������Ͳ��ܴ洢�㳤���ַ�����"
            bytReturnType = 0
        Case 1413
            strNote = "�����С��λ��������ʧ�ܡ�"
            bytReturnType = 0
        Case 1415
            strNote = "���ܶ�һ����ǩα��ָ��������[Outer-Join(+)]"
        Case 1416
            strNote = "���ű���ͬʱָ��һ��������[Outer-Join(+)]"
        Case 1417
            strNote = "һ�ű�ֻ��ָ��ָ�򲻳���һ�ű��������[Outer-Join(+)]"
        Case 1418
            strNote = "ָ�������������ڡ�"
        Case 1424
            strNote = "�������Ч�Ļ����ַ�(ͨ�����ֻ����'%'��'_')��"
        Case 1425
            strNote = "�����ַ������ǳ���Ϊ1���ַ���"
        Case 1426
            strNote = "��ֵ���ʽ���������(̫���̫С)��"
        Case 1427
            strNote = "�����Ӳ�ѯ�����˶��С�"
        Case 1428
            strNote = "�����Ĳ�������򳬽硣"
        Case 1429
            strNote = "һ�����������ڸ�ʽ���硣"
        Case 1430
            strNote = "ϣ�����ӵ����Ѿ����ڡ�"
        Case 1431
            strNote = "��Ȩ����(GRANT)�������ڵĲ�һ�¡�"
        Case 1432
            strNote = "ϣ��ɾ���Ĺ���ͬ����Ѿ������ڡ�"
        Case 1433
            strNote = "ϣ��������ͬ����Ѿ����ڡ�"
        Case 1434
            strNote = "ϣ��ɾ����ͬ����Ѿ������ڡ�"
        Case 1435
            strNote = "ָ�����û������ڡ�"
            bytReturnType = 0
        Case 1438
            strNote = "��ֵ������������ľ�ȷ�̶ȡ�"
        Case 1439, 1440, 1441
            strNote = "ֻ�п�ֵ�в����޸��������͡������Ȼ�ߴ��С"
        Case 1536
            strNote = "ĳ��������ռ�Ŀռ�������"
        Case 2290
            strNote = "������Ŀֵ��������ķ�Χ��Υ���˼��Լ�������������ӻ����ʧ�ܡ�"
            bytReturnType = 0
        Case 2291
            strNote = "����δ��д��ر��д��ڵ���Ŀֵ(Υ�������Լ��)���������ӻ����ʧ�ܡ�"
        Case 2292
            strNote = "��Ϊ�ü�¼�Ѿ�ʹ�ã��ʲ���ɾ���˼�¼��"
            bytReturnType = 0
        Case 12203
            strNote = "������������д�����û���������⣬�����������ӡ�"
            bytReturnType = 0
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "���ڷ����������ߵĽ�ɫ������������ӶԹ��̡�" & strTemp & "������Ȩ��"
            Else
                strNote = "δ֪���󣬷�����" & gcnOracle.Errors(0).Source
            End If
        End Select
        
    Else
        'VB��׼����
        '��־����
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
        Select Case Err.Number
            Case 3, 3 - 2146828288
                strNote = "δ���ñ�׼���ع���"
            Case 5, 5 - 2146828288
                strNote = "��Ч�Ĺ��̻����"
            Case 6, 6 - 2146828288
                strNote = "�������"
            Case 7, 7 - 2146828288
                strNote = "�ڴ����"
            Case 9, 9 - 2146828288
                strNote = "�±곬��"
            Case 10, 10 - 2146828288
                strNote = "�����ǹ̶��������ʱ����"
            Case 11, 11 - 2146828288
                strNote = "����Ϊ��̫С"
            Case 13, 13 - 2146828288
                strNote = "���Ͳ�ƥ��"
            Case 14, 14 - 2146828288
                strNote = "�����ַ���������"
            Case 16, 16 - 2146828288
                strNote = "���ʽ̫����"
            Case 17, 17 - 2146828288
                strNote = "��֧��Ҫ��Ĳ���"
            Case 18, 18 - 2146828288
                strNote = "�������û��ж�"
            Case 20, 20 - 2146828288
                strNote = "�޴��󷵻�"
            Case 28, 28 - 2146828288
                strNote = "��ջ�ռ����"
            Case 35, 35 - 2146828288
                strNote = "���̻���δ����"
            Case 47, 47 - 2146828288
                strNote = " ̫��Ķ�̬����⣨DLL��Ӧ�ÿͻ�"
            Case 48, 48 - 2146828288
                strNote = " ���ö�̬����⣨DLL������"
            Case 49, 49 - 2146828288
                strNote = " ��̬����⣨DLL��Լ������"
            Case 51, 51 - 2146828288
                strNote = "�ڲ�����"
            Case 52, 52 - 2146828288
                strNote = "������ļ������ļ���"
            Case 53, 53 - 2146828288
                strNote = "�ļ�δ�ҵ�"
            Case 54, 54 - 2146828288
                strNote = "�ļ���ʽ����"
            Case 55, 55 - 2146828288
                strNote = "�ļ��Ѿ���"
            Case 57, 57 - 2146828288
                strNote = "�豸���� / �������"
            Case 58, 58 - 2146828288
                strNote = "�ļ��Ѿ�����"
            Case 59, 59 - 2146828288
                strNote = "����ļ�¼����"
            Case 61, 61 - 2146828288
                strNote = "������"
            Case 62, 62 - 2146828288
                strNote = "���볬���ļ�β"
            Case 63, 63 - 2146828288
                strNote = "����ļ�¼��"
            Case 67, 67 - 2146828288
                strNote = "�ļ�̫��"
            Case 68, 68 - 2146828288
                strNote = "�豸��Ч��֧��"
            Case 70, 70 - 2146828288
                strNote = "�ܾ�����"
            Case 71, 71 - 2146828288
                strNote = "����δ׼����"
            Case 74, 74 - 2146828288
                strNote = "��������Ϊ��ͬ��������"
            Case 75, 75 - 2146828288
                strNote = "·�� / �ļ����ʴ���"
            Case 76, 76 - 2146828288
                strNote = "·��δ�ҵ�"
            Case 91, 91 - 2146828288
                strNote = "�������������Ϊ����(δ�½�ʵ��)"
            Case 92, 92 - 2146828288
                strNote = "ѭ��δ��ʼ��"
            Case 93, 93 - 2146828288
                strNote = "�����ģʽ�ַ���"
            Case 94, 94 - 2146828288
                strNote = "�����ʹ�ÿ�(Null)"
            Case 96, 96 - 2146828288
                strNote = " �����Ѿ�ʹ�õĶ���ʱ�䳬���������õ����Ԫ�غţ����²����ܽ����¼�"
            Case 97, 97 - 2146828288
                strNote = "���ܵ���һ��δ����ʵ�����������"
            Case 98, 98 - 2146828288
                strNote = " ����ʹ��һ��˽�ж�������Ժͷ���?�����ͷ���ֵ"
            Case 321, 321 - 2146828288
                strNote = "������ļ���ʽ"
            Case 322, 322 - 2146828288
                strNote = "���ܴ�����Ҫ����ʱ�ļ�"
            Case 325, 325 - 2146828288
                strNote = "��Դ�ļ��д���ĸ�ʽ"
            Case 380, 380 - 2146828288
                strNote = "���������ֵ"
            Case 381, 381 - 2146828288
                strNote = "�����������������"
            Case 382, 382 - 2146828288
                strNote = "��֧�ֵ�����ʱ����"
            Case 383, 383 - 2146828288
                strNote = "��֧�ֵ�ֻ����������"
            Case 385, 384 - 2146828288
                strNote = "��Ҫ������������"
            Case 387, 387 - 2146828288
                strNote = "�����������"
            Case 393, 393 - 2146828288
                strNote = "��֧�ֵ�����ʱ��ȡ"
            Case 394, 394 - 2146828288
                strNote = "��֧�ֵ�ֻд���Զ�ȡ"
            Case 422, 422 - 2146828288
                strNote = "�����ڵ�����"
            Case 423, 423 - 2146828288
                strNote = "�����ڵ����Ի򷽷�"
            Case 424, 424 - 2146828288
                strNote = "Ҫ��һ������"
            Case 429, 429 - 2146828288
                strNote = "ActiveX���ܴ�������"
            Case 430, 430 - 2146828288
                strNote = "�಻֧�ֵ��Զ���������֧�ֵĽ���"
            Case 432, 432 - 2146828288
                strNote = "���Զ������ڼ�δ�ҵ��ļ�����������"
            Case 438, 438 - 2146828288
                strNote = "����֧�ָ����Ի򷽷�"
            Case 440, 440 - 2146828288
                strNote = "�Զ����������"
            Case 442, 442 - 2146828288
                strNote = "��Զ��������������ᶪʧ����OK����Ի���ȥ����"
            Case 443, 443 - 2146828288
                strNote = "�Զ�������û��ȱʡֵ"
            Case 445, 445 - 2146828288
                strNote = "����֧�����ֲ���"
            Case 446, 446 - 2146828288
                strNote = "����֧����������"
            Case 447, 447 - 2146828288
                strNote = "����֧�ֵ�ǰ��������"
            Case 448, 448 - 2146828288
                strNote = "��������δ�ҵ�"
            Case 449, 449 - 2146828288
                strNote = "�������ǿ�ѡ��"
            Case 450, 450 - 2146828288
                strNote = "����Ĳ������������Է���"
            Case 451, 451 - 2146828288
                strNote = "���Ը�ֵ(Let)���̺Ͷ�ȡ(Get)���̲����ض���"
            Case 452, 452 - 2146828288
                strNote = "��Ч�����"
            Case 453, 453 - 2146828288
                strNote = "ָ����DLL����δ�ҵ�"
            Case 454, 454 - 2146828288
                strNote = "������Դδ�ҵ�"
            Case 455, 455 - 2146828288
                strNote = "������Դ��������"
            Case 457, 457 - 2146828288
                strNote = "�ùؼ�ֵ�Ѿ��뼯�ϵ���һԪ�ؽ��"
            Case 458, 458 - 2146828288
                strNote = "VB��֧�ֵĿɱ��Զ�������"
            Case 459, 459 - 2146828288
                strNote = "������಻֧�ֵ��¼���"
            Case 460, 460 - 2146828288
                strNote = "����ļ������ʽ"
            Case 461, 461 - 2146828288
                strNote = "���������ݳ�Աδ�ҵ�"
            Case 462, 462 - 2146828288
                strNote = "Զ�̷����������ڻ���Ч"
            Case 463, 463 - 2146828288
                strNote = "��û���ڱ���ע��"
            Case 481, 481 - 2146828288
                strNote = "��Ч��ͼƬ��ʽ"
            Case 482, 482 - 2146828288
                strNote = "��ӡ������"
            Case 735, 735 - 2146828288
                strNote = "���ܽ��洢Ϊ��ʱ�ļ�"
            Case 744, 744 - 2146828288
                strNote = "δ�ҵ�����������"
            Case 746, 746 - 2146828288
                strNote = "̫���ĸ���"
            'ADO����
            Case 3001
                strNote = "�������ʹ��󣬻���ֵ������Χ�������ͻ��"
            Case 3021
                strNote = "��¼����(EOF/BOF)�����ߵ�ǰ��¼��ɾ������ǰӦ�ò�����Ҫ��λ��ǰ��¼��"
            Case 3219
                strNote = "�����Ļ���������ǰӦ�ò����������Ǵ�����δ���������񣩡�"
            Case 3246
                strNote = "������ִ���У����ܹر�һ���������"
            Case 3251
                strNote = "��ǰ������֧����һӦ�ò�����"
            Case 3265
                strNote = "ADOû�ҵ�Ӧ�ó���Ҫ��Ķ�Ӧ���ƻ���š�"
            Case 3367
                strNote = "�����Ѿ����ڣ�������ӡ�"
            Case 3420
                strNote = "����δ���á�"
            Case 3421
                strNote = "��ǰ����ʹ���˴������ֵ���͡�"
            Case 3704
                strNote = "����ر�ʱ����ǰ��������ִ�С�"
            Case 3705
                strNote = "������ʱ����ǰ��������ִ�С�"
            Case 3706
                strNote = "ADOû�ҵ�ָ����֧�֡�"
            Case 3707
                strNote = "���ܲ����������ı�һ����¼���Ļ����Դ�����ԡ�"
            Case 3708
                strNote = "Ӧ�ó�����ִ���Ĳ������塣"
            Case 3709
                strNote = "Ӧ�ó���Ҫ��һ���رյ����ö������Ч���������"
            Case Else
                strNote = "�����ڽ���δ֪����"
        End Select
        bytReturnType = 0
    End If

'    If bytReturnType = 1 Then
'        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
'    Else
'        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
'        ErrCenter = 0
'    End If
    
    '�����ش����
    Err.Clear
End Function

Private Function mGetInvalidTable() As String
'���ܣ��õ������ʹ�õ�SQL����в��ܷ��ʵı����ͼ
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(mstrRecentSQL), ",")
    
    On Error Resume Next
    For lngCount = LBound(varTables) To LBound(varTables)
        strTable = varTables(lngCount)
        
        '���Ըö����Ƿ����
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        'ȥ����һ������
        mGetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function
Public Function SQLObject(ByVal strSQL As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����(��ʱstrAnalΪ�򵥲�ѯ,���ܴ�Union������)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & ",", "," & strTrue & ",") = 0 And strTrue <> "Ƕ�ײ�ѯ" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Private Function TrueObject(ByVal strObject As String) As String
'���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim i As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    'Ѱ�Һ����һ���������ַ�
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function
Public Function TrimChar(str As String) As String
'����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
'    i = InStr(strTmp, "  ")
'    Do While i > 0
'        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
'        i = InStr(strTmp, "  ")
'    Loop
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
'    i = InStr(1, strTmp, vbCrLf & vbCrLf)
'    Do While i > 0
'        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
'        i = InStr(1, strTmp, vbCrLf & vbCrLf)
'    Loop

    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'���ܣ�ͨ��Command����򿪴�����SQL�ļ�¼��
'������strSQL=�����а���������SQL���,������ʽΪ"[x]"
'             x>=1Ϊ�Զ��������,"[]"֮�䲻���пո�
'             ͬһ�������ɶദʹ��,�����Զ���ΪADO֧�ֵ�"?"����ʽ
'             ʵ��ʹ�õĲ����ſɲ�����,������Ĳ���ֵ��������(��SQL���ʱ��һ��Ҫ�õ��Ĳ���)
'      arrInput=���������Ĳ���ֵ,��������˳�����δ���,��������ȷ����
'      strTitle=����SQLTestʶ��ĵ��ô���/ģ�����
'���أ���¼����CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'������
'SQL���Ϊ="Select ���� From ������Ϣ Where (����ID=[3] Or �����=[3] Or ���� Like [4]) And �Ա�=[5] And �Ǽ�ʱ�� Between [1] And [2] And ���� IN([6],[7])"
'���÷�ʽΪ��Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!ת������,"yyyy-MM-dd")),dtpʱ��.Value, lng����ID, "��%", "��", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '�����Զ���[x]����
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '������������"[����]����"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '�滻Ϊ"?"����
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '��������SQL���ٵ����
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '�ַ�
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '����
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '�����ش�ԭ�в���:��Ȼ�����ظ�ִ��
    cmdData.CommandText = "" '��Ϊ����ʱ�����ش��������
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '�����µĲ���
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '�ַ�
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax = 0 Or intMax < 200 Then intMax = 200
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
        Case "Date" '����
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '����
            '���ַ�ʽ������һЩIN�Ӿ��Union���
            '��ʾͬһ�������Ķ��ֵ,�����Ų�������������Ĳ����Ž���,��Ҫ��֤�����ֵ��������
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '�ַ�
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax = 0 Or intMax < 200 Then intMax = 200
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '����
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '�ò������������õ��ڼ���ֵ��
        End Select
    Next

    'ִ�з��ؼ�¼��
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '���Ƚ���
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Public Sub SQLTest(Optional ByVal strProject As String, Optional ByVal strForm As String, Optional ByVal strSQL As String, Optional ByVal strNote As String)
'���ܣ���������ִ�е�SQL��������������ļ��У������ӿ�ʼ����ʱ�䣬ִ��ʱ��
'������strProject=��������,�����ȡApp.Title
'      strForm=������,�����ȡForm.Caption
'      strSQL=��Ҫִ�е�SQL���,��Openʱ����,�����������ʾ���һ��SQLִ�����
'      strNote=SQL���˵��
    Dim strTmp As String, sngEnd As Single
    
    mstrRecentSQL = strSQL  '�������ִ�е�SQL���
    
    If UCase(gcnOracle.ConnectionString) Like "*;SERVER=SQLLOG*" Then
        If strSQL <> "" Then
            If mobjLogText Is Nothing Then
                On Local Error Resume Next
                Set mobjLogText = gobjFile.OpenTextFile("ReportSQL_" & gstrDBUser & "_" & Format(date, "yyyyMMdd") & ".log", ForAppending, True, TristateFalse)
                On Local Error GoTo 0
            End If
            If Not mobjLogText Is Nothing Then
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Application:" & strProject & "\" & strForm & IIf(strNote <> "", "," & strNote, "")
                mobjLogText.WriteLine strTmp & "SQL:" & strSQL
                msngTime = Timer
            End If
        Else
            If Not mobjLogText Is Nothing Then
                sngEnd = Timer
                strTmp = "[" & Format(Time, "HH:mm:ss") & "]"
                mobjLogText.WriteLine strTmp & "Expend:" & Format(sngEnd - msngTime, "0.0000")
                mobjLogText.WriteBlankLines 1
            End If
        End If
    End If
End Sub
Public Function GetPrivFunc(lngSys As Long, lngProgID As Long) As String
'���ܣ����ص�ǰ�û����е�ָ������Ĺ��ܴ�
'������lngSys     ����ǹ̶�ģ�飬��Ϊ0
'      lngProgId  �������
'���أ��ֺż���Ĺ��ܴ�,Ϊ�ձ�ʾû��Ȩ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPrivs As String
    Dim strWhere As String
    
    On Error GoTo errH
    
'    If zlRegCheck <> "" Then Exit Function
    
    strSQL = "Select Text as ���� From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
    Set rsTmp = OpenSQLRecord(strSQL, "GetPrivFunc", lngSys, lngProgID)
    Do While Not rsTmp.EOF
        strPrivs = strPrivs & ";" & rsTmp!����
        rsTmp.MoveNext
    Loop
    GetPrivFunc = Mid(strPrivs, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Public Function OutRow(ByVal strPrint As String, ByVal x As Single, ByVal sngLeft As Single, ByVal Width As Single) As Boolean
    '------------------------------------------------
    '���ܣ� ���һ������
    '������X�����ܿ�ȵ�Left ΪX����ʼ��ӡ������������Left
    '      Y:��������Y����
    '      Width: ��ӡ��ʵ�ʿ��
    '���أ���
    '------------------------------------------------
    Dim strTemp As String
    With gobjOutTo
        If sngLeft >= x Then 'ǰ�滹��һ�οհ�
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft - x
        Else
            Do While sngLeft + .TextWidth(strTemp) < x
                    If Len(strPrint) = 0 Then Exit Do
                    strTemp = strTemp & Left(strPrint, 1)
                    strPrint = Mid(strPrint, 2)
            Loop
            .CurrentX = gsngLeft * conRatemmToTwip + sngLeft + .TextWidth(strTemp) - x
        End If
        Dim intPageCol As Long
        intPageCol = gintPage Mod gintColTotal
        If intPageCol = 0 Then intPageCol = gintColTotal
        strTemp = ""
        Do While (.TextWidth(strPrint) > Width) Or (.CurrentX + .TextWidth(strPrint) > gsngLeft * conRatemmToTwip + gsngPrintedWidth(intPageCol))
            If Len(strPrint) = 0 Then Exit Do
            strTemp = Right(strPrint, 1) & strTemp
            strPrint = Mid(strPrint, 1, Len(strPrint) - 1)
        Loop
        If Len(strTemp) > 0 And .CurrentX < gsngLeft * conRatemmToTwip + gsngPrintedWidth(intPageCol) + .TextWidth("��") Then strPrint = strPrint & Left(strTemp, 1)
        If Len(strPrint) = 0 Then Exit Function
        gobjOutTo.Print strPrint
    End With
End Function

Public Function ConvHF(ByVal strSource As String) As String
    '------------------------------------------------
    '���ܣ���ҳü��ҳ��ת����ʵ�ʴ�ӡ������
    '������strSource    ҳü��ҳ��
    '���أ�ʵ�ʴ�ӡ������
    '------------------------------------------------
    Dim strTemp As String
    
    strTemp = Replace(strSource, "[ҳ��]", CStr(gintPage + gintBegin - 1))
    strTemp = Replace(strTemp, "[ҳ��]", CStr(gintColTotal * gintRowTotal))
    strTemp = Replace(strTemp, "[ʱ��]", Format(Time, "HH:MM:SS"))
    strTemp = Replace(strTemp, "[����]", Format(date, "YYYY��mm��dd��"))
    strTemp = Replace(strTemp, "[�û���]", gstrUserName)
    strTemp = Replace(strTemp, "[��λ��]", gstrUnitName)
    ConvHF = strTemp
End Function

Public Sub RealPrint(ByVal intBegin As Long, ByVal intEnd As Long)
    '���ܣ� �����ӡ����
    '������intBegin     ��ʼҳ��
    '      intEnd       ����ҳ��
    '���أ���
    '------------------------------------------------
    Dim frmOutTemp As New frmOutStatus
    On Error Resume Next
    Screen.MousePointer = 11
    frmOutTemp.mintBegin = intBegin
    frmOutTemp.mintEnd = intEnd
    frmOutTemp.Show 1
    Unload frmOutTemp
    Set frmOutTemp = Nothing
    Screen.MousePointer = 0
End Sub


Public Sub ApplyOEM(objStatus As Object)
'���״̬��Ӧ��OEM����
    Dim strOEM As String
    On Error Resume Next
    
    If gstrSysName <> "-" Then
        objStatus.Panels(1).Text = gstrSysName
        '����״̬��ͼ���OEM����
        If gstrSysName = "�������" Then
            Set objStatus.Panels(1).Picture = LoadCustomPicture("Logo")
        Else
            strOEM = GetOEM(Mid(gstrSysName, 1, Len(gstrSysName) - 2))
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

Private Function GetOEM(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ÿ�����ߵ�ASCII��
    '������
    '���أ�
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    strCode = "OEM_"
    For intBit = 1 To Len(strAsk)
        'ȡÿ���ֵ�ASCII��
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function SetNTPrinterPaper(ByVal lngHwnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'���ܣ�NT�����У����ô�ӡ�����Զ���ֽ�ųߴ�
'������lngWidth��lngHeight=mm(����)
'     intOrient=1-����,2-����
'     intCopys=��ӡ����(�����ӡ��֧��,1-9999,��֧��ʱ�������,Ҳ��Ӱ����������)
'˵��������Width,Height�⣬����ͨ�����������õ����Բ�ֱ�ӷ�ӳ��Printer�ϣ�
'      (ȡDevModeҲ��ӳ������������Ҫ��GetJob���ܻ�ȡ����Ĵ�ӡ�ĵ�����)
    Dim vDevMode As DEVMODE
    Dim arrDevMode() As Byte
    Dim lngSize As Long
    
    Dim lngPrtDC As Long
    Dim lngHandle As Long
    Dim strPrtName As String
    
    lngPrtDC = Printer.hDC
    strPrtName = Printer.DeviceName
    
    If OpenPrinter(strPrtName, lngHandle, 0&) Then
        'Retrieve the size of the DEVMODE:fMode=0
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
        'Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(vDevMode, arrDevMode(1), Len(vDevMode))
        
        '���ô�ӡ�ĵ�����
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '�߼���ӡ����(��ȡ��ʱ,Copiesֻ֧��1;����֪��ôȡ����)
        vDevMode.dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERLENGTH Or DM_PAPERWIDTH Or DM_COPIES 'Or DM_COLLATE
        
        'Copy your changes back, then update DEVMODE.
        Call CopyMemory(arrDevMode(1), vDevMode, Len(vDevMode))
        If blnPrompt Then
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHwnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetCustonPager(ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'���ܣ��������Զ���ֽ��
'�����������Ϊ��λ
    If IsWindowsNT Then
        '��Ȼ����ʹ�����Ч�����ܸı�PaperSize������ֵ
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustonPager = SetNTPrinterPaper(gfrmTemp.Hwnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98ϵ�л�����ͨ����������
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetPrinterSet() As Boolean
'------------------------------------------------
    '���ܣ���ȡ��ϵͳע���Ĵ�ӡȱʡ����
    '------------------------------------------------
    Dim iCount As Long
    Dim strDeviceName As String
    Dim intPaperSize As Integer
    Dim intPaperBin As Integer
    Dim intOrientation As Long
    
    If Printers.Count = 0 Then
        GetPrinterSet = False
        Exit Function
    End If
    
    strDeviceName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
    If Printer.DeviceName <> strDeviceName Then
        For iCount = 0 To Printers.Count - 1
            If Printers(iCount).DeviceName = strDeviceName Then
                Set Printer = Printers(iCount)
                Exit For
            End If
        Next
    End If
    
    Err = 0
    On Error Resume Next
    Printer.PaperBin = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperBin", Printer.PaperBin)
    Printer.Orientation = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Orientation", Printer.Orientation)
    
    intPaperSize = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "PaperSize", Printer.PaperSize)
    If intPaperSize = 256 Then
        Dim lngWidth As Long
        Dim lngHeight As Long
        
        lngWidth = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Width", Printer.Width)
        lngHeight = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "Height", Printer.Height)
        
        Call SetCustonPager(lngWidth, lngHeight)
    Else
        Printer.PaperSize = intPaperSize
    End If
    GetPrinterSet = True
End Function

Public Function ReadPageHead(objHead As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Public Function ReadPageFoot(objFoot As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = GetTmpPath
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
errHand:
    Exit Function
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

