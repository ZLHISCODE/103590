Attribute VB_Name = "mdlPublic"
Option Explicit
Public gcnOracle As ADODB.Connection
Public gstrPrivs As String
Public gstrSQL As String
Public gblnMoved As Boolean

Global gfrmTemp  As New frmSample
Public mfrmPartogram As Object '�����������
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

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

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
' ����ָ����Ϣ�����壬�ȴ�������ŷ��أ��� PostMessage() ����������Ϣ���������أ�HWND hWnd Ŀ�괰��ľ����wMsg�����͵���Ϣ��wParam��Ϣ��һ������lParam��Ϣ�ڶ�������
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const SW_RESTORE = 9
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal Hwnd As Long) As Long 'Ҫ����ˢ��
Public Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal Hwnd As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWHEELSCROLLLINES = 104
Public WHEEL_SCROLL_LINES As Long
Global glngPrevWndProc As Long

Public Const WM_MOUSEWHEEL = &H20A

'����Ļ��ĳ�������Ļ����ת��Ϊ�ͻ���������
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long

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
Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal Hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'��ȡ��ʾ�����ߴ�ӡ������Ϣ
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Public Const conRatemmToTwip = 56.6857142857143    '������羵ı���
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Function ExistsPrinter() As Boolean
    Dim lngHDc As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDc = Printer.hDC
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
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

Public Function DelCustomPaper() As Boolean
    '���ܣ�ɾ���ղŴ������Զ���ֽ��
    Dim lngHandle As Long
    Dim strName As String
    
    strName = Printer.DeviceName
    If OpenPrinter(strName, lngHandle, 0&) Then
        DelCustomPaper = (DeleteForm(lngHandle, "zlBillPaper" & Chr(0)) <> 0)
        Call ClosePrinter(lngHandle)
    End If
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
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
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


'---------------------------------------------------------------------------------
'�����ǻ������������
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, _
                      ByVal strFields As String, _
                      ByVal strValues As String)

    Dim arrFields, arrValues, intField As Integer

    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)

    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew

        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next

        .Update
    End With

End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, _
                         ByVal strFields As String, _
                         ByVal strValues As String, _
                         ByVal strPrimary As String, _
                         Optional ByVal blnDelete As Boolean = False)

    Dim arrFields, arrValues, intField As Integer

    '���¼�¼,���������,������
    'strPrimary:�ֶ���,ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)

    If intField < 0 Then Exit Sub

    With rsObj

        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew

        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next

        .Update
    End With

End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, _
                              ByVal strPrimary As String, _
                              Optional ByVal blnDelete As Boolean = False) As Boolean

    Dim arrTmp

    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")

    With rsObj

        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"

        If .EOF Then Exit Function
        If blnDelete Then

            Do While Not .EOF

                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop

        Else
            Record_Locate = True
        End If

    End With

End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)

    Dim arrFields, intField As Integer

    Dim strFieldName As String, intType As Integer, lngLength As Long

    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj

        If .State = 1 Then .Close

        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then

                Select Case intType

                    Case adDouble
                        lngLength = madDoubleDefault

                    Case adVarChar
                        lngLength = madLongVarCharDefault

                    Case adLongVarChar
                        lngLength = madLongVarCharDefault

                    Case Else
                        lngLength = madDbDateDefault
                End Select

            End If

            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

End Sub

Public Sub OutputRsData(ByVal rsObj As ADODB.Recordset, _
                         Optional ByVal blnMod_Add As Boolean = False)

    Dim strOutput As String
    Dim intCol    As Integer, intCols As Integer
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strOutput = ""
            intCols = .Fields.Count
            For intCol = 1 To intCols
                If Not blnMod_Add Then
                    strOutput = strOutput & "," & .Fields(intCol - 1).Name & ":" & .Fields(intCol - 1).Value
                Else
                    strOutput = strOutput & "|" & .Fields(intCol - 1).Value
                End If
            Next
            Debug.Print Mid(strOutput, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub


Public Sub Hook(ByVal Hwnd As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    glngPrevWndProc = SetWindowLong(Hwnd, GWL_WNDPROC, AddressOf WindowProc)

    '��ȡ"�������"�еĹ�������ֵ

    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WHEEL_SCROLL_LINES, 0)

    If WHEEL_SCROLL_LINES > mfrmPartogram.ScrollBarY.Max Then WHEEL_SCROLL_LINES = mfrmPartogram.ScrollBarY.Max
End Sub

Public Sub UnHook(ByVal Hwnd As Long)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngReturnValue As Long

    lngReturnValue = SetWindowLong(Hwnd, GWL_WNDPROC, glngPrevWndProc)
    Set mfrmPartogram = Nothing
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '******************************************************************************************************************
    '���ܣ�����ϵͳ�¼������д���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim pt As POINTAPI
    Dim wzDelta
    Dim wKeys As Integer
    
    Select Case uMsg
    Case WM_MOUSEWHEEL                          '�����¼�
        wzDelta = HIWORD(wParam)
        wKeys = LOWORD(wParam)
        pt.x = LOWORD(lParam)
        pt.Y = HIWORD(lParam)
    
        '����Ļ����ת��ΪfrmCaseTendBody��������
    
        ScreenToClient mfrmPartogram.Hwnd, pt

        With mfrmPartogram
        
            '�ж������Ƿ���frmCaseTendBody.BodyEdit������
    
            If pt.x > .Left / Screen.TwipsPerPixelX And pt.x < (.Left + .Width) / Screen.TwipsPerPixelX And pt.Y > .Top / Screen.TwipsPerPixelY And pt.Y < (.Top + .Height) / Screen.TwipsPerPixelY Then
    
                If wKeys = 16 Then
                    'ˮƽ����
                    
                Else
                    '��ֱ����
                    If Sgn(wzDelta) = 1 Then
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value - WHEEL_SCROLL_LINES < .ScrollBarY.Min, .ScrollBarY.Min, .ScrollBarY.Value = .ScrollBarY.Value - WHEEL_SCROLL_LINES)
                    Else
                        .ScrollBarY.Value = IIf(.ScrollBarY.Value + WHEEL_SCROLL_LINES > .ScrollBarY.Max, .ScrollBarY.Max, .ScrollBarY.Value + WHEEL_SCROLL_LINES)
                    End If
                End If
            End If
        End With
    Case Else                                   '�����¼�����ϵͳȱʡ����
        WindowProc = CallWindowProc(glngPrevWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function

'*************************************************************************
'**�� �� ����HIWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĸ�16λ
'*************************************************************************
Public Function HIWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĸ�16λ
     HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

'*************************************************************************
'**�� �� ����LOWORD
'**��    �룺LongIn(Long) - 32λֵ
'**��    ����(Integer) - 32λֵ�ĵ�16λ
'**����������ȡ��32λֵ�ĵ�16λ
'*************************************************************************
Public Function LOWORD(LongIn As Long) As Integer
   ' ȡ��32λֵ�ĵ�16λ
     LOWORD = LongIn And &HFFFF&
End Function
