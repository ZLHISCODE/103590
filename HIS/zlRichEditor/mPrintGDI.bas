Attribute VB_Name = "mPrintGDI"
Option Explicit
Public Const conRatemmToTwip As Single = 56.6857142857143      '������羵ı���

'######################################################################################
'## �ж�ϵͳ�Ƿ���NT
'######################################################################################

Public Function IsWindowsNT() As Boolean
'���ܣ��Ƿ�WindowNT����ϵͳ
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
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

'######################################################################################
'## ���ô�ӡ�����Զ���ֽ�ųߴ�
'######################################################################################

Public Function SetNTPrinterPaper(ByVal lngHWnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
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
        lngSize = DocumentProperties(lngHWnd, lngHandle, strPrtName, 0&, 0&, 0&)
        'Reserve memory for the actual size of the DEVMODE.
        ReDim arrDevMode(1 To lngSize)
    
        'Fill the DEVMODE from the printer.
        lngSize = DocumentProperties(lngHWnd, lngHandle, strPrtName, arrDevMode(1), 0&, DM_OUT_BUFFER)
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
            lngSize = DocumentProperties(lngHWnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_IN_PROMPT Or DM_OUT_BUFFER)
        Else
            lngSize = DocumentProperties(lngHWnd, lngHandle, strPrtName, arrDevMode(1), arrDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
        End If
        If lngSize = IDOK Then SetNTPrinterPaper = True
        'Reset the DEVMODE for the DC.
        lngSize = ResetDC(lngPrtDC, arrDevMode(1))
        If lngSize = 0 Then SetNTPrinterPaper = False
        
        'Close the handle when you are finished with it.
        Call ClosePrinter(lngHandle)
    End If
End Function

Public Function SetCustomPager(ByRef lngHWnd As Long, ByVal lngWidth As Long, ByVal lngHeight As Long) As Integer
'���ܣ��������Զ���ֽ��
'�����������Ϊ��λ
    If IsWindowsNT Then
        '��Ȼ����ʹ�����Ч�����ܸı�PaperSize������ֵ
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustomPager = SetNTPrinterPaper(lngHWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98ϵ�л�����ͨ����������
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    End If
End Function

Public Function ExistsPrinter() As Boolean
    Dim lngHDC As Long
    
    If Printers.Count = 0 Then Exit Function
    
    On Error Resume Next
    lngHDC = Printer.hDC
    If Err.Number = 0 Then ExistsPrinter = True
    Err.Clear: On Error GoTo 0
End Function
