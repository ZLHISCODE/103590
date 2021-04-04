Attribute VB_Name = "mPrintGDI"
Option Explicit
Public Const conRatemmToTwip As Single = 56.6857142857143      '毫米与缇的比率

'######################################################################################
'## 判断系统是否是NT
'######################################################################################

Public Function IsWindowsNT() As Boolean
'功能：是否WindowNT操作系统
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

Private Function GetWinPlatform() As Long
'功    能：返回当前的系统版本代号
'参    数：无
'返    回：
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformId
End Function

'######################################################################################
'## 设置打印机的自定义纸张尺寸
'######################################################################################

Public Function SetNTPrinterPaper(ByVal lngHWnd As Long, ByVal intWidth As Integer, ByVal intHeight As Integer, _
    ByVal intOrient As Integer, ByVal intCopys As Integer, Optional ByVal blnPrompt As Boolean) As Boolean
'功能：NT环境中，设置打印机的自定义纸张尺寸
'参数：lngWidth、lngHeight=mm(毫米)
'     intOrient=1-纵向,2-横向
'     intCopys=打印份数(如果打印机支持,1-9999,不支持时不会出错,也不影响其它设置)
'说明：除了Width,Height外，其它通过本函数设置的属性不直接反映在Printer上，
'      (取DevMode也反映不出来，可能要用GetJob才能获取最近的打印文档属性)
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
        
        '设置打印文档属性
        vDevMode.dmOrientation = intOrient
        vDevMode.dmPaperSize = 256
        vDevMode.dmPaperWidth = intWidth * 10 'in tenths of a millimeter
        vDevMode.dmPaperLength = intHeight * 10 'in tenths of a millimeter
        vDevMode.dmCopies = intCopys
        'vDevMode.dmCollate = 0& '高级打印功能(当取消时,Copies只支持1;但不知怎么取不了)
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
'功能：在设置自定义纸张
'参数：是以绨为单位
    If IsWindowsNT Then
        '虽然不能使宽度生效，但能改变PaperSize的属性值
        Printer.Width = lngWidth
        Printer.Height = lngHeight
        SetCustomPager = SetNTPrinterPaper(lngHWnd, lngWidth / conRatemmToTwip, lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
    Else
        'Windows98系列还是以通常方法处理
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
