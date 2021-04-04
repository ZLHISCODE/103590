Attribute VB_Name = "mdlPrint"
Option Explicit

Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Type JOB_INFO_2
   JobId As Long
   pPrinterName As Long
   pMachineName As Long
   pUserName As Long
   pDocument As Long
   pNotifyName As Long
   pDatatype As Long
   pPrintProcessor As Long
   pParameters As Long
   pDriverName As Long
   pDevMode As Long
   pStatus As Long
   pSecurityDescriptor As Long
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   time As Long
   PagesPrinted As Long
End Type

Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
   (ByVal hPrinter As Long, _
   ByVal Level As Long, _
   pPrinter As Byte, _
   ByVal cbBuf As Long, _
   pcbNeeded As Long) _
   As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
   (ByVal hPrinter As Long) _
   As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
   (ByVal hPrinter As Long, _
   ByVal FirstJob As Long, _
   ByVal NoJobs As Long, _
   ByVal Level As Long, _
   pJob As Byte, _
   ByVal cdBuf As Long, _
   pcbNeeded As Long, _
   pcReturned As Long) _
   As Long


Public Declare Function OpenPrinter Lib "winspool.drv" _
   Alias "OpenPrinterA" _
   (ByVal pPrinterName As String, _
   phPrinter As Long, _
   pDefault As PRINTER_DEFAULTS) _
   As Long

Public Const MCONERROR_INSUFFICIENT_BUFFER = 122
Public Const MCONPRINTER_STATUS_BUSY = &H200
Public Const MCONPRINTER_STATUS_DOOR_OPEN = &H400000
Public Const MCONPRINTER_STATUS_ERROR = &H2
Public Const MCONPRINTER_STATUS_INITIALIZING = &H8000
Public Const MCONPRINTER_STATUS_IO_ACTIVE = &H100
Public Const MCONPRINTER_STATUS_MANUAL_FEED = &H20
Public Const MCONPRINTER_STATUS_NO_TONER = &H40000
Public Const MCONPRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const MCONPRINTER_STATUS_OFFLINE = &H80
Public Const MCONPRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const MCONPRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const MCONPRINTER_STATUS_PAGE_PUNT = &H80000
Public Const MCONPRINTER_STATUS_PAPER_JAM = &H8
Public Const MCONPRINTER_STATUS_PAPER_OUT = &H10
Public Const MCONPRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const MCONPRINTER_STATUS_PAUSED = &H1
Public Const MCONPRINTER_STATUS_PENDING_DELETION = &H4
Public Const MCONPRINTER_STATUS_PRINTING = &H400
Public Const MCONPRINTER_STATUS_PROCESSING = &H4000
Public Const MCONPRINTER_STATUS_TONER_LOW = &H20000
Public Const MCONPRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const MCONPRINTER_STATUS_WAITING = &H2000
Public Const MCONPRINTER_STATUS_WARMING_UP = &H10000
Public Const MCONJOB_STATUS_PAUSED = &H1
Public Const MCONJOB_STATUS_ERROR = &H2
Public Const MCONJOB_STATUS_DELETING = &H4
Public Const MCONJOB_STATUS_SPOOLING = &H8
Public Const MCONJOB_STATUS_PRINTING = &H10
Public Const MCONJOB_STATUS_OFFLINE = &H20
Public Const MCONJOB_STATUS_PAPEROUT = &H40
Public Const MCONJOB_STATUS_PRINTED = &H80
Public Const MCONJOB_STATUS_DELETED = &H100
Public Const MCONJOB_STATUS_BLOCKED_DEVQ = &H200
Public Const MCONJOB_STATUS_USER_INTERVENTION = &H400
Public Const MCONJOB_STATUS_RESTART = &H800

Public Function CheckPrinter(ByRef strPrinterStr As String, ByRef strJobStr As String) As Boolean
    Dim hPrinter As Long
    Dim ByteBuf As Long
    Dim BytesNeeded As Long
    Dim PI2 As PRINTER_INFO_2
    Dim JI2 As JOB_INFO_2
    Dim PrinterInfo() As Byte
    Dim JobInfo() As Byte
    Dim result As Long
    Dim LastError As Long
    Dim PrinterName As String
    Dim tempStr As String
    Dim NumJI2 As Long
    Dim pDefaults As PRINTER_DEFAULTS
    Dim i As Integer
   
    
    On Error GoTo errHandle
    CheckPrinter = True
    
    PrinterName = Printer.DeviceName
    result = OpenPrinter(PrinterName, hPrinter, pDefaults)
    If result = 0 Then
        CheckPrinter = False
        ClosePrinter hPrinter
        Exit Function
    End If
    
    
    result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)

    ReDim PrinterInfo(1 To BytesNeeded)
    
    ByteBuf = BytesNeeded
    
    '获取打印机状态
    result = GetPrinter(hPrinter, 2, PrinterInfo(1), ByteBuf, _
      BytesNeeded)
    
    '检查错误
    If result = 0 Then
        '当错误发生时
        LastError = err.LastDllError()
        
        '显示错误提示
        CheckPrinter = False
        ClosePrinter hPrinter
        Exit Function
    End If

   '打印机状态字节数组的内容复制到结构PRINTER_INFO_2 structure
    CopyMemory PI2, PrinterInfo(1), Len(PI2)
    
    '调用API来获得所需的缓冲区的大小
    result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, ByVal 0&, 0&, _
       BytesNeeded, NumJI2)
    
   '检查如果没有当前工作,然后显示适当的消息
    If BytesNeeded = 0 Then
        strJobStr = "无打印任务！"
    Else
      '对打印作业Redim字节数组来保存信息
      ReDim JobInfo(0 To BytesNeeded)
      
      '调用API获取打印工作信息
      result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, JobInfo(0), _
        BytesNeeded, ByteBuf, NumJI2)
      
      '检查错误.
      If result = 0 Then
            LastError = err.LastDllError
            
            '显示错误提示
            CheckPrinter = False
            ClosePrinter hPrinter
            Exit Function
      End If
      
      
      'JOB_INFO_2 结构
        For i = 0 To NumJI2 - 1
            CopyMemory JI2, JobInfo(i * Len(JI2)), Len(JI2)
            tempStr = ""
            '检查打印机状态
            If JI2.pStatus = 0& Then
                If JI2.Status = 0 Or (JI2.Status And MCONJOB_STATUS_PRINTING) Then
                    CheckPrinter = True
                Else
                    CheckPrinter = False
                End If
            Else
              ' 不同的打印机状态
              CheckPrinter = False
            End If
        Next
    End If
   
   '关闭打印机操作
    ClosePrinter hPrinter
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    If err.LastDllError <> MCONERROR_INSUFFICIENT_BUFFER Then
       '显示错误提示
        CheckPrinter = False
        ClosePrinter hPrinter
    End If
    Call SaveErrLog
End Function




