VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "dicomobjects.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDelRptImg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "删除报告图像"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   Icon            =   "frmDelRptImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   3
      Top             =   6600
      Width           =   1100
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   600
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4080
      TabIndex        =   2
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   1680
      TabIndex        =   1
      Top             =   6600
      Width           =   1100
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _Version        =   262147
      _ExtentX        =   16113
      _ExtentY        =   11033
      _StockProps     =   35
      BackColor       =   -2147483640
   End
End
Attribute VB_Name = "frmDelRptImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As frmViewer
Public pSeriesUID As String      '记录当前报告图像所关联的图像的序列UID，
                                '因为图像的序列UID在数据库中没有被修改，所以使用序列UID来查找报告图像比检查UID查找更加准确
Private iCurImageIndex As Integer
Private bDelete As Boolean
Private aImages() As String
Private aDelRecord As String
Private strRepImgField As String      '保存数据库中原来的报告图像字段值
Private strURL As String                '保存FTP网络路径

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    If iCurImageIndex > 0 Then
        aDelRecord = IIf(aDelRecord = vbNullString, " ", aDelRecord & ",") & Me.DViewer.Images(iCurImageIndex).Tag
        Me.DViewer.Images.Remove iCurImageIndex
        iCurImageIndex = iCurImageIndex - 1
        If iCurImageIndex <= 0 And Me.DViewer.Images.Count > 0 Then
            iCurImageIndex = 1
        End If
        If iCurImageIndex > 0 Then
            Me.DViewer.Images(iCurImageIndex).BorderColour = vbRed
            Me.DViewer.Refresh
        End If
        bDelete = True
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim aDelImages() As String
    Dim aRptImg As String       '记录报告图像文件名串
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim aDelNum() As String
    
    If bDelete Then '处理删除结果
        aDelNum = Split(aDelRecord, ",")
        
        If UBound(aDelNum) < 0 Then
            Unload Me
            Exit Sub
        End If
        
        ReDim aDelImages(UBound(aDelNum, 1)) As String
        For i = 0 To UBound(aDelNum, 1)
            aDelImages(i) = aImages(aDelNum(i))
            aImages(aDelNum(i)) = "DEL"
        Next
        For i = 0 To UBound(aImages, 1) '组合报告图像文件名串
            If aImages(i) <> "DEL" Then
                aRptImg = IIf(aRptImg = vbNullString, " ", aRptImg & ";") & aImages(i)
            End If
        Next
        '将报告图像保存到数据库中
        If InStr(strRepImgField, "|") <> 0 Then     '有录音报告
            aRptImg = aRptImg & Right(strRepImgField, Len(strRepImgField) - InStr(strRepImgField, "|") + 1)
        End If
                 
        strSQL = "ZL_影像检查报告_UPDATE('" & PstrCheckUID & "','" & aRptImg & "')"
        zlDatabase.ExecuteProcedure strSQL, App.ProductName
        
        '删除FTP中被删掉的报告图像
        For i = 0 To UBound(aDelImages, 1)
            Inet.Execute , "DELETE " & strURL & Trim(aDelImages(i))
            Do While Inet.StillExecuting
                DoEvents
            Loop
        Next
    End If
    Unload Me
End Sub


Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DViewer
        i = .ImageIndex(x, y)
        If i <> iCurImageIndex Then
            If iCurImageIndex > 0 And iCurImageIndex <= .Images.Count Then
                .Images(iCurImageIndex).BorderColour = vbWhite
            End If
            If i > 0 And i <= .Images.Count Then
                .Images(i).BorderColour = vbRed
                iCurImageIndex = i
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
     ShowImages pSeriesUID
     aDelRecord = vbNullString
End Sub

Public Sub ShowImages(SeriesUID As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double, lngRecID As Long
    Dim curImage As DicomImage
    Dim iRows As Integer, iCols As Integer
    Dim i As Integer, iNum As Integer
    Dim aTemp() As String
    
    Dim strTempPath As String, lngBuffSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    
    If gcnOracle Is Nothing Then Exit Sub
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    '现根据序列UID，查找对应的检查UID
    strSQL = "select 检查UID FROM 影像检查序列  where 序列UID =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SeriesUID)
    If rsTmp.RecordCount = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "该图像没有保存到网络中，没有对应的报告图。", vbInformation, gstrSysName
        Me.Hide
        Exit Sub
    End If
    PstrCheckUID = Nvl(rsTmp!检查UID)
    
    '先查设备一，没有再查设备二
    strSQL = "Select 报告图象," & _
        "'ftp://'||Decode(FTP用户名,Null,'',FTP用户名||Decode(FTP密码,Null,'',':'||FTP密码))" & _
        "||'@'||IP地址 As Host,'/'||Ftp目录||'/'" & _
        "||Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID||'/' As URL " & _
        "From 影像检查记录 C,影像设备目录 D " & _
        "Where C.位置一=D.设备号  " & _
        "And C.检查UID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, PstrCheckUID)
    If rsTmp.RecordCount = 0 Then
        '查设备二
        strSQL = "Select 报告图象," & _
            "'ftp://'||Decode(FTP用户名,Null,'',FTP用户名||Decode(FTP密码,Null,'',':'||FTP密码))" & _
            "||'@'||IP地址 As Host,'/'||Ftp目录||'/'" & _
            "||Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
            "||C.检查UID||'/' As URL " & _
            "From 影像检查记录 C,影像设备目录 D " & _
            "Where C.位置二=D.设备号  " & _
            "And C.检查UID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, PstrCheckUID)
    End If
    Screen.MousePointer = vbHourglass
    
    iCurImageIndex = 0
    With DViewer
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            If Len(Nvl(rsTmp(0))) = 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "没有报告图像，请在观片站中保存报告图！", vbInformation, gstrSysName
                Me.Hide
                Exit Sub
            End If
            
            strRepImgField = Nvl(rsTmp(0))
            If InStr(strRepImgField, "|") <> 0 Then      '有录音报告
                aTemp = Split(Nvl(rsTmp(0)), "|")
                If UBound(aTemp, 1) > 0 Then
                    aImages = Split(aTemp(0), ";")
                End If
            Else
                aImages = Split(Nvl(rsTmp(0)), ";")
            End If
            iNum = UBound(aImages, 1)
            If iNum = -1 Then
                Screen.MousePointer = vbDefault
                MsgBox "没有报告图像，请进入观片站设置！", vbInformation, gstrSysName
                Me.Hide
                Exit Sub
            End If
            Inet.AccessType = icUseDefault: Inet.URL = rsTmp("Host")
            strURL = rsTmp("URL")
            
            ResizeRegion iNum + 1, .width, .height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows

            lngRecID = 1
            For i = 0 To iNum
                If Len(LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 And _
                    InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 Then
                
                    Set curImage = New DicomImage
                    strTmpFile = strTempPath & objFileSystem.GetFileName(Trim(aImages(i)))
                    Inet.Execute , "Get " & rsTmp("URL") & Trim(aImages(i)) & " " & strTmpFile
                    Do While Inet.StillExecuting
                        DoEvents
                    Loop
                    If Dir(strTmpFile) <> "" Then
                        curImage.FileImport strTmpFile, ""
                        objFileSystem.DeleteFile strTmpFile, True
                        .Images.Add curImage: Set curImage = .Images(.Images.Count)
                    End If
                Else
                    Set curImage = .Images.ReadURL(Inet.URL & rsTmp("URL") & Trim(aImages(i)))
                End If
                If Not curImage Is Nothing Then
                    With curImage
                        .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                        .ShowLabels = True: .Tag = i
                    End With
                End If
            Next
            If .Images.Count > 0 Then
                iCurImageIndex = 1
                .Images(iCurImageIndex).BorderColour = vbRed
            End If
        Else
            Screen.MousePointer = vbDefault
            MsgBox "该检查未进行，没有报告图像。", vbInformation, gstrSysName
            Me.Hide
            Exit Sub
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Screen.MousePointer = vbDefault
    Call SaveErrLog
End Sub

