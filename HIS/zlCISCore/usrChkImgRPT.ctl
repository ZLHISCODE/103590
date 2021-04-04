VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Begin VB.UserControl usrChkImgRPT 
   BackColor       =   &H80000009&
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   ScaleHeight     =   5055
   ScaleWidth      =   7230
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   4755
      Width           =   7125
   End
   Begin VB.PictureBox picLable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   0
      ScaleHeight     =   1830
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   2925
      Width           =   255
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "可选报告图"
         ForeColor       =   &H8000000E&
         Height          =   1005
         Left            =   45
         TabIndex        =   4
         Top             =   420
         Width           =   195
      End
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7095
      _Version        =   262146
      _ExtentX        =   12515
      _ExtentY        =   4524
      _StockProps     =   35
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":0000
            Key             =   "one"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":015A
            Key             =   "two"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":02B4
            Key             =   "four"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":040E
            Key             =   "pic"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":07A8
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":0B42
            Key             =   "n1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":10DC
            Key             =   "n2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":1676
            Key             =   "n3"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":1C10
            Key             =   "n4"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":21AA
            Key             =   "n0"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":2744
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrChkImgRPT.ctx":295E
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   635
      ButtonWidth     =   1349
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "1× 1"
            Key             =   "one"
            Object.ToolTipText     =   "单张显示"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "1× 2"
            Key             =   "two"
            Object.ToolTipText     =   "两张显示"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "2× 2"
            Key             =   "four"
            Object.ToolTipText     =   "四张显示"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "选择"
            Key             =   "Select"
            Object.ToolTipText     =   "在当前位置加入新的报告图像"
            ImageKey        =   "pic"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "添加影像"
            Key             =   "Append"
            Object.ToolTipText     =   "在最后添加新的报告图像"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "清除"
            Key             =   "Clear"
            Object.ToolTipText     =   "清除当前的报告图像"
            ImageKey        =   "clear"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "图1"
            Key             =   "Index"
            Object.ToolTipText     =   "显示当前图的顺序号"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全清"
            Key             =   "ClearAll"
            Object.ToolTipText     =   "清除所有选择的报告图"
            ImageKey        =   "ClearAll"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刷新"
            Key             =   "Refresh"
            Object.ToolTipText     =   "刷新可选报告图"
            ImageKey        =   "Refresh"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin DicomObjects.DicomViewer DicomImages 
      Height          =   1830
      Left            =   255
      TabIndex        =   2
      Top             =   2925
      Width           =   6870
      _Version        =   262146
      _ExtentX        =   12118
      _ExtentY        =   3228
      _StockProps     =   35
      BackColor       =   0
   End
End
Attribute VB_Name = "usrChkImgRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDispMode As Boolean
Private mImageStep As Single
Private mlng病历id As Long
Private mblnMoved As Boolean '数据是否已转出

Private mblnDown As Boolean
Private sinOldX As Single, sinOldY As Single
Private lngCurIndex As Long
Private iCurImageIndex As Integer, iCurReportImage As Integer

Private AdviceID As Long, lngSendNO As Long
Private Inet1 As New clsFtp
Private Inet2 As New clsFtp
Private strDeviceNO1 As String
Private strDeviceNO2 As String
Private strVirtualPath As String
Private strLocalCachPath As String

Event ImageClick(ByVal Index As Byte)
Event ImageMouseDown(ByVal Index As Byte, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub SetDiagItem(ByVal lngAdviceID As Long, ByVal SendNO As Long)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage
    Dim i As Integer, iNum As Integer
    Dim iRows As Integer, iCols As Integer
    
    On Error GoTo ErrHandle
    AdviceID = lngAdviceID: lngSendNO = SendNO
    
    strSQL = "Select Count(*)" & _
        " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
        " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
        " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
        " And B.名称 Not Like '标记%'" & _
        " And E.应用场合=Decode(D.主页ID,Null,1,2) And D.ID=[1]" & _
        " Order BY Trunc(Y/567),Trunc(X/567)"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lngAdviceID)
    
    iNum = IIf(rsTmp(0) > 1, rsTmp(0), 4)
    With DViewer
        .Images.Clear
        
        ResizeRegion iNum, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
        For i = 1 To iNum
            Set curImage = .Images.AddNew
            With curImage
                .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
            End With
        Next
        iCurImageIndex = 1
        .Images(iCurImageIndex).BorderColour = vbRed
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowMe(ByVal lng病历ID As Long)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage
    Dim iRows As Integer, iCols As Integer
    Dim MaxImages As Integer, i As Integer, iNum As Integer, aImages() As String, aReportImages() As String, iSelectNum As Integer
    Dim strTempPath As String, lngBuffSize As Long, strHost As String
    Dim strImages As String
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    
    On Error GoTo ErrHandle
    UserControl.MousePointer = vbHourglass
    
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    '获取选中的报告图
    strSQL = "Select C.医嘱ID,C.发送号," & _
        "E.用户名 As User1,E.密码 As Pwd1," & _
        "E.IP地址 As Host1,e.设备号 as 设备号1," & _
        "'/'||E.Ftp目录||'/' As Root1,Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
        "||D.检查UID As URL1," & _
        "F.用户名 As User2,F.密码 As Pwd2," & _
        "F.IP地址 As Host2," & _
        "'/'||F.Ftp目录||'/' As Root2,Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
        "||D.检查UID As URL2,A.图象文件 As URL,f.设备号 as 设备号2  " & _
        "From 病人病历外部图 A,病人病历内容 B,病人医嘱发送 C,影像检查记录 D,影像设备目录 E,影像设备目录 F " & _
        "Where A.病历ID=B.ID And B.病历记录ID=C.报告ID(+) And C.医嘱ID=D.医嘱ID And C.发送号=D.发送号 And D.位置一=E.设备号(+) And D.位置二=F.设备号(+) " & _
        "And A.病历id = [1] Order By A.序号"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
        strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lng病历ID)
    
    iSelectNum = rsTmp.RecordCount
    With DViewer
        MaxImages = .Images.Count
        
        .Images.Clear
        If rsTmp.RecordCount > 0 Then
            AdviceID = rsTmp("医嘱ID"): lngSendNO = rsTmp("发送号")
'            Inet.strIPAddress = NVL(rsTmp("Host1")): Inet.strUser = NVL(rsTmp("User1")): Inet.strPsw = NVL(rsTmp("Pwd1"))
            ReDim aImages(rsTmp.RecordCount - 1)
            i = 0: iNum = UBound(aImages, 1)
            Do While Not rsTmp.EOF
                aImages(i) = rsTmp("URL"): i = i + 1
                
                rsTmp.MoveNext
            Loop
            rsTmp.MoveFirst
        End If
    End With
    '获取所有报告图
    iNum = ShowReportImages(aReportImages)
    If iSelectNum = 0 And iNum > 0 Then aImages = aReportImages: iSelectNum = UBound(aImages, 1) + 1
    
    With DViewer
        iNum = iSelectNum - 1
        ResizeRegion IIf(MaxImages > 0, MaxImages, IIf(iNum = -1, 1, iNum + 1)), _
            .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
        iCurImageIndex = 0
        strDeviceNO1 = ""
        strDeviceNO2 = ""

        For i = 0 To iNum
            If .Images.Count >= MaxImages And MaxImages > 0 Then Exit For
                
            Set curImage = New DicomImage
            strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & strLocalCachPath
            strTmpFile = Replace(strTmpFile, "/", "\")
            MkLocalDir strTmpFile
            strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(Trim(aImages(i)))
            
            If Len(LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 And _
                InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aImages(i))))) > 0 Then
                curImage.FileImport strTmpFile, objFileSystem.GetExtensionName(Trim(aImages(i)))
                .Images.Add curImage: Set curImage = .Images(.Images.Count)
            Else
                Set curImage = .Images.ReadFile(strTmpFile)
            End If
    
            With curImage
                .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                .ShowLabels = True: .Tag = Trim(aImages(i))
            End With
        Next
        
        For i = .Images.Count + 1 To MaxImages
            Set curImage = .Images.AddNew
            With curImage
                .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
            End With
        Next
            
        If .Images.Count > 0 Then
            iCurImageIndex = 1
            .Images(iCurImageIndex).BorderColour = vbRed
        End If
    End With
    
    UserControl.MousePointer = vbDefault
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    UserControl.MousePointer = vbDefault
    Call SaveErrLog
End Sub

Private Function ShowReportImages(aReportImages() As String) As Integer
'返回可选报告图数量
'aReportImages：返回可选报告图文件名数组
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim curImage As DicomImage
    Dim iRows As Integer, iCols As Integer
    Dim MaxImages As Integer, i As Integer, iNum As Integer
    Dim strTempPath As String, lngBuffSize As Long, strHost As String
    Dim strImages As String
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    
    On Error GoTo ErrHandle
    UserControl.MousePointer = vbHourglass
    
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    
    '获取所有报告图
    strSQL = "Select 报告图象," & _
        "D.用户名 As User1,D.密码 As Pwd1," & _
        "D.IP地址 As Host1," & _
        "'/'||D.Ftp目录||'/' As Root1,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL1,d.设备号 as 设备号1," & _
        "E.用户名 As User2,E.密码 As Pwd2," & _
        "E.IP地址 As Host2," & _
        "'/'||E.Ftp目录||'/' As Root2,Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')" & _
        "||C.检查UID As URL2,e.设备号 as 设备号2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) " & _
        "And C.医嘱ID=[1] And C.发送号=[2]"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    If mblnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "检查报告", AdviceID, lngSendNO)
    If rsTmp.EOF Then
        iNum = -1
    Else
        strImages = Trim(Split(NVL(rsTmp(0), " "), "|")(0))
        aReportImages = Split(strImages, ";"): iNum = UBound(aReportImages, 1)
'                Inet.strIPAddress = NVL(rsTmp("Host1")): Inet.strUser = NVL(rsTmp("User1")): Inet.strPsw = NVL(rsTmp("Pwd1"))
    End If
    If iNum > -1 Then
        strLocalCachPath = IIf(IsNull(rsTmp("URL1")), "", rsTmp("URL1"))
        With DicomImages
            .Images.Clear
            
            .MultiColumns = CInt(.Width / .Height): .MultiRows = 1
            
            strDeviceNO1 = ""
            strDeviceNO2 = ""
    
            For i = 0 To iNum
                Set curImage = New DicomImage
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & rsTmp("URL1")
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(Trim(aReportImages(i)))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("User1")), NVL(rsTmp("Pwd1"))
                    End If
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("User2")), NVL(rsTmp("Pwd2"))
                    End If
                    strHost = "ftp://" & NVL(rsTmp("User1")) & IIf(IsNull(rsTmp("Pwd1")), "", ":" & rsTmp("Pwd1")) & _
                        "@" & NVL(rsTmp("Host1"))
                    strVirtualPath = objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1") & "/" & aReportImages(i))
                    If Inet1.FuncDownloadFile(strVirtualPath, strTmpFile, Trim(aReportImages(i))) <> 0 Then
                        strHost = "ftp://" & NVL(rsTmp("User2")) & IIf(IsNull(rsTmp("Pwd2")), "", ":" & rsTmp("Pwd2")) & _
                            "@" & NVL(rsTmp("Host2"))
                        strVirtualPath = objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2") & "/" & aReportImages(i))
                        Call Inet2.FuncDownloadFile(strVirtualPath, strTmpFile, Trim(aReportImages(i)))
                    End If
                End If
                If Len(LCase(objFileSystem.GetExtensionName(Trim(aReportImages(i))))) > 0 And _
                    InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(Trim(aReportImages(i))))) > 0 Then
                    curImage.FileImport strTmpFile, objFileSystem.GetExtensionName(Trim(aReportImages(i)))
                    .Images.Add curImage: Set curImage = .Images(.Images.Count)
                Else
                    Set curImage = .Images.ReadFile(strTmpFile)
                End If
        
                With curImage
                    .BorderStyle = 6: .BorderWidth = 1: .BorderColour = vbWhite
                    .ShowLabels = True: .Tag = strHost & "," & strVirtualPath & "/" & Trim(aReportImages(i))
                End With
            Next
            
            .CurrentIndex = 1
            iCurReportImage = 1
            .Images(iCurReportImage).BorderColour = vbRed
            If .MultiColumns < .Images.Count Then
                HScroll.Enabled = True
            Else
                HScroll.Enabled = False
            End If
        End With
        SetScrollBar
    End If
    ShowReportImages = iNum + 1
    
    UserControl.MousePointer = vbDefault
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    UserControl.MousePointer = vbDefault
    Call SaveErrLog
End Function

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
'保存检验结果项目数据
    Dim strType As String, strPath As String, strFile As String
    Dim strTmp As String
    Dim lngPic As Long
    Dim i As Integer, objFileSystem As New Scripting.FileSystemObject
    
    strTmp = ""
    With DViewer
        For i = 1 To .Images.Count
            If .Images(i).Tag <> "" Then
                strType = "": strPath = ""
                strFile = objFileSystem.GetFileName(.Images(i).Tag)
                strTmp = strTmp & i & "\" & strType & "\" & strPath & "\" & strFile & "\"
            End If
        Next
'        If strTmp = "" Then strError = "无任何图像可以保存！": Exit Function
        '病历ID_IN
        '图像串_IN
        If strTmp <> "" Then strReturnSQL = "zl_病人病历外部图_检查(" & lng病历ID & ",'" & strTmp & "')"
        SaveData = True
    End With
End Function

Public Property Get Text() As String
'为每一个控件加上文本转储属性
    Text = ""
End Property

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get DispMode() As Boolean
    '是否为显示模式
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    Dim ObjButton As MSComctlLib.Button
    mDispMode = New_DispMode
    tb1.Visible = Not mDispMode: UserControl_Resize
    PropertyChanged "DispMode"
End Property

Public Property Get ID病人病历() As Long
'返回病人病历ID
    ID病人病历 = mlng病历id
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
'设置病人病历ID,并检查该病历是不是存在
    mlng病历id = New_ID病人病历
    ShowMe mlng病历id
End Property

Public Function GetPicture(Index As Byte) As Picture
'得到指定索引的图片里的图像
End Function

Public Sub SetPicture(Index As Byte, picTmp As StdPicture)
'根据指定索引图片设置图像
End Sub

Private Sub DicomImages_DblClick()
    Dim i As Integer
    
    If DicomImages.Images.Count = 0 Then Exit Sub
        
    tb1_ButtonClick tb1.Buttons("Select")
    
    On Error Resume Next
    With DViewer
        If iCurImageIndex >= .Images.Count Then Exit Sub
        i = iCurImageIndex + 1
        .Images(iCurImageIndex).BorderColour = vbWhite
        .Images(i).BorderColour = vbRed
        iCurImageIndex = i
    End With
End Sub

Private Sub DViewer_DblClick()
    tb1_ButtonClick tb1.Buttons("Clear")
End Sub

Private Sub HScroll_Change()
    On Error Resume Next
    DicomImages.CurrentIndex = HScroll.Value
    DicomImages.SetFocus
End Sub

Private Sub tb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strURL As String, astrFileInfo() As String
    Dim strTempPath As String, lngBuffSize As Long
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim curImage As DicomImage, i As Integer, iNum As Integer
    
    On Error GoTo ErrHandle
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    
    Select Case Button.Key
        Case "Select"
            If DicomImages.Images.Count = 0 Then Exit Sub
            
            strURL = DicomImages.Images(iCurReportImage).Tag
            If Len(strURL) > 0 Then
                astrFileInfo = Split(strURL, ",")
                With DViewer
                    If iCurImageIndex > 0 Then .Images.Remove iCurImageIndex
'                    If Len(LCase(objFileSystem.GetExtensionName(strURL))) > 0 And _
'                        InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(strURL))) > 0 Then
                
                        Set curImage = New DicomImage
'                        strTmpFile = strTempPath & objFileSystem.GetFolder(Trim(strURL))
                        strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & strLocalCachPath
                        strTmpFile = Replace(strTmpFile, "/", "\")
                        MkLocalDir strTmpFile
                        strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(Trim(strURL))
                        If Dir(strTmpFile, vbDirectory) = "" Then
                        
                            If Inet1.FuncDownloadFile(strVirtualPath, strTmpFile, objFileSystem.GetFileName(Trim(strURL))) <> 0 Then
                                Call Inet2.FuncDownloadFile(strVirtualPath, strTmpFile, objFileSystem.GetFileName(Trim(strURL)))
                            End If
                        End If
                        If Len(LCase(objFileSystem.GetExtensionName(strURL))) > 0 And _
                            InStr("bmp;jpg;jpeg;gif;ico", LCase(objFileSystem.GetExtensionName(strURL))) > 0 Then
                            curImage.FileImport strTmpFile, objFileSystem.GetExtensionName(strURL)
'                            objFileSystem.DeleteFile strTmpFile, True
                            .Images.Add curImage
                        Else
                            .Images.ReadFile strTmpFile
                        End If
'                    Else
'                        .Images.ReadURL Inet.URL & strURL
'                    End If
                    With .Images(.Images.Count)
                        .BorderStyle = 6: .BorderWidth = 1: .Tag = astrFileInfo(0) & strURL
                    End With
                    If iCurImageIndex > 0 Then
                        .Images.Move .Images.Count, iCurImageIndex
                    Else
                        iCurImageIndex = 1
                    End If
                    .Images(iCurImageIndex).BorderColour = vbRed
                End With
                UserControl_Resize
            End If
        Case "Clear"
            If iCurImageIndex > 0 Then
                With DViewer
                    .Images.Remove (iCurImageIndex)
                    .Images.AddNew
                    With .Images(.Images.Count)
                        .BorderStyle = 6: .BorderWidth = 1
                    End With
                    .Images.Move .Images.Count, iCurImageIndex
                    .Images(iCurImageIndex).BorderColour = vbRed
                End With
                UserControl_Resize
            End If
        Case "ClearAll"
            With DViewer
                iNum = .Images.Count
                .Images.Clear
                For i = 1 To iNum
                    .Images.AddNew
                    .Images(.Images.Count).BorderStyle = 6: .Images(.Images.Count).BorderWidth = 1
                Next
                iCurImageIndex = 1
                .Images(iCurImageIndex).BorderColour = vbRed
            End With
        Case "Refresh"
            Call ShowReportImages(astrFileInfo)
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    On Error Resume Next
    With DViewer
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And i <> iCurImageIndex Then
            .Images(iCurImageIndex).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurImageIndex = i
        End If
    End With
End Sub

Private Sub UserControl_InitProperties()
    mImageStep = Screen.TwipsPerPixelY * 1
    mDispMode = False
    mblnMoved = False
    lngCurIndex = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mImageStep = PropBag.ReadProperty("ImageStep", Screen.TwipsPerPixelY * 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
    mDispMode = PropBag.ReadProperty("DispMode", False)
    mblnMoved = PropBag.ReadProperty("DataMoved", False)
    lngCurIndex = PropBag.ReadProperty("ImageIndex", 0)
End Sub

Private Sub UserControl_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With tb1
        .Left = 0: .Top = 0
        .Width = UserControl.ScaleWidth: .Height = UserControl.ScaleHeight
    End With
    With DicomImages
        .Top = UserControl.ScaleHeight - HScroll.Height - .Height
        .Width = UserControl.ScaleWidth - .Left
        
        .MultiColumns = CInt(.Width / .Height): .MultiRows = 1
        If .MultiColumns < .Images.Count Then
            HScroll.Enabled = True
        Else
            HScroll.Enabled = False
        End If
    End With
    With picLable
        .Top = DicomImages.Top
    End With
    With HScroll
        .Top = UserControl.ScaleHeight - .Height
        .Width = UserControl.ScaleWidth - .Left
    End With
    With DViewer
        .Left = 0: .Top = IIf(mDispMode, 0, tb1.Top + tb1.Height)
        .Width = UserControl.ScaleWidth: .Height = DicomImages.Top - .Top
        
        If .Images.Count > 0 Then
            ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
            .MultiColumns = iCols: .MultiRows = iRows
        End If
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, False)
    Call PropBag.WriteProperty("DataMoved", mblnMoved, False)
    Call PropBag.WriteProperty("ImageStep", mImageStep, Screen.TwipsPerPixelY * 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
    Call PropBag.WriteProperty("ImageIndex", lngCurIndex, 0)
End Sub
 
Private Sub UserControl_EnterFocus()
    On Error Resume Next
    UserControl.Parent.CallBack_GotFocus
End Sub

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

'数据是否转出
Public Property Get DataMoved() As Boolean
    DataMoved = mblnMoved
End Property

Public Property Let DataMoved(ByVal vNewValue As Boolean)
    mblnMoved = vNewValue
End Property

Private Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Sub DicomImages_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    If Button <> 1 Then Exit Sub
    
    With DicomImages
        i = .ImageIndex(x, y)
        If i > 0 And i <= .Images.Count And iCurReportImage > 0 And i <> iCurReportImage Then
            .Images(iCurReportImage).BorderColour = vbWhite
            .Images(i).BorderColour = vbRed
            iCurReportImage = i
        End If
    End With
End Sub

Private Sub SetScrollBar()
    With HScroll
        .Min = 1: .Max = DicomImages.Images.Count: .SmallChange = 1: .LargeChange = 1
    End With
End Sub
