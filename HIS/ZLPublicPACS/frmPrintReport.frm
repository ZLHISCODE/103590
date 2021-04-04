VERSION 5.00
Object = "{257A5750-6F4D-4A7A-A149-21D28B3E6EAA}#6.1#0"; "ZLPACSRICHPAGES.OCX"
Begin VB.Form frmPrintReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报告打印"
   ClientHeight    =   5445
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5610
   Icon            =   "frmPrintReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin ZLPacsRichPageScale.ZLRichPageScaleAct zlDocEditor 
      Height          =   4332
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4452
      Object.Visible         =   -1  'True
      AutoScroll      =   0   'False
      AutoSize        =   0   'False
      AxBorderStyle   =   1
      BorderWidth     =   0
      Caption         =   "ZLRichPages"
      Color           =   -16777201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      KeyPreview      =   0   'False
      PixelsPerInch   =   96
      PrintScale      =   1
      Scaled          =   -1  'True
      DropTarget      =   0   'False
      HelpFile        =   ""
      PopupMode       =   0
      ScreenSnap      =   0   'False
      SnapBuffer      =   10
      DockSite        =   0   'False
      DoubleBuffered  =   0   'False
      ParentDoubleBuffered=   0   'False
      UseDockManager  =   0   'False
      Enabled         =   -1  'True
      AlignWithMargins=   0   'False
      HMenuVisible    =   -1  'True
      VMenuVisible    =   -1  'True
      ReadOnly        =   0   'False
      Orientation     =   0
      BottomMagin     =   2.54
      BoundLeftRight  =   20
      FooterVisible   =   -1  'True
      FooterY         =   10
      MaxPageBreakHeight=   25
      MinPageBreakHeight=   5
      PageBreakHeight =   20
      PageNoFirst     =   1
      PageNoFromNumber=   1
      PageNoHAlign    =   0
      PageNoVAlign    =   0
      PageNoVisible   =   -1  'True
      PageViewMode    =   -1  'True
      RightMargin     =   3.17
      TopMargin       =   2.54
      BackgroundStyle =   3
      CtlColor        =   10070188
      IsShowHint      =   -1  'True
      TabNavigation   =   1
      NoReadOnlyJumps =   0   'False
      NoCaretHighLightJumps=   0   'False
      NoImageResize   =   0   'False
      HideReadOnlyCaret=   -1  'True
      AutoSwitchLang  =   0   'False
      WantTabs        =   -1  'True
      DoNotWantShiftReturns=   0   'False
      DoNotWantReturns=   0   'False
      CtrlJumps       =   -1  'True
      ClearTagOnStyleApp=   0   'False
      IsShowCheckPoints=   0   'False
      IsShowPageBreaks=   0   'False
      IsShowSpecialCharacters=   0   'False
      IsShowHiddenText=   0   'False
      IsShowItemHints =   0   'False
      IsDblClickSelectsWord=   -1  'True
      IsRClickDeselects=   -1  'True
      AlignPageH      =   0
      AlignPageV      =   0
      ViewMode        =   0
      ZoomMode        =   0
      EditZoomMode    =   2
      ZoomPercent     =   68
      ZoomPercentEdit =   100
      ParentCustomHint=   0   'False
      Modified        =   0   'False
      UndoLimit       =   -1
      IsMarginRectVisible=   -1  'True
      StateView       =   -1  'True
      BackGroundPicture=   "frmPrintReport.frx":000C
      BeginProperty PageNoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderVisible   =   -1  'True
      HeaderY         =   10
      TableAutoAddRow =   -1  'True
      ThumbilsVisible =   -1  'True
      SimpleViewMode  =   0   'False
      SclRVRulerVVisible=   0   'False
      SclRVRulerHVisible=   0   'False
      ScrollBarVVisible=   -1  'True
      ScrollBarHVisible=   -1  'True
      BackGroudVisible=   -1  'True
      BorderPenStyle  =   0
      Ver             =   "2.1"
      StatusBarVisible=   -1  'True
      CanEdit         =   -1  'True
      DisableCopyElement=   0   'False
      PageWidth       =   21
      PageHeight      =   29.7
      CanPopMenu      =   0   'False
      LeftMargin      =   3.17
      CanInput        =   -1  'True
      TableGridVisible=   0   'False
      CanEditHeader   =   -1  'True
      CanEditFooter   =   -1  'True
      IsRevision      =   0   'False
      RevisionTag     =   ""
      RevisionAddColor=   0
      RevisionDelColor=   0
      MaskText        =   ""
      AllowSelection  =   -1  'True
      BeginProperty MaskTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FinalShowMode   =   0   'False
      DocMasterId     =   ""
      PageSetupInPre  =   0   'False
      ServerTime      =   "1899-12-30"
      XMLEncoding     =   ""
      HScrollPos      =   68
      VScrollPos      =   0
      IsShowMargin    =   0   'False
      IsAutoPageWidth =   0   'False
   End
End
Attribute VB_Name = "frmPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjFtpInfo As tFtpInfo
Private mobjFtp As New clsFtp

Private mcnOracle As New ADODB.Connection


Private Function GetFileExt(strFile As String) As String
    Dim index As Integer
    
    GetFileExt = ""
    
    index = InStrRev(strFile, ".")
    If index <= 0 Then Exit Function
    
    GetFileExt = Mid(strFile, index)
End Function


Public Sub PrintReport(ByVal strDocId As String, Optional ByVal strPrinterName As String = "", Optional ByVal blnIsPreview As Boolean = False)
'功能：打印报告
'参数说明:
'strDocId---报告ID
'strPrinterName---打印机名称，为空时弹出打印设置框
'blnIsPreview---为True时显示预览窗口
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strContent As String
    Dim intLoop As Integer
    
    On Error GoTo ErrH
    
    If Trim(strDocId) = "" Or Trim(strDocId) = "0" Then Exit Sub
    
    '加载报告内容
    strSql = "Select Length(a.报告内容.GetClobVal()) as ContentLength,A.ID From 影像报告记录 a Where a.ID = '" & strDocId & "'"
    Set rsTemp = GetRecordset(strSql)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If rsTemp("ContentLength").Value > 2000 Then
        For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
            strSql = "select to_char(substr(a.报告内容.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                     " from 影像报告记录 a where a.ID = '" & strDocId & "'"
                     
            Set rsTemp = GetRecordset(strSql)
            
            If rsTemp.EOF = False Then
                strContent = strContent & Nvl(rsTemp("send_content").Value)
            End If
        Next
    Else
        strSql = "Select a.报告内容.getclobval() as send_content From 影像报告记录 a Where a.ID = '" & strDocId & "'"
        
        Set rsTemp = GetRecordset(strSql)
            
        If rsTemp.EOF = False Then
            strContent = Nvl(rsTemp("send_content").Value)
        End If
    End If
    
    If strContent = "" Then
        MsgBox "报告内容不存在。"
        Exit Sub
    End If
    
    zlDocEditor.IsShowMargin = True
    zlDocEditor.FinalShowMode = True
    zlDocEditor.SimpleViewMode = False
    
    zlDocEditor.ViewMode = PreviewMode
    
    Call zlDocEditor.SetZoom(100)
    
    Call LoadReportContent(strContent, strDocId)
    
    If UCase(GetFileExt(strPrinterName)) = ".PDF" Then
      '另存为pdf文件
      Call zlDocEditor.SaveAs(strPrinterName)
    Else
        If blnIsPreview Then
            Call zlDocEditor.PrintPreview(False, False, False, False, True)
        Else
            Call zlDocEditor.PrintPages(strPrinterName)
        End If
    End If
    
    Exit Sub
ErrH:
    MsgBox err.Description, vbCritical, "系统信息"
    err.Clear
End Sub

Private Sub LoadReportContent(ByVal strContent As String, ByVal strDocId As String)
    Dim strXml As String
    
    If strContent = "" Then Exit Sub
    
    '解析xml文档，将图像信息加入文档中
    strXml = AddImageInfoToXml(strContent, strDocId)
    strXml = Replace(strXml, "吠", "名")
    strXml = Replace(strXml, "朠", "服")
    strXml = Replace(strXml, "丠", "不")
    
    zlDocEditor.OpenWithXML strXml
    zlDocEditor.FinalShowMode = True
End Sub

Private Function AddImageInfoToXml(ByVal strXml As String, ByVal strDocId As String) As String
'功能：解析xml文档，将图像信息加入文档中
'返回：
    Dim objXml As New DOMDocument
       
    Dim objXmlNodes As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeAttribute As IXMLDOMNode
    Dim strImgSVG As String
    
On Error GoTo Errorhand
    
    If objXml.loadXML(strXml) = False Then
        MsgBox "报告内容加载失败！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Set objXmlNodes = objXml.selectNodes("*//image")
    
    If objXmlNodes.length <= 0 Then
        AddImageInfoToXml = strXml
        Exit Function
    End If
    
    '初始化FTP相关信息
    Call InitFtpInfo(strDocId)
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '从FTP上获取图像文件后返回图像串
            strImgSVG = GetFtpImgSVG(objXmlNodeAttribute.Text)
            
            Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("img")
            '将图像信息写入xml
            objXmlNodeAttribute.Text = strImgSVG
        End If
    Next
    
    AddImageInfoToXml = objXml.xml
    
    Exit Function
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

'从FTP上获取SVG格式图像
Private Function GetFtpImgSVG(ByVal strKey As String, Optional ByRef strMsg As String = "") As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    
    If strKey = "" Then Exit Function
    
    strLocalFileName = Replace(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir & strKey, "/", "\")
    strVirtualPath = Replace(mobjFtpInfo.FtpDir & mobjFtpInfo.SubDir, "\", "/")
    
    '创建本地路径
    If Not objFSO.FolderExists(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir) Then
        Call MkLocalDir(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir)
    End If
    
    '如果本地存在则删除
    If objFSO.FileExists(strLocalFileName) Then Call objFSO.DeleteFile(strLocalFileName, True)
    
    '连接FTP
    If ConnFtp() = False Then
        strMsg = "FTP不能正常连接，请检查网络设置。"
        Exit Function
    End If
    
    If mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, objFSO.GetFileName(strLocalFileName)) <> 0 Then
        strMsg = "图像内容从FTP服务器上获取失败！"
        Exit Function
    End If
    
    '下载后读取
    GetFtpImgSVG = GetFileContent(strLocalFileName)
End Function

Private Function ConnFtp() As Boolean
    If mobjFtp.hConnection = 0 Then
        '连接FTP存储设备
        If mobjFtp.FuncFtpConnect(mobjFtpInfo.FtpIP, mobjFtpInfo.FTPUser, mobjFtpInfo.FtpPswd) = 0 Then
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

Private Function InitFtpInfo(ByVal strDocId As String) As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 'ReportImages/' || to_Char(b.创建时间,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
            "a.设备号, a.FTP用户名, a.FTP密码, a.IP地址,'/'||a.Ftp目录||'/' As Root " & _
            "From 影像设备目录 a, 影像报告记录 b where a.设备号 = b.设备号 And b.id = [1]"
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", strDocId)
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInfo.FtpDir = Nvl(rsTmp("Root"))
    mobjFtpInfo.FtpIP = Nvl(rsTmp("IP地址"))
    mobjFtpInfo.FtpPswd = Nvl(rsTmp("FTP密码"))
    mobjFtpInfo.FTPUser = Nvl(rsTmp("FTP用户名"))
    mobjFtpInfo.DiviceId = Trim(Nvl(rsTmp("设备号")))
    
    mobjFtpInfo.SubDir = Nvl(rsTmp("URL"))
    mobjFtpInfo.DestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInfo = True
End Function
