VERSION 5.00
Object = "{257A5750-6F4D-4A7A-A149-21D28B3E6EAA}#6.1#0"; "ZLPacsRichPages.ocx"
Begin VB.Form frmRichReportV2 
   Caption         =   "�����ĵ��༭��"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9510
   Icon            =   "frmRichReportV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9510
   StartUpPosition =   3  '����ȱʡ
   Begin ZLPacsRichPageScale.ZLRichPageScaleAct zlDocEditor 
      Height          =   4335
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Object.Visible         =   -1  'True
      AutoScroll      =   0   'False
      AutoSize        =   0   'False
      AxBorderStyle   =   1
      BorderWidth     =   0
      Caption         =   "ZLRichPages"
      Color           =   -16777201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Modified        =   -1  'True
      UndoLimit       =   -1
      IsMarginRectVisible=   -1  'True
      StateView       =   -1  'True
      BackGroundPicture=   "frmRichReportV2.frx":000C
      BeginProperty PageNoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
      HScrollPos      =   69
      VScrollPos      =   0
      IsShowMargin    =   0   'False
      IsAutoPageWidth =   0   'False
   End
   Begin VB.Label labInfo 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "ע����ǰ��������Ϊ�����ĵ��༭����������鿴�����ܱ༭��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmRichReportV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tFtpInfo
    FtpDir As String
    FtpIp As String
    FtpPswd As String
    FtpUser As String
    DiviceId As String
    
    SubDir As String
    DestMainDir As String
End Type

Private mlngAdviceId As Long
Private mblnMoved As Boolean
Private mcnOledb As ADODB.Connection
Private mObjNotify As IEventNotify
Private mlngModuleNo As Long
Private mlngDeptID As Long
Private mstrPrivs As String
Private mstrXmlVer As String

Private mobjFtpInfo As tFtpInfo
Private mobjFtp As New clsFtp

Public Sub zlRefresh(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsForceRefresh As Boolean, Optional ByVal strSpecifyDocId As String = "")
 
    If mlngAdviceId = lngAdviceId And blnIsForceRefresh = False Then Exit Sub
 
    mlngAdviceId = lngAdviceId
    mblnMoved = blnMoved
    
    
    Call PrintReport(True, strSpecifyDocId)
End Sub

Public Sub zlInit(objNotify As IEventNotify, ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String)
    
    Set mObjNotify = objNotify
    
    mlngModuleNo = lngModuleNo
    mlngDeptID = lngDeptId
    mstrPrivs = strPrivs
    mstrXmlVer = GetXMLVersion
    
    Call InitOledbConn
    Call InitReportEditor
End Sub


Private Function GetXMLVersion() As String
'��ȡxml��Ӧ��֧�ְ汾
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo errHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
        For intLoop = 0 To UBound(varXMLVersion)
            err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        
    On Error GoTo errHand
        
    If strXMLVer = "" Then
        MsgBox "����MSXML2.DOMDocument����ʧ��", vbInformation, "��ʾ"
        Exit Function
    End If
 
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description, vbInformation, "��ʾ"
End Function


Public Function InitOledbConn() As Boolean
    Dim objRegister As Object
    Dim strError As String
    
On Error GoTo err
    InitOledbConn = False
    
    If mcnOledb Is Nothing Then
        Set objRegister = VBA.Interaction.GetObject("", "zlRegister.clsRegister")
        Set mcnOledb = objRegister.ReGetConnection(1, strError)
    End If
    
    InitOledbConn = True
    Exit Function
err:
    InitOledbConn = False
    MsgBoxD Me, "��ʼ�����ݿ������쳣: & " & err.Description & "��", vbOKOnly, "��ʾ"
End Function


Private Function GetRecordset(ByVal strSQL As String) As ADODB.Recordset
On Error GoTo errHand
    Set GetRecordset = New ADODB.Recordset
    
    If mcnOledb Is Nothing Then Exit Function
    
    If GetRecordset.State = adStateOpen Then GetRecordset.Close
    '��
    GetRecordset.Open strSQL, mcnOledb, adOpenKeyset, adLockOptimistic

    Exit Function
errHand:
    If err <> 0 Then
        MsgBoxD Me, "��������" & err.Description, vbInformation, "��ʾ"
    End If
End Function

Public Sub PrintReport(Optional ByVal blnIsPreview As Boolean = False, Optional ByVal strSpecifyDocId As String = "")
'���ܣ���ӡ����
'����˵��:
'strDocId---����ID
'strPrinterName---��ӡ�����ƣ�Ϊ��ʱ������ӡ���ÿ�
'blnIsPreview---ΪTrueʱ��ʾԤ������
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strContent As String
    Dim intLoop As Integer
    Dim strDocId As String
    
    On Error GoTo errH
    
    Call zlDocEditor.ClearAll
    
    If mlngAdviceId = 0 Then Exit Sub
    If mcnOledb Is Nothing Then Exit Sub
    
    If strSpecifyDocId = "" Then
        strSQL = "select RAWTOHEX(��鱨��id) as ����ID from ����ҽ������ where ҽ��ID=[1]"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�ĵ��༭������ID", mlngAdviceId)
        If rsTemp.RecordCount <= 0 Then
            Exit Sub
        End If
        
        strDocId = nvl(rsTemp!����ID)
    Else
        strDocId = strSpecifyDocId
    End If
    
    If strDocId = "" Then
'        zlDocEditor.ClearAll
        Exit Sub
    End If
    
    '���ر�������
    strSQL = "Select Length(a.��������.GetClobVal()) as ContentLength,A.ID From Ӱ�񱨸��¼ a Where a.ID = '" & strDocId & "'"
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ�񱨸��¼", "HӰ�񱨸��¼")
    End If
    
    Set rsTemp = GetRecordset(strSQL)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If rsTemp("ContentLength").value > 2000 Then
        For intLoop = 1 To rsTemp("ContentLength").value / 2000 + 1
            strSQL = "select to_char(substr(a.��������.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                     " from Ӱ�񱨸��¼ a where a.ID = '" & strDocId & "'"
                     
            If mblnMoved Then
                strSQL = Replace(strSQL, "Ӱ�񱨸��¼", "HӰ�񱨸��¼")
            End If
            
            Set rsTemp = GetRecordset(strSQL)
            
            If rsTemp.EOF = False Then
                strContent = strContent & nvl(rsTemp("send_content").value)
            End If
        Next
    Else
        strSQL = "Select a.��������.getclobval() as send_content From Ӱ�񱨸��¼ a Where a.ID = '" & strDocId & "'"
        
        Set rsTemp = GetRecordset(strSQL)
            
        If rsTemp.EOF = False Then
            strContent = nvl(rsTemp("send_content").value)
        End If
    End If
    
    If strContent = "" Then
        MsgBoxD Me, "�������ݲ����ڡ�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
'    zlDocEditor.IsShowMargin = False
'    zlDocEditor.FinalShowMode = True
'    zlDocEditor.SimpleViewMode = True
    
'    zlDocEditor.ViewMode = PreviewMode
    
    Call LoadReportContent(strContent, strDocId)
    
'    Call zlDocEditor.SetZoom(100)
    zlDocEditor.VScrollPos = 0
    
    Exit Sub
errH:
    MsgBoxD Me, err.Description, vbCritical, "��ʾ"
    err.Clear
End Sub

Public Sub PrintPreview(Optional ByVal blnIsPreview As Boolean = True)
    If blnIsPreview Then
        Call zlDocEditor.PrintPreview(False, False, False, False, True)
    Else
        Call zlDocEditor.PrintPages
    End If
End Sub

Private Sub LoadReportContent(ByVal strContent As String, ByVal strDocId As String)
    Dim strXml As String
    
    If strContent = "" Then Exit Sub
    
    '����xml�ĵ�����ͼ����Ϣ�����ĵ���
    strXml = AddImageInfoToXml(strContent, strDocId)
    strXml = Replace(strXml, "��", "��")
    strXml = Replace(strXml, "�P", "��")
    strXml = Replace(strXml, "�H", "��")
    
    zlDocEditor.OpenWithXML strXml
'    zlDocEditor.FinalShowMode = True
End Sub

Private Function AddImageInfoToXml(ByVal strXml As String, ByVal strDocId As String) As String
'���ܣ�����xml�ĵ�����ͼ����Ϣ�����ĵ���
'���أ�
    Dim objXML As Object ' New DOMDocument
       
    Dim objXmlNodes As Object 'IXMLDOMNodeList
    Dim objXmlNode As Object 'IXMLDOMNode
    Dim objXmlNodeAttribute As Object 'IXMLDOMNode
    Dim strImgSVG As String
    
On Error GoTo ErrorHand

    Set objXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVer)
    If objXML Is Nothing Then
        MsgBoxD Me, "ʵ�������� MSXML2.DOMDocument" & mstrXmlVer & "ʧ�ܡ�", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    If objXML.loadXML(strXml) = False Then
        MsgBoxD Me, "�������ݼ���ʧ�ܣ�", vbExclamation, "��ʾ"
        Exit Function
    End If
    
    Set objXmlNodes = objXML.selectNodes("*//image")
    
    If objXmlNodes.Length <= 0 Then
        AddImageInfoToXml = strXml
        Exit Function
    End If
    
    '��ʼ��FTP�����Ϣ
    Call InitFtpInfo(strDocId)
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '��FTP�ϻ�ȡͼ���ļ��󷵻�ͼ��
            strImgSVG = GetFtpImgSVG(objXmlNodeAttribute.Text)
            
            Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("img")
            '��ͼ����Ϣд��xml
            objXmlNodeAttribute.Text = strImgSVG
        End If
    Next
    
    AddImageInfoToXml = objXML.XML
    
    Exit Function
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

'��FTP�ϻ�ȡSVG��ʽͼ��
Private Function GetFtpImgSVG(ByVal strKey As String, Optional ByRef strMsg As String = "") As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    
    If strKey = "" Then Exit Function
    
    strLocalFileName = Replace(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir & strKey, "/", "\")
    strVirtualPath = Replace(mobjFtpInfo.FtpDir & mobjFtpInfo.SubDir, "\", "/")
    
    '��������·��
    If Not objFSO.FolderExists(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir) Then
        Call MkLocalDir(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir)
    End If
    
    '������ش�����ɾ��
    If objFSO.FileExists(strLocalFileName) Then Call objFSO.DeleteFile(strLocalFileName, True)
    
    '����FTP
    If ConnFtp() = False Then
        strMsg = "FTP�����������ӣ������������á�"
        Exit Function
    End If
    
    If mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, objFSO.GetFileName(strLocalFileName)) <> 0 Then
        strMsg = "ͼ�����ݴ�FTP�������ϻ�ȡʧ�ܣ�"
        Exit Function
    End If
    
    '���غ��ȡ
    GetFtpImgSVG = GetFileContent(strLocalFileName)
End Function

Private Function GetFileContent(ByVal strFileName As String) As String
'��ȡ�����ļ�����
    Dim i As Integer, strContent As String, bty() As Byte
    
    If Dir(strFileName) = "" Then Exit Function
    
    i = FreeFile
    
    ReDim bty(FileLen(strFileName) - 1)
    
    Open strFileName For Binary As #i
    Get #i, , bty
    Close #i
    strContent = StrConv(bty, vbUnicode)
    
    GetFileContent = strContent
End Function

Private Function ConnFtp() As Boolean
    If mobjFtp.hConnection = 0 Then
        '����FTP�洢�豸
        If mobjFtp.FuncFtpConnect(mobjFtpInfo.FtpIp, mobjFtpInfo.FtpUser, mobjFtpInfo.FtpPswd) = 0 Then
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

Private Function InitFtpInfo(ByVal strDocId As String) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "Select 'ReportImages/' || to_Char(b.����ʱ��,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
            "a.�豸��, a.FTP�û���, a.FTP����, a.IP��ַ,'/'||a.FtpĿ¼||'/' As Root " & _
            "From Ӱ���豸Ŀ¼ a, Ӱ�񱨸��¼ b where a.�豸�� = b.�豸�� And b.id = [1]"
            
    If mblnMoved Then
        strSQL = Replace(strSQL, "Ӱ�񱨸��¼", "HӰ�񱨸��¼")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡFTP��Ϣ", strDocId)
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInfo.FtpDir = nvl(rsTmp("Root"))
    mobjFtpInfo.FtpIp = nvl(rsTmp("IP��ַ"))
    mobjFtpInfo.FtpPswd = nvl(rsTmp("FTP����"))
    mobjFtpInfo.FtpUser = nvl(rsTmp("FTP�û���"))
    mobjFtpInfo.DiviceId = Trim(nvl(rsTmp("�豸��")))
    
    mobjFtpInfo.SubDir = nvl(rsTmp("URL"))
    mobjFtpInfo.DestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInfo = True
End Function

Private Sub InitReportEditor()
    zlDocEditor.FooterVisible = False
    zlDocEditor.HeaderVisible = False
    zlDocEditor.HMenuVisible = False
    zlDocEditor.PageNoVisible = False
    zlDocEditor.ThumbilsVisible = False
    zlDocEditor.VMenuVisible = False
    zlDocEditor.ZoomPercent = 100
    zlDocEditor.CanEdit = False
    zlDocEditor.CanInput = False
    zlDocEditor.TableGridVisible = False
    zlDocEditor.InitOCX hwnd
    zlDocEditor.IsShowMargin = False
    zlDocEditor.FinalShowMode = True
    zlDocEditor.ViewMode = PreviewMode
    zlDocEditor.SimpleViewMode = True
End Sub

Private Sub Form_Resize()
On Error GoTo errhandle
    labInfo.Left = 0
    labInfo.Top = 0
    labInfo.Width = Me.ScaleWidth
    
    zlDocEditor.Left = 0
    zlDocEditor.Top = labInfo.Height
    zlDocEditor.Width = Me.ScaleWidth
    zlDocEditor.Height = Me.ScaleHeight - labInfo.Top
Exit Sub
errhandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mObjNotify = Nothing
    Set mobjFtp = Nothing
'    Set mobjStudyInfo = Nothing
    Set mcnOledb = Nothing
End Sub
