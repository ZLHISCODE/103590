VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmXMLValidator 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2580
      Index           =   2
      Left            =   1110
      ScaleHeight     =   2580
      ScaleWidth      =   5025
      TabIndex        =   4
      Top             =   4680
      Width           =   5025
      Begin VB.TextBox txtResult 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   1470
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   285
         Width           =   3900
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2250
      Index           =   0
      Left            =   465
      ScaleHeight     =   2250
      ScaleWidth      =   4110
      TabIndex        =   2
      Top             =   1590
      Width           =   4110
      Begin RichTextLib.RichTextBox txtXSD 
         Height          =   1875
         Left            =   150
         TabIndex        =   3
         Top             =   135
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3307
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmXMLValidator.frx":0000
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   5190
      ScaleHeight     =   2535
      ScaleWidth      =   4845
      TabIndex        =   0
      Top             =   1140
      Width           =   4845
      Begin RichTextLib.RichTextBox txtXML 
         Height          =   1875
         Left            =   225
         TabIndex        =   1
         Top             =   420
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3307
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmXMLValidator.frx":009D
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   3195
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Bindings        =   "frmXMLValidator.frx":013A
      Left            =   1620
      Top             =   345
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmXMLValidator.frx":014E
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   675
      Top             =   330
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmXMLValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrXmlSchemaFile As String
Private mobjFso As New FileSystemObject
Private mobjTextStream As TextStream

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "XML格式(*.xsd)"
    objPane.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "XML内容(*.xml)"
    objPane.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        
    Set objPane = dkpMain.CreatePane(3, 100, 100, DockBottomOf, Nothing)
    objPane.Title = "验证结果"
    objPane.Options = PaneNoCloseable + PaneNoFloatable
    
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
    
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
'    Dim objLbl As CommandBarControlCustom
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = ImageManager1.Icons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
   
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
            
    
'    Set mobjCaption = NewToolBar(objBar, xtpControlLabel, 0, "修改", True)
'    mobjCaption.Caption = "验证合法"
'    cbsMain.RecalcLayout
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 1, "打开格式")
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 2, "打开内容")
    objControl.IconId = 1
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 3, "验证格式", True)
        
        
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, 3
    End With
    
    Exit Function
    
errHand:
'    If zlComLib.ErrCenter = 1 Then
'        Resume
'    End If
'
End Function


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strFile As String
    Dim objFile As TextStream
    Dim strResult As String
    
    Select Case Control.id
    Case 1
        strFile = OpenDialog(cmdlg, "请选择XML格式", "XML格式(*.xsd)|*.xsd")
        If strFile <> "" Then
            If mobjFso.FileExists(strFile) Then
                Set objFile = mobjFso.OpenTextFile(strFile, ForReading)
                txtXSD.Text = objFile.ReadAll
                Set objFile = Nothing
            End If
        End If
    Case 2
        strFile = OpenDialog(cmdlg, "请选择XML内容", "XML内容(*.xml)|*.xml")
        If strFile <> "" Then
            If mobjFso.FileExists(strFile) Then
                Set objFile = mobjFso.OpenTextFile(strFile, ForReading)
                txtXML.Text = objFile.ReadAll
                Set objFile = Nothing
            End If
        End If
    Case 3
        
        If txtXSD.Text = "" Then
            MsgBox "请输入或从文件中选择XML格式内容(*.xsd)"
            Exit Sub
        End If
        
        If txtXML.Text = "" Then
            MsgBox "请输入或从文件中选择XML内容(*.xml)"
            Exit Sub
        End If
        
        
        If ValidXML(txtXSD.Text, txtXML.Text, strResult) Then

            
        End If
        
        If strResult = "" Then
            txtResult.Text = "验证合法！"
            txtResult.ForeColor = 0
        Else
            txtResult.Text = strResult
            txtResult.ForeColor = 255
        End If
            
    Case 4
        Unload Me
    End Select
    
End Sub

Private Function ValidXML(ByVal strXmlXsd As String, ByVal strXmlContent As String, Optional ByRef strErrorReason As String) As Boolean
    '******************************************************************************************************************
    '功能：校验XML格式
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFile As String
    Dim objXmlSchema As MSXML2.XMLSchemaCache60
    Dim objXmlMessage As MSXML2.DOMDocument60
    
    On Error GoTo errHand
    
    mstrXmlSchemaFile = gstrDataPath & "\zlXMLValidator.xsd"
    Set mobjTextStream = mobjFso.CreateTextFile(mstrXmlSchemaFile, True)
    mobjTextStream.Write strXmlXsd
    mobjTextStream.Close
    
    Set objXmlSchema = New MSXML2.XMLSchemaCache60
    objXmlSchema.Add "", mstrXmlSchemaFile

    Set objXmlMessage = New MSXML2.DOMDocument60
    objXmlMessage.async = False
    objXmlMessage.validateOnParse = True
    objXmlMessage.resolveExternals = False
    Set objXmlMessage.schemas = objXmlSchema

    strFile = gstrDataPath & "\zlXMLValidator.xml"

    strXmlContent = Replace(LCase(strXmlContent), "encoding=""utf-8""?", "encoding=""gbk""?")

    Set mobjTextStream = mobjFso.CreateTextFile(strFile, True)
    mobjTextStream.Write strXmlContent
    mobjTextStream.Close

    Call objXmlMessage.Load(strFile)
    
    Call objXmlMessage.Validate
    If objXmlMessage.parseError.errorCode <> 0 Then
        strErrorReason = objXmlMessage.parseError.reason
        ValidXML = False
    Else
        ValidXML = True
    End If
    
    Exit Function
    
errHand:
    MsgBox Err.Description
    
'    Call mobjFso.DeleteFile(mstrXmlSchemaFile, True)
'    Call mobjFso.DeleteFile(strFile, True)
    
End Function


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picBack(0).hWnd
    Case 2
        Item.Handle = picBack(1).hWnd
        
    Case 3
        Item.Handle = picBack(2).hWnd
        
    End Select
End Sub

Private Sub Form_Load()
        
    Call InitCommandBar
    Call InitDockPannel
    
    If mobjFso.FileExists(gstrDataPath & "\zlXMLValidator.xsd") Then
        Set mobjTextStream = mobjFso.OpenTextFile(gstrDataPath & "\zlXMLValidator.xsd", ForReading)
        txtXSD.Text = mobjTextStream.ReadAll
        mobjTextStream.Close
    End If
    
    If mobjFso.FileExists(gstrDataPath & "\zlXMLValidator.xml") Then
        Set mobjTextStream = mobjFso.OpenTextFile(gstrDataPath & "\zlXMLValidator.xml", ForReading)
        txtXML.Text = mobjTextStream.ReadAll
        mobjTextStream.Close
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 3, 15, 80, Me.ScaleWidth, 80)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjFso = Nothing
    Set mobjTextStream = Nothing
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        txtXSD.Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 1
        txtXML.Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 2
        txtResult.Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
End Sub


