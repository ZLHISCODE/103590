VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "ZLSUBCLASS.OCX"
Begin VB.UserControl Document 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   ScaleHeight     =   1860
   ScaleWidth      =   2430
   ToolboxBitmap   =   "Document.ctx":0000
   Begin zlSubclass.Subclass Subclass2 
      Left            =   1845
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin zlSubclass.Subclass Subclass1 
      Left            =   1485
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin RichTextLib.RichTextBox rtbThis 
      Height          =   1545
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2725
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Document.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblThis 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3930
   End
End
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##模 块 名：Document.ctl
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：普通视图的基础编辑控件。
'##版    本：
'######################################################################################

Option Explicit

'#############################################################################################################
'##     局部变量
'#############################################################################################################

Private m_hWndRTB As Long           'RTB的 hWnd
Private m_hWnd As Long              '控件的 hWnd
Private m_hWndParent  As Long       '父窗体的 hWnd

Private m_bSubClassing As Boolean   '是否继承子类
Private m_TOM As New cTextDocument  'TOM 3.0 模型，核心对象

'#############################################################################################################
'##     独立属性
'#############################################################################################################

Private mvarAutoDetectURL As Boolean
Private mvarBackColor As OLE_COLOR
Private mvarBorder As Boolean
Private mvarDefaultTabStop As Single
Private mvarDoDefaultURLClick As Boolean
Private mvarEnabled As Boolean
Private mvarFileName As String
Private mvarFoot As String
Private mvarForceEdit As Boolean
Private mvarHead As String
Private mvarMarginBottom As Long
Private mvarMarginLeft As Long
Private mvarMarginRight As Long
Private mvarMarginTop As Long
Private mvarModified As Boolean
Private mvarPaperColor As OLE_COLOR
Private mvarPaperHeight As Long
Private mvarPaperWidth As Long
Private mvarPaperKind As PaperKindEnum          '纸张类型属性
Private mvarPaperOrient As PaperOrientEnum      '纸张方向
Private mvarPicture As StdPicture
Private mvarReadOnly As Boolean
Private mvarTitle As String
Private mvarTransparent As Boolean
Private mvarViewMode As ViewModeEnum
Private mvarZoomFactor As Double
Private mvarWYSIWYG   As Boolean                '是否启用所见即所得
Private mvarAuditMode As Boolean                '审核模式

'#############################################################################################################
'##     事件声明
'#############################################################################################################

Public Event Change()       '内容改变！
Public Event Focuse()       '获取输入焦点！
Public Event MouseWheel(bBackDirection As Boolean, Shift As Integer, X As Single, Y As Single, Value As Single)    '鼠标滚轮事件
Public Event Zoom(NewFactor As Double)    '用户通过Ctrl＋鼠标来改变了缩放比例！
Public Event Resize()    '控件尺寸改变
Public Event RequestLine()              '请求行数改变
Public Event SelChange(ByVal lStart As Long, ByVal lEnd As Long)   '选择区域改变
Public Event LinkEvent(ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)      '链接事件
Public Event ModifyProtected(ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)            '试图编辑受保护区域
Public Event BeforeKeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event RequestRightMenu(ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event Click()        '单击
Public Event DblClick()     '双击
Public Event PressTabKey()  '按下Tab按钮
Public Event GetDelCharColor(ByRef COLOR As OLE_COLOR)     '获取删除字符的颜色
Public Event GetNewCharColor(ByRef COLOR As OLE_COLOR)     '获取新增字符的颜色
Public Event IsDelCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsDelCharColor As Boolean)   '判断是否是删除字符的颜色
Public Event IsNewCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsNewCharColor As Boolean)   '判断是否是新增字符的颜色

'#############################################################################################################
'##     公共属性
'#############################################################################################################

Public Property Get OriginRTB() As Object
    Set OriginRTB = rtbThis
End Property

Public Property Let AuditMode(ByVal vData As Boolean)
    mvarAuditMode = vData
    PropertyChanged "AuditMode"
End Property

Public Property Get AuditMode() As Boolean
    AuditMode = mvarAuditMode
End Property

Public Property Get WYSIWYG() As Boolean
    WYSIWYG = mvarWYSIWYG
End Property

Public Property Let WYSIWYG(ByVal vData As Boolean)
    mvarWYSIWYG = vData
    PropertyChanged "WYSIWYG"
End Property

Public Property Get AutoDetectURL() As Boolean
    AutoDetectURL = mvarAutoDetectURL
End Property

Public Property Let AutoDetectURL(ByVal vData As Boolean)
    mvarAutoDetectURL = vData
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.Freeze
        SendMessageLong m_hWndRTB, EM_AUTOURLDETECT, Abs(vData), 0
        m_TOM.TextDocument.UnFreeze
    End If
    PropertyChanged "AutoDetectURL"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    mvarBackColor = oColor
    PropertyChanged "BackColor"
End Property

Public Property Get Border() As Boolean
    Border = mvarBorder
End Property

Public Property Let Border(ByVal vData As Boolean)
    Dim dwStyle As Long
    Dim dwExStyle As Long

    If m_hWndRTB <> 0 Then
        ' Make sure that the RichEdit never has a border:
        dwStyle = GetWindowLong(m_hWndRTB, GWL_STYLE)
        dwExStyle = GetWindowLong(m_hWndRTB, GWL_EXSTYLE)
        dwStyle = dwStyle And Not ES_SUNKEN
        dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
        SetWindowLong m_hWndRTB, GWL_STYLE, dwStyle
        SetWindowLong m_hWndRTB, GWL_EXSTYLE, dwExStyle
        pStyleChanged
    End If
    UserControl.BorderStyle() = Abs(vData)
    
    mvarBorder = vData
    PropertyChanged "Border"
End Property

Public Property Get CanCopy() As Boolean
    CanCopy = (m_TOM.TextDocument.Selection.End > m_TOM.TextDocument.Selection.Start)
End Property

Public Property Get CanPaste() As Boolean
   CanPaste = SendMessageLong(m_hWndRTB, EM_CANPASTE, 0, 0)
End Property

Public Property Get CanRedo() As Boolean
    CanRedo = SendMessageLong(m_hWndRTB, EM_CANREDO, 0, 0)
End Property

Public Property Get CanUndo() As Boolean
   CanUndo = SendMessageLong(m_hWndRTB, EM_CANUNDO, 0, 0)
End Property

Public Property Get CanDelete() As Boolean
    CanDelete = (m_TOM.TextDocument.Selection.End > m_TOM.TextDocument.Selection.Start)
End Property

Public Property Get CurrentColumn() As Long
    Dim pt As POINTAPI
    pt = GetCurPos
    CurrentColumn = pt.X
End Property

Public Property Get CurrentLine() As Long
'    CurrentLine = SendMessageLong(m_hWndRTB, EM_EXLINEFROMCHAR, 0, m_TOM.TextDocument.Selection.Start) + 1
    Dim L  As Long
    L = SendMessage(m_hWndRTB, EM_LINEINDEX, -1, 0)
    CurrentLine = SendMessage(m_hWndRTB, EM_LINEFROMCHAR, L, 0) + 1
End Property

Public Property Get DefaultTabStop() As Single
    DefaultTabStop = mvarDefaultTabStop
End Property

Public Property Let DefaultTabStop(ByVal vData As Single)
    mvarDefaultTabStop = vData
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.Freeze
        m_TOM.TextDocument.DefaultTabStop = vData
        m_TOM.TextDocument.UnFreeze
    End If
    PropertyChanged "DefaultTabStop"
End Property

Public Property Get DoDefaultURLClick() As Boolean
    DoDefaultURLClick = mvarDoDefaultURLClick
End Property

Public Property Let DoDefaultURLClick(ByVal vData As Boolean)
    mvarDoDefaultURLClick = vData
    PropertyChanged "DoDefaultURLClick"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
    UserControl.Enabled = vData
    If Not m_hWndRTB = 0 Then
        m_TOM.TextDocument.Freeze
        EnableWindow m_hWndRTB, Abs(vData)
        m_TOM.TextDocument.UnFreeze
    End If
    PropertyChanged "Enabled"
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
'    mvarTitle = Mid(vData, InStrRev(vData, "\") + 1)
    PropertyChanged "FileName"
End Property

Public Property Get FirstVisibleLine() As Long
   FirstVisibleLine = SendMessageLong(m_hWndRTB, EM_GETFIRSTVISIBLELINE, 0, 0)
End Property

Public Property Get ForceEdit() As Boolean
    ForceEdit = mvarForceEdit
End Property

Public Property Let ForceEdit(ByVal vData As Boolean)
    mvarForceEdit = vData
    PropertyChanged "ForceEdit"
End Property

Public Property Get Head() As String
    Head = mvarHead
End Property

Public Property Let Head(ByVal vData As String)
    mvarHead = vData
    PropertyChanged "Head"
End Property

Public Property Get Hwnd() As Long
   Hwnd = UserControl.Hwnd
End Property

Public Property Get hWndRTB() As Long
   hWndRTB = rtbThis.Hwnd
End Property

Public Property Get LineCount() As Long
   LineCount = SendMessageLong(m_hWndRTB, EM_GETLINECOUNT, 0, 0)
End Property

Public Property Get MarginBottom() As Long
    MarginBottom = mvarMarginBottom
End Property

Public Property Let MarginBottom(vData As Long)
    mvarMarginBottom = vData
    PropertyChanged "MarginBottom"
End Property

Public Property Get MarginLeft() As Long
    MarginLeft = mvarMarginLeft
End Property

Public Property Let MarginLeft(vData As Long)
    mvarMarginLeft = vData
    PropertyChanged "MarginLeft"
End Property

Public Property Get MarginRight() As Long
    MarginRight = mvarMarginRight
End Property

Public Property Let MarginRight(vData As Long)
    mvarMarginRight = vData
    PropertyChanged "MarginRight"
End Property

Public Property Get MarginTop() As Long
    MarginTop = mvarMarginTop
End Property

Public Property Let MarginTop(vData As Long)
    mvarMarginTop = vData
    PropertyChanged "MarginTop"
End Property

Public Property Get Modified() As Boolean
   If (m_hWndRTB <> 0) Then
      Modified = (SendMessageLong(m_hWndRTB, EM_GETMODIFY, 0, 0) <> 0)
   End If
End Property

Public Property Let Modified(ByVal bModified As Boolean)
   If (m_hWndRTB <> 0) Then
      SendMessageLong m_hWndRTB, EM_SETMODIFY, Abs(bModified), 0
   End If
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = rtbThis.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtbThis.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = rtbThis.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    rtbThis.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get PaperColor() As OLE_COLOR
    PaperColor = mvarPaperColor
End Property

Public Property Let PaperColor(vData As OLE_COLOR)
    mvarPaperColor = vData
    If (m_hWndRTB <> 0) Then
       SendMessageLong m_hWndRTB, EM_SETBKGNDCOLOR, 0, TranslateColor(vData)
    End If
    PropertyChanged "PaperColor"
End Property

Public Property Get PaperHeight() As Long
    PaperHeight = mvarPaperHeight
End Property

Public Property Let PaperHeight(vData As Long)
    mvarPaperHeight = vData
    ResetWYSIWYG
    PropertyChanged "PaperHeight"
End Property

Public Property Get PaperWidth() As Long
    PaperWidth = mvarPaperWidth
End Property

Public Property Let PaperWidth(vData As Long)
    mvarPaperWidth = vData
    ResetWYSIWYG
    PropertyChanged "PaperWidth"
End Property

Public Property Let PaperKind(ByVal vData As PaperKindEnum)
    mvarPaperKind = vData
    ResetWYSIWYG
    PropertyChanged "PaperKind"
End Property

Public Property Get PaperKind() As PaperKindEnum
    PaperKind = mvarPaperKind
End Property

Public Property Let PaperOrient(ByVal vData As PaperOrientEnum)
    mvarPaperOrient = vData
    ResetWYSIWYG
    PropertyChanged "PaperOrient"
End Property

Public Property Get PaperOrient() As PaperOrientEnum
    PaperOrient = mvarPaperOrient
End Property

Public Property Get ReadOnly() As Boolean
    Dim lStyle As Long
    If (m_hWndRTB <> 0) Then
        lStyle = GetWindowLong(m_hWndRTB, GWL_STYLE)
        If (lStyle And ES_READONLY) = ES_READONLY Then
            ReadOnly = True
        End If
    Else
        ReadOnly = mvarReadOnly
    End If
End Property

Public Property Let ReadOnly(ByVal vData As Boolean)
    mvarReadOnly = vData
    If m_hWndRTB <> 0 Then
        SendMessageLong m_hWndRTB, EM_SETREADONLY, Abs(vData), 0
    End If
    PropertyChanged "ReadOnly"
End Property

Public Property Get SelLength() As Long
    SelLength = rtbThis.SelLength
End Property

Public Property Let SelLength(vData As Long)
    rtbThis.SelLength = vData
End Property

Public Property Get SelRTF() As String
    SelRTF = rtbThis.SelRTF
End Property

Public Property Let SelRTF(vData As String)
    On Error Resume Next
    rtbThis.SelRTF = vData
    Err.Clear
End Property

Public Property Get SelStart() As Long
    SelStart = rtbThis.SelStart
End Property

Public Property Let SelStart(vData As Long)
    rtbThis.SelStart = vData
End Property

Public Property Get SelText() As String
    If m_hWndRTB <> 0 Then
        SelText = m_TOM.TextDocument.Selection
    Else
        SelText = rtbThis.SelText
    End If
End Property

Public Property Let SelText(vData As String)
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.Selection = vData
    Else
        rtbThis.SelText = vData
    End If
End Property

Public Property Get Text() As String
    Text = rtbThis.Text
End Property

Public Property Let Text(ByRef vData As String)
    rtbThis.Text = vData
End Property

Public Property Get TextRTF() As String
    TextRTF = rtbThis.TextRTF
End Property

Public Property Let TextRTF(ByRef vData As String)
    rtbThis.TextRTF = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
    PropertyChanged "Title"
End Property

Public Property Get TOM() As cTextDocument
    Set TOM = m_TOM
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    mvarTransparent = vData
    PropertyChanged "Transparent"
End Property

Public Property Get ViewMode() As ViewModeEnum
    ViewMode = mvarViewMode
End Property

Public Property Let ViewMode(ByVal vData As ViewModeEnum)
    If m_hWndRTB <> 0 Then m_TOM.TextDocument.Freeze
    If vData = cprPaper Then vData = cprNormal
    mvarViewMode = vData
    Call UserControl_Show
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.UnFreeze
        If rtbThis.Visible And rtbThis.Enabled Then
            rtbThis.SetFocus
        End If
        If Ambient.UserMode Then
            Range(0, Len(rtbThis.Text)).Para.WidowControl = True
        End If
    End If
    PropertyChanged "ViewMode"
End Property

Public Property Get ZoomFactor() As Double
    Dim lngA As Long, lngB As Long, lngValue As Double
    If m_hWndRTB <> 0 Then
        SendMessageRef m_hWndRTB, EM_GETZOOM, lngA, lngB
        If lngB = 0 Then
            lngValue = 1
        Else
            lngValue = Abs(lngA / lngB)
        End If
        mvarZoomFactor = lngValue     '保存该值
    End If
    ZoomFactor = mvarZoomFactor
End Property

Public Property Let ZoomFactor(ByVal vData As Double)
    Dim lVal As Long
    mvarZoomFactor = vData
    lVal = Abs(Round(vData * 100))
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.Freeze
        SendMessageLong m_hWndRTB, EM_SETZOOM, lVal, 100
        m_TOM.TextDocument.UnFreeze
    End If
    Call ResetWYSIWYG
    PropertyChanged "ZoomFactor"
End Property

'#############################################################################################################
'##     公共方法
'#############################################################################################################

Public Sub CopyWithFormat()
    '带格式复制
    Dim lngr As Long
    
    lngr = OpenClipboard(m_hWndRTB)
    If lngr = 0 Then '不能打开粘贴板
        DoEvents '加上这一句，奇怪，360没再作怪了
        Clipboard.Clear '防止以下Copy内容为空，但粘贴板中有内容  ///360拦截、搜狗云、远程桌面 会报“不能打开粘贴板” '用API替代
    Else
        Call EmptyClipboard
        Call CloseClipboard
    End If
    SendMessageLong m_hWndRTB, WM_COPY, 0, 0
End Sub

Public Sub PasteWithFormat()
    '带格式粘贴
    SendMessageLong m_hWndRTB, WM_PASTE, 0, 0
End Sub

Public Sub Copy()
    '过滤掉内嵌关键字
    Dim strtmp As String, i As Long, lS As Long, lE As Long, j As Long
    lS = Selection.StartPos
    lE = Selection.EndPos
    strtmp = Space(lE - lS)
    For i = lS To lE - 1
        If Range(i, i + 1).Font.Hidden = False Then
            j = j + 1
            Mid(strtmp, j, 1) = Range(i, i + 1).Text
        End If
    Next
    
    Dim lngr As Long
    
    lngr = OpenClipboard(m_hWndRTB)
    If lngr = 0 Then '不能打开粘贴板
        DoEvents '加上这一句，奇怪，360没再作怪了
        Clipboard.Clear '防止以下Copy内容为空，但粘贴板中有内容  ///360拦截、搜狗云、远程桌面 会报“不能打开粘贴板” '用API替代
    Else
        Call EmptyClipboard
        Call CloseClipboard
    End If
    Clipboard.SetText Left(strtmp, j)
End Sub

Public Sub Cut()
    Call CloseClipboard
    SendMessageLong m_hWndRTB, WM_CUT, 0, 0
End Sub

Public Sub Delete()
    Selection.Delete
End Sub

Public Function FindText(sText As String, Optional ByVal iFlag As Long) As Boolean
    '功能：从文档当前位置向后查找指定字符串，查到则选中
    '参数：
    '   sText,要查找的字符串
    '   iFlag,匹配方式,默认为0(不区分大小写、全半角)，可以为以下标志的组合：
    '       tomMatchCase,2-大小写匹配
    '       tomMatchWord,4-完全匹配
    '       实际测试，尚不支持模式匹配等
    On Error Resume Next
    Dim i As Long, j As Long, r As Long
    i = m_TOM.TextDocument.Selection.End
    j = Len(rtbThis.Text) + 2 '在控件末尾有个VbCrLf，通过Text取不到
    m_TOM.TextDocument.Freeze
    m_TOM.TextDocument.Range(i, j).Select
    r = m_TOM.TextDocument.Selection.FindText(sText, tomForward, iFlag)
    If r > 0 And sText <> "" Then
        rtbThis.SelLength = Len(sText)
        FindText = True
    Else
        rtbThis.SelLength = 0
        FindText = False
    End If
    m_TOM.TextDocument.UnFreeze
End Function

Public Sub Freeze()
'用途: 禁止刷新编辑器
    m_TOM.TextDocument.Freeze
End Sub

Public Function GetLineString(lLine As Long) As String
'获取指定行字符串。
    Dim str5(255) As Byte '如果您的字串 > 255 byte请自行增加该Byte Array
    Dim str6 As String, i As Long
    
    str5(0) = 255 '字串的前两个Byte存该字串的最大长度
    str5(1) = 0
    i = SendMessage(m_hWndRTB, EM_GETLINE, lLine - 1, str5(0))
    If i = 0 Then
       GetLineString = ""
    Else
       str6 = StrConv(str5, vbUnicode)
       GetLineString = Left(str6, InStr(1, str6, Chr(0)) - 1)
    End If
End Function

Public Sub InsertOLEObject()
'插入OLE对象
    Dim UIInsertObj As OleUIInsertObjectType
    Dim retValue   As Long
    Dim lpolestr   As Long
    Dim strSize   As Long
    Dim ProgId   As String
    
    On Error GoTo Err
    UIInsertObj.cbStruct = LenB(UIInsertObj)
    UIInsertObj.dwFlags = IOF_SELECTCREATENEW
    UIInsertObj.hWndOwner = Me.Hwnd
    UIInsertObj.lpszFile = String(256, "  ")
    UIInsertObj.cchFile = Len(UIInsertObj.lpszFile)
    retValue = OleUIInsertObject(UIInsertObj)
    If (retValue = OLEUI_OK) Then
        If ((UIInsertObj.dwFlags And IOF_SELECTCREATENEW) = _
                            IOF_SELECTCREATENEW) Then
            retValue = ProgIDFromCLSID(UIInsertObj.clsid, lpolestr)
            strSize = lstrlenW(lpolestr) + 1
            ProgId = String(strSize, 0)
            CopyMemory ByVal StrPtr(ProgId), ByVal lpolestr, strSize * 2
            CoTaskMemFree lpolestr
            rtbThis.OLEObjects.Add , , "", ProgId
        Else    '  If  we  select  to  insert  from  file
            rtbThis.OLEObjects.Add , , UIInsertObj.lpszFile
        End If
    End If
    Exit Sub
Err:
    MsgBox Err.Description
End Sub

Public Sub NewDoc()
    m_TOM.TextDocument.Freeze
    m_TOM.TextDocument.New
    rtbThis.FileName = ""
    m_TOM.TextDocument.DefaultTabStop = mvarDefaultTabStop
    m_TOM.TextDocument.Selection.Para.ClearAllTabs
    m_TOM.TextDocument.UnFreeze
    FileName = ""
'    Title = "未命名文档"
    Modified = False
End Sub

Public Sub OpenDoc(Optional strFile As String = "")
    If strFile <> "" Then FileName = strFile
'    m_TOM.TextDocument.Freeze
    m_TOM.TextDocument.New
    m_TOM.TextDocument.DefaultTabStop = mvarDefaultTabStop
    m_TOM.TextDocument.Selection.Para.ClearAllTabs
    rtbThis.FileName = strFile
'    m_TOM.TextDocument.UnFreeze
    Modified = False
End Sub

Public Sub Paste()
    If Me.Selection.Font.Protected Or Me.Selection.Font.Hidden Then Exit Sub
    If Me.AuditMode Then
        Me.SelStart = Me.SelStart + Me.SelLength
        Me.SelLength = 0
        '设置当前位置为新增文本（红色）
        Me.ForceEdit = True
        On Error Resume Next
        rtbThis.SelColor = GetNewCharColor(tomAutoColor)    '新增文本
        rtbThis.SelStrikeThru = False   '去掉删除线
        Me.ForceEdit = False
        
    End If
    rtbThis.SelText = Clipboard.GetText
    Err.Clear
End Sub

Public Function Range(lStart As Long, lEnd As Long) As cRange
    Dim cR As New cRange
    cR.Init m_TOM, lStart, lEnd, mvarReadOnly
    Set Range = cR
End Function

Public Sub Redo()
    m_TOM.TextDocument.Redo 1
End Sub

Public Sub SaveDoc(Optional strFile As String = "")
    If strFile <> "" Then FileName = strFile
    rtbThis.SaveFile CStr(mvarFileName), rtfRTF
    Modified = False
End Sub

Public Sub SelectAll()
    SetSelection m_hWndRTB, 0, -1
'    Range(0, Len(rtbThis)).Selected
End Sub

Public Function Selection() As cSelection
    Dim cS As New cSelection
    cS.Init m_TOM, mvarReadOnly
    Set Selection = cS
End Function

Public Sub Undo()
    m_TOM.TextDocument.Undo 1
End Sub

Public Sub UnFreeze()
    m_TOM.TextDocument.UnFreeze
End Sub

Public Sub ResetWYSIWYG()
    '调整所见即所得
    If ViewMode = cprNormal And WYSIWYG Then Call WYSIWYG_RTF(rtbThis, MarginLeft * mvarZoomFactor, MarginRight * mvarZoomFactor, MarginTop * mvarZoomFactor, MarginBottom * mvarZoomFactor, PaperWidth * mvarZoomFactor, PaperHeight * mvarZoomFactor)
End Sub

'#############################################################################################################
'##     局部方法
'#############################################################################################################

Private Function GetCurPos() As POINTAPI
''取得光标所在的行和列
    Dim LineIndex As Long
    Dim SelRange As CHARRANGE
    Dim TempStr As String
    Dim TempArray() As Byte
    Dim CurRow As Long
    Dim CurPos As POINTAPI

    TempArray = StrConv(rtbThis.Text, vbFromUnicode)

    ''取得当前被选中文本的位置 适用于 RichTextBox
    ''rtbThis 用 EM_GETSEL 消息
    Call SendMessage(m_hWndRTB, EM_EXGETSEL, 0, SelRange)

    ''根据参数wParam指定的字符位置返回该字符所在的行号
    CurRow = SendMessage(m_hWndRTB, EM_LINEFROMCHAR, SelRange.cpMin, 0)

    ''取得指定行第一个字符的位置
    LineIndex = SendMessage(m_hWndRTB, EM_LINEINDEX, CurRow, 0)

    If SelRange.cpMin = LineIndex Then
        GetCurPos.X = 1
    Else

        TempStr = String(SelRange.cpMin - LineIndex, 13)

        ''复制当前行开始到选择文本开始的文本
        CopyMemory ByVal StrPtr(TempStr), ByVal StrPtr(TempArray) + LineIndex, SelRange.cpMin - LineIndex
        TempArray = TempStr

        ''删除无用的信息
        ReDim Preserve TempArray(SelRange.cpMin - LineIndex - 1)

        ''转换为 Unicode
        TempStr = StrConv(TempArray, vbUnicode)

        GetCurPos.X = Len(TempStr) + 1
    End If
    GetCurPos.Y = CurRow + 1
End Function

Private Sub pStyleChanged(Optional ByVal Hwnd As Long = 0)
    On Error Resume Next
    If Hwnd = 0 Then Hwnd = m_hWndRTB
    SetWindowPos m_hWndRTB, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
    Err.Clear
End Sub

Private Sub pAttachMessages()
'消息捕获绑定
    On Error Resume Next
    Dim dwMask As Long
    Subclass1.Hwnd = UserControl.Hwnd
    Subclass1.Messages(WM_NOTIFY) = True
    Subclass1.Messages(WM_COMMAND) = True
    Subclass2.Hwnd = rtbThis.Hwnd
    Subclass2.Messages(WM_MOUSEWHEEL) = True
    
    ' 键盘鼠标事件
    dwMask = ENM_KEYEVENTS Or ENM_MOUSEEVENTS
    ' 选择区域改变事件
    dwMask = dwMask Or ENM_SELCHANGE
    ' 放下文件
    dwMask = dwMask Or ENM_DROPFILES
    ' 滚动事件
    dwMask = dwMask Or ENM_SCROLL
    ' 更新
    dwMask = dwMask Or ENM_UPDATE
    ' 内容改变
    dwMask = dwMask Or ENM_CHANGE
    '自动适应尺寸
    dwMask = dwMask Or ENM_REQUESTRESIZE
    ' 超链接事件
    dwMask = dwMask Or ENM_LINK
    ' 保护区事件
    dwMask = dwMask Or ENM_PROTECTED
    '  拖动放下完毕
    dwMask = dwMask Or ENM_DRAGDROPDONE
    '  滚动事件
    dwMask = dwMask Or ENM_SCROLLEVENTS
    '  滚动事件
    dwMask = dwMask Or ENM_OBJECTPOSITIONS
    
    SendMessageLong m_hWndRTB, EM_SETEVENTMASK, 0, dwMask      '设置事件掩码
    
    m_bSubClassing = True
    Err.Clear
End Sub

Private Sub pDetachMessages()
'取消消息捕获
    m_bSubClassing = False
End Sub

Private Sub pInitialise()
'句柄初始化
    On Error Resume Next
    pTerminate
    If (UserControl.Ambient.UserMode) Then
        m_hWnd = UserControl.Hwnd
        m_hWndParent = UserControl.Parent.Hwnd
        m_hWndRTB = rtbThis.Hwnd
        
        SendMessageLong m_hWndRTB, EM_HIDESELECTION, 0, 0          '禁止失去焦点
        '多级Undo
        Dim lStyle As Long
        lStyle = TM_RICHTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
        SendMessageLong m_hWndRTB, EM_SETTEXTMODE, lStyle, 0
        '最多可以撤销100次操作！
        SendMessageLong m_hWndRTB, EM_SETUNDOLIMIT, 0, 0            '禁止撤销
'        '设置长度限制范围
'        SendMessageLong m_hWnd, EM_EXLIMITTEXT, 0, 9E+16
        '所见即所得样式！！！
'        SendMessageLong m_hWndRTB, EM_SETTARGETDEVICE, Printer.hdc, Printer.Width
        '显示滚动条

        
        If (m_hWndRTB <> 0) Then
           EnableWindow m_hWndRTB, 1
           Call pAttachMessages     '消息绑定
        End If
    End If
    Err.Clear
End Sub

Private Function pTerminate()
'销毁句柄
    On Error Resume Next
    If (m_hWndRTB <> 0) Then
        '终止消息绑定（很重要！）
        Call pDetachMessages          '取消消息绑定
        '删除窗体！
        ShowWindow m_hWndRTB, SW_HIDE
        SetParent m_hWndRTB, 0
        DestroyWindow m_hWndRTB
        Call SendMessage(m_hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
        '表示窗体清空！
        m_hWndRTB = 0
    End If
    Err.Clear
End Function

Public Sub ResetAuditText()
    '恢复所选文本修订内容
    Dim lS As Long, lE As Long, i As Long, lG As Long, COLOR As OLE_COLOR
    With Me
        .Freeze
        lS = .Selection.StartPos
        lE = .Selection.EndPos
        i = lS
        Do While i < lE
            If .Range(i, i + 1).Font.Protected = False And .Range(i, i + 1).Font.Hidden = False Then
                COLOR = IIf(.Range(i, i + 1).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, i + 1).Font.ForeColor)
                If IsNewCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = False Then
                    '下一个字符为新增文本，则直接删除之！
                    .Range(i, i + 1) = ""
                    lE = lE - 1
                ElseIf IsDelCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = True Then
                    '下一个字符为删除文本，则恢复文本为（无删除线＋删除前的颜色）。
                    lG = RGBGreen(COLOR)
                    If lG <> 0 Then
                        '表示该文本在删除前是新增文本，那么应该恢复为新增状态
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = RGB(255, lG, 0)
                    Else
                        '否则恢复为黑色
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = tomAutoColor
                    End If
                    i = i + 1
                Else
                    i = i + 1
                End If
            Else
                '若为保护/隐藏文本，则直接后移一位。
                i = i + 1
            End If
        Loop
        .Range(i, i).Selected
        .UnFreeze
    End With
End Sub

Public Sub AcceptAuditText()
    '接受所选文本修订内容（不包括对象）
    Dim lS As Long, lE As Long, i As Long, bForce As Boolean
    Dim r As Long, g As Long, b As Long, COLOR As OLE_COLOR
    With Me
        .Freeze
        bForce = .ForceEdit
        .ForceEdit = True
        lS = .Selection.StartPos
        lE = Len(.Text)
        i = lS
        Do While i < lE
            If .Range(i, i + 1).Font.Hidden = False Then
                COLOR = .Range(i, i + 1).Font.ForeColor
                If COLOR = tomAutoColor Or COLOR = tomUndefined Then COLOR = vbBlack
                r = RGBRed(COLOR): g = RGBGreen(COLOR): b = RGBBlue(COLOR)
                If r = 255 And g > 0 And b = 0 And .Range(i, i + 1).Font.Strikethrough = False Then
                    '下一个字符为新增文本，转为普通文本
                    .Range(i, i + 1).Font.Strikethrough = False
                    .Range(i, i + 1).Font.ForeColor = tomAutoColor
                    .Range(i, i + 1).Font.Italic = False
                    i = i + 1
                ElseIf r = 255 And b > 0 And .Range(i, i + 1).Font.Strikethrough = True Then
                    '下一个字符为删除文本，则直接删除之！
                    .Range(i, i + 1) = ""
                    lE = lE - 1
                Else
                    '否则，不变
                    i = i + 1
                End If
            Else
                '若为保护/隐藏文本，则直接后移一位。
                i = i + 1
            End If
        Loop
        .Range(i, i).Selected
        .ForceEdit = bForce
        .UnFreeze
    End With
End Sub

Private Sub rtbThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.AuditMode And KeyCode = 45 Then KeyCode = 0: Exit Sub
    RaiseEvent BeforeKeyDown(KeyCode, Shift)
    Dim lS As Long, lE As Long, lSS As Long, lSS2 As Long
    Dim lF As Single, LL As Single, lR As Single
    Dim W As Long
    Dim COLOR As OLE_COLOR
    Const LIMITWIDTH = 3000
    
    If Me.ReadOnly Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyEscape Then Exit Sub
    If Shift = 2 And KeyCode = 17 Then Exit Sub
    
    If mvarAuditMode Then
        '审核模式
        Select Case KeyCode
        Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
            vbKeyEscape, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
            vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
            vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, _
            vbKeyF10, vbKeyF11, vbKeyF12, vbKeyShift, vbKeyControl
            
            Exit Sub
        End Select
        
        Dim i As Long, j As Long, lLen As Long
        With Me
            i = .SelStart
            j = .SelStart + .SelLength
            If (.Range(i, j).Font.Protected Or .Range(i, j).Font.Hidden) And KeyCode <> vbKeyTab Then
'                KeyCode = 0
'                RaiseEvent ModifyProtected(False, rtbThis.SelStart, rtbThis.SelStart, 0, 0)
                Exit Sub
            End If
            
            Select Case KeyCode
            Case vbKeyBack
                '退格键处理
                If i = j Then
                    If .Range(i - 2, i) = vbCrLf Then
'                        Color = IIf(.Range(i - 2, i).Font.ForeColor = tomAutoColor Or .Range(i - 2, i).Font.ForeColor = tomUndefined, vbBlack, .Range(i - 2, i).Font.ForeColor)
                        If .Range(i - 2, i).Font.Protected Or .Range(i - 2, i).Font.Hidden Then KeyCode = 0: Exit Sub
'                        If IsNewCharColor(Color) And .Range(i - 2, i).Font.Strikethrough = False Then
                            '前面一个字符已经是新增文本，则直接删除之
                            .Range(i - 2, i).Text = ""
'                        ElseIf RGBBlue(Color) <> 0 And IsDelCharColor(Color) = False Then
'                            '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
'                            .Range(i - 2, i - 2).Selected
'                            KeyCode = 0
'                        Else
'                            '否则，标记前面文本为删除文本
'                            .Range(i - 2, i).Text = ""
''                            .Range(i - 2, i).Font.Strikethrough = True
''                            .Range(i - 2, i).Font.ForeColor = GetDelCharColor(Color)
'                            .Range(i - 2, i - 2).Selected
'                        End If
                    Else
                        COLOR = IIf(.Range(i - 1, i).Font.ForeColor = tomAutoColor, vbBlack, .Range(i - 1, i).Font.ForeColor)
                        If .Range(i - 1, i).Font.Protected Or .Range(i - 1, i).Font.Hidden Then KeyCode = 0: Exit Sub
                        If IsNewCharColor(COLOR) And .Range(i - 1, i).Font.Strikethrough = False Then
                            '前面一个字符已经是新增文本，则直接删除之
                            .Range(i - 1, i).Text = ""
                        ElseIf RGBBlue(COLOR) <> 0 And IsDelCharColor(COLOR) = False Then
                            '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                            .Range(i - 1, i - 1).Selected
                            KeyCode = 0
                        Else
                            '否则，标记前面文本为删除文本
                            .Range(i - 1, i).Font.Strikethrough = True
                            .Range(i - 1, i).Font.ForeColor = GetDelCharColor(COLOR)
                            .Range(i - 1, i - 1).Selected
                        End If
                    End If
                Else
                    COLOR = IIf(.Range(i, j).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, j).Font.ForeColor)
                    If IsNewCharColor(COLOR) And .Range(i, j).Font.Strikethrough = False Then
                        '选中文本为新增文本，直接删除之
                        .Range(i, j) = ""
                    ElseIf RGBBlue(COLOR) <> 0 And IsDelCharColor(COLOR) = False Then
                        '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                        .Range(i, i).Selected
                        KeyCode = 0
                    ElseIf IsNewCharColor(COLOR) = False And IsDelCharColor(COLOR) = False Then
                        '否则，如果为普通文本
                        .ForceEdit = True
                        .Range(i, j).Font.Strikethrough = True
                        .Range(i, j).Font.ForeColor = GetDelCharColor(COLOR)
                        .ForceEdit = False
                    ElseIf .Range(i, j).Font.ForeColor = tomUndefined Then
                        '否则，如果为混合文本，则不处理。
                        .Range(j, j).Selected
                    End If
                End If
                KeyCode = 0     '清除按键
            Case vbKeyDelete
                '删除键由外部处理
            Case vbKeyTab
                '检测Tab键！
                Dim iRetVal1 As Integer
                iRetVal1 = GetKeyState(VK_SHIFT)
                ' 如果没有按shift，检查tab
                If iRetVal1 <> -128 And iRetVal1 <> -127 Then
                    iRetVal1 = GetKeyState(VK_TAB)
                    If iRetVal1 = -128 Or iRetVal1 = -127 Then ' tab键按下
                        RaiseEvent PressTabKey
                        KeyCode = 0
                        lS = Selection.StartPos
                        lE = Selection.EndPos
                        lSS = IIf(lS - 2 > 0, lS - 2, 0)
                        lSS2 = IIf(lS - 16 > 0, lS - 16, 0)
                        If Range(lSS, lS) = vbCrLf Or lS = 0 Or (Range(lSS2, lSS2 + 3) = "OE(" And Range(lSS2, lSS2 + 3).Font.Hidden = True) Then
                            '行首，不处理
                        Else
                            '行中，增加一个Tab
                            If Range(lS, lE).Font.Protected = False Then
                                '否则，如果前面文本为以前版本的删除文本，则不作任何处理！ Then
                                .ForceEdit = True
                                .Range(lS, lE).Font.ForeColor = GetDelCharColor(.Range(lS, lE).Font.ForeColor)
                                .Range(lS, lE).Font.Strikethrough = True
                                .Range(lE, lE).Text = vbTab
                                .Range(lE, lE + 1).Font.ForeColor = GetNewCharColor(tomAutoColor)
                                .Range(lE, lE + 1).Font.Strikethrough = False
                                .Range(lE + 1, lE + 1).Selected
                                .ForceEdit = False
                            End If
                        End If
                        '设置当前位置为新增文本（红色）
                        If lS = lE Then
                            .ForceEdit = True
                            On Error Resume Next
                            rtbThis.SelColor = GetNewCharColor(tomAutoColor)
                            rtbThis.SelStrikeThru = False   '去掉删除线
                            .ForceEdit = False
                        End If
                        If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                            UserControl.Extender.SetFocus
                        End If
                    End If
                End If
                
            Case Else
                '否则，为普通按键处理
                If i <> j Then
                    '选中多个文本
                    COLOR = IIf(.Range(i, j).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, j).Font.ForeColor)
                    If IsDelCharColor(COLOR) And .Range(i, j).Font.Strikethrough = True Then
                        '选中文本为已删除文本，不处理
                        .Range(j, j).Selected   '选中末尾
                    ElseIf IsNewCharColor(COLOR) And .Range(i, j).Font.Strikethrough = False Then
                        '选中文本为新插入文本，直接覆盖之
                        .Range(i, j) = ""
                        .Range(i, i).Selected   '选中末尾
                    ElseIf RGBBlue(COLOR) <> 0 Then
                        '否则，如果前面文本为以前版本的删除文本，则不作任何处理！
                        .Range(j, j).Selected   '选中末尾
                    ElseIf .Range(i, j).Font.ForeColor = tomUndefined Then
                        '选中混合文本
                        .Range(j, j).Selected   '选中末尾
                    Else
                        '普通非保护/隐藏的文本，则标记为删除状态
                        .ForceEdit = True
                        .Range(i, j).Font.Strikethrough = True
                        .Range(i, j).Font.ForeColor = GetDelCharColor(COLOR)
                        .ForceEdit = False
                        .Range(j, j).Selected   '选中末尾
                    End If
                End If
                '设置当前位置为新增文本（红色）
                .ForceEdit = True
                On Error Resume Next
                rtbThis.SelColor = GetNewCharColor(tomAutoColor)
                rtbThis.SelStrikeThru = False   '去掉删除线
                .ForceEdit = False
            End Select
        End With
    Else
        '普通书写模式
        If KeyCode = vbKeyBack Then
            lS = Selection.StartPos
            lE = Selection.EndPos
            lSS = IIf(lS - 2 > 0, lS - 2, 0)
            lSS2 = IIf(lS - 16 > 0, lS - 16, 0)
            If Range(lSS, lS) = vbCrLf Or lS = 0 Or (Range(lSS2, lSS2 + 3) = "OE(" And Range(lSS2, lSS2 + 3).Font.Hidden = True) Then
    '            KeyAscii = 0
                '行首，增加首行缩进
                lF = Range(lS, lE).Para.FirstLineIndent
                LL = Range(lS, lE).Para.LeftIndent
                lR = Range(lS, lE).Para.RightIndent
                If lF = tomUndefined Then lF = 0
                If LL = tomUndefined Then LL = 0
                If lR = tomUndefined Then lR = 0
                
                If lF <> 0 Or LL <> 0 Then KeyCode = 0
                
                W = (mvarPaperWidth - mvarMarginLeft - mvarMarginRight - LIMITWIDTH) * mvarZoomFactor / 20
                
                If lF > 0 Then
                    lF = 0
                Else
                    LL = LL - DefaultTabStop
                End If
                If LL < 0 Then LL = 0
                ForceEdit = True
                If Range(lS - 2, lS) = vbCrLf And Range(lS, lE).Para.ListType <> cprLTNone Then
                    '如果在行首，则清除项目符号
                    Range(lS, lE).Para.ListType = cprLTNone
                Else
                    '否则，缩小左缩进
                    Range(lS, lE).Para.SetIndents lF, LL, lR
                End If
                
                ForceEdit = False
                If lF = 0 And LL = 0 And Range(lS - 32, lS - 29) = "OS(" And Range(lS - 32, lS - 29).Font.Hidden Then
                    If Range(lS - 32, lS - 34) = vbCrLf Then
                        Range(lS - 34, lS - 34).Selected
                    Else
                        Range(lS - 32, lS - 32).Selected
                    End If
                    KeyCode = 0
                ElseIf Range(lS - 2, lS) = vbCrLf And Range(lS, lS + 3) = "OS(" And Range(lS, lS + 3).Font.Hidden Then
                    Range(lS - 2, lS - 2).Selected
                    KeyCode = 0
                ElseIf Range(lS - 2, lS).Font.Protected And Range(lS - 2, lS) = vbCrLf Then
                    Range(lS - 2, lS - 2).Selected
                    KeyCode = 0
                Else
                    RaiseEvent SelChange(lS, lE)
                End If
            ElseIf Range(lS - 2, lS).Font.Protected And Range(lS - 2, lS) = vbCrLf Then
                Range(lS - 2, lS - 2).Selected
                KeyCode = 0
            ElseIf Range(lS - 1, lS).Font.Protected And Range(lS - 2, lS) <> vbCrLf Then
                Range(lS - 1, lS - 1).Selected
                KeyCode = 0
            End If
            If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                UserControl.Extender.SetFocus
            End If
        ElseIf KeyCode = vbKeyTab Then
            '检测Tab键！
            Dim iRetVal2 As Integer
            iRetVal2 = GetKeyState(VK_SHIFT)
            ' 如果没有按shift，检查tab
            If iRetVal2 <> -128 And iRetVal2 <> -127 Then
                iRetVal2 = GetKeyState(VK_TAB)
                If iRetVal2 = -128 Or iRetVal2 = -127 Then ' tab键按下
                    RaiseEvent PressTabKey
                    KeyCode = 0
                    lS = Selection.StartPos
                    lE = Selection.EndPos
                    lSS = IIf(lS - 2 > 0, lS - 2, 0)
                    lSS2 = IIf(lS - 16 > 0, lS - 16, 0)
                    If Range(lSS, lS) = vbCrLf Or lS = 0 Or (Range(lSS2, lSS2 + 3) = "OE(" And Range(lSS2, lSS2 + 3).Font.Hidden = True) Then
                        '行首，增加首行缩进
                        lF = Range(lS, lE).Para.FirstLineIndent
                        LL = Range(lS, lE).Para.LeftIndent
                        lR = Range(lS, lE).Para.RightIndent
                        If lF = tomUndefined Then lF = 0
                        If LL = tomUndefined Then LL = 0
                        If lR = tomUndefined Then lR = 0
                        
                        W = (mvarPaperWidth - mvarMarginLeft - mvarMarginRight - LIMITWIDTH) * mvarZoomFactor / 20
                        
                        If lF < DefaultTabStop Then
                            lF = DefaultTabStop
                        Else
                            LL = LL + DefaultTabStop
                        End If
                        
                        '不能超出范围
                        If LL < 0 Then LL = 0
                        If LL > W Then
                            If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                                UserControl.Extender.SetFocus: Exit Sub
                            End If
                        End If
                        
                        If lF < 0 Then
                            If Abs(lF) > LL Then lF = -LL
                        Else
                            If lF + LL > W Then
                                If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                                    UserControl.Extender.SetFocus: Exit Sub
                                End If
                            End If
                        End If
                        ForceEdit = True
                        Range(lS, lE).Para.SetIndents lF, LL, lR
                        ForceEdit = False
                        RaiseEvent SelChange(lS, lE)
                    Else
                        '行中，增加一个Tab
                        If Range(lS, lE).Font.Protected = False Then
                            Me.ForceEdit = True
                            Range(lS, lE).Text = vbTab
                            Range(lS + 1, lS + 1).Selected
                            Me.ForceEdit = False
                        End If
                    End If
                    If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                        UserControl.Extender.SetFocus
                    End If
                End If
            End If
        End If
    End If
    Err.Clear
End Sub

Private Sub Subclass2_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim tP As POINTAPI
    Select Case Msg
    Case WM_MOUSEWHEEL   '滚动
        Dim wzDelta As Long, wKeys As Long
        'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
        '大于零表示滚轮向前滚动（朝显示器方向）
        wzDelta = HIWORD(wParam)
        'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
        wKeys = LOWORD(wParam)
        tP.X = LOWORD(lParam)    'pt鼠标的坐标
        tP.Y = HIWORD(lParam)
        '--------------------------------------------------
        If wzDelta < 0 Then  '朝用户方向
           bWay = True
        Else                 '朝显示器方向
           bWay = False
        End If
        '--------------------------------------------------
        '将屏幕坐标转换为Form1.窗口坐标
        ScreenToClient Hwnd, tP
        sngX = tP.X
        sngY = tP.Y
        intShift = wKeys
        bMouseFlag = True  '置滚动标志
        If bMouseFlag = True Then
            bMouseFlag = False
            RaiseEvent MouseWheel(bWay, intShift, sngX, sngY, CLng(wzDelta)) '激活事件
        End If
        Result = Subclass2.CallWndProc(Msg, wParam, lParam)
    Case Else
        Result = Subclass2.CallWndProc(Msg, wParam, lParam)
    End Select
End Sub

'#############################################################################################################
'##     内部控件事件
'#############################################################################################################

Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_InitProperties()
'创建对象新实例时发生，即新属性的最初初始化代码！（即，当用户在窗体上放置一个控件时触发此事件！运行时不再触发！）
    AutoDetectURL = True
    BackColor = vbWhite
    Border = True
    DoDefaultURLClick = False
    Enabled = True
    FileName = ""
    ForceEdit = False
    Modified = False
    ReadOnly = False
    Text = ""
    Title = "未命名文档"
    ViewMode = cprNormal
    ZoomFactor = 1#
    AuditMode = False
    DefaultTabStop = rtbThis.Font.SIZE * 2
    
    PaperKind = cprPKA4
    PaperOrient = cprPOPortrait
    PaperHeight = 16840
    PaperWidth = 11907
    MarginTop = 1400
    MarginBottom = 1400
    MarginLeft = 1800
    MarginRight = 1800
    
    WYSIWYG = True
End Sub

Private Sub UserControl_Resize()
    rtbThis.Move 0, 0, ScaleWidth, ScaleHeight
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        rtbThis.Visible = True
        lblThis.Visible = False
        If rtbThis.Enabled And rtbThis.Visible Then
            rtbThis.SetFocus
        End If
    Else
        rtbThis.Visible = False
        lblThis.Visible = True
        lblThis.Caption = "(普通视图)"
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'当保存对象的实例时，发生该事件。该事件通知对象此时需要保存对象的状态，以便将来可恢复该状态。大多数情况下，对象的状态仅包括属性值。
'属性保存（静态属性的保存）
    PropBag.WriteProperty "AutoDetectURL", AutoDetectURL, True
    PropBag.WriteProperty "BackColor", BackColor, vbWhite
    PropBag.WriteProperty "Border", Border, True
    PropBag.WriteProperty "DefaultTabStop", DefaultTabStop, 21
    PropBag.WriteProperty "DoDefaultURLClick", DoDefaultURLClick, False
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "FileName", FileName, ""
    PropBag.WriteProperty "ForceEdit", ForceEdit, False
    PropBag.WriteProperty "Modified", Modified, False
    PropBag.WriteProperty "ReadOnly", ReadOnly, False
    PropBag.WriteProperty "Title", Title, "未命名文档"
    PropBag.WriteProperty "ViewMode", ViewMode, cprNormal
    PropBag.WriteProperty "ZoomFactor", ZoomFactor, 1#
    PropBag.WriteProperty "PaperKind", PaperKind, cprPKA4
    PropBag.WriteProperty "PaperOrient", PaperOrient, cprPOPortrait
    PropBag.WriteProperty "PaperHeight", PaperHeight, 16840
    PropBag.WriteProperty "PaperWidth", PaperWidth, 11907
    PropBag.WriteProperty "MarginTop", MarginTop, 1400
    PropBag.WriteProperty "MarginBottom", MarginBottom, 1400
    PropBag.WriteProperty "MarginLeft", MarginLeft, 1800
    PropBag.WriteProperty "MarginRight", MarginRight, 1800
    PropBag.WriteProperty "WYSIWYG", WYSIWYG, True
    PropBag.WriteProperty "AuditMode", AuditMode, False
    
    PropertyChanged "AutoDetectURL"
    PropertyChanged "BackColor"
    PropertyChanged "Border"
    PropertyChanged "DefaultTabStop"
    PropertyChanged "DoDefaultURLClick"
    PropertyChanged "Enabled"
    PropertyChanged "FileName"
    PropertyChanged "ForceEdit"
    PropertyChanged "Modified"
    PropertyChanged "ReadOnly"
    PropertyChanged "Text"
    PropertyChanged "Title"
    PropertyChanged "ViewMode"
    PropertyChanged "ZoomFactor"
    PropertyChanged "PaperKind"
    PropertyChanged "PaperOrient"
    PropertyChanged "PaperHeight"
    PropertyChanged "PaperWidth"
    PropertyChanged "MarginTop"
    PropertyChanged "MarginBottom"
    PropertyChanged "MarginLeft"
    PropertyChanged "MarginRight"
    PropertyChanged "WYSIWYG"
    PropertyChanged "AuditMode"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'当加载具有保存状态的对象的旧实例时，发生该事件。
'属性读取（静态属性的读取，从而转化为动态属性，此时调用pInitialise函数初始化句柄！）
    If Ambient.UserMode Then
        Call pInitialise
        m_TOM.Init rtbThis, UserControl.Extender
    End If
    AutoDetectURL = PropBag.ReadProperty("AutoDetectURL", True)
    BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    Border = PropBag.ReadProperty("Border", True)
    DefaultTabStop = PropBag.ReadProperty("DefaultTabStop", 21)
    DoDefaultURLClick = PropBag.ReadProperty("DoDefaultURLClick", False)
    Enabled = PropBag.ReadProperty("Enabled", True)
    FileName = PropBag.ReadProperty("FileName", "")
    ForceEdit = PropBag.ReadProperty("ForceEdit", False)
    ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    Title = PropBag.ReadProperty("Title", "未命名文档")
    ViewMode = PropBag.ReadProperty("ViewMode", cprNormal)
    ZoomFactor = PropBag.ReadProperty("ZoomFactor", 1#)
    PaperKind = PropBag.ReadProperty("PaperKind", cprPKA4)
    PaperOrient = PropBag.ReadProperty("PaperOrient", cprPOPortrait)
    PaperHeight = PropBag.ReadProperty("PaperHeight", 16840)
    PaperWidth = PropBag.ReadProperty("PaperWidth", 11907)
    MarginTop = PropBag.ReadProperty("MarginTop", 1400)
    MarginBottom = PropBag.ReadProperty("MarginBottom", 1400)
    MarginLeft = PropBag.ReadProperty("MarginLeft", 1800)
    MarginRight = PropBag.ReadProperty("MarginRight", 1800)
    WYSIWYG = PropBag.ReadProperty("WYSIWYG", True)
    AuditMode = PropBag.ReadProperty("AuditMode", False)
    
    '------------------------------------------
    '以下是运行用户保存的默认属性：
    
    If Ambient.UserMode Then
        '获取默认的字体属性
        With rtbThis.Font
            .Name = GetSetting(UCase(App.ProductName), "FONT", UCase("Name"), "宋体")
            .Italic = CBool(GetSetting(UCase(App.ProductName), "FONT", UCase("Italic"), 0))
            .Bold = CBool(GetSetting(UCase(App.ProductName), "FONT", UCase("Bold"), 0))
            .SIZE = GetSetting(UCase(App.ProductName), "FONT", UCase("Size"), 10.5)
        End With
        If Me.ReadOnly = False Then
            With Me.TOM.TextDocument.Selection.Font
                .Name = GetSetting(UCase(App.ProductName), "FONT", UCase("Name"), "宋体")
                .Italic = CBool(GetSetting(UCase(App.ProductName), "FONT", UCase("Italic"), 0))
                .Bold = CBool(GetSetting(UCase(App.ProductName), "FONT", UCase("Bold"), 0))
                .SIZE = GetSetting(UCase(App.ProductName), "FONT", UCase("Size"), 10.5)
            End With
        End If
        DefaultTabStop = Me.TOM.TextDocument.Selection.Font.SIZE * 2
        
        '获取默认的段落间距属性
        Dim lngSpacingRule As Long, dblSpacing As Double
        Me.Range(0, 0).Para.SpaceAfter = GetSetting(UCase(App.ProductName), "PARA", UCase("SpaceAfter"), 0)
        Me.Range(0, 0).Para.SpaceBefore = GetSetting(UCase(App.ProductName), "PARA", UCase("SpaceBefore"), 0)
        lngSpacingRule = GetSetting(UCase(App.ProductName), "PARA", UCase("LineSpacingRule"), cprLSSignle)
        Select Case lngSpacingRule
        Case cprLSSignle, cprLS1pt5, cprLSDouble
            dblSpacing = 0
        Case cprLSAtLeast, cprLSExactly
            dblSpacing = GetSetting(UCase(App.ProductName), "PARA", UCase("LineSpacing"), 12)
        Case cprLSMultiple
            dblSpacing = GetSetting(UCase(App.ProductName), "PARA", UCase("LineSpacing"), 1.5)
        End Select
        Call Me.Range(0, Len(Me.Text)).Para.SetLineSpacing(lngSpacingRule, dblSpacing)
        If ViewMode = cprNormal And WYSIWYG Then Call WYSIWYG_RTF(rtbThis, MarginLeft * mvarZoomFactor, MarginRight * mvarZoomFactor, MarginTop * mvarZoomFactor, MarginBottom * mvarZoomFactor, PaperWidth * mvarZoomFactor, PaperHeight * mvarZoomFactor)
    End If
    
    Modified = False    '此属性应该放到最后，避免ViewMode使得内容改变。
    
End Sub

Private Sub UserControl_Terminate()
'销毁控件时发生
    Call pTerminate
    Set m_TOM = Nothing
End Sub

Private Sub rtbThis_Click()
    RaiseEvent Click
End Sub

Private Sub rtbThis_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtbThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub rtbThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub rtbThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = vbRightButton Then
        If Selection.GetType = cprSTPicture Or (rtbThis.SelLength = 0) Then
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            DoEvents
        End If
        RaiseEvent RequestRightMenu(Shift, X, Y)
    End If
End Sub

Private Sub UserControl_Click()
    If rtbThis.Visible And rtbThis.Enabled Then
        rtbThis.SetFocus
    End If
End Sub

Public Function LockAllOLEObjectSize() As Boolean
    '锁定所有OLE对象尺寸
    Dim blnForce As Boolean
    blnForce = Me.ForceEdit
    Me.ForceEdit = True
    Me.TOM.TextDocument.Freeze
    LockAllOLEObjectSize = ChangeReObjectsFlag(Me.TOM, REO_DYNAMICSIZE, -1)
    Me.TOM.TextDocument.UnFreeze
    Me.ForceEdit = blnForce
End Function

Public Function LockOLEObjectSize(ByVal Index As Long) As Boolean
    '锁定指定Index的OLE对象尺寸
    Dim blnForce As Boolean
    blnForce = Me.ForceEdit
    Me.ForceEdit = True
    Me.TOM.TextDocument.Freeze
    LockOLEObjectSize = ChangeReObjectsFlag(Me.TOM, REO_DYNAMICSIZE, Index)
    Me.TOM.TextDocument.UnFreeze
    Me.ForceEdit = blnForce
End Function

Public Function GetDelCharColor(ByRef COLOR As OLE_COLOR) As OLE_COLOR
    Dim mColor As OLE_COLOR
    mColor = COLOR
    RaiseEvent GetDelCharColor(mColor)
    GetDelCharColor = mColor
End Function

Public Function GetNewCharColor(ByRef COLOR As OLE_COLOR) As OLE_COLOR
    Dim mColor As OLE_COLOR
    mColor = COLOR
    RaiseEvent GetNewCharColor(mColor)
    GetNewCharColor = mColor
End Function

Public Function IsDelCharColor(ByRef COLOR As OLE_COLOR) As Boolean
    Dim mColor As OLE_COLOR, blnIsDelCharColor As Boolean
    mColor = COLOR
    RaiseEvent IsDelCharColor(mColor, blnIsDelCharColor)
    IsDelCharColor = blnIsDelCharColor
End Function

Public Function IsNewCharColor(ByRef COLOR As OLE_COLOR) As Boolean
    Dim mColor As OLE_COLOR, blnIsNewCharColor As Boolean
    mColor = COLOR
    RaiseEvent IsNewCharColor(mColor, blnIsNewCharColor)
    IsNewCharColor = blnIsNewCharColor
End Function

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim tNMH As NMHDR
    Dim tSC As SelChange
    Dim tEN As ENLINK
    Dim tMF As MSGFILTER
    Dim tPR As ENPROTECTED
    Dim tP As POINTAPI
    Dim tR As RECT
    Dim tPS As PAINTSTRUCT
    Dim X As Single, Y As Single
    Dim iKeyCode As Integer, iKeyAscii As Integer, iShift As Integer
    Dim iBtn As Integer
    Dim bDefault As Boolean
    Dim bDoIt As Boolean
    Dim id As Long
    Dim Block As Boolean
    Dim iNotifyMsg As Long
    Dim lLen As Long
    Dim sText As String
    Dim rResize As REQRESIZE

    Select Case Msg
    Case WM_COMMAND
        iNotifyMsg = (wParam And &H7FFF0000) \ &H10000
        Select Case iNotifyMsg
        Case EN_CHANGE
            RaiseEvent Change
        Case EN_SETFOCUS
            '输入焦点获取
'            RaiseEvent Focuse
        Case Else
            Result = Subclass1.CallWndProc(Msg, wParam, lParam)
        End Select
    Case WM_NOTIFY  '系统通知
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        If (tNMH.hwndFrom = m_hWndRTB) Then
            Select Case tNMH.code
            Case EN_REQUESTRESIZE
                Dim lngH As Long
                RaiseEvent Zoom(ZoomFactor)     '在用户通过Ctrol加上鼠标滚轮缩放时自动通知客户端
                RaiseEvent RequestLine
            Case EN_SELCHANGE
                CopyMemory tSC, ByVal lParam, Len(tSC)
                RaiseEvent SelChange(tSC.chrg.cpMin, tSC.chrg.cpMax)
            Case EN_LINK
                CopyMemory tEN, ByVal lParam, Len(tEN)
                If mvarDoDefaultURLClick And tEN.Msg = cprLButtonDown Then
                    '打开浏览器
                    Dim eText As TEXTRANGE
                    eText.chrg.cpMin = tEN.chrg.cpMin
                    eText.chrg.cpMax = tEN.chrg.cpMax
                    eText.lpstrText = Space$(1024)
                    lLen = SendMessage(m_hWndRTB, EM_GETTEXTRANGE, 0, eText)
                    sText = Left$(eText.lpstrText, lLen)
                    ShellExecute m_hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
                Else
                    RaiseEvent LinkEvent(tEN.Msg, tEN.chrg.cpMin, tEN.chrg.cpMax)
                End If
            Case EN_PROTECTED
                CopyMemory tPR, ByVal lParam, Len(tPR)
                bDoIt = False
                If mvarForceEdit Then
                    Result = 0
                Else
                    RaiseEvent ModifyProtected(bDoIt, tPR.chrg.cpMin, tPR.chrg.cpMax, tPR.wPad1, giGetShiftState)
                    If bDoIt Then
                       Result = 0
                    Else
                       Result = 1
                    End If
                End If
            Case EN_MSGFILTER
                bDefault = True '采用默认处理方法
                CopyMemory tMF, ByVal lParam, Len(tMF)
                Select Case tMF.Msg
                Case WM_KEYDOWN
                    iShift = giGetShiftState()
                    iKeyCode = tMF.wParam
                    RaiseEvent KeyDown(iKeyCode, iShift)
                    '屏蔽默认快捷键 ^C/V/X/A ！
                    '如果按 SHIFT 键，则 shift 为 1；如果按 CTRL 键，则 shift 为 2；如果按 ALT 键，则 shift 为 4。
                    If iShift = 2 And (iKeyCode = vbKeyC Or iKeyCode = vbKeyV Or iKeyCode = vbKeyX Or iKeyCode = vbKeyA Or iKeyCode = vbKeyZ Or iKeyCode = vbKeyY) Then
                       bDefault = False
                    End If
                Case WM_CHAR
                    iShift = giGetShiftState()
                    iKeyAscii = tMF.wParam
                    RaiseEvent KeyPress(iKeyAscii)
                    If iShift = 2 And (iKeyAscii = vbKeyC Or iKeyAscii = vbKeyV Or iKeyAscii = vbKeyX Or iKeyAscii = vbKeyA Or iKeyCode = vbKeyZ Or iKeyCode = vbKeyY) Then
                       bDefault = False
                    End If
                Case WM_KEYUP
                    iShift = giGetShiftState()
                    iKeyCode = tMF.wParam
                    RaiseEvent KeyUp(iKeyCode, iShift)
                    If iShift = 2 And (iKeyCode = vbKeyC Or iKeyCode = vbKeyV Or iKeyCode = vbKeyX Or iKeyCode = vbKeyA Or iKeyCode = vbKeyZ Or iKeyCode = vbKeyY) Then
                       bDefault = False
                    End If
                Case Else
                    Result = Subclass1.CallWndProc(Msg, wParam, lParam)
                End Select
                If Not bDefault Then
                    Result = 1&
                End If
            Case Else
                Result = Subclass1.CallWndProc(Msg, wParam, lParam)
            End Select
        End If
    Case Else
        Result = Subclass1.CallWndProc(Msg, wParam, lParam)
    End Select
End Sub
Public Sub ClearEndCrlfChar()
Dim strText As String
    strText = Me.Text
    Do While strText <> ""
        If Mid(strText, Len(strText)) = vbCrLf Or Asc(Right(strText, 1)) = 13 Or Asc(Right(strText, 1)) = 10 Then
            Range(Len(strText) - 1, Len(strText)).Text = ""
            strText = Me.Text
        Else
            Exit Do
        End If
    Loop
End Sub
