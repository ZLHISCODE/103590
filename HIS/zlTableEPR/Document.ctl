VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.UserControl Document 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ScaleHeight     =   2235
   ScaleWidth      =   3120
   ToolboxBitmap   =   "Document.ctx":0000
   Begin zlSubclass.Subclass Subclass1 
      Left            =   3090
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   2040
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   3598
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Document.ctx":0312
   End
End
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_hWndRTB As Long           'RTB的 hWnd
Private m_hWnd As Long              '控件的 hWnd
Private m_hWndParent  As Long       '父窗体的 hWnd

Private m_bSubClassing As Boolean   '是否继承子类
Private m_TOM As New cTabTextDocument  'TOM 3.0 模型，核心对象

'#############################################################################################################
'##     独立属性
'#############################################################################################################

Private mvarAutoDetectURL As Boolean
Private mvarBackColor As OLE_COLOR
Private mvarBorder As Boolean
Private mvarDefaultTabStop As Single
Private mvarEnabled As Boolean
Private mvarFileName As String
Private mvarFoot As String
Private mvarForceEdit As Boolean
Private mvarHead As String
Private mvarReadOnly As Boolean
Private mvarTitle As String
Private mvarTransparent As Boolean
'用途: 超链接键盘鼠标事件
Public Enum LinkEventTypeEnum
   cprLButtonDblClick = WM_LBUTTONDBLCLK
   cprLButtonDown = WM_LBUTTONDOWN
   cprLButtonUp = WM_LBUTTONUP
   cprMouseMove = WM_MOUSEMOVE
   cprRButtonDblClick = WM_RBUTTONDBLCLK
   cprRButtonDown = WM_RBUTTONDOWN
   cprRBUttonUp = WM_RBUTTONUP
   cprSetCursor = WM_SETCURSOR
End Enum
'#############################################################################################################
'##     事件声明
'#############################################################################################################
Public Event Change()       '内容改变！
Public Event Resize()    '控件尺寸改变
Public Event RequestLine()              '请求行数改变
Public Event SelChange(ByVal lStart As Long, ByVal lEnd As Long)   '选择区域改变
Public Event ModifyProtected(ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)  '试图编辑受保护区域
Public Event BeforeKeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event RequestRightMenu(ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event Click()        '单击
Public Event DblClick()     '双击
Public Event GetNewCharColor(ByRef Color As OLE_COLOR)     '获取新增字符的颜色
Public Event IsNewCharColor(ByVal Color As OLE_COLOR, ByRef blnIsNewCharColor As Boolean)   '判断是否是新增字符的颜色

'#############################################################################################################
'##     公共属性
'#############################################################################################################

Public Property Get OriginRTB() As Object
    Set OriginRTB = rtf
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
    CurrentColumn = pt.x
End Property

Public Property Get CurrentLine() As Long
'    CurrentLine = SendMessageLong(m_hWndRTB, EM_EXLINEFROMCHAR, 0, m_TOM.TextDocument.Selection.Start) + 1
    Dim l  As Long
    l = SendMessage(m_hWndRTB, EM_LINEINDEX, -1, 0)
    CurrentLine = SendMessage(m_hWndRTB, EM_LINEFROMCHAR, l, 0) + 1
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

Public Property Get Filename() As String
    Filename = mvarFileName
End Property

Public Property Let Filename(ByVal vData As String)
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
    Head = mvarForceEdit
End Property

Public Property Let Head(ByVal vData As String)
    mvarHead = vData
    PropertyChanged "Head"
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get hWndRTB() As Long
   hWndRTB = rtf.hWnd
End Property

Public Property Get LineCount() As Long
   LineCount = SendMessageLong(m_hWndRTB, EM_GETLINECOUNT, 0, 0)
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
    Set MouseIcon = rtf.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtf.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = rtf.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    rtf.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
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
    SelLength = rtf.SelLength
End Property

Public Property Let SelLength(vData As Long)
    rtf.SelLength = vData
End Property

Public Property Get SelRTF() As String
    SelRTF = rtf.SelRTF
End Property

Public Property Let SelRTF(vData As String)
    On Error Resume Next
    rtf.SelRTF = vData
    Err.Clear
End Property

Public Property Get SelStart() As Long
    SelStart = rtf.SelStart
End Property

Public Property Let SelStart(vData As Long)
    rtf.SelStart = vData
End Property

Public Property Get SelText() As String
    If m_hWndRTB <> 0 Then
        SelText = m_TOM.TextDocument.Selection
    Else
        SelText = rtf.SelText
    End If
End Property

Public Property Let SelText(vData As String)
    If m_hWndRTB <> 0 Then
        m_TOM.TextDocument.Selection = vData
    Else
        rtf.SelText = vData
    End If
End Property

Public Property Get Text() As String
    Text = rtf.Text
End Property

Public Property Let Text(ByRef vData As String)
    rtf.Text = vData
End Property

Public Property Get TextRTF() As String
    TextRTF = rtf.TextRTF
End Property

Public Property Let TextRTF(ByRef vData As String)
    rtf.TextRTF = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
    PropertyChanged "Title"
End Property

Public Property Get TOM() As cTabTextDocument
    Set TOM = m_TOM
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    mvarTransparent = vData
    PropertyChanged "Transparent"
End Property
'#############################################################################################################
'##     公共方法
'#############################################################################################################

Public Sub CopyWithFormat()
    '带格式复制
    SendMessageLong m_hWndRTB, WM_COPY, 0, 0
End Sub

Public Sub PasteWithFormat()
    '带格式复制
    SendMessageLong m_hWndRTB, WM_PASTE, 0, 0
End Sub

Public Sub Copy()
    'SendMessageLong m_hWndRTB, WM_COPY, 0, 0
    '过滤掉内嵌关键字
    Dim strTmp As String, i As Long, lS As Long, lE As Long, j As Long
    lS = Selection.StartPos
    lE = Selection.EndPos
    strTmp = Space(lE - lS)
    For i = lS To lE - 1
        If Range(i, i + 1).Font.Hidden = False Then
            j = j + 1
            Mid(strTmp, j, 1) = Range(i, i + 1).Text
        End If
    Next
    
    Clipboard.Clear
    Clipboard.SetText Left(strTmp, j)
End Sub

Public Sub Cut()
    SendMessageLong m_hWndRTB, WM_CUT, 0, 0
'    Clipboard.Clear
'    Clipboard.SetText rtf.SelText
'    Selection.Text = ""
End Sub

Public Sub Delete()
    Selection.Delete
End Sub
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
Public Sub NewDoc()
    m_TOM.TextDocument.Freeze
    m_TOM.TextDocument.New
    rtf.Filename = ""
    m_TOM.TextDocument.DefaultTabStop = mvarDefaultTabStop
    m_TOM.TextDocument.Selection.Para.ClearAllTabs
    m_TOM.TextDocument.UnFreeze
    Filename = ""
'    Title = "未命名文档"
    Modified = False
End Sub

Public Sub OpenDoc(Optional strFile As String = "")
    If strFile <> "" Then Filename = strFile
'    m_TOM.TextDocument.Freeze
    m_TOM.TextDocument.New
    m_TOM.TextDocument.DefaultTabStop = mvarDefaultTabStop
    m_TOM.TextDocument.Selection.Para.ClearAllTabs
    rtf.Filename = strFile
'    m_TOM.TextDocument.UnFreeze
    Modified = False
End Sub

Public Sub Paste()
    If Me.Selection.Font.Protected Or Me.Selection.Font.Hidden Then Exit Sub

    Me.SelStart = Me.SelStart + Me.SelLength
    Me.SelLength = 0
    '设置当前位置为新增文本（红色）
    Me.ForceEdit = True
    On Error Resume Next
    rtf.SelColor = GetNewCharColor(tomAutoColor)    '新增文本
    rtf.SelStrikeThru = False   '去掉删除线
    Me.ForceEdit = False
    rtf.SelText = Clipboard.GetText
    Err.Clear
End Sub

Public Function Range(lStart As Long, lEnd As Long) As cTabRange
    Dim cR As New cTabRange
    cR.Init m_TOM, lStart, lEnd, mvarReadOnly
    Set Range = cR
End Function

Public Sub Redo()
    m_TOM.TextDocument.Redo 1
End Sub

Public Sub SaveDoc(Optional strFile As String = "")
    If strFile <> "" Then Filename = strFile
    rtf.SaveFile CStr(mvarFileName), rtfRTF
    Modified = False
End Sub

Public Sub SelectAll()
    SetSelection m_hWndRTB, 0, -1
'    Range(0, Len(rtf)).Selected
End Sub

Public Function Selection() As cTabSelection
    Dim cS As New cTabSelection
    cS.Init m_TOM, mvarReadOnly
    Set Selection = cS
End Function
Public Function GetCleanTxt(ByVal strData As String) As String
'功能:去除字串中的关键字
Dim strKey As String
    Do Until InStr(strData, "ES(") = 0
        strKey = Mid(strData, InStr(strData, "ES("), 16)
        strData = Replace(strData, strKey, "")
    Loop

    Do Until InStr(strData, "EE(") = 0
        strKey = Mid(strData, InStr(strData, "EE("), 16)
        strData = Replace(strData, strKey, "")
    Loop
    GetCleanTxt = strData
End Function
Public Sub Undo()
    m_TOM.TextDocument.Undo 1
End Sub

Public Sub UnFreeze()
    m_TOM.TextDocument.UnFreeze
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

    TempArray = StrConv(rtf.Text, vbFromUnicode)

    ''取得当前被选中文本的位置 适用于 RichTextBox
    ''rtf 用 EM_GETSEL 消息
    Call SendMessage(m_hWndRTB, EM_EXGETSEL, 0, SelRange)

    ''根据参数wParam指定的字符位置返回该字符所在的行号
    CurRow = SendMessage(m_hWndRTB, EM_LINEFROMCHAR, SelRange.cpMin, 0)

    ''取得指定行第一个字符的位置
    LineIndex = SendMessage(m_hWndRTB, EM_LINEINDEX, CurRow, 0)

    If SelRange.cpMin = LineIndex Then
        GetCurPos.x = 1
    Else

        TempStr = String(SelRange.cpMin - LineIndex, 13)

        ''复制当前行开始到选择文本开始的文本
        CopyMemory ByVal StrPtr(TempStr), ByVal StrPtr(TempArray) + LineIndex, SelRange.cpMin - LineIndex
        TempArray = TempStr

        ''删除无用的信息
        ReDim Preserve TempArray(SelRange.cpMin - LineIndex - 1)

        ''转换为 Unicode
        TempStr = StrConv(TempArray, vbUnicode)

        GetCurPos.x = Len(TempStr) + 1
    End If
    GetCurPos.y = CurRow + 1
End Function

Private Sub pStyleChanged(Optional ByVal hWnd As Long = 0)
    On Error Resume Next
    If hWnd = 0 Then hWnd = m_hWndRTB
    SetWindowPos m_hWndRTB, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
    Err.Clear
End Sub

Private Sub pAttachMessages()
'消息捕获绑定
    On Error Resume Next
    Dim dwMask As Long
    Subclass1.hWnd = UserControl.hWnd
    Subclass1.Messages(WM_NOTIFY) = True
    Subclass1.Messages(WM_COMMAND) = True
    
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
        m_hWnd = UserControl.hWnd
        m_hWndParent = UserControl.Parent.hWnd
        m_hWndRTB = rtf.hWnd
        
        SendMessageLong m_hWndRTB, EM_HIDESELECTION, 0, 0          '禁止失去焦点
        '多级Undo
        Dim lStyle As Long
        lStyle = TM_RICHTEXT Or TM_MULTILEVELUNDO Or TM_MULTICODEPAGE
        SendMessageLong m_hWndRTB, EM_SETTEXTMODE, lStyle, 0
        '最多可以撤销100次操作！
        SendMessageLong m_hWndRTB, EM_SETUNDOLIMIT, 0, 0            '禁止撤销

        
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
        '表示窗体清空！
        m_hWndRTB = 0
    End If
    Err.Clear
End Function

'Private Sub rtf_Change()
'    RaiseEvent Change
'End Sub
'
Private Sub rtf_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent BeforeKeyDown(KeyCode, Shift)
End Sub
'
'Private Sub rtf_KeyPress(KeyAscii As Integer)
' RaiseEvent KeyPress(KeyAscii)
'End Sub

'Private Sub rtf_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
''

Private Sub UserControl_InitProperties()
'创建对象新实例时发生，即新属性的最初初始化代码！（即，当用户在窗体上放置一个控件时触发此事件！运行时不再触发！）
    AutoDetectURL = True
    BackColor = vbWhite
    Border = True
    Enabled = True
    Filename = ""
    ForceEdit = False
    Modified = False
    ReadOnly = False
    Text = ""
    Title = "未命名文档"
    DefaultTabStop = rtf.Font.Size * 2

End Sub

Private Sub UserControl_Resize()
    rtf.Move 0, 0, ScaleWidth, ScaleHeight
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        rtf.Visible = True
        If rtf.Enabled And rtf.Visible Then rtf.SetFocus
    Else
        rtf.Visible = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'当保存对象的实例时，发生该事件。该事件通知对象此时需要保存对象的状态，以便将来可恢复该状态。大多数情况下，对象的状态仅包括属性值。
'属性保存（静态属性的保存）
    PropBag.WriteProperty "AutoDetectURL", AutoDetectURL, True
    PropBag.WriteProperty "BackColor", BackColor, vbWhite
    PropBag.WriteProperty "Border", Border, True
    PropBag.WriteProperty "DefaultTabStop", DefaultTabStop, 21
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "FileName", Filename, ""
    PropBag.WriteProperty "ForceEdit", ForceEdit, False
    PropBag.WriteProperty "Modified", Modified, False
    PropBag.WriteProperty "ReadOnly", ReadOnly, False
    PropBag.WriteProperty "Title", Title, "未命名文档"
    
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
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'当加载具有保存状态的对象的旧实例时，发生该事件。
'属性读取（静态属性的读取，从而转化为动态属性，此时调用pInitialise函数初始化句柄！）
    If Ambient.UserMode Then
        Call pInitialise
        m_TOM.Init rtf, UserControl.Extender
    End If
    AutoDetectURL = PropBag.ReadProperty("AutoDetectURL", True)
    BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    Border = PropBag.ReadProperty("Border", True)
    DefaultTabStop = PropBag.ReadProperty("DefaultTabStop", 21)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Filename = PropBag.ReadProperty("FileName", "")
    ForceEdit = PropBag.ReadProperty("ForceEdit", False)
    ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    Title = PropBag.ReadProperty("Title", "未命名文档")
    
    '------------------------------------------
    '以下是运行用户保存的默认属性：
    If Ambient.UserMode Then '获取默认的字体属性
        With rtf.Font
            .Name = "宋体"
            .Italic = False
            .Bold = False
            .Size = 10.5
        End With
        If Me.ReadOnly = False Then
            With Me.TOM.TextDocument.Selection.Font
                .Name = "宋体"
                .Italic = False
                .Bold = False
                .Size = 10.5
            End With
        End If
        DefaultTabStop = Me.TOM.TextDocument.Selection.Font.Size * 2
        
        '获取默认的段落间距属性
        Dim lngSpacingRule As Long, dblSpacing As Double
        Me.Range(0, 0).Para.SpaceAfter = 0
        Me.Range(0, 0).Para.SpaceBefore = 0
        lngSpacingRule = cprLSSignle
        Select Case lngSpacingRule
        Case cprLSSignle, cprLS1pt5, cprLSDouble
            dblSpacing = 0
        Case cprLSAtLeast, cprLSExactly
            dblSpacing = 12
        Case cprLSMultiple
            dblSpacing = 1.5
        End Select
        Call Me.Range(0, Len(Me.Text)).Para.SetLineSpacing(lngSpacingRule, dblSpacing)
    End If
    
    Modified = False    '此属性应该放到最后，避免ViewMode使得内容改变。
End Sub

Private Sub UserControl_Terminate()
'销毁控件时发生
    Call pTerminate
End Sub

Private Sub rtf_Click()
    RaiseEvent Click
End Sub

Private Sub rtf_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    If Button = vbRightButton Then
        If Selection.GetType = cprSTPicture Or (rtf.SelLength = 0) Then
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            DoEvents
        End If
        RaiseEvent RequestRightMenu(Shift, x, y)
    End If
End Sub

Private Sub UserControl_Click()
    If rtf.Visible And rtf.Enabled Then rtf.SetFocus
End Sub
Public Function GetNewCharColor(ByRef Color As OLE_COLOR) As OLE_COLOR
    Dim mColor As OLE_COLOR
    mColor = Color
    RaiseEvent GetNewCharColor(mColor)
    GetNewCharColor = mColor
End Function
Public Function IsNewCharColor(ByRef Color As OLE_COLOR) As Boolean
    Dim mColor As OLE_COLOR, blnIsNewCharColor As Boolean
    mColor = Color
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
    Dim x As Single, y As Single
    Dim iKeyCode As Integer, iKeyAscii As Integer, iShift As Integer
    Dim iBtn As Integer
    Dim bDefault As Boolean
    Dim bDoIt As Boolean
    Dim ID As Long
    Dim Block As Boolean
    Dim iNotifyMsg As Long
    Dim lLen As Long
    Dim sText As String
    Dim rResize As REQRESIZE

    Select Case Msg
    Case WM_COMMAND                             '当菜单项或下接菜单被点击
        iNotifyMsg = (wParam And &H7FFF0000) \ &H10000
        Select Case iNotifyMsg
        Case EN_CHANGE                          '控件内容文字发生改变
            RaiseEvent Change
        Case EN_SETFOCUS                        '控件接收输入焦点
        
        Case Else                               '否则回调WIN的消息处理
            Result = Subclass1.CallWndProc(Msg, wParam, lParam)
        End Select
    Case WM_NOTIFY                              '系统通知
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        If (tNMH.hwndFrom = m_hWndRTB) Then
            Select Case tNMH.code
            Case EN_REQUESTRESIZE               '通知RichEdit的父窗口，RichEdit的内容变小或过大，用于产生滚动条
                RaiseEvent RequestLine          '行数改变
            Case EN_SELCHANGE                   '通知Rich Edit的父窗口，选择内容改变
                CopyMemory tSC, ByVal lParam, Len(tSC)
                RaiseEvent SelChange(tSC.chrg.cpMin, tSC.chrg.cpMax)
            Case EN_LINK                         '向RichEdit的父窗口发出 点击链接消息
                Result = Subclass1.CallWndProc(Msg, wParam, lParam)
            Case EN_PROTECTED                   '通知RichEdit的父窗口，用户对受保护范围的文本进行操作
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
            Case EN_MSGFILTER                  '通知RichEdit的父窗口，鼠标、键盘事件
                bDefault = True '采用默认处理方法
                CopyMemory tMF, ByVal lParam, Len(tMF)
                Select Case tMF.Msg
                Case WM_KEYDOWN                '当非系统键被KeyPressed
                    iShift = giGetShiftState()
                    iKeyCode = tMF.wParam
                    RaiseEvent KeyDown(iKeyCode, iShift)
                    '屏蔽默认快捷键 ^C/V/X/A ！
                    '如果按 SHIFT 键，则 shift 为 1；如果按 CTRL 键，则 shift 为 2；如果按 ALT 键，则 shift 为 4。
                    If (iShift = 2 And (iKeyCode = vbKeyC Or iKeyCode = vbKeyV Or iKeyCode = vbKeyX Or iKeyCode = vbKeyA Or iKeyCode = vbKeyZ Or iKeyCode = vbKeyY)) Or iKeyCode = 0 Then
                       bDefault = False
                    End If
                Case WM_CHAR                   '快捷键 CTRL+C V X A Z Y
                    iShift = giGetShiftState()
                    iKeyAscii = tMF.wParam
                    RaiseEvent KeyPress(iKeyAscii)
                    If (iShift = 2 And (iKeyAscii = vbKeyC Or iKeyAscii = vbKeyV Or iKeyAscii = vbKeyX Or iKeyAscii = vbKeyA Or iKeyAscii = vbKeyZ Or iKeyAscii = vbKeyY)) Or iKeyAscii = 0 Then
                       bDefault = False
                    End If
                Case WM_KEYUP                  '按键按下
                    iShift = giGetShiftState()
                    iKeyCode = tMF.wParam
                    RaiseEvent KeyUp(iKeyCode, iShift)
                    If (iShift = 2 And (iKeyCode = vbKeyC Or iKeyCode = vbKeyV Or iKeyCode = vbKeyX Or iKeyCode = vbKeyA Or iKeyCode = vbKeyZ Or iKeyCode = vbKeyY)) Or iKeyCode = 0 Then
                       bDefault = False
                    End If
                Case Else
                    Result = Subclass1.CallWndProc(Msg, wParam, lParam)
                End Select
                If Not bDefault Then
                    Result = 1&
                End If
            End Select
        End If
    End Select
'    Dim tNMH As NMHDR
'    Dim tSC As SelChange
'    Dim tMF As MSGFILTER
'    Dim tPR As ENPROTECTED
'    Dim x As Single, y As Single
'    Dim iKeyCode As Integer, iKeyAscii As Integer, iShift As Integer
'    Dim bDefault As Boolean
'    Dim bDoIt As Boolean
'    Dim iNotifyMsg As Long
'    Dim lLen As Long
'    Dim sText As String
'
'    Select Case Msg
'    Case WM_COMMAND                             '当菜单项或下接菜单被点击
'        Result = Subclass1.CallWndProc(Msg, wParam, lParam)
'    Case WM_NOTIFY                              '系统通知
'        CopyMemory tNMH, ByVal lParam, Len(tNMH)
'        If (tNMH.hwndFrom = m_hWndRTB) Then
'            Select Case tNMH.code
'            Case EN_PROTECTED                   '通知RichEdit的父窗口，用户对受保护范围的文本进行操作
'                CopyMemory tPR, ByVal lParam, Len(tPR)
'                bDoIt = False
'                If mvarForceEdit Then
'                    Result = 0
'                Else
'                    RaiseEvent ModifyProtected(bDoIt, tPR.chrg.cpMin, tPR.chrg.cpMax, tPR.wPad1, giGetShiftState)
'                    If bDoIt Then
'                       Result = 0
'                    Else
'                       Result = 1
'                    End If
'                End If
'            Case Else                  '通知RichEdit的父窗口，鼠标、键盘事件
'                Result = Subclass1.CallWndProc(Msg, wParam, lParam)
'            End Select
'        End If
'    Case Else
'        Result = Subclass1.CallWndProc(Msg, wParam, lParam)
'    End Select
End Sub

