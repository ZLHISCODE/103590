VERSION 5.00
Begin VB.UserControl PatiAddress 
   BackColor       =   &H80000005&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   4920
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Text            =   "详细地址"
      Top             =   30
      Width           =   1417
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   3
      Left            =   2625
      TabIndex        =   3
      Text            =   "乡(镇)"
      Top             =   30
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Text            =   "县(区)"
      Top             =   30
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   1
      Left            =   850
      TabIndex        =   1
      Text            =   "市"
      Top             =   30
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Text            =   "省(区,市)"
      Top             =   30
      Width           =   945
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuCopyAll 
         Caption         =   "复制完整地址"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPopuMenuCopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPopuMenuPasteAll 
         Caption         =   "粘贴完整地址"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuPopuMenuPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPopuMenuDelete 
         Caption         =   "清空地址"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "PatiAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum enum_txtInfo
    txt省 = 0
    txt市 = 1
    txt区县 = 2
    txt乡镇 = 3
    txt详细地址 = 4
End Enum

Public Enum enum_Items
    One = 1
    Two = 2
    Three = 3
    Four = 4
    Five = 5
End Enum

Public Enum enum_Style
    TextBox = 0
    Underline = 1
End Enum

Private Type ItemInfo
    strInfo As String '匹配的地址名称
    strCode As String '匹配的地址编码
    strNullInfo As String '没有输入时默认现实
    strStName  As String '标准名称
    bln匹配 As Boolean '是否经过匹配检验
    bln虚拟 As Boolean '是否是虚拟地址
    bln不显示 As Boolean '是否是虚拟不显示内容的地址
    bln无效 As Boolean '是否未使用
    bln隐藏 As Boolean '是否隐藏输入
End Type

'属性变量
Private mstrTag As String
Private mblnShowTown As Boolean 'ShowTown属性
Private mblnLocked As Boolean 'ControlLock属性
Private mcolForeColor As OLE_COLOR
Private mEnumStyle As enum_Style
Private mEnumItemCount As enum_Items
Private mtxtBackColor As OLE_COLOR

'内部变量
Private marrItems(4) As ItemInfo
Private mstrLike As String
Private mblnLike As Boolean
Private mblnFocus As Boolean
Private mblnCancel As Boolean
Private mblnResize As Boolean '防止循环调用Resize
Private mblnSetItems As Boolean '防止循环触发TxtInfo_change
Private mblnChange As Boolean
Private mstrOldAddress As String
Private mblnChangeOld As Boolean '是否修改老地址
Private mblnEdit As Boolean    '是否编辑成功
Private mblnLineFeed As Boolean '详细地址是否换行显示

Public Event Change()
Public Event SetEdit(blnEdit As Boolean)
Public Event SetInput(ByVal intLevel As Integer, rsReturn As ADODB.Recordset)

'==============================================================
'===自定义控件属性
'==============================================================
'hwnd:窗口句柄
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
'Style:控件外观样式
Public Property Get Style() As enum_Style
    Style = mEnumStyle
End Property

Public Property Let Style(ByVal vNewValue As enum_Style)
    Dim i As Long
    mEnumStyle = vNewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).BorderStyle = IIf(vNewValue = 0, 1, 0)
    Next
    Call UserControl_Resize
    PropertyChanged "Style"
End Property
'Enabled:控件可用状态
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Dim i As Long
    UserControl.Enabled = NewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Enabled = NewValue
        txtInfo(i).BackColor = IIf(NewValue, vbWindowBackground, vbButtonFace)
    Next
    If Not NewValue Then Me.ControlLock = True
    PropertyChanged "Enabled"
End Property

'ControlLock:控件的Lock状态
Public Property Get ControlLock() As Boolean
    ControlLock = mblnLocked
End Property

Public Property Let ControlLock(ByVal NewValue As Boolean)
    Dim i As Long
    mblnLocked = NewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Locked = NewValue
        txtInfo(i).TabStop = Not NewValue
        If txtInfo(i).Enabled Then txtInfo(i).BackColor = IIf(NewValue, vbButtonFace, vbWindowBackground)
    Next
    PropertyChanged "ControlLock"
End Property

Public Property Get LineFeed() As Boolean
    LineFeed = mblnLineFeed
End Property

Public Property Let LineFeed(ByVal NewValue As Boolean)
    Dim i As Long
    mblnLineFeed = NewValue
    If Items() = Four Or Items() = Five Then
        Call UserControl_Resize
    End If
    PropertyChanged "LineFeed"
End Property
'Font:控件字体
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        Set txtInfo(i).Font = New_Font
    Next
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property
'Tag:存储控件相关的额外数据
Public Property Get Tag() As String
    Tag = mstrTag
End Property

Public Property Let Tag(ByVal vNewValue As String)
    mstrTag = vNewValue
    PropertyChanged "Tag"
End Property
'ForeColor:控件内文字颜色
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mcolForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).ForeColor() = New_ForeColor
    Next
    mcolForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'TextBackColor:输入框的背景颜色
Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = mtxtBackColor
End Property

Public Property Let TextBackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Long
    mtxtBackColor = New_BackColor
    For i = 0 To txtInfo.Count - 1
       txtInfo(i).BackColor = New_BackColor
    Next
    PropertyChanged "TextBackColor"
End Property
'BackColor:控件的背景颜色
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
'MaxLength:单个单元格允许输入的最大长度
Public Property Get MaxLength() As Long
    MaxLength = txtInfo(txt省).MaxLength
End Property

Public Property Let MaxLength(ByVal vNewValue As Long)
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).MaxLength = vNewValue
    Next
    PropertyChanged "MaxLength"
End Property
'Items:控件展示的项目个数
Public Property Get Items() As enum_Items
    Items = IIf(mEnumItemCount = 0, enum_Items.Four, mEnumItemCount)
End Property

Public Property Let Items(ByVal vNewValue As enum_Items)
    Dim i As Long, lngCount As Long
    
    For i = 0 To txt详细地址
        marrItems(i).bln无效 = Not (i < vNewValue)
    Next
    If vNewValue = Four Then
        marrItems(txt详细地址).bln无效 = False
        marrItems(txt乡镇).bln无效 = True
    End If
    mEnumItemCount = vNewValue
    PropertyChanged "Items"
    If vNewValue = Five And Not mblnShowTown Or vNewValue <> Four And mblnShowTown Then
        mblnShowTown = (vNewValue = Five)
        PropertyChanged "ShowTown"
    End If
    Call UserControl_Resize
    mblnChange = False
End Property
'ShowTown:是否展示乡镇级，Items=4时起作用
Public Property Get ShowTown() As Boolean
    ShowTown = mblnShowTown
End Property

Public Property Let ShowTown(ByVal vNewValue As Boolean)
    '5级与四级且显示乡镇都是5级地址，低于四级的地址，该属性永远为False
    Dim i As Integer

    If mEnumItemCount = Four And vNewValue Or mEnumItemCount = Five And Not vNewValue Then
        mEnumItemCount = IIf(vNewValue, Five, Four)
    ElseIf vNewValue <> (mEnumItemCount = Five) Then '五级地址
        mblnShowTown = mEnumItemCount = Five
    Else
        mblnShowTown = vNewValue
    End If
    For i = 0 To txt详细地址
        marrItems(i).bln无效 = Not (i < Me.Items)
    Next
    If Me.Items = Four Then
        marrItems(txt详细地址).bln无效 = False
        marrItems(txt乡镇).bln无效 = True
    End If
    PropertyChanged "ShowTown"
    PropertyChanged "Items"
    Call UserControl_Resize
End Property
'value:控件的值
Public Property Get value() As String
    value = Me.value省 & Me.value市 & Me.value区县 & Me.value乡镇 & Me.value详细地址
End Property

Public Property Let value(ByVal vNewValue As String)
    Call LoadAllAdress(vNewValue, Me.Items)
    PropertyChanged "value"
    Call UserControl_Resize
End Property
'value省:省级(第一级)地址的值
Public Property Get value省() As String
    value省 = marrItems(txt省).strInfo
End Property

'value市:市级(第二级)地址的值
Public Property Get value市() As String
    value市 = IIf(marrItems(txt市).bln不显示 Or marrItems(txt市).bln无效, "", marrItems(txt市).strInfo)
End Property

'value区县:区县级(第三级)地址的值
Public Property Get value区县() As String
    value区县 = IIf(marrItems(txt区县).bln不显示 Or marrItems(txt区县).bln无效, "", marrItems(txt区县).strInfo)
End Property

'value乡镇:乡镇级(第四级)地址的值
Public Property Get value乡镇() As String
    value乡镇 = IIf(marrItems(txt乡镇).bln不显示 Or marrItems(txt乡镇).bln无效, "", marrItems(txt乡镇).strInfo)
End Property

'value详细地址:最后级(第五级)地址的值
Public Property Get value详细地址() As String
    value详细地址 = IIf(marrItems(txt详细地址).bln不显示 Or marrItems(txt详细地址).bln无效, "", marrItems(txt详细地址).strInfo)
End Property

'Code:标准化地址单元格中最小一级地址对应的编码
Public Property Get Code() As String
    Dim i As Integer, strTmp As String
    For i = txt详细地址 To 0 Step -1
        strTmp = marrItems(i).strCode
        If strTmp <> "" Then Exit For
    Next
    Code = strTmp
End Property
'AllCodes:所有地址单元格对应编码，以逗号分割
Public Property Get AllCodes() As String
    AllCodes = marrItems(txt省).strCode & "," & marrItems(txt市).strCode & "," & marrItems(txt区县).strCode & "," & marrItems(txt乡镇).strCode & "," & marrItems(txt详细地址).strCode
End Property

'==============================================================
'===控件方法
'==============================================================
Public Sub LoadAllAdress(ByVal strAdress As String, Optional ByVal intType As Integer)
    Call StructAdress(strAdress, intType)
End Sub

Public Sub LoadStructAdress(ByVal str省 As String, ByVal str市 As String, ByVal str区县 As String, ByVal str乡镇 As String, ByVal str详细地址 As String, Optional ByVal intType As Integer)
   Call StructAdress(str省 & "," & str市 & "," & str区县 & "," & str乡镇 & "," & str详细地址, intType)
End Sub

Public Function CheckNullValue(Optional ByVal blnNotCheck详细地址 As Boolean = True, Optional ByVal blnOnlyChangeCheck As Boolean, Optional ByVal blnMustInput As Boolean) As String
'功能：存在数据时进行空值检查，保证按照顺序输入。
'参数：blnOnlyChangeCheck=只有变化才检查,为空，且必须输入时，该参数对于识别老地址有问题，请慎用
'          blnMustInput=是否必须输入
'          blnNotCheck详细地址=详细地址没输是否不检查,
'说明：
    Dim i As Long, blnNull As Boolean
    Dim blnCheck As Boolean
    
    If Me.value = "" And blnMustInput Then
        blnCheck = True
    ElseIf Me.value <> "" Then
        If blnOnlyChangeCheck Then
            blnCheck = mstrOldAddress <> Me.value
        Else
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        For i = 0 To txt详细地址
            If Not marrItems(i).bln无效 And txtInfo(i).Visible Then
                If marrItems(i).strInfo = "" Then
                    If Not (i = txt市 And InStr(marrItems(0).strInfo, "市") > 0) Then
                        If i = txt详细地址 And Not blnNotCheck详细地址 Or i <> txt详细地址 And i <> txt乡镇 Then
                            CheckNullValue = marrItems(i).strNullInfo
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    End If
End Function

Public Function CheckDefrentValue(ByVal NowAdress As String, Optional ByVal PreAdress As String = "") As Boolean
'功能：身份证地址与录入的地址进行校验
'参数：NowAdress=数据库读取出来的户口地址信息
'          PreAdress=界面上录入的地址信息
    Dim i As Integer
    Dim strPatiAdress As String
    If Trim(NowAdress) = "" Then
        CheckDefrentValue = True
    ElseIf Trim(NowAdress) <> "" Then
        If Me.Items >= Four Then
            strPatiAdress = Trim(NowAdress)
            Me.value = PreAdress
            If Me.value = strPatiAdress Then
                CheckDefrentValue = True
            Else
                CheckDefrentValue = False
            End If
            Me.value = strPatiAdress
        End If
    End If
    Exit Function
End Function

'==============================================================
'===自定义控件事件
'==============================================================
Private Sub UserControl_GotFocus()
    Set gobjPati = Me
    If txtInfo(txt省).Enabled Then Call txtInfo(txt省).SetFocus
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    marrItems(txt省).strNullInfo = "省(区,市)"
    marrItems(txt市).strNullInfo = "市"
    marrItems(txt区县).strNullInfo = "县(区)"
    marrItems(txt乡镇).strNullInfo = "乡(镇)"
    marrItems(txt详细地址).strNullInfo = "详细地址"
End Sub

Private Sub UserControl_InitProperties()
    If mEnumItemCount = 0 Then
        mEnumItemCount = Four
        mblnShowTown = False
        marrItems(txt乡镇).bln无效 = True
    End If
    mEnumStyle = enum_Style.TextBox
    mtxtBackColor = &H80000005
    mcolForeColor = &H80000000
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If Me.value <> "" Then
            Call mnuPopuMenuDelete_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.value <> "" Then
            Call mnuPopuMenuCopyAll_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        If Clipboard.GetText <> "" Then
            Call mnuPopuMenuPasteAll_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
            Call mnuPopuMenuPasteAll_Click
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mblnResize = True
    Me.Style = PropBag.ReadProperty("Style", enum_Style.TextBox)
    Me.Items = PropBag.ReadProperty("Items", enum_Items.Four)
    Me.ShowTown = PropBag.ReadProperty("ShowTown", Me.ShowTown)  '因为Items属性与该属性相关，因此默认值为Me.ShowTown
    Me.ControlLock = PropBag.ReadProperty("ControlLock", False)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.TextBackColor = PropBag.ReadProperty("TextBackColor", &H80000005)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000000)
    Me.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Me.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Me.value = PropBag.ReadProperty("value", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
    mblnResize = False: mblnChangeOld = True
    Me.LineFeed = PropBag.ReadProperty("LineFeed", False)
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
'功能: 设置控件大小位置
    On Error Resume Next
    Dim i As Long, intPreVisual As Integer
    Dim lngHeight As Long, lngPerWidth As Long, lngMinWidth As Long, lngMinHeight As Long
    Dim arrWidthShare As Variant, lngTotal As Double
    Dim lngCount As Long, lngLastItem As Long
    Dim lngDisH As Long, lngDisV As Long

    If mblnResize Then Exit Sub
    mblnResize = True
    For i = 0 To txt详细地址
        txtInfo(i).Visible = Not marrItems(i).bln无效
    Next
    If Me.Items = Two Then
        txtInfo(txt市).Visible = marrItems(txt区县).bln无效
    End If
    lngMinWidth = UserControl.TextWidth("省(区,市)") + 60
    '各个文本框宽度比例，以四个项目为计算标准，五级地址时通过第四级拆分
    arrWidthShare = Array(1, 1, 1, 2.5)
    lngCount = Me.Items
    lngLastItem = lngCount - 1
    If Me.Items = Four Then lngLastItem = 4
    If Me.Items = Five Then lngCount = 4
    For i = 0 To lngCount - 1
        If mblnLineFeed And Me.Items <> Five Then
            If i <> 3 And i <> 4 Then
                lngTotal = lngTotal + arrWidthShare(i)
            End If
        Else
            lngTotal = lngTotal + arrWidthShare(i)
        End If
    Next
    lngDisH = 0: lngDisV = 0
    lngPerWidth = (UserControl.Width - (lngCount + 1) * lngDisH) / lngTotal
    If lngPerWidth < lngMinWidth Then lngPerWidth = lngMinWidth
    lngMinHeight = UserControl.TextHeight("中")
    If UserControl.Height >= txtInfo(txt省).Height And mblnLineFeed And lngCount = 4 Then
        lngHeight = UserControl.Height / 2
    Else
        lngHeight = UserControl.Height - IIf(Me.Style = Underline, 30, 0)
    End If
    If lngHeight < lngMinHeight Then lngHeight = lngMinHeight
    '控件位置摆正
    For i = 0 To txtInfo.Count - 1
        If i = 0 Then
            txtInfo(i).Move lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
        ElseIf i < lngCount Then
            If mblnLineFeed And Me.Items <> Five Then
                If i <> 3 And i <> 4 Then
                    txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
                End If
            Else
                txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
            End If
        Else '不展示的控件也需要摆正位置，防止划线出问题
            If Me.Items = Four Or Me.Items = Five Then
                If mblnLineFeed Then
                    If i = txt详细地址 Then
                        txtInfo(i).Move lngDisH, lngDisV + lngHeight, lngPerWidth * 2.5, lngHeight
                    Else
                        txtInfo(i).Move lngDisH, lngDisV + lngHeight, lngPerWidth * 1, lngHeight
                    End If
                Else
                    txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * 1, lngHeight
                End If
            End If
        End If
    Next
    If Me.Items = Four Then
        If mblnLineFeed Then
            txtInfo(txt详细地址).Move lngDisH, lngDisV + lngHeight, txtInfo(txt省).Width + txtInfo(txt市).Width + txtInfo(txt区县).Width, lngHeight
        Else
            txtInfo(txt详细地址).Move txtInfo(txt区县).Left + txtInfo(txt区县).Width + lngDisH, lngDisV, txtInfo(txt乡镇).Width, lngHeight
        End If
    ElseIf Me.Items = Five Then
        '乡镇与详细地址比例为1:1.5
        lngPerWidth = (txtInfo(txt乡镇).Width - lngDisH) / (1 + 1.5)
        If mblnLineFeed Then
            txtInfo(txt乡镇).Width = lngPerWidth * 2.5
        Else
            txtInfo(txt乡镇).Width = lngPerWidth * 1
        End If
        If mblnLineFeed Then
            txtInfo(txt详细地址).Move lngDisH, lngDisV + lngHeight, txtInfo(txt省).Width + txtInfo(txt市).Width + txtInfo(txt区县).Width + txtInfo(txt乡镇).Width, lngHeight
        Else
            txtInfo(txt详细地址).Move txtInfo(txt乡镇).Left + txtInfo(txt乡镇).Width + lngDisH, lngDisV, lngPerWidth * 1.5, lngHeight
        End If
    ElseIf Me.Items = Two Then
        If txtInfo(txt市).Visible Then
            txtInfo(txt区县).Move txtInfo(txt市).Left + txtInfo(txt市).Width, 0, txtInfo(txt市).Width, txtInfo(txt市).Height
            txtInfo(txt市).ZOrder
        Else
            txtInfo(txt区县).Move txtInfo(txt市).Left, 0, txtInfo(txt市).Width, txtInfo(txt市).Height
            txtInfo(txt区县).ZOrder
        End If
    End If
    If mblnLineFeed Then
        If Me.Items = Four Or Me.Items = Five Then
            UserControl.Height = txtInfo(txt省).Top + txtInfo(txt省).Height * 2 + IIf(Me.Style = Underline, 30, 0)
            If Me.Items = Four Then
                UserControl.Width = txtInfo(txt省).Width + txtInfo(txt市).Width + txtInfo(txt区县).Width
            Else
                UserControl.Width = txtInfo(txt省).Width + txtInfo(txt市).Width + txtInfo(txt区县).Width + txtInfo(txt乡镇).Width
            End If
        Else
            UserControl.Width = txtInfo(lngLastItem).Left + txtInfo(lngLastItem).Width + lngDisH
            UserControl.Height = txtInfo(txt省).Top + txtInfo(txt省).Height + IIf(Me.Style = Underline, 30, 0)
        End If
    Else
        UserControl.Height = txtInfo(txt省).Top + txtInfo(txt省).Height + IIf(Me.Style = Underline, 30, 0)
        UserControl.Width = txtInfo(lngLastItem).Left + txtInfo(lngLastItem).Width + lngDisH
    End If
    UserControl.Refresh
    Call SetLine(Me.Style)
    mblnResize = False
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", mblnLocked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", mtxtBackColor, &H80000005)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Items", Me.Items, enum_Items.Four)
    Call PropBag.WriteProperty("ShowTown", Me.ShowTown, mblnShowTown)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(txt省).MaxLength, 0)
    Call PropBag.WriteProperty("value", Me.value, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
    Call PropBag.WriteProperty("LineFeed", mblnLineFeed, False)
End Sub
'==============================================================
'===自定义控件内部空间事件
'==============================================================
Private Sub mnuPopuMenuCopyAll_Click()
    Dim i As Long, strAdress As String
    Dim strTmp As String
    
    Clipboard.Clear
    For i = 0 To txt详细地址
        strTmp = marrItems(i).strInfo & "," & marrItems(i).strCode & "," & IIf(marrItems(i).bln虚拟, 1, 0) & "," & _
                        IIf(marrItems(i).bln不显示, 1, 0) & ",," & IIf(marrItems(i).bln匹配, 1, 0)
        strAdress = strAdress & IIf(strAdress = "", "", "|") & strTmp
    Next
    strAdress = "ZLSOFT:" & strAdress
    Clipboard.SetText strAdress
End Sub

Private Sub mnuPopuMenuCopy_Click()
    Clipboard.Clear
    If Not UserControl.ActiveControl Is Nothing Then
         Clipboard.SetText UserControl.ActiveControl.SelText
    End If
End Sub

Private Sub mnuPopuMenuDelete_Click()
    Dim i As Long, intCur As Integer
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    intCur = -1
    If Not UserControl.ActiveControl Is Nothing Then
        intCur = UserControl.ActiveControl.Index
    End If
    For i = 0 To txt详细地址
        txtInfo(i).Text = ""
        Call FillItems(i, , i = intCur)
    Next
End Sub

Private Sub mnuPopuMenuPasteAll_Click()
    Dim i As Long, intCur As Integer
    Dim strTmp As String
    
    
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    intCur = -1
    If Not UserControl.ActiveControl Is Nothing Then
        intCur = UserControl.ActiveControl.Index
    End If
    strTmp = Clipboard.GetText
    If zlCommFun.ActualLen(strTmp) > 500 Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
        Else
            strTmp = SubB(strTmp, 1, 500)
        End If
    End If
    mblnChangeOld = False
    Call StructAdress(strTmp)
    For i = 0 To txt详细地址
        Call FillItems(i, , i = intCur)
    Next
    mblnChangeOld = True
End Sub

Private Sub mnuPopuMenuPaste_Click()
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    If Not UserControl.ActiveControl Is Nothing Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
            Call mnuPopuMenuPasteAll_Click
        Else
            Call SendMessage(UserControl.ActiveControl.hWnd, WM_PASTE, 0, 0)
        End If
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If mblnSetItems Then Exit Sub
    marrItems(Index).strInfo = txtInfo(Index).Text
    Call ClearItems(Index)
    RaiseEvent Change
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Set gobjPati = Me
    Call FillItems(Index)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = 13 Or ((KeyCode = vbKeySpace Or Chr(KeyCode) = "*") And txtInfo(Index).Text = "") Then
        '输入区域数据
       
       If (txtInfo(Index).Tag = marrItems(Index).strCode And marrItems(Index).bln匹配 Or txtInfo(Index).Text = "") And KeyCode = 13 Then
            KeyCode = 0
            If txtInfo(Index).Text = "" And Not marrItems(Index).bln匹配 Then
                Call ClearItems(Index)
            End If
            Call LocateItem(Index, 1, marrItems(Index).bln匹配)
            Exit Sub
        Else
            KeyCode = 0
            Call SetInput(Index)
        End If
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
        KeyCode = 0
        Call LocateItem(Index, 1)
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        KeyCode = 0
        Call LocateItem(Index, -1)
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        gblnCanPaste = txtInfo(Index).Enabled And Not txtInfo(Index).Locked
        If gblnCanPaste And Clipboard.GetText Like "ZLSOFT:*" Then gblnCanPaste = False
        If glngTXTProc = 0 Then
            glngTXTProc = GetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC)
            Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, AddressOf WndMessagePaste)
        End If
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        txtInfo(Index).Text = ""
    End If
End Sub

Private Sub txtInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = vbCtrlMask Then
        If glngTXTProc <> 0 Then
            Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, glngTXTProc)
            glngTXTProc = 0
        End If
    End If
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = 0
    Call FillItems(Index, , False)
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtMouseDown(txtInfo(Index), Button, Shift, X, Y)
End Sub

Private Sub txtInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInfo(Index).ToolTipText = txtInfo(Index).Text
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtMouseUp(txtInfo(Index), Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    If Not marrItems(Index).bln匹配 And txtInfo(Index).Text <> marrItems(Index).strNullInfo And txtInfo(Index).Text <> "" Then
        Cancel = SetInput(Index)
    End If
End Sub
'==============================================================
'===内部方法
'==============================================================
Private Function SetInput(ByVal intIndex As Integer, Optional ByVal strInputCode As String, Optional ByRef strPreCode As String, Optional ByVal strName As String, Optional ByVal blnClare As Boolean = True) As Boolean
'功能：根据输入内容设置文本框内容
' 参数：intIndex:进行处理的控件
'          strInputCode:="",对当前单元格进行输入匹配，<>"" 对当前单元格进行精确查找（当传入的strPreCode与strName不为空，则不会进行查找，直接加载）
'          strPreCode=上级编码
'          strName=当前单元格匹配到的名称
'          blnClare=是否清除变动项目
' 返回：
'         strPreCode=上级编码
'         是否禁止光标移动
    Dim intPreIndex As Integer
    Dim strTmpSQL As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String, strCode As String, i As Integer
    Dim vPoint As POINTAPI, blnCancel As Boolean
    Dim intLevel As Integer
    
    If strInputCode = "" Then
        intPreIndex = GetSetableItem(intIndex, -1)
        strInput = Trim(txtInfo(intIndex).Text)
        If intPreIndex >= 0 Then strCode = marrItems(intPreIndex).strCode
        intLevel = intIndex
        If strCode = "" And strInput = "" And (intIndex = txt详细地址 Or intIndex = txt乡镇) Then Exit Function
        '乡镇不输入时，在详细地址输入
        If intIndex = txt详细地址 And marrItems(txt乡镇).bln无效 Then
            intLevel = txt乡镇
            If intPreIndex + 1 = intLevel Then
                strTmpSQL = IIf(strCode <> "", " And B.上级编码=[4]", "")
            Else
                strTmpSQL = IIf(strCode <> "", " And B.上级编码 In(Select D.编码 From 区域 D Where D.上级编码=[4])", "")
            End If
            
            If strInput <> "" Then
                If zlCommFun.IsCharChinese(strInput) Then
                    strTmpSQL = strTmpSQL & " And Nvl(a.简码,b.简码) Like [1] "
                ElseIf IsNumeric(strInput) Then
                    strTmpSQL = strTmpSQL & " And Nvl(a.简码,b.简码) Like [1]  "
                Else
                    strTmpSQL = strTmpSQL & " And Nvl(a.简码,b.简码) Like [1] "
                End If
            End If
            strSQL = "Select Rownum as ID,Nvl(a.编码,b.编码) 编码, b.名称 || a.名称 名称, Nvl(a.简码,b.简码) 简码, b.上级编码, a.是否虚拟, a.是否不显示" & vbNewLine & _
                            "From 区域 a, 区域 b" & vbNewLine & _
                            "Where a.上级编码(+) = b.编码 And NVL(B.级数,0)=[3] " & strTmpSQL & " Order by Nvl(a.编码,b.编码)"
        Else
            If intPreIndex + 1 = intLevel Then
                strTmpSQL = IIf(strCode <> "", " And A.上级编码=[4]", "")
            Else
                strTmpSQL = IIf(strCode <> "", " And A.上级编码 In(Select 编码 From 区域  Where 上级编码=[4])", "")
            End If
            If strInput <> "" Then
                If zlCommFun.IsCharChinese(strInput) Then
                    strTmpSQL = strTmpSQL & " And A.名称 Like [1] "
                ElseIf IsNumeric(strInput) Then
                    strTmpSQL = strTmpSQL & " And A.编码 Like [1]  "
                Else
                    strTmpSQL = strTmpSQL & " And A.简码 Like [1] "
                End If
            End If
            
            strSQL = "Select Rownum as ID,编码,名称,简码,上级编码,是否虚拟,是否不显示,邮编  From 区域 A " & _
                            "Where NVL(A.级数,0)=[3]" & strTmpSQL & " Order by 编码"
        End If
        If mblnLike = False Then mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", ""): mblnLike = True
        vPoint = GetCoordPos(txtInfo(txt省).hWnd, txtInfo(txt省).Left, txtInfo(txt省).Top)
        For i = 1 To intIndex
            If txtInfo(i - 1).Visible Then vPoint.X = vPoint.X + txtInfo(i - 1).Width
        Next
        Set rsTmp = zlDatabase.ShowSQLSelect(UserControl.Parent, strSQL, 0, "区域", False, "", "", False, _
            False, True, vPoint.X, vPoint.Y, txtInfo(intIndex).Height, blnCancel, False, False, _
            UCase(strInput) & "%", mstrLike & UCase(strInput) & "%", intLevel, strCode)
        '可以任意输入,不一定要匹配
        If Not rsTmp Is Nothing Then
            mblnEdit = True
            Call FillItems(intIndex, rsTmp, blnClare)
            strPreCode = rsTmp!上级编码 & ""
            '自动补缺
            strCode = strPreCode: strPreCode = ""
            For i = intIndex - 1 To 0 Step -1
                If Not marrItems(i).bln无效 Then
                    Call SetInput(i, strCode, strPreCode, False)
                    strCode = strPreCode: strPreCode = ""
                    If strCode = "" Then Exit For
                End If
            Next
            If Not SetNoNaturalAd(intIndex) Then
                If intIndex = txt详细地址 Then
                    txtInfo(txt详细地址).SelLength = 0
                    txtInfo(txt详细地址).SelStart = 0
                    txtInfo(txt详细地址).SelLength = Len(txtInfo(txt详细地址).Text)
                    Exit Function
                End If
                Call LocateItem(intIndex, 1, True)
            End If
        Else
            mblnEdit = False
            Call zlControl.TxtSelAll(txtInfo(intIndex))
            If Not blnCancel And intIndex <> txt详细地址 Then
                If zlCommFun.IsCharChinese(txtInfo(intIndex).Text) Then
                    If MsgBox("字典表中未找到您输入的区域，是否要使用输入的值？", vbQuestion + vbYesNo + vbDefaultButton1, "查找区域") = vbYes Then
                        marrItems(intIndex).bln匹配 = True
                        Call LocateItem(intIndex, 1, True)
                        mblnEdit = True
                    End If
                Else
                    MsgBox "字典表中未找到您输入的区域。", vbInformation, "查找区域"
                    SetInput = True
                End If
            ElseIf intIndex <> txt详细地址 Or blnCancel Then
                Call LocateItem(intIndex, 0, True)
                Call FillItems(intIndex)
            Else
                marrItems(intIndex).bln匹配 = True
                Call LocateItem(intIndex, 1, True)
                mblnEdit = True
            End If
        End If
        If Not rsTmp Is Nothing Then
            RaiseEvent SetInput(intLevel, rsTmp)
        End If
        RaiseEvent SetEdit(mblnEdit)
    Else
        '没有名称则需要匹配
        strSQL = "select 编码,名称,上级编码,是否虚拟,是否不显示 from 区域 where 编码=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "区域精确查找", strInputCode & "")
        If Not rsTmp.EOF Then
            strPreCode = rsTmp!上级编码 & ""
        End If
        Call FillItems(intIndex, rsTmp, False)
    End If
End Function

Private Function SetNoNaturalAd(ByVal intIndex As Integer) As Boolean
'功能：判断一个地址的下级是否没有实际地址，全是虚拟地址
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnOnlyVir As Boolean
    
    If Me.Items < Two Then Exit Function
    If intIndex = txt乡镇 Then Exit Function
    If marrItems(intIndex).strCode <> "" And intIndex < Me.Items - 1 Then
        strSQL = "Select 1 计数 from 区域 Where 上级编码 =[1]  And Nvl(是否虚拟,0)=0 And Nvl(是否不显示,0)=0 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断是否没有有效下级", marrItems(intIndex).strCode)
        If Me.Items = Two Then
            strSQL = "Select 1 计数 from 区域 Where 上级编码 =[1] And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断是否没有有效下级", marrItems(intIndex).strCode)
            If rsTmp.RecordCount > 0 Then
                blnOnlyVir = True
            End If
        Else
            If rsTmp.RecordCount = 0 Then
                strSQL = "Select 1 计数 from 区域 Where 上级编码 =[1] And Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "判断是否没有有效下级", marrItems(intIndex).strCode)
                If rsTmp.RecordCount > 0 Then
                    blnOnlyVir = True
                End If
            End If
        End If
        If Not blnOnlyVir Then
            txtInfo(intIndex + 1).Locked = False
            If Me.Enabled And Not mblnLocked Then txtInfo(intIndex + 1).BackColor = vbWindowBackground
            marrItems(intIndex + 1).bln虚拟 = False
            marrItems(intIndex + 1).bln不显示 = False
             If Me.Items = Two Then  '两级展示时，如果市是虚拟地址或不显示地址,则讲区县移动到市的位置
                Call FillItems(txt市, , False)
             End If
             Call LocateItem(intIndex, 1, True)
        Else
            If Me.Enabled Then txtInfo(intIndex + 1).Locked = True
            If Me.ControlLock = False Then
                txtInfo(intIndex + 1).BackColor = vbButtonFace
            End If
            marrItems(intIndex + 1).bln虚拟 = True
            marrItems(intIndex + 1).bln不显示 = True
            If Me.Items = Two Then  '两级展示时，如果市是虚拟地址或不显示地址,则讲区县移动到市的位置
                Call FillItems(txt区县, , False)
             End If
             Call LocateItem(intIndex + 1, 1, True)
        End If
        SetNoNaturalAd = True
    End If
End Function

Private Function StructAdress(ByVal strInput As String, Optional ByVal intType As Integer) As String
'功能：结构化地址，并读取结构化地址信息。
    Dim rsTmp  As ADODB.Recordset, strSQL As String
    Dim arrAddress As Variant
    Dim arrTmp As Variant
    Dim i As Long, j As Long, blnClare As Boolean
    Dim blnCopyStruct As Boolean
    Dim str乡镇Code As String
    
    If strInput Like "ZLSOFT:*" Then '结构化地址复制
        If strInput Like "*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*" Then
            blnCopyStruct = True
        End If
        strInput = Mid(strInput, Len("ZLSOFT:") + 1)
    End If
    
    If strInput = "" Then
        strInput = ",,,,|,,,,|,,,,|,,,,|,,,,"
    ElseIf Not blnCopyStruct Then
        strSQL = "Select Zl_Adderss_Structure([1],[2]) 地址分解 From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "地址结构分解", strInput, intType)
        strInput = rsTmp!地址分解 & ""
    End If
    blnClare = True
    arrAddress = Split(strInput, "|")
    For i = LBound(arrAddress) To UBound(arrAddress)
        arrTmp = Split(arrAddress(i), ",")
        marrItems(i).strInfo = arrTmp(0)
        marrItems(i).strCode = arrTmp(1)
        marrItems(i).bln虚拟 = Val(arrTmp(2)) = 1
        marrItems(i).bln不显示 = Val(arrTmp(3)) = 1
        If UBound(arrTmp) = 5 Then '地址复制时会有第五个元素
            marrItems(i).bln匹配 = Val(arrTmp(5)) = 1
        Else
            marrItems(i).bln匹配 = True
        End If
        '该级数只有虚拟下级
        If Val(arrTmp(4)) = 1 Then
            marrItems(i).bln虚拟 = True
            marrItems(i).bln不显示 = True
        End If
        If Me.Items = Two And i = txt市 Then
            marrItems(i).bln虚拟 = True
            marrItems(i).bln不显示 = True
            marrItems(txt区县).bln无效 = Not marrItems(i).bln虚拟
        End If
        If Me.Items = Two And i = txt区县 And marrItems(txt区县).strInfo = "" Then
            marrItems(i).strInfo = marrItems(txt市).strInfo
        End If
        If marrItems(i).bln无效 Then
            If i = txt乡镇 Then str乡镇Code = marrItems(i).strCode
            If blnClare Then Call ClearItems(i, False): blnClare = False
        Else
             If i = txt详细地址 Then
                If Me.Items = Four Then '乡镇不显示，讲乡镇合并详细地址
                    marrItems(i).strInfo = marrItems(txt乡镇).strInfo & marrItems(i).strInfo
                    '详细地址没有编码，就取乡镇的编码
                    If marrItems(i).strCode = "" And str乡镇Code <> "" Then
                        marrItems(i).strCode = str乡镇Code
                        marrItems(i).strStName = marrItems(txt乡镇).strInfo
                    End If
                    marrItems(txt乡镇).strInfo = ""
                    Call FillItems(txt乡镇, , False)
                End If
             End If
        End If
        Call FillItems(i, , False)
    Next
    If Me.value = marrItems(txt乡镇).strInfo Then
        If marrItems(txt乡镇).bln无效 Then
            marrItems(txt省).strInfo = marrItems(txt乡镇).strInfo
            marrItems(txt乡镇).strInfo = ""
            Call FillItems(txt省, , False)
        End If
    End If
    If mblnChangeOld Then
        mstrOldAddress = Me.value
    End If
End Function

Private Function GetSetableItem(ByVal intIndex As Integer, Optional ByVal intStep As Integer) As Integer
'功能：搜索可以输入的项目
'参数：intIndex=起始索引
'         intStep=定位方向，0-当前单元格，-1-向前定位，定位到前面最近一个可输入的单元格，1-向后定位，定位到可输入的单元格
'返回：可以定位的单元格
    Dim i As Integer, intReturn As Integer
    
    intReturn = -1
    If intStep = 0 Then
        '当前单元格能否定位，不能定位，则向后寻找
        If Not (marrItems(intIndex).bln虚拟 Or marrItems(intIndex).bln不显示) And txtInfo(intIndex).Visible Then
            intReturn = intIndex
        Else
            intReturn = GetSetableItem(intIndex, 1)
        End If
    ElseIf intStep = -1 Then '想前寻找
        For i = intIndex - 1 To 0 Step -1
            If Not (marrItems(i).bln虚拟 Or marrItems(i).bln不显示) And Not marrItems(i).bln无效 Then intReturn = i: Exit For
        Next
    ElseIf intStep = 1 Then
        For i = intIndex + 1 To txt详细地址
            If Not (marrItems(i).bln虚拟 Or marrItems(i).bln不显示) And Not marrItems(i).bln无效 Then intReturn = i: Exit For
        Next
    End If
    GetSetableItem = intReturn
End Function

Private Function LocateItem(ByVal intIndex As Integer, Optional ByVal intStep As Integer, Optional ByVal blnNotCheckSel As Boolean) As Integer
'功能：功能定位项目
'参数：intIndex=起始索引
'         intStep=定位方向，0-当前单元格，-1-向前定位，定位到前面最近一个可输入的单元格，1-向后定位，定位到可输入的单元格
'返回：可以定位的单元格
    Dim intReturn As Integer
    Dim intStart As Integer, intEnd As Integer
    Dim i As Integer
    intReturn = -1
    If intStep = 0 Then
        intStart = intIndex
        intReturn = GetSetableItem(intIndex)
        intEnd = intReturn - 1
    ElseIf intStep = -1 Then
        intEnd = intIndex - 1
        If txtInfo(intIndex).SelStart = 0 Or blnNotCheckSel Then
            intReturn = GetSetableItem(intIndex, intStep)
        End If
        intStart = intReturn
    ElseIf intStep = 1 Then
        intStart = intIndex
        If txtInfo(intIndex).SelStart = Len(txtInfo(intIndex).Text) Or blnNotCheckSel Then
            intReturn = GetSetableItem(intIndex, intStep)
        End If
        intEnd = intReturn - 1
    End If
    If intReturn <> -1 Then
        txtInfo(intReturn).SetFocus
    ElseIf intStep >= 0 Then
        zlCommFun.PressKey (vbKeyTab)
        intEnd = txt详细地址
    End If
    If intStart >= 0 Then
        For i = intStart To intEnd
            Call FillItems(i, , False)
        Next
    End If
End Function

Private Sub ClearItems(ByVal intIndex As Integer, Optional ByVal blnLocate As Boolean = True)
    Dim i As Integer, intEnd As Integer
    For i = intIndex To txt详细地址
        marrItems(i).strCode = ""
        marrItems(i).bln虚拟 = False
        marrItems(i).bln不显示 = False
        If Me.Items = Two And i = txt市 Then
            marrItems(i).bln虚拟 = True
            marrItems(i).bln不显示 = True
        End If
        marrItems(i).bln匹配 = False
        Call FillItems(i, , blnLocate And i = intIndex)
    Next
End Sub

Private Sub FillItems(ByVal intIndex As Integer, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnLocate As Boolean = True)
    mblnSetItems = True
    '根据记录集填充数据
    If Not rsInput Is Nothing Then
        marrItems(intIndex).strInfo = rsInput!名称 & ""
        marrItems(intIndex).strCode = rsInput!编码 & ""
        marrItems(intIndex).bln虚拟 = Val(rsInput!是否虚拟 & "") = 1
        If Me.Items = Two And intIndex = txt市 Then
            marrItems(intIndex).bln虚拟 = True
        End If
        marrItems(intIndex).bln不显示 = Val(rsInput!是否不显示 & "") = 1
        marrItems(intIndex).bln匹配 = True
    End If
    If intIndex > txt详细地址 Then Exit Sub
    '二级地址第二级为虚拟时的特殊处理
    If Me.Items = Two Then
        marrItems(txt区县).bln无效 = Not marrItems(txt市).bln虚拟
        If marrItems(txt区县).bln无效 Then
            If txtInfo(txt区县).Visible Then marrItems(txt区县).strInfo = ""
        Else
            If txtInfo(txt市).Visible Then marrItems(txt市).strInfo = ""
        End If
        Call UserControl_Resize
    End If
    '填充界面数据，并设置展示样式
    txtInfo(intIndex).Text = marrItems(intIndex).strInfo
    txtInfo(intIndex).Tag = marrItems(intIndex).strCode
    If blnLocate Then
        If txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo Then txtInfo(intIndex).Text = ""
        txtInfo(intIndex).ForeColor = &H80000008
    Else
        If txtInfo(intIndex).Text = "" Then txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo
        If txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo Then
            txtInfo(intIndex).ForeColor = mcolForeColor
            txtInfo(intIndex).SelStart = 0
            txtInfo(intIndex).SelLength = 0
        Else
            txtInfo(intIndex).ForeColor = &H80000008
        End If
    End If
    If marrItems(intIndex).bln无效 Then
        txtInfo(intIndex).BackColor = vbButtonFace
    Else
        If marrItems(intIndex).bln虚拟 Then
            If marrItems(intIndex).bln不显示 Then
                txtInfo(intIndex).Text = ""
            End If
            txtInfo(intIndex).Enabled = Me.Enabled And Not marrItems(intIndex).bln不显示
            txtInfo(intIndex).Locked = Not Me.Enabled
            txtInfo(intIndex).BackColor = vbButtonFace
        Else
            txtInfo(intIndex).Enabled = Me.Enabled
            txtInfo(intIndex).Locked = mblnLocked Or Not Me.Enabled
            If Me.Enabled Then
                txtInfo(intIndex).BackColor = IIf(mblnLocked, vbButtonFace, Me.TextBackColor)
            Else
                txtInfo(intIndex).BackColor = vbButtonFace
            End If
        End If
    End If
    mblnSetItems = False
End Sub

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Sub SetLine(ByVal lngStyle As Long)
'功能：设置下划线
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    UserControl.Cls
    For i = 0 To txtInfo.Count - 1
        If lngStyle = Underline Then
            x1 = txtInfo(i).Left
            y1 = txtInfo(i).Top + txtInfo(i).Height
            x2 = txtInfo(i).Left + txtInfo(i).Width - 30
            y2 = y1
            UserControl.Line (x1, y1)-(x2, y2)
        End If
    Next
End Sub

Private Sub TxtMouseDown(ByRef ObjText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：TextBox的默认右键消息提示的修改
    If intButton = vbRightButton Then
        If glngTXTProc = 0 Then
            glngTXTProc = GetWindowLong(ObjText.hWnd, GWL_WNDPROC)
            Call SetWindowLong(ObjText.hWnd, GWL_WNDPROC, AddressOf WndMessageMenu)
        End If
    End If
End Sub

Private Sub TxtMouseUp(ByRef ObjText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'功能：TextBox的默认右键消息提示的修改
    If intButton = vbRightButton Then
        If glngTXTProc <> 0 Then
            Call SetWindowLong(ObjText.hWnd, GWL_WNDPROC, glngTXTProc)
            glngTXTProc = 0
        End If
    End If
End Sub

Friend Sub PopMenu()
    Dim strTxt As String
    mnuPopuMenuCopyAll.Enabled = Me.value <> ""
    mnuPopuMenuDelete.Enabled = Me.value <> "" And Me.Enabled And Not Me.ControlLock
    strTxt = Clipboard.GetText
    mnuPopuMenuPasteAll.Enabled = strTxt <> "" And Me.Enabled And Not Me.ControlLock
    mnuPopuMenuPaste.Enabled = strTxt <> "" And Me.Enabled And Not Me.ControlLock
    PopupMenu mnuPopuMenu
End Sub

