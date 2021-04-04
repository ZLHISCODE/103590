VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl ComboxExpend 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   5400
   ToolboxBitmap   =   "ComboxExpend.ctx":0000
   Begin VB.PictureBox picDownList 
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   1605
      Picture         =   "ComboxExpend.ctx":0312
      ScaleHeight     =   2340
      ScaleWidth      =   2265
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   825
      Width           =   2265
      Begin MSComctlLib.TreeView tvwList 
         Height          =   3135
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   150
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   5530
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         Style           =   5
         Checkboxes      =   -1  'True
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   3630
      ScaleHeight     =   780
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   255
      Width           =   255
   End
   Begin VB.TextBox txtInput 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   225
      Width           =   1845
   End
   Begin VB.Line lnSplit 
      BorderColor     =   &H8000000A&
      Visible         =   0   'False
      X1              =   4185
      X2              =   4185
      Y1              =   450
      Y2              =   780
   End
   Begin VB.Shape shapBack 
      BorderColor     =   &H8000000A&
      Height          =   825
      Left            =   300
      Top             =   -45
      Visible         =   0   'False
      Width           =   2970
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   390
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "ComboxExpend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTyCtl_Locale As ty_ctlObject_Locale
Private vRect As RECT
Private mobjPopup As CommandBarPopup
Private mobjCommandBar As CommandBar
Private mobjControl As CommandBarControl
'缺省属性值:
Const m_def_BackColor = &H80000005
Const m_def_ForeColor = &H80000012
Const m_def_FontTransparent = 0
Const m_def_BorderStyle = Em_BorderStyle.Show_None

'属性变量:
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Nodes As MSComctlLib.Nodes
Dim m_Node As MSComctlLib.Node
Dim m_FontTransparent As Boolean
Dim m_BorderStyle As Integer

'事件声明:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtInput,txtInput,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtInput,txtInput,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtInput,txtInput,-1,MouseUp
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtInput,txtInput,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtInput,txtInput,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtInput,txtInput,-1,KeyUp
 
Event NodeCheck(ByVal Node As Node, ByRef strCaption As String) 'MappingInfo=tvwList,tvwList,-1,NodeCheck
Event NodeClick(ByVal Node As Node) 'MappingInfo=tvwList,tvwList,-1,NodeClick
Private mstrNodeKeySeled As String  '选择的接点
Private mstrNodeSeled As String     '选择的接点
Private mblnNotChange As Boolean
Private mblnAddItemAddNode As Boolean '通过AddItem增加的节点

Public Sub AddItem(ByVal strID As String, ByVal strCatpion As String, _
    Optional blnRoot As Boolean, Optional ByVal blnCheck As Boolean, _
    Optional ByVal blnSelected As Boolean, _
    Optional ByVal lngForeColor As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加子项
    '编制:刘兴洪
    '日期:2014-12-25 15:18:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    If blnRoot Then
        Set objNode = tvwList.Nodes.Add(, 4, "Root", strCatpion)
    Else
        Set objNode = tvwList.Nodes.Add("Root", 4, "K" & strID, strCatpion)
    End If
    With objNode
        .Expanded = True
        .Selected = blnSelected
        If lngForeColor <> -1 Then .ForeColor = lngForeColor
        If Checkboxes Then .Checked = blnCheck
        .Tag = strID
    End With
    Call RefreshTextData
    mblnAddItemAddNode = True
End Sub
Public Sub Clear()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2014-12-25 15:17:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    tvwList.Nodes.Clear
    mblnAddItemAddNode = False
End Sub

Public Function ListCount() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:统计列表个数
    '返回:返回个数
    '编制:刘兴洪
    '日期:2015-01-09 16:13:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ListCount = tvwList.Nodes.Count
End Function

Private Sub picDown_KeyPress(KeyAscii As Integer)
       Debug.Print KeyAscii
End Sub

Private Sub picDownList_Resize()
    Err = 0: On Error Resume Next
    With picDownList
        tvwList.Left = .ScaleLeft
        tvwList.Top = .ScaleTop
        tvwList.Width = .ScaleWidth
        tvwList.Height = .ScaleHeight
    End With
End Sub
Private Sub tvwList_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Expanded = True
End Sub
Private Sub txtInput_GotFocus()
    If Not picDownList.Visible Then
        Call SetDownDefautStatu
        Call ClearTag
    End If
    If picDown.Enabled And picDown.Visible Then picDown.SetFocus
End Sub

Private Sub txtInput_LostFocus()
    Call SetControlColor
End Sub

Private Sub UserControl_EnterFocus()
    Static sngBegin As Single
    Dim sngNow As Single
    sngNow = Timer
    If Format((sngNow - sngBegin), "0.000") < 0.5 Then
        Exit Sub
    End If
    sngBegin = Timer
    txtInput.BackColor = &H8000000D
    UserControl.BackColor = &H8000000D
    txtInput.ForeColor = vbWhite
    If picDown.Enabled And picDown.Visible Then picDown.SetFocus
    If Not picDownList.Visible Then
        Call SetDownDefautStatu
        Call ClearTag
    End If
End Sub

Private Sub UserControl_ExitFocus()
    Call SetControlColor
End Sub

Private Sub SetControlColor()
    txtInput.BackColor = m_BackColor
    UserControl.BackColor = m_BackColor
    txtInput.ForeColor = m_ForeColor
End Sub

Private Sub UserControl_Initialize()
    Call zlCommandBarDef
    If Appearance = Show_3D Then
        UserControl.BorderStyle = 1
    Else
        UserControl.BorderStyle = 0
    End If
    
    Call UserControl_Resize
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If vbAltMask = Shift And KeyCode = vbKeyDown Then
        AddPopu 0, 0
       If tvwList.Enabled And tvwList.Visible Then tvwList.SetFocus
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii = vbKeyTab Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl.MousePointer = 1
End Sub

Private Sub UserControl_Resize()
    Dim lngLnWidth As Long, lngTemp As Long
    Dim lngStep As Long
    
    
    Err = 0: On Error Resume Next
    picDownList.Width = ScaleWidth - 50
    If Appearance = Show_3D Then
        With UserControl
            txtInput.Left = .ScaleLeft
            txtInput.Width = .ScaleWidth - picDown.Width
            txtInput.Top = .ScaleTop  '+ (.ScaleHeight - txtInput.Height) \ 2
            txtInput.Height = .ScaleHeight
            picDown.Top = .ScaleTop + 5
            picDown.Height = .ScaleHeight
            picDown.Left = .ScaleWidth - picDown.Width
            picDownList.Width = .ScaleWidth
        End With
        Call SetCtrlVisible
        Call SetDownDefautStatu
        
        Exit Sub
    End If
    lngStep = 10
    With shapBack
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
        If Not BorderStyle = Show_None Then
            shapBack.Visible = True
        End If
    End With
    With picDown
        .Top = shapBack.Top + lngStep
        .Left = ScaleWidth - .Width - lngStep * 2
        .Height = ScaleHeight - .Top - lngStep * 2
    End With
    lngTemp = lnSplit.BorderWidth * 20
    With UserControl
        lnSplit.X1 = picDown.Left - lngTemp * 2
        lnSplit.X2 = lnSplit.X1
        lnSplit.Y1 = shapBack.Top
        lnSplit.Y2 = shapBack.Top + picDown.Height + lngTemp * 2
        
        txtInput.Left = shapBack.Left + lngStep * 2
        txtInput.Width = lnSplit.X1 - txtInput.Left - lngStep
        txtInput.Top = picDown.Top
        txtInput.Height = picDown.Height
        Call SetDownDefautStatu
        Exit Sub
    End With
End Sub

Private Sub picDown_mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    Call SetDownStatu
    Call AddPopu(x, y)
     
End Sub
Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.Appearance = Show_3D Then Exit Sub
    
    If picDown.Tag = "In" Then
        If x < 0 Or y < 0 Or x > picDown.Width Or y > picDown.Height Then
            picDown.Tag = ""
            ReleaseCapture
            shapBack.BorderColor = &H8000000A
            lnSplit.BorderColor = shapBack.BorderColor
            If Appearance = Show_Flat Then
                SetCommandStatu 0
            End If
        End If
    Else
        picDown.Tag = "In"
        SetCapture picDown.hWnd
        shapBack.BorderColor = &H80000012 ' vbBlue
        lnSplit.BorderColor = shapBack.BorderColor
        Select Case Appearance
        Case Em_Appearance.Show_Flat
             SetCommandStatu 1
        Case Else
        End Select
    End If
End Sub
Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    Call SetDownDefautStatu
    Call ClearTag
   ' tmrRefreah.Enabled = False
End Sub

Private Sub SetCommandStatu(ByVal intStyle As EM_DrawStyle, Optional intTYPE As Integer)
    zlRaisEffect picDown, intStyle  ', "", intAlignMent
    Call DrawIco
End Sub
Private Sub SetDownDefautStatu()
    Dim intTYPE As Integer
    Select Case Appearance
    Case Em_Appearance.Show_Flat
         intTYPE = 0
    Case Else
         intTYPE = 1
    End Select
    SetCommandStatu intTYPE
End Sub
Private Sub ClearTag()
     shapBack.BorderColor = &H8000000A
     lnSplit.BorderColor = shapBack.BorderColor
     picDown.Tag = ""
End Sub

Private Sub SetDownStatu()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置按下状态
    '编制:刘兴洪
    '日期:2014-12-19 16:03:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
      Dim intTYPE As Integer
   
     Select Case Appearance
     Case Em_Appearance.Show_3D
        UserControl.BorderStyle = 1
        Call zlControl.PicShowFlat(picDown, -1, , taCenterAlign)
        Call DrawIco
        UserControl_Resize
        Exit Sub
    Case Else
        SetCtrlVisible
        UserControl.BorderStyle = Em_BorderStyle.Show_None
        SetCommandStatu 0
        Call DrawIco
        UserControl_Resize
        Exit Sub
    End Select
    
End Sub

Private Sub DrawIco()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:绘制图标
    '编制:刘兴洪
    '日期:2014-12-19 14:52:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngX As Single, sngY As Single, sngPicW As Single, sngPicH As Single
    Dim sdPic As StdPicture
    Set sdPic = picDownList.Picture
    With sdPic
        sngPicW = sdPic.Width / 1.766667
        sngPicH = sdPic.Height / 1.766667
        sngX = (picDown.ScaleWidth - sngPicW) / 2
        sngY = (picDown.ScaleHeight - sngPicH) / 2
    End With
   
    picDown.PaintPicture sdPic, sngX, sngY
End Sub
 

Private Sub zlCommandBarDef()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义类
    '编制:刘兴洪
    '日期:2012-08-15 15:51:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.DeleteAll
    Exit Sub
errHandle:
    MsgBox Err.Description
End Sub
Private Sub AddPopu(ByVal x As Long, ByVal y As Long)
    vRect = zlControl.GetControlRect(UserControl.hWnd)
    vRect.Left = vRect.Left + 10
    vRect.Top = vRect.Top + 50
    Call CreatePopuMenu
    If Not mobjCommandBar Is Nothing Then Call mobjCommandBar.ShowPopup(, vRect.Left, vRect.Top + picDown.Height)
End Sub

Private Sub CreatePopuMenu()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建临时菜单
    '编制:刘兴洪
    '日期:2012-11-21 09:49:35
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim objCustom As CommandBarControlCustom
   
    Set mobjCommandBar = cbsThis.Add("PopupPati", xtpBarPopup)
    With mobjCommandBar.Controls
        Set objCustom = .Add(xtpControlCustom, 9999, "")
        objCustom.Handle = picDownList.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With
End Sub

Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
     
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Em_Appearance
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Em_Appearance)
    Call SetCtrlVisible
    UserControl.Appearance() = New_Appearance
    If New_Appearance = Show_3D Then
        UserControl.BorderStyle = 1
    Else
        UserControl.BorderStyle = 0
    End If
    PropertyChanged "Appearance"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Em_BorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Em_BorderStyle)
     m_BorderStyle = New_BorderStyle
     
    Call SetCtrlVisible
    Call UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
    
    Set m_Nodes = PropBag.ReadProperty("Nodes", Nothing)
    Set m_Node = PropBag.ReadProperty("Node", Nothing)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)

    txtInput.Text = PropBag.ReadProperty("Text", "Text1")
    Set txtInput.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtInput.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtInput.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtInput.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    txtInput.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
    txtInput.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    txtInput.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
   
    Set tvwList.Font = Font
 
    txtInput.Locked = PropBag.ReadProperty("Locked", True)
    If Appearance = Show_3D Then
        UserControl.BorderStyle = 1
    Else
        UserControl.BorderStyle = 0
    End If
    Call SetCtrlVisible
    Call UserControl_Resize
    Call SetDownDefautStatu
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    tvwList.Checkboxes = PropBag.ReadProperty("Checkboxes", True)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    Call SetControlColor
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Text", txtInput.Text, "Text1")
    Call PropBag.WriteProperty("Font", txtInput.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", txtInput.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtInput.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtInput.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtInput.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", txtInput.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontTransparent", m_FontTransparent, m_def_FontTransparent)
    Call PropBag.WriteProperty("FontUnderline", txtInput.FontUnderline, 0)
'    Call PropBag.WriteProperty("ForeColor", txtInput.ForeColor, &H80000008)
'    Call PropBag.WriteProperty("BackColor", txtInput.BackColor, &H80000005)
    Call PropBag.WriteProperty("Nodes", m_Nodes, Nothing)
    Call PropBag.WriteProperty("Node", m_Node, Nothing)
    Call PropBag.WriteProperty("Locked", txtInput.Locked, True)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Checkboxes", tvwList.Checkboxes, True)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,Text
Public Property Get Text() As String
    Text = txtInput.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtInput.Text() = New_Text
    PropertyChanged "Text"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,Font
Public Property Get Font() As Font
    Set Font = txtInput.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtInput.Font = New_Font
    Set tvwList.Font = Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = txtInput.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtInput.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = txtInput.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtInput.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontName
Public Property Get FontName() As String
    FontName = txtInput.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtInput.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = txtInput.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtInput.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = txtInput.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtInput.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get FontTransparent() As Boolean
    FontTransparent = m_FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    m_FontTransparent = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = txtInput.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtInput.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
 

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_FontTransparent = m_def_FontTransparent
    m_BorderStyle = m_def_BorderStyle
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
End Sub

Private Sub tvwList_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim objNode As Node
    Dim strCaption As String
    
    If mblnNotChange = True Then Exit Sub
    
    mblnNotChange = True
    If mblnAddItemAddNode Then
        If Node.Key = "Root" Then
            If Node.Checked = False Then
                txtInput.Text = "": txtInput.Tag = ""
            Else
                txtInput.Text = Node.Text: txtInput.Tag = Node.Tag
            End If
            
            For Each objNode In tvwList.Nodes
                If objNode.Key <> Node.Key Then
                    objNode.Checked = Node.Checked
                End If
            Next
            If Node.Checked = False Then
                For Each objNode In tvwList.Nodes
                    If objNode.Key <> Node.Key Then
                        objNode.Checked = True
                        Exit For
                    End If
                Next
            End If
        Else
            tvwList.Nodes("Root").Checked = False
        End If
    End If
    
    RaiseEvent NodeCheck(Node, strCaption)
    If mblnAddItemAddNode Then
        Call RefreshTextData
    Else
        txtInput.Text = strCaption
    End If
    mblnNotChange = False
End Sub
Private Sub RefreshTextData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新刷新数据
    '编制:刘兴洪
    '日期:2014-12-25 15:57:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node
    Dim mstrRootCaption As String
    
    If Checkboxes = False Then
        If tvwList.SelectedItem Is Nothing Then txtInput.Text = "": Exit Sub
        txtInput.Text = tvwList.SelectedItem.Text
        Exit Sub
    End If
    
    txtInput.Text = "": txtInput.Tag = ""
    For Each objNode In tvwList.Nodes
        If objNode.Checked Then
            If objNode.Key <> "Root" Then
                txtInput.Text = txtInput.Text & "," & objNode.Text
                If objNode.Tag <> "" Then
                    txtInput.Tag = txtInput.Tag & "," & objNode.Tag
                End If
            Else
                
                txtInput.Text = objNode.Text: txtInput.Tag = objNode.Tag
                Exit For
            End If
        End If
    Next
    If Left(txtInput.Text, 1) = "," Then txtInput.Text = Mid(txtInput.Text, 2)
    If Left(txtInput.Tag, 1) = "," Then txtInput.Tag = Mid(txtInput.Tag, 2)
End Sub


Private Sub tvwList_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent NodeClick(Node)
    If Checkboxes Then Exit Sub
    Call RefreshTextData
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get Nodes() As Nodes
    Set Nodes = tvwList.Nodes
End Property

Public Property Set Nodes(ByVal New_Nodes As MSComctlLib.Nodes)
    Set m_Nodes = New_Nodes
    PropertyChanged "Nodes"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get Node() As MSComctlLib.Node
    Set Node = m_Node
End Property

Public Property Set Node(ByVal New_Node As MSComctlLib.Node)
    Set m_Node = New_Node
    PropertyChanged "Node"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=txtInput,txtInput,-1,Locked
Public Property Get Locked() As Boolean
    Locked = txtInput.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    txtInput.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的显示
    '编制:刘兴洪
    '日期:2014-12-25 16:36:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisble As Boolean
    blnVisble = Appearance = 0
    blnVisble = blnVisble And BorderStyle = Show_Fixed_Single
    shapBack.Visible = blnVisble
    lnSplit.Visible = False 'blnVisble
End Sub
Public Sub RaisEffect(picBox As Object, Optional intStyle As EM_DrawStyle, _
    Optional strName As String = "", Optional TxtAlignment As gAlignment)
    Call zlRaisEffect(picBox, intStyle, strName, TxtAlignment)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=tvwList,tvwList,-1,Checkboxes
Public Property Get Checkboxes() As Boolean
    Checkboxes = tvwList.Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
    tvwList.Checkboxes() = New_Checkboxes
    PropertyChanged "Checkboxes"
End Property

Public Function GetNodesCheckedDatas(Optional ByVal blnRootSelReturnIsNull As Boolean = True) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择项目
    '入参:blnRootSelReturnIsNull-父项选中，返回空
    '返回:返回选中的接点项目数据,多个用逗号分离,如果blnRootSelReturnIsNull=true,则选中根目录的，则返回空
    '编制:刘兴洪
    '日期:2015-01-07 18:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node, strData As String
    strData = ""
    For Each objNode In tvwList.Nodes
        If objNode.Key = "Root" And objNode.Checked And blnRootSelReturnIsNull Then Exit Function
        If objNode.Checked And objNode.Key <> "Root" Then strData = strData & "," & objNode.Tag
    Next
    If strData <> "" Then strData = Mid(strData, 2)
    GetNodesCheckedDatas = strData
End Function

Public Function GetSelectNodeData() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选择项目
    '返回:返回选中的接点项目数据,选中根目录的，则返回空
    '编制:刘兴洪
    '日期:2015-01-07 18:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node, strData As String
    If tvwList.SelectedItem Is Nothing Then Exit Function
    If tvwList.SelectedItem.Key = "Root" Then Exit Function
    GetSelectNodeData = tvwList.SelectedItem.Tag
    Exit Function
End Function
Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub txtInput_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtInput.MousePointer = 1
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub txtInput_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

 

'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,&H80000005&
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Call SetControlColor
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,&H00404040&
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    txtInput.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
    'Call SetControlColor
End Property

