VERSION 5.00
Begin VB.UserControl zlOutLook 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   ScaleHeight     =   3600
   ScaleWidth      =   5565
   ToolboxBitmap   =   "Outlook.ctx":0000
   Begin VB.Timer timUD 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4110
      Top             =   2790
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   360
      MouseIcon       =   "Outlook.ctx":0312
      MousePointer    =   99  'Custom
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   270
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Index           =   0
      Left            =   420
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   2805
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   420
      End
      Begin VB.Image imgItem 
         Height          =   525
         Index           =   0
         Left            =   690
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "图标"
      Visible         =   0   'False
      Begin VB.Menu mnuIconBig 
         Caption         =   "大图标(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIconSmall 
         Caption         =   "小图标(&S)"
      End
   End
End
Attribute VB_Name = "zlOutLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'事件声明:
Event ItemClick(Item As OutItem)
'属性变量:
Dim m_AutoRedraw As Boolean
Dim m_SmallImageList As Object
Dim m_View As Integer
Dim m_SelectGroup As String
Dim m_SelectItem As String
Dim m_ImageList As Object
Dim WithEvents m_Groups As OutGroups
Attribute m_Groups.VB_VarHelpID = -1
Dim WithEvents m_Items As OutItems
Attribute m_Items.VB_VarHelpID = -1
'缺省属性值:
Const m_def_AutoRedraw = 0
Const m_def_View = 0        '如果为0就表示是大图标，否则是小图标
Const m_def_SelectGroup = ""
Const m_def_SelectItem = ""


'本地变量
Dim mblnDraw As Boolean
Dim mintTimer As Long

Private Sub cmdUD_Click(ByVal strText As String)
    Dim sngTemp As Single, intSelect As Integer
    
    If strText = "Down" Then
        intSelect = GetIndexFromCaption(SelectGroup)
        sngTemp = Val(cmdDown.Tag)
        If sngTemp <= 660 Then
            cmdUp.Tag = Val(cmdUp.Tag) + cmdDown.Tag
            cmdDown.Tag = ""
            picBody(intSelect).Top = picBody(intSelect).Top - sngTemp
            cmdUp.Top = picHead(intSelect).Top + picHead(intSelect).Height + 60
            cmdUp.Left = UserControl.ScaleWidth - cmdUp.Width - 60
            cmdDown.Visible = False
            timUD.Enabled = False
            mintTimer = 0
        Else
            cmdUp.Tag = Val(cmdUp.Tag) + 660
            cmdDown.Tag = sngTemp - 660
            picBody(intSelect).Top = picBody(intSelect).Top - 660
            cmdUp.Top = picHead(intSelect).Top + picHead(intSelect).Height + 60
            cmdUp.Left = UserControl.ScaleWidth - cmdUp.Width - 60
            timUD.Enabled = True
            mintTimer = 1
        End If
        cmdUp.Visible = True
    Else
        intSelect = GetIndexFromCaption(SelectGroup)
        sngTemp = Val(cmdUp.Tag)
        If sngTemp <= 660 Then
            cmdDown.Tag = Val(cmdDown.Tag) + cmdUp.Tag
            cmdUp.Tag = ""
            picBody(intSelect).Top = picHead(intSelect).Top + picHead(intSelect).Height
            cmdUp.Visible = False
            timUD.Enabled = False
            mintTimer = 0
        Else
            cmdDown.Tag = Val(cmdDown.Tag) + 660
            cmdUp.Tag = sngTemp - 660
            picBody(intSelect).Top = picBody(intSelect).Top + 660
            timUD.Enabled = True
            mintTimer = 2
        End If
        cmdDown.Visible = True
    End If
End Sub

Private Sub cmdDown_GotFocus()
    SendKeys "{TAB}"
End Sub

Private Sub cmdDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        cmdUD_Click "Down"
    End If
    SendKeys "{TAB}"
End Sub

Private Sub cmdDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or Y < 0 Or X > cmdDown.Width Or Y > cmdDown.Height Then timUD.Enabled = False
End Sub

Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timUD.Enabled = False
    mintTimer = 0
End Sub

Private Sub cmdUp_GotFocus()
    SendKeys "{TAB}"
End Sub

Private Sub cmdUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        cmdUD_Click "Up"
    End If
    SendKeys "+{TAB}"
End Sub

Private Sub cmdUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or Y < 0 Or X > cmdDown.Width Or Y > cmdDown.Height Then timUD.Enabled = False
End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timUD.Enabled = False
    mintTimer = 0
End Sub

Private Sub timUD_Timer()
    If mintTimer = 1 Then
         cmdUD_Click "Down"
    ElseIf mintTimer = 2 Then
         cmdUD_Click "Up"
    Else
        timUD.Enabled = False
    End If
End Sub

Private Sub imgItem_Click(Index As Integer)
    RaiseEvent ItemClick(m_Items(lblItem(Index).Tag))
End Sub

Private Sub imgItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc As RECT
    
    If View = 0 Then
        rc.Left = imgItem(Index).Left - 4
        rc.Right = imgItem(Index).Left + imgItem(Index).Width + 4
        rc.Top = imgItem(Index).Top - 4
        rc.Bottom = imgItem(Index).Top + imgItem(Index).Height - 2
    Else
        rc.Left = imgItem(Index).Left - 4
        rc.Right = imgItem(Index).Left + imgItem(Index).Width
        rc.Top = imgItem(Index).Top - 4
        rc.Bottom = imgItem(Index).Top + imgItem(Index).Height + 4
    End If
    DrawEdge picBody(GetIndexFromCaption(SelectGroup)).hdc, rc, BDR_SUNKENINNER, BF_RECT
    
    mblnDraw = True
End Sub

Private Sub imgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc As RECT
    
    If View = 0 Then
        rc.Left = imgItem(Index).Left - 4
        rc.Right = imgItem(Index).Left + imgItem(Index).Width + 4
        rc.Top = imgItem(Index).Top - 4
        rc.Bottom = imgItem(Index).Top + imgItem(Index).Height - 2
    Else
        rc.Left = imgItem(Index).Left - 4
        rc.Right = imgItem(Index).Left + imgItem(Index).Width
        rc.Top = imgItem(Index).Top - 4
        rc.Bottom = imgItem(Index).Top + imgItem(Index).Height + 4
    End If
    DrawEdge picBody(GetIndexFromCaption(SelectGroup)).hdc, rc, BDR_RAISEDOUTER, BF_RECT

    mblnDraw = True
End Sub

Private Sub imgItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgItem_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblItem_Click(Index As Integer)
    RaiseEvent ItemClick(m_Items(lblItem(Index).Tag))
End Sub

Private Sub lblItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgItem_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgItem_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub lblItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgItem_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub m_Groups_AddGroup(Key As String, Caption As String)
    Dim intMax As Integer
    intMax = picHead.UBound + 1
    Load picHead(intMax)
    Load picBody(intMax)
    picHead(intMax).Visible = True
    picBody(intMax).Visible = True
    picHead(intMax).Tag = Caption
    picBody(intMax).Tag = Caption
    picBody(intMax).Height = 200
    If m_Groups.Count = 1 Then SelectGroup = Caption
End Sub

Private Sub m_Groups_RemoveGroup(vntIndexKey As Variant)
    Dim strSelect As String, intSelect As Integer
    Dim objTemp As Variant
    strSelect = m_Groups(vntIndexKey)
    
    For Each objTemp In m_Items
        If objTemp.GroupName = strSelect Then
            m_Items.Remove objTemp.Key
        End If
    Next
    
    For Each objTemp In picHead
        If objTemp.Tag = strSelect Then
            intSelect = objTemp.Index
            Unload picBody(intSelect)
            Unload picHead(intSelect)
            If strSelect = m_SelectGroup Then SelectGroup = picHead(picHead.LBound).Tag
            Exit For
        End If
    Next
    
End Sub

Private Sub m_Items_AddItem(Key As String, Caption As String, Icon As String, GroupName As String)
    Dim intMax As Integer, picTemp As PictureBox
    Dim intSelect As Integer
    
    intMax = lblItem.UBound + 1
    Load imgItem(intMax)
    Load lblItem(intMax)
    imgItem(intMax).Visible = True
    lblItem(intMax).Visible = True
    
    If m_View = 0 Then
        imgItem(intMax).Picture = m_ImageList.ListImages(Icon).ExtractIcon
    Else
        imgItem(intMax).Picture = m_SmallImageList.ListImages(Icon).ExtractIcon
    End If
    imgItem(intMax).Tag = Icon
    lblItem(intMax).Tag = IIf(Key = "", "K" & Caption, Key)
    lblItem(intMax).Caption = Caption
    
    For Each picTemp In picBody
        If picTemp.Tag = GroupName Then Exit For
    Next
    Set imgItem(intMax).Container = picTemp
    Set lblItem(intMax).Container = picTemp
    
    If m_AutoRedraw = True Then
        Call ReListItem(picTemp)
    End If
End Sub

Private Sub m_Items_RemoveItem(vntIndexKey As Variant)
    Dim strSelect As String, intSelect As Integer
    Dim objTemp As Variant, picTemp As PictureBox
    strSelect = m_Items(vntIndexKey).Key
    
    For Each objTemp In lblItem
        If objTemp.Tag = strSelect Then
            intSelect = objTemp.Index
            Set picTemp = lblItem(intSelect).Container
            Unload lblItem(intSelect)
            Unload imgItem(intSelect)
            If strSelect = m_SelectItem Then SelectItem = lblItem.LBound
            Exit For
        End If
    Next
    
    If m_AutoRedraw = True Then
        Call ReListItem(picTemp)
        Call UserControl_Resize
    End If
End Sub

Private Sub ReListItem(picContainer As PictureBox)
    Dim sngTop As Single, sngLeft As Single
    Dim lblTemp As Label
    If picContainer Is Nothing Then Exit Sub
    
    picContainer.Height = 10
    sngTop = 10
    
    If m_View = 0 Then
        '大图标显示
        sngLeft = (picContainer.ScaleWidth - 32) / 2
        For Each lblTemp In lblItem
            If lblTemp.Container.Tag = picContainer.Tag Then
                imgItem(lblTemp.Index).Top = sngTop
                imgItem(lblTemp.Index).Height = 40
                imgItem(lblTemp.Index).Left = sngLeft
                lblTemp.Left = (picContainer.ScaleWidth - lblTemp.Width) / 2
                lblTemp.Top = sngTop + 40
                sngTop = sngTop + 70
                picContainer.Height = picContainer.Height + 1050
            End If
        Next
    Else
        '小图标显示
        For Each lblTemp In lblItem
            If lblTemp.Container.Tag = picContainer.Tag Then
                imgItem(lblTemp.Index).Top = sngTop
                imgItem(lblTemp.Index).Width = 20
                imgItem(lblTemp.Index).Left = picContainer.ScaleLeft + 8
                lblTemp.Left = imgItem(lblTemp.Index).Left + imgItem(lblTemp.Index).Width
                lblTemp.Top = sngTop
                sngTop = sngTop + 25
                picContainer.Height = picContainer.Height + 400
            End If
        Next
    End If
End Sub

Private Sub mnuIconBig_Click()
    If View = 0 Then Exit Sub
    View = 0
    mnuIconBig.Checked = True
    mnuIconSmall.Checked = False
End Sub

Private Sub mnuIconSmall_Click()
    If View = 1 Then Exit Sub
    View = 1
    mnuIconBig.Checked = False
    mnuIconSmall.Checked = True
End Sub

Private Sub picBody_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnDraw = True Then
        picBody(Index).Cls
        mblnDraw = False
    End If
    timUD.Enabled = False
End Sub

Private Sub picBody_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuIcon, vbPopupMenuRightButton
End Sub

Private Sub picBody_Resize(Index As Integer)
    Dim sngLeft As Single
    Dim lblTemp As Label
    
    If m_View = 1 Then Exit Sub '小图标就不处理
    sngLeft = (ScaleWidth / 15 - 32) / 2
    For Each lblTemp In lblItem
        If lblTemp.Container = picBody(Index) Then
            imgItem(lblTemp.Index).Left = sngLeft
            lblItem(lblTemp.Index).Left = (ScaleWidth / 15 - lblItem(lblTemp.Index).Width) / 2
        End If
    Next
End Sub

Private Sub picHead_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc As RECT
    rc.Left = picHead(Index).ScaleLeft
    rc.Right = picHead(Index).ScaleLeft + picHead(Index).ScaleWidth
    rc.Top = picHead(Index).ScaleTop
    rc.Bottom = picHead(Index).ScaleTop + picHead(Index).ScaleHeight
    
    DrawEdge picHead(Index).hdc, rc, EDGE_SUNKEN, BF_RECT
End Sub

Private Sub picHead_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim rc As RECT
    rc.Left = picHead(Index).ScaleLeft
    rc.Right = picHead(Index).ScaleLeft + picHead(Index).ScaleWidth
    rc.Top = picHead(Index).ScaleTop
    rc.Bottom = picHead(Index).ScaleTop + picHead(Index).ScaleHeight
    
    DrawEdge picHead(Index).hdc, rc, EDGE_RAISED, BF_RECT
    
    
    SelectGroup = picHead(Index).Tag
    If Button = 2 Then PopupMenu mnuIcon, vbPopupMenuRightButton
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuIcon, vbPopupMenuRightButton
End Sub

Private Sub picHead_Paint(Index As Integer)
    Dim rc As RECT
    picHead(Index).Cls
    
    picHead(Index).CurrentX = (picHead(Index).ScaleWidth - picHead(Index).TextWidth(picHead(Index).Tag)) / 2
    picHead(Index).CurrentY = (picHead(Index).ScaleHeight - picHead(Index).TextHeight(picHead(Index).Tag)) / 2
    
    picHead(Index).Print picHead(Index).Tag
    
    rc.Left = picHead(Index).ScaleLeft
    rc.Right = picHead(Index).ScaleLeft + picHead(Index).ScaleWidth
    rc.Top = picHead(Index).ScaleTop
    rc.Bottom = picHead(Index).ScaleTop + picHead(Index).ScaleHeight
    DrawEdge picHead(Index).hdc, rc, EDGE_RAISED, BF_RECT
End Sub

Private Sub picHead_Resize(Index As Integer)
    Call picHead_Paint(Index)
End Sub

Private Sub UserControl_Initialize()
    Set m_Groups = New OutGroups
    Set m_Items = New OutItems
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
'    m_SelectGroup = m_def_SelectGroup
    m_SelectItem = m_def_SelectItem
    m_View = m_def_View
    m_SelectGroup = m_def_SelectGroup
    m_AutoRedraw = m_def_AutoRedraw
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
'    m_SelectGroup = PropBag.ReadProperty("SelectGroup", m_def_SelectGroup)
    m_SelectItem = PropBag.ReadProperty("SelectItem", m_def_SelectItem)
    m_View = PropBag.ReadProperty("View", m_def_View)
    Set m_SmallImageList = PropBag.ReadProperty("SmallImageList", Nothing)
    m_SelectGroup = PropBag.ReadProperty("SelectGroup", m_def_SelectGroup)
    m_AutoRedraw = PropBag.ReadProperty("AutoRedraw", m_def_AutoRedraw)
End Sub

Private Sub UserControl_Resize()
    Dim intSelect As Integer, intCount As Integer
    Dim sngTop As Single, sngBottom As Single
    
    On Error Resume Next
    intSelect = GetIndexFromCaption(SelectGroup)
    sngTop = UserControl.ScaleTop
    sngBottom = UserControl.ScaleHeight
    
    For intCount = 1 To intSelect
        picHead(intCount).Left = UserControl.ScaleLeft
        picHead(intCount).Width = UserControl.ScaleWidth
        picHead(intCount).Top = sngTop
        picHead(intCount).Height = 300
        sngTop = sngTop + 300
        picBody(intCount).Visible = False
    Next
    
    For intCount = picHead.UBound To intSelect + 1 Step -1
        picHead(intCount).Left = UserControl.ScaleLeft
        picHead(intCount).Width = UserControl.ScaleWidth
        sngBottom = sngBottom - 300
        picHead(intCount).Top = sngBottom
        picHead(intCount).Height = 300
        picBody(intCount).Visible = False
    Next
    
    picBody(intSelect).Left = UserControl.ScaleLeft
    picBody(intSelect).Width = UserControl.ScaleWidth
    picBody(intSelect).Top = sngTop
    picBody(intSelect).ZOrder 1
    picBody(intSelect).Visible = True
    '把其它的放在下面
    For intCount = picBody.LBound To picBody.UBound
        If intCount <> intSelect Then picBody(intCount).ZOrder 1
    Next
    
    If sngBottom - sngTop < picBody(intSelect).Height Then
        cmdUp.Tag = ""
        cmdDown.Tag = picBody(intSelect).Height - (sngBottom - sngTop)
        cmdDown.Left = UserControl.ScaleWidth - cmdDown.Width - 60
        cmdDown.Top = sngBottom - cmdDown.Height - 60
        cmdDown.Visible = True
    Else
        cmdDown.Visible = False
    End If
    cmdUp.Visible = False
End Sub

Private Sub UserControl_Terminate()
    Set m_Groups = Nothing
    Set m_Items = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
'    Call PropBag.WriteProperty("SelectGroup", m_SelectGroup, m_def_SelectGroup)
    Call PropBag.WriteProperty("SelectItem", m_SelectItem, m_def_SelectItem)
    Call PropBag.WriteProperty("View", m_View, m_def_View)
    Call PropBag.WriteProperty("SmallImageList", m_SmallImageList, Nothing)
    Call PropBag.WriteProperty("SelectGroup", m_SelectGroup, m_def_SelectGroup)
    Call PropBag.WriteProperty("AutoRedraw", m_AutoRedraw, m_def_AutoRedraw)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get Groups() As OutGroups
    Set Groups = m_Groups
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get Items() As OutItems
    Set Items = m_Items
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
End Property
'
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get SelectGroup() As String
    SelectGroup = m_SelectGroup
End Property

Public Property Let SelectGroup(ByVal New_SelectGroup As String)
    m_SelectGroup = New_SelectGroup
    PropertyChanged "SelectGroup"

    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,2,
Public Property Get SelectItem() As String
Attribute SelectItem.VB_MemberFlags = "400"
    SelectItem = m_SelectItem
End Property

Public Property Let SelectItem(ByVal New_SelectItem As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_SelectItem = New_SelectItem
    PropertyChanged "SelectItem"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = m_AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    m_AutoRedraw = New_AutoRedraw
    PropertyChanged "AutoRedraw"
    
    If m_AutoRedraw = True Then
        Dim objTemp As Object
        For Each objTemp In picBody
            ReListItem objTemp
        Next
        
        Call UserControl_Resize
    End If
End Property

Private Function GetIndexFromCaption(Caption As String) As Integer
    Dim picTemp As PictureBox
    
    For Each picTemp In picBody
        If picTemp.Tag = Caption Then
            GetIndexFromCaption = picTemp.Index
            Exit Function
        End If
    Next
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get View() As Integer
    View = m_View
End Property

Public Property Let View(ByVal New_View As Integer)
    Dim picTemp As PictureBox
    
    m_View = New_View
    PropertyChanged "View"
    mnuIconBig.Checked = (New_View = 0)
    mnuIconSmall.Checked = (New_View = 1)
    
    If UserControl.Ambient.UserMode = False Then Exit Property

    Dim img As Image
'    For Each picTemp In picBody
'        picTemp.Visible = False
'    Next
    If m_View = 0 Then
        For Each img In imgItem
             If img.Tag <> "" Then img.Picture = m_ImageList.ListImages(img.Tag).ExtractIcon
        Next
    Else
        For Each img In imgItem
            If img.Tag <> "" Then img.Picture = m_SmallImageList.ListImages(img.Tag).ExtractIcon
        Next
    End If
        
    For Each picTemp In picBody
        ReListItem picTemp
'        If picTemp.Index <> 0 Then picTemp.Visible = True
    Next
    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get SmallImageList() As Object
    Set SmallImageList = m_SmallImageList
End Property

Public Property Set SmallImageList(ByVal New_SmallImageList As Object)
    Set m_SmallImageList = New_SmallImageList
    PropertyChanged "SmallImageList"
End Property
