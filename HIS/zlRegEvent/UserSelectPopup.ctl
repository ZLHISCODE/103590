VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl UserSelectPopup 
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2490
   ScaleWidth      =   4800
   Begin VB.PictureBox picSelect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2280
      ScaleHeight     =   1215
      ScaleWidth      =   825
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   825
      Begin VSFlex8Ctl.VSFlexGrid vsfSelect 
         Height          =   1215
         Left            =   30
         TabIndex        =   2
         Top             =   0
         Width           =   1425
         _cx             =   2514
         _cy             =   2143
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   240
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"UserSelectPopup.ctx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1470
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":0067
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":0571
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":0A7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":0F85
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":148F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserSelectPopup.ctx":1999
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   90
      ScaleHeight     =   315
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   60
      Width           =   1965
      Begin VB.PictureBox picItem 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   630
         ScaleHeight     =   345
         ScaleWidth      =   1365
         TabIndex        =   5
         Top             =   0
         Width           =   1365
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "第 1 周"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   30
            Width           =   585
         End
      End
      Begin VB.PictureBox picChangeButton 
         BorderStyle     =   0  'None
         Height          =   165
         Index           =   1
         Left            =   360
         Picture         =   "UserSelectPopup.ctx":1EA3
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   4
         Top             =   75
         Width           =   165
      End
      Begin VB.PictureBox picChangeButton 
         BorderStyle     =   0  'None
         Height          =   165
         Index           =   0
         Left            =   90
         Picture         =   "UserSelectPopup.ctx":239D
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   3
         Top             =   75
         Width           =   165
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   420
      Top             =   570
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "UserSelectPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Private mobjCommandBar As CommandBar
Private mobjDict As Object
Private mlngListIndex As Long

Private Enum m_ChangeType
    E_Previous = 0
    E_Next = 1
End Enum

Private m_PopupWidth As Single

'事件声明
Public Event ValueChanged(ByVal strKey As String, ByVal strValue As String) '值改变时触发

'弹出菜单宽度
Public Property Get PopupWidth() As Single
Attribute PopupWidth.VB_Description = "弹出菜单宽度。"
    PopupWidth = m_PopupWidth
End Property

Public Property Let PopupWidth(ByVal vNewValue As Single)
    m_PopupWidth = vNewValue
    PropertyChanged "PopupWidth"
End Property

Public Property Get SelectedKey() As String
    Dim varKeys As Variant
    
    If mobjDict Is Nothing Then Exit Property
    If mobjDict.Count = 0 Then Exit Property
    If mlngListIndex = -1 Then Exit Property
    If mlngListIndex > mobjDict.Count - 1 Then Exit Property
    
    varKeys = mobjDict.Keys
    SelectedKey = varKeys(mlngListIndex)
End Property

Public Property Let SelectedKey(ByVal vNewValue As String)
    Dim varKeys As Variant, i As Integer
    
    If mobjDict Is Nothing Then Exit Property
    If mobjDict.Count = 0 Then Exit Property
    varKeys = mobjDict.Keys
    For i = 0 To UBound(varKeys)
        If varKeys(i) = vNewValue Then
            mlngListIndex = i
            Call SetChangeValue
            Exit For
        End If
    Next
End Property

Public Property Get ListCount() As Integer
    Dim varKeys As Variant
    
    If mobjDict Is Nothing Then Exit Property
    ListCount = mobjDict.Count
End Property

Public Property Get SelectedValue() As String
    Dim varItems As Variant
    
    If mobjDict Is Nothing Then Exit Property
    If mobjDict.Count = 0 Then Exit Property
    If mlngListIndex = -1 Then Exit Property
    If mlngListIndex > mobjDict.Count - 1 Then Exit Property
    
    varItems = mobjDict.Items
    SelectedValue = varItems(mlngListIndex)
End Property

Public Sub AddItem(ByVal strKey As String, ByVal strValue As String)
    '功能:增加子项
    On Error Resume Next
    If mobjDict Is Nothing Then Set mobjDict = CreateObject("Scripting.Dictionary")
    
    If mobjDict.Exists(strKey) Then Exit Sub
    mobjDict.Add strKey, strValue
    
    With vsfSelect
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = strValue
    End With
    
'    If mobjDict.Count = 1 Then
'        mlngListIndex = 0
'        Call SetChangeValue
'    End If
'
    If mobjDict.Count > 0 Then
        picItem.Enabled = True
    End If
End Sub

Public Sub Clear()
    '功能:清除数据
    On Error Resume Next
    vsfSelect.Clear
    vsfSelect.Rows = 0
    
    lblItem.Caption = ""
    lblItem.Tag = ""
    picItem.Enabled = False
    
    picChangeButton(E_Previous).Enabled = False
    picChangeButton(E_Next).Enabled = False
    picChangeButton(E_Previous).Picture = imgList.ListImages(5).Picture
    picChangeButton(E_Next).Picture = imgList.ListImages(6).Picture
    
    If Not mobjDict Is Nothing Then mobjDict.RemoveAll
    mlngListIndex = -1
End Sub

Private Sub picChangeButton_Click(index As Integer)
    On Error Resume Next
    Select Case index
    Case E_Previous
        mlngListIndex = mlngListIndex - 1
    Case E_Next
        mlngListIndex = mlngListIndex + 1
    End Select
    Call SetChangeValue
End Sub

Private Sub picChangeButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    On Error Resume Next
    Select Case index
    Case E_Previous
        picChangeButton(E_Previous).Picture = imgList.ListImages(3).Picture
    Case E_Next
        picChangeButton(E_Next).Picture = imgList.ListImages(4).Picture
    End Select
End Sub

Private Sub picChangeButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    On Error Resume Next
    Select Case index
    Case E_Previous
        picChangeButton(E_Previous).Picture = imgList.ListImages(1).Picture
    Case E_Next
        picChangeButton(E_Next).Picture = imgList.ListImages(2).Picture
    End Select
End Sub

Private Sub picSelect_Resize()
    On Error Resume Next
    vsfSelect.Top = 0
    vsfSelect.Height = picSelect.ScaleHeight
    vsfSelect.Width = picSelect.ScaleWidth
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    Call zlCommandBarDef
    Set mobjDict = CreateObject("Scripting.Dictionary")
    
    Call SetChangeValue
    m_PopupWidth = picSelect.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PopupWidth = PropBag.ReadProperty("PopupWidth", picSelect.Width)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Height = picBack.Height
    picBack.Left = 0
    picBack.Top = 0
    picBack.Width = UserControl.Width
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    picItem.Top = 0
    picItem.Width = lblItem.Width
    picItem.Height = lblItem.Height
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    lblItem.Left = 0
'    lblItem.Top = 0
End Sub

Private Sub lblItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picItem_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    On Error Resume Next
    Call AddPopupMenu(X, Y)
End Sub

Private Sub UserControl_Terminate()
    Set mobjCommandBar = Nothing
    Set mobjDict = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PopupWidth", m_PopupWidth, picSelect.Width)
End Sub

Private Sub vsfSelect_Click()
    On Error Resume Next
    mlngListIndex = vsfSelect.Row
    Set mobjCommandBar = Nothing
    Call SetChangeValue
End Sub

Private Sub vsfSelect_DblClick()
'    On Error Resume Next
'    mlngListIndex = vsfSelect.Row
'    Set mobjCommandBar = Nothing
'    Call SetChangeValue
End Sub

Private Sub vsfSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    With vsfSelect
        .Redraw = flexRDBuffered
        .Row = .TopRow + Y \ .RowHeightMin
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = .BackColor
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, 0) = .ForeColor
        .Cell(flexcpBackColor, .Row, 0) = .BackColorSel
        .Cell(flexcpForeColor, .Row, 0) = vbWhite
    End With
End Sub

Private Sub SetChangeValue()
    Dim varKeys As Variant, varItems As Variant
    
    If lblItem.Tag = CStr(mlngListIndex) Then Exit Sub
    lblItem.Tag = CStr(mlngListIndex)
    picChangeButton(E_Previous).Enabled = False
    picChangeButton(E_Next).Enabled = False
    picChangeButton(E_Previous).Picture = imgList.ListImages(5).Picture
    picChangeButton(E_Next).Picture = imgList.ListImages(6).Picture
    picItem.Enabled = False
    
    If mobjDict Is Nothing Then Exit Sub
    If mobjDict.Count = 0 Then Exit Sub
    picItem.Enabled = True
    If mlngListIndex < 0 Then mlngListIndex = 0
    If mlngListIndex > mobjDict.Count - 1 Then mlngListIndex = mobjDict.Count - 1
    
    If mlngListIndex > 0 Then
        picChangeButton(E_Previous).Enabled = True
        picChangeButton(E_Previous).Picture = imgList.ListImages(1).Picture
    End If
    If mlngListIndex < mobjDict.Count - 1 Then
        picChangeButton(E_Next).Enabled = True
        picChangeButton(E_Next).Picture = imgList.ListImages(2).Picture
    End If
    
    varKeys = mobjDict.Keys: varItems = mobjDict.Items
    lblItem.Caption = varItems(mlngListIndex)
    picItem.Width = lblItem.Width '根据文本调整宽度
    RaiseEvent ValueChanged(varKeys(mlngListIndex), varItems(mlngListIndex))
End Sub

Private Sub AddPopupMenu(ByVal X As Long, ByVal Y As Long)
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(UserControl.Hwnd)
    vRect.Left = vRect.Left + picItem.Left - 30
    vRect.Top = vRect.Top + picItem.Height
    
    If mobjDict Is Nothing Then Exit Sub
    If mobjDict.Count < 2 Then Exit Sub
    Call CreatePopupMenu
    If Not mobjCommandBar Is Nothing Then
        vsfSelect.Row = vsfSelect.FindRow(lblItem.Caption, , 0)
        vsfSelect.ShowCell vsfSelect.Row, 0 '定位选中行
        Call mobjCommandBar.ShowPopup(, vRect.Left, vRect.Top)
    End If
End Sub

Private Sub zlCommandBarDef()
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
End Sub

Private Sub CreatePopupMenu()
    '功能:创建临时菜单
    Dim j As Long, sngHight As Single
    Dim objCustom As CommandBarControlCustom
    
    Set mobjCommandBar = cbsThis.Add("PopuMenu", xtpBarPopup)
    With mobjCommandBar.Controls
        Set objCustom = .Add(xtpControlCustom, 9999, "")
        picSelect.Width = m_PopupWidth
        sngHight = (vsfSelect.RowHeightMin + 5) * mobjDict.Count
        picSelect.Height = IIf(sngHight > 2000, 2000, sngHight)
        objCustom.handle = picSelect.Hwnd
        objCustom.flags = xtpFlagRightAlign
    End With
End Sub
