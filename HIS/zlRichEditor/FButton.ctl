VERSION 5.00
Begin VB.UserControl FButton 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   ControlContainer=   -1  'True
   ScaleHeight     =   930
   ScaleWidth      =   1170
   ToolboxBitmap   =   "FButton.ctx":0000
End
Attribute VB_Name = "FButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##模 块 名：FButton.ctl
'##创 建 人：吴庆伟
'##日    期：2005年6月20日
'##修 改 人：
'##日    期：
'##描    述：自定义的Office风格按钮控件。IsOptButton属性表示其是否为单选按钮。
'##版    本：
'######################################################################################

Option Explicit

Private mvarValue As OLE_OPTEXCLUSIVE
Private mvarIsOptButton As Boolean            '是否是单选按钮，否则是普通按钮
Private mvarPicture As StdPicture
Private mvarMaskColor As OLE_COLOR

Public Event Click()

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = mvarMaskColor
End Property

Public Property Let MaskColor(vData As OLE_COLOR)
    mvarMaskColor = vData
    PropertyChanged "MaskColor"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Set Picture(vData As StdPicture)
    Set mvarPicture = vData
    PropertyChanged "Picture"
End Property

Public Property Let Picture(vData As StdPicture)
    Set mvarPicture = vData
    PropertyChanged "Picture"
End Property

Public Property Get IsOptButton() As Boolean
    IsOptButton = mvarIsOptButton
End Property

Public Property Let IsOptButton(vData As Boolean)
    mvarIsOptButton = vData
    PropertyChanged "IsOptButton"
End Property

Public Property Get Value() As OLE_OPTEXCLUSIVE
Attribute Value.VB_UserMemId = 0
    Value = mvarValue
End Property

Public Property Let Value(vData As OLE_OPTEXCLUSIVE)
    mvarValue = vData
    If vData Then
        DrawButton 2
    Else
        DrawButton 0
    End If
    PropertyChanged "Value"
End Property

Private Sub UserControl_Click()
    If IsOptButton Then
        If Not Value Then
            RaiseEvent Click
            Value = True
        End If
    Else
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton 3
    UserControl.Tag = "Down"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And X <= ScaleWidth And Y >= 0 And Y <= ScaleHeight Then
        SetCapture UserControl.Hwnd         '导致ToolTipText不起作用了！
        '鼠标移入！！！
        If UserControl.Tag = "Down" Then
            DrawButton 3
        Else
            DrawButton 1
        End If
    Else
        If UserControl.Tag <> "" Then
            DrawButton 3
        Else
            '鼠标移出！！！                 '导致ToolTipText不起作用了！
            ReleaseCapture
            If Value Then
                DrawButton 2
            Else
                DrawButton 0
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Tag = ""
    If Value Then
        DrawButton 2
    Else
        DrawButton 0
    End If
End Sub

Private Sub DrawButton(lDrawStyle As Long)
    '0:普通 &H8000000F    1:移动  &HEED2C1   2:选中 &HE8E6E1    3:按下  &HE2B598          边框:&HC56A31
    On Error Resume Next
    If mvarIsOptButton = False And lDrawStyle = 2 Then lDrawStyle = 0
    Cls
    Select Case lDrawStyle
    Case 0  '普通
        BackColor = &H8000000F
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 1  '移动
        BackColor = &HEED2C1
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 2  '选中
        BackColor = &HE8E6E1
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 3  '按下
        BackColor = &HE2B598
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    End Select
    Refresh
    Err.Clear
End Sub

Private Sub UserControl_InitProperties()
    Value = False
    IsOptButton = False
    Set Picture = LoadPicture("")
    MaskColor = vbWhite
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", False)
    IsOptButton = PropBag.ReadProperty("IsOptButton", False)
    Picture = PropBag.ReadProperty("Picture", LoadPicture(""))
    MaskColor = PropBag.ReadProperty("MaskColor", vbGreen)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Show()
    If Value Then
        DrawButton 2
    Else
        DrawButton 0
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mvarPicture = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Value", mvarValue, False
    PropBag.WriteProperty "IsOptButton", mvarIsOptButton, False
    PropBag.WriteProperty "Picture", mvarPicture, LoadPicture("")
    PropBag.WriteProperty "MaskColor", mvarMaskColor, vbGreen
    
    PropertyChanged "Value"
    PropertyChanged "IsOptButton"
    PropertyChanged "Picture"
    PropertyChanged "MaskColor"
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub
