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
'##ģ �� ����FButton.ctl
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��6��20��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    �����Զ����Office���ť�ؼ���IsOptButton���Ա�ʾ���Ƿ�Ϊ��ѡ��ť��
'##��    ����
'######################################################################################

Option Explicit

Private mvarValue As OLE_OPTEXCLUSIVE
Private mvarIsOptButton As Boolean            '�Ƿ��ǵ�ѡ��ť����������ͨ��ť
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
        SetCapture UserControl.Hwnd         '����ToolTipText���������ˣ�
        '������룡����
        If UserControl.Tag = "Down" Then
            DrawButton 3
        Else
            DrawButton 1
        End If
    Else
        If UserControl.Tag <> "" Then
            DrawButton 3
        Else
            '����Ƴ�������                 '����ToolTipText���������ˣ�
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
    '0:��ͨ &H8000000F    1:�ƶ�  &HEED2C1   2:ѡ�� &HE8E6E1    3:����  &HE2B598          �߿�:&HC56A31
    On Error Resume Next
    If mvarIsOptButton = False And lDrawStyle = 2 Then lDrawStyle = 0
    Cls
    Select Case lDrawStyle
    Case 0  '��ͨ
        BackColor = &H8000000F
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 1  '�ƶ�
        BackColor = &HEED2C1
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 2  'ѡ��
        BackColor = &HE8E6E1
        Line (0, 0)-(ScaleWidth - Screen.TwipsPerPixelX, ScaleHeight - Screen.TwipsPerPixelY), &HC56A31, B
        PaintTransparentStdPic UserControl.hDC, 4, 4, 9, 9, mvarPicture, 0, 0, mvarMaskColor
    Case 3  '����
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
