VERSION 5.00
Begin VB.UserControl CustomButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ScaleHeight     =   2010
   ScaleWidth      =   2475
   Begin VB.Shape shpLine 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000A&
      Height          =   1110
      Left            =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CustomButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ȱʡ����ֵ
Const m_def_Enabled = 0
'���Ա���:
Dim m_Picture As Picture
Dim m_Enabled As Boolean

Event Click()
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    SetEnabled UserControl.Controls, New_Enabled
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Or m_Enabled = False Then Exit Sub
    '0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    Call PicShowFlat(-1)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Tag = "In" Then
        If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
            Tag = "": ReleaseCapture
            '0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
'            shpLine.Visible = False
            UserControl.Cls
            Call PicShowFlat(0)
        End If
    Else
        Tag = "In"
        SetCapture Hwnd
        '0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
'        shpLine.Visible = True
        Call PicShowFlat(0)
        'Call PicShowFlat(1)
        'Call zlControl.PicShowFlat(UserControl, 1)
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Or m_Enabled = False Then Exit Sub
    Tag = "": shpLine.Visible = False
    Call PicShowFlat(0)
    RaiseEvent Click
End Sub
Private Sub UserControl_Paint()
     Call DrawPicture(m_Picture, True)
End Sub

Private Sub UserControl_Resize()
    With shpLine
        .Left = UserControl.ScaleLeft
        .Top = UserControl.ScaleTop
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With
    Call DrawPicture(m_Picture, True)
End Sub

Public Sub PicShowFlat(Optional IntStyle As Integer = -1, Optional strName As String = "", Optional intAlign As mTextAlign, Optional blnFontBold As Boolean)
    '���ܣ���PictureBoxģ��ɰ��»�͹������
    '������'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    '      intAlign=���Ҫ��ʾ�ı�,��ָ�����뷽ʽ
    
    Dim vRect As RECT, lngTmp As Long
    
    With UserControl
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            vRect.Left = .ScaleLeft
            vRect.Top = .ScaleTop
            vRect.Right = .ScaleWidth
            vRect.Bottom = .ScaleHeight
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, vRect, CLng(BDR_RAISEDINNER Or BF_SOFT), BF_RECT
                Case 2
                    DrawEdge .hDC, vRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, vRect, CLng(BDR_SUNKENOUTER Or BF_SOFT), BF_RECT
                Case -2
                    DrawEdge .hDC, vRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If intAlign = taCenterAlign Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2 '�м����
            ElseIf intAlign = taRightAlign Then
                .CurrentX = .ScaleWidth - .TextWidth(strName) - 2 '�ұ߶���
            Else
                .CurrentX = 2 '��߶���
            End If
            .FontBold = blnFontBold
            UserControl.Print strName
        End If
        Call DrawPicture(m_Picture)
    End With
End Sub
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "����/���ÿؼ�����ʾ��ͼ�Ρ�"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Call DrawPicture(m_Picture)
    PropertyChanged "Picture"
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    Set m_Picture = LoadPicture("")
    m_Enabled = m_def_Enabled
    Call DrawPicture(m_Picture)
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Call DrawPicture(m_Picture)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
Private Sub DrawLine(ByVal bytLine As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����
    '    bytLine-0-����;1-����;2-����
    '����:���˺�
    '����:2016-01-14 13:45:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ForeColor = &H8000000A
    Line (ScaleLeft + 20, ScaleTop + 20)-(ScaleWidth - 40, ScaleHeight - 40), , B
End Sub
Private Sub DrawPicture(objPic As StdPicture, Optional blnCls As Boolean)
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    Dim lngPicW As Long, lngPicH As Long
    If objPic Is Nothing Then Cls: Exit Sub
    
    lngPicW = objPic.Width * 0.567: lngPicH = objPic.Height * 0.567
    On Error Resume Next
    If blnCls Then Cls

    UserControl.PaintPicture objPic, (ScaleWidth - lngPicW) / 2, (ScaleHeight - lngPicH) / 2
    If Err.Number <> 0 Then Err.Clear
End Sub
