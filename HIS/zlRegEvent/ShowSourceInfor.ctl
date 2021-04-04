VERSION 5.00
Begin VB.UserControl ShowSourceInfor 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   ScaleHeight     =   6255
   ScaleWidth      =   6780
End
Attribute VB_Name = "ShowSourceInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�¼�����:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click()
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"
Event DblClick()
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Private mobj�����Դ As �����Դ

Public Function LoadData(ByVal obj�����Դ As �����Դ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:obj�����Դ-�����Դ
    '����:
    '����:���سɹ�������true, ���򷵻�False
    '����:���˺�
    '����:2016-01-19 10:00:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj�����Դ = obj�����Դ
    Call PrintSoureInfor
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub PrintSoureInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ��Դ��Ϣ
    '���:lng��ԴID-��ԴID
    '����:���˺�
    '����:2016-01-11 13:06:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim sngTop As Single, sngLeft As Single
    Dim sngTopSkip As Single, sngLeftSkip As Single
    Dim fntTittleFont As StdFont
    Dim fntValueFont As StdFont
    Dim sngWidth As Single, sngHight As Single
    
    Set fntTittleFont = New StdFont
    Set fntValueFont = New StdFont
    On Error GoTo errHandle
    sngTop = ScaleTop: sngLeft = ScaleLeft
    
    With fntTittleFont
        .Charset = UserControl.Font.Charset
        .Italic = UserControl.Font.Italic
        .Name = UserControl.Font.Name
        .Size = UserControl.Font.Size
        .Strikethrough = UserControl.Font.Strikethrough
        .Underline = UserControl.Font.Underline
        .Weight = UserControl.Font.Weight
        .Bold = True
        .Size = 9
    End With
    With fntValueFont
        .Charset = UserControl.Font.Charset
        .Italic = UserControl.Font.Italic
        .Name = UserControl.Font.Name
        .Size = UserControl.Font.Size
        .Strikethrough = UserControl.Font.Strikethrough
        .Underline = UserControl.Font.Underline
        .Weight = UserControl.Font.Weight
        .Bold = False
        .Size = 9
    End With
 
    sngTopSkip = 0
    If mobj�����Դ Is Nothing Then
        Set mobj�����Դ = New �����Դ
        With mobj�����Դ
            .ID = 1
            .���� = "��ͨ"
            .���� = "01001"
            .���տ���״̬ = 1
            .����ID = 2
            .�������� = "���ﲿ��¥�ڿ�"
            .�Ƿ񽨲��� = False
            .��Ŀ���� = "����ҽ����"
            .ҽ������ = "�����"
        End With
    End If
    
    With UserControl
        sngWidth = .Width: sngHight = .Height
        .Width = 5000 '�����ǰ̫խ��ӡ������
        .Height = 5000
        
        sngTopSkip = .TextHeight("��") * 2 / 3
        .Cls
        Set .Font = fntTittleFont
        .CurrentX = sngLeft: .CurrentY = sngTop
        UserControl.Print "����:"
        
 
        sngLeftSkip = sngLeft + .TextWidth("����:") + 10
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj�����Դ.����
        Set .Font = fntTittleFont
        
        sngLeftSkip = sngLeftSkip + .TextWidth(mobj�����Դ.����)
        
        sngLeftSkip = sngLeftSkip + .TextWidth(String(10 - Len(mobj�����Դ.����), " "))
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        UserControl.Print "����:"
        
        sngLeftSkip = sngLeftSkip + .TextWidth("����:") + 10
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj�����Դ.����
        
        
        Set .Font = fntTittleFont
        sngTop = .CurrentY + sngTopSkip
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "����:"
        sngLeftSkip = sngLeft + .TextWidth("����:") + 10
        
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj�����Դ.��������
        
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "��Ŀ:"
        
        sngLeftSkip = sngLeft + .TextWidth("����:") + 10
        
        Set .Font = fntValueFont
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        UserControl.Print mobj�����Դ.��Ŀ����
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "ҽ��:"
        
        sngLeftSkip = sngLeft + .TextWidth("ҽ��:") + 10
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj�����Դ.ҽ������ & IIf(mobj�����Դ.ҽ��ְ�� = "", "", "(" & mobj�����Դ.ҽ��ְ�� & ")")
        
        Set .Font = fntTittleFont
        sngTop = .CurrentY + sngTopSkip
        .CurrentY = sngTop: .CurrentX = sngLeft
        UserControl.Print "���տ���:"
        
        sngLeftSkip = sngLeft + .TextWidth("���տ���:") + 10
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        '0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ
        Set .Font = fntValueFont
        UserControl.Print Decode(mobj�����Դ.���տ���״̬, 1, "����ԤԼ", 2, "��ֹԤԼ", 3, "�ܽڼ������ÿ���", "���ϰ�")
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "�Һű��뽨��:"
        
        sngLeftSkip = sngLeft + .TextWidth("�Һű��뽨��:") + 10
        Set .Font = fntValueFont
        .CurrentX = sngLeftSkip: .CurrentY = sngTop
        UserControl.Print IIf(mobj�����Դ.�Ƿ񽨲���, "��", "��")
        
        .Width = sngWidth: .Height = sngHight
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub UserControl_Initialize()
    Call PrintSoureInfor
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With picInfor
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = .ScaleHeight
    End With
End Sub
 
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "ǿ����ȫ�ػ�һ������"
     
End Sub
 
'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
End Sub

Private Sub UserControl_Terminate()
    Set mobj�����Դ = Nothing
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

