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
'事件声明:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click()
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick()
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Private mobj出诊号源 As 出诊号源

Public Function LoadData(ByVal obj出诊号源 As 出诊号源) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '入参:obj出诊号源-出诊号源
    '出参:
    '返回:加载成功，返加true, 否则返回False
    '编制:刘兴洪
    '日期:2016-01-19 10:00:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj出诊号源 = obj出诊号源
    Call PrintSoureInfor
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub PrintSoureInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印号源信息
    '入参:lng号源ID-号源ID
    '编制:刘兴洪
    '日期:2016-01-11 13:06:35
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
    If mobj出诊号源 Is Nothing Then
        Set mobj出诊号源 = New 出诊号源
        With mobj出诊号源
            .ID = 1
            .号类 = "普通"
            .号码 = "01001"
            .假日控制状态 = 1
            .科室ID = 2
            .科室名称 = "门诊部二楼内科"
            .是否建病案 = False
            .项目名称 = "主治医生号"
            .医生姓名 = "李大嘴"
        End With
    End If
    
    With UserControl
        sngWidth = .Width: sngHight = .Height
        .Width = 5000 '如果当前太窄打印不出来
        .Height = 5000
        
        sngTopSkip = .TextHeight("号") * 2 / 3
        .Cls
        Set .Font = fntTittleFont
        .CurrentX = sngLeft: .CurrentY = sngTop
        UserControl.Print "号码:"
        
 
        sngLeftSkip = sngLeft + .TextWidth("号码:") + 10
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj出诊号源.号码
        Set .Font = fntTittleFont
        
        sngLeftSkip = sngLeftSkip + .TextWidth(mobj出诊号源.号码)
        
        sngLeftSkip = sngLeftSkip + .TextWidth(String(10 - Len(mobj出诊号源.号码), " "))
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        UserControl.Print "号类:"
        
        sngLeftSkip = sngLeftSkip + .TextWidth("号类:") + 10
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj出诊号源.号类
        
        
        Set .Font = fntTittleFont
        sngTop = .CurrentY + sngTopSkip
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "科室:"
        sngLeftSkip = sngLeft + .TextWidth("科室:") + 10
        
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj出诊号源.科室名称
        
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "项目:"
        
        sngLeftSkip = sngLeft + .TextWidth("科室:") + 10
        
        Set .Font = fntValueFont
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        UserControl.Print mobj出诊号源.项目名称
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "医生:"
        
        sngLeftSkip = sngLeft + .TextWidth("医生:") + 10
        .CurrentY = sngTop: CurrentX = sngLeftSkip
        Set .Font = fntValueFont
        UserControl.Print mobj出诊号源.医生姓名 & IIf(mobj出诊号源.医生职称 = "", "", "(" & mobj出诊号源.医生职称 & ")")
        
        Set .Font = fntTittleFont
        sngTop = .CurrentY + sngTopSkip
        .CurrentY = sngTop: .CurrentX = sngLeft
        UserControl.Print "假日控制:"
        
        sngLeftSkip = sngLeft + .TextWidth("假日控制:") + 10
        
        .CurrentY = sngTop: .CurrentX = sngLeftSkip
        '0-不上班;1-上班且开放预约;2-上班但不开放预约
        Set .Font = fntValueFont
        UserControl.Print Decode(mobj出诊号源.假日控制状态, 1, "开放预约", 2, "禁止预约", 3, "受节假日设置控制", "不上班")
        
        sngTop = .CurrentY + sngTopSkip
        Set .Font = fntTittleFont
        .CurrentX = sngLeft
        .CurrentY = sngTop
        UserControl.Print "挂号必须建档:"
        
        sngLeftSkip = sngLeft + .TextWidth("挂号必须建档:") + 10
        Set .Font = fntValueFont
        .CurrentX = sngLeftSkip: .CurrentY = sngTop
        UserControl.Print IIf(mobj出诊号源.是否建病案, "是", "否")
        
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
 
'注意！不要删除或修改下列被注释的行！
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "强制完全重画一个对象。"
     
End Sub
 
'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
End Sub

Private Sub UserControl_Terminate()
    Set mobj出诊号源 = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
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

