VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl ClinicPlanDetailPages 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   ScaleHeight     =   6120
   ScaleWidth      =   11835
   Begin VB.PictureBox picTimeWork 
      BorderStyle     =   0  'None
      Height          =   4110
      Index           =   0
      Left            =   2490
      ScaleHeight     =   4110
      ScaleWidth      =   5700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   5700
      Begin zl9RegEvent.ClinicPlanDetail ClinicDetail 
         Height          =   2925
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   570
         Width           =   5205
         _ExtentX        =   12938
         _ExtentY        =   8176
      End
   End
   Begin XtremeSuiteControls.TabControl tbPageTimeWork 
      Height          =   930
      Left            =   570
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1680
      _Version        =   589884
      _ExtentX        =   2963
      _ExtentY        =   1640
      _StockProps     =   64
   End
   Begin VB.Shape shpLine 
      BorderColor     =   &H80000003&
      Height          =   630
      Left            =   120
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "ClinicPlanDetailPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'缺省属性值:
Const m_def_医生姓名 = ""
Const m_def_BackStyle = 0
Const m_def_诊疗频次 = 5
Const m_def_EditMode = 0
Const m_def_ForeColor = 0
Const m_def_BorderStyle = 0
'属性变量:
Dim m_医生姓名 As String
Dim m_BackStyle As Integer
Dim m_诊疗频次 As Integer
Dim m_EditMode As gRegistPlanEditMode
Dim m_ForeColor As Long
Dim m_Font As Font
Dim m_BorderStyle As Integer
'事件声明:
Event DataIsChanged(index As Integer)
Event Click()
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick()
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "当用户在拥有焦点的对象上按下任意键时发生。"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "当用户按下和释放 ANSI 键时发生。"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "当用户在拥有焦点的对象上释放键时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mobj出诊记录集 As 出诊记录集
Private mobj所有合作单位 As 合作单位控制集
Private mobj所有门诊诊室 As 分诊诊室集
Private mblnNotClick As Boolean
Private mstrPreTabPage As String '上一个选择页面
Private mblnLoaded As Boolean
Private mblnShowFirstPage As Boolean '是否缺省显示第一页，否则显示最后一页

Public Function LoadData(ByVal obj出诊记录集 As 出诊记录集, _
    Optional ByVal obj所有门诊诊室 As 分诊诊室集, _
    Optional ByVal obj所有合作单位 As 合作单位控制集, _
    Optional ByVal blnShowFirstPage As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '入参:obj出诊号源-出诊号源信息
    '     obj所有门诊诊室-所有有效的门诊诊室
    '     obj所有合作单位-所有合作单位
    '     blnShowFirstPage-是否缺省显示第一页，否则显示最后一页
    '返回:加载成功, 返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-11 14:26:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj出诊记录集 = obj出诊记录集
    mblnShowFirstPage = blnShowFirstPage
    Set mobj所有门诊诊室 = obj所有门诊诊室: Set mobj所有合作单位 = obj所有合作单位
    LoadData = InitPageAndData   '加载页面及数据
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetClinicRecord(ByVal obj出诊记录集 As 出诊记录集, ByVal str时间段 As String) As 出诊记录
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据时间段来获取对应的出诊记录集
    '入参:obj出诊记录集-出诊记录集
    '     str时间段-时间段
    '返回:出诊记录对象
    '编制:刘兴洪
    '日期:2016-03-24 15:37:50
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj出诊记录 As 出诊记录
    If obj出诊记录集 Is Nothing Then Exit Function
    
    On Error GoTo errHandle
    For Each obj出诊记录 In obj出诊记录集
        If obj出诊记录.时间段 = str时间段 Then
            Set GetClinicRecord = obj出诊记录.Clone: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function InitPageAndData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2016-01-11 14:23:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim objWork As 上班时段, lngRow As Long
    Dim objPlan As 出诊记录, obj出诊记录 As 出诊记录
    Dim intPageCount As Integer, intSelectedPageIndex As Integer
    
    Err = 0: On Error GoTo Errhand:
    Call LockWindowUpdate(UserControl.Hwnd)
    
    tbPageTimeWork.RemoveAll
    mstrPreTabPage = "" '标记当前为未选择任何页签
    
    '缺省加载5个时段控件
'    If ClinicDetail.Count < 5 Then
'        For i = 1 To 4
'            Load picTimeWork(i): picTimeWork(i).Visible = True
'            Load ClinicDetail(i): ClinicDetail(i).Visible = True
'            Set ClinicDetail(i).Container = picTimeWork(i)
'        Next
'    End If
    intPageCount = ClinicDetail.Count
    
    If Not mobj出诊记录集 Is Nothing Then
        If mobj出诊记录集.Count > 0 Then
            lngRow = 0
            For Each objPlan In mobj出诊记录集
                If lngRow > intPageCount - 1 Then
                    Load picTimeWork(lngRow): picTimeWork(lngRow).Visible = True
                    Load ClinicDetail(lngRow): ClinicDetail(lngRow).Visible = True
                    Set ClinicDetail(lngRow).Container = picTimeWork(lngRow)
                End If
                Set ObjItem = tbPageTimeWork.InsertItem(lngRow + 1, objPlan.时间段, picTimeWork(lngRow).Hwnd, 0)
                ClinicDetail(0).EditMode = ED_RegistPlan_View
                lngRow = lngRow + 1
            Next
        Else
            '清除第一个页签的数据
            ClinicDetail(0).LoadData Nothing, Nothing
            ClinicDetail(0).EditMode = ED_RegistPlan_View
        End If
    End If
    If tbPageTimeWork.ItemCount = 0 Then
        lngRow = 0
        Set ObjItem = tbPageTimeWork.InsertItem(lngRow + 1, "无上班时段", picTimeWork(lngRow).Hwnd, 0)
        ClinicDetail(lngRow).EditMode = ED_RegistPlan_View:
    End If
    Call LockWindowUpdate(0)
    
    With tbPageTimeWork
        If mblnShowFirstPage Then
            intSelectedPageIndex = 0
        Else
            intSelectedPageIndex = tbPageTimeWork.ItemCount - 1
        End If
        .Enabled = False
        .Item(intSelectedPageIndex).Selected = True
        '手动触发SelectedChanged事件
        Call tbPageTimeWork_SelectedChanged(.Item(intSelectedPageIndex))
        .Enabled = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Color = xtpTabColorVisualStudio
    End With
    mblnLoaded = True
    InitPageAndData = True
    Exit Function
Errhand:
    tbPageTimeWork.Visible = True
    mblnNotClick = False
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    shpLine.Visible = New_BorderStyle <> 0
    UserControl_Resize
End Property

Private Sub ClinicDetail_DataIsChanged(index As Integer)
    RaiseEvent DataIsChanged(index)
End Sub

'根据页签名称获取页签索引
Public Property Get ItemIndex(ByVal Caption As String) As Integer
    Dim i As Integer
    
    ItemIndex = -1
    For i = 0 To tbPageTimeWork.ItemCount - 1
        If tbPageTimeWork.Item(i).Caption = Caption Then
            ItemIndex = i: Exit For
        End If
    Next
End Property

Private Sub picTimeWork_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picTimeWork(index)
        ClinicDetail(index).Left = .ScaleLeft
        ClinicDetail(index).Top = .ScaleTop
        ClinicDetail(index).Height = .ScaleHeight - 10
        ClinicDetail(index).Width = .ScaleWidth - 10
    End With
End Sub
 
Private Sub tbPageTimeWork_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim obj出诊记录 As 出诊记录
    
    If mstrPreTabPage = Item.Caption Then Exit Sub
    
    mstrPreTabPage = Item.Caption
    If Val(Item.Tag) = 1 Then Exit Sub
    
    Item.Tag = "1"
    '加载数据
    If mobj出诊记录集 Is Nothing Then Exit Sub
    If mobj出诊记录集.Exits("K" & Item.Caption) = False Then Exit Sub
    Set obj出诊记录 = mobj出诊记录集("K" & Item.Caption) 'GetClinicRecord(mobj出诊记录集, Item.Caption)
    If obj出诊记录 Is Nothing Then Exit Sub
    
    ClinicDetail(Item.index).LoadData obj出诊记录, mobj所有合作单位, mobj所有门诊诊室
    ClinicDetail(Item.index).EditMode = m_EditMode
    If obj出诊记录.是否固定 Then ClinicDetail(Item.index).EditMode = ED_RegistPlan_View
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_BorderStyle = m_def_BorderStyle
    m_EditMode = m_def_EditMode
    m_BackStyle = m_def_BackStyle
    m_诊疗频次 = m_def_诊疗频次
    Call Set诊疗频次(m_诊疗频次)
    m_医生姓名 = m_def_医生姓名
    Call Set医生姓名(m_医生姓名)
    mblnLoaded = False
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    
    shpLine.Visible = m_BorderStyle <> 0
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
    
    Call ReSetPageEditMode
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_诊疗频次 = PropBag.ReadProperty("诊疗频次", m_def_诊疗频次)
    Call Set诊疗频次(m_诊疗频次)
    m_医生姓名 = PropBag.ReadProperty("医生姓名", m_def_医生姓名)
    Call Set医生姓名(m_医生姓名)
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With shpLine
        .Top = ScaleTop
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    With tbPageTimeWork
        .Top = IIf(shpLine.Visible, 10, 0)
        .Left = IIf(shpLine.Visible, 10, 0)
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top * 2
    End With
End Sub
 
Private Sub UserControl_Show()
    If mblnLoaded Then Exit Sub
    Call InitPageAndData
End Sub

Private Sub UserControl_Terminate()
    Set mobj出诊记录集 = Nothing
    Set mobj所有合作单位 = Nothing
    Set mobj所有门诊诊室 = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("诊疗频次", m_诊疗频次, m_def_诊疗频次)
    Call PropBag.WriteProperty("医生姓名", m_医生姓名, m_def_医生姓名)
End Sub
 
 
Public Property Get Get出诊记录集() As 出诊记录集
    Dim obj出诊记录集 As New 出诊记录集, obj出诊记录 As New 出诊记录
    Dim intPage As Integer
    
    On Error GoTo Errhand
    If mobj出诊记录集 Is Nothing Then Exit Property
    If mobj出诊记录集.出诊日期 <> "" Then
        obj出诊记录集.出诊日期 = mobj出诊记录集.出诊日期
        For intPage = 0 To tbPageTimeWork.ItemCount - 1
            If tbPageTimeWork(intPage).Caption <> "无上班时段" And tbPageTimeWork(intPage).Caption <> "" Then
                If Val(tbPageTimeWork.Item(intPage).Tag) = 1 Then
                    Set obj出诊记录 = ClinicDetail(intPage).Get出诊记录
                    If obj出诊记录.是否修改 Then obj出诊记录集.是否修改 = True
                Else
                    '未加载的数据
                    Set obj出诊记录 = GetClinicRecord(mobj出诊记录集, tbPageTimeWork.Item(intPage).Caption)
                End If
                obj出诊记录集.AddItem obj出诊记录, "K" & obj出诊记录.时间段
            End If
        Next
    End If
    Set Get出诊记录集 = obj出诊记录集
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Function IsValied() As Boolean
    '检查数据
    Dim intPage As Integer
    
    Err = 0: On Error GoTo errHandler
    For intPage = 0 To tbPageTimeWork.ItemCount - 1
        If ClinicDetail(intPage).IsValied() = False Then
            tbPageTimeWork.Enabled = False
            tbPageTimeWork(intPage).Selected = True
            tbPageTimeWork.Enabled = True
            Exit Function
        End If
    Next
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=24,0,0,0
Public Property Get EditMode(ByVal index As Integer) As gRegistPlanEditMode
    If index < 0 Or index > tbPageTimeWork.ItemCount - 1 Then Exit Property
    EditMode = ClinicDetail(index).EditMode
End Property

Public Property Let EditMode(Optional ByVal index As Integer = -1, ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    If index = -1 Then
        '设置所有页签
        For index = 0 To tbPageTimeWork.ItemCount - 1
            If tbPageTimeWork(index).Caption = "无上班时段" Then
                '已设置可用状态，且不能改
                ClinicDetail(index).EditMode = ED_RegistPlan_View
            ElseIf Not mobj出诊记录集 Is Nothing Then
                If index < mobj出诊记录集.Count Then
                    ClinicDetail(index).EditMode = m_EditMode
                    If mobj出诊记录集(index + 1).是否固定 Then ClinicDetail(index).EditMode = ED_RegistPlan_View
                End If
            Else
                ClinicDetail(index).EditMode = m_EditMode
            End If
        Next
        Exit Property
    End If
    
    If index < 0 Or index > tbPageTimeWork.ItemCount - 1 Then Exit Property
    If Not mobj出诊记录集 Is Nothing Then
        If index < mobj出诊记录集.Count Then
            ClinicDetail(index).EditMode = m_EditMode
            If mobj出诊记录集(index + 1).是否固定 Then ClinicDetail(index).EditMode = ED_RegistPlan_View
        End If
    Else
        ClinicDetail(index).EditMode = m_EditMode
    End If
End Property

Private Sub ReSetPageEditMode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置页面的编辑模式
    '编制:刘兴洪
    '日期:2016-03-25 15:30:59
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.UBound
        ClinicDetail(i).EditMode = m_EditMode
    Next
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,5
Public Property Get 诊疗频次() As Integer
    诊疗频次 = m_诊疗频次
End Property

Public Property Let 诊疗频次(ByVal New_诊疗频次 As Integer)
    m_诊疗频次 = New_诊疗频次
    PropertyChanged "诊疗频次"
    Call Set诊疗频次(m_诊疗频次)

End Property
Private Sub Set诊疗频次(ByVal int诊疗频次 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新设置诊疗频次
    '入参:int诊疗频次
    '编制:刘兴洪
    '日期:2016-03-30 16:25:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.Count
        ClinicDetail(i).诊疗频次 = int诊疗频次
        mobj出诊记录集(i).号序信息集.出诊频次 = int诊疗频次
    Next
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,5
Public Property Get 医生姓名() As String
    医生姓名 = m_医生姓名
End Property

Public Property Let 医生姓名(ByVal New_医生姓名 As String)
    m_医生姓名 = New_医生姓名
    PropertyChanged "医生姓名"
    Call Set医生姓名(m_医生姓名)

End Property
Private Sub Set医生姓名(ByVal str医生姓名 As String)
    '功能:重新设置医生姓名
    '入参:str医生姓名
    Dim i As Integer
    Err = 0: On Error Resume Next
    For i = 0 To ClinicDetail.Count - 1
        ClinicDetail(i).医生姓名 = str医生姓名
        If mobj出诊记录集.Count > i Then
            mobj出诊记录集(i + 1).医生姓名 = str医生姓名
            mobj出诊记录集(i + 1).安排门诊诊室集.医生姓名 = str医生姓名
        End If
    Next
End Sub
