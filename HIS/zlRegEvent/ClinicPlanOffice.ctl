VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ClinicPlanOffice 
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   ScaleHeight     =   5415
   ScaleWidth      =   7725
   Begin MSComctlLib.ListView lvwDoctorRoom 
      Height          =   2925
      Left            =   2130
      TabIndex        =   5
      Top             =   1140
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   5159
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.OptionButton opt分诊方式 
      Caption         =   "不分诊"
      Height          =   300
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   50
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton opt分诊方式 
      Caption         =   "指定诊室"
      Height          =   300
      Index           =   1
      Left            =   1710
      TabIndex        =   2
      Top             =   50
      Width           =   1035
   End
   Begin VB.OptionButton opt分诊方式 
      Caption         =   "动态分诊"
      Height          =   300
      Index           =   2
      Left            =   2775
      TabIndex        =   3
      Top             =   50
      Width           =   1035
   End
   Begin VB.OptionButton opt分诊方式 
      Caption         =   "平均分诊"
      Height          =   300
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   50
      Width           =   1035
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H80000000&
      Height          =   3975
      Left            =   330
      Top             =   540
      Width           =   6225
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      Caption         =   "分诊方式"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "ClinicPlanOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'属性变量:
Dim m_EditMode As gRegistPlanEditMode
Dim m_医生姓名 As String
Dim m_IsDataChanged As Boolean

'缺省属性值:
Const m_def_EditMode = 0
Const m_def_医生姓名 = ""
Const m_def_IsDataChanged = False

Private mblnNotClick As Boolean
Private mobj分诊诊室集 As 分诊诊室集
Private mobj所有分诊诊室 As 分诊诊室集
'事件声明:
Event DataIsChanged()


Public Function LoadData(ByVal obj分诊诊室集 As 分诊诊室集, Optional ByVal obj所有分诊诊室 As 分诊诊室集, _
    Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载分诊诊室
    '入参:
    '       obj分诊诊室集 - 分诊诊室集
    '       obj所有分诊诊室 - 所有分诊诊室集 ,不传表示查看
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOK As Boolean
    Err = 0: On Error GoTo Errhand:
    Set mobj分诊诊室集 = obj分诊诊室集
    If mobj分诊诊室集 Is Nothing Then Set mobj分诊诊室集 = New 分诊诊室集
    Set mobj所有分诊诊室 = obj所有分诊诊室
    
    m_IsDataChanged = blnChanged
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetEnabled分诊方式(ByVal str医生姓名 As String)
    If str医生姓名 = "" Then
        opt分诊方式(2).Enabled = m_EditMode = ED_RegistPlan_Edit
        opt分诊方式(2).Tag = ""
        opt分诊方式(3).Enabled = m_EditMode = ED_RegistPlan_Edit
        opt分诊方式(3).Tag = ""
    Else
        opt分诊方式(2).Enabled = False
        opt分诊方式(2).Tag = "1"
        opt分诊方式(3).Enabled = False
        opt分诊方式(3).Tag = "1"
        If opt分诊方式(2).Value Or opt分诊方式(3).Value Then
            opt分诊方式(0).Value = True
        End If
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 15:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objRoom As 分诊诊室, objCheckedRoom As 分诊诊室
    Dim intCol As Integer, intRow As Integer
    Dim ObjItem As ListItem
    
    Err = 0: On Error GoTo Errhand:
    If mobj分诊诊室集 Is Nothing Then Set mobj分诊诊室集 = New 分诊诊室集
    
    mblnNotClick = True
    opt分诊方式(mobj分诊诊室集.分诊方式).Value = True
    医生姓名 = mobj分诊诊室集.医生姓名
    mblnNotClick = False
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And mobj分诊诊室集.分诊方式 <> 0
    
    With lvwDoctorRoom
        .ListItems.Clear
        .Refresh
        If mobj所有分诊诊室 Is Nothing Then
            '只加载已选择诊室
            For Each objCheckedRoom In mobj分诊诊室集
                Set ObjItem = .ListItems.Add(, "K" & objCheckedRoom.诊室ID, objCheckedRoom.诊室名称)
                ObjItem.SubItems(1) = objCheckedRoom.诊室ID
                ObjItem.Checked = True
            Next
        Else
            For Each objRoom In mobj所有分诊诊室
                Set ObjItem = .ListItems.Add(, "K" & objRoom.诊室ID, objRoom.诊室名称)
                ObjItem.SubItems(1) = objRoom.诊室ID
                '加载已选择诊室
                For Each objCheckedRoom In mobj分诊诊室集
                    If objRoom.诊室ID = objCheckedRoom.诊室ID Then
                        ObjItem.Checked = True: Exit For
                    End If
                Next
            Next
        End If
    End With
    lvwDoctorRoom.BackColor = lvwDoctorRoom.BackColor
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get分诊诊室集() As 分诊诊室集
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取门诊诊室集
    '出参:
    '返回:获取成功，返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim objRoom As 分诊诊室, obj分诊诊室集 As New 分诊诊室集
    
    Err = 0: On Error GoTo Errhand:
    '数据未改变，直接返回原集合的副本
    If m_IsDataChanged = False Then
        Set Get分诊诊室集 = mobj分诊诊室集.Clone
        Exit Function
    End If
    
    '数据已改变，重新构造集合对象
    With obj分诊诊室集
        .分诊方式 = GetSelectedIndex(opt分诊方式)
        .医生姓名 = 医生姓名
        .是否修改 = True
    End With
    For i = 1 To lvwDoctorRoom.ListItems.Count
        If lvwDoctorRoom.ListItems(i).Checked Then
            Set objRoom = New 分诊诊室
            With objRoom
                .诊室ID = lvwDoctorRoom.ListItems(i).SubItems(1)
                .诊室名称 = lvwDoctorRoom.ListItems(i).Text
            End With
            obj分诊诊室集.AddItem objRoom, "K" & objRoom.诊室ID
        End If
    Next
    Set Get分诊诊室集 = obj分诊诊室集
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub lvwDoctorRoom_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        '指定诊室只能选一个
        If opt分诊方式(1).Value Then Call ClearAllGridChecked
        Item.Checked = True
    End If
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub lvwDoctorRoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub opt分诊方式_GotFocus(index As Integer)
    opt分诊方式(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub opt分诊方式_LostFocus(index As Integer)
     opt分诊方式(index).BackColor = Me.BackColor
End Sub

 
Private Sub UserControl_Initialize()
    Call InitFace
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_IsDataChanged = m_def_IsDataChanged
    m_医生姓名 = m_def_医生姓名
    m_EditMode = m_def_EditMode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_医生姓名 = PropBag.ReadProperty("医生姓名", m_def_医生姓名)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    With shpBack
        .Left = 0
        .Top = opt分诊方式(0).Top + opt分诊方式(0).Height + 60
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top
    End With
    With lvwDoctorRoom
        .Left = 10
        .Top = shpBack.Top + 10
        .Width = ScaleWidth - .Left - 20
        .Height = ScaleHeight - .Top - 10
    End With
End Sub

Public Property Get Get安排门诊诊室集() As 分诊诊室集
    Set Get安排门诊诊室集 = Get分诊诊室集
End Property

Public Function IsValied() As Boolean
    '检查数据
    Dim intCount As Integer, i As Long, j As Long
    
    Err = 0: On Error GoTo errHandler
    '数据未改变不检查
    If m_IsDataChanged = False Or m_EditMode <> ED_RegistPlan_Edit Then IsValied = True: Exit Function
    
    '诊室判断
    If opt分诊方式(0).Value = False Then
        '不分诊时才检查
        For i = 1 To lvwDoctorRoom.ListItems.Count
            If lvwDoctorRoom.ListItems(i).Checked Then
                intCount = intCount + 1
            End If
        Next
        If opt分诊方式(1).Value Then
            '指定诊室有且只能选择一个
            If intCount = 0 Then
                MsgBox "指定诊室时必须选择一个对应的门诊诊室！", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            ElseIf intCount > 1 Then
                MsgBox "指定诊室时只能选择一个对应的门诊诊室！", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            End If
        Else
            If intCount < 2 Then
                MsgBox "动态分诊或平均分诊时至少要选择两个对应的门诊诊室！", vbInformation, gstrSysName
                If lvwDoctorRoom.Visible And lvwDoctorRoom.Enabled Then lvwDoctorRoom.SetFocus
                Exit Function
            End If
        End If
    End If
    IsValied = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub opt分诊方式_Click(index As Integer)
    Dim i As Integer
    
    Err = 0: On Error GoTo errHandler
    If mblnNotClick Then Exit Sub
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And index <> 0
    lvwDoctorRoom.BackColor = lvwDoctorRoom.BackColor
    Call ClearAllGridChecked(index)
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ClearAllGridChecked(Optional ByVal byt分诊方式 As Byte)
    '功能：清空选择项目
    Dim i As Integer, j As Integer
    Dim intSelectedCount As Integer
    
    Err = 0: On Error GoTo errHandler
    For i = 1 To lvwDoctorRoom.ListItems.Count
        Select Case byt分诊方式
        Case 0 '不分诊
            lvwDoctorRoom.ListItems(i).Checked = False
        Case 1 '指定诊室
            If intSelectedCount = 1 Then lvwDoctorRoom.ListItems(i).Checked = False
            If lvwDoctorRoom.ListItems(i).Checked Then intSelectedCount = 1
        Case 2 '动态分诊
        Case 3 '平均分诊
        End Select
    Next
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub opt分诊方式_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub InitFace()
    Err = 0: On Error GoTo errHandler
    With lvwDoctorRoom
        .Checkboxes = True
        .FullRowSelect = True
        .GridLines = False
        .HideSelection = True
        .LabelEdit = lvwManual
        .MultiSelect = True
        .View = lvwList
        .TextBackground = lvwTransparent
        
        '添加列
        .ColumnHeaders.Add , "K_名称", "诊室名称", 3500
        .ColumnHeaders.Add , "K_ID", "诊室ID", 0
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub UserControl_Show()
    lvwDoctorRoom.View = lvwReport
    lvwDoctorRoom.View = lvwList
End Sub

Private Sub UserControl_Terminate()
    Set mobj分诊诊室集 = Nothing
    Set mobj所有分诊诊室 = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("医生姓名", m_医生姓名, m_def_医生姓名)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Dim i As Integer
    For i = opt分诊方式.LBound To opt分诊方式.UBound
        opt分诊方式(i).BackColor = New_BackColor
    Next
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,1,1,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get 医生姓名() As String
    医生姓名 = m_医生姓名
End Property

Public Property Let 医生姓名(ByVal New_医生姓名 As String)
    m_医生姓名 = New_医生姓名
    PropertyChanged "医生姓名"
    
    m_医生姓名 = New_医生姓名
    SetEnabled分诊方式 m_医生姓名
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    m_EditMode = New_EditMode
    PropertyChanged "EditMode"
    
    Dim i As Integer
    For i = opt分诊方式.LBound To opt分诊方式.UBound
        'opt分诊方式(i).Tag = "1"表示不能修改状态
        opt分诊方式(i).Enabled = m_EditMode = ED_RegistPlan_Edit And opt分诊方式(i).Tag = ""
    Next
    lvwDoctorRoom.Enabled = m_EditMode = ED_RegistPlan_Edit And opt分诊方式(0).Value = False
End Property

