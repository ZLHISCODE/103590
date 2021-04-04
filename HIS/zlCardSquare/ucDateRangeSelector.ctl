VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ucDateRangeSelector 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   315
   ScaleWidth      =   6000
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   795
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   15
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   178782211
      CurrentDate     =   40777
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   285
      Left            =   3765
      TabIndex        =   2
      Top             =   15
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   178782211
      CurrentDate     =   40777
   End
   Begin VB.Label lblDateShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1900-01-01 00:00:00 ～ 1900-01-01 23:59:59"
      Height          =   180
      Left            =   2160
      TabIndex        =   5
      Top             =   60
      Width           =   3780
   End
   Begin VB.Label lblSplit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      Height          =   180
      Left            =   3540
      TabIndex        =   4
      Top             =   60
      Width           =   180
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缺省显示"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "ucDateRangeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'事件声明:
Event ValueChanged(ByVal dtBegin As Date, ByVal dtEnd As Date)
'属性变量:
Dim m_BeginTime As Date
Dim m_EndTime As Date
Dim m_ListIndex As Integer


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

Private Sub cboDate_Click()
    dtpStartDate.Visible = cboDate.ListIndex = 5
    lblSplit.Visible = cboDate.ListIndex = 5
    dtpEndDate.Visible = cboDate.ListIndex = 5
    lblDateShow.Visible = cboDate.ListIndex <> 5
    
    m_EndTime = CDate(Format(Now(), "yyyy-mm-dd") & " 23:59:59")
    Select Case cboDate.ListIndex
        Case 0 '今日
            m_BeginTime = CDate(Format(m_EndTime, "yyyy-mm-dd"))
        Case 1 '最近2天
            m_BeginTime = CDate(Format(DateAdd("d", -1, m_EndTime), "yyyy-mm-dd"))
        Case 2 '最近3天
            m_BeginTime = CDate(Format(DateAdd("d", -2, m_EndTime), "yyyy-mm-dd"))
        Case 3  '最近一周
            m_BeginTime = CDate(Format(DateAdd("d", -7, m_EndTime), "yyyy-mm-dd"))
        Case 4  '本月
            m_BeginTime = CDate(Format(m_EndTime, "yyyy-mm") & "-01")
        Case Else
            m_BeginTime = CDate(Format(dtpStartDate.value, "yyyy-mm-dd"))
            m_EndTime = CDate(Format(dtpEndDate.value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    lblDateShow.Caption = Format(m_BeginTime, "yyyy-mm-dd HH:mm:ss")
    lblDateShow.Caption = lblDateShow.Caption & " ～ " & Format(m_EndTime, "yyyy-mm-dd HH:mm:ss")
    
    RaiseEvent ValueChanged(m_BeginTime, m_EndTime)
End Sub

Private Sub dtpEndDate_Change()
    m_BeginTime = dtpStartDate.value: m_EndTime = dtpEndDate.value
    RaiseEvent ValueChanged(m_BeginTime, m_EndTime)
End Sub

Private Sub dtpStartDate_Change()
    m_BeginTime = dtpStartDate.value: m_EndTime = dtpEndDate.value
    RaiseEvent ValueChanged(m_BeginTime, m_EndTime)
End Sub

Private Sub UserControl_Initialize()
    With cboDate
        .Clear
        .AddItem "今日"
        .AddItem "最近两天"
        .AddItem "最近三天"
        .AddItem "最近一周"
        .AddItem "本月"
        .AddItem "自定义"
    End With
    ListIndex = 0
    dtpStartDate.value = CDate(Format(Now(), "yyyy-mm-dd"))
    dtpEndDate.value = CDate(Format(Now(), "yyyy-mm-dd") & " 23:59:59")
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ListIndex = PropBag.ReadProperty("ListIndex", 0)
    
    cboDate.Enabled = UserControl.Enabled
    dtpStartDate.Enabled = UserControl.Enabled
    dtpEndDate.Enabled = UserControl.Enabled
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 6000
    UserControl.Height = 300
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ListIndex", m_ListIndex, 0)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get ListIndex() As Integer
    ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    If New_ListIndex < 0 Or New_ListIndex > cboDate.ListCount - 1 Then Exit Property
    cboDate.ListIndex = New_ListIndex
    
    m_ListIndex = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=3,0,0,2011/8/22
Public Property Get BeginTime() As Date
    BeginTime = m_BeginTime
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=3,0,0,3000-01-01
Public Property Get EndTime() As Date
    EndTime = m_EndTime
End Property

