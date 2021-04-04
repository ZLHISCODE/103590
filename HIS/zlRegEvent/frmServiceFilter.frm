VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmServiceFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "消息过滤"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   60
      TabIndex        =   10
      Top             =   2895
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2730
      Left            =   75
      TabIndex        =   8
      Top             =   0
      Width           =   5310
      Begin VB.Frame Frame3 
         Caption         =   "时间范围"
         Height          =   795
         Left            =   225
         TabIndex        =   13
         Top             =   165
         Width           =   4935
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   2535
            TabIndex        =   1
            Top             =   330
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   103874563
            CurrentDate     =   42338
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   120
            TabIndex        =   0
            Top             =   330
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   103874563
            CurrentDate     =   42328
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   2280
            TabIndex        =   14
            Top             =   390
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "消息类型"
         Height          =   675
         Left            =   225
         TabIndex        =   12
         Top             =   1095
         Width           =   4935
         Begin VB.CheckBox chkType 
            Caption         =   "停诊"
            Height          =   345
            Index           =   0
            Left            =   675
            TabIndex        =   2
            Top             =   225
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkType 
            Caption         =   "替诊"
            Height          =   345
            Index           =   1
            Left            =   1920
            TabIndex        =   3
            Top             =   225
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkType 
            Caption         =   "预约登记"
            Height          =   345
            Index           =   2
            Left            =   3210
            TabIndex        =   4
            Top             =   225
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkShowRead 
         Caption         =   "显示已经处理完成的消息"
         Height          =   180
         Left            =   225
         TabIndex        =   6
         Top             =   2340
         Width           =   2310
      End
      Begin VB.ComboBox cbo登记人 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   840
         TabIndex        =   5
         Text            =   "cbo登记人"
         Top             =   1875
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   1935
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   9
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2925
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
End
Attribute VB_Name = "frmServiceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean

Private Sub cbo登记人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkShowRead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDef_Click()
    Dim i As Integer
    dtpEnd.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    dtpBegin.Value = CDate(Format(zlDatabase.Currentdate - 7, "yyyy-mm-dd 00:00:00"))
    chkType(0).Value = 1
    chkType(1).Value = 1
    chkType(2).Value = 1
    chkShowRead.Value = 0
    cbo登记人.ListIndex = 0
'    For i = 0 To cbo登记人.ListCount - 1
'        If NeedName(cbo登记人.List(i)) = UserInfo.姓名 Then cbo登记人.ListIndex = i: Exit For
'    Next
End Sub

Private Sub cmdOK_Click()
    If dtpEnd.Value <= dtpBegin.Value Then
        MsgBox "过滤的开始时间不能大于结束时间!", vbInformation, gstrSysName
        Exit Sub
    End If
    If chkType(0).Value = 0 And chkType(0).Value = 0 And chkType(0).Value = 0 Then
        MsgBox "请至少选择一项消息类型!", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True
    Me.Hide
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
    dtpEnd.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    dtpBegin.Value = CDate(Format(zlDatabase.Currentdate - 7, "yyyy-mm-dd 00:00:00"))
    dtpBegin.MaxDate = dtpEnd.Value
    Call LoadData
End Sub

Public Function Get消息类型() As String
    Dim strMes As String
    With chkType
        If .Item(0).Value Then strMes = strMes & ",1"
        If .Item(1).Value Then strMes = strMes & ",2"
        If .Item(2).Value Then strMes = strMes & ",3"
        If strMes <> "" Then strMes = Mid(strMes, 2)
    End With
    Get消息类型 = strMes
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：加载基础数据
    '编制：刘尔旋
    '日期：2016-01-11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset

    Set rsTemp = GetPersonnel("", True)

    cbo登记人.Clear
    cbo登记人.AddItem "所有登记人-"
    If rsTemp.RecordCount > 0 Then
        Call rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            cbo登记人.AddItem rsTemp!简码 & "-" & rsTemp!姓名
'            If Nvl(rsTemp!姓名) = UserInfo.姓名 Then cbo登记人.ListIndex = cbo登记人.NewIndex
            rsTemp.MoveNext
        Next
    End If
    cbo登记人.ListIndex = 0
    LoadData = True
    
End Function
