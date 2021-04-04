VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDistRoom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "重新分配诊室"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmb 
      Height          =   360
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   1920
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   390
      Left            =   4680
      TabIndex        =   6
      Top             =   420
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   390
      Left            =   4680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   915
      Width           =   1350
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   390
      Left            =   4680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1350
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   4095
      Left            =   90
      TabIndex        =   5
      Tag             =   "可变化的"
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "诊室"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "位置"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   450
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   103481347
      CurrentDate     =   36588
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分诊时间"
      Height          =   240
      Left            =   2160
      TabIndex        =   2
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "选择诊室"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "选择医生"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "frmDistRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrRoom As String
Private Const STR_COMP = "|',~" '分隔字符串
Private mblnExit As Boolean
Private mbln签到 As Boolean
Private mlng挂号ID As Long

Public Sub ShowMe(strRoom As String, frmParent As Form, Optional bln签到 As Boolean = False, _
                    Optional lng挂号ID As Long)
    '显示本窗体并返回选择的诊室
    '
    mbln签到 = bln签到
    mlng挂号ID = lng挂号ID
    If mbln签到 Then Me.Caption = "重新分配诊室『病人签到』"
    Me.Show 1, frmParent
    strRoom = mstrRoom
End Sub

Private Sub cmb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mstrRoom = STR_COMP
    mblnExit = True
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    '95637:李南春,2016/7/18,不分诊也支持签到
    If lvwMain.ListItems.Count > 0 And lvwMain.SelectedItem Is Nothing Then
        MsgBox "请选择一个诊室！", vbInformation, gstrSysName
        Exit Sub
'    ElseIf lvwMain.SelectedItem Is Nothing And lvwMain.ListItems.Count = 0 Then
'        MsgBox "无诊室可以选择！", vbInformation, gstrSysName
'        Exit Sub
    End If
    
    If lvwMain.SelectedItem Is Nothing Then
        If ExcPlugInFun(0, mlng挂号ID, IIf(cmb.Text = "无", "", zlCommFun.GetNeedName(cmb.Text))) = False Then Exit Sub
    Else
        If ExcPlugInFun(0, mlng挂号ID, IIf(cmb.Text = "无", "", zlCommFun.GetNeedName(cmb.Text)), lvwMain.SelectedItem.Text) = False Then Exit Sub
    End If
    
    If cmb.Text = "无" Then
        If lvwMain.SelectedItem Is Nothing Then
            mstrRoom = "" & STR_COMP & " " & STR_COMP & Format(dtpBegin.Value, "YYYY-MM-DD HH:MM:SS")
        Else
            mstrRoom = lvwMain.SelectedItem.Text & STR_COMP & " " & STR_COMP & Format(dtpBegin.Value, "YYYY-MM-DD HH:MM:SS")
        End If
    Else
        If lvwMain.SelectedItem Is Nothing Then
            mstrRoom = "" & STR_COMP & zlCommFun.GetNeedName(cmb.Text) & STR_COMP & Format(dtpBegin.Value, "YYYY-MM-DD HH:MM:SS")
        Else
            mstrRoom = lvwMain.SelectedItem.Text & STR_COMP & zlCommFun.GetNeedName(cmb.Text) & STR_COMP & Format(dtpBegin.Value, "YYYY-MM-DD HH:MM:SS")
        End If
    End If
    Beep
    Beep
    mblnExit = True
    Unload Me
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Me.lvwMain.Sorted = False
    mblnExit = False
    If mbln签到 Then Me.Caption = "重新分配诊室『病人签到』"
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnExit Then Call cmdCancel_Click
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.index - 1
    End If
End Sub

Private Sub lvwMain_DblClick()
    If lvwMain.ListItems.Count > 0 And Not (lvwMain.SelectedItem Is Nothing) Then
        cmdOK_Click
    End If
End Sub

Private Sub lvwMain_KeyPress(KeyAscii As Integer)
    '96764:李南春,2016/7/18,分诊快速定位诊室
    Dim intSequence As Integer
    intSequence = Chr(KeyAscii)
    If 0 < intSequence And intSequence < 10 Then
        If lvwMain.ListItems.Count >= intSequence Then
            lvwMain.ListItems.Item(intSequence).Selected = True
            lvwMain.ListItems.Item(intSequence).EnsureVisible
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
