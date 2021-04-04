VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockSeating 
   BorderStyle     =   0  'None
   ClientHeight    =   5220
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDockSeating.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picWE 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3750
      Left            =   6390
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3750
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   390
      Width           =   45
   End
   Begin VB.PictureBox picNS2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   6345
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   2520
      TabIndex        =   1
      Top             =   2595
      Width           =   2520
   End
   Begin VB.PictureBox picNS1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   2820
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   3765
      Width           =   3825
   End
   Begin VB.PictureBox picSeating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Index           =   3
      Left            =   4755
      ScaleHeight     =   1950
      ScaleWidth      =   2490
      TabIndex        =   11
      Top             =   2910
      Width           =   2490
      Begin MSComctlLib.ListView lvwSeating 
         Height          =   1455
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   285
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2566
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "座位"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "性别"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "门诊号"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "收费价格"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "状态"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "VIP座位"
         ForeColor       =   &H80000013&
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   12
         Top             =   45
         Width           =   1110
      End
   End
   Begin VB.PictureBox picSeating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000013&
      Height          =   2430
      Index           =   2
      Left            =   4770
      ScaleHeight     =   2430
      ScaleWidth      =   2505
      TabIndex        =   8
      Top             =   420
      Width           =   2505
      Begin MSComctlLib.ListView lvwSeating 
         Height          =   1965
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   345
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   3466
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "座位"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "性别"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "门诊号"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "收费价格"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "状态"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         Caption         =   "特殊药座位"
         ForeColor       =   &H80000013&
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   45
         Width           =   1110
      End
   End
   Begin VB.PictureBox picSeating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   -45
      ScaleHeight     =   1470
      ScaleWidth      =   4740
      TabIndex        =   5
      Top             =   3390
      Width           =   4740
      Begin MSComctlLib.ListView lvwSeating 
         Height          =   1080
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   315
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1905
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "座位"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "性别"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "门诊号"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "收费价格"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "状态"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "加座"
         ForeColor       =   &H80000013&
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.PictureBox picSeating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Index           =   0
      Left            =   30
      ScaleHeight     =   2940
      ScaleWidth      =   4665
      TabIndex        =   3
      Top             =   420
      Width           =   4665
      Begin MSComctlLib.ListView lvwSeating 
         Height          =   2565
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   240
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   4524
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Caption         =   "普通座位"
         ForeColor       =   &H80000013&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   45
         Width           =   6150
      End
   End
   Begin MSComctlLib.ImageList imgRpt 
      Left            =   2025
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":000C
            Key             =   "未执行"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":05A6
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":0B40
            Key             =   "拒绝执行"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":10DA
            Key             =   "正在执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":1674
            Key             =   "已报到"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":1C0E
            Key             =   "Calling"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2925
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":8470
            Key             =   "Add"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":ECD2
            Key             =   "Kong"
            Object.Tag             =   "Kong"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":15534
            Key             =   "YouRen"
            Object.Tag             =   "YouRen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":1BD96
            Key             =   "WeiXiu"
            Object.Tag             =   "WeiXiu"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":225F8
            Key             =   "Modi"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":28E5A
            Key             =   "Del"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":2F6BC
            Key             =   "View"
            Object.Tag             =   "200"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":3113E
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":379A0
            Key             =   "no"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":3E202
            Key             =   "Z_K"
            Object.Tag             =   "Z_K"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":44A64
            Key             =   "Z_Y"
            Object.Tag             =   "Z_Y"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":4B2C6
            Key             =   "Z_X"
            Object.Tag             =   "Z_X"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3525
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":51B28
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":5838A
            Key             =   "no"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":5EBEC
            Key             =   "Kong"
            Object.Tag             =   "Kong"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":6544E
            Key             =   "YouRen"
            Object.Tag             =   "YouRen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":6BCB0
            Key             =   "WeiXiu"
            Object.Tag             =   "WeiXiu"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":72512
            Key             =   "Z_K"
            Object.Tag             =   "Z_K"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":78D74
            Key             =   "Z_Y"
            Object.Tag             =   "Z_Y"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeating.frx":7F5D6
            Key             =   "Z_X"
            Object.Tag             =   "Z_X"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   60
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockSeating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Event Activate() '自已激活时
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字

Private Const subMenu_Add = 101
Private Const subMenu_Modify = 102
Private Const subMenu_Delete = 103

Private Const subMenu_View = 200
Private Const subMenu_Icon = 201
Private Const subMenu_List = 202
Private Const subMenu_Report = 203

Private Const subMenu_Clear = 300 '清除占用的座位
Private Const subMenu_SetSeating = 400 '安排座位

Private mintActiveLvw As Integer '当前选定的类别
Private mcurSeatings As Seatings

Public lng病人ID As Long '主窗体传进来， 用于安排座位
Public objPati As cPatient

Private mSourceItem As ListItem '换座时的源座位
Private mObjItem As ListItem '换座时的目标座位
Private mcbsMain As CommandBars

Private mSelectKey As String '当前选择的座位

Public Sub zlRefresh(ByVal curSeatings As Seatings)
    Dim strIcon As String, str状态 As String
    Dim lstItem As ListItem
    
    lvwSeating(0).ListItems.Clear
    lvwSeating(1).ListItems.Clear
    lvwSeating(2).ListItems.Clear
    lvwSeating(3).ListItems.Clear
    Dim curSeating As Seating
    
    Set mcurSeatings = Nothing
    Set mSourceItem = Nothing
    Set mObjItem = Nothing
    Set mcurSeatings = curSeatings
    
    For Each curSeating In mcurSeatings
        With curSeating
            If .状态 = 0 Then
                strIcon = "Z_K"
                If .类型 = 1 Then strIcon = "Kong"
                str状态 = "空座位"
            ElseIf .状态 = 1 Then
                strIcon = "Z_Y"
                If .类型 = 1 Then strIcon = "YouRen"
                str状态 = "被占用"
            Else
                strIcon = "Z_X"
                If .类型 = 1 Then strIcon = "WeiXiu"
                str状态 = "在维护"
            End If
            
            Set lstItem = lvwSeating(.类别).ListItems.Add(, .类别 & "_" & .编号, IIf(.姓名 = "", .编号, .编号 & ":" & .姓名), strIcon)
            lstItem.ListSubItems.Add , , .性别
            lstItem.ListSubItems.Add , , .门诊号
            lstItem.ListSubItems.Add , , Format(.现价, "0.00")
            lstItem.ListSubItems.Add , , str状态, strIcon
        End With
    Next
    If mSelectKey <> "" Then
        On Error Resume Next
        lvwSeating(mintActiveLvw).ListItems(mSelectKey).Selected = True
    End If
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case Else
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
     '#
    Dim StrKey As String, strObjKey As String
    
    Select Case Control.ID
        Case conMenu_Edit_Seat_Icon
            lvwSeating(mintActiveLvw).View = lvwIcon
        Case conMenu_Edit_Seat_Report
            lvwSeating(mintActiveLvw).View = lvwReport
        Case conMenu_Edit_Seat_List
            lvwSeating(mintActiveLvw).View = lvwList
        Case conMenu_Edit_Seat_Add
            If frmSeatingMana.SeatingMana(0, mcurSeatings, mintActiveLvw, "", Me) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Modify
            Call ModiSeat
        Case conMenu_Edit_Seat_Delete
            StrKey = lvwSeating(mintActiveLvw).SelectedItem.Key
            If mcurSeatings.Delete(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Set
            '安排座位
            If lng病人ID <> 0 And Not objPati Is Nothing Then
                StrKey = lvwSeating(mintActiveLvw).SelectedItem.Key
                If MsgBox("是否安排[" & objPati.姓名 & "]到[" & lvwSeating(mintActiveLvw).SelectedItem.Text & "]座位", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                    If mcurSeatings.SetSeating(lng病人ID, StrKey) Then
                        RaiseEvent RequestRefresh
                    End If
                End If
            End If
        Case conMenu_Edit_Seat_Clear
        '清除占用的座位
            StrKey = lvwSeating(mintActiveLvw).SelectedItem.Key
            If mcurSeatings.Clear(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Swap
            '换座位
            StrKey = lvwSeating(mintActiveLvw).SelectedItem.Key
            strObjKey = frmSeatingSwap.ObjectKey(StrKey, mcurSeatings, Me)
            If strObjKey <> "" Then
                If mcurSeatings.SwapSeating(StrKey, strObjKey) Then
                    RaiseEvent RequestRefresh
                End If
            End If
    End Select
End Sub

Private Sub ModiSeat()
    '修改座位
    Dim StrKey As String
    If Not lvwSeating(mintActiveLvw).SelectedItem Is Nothing Then
        If InStr(lvwSeating(mintActiveLvw).SelectedItem.Text, ":") <= 0 Then
            StrKey = lvwSeating(mintActiveLvw).SelectedItem.Key
            If frmSeatingMana.SeatingMana(1, mcurSeatings, mintActiveLvw, StrKey, Me) Then
                RaiseEvent RequestRefresh
            End If
        End If
    End If

End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim objItem As ListItem
    Set objItem = lvwSeating(mintActiveLvw).SelectedItem
    
    Select Case Control.ID
        Case conMenu_Edit_Seat_Modify, conMenu_Edit_Seat_Delete
            If objItem Is Nothing Then
                Control.Enabled = False
            ElseIf InStr(objItem.Text, ":") > 0 Then
                Control.Enabled = False
            Else
                Select Case mintActiveLvw
                Case 0
                    Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
                Case 1
                    Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
                Case 2
                    Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
                Case 3
                    Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
                End Select
            End If
        Case conMenu_Edit_Seat_Add
            Select Case mintActiveLvw
            Case 0
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
            Case 1
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
            Case 2
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
            Case 3
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
            End Select
        Case conMenu_Edit_Seat_Icon
            Control.Checked = lvwSeating(mintActiveLvw).View = lvwIcon
        Case conMenu_Edit_Seat_List
            Control.Checked = lvwSeating(mintActiveLvw).View = lvwList
        Case conMenu_Edit_Seat_Report
            Control.Checked = lvwSeating(mintActiveLvw).View = lvwReport
        Case conMenu_Edit_Seat_Set
            If Not objItem Is Nothing Then
                Control.Enabled = (lng病人ID <> 0 And (objItem.Icon = "Kong" Or objItem.Icon = "Z_K"))
            Else
                Control.Enabled = False
            End If
            If Control.Enabled Then
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0
            End If
        Case conMenu_Edit_Seat_Clear
            If Not objItem Is Nothing Then
                Control.Enabled = InStr(objItem, ":") > 0
            Else
                Control.Enabled = False
            End If
            If Control.Enabled Then
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0
            End If
        Case conMenu_Edit_Seat_Swap
            If Not objItem Is Nothing Then
                Control.Enabled = InStr(objItem, ":") > 0
            Else
                Control.Enabled = False
            End If
            If Control.Enabled Then
                Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0
            End If
    End Select
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As CommandBars, ByVal int场合 As Integer)
    '主窗体要求初始化主窗体上的菜单
    Dim objMenu As CommandBarPopup, objViewMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '病人项目的菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    Set mcbsMain.Icons = zlCommFun.GetPubIcons
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "座位管理(&S)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "安排座位(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "清除座位(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Swap, "调换座位(&W)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Add, "增加座位(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Modify, "修改座位(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Delete, "删除座位(&D)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_View_Seat, "座位图例")
        objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_GBed, "普通床位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_RBed, "占用床位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_YBed, "维护床位")
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Gseat, "普通座位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Rseat, "占用座位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Yseat, "维护座位")
        End With
    End With
    
    Set objViewMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    If objViewMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "查看(&V)")
            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "图标方式(&I)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "列表方式(&L)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "报表方式(&R)")
            End With
        End With
    Else
        With objViewMenu.CommandBar.Controls
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "查看方式(&V)")
            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "图标(&I)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "列表(&L)")
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "报表(&R)")
            End With
        End With
    
    End If
    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '先求出前面的最后一个Control
        
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Seat, "座位", objControl.Index + 1)
        objPopup.ID = conMenu_Edit_Seat: objPopup.BeginGroup = True
        
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "安排座位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Clear, "清除座位")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Swap, "调换座位")
            
        End With
        
        
    End With
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    'lngBottom = lngBottom - stbCutline.Height
    With picSeating(0)
        .Top = lngTop
        .Left = lngLeft
    End With
    picSeating_Resize (0)
    
    With picNS1
        .Top = lngTop + picSeating(0).Height
        .Left = lngLeft
        .Width = picSeating(0).Width
    End With
    
    With picWE
        .Top = lngTop
        .Left = lngLeft + picSeating(0).Width
        .Height = picSeating(0).Height + picSeating(1).Height + picNS1.Height
    End With
    
    With picSeating(1)
        .Top = picNS1.Top + picNS1.Height
        .Left = lngLeft
        .Height = lngBottom - .Top
        .Width = picNS1.Width
    End With
    
    With picSeating(2)
        .Top = lngTop
        .Left = picWE.Left + picWE.Width
        .Width = lngRight - picSeating(0).Width - 45
    End With
    With picNS2
        .Top = picSeating(2).Top + picSeating(2).Height
        .Left = lngLeft + picSeating(0).Width + 45
        .Width = lngRight - picSeating(0).Width - 45
    End With
    
    
    With picSeating(3)
        .Top = picNS2.Top + picNS2.Height
        .Left = picWE.Left + picWE.Width
        .Width = lngRight - .Left
        .Height = lngBottom - .Top
    End With
End Sub

Private Sub Form_Load()
'
    cbsSub.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsSub_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mSelectKey = ""
End Sub

Private Sub lvwSeating_DblClick(Index As Integer)
    '#
    Call ModiSeat
End Sub

Private Sub lvwSeating_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim strObjKey As String
    
    If TypeName(Source) = "ListView" And Not lvwSeating(Index).DropHighlight Is Nothing Then
        '放下
        Set mObjItem = lvwSeating(Index).DropHighlight
        lvwSeating(Index).MousePointer = ccDefault
        If Not mSourceItem Is Nothing And Not mObjItem Is Nothing Then
            strObjKey = frmSeatingSwap.ObjectKey(mSourceItem.Key, mcurSeatings, Me, mObjItem.Key)
            If strObjKey <> "" Then
                If mcurSeatings.SwapSeating(mSourceItem.Key, strObjKey) = True Then
                    RaiseEvent RequestRefresh
                End If
            End If
            Set lvwSeating(Index).DropHighlight = Nothing
        End If
    End If

    If TypeName(Source) = "ReportControl" And Not lvwSeating(Index).DropHighlight Is Nothing And lng病人ID <> 0 Then
        '放下
        Set mObjItem = lvwSeating(Index).DropHighlight
        'lvwSeating(Index).MousePointer = ccDefault
        If Not mObjItem Is Nothing And Not objPati Is Nothing Then
            If MsgBox("是否安排[" & objPati.姓名 & "]到[" & mObjItem & "]座位", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                Call mcurSeatings.SetSeating(lng病人ID, mObjItem.Key)
                RaiseEvent RequestRefresh
            End If
            Set lvwSeating(Index).DropHighlight = Nothing
        End If
    End If
    
End Sub

Private Sub lvwSeating_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim objOver As ListItem
    If TypeName(Source) = "ListView" Then
        Set objOver = lvwSeating(Index).HitTest(X, Y)
        If Not objOver Is Nothing Then
            If objOver.Icon = "Kong" Or objOver.Icon = "Z_K" Then
                '经过空座位
                Set lvwSeating(Index).DropHighlight = objOver
                lvwSeating(Index).DropHighlight.EnsureVisible
                Set Source.DragIcon = img16.ListImages("yes").Picture
            Else
                Set Source.DragIcon = img16.ListImages("no").Picture
            End If
        Else
            Set lvwSeating(Index).DropHighlight = Nothing
            Set Source.DragIcon = img32.ListImages("YouRen").Picture
        End If
    End If
    If TypeName(Source) = "ReportControl" Then
        Set objOver = lvwSeating(Index).HitTest(X, Y)
        If Not objOver Is Nothing Then
            If (objOver.Icon = "Kong" Or objOver.Icon = "Z_K") And lng病人ID <> 0 Then
                Set lvwSeating(Index).DropHighlight = objOver
                lvwSeating(Index).DropHighlight.EnsureVisible
                Set Source.DragIcon = img16.ListImages("yes").Picture
            Else
                Set Source.DragIcon = img16.ListImages("no").Picture
            End If
        Else
            Set Source.DragIcon = imgRpt.ListImages("未执行").Picture
        End If
    End If
    
    If State = 1 Then Set lvwSeating(Index).DropHighlight = Nothing
End Sub

Private Sub lvwSeating_GotFocus(Index As Integer)
    Dim i As Integer
    For i = 0 To 3
        If i = Index Then
            lblTitle(Index).BackColor = vbActiveTitleBar
            lblTitle(Index).ForeColor = vbActiveTitleBarText
        Else
            lblTitle(i).BackColor = vbInactiveTitleBar
            lblTitle(i).ForeColor = vbInactiveTitleBarText
        End If
    Next
    
    mintActiveLvw = Index
    If lvwSeating(Index).SelectedItem Is Nothing Then Exit Sub
    Dim StrKey As String
    StrKey = lvwSeating(Index).SelectedItem.Key
    With mcurSeatings.Item(StrKey)
'        stbCutline.Panels(1).Text = .编号 & IIf(.病人ID = 0, "", " 病人ID:" & .病人ID) & _
'                                    IIf(Trim(.收费项目) = "", "", " 收费项目:" & .收费项目) & _
'                                    IIf(.现价 = 0, "", " 价格(次):" & Format(.现价, "0.00"))
        RaiseEvent StatusTextUpdate(.类别 & "_" & .编号)
    End With
End Sub

Private Sub lvwSeating_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim StrKey As String
    StrKey = lvwSeating(Index).SelectedItem.Key
    mSelectKey = StrKey
    With mcurSeatings.Item(StrKey)
'        stbCutline.Panels(1).Text = .编号 & IIf(.病人ID = 0, "", " 病人ID:" & .病人ID) & _
'                                    IIf(Trim(.收费项目) = "", "", " 收费项目:" & .收费项目) & _
'                                    IIf(.现价 = 0, "", " 价格(次):" & Format(.现价, "0.00"))
        RaiseEvent StatusTextUpdate(.类别 & "_" & .编号)
    End With
End Sub

Private Sub lvwSeating_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ModiSeat
    End If
End Sub

Private Sub lvwSeating_LostFocus(Index As Integer)
'    stbCutline.Panels(1).Text = ""
End Sub

Private Sub lvwSeating_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        If lvwSeating(Index).HitTest(X, Y) Is Nothing Then
'            stbCutline.Panels(1).Text = ""
        End If
    End If
End Sub

Private Sub lvwSeating_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not lvwSeating(Index).SelectedItem Is Nothing Then
        '空位不能操作
        If InStr(lvwSeating(Index).SelectedItem.Text, ":") <= 0 Then
            lvwSeating(Index).Drag vbCancel
            Exit Sub
        End If
        Set mSourceItem = lvwSeating(Index).SelectedItem
        Set lvwSeating(Index).DragIcon = img32.ListImages("YouRen").Picture
        lvwSeating(Index).Drag vbBeginDrag
    End If
End Sub

Private Sub lvwSeating_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub picNS1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picSeating(0).Height + Y < 1000 Or picSeating(1).Height - Y < 1000 Then Exit Sub
        picNS1.Top = picNS1.Top + Y
        picSeating(0).Height = picSeating(0).Height + Y
        picSeating(1).Top = picSeating(1).Top + Y
        picSeating(1).Height = picSeating(1).Height - Y
    End If
End Sub

Private Sub picNS2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picSeating(2).Height + Y < 1000 Or picSeating(2).Height - Y < 1000 Then Exit Sub
        picNS2.Top = picNS2.Top + Y
        picSeating(2).Height = picSeating(2).Height + Y
        picSeating(3).Top = picSeating(3).Top + Y
        picSeating(3).Height = picSeating(3).Height - Y
    End If
End Sub

Private Sub picSeating_Resize(Index As Integer)
    On Error Resume Next
    lblTitle(Index).Left = 45
    lvwSeating(Index).Left = 45
    
    lblTitle(Index).Width = picSeating(Index).ScaleWidth - 90
    lvwSeating(Index).Width = picSeating(Index).ScaleWidth - 90
    
    lblTitle(Index).Top = 45
    lvwSeating(Index).Top = lblTitle(Index).Top + lblTitle(Index).Height
    lvwSeating(Index).Height = picSeating(Index).ScaleHeight - lblTitle(Index).Top - lblTitle(Index).Height - 45
    
End Sub

Private Sub picWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picSeating(0).Width + X < 1000 Or picSeating(2).Width - X < 1000 Then Exit Sub
        picWE.Left = picWE.Left + X
        picSeating(0).Width = picSeating(0).Width + X
        picSeating(1).Width = picSeating(1).Width + X
        picNS1.Width = picNS1.Width + X
        
        picSeating(2).Left = picSeating(2).Left + X
        picSeating(2).Width = picSeating(2).Width - X
        picSeating(3).Left = picSeating(3).Left + X
        picSeating(3).Width = picSeating(3).Width - X
        picNS2.Left = picNS2.Left + X
        picNS2.Width = picNS2.Width - X
    End If
End Sub

Public Sub RefreshMain()
    RaiseEvent RequestRefresh
End Sub
