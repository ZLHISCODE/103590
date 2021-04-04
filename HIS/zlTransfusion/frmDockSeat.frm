VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDockSeat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicTab 
      Height          =   5670
      Left            =   330
      ScaleHeight     =   5610
      ScaleWidth      =   11355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1095
      Width           =   11415
      Begin VB.PictureBox picSeats 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5985
         Left            =   180
         ScaleHeight     =   5955
         ScaleWidth      =   10800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Width           =   10830
         Begin zl9Transfusion.udSeat ctlSeat 
            Height          =   2310
            Index           =   0
            Left            =   165
            TabIndex        =   4
            Top             =   300
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   4075
         End
      End
      Begin VB.VScrollBar vsbSeat 
         Height          =   5955
         Left            =   11100
         TabIndex        =   2
         Top             =   240
         Width           =   260
      End
   End
   Begin MSComctlLib.TabStrip TabSeat 
      Height          =   6675
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   11774
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgRpt 
      Left            =   3420
      Top             =   105
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
            Picture         =   "frmDockSeat.frx":0000
            Key             =   "move"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":059A
            Key             =   "已执行"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":0B34
            Key             =   "no"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":10CE
            Key             =   "正在执行"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":1668
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockSeat.frx":1C02
            Key             =   "Calling"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmDockSeat"
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

Private mcurSeatings As Seatings        '坐位记录集

Public lng病人ID As Long                '主窗体传进来， 用于安排座位
Public objPati As cPatient

Private mSourceItem As String           '换座时的源座位
Private mObjItem As String              '换座时的目标座位
Private mcbsMain As CommandBars         '本窗体用的工具栏

Private mSelectKey As String            '当前选择的座位
Private mSelectIndex As Integer         '当前选择的座位索引号

Private mSelectType As String           '当前选择的分类页

Private mblnFormResize As Boolean           '窗体是否在变化，变化时不刷新座位
Private mlngMax As Long
    
Public Sub zlRefresh(ByVal curSeatings As Seatings)
     
    Dim intIndex As Integer
    Dim curSeating As Seating
    Dim strType As String
    Dim strTmp As String
    Set mcurSeatings = Nothing
    mSourceItem = ""
    mObjItem = ""
    Set mcurSeatings = curSeatings
    
    strType = ""
    TabSeat.Tabs.Clear
    
    For Each curSeating In mcurSeatings
        With curSeating
            
            '加分类
            
            strTmp = "" & .分类
            If strTmp = "" Then strTmp = "普通座位"
            
            If strType = "" Then
                TabSeat.Tabs.Add , strTmp, strTmp
                strType = strTmp
            ElseIf InStr("," & strType & ",", "," & strTmp & ",") <= 0 Then
                TabSeat.Tabs.Add , strTmp, strTmp
                strType = strType & "," & strTmp
            End If
             
        End With
    Next
    
    If strType = "" Then
        TabSeat.Tabs.Add , "普通座位", "普通座位"
    End If
    
    mSelectIndex = -1
    mSelectKey = ""
    
    If mSelectType <> "" Then
        On Error Resume Next
        TabSeat.Tabs(mSelectType).Selected = True
        Call TabSeat_Click
    Else

        TabSeat.Tabs("普通座位").Selected = True
        Call TabSeat_Click
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
'        Case conMenu_Edit_Seat_Icon
'            lvwSeating(mintActiveLvw).View = lvwIcon
'        Case conMenu_Edit_Seat_Report
'            lvwSeating(mintActiveLvw).View = lvwReport
'        Case conMenu_Edit_Seat_List
'            lvwSeating(mintActiveLvw).View = lvwList
        Case conMenu_Edit_Seat_Add
            If frmSeatingMana.SeatingMana(0, mcurSeatings, 0, "", Me, mSelectType) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Modify
            Call ModiSeat
        Case conMenu_Edit_Seat_Delete
            StrKey = mSelectKey
            If mcurSeatings.Delete(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Set
            '安排座位
            If lng病人ID <> 0 And Not objPati Is Nothing Then
                StrKey = mSelectKey
                If MsgBox("是否安排[" & objPati.姓名 & "]到[" & StrKey & "]座位", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                    If mcurSeatings.SetSeating(lng病人ID, objPati.挂号单, StrKey) Then
                        RaiseEvent RequestRefresh
                    End If
                End If
            End If
        Case conMenu_Edit_Seat_Clear
        '清除占用的座位
            StrKey = mSelectKey
            If mcurSeatings.Clear(StrKey) Then
                RaiseEvent RequestRefresh
            End If
        Case conMenu_Edit_Seat_Swap
            '换座位
            StrKey = mSelectKey
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
    If mSelectKey <> "" Then
        StrKey = mSelectKey
        If ctlSeat(mSelectIndex).PatiName = "" Then
            If frmSeatingMana.SeatingMana(1, mcurSeatings, 0, StrKey, Me) Then
                RaiseEvent RequestRefresh
            End If
        End If
    End If

End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
 
    
    Select Case Control.ID
        Case conMenu_Edit_Seat_Modify, conMenu_Edit_Seat_Delete
        
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
            
            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName <> "" Then
                    Control.Enabled = False
                End If
            End If
            
        
        Case conMenu_Edit_Seat_Add
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位管理" & ";") > 0
'        Case conMenu_Edit_Seat_Icon
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwIcon
'        Case conMenu_Edit_Seat_List
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwList
'        Case conMenu_Edit_Seat_Report
'            Control.Checked = lvwSeating(mintActiveLvw).View = lvwReport
        Case conMenu_Edit_Seat_Set
        
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0

            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf Not (lng病人ID <> 0 And ctlSeat(mSelectIndex).Stat = 0) Then
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Seat_Clear
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0
            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName = "" Then
                    Control.Enabled = False
                End If
                
            End If
        Case conMenu_Edit_Seat_Swap
            Control.Enabled = InStr(";" & gstrPrivs & ";", ";" & "座位安排" & ";") > 0

            If Control.Enabled Then
                If mSelectIndex = -1 Then
                    Control.Enabled = False
                ElseIf Not (ctlSeat.UBound >= mSelectIndex And ctlSeat.lBound <= mSelectIndex) Then
                    Control.Enabled = False
                ElseIf ctlSeat(mSelectIndex).PatiName = "" Then
                    Control.Enabled = False
                End If
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
        
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_View_Seat, "座位图例")
'        objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'        With objPopup.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_GBed, "普通床位")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_RBed, "占用床位")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_YBed, "维护床位")
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Gseat, "普通座位")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Rseat, "占用座位")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_View_Yseat, "维护座位")
'        End With
    End With
    
'    Set objViewMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'    If objViewMenu Is Nothing Then
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "查看(&V)")
'            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "图标方式(&I)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "列表方式(&L)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "报表方式(&R)")
'            End With
'        End With
'    Else
'        With objViewMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Seat_View, "查看方式(&V)")
'            objPopup.ID = conMenu_Edit_Seat_View: objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Icon, "图标(&I)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_List, "列表(&L)")
'                Set objControl = .Add(xtpControlButton, conMenu_Edit_Seat_Report, "报表(&R)")
'            End With
'        End With
'
'    End If
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
 
    mblnFormResize = True
    
    
    Me.TabSeat.Left = lngLeft
    Me.TabSeat.Top = lngTop
    Me.TabSeat.Width = lngRight - lngLeft
    Me.TabSeat.Height = lngBottom - lngTop
    mblnFormResize = False
    Call picTab_Resize
    
    Me.Refresh
End Sub

Private Sub ctlSeat_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim strObjKey As String
    
    
    If TypeName(Source) = "udSeat" Then
        '放下
         
        
        If Not mSourceItem = "" And Not mObjItem = "" Then
            strObjKey = frmSeatingSwap.ObjectKey(mSourceItem, mcurSeatings, Me, mObjItem)
            If strObjKey <> "" Then
                If mcurSeatings.SwapSeating(mSourceItem, strObjKey) = True Then
                    RaiseEvent RequestRefresh
                End If
            End If
             
        End If
    End If

    If TypeName(Source) = "ReportControl" And lng病人ID <> 0 Then
        '放下
         
        'lvwSeating(Index).MousePointer = ccDefault
        If Not mObjItem = "" And Not objPati Is Nothing Then
            If MsgBox("是否安排[" & objPati.姓名 & "]到[" & mcurSeatings(mObjItem).编号 & "]座位", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
                Call mcurSeatings.SetSeating(lng病人ID, objPati.挂号单, mObjItem)
                RaiseEvent RequestRefresh
            End If
        End If
    End If
    
End Sub

Private Sub ctlSeat_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    Dim objOver As ListItem
    
    Source.DragIcon = imgRpt.ListImages("move").Picture
    If TypeName(Source) = "udSeat" Then
         
        
        If ctlSeat(Index).Stat = 0 Then
                '经过空座位
                
            Set Source.DragIcon = imgRpt.ListImages("yes").Picture
            mObjItem = ctlSeat(Index).Key
        Else
            Set Source.DragIcon = imgRpt.ListImages("no").Picture
            mObjItem = ""
        End If
 
    End If
    If TypeName(Source) = "ReportControl" Then
        
        If ctlSeat(Index).Stat = 0 And lng病人ID <> 0 Then

            Set Source.DragIcon = imgRpt.ListImages("yes").Picture
            mObjItem = ctlSeat(Index).Key
        Else
            Set Source.DragIcon = imgRpt.ListImages("no").Picture
            mObjItem = ""
        End If
    End If
    
    If State = 1 Then
        '离开
        'ctlSeat(Index).GridColor = vbBlue
    Else
        'ctlSeat(Index).GridColor = vbMagenta
    End If
    
End Sub

Private Sub ctlSeat_GotFocus(Index As Integer)
    
    ctlSeat(Index).GridColor = vbRed
    mSelectIndex = Index
    mSelectKey = ctlSeat(Index).Key
    
    RaiseEvent StatusTextUpdate(mSelectKey)
End Sub

Private Sub ctlSeat_LostFocus(Index As Integer)
    ctlSeat(Index).GridColor = vbBlue
End Sub

Private Sub ctlSeat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        '空位不能操作
        If ctlSeat(Index).PatiName = "" Then
            ctlSeat(Index).Drag vbCancel
            Exit Sub
        End If
        mSourceItem = ctlSeat(Index).Key
        
        Set ctlSeat(Index).DragIcon = imgRpt.ListImages("move").Picture
        ctlSeat(Index).Drag vbBeginDrag
    End If
    
End Sub

Private Sub ctlSeat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub Form_Load()
   cbsSub.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
End Sub

Private Sub picSeats_Resize()
    
    Dim iCount As Integer, iRow As Integer
    Dim lngCurLeft As Long, lngCurTop As Long
    Dim lngCarHeight As Long, lngCarWidth As Long '卡片宽，高
    Dim lngSplitWidth  As Long  '间隔
    
    On Error Resume Next
    
    If mblnFormResize Then Exit Sub
    With picSeats
        
        lngCarHeight = ctlSeat(ctlSeat.lBound).Height
        lngCarWidth = ctlSeat(ctlSeat.lBound).Width
        lngSplitWidth = 35
        
        '-- 卡片
        For iCount = ctlSeat.lBound To ctlSeat.UBound
        
            If iCount = ctlSeat.lBound Then
                '第一个卡片
                ctlSeat(iCount).Left = .ScaleTop + lngSplitWidth
                ctlSeat(iCount).Top = .ScaleTop + lngSplitWidth
                
                lngCurLeft = ctlSeat(iCount).Left
                lngCurTop = ctlSeat(iCount).Top
                iRow = 0
            Else
                '之后的卡片根据第一个卡片的位置排列，超过宽度换行
                lngCurLeft = lngCurLeft + lngCarWidth + lngSplitWidth
                If lngCurLeft + lngCarWidth > picSeats.ScaleWidth Then
                    iRow = iRow + 1
                    lngCurLeft = ctlSeat(ctlSeat.lBound).Left
                    lngCurTop = lngCarHeight * iRow + lngSplitWidth * iRow + lngSplitWidth
                    
                End If
                ctlSeat(iCount).Left = lngCurLeft
                ctlSeat(iCount).Top = lngCurTop
            End If
        Next
        mblnFormResize = True
        If ctlSeat(ctlSeat.UBound).Top + lngCarHeight + lngSplitWidth > Me.PicTab.ScaleHeight Then
            picSeats.Height = ctlSeat(ctlSeat.UBound).Top + lngCarHeight + lngSplitWidth
        End If
        mblnFormResize = False
    End With
    
    '--- 滚动条
    vsbSeat.Top = Me.PicTab.ScaleTop
    vsbSeat.Left = Me.PicTab.ScaleWidth - vsbSeat.Width
    vsbSeat.Height = Me.PicTab.ScaleHeight
    vsbSeat.Max = (picSeats.Height - Me.PicTab.ScaleHeight) / Screen.TwipsPerPixelX  '转换为像素为单位
    Call ShowSize(0)
    vsbSeat.Value = 0
    

End Sub


Private Sub ShowSize(Optional lngTop As Single = 0, Optional lngLeft As Single = 0)
    '功能:显示整个坐位
    picSeats.Left = lngLeft
    picSeats.Top = lngTop
    
    Me.Refresh
End Sub

Private Sub picTab_Resize()
    Dim lngCurLeft As Long
    Dim lngCarWidth As Long
    Dim lngCarHeight As Long
    Dim lngSplitWidth As Long
    Dim iRow As Integer
    Dim i As Integer
    On Error Resume Next
    mblnFormResize = True
    Me.PicTab.Move Me.TabSeat.ClientLeft, Me.TabSeat.ClientTop, Me.TabSeat.ClientWidth, Me.TabSeat.ClientHeight
    
    Me.picSeats.Left = Me.PicTab.ScaleLeft
    Me.picSeats.Width = Me.PicTab.ScaleWidth - Me.vsbSeat.Width
    Me.picSeats.Top = Me.PicTab.ScaleTop
    
    lngCarWidth = ctlSeat(ctlSeat.lBound).Width
    lngCarHeight = ctlSeat(ctlSeat.lBound).Height

    lngSplitWidth = 15
    iRow = 0
    For i = 1 To mlngMax
        lngCurLeft = lngSplitWidth + lngCarWidth * i + lngSplitWidth + i
        If lngCurLeft > Me.picSeats.ScaleWidth Then
            iRow = iRow + 1
        End If
    Next
    picSeats.Height = lngSplitWidth + iRow + 1 * lngCarHeight + lngSplitWidth * iRow + 1

    If picSeats.Height < Me.PicTab.ScaleHeight Then
        picSeats.Height = Me.PicTab.ScaleHeight
    End If
    
    mblnFormResize = False
    
    Call picSeats_Resize
End Sub

Private Sub TabSeat_Click()
    Dim i As Integer, iCur As Integer, iMax As Integer
        
    '--装入座位
    mSelectType = TabSeat.SelectedItem.Key
    
    Dim curSeating As Seating
    iMax = 0
    For Each curSeating In mcurSeatings
        If IIf(curSeating.分类 = "", "普通座位", curSeating.分类) = mSelectType Then
            iMax = iMax + 1
        End If
    Next
    
    iCur = ctlSeat.UBound + 1
    If iMax <> iCur Then
        If iMax > iCur Then
            For i = iCur To iMax - 1
                'If ctlSeat.UBound = i - 1 Then Exit For
                    Load ctlSeat(i)
            Next
        Else
            For i = iMax To iCur - 1
                If i = ctlSeat.lBound Then
                    ctlSeat(i).Visible = False
                Else
                    Unload ctlSeat(i)
                End If
                
            Next
        End If
    End If
    
    i = 0
    For Each curSeating In mcurSeatings
        If IIf(curSeating.分类 = "", "普通座位", curSeating.分类) = mSelectType Then
            ctlSeat(i).SeatNo = curSeating.编号
            ctlSeat(i).PatiName = curSeating.姓名 '& " " & curSeating.年齡
            ctlSeat(i).SeatType = curSeating.类型
            ctlSeat(i).Sex = "" & curSeating.性别
            ctlSeat(i).Diagnosis = curSeating.诊断
            ctlSeat(i).Start = curSeating.开始时间 '& " " & curSeating.操作員
            ctlSeat(i).Stat = curSeating.状态
            ctlSeat(i).Key = curSeating.Key
            
            ctlSeat(i).GridColor = vbBlue
            ctlSeat(i).GridWidth = 1
            ctlSeat(i).Visible = True
            i = i + 1
        End If
    Next
    
    mlngMax = iMax
    
    Call picTab_Resize
End Sub

Private Sub vsbSeat_Change()
    Call ShowSize(-vsbSeat.Value * 15#)
End Sub

Private Sub vsbSeat_Scroll()
    Call ShowSize(-vsbSeat.Value * 15#)
End Sub
