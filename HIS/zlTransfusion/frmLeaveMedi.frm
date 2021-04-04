VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLeaveMedi 
   BorderStyle     =   0  'None
   Caption         =   "药品寄存管理"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   Icon            =   "frmLeaveMedi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTow 
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   345
      ScaleHeight     =   2580
      ScaleWidth      =   7005
      TabIndex        =   5
      Top             =   660
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid vsListUsed 
         Height          =   3255
         Left            =   150
         TabIndex        =   6
         Top             =   165
         Width           =   7800
         _cx             =   13758
         _cy             =   5741
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLeaveMedi.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picOne 
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   960
      ScaleHeight     =   2580
      ScaleWidth      =   7005
      TabIndex        =   3
      Top             =   3300
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   3255
         Left            =   -435
         TabIndex        =   4
         Top             =   195
         Width           =   7800
         _cx             =   13758
         _cy             =   5741
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLeaveMedi.frx":68ED
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   3045
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   5371
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2145
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeaveMedi.frx":6988
            Key             =   "mx"
            Object.Tag             =   "mx"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeaveMedi.frx":D1EA
            Key             =   "use"
            Object.Tag             =   "use"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picNS 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   1875
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   2955
      Width           =   6615
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMaster 
      Height          =   2445
      Left            =   270
      TabIndex        =   0
      Top             =   450
      Width           =   7800
      _cx             =   13758
      _cy             =   4313
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmLeaveMedi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const subMenu_Add = 101 '增加
Private Const subMenu_Modify = 102 '修改
Private Const subMenu_Delete = 103 '删除
Private Const subMenu_Post = 104 '使用登记
Private Const subMenu_Repertory = 105 '库存查询
Private Const subMenu_AccountBook = 106 '库存台帐

Private mlng病人ID As Long
Private mlng科室ID As Long
Private mdateBeging As Date
Private mdateEnd As Date
Private mobjMasters As MediMasters
Private mstr科室 As String
Private mstr姓名 As String
Private mstr性别 As String
Private mstr年龄 As String
Private mstr挂号单 As String

Private mblnEditList As Boolean '编辑状态
Private mcbsMain As CommandBars

Public Property Let 年龄(ByVal vData As String)
    mstr年龄 = vData
End Property

Public Property Let 性别(ByVal vData As String)
    mstr性别 = vData
End Property

Public Property Let 科室(ByVal vData As String)
    mstr科室 = vData
End Property

Public Property Let 姓名(ByVal vData As String)
    mstr姓名 = vData
End Property

Public Property Let 挂号单(ByVal vData As String)
    mstr挂号单 = vData
End Property

Public Property Let 科室ID(ByVal vData As Long)
    mlng科室ID = vData
End Property

Public Property Let 病人ID(ByVal vData As Long)
    mlng病人ID = vData
End Property

Public Property Let dateBeging(ByVal vData As Date)
    mdateBeging = vData
End Property

Public Property Let DateEnd(ByVal vData As Date)
    mdateEnd = vData
End Property

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    'Call dkpSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'    With vsMaster
'        .Left = 0
'        .Top = lngTop
'        .Width = lngRight
'    End With
''    With picNS
''        .Left = 0
''        .Top = vsMaster.Top + vsMaster.Height
''        .Width = lngRight
''    End With
''    With tabRecord
''        .Left = 0
''        .Top = picNS.Top + picNS.Height
''    End With
'    With vsList
'        .Left = 0
'        .Top = tabRecord.Top + tabRecord.Height + 15
'        .Width = lngRight
'        .Height = lngBottom - lngTop - vsMaster.Height - picNS.Height - tabRecord.Height - 60
'    End With

    vsList.Top = lngTop
    vsList.Left = lngLeft
    vsList.Height = (lngBottom - lngTop) / 2
    vsList.Width = lngRight - lngLeft

    vsMaster.Top = lngTop
    vsMaster.Left = lngLeft
    vsMaster.Height = (lngBottom - lngTop) / 2
    vsMaster.Width = lngRight - lngLeft

    vsListUsed.Top = lngTop
    vsListUsed.Left = lngLeft
    vsListUsed.Height = (lngBottom - lngTop) / 2
    vsListUsed.Width = lngRight - lngLeft
End Sub

Private Sub dkpSub_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    'If Action = PaneActionFloating Then
'    dkpSub.PanelPaintManager.Position = xtpTabPositionBottom
    'End If
End Sub

Private Sub picNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsMaster.Height + Y < 1000 Or tbcSub.Height - Y < 800 Then Exit Sub
        picNS.Top = picNS.Top + Y
        vsMaster.Height = vsMaster.Height + Y
        
    
        tbcSub.Top = tbcSub.Top + Y
        tbcSub.Height = tbcSub.Height - Y
         
    End If
End Sub
Private Sub Form_Load()
    Call Fill_Master
End Sub

Public Sub ShowLeaveMedi(ByVal lng科室ID As Long, Optional ByVal lng病人ID As Long)
    mlng病人ID = lng病人ID
    mlng科室ID = lng科室ID
End Sub

Public Sub Fill_Master()
    Dim strHead As String, objMaster As MediMaster
    strHead = "NO,1000,1;姓名,900,1;性别,450,4;年龄,450,1;操作员,900,1;登记时间,1500,1;合计,900,7;使用情况,900,1;摘要,1000,1;作废时间,0,1;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsMaster)
    Set mobjMasters = New MediMasters
    Call mobjMasters.GetMediMasters(mdateBeging, mdateEnd, mlng科室ID, mlng病人ID)
    
    If mobjMasters Is Nothing Then Exit Sub
    
    For Each objMaster In mobjMasters
        With vsMaster
            .TextMatrix(.Rows - 1, .ColIndex("NO")) = objMaster.NO
            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = objMaster.姓名
            .TextMatrix(.Rows - 1, .ColIndex("性别")) = objMaster.性别
            .TextMatrix(.Rows - 1, .ColIndex("年龄")) = objMaster.年龄
            .TextMatrix(.Rows - 1, .ColIndex("操作员")) = objMaster.操作员
            .TextMatrix(.Rows - 1, .ColIndex("登记时间")) = Format(objMaster.登记时间, "yy-MM-dd hh:mm")
            .TextMatrix(.Rows - 1, .ColIndex("合计")) = Format(objMaster.合计, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("使用情况")) = objMaster.使用情况
            .TextMatrix(.Rows - 1, .ColIndex("摘要")) = objMaster.摘要
            .TextMatrix(.Rows - 1, .ColIndex("作废时间")) = IIf(objMaster.作废时间 = CDate(0), "", Format(objMaster.作废时间, "yy-MM-dd hh:mm"))
            .TextMatrix(.Rows - 1, .ColIndex("Key")) = objMaster.NO & "_" & IIf(objMaster.作废时间 = CDate(0), "0", Format(objMaster.作废时间, "yyMMddhhmmss"))
            .Rows = .Rows + 1
        End With
    Next
    If vsMaster.Rows > 2 Then
        Call vsMaster.RemoveItem(vsMaster.Rows - 1)
    End If
    vsMaster.Editable = flexEDNone
    vsMaster_RowColChange
    'vsMaster.Select vsMaster.Rows - (Rows - 1), 1
End Sub

Private Sub initVsList()
    Dim strHead As String
    
    strHead = "药品来源,650,1;药品名称与编码,2800,1;规格,1800,1;用途,550,1;数量,750,7;已用数量,600,7;计算单位,450,4;单价,750,7;金额,1000,7;原使用数量,0,7;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsList)
    
    strHead = "使用时间,1200,1;药品来源,650,1;药品名称与编码,2800,1;规格,1800,1;用途,550,1;使用数量,600,7;计算单位,450,4;单价,750,7;金额,1000,7;填制人,650,1;摘要,1000,1;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsListUsed)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsMaster
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Width = Me.ScaleWidth
    End With
    With picNS
        .Left = Me.ScaleLeft
        .Top = vsMaster.Top + vsMaster.Height
        .Width = Me.ScaleWidth
    End With
    With tbcSub
        .Left = Me.ScaleLeft
        .Top = picNS.Top + picNS.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsMaster.Height - picNS.Height - 60
    End With
    With vsList
        
    End With
End Sub

Private Sub picOne_Resize()
    With vsList
        .Top = picOne.Top
        .Width = picOne.Width
        .Height = picOne.Height
        .Left = picOne.Left
    End With
End Sub



Private Sub picTow_Resize()
    With vsListUsed
        .Top = picTow.Top
        .Width = picTow.Width
        .Height = picTow.Height
        .Left = picTow.Left
    End With
End Sub



Private Sub vsMaster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsMaster_RowColChange()
    
    With vsMaster
    If .ColIndex("Key") > 0 Then
        Call Fill_List(.TextMatrix(.Row, .ColIndex("Key")))
    Else
        Call Fill_List("")
    End If
    End With
End Sub

Private Sub Fill_List(ByVal strNo As String)

    Dim strHead As String, objBIll As MediBill, str来源 As String, str用途 As String
    Dim objMaster As MediMaster, i As Integer
    Call initVsList

        If strNo = "" Or strNo = "NO" Then Exit Sub
        If mobjMasters Is Nothing Then Exit Sub
        Set objMaster = mobjMasters.Item(strNo)
        With vsList
            For i = 1 To objMaster.BillCount
                Set objBIll = objMaster.BillItem(i)
                If objBIll.入出系数 = 1 Then
                    Select Case objBIll.执行分类
                    Case 1
                        str用途 = "输液"
                    Case 2
                        str用途 = "注射"
                    Case 3
                        str用途 = "皮试"
                    Case Else
                        str用途 = "治疗"
                    End Select
                    If objBIll.药品ID = 0 And objBIll.医嘱ID = 0 Then
                        str来源 = "目录外"
                    ElseIf objBIll.药品ID <> 0 And objBIll.医嘱ID = 0 Then
                        str来源 = "目录内"
                    ElseIf objBIll.药品ID <> 0 And objBIll.医嘱ID <> 0 Then
                        str来源 = "医嘱"
                    Else
                        str来源 = "错误"
                    End If
                    
                    .TextMatrix(.Rows - 1, .ColIndex("药品来源")) = str来源
                    .TextMatrix(.Rows - 1, .ColIndex("药品名称与编码")) = objBIll.药品名称
                    .TextMatrix(.Rows - 1, .ColIndex("规格")) = objBIll.规格
                    .TextMatrix(.Rows - 1, .ColIndex("用途")) = str用途
                    .TextMatrix(.Rows - 1, .ColIndex("数量")) = objBIll.数量
                    .TextMatrix(.Rows - 1, .ColIndex("已用数量")) = objBIll.已用数量
                    .TextMatrix(.Rows - 1, .ColIndex("计算单位")) = objBIll.计算单位
                    .TextMatrix(.Rows - 1, .ColIndex("单价")) = Format(objBIll.单价, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(objBIll.金额, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("原使用数量")) = objBIll.已用数量
                    .TextMatrix(.Rows - 1, .ColIndex("Key")) = objBIll.序号 & "_" & objBIll.入出系数 & "_" & Format(objBIll.登记时间, "yyMMddhhmmss")
                    .Rows = .Rows + 1
                End If
            Next
            If .Rows > 2 Then
                .RemoveItem (.Rows - 1)
            End If
        End With
        

        '------使用记录
        
        If strNo = "" Or strNo = "NO" Then Exit Sub
        If mobjMasters Is Nothing Then Exit Sub
        Set objMaster = mobjMasters.Item(strNo)
        With vsListUsed
            For i = 1 To objMaster.BillCount
                Set objBIll = objMaster.BillItem(i)
                If objBIll.入出系数 = -1 Then
                    Select Case objBIll.执行分类
                    Case 1
                        str用途 = "输液"
                    Case 2
                        str用途 = "注射"
                    Case 3
                        str用途 = "皮试"
                    Case Else
                        str用途 = "治疗"
                    End Select
                    If objBIll.药品ID = 0 And objBIll.医嘱ID = 0 Then
                        str来源 = "目录外"
                    ElseIf objBIll.药品ID <> 0 And objBIll.医嘱ID = 0 Then
                        str来源 = "目录内"
                    ElseIf objBIll.药品ID <> 0 And objBIll.医嘱ID <> 0 Then
                        str来源 = "医嘱"
                    Else
                        str来源 = "错误"
                    End If
                    .TextMatrix(.Rows - 1, .ColIndex("使用时间")) = Format(objBIll.登记时间, "yy-MM-dd hh:mm")
                    .TextMatrix(.Rows - 1, .ColIndex("药品来源")) = str来源
                    .TextMatrix(.Rows - 1, .ColIndex("药品名称与编码")) = objBIll.药品名称
                    .TextMatrix(.Rows - 1, .ColIndex("规格")) = objBIll.规格
                    .TextMatrix(.Rows - 1, .ColIndex("用途")) = str用途
                    .TextMatrix(.Rows - 1, .ColIndex("使用数量")) = objBIll.数量
                    .TextMatrix(.Rows - 1, .ColIndex("计算单位")) = objBIll.计算单位
                    .TextMatrix(.Rows - 1, .ColIndex("单价")) = Format(objBIll.单价, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = Format(objBIll.金额, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("填制人")) = objBIll.填制人
                    .TextMatrix(.Rows - 1, .ColIndex("摘要")) = objBIll.使用摘要
                    .TextMatrix(.Rows - 1, .ColIndex("Key")) = objBIll.序号 & "_" & objBIll.入出系数 & "_" & Format(objBIll.登记时间, "yyMMddhhmmss")
                    .Rows = .Rows + 1
                End If
            Next
            If .Rows > 2 Then
                .RemoveItem (.Rows - 1)
            End If
        End With
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    '主窗体要求初始化主窗体上的菜单
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '暂存药品的菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    Set mcbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "寄存药品(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Add, "增加(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Delete, "删除(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "使用登记(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "撤消登记(&U)")
    End With
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
'    If objMenu Is Nothing Then
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
'        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
'    End If
'    With objMenu.CommandBar.Controls
'        '子项放在最前面,反序加入
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Repertory, "库存查询(&R)", 1)
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_AccountBook, "库存台帐(&A)", 2)
'    End With
    
    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Add, "增加", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Modify, "修改", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Delete, "删除", objControl.Index + 1)
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "使用登记", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "撤消登记", objControl.Index + 1)
    End With
'    cbsSub.ActiveMenuBar.Visible = False
    
    
    '-- vsLIst
'    Dim objPane As Pane, objPaneBase As Pane
'
'    Me.dkpSub.Options.UseSplitterTracker = False '实时拖动
'    Me.dkpSub.Options.ThemedFloatingFrames = True
'    Me.dkpSub.VisualTheme = ThemeOffice2003
'    Me.dkpSub.Options.HideClient = True
'
'    If dkpSub.FindPane(1) Is Nothing Then
'        Set objPaneBase = Me.dkpSub.CreatePane(1, 700, 200, DockTopOf, Nothing)
'        objPaneBase.Title = "暂存药品"
'        objPaneBase.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If
'
'    If dkpSub.FindPane(2) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(2, 700, 200, DockBottomOf, objPaneBase)
'        objPaneBase.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If
    
    
    With tbcSub
        If tbcSub.ItemCount <= 0 Then
            With .PaintManager
                .Appearance = xtpTabAppearanceExcel
                .ClientFrame = xtpTabFrameSingleLine
                
                .Position = xtpTabPositionBottom '选项卡在底部
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
    
            .InsertItem(0, "暂存明细", picOne.hwnd, 0).Tag = "暂存明细"
            .InsertItem(1, "使用记录", picTow.hwnd, 0).Tag = "使用记录"
            .Item(0).Selected = True
        End If
    End With
'    If dkpSub.FindPane(2) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(2, 700, 200, DockBottomOf, objPaneBase)
'        objPane.Title = "暂存明细"
'        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'
'        vsList.Visible = False
'    End If
'
'    If dkpSub.FindPane(3) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(3, 700, 200, DockBottomOf, objPaneBase)
'        objPane.Title = "使用明细"
'        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If


   '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    'Me.dkpSub.SetCommandBars cbsSub
    
End Sub

Public Sub zlRefresh()
    Call Fill_Master
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '
    Select Case Control.ID
        Case conMenu_Edit_Leave_Add
        '增加
            Control.Enabled = mlng病人ID <> 0 And (Not mblnEditList) And InStr(gstrPrivs, ";" & "药品寄存" & ";") <> 0
        Case conMenu_Edit_Leave_Modify
        '修改
             
            If vsMaster.Row > 0 And vsMaster.ColIndex("NO") < vsMaster.Cols Then
                Control.Enabled = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("NO")) <> "" And (Not mblnEditList) _
                                  And vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("使用情况")) = "未用" _
                                  And InStr(gstrPrivs, ";" & "药品寄存" & ";") <> 0
            Else
                Control.Enabled = False
            End If
            If vsList.Editable = flexEDKbd Then Control.Enabled = False
        Case conMenu_Edit_Leave_Delete
        '删除
            
            If vsMaster.Row > 0 And vsMaster.ColIndex("NO") < vsMaster.Cols Then
                Control.Enabled = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("NO")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "药品寄存" & ";") <> 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Leave_Post
        '登记
            If vsList.Row > 0 And vsList.ColIndex("Key") < vsList.Cols Then
                Control.Enabled = vsList.TextMatrix(vsList.Row, vsList.ColIndex("Key")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "药品寄存" & ";") <> 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Leave_UndoPost
        '撤消登记
            
            If tbcSub.Item(1).Selected Then
                If vsListUsed.Row > 0 And vsListUsed.ColIndex("Key") < vsListUsed.Cols Then
                    Control.Enabled = vsListUsed.TextMatrix(vsListUsed.Row, vsListUsed.ColIndex("Key")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "药品寄存" & ";") <> 0
                End If
            Else
                Control.Enabled = False
            End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl, ByVal frmMain As frmTransfusion)
    '主窗体调用本窗体的执行功能
    Dim lngRow As Long, strMasterKey As String, lngMastRow As Long
    Select Case Control.ID
        Case conMenu_Edit_Leave_Add '增加
            If mlng病人ID <> 0 Then
                Set frmLeaveMediMana.pMediMaster = New MediMaster
                frmLeaveMediMana.pintType = 1
                With frmLeaveMediMana.pMediMaster
                    .科室名称 = mstr科室
                    .姓名 = mstr姓名
                    .性别 = mstr性别
                    .年龄 = mstr年龄
                    .病人ID = mlng病人ID
                    .科室ID = mlng科室ID
                    .挂号单 = mstr挂号单
                    .操作员 = UserInfo.姓名
                    .登记时间 = zlDatabase.Currentdate
                End With
                frmLeaveMediMana.Show vbModal, Me
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                Call frmMain.刷新(lngRow)
                vsMaster.Row = lngMastRow
                vsMaster_RowColChange
            End If
        Case conMenu_Edit_Leave_Modify '修改
            With vsMaster
            If .TextMatrix(.Row, .ColIndex("Key")) <> "" Then
                frmLeaveMediMana.pintType = 2
                Set frmLeaveMediMana.pMediMaster = mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key")))
                With frmLeaveMediMana.pMediMaster
                    .挂号单 = mstr挂号单
                    .病人ID = mlng病人ID
                    .科室ID = mlng科室ID
                    .科室名称 = mstr科室
                    .操作员 = UserInfo.姓名
                End With
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                frmLeaveMediMana.Show vbModal, Me
                Call frmMain.刷新(lngRow)
                If vsMaster.Rows > lngMastRow Then
                    vsMaster.Row = lngMastRow
                End If
                vsMaster_RowColChange

            End If
            End With
        Case conMenu_Edit_Leave_Delete '删除
            With vsMaster
                lngRow = GetMainCurRowIndex(frmMain)
                Call mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key"))).DeleteBill(0)
                Call frmMain.刷新(lngRow)
            End With
        Case conMenu_Edit_Leave_Post '使用登记
            With vsMaster
            If .TextMatrix(.Row, .ColIndex("Key")) <> "" Then
                frmLeaveMediMana.pintType = 3
                Set frmLeaveMediMana.pMediMaster = mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key")))
                With frmLeaveMediMana.pMediMaster
                    .挂号单 = mstr挂号单
                    .病人ID = mlng病人ID
                    .科室ID = mlng科室ID
                    .科室名称 = mstr科室
                    .操作员 = UserInfo.姓名
                End With
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                frmLeaveMediMana.Show vbModal, Me
                Call frmMain.刷新(lngRow)
                If vsMaster.Rows > lngMastRow Then
                    vsMaster.Row = lngMastRow
                End If
                vsMaster_RowColChange
            End If
            End With
        Case conMenu_Edit_Leave_UndoPost '撤消使用登记
            strMasterKey = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("Key"))
            If strMasterKey <> "" Then
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                With vsListUsed
                    If .TextMatrix(.Row, .ColIndex("药品来源")) <> "医嘱" Then
                        If MsgBox("请确认，是否撤消【" & .TextMatrix(.Row, .ColIndex("药品名称与编码")) & "】，数量为 【" & .TextMatrix(.Row, .ColIndex("使用数量")) & "】的使用记录？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            Call mobjMasters.Item(strMasterKey).UndoUse(.TextMatrix(.Row, .ColIndex("Key")))
                            Call frmMain.刷新(lngRow)
                            If vsMaster.Rows > lngMastRow Then
                                vsMaster.Row = lngMastRow
                            End If
                            vsMaster_RowColChange
                        End If
                    Else
                        MsgBox "医嘱记录不能在此撤消使用！", vbInformation, gstrSysName
                    End If
                End With
            End If
        Case conMenu_Edit_Leave_Repertory '库存查询
        Case conMenu_Edit_Leave_AccountBook '台帐
        
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '主窗框体调用本窗体的弹出菜单
End Sub

Private Function GetMainCurRowIndex(ByVal frmMain As frmTransfusion) As Long
    '取主窗体的当前活动面面的RPT控件中的选中行的index
    Dim objRpt As ReportControl
    
    If frmMain.tbcList.Selected.Tag = "未接单" Then
        Set objRpt = frmMain.rptQueue0
    ElseIf frmMain.tbcList.Selected.Tag = "待配液" Then
        Set objRpt = frmMain.rptQueue1
    ElseIf frmMain.tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = frmMain.rptQueue5
    ElseIf frmMain.tbcList.Selected.Tag = "待执行" Then
        Set objRpt = frmMain.rptQueue6
    ElseIf frmMain.tbcList.Selected.Tag = "执行中" Then
        Set objRpt = frmMain.rptQueue7
    ElseIf frmMain.tbcList.Selected.Tag = "已结束" Then
        Set objRpt = frmMain.rptPati
    End If
    GetMainCurRowIndex = objRpt.SelectedRows(0).Index
End Function

