VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCardSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "消费卡接口设置"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7965
   Icon            =   "frmCardSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7965
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCard 
      Caption         =   "取消(&C)"
      Height          =   350
      Index           =   3
      Left            =   6690
      TabIndex        =   3
      Top             =   855
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "确定(&O)"
      Height          =   350
      Index           =   0
      Left            =   6690
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "启用(&E)"
      Height          =   350
      Index           =   2
      Left            =   6690
      TabIndex        =   1
      Top             =   525
      Width           =   1100
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "记录日志"
      Height          =   225
      Left            =   210
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   5475
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   5175
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   6300
      _cx             =   11112
      _cy             =   9128
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
      Rows            =   6
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCardSelect.frx":030A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
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
   Begin VB.CommandButton cmdCard 
      Caption         =   "设置(&S)"
      Height          =   350
      Index           =   1
      Left            =   6690
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "frmCardSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytCardType As Byte '1-消费卡;2-医疗卡
Private mstrRegSection As String, mblnFirst As Boolean
Private mobjCard As Object
Attribute mobjCard.VB_VarHelpID = -1

Public Function SelectCard(Optional ByVal frmMain As Object, _
    Optional bytCardType As Byte = 0) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择卡
    '入参:frmMain-调用的主窗体
    '       bytCardType-1-消费卡;2-医疗卡;0-不区分消费卡和医疗卡,统一设置
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-06-25 23:20:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    mbytCardType = bytCardType
    mstrRegSection = "公共模块\zlSquareCard\" & IIf(bytCardType = 1, "", "医疗卡\")
    
    If frmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmMain
    End If
   Exit Function
errHandle:
    Call zlWritLog(glngModul, "一卡通部件对象创建", App.ProductName & ".frmCardSelect.SelectCard", "详细的信息为:" & Err.Description, 2)
End Function
Public Function LoadData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载接口数据
    '编制:刘兴洪
    '日期:2009-12-15 16:57:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As clsCard, blnIsCardType As Boolean '是否该对接口
    Dim strPreCode As String, i As Long
    Dim strRegSection  As String
    Dim objYLCards As clsCards
    '59760
    If zlGetCards_YL(objYLCards) = False Then Exit Function
    
    On Error GoTo errHandle

    strPreCode = Val(GetSetting("ZLSOFT", mstrRegSection, "最后选择序号", 0))
    With vfgList
        .Clear 1
        .Rows = 2
        For i = 1 To objYLCards.Count
            'strRegSection = "公共模块\zlSquareCard\" & IIf(objYLCards(i).消费卡, "", "医疗卡\")
            '问题:48005
            If mbytCardType >= 1 And mbytCardType <= 2 Then
                blnIsCardType = IIf(mbytCardType = 1, objYLCards.Item(i).消费卡, (Not objYLCards.Item(i).消费卡 And Not objYLCards.Item(i).系统 And objYLCards.Item(i).自制卡 = False) Or (objYLCards.Item(i).自制卡 = True And objYLCards.Item(i).接口程序名 <> ""))
            Else
                 blnIsCardType = objYLCards.Item(i).消费卡 Or (Not objYLCards.Item(i).消费卡 And Not objYLCards.Item(i).系统 And objYLCards.Item(i).自制卡 = False) Or (objYLCards.Item(i).自制卡 = True And objYLCards.Item(i).接口程序名 <> "")
                 '问题号:54098
                 If (objYLCards.Item(i).名称 Like "*身份证*" Or objYLCards.Item(i).名称 Like "*IC卡*" Or objYLCards.Item(i).名称 = "居民健康卡") And objYLCards.Item(i).系统 = True And objYLCards.Item(i).接口程序名 <> "" Then blnIsCardType = True
            End If
            
            If blnIsCardType Then
                 Set objCard = objYLCards.Item(i)
                 .TextMatrix(.Rows - 1, .ColIndex("编码")) = objCard.接口编码
                 .RowData(.Rows - 1) = IIf(objCard.消费卡, "X", "K") & objCard.接口序号
                 .Cell(flexcpData, .Rows - 1, .ColIndex("编码")) = IIf(objCard.自制卡, 0, 1)
                 .TextMatrix(.Rows - 1, .ColIndex("名称")) = objCard.名称
                 .TextMatrix(.Rows - 1, .ColIndex("启用")) = IIf(objCard.启用, "√", "×")
                 .TextMatrix(.Rows - 1, .ColIndex("自动")) = IIf(objCard.是否自动读取, "√", "×")
                 .TextMatrix(.Rows - 1, .ColIndex("部件")) = objCard.接口程序名
                If strPreCode = objCard.接口编码 Then .Row = .Rows - 1
                .Rows = .Rows + 1
            End If
        Next
        If .Rows > 0 Then .Rows = .Rows - 1
        If .RowIsVisible(.Row) Then .TopRow = .Row
    End With
    LoadData = True
    Exit Function
errHandle:
    Call zlWritLog(glngModul, "一卡通部件对象创建", App.ProductName & ".frmCardSelect.LoadData", "详细的信息为:" & Err.Description, 2)
End Function
Private Sub chkDebug_Click()
    Call SaveSetting("ZLSOFT", mstrRegSection, "调试", chkDebug.value)
End Sub

Private Sub cmdCard_Click(Index As Integer)
    Dim strPreCode As String, strKey As String
    Dim bln消费卡 As Boolean, lngCardTypeID As Long
    
    With vfgList
    Select Case Index
        Case 0 '-确定
            If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
                Call SaveSetting("ZLSOFT", mstrRegSection, "调试", chkDebug.value)
                strPreCode = Trim(.TextMatrix(.Row, .ColIndex("编码")))
                Call SaveSetting("ZLSOFT", mstrRegSection, "LastSelect", strPreCode)
                Unload Me
            End If
        Case 1 '-设置
            strKey = .RowData(.Row)
            If strKey = "" Then Exit Sub
            
            bln消费卡 = Left(strKey, 1) = "X"
            lngCardTypeID = Val(Mid(strKey, 2))
            
            If lngCardTypeID > 0 Then
                Set mobjCard = zlGetComponentObject(lngCardTypeID, bln消费卡)
                If mobjCard Is Nothing Then
                    MsgBox "注意:" & vbCrLf & .TextMatrix(.Row, .ColIndex("名称")) & " 未启用或未找到指定部件，请检查！", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
                If mobjCard.zlCardDevSet(Me, lngCardTypeID) Then Exit Sub
                Set gObjYLCardObjs = Nothing: Set gObjYLCards = Nothing '88185,将这两个全局变量置为nothing，以便其它地方更新医疗卡缓存数据
            End If
        Case 2 '-启用,停用
            Call setCardStopOrResume
        Case 3 '-退出
            Unload Me
    End Select
    End With
End Sub
Private Sub setCardStopOrResume()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置卡的启用和停用
    '编制:刘兴洪
    '日期:2009-12-15 17:29:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strKey As String, lngCardTypeID As Long, bln消费卡 As Boolean
    Dim strRegSection  As String
'    Dim objYLCards As clsCards
'    Dim objYlCardObjs As clsCardObjects
'    '59760
'    If zlGetCards_YL(objYLCards) = False Then Exit Sub
'    If zlGetYLCardObjs(objYlCardObjs) = False Then Exit Sub
    
    With vfgList
        strKey = .RowData(.Row)
        If strKey = "" Then Exit Sub
        bln消费卡 = Left(strKey, 1) = "X"
        strRegSection = "公共模块\zlSquareCard\" & IIf(bln消费卡, "", "医疗卡\")
            
        lngCardTypeID = Val(Mid(strKey, 2))
        If InStr(.RowData(.Row), "K") = 1 And .Cell(flexcpData, .Row, .ColIndex("编码")) = 0 Then Exit Sub
        If lngCardTypeID > 0 Then
            If .TextMatrix(.Row, .ColIndex("启用")) = "√" Then
                Call SaveSetting("ZLSOFT", strRegSection & Trim(.TextMatrix(.Row, .ColIndex("编码"))), "启用", 0)
                .TextMatrix(.Row, .ColIndex("启用")) = "×"
                cmdCard(2).Caption = "启用(&E)"
            Else
                Call SaveSetting("ZLSOFT", strRegSection & Trim(.TextMatrix(.Row, .ColIndex("编码"))), "启用", 1)
                .TextMatrix(.Row, .ColIndex("启用")) = "√"
                cmdCard(2).Caption = "停用(&E)"
            End If
        Else
            Exit Sub
        End If
        If lngCardTypeID <> 0 Then
            Set gObjYLCardObjs = Nothing: Set gObjYLCards = Nothing '88185,将这两个全局变量置为nothing，以便其它地方更新医疗卡缓存数据
        End If
    End With
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadData = False Then Unload Me: Exit Sub
    chkDebug.value = GetSetting("ZLSOFT", mstrRegSection, "调试", 0)
    Call vfgList_GotFocus: Call vfgList_EnterCell
    Call vfgList_AfterRowColChange(0, 0, 1, 0)
     If vfgList.Enabled And vfgList.Visible Then vfgList.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdCard_Click(0)
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCard_Click(3)
    ElseIf KeyAscii = vbKeySpace Then
        If cmdCard(1).Visible Then
            Call setCardStopOrResume
        End If
    End If
End Sub

Private Sub Form_Load()
    frmCardBrush.tmrMain.Enabled = False
    mblnFirst = True
End Sub
 Private Sub vfgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vfgList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    With vfgList
        If .Rows <= 1 Then Exit Sub
        cmdCard(1).Enabled = Val(.Cell(flexcpData, NewRow, .ColIndex("编码"))) = 1
        '问题48005
        '83399:李南春,2015/7/21,消费卡也能设置部件
        cmdCard(1).Enabled = True
        cmdCard(2).Enabled = Val(.Cell(flexcpData, NewRow, .ColIndex("编码"))) = 1
    End With
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If cmdCard(1).Visible Then
            If Val(.TextMatrix(.MouseRow, .ColIndex("编码"))) > 0 Then
                .Select .MouseRow, .ColIndex("启用")
                Call setCardStopOrResume
            End If
        Else
            Call cmdCard_Click(2)
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
    With vfgList
        If .ColIndex("编码") < 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
            If .TextMatrix(.Row, .ColIndex("启用")) = "√" Then
                cmdCard(2).Caption = "停用(&E)"
            Else
                cmdCard(2).Caption = "启用(&E)"
            End If
        End If
    End With
End Sub

 
Private Sub vfgList_GotFocus()
  zl_VsGridGotFocus vfgList, gSysColor.lngGridColorSel
End Sub

Private Sub vfgList_LostFocus()
  zl_VsGridLostFocus vfgList, gSysColor.lngGridColorLost
End Sub
