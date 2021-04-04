VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCardSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择读卡器"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmCardSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleMode       =   0  'User
   ScaleWidth      =   5969.773
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkDebug 
      Caption         =   "记录日志"
      Height          =   225
      Left            =   150
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   2400
      Width           =   1020
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "启用(&E)"
      Height          =   350
      Index           =   2
      Left            =   2325
      TabIndex        =   4
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "设置(&S)"
      Height          =   350
      Index           =   1
      Left            =   1215
      TabIndex        =   3
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "取消(&C)"
      Height          =   350
      Index           =   3
      Left            =   4620
      TabIndex        =   2
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "确定(&O)"
      Height          =   350
      Index           =   0
      Left            =   3510
      TabIndex        =   1
      Top             =   2340
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2145
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5685
      _cx             =   10028
      _cy             =   3784
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
End
Attribute VB_Name = "frmCardSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCardN0 As String
Private mLasterRow As Integer

Friend Function SelectCard(ByVal colCards As Collection, ByVal intCount As Integer, Optional ByVal FrmMain As Object) As Integer
    Dim objCard As clsCard, lstItem As ListItem
    On Error GoTo errHandle
    If intCount > 0 Then
        '读卡状态，只显示已启用的卡
        Call LoadData(colCards, True)
    ElseIf intCount = 0 Then
        '未设置的情况下读卡，显示所有卡
        Call LoadData(colCards, False)
    ElseIf intCount = -1 Then
        '设置状态下，显示所有的卡
        Call LoadData(colCards, False)
    
    End If
    
    If intCount <> -1 Then
        Me.cmdCard(1).Visible = False
        Me.cmdCard(2).Visible = False
        chkDebug.Visible = False
    End If
    If FrmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, FrmMain
    End If
    SelectCard = Val(mstrCardN0)
    mstrCardN0 = ""
    Exit Function
errHandle:
    Call WritLog("CardSelect.SelectCard", "", err.Description)
End Function

Private Sub cmdCard_Click(Index As Integer)
    With vfgList
    Select Case Index
        Case 0 '-确定
            If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
                mstrCardN0 = Val(.TextMatrix(.Row, .ColIndex("编码")))
                
                Call SaveSetting("ZLSOFT", "公共模块\zlICCard", "调试", chkDebug.value)
                mLasterRow = .Row
                If mLasterRow <= 0 Then mLasterRow = 1
                Call SaveSetting("ZLSOFT", "公共模块\zlICCard", "LastSelect", mLasterRow)
               
                Unload Me
            End If
        Case 1 '-设置
            Dim objCard As clsCardDev
            If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
                Set objCard = CreateObject(.TextMatrix(.Row, .ColIndex("接口")))
                mLasterRow = .Row
                objCard.SetCard
            End If
        Case 2 '-启用,停用
            Call CardEnable
        Case 3 '-退出
            Unload Me
    End Select
    End With
End Sub


Private Sub CardEnable()
    Dim i As Integer
    Dim intCardNo As Integer
    With vfgList
            If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
                intCardNo = Val(.TextMatrix(.Row, .ColIndex("编码")))
                If .TextMatrix(.Row, .ColIndex("启用")) = "√" Then
                    Call SaveSetting("ZLSOFT", "公共模块\zlICCard", Val(.TextMatrix(.Row, .ColIndex("编码"))), 0)
                    .TextMatrix(.Row, .ColIndex("启用")) = "×"
                    cmdCard(2).Caption = "启用(&S)"
                Else
                    Call SaveSetting("ZLSOFT", "公共模块\zlICCard", Val(.TextMatrix(.Row, .ColIndex("编码"))), 1)
                    .TextMatrix(.Row, .ColIndex("启用")) = "√"
                    cmdCard(2).Caption = "停用(&E)"
                End If
            End If

    End With
    For i = 1 To Cards.Count
        If Item(i).编码 = intCardNo Then
            Item(i).启用 = IIf(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("启用")) = "×", False, True)
        End If
    Next
    
End Sub

Private Sub Form_Activate()
    chkDebug.value = GetSetting("ZLSOFT", "公共模块\zlICCard", "调试", 0)
    
    mLasterRow = Val(GetSetting("ZLSOFT", "公共模块\zlICCard", "LastSelect", 1))
    If mLasterRow = 0 Then mLasterRow = 1
    
    Call vfgList_EnterCell
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdCard_Click(0)
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCard_Click(3)
    ElseIf KeyAscii = vbKeySpace Then
        If cmdCard(1).Visible Then
            'Call lvwCardType_DblClick
            Call CardEnable
        End If
    End If
End Sub

Private Sub Form_Load()
    frmTimer.tmrMain.Enabled = False
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If cmdCard(1).Visible Then
            If Val(.TextMatrix(.MouseRow, .ColIndex("编码"))) > 0 Then
                .Select .MouseRow, .ColIndex("启用")
                Call CardEnable
            End If
        Else
            Call cmdCard_Click(0)
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
    cmdCard(1).Enabled = False
    With vfgList
        If .ColIndex("编码") < 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("编码"))) > 0 Then
            If .TextMatrix(.Row, .ColIndex("设置")) = "1" Then
                cmdCard(1).Enabled = True
            End If
            If .TextMatrix(.Row, .ColIndex("启用")) = "√" Then
                cmdCard(2).Caption = "停用(E)"
            Else
                cmdCard(2).Caption = "启用(S)"
            End If
        End If
    End With
End Sub

'---


Public Sub LoadData(ByVal objCards As Collection, ByVal bln启用 As Boolean)
    
    Dim strHead As String, objCard As clsCard
    '1 左对齐 4 居中 7 右对齐
    On Error GoTo errHandle
    If bln启用 Then
        strHead = "编码,600,4;名称,4200,1;启用,1,4;自动,1,4;接口,1,0;险类,1,0;设置,1,0"
    Else
        strHead = "编码,600,4;名称,3600,1;启用,600,4;自动,600,4;接口,1,0;险类,1,0;设置,1,0"
    End If
    
    With vfgList
        .Clear
        Call SetVsFlexGridHead(strHead, vfgList)
        
        For Each objCard In objCards
             If bln启用 = True Then
                 If objCard.启用 Then
                    '仅显示启用
                    .TextMatrix(.Rows - 1, .ColIndex("编码")) = objCard.编码
                    .TextMatrix(.Rows - 1, .ColIndex("名称")) = objCard.名称
                    .TextMatrix(.Rows - 1, .ColIndex("启用")) = IIf(objCard.启用, "√", "×")
                    .TextMatrix(.Rows - 1, .ColIndex("自动")) = IIf(objCard.是否自动读取, "√", "×")
                    .TextMatrix(.Rows - 1, .ColIndex("接口")) = objCard.接口程序名
                    .TextMatrix(.Rows - 1, .ColIndex("险类")) = objCard.险类
                    .TextMatrix(.Rows - 1, .ColIndex("设置")) = objCard.可否设置
                    .Rows = .Rows + 1
                End If
            Else
                '全部
                .TextMatrix(.Rows - 1, .ColIndex("编码")) = objCard.编码
                .TextMatrix(.Rows - 1, .ColIndex("名称")) = objCard.名称
                .TextMatrix(.Rows - 1, .ColIndex("启用")) = IIf(objCard.启用, "√", "×")
                .TextMatrix(.Rows - 1, .ColIndex("自动")) = IIf(objCard.是否自动读取, "√", "×")
                .TextMatrix(.Rows - 1, .ColIndex("接口")) = objCard.接口程序名
                .TextMatrix(.Rows - 1, .ColIndex("险类")) = objCard.险类
                .TextMatrix(.Rows - 1, .ColIndex("设置")) = objCard.可否设置
                .Rows = .Rows + 1
            End If
            
        Next
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '行选择
        .SelectionMode = flexSelectionByRow
        
        If mLasterRow > 0 And mLasterRow < .Rows Then
            .Select mLasterRow, 1
            .TopRow = mLasterRow
        End If
    End With
    Exit Sub
errHandle:
    Call WritLog("CardSelect.LoadData", "", err.Description)
End Sub
Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub
