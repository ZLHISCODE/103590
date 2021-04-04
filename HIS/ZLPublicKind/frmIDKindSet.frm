VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIDKindSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "类别属性设置"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7875
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboDefaultCard 
      Height          =   300
      Left            =   1380
      TabIndex        =   22
      Text            =   "cboDefaultCard"
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "恢复缺省(&D)"
      Height          =   405
      Left            =   6465
      TabIndex        =   20
      Top             =   4905
      Width           =   1320
   End
   Begin VB.Frame fraSplitRight 
      Caption         =   "Frame2"
      Height          =   5910
      Left            =   6375
      TabIndex        =   19
      Top             =   -165
      Width           =   30
   End
   Begin VB.ComboBox cboFastkey 
      Height          =   300
      Index           =   2
      Left            =   2310
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1455
      Width           =   1140
   End
   Begin VB.ComboBox cboFun 
      Height          =   300
      Index           =   2
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1455
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "卡类别滚动快键设置"
      Height          =   990
      Left            =   90
      TabIndex        =   3
      Top             =   240
      Width           =   6240
      Begin VB.ComboBox cboFastkey 
         Height          =   300
         Index           =   1
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   960
      End
      Begin VB.ComboBox cboFun 
         Height          =   300
         Index           =   1
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
      Begin VB.ComboBox cboFastkey 
         Height          =   300
         Index           =   0
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   345
         Width           =   1155
      End
      Begin VB.ComboBox cboFun 
         Height          =   300
         Index           =   0
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lblFun 
         Caption         =   "快键:Ctrl+F4"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   3750
         TabIndex        =   13
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   150
         Index           =   1
         Left            =   5145
         TabIndex        =   11
         Top             =   450
         Width           =   210
      End
      Begin VB.Label lblEdit 
         Caption         =   "向后滚动"
         Height          =   210
         Index           =   1
         Left            =   3390
         TabIndex        =   9
         Top             =   405
         Width           =   780
      End
      Begin VB.Label lblFun 
         Caption         =   "快键:Ctrl+F4"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   8
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   150
         Index           =   0
         Left            =   1980
         TabIndex        =   6
         Top             =   435
         Width           =   210
      End
      Begin VB.Label lblEdit 
         Caption         =   "向前滚动"
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6645
      TabIndex        =   1
      Top             =   315
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6645
      TabIndex        =   0
      Top             =   750
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3465
      Left            =   120
      TabIndex        =   2
      Top             =   1905
      Width           =   6195
      _cx             =   10927
      _cy             =   6112
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIDKindSet.frx":0000
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
   Begin VB.Label lblDefaultCard 
      Caption         =   "缺省读卡类别"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   5460
      Width           =   1155
   End
   Begin VB.Label lblFun 
      Caption         =   "快键:Ctrl+F4"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   3810
      TabIndex        =   18
      Top             =   1515
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   150
      Index           =   2
      Left            =   2070
      TabIndex        =   16
      Top             =   1545
      Width           =   210
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "读卡快键"
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   14
      Top             =   1500
      Width           =   720
   End
End
Attribute VB_Name = "frmIDKindSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mobjCards As Cards
Private mblnOk As Boolean
Private mstrNotContainFastKey As String
Private mvarNotKey As Variant
Private mstrMustSelectItems As String
Private mRegType As gRegType
Private mblnConn As Boolean
Private mstrPrivs As String
Private mcnOracle As ADODB.Connection

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的有效性
    '返回: 数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-23 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, i As Long
    On Error GoTo errHandle
    
    With vsGrid
        For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("启用"))) <> 0 And .RowData(i) <> "" Then
                blnFind = True: Exit For
             End If
        Next
    End With
    If Not blnFind Then
        MsgBox "必须至少启用一个类别,请检查", vbInformation + vbOKOnly, gstrSysName
        If vsGrid.Enabled And vsGrid.Visible Then vsGrid.SetFocus
        Exit Function
    End If
    For i = 0 To cboFun.UBound
        If Trim(cboFun(i).Text) <> "" And Trim(cboFastkey(i).Text) = "" Then
            MsgBox lblEdit(i).Caption & "未设置快键,请检查", vbInformation + vbOKOnly, gstrSysName
            If cboFastkey(i).Enabled And cboFastkey(i).Visible Then cboFastkey(i).SetFocus
            Exit Function
        End If
        
    Next
    '78768
    With vsGrid
        For i = 1 To .Rows - 1
            If cboDefaultCard.Text = .TextMatrix(i, .ColIndex("名称")) And .TextMatrix(i, .ColIndex("启用")) <> 1 Then
                MsgBox "选用的缺省读卡类别已被停用，请重新设置。", vbInformation + vbOKOnly, gstrSysName
                '刷新卡列表
                Call InitDefaultCard
            isValied = False: Exit Function
        End If
    Next
End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function


Private Sub SaveParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存参数设置
    '编制:刘兴洪
    '日期:2018-12-05 14:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, strValue As String
    Dim objPubOneCard As Object
    
    On Error GoTo errHandle
    
    Call zlGetPubOneCard(mcnOracle, objPubOneCard)
    strValue = Trim(cboFun(0).Text)
    If mblnConn Then Call objPubOneCard.SetPara("向前滚动-功能键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "向前滚动-功能键", strValue)
    strValue = Trim(cboFastkey(0).Text)
    If mblnConn Then Call objPubOneCard.SetPara("向前滚动-快键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "向前滚动-快键", strValue)
    
    strValue = Trim(cboFun(1).Text)
    If mblnConn Then Call objPubOneCard.SetPara("向后滚动-功能键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "向后滚动-功能键", strValue)
     strValue = Trim(cboFastkey(1).Text)
     If mblnConn Then Call objPubOneCard.SetPara("向后滚动-快键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "向后滚动-快键", strValue)
    
    strValue = Trim(cboFun(2).Text)
    If mblnConn Then Call objPubOneCard.SetPara("读卡-功能键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "读卡-功能键", strValue)
    strValue = Trim(cboFastkey(2).Text)
    If mblnConn Then Call objPubOneCard.SetPara("读卡-快键", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "医疗卡类别", "读卡-快键", strValue)
    With vsGrid
        For i = 1 To .Rows - 1
            If Left(.RowData(i), 1) = "K" Then
                strValue = IIf(Val(.TextMatrix(i, .ColIndex("启用"))) = 0, 0, 1)
                Call SaveRegInFor(mRegType, "医疗卡类别\" & .TextMatrix(i, .ColIndex("名称")), "启用", strValue)
                '103310:李南春,2016/12/7,启用回车后增加卡号长度
                strValue = Trim(.TextMatrix(i, .ColIndex("回车符")))
                Call SaveRegInFor(mRegType, "医疗卡类别\" & .TextMatrix(i, .ColIndex("名称")), "回车符", strValue)
                strValue = Trim(.TextMatrix(i, .ColIndex("功能键")))
                Call SaveRegInFor(mRegType, "医疗卡类别\" & .TextMatrix(i, .ColIndex("名称")), "读卡-功能键", strValue)
                strValue = Trim(.TextMatrix(i, .ColIndex("快键")))
                Call SaveRegInFor(mRegType, "医疗卡类别\" & .TextMatrix(i, .ColIndex("名称")), "读卡-快键", strValue)
            End If
        Next
    End With
    '78768:李南春,2014/11/26,缺省读卡类别
    '103309：李南春，2016/12/7，禁止录入缺省读卡类别
    With cboDefaultCard
        If InStr(1, mstrPrivs, ";参数设置;") > 0 Then
            If .ListIndex < 0 Then .ListIndex = 0
            Call objPubOneCard.SetPara("缺省读卡类别", .ItemData(.ListIndex), glngSys, 1153, True)
            Call SaveRegInFor(mRegType, "医疗卡类别", "缺省读卡类别", .ItemData(.ListIndex))
        End If
        If Not mblnConn Then Call SaveRegInFor(mRegType, "医疗卡类别", "缺省读卡类别", .ItemData(.ListIndex))
    End With
    Set objPubOneCard = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function InitData(Optional blnRestoreDefault As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-12-05 14:59:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strValue As String, strA1 As String
    Dim objPubOneCard As Object
    
    On Error GoTo errHandle
    
    Call zlGetPubOneCard(mcnOracle, objPubOneCard)
 
        
    If Not blnRestoreDefault Then
        '78768:李南春,2014/11/26,将快键保存到参数表
        If mblnConn Then
            strA1 = objPubOneCard.getPara("向前滚动-功能键", glngSys, 1153)
            strValue = objPubOneCard.getPara("向前滚动-快键", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "医疗卡类别", "向前滚动-功能键", strA1)
            Call GetRegInFor(mRegType, "医疗卡类别", "向前滚动-快键", strValue)
        End If
    End If
    
    strA1 = IIf(strA1 = "", "SHIFT", strA1)
    Call SetCboFun(cboFun(0), strA1)
    Call SetCboFastkey(cboFastkey(0), cboFun(0).Text, IIf(strValue = "", "F4", strValue))
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("向后滚动-功能键", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "医疗卡类别", "向后滚动-功能键", strValue)
        End If
    End If
    Call SetCboFun(cboFun(1), strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("向后滚动-快键", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "医疗卡类别", "向后滚动-快键", strValue)
        End If
    End If
    strValue = IIf(strValue = "", "F4", strValue) '缺省F4
    Call SetCboFastkey(cboFastkey(1), cboFun(1).Text, strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("读卡-功能键", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "医疗卡类别", "读卡-功能键", strValue)
        End If
    End If
    Call SetCboFun(cboFun(2), strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("读卡-快键", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "医疗卡类别", "读卡-快键", strValue)
        End If
    End If
    strValue = IIf(strValue = "", "空格键", strValue) '缺省空格键
    Call SetCboFastkey(cboFastkey(2), cboFun(2).Text, strValue)
    With vsGrid
        .Clear 1

        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .Rows = mobjCards.Count + 1
        For i = 1 To mobjCards.Count
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "医疗卡类别\" & mobjCards(i).名称, "启用", strValue)
            End If
            If strValue = "" Then   '缺省启用
                .TextMatrix(i, .ColIndex("启用")) = 1
            Else
                .TextMatrix(i, .ColIndex("启用")) = Val(strValue)
            End If
            
            '103310:李南春,2016/12/7,启用回车后增加卡号长度
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "医疗卡类别\" & mobjCards(i).名称, "回车符", strValue)
            End If
            If strValue = "" Then   '缺省启用
                .TextMatrix(i, .ColIndex("回车符")) = "缺省"
            Else
                .TextMatrix(i, .ColIndex("回车符")) = strValue
            End If
            
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "医疗卡类别\" & mobjCards(i).名称, "读卡-功能键", strValue)
            End If
            .TextMatrix(i, .ColIndex("功能键")) = IIf(strValue = "", " ", strValue)
            
            strValue = ""
            If Not blnRestoreDefault Then
                    Call GetRegInFor(mRegType, "医疗卡类别\" & mobjCards(i).名称, "读卡-快键", strValue)
            End If
            .TextMatrix(i, .ColIndex("快键")) = IIf(strValue = "", " ", strValue)
            .TextMatrix(i, .ColIndex("名称")) = mobjCards(i).名称
            If InStr(1, "," & mstrMustSelectItems & ",", "," & .TextMatrix(i, .ColIndex("名称")) & ",") > 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000011
                .Cell(flexcpForeColor, i, .ColIndex("回车符")) = &H80000008
            End If
            '78768:李南春,2014/11/26,缺省读卡类别
            .TextMatrix(i, .ColIndex("刷卡")) = IIf(mobjCards(i).是否刷卡, 1, 0)
            .Cell(flexcpData, i, .ColIndex("名称")) = mobjCards(i).接口序号
            .TextMatrix(i, .ColIndex("扫描")) = IIf(mobjCards(i).是否扫描, 1, 0)
            .TextMatrix(i, .ColIndex("接触式读卡")) = IIf(mobjCards(i).是否接触式读卡, 1, 0)
            .TextMatrix(i, .ColIndex("非接触式读卡")) = IIf(mobjCards(i).是否非接触式读卡, 1, 0)
            .RowData(i) = "K" & i
            
        Next
        .Editable = flexEDKbdMouse
    End With
    InitData = True
    Set objPubOneCard = Nothing
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AddCboKey(ByVal cboFun As ComboBox, ByVal strFunKey As String, ByVal strKey As String, strDeult As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:增加合法的快键
    '入参:strKey - String
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-27 16:31:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varTemp As Variant
    For i = 0 To UBound(mvarNotKey)
        varTemp = Split(mvarNotKey(i) & "+", "+")
        If strFunKey = "" Then
            If varTemp(0) = "" And varTemp(1) = strKey Then Exit Sub
        ElseIf strFunKey = "CTRL" Then
            If varTemp(0) = "CTRL" And varTemp(1) = strKey Then Exit Sub
        ElseIf strFunKey = "SHIFT" Then
            If varTemp(0) = "SHIFT" And varTemp(1) = strKey Then Exit Sub
        End If
    Next
    With cboFun
        .AddItem strKey
        If strKey = strDeult Then .ListIndex = .NewIndex
    End With
End Sub

Private Sub SetCboFastkey(ByVal cboFun As ComboBox, ByVal strFunKey As String, ByVal strDefaultValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置快键
    '编制:刘兴洪
    '日期:2012-08-22 15:54:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFastkey As String
    Dim i As Long
    With cboFun
        .Clear
        .AddItem " "
        Call AddCboKey(cboFun, strFunKey, "空格键", strDefaultValue)
    End With
    For i = 1 To 12
        Call AddCboKey(cboFun, strFunKey, "F" & i, strDefaultValue)
    Next
    Call AddCboKey(cboFun, strFunKey, "↑", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "↓", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "←", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "→", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "<", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, ">", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "[", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "]", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "/", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "\", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "*", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "`", strDefaultValue)
        
    Call AddCboKey(cboFun, strFunKey, "+", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "-", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "=", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "(", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, ")", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "?", strDefaultValue)
    For i = Asc("A") To Asc("Z")
        Call AddCboKey(cboFun, strFunKey, Chr(i), strDefaultValue)
    Next
    For i = Asc("0") To Asc("9")
        Call AddCboKey(cboFun, strFunKey, Chr(i), strDefaultValue)
    Next
        
    '数字键盘
    For i = Asc("0") To Asc("9")
        Call AddCboKey(cboFun, strFunKey, "NUM" & Chr(i), strDefaultValue)
    Next
    Call AddCboKey(cboFun, strFunKey, "NUM*", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM+", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM-", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM/", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM.", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUMENTER", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, UCase("KeyLButton"), strDefaultValue)      'vbKeyLButton 0x1 鼠标左键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMButton"), strDefaultValue)     'vbKeyRButton 0x2 鼠标右键'
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMButton"), strDefaultValue)     'vbKeyMButton 0x4 鼠标中键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyCancel"), strDefaultValue)     ' vbKeyCancel 0x3 CANCEL 键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyBack"), strDefaultValue)     ' vbKeyBack 0x8 BACKSPACE 键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyTab"), strDefaultValue)     ' vbKeyTab 0x9 TAB 键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyClear"), strDefaultValue)     ' vbKeyClear 0xC CLEAR 键
    Call AddCboKey(cboFun, strFunKey, UCase("ENTER"), strDefaultValue)     ' vbKeyReturn 0xD ENTER 键
    Call AddCboKey(cboFun, strFunKey, UCase("SHIFT"), strDefaultValue)     ' vbKeyShift 0x10 SHIFT 键
    Call AddCboKey(cboFun, strFunKey, UCase("CTRL"), strDefaultValue)     ' vbKeyControl 0x11 CTRL 键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMenu"), strDefaultValue)     ' vbKeyMenu 0x12 MENU 键
    Call AddCboKey(cboFun, strFunKey, UCase("KeyPause"), strDefaultValue)     ' vbKeyPause 0x13 PAUSE 键
    Call AddCboKey(cboFun, strFunKey, UCase(" CAPS LOCK"), strDefaultValue)     ' vbKeyCapital 0x14 CAPS LOCK 键
    Call AddCboKey(cboFun, strFunKey, UCase("ESC"), strDefaultValue)     ' vbKeyEscape 0x1B ESC 键
    Call AddCboKey(cboFun, strFunKey, UCase("SELECT"), strDefaultValue)     ' vbKeySelect 0x29 SELECT 键
    Call AddCboKey(cboFun, strFunKey, UCase("PRINT SCREEN"), strDefaultValue)     ' vbKeyPrint 0x2A PRINT SCREEN 键
    Call AddCboKey(cboFun, strFunKey, UCase("EXECUTE"), strDefaultValue)     ' vbKeyExecute 0x2B EXECUTE 键
    Call AddCboKey(cboFun, strFunKey, UCase("SNAPSHOT"), strDefaultValue)     ' vbKeySnapshot 0x2C SNAPSHOT 键
    Call AddCboKey(cboFun, strFunKey, UCase("INSERT"), strDefaultValue)    ' vbKeyInsert 0x2D INSERT 键
    Call AddCboKey(cboFun, strFunKey, UCase("DELETE"), strDefaultValue)     ' vbKeyDelete 0x2E DELETE 键
    Call AddCboKey(cboFun, strFunKey, UCase("HELP"), strDefaultValue)     ' vbKeyHelp 0x2F HELP 键
    Call AddCboKey(cboFun, strFunKey, UCase("NUM LOCK"), strDefaultValue)     ' vbKeyNumlock 0x90 NUM LOCK 键
    If cboFun.ListIndex < 0 Then cboFun.ListIndex = 0
End Sub

Private Sub SetCboFun(ByVal cboFun As ComboBox, ByVal strDefaultValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能键
    '编制:刘兴洪
    '日期:2012-08-22 15:54:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With cboFun
        .Clear
        .AddItem " "
        If strDefaultValue = "" Then .ListIndex = .NewIndex
        .AddItem "CTRL"
        If strDefaultValue = "CTRL" Then .ListIndex = .NewIndex
        .AddItem "SHIFT"
        If strDefaultValue = "SHIFT" Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub

Public Function ShowSetWin(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, _
    ByVal objCards As Cards, ByVal RegType As gRegType, _
    Optional strNotContainFastKey As String = "", _
    Optional strMustSelectItems As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:启用或停用卡类别
    '入参:strNotContainFastKey-不能包含的快键
    '返回:设置点击确定后,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-25 23:20:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjCards = objCards: mstrNotContainFastKey = strNotContainFastKey
    Set mcnOracle = cnOracle
    
    mstrMustSelectItems = strMustSelectItems
    mRegType = RegType
    mblnOk = False
    If frmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmMain
    End If
    ShowSetWin = mblnOk
End Function

Private Sub cboDefaultCard_KeyPress(KeyAscii As Integer)
    '不允许输入
    KeyAscii = 0
End Sub

Private Sub cboFastkey_Click(Index As Integer)
    lblFun(Index).Caption = "快键:" & cboFun(Index).Text & IIf(Trim(cboFun(Index).Text) = "", " ", " + ") & cboFastkey(Index).Text
End Sub

Private Sub cboFun_Click(Index As Integer)
    lblFun(Index).Caption = "快键:" & cboFun(Index).Text & IIf(Trim(cboFun(Index).Text) = "", " ", " + ") & cboFastkey(Index).Text
    Call SetCboFastkey(cboFastkey(Index), cboFun(Index).Text, cboFastkey(Index).Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefault_Click()
   Call InitData(True)
   If isValied = False Then Exit Sub
  Call SaveParaSet
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    Call SaveParaSet
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    Call InitDefaultCard
    If vsGrid.Enabled And vsGrid.Visible Then vsGrid.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        Unload Me: Exit Sub
    End If
End Sub
Private Sub Form_Load()
    Dim objPubOneCard As Object
    
    mblnFirst = True
    mvarNotKey = Split(mstrNotContainFastKey, ";")
    mblnConn = Not gcnOracle Is Nothing
    Call InitFace
    If mblnConn Then
        Call zlGetPubOneCard(mcnOracle, objPubOneCard)
        mstrPrivs = ";" & objPubOneCard.GetPrivFunc(glngSys, 1153) & ";"
        Set objPubOneCard = Nothing
    End If
    
End Sub

Private Sub InitFace()
    With vsGrid
        .Clear 1
        .ColComboList(.ColIndex("回车符")) = "缺省|启用|禁用"
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("启用"), .ColIndex("功能键"), .ColIndex("快键")
            If Left(.RowData(Row), 1) <> "K" Then Cancel = True: Exit Sub
            If InStr(1, "," & mstrMustSelectItems & ",", "," & .TextMatrix(Row, .ColIndex("名称")) & ",") > 0 Then Cancel = True
        Case .ColIndex("回车符")
            If Left(.RowData(Row), 1) <> "K" Then Cancel = True: Exit Sub
            If mobjCards(Row).接口序号 <= 0 Or Not (mobjCards(Row).是否刷卡 Or mobjCards(Row).是否扫描) Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub InitDefaultCard()
'将缺省读卡类型保存在参数表和注册表
    Dim lngDefaultCardID As Long, strValue As String, i As Integer, j As Integer
    Dim objPubOneCard As Object
    If mblnConn Then
        Call zlGetPubOneCard(mcnOracle, objPubOneCard)
        strValue = objPubOneCard.getPara("缺省读卡类别", glngSys, 1153, 0, Array(lblDefaultCard, cboDefaultCard), InStr(1, mstrPrivs, ";参数设置;") > 0)
    Else
        Call GetRegInFor(mRegType, "医疗卡类别", "缺省读卡类别", strValue)
    End If
    lngDefaultCardID = Val(strValue)
    cboDefaultCard.Clear
    cboDefaultCard.AddItem ""
    cboDefaultCard.ItemData(cboDefaultCard.NewIndex) = 0
    cboDefaultCard.ListIndex = cboDefaultCard.NewIndex
    With vsGrid
        For i = 1 To .Rows - 1
            '85565,李南春,2015/7/10:读卡性质
            'If .TextMatrix(i, .ColIndex("启用")) <> 0 And .TextMatrix(i, .ColIndex("刷卡")) = 0 Then
            If .TextMatrix(i, .ColIndex("启用")) <> 0 And (.TextMatrix(i, .ColIndex("接触式读卡")) = 1 Or .TextMatrix(i, .ColIndex("非接触式读卡")) = 1) Then
                cboDefaultCard.AddItem .TextMatrix(i, .ColIndex("名称"))
                cboDefaultCard.ItemData(cboDefaultCard.NewIndex) = .Cell(flexcpData, i, .ColIndex("名称"))
                If lngDefaultCardID = .Cell(flexcpData, i, .ColIndex("名称")) Then cboDefaultCard.ListIndex = cboDefaultCard.NewIndex
            End If
        Next
    End With
End Sub

Private Sub vsGrid_Validate(Cancel As Boolean)
    If cboDefaultCard.Enabled And cboDefaultCard.Visible Then
        Call InitDefaultCard
    End If
End Sub
