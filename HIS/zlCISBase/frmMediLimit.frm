VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediLimit 
   Caption         =   "药品储备设置"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   Icon            =   "frmMediLimit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10575
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   2445
      Left            =   6240
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   4313
      Appearance      =   0
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediLimit.frx":058A
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin MSComctlLib.ImageList imgStoreRoom 
      Left            =   9960
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediLimit.frx":060D
            Key             =   "StroeRoomPic"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt查找 
      Height          =   300
      Left            =   6600
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox cboDrugUnit 
      Height          =   300
      ItemData        =   "frmMediLimit.frx":6E6F
      Left            =   1800
      List            =   "frmMediLimit.frx":6E7F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1065
      Width           =   2400
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6300
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediLimit.frx":6EAB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   5040
      Width           =   9810
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&T)"
         Height          =   350
         Left            =   5350
         TabIndex        =   16
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "应用于本列(&O)"
         Height          =   350
         Left            =   3990
         TabIndex        =   11
         Top             =   165
         Width           =   1365
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "恢复(&R)"
         Height          =   350
         Left            =   2685
         Picture         =   "frmMediLimit.frx":773D
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全部清除(&C)"
         Height          =   350
         Left            =   1380
         Picture         =   "frmMediLimit.frx":7887
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   165
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   6720
         TabIndex        =   7
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   90
         Picture         =   "frmMediLimit.frx":79D1
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   165
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭(&X)"
         Height          =   350
         Left            =   7920
         TabIndex        =   8
         Top             =   165
         Width           =   1100
      End
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -195
      TabIndex        =   6
      Top             =   1440
      Width           =   9810
   End
   Begin VB.ComboBox cboRoom 
      Height          =   276
      Left            =   1800
      TabIndex        =   2
      Text            =   "cboRoom"
      Top             =   585
      Width           =   2400
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfLimit 
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   5895
      _cx             =   10398
      _cy             =   2355
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediLimit.frx":7B1B
      ScrollTrack     =   0   'False
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
   Begin VB.Label lbl查找 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5880
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lblComment1 
      AutoSize        =   -1  'True
      Caption         =   "按F3进行连续查找"
      Height          =   180
      Left            =   8550
      TabIndex        =   18
      Top             =   660
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblDrugUnit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单位(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   3
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品库房(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   645
      Width           =   990
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmMediLimit.frx":7EF7
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    选择药品库房后，指定该库房药品的储备限量；并根据药品的管理要求，可以同时指定其盘点属性和库房货位。"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   150
      Width           =   7725
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "药品在各库房的限额与盘点要求(&T)："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2970
   End
End
Attribute VB_Name = "frmMediLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前材质：由me.tag保存,分别为"5","6","7"
'   2、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序传入
'   3、指定药品：由me.lblMedi.tag保存，由上级程序传入可以传递，也可以不传递
'---------------------------------------------------
Public strPrivs As String       '当前用户具有的本程序权限

Private mrsNormal As New ADODB.Recordset
Private mintCount As Integer
Private mlng库房ID As Long
Private mlngFind As Long
Private mlngFindFirst As Long
Private mrsFindName As ADODB.Recordset
Private mblnChanged As Boolean
Private mstr分类 As String
Private mstr分类ID  As String
Private mstr剂型 As String
Private mblnActive As Boolean
Private mint药品名称显示 As Integer         '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
Private mlngRow As Long     '用来记录点击货位按钮时的行
Private Const mlngBorderColor As Long = &H8000000D     '选中行边框颜色
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' 没选中行边框颜色
Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long, lngRows As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim str其他名称 As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strFindStyle As String
    Dim strTmp As String
    
    '查找药品
    On Error GoTo errHandle
    If strInput = txt查找.Tag Then
        '表示查找下一条记录
        If mlngFind >= vsfLimit.Rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '表示新的查找
        lngStart = 0
        mlngFindFirst = 0
        txt查找.Tag = strInput
        
        strFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
        
        Set mrsFindName = New ADODB.Recordset

        gstrSql = "Select Distinct A.Id,A.编码 From 收费项目目录 A,收费项目别名 B" & _
                 " Where A.Id =B.收费细目id And A.类别=[1] "

        If IsNumeric(Replace(strInput, "-", "")) Then       '输入全是数字（或包含一个"-"）时只匹配编码
            gstrSql = gstrSql & " And A.编码 Like [2] Or B.简码 Like [2] And B.码类=3 "
        ElseIf zlStr.IsCharAlpha(strInput) Then          '输入全是字母时只匹配简码
            gstrSql = gstrSql & " And B.简码 Like [3] "
        ElseIf zlStr.IsCharChinese(strInput) Then        '输入全是汉字时只匹配名称
            gstrSql = gstrSql & " And B.名称 Like [3] "
        Else
            gstrSql = gstrSql & " And (A.编码 Like [2] Or B.名称 Like [3] Or B.简码 Like [3] )"
        End If
        
        gstrSql = gstrSql & " Order By A.编码 "
                 
        Set mrsFindName = zldatabase.OpenSQLRecord(gstrSql, "取匹配的药品ID", Me.Tag, strInput & "%", strFindStyle & strInput & "%")
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    lngStart = lngStart + 1
    lngRows = vsfLimit.Rows - 1
    
    With mrsFindName
        If .EOF Then .MoveFirst
        
        Do While Not .EOF
            lngFindRow = vsfLimit.FindRow(!编码, 0, vsfLimit.ColIndex("编码"), True, True)
            If lngFindRow > 0 Then
                vsfLimit.Select lngFindRow, 1, lngFindRow, vsfLimit.Cols - 1
                vsfLimit.TopRow = lngFindRow
                mlngFind = lngFindRow
                
                '记录找到的第1条记录
                If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
                
                mrsFindName.MoveNext
                Exit Do
            End If
            mrsFindName.MoveNext
    
            '如果到底了，则返回第1条记录
            If .EOF And lngFindRow = -1 Then
                vsfLimit.Select mlngFindFirst, 1, mlngFindFirst, vsfLimit.Cols - 1
                vsfLimit.TopRow = mlngFindFirst
                mlngFind = mlngFindFirst
            End If
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub IniGrid()
    With vsfLimit
        .Redraw = flexRDNone
        .Rows = 1
        .SelectionMode = flexSelectionFree
        .ExplorerBar = flexExSortShowAndMove
        .Editable = flexEDNone
        
        .ColComboList(.ColIndex("货位")) = "..."
        
        .ColWidth(.ColIndex("商品名")) = IIf(mint药品名称显示 = 2, 2000, 0)
        
        .TextMatrix(0, .ColIndex("产地")) = "生产商"
        .ColWidth(.ColIndex("规格")) = 1500 'IIf(Me.Tag <> "7", 1500, 0)
        .ColWidth(.ColIndex("产地")) = 1200
        .ColHidden(.ColIndex("原产地")) = IIf(Me.Tag = "7", False, True)
        
        If InStr(1, strPrivs, "上下限控制") > 0 Then
            .ColHidden(.ColIndex("上限")) = False
            .ColHidden(.ColIndex("下限")) = False
            
            If .ColWidth(.ColIndex("上限")) = 0 Then .ColWidth(.ColIndex("上限")) = 1050
            If .ColWidth(.ColIndex("下限")) = 0 Then .ColWidth(.ColIndex("下限")) = 1050
        Else
            .ColHidden(.ColIndex("上限")) = True
            .ColHidden(.ColIndex("下限")) = True
        End If
        
        If InStr(1, strPrivs, "盘点属性设置") > 0 Then
            .ColHidden(.ColIndex("日盘")) = False
            .ColHidden(.ColIndex("周盘")) = False
            .ColHidden(.ColIndex("月盘")) = False
            .ColHidden(.ColIndex("季盘")) = False
            
            If .ColWidth(.ColIndex("日盘")) = 0 Then .ColWidth(.ColIndex("日盘")) = 500
            If .ColWidth(.ColIndex("周盘")) = 0 Then .ColWidth(.ColIndex("周盘")) = 500
            If .ColWidth(.ColIndex("月盘")) = 0 Then .ColWidth(.ColIndex("月盘")) = 500
            If .ColWidth(.ColIndex("季盘")) = 0 Then .ColWidth(.ColIndex("季盘")) = 500
        Else
            .ColHidden(.ColIndex("日盘")) = True
            .ColHidden(.ColIndex("周盘")) = True
            .ColHidden(.ColIndex("月盘")) = True
            .ColHidden(.ColIndex("季盘")) = True
        End If
        
        .ColDataType(.ColIndex("上限")) = flexDTDouble
        .ColDataType(.ColIndex("下限")) = flexDTDouble
        
        .Cell(flexcpForeColor, 0, .ColIndex("允许领用")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("上限")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("下限")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("日盘")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("周盘")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("月盘")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("季盘")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("货位")) = vbBlue
        
        .Redraw = flexRDDirect
    End With

End Sub

Private Sub cboDrugUnit_Click()
'不再读服务器数据，直接界面换算刷新
    Dim i As Long
    
    If Val(cboDrugUnit.Tag) = cboDrugUnit.ListIndex Or cboDrugUnit.Tag = "-1" Then Exit Sub
    
    With Me.vsfLimit
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            '还原成售价单位
            Select Case Val(cboDrugUnit.Tag)
                Case 1  '住院单位
                    .TextMatrix(i, .ColIndex("上限")) = Val(.TextMatrix(i, .ColIndex("上限"))) * Val(.TextMatrix(i, .ColIndex("住院包装")))
                    .TextMatrix(i, .ColIndex("下限")) = Val(.TextMatrix(i, .ColIndex("下限"))) * Val(.TextMatrix(i, .ColIndex("住院包装")))
                Case 2  '门诊单位
                    .TextMatrix(i, .ColIndex("上限")) = Val(.TextMatrix(i, .ColIndex("上限"))) * Val(.TextMatrix(i, .ColIndex("门诊包装")))
                    .TextMatrix(i, .ColIndex("下限")) = Val(.TextMatrix(i, .ColIndex("下限"))) * Val(.TextMatrix(i, .ColIndex("门诊包装")))
                Case 3  '药库单位
                    .TextMatrix(i, .ColIndex("上限")) = Val(.TextMatrix(i, .ColIndex("上限"))) * Val(.TextMatrix(i, .ColIndex("药库包装")))
                    .TextMatrix(i, .ColIndex("下限")) = Val(.TextMatrix(i, .ColIndex("下限"))) * Val(.TextMatrix(i, .ColIndex("药库包装")))
            End Select
            
            '开始换算
            Select Case cboDrugUnit.ListIndex
                Case 0  '售价单位
                    .TextMatrix(i, .ColIndex("单位")) = .TextMatrix(i, .ColIndex("售价单位"))
                    .TextMatrix(i, .ColIndex("包装")) = 1
                    .TextMatrix(i, .ColIndex("零售价")) = Format(.TextMatrix(i, .ColIndex("固定零售价")), "0.000")
                    .TextMatrix(i, .ColIndex("库存数量")) = Format(Val(.TextMatrix(i, .ColIndex("实际数量"))), "0.00")
                Case 1  '住院单位
                    .TextMatrix(i, .ColIndex("单位")) = .TextMatrix(i, .ColIndex("住院单位"))
                    .TextMatrix(i, .ColIndex("包装")) = .TextMatrix(i, .ColIndex("住院包装"))
                    .TextMatrix(i, .ColIndex("零售价")) = Format(Val(.TextMatrix(i, .ColIndex("固定零售价"))) * Val(.TextMatrix(i, .ColIndex("住院包装"))), "0.000")
                    .TextMatrix(i, .ColIndex("库存数量")) = Format(Val(.TextMatrix(i, .ColIndex("实际数量"))) / Val(.TextMatrix(i, .ColIndex("住院包装"))), "0.00")
                    .TextMatrix(i, .ColIndex("上限")) = Format(Val(.TextMatrix(i, .ColIndex("上限"))) / Val(.TextMatrix(i, .ColIndex("住院包装"))), "0.00000")
                    .TextMatrix(i, .ColIndex("下限")) = Format(Val(.TextMatrix(i, .ColIndex("下限"))) / Val(.TextMatrix(i, .ColIndex("住院包装"))), "0.00000")
                Case 2  '门诊单位
                    .TextMatrix(i, .ColIndex("单位")) = .TextMatrix(i, .ColIndex("门诊单位"))
                    .TextMatrix(i, .ColIndex("包装")) = .TextMatrix(i, .ColIndex("门诊包装"))
                    .TextMatrix(i, .ColIndex("零售价")) = Format(Val(.TextMatrix(i, .ColIndex("固定零售价"))) * Val(.TextMatrix(i, .ColIndex("门诊包装"))), "0.000")
                    .TextMatrix(i, .ColIndex("库存数量")) = Format(Val(.TextMatrix(i, .ColIndex("实际数量"))) / Val(.TextMatrix(i, .ColIndex("门诊包装"))), "0.00")
                    .TextMatrix(i, .ColIndex("上限")) = Format(Val(.TextMatrix(i, .ColIndex("上限"))) / Val(.TextMatrix(i, .ColIndex("门诊包装"))), "0.00000")
                    .TextMatrix(i, .ColIndex("下限")) = Format(Val(.TextMatrix(i, .ColIndex("下限"))) / Val(.TextMatrix(i, .ColIndex("门诊包装"))), "0.00000")
                Case 3  '药库单位
                    .TextMatrix(i, .ColIndex("单位")) = .TextMatrix(i, .ColIndex("药库单位"))
                    .TextMatrix(i, .ColIndex("包装")) = .TextMatrix(i, .ColIndex("药库包装"))
                    .TextMatrix(i, .ColIndex("零售价")) = Format(Val(.TextMatrix(i, .ColIndex("固定零售价"))) * Val(.TextMatrix(i, .ColIndex("药库包装"))), "0.000")
                    .TextMatrix(i, .ColIndex("库存数量")) = Format(Val(.TextMatrix(i, .ColIndex("实际数量"))) / Val(.TextMatrix(i, .ColIndex("药库包装"))), "0.00")
                    .TextMatrix(i, .ColIndex("上限")) = Format(Val(.TextMatrix(i, .ColIndex("上限"))) / Val(.TextMatrix(i, .ColIndex("药库包装"))), "0.00000")
                    .TextMatrix(i, .ColIndex("下限")) = Format(Val(.TextMatrix(i, .ColIndex("下限"))) / Val(.TextMatrix(i, .ColIndex("药库包装"))), "0.00000")
            End Select
        Next
        .Redraw = flexRDBuffered
    End With
    cboDrugUnit.Tag = cboDrugUnit.ListIndex
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    If Me.Tag = "5" Then
        str工作性质 = "I,M,K"
    ElseIf Me.Tag = "6" Then
        str工作性质 = "N,J,K"
    Else
        str工作性质 = "L,H,K"
    End If

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRoom.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfLimit): Exit Sub
    
    If cboRoom.ListIndex >= 0 Then
        If Val(cboRoom.Tag) = cboRoom.ItemData(cboRoom.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfLimit, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboRoom, Trim(cboRoom.Text), str工作性质, IIf(InStr(1, strPrivs, "允许设置所有库房限额盘点") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cboRoom.ListIndex >= 0 Then
        cboRoom.Tag = cboRoom.ItemData(cboRoom.ListIndex)
    End If
End Sub

Private Sub cboRoom_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboRoom_Validate(Cancel As Boolean)
    If cboRoom.ListCount > 0 Then
        If cboRoom.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cmdClear_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    
    If MsgBox("将清除所有设置，是否确定？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    With Me.vsfLimit
        .Redraw = flexRDNone
        
        .Cell(flexcpText, 1, 0, .Rows - 1, 0) = ""
        
        If InStr(1, strPrivs, "上下限控制") > 0 Then
            .Cell(flexcpText, 1, .ColIndex("上限"), .Rows - 1, .ColIndex("上限")) = Format(0, "0.00000")
            .Cell(flexcpText, 1, .ColIndex("下限"), .Rows - 1, .ColIndex("下限")) = Format(0, "0.00000")
        End If
        
        If InStr(1, strPrivs, "盘点属性设置") > 0 Then
            .Cell(flexcpText, 1, .ColIndex("日盘"), .Rows - 1, .ColIndex("日盘")) = ""
            .Cell(flexcpText, 1, .ColIndex("周盘"), .Rows - 1, .ColIndex("周盘")) = ""
            .Cell(flexcpText, 1, .ColIndex("月盘"), .Rows - 1, .ColIndex("月盘")) = ""
            .Cell(flexcpText, 1, .ColIndex("季盘"), .Rows - 1, .ColIndex("季盘")) = ""
        End If
        
        .Cell(flexcpText, 1, .ColIndex("货位"), .Rows - 1, .ColIndex("货位")) = ""
        
        .Redraw = flexRDBuffered
    End With

End Sub
Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdFilter_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    
    If frmMediLimitFilter.GetCondition(Me, mlng库房ID, Me.Tag, mstr分类, mstr分类ID, mstr剂型) = True Then
        If mblnChanged = True Then
            If MsgBox("当前修改未保存，是否按过滤条件提取数据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Call zlLimitRef
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub cmdRestore_Click()
    If Me.vsfLimit.Rows = 1 Then Exit Sub
    If MsgBox("将恢复所有设置，是否确定？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlLimitRef
End Sub


Private Sub cmdSave_Click()
    Dim strMsgBox As String, strErrors As String
    Dim intNewStocks As Integer
    Dim intMaxStocks As Integer
    
    strErrors = ""
    
    If mblnChanged = False Then Exit Sub
    
    With Me.vsfLimit
        For mintCount = 1 To .Rows - 1
            If Val(.TextMatrix(mintCount, .ColIndex("上限"))) <> 0 _
                And Val(.TextMatrix(mintCount, .ColIndex("上限"))) < Val(.TextMatrix(mintCount, .ColIndex("下限"))) Then
                .TextMatrix(mintCount, 0) = "？"
                strErrors = strErrors & vbCrLf & .TextMatrix(mintCount, .ColIndex("编码")) & "-" & .TextMatrix(mintCount, .ColIndex("名称"))
                strMsgBox = "“" & .TextMatrix(mintCount, .ColIndex("编码")) & "-" & .TextMatrix(mintCount, .ColIndex("名称")) & "”的储备下限大于储备上限！" & _
                        vbCrLf & vbCrLf & "继续保存其他药品吗？"
                If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbThis.Panels(2).Text = ""
                    .TopRow = mintCount: .Row = mintCount: .SetFocus: Exit Sub
                End If
            ElseIf Val(.RowData(mintCount)) <> 0 Then
                gstrSql = "zl_药品储备限额_Update(" & Me.cboRoom.ItemData(Me.cboRoom.ListIndex)
                gstrSql = gstrSql & "," & .RowData(mintCount)
                gstrSql = gstrSql & "," & Format(Val(.TextMatrix(mintCount, .ColIndex("上限"))) * Val(.TextMatrix(mintCount, .ColIndex("包装"))), "0.00000")
                gstrSql = gstrSql & "," & Format(Val(.TextMatrix(mintCount, .ColIndex("下限"))) * Val(.TextMatrix(mintCount, .ColIndex("包装"))), "0.00000")
                gstrSql = gstrSql & ",'" & IIf(Trim(.TextMatrix(mintCount, .ColIndex("日盘"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("周盘"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("月盘"))) = "", "0", "1")
                gstrSql = gstrSql & IIf(Trim(.TextMatrix(mintCount, .ColIndex("季盘"))) = "", "0", "1")
                gstrSql = gstrSql & "','" & Trim(.TextMatrix(mintCount, .ColIndex("货位"))) & "'"
                gstrSql = gstrSql & "," & IIf(Trim(.TextMatrix(mintCount, .ColIndex("允许领用"))) = "", "0", "1")
                gstrSql = gstrSql & ")"
                err = 0: On Error Resume Next
                Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                If err <> 0 Then
                    Call SaveErrLog
                    err = 0
                    .TextMatrix(mintCount, 0) = "？"
                    strErrors = strErrors & vbCrLf & .TextMatrix(mintCount, .ColIndex("编码")) & "-" & .TextMatrix(mintCount, .ColIndex("名称"))
                    strMsgBox = "保存“" & .TextMatrix(mintCount, .ColIndex("编码")) & .TextMatrix(mintCount, .ColIndex("名称")) & "”时发生错误！" & _
                            vbCrLf & vbCrLf & "继续保存其他药品吗？"
                    If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Me.stbThis.Panels(2).Text = ""
                        .TopRow = mintCount: .Row = mintCount: .SetFocus: Exit Sub
                    End If
                End If
                If mintCount Mod IIf(.Rows > 20, .Rows \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "正在保存：" & String(mintCount \ IIf(.Rows > 20, .Rows \ 20, 1), "…")
                End If
            End If
        Next
    End With
    Me.stbThis.Panels(2).Text = ""
    strMsgBox = "“" & Me.cboRoom.Text & "”储备特性保存完毕！"
    If strErrors <> "" Then
        strMsgBox = strMsgBox & vbCrLf & "但以下药品发生错误，请检查：" & strErrors
    End If
    MsgBox strMsgBox, vbExclamation, gstrSysName

End Sub

Private Sub cmdApply_Click()
    Dim strValue As String
    
    With vsfLimit
        If .Rows = 1 Then Exit Sub
        
        Select Case .Col
            Case .ColIndex("日盘"), .ColIndex("周盘"), .ColIndex("月盘"), .ColIndex("季盘")
            Case .ColIndex("上限"), .ColIndex("下限")
                If InStr(1, strPrivs, "上下限控制") = 0 Then Exit Sub
            Case .ColIndex("货位")
                If InStr(1, strPrivs, "调整库位") = 0 Then Exit Sub
            Case Else
                Exit Sub
        End Select
        
        If MsgBox("将[" & .TextMatrix(0, .Col) & "]列的内容应用到所有药品，是否确定？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
        '将当前列的内容应用到所有药品相同列
        strValue = .TextMatrix(.Row, .Col)
        .Cell(flexcpText, 1, .Col, .Rows - 1, .Col) = strValue
    End With
End Sub
Private Sub Form_Activate()
    If mblnActive = True Then Exit Sub
    
    If Me.cmdClose.Tag = "查阅" Then
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    End If
    lbl查找.Visible = True
    txt查找.Visible = True
    lblComment1.Visible = True
    
    err = 0: On Error GoTo ErrHand
    gstrSql = "select ID,编码,名称" & _
            "  from 部门表 D"
    If Me.Tag = "5" Then
        gstrSql = gstrSql & " where ID in (select distinct 部门id from 部门性质说明 where 工作性质 like '西药%' or 工作性质='制剂室') and (d.撤档时间 is null or to_char(d.撤档时间,'yyyy-mm-dd')='3000-01-01')"
    ElseIf Me.Tag = "6" Then
        gstrSql = gstrSql & " where ID in (select distinct 部门id from 部门性质说明 where 工作性质 like '成药%' or 工作性质='制剂室') and (d.撤档时间 is null or to_char(d.撤档时间,'yyyy-mm-dd')='3000-01-01')"
    Else
        gstrSql = gstrSql & " where ID in (select distinct 部门id from 部门性质说明 where 工作性质 like '中药%' or 工作性质='制剂室') and (d.撤档时间 is null or to_char(d.撤档时间,'yyyy-mm-dd')='3000-01-01')"
    End If
    If InStr(1, strPrivs, "允许设置所有库房限额盘点") = 0 Then
        gstrSql = gstrSql & "      and ID in (select 部门ID from 部门人员 R where R.人员ID=[1])"
    End If
    
    Set mrsNormal = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
        
    With mrsNormal
        Me.cboRoom.Clear
        Do While Not .EOF
            Me.cboRoom.AddItem !编码 & "-" & !名称
            Me.cboRoom.ItemData(Me.cboRoom.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    If Me.cboRoom.ListCount <= 0 Then
        MsgBox "未设置" & IIf(Me.Tag = "5", "西成药", IIf(Me.Tag = "6", "中成药", "中草药")) & "库房，无法设置储备限量", vbExclamation, gstrSysName
        Unload Me: Exit Sub
    End If
    Me.cboRoom.ListIndex = 0
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    
    mblnActive = True
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboRoom_Click()
    err = 0: On Error GoTo ErrHand
    
    If mlng库房ID = cboRoom.ItemData(cboRoom.ListIndex) Then Exit Sub
    mlng库房ID = cboRoom.ItemData(cboRoom.ListIndex)
    cboRoom.Tag = GetDrugUnit(mlng库房ID)
    cboDrugUnit.Text = cboRoom.Tag
    mlngFind = 0
    mstr分类 = "所有"
    mstr分类ID = ""
    mstr剂型 = ""
    Call zlLimitRef
    cboDrugUnit.Tag = cboDrugUnit.ListIndex
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'取药品单位名称
Public Function GetDrugUnit(ByVal lng库房ID As Long) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim intUnit As Integer, strUnit As String
    Dim bln缺省 As Boolean
    Dim lngModul As Long
    
    On Error GoTo ErrHand
    
    gstrSql = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "读取药品单位", lng库房ID)

    '取服务对象及部门性质
    With rsProperty
        Do While Not .EOF
            strobjTemp = strobjTemp & .Fields(0)
            strWorkTemp = strWorkTemp & .Fields(1)
            .MoveNext
        Loop
        .Close
    End With
    If InStr(strWorkTemp, "药库") <> 0 Then
        '药库单位
        intUnit = 1
        strUnit = 4
    ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
        '门诊单位
        intUnit = 2
        strUnit = 2
    ElseIf InStr(strobjTemp, "2") <> 0 Then
        '住院单位
        intUnit = 3
        strUnit = 3
    Else
        '售价单位：主要是制剂室
        intUnit = 4
        strUnit = 1
    End If
    
    '取该药房缺省该使用的单位
    GetDrugUnit = GetSpecUnit(lng库房ID, intUnit)
        
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "售价单位"
End Function

'返回指定库房指定适用范围的单位
Public Function GetSpecUnit(ByVal lng库房ID As Long, ByVal int范围 As Integer) As String
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select Nvl(性质,1) AS 单位 From 药品库房单位 Where 库房ID=[1] And 适用范围=[2]"
    Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "提取单位", lng库房ID, int范围)

    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!单位
    Else
'        MsgBox "该库房未设置库房单位，根据部门性质以及服务对象取缺省单位！" & _
'            vbCrLf & "缺省单位的规则：" & _
'            vbCrLf & "  服务对象是住院或门诊和住院的，取住院单位" & _
'            vbCrLf & "  仅服务于门诊的，取门诊单位" & _
'            vbCrLf & "  具有药库属性的，取药库单位" & _
'            vbCrLf & "  其他取售价单位", vbInformation, gstrSysName
        
        gstrSql = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        Set rsProperty = zldatabase.OpenSQLRecord(gstrSql, "读取药品单位", lng库房ID)

        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '住院单位
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '门诊单位
            strUnit = 2
        ElseIf InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            strUnit = 4
        Else
            '售价单位：主要是制剂室
            strUnit = 1
        End If
    End If
    
    '转换为真实的单位返回给调用者
    GetSpecUnit = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub zlLimitRef()
    '--------------------------------------------------------
    '功能：刷新库存限额
    '--------------------------------------------------------
    Dim rsFind As ADODB.Recordset
    Dim lngRow As Long
    Dim int前缀 As Integer
    Dim int后缀 As Integer
    Dim str编码前缀 As String
    Dim str编码后缀 As String
    Dim blnLine As Boolean
    Dim lngCount As Long
    Dim strRule As String
    Dim lngId As Long
    
    err = 0: On Error GoTo ErrHand
    
    If mstr剂型 = "" And mstr分类ID = "" Then
        strRule = ""
    Else '这种情况下面会使用自定义的oracle函数
        strRule = "/*+ RULE*/"
    End If
    
    gstrSql = "Select " & strRule & " I.ID,I.编码,I.名称, i.商品名,I.规格,I.产地,I.原产地," & _
         Switch(Me.cboRoom.Tag = "售价单位", "I.计算单位 as 单位,1 as 包装,", _
                Me.cboRoom.Tag = "药库单位", "I.药库单位 as 单位,nvl(I.药库包装,1) as 包装,", _
                Me.cboRoom.Tag = "门诊单位", "I.门诊单位 as 单位,nvl(I.门诊包装,1) as 包装,", _
                Me.cboRoom.Tag = "住院单位", "I.住院单位 as 单位,nvl(I.住院包装,1) as 包装,") & _
            "I.计算单位 as 售价单位, 1 as 售价包装," & _
            "I.住院单位, nvl(I.住院包装,1) as 住院包装," & _
            "I.门诊单位, nvl(I.门诊包装,1) as 门诊包装," & _
            "I.药库单位, nvl(I.药库包装,1) as 药库包装," & _
            "   nvl(L.上限,0) as 上限,nvl(L.下限,0) as 下限,L.盘点属性,L.库房货位,l.领用标志,K.实际数量," & _
            "   Decode(I.是否变价, 0, P.现价, Decode(Sign(K.实际数量 - 1), -1, 0, K.实际金额 / K.实际数量)) As 零售价 " & _
            " From (Select  I.ID,I.编码,I.名称, b.名称 As 商品名,I.规格,I.产地,S.原产地,I.计算单位,S.门诊单位,S.门诊包装, " & _
            "           S.住院单位,S.住院包装,S.药库单位,S.药库包装, I.是否变价, S.药名id " & _
            "       From 收费项目目录 I, 收费项目别名 B,药品规格 S," & _
            "            (Select Distinct 诊疗项目id From 诊疗执行科室 Where 执行科室id=[1]) E,(select distinct 收费细目id from 收费执行科室 where 执行科室id=[1]) F " & _
            "       Where i.Id = b.收费细目id(+) And b.性质(+) = 3 And I.Id=S.药品id And S.药名id=E.诊疗项目id and I.类别=[2] And i.id=f.收费细目id " & _
            "            and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))) I," & _
            "      (Select 药品id,上限,下限,盘点属性,库房货位,领用标志 From 药品储备限额 L Where 库房id=[1]) L," & _
            " (Select 药品id, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额 From 药品库存 " & _
            "  Where 性质 = 1 And 库房id = [1] Group By 药品id) K, 收费价目 P "
            
    If mstr分类ID <> "" Then
        gstrSql = gstrSql & ", 诊疗项目目录 Z, 诊疗分类目录 M, Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) G "
    End If
    
    If mstr剂型 <> "" Then
        gstrSql = gstrSql & ", 药品特性 T, Table(Cast(f_Str2list([4]) As zlTools.t_strlist)) H "
    End If
    
    gstrSql = gstrSql & " Where I.ID = P.收费细目id And I.Id=L.药品id(+) And I.ID = K.药品id(+) And (p.终止日期 Is Null Or Sysdate Between p.执行日期 And Nvl(p.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
            GetPriceClassString("P")
    
    If mstr分类ID <> "" Then
        gstrSql = gstrSql & " And I.药名id = Z.ID And Z.分类id = M.ID And M.ID = G.Column_Value "
    End If
    
    If mstr剂型 <> "" Then
        gstrSql = gstrSql & " And I.药名id = T.药名id And T.药品剂型 = H.Column_Value "
    End If
    
    gstrSql = gstrSql & " Order By I.编码"
    
    Set mrsNormal = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex), Me.Tag, mstr分类ID, mstr剂型)
    
    If Not mrsNormal.EOF Then lngCount = mrsNormal.RecordCount
    With mrsNormal
        Me.vsfLimit.Rows = 1
        Me.vsfLimit.Redraw = False
        Call IniGrid
        Do While Not .EOF
            If lngId <> mrsNormal!ID Then
                lngId = mrsNormal!ID
                Me.vsfLimit.Rows = vsfLimit.Rows + 1
                Me.vsfLimit.RowData(vsfLimit.Rows - 1) = Val(!ID)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("编码")) = !编码
            
                If mint药品名称显示 = 0 Then
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("名称")) = !名称
                ElseIf mint药品名称显示 = 1 Then
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("名称")) = IIf(IsNull(!商品名), !名称, !商品名)
                Else
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("名称")) = !名称
                    Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("商品名")) = IIf(IsNull(!商品名), "", !商品名)
                End If
                
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("规格")) = IIf(IsNull(!规格), "", !规格)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("产地")) = IIf(IsNull(!产地), "", !产地)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("原产地")) = IIf(IsNull(!原产地), "", !原产地)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("单位")) = IIf(IsNull(!单位), "", !单位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("包装")) = !包装
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("零售价")) = IIf(!零售价 = 0, "", Format(!零售价 * !包装, "0.000"))
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("库存数量")) = Format(!实际数量 / !包装, "0.00")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("上限")) = Format(!上限 / !包装, "0.00000")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("下限")) = Format(!下限 / !包装, "0.00000")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("日盘")) = IIf(Mid(!盘点属性, 1, 1) = "1", "√", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("周盘")) = IIf(Mid(!盘点属性, 2, 1) = "1", "√", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("月盘")) = IIf(Mid(!盘点属性, 3, 1) = "1", "√", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("季盘")) = IIf(Mid(!盘点属性, 4, 1) = "1", "√", "")
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("货位")) = IIf(IsNull(!库房货位), "", !库房货位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("允许领用")) = IIf(IsNull(!领用标志), "√", IIf(!领用标志 = 0, "", "√"))
                
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("售价单位")) = IIf(IsNull(!售价单位), "", !售价单位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("住院单位")) = IIf(IsNull(!住院单位), "", !住院单位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("住院包装")) = IIf(IsNull(!住院包装), "", !住院包装)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("门诊单位")) = IIf(IsNull(!门诊单位), "", !门诊单位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("门诊包装")) = IIf(IsNull(!门诊包装), "", !门诊包装)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("药库单位")) = IIf(IsNull(!药库单位), "", !药库单位)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("药库包装")) = IIf(IsNull(!药库包装), "", !药库包装)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("固定零售价")) = IIf(IsNull(!零售价), 0, !零售价)
                Me.vsfLimit.TextMatrix(vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("实际数量")) = IIf(IsNull(!实际数量), 0, !实际数量)
                
                If InStr(!编码, "-") > 0 Then
                    blnLine = True
                    If Len(Mid(!编码, 1, InStr(!编码, "-") - 1)) > int前缀 Then
                        int前缀 = Len(Mid(!编码, 1, InStr(!编码, "-") - 1))
                    End If
                    
                    If Len(Mid(!编码, InStr(!编码, "-") + 1)) > int后缀 Then
                        int后缀 = Len(Mid(!编码, InStr(!编码, "-") + 1))
                    End If
                Else
                    If Len(!编码) > int前缀 Then
                        int前缀 = Len(!编码)
                    End If
                End If
                
                If vsfLimit.Rows - 1 Mod IIf(.RecordCount > 20, .RecordCount \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "正在提取：" & String(vsfLimit.Rows - 1 \ IIf(.RecordCount > 20, .RecordCount \ 20, 1), "…")
                End If
            End If
            .MoveNext
        Loop
        
        For lngRow = 1 To Me.vsfLimit.Rows - 1
            If blnLine = False Then
                Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("排序编码")) = Format(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), String(int前缀, "0"))
            Else
                If InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), "-") > 0 Then
                    str编码前缀 = Mid(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), 1, InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), "-") - 1)
                    str编码后缀 = Mid(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), InStr(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), "-") + 1)
                    
                    str编码前缀 = Format(str编码前缀, String(int前缀, "0"))
                    str编码后缀 = Format(str编码后缀, String(int后缀, "0"))
                Else
                    str编码前缀 = Format(Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("编码")), String(int前缀, "0"))
                    str编码后缀 = String(int后缀, "0")
                End If
                
                Me.vsfLimit.TextMatrix(lngRow, Me.vsfLimit.ColIndex("排序编码")) = str编码前缀 & "-" & str编码后缀
            End If
        Next
        
        Me.vsfLimit.Col = Me.vsfLimit.ColIndex("排序编码")
        Me.vsfLimit.Sort = flexSortStringAscending
        
        If Me.vsfLimit.Rows > 1 Then
            Me.vsfLimit.Cell(flexcpBackColor, 1, Me.vsfLimit.ColIndex("编码"), Me.vsfLimit.Rows - 1, Me.vsfLimit.ColIndex("库存数量")) = &HEFEFEF
        End If
        
        Me.vsfLimit.Redraw = True
    End With
    Me.stbThis.Panels(2).Text = "共：" & lngCount & "种药品规格" & " " & " 当前分类：" & IIf(mstr分类 = "", "所有", mstr分类) & "  当前剂型：" & IIf(mstr剂型 = "", "所有", mstr剂型)
    mblnChanged = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsHaveStock(ByVal strStockName As String) As Boolean
    Dim rs As ADODB.Recordset
    On Error GoTo errHandle
    gstrSql = "Select 编码 From 药品库房货位 where 名称=[1]"
    Set rs = zldatabase.OpenSQLRecord(gstrSql, "判断是否存在药品库房货位", strStockName)
        
    IsHaveStock = (rs.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txt查找.Visible And KeyCode = vbKeyF3 Then
        Call txt查找_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub Form_Load()
    mint药品名称显示 = Val(zldatabase.GetPara("药品名称显示", , , 2))
    Call RestoreWinState(Me, App.ProductName)
    mlng库房ID = 0
    cboDrugUnit.Tag = "-1"
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.fraLine.Left = 0: Me.fraLine.Width = Me.ScaleWidth + 100
    Me.vsfLimit.Left = 0: Me.vsfLimit.Width = Me.ScaleWidth
    Me.vsfLimit.Height = Me.ScaleHeight - Me.vsfLimit.Top - Me.fraFunc.Height - Me.stbThis.Height
    Me.fraFunc.Left = 0: Me.fraFunc.Width = Me.ScaleWidth: Me.fraFunc.Top = Me.vsfLimit.Top + Me.vsfLimit.Height
    Me.cmdClose.Left = Me.fraFunc.Width - Me.cmdClose.Width - 90
    Me.cmdSave.Left = Me.cmdClose.Left - Me.cmdSave.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    mblnActive = False
End Sub




Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub





Private Sub txt查找_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txt查找.Text))
    If strInput = "" Then Exit Sub
    
    Call FindGridRow(strInput)
End Sub

Private Sub txt查找_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vsfLimit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfLimit
        Select Case Col
            Case .ColIndex("货位")
                .ColComboList(.ColIndex("货位")) = "..."
        End Select
    End With
End Sub


Private Sub vsfLimit_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfLimit
        If Col = .ColIndex("编码") Then
            .Col = .ColIndex("排序编码")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfLimit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfLimit
        Select Case Col
            Case .ColIndex("货位")
                mlngRow = Row
                If Select货位("") = False Then
                    Exit Sub
                End If
            Case Else
        End Select
    End With
End Sub

Private Sub vsfLimit_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChanged = True
End Sub

Private Sub vsfLimit_DblClick()
    Dim blnNext As Boolean
    
    With vsfLimit
        If .Row < 1 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .Col = .ColIndex("允许领用") Then
            If .TextMatrix(.Row, .Col) = "√" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
        End If
        If .Col = .ColIndex("日盘") Or .Col = .ColIndex("周盘") Or .Col = .ColIndex("月盘") Or .Col = .ColIndex("季盘") Then
            If InStr(1, strPrivs, "盘点属性设置") = 0 Then Exit Sub
            blnNext = True
        End If
        
        If blnNext = True Then
            If .TextMatrix(.Row, .Col) = "√" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
        End If
    End With
End Sub

Private Sub vsfLimit_EnterCell()
    Dim intRow As Integer
    
    With vsfLimit
        If .Row < 1 Then Exit Sub
        .FocusRect = flexFocusLight
        .Editable = flexEDNone
        Select Case .Col
            Case .ColIndex("日盘"), .ColIndex("周盘"), .ColIndex("月盘"), .ColIndex("季盘"), .ColIndex("允许领用")
                .FocusRect = flexFocusSolid
            Case .ColIndex("上限"), .ColIndex("下限")
                If InStr(1, strPrivs, "上下限控制") > 0 Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                End If
            Case .ColIndex("货位")
                If InStr(1, strPrivs, "调整库位") > 0 Then
                    .Editable = flexEDKbdMouse
                    .FocusRect = flexFocusSolid
                End If
        End Select
        
        '设置行选中边框
        If .Rows <> 1 Then
            For intRow = 0 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("货位"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("编码"), .Row, .ColIndex("编码"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("货位"), .Row, .ColIndex("货位"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub


Private Sub vsfLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfStore.Visible = False
    End If
    If vsfLimit.Col = vsfLimit.ColIndex("货位") Then
        If KeyCode <> vbKeyReturn Then
            vsfLimit.ColComboList(vsfLimit.ColIndex("货位")) = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.Col) = ""
        End If
    End If
    
    If txt查找.Visible And KeyCode = vbKeyF3 Then
        Call txt查找_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub vsfLimit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfLimit
        If Trim(.EditText) = "" Then Exit Sub
        
        If Col = .ColIndex("货位") Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
'                MsgBox "货位超长！最多50个字母或25个汉字", vbInformation, gstrSysName
'                vsfLimit.TextMatrix(Row, Col) = ""
'                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                Exit Sub
            End If
        ElseIf Col = .ColIndex("上限") Or Col = .ColIndex("下限") Then
            If Not IsNumeric(.EditText) Then
                KeyCode = 0
                Exit Sub
            End If
            If Val(.EditText) < 0 Then
                KeyCode = 0
                Exit Sub
            End If
            If Val(.EditText) > 10000000000000# Then
                KeyCode = 0
                Exit Sub
            End If
        End If

       Select Case Col
            Case .ColIndex("上限")
                .EditText = Format(.EditText, "0.00000"): .TextMatrix(Row, .ColIndex("上限")) = .EditText
            Case .ColIndex("下限")
                .EditText = Format(.EditText, "0.00000"): .TextMatrix(Row, .ColIndex("下限")) = .EditText
            Case .ColIndex("货位")
                If Select货位(.EditText) = False Then
                    vsfLimit.TextMatrix(Row, Col) = vsfLimit.EditText
                    vsfLimit.Cell(flexcpForeColor, Row, Col) = vbRed
'                    If MsgBox("没有找到该货位，是否增加该货位？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        vsfLimit.TextMatrix(Row, Col) = ""
'                    End If
                    Exit Sub
                End If
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
        End Select
    End With
End Sub

Private Function Select货位(ByVal strKey As String) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strID As String
    Dim str名称 As String
    Dim objNode As Node
    Dim str货位 As String
    
    err = 0: On Error GoTo ErrHand:
    
    strKey = UCase(strKey)

    If strKey <> "" Then
        gstrSql = " Select id,编码,名称,简码 From 药品库房货位 " & _
                " Where 库房id=[1] And(编码 Like [2] Or 名称 Like [3] Or 简码 Like [3]) Order By 编码"
    Else
        gstrSql = " Select id,编码,名称,简码 From 药品库房货位 " & _
                " Where 库房id=[1]"
    End If
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "货位", mlng库房ID, strKey & "%", gstrMatch & strKey & "%")
    
    If rsTemp.EOF Then
        vsfLimit.EditText = strKey
        Exit Function
    End If
    
    str货位 = vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("货位"))
    vsfStore.Rows = 1
    Do While Not rsTemp.EOF
        With vsfStore
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTemp!编码
            .TextMatrix(.Rows - 1, .ColIndex("货位")) = rsTemp!名称
            
            If str货位 <> "" Then
                If InStr(1, "," & str货位 & ",", "," & rsTemp!名称 & ",") > 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("选择")) = 1
                End If
            End If
        End With
        rsTemp.MoveNext
    Loop
    If rsTemp.RecordCount > 0 Then
        vsfStore.Move vsfLimit.CellLeft + 30, vsfLimit.CellTop + vsfStore.Height - 200
        vsfStore.Visible = True
        If strKey = "" Then
            vsfStore.SetFocus
        End If
        vsfStore.Row = 1
    End If
    
    Select货位 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfLimit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    Select Case Col
        Case vsfLimit.ColIndex("上限"), vsfLimit.ColIndex("下限")
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            ElseIf KeyAscii = Asc(".") Then
                If InStr(vsfLimit.EditText, ".") <> 0 Then     '只能存在一个小数点
                    KeyAscii = 0
                End If
            End If
    End Select
    
End Sub

Private Sub vsfLimit_RowColChange()
    With vsfLimit()
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfLimit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    With vsfLimit
        If Trim(.EditText) = "" Then Exit Sub
        
        If Col = .ColIndex("货位") Then
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                MsgBox "货位超长！最多50个字母或25个汉字", vbInformation, gstrSysName
                vsfLimit.TextMatrix(Row, Col) = ""
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                Exit Sub
            End If
        End If

       Select Case Col
            Case .ColIndex("货位")
                If Select货位(.EditText) = False Then
                    vsfLimit.TextMatrix(Row, Col) = vsfLimit.EditText
                    vsfLimit.Cell(flexcpForeColor, Row, Col) = vbRed
                    If MsgBox("没有找到该货位，是否增加该货位？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        vsfLimit.TextMatrix(Row, Col) = ""
                        vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
                    Else
                        gstrSql = "Zl_药品库房货位_Insert(Null, '" & Trim(vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("货位"))) & "', Null, " & Me.cboRoom.ItemData(Me.cboRoom.ListIndex) & ", Null)"
                        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
                    End If
                    Exit Sub
                End If
                vsfLimit.EditText = vsfLimit.TextMatrix(Row, Col)
        End Select
    End With
End Sub

Private Sub vsfStore_Click()
    With vsfStore
        If .Col = .ColIndex("选择") Then
            If Val(.TextMatrix(.Row, .ColIndex("选择"))) = 1 Then
                .TextMatrix(.Row, .ColIndex("选择")) = ""
            Else
                .TextMatrix(.Row, .ColIndex("选择")) = "1"
            End If
        End If
    End With
End Sub


Private Sub vsfStore_DblClick()
    Dim i As Integer
    Dim str货位 As String
    
    With vsfStore
        If .Rows <= 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) = 1 Then
                str货位 = str货位 & "," & .TextMatrix(i, .ColIndex("货位"))
            End If
        Next
        vsfStore.Visible = False
    End With
    
    If str货位 <> "" Then
        str货位 = Mid(str货位, 2)
    End If
    With vsfLimit
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("货位")) = str货位
        vsfLimit.Cell(flexcpForeColor, .Row, .ColIndex("货位"), .Row, .ColIndex("货位")) = vbBlack
        .Redraw = flexRDBuffered
    End With
End Sub


Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfStore.Visible = False
    End If
End Sub


Private Sub vsfStore_LostFocus()
    If vsfStore.Visible = True Then
        vsfStore.Visible = False
    End If
End Sub


