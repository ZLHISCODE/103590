VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗项目选择"
   ClientHeight    =   7695
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "frmClinicSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   12270
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7065
      Width           =   12270
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   9810
         TabIndex        =   10
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11055
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   870
         TabIndex        =   8
         Top             =   135
         Width           =   1845
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位(&L)"
         Height          =   350
         Left            =   2700
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "项目查找"
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "请输入查找条件"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   3975
         TabIndex        =   11
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdBound 
      Caption         =   "批量绑定"
      Height          =   350
      Left            =   4350
      TabIndex        =   5
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除所有"
      Height          =   350
      Left            =   11115
      TabIndex        =   4
      Top             =   3855
      Width           =   1100
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   3630
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5685
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   240
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   12515
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgItem 
      Height          =   3750
      Left            =   3675
      TabIndex        =   1
      Top             =   30
      Width           =   8550
      _cx             =   15081
      _cy             =   6615
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   930
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":6852
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":68B0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgOften 
      Left            =   0
      Top             =   645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":690E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7008
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7702
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":7DFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmClinicSelect.frx":84F6
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":8A90
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":902A
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":95C4
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":9B5E
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":A0F8
            Key             =   "方案"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgBound 
      Height          =   2790
      Left            =   3675
      TabIndex        =   3
      Top             =   4320
      Width           =   8550
      _cx             =   15081
      _cy             =   4921
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      ExplorerBar     =   3
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
Attribute VB_Name = "frmClinicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnDown As Boolean
Private mstrIDs As String
Private mstrNAMEs As String
Private mstrPreNode As String
Private mstrLike As String
Private mrsItem As New ADODB.Recordset
Private mrsFind As New ADODB.Recordset

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As node
    
    On Error GoTo errH
    
    strSQL = _
        " Select 0 as 级,类型,-类型 as ID,-Null as 上级ID,类型||'' as 编码," & _
        " 类型||'.'||Decode(类型,1,'西成药',2,'中成药',3,'中草药',4,'中药配方',5,'诊疗项目',6,'成套诊疗','7','卫生材料') as 名称" & _
        " From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Group by 类型"
    strSQL = strSQL & " Union ALL " & _
        " Select Level as 级,类型,ID,Nvl(上级ID,-类型) as 上级ID,编码,名称 From 诊疗分类目录" & _
        " Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Order by 级,编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FillList() As Boolean
'功能：根据当前界面条件装入诊疗项目目录
    Dim objNode As node, objItem As ListItem
    Dim strSQL As String, strInside As String
    Dim arrClass As Variant, strclass As String
    Dim strSub As String, str操作类型 As String
    Dim str性别 As String, strStock As String
    Dim strInput As String, lng药房ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str范围 As String, str药品 As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str库存限制 As String, strPriv As String
    Dim i As Long, j As Long
    Dim strCommIF As String, strScope As String
    
    Dim lng分类ID As Long, int类型 As Integer, str类别 As String

    Set objNode = tvw_s.SelectedItem '可能为Nothing
    
    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vfgItem.Rows = vfgItem.FixedRows
    vfgItem.Rows = vfgItem.FixedRows + 1
    Me.Refresh
    
    '读取数据
    int类型 = Val(objNode.Tag)
    lng分类ID = Val(Mid(objNode.Key, 2))
    If Val(Mid(objNode.Key, 2)) < 0 Then
        strSub = " And A.分类ID IN(" & _
            " Select ID From 诊疗分类目录 Where 类型=[1] And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " )"
    Else
        strSub = " And A.分类ID IN(" & _
            " Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[2] Connect by Prior ID=上级ID)"
    End If
    
    '按品种下达的长嘱
    blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID," & _
                " F.名称 As 类别,Null as 基本,A.编码,A.名称,Null as 商品名," & _
                "A.计算单位,Null as 规格,Null as 产地, D.药品剂型," & _
                "Null as 费用类型,Null as 说明,D.处方职务 as 处方职务ID" & _
            " From 药品特性 D,诊疗项目类别 F,诊疗项目目录 A" & _
            " Where A.ID=D.药名ID And A.类别=F.编码 And A.类别 IN ('5','6','7')" & strCommIF & strSub
    End If
        
    '2.非药品卫材的诊疗项目部份:分类不是药品类型时不必读取
    '--------------------------------------------------------------------------------------
    blnLoad = InStr(",1,2,3,7,", Val(objNode.Tag)) = 0
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select " & _
                " A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID,D.名称 As 类别,Null as 基本," & _
                " A.编码,A.名称,Null as 商品名,A.计算单位,A.标本部位 as 规格,Null as 产地," & _
                " Null as 药品剂型,Null as 费用类型,Null as 说明,Null as 处方职务ID" & _
            " From 诊疗项目类别 D,诊疗项目目录 A" & _
            " Where A.类别=D.编码 And A.类别 Not IN('4','5','6','7')" & strCommIF & strSub
    End If
    
    blnLoad = Val(objNode.Tag) = 7
    If blnLoad Then
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                " F.名称 AS 类别,Null as 基本,A.编码,A.名称 as 名称,Null as 商品名,A.计算单位,A.规格,A.产地," & _
                " Null as 药品剂型,Null as 项目特性,A.费用类型,A.说明,Null as 处方职务ID" & _
            " From 收费项目目录 A,材料特性 C,诊疗项目目录 E,收费项目类别 F" & _
            " Where A.ID=C.材料ID And C.诊疗ID=E.ID And A.类别=F.编码 And E.类别='4' And C.核算材料=0" & _
                " And A.类别='4'" & strCommIF & strSub & _
                " And (E.服务对象 IN(2,3)) " & _
                " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)"
    End If
    
    strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by Decode(类别ID,'4','Z',类别ID),类别,编码"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int类型, lng分类ID)
    
    '绑定数据
    '--------------------------------------------------------------------------
    vfgItem.Redraw = flexRDNone
    
    vfgItem.Rows = 2
    vfgItem.FixedRows = 1
    vfgItem.Cols = 2
    vfgItem.FixedCols = 1
    vfgItem.RowData(1) = 0
    
    vfgItem.ScrollBars = flexScrollBarNone
    Set vfgItem.DataSource = mrsItem
    vfgItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vfgItem.Rows = vfgItem.FixedRows Then
        vfgItem.Rows = vfgItem.FixedRows + 1
    End If
    
    '列属性调整
    vfgItem.ColAlignment(0) = 4
    vfgItem.Cell(flexcpAlignment, 0, 0, 0, vfgItem.Cols - 1) = 4
    vfgItem.RowHeight(0) = vfgItem.RowHeightMin
    
    '卡片相关数据计算
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vfgItem.TextMatrix(i, 0) = i
        vfgItem.RowHeight(i) = vfgItem.RowHeightMin
        vfgItem.RowData(i) = Val(mrsItem!诊疗项目ID)
        mrsItem.MoveNext
    Next
    
    '根据结果数据类别等情况隐藏一些不必要的列
    For i = 1 To vfgItem.Cols - 1
        If InStr(1, ",KEYID,类别ID,收费细目ID,基本,处方职务ID,", "," & vfgItem.TextMatrix(0, i) & ",") <> 0 Then vfgItem.ColHidden(i) = True
    Next
    
    '行号列宽度
    vfgItem.ColWidth(0) = Me.TextWidth(vfgItem.TextMatrix(vfgItem.Rows - 1, 0) & " ")
    If vfgItem.ColWidth(0) < 380 Then vfgItem.ColWidth(0) = 380
    
    vfgItem.FrozenCols = 0
    vfgItem.Editable = flexEDNone
    vfgItem.SheetBorder = vfgItem.BackColor
    
    vfgItem.Row = vfgItem.FixedRows: vfgItem.Col = vfgItem.FixedCols
    vfgItem.Redraw = flexRDDirect
        
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(ByVal frmParent As Form, strIDs As String, strNames As String) As Boolean
    mblnOK = False
    mstrIDs = strIDs
    mstrNAMEs = strNames
    Me.Show 1, frmParent
    ShowMe = mblnOK
    strIDs = mstrIDs
    strNames = mstrNAMEs
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdBound_Click()
    Dim node As MSComctlLib.node
    Dim strSel As String, strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim blnAdd As Boolean
    
    On Error GoTo ErrHand
    Set node = tvw_s.Nodes(1)
    strSel = GetSelNodes(node)
    
    If strSel <> "" Then
        strSQL = " Select /*+ Rule */ a.Id, b.名称 类别,a.编码,a.名称" & vbNewLine & _
                " From 诊疗项目目录 A,诊疗项目类别 B,(Select Column_Value From Table(f_Num2list([1]))) C" & vbNewLine & _
                " Where  a.类别 = b.编码 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.分类Id = C.Column_Value" & vbNewLine & _
                " Order By a.类别, a.编码"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取诊疗项目", strSel)
        Do While Not rsTemp.EOF
            blnAdd = True
            For lngRow = vfgBound.FixedRows To vfgBound.Rows - 1
                If Val(vfgBound.TextMatrix(lngRow, 1)) = Val(rsTemp!ID) Then
                    blnAdd = False
                    Exit For
                End If
            Next lngRow
            If blnAdd = True Then
                If Val(vfgBound.TextMatrix(vfgBound.Rows - 1, 1)) > 0 Or vfgBound.Rows = vfgBound.FixedRows Then
                    vfgBound.Rows = vfgBound.Rows + 1
                End If
                vfgBound.TextMatrix(vfgBound.Rows - 1, 0) = vfgBound.Rows - vfgBound.FixedRows
                vfgBound.TextMatrix(vfgBound.Rows - 1, 1) = Val(rsTemp!ID)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 2) = CStr(Nvl(rsTemp!类别))
                vfgBound.TextMatrix(vfgBound.Rows - 1, 3) = CStr(Nvl(rsTemp!编码))
                vfgBound.TextMatrix(vfgBound.Rows - 1, 4) = CStr(Nvl(rsTemp!名称))
                vfgBound.RowData(vfgBound.Rows - 1) = Val(rsTemp!ID)
                vfgBound.Row = vfgBound.Rows - 1
                vfgBound.TopRow = vfgBound.Rows - 1
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
            rsTemp.MoveNext
        Loop
        vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
    Else
        MsgBox "您至少要勾选一个要添加的节点，请在左侧列表中进行勾选！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdClear_Click()
    '清除所有绑定的项目
    Call FillBoundItem("")
End Sub

Private Sub cmdFind_Click()
'功能:词句查找
    Dim strText As String, strMatch As String
    Dim strFind As String, strSQL As String
    Dim lngRow As Long, lngTypeID As Long
    
    On Error GoTo ErrHand
    
    If mrsFind.State = adStateOpen Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocaItem
        Exit Sub
    End If
    
    If Trim(txtFind.Text) = "" Then
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtFind.Text, "'") > 0 Then
        MsgBox "输入的内容包含非法字符 ' ,请检查!", vbInformation, gstrSysName
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If Not tvw_s.SelectedItem Is Nothing Then
        lngTypeID = Val(Mid(tvw_s.SelectedItem.Key, 2))
    Else
        lngTypeID = 0
    End If
    
    strText = mstrLike & txtFind.Text & "%"
    If ZLCommFun.IsCharChinese(txtFind.Text) Then
        strFind = " And A.名称 Like '" & strText & "'"
    ElseIf IsNumeric(txtFind.Text) Then
        strFind = " And A.编码 Like '" & strText & "'"
    Else
        strFind = " And zlspellcode(A.名称) Like '" & UCase(strText) & "'"
    End If
    
    '根据输入的内容提取匹配的词句
    strSQL = " Select a.分类id, a.Id,b.上级ID" & vbNewLine & _
            " From 诊疗项目目录 a, 诊疗分类目录 b" & vbNewLine & _
            " Where a.分类id = b.Id And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "      (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & strFind & _
            " Order by" & IIf(lngTypeID = 0, "", " DECODE(A.分类ID," & lngTypeID & ",0,1),") & " b.类型,b.编码,a.编码"
    Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, "项目查找")

    Call LocaItem
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocaItem()
    Dim lngRow As Long
    
    If mrsFind.RecordCount = 0 Then
        lblInfo.Caption = "没有找到符合条件的信息"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    
    If mrsFind.EOF = True Then
        lblInfo.Caption = "已经完成所有定位，请重新输入条件"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    lblInfo.Caption = "共找到" & mrsFind.RecordCount & "条,当前是第" & mrsFind.AbsolutePosition & "条"
    lblInfo.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "下一个(&L)"
        Else
            cmdFind.Caption = "定位(&L)"
            lblInfo.Caption = "已经是最后一条，请重新输入条件"
        End If
    End If
    On Error Resume Next
    err.Clear: err = 0
    '开始进行定位
    tvw_s.Nodes("_" & mrsFind!分类id).Selected = True
    If err <> 0 Then err.Clear: Exit Sub
    Call tvw_s_NodeClick(tvw_s.Nodes("_" & mrsFind!分类id))
    
    For lngRow = vfgItem.FixedRows To vfgItem.Rows - 1
        If Val(vfgItem.RowData(lngRow)) = Val(mrsFind!ID) Then
            vfgItem.Row = lngRow
            vfgItem.TopRow = lngRow
            Exit For
        End If
    Next lngRow
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim strIDs As String, strNames As String
    With vfgBound
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) > 0 Then
                strIDs = strIDs & "," & Val(.TextMatrix(lngRow, 1))
                strNames = strNames & "," & .TextMatrix(lngRow, 4)
            End If
        Next
    End With
    
    mstrIDs = Mid(strIDs, 2)
    mstrNAMEs = Mid(strNames, 2)
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mstrPreNode = ""
    mblnDown = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    Call FillTree
    Call FillBoundItem(mstrIDs)
End Sub

Private Sub FillBoundItem(ByVal strIDs As String)
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    strSQL = "Select /*+ Rule */ a.Id, b.名称 类别, a.编码, a.名称" & vbNewLine & _
            " From 诊疗项目目录 a, 诊疗项目类别 b, (Select Column_Value From Table(f_Num2list([1]))) c" & vbNewLine & _
            " Where a.类别 = b.编码 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And a.Id = c.Column_Value" & vbNewLine & _
            " Order By a.类别, a.编码"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取已经绑定诊疗项目", strIDs)
    
    With vfgBound
        .Redraw = flexRDNone
        
        '可能统计常用项目时设置为了0行0列
        .Rows = 2
        .FixedRows = 1
        .Cols = 5
        .FixedCols = 1
        .RowData(1) = 0
        .ScrollBars = flexScrollBarNone
        Set .DataSource = rsTemp
        .ScrollBars = flexScrollBarBoth
        .ColHidden(1) = True
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
        End If
        
        '列属性调整
        .ColAlignment(0) = 4
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .RowHeight(0) = .RowHeightMin
        
        '卡片相关数据计算
        '------------------------------------------------------------------------
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, 0) = i
            .RowHeight(i) = .RowHeightMin
            .RowData(i) = Val(rsTemp!ID)
             rsTemp.MoveNext
        Next
        
        '行号列宽度
        .ColWidth(0) = Me.TextWidth(.TextMatrix(.Rows - 1, 0) & " ")
        If .ColWidth(0) < 380 Then .ColWidth(0) = 380
        .ColWidth(2) = 800
        .ColWidth(3) = 1000
        .ColWidth(4) = 2000
        
        .FrozenCols = 0
        .Editable = flexEDNone
        .SheetBorder = .BackColor
        
        .Row = .FixedRows: .Col = .FixedCols
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With picSplit
        .Top = 0
        .Height = Me.ScaleHeight - picBottom.Height
    End With
    With tvw_s
        .Height = picSplit.Height
        .Width = picSplit.Left
    End With
    With vfgItem
        .Left = picSplit.Left + picSplit.Width
        .Width = Me.ScaleWidth - .Left
        .Height = picSplit.Height - vfgBound.Height - cmdClear.Height - 200
    End With
    
    With cmdClear
        .Top = vfgItem.Height + 100
        .Left = Me.ScaleWidth - .Width
    End With
    
    With cmdBound
        .Top = cmdClear.Top
        .Left = vfgItem.Left
    End With
    
    With vfgBound
        .Top = cmdClear.Top + cmdClear.Height + 100
        .Left = vfgItem.Left
        .Width = vfgItem.Width
    End With
    cmdCancel.Left = picBottom.Width - cmdCancel.Width - 150
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
    Set mrsFind = Nothing
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = (Button = 1)
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnDown Then Exit Sub
    Dim blnAdjust As Boolean
    
    blnAdjust = True
    If picSplit.Left + X < 3000 Then picSplit.Left = 3000: blnAdjust = False
    If picSplit.Left + X > Me.ScaleWidth - 2000 Then picSplit.Left = Me.ScaleWidth - 2000: blnAdjust = False
    If blnAdjust Then
        picSplit.Left = picSplit.Left + X
        Call Form_Resize
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDown = False
    Call Form_Resize
End Sub

Private Sub tvw_s_NodeCheck(ByVal node As MSComctlLib.node)
    '自动勾选下级结点
    Call NodeCheck(node, node.Checked, True)
    Call NodeSelAll(node)
End Sub

Private Sub tvw_s_NodeClick(ByVal node As MSComctlLib.node)
    If node.Key = mstrPreNode Then Exit Sub
    '结点改变时,保存当前顺序(分类型)
    mstrPreNode = node.Key
    
    Call FillList
End Sub

Private Sub NodeCheck(ByVal node As MSComctlLib.node, ByVal blnSel As Boolean, Optional ByVal blnParent As Boolean = False)
    '递归调用,循环选择所有子结点
    node.Checked = blnSel
    If node.Children > 0 Then Call NodeCheck(node.Child, blnSel)
    If blnParent Then Exit Sub
    If Not node.Next Is Nothing Then Call NodeCheck(node.Next, blnSel)
End Sub

Private Function NodeSelAll(ByVal node As MSComctlLib.node) As Boolean
    '检查同级(只要选择了一个子结点,父结点都应该勾选;一个子结点都没选,父结点不需要勾选)
    Dim intCount As Integer
    Dim nodSource As MSComctlLib.node
    
    Set nodSource = node
    If Not node.Parent Is Nothing Then Set node = node.Parent.Child     '如果当前不是根结点，设置为第一个子结点
    If node.Checked Then intCount = 1
    Do While True
        If Not node.Next Is Nothing Then
            If node.Next.Checked Then intCount = intCount + 1
            If intCount > 0 Then Exit Do
            Set node = node.Next
        Else
            Exit Do
        End If
    Loop
    
    '向上回溯
    Set node = nodSource
    Do While True
        If Not node.Parent Is Nothing Then
            node.Parent.Checked = intCount
            Set node = node.Parent
        Else
            Exit Do
        End If
    Loop
End Function

Private Function GetSelNodes(ByVal node As MSComctlLib.node) As String
    Dim strSel As String
    Dim strReturn As String
    
    '获取所有选择的最末级结点
    If node.Checked Then
        If node.Children > 0 Then
            strSel = GetSelNodes(node.Child)
            If strSel <> "" Then strReturn = strReturn & IIf(strReturn <> "", ",", "") & strSel
        Else
            strReturn = strReturn & IIf(strReturn <> "", ",", "") & Mid(node.Key, 2)
        End If
    End If
    If Not node.Next Is Nothing Then
        strSel = GetSelNodes(node.Next)
        If strSel <> "" Then strReturn = strReturn & IIf(strReturn <> "", ",", "") & strSel
    End If
    GetSelNodes = strReturn
End Function

Private Sub txtFind_Change()
    If Trim(txtFind.Text) = "" Then
        lblInfo.Caption = "请输入查找条件"
        lblInfo.ForeColor = &H8000&
    Else
        lblInfo.Caption = "点击定位完成词句查找"
        lblInfo.ForeColor = &H8000000D
    End If
    
    cmdFind.Caption = "定位(&L)"
    Set mrsFind = New ADODB.Recordset
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdFind.SetFocus
        Call cmdFind_Click
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vfgBound_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    With vfgBound
        If KeyCode = vbKeyDelete And .Row >= .FixedRows Then
            If .RowData(.Row) > 0 Then
                If .Row = .FixedRows And .Row = .Rows - 1 Then
                    .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    .RowData(.Row) = 0
                Else
                    If .Row < .Rows - 1 Then
                        .RowPosition(.Row) = .Rows - 1
                        For lngRow = .Row To .Rows - 1
                            .TextMatrix(lngRow, 0) = lngRow + .FixedRows - 1
                        Next
                    End If
                    .RemoveItem .Rows - 1
                End If
                vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
        End If
    End With
End Sub

Private Sub vfgItem_DblClick()
    Dim i As Long
    Dim blnAdd As Boolean
    With vfgItem
        blnAdd = True
        If .Row >= .FixedRows And .Cols >= .FixedCols Then
            If Val(.TextMatrix(.Row, .FixedCols)) <= 0 Then Exit Sub
            For i = vfgBound.FixedRows To vfgBound.Rows - 1
                If Val(.TextMatrix(.Row, 3)) = Val(vfgBound.TextMatrix(i, 1)) Then
                    blnAdd = False
                    Exit For
                End If
            Next i
            
            If blnAdd = True Then
                If Val(vfgBound.TextMatrix(vfgBound.Rows - 1, 1)) > 0 Or vfgBound.Rows = vfgBound.FixedRows Then
                    vfgBound.Rows = vfgBound.Rows + 1
                End If
                vfgBound.TextMatrix(vfgBound.Rows - 1, 0) = vfgBound.Rows - vfgBound.FixedRows
                vfgBound.TextMatrix(vfgBound.Rows - 1, 1) = .TextMatrix(.Row, 3)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 2) = .TextMatrix(.Row, 5)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 3) = .TextMatrix(.Row, 7)
                vfgBound.TextMatrix(vfgBound.Rows - 1, 4) = .TextMatrix(.Row, 8)
                vfgBound.RowData(vfgBound.Rows - 1) = Val(.TextMatrix(.Row, 3))
                vfgBound.Row = vfgBound.Rows - 1
                vfgBound.TopRow = vfgBound.Rows - 1
                vfgBound.ColWidth(0) = Me.TextWidth(vfgBound.TextMatrix(vfgBound.Rows - 1, 0) & " ")
                If vfgBound.ColWidth(0) < 380 Then vfgBound.ColWidth(0) = 380
            End If
        End If
    End With
End Sub
