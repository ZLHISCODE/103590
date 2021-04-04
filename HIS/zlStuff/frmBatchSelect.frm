VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫材批量选择"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdGet 
      Caption         =   "提取"
      Height          =   300
      Left            =   4785
      TabIndex        =   19
      Top             =   750
      Width           =   510
   End
   Begin VB.TextBox txtCostEnd 
      Height          =   300
      Left            =   4185
      TabIndex        =   18
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostBegin 
      Height          =   300
      Left            =   3360
      TabIndex        =   16
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboCost 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   750
      Width           =   2055
   End
   Begin VB.PictureBox picDrug 
      Height          =   5775
      Left            =   9120
      ScaleHeight     =   5715
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   1060
      Visible         =   0   'False
      Width           =   3975
      Begin VSFlex8Ctl.VSFlexGrid vsfDrug 
         Height          =   5830
         Left            =   0
         TabIndex        =   14
         Top             =   -120
         Width           =   3975
         _cx             =   7011
         _cy             =   10283
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBatchSelect.frx":000C
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
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   12120
      TabIndex        =   6
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   11040
      TabIndex        =   5
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "清空(&O)"
      Height          =   300
      Left            =   9960
      TabIndex        =   4
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox txtSelect 
      Height          =   300
      Left            =   9120
      TabIndex        =   1
      Top             =   780
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":0081
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":061B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
      Height          =   6645
      Left            =   120
      TabIndex        =   8
      Top             =   1095
      Width           =   12975
      _cx             =   22886
      _cy             =   11721
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBatchSelect.frx":7417
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
   Begin VB.Frame fra调整额 
      Caption         =   "调整方式"
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   100
      Width           =   12980
      Begin VB.ComboBox cbo调整方式 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2580
      End
      Begin VB.TextBox txt调整额 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblInfor 
         Caption         =   "根据成本价，输入新的加成率重新加成调价"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   263
         Width           =   3420
      End
      Begin VB.Label lbl调整 
         AutoSize        =   -1  'True
         Caption         =   "％"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3540
         TabIndex        =   12
         Top             =   285
         Width           =   225
      End
   End
   Begin VB.Label lblCost 
      AutoSize        =   -1  'True
      Caption         =   "成本价范围"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "--"
      Height          =   180
      Left            =   3960
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "查找"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   7980
      Width           =   360
   End
   Begin VB.Label lblCalss 
      AutoSize        =   -1  'True
      Caption         =   "品种简码"
      Height          =   180
      Left            =   8355
      TabIndex        =   0
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '本模块中设置的显示单位 0-散装单位,1-包装单位
Private Const mlngRowHeight As Long = 300 '表格中各行行高
Private mrsReturn As ADODB.Recordset        '返回选定卫材数据
Private mblnOk As Boolean   '记录是否是点击的确定按钮
Private mrsFindName As ADODB.Recordset '记录查询数据集
Private mstrMatch  As String '0-双向匹配 1-单向右匹配
Private mintType As Integer '调整方式
Private mdbl比率 As Double  '调整方式的比率
Private mint调价 As Integer  '只调成本价时不能输入比率
Private Const mstrCaption As String = "卫材批量选择"
Private mstr调整额 As String
'各单位
Private mFMT As g_FmtString

Private Enum vsfSelectDrugCol
    材料ID = 0
    材料信息 = 1
    材料编码
    商品名
    通用名
    规格
    产地
    单位
    单位系数
    跟踪在用
    散装单位
    包装系数
    包装单位
    类型
    原售价
    售价
    成本价
    指导批价
    指导售价
    总列数
End Enum

Public Sub ShowMe(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOK As Boolean, ByRef intType As Integer, ByRef dbl比率 As Double, Optional int调价 As Integer, Optional str调整额 As String)
    mint调价 = int调价
    Me.Show vbModal, frmParent
    blnOK = mblnOk
    Set rsTemp = mrsReturn
    intType = mintType
    dbl比率 = mdbl比率
    str调整额 = mstr调整额
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.总列数
        .Rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMove '移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度

        '设置列宽
        .ColWidth(vsfSelectDrugCol.材料ID) = 0
        .ColWidth(vsfSelectDrugCol.材料信息) = 3000
        .ColWidth(vsfSelectDrugCol.材料编码) = 0
        .ColWidth(vsfSelectDrugCol.商品名) = 0
        .ColWidth(vsfSelectDrugCol.通用名) = 0
        .ColWidth(vsfSelectDrugCol.产地) = 1500
        .ColWidth(vsfSelectDrugCol.跟踪在用) = 0
        .ColWidth(vsfSelectDrugCol.单位) = 500
        .ColWidth(vsfSelectDrugCol.散装单位) = 0
        .ColWidth(vsfSelectDrugCol.包装系数) = 0
        .ColWidth(vsfSelectDrugCol.包装单位) = 0
        
        .ColWidth(vsfSelectDrugCol.类型) = 1000
        .ColWidth(vsfSelectDrugCol.售价) = 1500
        .ColWidth(vsfSelectDrugCol.原售价) = 0
        .ColWidth(vsfSelectDrugCol.成本价) = 1500
        .ColWidth(vsfSelectDrugCol.指导批价) = 1500
        .ColWidth(vsfSelectDrugCol.指导售价) = 1500
        .ColWidth(vsfSelectDrugCol.单位系数) = 0
        '设置列头
        .TextMatrix(0, vsfSelectDrugCol.材料ID) = "材料id"
        .TextMatrix(0, vsfSelectDrugCol.材料信息) = "材料信息"
        .TextMatrix(0, vsfSelectDrugCol.材料编码) = "材料编码"
        .TextMatrix(0, vsfSelectDrugCol.跟踪在用) = "跟踪在用"
        .TextMatrix(0, vsfSelectDrugCol.商品名) = "商品名"
        .TextMatrix(0, vsfSelectDrugCol.通用名) = "通用名"
        .TextMatrix(0, vsfSelectDrugCol.规格) = "规格"
        .TextMatrix(0, vsfSelectDrugCol.产地) = "产地"
        .TextMatrix(0, vsfSelectDrugCol.单位) = "单位"
        
        .TextMatrix(0, vsfSelectDrugCol.散装单位) = "散装单位"
        .TextMatrix(0, vsfSelectDrugCol.包装系数) = "包装系数"
        .TextMatrix(0, vsfSelectDrugCol.包装单位) = "包装单位"
        
        .TextMatrix(0, vsfSelectDrugCol.类型) = "类型"
        .TextMatrix(0, vsfSelectDrugCol.售价) = "售价"
        .TextMatrix(0, vsfSelectDrugCol.成本价) = "成本价"
        .TextMatrix(0, vsfSelectDrugCol.指导批价) = "指导批价"
        .TextMatrix(0, vsfSelectDrugCol.指导售价) = "指导售价"

        .ColAlignment(vsfSelectDrugCol.材料ID) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.材料信息) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.材料编码) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.规格) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.产地) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.单位) = flexAlignCenterCenter
        .ColAlignment(vsfSelectDrugCol.类型) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.售价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.成本价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导批价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub


'Private Sub setTvwInfo()
'    '为树表填充数据
'    Dim objNode As Node
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
'              " Where Instr([1],编码,1) > 0 " & _
'              " Order by 编码"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "4")
'
'    If rsTemp Is Nothing Then
'        Exit Sub
'    End If
'
'    With tvwDrug
'        .Nodes.Clear
'        Do While Not rsTemp.EOF
'            .Nodes.Add , , "Root" & rsTemp!名称, rsTemp!名称, 1, 1
'            .Nodes("Root" & rsTemp!名称).Tag = rsTemp!编码
'            rsTemp.MoveNext
'        Loop
'    End With
'
'
'    gstrSQL = "Select ID, 上级id, 编码, 名称, Decode(类型, 7, '卫材') 分类, '分类' As 类别" & _
'                " From 诊疗分类目录" & _
'                " Where 类型 ='7' And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                " Start With 上级id Is Null" & _
'                " Connect By Prior ID = 上级id"
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "分类查询")
'    With rsTemp
'        Do While Not .EOF
'           If IsNull(!上级ID) Then
'                Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'            Else
'                Set objNode = tvwDrug.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'            End If
'            objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'            .MoveNext
'        Loop
'    End With
'
'    If optVariety.Value = True Then
'        gstrSQL = "Select ID, 分类id, 编码, 名称, Decode(类别, 7, '卫材') 分类, '品种' As 类别" & _
'                  "  From 诊疗项目目录" & _
'                  "  Where 分类id In (Select ID" & _
'                                   " From 诊疗分类目录" & _
'                                   " Where 类型 ='7' And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                                   " Start With 上级id Is Null" & _
'                                   " Connect By Prior ID = 上级id)"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "品种")
'
'        With rsTemp
'            Do While Not .EOF
'                Set objNode = tvwDrug.Nodes.Add("K_" & !分类id, 4, !类别 & "K_" & !Id, !名称 & "-品种", 1, 1)
'                objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                .MoveNext
'            Loop
'        End With
'    End If
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub cboCost_Click()
    If cboCost.Text = "自定义" Then
        txtCostBegin.Visible = True
        txtCostEnd.Visible = True
        lblTo.Visible = True
        cmdGet.Left = txtCostEnd.Left + txtCostEnd.Width + 5
    Else
        txtCostBegin.Visible = False
        txtCostEnd.Visible = False
        lblTo.Visible = False
        cmdGet.Left = txtCostBegin.Left
    End If
End Sub

Private Sub cbo调整方式_Click()
    Dim intType As Integer
    Dim dbl比率 As Double
    Dim dbl换算系数  As Double
    Dim intRow As Integer
    
    If cbo调整方式.ListIndex < 0 Then Exit Sub
    Select Case cbo调整方式.ItemData(cbo调整方式.ListIndex)
        Case 1
            lblInfor.Caption = "根据成本价，输入新的加成率重新加成调价"
            lbl调整.Caption = "％"
            txt调整额.MaxLength = 3
        Case 2
            lblInfor.Caption = "在当前售价基础上按照比例调价"
            lbl调整.Caption = "％"
            txt调整额.MaxLength = 3
        Case 3
            lblInfor.Caption = "在当前售价基础上按固定金额加减调价"
            lbl调整.Caption = "元"
            txt调整额.MaxLength = 10
    End Select

    intType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
    
    With vsfSelectDrug
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(1, vsfSelectDrugCol.材料ID) = "" Then Exit Sub
        For intRow = 1 To .Rows - 1
            dbl比率 = Val(txt调整额.Text)
            If Trim(txt调整额.Text) = "" Then
                .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价)), mFMT.FM_零售价)
            Else
                Select Case intType
                    Case 1      '根据成本价加成
                        dbl比率 = 1 + Val(dbl比率) / 100
                        .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.成本价)) * dbl比率, mFMT.FM_零售价)
                    Case 2      '根据零售价按比例
                        dbl比率 = 1 + Val(dbl比率) / 100
                        .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价)) * dbl比率, mFMT.FM_零售价)
                    Case 3      '根据零售价按固定金额加减
                        dbl比率 = Val(dbl比率)
                        .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format((Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价))) + dbl比率, mFMT.FM_零售价)
                End Select
            End If
            If Val(.TextMatrix(intRow, vsfSelectDrugCol.售价)) > Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)) And Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)) <> 0 Then
                .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)), mFMT.FM_零售价)
            End If
        Next
    End With
End Sub

Private Sub cmdGet_Click()
    Dim dblBegin As Double
    Dim dblEnd As Double
    Dim strTemp As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If cboCost.Text = "自定义" Then
        If Trim(txtCostBegin.Text) = "" Then
            MsgBox "请输入要查询成本价的开始价格！", vbInformation, gstrSysName
            txtCostBegin.SetFocus
            Exit Sub
        End If
        If Trim(txtCostEnd.Text) = "" Then
            MsgBox "请输入要查询成本价的结束价格！", vbInformation, gstrSysName
            txtCostEnd.SetFocus
            Exit Sub
        End If
    End If
    
    vsfSelectDrug.Rows = 1
    
    If cboCost.Text <> "自定义" Then
        strTemp = cboCost.Text
        dblBegin = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
        dblEnd = Mid(strTemp, InStr(1, strTemp, "-") + 1, InStr(1, strTemp, "(") - InStr(1, strTemp, "-") - 1)
    Else
        dblBegin = Trim(txtCostBegin.Text)
        dblEnd = Trim(txtCostEnd.Text)
    End If
    
    If mintUnit = 0 Then
        '散装单位
        strTemp = "And (b.平均成本价 Between [1] And [2] Or d.成本价 Between [1] And [2])"
    Else
        '药库单位
        strTemp = "And (b.平均成本价 Between [1]/d.换算系数 And [2]/d.换算系数 Or d.成本价 Between [1]/d.换算系数 And [2]/d.换算系数)"
    End If
    
    gstrSQL = "Select Distinct a.Id As 材料id, a.编码 As 材料编码, a.名称 As 通用名, c.商品名, a.规格, a.是否变价 As 时价, a.产地, a.计算单位, d.换算系数, d.包装单位," & vbNewLine & _
                "                Decode(Nvl(b.平均成本价, 0), 0, d.成本价, b.平均成本价) As 成本价," & vbNewLine & _
                "                Decode(Nvl(b.实际数量, 0), 0, e.现价, b.实际金额 / b.实际数量) As 现价, d.指导批发价, d.指导零售价, d.跟踪在用" & vbNewLine & _
                "From 收费项目目录 A, 药品库存 B, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) C, 材料特性 D, 收费价目 E" & vbNewLine & _
                "Where a.Id = b.药品id(+) And a.Id = c.收费细目id(+) And a.Id = d.材料id And a.Id = e.收费细目id And Sysdate Between e.执行日期 And" & vbNewLine & _
                "      e.终止日期 And (a.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) " & GetPriceClassString("E") & _
                " And a.类别 = '4' " & strTemp
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "按成本价范围查询", dblBegin, dblEnd)
        
    If rsData.RecordCount > 0 Then
        Call setVSFValue(rsData)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCostBegin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
        Else
           KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txt调整额_KeyPress(KeyAscii As Integer)
    Dim intType As Integer
    Dim dbl比率 As Double
    Dim intRow As Integer
    Dim dbl换算系数 As Double
    
    If cbo调整方式.ItemData(cbo调整方式.ListIndex) = 3 Then
        Call zlControl.TxtCheckKeyPress(txt调整额, KeyAscii, m负金额式)
    Else
        Call zlControl.TxtCheckKeyPress(txt调整额, KeyAscii, m金额式)
    End If
    If KeyAscii <> 0 And KeyAscii = vbKeyReturn Then
        intType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
        
        With vsfSelectDrug
            For intRow = 1 To .Rows - 1
                dbl比率 = Val(txt调整额.Text)
                If Trim(txt调整额.Text) = "" Then
                    .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价)), mFMT.FM_零售价)
                Else
                    Select Case intType
                        Case 1      '根据成本价加成
                            dbl比率 = 1 + Val(dbl比率) / 100
                            .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.成本价)) * dbl比率, mFMT.FM_零售价)
                        Case 2      '根据零售价按比例
                            dbl比率 = 1 + Val(dbl比率) / 100
                            .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价)) * dbl比率, mFMT.FM_零售价)
                        Case 3      '根据零售价按固定金额加减
                            dbl比率 = Val(dbl比率)
                            .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format((Val(.TextMatrix(intRow, vsfSelectDrugCol.原售价))) + dbl比率, mFMT.FM_零售价)
                    End Select
                End If
                If Val(.TextMatrix(intRow, vsfSelectDrugCol.售价)) > Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)) And Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)) <> 0 Then
                    .TextMatrix(intRow, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)), mFMT.FM_零售价)
                End If
            Next
        End With
    End If
End Sub

Private Sub cbo调整方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

'Private Sub cmdSelect_Click()
'    picDrug.Visible = True
'    tvwDrug.Visible = True
'    Call setTvwInfo
'End Sub

Private Sub cmdCal_Click()
    With vsfSelectDrug
        If MsgBox("确定要清空所有已经选择的卫材？", vbYesNo, gstrSysName) = vbYes Then
            .Rows = 1
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset
    mintType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
    mdbl比率 = Val(txt调整额.Text)
    mstr调整额 = Trim(txt调整额.Text)
    With mrsReturn
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "时价", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "跟踪在用", adLongVarChar, 2, adFldIsNullable
        
        .Fields.Append "计算单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "换算系数", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "包装单位", adLongVarChar, 11, adFldIsNullable
        .Fields.Append "成本价", adDouble, 18, adFldIsNullable
        .Fields.Append "指导批发价", adDouble, 18, adFldIsNullable
        .Fields.Append "指导零售价", adDouble, 18, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    With vsfSelectDrug
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.材料ID) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!Id = .TextMatrix(intRow, vsfSelectDrugCol.材料ID)
            mrsReturn!编码 = .TextMatrix(intRow, vsfSelectDrugCol.材料编码)
            mrsReturn!商品名 = .TextMatrix(intRow, vsfSelectDrugCol.商品名)
            mrsReturn!通用名 = .TextMatrix(intRow, vsfSelectDrugCol.通用名)
            mrsReturn!规格 = .TextMatrix(intRow, vsfSelectDrugCol.规格)
            mrsReturn!时价 = .TextMatrix(intRow, vsfSelectDrugCol.类型)
            mrsReturn!产地 = .TextMatrix(intRow, vsfSelectDrugCol.产地)
            mrsReturn!跟踪在用 = .TextMatrix(intRow, vsfSelectDrugCol.跟踪在用)
            mrsReturn!计算单位 = .TextMatrix(intRow, vsfSelectDrugCol.散装单位)
            mrsReturn!换算系数 = .TextMatrix(intRow, vsfSelectDrugCol.包装系数)
            mrsReturn!包装单位 = .TextMatrix(intRow, vsfSelectDrugCol.包装单位)
            
            mrsReturn!成本价 = Val(.TextMatrix(intRow, vsfSelectDrugCol.成本价)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.单位系数))
            mrsReturn!指导批发价 = Val(.TextMatrix(intRow, vsfSelectDrugCol.指导批价)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.单位系数))
            mrsReturn!指导零售价 = Val(.TextMatrix(intRow, vsfSelectDrugCol.指导售价)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.单位系数))
            
            mrsReturn.Update
        Next
    End With
    mblnOk = True
    
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfDrug.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    '获取设置的单位
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, 1726, 1))
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    With cbo调整方式
        .AddItem "根据成本价按加成调价"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "根据售价按比例调价"
        .ItemData(.NewIndex) = 2
        .AddItem "根据售价按固定金额调价"
        .ItemData(.NewIndex) = 3
    End With
    
    With cboCost
        .AddItem "0-10(含10)"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "10-20(含20)"
        .ItemData(.NewIndex) = 2
        .AddItem "20-50(含50)"
        .ItemData(.NewIndex) = 3
        .AddItem "自定义"
        .ItemData(.NewIndex) = 4
    End With
    
    If mintUnit = 0 Then
        Me.Caption = "卫材批量选择(散装单位)"
    Else
        Me.Caption = "卫材批量选择(包装单位)"
    End If
    
    cmdGet.Left = txtCostBegin.Left
    
    mstrMatch = IIf(zlDatabase.GetPara("输入匹配", , , 0) = "0", "%", "")
    mblnOk = False
    
    If mint调价 = 1 Then
        txt调整额.Enabled = False
        txt调整额.BackColor = &H80000000
        cbo调整方式.Enabled = False
    Else
        txt调整额.Enabled = True
        txt调整额.BackColor = &H80000005
        cbo调整方式.Enabled = True
    End If
    
    Call initVsflexgrid
    
    Call RestoreWinState(Me, App.ProductName, mstrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrCaption)
End Sub

Private Sub optClass_Click()
    picDrug.Visible = False
    lblCalss.Caption = "分类"
End Sub

Private Sub optClassSub_Click()
    picDrug.Visible = False
    lblCalss.Caption = "分类(含子类)"
End Sub

Private Sub optVariety_Click()
    picDrug.Visible = False
    lblCalss.Caption = "品种"
End Sub

'Private Sub tvwDrug_NodeClick(ByVal Node As MSComctlLib.Node)
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    If Node.Key Like "Root" Then Exit Sub
'
'    gstrSQL = "select id,编码,名称,计算单位 from 诊疗项目目录 where  Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' and 分类id=[1]"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查询品种", Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
'
'    Set vsfDetails.DataSource = rsTemp
'
'    Exit Sub
'errHandle:
'    If errcenter() = 1 Then Resume
'    Call saveerrlog
'End Sub

'Private Sub tvwDrug_DblClick()
'    '用来向界面中传入值
'    Dim lngID As Long
'    Dim rsTemp As ADODB.Recordset
'    Dim intRow As Integer
'    Dim i As Integer
'    Dim blnDou As Boolean '重复数据
'    Dim dbl换算系数 As Double
'    Dim strUnit As String   '单位
'    Dim intType As Integer '调整方式
'    Dim dbl比率 As Double   '调整额
'
'    On Error GoTo errHandle
'
'    intType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
'    dbl比率 = Val(txt调整额.Text)
'    With tvwDrug
'        If optVariety.Value = True Then
'            If InStr(1, .SelectedItem.Text, "-品种") <= 0 Then
'                Exit Sub
'            End If
'            gstrSQL = "Select Distinct a.材料id, c.编码 As 材料编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位,a.换算系数, a.包装单位," & _
'                                        " a.成本价, e.现价, a.指导批发价, a.指导零售价,a.跟踪在用" & _
'                        " From 材料特性 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D,收费价目 E" & _
'                        " Where a.诊疗id = b.Id And a.材料id = c.Id And c.Id = d.收费细目id(+) and a.材料id=e.收费细目id and sysdate between e.执行日期 and e.终止日期 and b.id=[1] order by c.编码"
'        Else
'            If InStr(1, .SelectedItem.Text, "-分类") <= 0 Then
'                Exit Sub
'            End If
'                If optClassSub.Value = True Then '分类下子节点
'                    gstrSQL = "(Select ID From 诊疗分类目录 Where 类型 =7 Start With ID = [1] Connect By Prior ID = 上级id) A,"
'                Else '本分类
'                    gstrSQL = "(select id from 诊疗分类目录 where 类型 =7 and id=[1]) A,"
'                End If
'
'                gstrSQL = "Select Distinct c.材料id, d.编码 As 材料编码, d.名称 As 通用名, f.商品名, d.规格, d.是否变价 As 时价, d.产地, d.计算单位, c.换算系数, c.包装单位," & _
'                                        "  c.成本价, e.现价, c.指导批发价, c.指导零售价,c.跟踪在用 " & _
'                        " From " & gstrSQL & " 诊疗项目目录 B, 材料特性 C," & _
'                             " 收费项目目录 D, 收费价目 E, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) F" & _
'                        " Where a.Id = b.分类id And b.Id = c.诊疗id And c.材料id = d.Id And d.Id = e.收费细目id And e.收费细目id = f.收费细目id(+) And" & _
'                              " Sysdate Between e.执行日期 And e.终止日期 order by d.编码"
'        End If
'        lngID = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "K_") + 2)
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询卫材", lngID)
'        If rsTemp.RecordCount = 0 Then
'            picDrug.Visible = False
'            Exit Sub
'        End If
'    End With
'
'    With vsfSelectDrug
'        For intRow = 0 To rsTemp.RecordCount - 1
'            blnDou = False
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, vsfSelectDrugCol.材料ID) = rsTemp!材料ID Then
'                    blnDou = True
'                End If
'            Next
'            If blnDou = False Then
'                .Rows = .Rows + 1
'                .RowHeight(.Rows - 1) = mlngRowHeight
'
'                Select Case mintUnit
'                    Case 0
'                        dbl换算系数 = 1
'                        strUnit = rsTemp!计算单位
'                    Case 1
'                        dbl换算系数 = rsTemp!换算系数
'                        strUnit = rsTemp!包装单位
'                End Select
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料ID) = rsTemp!材料ID
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料信息) = "[" & rsTemp!材料编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料编码) = rsTemp!材料编码
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.商品名) = IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.通用名) = IIf(IsNull(rsTemp!通用名), "", rsTemp!通用名)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.单位) = strUnit
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.单位系数) = dbl换算系数
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.散装单位) = rsTemp!计算单位
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.包装单位) = rsTemp!包装单位
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.包装系数) = rsTemp!换算系数
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.类型) = IIf(rsTemp!时价 = 1, "时价", "定价")
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.跟踪在用) = NVL(rsTemp!跟踪在用)
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.成本价) = Format(dbl换算系数 * rsTemp!成本价, mFMT.FM_成本价)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.指导批价) = Format(dbl换算系数 * rsTemp!指导批发价, mFMT.FM_成本价)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价) = Format(dbl换算系数 * rsTemp!指导零售价, mFMT.FM_零售价)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.原售价) = Format(Val(NVL(rsTemp!现价)) * dbl换算系数, mFMT.FM_零售价)
'                If dbl比率 = 0 Then
'                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * dbl换算系数, mFMT.FM_零售价)
'                Else
'                    Select Case intType
'                        Case 1      '根据成本价加成
'                            dbl比率 = 1 + Val(dbl比率) / 100
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(NVL(rsTemp!成本价)) * dbl比率 * dbl换算系数, mFMT.FM_零售价)
'                        Case 2      '根据零售价按比例
'                            dbl比率 = 1 + Val(dbl比率) / 100
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(NVL(rsTemp!现价)) * dbl比率 * dbl换算系数, mFMT.FM_零售价)
'                        Case 3      '根据零售价按固定金额加减
'                            dbl比率 = Val(dbl比率)
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format((Val(NVL(rsTemp!现价)) * dbl换算系数) + dbl比率, mFMT.FM_零售价)
'                    End Select
'                End If
'                If Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.售价)) > Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)) And Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)) <> 0 Then
'                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)), mFMT.FM_零售价)
'                End If
'            End If
'            rsTemp.MoveNext
'        Next
'        picDrug.Visible = False
'    End With
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long
    
    '查找卫材
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 材料编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别='4' " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 材料编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "取匹配的卫材ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        str药名 = mrsFindName!材料编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        
        For lngRow = 1 To vsfSelectDrug.Rows - 1
            lngFindRow = vsfSelectDrug.FindRow(str药名, lngRow, CLng(vsfSelectDrugCol.材料信息), True, True)
            If lngFindRow > 0 Then
                vsfSelectDrug.Select lngFindRow, 1, lngFindRow, vsfSelectDrug.Cols - 1
                vsfSelectDrug.TopRow = lngFindRow
                Exit For
            End If
        Next
        
        If lngFindRow > 0 Then  '查询到数据后就移动下下一条并退出本次查询
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtSelect_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub txtSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim rsPinzhong As ADODB.Recordset
    Dim objNode As Node
    Dim lng分类id As Long
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
    
        On Error GoTo ErrHandle
        
        If Trim(txtSelect.Text) = "" Then Exit Sub
                
        gstrSQL = "Select Distinct a.id,a.编码,a.名称" & _
                  "  From 诊疗项目目录 A, 诊疗项目别名 B" & _
                    " Where a.Id = b.诊疗项目id(+) And a.类别 ='4' And Sysdate Between 建档时间 And 撤档时间 And" & _
                         " (a.编码 Like [1] Or a.名称 Like [1] Or b.名称 Like [1] Or b.简码 Like [1])"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
        If rsTemp.RecordCount = 0 Then
            MsgBox "未查询到品种！", vbInformation, gstrSysName
            txtSelect.SetFocus
            txtSelect.SelStart = 1
            txtSelect.SelLength = Len(txtSelect.Text)
        Else
            picDrug.Visible = True
            vsfDrug.Visible = True
            Set vsfDrug.DataSource = rsTemp
            vsfDrug.SetFocus
            vsfDrug.Row = 1
        End If
        With vsfDrug
            For i = 0 To .Rows - 1
                .RowHeight(i) = mlngRowHeight
            Next
        End With
        
'        picDrug.Visible = True
'        tvwDrug.Visible = True
'
'        gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
'                  " Where Instr([1],编码,1) > 0 " & _
'                  " Order by 编码"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "4")
'
'        If rsTemp Is Nothing Then
'            Exit Sub
'        End If
'
'        With tvwDrug
'            .Nodes.Clear
'            Do While Not rsTemp.EOF
'                .Nodes.Add , , "Root" & rsTemp!名称, rsTemp!名称, 1, 1
'                .Nodes("Root" & rsTemp!名称).Tag = rsTemp!编码
'                rsTemp.MoveNext
'            Loop
'        End With
'
'        If optVariety.Value = True Then '品种被选中
'            gstrSQL = "Select distinct a.Id, a.上级id, a.编码, a.名称, Decode(a.类型, 7, '卫材') 分类, '分类' As 类别" & _
'                        " From 诊疗分类目录 A" & _
'                        " Start With ID In (Select Distinct a.分类id" & _
'                                          " From 诊疗项目目录 A, 诊疗项目别名 B" & _
'                                          " Where a.Id = b.诊疗项目id(+) And a.类别 = '4' And Sysdate Between 建档时间 And 撤档时间 And" & _
'                                                " (a.编码 Like [1] Or a.名称 Like [1] Or b.名称 Like [1] Or b.简码 Like [1]))" & _
'                        " Connect By Prior a.上级id = a.Id" & _
'                        " order by a.id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!上级ID) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    End If
'                    objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                    .MoveNext
'                Loop
'
'                rsTemp.MoveFirst
'                Do While Not rsTemp.EOF
'                    lng分类id = rsTemp!Id
'                    gstrSQL = "Select Distinct a.Id, a.分类id, a.编码, a.名称, Decode(a.类别, '4','卫材') 分类, '品种' As 类别" & _
'                                " From 诊疗项目目录 A" & _
'                                " Where a.类别 ='4' And a.分类id=[1] and Sysdate Between a.建档时间 And a.撤档时间"
'                    Set rsPinzhong = zlDatabase.OpenSQLRecord(gstrSQL, "品种", lng分类id)
'
'                    Do While Not rsPinzhong.EOF
'                        Set objNode = tvwDrug.Nodes.Add("K_" & rsPinzhong!分类id, 4, rsPinzhong!类别 & "K_" & rsPinzhong!Id, rsPinzhong!名称 & "-品种", 1, 1)
'                        objNode.Tag = rsPinzhong!分类 & "-" & rsPinzhong!类别
'                        rsPinzhong.MoveNext
'                    Loop
'                    rsTemp.MoveNext
'                Loop
'            End With
'        Else
'            gstrSQL = "Select ID, 上级id, 编码, 名称, Decode(类型, 7, '卫材') 分类, '分类' As 类别" & _
'                        " From 诊疗分类目录" & _
'                        " Start With ID in (Select ID" & _
'                                         " From 诊疗分类目录" & _
'                                         " Where 类型 = '7' And (Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' Or 撤档时间 Is Null) And" & _
'                                               " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1]))" & _
'                        " Connect By Prior 上级id = ID" & _
'                        " order by id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!上级ID) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    End If
'                    objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                    .MoveNext
'                Loop
'            End With
'        End If
'        tvwDrug.SetFocus
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_DblClick()
    Dim lngId As Long
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    
    With vsfDrug
        If Val(.TextMatrix(.Row, 0)) = 0 Then
            Exit Sub
        End If
        gstrSQL = "Select Distinct a.材料id, c.编码 As 材料编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位,a.换算系数, a.包装单位," & _
                                    " a.成本价, e.现价, a.指导批发价, a.指导零售价,a.跟踪在用" & _
                    " From 材料特性 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D,收费价目 E" & _
                    " Where a.诊疗id = b.Id And a.材料id = c.Id And c.Id = d.收费细目id(+) and a.材料id=e.收费细目id and sysdate between e.执行日期 and e.终止日期 " & _
                    GetPriceClassString("E") & "and b.id=[1] And (c.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or c.撤档时间 is null ) order by c.编码"
    
        lngId = Val(.TextMatrix(.Row, 0))

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询卫材", lngId, gstrPriceClass)
        If rsTemp.RecordCount = 0 Then
            picDrug.Visible = False
            Exit Sub
        End If
    End With
    
    '为表格赋值
    Call setVSFValue(rsTemp)

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub setVSFValue(ByVal rsTemp As ADODB.Recordset)
    Dim lngId As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '重复数据
    Dim dbl换算系数 As Double
    Dim strUnit As String   '单位
    Dim intType As Integer '调整方式
    Dim dbl比率 As Double   '调整额
    
    intType = cbo调整方式.ItemData(cbo调整方式.ListIndex)
    '为表格赋值
    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .Rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.材料ID) = rsTemp!材料ID Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = mlngRowHeight
            
                Select Case mintUnit
                    Case 0
                        dbl换算系数 = 1
                        strUnit = rsTemp!计算单位
                    Case 1
                        dbl换算系数 = rsTemp!换算系数
                        strUnit = rsTemp!包装单位
                End Select
                                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料ID) = rsTemp!材料ID
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料信息) = "[" & rsTemp!材料编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)

                .TextMatrix(.Rows - 1, vsfSelectDrugCol.材料编码) = rsTemp!材料编码
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.商品名) = IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.通用名) = IIf(IsNull(rsTemp!通用名), "", rsTemp!通用名)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.单位) = strUnit
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.单位系数) = dbl换算系数
                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.散装单位) = rsTemp!计算单位
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.包装单位) = rsTemp!包装单位
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.包装系数) = rsTemp!换算系数
                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.类型) = IIf(rsTemp!时价 = 1, "时价", "定价")
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.跟踪在用) = zlStr.nvl(rsTemp!跟踪在用)
                                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.成本价) = Format(dbl换算系数 * rsTemp!成本价, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.指导批价) = Format(dbl换算系数 * rsTemp!指导批发价, mFMT.FM_成本价)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价) = Format(dbl换算系数 * rsTemp!指导零售价, mFMT.FM_零售价)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.原售价) = Format(Val(zlStr.nvl(rsTemp!现价)) * dbl换算系数, mFMT.FM_零售价)
                
                dbl比率 = Val(txt调整额.Text)
                If Trim(txt调整额.Text) = "" Then
                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * dbl换算系数, mFMT.FM_零售价)
                Else
                    Select Case intType
                        Case 1      '根据成本价加成
                            dbl比率 = 1 + Val(dbl比率) / 100
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(zlStr.nvl(rsTemp!成本价)) * dbl比率 * dbl换算系数, mFMT.FM_零售价)
                        Case 2      '根据零售价按比例
                            dbl比率 = 1 + Val(dbl比率) / 100
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(zlStr.nvl(rsTemp!现价)) * dbl比率 * dbl换算系数, mFMT.FM_零售价)
                        Case 3      '根据零售价按固定金额加减
                            dbl比率 = Val(dbl比率)
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format((Val(zlStr.nvl(rsTemp!现价)) * dbl换算系数) + dbl比率, mFMT.FM_零售价)
                    End Select
                End If
                If Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.售价)) > Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)) And Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)) <> 0 Then
                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.售价) = Format(Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.指导售价)), mFMT.FM_零售价)
                End If
            End If
            rsTemp.MoveNext
        Next
        picDrug.Visible = False
    End With
End Sub

Private Sub vsfDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfDrug_DblClick
    End If
End Sub

Private Sub vsfSelectDrug_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub vsfSelectDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsfSelectDrug.Rows > 1 Then
        vsfSelectDrug.RemoveItem vsfSelectDrug.Row
    End If
End Sub
