VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmImportFileCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查设置"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   Icon            =   "frmImportFileCondition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8970
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   4980
      Left            =   105
      TabIndex        =   5
      Top             =   495
      Width           =   8760
      _cx             =   15452
      _cy             =   8784
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   30
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.OptionButton optPartImport 
      Caption         =   "跳过错误"
      Height          =   240
      Left            =   1380
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton optNoImport 
      Caption         =   "错误禁止"
      Height          =   255
      Left            =   2805
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   300
      Left            =   6870
      TabIndex        =   1
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   300
      Left            =   7950
      TabIndex        =   0
      Top             =   60
      Width           =   885
   End
   Begin VB.Label lblImportMethod 
      AutoSize        =   -1  'True
      Caption         =   "导入方式"
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmImportFileCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MCONFIXECOLOR As Long = &H8000000F  '不能修改列背景色
Private strPara             As String             '参数值
Private mlngModal           As Long               '当前模块号
Private mstrCheck           As String             '检查对象

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTemp As String
    Dim intRow  As Integer
    
    With vsfError
        If optNoImport.Value = True Then
            strTemp = "1/"
        Else
            strTemp = "0/"
        End If
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = "禁止" Then
                strTemp = strTemp & "1|"
            Else
                strTemp = strTemp & "0|"
            End If
        Next
    End With
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 1, LenB(StrConv(strTemp, vbFromUnicode)) - 1)
    Else
        strTemp = "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    End If
    Call zlDatabase.SetPara("导入文件检查方式", strTemp, glngSys, mlngModal)
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitVsf
    Call LoadData
End Sub

Public Sub ShowMe(ByVal frmPar As Form, ByVal lngModal As Long)
    mlngModal = lngModal
    Me.Show vbModal, frmPar
End Sub

Private Sub optNoImport_Click()
    Dim intRow As Integer
    
    With vsfError
        If optNoImport.Value = True Then
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, 1) = "禁止"
            Next
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, 3) = MCONFIXECOLOR '不能修改行颜色
            .Cell(flexcpFontBold, 1, 1, .Rows - 1, 1) = True
            .Editable = flexEDNone
            .Row = 0
        End If
    End With
End Sub

Private Sub optPartImport_Click()
    Dim intRow As Integer
    
    If optPartImport.Value = True Then
        For intRow = 1 To vsfError.Rows - 1
            vsfError.TextMatrix(intRow, 1) = "提示"
        Next
        vsfError.Row = 0
        vsfError.Cell(flexcpBackColor, 1, 1, vsfError.Rows - 1, 1) = &H80000005    '能修改行颜色
    End If
End Sub

Private Sub vsfError_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfError
        If Col = 1 Then
            If .TextMatrix(Row, Col) = "禁止" Then
                .Cell(flexcpFontBold, Row, 1, Row, 1) = True
            Else
                .Cell(flexcpFontBold, Row, 1, Row, 1) = False
            End If
        End If
    End With
End Sub

Private Sub vsfError_EnterCell()
    With vsfError
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = MCONFIXECOLOR Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub LoadData()
    '加载数据
    Dim strImportMethod As String
    Dim strPara         As String
    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim arryPara        As Variant
    Dim arryTempPara    As Variant
    Dim strTemp         As String
    
    '导入格式(0-错误提示1-错误禁止/0-提示1-禁止|0-提示1-禁止|....)
    strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    
    arryPara = Split(strPara, "|")
    With vsfError
        For intRow = 0 To UBound(arryPara)
            strTemp = arryPara(intRow)
            If intRow = 0 Then
                strImportMethod = Split(strTemp, "/")(0)
                If strImportMethod = "0" Then
                    optNoImport.Value = False
                    optPartImport.Value = True
                Else
                    optNoImport.Value = True
                    optPartImport.Value = False
                End If
                strTemp = Split(strTemp, "/")(1)
                If strTemp = "0" Then
                    .TextMatrix(intRow + 1, 1) = "提示"
                Else
                    .TextMatrix(intRow + 1, 1) = "禁止"
                    .Cell(flexcpFontBold, intRow + 1, 1) = True
                End If
            End If
            If strTemp = "0" Then
                .TextMatrix(intRow + 1, 1) = "提示"
            Else
                .TextMatrix(intRow + 1, 1) = "禁止"
                .Cell(flexcpFontBold, intRow + 1, 1) = True
            End If
        Next
    End With
End Sub

Private Sub InitVsf()
    '初始化vsf控件
    With vsfError
        .Cols = 4
        .Rows = 28
        .FixedRows = 1
        .FixedCols = 1
        .RowHeight(-1) = 450
        .ColWidth(2) = 1500
        .Editable = flexEDNone
        .AllowSelection = False '仅选择一行
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ExtendLastCol = True '最后一列填充满
        .ColComboList(1) = "禁止|提示"
        .WordWrap = True
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 3) = MCONFIXECOLOR '不能修改行颜色
        .Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter '列头居中加粗
        .Cell(flexcpFontBold, 0, 0, 0, 3) = True
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, 3) = flexAlignLeftCenter
        .AutoResize = True
        .WordWrap = True '文字换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动换行
        .MergeCells = flexMergeFree '单元格合并
        .MergeCol(0) = True
        
        .TextMatrix(0, 0) = "检查类型"
        .TextMatrix(0, 1) = "检查方式"
        .Cell(flexcpFontBold, 0, 1, 0, 1) = True
        .TextMatrix(0, 2) = "检查对象"
        .TextMatrix(0, 3) = "备注"
        
        .TextMatrix(1, 0) = "分类"
        .TextMatrix(1, 2) = "类别"
        .TextMatrix(1, 3) = "类别只能是西成药、中成药、中草药；不能为空"
        .TextMatrix(2, 0) = "分类"
        .TextMatrix(2, 2) = "上级名称"
        .TextMatrix(2, 3) = "上级分类参照表中已有的数据；格式：用\分隔各个级别"
        .TextMatrix(3, 0) = "分类"
        .TextMatrix(3, 2) = "编码"
        .TextMatrix(3, 3) = "不能含有非法字符；不能为空；长度不能超过数据库字段长度"
        .TextMatrix(4, 0) = "分类"
        .TextMatrix(4, 2) = "名称"
        .TextMatrix(4, 3) = "不能含有非法字符；不能为空；长度不能超过数据库字段长度"
        .TextMatrix(5, 0) = "分类"
        .TextMatrix(5, 2) = "编码和类别唯一检查"
        .TextMatrix(5, 3) = "编码和类别不能与界面现有或数据库中已有数据相同"
        .TextMatrix(6, 0) = "分类"
        .TextMatrix(6, 2) = "名称.类别.上级分类唯一检查"
        .TextMatrix(6, 3) = "名称、类别、上级分类不能与界面现有或数据库中已有数据相同"
        .TextMatrix(7, 0) = "明细"
        .TextMatrix(7, 2) = "类别"
        .TextMatrix(7, 3) = "只能是西成药、中成药、中草药；不能为空"
        .TextMatrix(8, 0) = "明细"
        .TextMatrix(8, 2) = "分类"
        .TextMatrix(8, 3) = "参照表中已有的数据；不能为空；格式：用\分隔各个级别"
        .TextMatrix(9, 0) = "明细"
        .TextMatrix(9, 2) = "品种编码"
        .TextMatrix(9, 3) = "不能含有非法字符；不能为空；长度不能超过数据库字段长度"
        .TextMatrix(10, 0) = "明细"
        .TextMatrix(10, 2) = "品种名称"
        .TextMatrix(10, 3) = "不能含有非法字符；不能为空；长度不能超过数据库字段长度"
        .TextMatrix(11, 0) = "明细"
        .TextMatrix(11, 2) = "规格编码"
        .TextMatrix(11, 3) = "不能含有非法字符；不能为空，长度不能超过数据库字段长度"
        .TextMatrix(12, 0) = "明细"
        .TextMatrix(12, 2) = "药品规格"
        .TextMatrix(12, 3) = "不能含有非法字符，不能为空，长度不能超过数据库字段长度"
        .TextMatrix(13, 0) = "明细"
        .TextMatrix(13, 2) = "生产商"
        .TextMatrix(13, 3) = "不能含有非法字符；长度不能超过数据库字段长度"
        .TextMatrix(14, 0) = "明细"
        .TextMatrix(14, 2) = "剂型"
        .TextMatrix(14, 3) = "参照表中已有的数据；不能为空；不能含有非法字符；长度不能超过数据库字段长度"
        .TextMatrix(15, 0) = "明细"
        .TextMatrix(15, 2) = "各级单位检查"
        .TextMatrix(15, 3) = "不能含有非法字符；不能为空；长度不能超过数据库字段长度"
        .TextMatrix(16, 0) = "明细"
        .TextMatrix(16, 2) = "各级单位换算检查"
        .TextMatrix(16, 3) = "不能为空；单位换算系数是数字且都>0；单位相同换算系数必须相同"
        .TextMatrix(17, 0) = "明细"
        .TextMatrix(17, 2) = "变价检查"
        .TextMatrix(17, 3) = "为空默认为定价，“√”表示时价；输入内容只能是“√”或空"
        .TextMatrix(18, 0) = "明细"
        .TextMatrix(18, 2) = "价格检查"
        .TextMatrix(18, 3) = "价格字段只能是数字型；精度不能超过最大设置精度"
        .TextMatrix(19, 0) = "明细"
        .TextMatrix(19, 2) = "效期"
        .TextMatrix(19, 3) = "必须是数字且只能是不小于0的整数"
        .TextMatrix(20, 0) = "明细"
        .TextMatrix(20, 2) = "收入项目"
        .TextMatrix(20, 3) = "不能含有非法字符；不能为空；只能是数据库已有收入项目"
        .TextMatrix(21, 0) = "明细"
        .TextMatrix(21, 2) = "门诊/住院分零"
        .TextMatrix(21, 3) = "只能是已有设置的分零方式；不能为空"
        .TextMatrix(22, 0) = "明细"
        .TextMatrix(22, 2) = "服务对象"
        .TextMatrix(22, 3) = "只能是已有数据库设置的服务对象"
        .TextMatrix(23, 0) = "明细"
        .TextMatrix(23, 2) = "分批属性"
        .TextMatrix(23, 3) = "药库分批时药房才能分批；输入内容只能是“√”或空；为空表示不分批"
        .TextMatrix(24, 0) = "明细"
        .TextMatrix(24, 2) = "供应商"
        .TextMatrix(24, 3) = "只能是数据库中已有的供应商"
        .TextMatrix(25, 0) = "明细"
        .TextMatrix(25, 2) = "日期格式"
        .TextMatrix(25, 3) = "日期格式：2015-10-10或者2015/10/10或者2015.10.10"
        .TextMatrix(26, 0) = "明细"
        .TextMatrix(26, 2) = "品种唯一性检查"
        .TextMatrix(26, 3) = "判断导入项目与界面现有数据或数据库中已有品种是否冲突"
        .TextMatrix(27, 0) = "明细"
        .TextMatrix(27, 2) = "规格唯一性检查"
        .TextMatrix(27, 3) = "判断导入项目与界面现有数据或数据库中已有规格是否冲突"
    End With
    Call GetCheck
End Sub

Private Sub GetCheck()
    Dim intRow As Integer
    
    mstrCheck = ""
    With vsfError
        For intRow = 1 To .Rows - 1
            mstrCheck = mstrCheck & "|" & .TextMatrix(intRow, 2)
        Next
    End With
End Sub
