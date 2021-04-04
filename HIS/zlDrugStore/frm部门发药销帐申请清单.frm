VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm部门发药销帐申请清单 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "销帐申请清单"
   ClientHeight    =   5760
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   7620
   Icon            =   "frm部门发药销帐申请清单.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7620
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk全部勾选 
      Caption         =   "全部勾选"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   5329
      Value           =   1  'Checked
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消发药"
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "继续发药"
      Height          =   350
      Left            =   5160
      TabIndex        =   3
      Top             =   5280
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   -240
      TabIndex        =   2
      Top             =   600
      Width           =   8292
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4452
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7332
      _cx             =   12933
      _cy             =   7853
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   12
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
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "[勾选]-允许该单据发药         [不勾选]-不允许该单据发药"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   5040
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "因以下单据存在未处理的【销帐申请】记录，需要手动进行处理！"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   732
      TabIndex        =   1
      Top             =   120
      Width           =   5220
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frm部门发药销帐申请清单.frx":6852
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frm部门发药销帐申请清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsData As Recordset        '用于接受父窗体的记录集
Private mbln允许发送 As Boolean

Private mint模式 As Integer         '0-部门发药;1-处方发药

Private mstrArray As String   '用于记录有数据的行

Private mIntCol标记 As Integer
Private mIntCol费用ID As Integer
Private mIntCol收发ID As Integer
Private mIntColNO As Integer
Private mIntCol药品名称 As Integer
Private mintcol批号 As Integer
Private mintcol数量 As Integer
Private mIntCol销帐申请数量 As Integer
Private mIntCol姓名 As Integer
Private mIntCol性别 As Integer
Private mIntCol年龄 As Integer
Private mIntCol领药部门 As Integer
Private mIntCol床号 As Integer
Private mIntCol病人科室 As Integer
Private Const mconIntCol列数 As Integer = 14

Public Sub ShowCard(FrmMain As Form, ByRef rsData As ADODB.Recordset, ByRef bln允许发送 As Boolean, Optional ByVal int模式 As Integer)
    Set mrsData = rsData
    mbln允许发送 = False
    mint模式 = int模式
    
    Me.Show vbModal, FrmMain
    '返回数据
    Set rsData = mrsData
    bln允许发送 = mbln允许发送
End Sub

Private Sub InitList()
    '初始化列头
    mIntCol标记 = 0
    mIntColNO = 1
    mIntCol药品名称 = 2
    mintcol批号 = 3
    mintcol数量 = 4
    mIntCol销帐申请数量 = 5
    mIntCol姓名 = 6
    mIntCol性别 = 7
    mIntCol年龄 = 8
    mIntCol领药部门 = 9
    mIntCol床号 = 10
    mIntCol病人科室 = 11
    mIntCol收发ID = 12
    mIntCol费用ID = 13
    
    With vsfList
        .Cols = mconIntCol列数
        .rows = 1
        
        .SelectionMode = flexSelectionByRow
        .AllowSelection = False
        .ColDataType(mIntCol标记) = flexDTBoolean
        
        VsfGridColFormat vsfList, mIntCol标记, "发送", 400, flexAlignCenterCenter, "标记"
        VsfGridColFormat vsfList, mIntColNO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList, mIntCol药品名称, "药品名称", 1500, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfList, mintcol批号, "批号", 1300, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList, mintcol数量, "数量", 800, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList, mIntCol销帐申请数量, "销帐申请数量", 1200, flexAlignRightCenter, "销帐申请数量"
        VsfGridColFormat vsfList, mIntCol姓名, "姓名", 800, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfList, mIntCol性别, "性别", 500, flexAlignLeftCenter, "性别"
        VsfGridColFormat vsfList, mIntCol年龄, "年龄", 500, flexAlignRightCenter, "年龄"
        
        If mint模式 = 0 Then
            VsfGridColFormat vsfList, mIntCol领药部门, "领药部门", 1000, flexAlignRightCenter, "领药部门"
            VsfGridColFormat vsfList, mIntCol床号, "床号", 500, flexAlignRightCenter, "床号"
            VsfGridColFormat vsfList, mIntCol病人科室, "病人科室", 1000, flexAlignLeftCenter, "病人科室"
        End If
        
        VsfGridColFormat vsfList, mIntCol收发ID, "收发ID", 0, flexAlignLeftCenter, "收发ID"
        VsfGridColFormat vsfList, mIntCol费用ID, "费用ID", 0, flexAlignLeftCenter, "费用ID"
        
        .ColHidden(mIntCol收发ID) = True
        .ColHidden(mIntCol费用ID) = True
        
        If mint模式 = 1 Then
            .ColHidden(mIntCol领药部门) = True
            .ColHidden(mIntCol床号) = True
            .ColHidden(mIntCol病人科室) = True
        End If
        
    End With
End Sub

Private Sub LoadList(ByVal rsData As ADODB.Recordset)
    Dim lngRow As Long
    Dim lng费用id As Long
    
    '---------加载数据---------
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        With vsfList
            .Redraw = flexRDNone
            
            Do While Not rsData.EOF
                '添加空数据隐藏行，防止不同的费用ID行的“销帐申请数量”上下合并
                If lng费用id <> rsData!费用ID Then
                    lngRow = lngRow + 2
                    .rows = lngRow + 1
                    '隐藏无数据列
                    .RowHidden(lngRow - 1) = True
                Else
                    lngRow = lngRow + 1
                    .rows = lngRow + 1
                End If
                
                .TextMatrix(lngRow, mIntCol标记) = True    '默认都勾选
                .TextMatrix(lngRow, mIntColNO) = zlCommFun.NVL(rsData!NO, "")
                .TextMatrix(lngRow, mIntCol药品名称) = zlCommFun.NVL(rsData!药品名称, "")
                .TextMatrix(lngRow, mintcol批号) = zlCommFun.NVL(rsData!批号, "")
                .TextMatrix(lngRow, mintcol数量) = rsData!数量
                .TextMatrix(lngRow, mIntCol销帐申请数量) = rsData!销帐申请数量
                .TextMatrix(lngRow, mIntCol姓名) = rsData!姓名
                .TextMatrix(lngRow, mIntCol性别) = zlCommFun.NVL(rsData!性别, "")
                .TextMatrix(lngRow, mIntCol年龄) = zlCommFun.NVL(rsData!年龄, "")
                
                If mint模式 = 0 Then
                    .TextMatrix(lngRow, mIntCol领药部门) = zlCommFun.NVL(rsData!领药部门, "")
                    .TextMatrix(lngRow, mIntCol床号) = zlCommFun.NVL(rsData!床号, "")
                    .TextMatrix(lngRow, mIntCol病人科室) = zlCommFun.NVL(rsData!病人科室, "")
                End If
                
                .TextMatrix(lngRow, mIntCol收发ID) = rsData!收发ID
                .TextMatrix(lngRow, mIntCol费用ID) = rsData!费用ID
                   
                lng费用id = rsData!费用ID
                
                rsData.MoveNext
            Loop
            
            .RowHeight(-1) = 300
            
            .MergeCells = flexMergeRestrictColumns
            .MergeCol(mIntCol销帐申请数量) = True
            
            .Redraw = flexRDDirect
        End With
    End If
    
End Sub

Private Sub chk全部勾选_Click()
    Dim i As Integer
    
    If chk全部勾选.Value = 1 Then
        vsfList.Cell(flexcpText, 1, mIntCol标记, vsfList.rows - 1, mIntCol标记) = True
    ElseIf chk全部勾选.Value = 0 Then
        vsfList.Cell(flexcpText, 1, mIntCol标记, vsfList.rows - 1, mIntCol标记) = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '统计不允许发药的数据,并更改执行状态
    Dim i As Integer
    
    If MsgBox("只有已勾选的单据才会被发药，请问是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    With vsfList
        For i = 1 To .rows - 1
            If .TextMatrix(i, mIntCol收发ID) <> "" Then
                If .TextMatrix(i, mIntCol标记) = False Then
                    mrsData.Filter = "收发ID =" & Val(.TextMatrix(i, mIntCol收发ID))
                    If mint模式 = 0 Then
                        mrsData!执行状态 = 3
                    ElseIf mint模式 = 1 Then
                        mrsData!标志 = 0
                    End If
                    
                    mrsData.Update
                End If
            End If
        Next
    End With
    
    mbln允许发送 = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitList
    Call LoadList(mrsData)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbln允许发送 = False Then
        If MsgBox("是否取消本次发药操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim blnHaveCheck As Integer      '至少有一行勾选
    Dim blnNoCheck As Integer       '至少有一行没有勾选
    Dim strNo As String
    
    With vsfList
        '【处方发药】模块因为不能拆分发药，所以同单据的勾选状态需要进行联动
        If mint模式 = 1 Then
            strNo = .TextMatrix(Row, mIntColNO)
            
            For i = 1 To .rows - 1
                If .TextMatrix(i, mIntColNO) = strNo Then
                    .TextMatrix(i, mIntCol标记) = .TextMatrix(Row, mIntCol标记)
                End If
            Next
        End If
        
        For i = 1 To .rows - 1
            If .TextMatrix(i, mIntCol收发ID) <> "" Then
                If .TextMatrix(i, mIntCol标记) = True Then
                    blnHaveCheck = blnHaveCheck + 1
                Else
                    blnNoCheck = blnNoCheck + 1
                End If
            End If
        Next
        
        If blnHaveCheck > 0 And blnNoCheck > 0 Then
            chk全部勾选.Value = 2
        ElseIf blnHaveCheck = 0 And blnNoCheck > 0 Then
            chk全部勾选.Value = 0
        ElseIf blnHaveCheck > 0 And blnNoCheck = 0 Then
            chk全部勾选.Value = 1
        End If
        
    End With
End Sub

Private Sub vsfList_EnterCell()
    vsfList.Editable = flexEDNone
    
    If vsfList.ColSel = mIntCol标记 Then
        vsfList.Editable = flexEDKbd
    End If
End Sub
