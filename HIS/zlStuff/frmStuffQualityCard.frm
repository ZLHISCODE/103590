VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStuffQualityCard 
   Caption         =   "卫材质量管理编辑"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmStuffQualityCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10890
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   4275
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   10695
      _cx             =   18865
      _cy             =   7541
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
      BackColorBkg    =   15132390
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffQualityCard.frx":000C
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
   Begin VB.TextBox txtCheck 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   10
      Top             =   5430
      Width           =   1000
   End
   Begin VB.TextBox txtVerify 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6360
      TabIndex        =   9
      Top             =   5430
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   8
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   7
      Top             =   6240
      Width           =   1100
   End
   Begin VB.TextBox txt备注 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   6
      Top             =   5820
      Width           =   9975
   End
   Begin VB.ComboBox cboStock 
      Enabled         =   0   'False
      Height          =   300
      Left            =   585
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1920
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   2  'OFF
      Left            =   9360
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   1425
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "登记人"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   5490
      Width           =   540
   End
   Begin VB.Label lblVerify 
      AutoSize        =   -1  'True
      Caption         =   "处理人"
      Height          =   180
      Left            =   5640
      TabIndex        =   16
      Top             =   5490
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl备注 
      AutoSize        =   -1  'True
      Caption         =   "备注"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   360
   End
   Begin VB.Label txtCheckDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   2955
      TabIndex        =   14
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label lblCheckDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登记日期"
      Height          =   180
      Left            =   2100
      TabIndex        =   13
      Top             =   5475
      Width           =   720
   End
   Begin VB.Label txtVerifyDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   8730
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblVerifyDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "处理日期"
      Height          =   180
      Left            =   7875
      TabIndex        =   11
      Top             =   5475
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "库房"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8880
      TabIndex        =   4
      Top             =   202
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "材料质量管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmStuffQualityCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑状态 As Integer '1-新增 2-修改 3-处理 4-查看
Private mlng质量id As Long      '质量id,修改和查看状态下有值
Private mstrMatch As String     '匹配方式
Private mintUnit As Long        '0-散装单位，1-包装单位
Private mlng库房id As Long      '库房id
Private mblnChange As Boolean   '是否进行过编辑
Private mint库存检查 As Integer  '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mFMT As g_FmtString     '小数位数的格式串
Private ArrNum As Variant       '记录毁损数量
Private mblnUsableNum As Boolean '是否卫材填单下可用库存
Private mintcboIndex As Integer
Private mblnValidateEdit As Boolean

Public Sub ShowMe(ByVal int编辑状态 As Integer, ByVal fraPar As Form, ByVal lng库房ID As Long, ByVal lng质量id As Long, ByVal intUnit As Integer)
    mint编辑状态 = int编辑状态
    mlng库房id = lng库房ID
    mlng质量id = lng质量id
    mintUnit = intUnit
    
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(2, g_售价)
    End With
    
    Me.Show vbModal, fraPar
End Sub

'检查数据依赖性
Private Sub CheckDepend()
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    
    '获取可操作的库房性质编码
    strStock = "VKW"
    
    '检查当前人员所属科室是否为“卫材库”、“制剂室”、“发料部门”
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr([2],b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(InStr(1, gstrPrivs, ";所有库房;") > 0, "", " and a.id in (Select 部门id from 部门人员 where 人员id =[1])")
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, strStock, gstrNodeNo)
    
    If rsDepend.EOF Then
        MsgBox "没有设置卫材库性质的部门或不具备相关的权限,请查看部门管理或找系统管事员授权！", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Sub
    End If
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        .Text = frmStuffQualityList.cboStock.Text
        If .ListIndex = -1 Then .ListIndex = 0
        mintcboIndex = .ListIndex
        rsDepend.Close
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboStock_Click()
    mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To VSFDetail.Rows - 1
                If Val(VSFDetail.TextMatrix(i, VSFDetail.ColIndex("材料id"))) <> 0 Then
                    Exit For
                End If
            Next
            If i <> VSFDetail.Rows Then
                If MsgBox("如果改变库房，有可能要改变相应卫材的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理卫材单位改变
                    mintcboIndex = .ListIndex
                    VSFDetail.Rows = 1
                    VSFDetail.Rows = 2
                    VSFDetail.Row = 1
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
        mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        If ValiData = False Then Exit Sub
        
        If SaveCard = True Then
            MsgBox "保存成功！", vbInformation, gstrSysName
            VSFDetail.Rows = 1
            VSFDetail.Rows = 2
            VSFDetail.SetFocus
            VSFDetail.Row = 1
            VSFDetail.Col = 0
            mblnChange = False
            If mint编辑状态 = 2 Then Unload Me
            Exit Sub
        End If
    End If
    If mint编辑状态 = 3 Then
        If mlng质量id = 0 Then
            Exit Sub
        End If
        
        If SaveCheck(mlng质量id) = True Then
            MsgBox "处理成功！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Function SaveCheck(ByVal lng质量id As Long) As Boolean
    '处理单据
    Dim strVerifyDate As String
    Dim lngRow As Long
    Dim arrSQL As Variant
    Dim dblTemp As Double
    Dim dbl包装系数 As Double
    Dim blnTran As Boolean
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        arrSQL = Array()
        strVerifyDate = txtVerifyDate.Caption
        
        gstrSQL = "Zl_材料质量主表_Verify("
        '质量id
        gstrSQL = gstrSQL & lng质量id & ","
        '处理人
        gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
        '处理日期
        gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'))"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("材料id"))) <> 0 Then
                dbl包装系数 = Val(.TextMatrix(lngRow, .ColIndex("换算系数")))
                
                gstrSQL = "Zl_材料其他出库_Verify("
                '序号
                gstrSQL = gstrSQL & lngRow & ","
                'NO
                gstrSQL = gstrSQL & "'" & TxtNo.Text & "',"
                '库房id
                gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                '材料id
                gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("材料id"))) & ","
                '批次
                gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("批次"))) & ","
                '毁损数量
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))) * dbl包装系数, g_小数位数.obj_最大小数.数量小数)
                gstrSQL = gstrSQL & dblTemp & ","
                '成本价
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) / dbl包装系数, g_小数位数.obj_最大小数.成本价小数)
                gstrSQL = gstrSQL & dblTemp & ","
                '成本金额
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                gstrSQL = gstrSQL & dblTemp & ","
                '售价金额
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                gstrSQL = gstrSQL & dblTemp & ","
                '金额差
                dblTemp = Round((Val(.TextMatrix(lngRow, .ColIndex("零售价"))) - Val(.TextMatrix(lngRow, .ColIndex("成本价")))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                gstrSQL = gstrSQL & dblTemp & ","
                '入出类别ID
                gstrSQL = gstrSQL & "19,"
                '处理人
                gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
                '处理日期
                gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'),"
                '药库(药房)业务
                gstrSQL = gstrSQL & "1)"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
        Next
    End With
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngRow = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(lngRow)), "SaveCheck")
    Next
    gcnOracle.CommitTrans
    SaveCheck = True
    
    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCard() As Boolean
    '新增、修改保存
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strNo As String
    Dim lng质量id As Long
    Dim strCheckDate As String
    Dim dbl包装系数 As Double
    Dim blnTran As Boolean
    Dim arrSQL As Variant
    Dim dblTemp As Double
    
    On Error GoTo ErrHandle
    
    arrSQL = Array()
    
    With VSFDetail
        
        If TxtNo.Text = "" Then
            strNo = zldatabase.GetNextNo(74, cboStock.ItemData(cboStock.ListIndex)) '根据药品其他出库取NO
        Else
            strNo = TxtNo.Text
        End If
        
        If .Rows > 1 And .TextMatrix(1, .ColIndex("材料id")) <> "" Then
            If mint编辑状态 = 2 Then '修改
                gstrSQL = "Zl_材料质量主表_Delete(" & mlng质量id & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                
                gstrSQL = "zl_材料其他出库_Delete('" & strNo & "')"
            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            Else
                gstrSQL = "Select 材料质量主表_ID.Nextval as id From Dual"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "卫材质量管理")
                mlng质量id = rsTemp!Id
            End If
            
            strCheckDate = Format(txtCheckDate.Caption, "yyyy-mm-dd hh:mm:ss")
            
            '主表插入
            gstrSQL = "Zl_材料质量主表_Insert ("
            'Id
            gstrSQL = gstrSQL & mlng质量id & ","
            'No
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '登记人
            gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
            '登记日期
            gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
            '处理人
            gstrSQL = gstrSQL & "Null,"
            '处理日期
            gstrSQL = gstrSQL & "Null,"
            '备注
            gstrSQL = gstrSQL & IIf(txt备注.Text = "", "NULL", "'" & txt备注.Text & "'") & ")"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            '次表插入
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("材料id"))) <> 0 Then
                    dbl包装系数 = Val(.TextMatrix(lngRow, .ColIndex("换算系数")))
                    
                    gstrSQL = "Zl_材料质量记录_Insert ("
                    '质量id
                    gstrSQL = gstrSQL & mlng质量id & ","
                    '库房id
                    gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                    '材料id
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("材料id"))) & ","
                    '批次
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("批次"))) & ","
                    '批号
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("批号")) & "',"
                    '产地
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("产地")) & "',"
                    '成本价
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) / dbl包装系数, g_小数位数.obj_最大小数.成本价小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '零售价
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) / dbl包装系数, g_小数位数.obj_最大小数.零售价小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '毁损数量
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))) * dbl包装系数, g_小数位数.obj_最大小数.数量小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '毁损原因
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("毁损原因")) & "',"
                    '解决办法
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("解决办法")) & "',"
                    '供药单位id
                    gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(lngRow, .ColIndex("供应商"))) = "", "Null", Val(.TextMatrix(lngRow, .ColIndex("供药单位id")))) & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_材料其他出库_Insert("
                    '入出类别ID
                    gstrSQL = gstrSQL & "19,"
                    'NO
                    gstrSQL = gstrSQL & "'" & strNo & "',"
                    '序号
                    gstrSQL = gstrSQL & lngRow & ","
                    '库房id
                    gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                    '材料id
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("材料id"))) & ","
                    '批次
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("批次"))) & ","
                    '毁损数量
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))) * dbl包装系数, g_小数位数.obj_最大小数.数量小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '成本价
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) / dbl包装系数, g_小数位数.obj_最大小数.成本价小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '成本金额
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("成本价"))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '零售价
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) / dbl包装系数, g_小数位数.obj_最大小数.零售价小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '售价金额
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("零售价"))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '金额差
                    dblTemp = Round((Val(.TextMatrix(lngRow, .ColIndex("零售价"))) - Val(.TextMatrix(lngRow, .ColIndex("成本价")))) * Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))), g_小数位数.obj_最大小数.金额小数)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '填制人
                    gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
                    '填制日期
                    gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
                    '产地
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("产地")) & "',"
                    '批号
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("批号")) & "',"
                    '效期，灭菌效期
                    gstrSQL = gstrSQL & "Null,Null,"
                    '摘要
                    gstrSQL = gstrSQL & IIf(txt备注.Text = "", "NULL", "'" & txt备注.Text & "'") & ","
                    '外调价，外调单位,增值税率
                    gstrSQL = gstrSQL & "Null,Null,Null,"
                    '药库(药房)业务
                    gstrSQL = gstrSQL & "1,"
                    '质量id
                    gstrSQL = gstrSQL & mlng质量id & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
            Next
            
            blnTran = True
            gcnOracle.BeginTrans
            For lngRow = 0 To UBound(arrSQL)
                Call zldatabase.ExecuteProcedure(CStr(arrSQL(lngRow)), "SaveCard")
            Next
            gcnOracle.CommitTrans
            SaveCard = True
        Else
            SaveCard = False
            Exit Function
        End If
    End With

    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValiData() As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    
    '保存时数据检查
    With VSFDetail
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("材料id")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("毁损数量"))) = 0 Then
                MsgBox "第" & lngRow & "行数据毁损数量不能为0或空！", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("毁损数量")
                .SetFocus
                Exit Function
            End If
        Next
    End With
    ValiData = True
End Function

Private Sub Form_Load()
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Call CheckDepend
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "") '匹配方式
    mblnUsableNum = (Val(zldatabase.GetPara("卫材填单下可用库存", 100)) = 1) '卫材下可用库存
    
    If InStr(1, ";" & gstrPrivs & ";", ";查看成本价;") = 0 Then '权限“查看成本价”
        VSFDetail.ColHidden(VSFDetail.ColIndex("成本价")) = True
        VSFDetail.ColHidden(VSFDetail.ColIndex("成本金额")) = True
    End If
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        cboStock.Enabled = True
        txtCheck.Text = UserInfo.用户名
        txt备注.Enabled = True
        Me.Caption = "卫材质量管理修改"
        If mint编辑状态 = 1 Then
            Me.Caption = "卫材质量管理新增"
            txtCheckDate.Caption = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    If mint编辑状态 = 3 Then
        txtVerify.Text = UserInfo.用户名
        txtVerifyDate.Caption = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Me.Caption = "卫材质量管理处理"
    End If
    
    If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 4 Then
        Call initCard
        If mint编辑状态 = 3 Then
            CmdSave.Caption = "处理(&O)"
        End If
        If mint编辑状态 = 4 Then
            CmdSave.Visible = False
            Me.Caption = "卫材质量管理查阅"
        End If
    End If
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 1 Then RestoreWinState Me, App.ProductName, "卫材质量管理"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LblTitle
        .Left = 0
        .Top = 120
        .Width = Me.ScaleWidth
    End With
    
    With TxtNo
        .Move Me.ScaleWidth - .Width - 150
    End With
    lblNo.Move TxtNo.Left - lblNo.Width - 100
    
    With lblStore
        .Move 150, 720
    End With
    
    With cboStock
        .Move lblStore.Left + lblStore.Width + 50, lblStore.Top - 60
    End With
    
    With VSFDetail
        .Move lblStore.Left, lblStore.Top + lblStore.Height + 100, Me.ScaleWidth - lblStore.Left - 150, Me.ScaleHeight - .Top - txtVerify.Height - 1200
    End With
    
    lblCheck.Move VSFDetail.Left, VSFDetail.Top + VSFDetail.Height + 200
    txtCheck.Move lblCheck.Left + lblCheck.Width + 100, lblCheck.Top - 60
    lblCheckDate.Move txtCheck.Left + txtCheck.Width + 100, lblCheck.Top
    txtCheckDate.Move lblCheckDate.Left + lblCheckDate.Width + 100, txtCheck.Top
    
    If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
        lblVerify.Visible = True
        txtVerify.Visible = True
        lblVerifyDate.Visible = True
        txtVerifyDate.Visible = True
        
        lblVerify.Move txtCheckDate.Left + txtCheckDate.Width + 500, lblCheck.Top
        txtVerify.Move lblVerify.Left + lblVerify.Width + 100, txtCheck.Top
        
        lblVerifyDate.Move txtVerify.Left + txtVerify.Width + 100, lblVerify.Top
        txtVerifyDate.Move lblVerifyDate.Left + lblVerifyDate.Width + 100, txtVerify.Top
    End If
    
    lbl备注.Move lblCheck.Left, lblCheck.Top + lblCheck.Height + 200
    txt备注.Move txtCheck.Left, lbl备注.Top - 20, VSFDetail.Width - lbl备注.Left - 530
    
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 200, lbl备注.Top + lbl备注.Height + 280
    CmdSave.Move cmdCancel.Left - CmdSave.Width - 200, lbl备注.Top + lbl备注.Height + 280
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, "卫材质量管理"
        Exit Sub
    End If
    
    If MsgBox("数据可能已改变，但未保存，是否退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        mblnChange = False
        SaveWinState Me, App.ProductName, "卫材质量管理"
    End If
    
End Sub

Private Sub txt备注_GotFocus()
    zlControl.TxtSelAll txt备注
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As Boolean
    '功能：检查重复记录
    Dim i As Integer
    Dim str批次 As String
    Dim str材料ID As String
    
    With VSFDetail
        CheckRedo = False
        str材料ID = rsTemp!材料ID
        str批次 = IIf(IsNull(rsTemp!批次), "", rsTemp!批次)
        
        For i = 1 To .Rows - 1
            If str材料ID = .TextMatrix(i, .ColIndex("材料ID")) And str批次 = .TextMatrix(i, .ColIndex("批次")) And .TextMatrix(i, .ColIndex("材料ID")) <> "" Then
                If str材料ID <> .TextMatrix(.Row, .ColIndex("材料ID")) Then
                    MsgBox "[" & rsTemp!编码 & "]" & rsTemp!名称 & "，该材料列表中已存在！", vbInformation, gstrSysName
                    If .TextMatrix(.Row, .ColIndex("材料id")) = "" Then .TextMatrix(.Row, .ColIndex("材料信息")) = ""
                    CheckRedo = True
                    Exit For
                End If
            End If
        Next
        
    End With
End Function

Private Sub vsfDetail_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    
    With VSFDetail
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("材料id"))) = 0 Then
                    .RemoveItem lngRow
                    .Rows = .Rows + 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim dblTop As Double
    Dim dblLeft As Double
    Dim vRect As RECT
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        Select Case Col
            Case .ColIndex("材料信息")
                Set RecReturn = Frm材料选择器.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , , , , , gstrPrivs)
                If RecReturn.RecordCount = 0 Then Exit Sub
                If CheckRedo(RecReturn) = True Then Exit Sub
                
                Call SetColValue(.Row, RecReturn!材料ID, RecReturn!编码, RecReturn!名称, _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!批次), "", RecReturn!批次), _
                                IIf(IsNull(RecReturn!批号), "", RecReturn!批号), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                IIf(IsNull(RecReturn!供药单位ID), "", RecReturn!供药单位ID), IIf(IsNull(RecReturn!散装单位), "", RecReturn!散装单位), _
                                IIf(IsNull(RecReturn!包装单位), "", RecReturn!包装单位), IIf(IsNull(RecReturn!换算系数), "", RecReturn!换算系数), _
                                IIf(IsNull(RecReturn!售价), "", RecReturn!售价), IIf(IsNull(RecReturn!时价), "", RecReturn!时价), _
                                IIf(IsNull(RecReturn!在用分批), "", RecReturn!在用分批), IIf(IsNull(RecReturn!库房分批), "", RecReturn!库房分批), _
                                IIf(IsNull(RecReturn!可用数量), 0, RecReturn!可用数量))
                .Col = .ColIndex("供应商")
            Case .ColIndex("供应商")
                vRect = zlControl.GetControlRect(.hwnd) '获取位置
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
                          "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                          "  And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
                          "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("供应商")) = RecReturn!名称
                    .TextMatrix(Row, .ColIndex("供药单位id")) = RecReturn!Id
                End If
                .Col = .ColIndex("毁损数量")
            Case .ColIndex("毁损原因")
                vRect = zlControl.GetControlRect(.hwnd)
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select 编码 id,'' 上级ID,0 末级,编码,简码,名称 From 毁损发生原因 " & _
                          " Where 名称 is not null " & _
                          " order by ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "毁损原因", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("毁损原因")) = RecReturn!名称
                End If
                .Col = .ColIndex("解决办法")
            Case .ColIndex("解决办法")
                vRect = zlControl.GetControlRect(.hwnd)
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select 编码 id,'' 上级ID,0 末级,编码,简码,名称 From 毁损解决办法 " & _
                          " Where 名称 is not null " & _
                          " order by ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "解决办法", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("解决办法")) = RecReturn!名称
                End If
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsfDetail_DblClick()
    With VSFDetail
        If .Col = .ColIndex("材料信息") Or .Col = .ColIndex("供应商") Or .Col = .ColIndex("毁损数量") Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.TextMatrix(.Row, .Col)) * 2
        End If
    End With
End Sub

Private Sub VSFDetail_EnterCell()
    With VSFDetail
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim dblLeft As Single
    Dim dblTop As Single
    Dim vRect As RECT
    Dim RecReturn As Recordset
    Dim strKey As String
    
    With VSFDetail
        If KeyCode = vbKeyDelete Then
            If MsgBox("将删除此行，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Shift = 0
            Else
                .RemoveItem VSFDetail.Row
            End If
        ElseIf KeyCode = vbKeyReturn Then
            If Trim(.TextMatrix(.Row, .ColIndex("材料信息"))) = "" Then KeyCode = 0: Exit Sub
            
            If .Col = .ColIndex("材料信息") Then
                .Col = .ColIndex("供应商")
            ElseIf .Col = .ColIndex("供应商") Then
                .Col = .ColIndex("毁损数量")
            ElseIf .Col = .ColIndex("毁损数量") Then
                .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), mFMT.FM_数量)
                .TextMatrix(.Row, .ColIndex("成本金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("成本价"))) * Val(.TextMatrix(.Row, .Col)), mFMT.FM_金额)
                .TextMatrix(.Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("零售价"))) * Val(.TextMatrix(.Row, .Col)), mFMT.FM_金额)
                .Col = .ColIndex("毁损原因")
            ElseIf .Col = .ColIndex("毁损原因") Then
                .Col = .ColIndex("解决办法")
            ElseIf .Col = .ColIndex("解决办法") Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = .ColIndex("材料信息")
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("材料信息")
                End If
            Else
                If InStr(1, ";" & gstrPrivs & ";", ";查看成本价;") = 0 Then
                    If .Col = .ColIndex("单位") Then
                        .Col = .ColIndex("零售价")
                    ElseIf .Col = .ColIndex("零售价") Then
                        .Col = .ColIndex("售价金额")
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
                .Editable = flexEDNone
            Else
                If .Col = .ColIndex("材料信息") Or .Col = .ColIndex("供应商") Or .Col = .ColIndex("毁损数量") Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(.Col) = ""
                Else
                    .Editable = flexEDNone
                End If
            End If
        End If
    End With
End Sub

Private Sub SetColValue(ByVal intRow As Integer, ByVal lng材料ID As Long, ByVal str编码 As String, ByVal str名称 As String, _
                        ByVal str规格 As String, ByVal str批次 As String, ByVal str批号 As String, ByVal str产地 As String, _
                        ByVal str供药单位id As String, ByVal str散装单位 As String, ByVal str包装单位 As String, _
                        ByVal str换算系数 As String, ByVal str售价 As String, ByVal int是否变价 As Integer, _
                        ByVal int在用分批 As String, ByVal int库房分批 As String, ByVal dbl可用数量 As Double)
    '表格赋值
    Dim rsTemp As Recordset
    Dim str包装系数 As String
    Dim bln分批 As Boolean
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        .TextMatrix(intRow, .ColIndex("材料id")) = lng材料ID
        .EditText = "[" & str编码 & "]" & str名称
        .TextMatrix(intRow, .ColIndex("材料信息")) = "[" & str编码 & "]" & str名称
        .TextMatrix(intRow, .ColIndex("规格")) = str规格
        .TextMatrix(intRow, .ColIndex("产地")) = str产地
        .TextMatrix(intRow, .ColIndex("单位")) = IIf(mintUnit = 0, str散装单位, str包装单位)
        .TextMatrix(intRow, .ColIndex("供药单位id")) = str供药单位id
        .TextMatrix(intRow, .ColIndex("换算系数")) = str换算系数
        .TextMatrix(intRow, .ColIndex("批次")) = str批次
        .TextMatrix(intRow, .ColIndex("批号")) = str批号
        .TextMatrix(intRow, .ColIndex("是否变价")) = int是否变价
        .TextMatrix(intRow, .ColIndex("毁损数量")) = ""
        .TextMatrix(intRow, .ColIndex("供应商")) = ""
        .TextMatrix(intRow, .ColIndex("成本金额")) = ""
        .TextMatrix(intRow, .ColIndex("售价金额")) = ""
        .TextMatrix(intRow, .ColIndex("毁损原因")) = ""
        .TextMatrix(intRow, .ColIndex("解决办法")) = ""
        
        str包装系数 = IIf(mintUnit = 0, 1, str换算系数)
        
        '可用数量
        .TextMatrix(intRow, .ColIndex("可用数量")) = Format(Val(dbl可用数量) / Val(str包装系数), mFMT.FM_数量)
        
        '供应商
        gstrSQL = "select 名称 from 供应商 where substr(类型,5,1)=1 and id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "供应商查询", Val(str供药单位id))
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, .ColIndex("供应商")) = IIf(IsNull(rsTemp!名称), "", rsTemp!名称)
        End If
        
        '成本价
        gstrSQL = "Select Decode(Nvl(a.平均成本价, 0), 0, b.成本价, a.平均成本价) 成本价" & vbNewLine & _
                "From 药品库存 A, 材料特性 B" & vbNewLine & _
                "Where a.药品id = b.材料id And a.库房id = [1] And a.药品id = [2] And Nvl(a.批次, 0) = [3]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "成本价查询", Val(cboStock.ItemData(cboStock.ListIndex)), Val(lng材料ID), Val(str批次))
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, .ColIndex("成本价")) = Format(Val(rsTemp!成本价) * Val(str包装系数), mFMT.FM_成本价)
        End If
        
        '售价
        .TextMatrix(intRow, .ColIndex("零售价")) = Format(Val(str售价) * Val(str包装系数), mFMT.FM_零售价)
        If int是否变价 = 1 Then
            If int在用分批 = 0 Then
                If int库房分批 = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From 部门性质说明 " & _
                            "Where ((工作性质 Like '发料部门') Or (工作性质 Like '制剂室')) And 部门id = [1]"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "部门查询", Val(cboStock.ItemData(cboStock.ListIndex)))
                    If rsTemp.RecordCount = 0 Then
                        bln分批 = True
                    End If
                End If
            Else
                bln分批 = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(零售价,0) as 分批售价,nvl(实际金额,0)/实际数量 as 平均零售价" & _
                "   From 药品库存 " & _
                "   Where 库房id=[1]" & _
                "       and 药品id=[2]" & _
                "       and 性质=1 and 实际数量>0 and " & _
                "       nvl(批次,0)=[3]"
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "售价查询", Val(cboStock.ItemData(cboStock.ListIndex)), Val(lng材料ID), Val(str批次))
            If Not rsTemp.EOF Then
                If bln分批 = True Then
                    .TextMatrix(intRow, .ColIndex("零售价")) = Format(Val(rsTemp!分批售价) * Val(str包装系数), mFMT.FM_零售价)
                Else
                    .TextMatrix(intRow, .ColIndex("零售价")) = Format(Val(rsTemp!平均零售价) * Val(str包装系数), mFMT.FM_零售价)
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With VSFDetail
        If .Col = .ColIndex("材料信息") Or .Col = .ColIndex("供应商") Then
            .ColComboList(.Col) = "|..."
        ElseIf .Col = .ColIndex("毁损原因") Or .Col = .ColIndex("解决办法") Then
            .ColComboList(.Col) = "..."
        End If
    End With
End Sub

Private Sub VSFDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim strKey As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intOldRow As Integer
    Dim i As Integer
    Dim intRow As Integer
    Dim intPosition As Integer

    On Error GoTo ErrHandle
    
    With VSFDetail
        intOldRow = .Row
        strKey = UCase(Trim(.EditText))
        
        Select Case Col
            Case .ColIndex("材料信息")
                If KeyAscii <> vbKeyReturn Then Exit Sub
                If strKey = "" Then Exit Sub
                
                dblLeft = Me.Left + .Left + .CellLeft + 130
                dblTop = Me.Top + .Top + .CellTop + .CellHeight + 500
                If dblTop + 4300 > Screen.Height Then
                    dblTop = dblTop - .CellHeight - 3680
                End If
                
                If Mid(strKey, 1, 1) = "[" Then
                    If InStr(2, strKey, "]") <> 0 Then
                        strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                    Else
                        strKey = Mid(strKey, 2)
                    End If
                End If
                .TextMatrix(.Row, .Col) = strKey
                
                Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strKey, dblLeft, dblTop, .CellWidth, .CellHeight, , , , , , , , , , , , gstrPrivs)
                
                If RecReturn.RecordCount = 0 Then .TextMatrix(.Row, .ColIndex("材料信息")) = "": Exit Sub
                If CheckRedo(RecReturn) = True Then Exit Sub
                
                Call SetColValue(.Row, RecReturn!材料ID, RecReturn!编码, RecReturn!名称, _
                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!批次), "", RecReturn!批次), _
                            IIf(IsNull(RecReturn!批号), "", RecReturn!批号), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                            IIf(IsNull(RecReturn!供药单位ID), "", RecReturn!供药单位ID), IIf(IsNull(RecReturn!散装单位), "", RecReturn!散装单位), _
                            IIf(IsNull(RecReturn!包装单位), "", RecReturn!包装单位), IIf(IsNull(RecReturn!换算系数), "", RecReturn!换算系数), _
                            IIf(IsNull(RecReturn!售价), "", RecReturn!售价), IIf(IsNull(RecReturn!时价), "", RecReturn!时价), _
                            IIf(IsNull(RecReturn!在用分批), "", RecReturn!在用分批), IIf(IsNull(RecReturn!库房分批), "", RecReturn!库房分批), _
                            IIf(IsNull(RecReturn!可用数量), 0, RecReturn!可用数量))
                .Col = .ColIndex("供应商")
            Case .ColIndex("供应商")
                If KeyAscii <> vbKeyReturn Or mblnValidateEdit = False Then Exit Sub
                If strKey = "" Then .Col = .ColIndex("毁损数量"): Exit Sub
                .TextMatrix(.Row, .ColIndex("供应商")) = strKey
                
                vRect = zlControl.GetControlRect(.hwnd) '获取位置
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                          "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                          "  And 末级=1 And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
                          "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"

                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                                True, dblLeft, dblTop, 1000, False, False, True, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", gstrNodeNo)
                If RecReturn Is Nothing Then
                    MsgBox "未找到供应商“" & Trim(.EditText) & "”，请重新输入！", vbInformation, gstrSysName
                    .EditText = ""
                    .TextMatrix(.Row, .ColIndex("供应商")) = ""
                    .TextMatrix(.Row, .ColIndex("供药单位id")) = ""
                    Exit Sub
                Else
                    .EditText = RecReturn!名称
                    .TextMatrix(.Row, .ColIndex("供应商")) = RecReturn!名称
                    .TextMatrix(.Row, .ColIndex("供药单位id")) = RecReturn!Id
                    .Col = .ColIndex("毁损数量")
                End If
            Case .ColIndex("毁损数量")
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strKey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        
                        If .EditSelLength = Len(strKey) Then Exit Sub
                        If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= Len(Mid(mFMT.FM_数量, InStr(1, mFMT.FM_数量, ".") + 1)) And strKey Like "*.*" Then
                            KeyAscii = 0
                            Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                    
                    If InStr(1, strKey, "-") > 0 Then
                        If Chr(KeyAscii) = "-" Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        If Val(strKey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                    
                End If
        End Select
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsTemp As ADODB.Recordset
    Dim str包装系数 As String
    Dim dbl成本价 As Double
    Dim dbl零售价 As Double
    Dim dbl成本金额 As Double
    Dim dbl售价金额 As Double
    Dim dbl毁损数量 As Double
    Dim dbl可用数量 As Double
    
    On Error GoTo ErrHandle
    
    ArrNum = Array()
    
    gstrSQL = "Select e.No, c.材料id, b.编码, b.名称, b.规格, d.名称 As 供应商, a.批号, a.批次, a.产地, a.毁损原因, a.解决办法, " & IIf(mintUnit = 0, " b.计算单位", " c.包装单位") & " as 单位, " & vbNewLine & _
            " a.毁损数量, a.成本价, a.零售价,a.供药单位id, c.换算系数, e.登记人, e.登记日期, e.处理人, e.处理日期, e.备注, b.是否变价, f.可用数量 " & vbNewLine & _
            " From 材料质量记录 A, 收费项目目录 B, 材料特性 C, 供应商 D, 材料质量主表 E, 药品库存 F, 药品收发记录 H " & vbNewLine & _
            " Where a.质量id = e.Id And a.材料id = b.Id And b.Id = c.材料id And a.库房id=f.库房id And a.材料id=f.药品id And nvl(a.批次,0) = nvl(f.批次,0) " & vbNewLine & _
            " And e.No=h.No And h.单据=21 And a.库房id=h.库房id And a.材料id=h.药品id And nvl(a.批次,0) = nvl(h.批次,0) " & vbNewLine & _
            " And a.供药单位id = d.Id(+) And a.质量id = [1] " & vbNewLine & _
            " Order by h.序号 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng质量id)
    
    With VSFDetail
        Do While Not rsTemp.EOF
            TxtNo.Text = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("材料信息")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(.Rows - 1, .ColIndex("供应商")) = IIf(IsNull(rsTemp!供应商), "", rsTemp!供应商)
            .TextMatrix(.Rows - 1, .ColIndex("产地")) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(.Rows - 1, .ColIndex("单位")) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
            .TextMatrix(.Rows - 1, .ColIndex("毁损原因")) = IIf(IsNull(rsTemp!毁损原因), "", rsTemp!毁损原因)
            .TextMatrix(.Rows - 1, .ColIndex("解决办法")) = IIf(IsNull(rsTemp!解决办法), "", rsTemp!解决办法)
            .TextMatrix(.Rows - 1, .ColIndex("材料id")) = IIf(IsNull(rsTemp!材料ID), "", rsTemp!材料ID)
            .TextMatrix(.Rows - 1, .ColIndex("供药单位id")) = IIf(IsNull(rsTemp!供药单位ID), "", rsTemp!供药单位ID)
            .TextMatrix(.Rows - 1, .ColIndex("换算系数")) = IIf(IsNull(rsTemp!换算系数), "", rsTemp!换算系数)
            .TextMatrix(.Rows - 1, .ColIndex("批次")) = IIf(IsNull(rsTemp!批次), "", rsTemp!批次)
            .TextMatrix(.Rows - 1, .ColIndex("批号")) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(.Rows - 1, .ColIndex("是否变价")) = IIf(IsNull(rsTemp!是否变价), "", rsTemp!是否变价)
            
            str包装系数 = IIf(mintUnit = 0, 1, rsTemp!换算系数)
            
            If IsNull(rsTemp!毁损数量) = False Then dbl毁损数量 = Val(rsTemp!毁损数量) / Val(str包装系数)
            .TextMatrix(.Rows - 1, .ColIndex("毁损数量")) = Format(dbl毁损数量, mFMT.FM_数量)
            
            If IsNull(rsTemp!可用数量) = False Then dbl可用数量 = Val(rsTemp!可用数量) / Val(str包装系数)
            .TextMatrix(.Rows - 1, .ColIndex("可用数量")) = Format(dbl可用数量, mFMT.FM_数量)
            
            If IsNull(rsTemp!成本价) = False Then dbl成本价 = Val(rsTemp!成本价) * Val(str包装系数)
            .TextMatrix(.Rows - 1, .ColIndex("成本价")) = Format(dbl成本价, mFMT.FM_成本价)
            
            If IsNull(rsTemp!零售价) = False Then dbl零售价 = Val(rsTemp!零售价) * Val(str包装系数)
            .TextMatrix(.Rows - 1, .ColIndex("零售价")) = Format(dbl零售价, mFMT.FM_零售价)
            
            dbl成本金额 = Val(rsTemp!成本价) * Val(rsTemp!毁损数量)
            .TextMatrix(.Rows - 1, .ColIndex("成本金额")) = Format(dbl成本金额, mFMT.FM_金额)
            
            dbl售价金额 = Val(rsTemp!零售价) * Val(rsTemp!毁损数量)
            .TextMatrix(.Rows - 1, .ColIndex("售价金额")) = Format(dbl售价金额, mFMT.FM_金额)
            
            txtCheck.Text = IIf(IsNull(rsTemp!登记人), "", rsTemp!登记人)
            If IsNull(rsTemp!登记日期) = False Then
                txtCheckDate.Caption = Format(rsTemp!登记日期, "yyyy-mm-dd hh:mm:ss")
            End If
            
            If mint编辑状态 = 4 Then
                txtVerify.Text = IIf(IsNull(rsTemp!处理人), "", rsTemp!处理人)
                If IsNull(rsTemp!处理日期) = False Then
                    txtVerifyDate.Caption = Format(rsTemp!处理日期, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            txt备注.Text = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
            
            If mint编辑状态 = 2 And mblnUsableNum = True Then
                ReDim Preserve ArrNum(UBound(ArrNum) + 1)
                ArrNum(UBound(ArrNum)) = Val(.TextMatrix(.Rows - 1, .ColIndex("材料id"))) & "," & Val(.TextMatrix(.Rows - 1, .ColIndex("批次"))) & "," & dbl毁损数量
            End If
            
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        
        If .Rows > 1 Then
            .Row = 1
            .Col = 0
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With VSFDetail
            If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
                .Editable = flexEDNone
            ElseIf .Col = .ColIndex("材料信息") Or .Col = .ColIndex("供应商") Then
                .Editable = flexEDKbdMouse
                .ColComboList(.Col) = "|..."
            ElseIf .Col = .ColIndex("毁损原因") Or .Col = .ColIndex("解决办法") Then
                .Editable = flexEDKbdMouse
                .ColComboList(.Col) = "..."
            ElseIf .Col = .ColIndex("毁损数量") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
                .ColComboList(.Col) = ""
            End If
        End With
    End If
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    With VSFDetail
        Select Case .ColKey(Col)
            Case "材料信息"
                VSFDetail_KeyPressEdit Row, Col, vbKeyReturn
            Case "供应商"
                If Val(.TextMatrix(Row, .ColIndex("材料id"))) = 0 Then Exit Sub
                mblnValidateEdit = True
                VSFDetail_KeyPressEdit Row, Col, vbKeyReturn
                mblnValidateEdit = False
            Case "毁损数量"
                If Val(.TextMatrix(Row, .ColIndex("材料id"))) = 0 Then Exit Sub
                If Val(Trim(.EditText)) = 0 Then
                    .TextMatrix(.Row, .Col) = .EditText
                    MsgBox "毁损数量不能为0或空,且只能是数字！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                ElseIf Val(Trim(.EditText)) < 0 Then
                    If Val(.TextMatrix(.Row, .ColIndex("是否变价"))) <> 0 Or Val(.TextMatrix(.Row, .ColIndex("批次"))) <> 0 Then
                        .TextMatrix(.Row, .Col) = .EditText
                        MsgBox "该材料不能负数出库！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
                
                If Not CompareUsableQuantity(Row, Val(Trim(.EditText))) Then
                    Cancel = True
                    .TextMatrix(.Row, .Col) = .EditText
                    Exit Sub
                Else
                    .EditText = Format(Val(Trim(.EditText)), mFMT.FM_数量)
                    .TextMatrix(.Row, .Col) = Format(Val(Trim(.EditText)), mFMT.FM_数量)
                    .TextMatrix(.Row, .ColIndex("成本金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("成本价"))) * Val(Trim(.EditText)), mFMT.FM_金额)
                    .TextMatrix(.Row, .ColIndex("售价金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("零售价"))) * Val(Trim(.EditText)), mFMT.FM_金额)
                    .Col = .ColIndex("毁损原因")
                End If
        End Select
    End With
End Sub

Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double) As Boolean
    '与可用数量进行比较
    Dim dblUsableQuantity As Double '可用数量
    Dim dblOldNum As Double '修改前毁损数量
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim intLop As Integer
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False
    
    With VSFDetail
        If .TextMatrix(intRow, .ColIndex("材料id")) = "" Or Val(dbl填写数量) = 0 Then Exit Function
        
        If mint编辑状态 = 2 And mblnUsableNum = True Then
            For intLop = 0 To UBound(ArrNum)
                lng材料ID = Val(Split(ArrNum(intLop), ",")(0))
                lng批次 = Val(Split(ArrNum(intLop), ",")(1))
                If lng材料ID = Val(.TextMatrix(intRow, .ColIndex("材料id"))) And lng批次 = Val(.TextMatrix(intRow, .ColIndex("批次"))) Then
                    dblOldNum = Val(Split(ArrNum(intLop), ",")(2))
                    Exit For
                End If
            Next
        End If
        
        dblUsableQuantity = Val(.TextMatrix(intRow, .ColIndex("可用数量")))
        .TextMatrix(intRow, .ColIndex("毁损数量")) = dbl填写数量
        
        If Val(.TextMatrix(intRow, .ColIndex("批次"))) > 0 Or Val(.TextMatrix(intRow, .ColIndex("是否变价"))) = 1 Then '对移出库房是库房且卫材是分批核算的卫材的判断
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                If dbl填写数量 > dblUsableQuantity + dblOldNum Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + dblOldNum & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            .EditText = Format(dbl填写数量, mFMT.FM_数量)
            .TextMatrix(intRow, .ColIndex("毁损数量")) = Format(dbl填写数量, mFMT.FM_数量)
            CompareUsableQuantity = True
            Exit Function
        End If
        
        ' 不分批卫材
        
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                If dbl填写数量 > dblUsableQuantity + dblOldNum Then
                    If MsgBox("你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + dblOldNum & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                If dbl填写数量 > dblUsableQuantity + dblOldNum Then
                    MsgBox "你输入的数量“" & dbl填写数量 & "”大于了该卫材的可用库存数量“" & dblUsableQuantity + dblOldNum & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
        .TextMatrix(intRow, .ColIndex("毁损数量")) = Format(dbl填写数量, mFMT.FM_数量)
    End With
    
    CompareUsableQuantity = True
    
End Function
