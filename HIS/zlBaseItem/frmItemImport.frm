VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmItemImport 
   Caption         =   "导入项目"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmItemImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9615
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOutput 
      Caption         =   "导出（&O）"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   10
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CheckBox chk供应商 
      Caption         =   "新的厂商自动增加"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox chkStop 
      Caption         =   "遇到错误终止导入"
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消（&C）"
      Height          =   350
      Left            =   8400
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "导入（&I）"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "合法性检查"
      Height          =   350
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "…"
      Height          =   300
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   280
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFList 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8175
      _cx             =   14420
      _cy             =   8705
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
      BackColorAlternate=   14737632
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemImport.frx":6852
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
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   225
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog dlgOutput 
      Left            =   720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin VB.Label lblFile 
      Caption         =   "文件"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   173
      Width           =   420
   End
   Begin VB.Label lblFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmItemImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTRCHARGETYPE As String = "上级分类,编码,名称"
Private Const mstrCharge As String = "类别,分类,编码,名称,是否变价,收入项目,现价,标识主码,标识子码,备选码,规格,计算单位,服务对象,费用类型"
Private Const MSTRSTUFFTYPE As String = "上级分类,编码,名称"
Private Const MSTRSTUFF As String = "分类,品种编码,品种名称,规格编码,规格,生产商,散装单位,包装单位,散装包装换算系数,是否变价,成本价,售价,收入项目,来源,服务对象,标识主码,标识子码,卫材库房分批,发料部门分批,效期(月),批准文号,产品注册商标,注册证号,许可证号,许可证效期,供应商名称,供应商许可证号,供应商许可证效期"
Private Const MSTRMEDICALTYPE As String = "类别,上级名称,编码,名称"
Private Const MSTRMEDICAL As String = "类别,分类,品种编码,品种名称,规格编码,药品规格,产地,剂型,剂量单位,售价单位,售价剂量换算系数,门诊单位,门诊单位换算系数,住院单位,住院单位换算系数,药库单位,药库包装换算系数,是否变价,成本价,售价,收入项目,住院可否分零,门诊可否分零,服务对象,药库分批,药房分批,效期(月),供应商名称,供应商许可证号,供应商许可证效期"
Private Const MINTTITLE As Integer = 2 '标题行
Private mobjWB As Object
Private mobjWS As Object
Private mobjWSType As Object
Private mobjXLS As Object
Private mRsError As Recordset
Private mintType As Integer
Private mLngCount As Long
Private mLngType As Long
Private mLngSumType As Long
Private mLngSumCount As Long
Private mCollTypeCols As Collection
Private mCollItemCols As Collection
Private mstrIn As String
Private mblnExists As Boolean

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim i As Long
    
    '清空表格内的数据
    Me.VSFList.Rows = 1
    Set mRsError = New Recordset
    
    '添加数据集的字段
    With mRsError
        If .State = 1 Then .Close

        .Fields.Append "Type", adVarChar, 200
        .Fields.Append "Error", adVarChar, 1000
        .Fields.Append "Row", adBigInt
        .Fields.Append "Col", adSmallInt
        .Fields.Append "Page", adVarChar, 20


        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    '检查表格格式
   If CheckExcel = False Then Exit Sub

    '检查分类信息
    CheckType mintType

    '检查项目信息
    '检查公共项目
    CheckPub
    '检查具体的项目
    If mintType = 1 Then
        CheckCharge
    ElseIf mintType = 2 Then
        CheckMedi
    Else
        CheckStuff
    End If
    
    '加载数据
    With mRsError
        If .RecordCount > 0 Then
            VSFList.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To .RecordCount
                Me.VSFList.RowHeight(i) = 300
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误类型")) = mRsError!Type
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误原因")) = mRsError!Error
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误行")) = mRsError!Row
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误列")) = mRsError!Col
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("类别")) = mRsError!Page
                .MoveNext
            Next
            
            Me.cmdCheck.Enabled = True
            Me.cmdImport.Enabled = False
            Me.cmdOutput.Enabled = True
        ElseIf .RecordCount = 0 Then
            Me.cmdCheck.Enabled = False
            Me.cmdImport.Enabled = True
            Me.cmdOutput.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdChoose_Click()
    OpenFile

    Me.cmdCheck.Enabled = Not (Me.lblFileName.Caption = "")
    Me.cmdImport.Enabled = False
End Sub

Private Sub cmdImport_Click()
    Dim strText As String
    Dim i As Long
    
    If cmdCheck.Enabled = True Then
        MsgBox "请先进行【合法性检查】操作！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    mLngType = 0
    mLngCount = 0
    
    Me.prg.Visible = True
    Me.VSFList.Rows = 1
    Set mRsError = New Recordset
    
    '添加数据集的字段
    With mRsError
        If .State = 1 Then .Close

        .Fields.Append "Type", adVarChar, 200
        .Fields.Append "Error", adVarChar, 1000
        .Fields.Append "Row", adBigInt
        .Fields.Append "Col", adSmallInt
        .Fields.Append "Page", adVarChar, 20
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    
    '保存分类
    SaveType
    '保存项目
    If mintType = 1 Then
        SaveData
    ElseIf mintType = 2 Then
        SaveMedi
    Else
        SaveStuff
    End If
    
    '加载数据
    With mRsError
        If .RecordCount > 0 Then
            VSFList.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To .RecordCount
                Me.VSFList.RowHeight(i) = 300
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误类型")) = mRsError!Type
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误原因")) = mRsError!Error
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误行")) = mRsError!Row
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("错误列")) = mRsError!Col
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("类别")) = mRsError!Page
                If mRsError!Page = "分类" Then
                    mobjWSType.Rows(Val(mRsError!Row)).Font.Color = vbRed
                Else
                    mobjWS.Rows(Val(mRsError!Row)).Font.Color = vbRed
                End If
                .MoveNext
            Next
            Me.cmdOutput.Enabled = True
        ElseIf .RecordCount = 0 Then
            Me.cmdCheck.Enabled = False
            Me.cmdOutput.Enabled = False
        End If
    End With
    
    '显示成功的数目
    With VSFList
        .Rows = .Rows + 1
        strText = "分类一共有" & mLngSumType & "条数据，成功导入" & mLngType & "条数据；细目一共有" & mLngSumCount & "条数据，成功导入" & mLngCount & "条数据！"
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = strText
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
        
        .MergeCells = flexMergeFree
        .MergeRow(.Rows - 1) = True
    End With
    
    Me.prg.Visible = False
End Sub

Private Sub SaveType()
    Dim strSql As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItemID As Long
    Dim strSqlTemp As String  '保存查询验证的sql语句
    Dim strTemp As String
    Dim blnStop As Boolean
    
    
    On Error Resume Next
    blnStop = (Me.chkStop.Value = 1)
    With mobjWSType.UsedRange
        For lngRow = 3 To .Rows.Count
            If mintType = 1 Then
                lngItemID = sys.NextId("收费分类目录")
                strSql = "ZL_收费分类目录_INSERT(" & lngItemID & ","
            Else
                lngItemID = sys.NextId("诊疗分类目录")
                strSql = "ZL_诊疗分类目录_INSERT(" & lngItemID & ","
            End If

            For lngCol = 1 To .Columns.Count
                If mintType = 2 And lngCol = 1 Then
                    If .cells(lngRow, lngCol) = "西成药" Then
                        strTemp = "1,"
                    ElseIf .cells(lngRow, lngCol) = "中成药" Then
                        strTemp = "2,"
                    ElseIf .cells(lngRow, lngCol) = "中草药" Then
                        strTemp = "3,"
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "类别的值只能为西成药，中成药或中草药", "分类"
                            
                            GoTo ErrHandle
                        End If
                    End If
                ElseIf (mintType = 2 And lngCol = 2) Or (mintType <> 2 And lngCol = 1) Then
                    If .cells(lngRow, lngCol) <> "" Then
                        If GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, IIf(mintType = 2, Val(strTemp), 7), False, True) <> 0 And .cells(lngRow, lngCol) <> "" Then
                            strSql = strSql & GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, IIf(mintType = 2, Val(strTemp), 7), False, True) & ","
                        Else
                            If blnStop Then
                                Exit Sub
                            Else
                                GoTo ErrHandle
                            End If
                        End If
                    Else
                        strSql = strSql & "null,"
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
            Next
            '简码
            strSql = strSql & "'" & zlStr.GetCodeByVB(.cells(lngRow, lngCol - 1)) & "',"

            '诊疗项目的类型
            If mintType = 2 Then
                strSql = strSql & strTemp
            ElseIf mintType = 3 Then
                strSql = strSql & "7,"
            End If

            strSql = strSql & "0)"

            zlDatabase.ExecuteProcedure strSql, "SaveType"
            If Err.Number <> 0 Then
                If blnStop Then
                    MsgBox "保存数据出错，导入终止", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                Else
                    AddErr "保存错误", lngRow, lngCol, Err.Description, "分类"
                    GoTo ErrHandle
                End If
            End If
        mLngType = mLngType + 1
ErrHandle:
        prg.Value = Int((lngRow - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub

Private Sub SaveData()
    Dim strSql As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItemID As Long
    Dim strSqlTemp As String  '保存查询验证的sql语句
    Dim strTemp As String
    Dim int特殊项目 As Integer
    Dim strID As String
    Dim lng价目ID As Long
    Dim str别名 As String
    Dim blnStop As Boolean
    Dim dateNow As Date
    Dim strNo As String
    Dim str编码 As String
    Dim rsTemp As Recordset
    Dim rs类别 As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng分类id As Long
    
    On Error Resume Next
    blnStop = chkStop.Value
    str编码 = ""
    
    '获取当前类别
    strSql = "Select 编码,名称 From 收费项目类别"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            lngCol = mCollItemCols.Item("类别")
'            mobjWS.cells(lngRow, lngCol) = "1111"
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "类别为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If .cells(lngRow, lngCol) = "挂号" Then
                    int特殊项目 = 1
                ElseIf .cells(lngRow, lngCol) = "护理" Then
                    int特殊项目 = 3
                Else
                    int特殊项目 = 0
                End If
                
                rsTemp.Filter = "名称='" & .cells(lngRow, lngCol) & "'"

                If Not rsTemp.EOF Then
                    strTemp = rsTemp!编码
                Else
                    '类别不存在的处理
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "类别不存在", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            strID = sys.NextId("收费项目目录")
            strSql = "zl_收费细目_insert(" & int特殊项目 & "," & strID & ",'" & strTemp & "',"
            
            '编码
            lngCol = mCollItemCols.Item("编码")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "编码为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '标识主码
            lngCol = mCollItemCols.Item("标识主码")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '标识子码
            lngCol = mCollItemCols.Item("标识子码")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '备选码
            lngCol = mCollItemCols.Item("备选码")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '名称
            lngCol = mCollItemCols.Item("名称")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "名称为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If


            '上级分类
            lngCol = mCollItemCols.Item("分类")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "上级分类为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                lng分类id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 0)
                If lng分类id <> 0 Then
                    strSql = strSql & lng分类id & ","
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        GoTo ErrHandle
                    End If
                End If
            End If

            '规格
            lngCol = mCollItemCols.Item("规格")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & Replace(.cells(lngRow, lngCol), "'", "''") & "',"
            End If

            '说明
            strSql = strSql & "'',"


            '计算单位
            lngCol = mCollItemCols.Item("计算单位")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '费用类型
            lngCol = mCollItemCols.Item("费用类型")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '屏蔽费别，是否变价，
            strSql = strSql & "0,"
            
            '是否变价
            lngCol = mCollItemCols.Item("是否变价")
            If .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                strSql = strSql & "0,"
            End If
    
            '加班加价，执行科室
            strSql = strSql & "0,0,"
            
            '服务对象
            lngCol = mCollItemCols.Item("服务对象")
            strSql = strSql & Val(.cells(lngRow, lngCol)) & ","

            '摘要
            strSql = strSql & "'',"

            '现价，原价
            lngCol = mCollItemCols.Item("现价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "现价为空", "明细"
                    
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    strSql = strSql & .cells(lngRow, lngCol) & ",0,"
                Else
                    If blnStop Then
                        
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "现价只能为数字", "明细"
                        
                        GoTo ErrHandle
                    End If
                End If
            End If

            '别名
            lngCol = mCollItemCols.Item("名称")
            If zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) <> "" Then
                str别名 = "1''" & .cells(lngRow, lngCol) & "''1''" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "''"
            End If

            If zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) <> "" Then
                str别名 = str别名 & "1''" & .cells(lngRow, lngCol) & "''2''" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "''"
            End If

            strSql = strSql & "'" & str别名 & "',"

            '录入限量，录入限量范围，费用确认，费用确认范围，自动计算,站点，病案费目
            strSql = strSql & "0,0,0,0,0,null,'')"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            '为材料更新产地
            If strTemp = "M" Then
                strSql = "ZL_收费细目_材料产地(" & strID & ",'" & Replace(.cells(lngRow, lngCol), "'", "''") & "')"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            End If

            '更新收费价目信息
            lng价目ID = sys.NextId("收费价目")
            lngCol = mCollItemCols.Item("现价")
            dateNow = sys.Currentdate
            strNo = sys.GetNextNo(9)
            strSql = ""
            strSql = "zl_收费价目_insert(" & lng价目ID & ",null," & strID & "," & GetTypeID(.cells(lngRow, mCollItemCols.Item("收入项目")), lngRow, mCollItemCols.Item("收入项目"), 0, True) & ",0," & .cells(lngRow, lngCol) & ",null,null,'初始定价',null,'" & gstrUserName & " ',to_date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss'),1,'" & strNo & "',1)"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "保存数据出错，导入终止", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        AddErr "保存错误", lngRow, lngCol, Err.Description, "明细"
                        gcnOracle.RollbackTrans
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
        

    End With
End Sub

Private Sub SaveMedi()
    Dim blnStop As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTemp As String
    Dim strSql As String
    Dim lng药名id As Long
    Dim lng药品ID As Long
    Dim lng收入id As Long
    Dim str供应商名称 As String
    Dim str供应商许可证号 As String
    Dim str供应商许可证效期 As String
    Dim intType As Integer
    Dim str编码 As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng分类id As Long
    Dim intKind As Integer
    
    On Error Resume Next
    blnStop = chkStop.Value
    str编码 = ""
    
    '获取品种编码
    strSql = "Select id,编码 From 诊疗项目目录 Where 类别 In ('5','6','7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    Do While Not rsTemp.EOF
        str编码 = str编码 & rsTemp!ID & "[" & rsTemp!编码 & "],"
        rsTemp.MoveNext
    Loop
    
    '获取供应商
    strSql = "Select 名称 From 供应商"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "CheckSupplier")
    
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            '类别
            lngCol = mCollItemCols.Item("品种编码")
            If InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") <= 0 Then
                lng药名id = sys.NextId("诊疗项目目录")
                str编码 = str编码 & lng药名id & "[" & .cells(lngRow, lngCol) & "],"
                lngCol = mCollItemCols.Item("类别")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "类别为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    If .cells(lngRow, lngCol) = "西成药" Then
                        strTemp = "5"
                        intKind = 1
                        strSql = "Zl_成药品种_Insert('" & strTemp & "',"
                    ElseIf .cells(lngRow, lngCol) = "中成药" Then
                        strTemp = "6"
                        intKind = 2
                        strSql = "Zl_成药品种_Insert('" & strTemp & "',"
                    ElseIf .cells(lngRow, lngCol) = "中草药" Then
                        strTemp = "7"
                        intKind = 3
                        strSql = "Zl_草药品种_Insert('" & strTemp & "',"
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "类别的值只能为西成药，中成药或中草药", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                End If

                '上级分类
                lngCol = mCollItemCols.Item("分类")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "上级分类为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    lng分类id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind)
                    If lng分类id <> 0 Then
                        strSql = strSql & lng分类id & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            GoTo ErrHandle
                        End If
                    End If
                End If
    
                '品种id
                strSql = strSql & lng药名id & ","

                '编码
                lngCol = mCollItemCols.Item("品种编码")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "品种编码为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                '名称
                lngCol = mCollItemCols.Item("品种名称")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "品种名称为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If

                '拼音简码
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "',"
    
                '五笔简码
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "',"
    
                '英文名称
                strSql = strSql & "'',"
    
                '剂量单位
                lngCol = mCollItemCols.Item("剂量单位")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "剂量单位为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If

                '剂型
                If strTemp <> "7" Then
                    lngCol = mCollItemCols.Item("剂型")
                    If .cells(lngRow, lngCol) = "" Then
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "空值错误", lngRow, lngCol, "剂型为空", "明细"
                            GoTo ErrHandle
                        End If
                    Else
                        strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                    End If
                End If
                
                '毒理分类,价值分类,货源情况,用药梯次
                strSql = strSql & "5,1,1,1)"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                lng药名id = Mid(Mid(str编码, 1, InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") - 1), InStrRev(Mid(str编码, 1, InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") - 1), ",") + 1)
            End If
            
            strSql = ""
            '保存药品规格
            If strTemp = "5" Then
                strSql = "Zl_成药规格_Insert(" & lng药名id & ","
            ElseIf strTemp = "6" Then
                strSql = "Zl_成药规格_Insert(" & lng药名id & ","
            Else
                strSql = "Zl_草药规格_Insert(" & lng药名id & ","
            End If

            '药品id
            lng药品ID = sys.NextId("收费项目目录")
            strSql = strSql & lng药品ID & ","

            '规格编码：药品品种编码后面连接1
            lngCol = mCollItemCols.Item("规格编码")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "规格编码为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '规格
            lngCol = mCollItemCols.Item("药品规格")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "药品规格为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '产地
            lngCol = mCollItemCols.Item("产地")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '商品名,拼音简码,五笔简码,数字码,标识码,药品来源
            strSql = strSql & "'','','','','','',"

            '批准文号
'            lngCol = mCollItemCols.Item("批准文号")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "'',"
'            Else
'
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            strSql = strSql & "'',"

            '注册商标
'            lngCol = mCollItemCols.Item("注册商标")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "'',"
'            Else
'
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            strSql = strSql & "'',"

            '售价单位
            lngCol = mCollItemCols.Item("售价单位")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价单位为空", "明细"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '售价剂量换算系数：剂量系数
            lngCol = mCollItemCols.Item("售价剂量换算系数")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价剂量换算系数为空", "明细"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If

            '门诊单位
            lngCol = mCollItemCols.Item("门诊单位")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "门诊单位为空", "明细"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

             '门诊单位换算系数：门诊包装
            lngCol = mCollItemCols.Item("门诊单位换算系数")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "门诊单位换算系数为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If
            
            If strTemp <> "7" Then
                '住院单位
                lngCol = mCollItemCols.Item("住院单位")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "住院单位为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
    
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                 '住院单位换算系数：住院包装
                lngCol = mCollItemCols.Item("住院单位换算系数")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "住院单位换算系数为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & .cells(lngRow, lngCol) & ","
                End If
            End If

            '药库单位
            lngCol = mCollItemCols.Item("药库单位")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "药库单位为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

             '药库单位换算系数：药库包装
            lngCol = mCollItemCols.Item("药库包装换算系数")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "药库包装换算系数为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If

            '申领单位,申领阀值
            strSql = strSql & "null,null,"

            '是否变价
            lngCol = mCollItemCols.Item("是否变价")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "是否变价的值只能为空或‘√’", "明细"
                    GoTo ErrHandle
                End If
            End If

            '指导批发价：成本价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "成本价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "成本价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "成本价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '扣率
            strSql = strSql & "100,"

            '指导零售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "售价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "售价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '指导差价率,管理费比例,药价级别,费用类型
            strSql = strSql & "13.0435,0,null,null,"

             '服务对象
            lngCol = mCollItemCols.Item("服务对象")
            
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 And Val(.cells(lngRow, lngCol)) <= 3 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "服务对象只能为0-3的数字或者为空", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "服务对象只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            'Gmp认证,招标药品,屏蔽费别,
            strSql = strSql & "0,0,0,"
            
            '住院可否分零
            lngCol = mCollItemCols.Item("住院可否分零")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            Else
                strSql = strSql & Mid(.cells(lngRow, lngCol), 1, 1) & ","
            End If

            '药库分批
            lngCol = mCollItemCols.Item("药库分批")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "药库分批的值只能为空或‘√’", "明细"
                    GoTo ErrHandle
                End If
            End If

            '药房分批
            lngCol = mCollItemCols.Item("药房分批")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "药房分批的值只能为空或‘√’", "明细"
                    GoTo ErrHandle
                End If
            End If

            '最大效期
            lngCol = mCollItemCols.Item("效期(月)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "效期(月)的值只能大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "效期(月)的值只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            Else
                strSql = strSql & "null,"
            End If

            '差价让利比
            strSql = strSql & "100,"

            '成本价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "成本价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "成本价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "成本价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "售价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "售价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '收入项目
            lngCol = mCollItemCols.Item("收入项目")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "收入项目为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                lng收入id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind, True)
                If lng收入id <> 0 Then
                    strSql = strSql & GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind, True)
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "收入项目不存在", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            If strTemp = "7" Then
                '合同单位id,说明,动态分零,发药类型,备选码,增值税率,基本药物,中药形态,站点,是否常备,病案费目
                strSql = strSql & ",Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,"
            Else
                '合同单位id,说明,动态分零,发药类型,备选码,增值税率,基本药物,站点,是否常备,存储温度,存储条件,配药类型,是否不予配置,容量,病案费目
                strSql = strSql & ",Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,"
            End If
            
            '门诊可否分零
            lngCol = mCollItemCols.Item("门诊可否分零")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0"
            Else
                strSql = strSql & Mid(.cells(lngRow, lngCol), 1, 1)
            End If
            
            
            strSql = strSql & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            '保存供应商
            lngCol = mCollItemCols.Item("供应商名称")
            If .cells(lngRow, lngCol) <> "" Then
                str供应商名称 = .cells(lngRow, lngCol)
                
                rsTemp.Filter = "名称='" & str供应商名称 & "'"
                
                If rsTemp.EOF Then
                    lngCol = mCollItemCols.Item("供应商许可证号")
                    str供应商许可证号 = .cells(lngRow, lngCol)
                    
                    lngCol = mCollItemCols.Item("供应商许可证效期")
                    str供应商许可证效期 = .cells(lngRow, lngCol)
                    
                    strSql = ""
                    intType = CheckSupplier(str供应商名称, str供应商许可证号, str供应商许可证效期, lngRow, lngCol, strSql)
                    If intType = 1 Then
                        Exit Sub
                    ElseIf intType = 2 Then
                        GoTo ErrHandle
                    End If
                    
                    If strSql <> "" Then
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = strSql
                    End If
                End If
            End If
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "保存数据出错，导入终止", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        gcnOracle.RollbackTrans
                        AddErr "保存错误", lngRow, lngCol, Err.Description, "明细"
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub

Private Function CheckSupplier(ByVal strVal As String, ByVal str许可证号 As String, ByVal str许可证效期 As String, ByVal lngRow As Long, ByVal lngCol As Long, ByRef strSql As String) As Integer
'检查产地是否存在，不存在是否新增
    Dim rsTemp As Recordset
    Dim lngId As Long
    Dim blnStop As Boolean
    Dim strTemp As String
    

    On Error Resume Next
    blnStop = (Me.chkStop.Value = 1)
    If chk供应商.Value = 1 Then

        strTemp = "Select max(编码) 编码,zlSpellCode([1], 10) 简码 From 供应商"
        Set rsTemp = zlDatabase.OpenSQLRecord(strTemp, "CheckSupplier", strVal)
        
        lngId = sys.NextId("供应商")
        strSql = "zl_供应商_insert("
        'id
        strSql = strSql & lngId & ","
        '上级id
        strSql = strSql & "null,"
        '编码
        If IsNull(rsTemp!编码) Then
            strSql = strSql & "'01',"
        Else
            strSql = strSql & "'" & Format(Val(rsTemp!编码) + 1, String(Len(Trim(rsTemp!编码)), "0")) & "',"
        End If
        '名称
        strSql = strSql & "'" & strVal & "',"
        
'            简码
        strSql = strSql & "'" & rsTemp!简码 & "',"
'            地址
        strSql = strSql & "null,"
'            电话
        strSql = strSql & "null,"
'            开户银行
        strSql = strSql & "null,"
'            帐号
        strSql = strSql & "null,"
'            联系人
        strSql = strSql & "null,"
'            税务登记号
        strSql = strSql & "null,"
'            许可证号
        strSql = strSql & "'" & str许可证号 & "',"
'            许可证效期
        strSql = strSql & IIf(str许可证效期 <> "", "to_date('" & str许可证效期 & "','YYYY-MM-dd'),", "Null,")
'            执照号
        strSql = strSql & "null,"
'            执照效期
        strSql = strSql & "null,"
'            授权号
        strSql = strSql & "null,"
'            授权期
        strSql = strSql & "null,"
'            类型
        strSql = strSql & IIf(mintType = 2, "'10000',", "'00001',")
        strSql = strSql & "0,0,Null,null,null,null,null,Null,Null,1,0)"
    Else
        If blnStop Then
            CheckSupplier = 1
        Else
            AddErr "供应商不存在", lngRow, lngCol, "当前供应商不存在", "明细"
            CheckSupplier = 2
        End If
    End If
End Function

Private Function GetTypeID(ByVal strVal As String, ByVal lngRow As Long, ByVal lngCol As Long, ByVal intType As Integer, Optional ByVal blnType As Boolean, Optional ByVal blnTypePage As Boolean) As Long
    Dim strSql As String
    Dim strType As String
    Dim strSecType As String
    Dim rsTemp As Recordset
    Dim Count As Long
    Dim i As Integer
    Dim strTemp As String

    On Error GoTo ErrHandle
    '检查分类是否只有一级
    If InStr(1, strVal, "\") > 1 Then
        strType = Mid(strVal, InStrRev(strVal, "\") + 1)
        strSecType = Mid(strVal, 1, InStrRev(strVal, "\") - 1)
    Else
        strType = strVal
        strSecType = ""
    End If

    If strSecType = "" Then
        '分类只有一级的情况
        strSql = "Select id,编码 " & _
                 "From 收费分类目录 " & _
                 "Where 名称 = [1] "
    Else
        strSql = "Select id,编码 From 收费分类目录" & vbNewLine & _
                "Where 名称 = [1] And 上级id In (Select 上级id From 收费分类目录 Where 名称 = [2] "
        Count = 2
        For i = UBound(Split(strSecType, "\")) - 1 To 0 Step -1
            If i < UBound(Split(strSecType, "\")) Then
                Count = Count + 1
                strTemp = "And 上级id In (Select 上级id From 收费分类目录 Where 名称 = '" & Split(strSecType, "\")(i) & "'"
                strSql = strSql & strTemp
            End If
        Next
        strSecType = Split(strSecType, "\")(UBound(Split(strSecType, "\")))
        strSql = strSql & String(Count - 1, ")")
    End If

    If mintType <> 1 And blnType = False Then
        strSql = strSql & " and 类型=[3]"
        strSql = Replace(strSql, "收费分类目录", "诊疗分类目录")
    ElseIf blnType = True Then
        strSql = Replace(strSql, "收费分类目录", "收入项目")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, strSecType, intType)

    If Not blnType Then
        If rsTemp.RecordCount = 1 Then
            GetTypeID = rsTemp!ID
        ElseIf rsTemp.RecordCount = 0 Then
            GetTypeID = 0
            AddErr "上级分类不存在", lngRow, lngCol, "上级分类【" & strVal & "】不存在", IIf(blnTypePage, "分类", "明细")
        Else
            GetTypeID = rsTemp!ID
            AddErr "上级分类不唯一", lngRow, lngCol, "上级分类【" & strVal & "】有多个，默认在编码为【" & rsTemp!编码 & "】的分类下", IIf(blnTypePage, "分类", "明细")
        End If
    Else
        If rsTemp.RecordCount = 1 Then
            GetTypeID = rsTemp!ID
        ElseIf rsTemp.RecordCount = 0 Then
            GetTypeID = 0
            AddErr "收入项目不存在", lngRow, lngCol, "收入项目【" & strVal & "】不存在", "明细"
        Else
            GetTypeID = rsTemp!ID
            AddErr "收入项目不唯一", lngRow, lngCol, "收入项目【" & strVal & "】有多个，默认在编码为【" & rsTemp!编码 & "】的收入项目下", "明细"
        End If
    End If
    Exit Function
ErrHandle:
    '收集错误信息
End Function

Private Sub SaveStuff()
    Dim blnStop As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTemp As String
    Dim strSql As String
    Dim lng诊疗ID As Long
    Dim lng材料ID As Long
    Dim lng收入id As Long
    Dim str供应商名称 As String
    Dim str供应商许可证号 As String
    Dim str供应商许可证效期 As String
    Dim intType As Integer
    Dim str编码 As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng分类id As Long
    
    On Error Resume Next
    blnStop = chkStop.Value
    str编码 = ""
    
    '获取品种编码
    strSql = "Select id,编码 From 诊疗项目目录 Where 类别 ='4'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    Do While Not rsTemp.EOF
        str编码 = str编码 & rsTemp!ID & "[" & rsTemp!编码 & "],"
        rsTemp.MoveNext
    Loop
    
    '获取供应商
    strSql = "Select 名称 From 供应商"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "CheckSupplier")
    
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            lngCol = mCollItemCols.Item("品种编码")
            If InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") <= 0 Then
                lng诊疗ID = sys.NextId("诊疗项目目录")
                str编码 = str编码 & lng诊疗ID & "[" & .cells(lngRow, lngCol) & "],"
                '卫材品种
                strSql = "zl_材料品种_INSERT("
    
                '上级分类
                lngCol = mCollItemCols.Item("分类")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "分类为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    lng分类id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 7)
                    If lng分类id <> 0 Then
                        strSql = strSql & lng分类id & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            GoTo ErrHandle
                        End If
                    End If
                End If
    
                '品种id
                strSql = strSql & lng诊疗ID & ","
    
                '品种编码
                lngCol = mCollItemCols.Item("品种编码")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "品种编码为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                '名称
                lngCol = mCollItemCols.Item("品种名称")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "品种名称为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
                
                '单位
                lngCol = mCollItemCols.Item("散装单位")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "空值错误", lngRow, lngCol, "散装单位为空", "明细"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                '拼音简码
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "',"
    
                '五笔简码
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "',"
    
                '英文名称
                strSql = strSql & "'',"
                
                '站点
                strSql = strSql & "null,"
                
                '适用性别
                strSql = strSql & "0,"
                
                '别名
                strSql = strSql & "'')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                lng诊疗ID = Mid(Mid(str编码, 1, InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") - 1), InStrRev(Mid(str编码, 1, InStr(1, str编码, "[" & .cells(lngRow, lngCol) & "]") - 1), ",") + 1)
            End If
            
            strSql = ""
            '卫材规格
            strSql = "Zl_卫生材料_Insert("
            
            '诊疗id
            strSql = strSql & lng诊疗ID & ","
            
            '材料id
            lng材料ID = sys.NextId("收费项目目录")
            strSql = strSql & lng材料ID & ","
            
            '编码:编码生成的规则
            lngCol = mCollItemCols.Item("规格编码")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "规格编码为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            
            '规格
            lngCol = mCollItemCols.Item("规格")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "规格为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '产地
            lngCol = mCollItemCols.Item("生产商")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '标识主码
            lngCol = mCollItemCols.Item("标识主码")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '标识子码
            lngCol = mCollItemCols.Item("标识子码")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '备选码
             strSql = strSql & "null,"
'            lngCol = mCollItemCols.Item("备选码")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "null,"
'            Else
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            
            '材料来源
            strSql = strSql & "'',"
            
            '货源情况
            strSql = strSql & "'',"
            
            '散装单位
            lngCol = mCollItemCols.Item("散装单位")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "散装单位为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '包装单位
            lngCol = mCollItemCols.Item("包装单位")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "包装单位为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '散装包装换算系数
            lngCol = mCollItemCols.Item("散装包装换算系数")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "散装包装换算系数为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If
            
            '是否变价：√
            lngCol = mCollItemCols.Item("是否变价")
            If .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            ElseIf .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "0,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "是否变价的值只能为‘√’或者空", "明细"
                    GoTo ErrHandle
                End If
            End If
            
            '指导批发价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "成本价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "成本价的值必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "成本价的值只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '扣率
            strSql = strSql & "100,"
            
            '指导零售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "售价的值必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "售价的值只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '指导差价率
            strSql = strSql & "13.0435,"
            
            '费用类型
            strSql = strSql & "null,"
            
            '服务对象
            lngCol = mCollItemCols.Item("服务对象")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 And Val(.cells(lngRow, lngCol)) <= 3 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "服务对象只能为0-3的数字或者为空", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "服务对象只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '屏蔽费别
            strSql = strSql & "0,"
            
            '卫材库房分批
            lngCol = mCollItemCols.Item("卫材库房分批")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "卫材库房分批的值只能为空或‘√’", "明细"
                    GoTo ErrHandle
                End If
            End If

            '发料部门分批
            lngCol = mCollItemCols.Item("发料部门分批")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "√" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "值错误", lngRow, lngCol, "发料部门分批的值只能为空或‘√’", "明细"
                    GoTo ErrHandle
                End If
            End If

            '最大效期
            lngCol = mCollItemCols.Item("效期(月)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "效期(月)的值只能大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "效期(月)的值只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            Else
                strSql = strSql & "null,"
            End If
            
            
            '灭菌效期
            strSql = strSql & "Null,"
            
            '无菌性材料
            strSql = strSql & "0,"
            
            '一次性材料
            strSql = strSql & "0,"
            
            '原材料
            strSql = strSql & "0,"
            
            '差价让利比
            strSql = strSql & "100,"

            '成本价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "成本价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "成本价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "成本价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '跟踪在用
            strSql = strSql & "0,"
            
            '核算材料
            strSql = strSql & "0,"

            '售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "售价为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "值错误", lngRow, lngCol, "售价必须大于0", "明细"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值类型错误", lngRow, lngCol, "售价只能为数字", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '收入项目
            lngCol = mCollItemCols.Item("收入项目")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "空值错误", lngRow, lngCol, "收入项目为空", "明细"
                    GoTo ErrHandle
                End If
            Else
                lng收入id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 7, True)
                If lng收入id <> 0 Then
                    strSql = strSql & lng收入id & ","
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "值错误", lngRow, lngCol, "收入项目不存在", "明细"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '批准文号
            lngCol = mCollItemCols.Item("批准文号")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            
            '产品注册商标
            lngCol = mCollItemCols.Item("产品注册商标")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '注册证号
            lngCol = mCollItemCols.Item("注册证号")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '许可证号
            lngCol = mCollItemCols.Item("许可证号")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '许可证有效期
            lngCol = mCollItemCols.Item("许可证效期")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & CDate(.cells(lngRow, lngCol)) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '材质分类
            strSql = strSql & "'',"
                
            '存储条件
            strSql = strSql & "'',"
            
            '跟踪病人
            strSql = strSql & "0,"
            
            '站点
            strSql = strSql & "'')"
            '品名
            '拼音
            '五笔
            '增值税率
            '说明
            '高值材料
            '条码管理
            '病案费目
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            '保存供应商
            lngCol = mCollItemCols.Item("供应商名称")
            If .cells(lngRow, lngCol) <> "" Then
                str供应商名称 = .cells(lngRow, lngCol)
                rsTemp.Filter = "名称='" & str供应商名称 & "'"
                
                If rsTemp.EOF Then
                    lngCol = mCollItemCols.Item("供应商许可证号")
                    str供应商许可证号 = .cells(lngRow, lngCol)
                    
                    lngCol = mCollItemCols.Item("供应商许可证效期")
                    str供应商许可证效期 = .cells(lngRow, lngCol)
                    
                    strSql = ""
                    intType = CheckSupplier(str供应商名称, str供应商许可证号, str供应商许可证效期, lngRow, lngCol, strSql)
                    If intType = 1 Then
                        Exit Sub
                    ElseIf intType = 2 Then
                        GoTo ErrHandle
                    End If
                    
                    If strSql <> "" Then
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = strSql
                    End If
                End If
            End If
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "保存数据出错，导入终止", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        
                        AddErr "保存错误", lngRow, lngCol, Err.Description, "明细"
                        gcnOracle.RollbackTrans
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub


Private Sub cmdOutput_Click()
    On Error GoTo ErrHandle
    Me.VSFList.SaveGrid "C:\APPSOFT\附加文件\错误信息.xls", flexFileExcel, True
    MsgBox "错误信息导出成功！", vbInformation + vbCritical, gstrSysName
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    Me.lblFileName.Caption = mstrIn
    dlgOpenFile.FileName = mstrIn

    Me.VSFList.RowHeight(0) = 300
    VSFList.Cell(flexcpFontBold, 0, 0, 0, VSFList.Cols - 1) = True
    Exit Sub
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub OpenFile()
    dlgOpenFile.Filter = "xlsx|*.xlsx|xls|*.xls"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        lblFileName.Caption = dlgOpenFile.FileName
    End If
    
End Sub

Private Function CheckExcel() As Boolean
'检查表格格式
    
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String
    Dim strItem As String
    
    Set mCollTypeCols = New Collection
    Set mCollItemCols = New Collection
    If mintType = 1 Then
        strType = MSTRCHARGETYPE
        strItem = mstrCharge
    ElseIf mintType = 2 Then
        strType = MSTRMEDICALTYPE
        strItem = MSTRMEDICAL
    Else
        strType = MSTRSTUFFTYPE
        strItem = MSTRSTUFF
    End If
    
    On Error GoTo ErrHandle
    Set mobjWB = mobjXLS.Workbooks.Open(Me.lblFileName.Caption)
    mblnExists = True
    
    Set mobjWSType = mobjWB.Sheets(1)
    With mobjWSType.UsedRange
        '列名和列顺序检查
        If .Columns.Count <> UBound(Split(strType, ",")) + 1 Then
            MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "外部文件的列数不正确，请检查！", vbInformation, gstrSysName
            CheckExcel = False
            Exit Function
        End If

        For lngCol = 1 To .Columns.Count
            If .cells(MINTTITLE, lngCol) <> Split(strType, ",")(lngCol - 1) Then
                MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "外部文件列名或顺序不正确，请检查！", vbInformation, gstrSysName
                CheckExcel = False
                Exit Function
            End If
            Call mCollTypeCols.Add(lngCol, .cells(MINTTITLE, lngCol))
        Next
        
        mLngSumType = .Rows.Count - MINTTITLE
    End With
    
    Set mobjWS = mobjWB.Sheets(2)
    With mobjWS.UsedRange
        '列名和列顺序检查
        If .Columns.Count <> UBound(Split(strItem, ",")) + 1 Then
            MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "外部文件的列数不正确，请检查！", vbInformation, gstrSysName
            CheckExcel = False
            Exit Function
        End If

        For lngCol = 1 To .Columns.Count
            If .cells(MINTTITLE, lngCol) <> Split(strItem, ",")(lngCol - 1) Then
                MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "外部文件列名或顺序不正确，请检查！", vbInformation, gstrSysName
                CheckExcel = False
                Exit Function
            End If
            Call mCollItemCols.Add(lngCol, .cells(MINTTITLE, lngCol))
        Next
        
        mLngSumCount = .Rows.Count - MINTTITLE
    End With
    CheckExcel = True
    Exit Function
ErrHandle:
    MsgBox "请检查路径[" & Me.lblFileName.Caption & "]下的文件是否存在！", vbInformation, gstrSysName
    CheckExcel = False
End Function

Private Sub CheckPub()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long

    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
           '检查编码
            If mintType = 1 Then
                lngCol = mCollItemCols.Item("编码")
            Else
                lngCol = mCollItemCols.Item("品种编码")
            End If
            
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "编码为空值", "明细"
            Else
                '检查编码中是否含有非法字符
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "值错误", lngRow, lngCol, "编码中含有非法字符", "明细"
                    End If
                Next
            End If

            '检查名称
            If mintType = 1 Then
                lngCol = mCollItemCols.Item("名称")
            Else
                lngCol = mCollItemCols.Item("品种名称")
            End If
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "名称为空值", "明细"
            End If

            '检查服务对象
            lngCol = mCollItemCols.Item("服务对象")
            If .cells(lngRow, lngCol) <> "" Then
                If CDbl(.cells(lngRow, lngCol)) > 3 Then
                    AddErr "值错误", lngRow, lngCol, "服务对象的值超出范围", "明细"
                End If
            End If

            lngCol = mCollItemCols.Item("收入项目")
            '如果分类为空，则不用检查
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "收入项目为空", "明细"
            End If
        Next
    End With
End Sub

Private Sub CheckCharge()
'收费项目检查
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strName As String   '保存已经检查过的上级名称
    Dim strNotExit As String  '保存已经检查且不存在的上级名称
    Dim strType As String     '保存直接上级的名称
    Dim strSecType As String  '保存第二个上级的名称
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long
    '检查项目部分
     With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
        '检查类别是否存在
            lngCol = mCollItemCols.Item("类别")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "类别为空", "明细"
            End If
            
            
            '检查分类是否合理
            lngCol = mCollItemCols.Item("分类")
            '如果分类为空，则不用检查
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "分类为空", "明细"
            End If

            '检查现价
            lngCol = mCollItemCols.Item("现价")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "现价为空值", "明细"
            Else
                If Not IsNumeric(.cells(lngRow, lngCol)) Then
                    AddErr "值类型错误", lngRow, lngCol, "现价不是数字", "明细"
                ElseIf CDbl(.cells(lngRow, lngCol)) < 0 Then
                    AddErr "值错误", lngRow, lngCol, "现价不能为负数", "明细"
                End If
            End If

            '检查服务对象
            lngCol = mCollItemCols.Item("服务对象")
            If .cells(lngRow, lngCol) <> "" Then
                If CInt(.cells(lngRow, lngCol)) > 3 Then
                    AddErr "值错误", lngRow, lngCol, "服务对象的值超出范围", "明细"
                End If
            End If
        Next
    End With
End Sub

Private Sub CheckMedi()
'药品项目检查
'检查药品部分数据的合法性
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strName As String   '保存已经检查过的上级名称
    Dim strNotExit As String  '保存已经检查且不存在的上级名称
    Dim strType As String     '保存直接上级的名称
    Dim strSecType As String  '保存第二个上级的名称
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '检查药品部分
    '检查药品的目录和规格信息
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            '检查类别信息
            lngCol = mCollItemCols.Item("类别")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "类型为空", "明细"
            ElseIf .cells(lngRow, lngCol) <> "西成药" And .cells(lngRow, lngCol) <> "中成药" And .cells(lngRow, lngCol) <> "中草药" Then
                AddErr "值错误", lngRow, lngCol, "类型只能为西成药，中成药或者中草药", "明细"
            End If

            '检查分类信息
            lngCol = mCollItemCols.Item("分类")
            '如果分类为空，则不用检查
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "分类为空", "明细"
            End If
            
            '检查编码
            lngCol = mCollItemCols.Item("规格编码")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "规格编码为空值", "明细"
            Else
                '检查编码中是否含有非法字符
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "值错误", lngRow, lngCol, "规格编码中含有非法字符", "明细"
                    End If
                Next
            End If
            
            '检查剂型
            lngCol = mCollItemCols.Item("剂型")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "剂型为空值", "明细"
            End If

            '检查售价剂量换算系数
            lngCol = mCollItemCols.Item("售价剂量换算系数")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "售价剂量换算系数为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "售价剂量换算系数的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "售价剂量换算系数的值必须是数字", "明细"
                End If
            End If


            '检查门诊单位换算系数
            lngCol = mCollItemCols.Item("门诊单位换算系数")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "门诊单位换算系数为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "门诊单位换算系数的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "门诊单位换算系数的值必须是数字", "明细"
                End If
            End If

            '检查住院单位转换系数
            lngCol = mCollItemCols.Item("住院单位换算系数")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "住院单位转换系数为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "住院单位转换系数的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "住院单位转换系数的值必须是数字", "明细"
                End If
            End If

            '检查药库包装换算系数
            lngCol = mCollItemCols.Item("药库包装换算系数")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "药库包装换算系数为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "药库包装换算系数的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "药库包装换算系数的值必须是数字", "明细"
                End If
            End If

            '检查成本价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "成本价为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "成本价的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "成本价的值必须是数字", "明细"
                End If
            End If

            '检查售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "售价为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "售价的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "售价的值必须是数字", "明细"
                End If
            End If


            '检查效期
            lngCol = mCollItemCols.Item("效期(月)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "效期的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "效期的值必须是数字", "明细"
                End If
            End If
        Next
    End With
End Sub

Private Sub CheckStuff()
'卫材项目检查
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String     '保存直接上级的名称
    Dim strSecType As String  '保存第二个上级的名称
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '检查卫材项目部分
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            '检查分类信息
            lngCol = mCollItemCols.Item("分类")
            '如果分类为空，则不用检查
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "分类为空", "明细"
            End If
            
            '检查编码
            lngCol = mCollItemCols.Item("规格编码")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "规格编码为空值", "明细"
            Else
                '检查编码中是否含有非法字符
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "值错误", lngRow, lngCol, "规格编码中含有非法字符", "明细"
                    End If
                Next
            End If
            
            '检查规格
            lngCol = mCollItemCols.Item("规格")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "规格为空", "明细"
            End If
            
            '散装单位
            lngCol = mCollItemCols.Item("散装单位")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "散装单位为空", "明细"
            End If
            
            '包装单位
            lngCol = mCollItemCols.Item("包装单位")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "包装单位为空", "明细"
            End If
            
            '换算系数
            lngCol = mCollItemCols.Item("散装包装换算系数")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "散装包装换算系数为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "散装包装换算系数的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "散装包装换算系数的值必须是数字", "明细"
                End If
            End If
            
            '成本价
            lngCol = mCollItemCols.Item("成本价")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "成本价为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "成本价的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "成本价的值必须是数字", "明细"
                End If
            End If

            '检查售价
            lngCol = mCollItemCols.Item("售价")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "售价为空值", "明细"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "售价的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "售价的值必须是数字", "明细"
                End If
            End If


            '检查效期
            lngCol = mCollItemCols.Item("效期(月)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "值错误", lngRow, lngCol, "效期的值必须大于0", "明细"
                    End If
                Else
                    AddErr "值类型错误", lngRow, lngCol, "效期的值必须是数字", "明细"
                End If
            End If
        Next
    End With
End Sub


Private Sub CheckType(ByVal intType As Integer)
'分类项目检查
'intType：1-收费项目分类，2-药品项目分类，3-卫材项目分类
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String     '保存直接上级的名称
    Dim strSecType As String  '保存第二个上级的名称
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '检查分类部分
    With mobjWSType.UsedRange
        For lngRow = 3 To .Rows.Count
            If intType = 2 Then
                '检查类别信息
                lngCol = mCollTypeCols.Item("类别")
                If .cells(lngRow, lngCol) = "" Then
                    AddErr "空值错误", lngRow, lngCol, "类型为空", "分类"
                ElseIf .cells(lngRow, lngCol) <> "西成药" And .cells(lngRow, lngCol) <> "中成药" And .cells(lngRow, lngCol) <> "中草药" Then
                    AddErr "值错误", lngRow, lngCol, "类型只能为西成药，中成药或者中草药", "分类"
                End If
            End If

            '检查编码
            lngCol = mCollTypeCols.Item("编码")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "编码为空值", "分类"
            Else
                '检查编码中是否含有非法字符
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "值错误", lngRow, lngCol, "编码中含有非法字符", "分类"
                    End If
                Next
            End If

            '检查名称
            lngCol = mCollTypeCols.Item("名称")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "空值错误", lngRow, lngCol, "名称为空值", "分类"
            End If
        Next
    End With
End Sub


Public Sub ShowMe(ByVal intType As Integer, ByVal frmParent As Form)
    On Error Resume Next
    Set mobjXLS = CreateObject("Excel.Application")
    
    If mobjXLS Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    mobjXLS.DisplayAlerts = False
    mintType = intType
    
    If mintType = 1 Then
        mstrIn = "C:\APPSOFT\附加文件\收费目录"
    ElseIf mintType = 2 Then
        mstrIn = "C:\APPSOFT\附加文件\药品目录"
    Else
        mstrIn = "C:\APPSOFT\附加文件\卫材目录"
    End If
    
    Me.Show 1, frmParent
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.cmdCheck.Left = Me.Width - 400 - Me.cmdCheck.Width
    Me.lblFileName.Width = Me.cmdCheck.Left - Me.cmdChoose.Width - 550
    Me.cmdChoose.Left = Me.cmdCheck.Left - 500
    Me.VSFList.Width = Me.cmdChoose.Width + Me.cmdChoose.Left - Me.lblFile.Left
    Me.VSFList.Height = Me.Height - Me.cmdCheck.Height - 900
    Me.cmdCancle.Left = Me.cmdCheck.Left
    Me.cmdCancle.Top = Me.VSFList.Top + Me.VSFList.Height - Me.cmdCancle.Height
    
    Me.cmdOutput.Left = Me.cmdCheck.Left
    Me.cmdOutput.Top = Me.cmdCancle.Top - Me.cmdOutput.Height - 100
    
    Me.cmdImport.Left = Me.cmdCheck.Left
    Me.cmdImport.Top = Me.cmdOutput.Top - Me.cmdImport.Height - 100
    
    Me.chkStop.Left = Me.cmdCheck.Left
    Me.chkStop.Top = Me.cmdImport.Top - Me.chkStop.Height
    
    Me.chk供应商.Left = Me.cmdCheck.Left
    Me.chk供应商.Top = Me.chkStop.Top - Me.chk供应商.Height
    
    Me.prg.Top = Me.Height - Me.prg.Height - 550
    Me.prg.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '关闭Excel进程
    If mblnExists And Not mobjWB Is Nothing Then mobjXLS.ActiveWorkbook.SaveAs (Me.lblFileName.Caption)
    Set mobjWB = Nothing
    Set mobjWS = Nothing
    Set mobjWSType = Nothing
    mobjXLS.quit
    Set mobjXLS = Nothing
    Set mRsError = Nothing
    mblnExists = False
    mintType = 0
    mLngCount = 0
    mLngType = 0
    mLngSumType = 0
    mLngSumCount = 0
    
    Set mCollTypeCols = Nothing
    Set mCollItemCols = Nothing
    mstrIn = ""
End Sub

Private Sub AddErr(ByVal strType As String, ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, ByVal strPage As String)
    mRsError.AddNew
    mRsError!Type = strType
    mRsError!Row = lngRow
    mRsError!Col = lngCol
    mRsError!Error = strContent
    mRsError!Page = strPage
    mRsError.Update
End Sub

