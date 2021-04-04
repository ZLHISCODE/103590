VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm卫材未发料清单 
   BorderStyle     =   0  'None
   Caption         =   "未发料清单"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   285
      Left            =   2955
      TabIndex        =   5
      Top             =   4935
      Width           =   285
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   660
      MaxLength       =   20
      TabIndex        =   4
      Top             =   4920
      Width           =   2595
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   0
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4560
      Width           =   2625
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm卫材未发料清单.frx":0000
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发料打印单据格式"
      Height          =   180
      Index           =   1
      Left            =   3165
      TabIndex        =   2
      Top             =   4650
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      Caption         =   "配料人"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4965
      Width           =   615
   End
End
Attribute VB_Name = "frm卫材未发料清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset       '内部记录集:未发料部分
Private mrsChargeOff As New ADODB.Recordset                   '用于显示销帐申请记录
Private mbln发料前收费或审核 As Boolean
Private mbln允许未收费的门诊划价处方发料 As Boolean
Private mbln允许未审核的记账处方发料 As Boolean
Private mbln发料时汇总销账 As Boolean
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '过滤条件
Private mfrmMain As Form        '父窗口
Private mlng缺料检查 As Long
Private mbln按单据发料 As Boolean
Private mbln显示领退料人 As Boolean     '暂无
Private mintUnit As Integer     '显示单位
Private mbln领料人签名 As Boolean   '暂无
Private mrsMatStock As ADODB.Recordset      '存储库房
Private mstrNo As String                    '当前选择的NO
Private Const mstrAllType As String = "临床,护理,检查,检验,手术,治疗,营养"
Private mfrmFilter As New frm卫材发放过滤
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Enum mcboIdx
    idx_单据格式 = 0
End Enum
Private Enum mtxtIdx
    idx_配料人 = 0
End Enum
Private Enum mlblIdx
    idx_lbl配料人 = 0
    idx_lbl单据格式 = 1
End Enum
Private mblnHave发料 As Boolean     '是否存在发料项目
Private mblnHave拒发 As Boolean         '是否存在拒发料项目
Private mstr默认单据格式 As String '
Private mstr汇总标识号 As String
Private mbln按发生时间过滤 As Boolean

Public Event zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset)

Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private mobjPlugIn As Object             '外挂接口对象

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Private Function CheckIsStockUp(ByVal lng发料部门ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, _
        ByVal lng费用ID As Long, ByVal dbl实际数量 As Double) As Boolean
    '1、检查当前记录是否是备货卫材的发料记录：通过查找是否有对应的未审核的虚拟库房其他出库单据来判断
    '2、检查在虚拟库房的实际数量是否足够
    Dim rsData As ADODB.Recordset
    Dim lng虚拟库房id As Long
    
    On Error GoTo ErrHand
    gstrSQL = "Select a.库房id From 药品收发记录 A, 虚拟库房对照 B " & _
        " Where a.库房id + 0 = b.虚拟库房id And a.单据 = 21 And a.审核日期 Is Null And b.科室id = [1] And a.药品id + 0 = [2] " & _
        " And Nvl(a.批次, 0) = [3] And a.费用id = [4] And Rownum = 1 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsStockUp", lng发料部门ID, lng材料ID, lng批次, lng费用ID)
    
    If rsData.EOF Then
        CheckIsStockUp = False
        Exit Function
    Else
        lng虚拟库房id = rsData!库房ID
    End If
    
    gstrSQL = "Select nvl(实际数量,0) As 实际数量 From 药品库存 Where 性质=1 And 库房ID=[1] And 药品ID=[2] And nvl(批次,0)=[3]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsStockUp", lng虚拟库房id, lng材料ID, lng批次)
    
    If rsData.EOF Then
        CheckIsStockUp = False
        Exit Function
    ElseIf Val(rsData!实际数量) < dbl实际数量 Then
         CheckIsStockUp = False
         Exit Function
    End If
    
    CheckIsStockUp = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefreshChargeOffStation(ByVal rsStuffSendData As ADODB.Recordset, ByRef rsChargeOff As ADODB.Recordset)
    '根据发料记录的执行状态更新销账记录的执行状态
    
    rsStuffSendData.Filter = 0
    If rsStuffSendData.RecordCount = 0 Then Exit Sub
    
    rsChargeOff.Filter = 0
    If rsChargeOff.RecordCount = 0 Then Exit Sub
    
    rsChargeOff.MoveFirst
    Do While Not rsChargeOff.EOF
        '先将执行状态置为0
        rsChargeOff!执行状态 = 0
        rsChargeOff!审核标志 = 0
        
        rsStuffSendData.MoveFirst
        Do While Not rsStuffSendData.EOF
            '只要有一个对应的发料科室，材料ID的执行状态=1，则更新对应的销账记录执行状态=1
            If rsChargeOff!领料部门id = rsStuffSendData!科室id And rsChargeOff!材料ID = rsStuffSendData!材料ID And rsStuffSendData!执行状态 = 1 Then
                rsChargeOff!执行状态 = 1
                rsChargeOff!审核标志 = 1
                Exit Do
            End If
            
            rsStuffSendData.MoveNext
        Loop
        
        rsChargeOff.Update
        
        rsChargeOff.MoveNext
    Loop
End Sub

Private Sub GetChargeOffRecord(ByVal rsStuffSendData As ADODB.Recordset)
    '1.统计发料记录集中有哪些领料部门(科室)
    '2.查询销帐记录，查找审核部门是该库房对应的所有销账数据，以申请部门id和状态=0为条件，提取并记录申请部门ID，收费细目ID，申请时间，费用ID，销帐数量等关键信息
    '3.循环发料数据集，判断从1中找到的收费细目ID，申请部门ID，费用ID在发料数据集中是否存在，如果存在表示存在同时发料和销帐的情况
    '4.对3中找到的申请部门ID，收费细目id，申请时间再根据费用ID关联收发记录等表来组织退料销帐数据
    '5.在4中可能一个费用ID对应多个收发ID（不同批次），根据销账数量判断各自的准退数量是否足够，如果不够则分解到不同的收发ID（批次）上
    
    Dim rsTmp As ADODB.Recordset
    Dim rsChargeOffTmp As ADODB.Recordset
    Dim strDeptIDs As String
    Dim lngDeptId As Long
    Dim lngStuffID As Long
    Dim lngChargeID As Long
    Dim str包装单位 As String
    Dim dbl剩余销账数量 As Double
    Dim str申请时间 As String
    
    On Error GoTo ErrHandle
    
    If mbln发料时汇总销账 = False Then Exit Sub
    
    rsStuffSendData.Filter = "执行状态=1"
    If rsStuffSendData.RecordCount = 0 Then Exit Sub
    
    Set rsChargeOffTmp = New ADODB.Recordset
    With rsChargeOffTmp
        If .State = 1 Then .Close
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收费细目ID", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '1.领料部门汇总
    rsStuffSendData.Sort = "科室id"
    Do While Not rsStuffSendData.EOF
        If lngDeptId <> rsStuffSendData!科室id Then
            lngDeptId = rsStuffSendData!科室id
            strDeptIDs = IIf(strDeptIDs = "", "", strDeptIDs & ",") & lngDeptId
        End If
        rsStuffSendData.MoveNext
    Loop
    
    '2.查询销帐记录
    If InStr(strDeptIDs, ",") = 0 Then
        gstrSQL = "Select 费用id, 申请部门id, 收费细目id, 数量, 申请时间 " & _
            " From 病人费用销帐 " & _
            " Where 审核部门id = [1] And 申请类别 = 1 And 状态 = 0 And 申请部门id = [2] " & _
            " Order By 收费细目id, 申请部门id "
    Else
        gstrSQL = "Select a.费用id, a.申请部门id, a.收费细目id, a.数量, a.申请时间 " & _
            " From 病人费用销帐 A, Table(f_Str2list([2])) T " & _
            " Where A.审核部门id = [1] And A.申请类别 = 1 And A.状态 = 0 And A.申请部门id = t.Column_Value " & _
            " Order By a.收费细目id, a.申请部门id "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetChargeOffRecord", Val(mArrFilter("发料部门ID")), strDeptIDs)
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    '3.查找匹配的同时销账发料项目，如果找到则保存到临时销账记录集
    rsStuffSendData.Sort = "材料id,科室id"
    
    lngChargeID = 0
    lngStuffID = 0
    lngDeptId = 0
    
    Do While Not rsTmp.EOF
        lngStuffID = rsTmp!收费细目id
        lngDeptId = rsTmp!申请部门id
        
        rsStuffSendData.MoveFirst
        Do While Not rsStuffSendData.EOF
            If lngStuffID = rsStuffSendData!材料ID And lngDeptId = rsStuffSendData!科室id Then
                rsChargeOffTmp.AddNew
                rsChargeOffTmp!费用ID = rsTmp!费用ID
                rsChargeOffTmp!申请部门id = rsTmp!申请部门id
                rsChargeOffTmp!收费细目id = rsTmp!收费细目id
                rsChargeOffTmp!数量 = rsTmp!数量
                rsChargeOffTmp!申请时间 = Format(rsTmp!申请时间, "yyyy-mm-dd hh:mm:ss")
                rsChargeOffTmp.Update
            End If
            rsStuffSendData.MoveNext
        Loop
        
        rsTmp.MoveNext
    Loop
    
    If rsChargeOffTmp.RecordCount = 0 Then Exit Sub
    
    '4.组织销帐数据
    If mintUnit = 0 Then
        str包装单位 = ",x.计算单位 as 单位,1 as 包装 "
    Else
        str包装单位 = ",d.包装单位 as 单位,d.换算系数 as 包装 "
    End If
    
    gstrSQL = "Select Distinct '[' || x.编码 || ']' || x.名称 As 材料名称, c.Id As 收发id, c.药品id as 材料id, c.单据, c.No, c.序号 As 收发序号, c.产地, c.批号, c.效期," & vbNewLine & _
        "              f.险类, p.名称 As 开单科室, e.名称 As 领料部门, e.Id As 领料部门id, a.费用id, b.序号 As 费用序号, b.记录性质, b.主页id, b.病人id, a.申请时间," & vbNewLine & _
        "              c.实际数量 As 准退数量, a.数量 As 销帐数量" & str包装单位 & vbNewLine & _
        " From 病人费用销帐 A, 住院费用记录 B," & vbNewLine & _
        "     (Select a.Id, a.单据, a.No, a.序号, a.药品id, a.产地, a.批号, a.效期, a.费用id, b.实际数量" & vbNewLine & _
        "       From 药品收发记录 A," & vbNewLine & _
        "            (Select c.单据, c.No, c.序号, c.药品id, Sum(Nvl(c.付数, 1) * c.实际数量) As 实际数量" & vbNewLine & _
        "              From 药品收发记录 C, 病人费用销帐 A, 住院费用记录 B" & vbNewLine & _
        "              Where a.申请类别 = 1 And a.状态 = 0 And a.费用id = b.Id And b.No = c.No And b.Id = c.费用id And c.单据 In (24, 25) And" & vbNewLine & _
        "                    c.审核日期 Is Not Null And c.库房id = a.审核部门id And c.库房id = [1] And c.费用id = [2] And a.申请部门id = [3] And a.申请时间 = [4] " & vbNewLine & _
        "              Group By c.单据, c.No, c.序号, c.药品id" & vbNewLine & _
        "              Having Sum(Nvl(c.付数, 1) * c.实际数量) > 0) B" & vbNewLine & _
        "       Where a.No = b.No And a.单据 = b.单据 And a.药品id + 0 = b.药品id And a.序号 = b.序号 And a.审核人 Is Not Null And" & vbNewLine & _
        "             (a.记录状态 = 1 Or Mod(a.记录状态, 3) = 0)) C, 材料特性 D, 收费项目目录 X, 部门表 P, 病案主页 F, 部门表 E" & vbNewLine & _
        " Where a.申请类别 = 1 And a.状态 = 0 And a.费用id = b.Id And b.No = c.No And b.Id = c.费用id And b.开单部门id = p.Id And" & vbNewLine & _
        "      b.收费细目id = d.材料id And b.收费细目id = x.Id And b.病人id = f.病人id And b.主页id = f.主页id And a.申请部门id = e.Id And" & vbNewLine & _
        "      f.出院日期 Is Null And b.执行部门id = [1] And a.费用id = [2] And a.申请部门id = [3] And a.申请时间 = [4] " & vbNewLine & _
        "Order By a.申请时间, c.单据, c.No, c.序号 Desc"
    
    rsChargeOffTmp.MoveFirst
    Do While Not rsChargeOffTmp.EOF
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetChargeOffRecord", Val(mArrFilter("发料部门ID")), _
            Val(rsChargeOffTmp!费用ID), Val(rsChargeOffTmp!申请部门id), CDate(rsChargeOffTmp!申请时间))
        
        Do While Not rsTmp.EOF
            With mrsChargeOff
                .AddNew
                !材料名称 = rsTmp!材料名称
                !领料部门 = rsTmp!领料部门
                !领料部门id = rsTmp!领料部门id
                !单据 = rsTmp!单据
                !NO = rsTmp!NO
                !材料ID = rsTmp!材料ID
                !申请时间 = Format(rsTmp!申请时间, "yyyy-mm-dd hh:mm:ss")
                !病人ID = rsTmp!病人ID
                !收发序号 = rsTmp!收发序号
                !产地 = rsTmp!产地
                !批号 = rsTmp!批号
                !效期 = rsTmp!效期
'                !准退数量 = Format(rsTmp!准退数量 / rsTmp!包装, mFMT.FM_数量)
'                !销帐数量 = Format(rsTmp!销帐数量 / rsTmp!包装, mFMT.FM_数量)
                !准退数量 = rsTmp!准退数量
                !销帐数量 = rsTmp!销帐数量
                !包装 = rsTmp!包装
                !单位 = rsTmp!单位
                !收发ID = rsTmp!收发ID
                !主页id = IIf(IsNull(rsTmp!主页id), 0, rsTmp!主页id)
                !费用序号 = rsTmp!费用序号
                !险类 = rsTmp!险类
                !费用ID = rsTmp!费用ID
                !记录性质 = rsTmp!记录性质
                !审核标志 = 0
                !执行状态 = 0
    
                .Update
            End With
            
            rsTmp.MoveNext
        Loop
        
        rsChargeOffTmp.MoveNext
    Loop
    
    If mrsChargeOff.RecordCount = 0 Then Exit Sub
    
    '5.同一费用ID有多个收发ID的按准退数量和销帐数量进行分配
    lngChargeID = 0
    dbl剩余销账数量 = 0
    str申请时间 = ""
    mrsChargeOff.Sort = "费用id,申请时间,收发序号 Desc"
    mrsChargeOff.MoveFirst
    
    Do While Not mrsChargeOff.EOF
        If lngChargeID = mrsChargeOff!费用ID And str申请时间 = mrsChargeOff!申请时间 Then
            '表示是多个批次，按上个批次剩余销账数量再来分配
            If dbl剩余销账数量 > 0 Then
                If dbl剩余销账数量 - mrsChargeOff!准退数量 > 0 Then
                    '还有剩余，本次只能按准退数量分配
                    mrsChargeOff!销帐数量 = mrsChargeOff!准退数量
                    dbl剩余销账数量 = dbl剩余销账数量 - mrsChargeOff!准退数量
                Else
                    '没有剩余，按剩余数量分配
                    mrsChargeOff!销帐数量 = dbl剩余销账数量
                End If
            Else
                '表示上次分配后销账数量没有剩余，剩下的批次销账数量设为0
                mrsChargeOff!销帐数量 = 0
            End If
            mrsChargeOff.Update
        Else
            '按费用ID，申请时间开始新的销账数量分配
            lngChargeID = mrsChargeOff!费用ID
            str申请时间 = mrsChargeOff!申请时间
            
            dbl剩余销账数量 = mrsChargeOff!销帐数量 - mrsChargeOff!准退数量
            If dbl剩余销账数量 > 0 Then
                '表示有剩余，本次只能按准退数量分配
                mrsChargeOff!销帐数量 = mrsChargeOff!准退数量
                mrsChargeOff.Update
            End If
        End If
         
        mrsChargeOff.MoveNext
    Loop
    
    '6.更新执行状态
    Call RefreshChargeOffStation(rsStuffSendData, mrsChargeOff)
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetMatStock(ByVal lng库房ID As Long)
    On Error GoTo ErrHandle
    gstrSQL = "Select 收费细目id From 收费执行科室 Where 执行科室id = [1] "
    Set mrsMatStock = zlDatabase.OpenSQLRecord(gstrSQL, "取存储库房", lng库房ID)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsGrid()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始网格控件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-12 10:27:06
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        '0-可选,1-必选,-1-隐藏
        .ColData(.ColIndex("状态")) = 1
        .ColData(.ColIndex("单据类型")) = 1
        .ColData(.ColIndex("单据号")) = 1
        .ColData(.ColIndex("付数")) = 1
        .ColData(.ColIndex("数量")) = 1
    End With
End Sub

Private Function SaveChargeOffData(ByVal strDate As String) As Boolean
    '销账审核+退料
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln是否有退料 As Boolean
    Dim str序号数量 As String
    Dim cllPro As Collection
    Dim strAudit As String
    Dim strReturnInfo As String
    Dim strReserve As String
    
    If mbln发料时汇总销账 = False Then
        SaveChargeOffData = True
        Exit Function
    End If

    Set cllPro = New Collection
    
    With mrsChargeOff
        .Filter = "执行状态=1 And 销帐数量>0 "
        If .RecordCount = 0 Then
            SaveChargeOffData = True
            Exit Function
        End If
        
        Do While Not .EOF
            '排除重复的销账记录
            If InStr("," & strAudit & ",", "," & !费用ID & !申请时间 & ",") = 0 Then
                strAudit = IIf(strAudit = "", !费用ID & !申请时间, strAudit & "," & !费用ID & !申请时间)
                    
                'Zl_病人费用销帐_Audit
                gstrSQL = "Zl_病人费用销帐_Audit("
                '  Id_In       病人费用销帐.费用id%Type,
                gstrSQL = gstrSQL & "" & Val(NVL(!费用ID)) & ","
                '  申请时间_In 病人费用销帐.申请时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & !申请时间 & "','YYYY-MM-DD HH24:MI:SS'),"
                '  审核人_In   病人费用销帐.审核人%Type,
                gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                '  审核时间_In 病人费用销帐.审核时间%Type,
                gstrSQL = gstrSQL & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),"
                '  状态_In     病人费用销帐.状态%Type,
                gstrSQL = gstrSQL & "1,"
                '  int自动退料 Integer:=1
                gstrSQL = gstrSQL & "0)"
                AddArray cllPro, gstrSQL
            End If
            
            '退料处理
            'Zl_材料收发记录_部门退料
            gstrSQL = "Zl_材料收发记录_部门退料("
            '    收发id_In   In 药品收发记录.ID%Type,
            gstrSQL = gstrSQL & "" & NVL(!收发ID) & ","
            '    审核人_In   In 药品收发记录.审核人%Type,
            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
            '    审核日期_In In 药品收发记录.审核日期%Type,
            gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
            '    批号_In     In 药品库存.上次批号%Type := Null,
            gstrSQL = gstrSQL & "'" & NVL(!批号) & "',"
            '    效期_In     In 药品库存.效期%Type := Null,
            gstrSQL = gstrSQL & "" & IIf(IsNull(!效期), "NULL", IIf(NVL(!效期) = "", "NULL", "To_Date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
            '    产地_In     In 药品库存.上次产地%Type := Null,
            gstrSQL = gstrSQL & "'" & NVL(!产地) & "',"
            '    退料数量_In In 药品收发记录.实际数量%Type := Null,
            gstrSQL = gstrSQL & "" & NVL(!销帐数量) & ","
            '    自动销帐_In Integer := 0,
            gstrSQL = gstrSQL & "" & 0 & ","
            '    退料人_In   In 药品收发记录.领用人%Type := Null
            gstrSQL = gstrSQL & "'" & gstrUserName & "')"
            
            AddArray cllPro, gstrSQL
            
            bln是否有退料 = True
            
            strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!收发ID) & "," & NVL(!销帐数量)
            
            '销帐处理
            str序号数量 = !费用序号 & ":" & !销帐数量
            '--序号：格式如"1,3,5,7,8",或"1:2,3:2,5:2,7:2,8:2",冒号前面的数字表示行号,后面的数字表示退的数量,目前仅在销帐审核时非药品才传入
            '--      为空表示冲销所有可冲销行

            If !主页id = 0 Then
                gstrSQL = "Zl_门诊记帐记录_Delete('" & !NO & "','" & !费用序号 & "','" & gstrUserCode & "','" & gstrUserName & "')"
            Else
                gstrSQL = "ZL_住院记帐记录_Delete('" & !NO & "','" & str序号数量 & "','" & gstrUserCode & "','" & gstrUserName & "'," & !记录性质 & ",1)"
            End If
            AddArray cllPro, gstrSQL
            
            '医保处理
            If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                        "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
            End If

            .MoveNext
        Loop
    End With
    
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    '医保，记帐作废上传，作废时上传
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:  Exit Function
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    
    '医保，记帐作废上传，完成后上传
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        Next
    End If

    err = 0: On Error GoTo ErrHandRpt:
    If bln是否有退料 = True Then
      If zlStr.IsHavePrivs(mstrPrivs, "退料通知单") Then
            If MsgBox("你需要打印退料清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "退料时间=" & strDate, "单位=" & mintUnit + 1, 2)
            End If
     End If
    End If
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And bln是否有退料 Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("发料部门id")), strReturnInfo, CDate(strDate), strReserve
    End If
    
    SaveChargeOffData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHandRpt:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    SaveChargeOffData = True
End Function

Public Function zlFullData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充相关的未发料单据信息
    '入参:frmMain-父窗口
    '     strPrivs-权限串
    '     lngModule-模块号
    '     intUnit-显示单位(0-散装单位,1-包装单位)
    '     arrFilter-条件过滤
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    mlng缺料检查 = Val(zlDatabase.GetPara("缺料检查", glngSys, lngModule))
    mbln按单据发料 = (Val(zlDatabase.GetPara("按单据号发料", glngSys, lngModule, "0")) = 1)
    mbln领料人签名 = (Val(zlDatabase.GetPara("领料人签名", glngSys, mlngModule, 0)) = 1)
    mbln发料时汇总销账 = (Val(zlDatabase.GetPara("发料时汇总退料销帐记录", glngSys, mlngModule, 0)) = 1)
    mbln按发生时间过滤 = Val(zlDatabase.GetPara("卫材医嘱按发生时间过滤", glngSys, 1723, 0))
    
    '初始控件数据
    Call InitData
    '填充数据给未发料网格
    If RefreshData() = False Then Exit Function
    zlFullData = False
End Function
Public Function zlPayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:发放卫生材料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String
          
    mstrNo = ""
    If vsGrid.Row > 0 Then
        If vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("单据号")) <> "" And vsGrid.IsSubtotal(vsGrid.Row) = False Then
            mstrNo = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("单据号"))
        End If
    End If
    
    If ISValied() = False Then Exit Function
    
    strDate = Format(sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    If SaveData(strDate) = False Then Exit Function
    
    If SaveChargeOffData(strDate) = False Then Exit Function
    
    zlPayStuff = True
End Function


Private Function GetNext汇总标识号() As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取汇总标识号
    '入参:
    '出参:
    '返回:成功,返回标识号
    '编制:刘兴洪
    '日期:2008-04-23 14:20:49
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    GetNext汇总标识号 = sys.GetNextNo(20)
    Exit Function
End Function
Private Function CheckBillStruct() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化检查对象集
    '入参:
    '出参:
    '返回:成功,返回空记录集结构
    '编制:刘兴洪
    '日期:2008-04-23 14:41:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set CheckBillStruct = rsTemp
End Function

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查发料是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 14:25:36
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Integer, lng材料ID As Long, rsCheck As ADODB.Recordset
    Dim str序号 As String
    Dim intCardCount As Integer '本次发料需要刷卡次数
    Dim int执行状态 As Integer
    
    ISValied = False
    
    '检查配料人
    If Trim(txtEdit(mtxtIdx.idx_配料人).Tag) = "" Then
        MsgBox "配料人未输入或输入不正确，请输入配料人！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '先初始检查
    Set rsCheck = CheckBillStruct
    
    '检查执行库房
    With mrsNotPayStuff
        .Filter = ""
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If mrsMatStock Is Nothing Then
            GetMatStock Val(mArrFilter("发料部门ID"))
            
            If mrsMatStock.RecordCount = 0 Then
                MsgBox "未设置存储库房，不能发料,请检查！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mbln按单据发料 = True Then
            .Filter = "NO='" & mstrNo & "'"
        End If
        
        .Sort = "材料ID Asc"
        Do While Not .EOF
            If lng材料ID <> !材料ID Then
                If !执行状态 = 1 Then
                    mrsMatStock.Filter = "收费细目id=" & Val(!材料ID)
                    If mrsMatStock.EOF Then
                        MsgBox !材料名称 & "未设置存储库房，不能发料,请检查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    lng材料ID = !材料ID
                Else
                    lng材料ID = 0
                End If
            End If
            
            If !执行状态 = 1 Then
                '主要是提供检查速度，先处理内部数据集
                rsCheck.Filter = "单据标识='" & NVL(!NO) & "|" & NVL(!单据) & "'"
                If rsCheck.RecordCount <> 0 Then
                    rsCheck.Find "病人ID=" & Val(NVL(!病人ID))
                    If rsCheck.EOF Then rsCheck.AddNew
                Else
                    rsCheck.AddNew
                End If
                
                rsCheck!单据标识 = NVL(!NO) & "|" & NVL(!单据)
                rsCheck!病人ID = Val(NVL(!病人ID))
                rsCheck!记录性质 = Val(NVL(!记录性质))
                rsCheck!门诊标志 = Val(NVL(!门诊标志))
                str序号 = NVL(rsCheck!序号)
                If InStr(1, "," & str序号 & ",", "," & Val(NVL(!序号)) & ",") = 0 Then
                    If str序号 = "" Then
                        str序号 = Val(NVL(!序号))
                    Else
                        str序号 = str序号 & "," & Val(NVL(!序号))
                    End If
                    rsCheck!序号 = str序号
                End If
                rsCheck.Update
                rsCheck.Filter = 0
            End If
            
            .MoveNext
        Loop
    End With
    Dim strNo As String, lng单据 As Long, lng病人id As Long
    '检查单据,主要是检查处方是否已经结帐,病人是否已经出院，差对权限进行相关的检查
    With rsCheck
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识 & "|"
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng病人id = !病人ID
            str序号 = NVL(!序号)
            
            '检查结帐处方是否能发料
            If Check结帐处方(mstrPrivs, lng单据, strNo, str序号, Val(!记录性质), Val(!门诊标志)) = False Then Exit Function
            If Check出院病人(mstrPrivs, lng单据, strNo, Val(!记录性质), Val(!门诊标志), lng病人id) = False Then Exit Function
            .MoveNext
        Loop
    End With
    
    strNo = ""
    lng单据 = 0
    lng病人id = 0
    
    '一卡通消费检查
    If mbln发料前收费或审核 = True Then
        With mrsNotPayStuff
            '检查1：发料如果需要刷卡，那么一次发料只能一个病人刷卡消费
            .Filter = "执行状态=1 And 已收费=0 And 病人ID>0"
            .Sort = "病人ID"
            Do While Not .EOF
                If lng病人id = 0 Then
                    lng病人id = !病人ID
                End If
                If lng病人id <> !病人ID Then
                    MsgBox "不支持多个病人发料时进行刷卡消费。本次不能发料，请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
                .MoveNext
            Loop
            
            '检查2：发料时如果需要刷卡，必须是整个处方都是发料状态
            If lng病人id > 0 Then
                .Filter = "已收费=0 And 病人ID=" & lng病人id
                .Sort = "单据,NO,执行状态"
                Do While Not .EOF
                    If lng单据 <> !单据 And strNo <> !NO Then
                        lng单据 = !单据
                        strNo = !NO
                        int执行状态 = !执行状态
                    ElseIf int执行状态 <> !执行状态 Then
                        MsgBox "发料时刷卡消费必须整个处方材料一起发料。本次不能发料，请检查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    ISValied = True
End Function

Private Function CardConfirm(ByVal rsData As ADODB.Recordset) As Boolean
    '消费卡消费确认接口
    '如果是批量发料，并且包含多个病人，按病人多次调用刷卡消费接口
    '实际在之前已进行校验，如果包含多个病人需要刷卡消费，则禁止发料，所以这里应该不包含多个病人刷卡消费
    '暂时保留这种处理方式，可能以后会变动
    Dim lngCard病人ID As Long
    Dim strCardNo As String
        
    On Error GoTo ErrHand
    
    If mbln发料前收费或审核 = False Then
        CardConfirm = True
        Exit Function
    End If
        
    '注意传入的记录集是处方明细
    '收费单据
     rsData.Filter = "执行状态=1 And 记录性质=1 And 已收费=0"
     rsData.Sort = "病人ID,NO"
     Do While Not rsData.EOF
         If lngCard病人ID <> rsData!病人ID Then
             If strCardNo <> "" Then
                 '刷卡消费
                If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard病人ID, mfrmFilter.PatiCardID, False, 1, strCardNo) = False Then
                    Exit Function
                End If
             End If
             
             lngCard病人ID = rsData!病人ID
             strCardNo = rsData!NO
         Else
             If strCardNo = "" Then
                 strCardNo = rsData!NO
             ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                 strCardNo = strCardNo & "," & rsData!NO
             End If
         End If
         rsData.MoveNext
     Loop
     
     If strCardNo <> "" Then
        '刷卡消费
        If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard病人ID, mfrmFilter.PatiCardID, False, 1, strCardNo) = False Then
            Exit Function
        End If
     End If
    
    lngCard病人ID = 0
    strCardNo = ""
    
    '记账单据：只对门诊病人进行处理
    rsData.Filter = "执行状态=1 And 记录性质=2 And 已收费=0"
    rsData.Sort = "病人ID,NO"
    Do While Not rsData.EOF
        If rsData!门诊标志 = 1 Or rsData!门诊标志 = 4 Then
            If lngCard病人ID <> rsData!病人ID Then
                If strCardNo <> "" Then
                    '刷卡消费
                    If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard病人ID, mfrmFilter.PatiCardID, False, 2, strCardNo) = False Then
                        Exit Function
                    End If
                    strCardNo = ""
                End If
                
                lngCard病人ID = rsData!病人ID
                strCardNo = rsData!NO
            Else
                If strCardNo = "" Then
                    strCardNo = rsData!NO
                ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                    strCardNo = strCardNo & "," & rsData!NO
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    If strCardNo <> "" Then
        '刷卡消费
        If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard病人ID, mfrmFilter.PatiCardID, False, 2, strCardNo) = False Then
            Exit Function
        End If
    End If
    
    CardConfirm = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CardConfirm = False
End Function
Private Function SaveData(ByVal strDate As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:对指定的发料项目进行发料处理
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:48:06
    '-----------------------------------------------------------------------------------------------------------
    Dim str领料人 As String, lng病人id As Long, strID批次 As String
    Dim cllPro As Collection
    Dim strReserve As String
        
    SaveData = False
    err = 0: On Error GoTo ErrHand:
    mstr汇总标识号 = GetNext汇总标识号()
   
    Set cllPro = New Collection
    With mrsNotPayStuff
        .Filter = ""
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If mbln按单据发料 = True Then
            If mstrNo = "" Then Exit Function
            If MsgBox("你现在确定要对单据[" & mstrNo & "]进行发料操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("你现在确定要进行发料操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        '新的消费卡刷卡消费接口
        If Not CardConfirm(mrsNotPayStuff) Then Exit Function
        
        '领药人签名
        str领料人 = ""
        If mbln领料人签名 Then
            str领料人 = zlDatabase.UserIdentify(Me, "领料人签名", glngSys, mlngModule, "")
            If str领料人 = "" Then
                Exit Function
            End If
        End If
        
        If mbln按单据发料 = True Then
            .Filter = "NO='" & mstrNo & "'"
        Else
            .Filter = ""
        End If
        
        '按病人ID，材料ID排序
        .Sort = "病人ID Asc ,材料ID Asc"
        
        Do While Not .EOF
            If !执行状态 = 1 Then
                            
                If lng病人id = 0 Then
                    lng病人id = !病人ID
                End If
                '病人ID相同时候
                If lng病人id = !病人ID Then
                    '如果传入的字符串大于3950时就提交事务（最大字符串为4000）
                    If zlCommFun.ActualLen(strID批次) > 3990 Then
                        'Zl_药品收发记录_批量发料
                        gstrSQL = "Zl_药品收发记录_批量发料("
                        '    收发id_In     In Varchar2, --格式:"id1,批次1|id2,批次2|....."
                        gstrSQL = gstrSQL & "'" & strID批次 & "',"
                        '    库房id_In     In 药品收发记录.库房id%Type,
                        gstrSQL = gstrSQL & "" & Val(mArrFilter("发料部门id")) & ","
                        '    审核人_In     In 药品收发记录.审核人%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '    审核日期_In   In 药品收发记录.审核日期%Type,
                        gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                        '    发料方式_In   In 药品收发记录.发药方式%Type := 3, --1-处方发料;2-批量发料;3-部门发料;-1 停止发料
                        gstrSQL = gstrSQL & "3,"
                        '    领料人_In     In 药品收发记录.领用人%Type := Null,
                        gstrSQL = gstrSQL & "'" & str领料人 & "',"
                        '    发料标识号_In In 药品收发记录.汇总发药号%Type := Null,
                        gstrSQL = gstrSQL & "" & Val(mstr汇总标识号) & ","
                        '    配料人_In     In 药品收发记录.配药人%Type := Null
                        gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_配料人).Text & "',"
                        '    操作员编码
                        gstrSQL = gstrSQL & "'" & UserInfo.编号 & "')"
                        Call AddArray(cllPro, gstrSQL)
                        lng病人id = 0
                        strID批次 = !Id & "," & NVL(!批次, 0)
                    Else
                        strID批次 = IIf(strID批次 = "", !Id & "," & NVL(!批次, 0), strID批次 & "|" & !Id & "," & NVL(!批次, 0))
                    End If
                Else
                    '如果病人ID不同则提交事务
                    'Zl_药品收发记录_批量发料
                    gstrSQL = "Zl_药品收发记录_批量发料("
                    '    收发id_In     In Varchar2, --格式:"id1,批次1|id2,批次2|....."
                    gstrSQL = gstrSQL & "'" & strID批次 & "',"
                    '    库房id_In     In 药品收发记录.库房id%Type,
                    gstrSQL = gstrSQL & "" & Val(mArrFilter("发料部门id")) & ","
                    '    审核人_In     In 药品收发记录.审核人%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    审核日期_In   In 药品收发记录.审核日期%Type,
                    gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                    '    发料方式_In   In 药品收发记录.发药方式%Type := 3, --1-处方发料;2-批量发料;3-部门发料;-1 停止发料
                    gstrSQL = gstrSQL & "3,"
                    '    领料人_In     In 药品收发记录.领用人%Type := Null,
                    gstrSQL = gstrSQL & "'" & str领料人 & "',"
                    '    发料标识号_In In 药品收发记录.汇总发药号%Type := Null,
                    gstrSQL = gstrSQL & "" & mstr汇总标识号 & ","
                    '    配料人_In     In 药品收发记录.配药人%Type := Null
                    gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_配料人).Text & "',"
                    '    操作员编码
                    gstrSQL = gstrSQL & "'" & UserInfo.编号 & "')"
                    Call AddArray(cllPro, gstrSQL)
                    lng病人id = !病人ID
                    strID批次 = !Id & "," & NVL(!批次, 0)
                End If
            End If
            .MoveNext
            
            '如果后面没有记录并且传入字符串不为空，则提交事务
            If .EOF And strID批次 <> "" Then
                    'Zl_药品收发记录_批量发料
                    gstrSQL = "Zl_药品收发记录_批量发料("
                    '    收发id_In     In Varchar2, --格式:"id1,批次1|id2,批次2|....."
                    gstrSQL = gstrSQL & "'" & strID批次 & "',"
                    '    库房id_In     In 药品收发记录.库房id%Type,
                    gstrSQL = gstrSQL & "" & Val(mArrFilter("发料部门id")) & ","
                    '    审核人_In     In 药品收发记录.审核人%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    审核日期_In   In 药品收发记录.审核日期%Type,
                    gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                    '    发料方式_In   In 药品收发记录.发药方式%Type := 3, --1-处方发料;2-批量发料;3-部门发料;-1 停止发料
                    gstrSQL = gstrSQL & "3,"
                    '    领料人_In     In 药品收发记录.领用人%Type := Null,
                    gstrSQL = gstrSQL & "'" & str领料人 & "',"
                    '    发料标识号_In In 药品收发记录.汇总发药号%Type := Null,
                    gstrSQL = gstrSQL & "" & mstr汇总标识号 & ","
                    '    配料人_In     In 药品收发记录.配药人%Type := Null
                    gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_配料人).Text & "',"
                    '    操作员编码
                    gstrSQL = gstrSQL & "'" & UserInfo.编号 & "')"
                    Call AddArray(cllPro, gstrSQL)
            End If
        Loop
    End With
        
    On Error GoTo ErrExcute:
    Call ExecuteProcedureArrAy(cllPro, Me.Caption)
    SaveData = True
    err = 0: On Error GoTo ErrHand:
    Call BillListPrint(strDate)
    mstrNo = ""
    
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        mobjPlugIn.StuffSendBySumID Val(mArrFilter("发料部门id")), mstr汇总标识号, strReserve
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrExcute:
      gcnOracle.RollbackTrans
      If ErrCenter = 1 Then Resume
      Call SaveErrLog
End Function

Private Sub SetExecuteStaut(ByVal lngRow As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置执行状态
    '入参:lngRow-指定的行
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-04-23 11:31:04
    '-----------------------------------------------------------------------------------------------------------
    Dim str状态 As String, int状态 As Integer, lng位置 As Long
    With vsGrid
        str状态 = Trim(.TextMatrix(lngRow, .ColIndex("状态")))
        int状态 = Decode(str状态, "缺料", 0, "发料", 1, "拒发", 2, "不处理", 3, 4)
        lng位置 = Val(.Cell(flexcpData, lngRow, .ColIndex("单据号")))
    End With
    
    With mrsNotPayStuff
         .Filter = 0
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        .Find "位置=" & lng位置
        If .EOF = False Then
            !执行状态 = int状态:
            !状态 = str状态
            .Update
            Call CheckStock(Val(NVL(!材料ID)))
        End If
        
        '可能库存检查后，需要处理当前的状态
        .MoveFirst
        .Find "位置=" & lng位置
        If .EOF = False Then
             vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("状态")) = Decode(NVL(!执行状态), 0, "缺料", 1, "发料", 2, "拒发", "不处理")
        End If
        .MoveFirst
        .Find "执行状态=1"
        mblnHave发料 = (.EOF = False)
        .MoveFirst
        .Find "执行状态=2"
        mblnHave拒发 = (.EOF = False)
    End With
    
    Call RefreshChargeOffStation(mrsNotPayStuff, mrsChargeOff)
End Sub
  
Private Sub cmdSel_Click()
   Call SelectItem(txtEdit(mtxtIdx.idx_配料人), "")
End Sub

Private Sub Form_Load()
    '刘兴宏:增加小数格式化串
     
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "未发料"
    
    mstr默认单据格式 = Trim(zlDatabase.GetPara("发料单据打印格式", glngSys, mlngModule, , Array(cboEdit(mcboIdx.idx_单据格式)), zlStr.IsHavePrivs(mstrPrivs, "参数设置")))

    mbln发料前收费或审核 = zlDatabase.GetPara("项目执行前必须先收费或先记帐审核", glngSys)
    mbln允许未收费的门诊划价处方发料 = Val(zlDatabase.GetPara("允许未收费的门诊划价处方发料", glngSys))
    mbln允许未审核的记账处方发料 = Val(zlDatabase.GetPara("允许未审核的记账处方发料", glngSys))
            
    Call InitVsGrid
    vsGrid.RowHeightMin = 300
    
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With txtEdit(mtxtIdx.idx_配料人)
        .Left = lblEdit(mlblIdx.idx_lbl配料人).Left + lblEdit(mlblIdx.idx_lbl配料人).Width
        cmdSel.Left = .Left + .Width - cmdSel.Width - 10
    End With
    With cboEdit(mcboIdx.idx_单据格式)
        .Top = ScaleHeight - .Height - 50
        .Left = ScaleWidth - .Width - 50
        lblEdit(mlblIdx.idx_lbl单据格式).Left = .Left - lblEdit(mlblIdx.idx_lbl单据格式).Width - 10
        lblEdit(mlblIdx.idx_lbl单据格式).Top = .Top + (.Height - lblEdit(mlblIdx.idx_lbl单据格式).Height) \ 2
        lblEdit(mlblIdx.idx_lbl配料人).Top = lblEdit(mlblIdx.idx_lbl单据格式).Top
        txtEdit(mtxtIdx.idx_配料人).Top = .Top
        cmdSel.Top = .Top + (txtEdit(mtxtIdx.idx_配料人).Height - cmdSel.Height) \ 2
    End With
    
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = cboEdit(mcboIdx.idx_单据格式).Top - .Top - 50
    End With
End Sub

Private Function InitRsStruct() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化内部记录集
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 09:54:46
    '-----------------------------------------------------------------------------------------------------------
    Set mrsNotPayStuff = New ADODB.Recordset
    With mrsNotPayStuff
        If .State = 1 Then .Close
        .Fields.Append "科室", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "开单医生", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "医生嘱托", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "状态", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "换算系数", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "在用分批", adDouble, 2, adFldIsNullable
        .Fields.Append "付", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "金额", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "记帐员", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单量单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "已收费", adDouble, 2, adFldIsNullable
        .Fields.Append "是否变价", adDouble, 2, adFldIsNullable
        .Fields.Append "位置", adDouble, 18, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 1, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable            '判断库存用
        .Fields.Append "单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "说明", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "记帐时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配料人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "审核人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "退料人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "库房货位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "科室ID", adDouble, 18, adFldIsNullable
        .Fields.Append "库存下限", adDouble, 18, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
    
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrsChargeOff = New ADODB.Recordset
    With mrsChargeOff
        If .State = 1 Then .Close
        .Fields.Append "领料部门", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "领料部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
        .Fields.Append "申请时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收发序号", adDouble, 18, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "准退数量", adDouble, 18, adFldIsNullable
        .Fields.Append "销帐数量", adDouble, 18, adFldIsNullable
        .Fields.Append "包装", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable
        .Fields.Append "主页ID", adDouble, 18, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 18, adFldIsNullable
        .Fields.Append "险类", adDouble, 18, adFldIsNullable
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "审核标志", adDouble, 18, adFldIsNullable
        .Fields.Append "材料名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "执行状态", adDouble, 2, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
        
    InitRsStruct = True
End Function
Private Function WhiteDataToRecord(ByVal rsSource As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据写入内部记录集(未发料部分)
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 10:03:41
    '-----------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:

    With rsSource
        If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
        Do While Not .EOF
            mrsNotPayStuff.AddNew
            mrsNotPayStuff!Id = !Id
            mrsNotPayStuff!状态 = "发料"    '全部默认为发料
            mrsNotPayStuff!科室 = !科室
            mrsNotPayStuff!开单医生 = !开单医生
            mrsNotPayStuff!类型 = Decode(NVL(!单据), 24, "收费单", 25, "记帐单", 26, "记帐表", "不知") & IIf(!已收费 = 0, "(未)", "")
            mrsNotPayStuff!材料ID = !材料ID
            mrsNotPayStuff!位置 = .AbsolutePosition
            mrsNotPayStuff!NO = !NO
            mrsNotPayStuff!单据 = !单据
            mrsNotPayStuff!病人ID = Val(NVL(!病人ID))
            mrsNotPayStuff!序号 = !序号
            mrsNotPayStuff!床号 = !床号
            mrsNotPayStuff!姓名 = NVL(!姓名)
            mrsNotPayStuff!住院号 = IIf(Val(NVL(!门诊标志)) = 2, NVL(!标识号), "")
            mrsNotPayStuff!材料名称 = NVL(!材料名称)
            mrsNotPayStuff!规格 = NVL(!规格)
            mrsNotPayStuff!产地 = NVL(!产地)
            mrsNotPayStuff!批次 = Val(NVL(!批次))
            mrsNotPayStuff!批号 = NVL(!批号)
            mrsNotPayStuff!换算系数 = Val(NVL(!换算系数))
            mrsNotPayStuff!在用分批 = Val(NVL(!分批))
            mrsNotPayStuff!是否变价 = Val(NVL(!是否变价))
            mrsNotPayStuff!付 = IIf(Val(NVL(!付)) = 0, 1, Val(NVL(!付)))
            mrsNotPayStuff!实际数量 = IIf(Val(NVL(!数量)) = 0, 1, Val(NVL(!数量)))
            mrsNotPayStuff!单位 = !单位
            mrsNotPayStuff!数量 = Format(IIf(Val(NVL(!数量)) = 0, 1, Val(NVL(!数量))) / !换算系数, mFMT.FM_数量) & !单位
            mrsNotPayStuff!单价 = !单价
            mrsNotPayStuff!金额 = !金额
            mrsNotPayStuff!记帐员 = NVL(!操作员姓名)
            mrsNotPayStuff!单量 = IIf(IsNull(!单量), "", zlStr.FormatEx(!单量, 5) & NVL(!计算单位))
            mrsNotPayStuff!单量单位 = NVL(!计算单位)
            mrsNotPayStuff!频次 = IIf(IsNull(!频次), "", !频次)
            mrsNotPayStuff!用法 = IIf(IsNull(!用法), "", !用法)
            mrsNotPayStuff!说明 = IIf(IsNull(!说明), "", !说明)
            mrsNotPayStuff!费用ID = Val(NVL(!费用ID))
            mrsNotPayStuff!记录性质 = Val(NVL(!记录性质))
            mrsNotPayStuff!门诊标志 = Val(NVL(!门诊标志))
            mrsNotPayStuff!医生嘱托 = IIf(IsNull(!医生嘱托), "", !医生嘱托)
            If IsNull(!登记时间) Then
                mrsNotPayStuff!记帐时间 = ""
            Else
                mrsNotPayStuff!记帐时间 = Format(!登记时间, "yyyy-MM-dd HH:mm:ss")
            End If
            
            mrsNotPayStuff!配料人 = NVL(!配料人)
            mrsNotPayStuff!审核人 = NVL(!审核人)
            mrsNotPayStuff!已收费 = !已收费                          '未收费或记帐处方，不允许发料
            mrsNotPayStuff!医嘱id = !医嘱id
            If mbln显示领退料人 = True Then
                mrsNotPayStuff!退料人 = !退料人         '主要是最后一次的退料人
            Else
                mrsNotPayStuff!退料人 = ""
            End If
            
            mrsNotPayStuff!库房货位 = IIf(IsNull(!库房货位), "", !库房货位)
            mrsNotPayStuff!科室id = IIf(IsNull(!科室id), 0, !科室id)
            mrsNotPayStuff!库存下限 = !库存下限
            
            '检查是否允许发料 :0-缺料,1-发料,2-拒发,3-不处理
            mrsNotPayStuff!执行状态 = 1     '缺省为发料
            If mrsNotPayStuff!已收费 = 0 And mbln发料前收费或审核 = False And mbln允许未收费的门诊划价处方发料 = False And mrsNotPayStuff!单据 = 24 Then
                mrsNotPayStuff!执行状态 = 3   '未收费的，不处理
                mrsNotPayStuff!状态 = "不处理"    '全部默认为发料
            End If
            
            '如果说明是拒发，则表明该材料已拒发，同时设置其执行状态
            If NVL(!说明) = "拒发" Then mrsNotPayStuff!执行状态 = 2
'            If mbln允许未审核处方发料 = False Then
                If NVL(mrsNotPayStuff!审核人) = "" And mbln发料前收费或审核 = False And mbln允许未审核的记账处方发料 = False And mrsNotPayStuff!单据 = 25 Then
                    mrsNotPayStuff!执行状态 = 3  '未审的，表示为不处理
                    mrsNotPayStuff!状态 = "不处理"    '全部默认为发料
'                End If
            Else
                '全部为发料
                mrsNotPayStuff!执行状态 = 1
            End If
            mrsNotPayStuff.Update
            If err <> 0 Then GoTo ErrHand
            .MoveNext
        Loop
    End With
   
    WhiteDataToRecord = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function FullDataToVsGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:将相关的数据填充到指定的网格控件中
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    FullDataToVsGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '填充数据到控件中
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    mblnHave发料 = False
    mblnHave拒发 = False
    
    With vsGrid
        .Clear (1)
        If mrsNotPayStuff.EOF Then '
            .Rows = 2
            FullDataToVsGrid = True
            Exit Function
        End If
        .Subtotal flexSTClear

        .Rows = mrsNotPayStuff.RecordCount + .FixedRows
        lngRow = .FixedRows
        Do While Not mrsNotPayStuff.EOF
            .RowData(lngRow) = Val(mrsNotPayStuff!Id)
            .TextMatrix(lngRow, .ColIndex("科室")) = NVL(mrsNotPayStuff!科室)
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = Val(NVL(mrsNotPayStuff!位置))
            .TextMatrix(lngRow, .ColIndex("开单医生")) = NVL(mrsNotPayStuff!开单医生)
            .TextMatrix(lngRow, .ColIndex("医生嘱托")) = NVL(mrsNotPayStuff!医生嘱托)
            .TextMatrix(lngRow, .ColIndex("状态")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "卫生材料发料") Or zlStr.IsHavePrivs(mstrPrivs, "卫生材料拒发"), NVL(mrsNotPayStuff!状态), "")
            '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
            .TextMatrix(lngRow, .ColIndex("单据类型")) = NVL(mrsNotPayStuff!类型)
            .TextMatrix(lngRow, .ColIndex("单据号")) = NVL(mrsNotPayStuff!NO)
            .TextMatrix(lngRow, .ColIndex("记帐员")) = NVL(mrsNotPayStuff!记帐员)
            .TextMatrix(lngRow, .ColIndex("床号")) = NVL(mrsNotPayStuff!床号)
            .TextMatrix(lngRow, .ColIndex("病人姓名")) = NVL(mrsNotPayStuff!姓名)
            .TextMatrix(lngRow, .ColIndex("住院号")) = NVL(mrsNotPayStuff!住院号)
            .TextMatrix(lngRow, .ColIndex("材料名称")) = NVL(mrsNotPayStuff!材料名称)
            .TextMatrix(lngRow, .ColIndex("规格")) = NVL(mrsNotPayStuff!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = NVL(mrsNotPayStuff!产地)
            .TextMatrix(lngRow, .ColIndex("批号")) = NVL(mrsNotPayStuff!批号)
            .TextMatrix(lngRow, .ColIndex("付数")) = Format(Val(NVL(mrsNotPayStuff!付)), "###")
            .TextMatrix(lngRow, .ColIndex("数量")) = NVL(mrsNotPayStuff!数量)
            .TextMatrix(lngRow, .ColIndex("单价")) = Format(Val(NVL(mrsNotPayStuff!单价)) * mrsNotPayStuff!换算系数, mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(NVL(mrsNotPayStuff!金额)), mFMT.FM_金额)
            .TextMatrix(lngRow, .ColIndex("说明")) = NVL(mrsNotPayStuff!说明)
            .TextMatrix(lngRow, .ColIndex("记帐时间")) = NVL(mrsNotPayStuff!记帐时间)
            If mrsNotPayStuff!执行状态 = 1 Then mblnHave发料 = True
            If mrsNotPayStuff!执行状态 = 2 Then mblnHave拒发 = True
            
            lngRow = lngRow + 1
            mrsNotPayStuff.MoveNext
         Loop
    End With
    
    '按单据进行汇总
    If SetTotalRowData = False Then Exit Function
    FullDataToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:刷新未发料单据数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 09:59:36
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, strWhere As String, strFields As String
    Dim str门诊 As String
    Dim rsTemp As New ADODB.Recordset
    Dim str病区发料 As String
    Dim str住院 As String
    Dim strSqlTmp As String
    
    On Error GoTo ErrHandle

    str病区发料 = zlDatabase.GetPara("病区发料方式", glngSys, mlngModule, "临床,护理,检查,检验,手术,治疗,营养")
    
    If mintUnit = 0 Then
        strFields = "x.计算单位 as 单位,1 as 换算系数,"
    Else
        strFields = "d.包装单位 as 单位,d.换算系数,"
    End If
    
    gstrSQL = "" & _
        "      Select Distinct s.Id, s.药品id AS 材料ID, Nvl(n.已收费, 0) 已收费, p.名称 科室, s.配药人 AS 配料人 ,S.费用ID, c.开单人 开单医生, " & _
        "          c.操作员姓名 审核人, s.单据, Nvl(s.扣率, 0) 扣率, s.No, s.序号, nvl(c.病人id,0) as 病人ID, '' 床号, c.姓名, " & _
        "          c.标识号, c.操作员姓名, '[' || x.编码 || ']' || x.名称 材料名称, s.付数 付, s.实际数量 数量, " & _
        "          Nvl(d.在用分批, 0) 分批, x.规格, c.登记时间," & strFields & _
        "          s.零售价 单价, s.零售金额 金额, s.单量, s.频次, s.用法, s.摘要 说明, " & _
        "          Decode(s.批号, Null, '', s.批号) || Decode(s.批次, Null, '', 0, '', '(' || s.批次 || ')') 批号, " & _
        "          Nvl(s.批次, 0) 批次, c.医嘱序号, i.计算单位, Nvl(s.产地, Nvl(x.产地, '')) 产地, " & _
        "          Nvl(m.审查结果, -1) 审查结果, Nvl(c.医嘱序号, -1) 医嘱id, '' 库房货位,x.是否变价, m.相关id,m.医生嘱托, " & _
        "          s.对方部门id As 科室id, c.序号 费用序号, C.记录性质,C.门诊标志,0 库存下限, z.名称 As 其它名 " & _
        "       From 未发药品记录 n,药品收发记录 s, 门诊费用记录 c,病人信息 c1, 病人医嘱记录 m,   " & _
        "          部门表 p, 材料特性 d, 收费项目目录 x, 收费项目别名 e,诊疗项目目录 i, 诊疗项目别名 z " & _
        "       Where n.单据 = s.单据 And  n.No = s.No AND nvl(n.库房id,[1])+0=nvl(s.库房id,[1])  " & _
        "             And s.费用id = c.Id AND s.对方部门id + 0 = p.Id  " & _
        "             And s.药品id = d.材料id And S.药品id = x.Id  " & _
        "             And s.药品id = e.收费细目id(+)  And e.性质(+) = 3 " & _
        "             And Nvl(Ltrim(Rtrim(s.摘要)), 'NOT拒发') <> '拒发'  AND s.审核人 Is Null And Nvl(s.发药方式, 0) <> -1 " & _
        "             And Mod(s.记录状态, 3) = 1 And instr([4],','||s.单据||',')>0 " & _
        "             AND d.诊疗ID=i.id  and C.病人ID=c1.病人ID(+) and C.病人ID=c1.病人ID(+) " & _
        "             AND D.诊疗id = z.诊疗项目id(+) And z.性质(+) = 2    " & _
        "             AND c.医嘱序号 = m.Id(+)  And Nvl(c.费用状态,0)<>1 " & _
        "             And Nvl(n.库房id, [1]) + 0 = [1]  " & _
        "             And n.填制日期 Between [2] And  [3]" & _
        "             "
    
    '排除对未发药品的销帐记录
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From 病人费用销帐 X " & _
        " Where X.申请类别 = 0 And X.状态+0 = 0 And X.收费细目id+0 = S.药品id And X.费用id = S.费用id) "
    
    '收费处方显示方式
    If Val(mArrFilter("收费处方")) = 1 Then
        gstrSQL = gstrSQL & " And n.已收费=1 "
    ElseIf Val(mArrFilter("收费处方")) = 2 Then
        gstrSQL = gstrSQL & " And n.已收费=0 "
    End If
        
    If Trim(mArrFilter("开单科室ID")) <> "" Then
        Select Case Val(mArrFilter("部门类型"))
        Case 0  '临床
            gstrSQL = gstrSQL & " And Instr([5], ',' || C.开单部门id || ',') > 0 And C.病人科室id=C.开单部门id"
        Case 1 '医技
            gstrSQL = gstrSQL & " And Instr([5], ',' || C.开单部门id || ',') > 0 And C.病人科室id<>C.开单部门id"
        Case Else
            '病区
            If str病区发料 = "" Then
                gstrSQL = gstrSQL & " And Instr([5], ',' || C.病人病区ID || ',') > 0 And C.病人科室id=C.开单部门id"
            Else
                gstrSQL = gstrSQL & " And Instr([5], ',' || C.病人病区ID || ',') > 0 "
                If str病区发料 <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.开单部门id Not In (Select Distinct 部门id From 部门性质说明 " & _
                        " Where Instr([13],',' || 工作性质 || ',') > 0) "
                End If
            End If
        End Select
    End If
    
    strWhere = ""
    If (Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) = "") Then
        strWhere = strWhere & "            AND s.NO =[6]  "
    ElseIf (Trim(mArrFilter("单据号")(1)) <> "" And Trim(mArrFilter("单据号")(0)) = "") Then
        strWhere = strWhere & "            AND s.NO =[7]  "
    ElseIf Trim(mArrFilter("单据号")(0)) <> "" And Trim(mArrFilter("单据号")(1)) <> "" Then
        strWhere = strWhere & "            AND ( s.NO between [6] and [7] )"
    End If
    
    gstrSQL = gstrSQL & strWhere
    
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("病人ID")) = 0 And Val(mArrFilter("IC卡号")) = 0, "", "       AND c.病人iD=[8]  ")
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("住院号")) = 0, "", "       AND c.标识号=[9] and c.门诊标志=2 ")
    gstrSQL = gstrSQL & IIf(Trim(mArrFilter("姓名")) = "", "", "       AND C.姓名 like [10] ")
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("门诊号")) = 0, "", "       AND c.标识号=[11] and c.门诊标志=1 ")
    gstrSQL = gstrSQL & IIf(Trim(mArrFilter("就诊卡号")) = "", "", "   AND c1.就诊卡号 =[12] ")
    
    If mbln按发生时间过滤 = True Then
        strSqlTmp = Replace(gstrSQL, "n.填制日期", "c.发生时间") & " And C.医嘱序号 Is Not Null"
        gstrSQL = gstrSQL & " And C.医嘱序号 Is Null"
        gstrSQL = gstrSQL & " Union All " & strSqlTmp
    End If
    
    If mbln显示领退料人 Then
        gstrSQL = " Select a.*, b.退药人 as 退料人 " & _
                 "  From ( " & gstrSQL & _
                 "          Order By s.No, s.单据 " & _
                 "       ) a, " & _
                 "      (Select a.单据, a.No, a.序号, a.领用人 退药人 " & _
                 "       From 药品收发记录 a, " & _
                 "          (   Select s.单据, s.No, s.序号, Max(s.记录状态) 记录状态 " & _
                 "              From 药品收发记录 s, 未发药品记录 n " & _
                 "              Where s.No = n.No And s.单据 = n.单据 And Nvl(s.发药方式, 0) <> -1 " & _
                 "                     And Nvl(s.库房id, [1]) + 0 = Nvl(n.库房id, [1])  " & _
                 "                     And Nvl(s.库房id, [1]) + 0 = [1]  " & _
                 "                     And n.填制日期 Between [2] And [3]  " & _
                 "                     And Mod(s.记录状态, 3) = 2 And instr([4],','||s.单据||',')>0 " & strWhere & _
                 "              Group By s.单据, s.No, s.序号 " & _
                 "            ) b " & _
                 "       Where a.单据 = b.单据 And a.No = b.No And a.序号 = b.序号 And a.记录状态 = b.记录状态) b " & _
                 "  Where a.单据 = b.单据(+) And a.No = b.No(+) And a.序号 = b.序号(+) "
    End If
    
    If Val(mArrFilter("请求类型")) = 0 Then
        '所有
        str门诊 = Replace(gstrSQL, "C.病人病区ID", "C.开单部门id")
        str住院 = Replace(gstrSQL, "'' 床号", "c.床号")
        str住院 = Replace(str住院, "C.姓名", "nvl(r.姓名,C.姓名)")
        str住院 = Replace(str住院, "c.姓名", "nvl(r.姓名,c.姓名) 姓名")
        str住院 = Replace(str住院, "门诊费用记录 c", "住院费用记录 c,病案主页 r")
        str住院 = Replace(str住院, "And Nvl(c.费用状态,0)<>1", " and r.病人id=c.病人id and r.主页id=c.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND c.床号 =[14] "))
        If Trim(mArrFilter("床号")) <> "" Then str门诊 = str门诊 & " and 1=0"
        gstrSQL = str门诊 & " Union All " & str住院
    ElseIf Val(mArrFilter("请求类型")) = 1 Then
        gstrSQL = Replace(gstrSQL, "C.病人病区ID", "C.开单部门id")
    ElseIf Val(mArrFilter("请求类型")) = 2 Then
        '住院记帐单
        gstrSQL = Replace(gstrSQL, "'' 床号", "c.床号")
        gstrSQL = Replace(gstrSQL, "C.姓名", "nvl(r.姓名,C.姓名)")
        gstrSQL = Replace(gstrSQL, "c.姓名", "nvl(r.姓名,c.姓名) 姓名")
        gstrSQL = Replace(gstrSQL, "门诊费用记录 c", "住院费用记录 c,病案主页 r")
        gstrSQL = Replace(gstrSQL, "And Nvl(c.费用状态,0)<>1", " and r.病人id=c.病人id and r.主页id=c.主页id " & IIf(Trim(mArrFilter("床号")) = "", "", "   AND c.床号 =[14] "))
    End If
    
    gstrSQL = gstrSQL & "  Order By No, 费用序号"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("发料部门ID")), _
        CDate(mArrFilter("日期范围")(0)), _
        CDate(mArrFilter("日期范围")(1)), _
        CStr("," & mArrFilter("单据") & ","), _
        "," & Trim(mArrFilter("开单科室ID")) & ",", _
        CStr(mArrFilter("单据号")(0)), _
        CStr(mArrFilter("单据号")(1)), _
        Val(mArrFilter("病人ID")), _
        Val(mArrFilter("住院号")), _
        CStr(mArrFilter("姓名")), _
        Val(mArrFilter("门诊号")), _
        CStr(mArrFilter("就诊卡号")), _
        "," & str病区发料 & ",", Val(mArrFilter("床号")))
    
    '初始化内容结构
    Call InitRsStruct
    '填充相关的数据到内部记录集
    Call WhiteDataToRecord(rsTemp)
    '检查库存是否充足
    Call CheckStock
    '组织销帐数据集
    Call GetChargeOffRecord(mrsNotPayStuff)
    
    RaiseEvent zlRefreshDataRecordSet(mrsNotPayStuff, mrsChargeOff)
    
    Call FullDataToVsGrid
    RefreshData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetTotalRowData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置行的汇总属性
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 10:22:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, strUpper As String
    Dim lngOldRow As Long, lngOldTopRow As Long
    With vsGrid
        .Redraw = flexRDNone
         .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        .Subtotal flexSTSum, -1, .ColIndex("金额"), "###.00", , vbBlue, True, "合计"
        .Subtotal flexSTSum, .ColIndex("单据类型"), .ColIndex("金额"), "###.00", , vbBlue, True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("金额"), "###.00", , vbBlue, True
        '.Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("数量"), "###.00", , , True, "单据小计"
        .Editable = flexEDKbdMouse
        If .Rows > 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("状态"), .Rows - 1, .ColIndex("状态")) = &HE7CFBA
            .Cell(flexcpBackColor, 1, .ColIndex("数量"), .Rows - 1, .ColIndex("数量")) = &HE7CFBA
        End If
        
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) = True Then
                If InStr(1, .TextMatrix(lngRow, .ColIndex("单据类型")), "Total") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("单据类型")) = Replace(.TextMatrix(lngRow, .ColIndex("单据类型")), "Total", "")
                    If Trim(.TextMatrix(lngRow, .ColIndex("单据类型"))) <> "" Then
                        .TextMatrix(lngRow, .ColIndex("单据类型")) = Trim(.TextMatrix(lngRow, .ColIndex("单据类型"))) & "(小计)"
                    End If
                End If
                
                .TextMatrix(lngRow, .ColIndex("单据号")) = Replace(.TextMatrix(lngRow, .ColIndex("单据号")), "Total", "")
                If Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) <> "" Then
                    .TextMatrix(lngRow, .ColIndex("单据号")) = Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) & "(小计)"
                End If

            End If
        Next
        
        '进行单据合并
'        .MergeCells = flexMergeRestrictRows
'        For lngRow = 1 To .Rows - 1
'            .MergeRow(lngRow) = False
'            If .IsSubtotal(lngRow) = True Then
'                .MergeRow(lngRow) = True
'                strUpper = Trim(.TextMatrix(lngRow, .ColIndex("金额"))) & " (大写:" & zlCommFun.UppeMoney(Val(.TextMatrix(lngRow, .ColIndex("金额")))) & ")"
'                For lngCol = .ColIndex("单据号") To .Cols - 1
'                    .TextMatrix(lngRow, lngCol) = strUpper
'                Next
'            End If
'        Next
        .OutlineBar = 1

        .Redraw = flexRDBuffered
    End With
End Function

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "未发料"
    Call zlDatabase.SetPara("发料单据打印格式", cboEdit(mcboIdx.idx_单据格式), glngSys, mlngModule)
End Sub
Private Sub txtEDIT_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    OS.OpenIme True
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Index <> mtxtIdx.idx_配料人 Then Exit Sub
    If Trim(txtEdit(Index).Text) = "" Then Exit Sub
    If txtEdit(Index).Tag <> "" Then OS.PressKey vbKeyTab
    Call SelectItem(txtEdit(Index), Trim(txtEdit(Index)))
End Sub

Private Sub txtEDIT_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, str状态 As String
    With vsGrid
        Select Case Col
        Case .ColIndex("状态")
            Call ChangeSelStaut(Row)
        End Select
    End With
End Sub
Private Sub ChangeSelStaut(ByVal Row As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:改变指定行的状态
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-01 23:58:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, str状态 As String
    Dim lng级数 As Long
    
    With vsGrid
            If .IsSubtotal(Row) Then
               str状态 = Trim(.TextMatrix(Row, .ColIndex("状态")))
               lng级数 = .RowOutlineLevel(Row)
               For lngRow = Row + 1 To .Rows - 1
                   If .RowOutlineLevel(lngRow) <> lng级数 Then
                       If .TextMatrix(lngRow, .ColIndex("状态")) <> "缺料" Then
                            .TextMatrix(lngRow, .ColIndex("状态")) = str状态
                            '设置相关的执行状态
                            Call SetExecuteStaut(lngRow)
                            
                       End If
                   Else
                        Exit For
                   End If
               Next
            Else
                '设置相关的执行状态
                Call SetExecuteStaut(Row)
            End If
    End With
End Sub
 

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case .Col
        Case .ColIndex("状态")
                  If zlStr.IsHavePrivs(mstrPrivs, "卫生材料发料") = False And zlStr.IsHavePrivs(mstrPrivs, "卫生材料拒发") = False Then
                        Cancel = True
                        Exit Sub
                  End If

                If .TextMatrix(Row, Col) = "缺料" Then
                    Cancel = True
                End If
'                If Row = 1 And .IsSubtotal(Row) = True Then
'                    Cancel = True
'                End If
'                If .IsSubtotal(Row) = True And InStr(1, Trim(.TextMatrix(Row, .ColIndex("单据类型"))), "小计") > 0 Then
'                    Cancel = True
'                End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Function InitCheckStock() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化检查库存的记录集
    '入参:
    '出参:
    '返回:返回记录集
    '编制:刘兴洪
    '日期:2008-04-22 20:47:33
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "材料ID", adDouble, 18
        .Fields.Append "批次", adDouble, 18
        .Fields.Append "变价", adDouble, 18
        .Fields.Append "数量", adDouble, 18
        .Fields.Append "序号", adDouble, 5
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set InitCheckStock = rsTemp
End Function
Private Function CheckStock(Optional lng材料ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查可用库存,确定是否存在缺料
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-22 20:38:29
    '-----------------------------------------------------------------------------------------------------------
    Dim rsStock As ADODB.Recordset
    Dim lngRow As Long, lng序号 As Long
    Dim arrtemp As Variant
    
    Set rsStock = InitCheckStock
    With mrsNotPayStuff
        '检查库存:
        '   1.由于可能存在相同的批次和材料，因此，需要逐步减少每笔库存，才能最终确定各卫材的库存数量是否允足
        '   2.
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        
        If lng材料ID <> 0 Then
            '可能对某种材料进行检查
            .Filter = "材料ID=" & lng材料ID
            If .RecordCount = 0 Then .Filter = 0: Exit Function
        End If
        Do While Not .EOF
            If !执行状态 <= 1 Then
                '只检查缺料和发料这种情况
                If LocaleStockData(rsStock, Val(mArrFilter("发料部门ID")), Val(NVL(!材料ID)), Val(NVL(!批次)), lng序号) = True Then
                   '找到了指定的库存:需要检查数量是否充足，备货卫材要单独判断，不考虑发料部门的库存数量
                        If Val(NVL(rsStock!数量)) - Val(NVL(mrsNotPayStuff!实际数量)) < 0 Then
                            '数量不够,需确定为缺料:（有批次或变价的,必需为检查库存）
                            If mrsNotPayStuff!已收费 = 0 And mbln发料前收费或审核 = False And mbln允许未收费的门诊划价处方发料 = False And mrsNotPayStuff!单据 = 24 Then
                                !执行状态 = 3
                            ElseIf NVL(mrsNotPayStuff!审核人) = "" And mbln发料前收费或审核 = False And mbln允许未审核的记账处方发料 = False And mrsNotPayStuff!单据 = 25 Then
                                !执行状态 = 3
                            Else
                                !执行状态 = IIf(mlng缺料检查 = 1 Or rsStock!批次 <> 0 Or rsStock!变价 = 1, 0, !执行状态)
                            End If
                            
                            .Update
                        Else
                            '数量充足:
                            !执行状态 = 1   '缺省为发料
                            If mrsNotPayStuff!已收费 = 0 And mbln发料前收费或审核 = False And mbln允许未收费的门诊划价处方发料 = False And mrsNotPayStuff!单据 = 24 Then !执行状态 = 3      '未收费的，强制为不处理
                            If NVL(mrsNotPayStuff!审核人) = "" And mbln发料前收费或审核 = False And mbln允许未审核的记账处方发料 = False And mrsNotPayStuff!单据 = 25 Then !执行状态 = 3 '未审核的单据,默认为不处理
                            .Update
                        End If
                        If !执行状态 = 1 Then
                            '如果发料,需要更改可库存数
                            With rsStock
                                !数量 = Val(NVL(!数量)) - Val(NVL(mrsNotPayStuff!实际数量))
                                .Update
                            End With
                        End If
                End If
                
                If !执行状态 = 0 Then
                    If CheckIsStockUp(Val(mArrFilter("发料部门ID")), Val(NVL(!材料ID)), Val(NVL(!批次)), Val(NVL(!费用ID)), Val(NVL(mrsNotPayStuff!实际数量))) = True Then
                        !执行状态 = 1
                    End If
                End If
                
                !状态 = Decode(!执行状态, 0, "缺料", 1, "发料", 2, "拒发", "不处理")
                .Update
            End If
            .MoveNext
        Loop
     End With
     mrsNotPayStuff.Filter = 0
End Function

Private Function LocaleStockData(ByRef rsStock As ADODB.Recordset, _
    ByVal lng发料部门ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, Optional ByRef lng序号 As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查指定材料指定批的库存是否充足
    '入参:rsStock-指定检查的库存数据(可以为空记录),可以自动扩展
    '     lng发料部门ID-发料部门id
    '     lng材料id-材料id
    '     lng批次-批次
    '
    '出参:lng序号-返回库存的序号
    '返回:成功,表示找到,否则表示未找到
    '编制:刘兴洪
    '日期:2008-04-22 21:07:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim dbl库存 As Double
    LocaleStockData = False
    
    err = 0: On Error GoTo ErrHand:
    With rsStock
        .Filter = "材料ID=" & lng材料ID & " and 批次=" & lng批次
        If .RecordCount = 0 Then
            .Filter = 0: lng序号 = .RecordCount + 1
            
            gstrSQL = "" & _
            " Select nvl(F.是否变价,0) 变价,nvl(A.实际数量,0) 数量" & _
            " From 材料特性 B,收费项目目录 F," & _
            "      (Select A.药品id as 材料ID,a.实际数量 From 药品库存 A Where 性质=1 And 库房ID=[1] And 药品ID=[2] And nvl(批次,0)=[3]) A" & _
            " Where B.材料ID=F.ID And B.材料ID=A.材料ID(+) And B.材料ID=[2] "
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng发料部门ID, lng材料ID, lng批次)
            
            dbl库存 = Val(NVL(rsTemp!数量))
            .AddNew
            !材料ID = lng材料ID
            !批次 = lng批次
            !变价 = rsTemp!变价
            !数量 = dbl库存
            !序号 = lng序号
            .Update
        Else
            lng序号 = Val(NVL(!序号))
            .Filter = 0
        End If
        .MoveFirst
        .Find "序号=" & lng序号
        LocaleStockData = True
    End With
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化控件数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-04-23 16:29:05
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, intDefault As Integer
    
    On Error GoTo ErrHandle
    If mbln按单据发料 = True Then
        vsGrid.ColHidden(vsGrid.ColIndex("状态")) = True
    Else
        vsGrid.ColHidden(vsGrid.ColIndex("状态")) = False
    End If
    
    intDefault = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\ZL1_BILL_1723_1", "Format", 1))
    txtEdit(mtxtIdx.idx_配料人).Text = gstrUserName: txtEdit(mtxtIdx.idx_配料人).Tag = gstrUserName
    '加载相关打印单据
    If cboEdit(mcboIdx.idx_单据格式).ListCount <> 0 Then Exit Sub
    mstr默认单据格式 = Trim(zlDatabase.GetPara("发料单据打印格式", glngSys, mlngModule, , Array(cboEdit(mcboIdx.idx_单据格式)), zlStr.IsHavePrivs(mstrPrivs, "参数设置")))
    gstrSQL = "Select  序号,说明 From zltools.zlRPTFMTs Where 报表id = (Select ID From zltools.zlReports Where 编号 = 'ZL1_BILL_1723_1') Order By 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取单据打印格式")
    With cboEdit(mcboIdx.idx_单据格式)
        Do While Not rsTemp.EOF
            .AddItem rsTemp!说明
            If mstr默认单据格式 <> "" Then
                If NVL(rsTemp!说明) = mstr默认单据格式 Then
                    .ListIndex = .NewIndex
                End If
            ElseIf Val(NVL(rsTemp!序号)) = intDefault Then
                    .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 And .ListIndex < 0 Then .ListIndex = 0
        If rsTemp.RecordCount = 1 Then .Enabled = False
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function SelectItem(ByVal objCtl As Control, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, strTittle As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    strTittle = "配料人选择"
    
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    gstrSQL = "" & _
        "   Select distinct a.编号 as 编码,A.姓名 As 名称,简码" & _
        "   From 人员表 A,部门人员 B,部门性质说明 C,人员性质说明 D " & _
        "   Where A.Id=B.人员id And B.部门id=C.部门Id And D.人员id=A.Id " & _
        "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND B.部门id in (Select 部门ID From 部门人员 where 人员id=[2] ) "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "    And  ((A.姓名) like [1] or  A.编号  like [1] or  简码  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = "Select rownum as ID,a.* from (" & gstrSQL & ") A" & _
        "   ORDER BY 编码 "
     
     strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey, UserInfo.Id)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgBox "没有找到满足条件的内容,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!名称)
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = NVL(rsTemp!名称)
        objCtl.Tag = NVL(rsTemp!名称)
        OS.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
 Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Private Sub vsGrid_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsGrid
        If Position <= .ColIndex("单据号") Then
            ShowMsgBox "不能将列移动到单据号以前的列!"
            Position = Col
        End If
    End With
End Sub

Private Sub vsGrid_DblClick()
    Dim str状态 As String
    If zlStr.IsHavePrivs(mstrPrivs, "卫生材料发料") = False And zlStr.IsHavePrivs(mstrPrivs, "卫生材料拒发") = False Then Exit Sub
    With vsGrid
'         If .Row = 1 And .IsSubtotal(.Row) = True Then
'            Exit Sub
'        End If
         
         If mbln按单据发料 = True Then Exit Sub
         
         str状态 = Trim(.TextMatrix(.Row, .ColIndex("状态")))
         
        .TextMatrix(.Row, .ColIndex("状态")) = Decode(str状态, "发料", "拒发", "拒发", "不处理", "缺料", "缺料", "发料")
        
        If .IsSubtotal(.Row) = True Then
            Call ChangeSelStaut(.Row)
            Exit Sub
        End If
     
        Call ChangeSelStaut(.Row)
    End With
End Sub
Public Property Get zlHaveSel发料() As Boolean
    zlHaveSel发料 = mblnHave发料
End Property
Public Property Get zlHaveSel拒发() As Boolean
    zlHaveSel拒发 = mblnHave拒发
End Property
Public Property Get zlHaveData() As Boolean
    zlHaveData = mrsNotPayStuff.RecordCount <> 0
End Property

Private Sub BillListPrint(Optional strDate As String = "", Optional IntStyle As Integer = 0)
    '-----------------------------------------------------------------------------------------------------------
    '功能:单据或清册打印
    '入参:
    '     intStyle:0-按发料方式打印,1-单据打印,2-退料单据
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-05 10:36:44
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim bln退料单 As Boolean
    Dim bln已发料清单 As Boolean
    Dim bln单据打印 As Boolean
    Dim intMsg As Integer   '0-提示打印,1-自动打印,2-不打印
    
    
    intMsg = Val(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0"))
    
    bln退料单 = zlStr.IsHavePrivs(mstrPrivs, "退料通知单")
    bln已发料清单 = zlStr.IsHavePrivs(mstrPrivs, "打印已发料清单")
    bln单据打印 = zlStr.IsHavePrivs(mstrPrivs, "单据打印")
    If intMsg = 0 Then
        '提示打印
        If bln已发料清单 = False Then Exit Sub
        If MsgBox("你需要打印相关单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intMsg = 1 Then
        '自动打印
    Else
        Exit Sub
    End If
        '部门发料
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
       "库房=" & Val(mArrFilter("发料部门ID")), _
       "发料方式=部门发料|3", _
       "部门性质=" & Val(mArrFilter("部门类型")), _
       "接收科室=" & 获取接收部门条件(strDate), _
       "单位=" & IIf(mintUnit = 0, 0, 1), _
       "发料号=" & strDate, _
       "汇总发料号=" & Val(mstr汇总标识号), _
       "ReportFormat=" & IIf(cboEdit(mcboIdx.idx_单据格式).ListIndex = -1, 1, cboEdit(mcboIdx.idx_单据格式).ListIndex + 1), "PrintEmpty=0", 2)
End Sub
Private Function 获取接收部门条件(ByVal strDate As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取接收部门的打印条件
    '入参:
    '出参:
    '返回:成功,返回 显示|IN(部门ID,..) ,否则返回""
    '编制:刘兴洪
    '日期:2008-05-05 13:31:28
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, str显示 As String, strIDIn As String
    
    On Error GoTo ErrHandle
    If mArrFilter("开单科室id") = "" Then
        '没有条件,则以根据选择的类别读取显示科室
        gstrSQL = "Select distinct D.ID,D.编码,D.名称 as 科室 " & _
                 " From 药品收发记录 S,门诊费用记录 C,部门表 d " & _
                 " Where S.费用ID=C.ID And Mod(S.记录状态,3) In (0,1) And S.审核人 Is Not Null " & _
                 "      And C.执行状态=1 And S.库房ID=[1] And S.发药方式=3 And S.审核日期=[2] " & _
                 "      And S.单据 In (24,25,26) "
        Select Case Val(mArrFilter("部门类型"))
            Case 0  '
                gstrSQL = gstrSQL & " and C.病人科室id=d.id(+) "
            Case 1 '医技
                gstrSQL = gstrSQL & "  and C.开单部门id =d.id(+)"
            Case Else '病区
                gstrSQL = gstrSQL & "  and C.病人病区ID =d.id(+)"
        End Select
        
        Select Case Val(mArrFilter("部门类型"))
            Case 0, 1
                gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            Case Else '病区
                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        End Select
        
        gstrSQL = gstrSQL & "order by 编码"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("发料部门id")), CDate(strDate))
        With rsTemp
            Do While Not .EOF
                str显示 = str显示 & "," & !科室
                strIDIn = strIDIn & "," & !Id
                rsTemp.MoveNext
            Loop
        End With
        strIDIn = "0" & strIDIn
        str显示 = str显示 & "|" & " IN (" & strIDIn & ")"
        获取接收部门条件 = str显示
        Exit Function
    End If
    gstrSQL = "Select ID, 名称 From 部门表 A, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) J Where ID = J.Column_Value order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mArrFilter("开单科室id")))
    With rsTemp
        Do While Not .EOF
            str显示 = str显示 & "," & !名称
            rsTemp.MoveNext
        Loop
    End With
    str显示 = str显示 & "|" & " IN (0" & CStr(mArrFilter("开单科室id")) & ")"
    获取接收部门条件 = str显示
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlSetFontSize(ByVal curFontSize As Currency)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-06 17:00:44
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
    lblEdit(mlblIdx.idx_lbl配料人).Font.Size = curFontSize
    lblEdit(mlblIdx.idx_lbl配料人).AutoSize = True
    lblEdit(mlblIdx.idx_lbl单据格式).FontSize = curFontSize
    lblEdit(mlblIdx.idx_lbl单据格式).AutoSize = True
    cboEdit(mcboIdx.idx_单据格式).Font.Size = curFontSize
    txtEdit(mtxtIdx.idx_配料人).Font.Size = curFontSize
    Call Form_Resize
End Sub
Public Property Get zl_上次汇总发料号() As String
    '返回上次汇总发料号
    zl_上次汇总发料号 = mstr汇总标识号
End Property

 
Private Sub vsGrid_EnterCell()
    Dim lngRow As Long
    Dim strNo As String
    Dim lng病人id As Long
    
    If mbln按单据发料 = True Then
        mblnHave发料 = False
        If vsGrid.Row > 0 Then
            If vsGrid.IsSubtotal(vsGrid.Row) = False Then
                If vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("状态")) <> "发料" Then Exit Sub
                strNo = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("单据号"))
                For lngRow = 1 To vsGrid.Rows - 1
                    If vsGrid.IsSubtotal(lngRow) = False Then
                        If strNo = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("单据号")) Then
                            If vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("状态")) <> "发料" Then
                                mblnHave发料 = False
                                Exit For
                            End If
                        End If
                    End If
                Next
                
                mblnHave发料 = True
                
                If mbln发料前收费或审核 = True Then
                    With mrsNotPayStuff
                        .Filter = ""
                        If Not .EOF Then
                            '默认全部执行状态为"不发料"
                            Do While Not .EOF
                                !执行状态 = 0
                                .MoveNext
                            Loop
                            
                            '寻找当前选择的病人
                            .Filter = "NO = '" & strNo & "'"
                            lng病人id = !病人ID
                            
                            .Filter = "病人id =" & lng病人id
                            Do While Not .EOF
                                !执行状态 = 1
                                .MoveNext
                            Loop
                            
                            .Filter = ""
                        End If
                    End With
                End If
            End If
        End If
    End If
End Sub


