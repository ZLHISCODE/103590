VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm库存变动 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "库存变动"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frm库存变动.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8895
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "说明：该窗体显示盘点表中对应药品在盘点日期后发生的库存变动情况！"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   75
         Width           =   5760
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   0
         Picture         =   "frm库存变动.frx":000C
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "退出(&E)"
      Height          =   345
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf库存变动 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      _cx             =   15690
      _cy             =   8281
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm库存变动.frx":685E
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
      ExplorerBar     =   5
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
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "注意：红色表示发生出库业务！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   5235
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Caption         =   "说明：该窗体显示盘点表中对应药品在盘点日期后发生的库存变动情况！"
      Height          =   180
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "frm库存变动"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintStyle As Integer '1:表示窗体显示库存变动；2：表示窗体显示可用数量占用
Private mlng库房ID As Long
Private mlng药品ID As Long
Private mlng批次 As Long
Private mstr盘点时间 As String
Private mstr单位 As String
Private mdbl比例系数 As Double
Private mstr单位小 As String
Private mdbl比例系数小 As Double
Private mbln区分大小单位 As Boolean
Private mintNumberDigit As Integer

Public Sub ShowME(ByVal intStyle As Integer, ByVal lng库房ID As Long, ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal str盘点时间 As String, ByVal frmPar As Form, ParamArray arrInput() As Variant)
    Dim arrPars() As Variant
    arrPars = arrInput
    
    If UBound(arrPars) = 2 Then
        mbln区分大小单位 = False
        mstr单位 = arrPars(0)
        mdbl比例系数 = arrPars(1)
        mintNumberDigit = arrPars(2)
    Else
        mbln区分大小单位 = True
        mstr单位 = arrPars(0)
        mdbl比例系数 = arrPars(1)
        mstr单位小 = arrPars(2)
        mdbl比例系数小 = arrPars(3)
        mintNumberDigit = arrPars(4)
    End If
    
    mintStyle = intStyle
    mlng库房ID = lng库房ID
    mlng药品ID = lng药品ID
    mlng批次 = lng批次
    mstr盘点时间 = str盘点时间
    
    Me.Show 1, frmPar
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim int大包装数量 As Integer
    
    On Error GoTo ErrHandle
    If mintStyle = 1 Then '库存变动
        gstrSQL = "Select * From (Select a.NO,Decode(a.单据," & vbNewLine & _
                    "               1," & vbNewLine & _
                    "               '外购入库'," & vbNewLine & _
                    "               2," & vbNewLine & _
                    "               '自制入库'," & vbNewLine & _
                    "               3," & vbNewLine & _
                    "               '协药入库'," & vbNewLine & _
                    "               4," & vbNewLine & _
                    "               '其他入库'," & vbNewLine & _
                    "               6," & vbNewLine & _
                    "               '药品移库'," & vbNewLine & _
                    "               7," & vbNewLine & _
                    "               '部门领用'," & vbNewLine & _
                    "               11," & vbNewLine & _
                    "               '其他出库'," & vbNewLine & _
                    "               12," & vbNewLine & _
                    "               '药品盘点'," & vbNewLine & _
                    "               '处方发药') As 业务类型,  a.入出系数 * a.实际数量 * a.付数 As 发生数量, To_Char(a.审核日期, 'yyyy-mm-dd HH24:Mi:SS') As 发生日期, a.填制人, a.审核人,a.记录状态,a.发药方式,a.单据" & vbNewLine & _
                    "From 药品收发记录 a" & vbNewLine & _
                    "Where  a.库房id = [1] And a.药品id = [2] And nvl(a.批次,0) = [3] And a.审核日期 > To_Date([4], 'YYYY-MM-DD HH24:MI:SS') And a.单据 not in (5,13)" & vbNewLine & _
                    "Order By a.审核日期 Desc )" & vbNewLine & _
                    "union all " & vbNewLine & _
                    "Select '','合计' As 业务类型 , sum(a.入出系数 * a.实际数量 * a.付数) As 发生数量, '', '', '',Null,Null,Null" & vbNewLine & _
                    "From 药品收发记录 a,收费项目目录 b" & vbNewLine & _
                    "Where a.药品id = b.id and a.库房id = [1] And a.药品id = [2] And nvl(a.批次,0) = [3] And a.审核日期 > To_Date([4], 'YYYY-MM-DD HH24:MI:SS')And a.单据 not in (5,13)"
    
    
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng库房ID, mlng药品ID, mlng批次, mstr盘点时间)
        
        If rsTemp.RecordCount <= 1 Then Exit Sub
        
        With vsf库存变动
            
            Do While Not rsTemp.EOF
                .rows = .rows + 1
                .Row = .Row + 1
                
                .TextMatrix(.Row, .ColIndex("NO")) = "" & rsTemp!NO
                .TextMatrix(.Row, .ColIndex("业务类型")) = "" & rsTemp!业务类型
                
                '显示冲销、退库单据提示
                If rsTemp!单据 = 1 Then '外购入库
                    If rsTemp!记录状态 Mod 3 = 2 Then '冲销单据
                        If rsTemp!发药方式 = 1 Then '退库
                            .TextMatrix(.Row, .ColIndex("业务类型")) = .TextMatrix(.Row, .ColIndex("业务类型")) & "(退、冲)"
                        Else
                            .TextMatrix(.Row, .ColIndex("业务类型")) = .TextMatrix(.Row, .ColIndex("业务类型")) & "(冲销)"
                        End If
                    Else
                        If rsTemp!发药方式 = 1 Then .TextMatrix(.Row, .ColIndex("业务类型")) = .TextMatrix(.Row, .ColIndex("业务类型")) & "(退库)" '退库
                    End If
                Else
                    If rsTemp!记录状态 Mod 3 = 2 Then .TextMatrix(.Row, .ColIndex("业务类型")) = .TextMatrix(.Row, .ColIndex("业务类型")) & "(冲销)" '冲销单据
                End If
                
                '颜色区分入、出库
                If rsTemp!发生数量 < 0 Then .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '出库红色
               
                
                If Not mbln区分大小单位 Then
                    .TextMatrix(.Row, .ColIndex("发生数量")) = zlStr.FormatEx(Abs(IIf(IsNull(rsTemp!发生数量), 0, rsTemp!发生数量)) / mdbl比例系数, mintNumberDigit, , True) & mstr单位
                Else
'                    If rsTemp!发生数量 < 0 Then .TextMatrix(.Row, .ColIndex("发生数量")) = "-"
                    
                    int大包装数量 = Int(Abs(IIf(IsNull(rsTemp!发生数量), 0, rsTemp!发生数量)) / mdbl比例系数)
                    .TextMatrix(.Row, .ColIndex("发生数量")) = .TextMatrix(.Row, .ColIndex("发生数量")) & zlStr.FormatEx(int大包装数量, mintNumberDigit, , True) & mstr单位
                    .TextMatrix(.Row, .ColIndex("发生数量")) = .TextMatrix(.Row, .ColIndex("发生数量")) & zlStr.FormatEx((Abs(IIf(IsNull(rsTemp!发生数量), 0, rsTemp!发生数量)) - int大包装数量 * mdbl比例系数) / mdbl比例系数小, mintNumberDigit, , True) & mstr单位小
                End If
                
                .TextMatrix(.Row, .ColIndex("发生日期")) = "" & rsTemp!发生日期
                .TextMatrix(.Row, .ColIndex("填制人")) = "" & rsTemp!填制人
                .TextMatrix(.Row, .ColIndex("审核人")) = "" & rsTemp!审核人
                
                rsTemp.MoveNext
    
            Loop
             
            .TextMatrix(.rows - 1, .ColIndex("业务类型")) = .TextMatrix(.rows - 1, .ColIndex("业务类型")) & IIf(.Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF, "(出库)", "(入库)")
        End With
       
        
    Else '可用数量占用
        Me.Caption = "可用数量占用"
        lblComment.Caption = "说明：该窗体显示盘点表中对应药品30天之内的可用数量占用情况！"
        gstrSQL = "Select * From (Select Decode(a.单据," & vbNewLine & _
                    "               1," & vbNewLine & _
                    "               '外购入库'," & vbNewLine & _
                    "               2," & vbNewLine & _
                    "               '自制入库'," & vbNewLine & _
                    "               3," & vbNewLine & _
                    "               '协药入库'," & vbNewLine & _
                    "               4," & vbNewLine & _
                    "               '其他入库'," & vbNewLine & _
                    "               6," & vbNewLine & _
                    "               '药品移库'," & vbNewLine & _
                    "               7," & vbNewLine & _
                    "               '部门领用'," & vbNewLine & _
                    "               11," & vbNewLine & _
                    "               '其他出库'," & vbNewLine & _
                    "               12," & vbNewLine & _
                    "               '药品盘点'," & vbNewLine & _
                    "               '处方发药') As 业务类型, a.实际数量 * a.付数 As 占用数量, To_Char(a.填制日期, 'yyyy-mm-dd HH24:Mi:SS') As 占用日期, a.填制人, a.审核人" & vbNewLine & _
                    "From 药品收发记录 a" & vbNewLine & _
                    "Where  a.入出系数 = -1 And a.库房id = [1] And a.药品id = [2] And nvl(a.批次,0) = [3] And a.审核日期 is null And a.填制日期 > (sysdate - 30) And a.单据 not in (5,13)" & vbNewLine & _
                    "Order By a.填制日期 Desc )" & vbNewLine & _
                    "union all " & vbNewLine & _
                    "Select '合计' As 业务类型, sum(a.实际数量 * a.付数) As 发生数量, '', '', ''" & vbNewLine & _
                    "From 药品收发记录 a,收费项目目录 b" & vbNewLine & _
                    "Where a.入出系数 = -1 And a.药品id = b.id and a.库房id = [1] And a.药品id = [2] And nvl(a.批次,0) = [3] And a.审核日期 is null And a.填制日期 > (sysdate - 30)And a.单据 not in (5,13)"
    
    
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng库房ID, mlng药品ID, mlng批次)
        
        If rsTemp.RecordCount <= 1 Then Exit Sub
        
        With vsf库存变动
            vsf库存变动.TextMatrix(0, 2) = "占用数量"
            vsf库存变动.TextMatrix(0, 4) = "占用日期"
            
            Do While Not rsTemp.EOF
                .rows = .rows + 1
                .Row = .Row + 1
                
                .TextMatrix(.Row, .ColIndex("业务类型")) = rsTemp!业务类型
                
                If Not mbln区分大小单位 Then
                    .TextMatrix(.Row, .ColIndex("发生数量")) = zlStr.FormatEx(IIf(IsNull(rsTemp!占用数量), 0, rsTemp!占用数量) / mdbl比例系数, mintNumberDigit, , True) & mstr单位
                Else
                    If rsTemp!占用数量 < 0 Then .TextMatrix(.Row, .ColIndex("发生数量")) = "-"
                    
                    int大包装数量 = Int(Abs(IIf(IsNull(rsTemp!占用数量), 0, rsTemp!占用数量)) / mdbl比例系数)
                    .TextMatrix(.Row, .ColIndex("发生数量")) = .TextMatrix(.Row, .ColIndex("发生数量")) & zlStr.FormatEx(int大包装数量, mintNumberDigit, , True) & mstr单位
                    .TextMatrix(.Row, .ColIndex("发生数量")) = .TextMatrix(.Row, .ColIndex("发生数量")) & zlStr.FormatEx((Abs(IIf(IsNull(rsTemp!占用数量), 0, rsTemp!占用数量)) - int大包装数量 * mdbl比例系数) / mdbl比例系数小, mintNumberDigit, , True) & mstr单位小
                End If
                
                .TextMatrix(.Row, .ColIndex("发生日期")) = "" & rsTemp!占用日期
                .TextMatrix(.Row, .ColIndex("填制人")) = "" & rsTemp!填制人
                .TextMatrix(.Row, .ColIndex("审核人")) = "" & rsTemp!审核人
                
              rsTemp.MoveNext
    
            Loop
        End With
        
        
    End If
    
    vsf库存变动.Cell(flexcpFontBold, vsf库存变动.rows - 1, 0, vsf库存变动.rows - 1, vsf库存变动.Cols - 1) = True '最后一行字体加粗（合计行）
    vsf库存变动.TopRow = vsf库存变动.Row
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

