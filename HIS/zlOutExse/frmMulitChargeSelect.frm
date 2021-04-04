VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMulitChargeSelect 
   Caption         =   "门诊收费单据选择"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11985
   Icon            =   "frmMulitChargeSelect.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11985
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   570
      ScaleHeight     =   765
      ScaleWidth      =   12465
      TabIndex        =   1
      Top             =   7185
      Width           =   12465
      Begin VB.PictureBox picNoInfo 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   9225
         TabIndex        =   4
         Top             =   210
         Width           =   9225
         Begin VB.TextBox txtInvoiceNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   3240
         End
         Begin VB.TextBox txtCurTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.TextBox txtAllTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lblInvoice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发票信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4875
            TabIndex        =   10
            Top             =   75
            Width           =   960
         End
         Begin VB.Label lblCurTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前单据"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   9
            Top             =   60
            Width           =   960
         End
         Begin VB.Label lblAllTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所属单据"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2385
            TabIndex        =   8
            Top             =   60
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9255
         TabIndex        =   3
         ToolTipText     =   "热键F2,右键弹出保存为划价单(或按CTRL+S)"
         Top             =   195
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10770
         TabIndex        =   2
         ToolTipText     =   "热键:Esc"
         Top             =   195
         Width           =   1440
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   5835
      Left            =   -525
      TabIndex        =   0
      Top             =   900
      Width           =   9510
      _cx             =   16775
      _cy             =   10292
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMulitChargeSelect.frx":0442
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmMulitChargeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNOs As String
Private mlngModule As Long
Private mblnUnLoad As Boolean
Private mblnNOMoved As Boolean
Private mblnOk As Boolean
Private mstrNo As String
Private mstrShowInVoiceNo As String
Private mblnOldDelSelect As Boolean

Private Function LoadData(ByVal strNos As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-12 16:40:13
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str医嘱序号 As String, i As Long, strTemp As String
    Dim cur合计 As Double, strReturnNos As String
    Dim strNoTemp As String
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    strSQL = "" & _
    " Select A.结帐ID,A.NO,A.记录状态,Nvl(A.价格父号,A.序号)  as 序号,A.收费类别,A.执行部门ID,A.开单部门ID, A.收费细目ID," & _
    "           A.费用类型,A.计算单位,A.医嘱序号,A.费别,A.从属父号," & _
    "          Avg(Nvl(A.付数,1)) as 付数,Avg(Nvl(A.数次,0)) as 数次,Sum(A.标准单价) as 单价,sum(A.应收金额) as 应收金额,sum(A.实收金额) as 实收金额," & _
    "          max(Decode(A.记录状态,2,NULL,A.操作员姓名)) as  操作员姓名,Max(decode(A.记录状态,2,NULL,A.登记时间)) as 登记时间,max(decode(A.记录状态,2,NULL,A.摘要)) as 摘要" & _
    " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,Table(f_Str2list([1])) J" & _
    " Where A.记录性质=1  And A.NO=J.Column_Value " & _
    "  Group by A.NO,A.记录状态,A.结帐ID,Nvl(A.价格父号,A.序号),A.收费类别 ,A.执行部门ID,A.开单部门ID, A.收费细目ID,A.费用类型,A.计算单位,A.医嘱序号,A.费别,A.从属父号"
    
    strSQL = _
    " Select  A.NO,A.序号,A.从属父号,A.费别," & _
    "        A.收费细目ID,C.编码 as 类别码,C.名称 as 类别名,B.编码,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
        IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
    "       Sum(Nvl(A.付数,1)*A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 剩余数量," & _
    "       Max(A.单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       D.名称 as 执行科室,E.名称 as 开单科室,Max(A.操作员姓名) as 操作员姓名,Max(A.登记时间) as 登记时间, " & _
    "       Max(A.摘要) as 摘要,Max(A.医嘱序号) as 医嘱序号" & _
    " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X" & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
    "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+)  " & _
    " Group by  A.NO,A.序号 ,A.从属父号,A.费别,A.收费细目ID,C.编码,C.名称,B.编码,B.名称," & _
    "       B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,E.名称,X.药品ID,X." & gstr药房单位
    
    strSQL = "" & _
    "   Select /*+ rule */ " & _
    "        A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.编码,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
    "       A.计算单位,A.医嘱序号 ,A.收费细目ID,A.剩余数量,A.单价,A.应收金额,A.实收金额," & _
    "       A.执行科室,A.开单科室,A.操作员姓名,A.登记时间, A.摘要 , M.医嘱内容  as 医嘱内容 " & _
    "   From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1,病人医嘱记录 M" & _
    "   Where       nvl(A.剩余数量,0)<>0 And A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    "       And A.医嘱序号=M.ID(+)  " & _
    " Order by A.NO,A.序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))

    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        strNoTemp = ""
        For i = 1 To rsTemp.RecordCount
            If InStr(strReturnNos & ",", "," & Nvl(rsTemp!NO) & ",") = 0 Then
                strReturnNos = strReturnNos & "," & Nvl(rsTemp!NO)
            End If
            .Cell(flexcpData, i, .ColIndex("项目")) = Nvl(rsTemp!从属父号)
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTemp!医嘱序号) & "," & Nvl(rsTemp!收费细目ID)
            If Val(Nvl(rsTemp!医嘱序号)) <> 0 And InStr(str医嘱序号 & ",", "," & Nvl(rsTemp!医嘱序号) & ",") = 0 Then
                str医嘱序号 = str医嘱序号 & "," & Nvl(rsTemp!医嘱序号)
            End If
            strTemp = ""
            If Val(Nvl(rsTemp!从属父号)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "┣"
                If rsTemp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTemp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
    
            .RowData(i) = CLng(rsTemp!序号)
            .TextMatrix(i, .ColIndex("单据号")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("类别")) = rsTemp!类别名
            .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称 & IIf(IsNull(rsTemp!规格), "", " " & rsTemp!规格)
            .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTemp!商品名)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(Val(Nvl(rsTemp!剩余数量)), 5)
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(rsTemp!单价, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(rsTemp!应收金额, gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(rsTemp!实收金额, gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTemp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTemp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = Nvl(rsTemp!操作员姓名)
            .TextMatrix(i, .ColIndex("时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("结帐ID")) = 0
            .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsTemp!医嘱内容)
            .TextMatrix(i, .ColIndex("原始数量")) = 0
            .TextMatrix(i, .ColIndex("准退数量")) = 0
            .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            If InStr(strNoTemp & ",", "," & rsTemp!NO & ",") = 0 Then
                '画出分隔线
                If strNoTemp <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                strNoTemp = strNoTemp & "," & rsTemp!NO
            End If
            cur合计 = cur合计 + rsTemp!实收金额
            rsTemp.MoveNext
        Next
        If .Rows <= 1 Then .Rows = 2
        .Row = .FixedRows: .Col = .ColIndex("项目")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .SelectionMode = flexSelectionByRow
        .Redraw = flexRDBuffered
    End With
    txtAllTotal.Text = Format(cur合计, gstrDec)
    Screen.MousePointer = 0
    If strReturnNos <> "" Then strReturnNos = Mid(strReturnNos, 2)
    '没有可选单据或者只有一张单据时，退出
    If strReturnNos = "" Or InStr(strReturnNos, ",") = 0 Then mstrNo = strReturnNos: mblnOk = True: Exit Function
    LoadData = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
End Function
Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化表头列信息
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-11 09:47:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    strHead = "单据号,1000,1;类别,720,1;项目,2800,1;商品名,2000,1;数量,750,7;单位,550,1;单价,1100,7;" & _
        "应收金额,1100,7;实收金额,1100,7;开单科室,1000,1;执行科室,1000,1;操作员,850,1;时间,1260,1;结帐ID,0,0;医嘱,1560,1;" & _
        "原始数量,0,4;准退数量,0,4;医嘱序号,0,4"
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = varTemp(0)
            .TextMatrix(.FixedRows - 1, intCol) = varTemp(0)
            If UBound(varTemp) > 0 Then
                .ColHidden(intCol) = False
                .ColWidth(intCol) = Val(varTemp(1))
                If .ColWidth(intCol) = 0 Then .ColHidden(intCol) = True
                .ColAlignment(intCol) = Val(varTemp(2))
            Else
                .ColHidden(intCol) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .COLS - 1) = 4
        .ColHidden(.ColIndex("商品名")) = gTy_System_Para.byt药品名称显示 <> 2
        .FrozenCols = 2
    End With
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    mstrNo = ""
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim strNo As String
    With vsBill
        If .Row < 1 Then Exit Sub
        strNo = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
        If strNo = "" Then Exit Sub
    End With
    mstrNo = strNo
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    picNoInfo.Visible = Not mblnOldDelSelect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Form_Load()
    mblnUnLoad = False
    txtInvoiceNo.Text = mstrShowInVoiceNo
    Call InitBillHead
    Call RestoreWinState(Me, App.ProductName)
    If LoadData(mstrNOs) = False Then mblnUnLoad = True: Unload Me
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Left = ScaleLeft
        .Top = ScaleHeight - .Height
        .Width = ScaleWidth
        
    End With
    With vsBill
        .Left = ScaleLeft + 50
        .Top = ScaleTop
        .Height = picDown.Top - .Top
        .Width = ScaleWidth - .Left * 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur合计 As Currency, i As Long
    If NewRow <> OldRow Then
        With vsBill
            If .TextMatrix(NewRow, .ColIndex("单据号")) <> "" Then
                For i = NewRow - 1 To .FixedRows Step -1
                    If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
                    cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                Next
                For i = NewRow To .Rows - 1
                    If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
                    cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                Next
            End If
            txtCurTotal.Text = Format(cur合计, gstrDec)
        End With
    End If
End Sub
 
Public Function zlShowSelect(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strNos As String, _
    strShowInVoiceNo As String, ByRef strNo As String, _
    Optional blnOldDelSelect As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对传入单据,选择其中一张单据
    '入参:frmMain-调用的主窗体
    '       strNos-单据号,用逗号分离:A0001,A0002
    '       strShowInVoiceNo-显示的发票号
    '出参:strNO-返回选中的单据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-04-12 17:41:39
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrNOs = strNos: mblnOldDelSelect = blnOldDelSelect
    mstrShowInVoiceNo = strShowInVoiceNo
    Screen.MousePointer = 0: mblnOk = False
    Err = 0: On Error Resume Next
    Me.Show 1, frmMain
    Screen.MousePointer = 11
    strNo = mstrNo
    zlShowSelect = mblnOk
End Function

Private Sub vsBill_DblClick()
    Call cmdOK_Click
End Sub
