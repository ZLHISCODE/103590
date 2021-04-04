VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBarCodePrint 
   Caption         =   "条码打印"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   Icon            =   "frmBarCodePrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   12375
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkNoPrint 
      Caption         =   "只显示未打印完单据"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   6405
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   12135
      _cx             =   21405
      _cy             =   11298
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBarCodePrint.frx":058A
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
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11160
      TabIndex        =   8
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   9840
      TabIndex        =   7
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "过滤"
      Height          =   300
      Left            =   11040
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboInputDate 
      Height          =   300
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   60
      Width           =   1695
   End
   Begin VB.ComboBox cboStore 
      Height          =   300
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtp开始时间 
      Height          =   315
      Left            =   7080
      TabIndex        =   11
      Top             =   60
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   162136067
      CurrentDate     =   36263
   End
   Begin MSComCtl2.DTPicker dtp结束时间 
      Height          =   315
      Left            =   8985
      TabIndex        =   12
      Top             =   60
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   162136067
      CurrentDate     =   36263
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "注意：条码打印只以散装单位打印"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2640
      TabIndex        =   14
      Top             =   600
      Width           =   2700
   End
   Begin VB.Label lbl至 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "～"
      Height          =   180
      Left            =   8760
      TabIndex        =   13
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblInputDate 
      AutoSize        =   -1  'True
      Caption         =   "入库时间"
      Height          =   180
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "库房"
      Height          =   180
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "frmBarCodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFormatNum As String '数量格式化语句
Private mlngStore As Long     '当前库房

Public Sub ShowMe(ByVal frmpar As Form, ByVal strFormatNum As String, ByVal objcboBox As ComboBox)
    Dim i As Integer
    
    mstrFormatNum = strFormatNum
    
    cboStore.Clear
    For i = 0 To objcboBox.ListCount - 1
        cboStore.AddItem objcboBox.List(i)
        cboStore.ItemData(cboStore.NewIndex) = objcboBox.ItemData(i)
        If objcboBox.List(i) = objcboBox.Text Then
            cboStore.ListIndex = cboStore.NewIndex
        End If
    Next
    
    With cboInputDate
        .AddItem "一天内"
        .AddItem "三天内"
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    Me.Show vbModal, frmpar
End Sub

Private Sub cboInputDate_Click()
    If cboInputDate.Text = "自定义" Then
        dtp开始时间.Visible = True
        lbl至.Visible = True
        dtp结束时间.Visible = True
    Else
        dtp开始时间.Visible = False
        lbl至.Visible = False
        dtp结束时间.Visible = False
    End If
End Sub

Private Sub chkNoPrint_Click()
    Call GetDetails
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    Call GetDetails
End Sub

Private Sub cmdPrint_Click()
    Dim i As Integer, j As Integer
    Dim blnRe As Boolean
    
    With vsfList
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("本次打印数量"))) > 0 Then
                
                For j = 1 To Val(.TextMatrix(i, .ColIndex("本次打印数量")))
                    If Trim(.TextMatrix(i, .ColIndex("商品条码"))) <> "" Then
                        blnRe = ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_4", Me, "库房=" & cboStore.Text & "|=  " & cboStore.ItemData(cboStore.ListIndex), "类型=1", "商品条码=" & .TextMatrix(i, .ColIndex("商品条码")), "内部条码=" & .TextMatrix(i, .ColIndex("内部条码")), 2)
                    Else
                        blnRe = ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_4", Me, "库房=" & cboStore.Text & "|=  " & cboStore.ItemData(cboStore.ListIndex), "类型=2", "商品条码=" & .TextMatrix(i, .ColIndex("内部条码")), "内部条码=" & .TextMatrix(i, .ColIndex("内部条码")), 2)
                    End If
                Next
                
                If blnRe = True Then
                    gstrSQL = ""
                    gstrSQL = "Zl_卫材条码打印记录_Update('" & _
                    .TextMatrix(i, .ColIndex("NO")) & "','" & _
                    .TextMatrix(i, .ColIndex("单据")) & "'," & _
                    .TextMatrix(i, .ColIndex("库房id")) & "," & _
                    .TextMatrix(i, .ColIndex("材料id")) & "," & _
                    .TextMatrix(i, .ColIndex("序号")) & "," & _
                    .TextMatrix(i, .ColIndex("本次打印数量")) & ")"
                    zlDatabase.ExecuteProcedure gstrSQL, "打印"
                End If
            End If
        Next
        
        If blnRe = True Then
            Call GetDetails
            MsgBox "打印完成！", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Form_Load()
    dtp开始时间.Value = DateAdd("d", -7, sys.Currentdate)
    dtp结束时间.Value = sys.Currentdate
    
    vsfList.Cell(flexcpForeColor, 0, vsfList.Cols - 1, 0, vsfList.Cols - 1) = vbBlue
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lblNO.Move 80, 120
    txtNO.Move lblNO.Left + lblNO.Width + 160, 60
    lblStore.Move txtNO.Left + txtNO.Width + 400, lblNO.Top
    cboStore.Move lblStore.Left + lblStore.Width + 160, 60
    lblInputDate.Move cboStore.Left + cboStore.Width + 400, lblNO.Top
    cboInputDate.Move lblInputDate.Left + lblInputDate.Width + 160, 60
    dtp开始时间.Move cboInputDate.Left + cboInputDate.Width + 200, 60
    lbl至.Move dtp开始时间.Left + dtp开始时间.Width + 75, 120
    dtp结束时间.Move lbl至.Left + lbl至.Width + 75, 60
    cmdFilter.Move Me.Width - cmdFilter.Width - 300, txtNO.Top
    
    chkNoPrint.Move lblNO.Left, txtNO.Top + txtNO.Height + 50
    lblMsg.Top = chkNoPrint.Top + (chkNoPrint.Height - lblMsg.Height) / 2
        
    CmdCancel.Move Me.Width - CmdCancel.Width - 300, Me.Height - CmdCancel.Height - 650
    cmdPrint.Move CmdCancel.Left - cmdPrint.Width - 50, CmdCancel.Top

    vsfList.Move lblNO.Left, chkNoPrint.Top + chkNoPrint.Height + 50, Me.Width - 400, Me.Height - chkNoPrint.Top - cmdPrint.Height - 1150
End Sub

Private Sub GetDetails()
    '提取数据
    Dim rsTemp As ADODB.Recordset
    Dim mdatBeginDate As Date
    Dim mdatEndDate As Date
    Dim intYear As Integer, strYear As String
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select a.名称, a.编码, a.规格,a.计算单位,c.包装单位,c.换算系数, b.No, b.单据, b.库房id, b.材料id, b.序号, b.商品条码, b.内部条码, b.入库数量, b.打印数量, b.入库时间" & vbNewLine & _
                "From 收费项目目录 A, 材料特性 C, 卫材条码打印记录 B" & vbNewLine & _
                "Where a.Id = c.材料id And a.Id = b.材料id And b.库房id = [1] And b.入库时间 Between [2] And [3]"


    If chkNoPrint.Value = 1 Then
        gstrSQL = gstrSQL & " and b.入库数量<>b.打印数量 "
    End If
    
    If Trim(txtNO.Text) <> "" Then
        gstrSQL = gstrSQL & " and b.no=[4]"
    End If
        
    With cboInputDate
        Select Case .Text
            Case "一天内"
                mdatBeginDate = CDate(Format(DateAdd("D", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = sys.Currentdate
            Case "三天内"
                mdatBeginDate = CDate(Format(DateAdd("D", -3, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = sys.Currentdate
            Case "一周内"
                mdatBeginDate = CDate(Format(DateAdd("D", -7, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = sys.Currentdate
            Case "一月内"
                mdatBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = sys.Currentdate
            Case "自定义"
                mdatBeginDate = CDate(Format(dtp开始时间, "yyyy-mm-dd") & " 00:00:00")
                mdatEndDate = CDate(Format(dtp结束时间, "yyyy-mm-dd") & " 23:59:59")
        End Select
    End With
    
    If Len(txtNO) < 8 And Len(txtNO) > 0 Then '处理单据号
        Me.txtNO = UCase(LTrim(Me.txtNO))
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txtNO) < 8 Then Me.txtNO = strYear & String(7 - Len(txtNO), "0") & Me.txtNO
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询明细", cboStore.ItemData(cboStore.ListIndex), mdatBeginDate, mdatEndDate, txtNO.Text)
    
    With vsfList
        .Rows = 1
        
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("材料名称")) = "【" & rsTemp!编码 & "】" & rsTemp!名称 & "-" & rsTemp!规格
            .TextMatrix(.Rows - 1, .ColIndex("单位")) = rsTemp!计算单位
            .TextMatrix(.Rows - 1, .ColIndex("商品条码")) = zlStr.NVL(rsTemp!商品条码)
            .TextMatrix(.Rows - 1, .ColIndex("内部条码")) = rsTemp!内部条码
            .TextMatrix(.Rows - 1, .ColIndex("入库数量")) = rsTemp!入库数量
            .TextMatrix(.Rows - 1, .ColIndex("已打印数量")) = rsTemp!打印数量
            .TextMatrix(.Rows - 1, .ColIndex("本次打印数量")) = Val(.TextMatrix(.Rows - 1, .ColIndex("入库数量"))) - Val(.TextMatrix(.Rows - 1, .ColIndex("已打印数量")))
            .TextMatrix(.Rows - 1, .ColIndex("单据")) = rsTemp!单据
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = rsTemp!序号
            .TextMatrix(.Rows - 1, .ColIndex("库房id")) = rsTemp!库房id
            .TextMatrix(.Rows - 1, .ColIndex("材料id")) = rsTemp!材料ID
            
            rsTemp.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    
    lng库房ID = cboStore.ItemData(cboStore.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txtNO) < 8 And Len(txtNO) > 0 Then
            txtNO.Text = zlCommFun.GetFullNO(txtNO.Text, 68, lng库房ID)
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        If .Col = .Cols - 1 Then '本次打印列
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfList
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Val(.EditText) > Val(.TextMatrix(Row, .ColIndex("入库数量"))) - Val(.TextMatrix(Row, .ColIndex("已打印数量"))) Then
            MsgBox "本次打印数量不能大于剩余数量！", vbInformation, gstrSysName
            Cancel = True
        End If
    End With
End Sub
