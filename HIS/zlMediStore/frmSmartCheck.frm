VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmSmartCheck 
   Caption         =   "盘点表智能检查"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "frmSmartCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8355
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   7815
      _cx             =   13785
      _cy             =   7646
      Appearance      =   0
      BorderStyle     =   0
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSmartCheck.frx":6852
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
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "无库存且未盘点"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         ToolTipText     =   "检查最近的盘点表中重复盘点的药品"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6000
         TabIndex        =   9
         Text            =   "3"
         Top             =   105
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   270
         Left            =   6375
         TabIndex        =   8
         Top             =   105
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDay"
         BuddyDispid     =   196612
         OrigLeft        =   4200
         OrigTop         =   120
         OrigRight       =   4455
         OrigBottom      =   375
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "检查(&C)"
         Height          =   300
         Left            =   6840
         TabIndex        =   6
         Top             =   90
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "盘点重复"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "检查最近的盘点表中重复盘点的药品"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "盘点遗漏"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "检查最近的盘点表中盘漏的药品"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "检查天数："
         Height          =   180
         Left            =   5160
         TabIndex        =   4
         Top             =   150
         Width           =   900
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "检查类型："
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5415
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1235
            Text            =   "检查进度"
            TextSave        =   "检查进度"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12594
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSmartCheck.frx":68D6
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSmartCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mcon打印 As Integer = 103
Private Const mcon预览 As Integer = 102
Private Const mcon退出 As Integer = 191

Private mcbrToolBar As CommandBar
Private mobjPopup As CommandBar
Private mblnSuccess As Boolean '单据是否发生变化


Private mlng库房ID As Long
Private mfraPar As Form

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case mcon打印
            cbsFilePrint
        Case mcon预览
            cbsFilePreView
        Case mcon退出
            Unload Me
    End Select
End Sub

Private Sub cbsFilePrint()
    '打印
    vsfGrid.Redraw = flexRDNone
    subPrint 1
    vsfGrid.Redraw = flexRDDirect
    vsfGrid.Col = 0
    vsfGrid.ColSel = vsfGrid.Cols - 1
End Sub

Private Sub cbsFilePreView()
    '打印预览
    vsfGrid.Redraw = flexRDNone
    subPrint 2
    vsfGrid.Redraw = flexRDDirect
    vsfGrid.Col = 0
    vsfGrid.ColSel = vsfGrid.Cols - 1
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "检查结果"
        
    objRow.Add "时间：" & zlDataBase.Currentdate
    objRow.Add "部门：" & mfraPar.cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfGrid
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub


Private Sub cmdCheck_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strNo As String
    Dim lng药品id As Long
    Dim lng批次 As Long
    Dim str单据状态 As String
    Dim str提示信息 As String
    Dim lng盘点Sum As Long
    Dim lng药品Sum As Long
    Dim lng药品Sum2 As Long
    Dim str提示药品 As String
    Dim lng提示药品sum As Long
    Dim str漏盘药品 As String
    Dim lng漏盘药品sum As Long
    Dim bln变色 As Boolean
    Dim str漏盘药品信息 As String
    
    On Error GoTo ErrHandle
    vsfGrid.Clear
    vsfGrid.rows = 1
    vsfGrid.RowHeight(-1) = 300
    
    '加载最近几天未冲销的盘点表
    gstrSQL = "Select a.No, a.库房id, a.药品id, Nvl(a.批次, 0) 批次, a.记录状态, a.产地,b.规格, a.批号, b.编码, b.名称, Decode(a.审核日期, Null, Null, '已审') 单据状态," & vbNewLine & _
            "       a.填制人, a.填制日期, a.审核人, a.审核日期" & vbNewLine & _
            "From 药品收发记录 A, 收费项目目录 B" & vbNewLine & _
            "Where a.记录状态 = 1 And a.单据 = 12 And a.填制日期 > Sysdate - [2] And a.库房id = [1] And a.药品id = b.Id" & vbNewLine & _
            "Order By a.药品id, 批次"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng库房ID, Val(txtDay.Text))
    '加载该库房所有设置存储属性及有库存的药品
    gstrSQL = "Select a.*,b.编码,b.名称,b.规格" & vbNewLine & _
            "From (Select a.Id 药品id, Null 批次,Null 批号" & vbNewLine & _
            "       From 收费项目目录 A, 药品规格 B, 收费执行科室 C" & vbNewLine & _
            "       Where a.Id = b.药品id And a.Id = c.收费细目id And c.执行科室id = [1] And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From 药品库存 D" & vbNewLine & _
            "              Where a.Id = d.药品id And d.库房id = [1] And (实际数量 <> 0 or 实际金额 <> 0 or 实际差价 <> 0))" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Distinct a.药品id, Nvl(a.批次, 0) 批次,上次批号 批号" & vbNewLine & _
            "       From 药品库存 A" & vbNewLine & _
            "       Where a.库房id = [1] And (实际数量 <> 0 or 实际金额 <> 0 or 实际差价 <> 0)) A, 收费项目目录 B" & vbNewLine & _
            "Where a.药品id = b.Id" & vbNewLine & _
            "Order By 药品id, 批次"
    Set rsPhysic = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng库房ID)
    
    If rsTemp.RecordCount = 0 And rsPhysic.RecordCount = 0 Then
        staThis.Panels(2).Text = "已完成！"
        Exit Sub
    End If
    
    '进度条各总数
    lng药品Sum = IIf(chkType(0).Value = 0, 0, rsPhysic.RecordCount)
    lng盘点Sum = IIf(chkType(1).Value = 0, 0, rsTemp.RecordCount)
    lng药品Sum2 = IIf(chkType(2).Value = 0, 0, rsPhysic.RecordCount)
    
    If chkType(1).Value = 1 Then '重复盘点药品
        Do While Not rsTemp.EOF
            strNo = rsTemp!NO
            lng药品id = rsTemp!药品id
            lng批次 = nvl(rsTemp!批次, 0)
            str单据状态 = nvl(rsTemp!单据状态, "")
            str提示信息 = rsTemp!填制人 & "于" & rsTemp!填制日期 & "填制" & IIf(IsNull(rsTemp!审核人), "", "；" & rsTemp!审核人 & "于" & rsTemp!审核日期 & "审核")
            
            rsTemp.MoveNext
            
            If Not rsTemp.EOF Then
                If lng药品id = rsTemp!药品id And lng批次 = nvl(rsTemp!批次, 0) Then
                
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "盘点重复：单据【" & strNo & "】 药品【[" & rsTemp!编码 & "]" & rsTemp!名称 & "(" & rsTemp!规格 & ")" & "】" & "批号：" & IIf(IsNull(rsTemp!批号), "无", rsTemp!批号) & IIf(str单据状态 = "", "", "(" & str单据状态 & ")")
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 1) = strNo
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = lng药品id
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 3) = rsTemp!批次
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 4) = IIf(str单据状态 = "", 2, 4)
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 5) = str提示信息
                    vsfGrid.rows = vsfGrid.rows + 1
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "          单据【" & rsTemp!NO & "】 药品【[" & rsTemp!编码 & "]" & rsTemp!名称 & "(" & rsTemp!规格 & ")" & "】" & "批号：" & IIf(IsNull(rsTemp!批号), "无", rsTemp!批号) & IIf(IsNull(rsTemp!单据状态), "", "(" & rsTemp!单据状态 & ")")
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 1) = rsTemp!NO
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = rsTemp!药品id
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 3) = rsTemp!批次
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 4) = IIf(IsNull(rsTemp!单据状态), 2, 4)
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 5) = rsTemp!填制人 & "于" & rsTemp!填制日期 & "填制" & IIf(IsNull(rsTemp!审核人), "", "；" & rsTemp!审核人 & "于" & rsTemp!审核日期 & "审核")
                    
                    If Not bln变色 Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 2, 0, vsfGrid.rows - 1, 0) = &H8000000F
                    bln变色 = Not bln变色
                    
                    vsfGrid.rows = vsfGrid.rows + 1
                    
                End If
                
                Call zlControl.StaShowPercent(rsTemp.AbsolutePosition / (lng药品Sum + lng盘点Sum + lng药品Sum2), staThis.Panels(2), frmSmartCheck)
            End If
            
        Loop
    End If
    
    If chkType(0).Value = 1 Then '盘点漏掉药品
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            If Not IsNull(rsPhysic!批次) Then '有库存但在指定时间内盘点表中不存在的药品
                rsTemp.Filter = "药品id = " & rsPhysic!药品id & " And 批次 = " & rsPhysic!批次
                
                If rsTemp.RecordCount = 0 Then
                    If str漏盘药品信息 = "" Then
                        str漏盘药品信息 = rsPhysic!药品id & ":" & rsPhysic!批次
                    Else
                        str漏盘药品信息 = str漏盘药品信息 & ";" & rsPhysic!药品id & ":" & rsPhysic!批次
                    End If
                    
                    If str漏盘药品 = "" Then
                        lng漏盘药品sum = lng漏盘药品sum + 1
                        str漏盘药品 = "[" & rsPhysic!编码 & "]" & rsPhysic!名称 & "(" & rsPhysic!规格 & ")" & rsPhysic!批号
                    Else
                        lng漏盘药品sum = lng漏盘药品sum + 1
                        If lng漏盘药品sum <= 3 Then str漏盘药品 = str漏盘药品 & "、[" & rsPhysic!编码 & "]" & rsPhysic!名称 & "(" & rsPhysic!规格 & ")" & rsPhysic!批号
                    End If
                End If
                
            End If
            
            Call zlControl.StaShowPercent((rsPhysic.AbsolutePosition + lng盘点Sum) / (lng药品Sum + lng盘点Sum + lng药品Sum2), staThis.Panels(2), frmSmartCheck)
            
            rsPhysic.MoveNext
        Loop
        
        If lng漏盘药品sum = 0 Then
            vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "不存在盘点遗漏的药品"
            vsfGrid.rows = vsfGrid.rows + 1
        Else
            If lng漏盘药品sum > 0 Then
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "有账面数量未盘点的药品有:" & str漏盘药品 & IIf(lng漏盘药品sum > 3, "等" & lng漏盘药品sum & "个", "")
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = str漏盘药品信息
                vsfGrid.RowHeight(vsfGrid.rows - 1) = IIf(lng漏盘药品sum > 1, 600, 300)
                If Not bln变色 Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 1, 0, vsfGrid.rows - 1, 0) = &H8000000F
                vsfGrid.rows = vsfGrid.rows + 1
            End If
        End If
        
    End If
    
    If chkType(2).Value = 1 Then '无库存未盘点
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            If IsNull(rsPhysic!批次) Then '设置存储属性且无库存，只检查药品id不检查批次
                rsTemp.Filter = "药品id = " & rsPhysic!药品id
                
                If rsTemp.RecordCount = 0 Then
                    If str提示药品 = "" Then
                        lng提示药品sum = lng提示药品sum + 1
                        str提示药品 = "[" & rsPhysic!编码 & "]" & rsPhysic!名称 & "(" & rsPhysic!规格 & ")"
                    Else
                        lng提示药品sum = lng提示药品sum + 1
                        If lng提示药品sum <= 3 Then str提示药品 = str提示药品 & "、[" & rsPhysic!编码 & "]" & rsPhysic!名称 & "(" & rsPhysic!规格 & ")"
                    End If
                End If
            
            End If
            
            Call zlControl.StaShowPercent((rsPhysic.AbsolutePosition + lng盘点Sum + lng药品Sum) / (lng药品Sum + lng盘点Sum + lng药品Sum2), staThis.Panels(2), frmSmartCheck)
            
            rsPhysic.MoveNext
        Loop
        
        If lng提示药品sum = 0 Then
            vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "不存在无库存未盘点药品"
            vsfGrid.rows = vsfGrid.rows + 1
        Else
            If lng提示药品sum > 0 Then
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "无账面数量也未盘点的药品有:" & str提示药品 & IIf(lng提示药品sum > 3, "等" & lng提示药品sum & "个", "")
                vsfGrid.RowHeight(vsfGrid.rows - 1) = IIf(lng提示药品sum > 1, 600, 300)
                If Not bln变色 Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 1, 0, vsfGrid.rows - 1, 0) = &H8000000F
                vsfGrid.rows = vsfGrid.rows + 1
            End If
        End If
        
    End If
    
    staThis.Panels(2).Text = "已完成！"
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowME(ByVal lng库房ID As Long, ByVal fraPar As Form, ByRef blnSuccess As Boolean)
    mlng库房ID = lng库房ID
    Set mfraPar = fraPar
    
    Me.Show 1, fraPar
    
    blnSuccess = mblnSuccess
End Sub


Private Sub CmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Load()
        InitComandBars
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 8595 Then Me.Width = 8595
    If Me.Height < 6345 Then Me.Height = 6345
    
    picCondition.Width = Me.ScaleWidth
    cmdCheck.Left = Me.ScaleWidth - cmdCheck.Width - 300
    vsfGrid.Move vsfGrid.Left, vsfGrid.Top, Me.ScaleWidth, Me.ScaleHeight - picCondition.Height - staThis.Height - 500
    vsfGrid.ColWidth(0) = vsfGrid.Width - 10
End Sub


Private Sub txtDay_Change()
    If txtDay.Text > 30 Then txtDay.Text = 30
    If txtDay.Text < 0 Then txtDay.Text = 1
End Sub

Private Sub txtDay_GotFocus()
    txtDay.SelStart = 0
    txtDay.SelLength = Len(txtDay.Text)
    txtDay.SelText = txtDay.Text
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub vsfGrid_DblClick()
    Dim blnSuccess As Boolean
    
    If vsfGrid.TextMatrix(vsfGrid.Row, 1) <> "" Then
    
        MousePointer = vbHourglass
        With vsfGrid
            frmNewCheckCard.ShowCard mfraPar, .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 4)), 1, blnSuccess, Val(.TextMatrix(.Row, 2)), Val(.TextMatrix(.Row, 3))
        End With
        MousePointer = vbDefault
        
    ElseIf vsfGrid.TextMatrix(vsfGrid.Row, 2) <> "" Then
        With vsfGrid
            frmNewCheckCard.ShowCard mfraPar, .TextMatrix(.Row, 1), 9, 1, blnSuccess, , , .TextMatrix(.Row, 2)
        End With
    End If
    
    mblnSuccess = blnSuccess
    If blnSuccess Then cmdCheck_Click
End Sub

Private Sub InitComandBars()
    '初始化工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With

    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons
    
    
    '工具栏定义
    Set mcbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny Or xtpFlagHideWrap
    
    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mcon打印, "打印")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon预览, "预览")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        Set cbrControlMain = .Add(xtpControlButton, mcon退出, "退出")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
    End With

    cbsMain.Item(1).Delete
     
     '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_ESCAPE, mcon退出
    End With

End Sub

Private Sub vsfGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If vsfGrid.MouseRow >= 0 Then vsfGrid.ToolTipText = vsfGrid.TextMatrix(vsfGrid.MouseRow, 5)
End Sub
