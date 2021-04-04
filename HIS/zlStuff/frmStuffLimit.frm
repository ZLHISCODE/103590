VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffLimit 
   Caption         =   "卫材储备设置"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmStuffLimit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9060
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3105
      Left            =   3285
      TabIndex        =   14
      Top             =   6015
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5477
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -15
      TabIndex        =   6
      Top             =   4350
      Width           =   9810
      Begin VB.CommandButton cmd应用于本列所有卫材 
         Caption         =   "应用于本列(&O)"
         Height          =   350
         Left            =   3990
         TabIndex        =   13
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "恢复(&R)"
         Height          =   350
         Left            =   2685
         Picture         =   "frmStuffLimit.frx":058A
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全部清除(&C)"
         Height          =   350
         Left            =   1395
         Picture         =   "frmStuffLimit.frx":06D4
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   5700
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   90
         Picture         =   "frmStuffLimit.frx":081E
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭(&X)"
         Height          =   350
         Left            =   6810
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
   End
   Begin ZL9BillEdit.BillEdit msfLimit 
      Height          =   2655
      Left            =   75
      TabIndex        =   1
      Top             =   1380
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -195
      TabIndex        =   5
      Top             =   1035
      Width           =   9810
   End
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      Text            =   "cboRoom"
      Top             =   585
      Width           =   3360
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5715
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffLimit.frx":0968
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.TextBox txt查找 
      Height          =   300
      Left            =   6480
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lbl查找 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5760
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    卫材库房(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   645
      Width           =   1350
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmStuffLimit.frx":11FA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    选择库房后，指定该库房卫生材料的储备限量；并根据卫生材料的管理要求，可以同时指定其盘点属性和库房货位。"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   135
      Width           =   7725
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "卫生材料在各库房的限额与盘点要求(&T)："
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   3330
   End
End
Attribute VB_Name = "frmStuffLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前材质：由me.tag保存4
'   2、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序传入
'   3、指定卫材：由me.lblMedi.tag保存，由上级程序传入可以传递，也可以不传递
'---------------------------------------------------
Public strPrivs As String       '当前用户具有的本程序权限

Dim mobjItem As ListItem
Dim mLngCount As Long

Dim mrsTemp As New ADODB.Recordset
Private Const col编码 As Integer = 1
Private Const col名称 As Integer = 2
Private Const col规格 As Integer = 3
Private Const col产地 As Integer = 4
Private Const col成本价 As Integer = 5
Private Const col零售价 As Integer = 6
Private Const col库存数量 As Integer = 7
Private Const col单位 As Integer = 8
Private Const col包装 As Integer = 9
Private Const col下限 As Integer = 10
Private Const col上限 As Integer = 11
Private Const col日盘 As Integer = 12
Private Const col周盘 As Integer = 13
Private Const col月盘 As Integer = 14
Private Const col季盘 As Integer = 15
Private Const col货位 As Integer = 16
Private mlngFind As Long
Private mblnFind As Boolean             '是否查询到值
Private mblnFindFrist As Boolean        '重来没有找到过数据
Private mblnNoClick As Boolean

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub cboRoom_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long
    
    err = 0: On Error GoTo ErrHand
    
    If mblnNoClick Then Exit Sub
    
    gstrSQL = "Select 工作性质,nvl(服务对象,0) as 服务对象 From 部门性质说明" & _
            " Where 部门id=[1]"
    gstrSQL = gstrSQL & " and  工作性质 In ('卫材库','发料部门', '虚拟库房') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex))
    
    With rsTemp
        Me.cboRoom.Tag = "售价"
        Do While Not .EOF
            If InStr(1, !工作性质, "卫材库") > 0 Or InStr(1, !工作性质, "虚拟库房") > 0 Then
                Me.cboRoom.Tag = "卫材库"
                Exit Do
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        
        Do While Not .EOF
            If InStr(1, !工作性质, "发料部门") > 0 And (!服务对象 = 1) Then
                Me.cboRoom.Tag = "门诊"
                Exit Do
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, !工作性质, "发料部门") > 0 And (!服务对象 = 2 Or !服务对象 = 3) Then Me.cboRoom.Tag = "住院": Exit Do
            .MoveNext
        Loop
    End With
    
    Call zlLimitRef
    
    With msfLimit
        If .Rows = 2 And .RowData(1) = 0 Then
            For lngCol = 0 To .Cols - 1
                .ColData(lngCol) = 5
            Next
        ElseIf .Rows > 2 Then
            .Redraw = False
            .ColData(col上限) = 4
            .ColData(col下限) = 4
            .ColData(col日盘) = -1
            .ColData(col周盘) = -1
            .ColData(col月盘) = -1
            .ColData(col季盘) = -1
            .ColData(col货位) = 1
            .SetColColor col上限, vbWhite
            .SetColColor col下限, vbWhite
            .SetColColor col日盘, vbWhite
            .SetColColor col周盘, vbWhite
            .SetColColor col月盘, vbWhite
            .SetColColor col季盘, vbWhite
            .SetColColor col货位, vbWhite
            .SetRowColor 0, &H8000000F
            .Redraw = True
        End If
        For lngCol = 0 To .Cols - 1
            If .ColData(lngCol) = 5 Then
                .SetColColor lngCol, &H8000000F
            End If
        Next
        .SetFocus
        .Col = col下限
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRoom.ListCount = 0 Then Call zlControl.ControlSetFocus(msfLimit): Exit Sub
    
    If cboRoom.ListIndex >= 0 Then
        If Val(cboRoom.Tag) = cboRoom.ItemData(cboRoom.ListIndex) Then
            Call zlControl.ControlSetFocus(msfLimit, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboRoom, Trim(cboRoom.Text), "V,K,12,W", IIf(InStr(1, strPrivs, "允许设置所有库房限额盘点") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cboRoom.ListIndex >= 0 Then
        cboRoom.Tag = cboRoom.ItemData(cboRoom.ListIndex)
    End If
End Sub


Private Sub cboRoom_LostFocus()
    Dim i As Long
    If cboRoom.ListCount = 0 Then Exit Sub
    If cboRoom.ListIndex < 0 Then
        For i = 0 To cboRoom.ListCount - 1
            If Val(cboRoom.Tag) = cboRoom.ItemData(i) Then
                mblnNoClick = True
                cboRoom.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub


Private Sub cmdClear_Click()
    With Me.msfLimit
        .Redraw = False
        For mLngCount = 1 To .Rows - 1
            .TextMatrix(mLngCount, 0) = ""
            If InStr(1, strPrivs, "上下限控制") > 0 Then
                .TextMatrix(mLngCount, col下限) = Format(0, mFMT.FM_数量)
                .TextMatrix(mLngCount, col上限) = Format(0, mFMT.FM_数量)
            End If
            If InStr(1, strPrivs, "盘点属性设置") > 0 Then
                .TextMatrix(mLngCount, col日盘) = ""
                .TextMatrix(mLngCount, col周盘) = ""
                .TextMatrix(mLngCount, col月盘) = ""
                .TextMatrix(mLngCount, col季盘) = ""
            End If
            .TextMatrix(mLngCount, col货位) = ""
        Next
        .Redraw = True
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlLimitRef
End Sub

Private Sub CmdSave_Click()
    Dim strMsgBox As String, strErrors As String
    
    strErrors = ""
    With Me.msfLimit
        For mLngCount = 1 To .Rows - 1
            If Val(.TextMatrix(mLngCount, col上限)) <> 0 _
                And Val(.TextMatrix(mLngCount, col上限)) < Val(.TextMatrix(mLngCount, col下限)) Then
                .TextMatrix(mLngCount, 0) = "？"
                strErrors = strErrors & vbCrLf & .TextMatrix(mLngCount, col编码) & "-" & .TextMatrix(mLngCount, col名称)
                strMsgBox = "“" & .TextMatrix(mLngCount, col编码) & "-" & .TextMatrix(mLngCount, col名称) & "”的储备下限大于储备上限！" & _
                        vbCrLf & vbCrLf & "继续保存其他卫材吗？"
                If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbThis.Panels(2).Text = ""
                    .MsfObj.TopRow = mLngCount: .Row = mLngCount: .SetFocus: Exit Sub
                End If
            ElseIf .RowData(mLngCount) <> 0 Then
                gstrSQL = "zl_卫生材料储备限额_Update(" & Me.cboRoom.ItemData(Me.cboRoom.ListIndex)
                gstrSQL = gstrSQL & "," & .RowData(mLngCount)
                gstrSQL = gstrSQL & "," & Round(Val(.TextMatrix(mLngCount, col上限)) * Val(.TextMatrix(mLngCount, col包装)), g_小数位数.obj_散装小数.数量小数)
                gstrSQL = gstrSQL & "," & Round(Val(.TextMatrix(mLngCount, col下限)) * Val(.TextMatrix(mLngCount, col包装)), g_小数位数.obj_散装小数.数量小数)
                gstrSQL = gstrSQL & ",'" & IIf(Trim(.TextMatrix(mLngCount, col日盘)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col周盘)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col月盘)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col季盘)) = "", "0", "1")
                gstrSQL = gstrSQL & "','" & Trim(.TextMatrix(mLngCount, col货位)) & "')"
                err = 0: On Error Resume Next
                zldatabase.ExecuteProcedure gstrSQL, Me.Caption
                
                If err <> 0 Then
                    Call SaveErrLog
                    err = 0
                    .TextMatrix(mLngCount, 0) = "？"
                    strErrors = strErrors & vbCrLf & .TextMatrix(mLngCount, col编码) & "-" & .TextMatrix(mLngCount, col名称)
                    strMsgBox = "保存“" & .TextMatrix(mLngCount, col编码) & .TextMatrix(mLngCount, col名称) & "”时发生错误！" & _
                            vbCrLf & vbCrLf & "继续保存其他卫材吗？"
                    If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Me.stbThis.Panels(2).Text = ""
                        .MsfObj.TopRow = mLngCount: .Row = mLngCount: .SetFocus: Exit Sub
                    End If
                End If
                If mLngCount Mod IIf(.Rows > 20, .Rows \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "正在保存：" & String(mLngCount \ IIf(.Rows > 20, .Rows \ 20, 1), "…")
                End If
            End If
        Next
    End With
    Me.stbThis.Panels(2).Text = ""
    strMsgBox = "“" & Me.cboRoom.Text & "”储备特性保存完毕！"
    If strErrors <> "" Then
        strMsgBox = strMsgBox & vbCrLf & "但以下卫材发生错误，请检查：" & strErrors
    End If
    MsgBox strMsgBox, vbExclamation, gstrSysName
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmd应用于本列所有卫材_Click()
    Dim lngRow As Long, lngRows As Long
    Dim strValue As String
    '将当前列的内容应用到所有药品相同列
    lngRows = msfLimit.Rows - 1
    strValue = msfLimit.TextMatrix(msfLimit.Row, msfLimit.Col)
    For lngRow = 1 To lngRows
        msfLimit.TextMatrix(lngRow, msfLimit.Col) = strValue
    Next
    msfLimit.SetFocus
End Sub

Private Sub Form_Activate()
    Dim lngCol As Long
    
    With Me.msfLimit
        .Cols = 17: .MsfObj.FixedCols = 1
        .TextMatrix(0, col编码) = "编码": .TextMatrix(0, col名称) = "名称"
        .TextMatrix(0, col规格) = "规格": .TextMatrix(0, col产地) = "产地"
        .TextMatrix(0, col单位) = "单位": .TextMatrix(0, col包装) = "包装":
        .TextMatrix(0, col成本价) = "成本价": .TextMatrix(0, col零售价) = "零售价": .TextMatrix(0, col库存数量) = "库存数量":
        .TextMatrix(0, col下限) = "下限": .TextMatrix(0, col上限) = "上限"
        .TextMatrix(0, col日盘) = "日盘": .TextMatrix(0, col周盘) = "周盘": .TextMatrix(0, col月盘) = "月盘": .TextMatrix(0, col季盘) = "季盘"
        .TextMatrix(0, col货位) = "货位"

        .ColWidth(0) = 250: .ColWidth(col编码) = 900: .ColWidth(col名称) = 2200
        .ColWidth(col规格) = 1500: .ColWidth(col产地) = 1200
        .ColWidth(col单位) = 500: .ColWidth(col包装) = 0: .ColWidth(col成本价) = 1200: .ColWidth(col零售价) = 1200: .ColWidth(col库存数量) = 1200
        If InStr(1, strPrivs, "上下限控制") > 0 Then
            .ColWidth(col上限) = 855: .ColWidth(col下限) = 855
        Else
            .ColWidth(col上限) = 0: .ColWidth(col下限) = 0
        End If
        If InStr(1, strPrivs, "盘点属性设置") > 0 Then
            .ColWidth(col日盘) = 500: .ColWidth(col周盘) = 500: .ColWidth(col月盘) = 500: .ColWidth(col季盘) = 500
        Else
            .ColWidth(col日盘) = 0: .ColWidth(col周盘) = 0: .ColWidth(col月盘) = 0: .ColWidth(col季盘) = 0
        End If
        .ColWidth(col货位) = 1700
        
        .ColAlignment(col编码) = 1: .ColAlignment(col名称) = 1
        .ColAlignment(col规格) = 1: .ColAlignment(col产地) = 1
        .ColAlignment(col单位) = 4: .ColAlignment(col包装) = 7: .ColAlignment(col成本价) = 7: .ColAlignment(col零售价) = 7: .ColAlignment(col库存数量) = 7
        .ColAlignment(col上限) = 7: .ColAlignment(col下限) = 7
        .ColAlignment(col日盘) = 4: .ColAlignment(col周盘) = 4: .ColAlignment(col月盘) = 4: .ColAlignment(col季盘) = 4
        .ColAlignment(col货位) = 1
        
        .ColData(col编码) = 5: .ColData(col名称) = 5
        .ColData(col规格) = 5: .ColData(col产地) = 5
        .ColData(col单位) = 5: .ColData(col包装) = 5: .ColData(col成本价) = 5: .ColData(col零售价) = 5: .ColData(col库存数量) = 5
        If InStr(1, strPrivs, "上下限控制") > 0 Then
            .ColData(col上限) = 4: .ColData(col下限) = 4
        Else
            .ColData(col上限) = 5: .ColData(col下限) = 5
        End If
        If InStr(1, strPrivs, "盘点属性设置") > 0 Then
            .ColData(col日盘) = -1: .ColData(col周盘) = -1: .ColData(col月盘) = -1: .ColData(col季盘) = -1
        Else
            .ColData(col日盘) = 5: .ColData(col周盘) = 5: .ColData(col月盘) = 5: .ColData(col季盘) = 5
        End If
        .ColData(col货位) = 1
        .PrimaryCol = col单位:
        
        '需确定相关权限
        If InStr(1, strPrivs, "上下限控制") = 0 And InStr(1, strPrivs, "上下限控制") = 0 Then
        Else
            .LocateCol = IIf(InStr(1, strPrivs, "上下限控制") <> 0, col上限, col日盘)
        End If
                    
        .Row = 1: .Col = IIf(InStr(1, strPrivs, "上下限控制") <> 0, col上限, col日盘)
    End With
    
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfLimit.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfLimit.Active = True
    End If
    
    err = 0: On Error GoTo ErrHand
    
    gstrSQL = "Select ID, 编码, 名称" & _
              "  From 部门表 D" & _
               " Where ID In (Select Distinct 部门id" & _
                "             From 部门性质说明 a" & _
                            " Where 工作性质 In ('发料部门', '物资库房', '卫材库', '制剂室', '虚拟库房')) And" & _
                     " exists (Select 1  b From 诊疗执行科室 b where d.id=b.执行科室id) and (d.撤档时间 is null or to_char(d.撤档时间,'yyyy-mm-dd')='3000-01-01')"
    If InStr(1, strPrivs, "允许设置所有库房限额盘点") = 0 Then
        gstrSQL = gstrSQL & "      and ID in (select 部门ID from 部门人员 R where R.人员ID=[1])"
    End If
    
    Set mrsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id)
    
    With mrsTemp
        Me.cboRoom.Clear
        Do While Not .EOF
            Me.cboRoom.AddItem !编码 & "-" & !名称
            Me.cboRoom.ItemData(Me.cboRoom.NewIndex) = !Id
            .MoveNext
        Loop
    End With
    If Me.cboRoom.ListCount <= 0 Then
        MsgBox "未设置相关的库房，无法设置储备限量", vbExclamation, gstrSysName
        Unload Me: Exit Sub
    End If
    Me.cboRoom.ListIndex = 0

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(1, g_成本价)
        .FM_金额 = GetFmtString(1, g_金额)
        .FM_零售价 = GetFmtString(1, g_售价)
        .FM_数量 = GetFmtString(1, g_数量)
    End With
    lbl查找.Visible = True
    txt查找.Visible = True
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.fraLine.Left = 0: Me.fraLine.Width = Me.ScaleWidth + 100
    Me.msfLimit.Left = 0: Me.msfLimit.Width = Me.ScaleWidth
    Me.msfLimit.Height = Me.ScaleHeight - Me.msfLimit.Top - Me.fraFunc.Height - Me.stbThis.Height
    Me.fraFunc.Left = 0: Me.fraFunc.Width = Me.ScaleWidth: Me.fraFunc.Top = Me.msfLimit.Top + Me.msfLimit.Height
    Me.cmdClose.Left = Me.fraFunc.Width - Me.cmdClose.Width - 90
    Me.cmdSave.Left = Me.cmdClose.Left - Me.cmdSave.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFindFrist = False
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        msfLimit.SetFocus
        msfLimit.Col = col货位
        Cancel = True
        Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub msfLimit_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub msfLimit_EditKeyPress(KeyAscii As Integer)
    If InStr("'!@#$%^&*|-""", Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub msfLimit_EnterCell(Row As Long, Col As Long)
    Dim lngCol As Long
    With msfLimit
      Select Case .Col
          Case col上限, col下限
              .TxtCheck = True
              .MaxLength = 15
              .TextMask = ".1234567890"
          Case col货位
               ImeLanguage True
              .MaxLength = 20
              .TextMask = ""
              .TxtSetFocus
          Case Else
              .TxtCheck = False
          End Select
          For lngCol = 0 To .Cols - 1
              If .ColData(lngCol) = 5 Then
                  .SetColColor lngCol, &H8000000F
              End If
          Next
      End With
End Sub

Private Sub msfLimit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Me.msfLimit
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
        If .TextMatrix(.Row, col上限) = "" Then
            .Text = " "
            .TextMatrix(.Row, col上限) = .Text
            .TextMatrix(.Row, col下限) = .Text
            Exit Sub
        End If
        
        If Trim(.Text) = "" Then
            .Text = IIf(.TextMatrix(.Row, .Col) = "", " ", IIf(.TxtVisible, .Text, .TextMatrix(.Row, .Col)))
            .TextMatrix(.Row, .Col) = .Text
        Else
            If .Col = col货位 Then
                If LenB(StrConv(.Text, vbFromUnicode)) > 50 Then
                    MsgBox "货位超长！最多50个字符或25个汉字", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
            Else
                If Not IsNumeric(.Text) Then
                    MsgBox "输入中含有非法字符！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
                If Val(.Text) < 0 Then
                    MsgBox "库存上下限不能小于零！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
                If Val(.Text) > 10000000000000# Then
                    MsgBox "输入值超过最大值！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
            End If
        End If
        
        Select Case .Col
        Case col上限
            .Text = Format(.Text, mFMT.FM_数量): .TextMatrix(.Row, col上限) = .Text
        Case col下限
            .Text = Format(.Text, mFMT.FM_数量): .TextMatrix(.Row, col下限) = .Text
        Case col货位
                '无处理
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, col货位) = ""
                    End If
                    Exit Sub
                Else
                    Dim strTemp As String
                    strTemp = GetMatchingSting(strKey)
                    
                    gstrSQL = " Select 编码,名称 From 材料库房货位 " & _
                          " Where (编码 Like [1]" & _
                          "     Or 名称 Like [1]" & _
                          "     Or 简码 Like [1])"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
                
                    
                    If rsTemp.EOF Then
                        If MsgBox("没有找到数据，是否增加名称为[" & .Text & "]的库房货位？ ", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                            .Text = ""
                            .TxtSetFocus
                            Cancel = True
                            Exit Sub
                        Else
                            .AllowAddRow = False
                            '货位_IN    IN 材料库房货位.名称%Type
                            gstrSQL = "zl_卫材库房货位_Update('" & strKey & "')"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            If .Cols - 1 = col货位 And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                                Exit Sub
                            End If
                        End If
                    Else
                        .AllowAddRow = False
                        If rsTemp.RecordCount = 1 Then
                            .TextMatrix(.Row, col货位) = rsTemp.Fields("名称")
                            .Text = rsTemp.Fields("名称")
                            If .Cols - 1 = col货位 And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                                Exit Sub
                            End If
                        Else
                            Set mshSelect.Recordset = rsTemp
                            Call setSelectLocal
                            Cancel = True
                            If .Cols - 1 = col货位 And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                End If
                OS.OpenIme False
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlLimitRef()
    '--------------------------------------------------------
    '功能：刷新库存限额
    '--------------------------------------------------------
    err = 0: On Error GoTo ErrHand
    gstrSQL = "Select  i.Id, i.编码, i.名称, i.规格, i.产地, i.包装单位 As 单位, i.换算系数 As 包装," & vbNewLine & _
                    "       Nvl(l.上限, 0) / i.换算系数 As 上限,Nvl(l.下限, 0) /i.换算系数 As 下限, l.盘点属性, l.库房货位," & vbNewLine & _
                    "       nvl(k.实际数量,0)/i.换算系数 as 库存数量," & vbNewLine & _
                    "       Decode(nvl(k.实际数量,0),0,i.成本价,(k.实际金额-k.实际差价) / k.实际数量)*i.换算系数 as 成本价," & vbNewLine & _
                    "       Decode(i.是否变价, 0, p.现价,Decode(nvl(k.实际数量,0),0,nvl(i.上次售价,p.现价),k.实际金额 / k.实际数量))*i.换算系数 as 零售价" & vbNewLine & _
                    "From (Select i.是否变价, i.Id, i.编码, i.名称, i.规格, i.产地, i.计算单位, s.包装单位," & vbNewLine & _
                    "            Decode(s.换算系数, 0, 1, Null, 1, s.换算系数) as 换算系数,s.成本价,s.上次售价" & vbNewLine & _
                    "       From 收费项目目录 I, 材料特性 S, (Select Distinct 诊疗项目id From 诊疗执行科室 Where 执行科室id =[1]) E" & vbNewLine & _
                    "       Where i.Id = s.材料id And s.诊疗id = e.诊疗项目id And i.类别 = '4' And" & vbNewLine & _
                    "             (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))) I," & vbNewLine & _
                    "     (Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 L Where 库房id =[1]) L," & vbNewLine & _
                    "     (Select 药品id, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价" & vbNewLine & _
                    "       From 药品库存" & vbNewLine & _
                    "       Where 性质 = 1 And 库房id =[1]" & vbNewLine & _
                    "       Group By 药品id) K, 收费价目 P" & vbNewLine & _
                    "Where i.Id = p.收费细目id And i.Id = l.材料id(+) And i.Id = k.药品id(+) And" & vbNewLine & _
                    "      (p.终止日期 Is Null Or Sysdate Between p.执行日期 And Nvl(p.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) And" & vbNewLine & _
                    "      p.价格等级 Is Null" & vbNewLine & _
                    "Order By i.编码"

    Set mrsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex))
        
    With mrsTemp
        Me.msfLimit.ClearBill
        Me.msfLimit.Redraw = False
        Do While Not .EOF
            If Me.msfLimit.Rows < .AbsolutePosition + 1 Then Me.msfLimit.Rows = Me.msfLimit.Rows + 1
            Me.msfLimit.RowData(.AbsolutePosition) = !Id
            Me.msfLimit.TextMatrix(.AbsolutePosition, col编码) = !编码
            Me.msfLimit.TextMatrix(.AbsolutePosition, col名称) = !名称
            Me.msfLimit.TextMatrix(.AbsolutePosition, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col单位) = IIf(IsNull(!单位), "", !单位)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col包装) = zlStr.Nvl(!包装)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col上限) = Format(!上限, mFMT.FM_数量)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col下限) = Format(!下限, mFMT.FM_数量)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col日盘) = IIf(Mid(!盘点属性, 1, 1) = "1", "√", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col周盘) = IIf(Mid(!盘点属性, 2, 1) = "1", "√", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col月盘) = IIf(Mid(!盘点属性, 3, 1) = "1", "√", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col季盘) = IIf(Mid(!盘点属性, 4, 1) = "1", "√", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col货位) = IIf(IsNull(!库房货位), "", !库房货位)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col成本价) = Format(!成本价, mFMT.FM_成本价)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col零售价) = Format(!零售价, mFMT.FM_零售价)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col库存数量) = Format(!库存数量, mFMT.FM_数量)
            
            If .AbsolutePosition Mod IIf(.RecordCount > 20, .RecordCount \ 20, 1) = 0 Then
                Me.stbThis.Panels(2).Text = "正在提取：" & String(.AbsolutePosition \ IIf(.RecordCount > 20, .RecordCount \ 20, 1), "…")
            End If
            .MoveNext
        Loop
        Me.msfLimit.Redraw = True
    End With
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt查找_Change()
    mblnFindFrist = False
End Sub

Private Sub txt查找_GotFocus()
    zlControl.TxtSelAll txt查找
End Sub

Private Sub txt查找_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String, lngStart As Long, lngRows As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim strTemp编码 As String, strTemp名称 As String, strTemp简码 As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txt查找.Text))
    If strInput = "" Then Exit Sub
    
    '查找药品
    If strInput = txt查找.Tag Then
        '表示查找下一条记录
        If mlngFind >= msfLimit.Rows - 1 And mblnFindFrist = True Then
            MsgBox "已查询到最后！", vbInformation, gstrSysName
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '表示新的查找
        lngStart = 0
        txt查找.Tag = strInput
    End If
    
    '开始查找
    lngStart = lngStart + 1
    lngRows = msfLimit.Rows - 1
    mblnFind = False
    For lngStart = lngStart To lngRows
        str编码 = Trim(UCase(msfLimit.TextMatrix(lngStart, col编码)))
        str名称 = Trim(UCase(msfLimit.TextMatrix(lngStart, col名称)))
        str简码 = UCase(zlStr.GetCodeByVB(str名称))
        If str编码 Like "*" & strInput & "*" Or _
            str名称 Like "*" & strInput & "*" Or _
            str简码 Like "*" & strInput & "*" Then
            msfLimit.Row = lngStart
            msfLimit.MsfObj.TopRow = lngStart
            msfLimit.SetFocus
            mblnFind = True '记录已经查询到值
            mblnFindFrist = True
            Exit For
        End If
    Next

    mlngFind = lngStart
    If mlngFind = lngRows + 1 And mblnFind = False And mblnFindFrist = False Then
        MsgBox "没有找到你想要的数据！", vbInformation, gstrSysName
        zlControl.TxtSelAll txt查找
        Exit Sub
    End If
End Sub

Private Sub txt查找_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'Private Sub cmdFind_Click()
'    Dim blnVisible As Boolean
'    '查找或查找下一条
'    blnVisible = lbl查找.Visible Xor True
'    lbl查找.Visible = blnVisible
'    txt查找.Visible = blnVisible
'    If blnVisible Then txt查找.SetFocus
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txt查找.Visible And KeyCode = vbKeyF3 Then
        Call txt查找_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub msfLimit_CommandClick()
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 编码,名称,简码 From 材料库房货位 Order by 编码"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "提取材料库房货位")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "材料库房货位还未初始化！[字典管理]", vbInformation, gstrSysName
        Exit Sub
    End If
    With msfLimit
        If rsTemp.RecordCount = 1 Then
            .TextMatrix(.Row, col货位) = rsTemp.Fields("名称")
            .Text = rsTemp.Fields("名称")
        Else
            Set mshSelect.Recordset = rsTemp
            Call setSelectLocal
            Exit Sub
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    With msfLimit
        If KeyCode = vbKeyEscape Then
            mshSelect.Visible = False
            .SetFocus
        End If
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = mshSelect.TextMatrix(mshSelect.Row, 1)
            mshSelect.Visible = False
            .Col = col货位
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            End If
            .SetFocus
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    If mshSelect.Visible Then
        mshSelect.Visible = False
    End If
End Sub
Private Sub setSelectLocal()
    '功能:设置选择器的位置
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngTemp As Single
    sngLeft = msfLimit.Left + msfLimit.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = msfLimit.Top + msfLimit.MsfObj.CellTop + msfLimit.MsfObj.CellHeight
    With mshSelect
        .Redraw = False
        If sngLeft + .Width > Me.ScaleWidth Then
            If sngLeft - .Width < 0 Then
                sngLeft = 0
            Else
                sngLeft = sngLeft - .Width + msfLimit.MsfObj.CellWidth
            End If
        End If
        sngTemp = sngTop - msfLimit.MsfObj.CellHeight
        
        If Me.ScaleHeight - sngTop > sngTemp Then
            .Height = Me.ScaleHeight - sngTop
        Else
            .Height = IIf(sngTemp < 0, 0, sngTemp)
            sngTemp = sngTop
            sngTop = sngTop - .Height
            
        End If
        If .Rows * .RowHeight(0) + (.Rows * 15) + .RowHeight(0) <= .Height Then
            .Height = .Rows * .RowHeight(0) + (.Rows * 15) + .RowHeight(0)
            If msfLimit.Top + msfLimit.MsfObj.CellTop > sngTop Then
                sngTop = msfLimit.Top + msfLimit.MsfObj.CellTop - .Height
            End If
        End If
        
        .Left = sngLeft
        .Top = sngTop
        .ColWidth(0) = 1000
        .ColWidth(1) = IIf(.Width - .ColWidth(0) - 15 < 0, 500, .Width - .ColWidth(0) - 15)
        .Row = 1
        .Col = 0
        .TopRow = 1
        .ColSel = .Cols - 1
        .Visible = True
        .SetFocus
        .Redraw = True
        Exit Sub
    End With
End Sub
