VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm带量供应商 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批量设置中选药品供应商"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10905
   Icon            =   "frm带量供应商.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd应用所有供应商 
      Caption         =   "应用所有药品(&Y)"
      Height          =   350
      Left            =   6000
      Picture         =   "frm带量供应商.frx":6852
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frm带量供应商.frx":699C
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   2
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   8520
      TabIndex        =   1
      Top             =   5880
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10680
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   830
      Width           =   10695
      _cx             =   18865
      _cy             =   8705
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm带量供应商.frx":6AE6
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
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "提示：提取属性设置了带量采购的药品，批量设置供应商"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   300
      Width           =   5325
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frm带量供应商.frx":6C00
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frm带量供应商"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mcstEditColor = &H80000003   '能编辑的颜色

Public Sub ShowMe(ByVal objFra As frmMediLists)
    Me.Show vbModal, objFra
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim arrSql() As Variant     '纪录存储过程的数组
    Dim blnTrans As Boolean
    Dim i As Integer

    On Error GoTo ErrHand:

    If vsfList.Rows < 2 Then Exit Sub
    If MsgBox("是否确定保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    arrSql = Array()
    
    With vsfList
        For i = 1 To .Rows - 1
            gstrSql = "Zl_带量采购供应商_Update("
            
            '药品id_In       In 药品规格.药品id%Type,
            gstrSql = gstrSql & Val(.TextMatrix(i, .ColIndex("药品ID")))
            
            '带量供应商id_In In 药品规格.带量供应商id%Type
            gstrSql = gstrSql & "," & IIf(Val(.TextMatrix(i, .ColIndex("供应商ID"))) = 0, "null", Val(.TextMatrix(i, .ColIndex("供应商ID"))))
            
            gstrSql = gstrSql & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSql
        Next
    End With

                
    gcnOracle.BeginTrans: blnTrans = True          '开启事务
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False     '提交事物
    
    MsgBox "保存成功！", vbOKOnly + vbInformation, gstrSysName
    
    Call FillVSF
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd帮助_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmd应用所有供应商_Click()
    Dim i As Integer
    If vsfList.Rows < 2 Then Exit Sub
    If vsfList.Row = 0 Then Exit Sub
    
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("供应商id")) = Val(.TextMatrix(.Row, .ColIndex("供应商id")))
            .TextMatrix(i, .ColIndex("供应商")) = .TextMatrix(.Row, .ColIndex("供应商"))
        Next
    End With
End Sub

Private Sub Form_Load()
    Call IniGrid
    Call FillVSF
End Sub

Private Sub IniGrid()
    With vsfList
        .Editable = flexEDNone
        .Rows = 1
        .ColWidth(0) = 350
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = 400
        .AllowSelection = False '不能多选
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
    End With
End Sub

Private Sub FillVSF()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSql = "Select a.药品id, b.编码, b.名称, b.规格, b.产地 As 生产商, c.名称 As 供应商,n.名称 As 商品名,c.id as 供应商id " & vbNewLine & _
                    "From 药品规格 A, 收费项目目录 B, 供应商 C, 收费项目别名 N" & vbNewLine & _
                    "Where a.药品id = b.Id And a.带量供应商id = c.Id(+) And b.Id = n.收费细目id(+)" & vbNewLine & _
                    "      And n.码类(+) = 1 And n.性质(+) = 3 And a.是否带量采购 = 1" & vbNewLine & _
                    "Order By b.编码"


    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "否带量采购药品")
    
    With vsfList
        .Rows = 1
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("药品ID")) = rsTemp!药品id
            .TextMatrix(.Rows - 1, .ColIndex("药品名称")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
            .TextMatrix(.Rows - 1, .ColIndex("商品名")) = nvl(rsTemp!商品名)
            .TextMatrix(.Rows - 1, .ColIndex("规格")) = rsTemp!规格
            .TextMatrix(.Rows - 1, .ColIndex("生产商")) = nvl(rsTemp!生产商)
            .TextMatrix(.Rows - 1, .ColIndex("供应商id")) = nvl(rsTemp!供应商id, 0)
            .TextMatrix(.Rows - 1, .ColIndex("供应商")) = nvl(rsTemp!供应商)
            
            .Cell(flexcpBackColor, .Rows - 1, .ColIndex("供应商"), .Rows - 1, .ColIndex("供应商")) = mcstEditColor
            
            rsTemp.MoveNext
        Loop
        
    End With
    
    Call VsfRowHeight(vsfList)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_EnterCell()
    
    With vsfList
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstEditColor Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    With vsfList
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("药品ID"))) = 0 Then Exit Sub
        
        If KeyCode = vbKeyDelete Then
            If MsgBox("是否确定删除第" & .Row & "行数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            .RemoveItem .Row
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("序号")) = i
            Next
        End If
  
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsfList.ColIndex("供应商") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsfList_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsfList.hwnd) '获取位置
    dblLeft = vRect.Left + vsfList.CellLeft
    dblTop = vRect.Top + vsfList.CellTop + vsfList.CellHeight + 3300
    With vsfList
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Col = .ColIndex("供应商") And .EditText = "" Then .TextMatrix(.Row, .ColIndex("供应商id")) = 0: Exit Sub
        If Col = .ColIndex("供应商") Then
        
            gstrSql = "Select id,编码,名称,简码 From 供应商 " & _
                      "Where (站点 = [3] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                      "  And 末级=1 And substr(类型,1,1)=1 " & _
                      "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] Or zlSpellCode(名称) Like [2] Or zlWbCode(名称) Like [2])" & _
                      "  Start with 上级ID is null and (站点 = [3] Or 站点 is Null) connect by prior ID =上级ID and (站点 = [3] Or 站点 is Null) "
    
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "供药单位", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, gstrMatch & UCase(.EditText) & "%", UCase(.EditText) & "%", gstrNodeNo)
    
            If blnCancel = True Then Exit Sub  '打开选择器时，点Esc不做以下处理
  
            If rsRecord Is Nothing Then
                MsgBox "没有找到该供应商！", vbInformation, gstrSysName
                Exit Sub
            Else
                .EditText = rsRecord!名称
                .TextMatrix(.Row, .ColIndex("供应商")) = rsRecord!名称
                .TextMatrix(.Row, .ColIndex("供应商id")) = rsRecord!ID
            End If
            
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfList
        If .Col = .ColIndex("供应商") Then
            .ColComboList(.ColIndex("供应商")) = "|..."
        Else
            .ColComboList(.ColIndex("供应商")) = ""
        End If
    End With
End Sub

Private Sub vsfList_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsfList
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        .EditMaxLength = 50
    End With
End Sub


Private Sub vsfList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsfList.hwnd) '获取位置
    dblLeft = vRect.Left + vsfList.CellLeft
    dblTop = vRect.Top + vsfList.CellTop + vsfList.CellHeight + 3300
    With vsfList
        If Col = .ColIndex("供应商") Then
            gstrSql = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
                      "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                      "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                      "Start with 上级ID is null and (站点 = [1] Or 站点 is Null) connect by prior ID =上级ID and (站点 = [1] Or 站点 is Null) order by level,ID"
                      
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 1, "供药单位", True, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, gstrNodeNo)
            
            If blnCancel = True Then Exit Sub '打开选择器时，点Esc不做以下处理
            
            If rsRecord Is Nothing Then
                Exit Sub
            Else
                .TextMatrix(.Row, .ColIndex("供应商")) = rsRecord!名称
                .TextMatrix(.Row, .ColIndex("供应商id")) = rsRecord!ID
            End If
        End If
    End With
    
End Sub

Private Sub VsfRowHeight(ByVal VsfObj As VSFlexGrid)
    Dim i As Long
    With VsfObj
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If .Rows < 2 Or .Row = 0 Then Exit Sub
        If Col = .ColIndex("供应商") And .EditText = "" Then
            .TextMatrix(.Row, .ColIndex("供应商id")) = 0
        End If
    End With
End Sub


