VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendBodyActiveItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "活动项目"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBodyActiveItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCloumn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5955
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2610
         MaxLength       =   20
         TabIndex        =   5
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "frmCaseTendBodyActiveItem.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "确认"
         Top             =   2460
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "frmCaseTendBodyActiveItem.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "取消"
         Top             =   2460
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "选用(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   7
         Top             =   1515
         Width           =   1095
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "删除(&E)"
         Height          =   300
         Index           =   1
         Left            =   2430
         TabIndex        =   8
         Top             =   1845
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstColumnItems 
         Height          =   2490
         Left            =   60
         TabIndex        =   4
         Top             =   480
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "项目序号"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "项目名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "部位"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstColumnUsed 
         Height          =   2490
         Left            =   3855
         TabIndex        =   6
         Top             =   480
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "项目序号"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "项目名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "部位"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找"
         Height          =   180
         Left            =   2160
         TabIndex        =   11
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已发生数据,不允许删除."
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   945
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可选的护理活动项目"
         Height          =   180
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已经选择的护理活动项目"
         Height          =   180
         Left            =   3855
         TabIndex        =   2
         Top             =   180
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyActiveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjVsf As Object
Private mblnChage As Boolean
Private mstrSQL As String
Private mstrActiveItem As String
Private mblnInit As Boolean
Private mlng护理等级 As Long
Private mlng婴儿 As Long
Private mlng科室ID As Long

Private Enum TYPE_Tab
    COL_tab分组名 = 0
    COL_tab字符串 = 1
    COL_tab项目序号 = 2
    COL_tab项目名 = 3  '--不包含单位
    COL_TabNull = 4
    COL_tab项目名称 = 5 '--包含单位
End Enum

Public Function ShowMe(objVsf As Object, ByVal frmParent As Form, ByVal lng护理等级 As Long, ByVal lng婴儿 As Long, ByVal lng科室ID As Long) As Boolean
    mblnChage = False
    mstrActiveItem = ""
    Set mobjVsf = objVsf
    mlng护理等级 = lng护理等级
    mlng婴儿 = lng婴儿
    mlng科室ID = lng科室ID
    If Not BoundItems Then Exit Function
    lblNote.Visible = False
    mblnInit = True
    Me.Show 1, frmParent
    ShowMe = mblnChage
End Function

Private Sub cmdFilterCancel_Click()
    mblnChage = False
    Unload Me
End Sub

Private Sub cmdFilterOK_Click()
'
    Dim intItem As Integer, intRow As Integer, i As Integer
    Dim lngItemCode As Integer, strItemName As String
    Dim blnAdd As Boolean, blnDelete As Boolean
    Dim strPart As String
    Dim arrStr() As String
    Dim arrTmp() As String, varCode() As String
    
    arrTmp = Split(mstrActiveItem, ";")
    
    '添加活动项目
    For intItem = 1 To lstColumnUsed.ListItems.Count
        lngItemCode = Val(lstColumnUsed.ListItems(intItem).Text)
        strItemName = lstColumnUsed.ListItems(intItem).SubItems(1)
        strPart = lstColumnUsed.ListItems(intItem).SubItems(2)
        blnAdd = True
        For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
            If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 Then
                If lngItemCode = Val(mobjVsf.TextMatrix(intRow, COL_tab项目序号)) And strItemName = mobjVsf.TextMatrix(intRow, COL_tab项目名) Then
                    blnAdd = False
                    Exit For
                End If
            End If
        Next intRow
        
        If blnAdd = True Then
            mblnChage = True
            For i = LBound(arrTmp) To UBound(arrTmp)
                varCode = Split(arrTmp(i), "'")
                If Val(varCode(2)) = lngItemCode And varCode(4) = strItemName Then
                    mobjVsf.Rows = mobjVsf.Rows + 1
                    arrStr = Split(varCode(1), ",")
                    If UBound(arrStr) > 6 Then arrStr(7) = strPart
                    varCode(1) = Join(arrStr, ",")
                    mobjVsf.TextMatrix(intRow, COL_tab分组名) = varCode(0)
                    mobjVsf.TextMatrix(intRow, COL_tab字符串) = varCode(1)
                    mobjVsf.TextMatrix(intRow, COL_tab项目序号) = lngItemCode
                    mobjVsf.TextMatrix(intRow, COL_tab项目名) = strItemName
                    mobjVsf.TextMatrix(intRow, COL_TabNull) = ""
                    mobjVsf.TextMatrix(intRow, COL_tab项目名称) = varCode(3)
                    '定位到新添加的行
                    mobjVsf.Row = mobjVsf.Rows - 1: mobjVsf.Col = mobjVsf.FixedCols
                End If
            Next i
        End If
    Next intItem
    '主要处理可能没有绑定固定项目的情况
    If mobjVsf.Rows > mobjVsf.FixedRows + 1 And mobjVsf.Tag = "NO" Then
        mobjVsf.Tag = ""
        Call mobjVsf.RemoveItem(mobjVsf.FixedRows)
    End If
    '删除活动项目
    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If intRow > mobjVsf.Rows - 1 Then Exit For
        If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 Then
            lngItemCode = Val(mobjVsf.TextMatrix(intRow, COL_tab项目序号))
            strItemName = mobjVsf.TextMatrix(intRow, COL_tab项目名)
            blnDelete = True
            For intItem = 1 To lstColumnUsed.ListItems.Count
                If lngItemCode = Val(lstColumnUsed.ListItems(intItem).Text) And strItemName = lstColumnUsed.ListItems(intItem).SubItems(1) Then
                    blnDelete = False
                    Exit For
                End If
            Next intItem
            
            If blnDelete = True Then
                mblnChage = True
                If mobjVsf.Rows = mobjVsf.FixedRows + 1 And intRow = mobjVsf.FixedRows Then
                    '主要处理可能没有绑定固定项目的情况
                    mobjVsf.Cell(flexcpText, intRow, 0, intRow, mobjVsf.Cols - 1) = ""
                    varCode = Split("',0,0,0,0,1,0,,0'-999''", "'")
                    mobjVsf.TextMatrix(intRow, COL_tab分组名) = varCode(0)
                    mobjVsf.TextMatrix(intRow, COL_tab字符串) = varCode(1)
                    mobjVsf.TextMatrix(intRow, COL_tab项目序号) = varCode(2)
                    mobjVsf.TextMatrix(intRow, COL_tab项目名) = varCode(4)
                    mobjVsf.TextMatrix(intRow, COL_TabNull) = ""
                    mobjVsf.TextMatrix(intRow, COL_tab项目名称) = varCode(3)
                    mobjVsf.Tag = "NO"
                    '定位到新添加的行
                    mobjVsf.Row = mobjVsf.Rows - 1: mobjVsf.Col = mobjVsf.FixedCols
                Else
                    Call mobjVsf.RemoveItem(intRow)
                    intRow = intRow - 1
                End If
            End If
        End If
    Next intRow
    
    Unload Me
End Sub


Private Function BoundItems() As Boolean
'---------------------------------------------------------------------
'功能:提取活动项目信息
'---------------------------------------------------------------------
    Dim lstItem As ListItem
    Dim rsActive As New ADODB.Recordset
    Dim arrActive() As String, ArrCode() As String
    Dim strSQL As String, strText As String
    Dim i As Integer, j As Integer
    Dim strItemCode As String, str值域 As String
    Dim intRow As Integer
    On Error GoTo Errhand
    
    If mobjVsf Is Nothing Then Exit Function
    
    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 Then
            strText = ""
            strText = "" & mobjVsf.TextMatrix(intRow, COL_tab项目序号) & " 项目序号,'" & mobjVsf.TextMatrix(intRow, COL_tab项目名) & "' 项目名称,1 标识"
            strSQL = strSQL & IIf(strSQL = "", "", "UNION ALL") & " SELECT " & strText & " FROM Dual "
        End If
    Next intRow
    
    mstrSQL = "" & _
            "Select a.项目序号, a.项目名称,a.部位 ,a.项目值域,a.项目性质,a.项目类型, a.项目长度, a.项目小数, a.记录频次,a.入院首测, a.分组名,a.项目单位, a.项目表示," & vbNewLine & _
            IIf(strSQL = "", "0 标识", "            Nvl(b.标识, 0) 标识") & vbNewLine & _
            "From (Select a.项目序号, c.部位 || b.项目名称 项目名称,c.部位, b.项目值域, b.项目类型, b.项目长度, b.项目小数," & vbNewLine & _
            "                           Nvl(a.记录频次, 2) 记录频次,A.入院首测, b.分组名, b.项目表示,b.项目性质,b.项目单位" & vbNewLine & _
            "            From 体温记录项目 a, 体温部位 c, 护理记录项目 b" & vbNewLine & _
            "            Where a.项目序号 = b.项目序号 And b.项目序号 = c.项目序号(+) And b.项目性质 = 2 And Nvl(b.应用方式, 0) = 1 And" & vbNewLine & _
            "                        b.护理等级 >= [1] And Nvl(b.适用病人, 0) In (0, [2]) And" & vbNewLine & _
            "                        (b.适用科室 = 1 Or" & vbNewLine & _
            "                        (b.适用科室 = 2 And Exists (Select 1 From 护理适用科室 d Where d.项目序号 = b.项目序号 And d.科室id = [3])))) a"
    If strSQL <> "" Then
        mstrSQL = mstrSQL & vbNewLine & ",(" & strSQL & ") b" & vbNewLine & _
            "Where a.项目序号 = b.项目序号(+) And a.项目名称 = b.项目名称(+)"
    End If
    mstrSQL = mstrSQL & vbNewLine & "   Order By a.项目序号, a.项目名称"
            
    Set rsActive = zlDatabase.OpenSQLRecord(mstrSQL, "提取未设置的活动项目", mlng护理等级, IIf(mlng婴儿 = 0, 1, 2), mlng科室ID)
    
    If rsActive.RecordCount = 0 Then
        MsgBox "没有可供选择的活动项目，请在护理项目管理模块中进行设置！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '加入活动项目
    txtFind.Text = ""
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    strItemCode = ""
    mstrActiveItem = ""
    
    With rsActive
        Do While Not .EOF
            str值域 = zlCommFun.Nvl(!项目值域)
            If zlCommFun.Nvl(!项目类型) = 0 Then
                If InStr(1, str值域, ";") <> 0 Then str值域 = Split(str值域, ";")(0) & "～" & Split(str值域, ";")(1)
            End If
            str值域 = Replace(Replace(Replace(str值域, ";", ":"), "'", ""), ",", "")
            If strItemCode = "" Then
                strItemCode = !项目序号 & "'" & Nvl(!项目名称)
                mstrActiveItem = zlCommFun.Nvl(!分组名, "2)体温表格项目") & "'" & str值域 & "," & zlCommFun.Nvl(!项目类型) & "," & _
                    zlCommFun.Nvl(!项目小数) & "," & zlCommFun.Nvl(!记录频次) & "," & zlCommFun.Nvl(!项目表示) & "," & zlCommFun.Nvl(!项目性质) & "," & _
                    zlCommFun.Nvl(!项目长度) & "," & zlCommFun.Nvl(!部位) & "," & zlCommFun.Nvl(!入院首测, 0) & "'" & _
                    zlCommFun.Nvl(!项目序号) & "'" & Replace(zlCommFun.Nvl(!项目名称) & IIf(zlCommFun.Nvl(!项目单位, "") = "", "", "(" & !项目单位 & ")"), ";", ":") & "'" & zlCommFun.Nvl(!项目名称)

            Else
                If InStr(1, "," & strItemCode & ",", "," & !项目序号 & "'" & Nvl(!项目名称) & ",") = 0 Then
                    strItemCode = strItemCode & "," & !项目序号 & "'" & Nvl(!项目名称)
                    mstrActiveItem = mstrActiveItem & ";" & zlCommFun.Nvl(!分组名, "2)体温表格项目") & "'" & str值域 & "," & zlCommFun.Nvl(!项目类型) & "," & _
                        zlCommFun.Nvl(!项目小数) & "," & zlCommFun.Nvl(!记录频次) & "," & zlCommFun.Nvl(!项目表示) & "," & zlCommFun.Nvl(!项目性质) & "," & _
                        zlCommFun.Nvl(!项目长度) & "," & zlCommFun.Nvl(!部位) & "," & zlCommFun.Nvl(!入院首测, 0) & "'" & _
                        zlCommFun.Nvl(!项目序号) & "'" & Replace(zlCommFun.Nvl(!项目名称) & IIf(zlCommFun.Nvl(!项目单位, "") = "", "", "(" & !项目单位 & ")"), ";", ":") & "'" & zlCommFun.Nvl(!项目名称)
                End If
            End If
            
            If !标识 = 1 Then
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !项目序号 & "_" & lstColumnUsed.ListItems.Count, !项目序号)
                lstItem.SubItems(1) = zlCommFun.Nvl(!项目名称)
                lstItem.SubItems(2) = zlCommFun.Nvl(!部位)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !项目序号 & "_" & lstColumnItems.ListItems.Count + 100, !项目序号)
                lstItem.SubItems(1) = zlCommFun.Nvl(!项目名称)
                lstItem.SubItems(2) = zlCommFun.Nvl(!部位)
            End If
            .MoveNext
        Loop
    End With
    
    BoundItems = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnItems_DblClick
End Sub

Private Sub lstColumnUsed_DblClick()
    Call cmdColumn_Click(1)
End Sub

Private Sub lstColumnUsed_ItemClick(ByVal Item As MSComctlLib.ListItem)
        '检查是否存在数据,存在数据则提示用不不允许删除
    If Not Item Is Nothing Then
        If CheckGridData(Val(Item.Text), Item.SubItems(1)) Then
            lblNote.Caption = Item.SubItems(1) & "已经发生数据,不能进行删除."
            lblNote.Visible = True
            cmdColumn(1).Enabled = False
        Else
            lblNote.Caption = ""
            lblNote.Visible = False
            cmdColumn(1).Enabled = True
        End If
    End If
End Sub

Private Sub lstColumnUsed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnUsed_DblClick
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim lstItem As ListItem
    
    If Index = 0 Then
        'add
        If Not lstColumnItems.SelectedItem Is Nothing Then
            Set lstItem = lstColumnUsed.ListItems.Add(, lstColumnItems.SelectedItem.Key, lstColumnItems.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnItems.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnItems.SelectedItem.SubItems(2)
            lstColumnItems.ListItems.Remove lstColumnItems.SelectedItem.Index
        End If
    Else
        'del
        If Not lstColumnUsed.SelectedItem Is Nothing Then
            If CheckGridData(Val(lstColumnUsed.SelectedItem.Text), lstColumnUsed.SelectedItem.SubItems(1)) = True Then Exit Sub
            Set lstItem = lstColumnItems.ListItems.Add(, lstColumnUsed.SelectedItem.Key, lstColumnUsed.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnUsed.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnUsed.SelectedItem.SubItems(2)
            lstColumnUsed.ListItems.Remove lstColumnUsed.SelectedItem.Index
        End If
    End If
End Sub

Private Function CheckGridData(ByVal lngID As Long, ByVal strName As String) As Boolean
'-------------------------------------------------------------------
'检查当天活动项目是否存在数据,有数据则不允许删除
'参数:lngID 项目序号 strName 项目名称（项目名称+部位）
'-------------------------------------------------------------------
    CheckGridData = True
    Dim intRow As Integer, intCOl As Integer

    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If Val(mobjVsf.TextMatrix(intRow, COL_tab项目序号)) = lngID And mobjVsf.TextMatrix(intRow, COL_tab项目名称) = strName Then
            Exit For
        End If
    Next intRow
    
    If intRow > mobjVsf.Rows - 1 Then CheckGridData = False: Exit Function
    
    '检查活动项目列是否存在数据
    For intCOl = mobjVsf.FixedCols To Val(Split(mobjVsf.TextMatrix(intRow, COL_tab字符串), ",")(3)) + mobjVsf.FixedCols - 1  '记录频次+固定列
        If Trim(mobjVsf.TextMatrix(intRow, intCOl)) <> "" Then
            Exit Function
        End If
    Next intCOl
    
    CheckGridData = False
End Function

Private Sub txtFind_Change()
    Call txtFind_KeyDown(10, 0)
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = 100
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Static lngPreIndex As Long
    Dim strText As String
    Dim lngIndex As Long
    
    '61855:刘鹏飞,2013-11-07,绑定活动项目怎么加搜索功能
    strText = Trim(txtFind.Text)
    If KeyCode = 10 Or strText = "" Then
        '主要是用于清除变量值
        lngPreIndex = 0
    ElseIf KeyCode = vbKeyReturn And strText <> "" Then
        If Not (lngPreIndex > 0 And lngPreIndex < lstColumnItems.ListItems.Count) Then lngPreIndex = 1
        For lngIndex = lngPreIndex To lstColumnItems.ListItems.Count
            If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                lstColumnItems.ListItems(lngIndex).Selected = True
                lstColumnItems.ListItems(lngIndex).EnsureVisible
                Exit For
            End If
        Next
        
        If lngIndex > lstColumnItems.ListItems.Count Then
            If lngPreIndex > 1 Then
                For lngIndex = 1 To lstColumnItems.ListItems.Count
                    If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                        lstColumnItems.ListItems(lngIndex).Selected = True
                        lstColumnItems.ListItems(lngIndex).EnsureVisible
                        Exit For
                    End If
                Next
            End If
            lngPreIndex = 1
        Else
            lngPreIndex = lngIndex + 1
        End If
    End If
End Sub

