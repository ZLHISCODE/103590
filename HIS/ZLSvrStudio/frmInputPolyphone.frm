VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInputPolyphone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "多音字配置"
   ClientHeight    =   4335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7275
   Icon            =   "frmInputPolyphone.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6030
      TabIndex        =   5
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6030
      TabIndex        =   4
      Top             =   345
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   6015
      TabIndex        =   6
      Top             =   1500
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3900
      Left            =   60
      TabIndex        =   1
      Top             =   345
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   6879
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "数据表"
         Object.Width           =   2646
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3900
      Left            =   2310
      TabIndex        =   3
      Top             =   345
      Width           =   3645
      _cx             =   6429
      _cy             =   6879
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483628
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      Editable        =   2
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "多音字配置(&L)"
      Height          =   180
      Index           =   2
      Left            =   2295
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "可用数据表(&T)"
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "frmInputPolyphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private rsTmp As New ADODB.Recordset
Private mrsPolyphone As New ADODB.Recordset
Private mlngSys As Long
Private mblnOK As Boolean
Private mstrKey As String

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngSys As Long) As Boolean
    
    mblnOK = False
    mblnStartUp = True
    
    mlngSys = lngSys
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Sub AdjustRowFlag(ByRef bill As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To bill.Rows - 1
        bill.TextMatrix(lngLoop, 0) = ""
    Next
    
    bill.TextMatrix(intRow, 0) = "●"
    
End Sub

Private Sub InsertNewRow()
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    
    If vsf.Editable <> flexEDNone Then
        vsf.AddItem "", vsf.Rows
        
        vsf.Row = vsf.Rows - 1
        
    Else
        vsf.Row = vsf.Rows - 1
    End If
End Sub

Private Sub GoNextCell()
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-----------------------------------------------------------------------------------------
    If vsf.Col + 1 > vsf.Cols - 1 Then
        
        '换行之前，先检查是否允许换行，即是否有必输的项目没有输入
        If Trim(vsf.TextMatrix(vsf.Row, 1)) = "" Or Trim(vsf.TextMatrix(vsf.Row, 2)) = "" Then
            Exit Sub
        End If
                        
        If vsf.Row = vsf.Rows - 1 Then
            Call InsertNewRow
        Else
            vsf.Row = vsf.Row + 1
        End If
        
        '找第一个可以编辑的列
        vsf.Col = 1
    Else
        '找下一个可以编辑的列
        vsf.Col = vsf.Col + 1
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrKey = ""
    Call lvw_ItemClick(lvw.SelectedItem)
    
    lvw.Enabled = True
    cmdSave.Enabled = False
    cmdRestore.Enabled = False
    
End Sub

Private Sub cmdSave_Click()
    Dim lngLoop As Long
    Dim lngCount As Long
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    '1.先检查合法性(重复)
    For lngLoop = 1 To vsf.Rows - 1
        If Not (lngLoop = vsf.Rows - 1 And Trim(vsf.TextMatrix(lngLoop, 1)) = "") Then
            If Trim(vsf.TextMatrix(lngLoop, 1)) = "" Then
                MsgBox "必须指定多音字！", vbInformation, gstrSysName
                vsf.Row = lngLoop
                vsf.Col = 1
                vsf.SetFocus
                vsf.ShowCell vsf.Row, vsf.Col
                
                Exit Sub
            End If
            
            If Trim(vsf.TextMatrix(lngLoop, 2)) = "" Then
                MsgBox "必须多音字的读音！", vbInformation, gstrSysName
                vsf.Row = lngLoop
                vsf.Col = 2
                vsf.SetFocus
                vsf.ShowCell vsf.Row, vsf.Col
                
                Exit Sub
            End If
            
            For lngCount = 1 To lngLoop - 1
                If vsf.TextMatrix(lngLoop, 1) = vsf.TextMatrix(lngCount, 1) Then
                    MsgBox "不能重复设置多音字的读音！", vbInformation, gstrSysName
                    vsf.Row = lngLoop
                    vsf.Col = 1
                    vsf.SetFocus
                    vsf.ShowCell vsf.Row, vsf.Col
                    Exit Sub
                End If
            Next
            
        End If
    Next
    
    '2.保存
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    
    gstrSQL = "ZL_zlWordPolyphone_DELETE(" & mlngSys & ",'" & lvw.SelectedItem.Text & "')"
    
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    For lngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, 1)) <> "" Then
            
            gstrSQL = "ZL_zlWordPolyphone_INSERT(" & mlngSys & ",'" & lvw.SelectedItem.Text & "','" & vsf.TextMatrix(lngLoop, 1) & "','" & vsf.TextMatrix(lngLoop, 2) & "')"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            
        End If
    Next
    gcnOracle.CommitTrans
    
    cmdSave.Enabled = False
    cmdRestore.Enabled = False
    
    lvw.Enabled = True
    
    mblnOK = True
    
    Exit Sub
    
errHand:
    gcnOracle.RollbackTrans
    MsgBox "保存多音字配置失败！" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    mstrKey = ""
    
    With vsf
        .FixedCols = 1
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "多音字"
        .TextMatrix(0, 2) = "读音"
        
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 1500
        
        .Editable = flexEDKbdMouse
    End With
    
    DoEvents
        
    gstrSQL = "select 字词 AS 名称 from " & _
                                "(select 字词, substr(输入码, 1, 1) as item " & _
                                "from zlwordbasic " & _
                                "where 输入法 = 1 and length(字词) = 1 group by 字词, substr(输入码, 1, 1) " & _
                                ") B " & _
                        "GROUP BY 字词 " & _
                        "Having Count(Item) > 1"
                        
    If mrsPolyphone.State = adStateOpen Then mrsPolyphone.Close
    mrsPolyphone.Open gstrSQL, gcnOracle
    
    lvw.ListItems.Clear
    
    gstrSQL = "SELECT * " & _
             "FROM (SELECT A.TABLE_NAME " & _
                     "FROM all_col_comments A, All_Tab_Columns B, all_objects E " & _
                    "WHERE A.OWNER = '" & UCase(gstrUserName) & "' AND B.OWNER = '" & UCase(gstrUserName) & "'  AND E.OWNER = '" & UCase(gstrUserName) & "' AND " & _
                          "INSTR(',名称,中文名,', ',' || A.COLUMN_NAME || ',') > 0 AND " & _
                          "A.TABLE_NAME = B.TABLE_NAME AND B.DATA_TYPE = 'VARCHAR2' AND " & _
                          "E.OBJECT_NAME = A.TABLE_NAME AND E.object_type = 'TABLE' And Instr(E.OBJECT_NAME,'BIN$')<=0 " & _
                   "Union " & _
                     "SELECT table_name from user_tables where table_name = '人员表' " & _
                   ") " & _
            "GROUP BY TABLE_NAME"
            
    If rsTmp.State = adStateOpen Then rsTmp.Close
    rsTmp.Open gstrSQL, gcnOracle
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            lvw.ListItems.Add , "K" & rsTmp("TABLE_NAME").Value, rsTmp("TABLE_NAME").Value
            rsTmp.MoveNext
        Loop
    End If
    
    Call AdjustRowFlag(vsf, 1)
    vsf.ComboList = " "
    If mrsPolyphone.BOF = False Then
        Do While Not mrsPolyphone.EOF
            vsf.ComboList = vsf.ComboList & "|" & mrsPolyphone("名称").Value
            mrsPolyphone.MoveNext
        Loop
    End If
        
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub Form_Load()
    mblnStartUp = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled Then
        Cancel = (MsgBox("修改后的基本字词必须保存后才生效，是否放弃保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As New ADODB.Recordset
    
    If mstrKey <> Item.Key Then
        mstrKey = Item.Key
        
        vsf.Rows = 2
        vsf.TextMatrix(1, 1) = ""
        vsf.TextMatrix(1, 2) = ""
        
        gstrSQL = "SELECT * FROM zlWordPolyphone WHERE 系统=" & mlngSys & " AND 表名='" & Item.Text & "'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                If vsf.TextMatrix(vsf.Rows - 1, 1) <> "" Then
                    vsf.AddItem ""
                End If
                
                vsf.TextMatrix(vsf.Rows - 1, 1) = rs("字词").Value
                vsf.TextMatrix(vsf.Rows - 1, 2) = rs("读音").Value
                
                rs.MoveNext
            Loop
        End If
        
    End If
    
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Trim(vsf.Cell(flexcpData, Row, Col)) <> Trim(vsf.TextMatrix(Row, Col)) Then
        cmdRestore.Enabled = True
        cmdSave.Enabled = True
        lvw.Enabled = False
        If Col = 1 Then vsf.TextMatrix(Row, 2) = ""
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
            
    vsf.FocusRect = flexFocusSolid
                    
    If OldRow <> NewRow Then Call AdjustRowFlag(vsf, NewRow)
    
    If OldCol <> NewCol Then
        vsf.ComboList = " "
            
        '下拉,传入记录集
        If NewCol = 1 Then
            
            If mrsPolyphone.RecordCount > 0 Then mrsPolyphone.MoveFirst
            If mrsPolyphone.BOF = False Then
                Do While Not mrsPolyphone.EOF
                    vsf.ComboList = vsf.ComboList & "|" & mrsPolyphone("名称").Value
                    mrsPolyphone.MoveNext
                Loop
            End If
            
        ElseIf NewCol = 2 Then
            gstrSQL = "select 输入码 AS 名称 from zlwordbasic where 输入法 = 1 and 字词='" & vsf.TextMatrix(NewRow, 1) & "'"
            If rsTmp.State = adStateOpen Then rsTmp.Close
            rsTmp.Open gstrSQL, gcnOracle
        
            If rsTmp.BOF = False Then
                Do While Not rsTmp.EOF
                    vsf.ComboList = vsf.ComboList & "|" & rsTmp("名称").Value
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        
        If Shift = 0 And vsf.Editable <> flexEDNone Then
            '删除整行及内容
            
            If vsf.Rows > 1 Then
                If vsf.Rows = 2 And vsf.Row = 1 Then
                    vsf.TextMatrix(1, 1) = ""
                    vsf.TextMatrix(1, 2) = ""
                Else
                    vsf.RemoveItem vsf.Row
                End If
            End If
            
        End If
        
        If Shift = 2 And vsf.Editable <> flexEDNone Then
            '删除当前单元格的内容
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
        End If
        
    End Select
    
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call GoNextCell
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call GoNextCell
    End If
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsf.EditSelStart = 0
    vsf.EditSelLength = Len(vsf.EditText)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
End Sub
