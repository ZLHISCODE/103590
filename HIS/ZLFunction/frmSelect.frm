VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView lvw 
      Height          =   2850
      Left            =   2535
      TabIndex        =   2
      Top             =   555
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   5027
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6120
      Begin VB.Image Image1 
         Height          =   240
         Left            =   165
         Picture         =   "frmSelect.frx":014A
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2355
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3660
      Width           =   6120
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4395
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3150
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   165
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelect.frx":06D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   2760
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   4868
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   2625
      Left            =   2190
      TabIndex        =   8
      Top             =   615
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   4630
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmSelect.frx":082E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4905
      TabIndex        =   9
      Top             =   315
      Width           =   435
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入：SQL及字段描述
Public mstrSQLList As String
Public mstrSQLTree As String
Public mstrFLDList As String
Public mstrFLDTree As String
Public mstrParName As String '参数名称
Public mbytDataType As Byte      '参数数据类型
Public mstrMatch As String '输入匹配的内容
Public mlngSeekHwnd As Long '用于定位窗体位置的控件

'出：未作格式处理的数据原始值
Public mstrOutBand As String '选择的绑定值,对应&B
Public mstrOutDisp As String '选择的显示值,对应&D

Private intPreNode As Long
Private blnItem As Boolean
Private blnSetFlex As Boolean, blnSetLvw As Boolean
Private rsList As ADODB.Recordset
Private strList As String
Private blnSave As Boolean
Private rParent As RECT

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strDisp As String, strBand As String
    
    strDisp = GetScript(mstrFLDList, "&D") '显示的字段名
    strBand = GetScript(mstrFLDList, "&B") '绑定的字段名
    
    If strDisp = "" Or strBand = "" Then
        MsgBox "选择器中没有定义条件的绑定及显示字段项目！", vbInformation, App.Title
        Exit Sub
    End If
    
    If strList = "lvw" Then
        If lvw.SelectedItem Is Nothing Then
            MsgBox "没有选择任何内容！", vbInformation, App.Title
            If tvw_s.Visible Then tvw_s.SetFocus
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(0) = "" Then
            MsgBox "在条件""" & mstrParName & """中显示的项目""" & strDisp & """为空，请选择其它内容！", vbInformation, App.Title
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(1) = "" Then
            MsgBox "与条件值""" & mstrParName & """绑定的项目""" & strBand & """为空，请选择其它内容！", vbInformation, App.Title
            Exit Sub
        End If
        '类型检查
        Select Case mbytDataType
            Case 1
                If Not IsNumeric(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "项目""" & strBand & """的内容非数字型,不能被选择！", vbInformation, App.Title
                    Exit Sub
                End If
            Case 2
                If Not IsDate(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "项目""" & strBand & """的内容非日期型,不能被选择！", vbInformation, App.Title
                    Exit Sub
                End If
        End Select
        
        mstrOutDisp = Split(lvw.SelectedItem.Tag, "|")(0)
        mstrOutBand = Split(lvw.SelectedItem.Tag, "|")(1)
    Else
        '如果FlexGrid可见,则rsList一定有数据
        If msh.TextMatrix(msh.Row, GetColNum(strDisp)) = "" Then
            MsgBox "在条件""" & mstrParName & """中显示的项目""" & strDisp & """为空，请选择其它内容！", vbInformation, App.Title
            Exit Sub
        End If
        If msh.TextMatrix(msh.Row, GetColNum(strBand)) = "" Then
            MsgBox "与条件值""" & mstrParName & """绑定的项目""" & strBand & """为空，请选择其它内容！", vbInformation, App.Title
            Exit Sub
        End If
        '类型检查
        Select Case mbytDataType
            Case 1
                If Not IsNumeric(msh.TextMatrix(msh.Row, GetColNum(strBand))) Then
                    MsgBox "项目""" & strBand & """的内容非数字型,不能被选择！", vbInformation, App.Title
                    Exit Sub
                End If
            Case 2
                If Not IsDate(msh.TextMatrix(msh.Row, GetColNum(strBand))) Then
                    MsgBox "项目""" & strBand & """的内容非日期型,不能被选择！", vbInformation, App.Title
                    Exit Sub
                End If
        End Select
        
        mstrOutDisp = msh.TextMatrix(msh.Row, GetColNum(strDisp))
        mstrOutBand = msh.TextMatrix(msh.Row, GetColNum(strBand))
    End If
    gblnOK = True
    
    On Error Resume Next
    Hide
End Sub

Private Sub Form_Activate()
    If tvw_s.Visible Then
        If Not tvw_s.SelectedItem Is Nothing Then
            If tvw_s.SelectedItem.Key = "ALL" Then
                If lvw.Visible Then
                    lvw.SetFocus
                ElseIf msh.Visible Then
                    msh.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub msh_DblClick()
    If msh.MouseRow = 0 Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub msh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim lngW As Long, i As Integer
    
    If Not InDesign Then
        glngSelProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SelMessage)
    End If
    
    gblnOK = False
    blnSave = True
    blnSetFlex = False '是否已经对表格恢复宽度
    blnSetLvw = False
    intPreNode = 0
    
    mstrOutBand = ""
    mstrOutDisp = ""
    
    msh.Tag = mstrParName
    lvw.Tag = mstrParName
    
    Me.Caption = mstrParName & "选择器"
    
    mstrSQLList = Replace(mstrSQLList, "[*]", mstrMatch)
    mstrSQLTree = Replace(mstrSQLTree, "[*]", mstrMatch)
    
    If mstrSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
        If Not FillList Then blnSave = False: Unload Me: Exit Sub
    Else
        tvw_s.Visible = True
        If Not FillTree Then blnSave = False: Unload Me: Exit Sub
        If tvw_s.Nodes.Count > 0 Then
            tvw_s.Nodes(1).Selected = True
            If Not tvw_s.Nodes(1).Child Is Nothing And mstrMatch = "" Then
                tvw_s.Nodes(1).Child.Selected = True
            End If
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    '输入匹配自动返回
    If mstrMatch <> "" Then
        If rsList.RecordCount = 1 Then
            blnSave = False
            Call cmdOK_Click
            Unload Me: Exit Sub
        ElseIf rsList.RecordCount = 0 Then
            MsgBox "没有找到相匹配的项目,请重新输入！", vbInformation, App.Title
            blnSave = False
            Call cmdCancel_Click: Exit Sub
        End If
    End If
    
    Call Form_Resize
    
    '窗体及列表缺省宽度
    Select Case strList
        Case "lvw"
            If lvw.ColumnHeaders.Count = 1 Then
                lvw.ColumnHeaders(1).Width = 2500
                Me.Width = 3000 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
            Else
                For i = 1 To lvw.ColumnHeaders.Count
                    lngW = lngW + lvw.ColumnHeaders(i).Width
                Next
                Me.Width = lngW + 500 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
                If Me.Width < 3000 Then Me.Width = 3000
            End If
        Case "msh"
            If msh.Cols = 1 Then
                msh.ColWidth(0) = 2500
                Me.Width = 3000 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
            Else
                For i = 0 To msh.Cols - 1
                    lngW = lngW + msh.ColWidth(i)
                Next
                Me.Width = lngW + 500 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
                If Me.Width < 3000 Then Me.Width = 3000
            End If
    End Select
    If mstrSQLTree <> "" Then
        If Me.Width < (tvw_s.Width + pic.Width) * 2.2 Then Me.Width = (tvw_s.Width + pic.Width) * 2.2
    End If
    
    RestoreWinState Me, App.ProductName, mstrParName
    
    If mstrSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
    Else
        tvw_s.Visible = True
    End If
    
    '定位
    If mlngSeekHwnd <> 0 Then
        Call Form_Resize
        GetWindowRect mlngSeekHwnd, rParent
        If rParent.Top >= Me.Height / 15 Then
            Me.Top = rParent.Bottom * 15 - Me.Height + 30
        Else
            Me.Top = (rParent.Bottom - rParent.Top) * 15 + 30
        End If
        If rParent.Left >= Me.Width / 15 Then
            Me.Left = rParent.Right * 15 - Me.Width + 30
        Else
            Me.Left = (rParent.Right - rParent.Left) * 15 + 30
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim lngTVW As Long
    lngTVW = IIf(tvw_s.Visible, tvw_s.Width + pic.Width, 0)
    
    tvw_s.Left = Me.ScaleLeft
    tvw_s.Top = picInfo.Top + picInfo.Height + 15
    tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - 15
    
    pic.Left = tvw_s.Left + tvw_s.Width
    pic.Top = tvw_s.Top
    pic.Height = tvw_s.Height
    
    lvw.Left = Me.ScaleLeft + lngTVW
    lvw.Top = tvw_s.Top
    lvw.Height = tvw_s.Height
    lvw.Width = Me.ScaleWidth - lngTVW
    
    msh.Left = lvw.Left
    msh.Top = lvw.Top
    msh.Width = lvw.Width
    msh.Height = lvw.Height
    
    lbl.Left = lvw.Left
    lbl.Top = lvw.Top
    lbl.Width = lvw.Width
    lbl.Height = lvw.Height
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrMatch = ""
    mlngSeekHwnd = 0
    If blnSave Then SaveWinState Me, App.ProductName, mstrParName
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngSelProc)
End Sub

Private Sub lvw_DblClick()
    If blnItem Then Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnItem = True
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnItem = False
End Sub

Private Sub msh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If msh.MouseRow = 0 Then
        msh.MousePointer = 99
    Else
        msh.MousePointer = 0
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        msh.Left = msh.Left + X
        msh.Width = msh.Width - X
        
        lbl.Left = lbl.Left + X
        lbl.Width = lbl.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Index = intPreNode Then Exit Sub
    intPreNode = Node.Index
    DoEvents
    Call FillList(Node.Tag)
End Sub

Private Function FillTree() As Boolean
'功能：根据定义数据源及字段属性，将分类数据显示在TreeView中
'返回：操作是否成功(用户非正常定义)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objNode As Node
    Dim strSel As String, strRela As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSel = GetScript(mstrFLDTree, "&S")
    strRela = GetScript(mstrFLDTree, "&R")
    
    If strSel = "" Or strRela = "" Then
        MsgBox "未发现用于选择或与明细列表相关联的字段项目！", vbInformation, App.Title
        Exit Function
    End If
    strSQL = RemoveNote(mstrSQLTree)
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "FillTree")
    
    tvw_s.Nodes.Clear
        
    If Not rsTmp.EOF Then
        If InStr("|" & UCase(mstrFLDTree), "|ID,") > 0 And InStr("|" & UCase(mstrFLDTree), "|上级ID,") > 0 Then
            '采用树形列表显示
            Set objNode = tvw_s.Nodes.Add(, , "ALL", "所有项目", 1)
            objNode.Tag = "ALL"
            objNode.Expanded = True
            
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!上级ID) Then
                    Set objNode = tvw_s.Nodes.Add("ALL", 4, "_" & rsTmp!ID, IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                Else
                    Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                End If
                objNode.Tag = IIf(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
                rsTmp.MoveNext
            Next
        Else
            '采用一般列表显示
            For i = 1 To rsTmp.RecordCount
                Set objNode = tvw_s.Nodes.Add(, , , IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                objNode.Tag = IIf(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
                rsTmp.MoveNext
            Next
        End If
    End If
    FillTree = True
    Exit Function
errH:
    If Err.Number = 35601 Then
        MsgBox "不能正常处理树形列表，条件选择器不能使用！", vbExclamation, App.Title
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Private Function GetRelaSQL(ByVal strSQL As String, ByVal strFld As String, ByVal strKey As String) As String
'功能：处理关联的SQL
    Dim i As Integer, strRela As String
    
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), "&R") > 0 Then
            strRela = Split(Split(strFld, "|")(i), ",")(0)
            If strKey = "" Then
                GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & " is NULL"
            Else
                Select Case Split(Split(strFld, "|")(i), ",")(1)
                    Case adNumeric, adVarNumeric
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=" & strKey
                    Case adChar, adVarChar
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "='" & strKey & "'"
                    Case adDBTimeStamp
                        If Format(strKey, "hh:mm:ss") = "00:00:00" Then
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & ">=To_Date('" & Format(strKey, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & strRela & "<=To_Date('" & Format(strKey, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=To_Date('" & Format(strKey, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                End Select
            End If
            Exit Function
        End If
    Next
End Function

Private Function GetScript(strFld As String, strType As String) As String
'功能：根据指定的字段描述返回字段名
'参数：strType="&S &D &B &R"
'说明：适用于唯一性描述字段(如绑定字段)
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), strType) > 0 Then
            GetScript = Split(Split(strFld, "|")(i), ",")(0)
            Exit Function
        End If
    Next
End Function

Private Function HaveScript(strFld As String, StrName As String, strType As String) As Boolean
'功能：判断在字段描述中，指定的字段是否具有指定的描述属性
'参数：strName=字段名,strFld=字段描述串,strType="&S &D &B &R"
'返回：False=未发现字段或字段不具有指定描述
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = StrName Then
            If InStr(Split(Split(strFld, "|")(i), ",")(2), strType) > 0 Then
                HaveScript = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function FillList(Optional strKey As String, Optional blnSort As Boolean) As Boolean
'功能：根据当前选择的分类或在无分类时处理对应的明细列表
'参数：strKey=分类列表中的当前关联值
'说明：根据数据量的多少，确定用ListView还是DataGrid
    Dim strSQL As String, i As Long, j As Integer
    Dim objItem As ListItem, strValue As String
    Dim strDisp As String, strBand As String
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    
    lvw.Visible = False
    msh.Visible = False
    strList = ""
    msh.Clear
        
    '可能为只处理排序
    If Not blnSort Then
        If mstrSQLTree = "" Then
            strSQL = mstrSQLList
        Else
            '动态将明细数据处理为只读取关联的分类部分(处理 Order by 子句)
            If strKey = "ALL" Then
                strSQL = mstrSQLList
            Else
                strSQL = GetRelaSQL(RemoveOrderBy(mstrSQLList), mstrFLDList, strKey)
            End If
            
            If strSQL = "" Then
                MsgBox "该类数据读取失败！", vbInformation, App.Title
                Exit Function
            End If
        End If
        
        Set rsList = New ADODB.Recordset
        rsList.CursorLocation = adUseClient
        Screen.MousePointer = 11
        Me.Refresh
        strSQL = RemoveNote(strSQL)
        Set rsList = zldatabase.OpenSQLRecord(strSQL, "FillList")
        
    End If
    
    If Not rsList.EOF Then
        If rsList.RecordCount <= 500 Then
            If lvw.ColumnHeaders.Count = 0 Then Call AddListCols
            
            strDisp = GetScript(mstrFLDList, "&D") '显示值项目
            strBand = GetScript(mstrFLDList, "&B") '绑定值项目
            
            For i = 1 To rsList.RecordCount
                strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(1).Text))
                If lvw.ColumnHeaders(1).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(1).Tag)
                Set objItem = lvw.ListItems.Add(, , strValue, , 1)
                For j = 2 To lvw.ColumnHeaders.Count
                    strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(j).Text))
                    If lvw.ColumnHeaders(j).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(j).Tag)
                    objItem.SubItems(j - 1) = strValue
                Next
                
                '将显示值及绑定值保存在TAG中,因为不一定这些字段会为选择字段
                '格式为"显示值|绑定值"
                If strDisp <> "" Then
                    objItem.Tag = IIf(IsNull(rsList.Fields(strDisp).Value), "", rsList.Fields(strDisp).Value)
                End If
                objItem.Tag = objItem.Tag & "|"
                If strBand <> "" Then
                    objItem.Tag = objItem.Tag & IIf(IsNull(rsList.Fields(strBand).Value), "", rsList.Fields(strBand).Value)
                End If
                                
                rsList.MoveNext
            Next
            
            '自动调整列宽
            Call AutoSizeCol(lvw)
            
            If Not Visible Or Not blnSetLvw Then
                Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrParName)
                blnSetLvw = True
            End If
            lvw.Visible = True
            strList = "lvw"
        Else
            msh.Redraw = False
            msh.Clear
            Set msh.DataSource = rsList
           
            For i = 0 To msh.Cols - 1
                '删除不显示的列(&S)
                If Not HaveScript(mstrFLDList, msh.TextMatrix(0, i), "&S") Then
                    msh.ColWidth(i) = 0
                Else
                    '设置列对齐
                    Select Case rsList.Fields(msh.TextMatrix(0, i)).Type
                        Case adNumeric, adVarNumeric
                            If rsList.Fields(msh.TextMatrix(0, i)).NumericScale > 0 Then
                                j = rsList.Fields(msh.TextMatrix(0, i)).NumericScale
                                msh.ColAlignment(i) = 7
                            Else
                                If rsList.Fields(msh.TextMatrix(0, i)).Precision < 3 Then
                                    msh.ColAlignment(i) = 4
                                Else
                                    msh.ColAlignment(i) = 1
                                End If
                            End If
                        Case adDBTimeStamp
                            msh.ColAlignment(i) = 4
                        Case Else
                            msh.ColAlignment(i) = 1
                    End Select
                    If msh.TextMatrix(0, i) Like "*单位*" Then msh.ColAlignment(i) = 4
                    If msh.TextMatrix(0, i) Like "*否*" Then msh.ColAlignment(i) = 4
                End If
            Next
            '设置列宽度
            Call SetColWidth(msh, Me)
            
            msh.Col = 0: msh.ColSel = msh.Cols - 1
            If Not Visible Or Not blnSetFlex Then
                blnSetFlex = True
                RestoreFlexState msh, App.ProductName & "\" & Me.Name & mstrParName
            End If
            msh.Redraw = True
            msh.Visible = True
            strList = "msh"
        End If
        lblInfo.Caption = "共 " & rsList.RecordCount & " 个明细项目."
    Else
        '没有数据时，显示空的ListView(带列头)
        If lvw.ColumnHeaders.Count = 0 Then Call AddListCols
        lvw.Visible = True
        strList = "lvw"
        lblInfo.Caption = "没有明细项目."
    End If
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddListCols()
'功能：根据mstrFLDList字段描述值,为ListView增加列头
    Dim i As Integer, j As Integer, strFld As String
    Dim objCol As ColumnHeader
    
    For i = 0 To UBound(Split(mstrFLDList, "|"))
        strFld = Split(mstrFLDList, "|")(i)
        If strFld Like "*&S*" Then
            Set objCol = lvw.ColumnHeaders.Add(, "_" & Split(strFld, ",")(0), Split(strFld, ",")(0))
            
            objCol.Width = Me.TextWidth(Split(strFld, ",")(0) & "字")
            
            '根据字段名及类型设置对齐(列1只能左对齐)
            Select Case Split(strFld, ",")(1)
                Case adNumeric, adVarNumeric
                    If rsList.Fields(objCol.Text).NumericScale > 0 Then
                        j = rsList.Fields(objCol.Text).NumericScale
                        objCol.Tag = "0." & String(IIf(j > 2, 2, j), "0; ;")
                        If objCol.Index <> 1 Then objCol.Alignment = lvwColumnRight
                    ElseIf objCol.Index <> 1 Then
                        If rsList.Fields(objCol.Text).Precision < 3 Then
                            objCol.Alignment = lvwColumnCenter
                        Else
                            objCol.Alignment = lvwColumnLeft
                        End If
                    End If
                    If objCol.Text Like "*价" Then objCol.Tag = "0.000"
                    If objCol.Text Like "*额" Then objCol.Tag = "0.00"
                Case adDBTimeStamp
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
                Case Else
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
            End Select
            If objCol.Text Like "*单位*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
            If objCol.Text Like "*否*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
        End If
    Next
End Sub

Private Function GetValue(objFld As Field) As String
'功能:根据字段内容取合适的显示值
    Dim strValue As String
    Select Case objFld.Type
        Case adChar, adVarChar, adLongVarChar
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
        Case adNumeric, adVarNumeric
            strValue = IIf(IsNull(objFld.Value), 0, objFld.Value)
        Case adDBTimeStamp
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
            If Format(strValue, "HH:mm:ss") = "00:00:00" Then
                strValue = Format(strValue, "yyyy-MM-dd")
            Else
                strValue = Format(strValue, "yyyy-MM-dd HH:mm:ss")
            End If
        Case Else
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
    End Select
    GetValue = strValue
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'功能：按列排序
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub msh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    If msh.MouseRow <> 0 Then Exit Sub
    
    lngCol = msh.MouseCol
    
    If Button = 1 And msh.MousePointer = 99 Then
        If msh.TextMatrix(0, lngCol) = "" Then Exit Sub
        If rsList Is Nothing Then Exit Sub
        If rsList.State = 0 Then Exit Sub
        
        Set msh.DataSource = Nothing

        rsList.Sort = msh.TextMatrix(0, lngCol) & IIf(msh.ColData(lngCol) = 0, "", " DESC")
        msh.ColData(lngCol) = (msh.ColData(lngCol) + 1) Mod 2
        
        Call FillList(, True)
    End If
End Sub
