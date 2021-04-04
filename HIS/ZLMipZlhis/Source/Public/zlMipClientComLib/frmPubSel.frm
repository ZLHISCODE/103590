VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   6840
      TabIndex        =   9
      Top             =   0
      Width           =   6840
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   120
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6840
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "全选(&A)"
         Height          =   360
         Left            =   300
         TabIndex        =   4
         ToolTipText     =   "Ctrl+A"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "全清(&R)"
         Height          =   360
         Left            =   1290
         TabIndex        =   5
         ToolTipText     =   "Ctrl+R"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5265
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4170
         TabIndex        =   2
         Top             =   105
         Width           =   1100
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   1
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
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
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4725
      Top             =   1425
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
            Picture         =   "frmPubSel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2400
      ScaleHeight     =   1110
      ScaleWidth      =   2220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String
Private mblnMulti As Boolean

'入口参数
Private mstrTitle As String
Private mstrNote As String
Private mbytStyle As Byte
Private mstrSeek As String
Private mbln末级 As Boolean
Private mblnShowSub As Boolean
Private mblnShowRoot As Boolean
Private mblnMultiOne As Boolean

Private mstrSaveTag As String '注册表区分键
Private mstrSQL As String
Private marrInput() As Variant

Private mblnSearch As Boolean '是否通过输入行号检索

Private mblnNoneWin As Boolean
Private mlngX As Long, mlngY As Long, mlngTxtH As Long
'出口参数
Private mrsSel As ADODB.Recordset

'程序变量
Private mblnOK As Boolean

Public Function ShowSelect(frmParent As Object, ByVal strSQL As String, bytStyle As Byte, _
    ByVal strTitle As String, bln末级 As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
    Optional arrInput As Variant) As ADODB.Recordset
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
'     blnMulti=是否允许多选
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    mstrSQL = strSQL
    If TypeName(arrInput) <> "Error" Then
        marrInput = arrInput
    Else
        marrInput = Array()
    End If
    
    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mbln末级 = bln末级
    mstrSeek = strSeek
    mblnShowSub = blnShowSub
    mblnShowRoot = blnShowRoot
    mblnMultiOne = blnMultiOne
    mblnNoneWin = blnNoneWin
    mlngX = X: mlngY = Y: mlngTxtH = txtH
    mblnSearch = blnSearch
    mblnMulti = blnMulti
    
    If Not frmParent Is Nothing Then
        mstrSaveTag = frmParent.Name & "_" & strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    End If
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        Cancel = False
        Set ShowSelect = mrsSel
    Else
        Cancel = True
        Set ShowSelect = Nothing
    End If
End Function

Private Sub cmdCancel_Click()
    Set mrsSel = Nothing
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    For i = 1 To lvw.ListItems.count
        lvw.ListItems(i).Checked = False
    Next
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub cmdOK_Click()
    If mrsSel.RecordCount = 0 Then Exit Sub
    
    If mbln末级 And mbytStyle = 1 Then
        If mrsSel!末级 <> 1 Then Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To lvw.ListItems.count
        lvw.ListItems(i).Checked = True
    Next
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If lvw.Visible Then
        lvw.SetFocus
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOK.Enabled Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdSelALL_Click
    ElseIf (KeyCode = vbKeyR Or KeyCode = vbKeyC) And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdClear_Click
    End If
End Sub

Private Function OpenSQLRecord(ByVal strSQL As String, strLog As String) As ADODB.Recordset
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim intMax As Integer, i As Integer
    Dim strSeq As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = marrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.count > 0
        cmdData.Parameters.Delete 0
    Loop

    '创建新的参数
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = marrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = gobjComLib.zlCommFun.ActualLen(varValue)
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call gobjComLib.SQLTest(App.ProductName, mstrTitle & "选择", strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call gobjComLib.SQLTest
End Function

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    Dim lngColW As Long, i As Integer
    Dim blnSame As Boolean, strItemID As String
    Dim strCode As String, strName As String
    Dim objNode As Node, strLog As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    mblnOK = False
    mstrKey = ""
        
    '打开SQL语句
    If UBound(marrInput) >= 0 Then
        Set mrsSel = OpenSQLRecord(mstrSQL, strLog)
    Else
        Set mrsSel = New ADODB.Recordset
        mrsSel.CursorLocation = adUseClient
        Call gobjComLib.SQLTest(App.ProductName, mstrTitle & "选择", mstrSQL)
        mrsSel.Open mstrSQL, gcnOracle, adOpenKeyset
        Call gobjComLib.SQLTest
    End If
    
    '没有数据则返回
    If mrsSel.EOF Then
        Screen.MousePointer = 0
        Set mrsSel = Nothing
        mblnOK = True: Unload Me: Exit Sub
    End If
     
    '输入匹配时自动返回的情况
    If mstrSQL Like "*%*" Or strLog Like "*%*" Then
        If mrsSel.RecordCount = 1 Then '只有一行数据
            Screen.MousePointer = 0
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mblnMultiOne And mbytStyle = 0 Then '多行相同数据
            blnSame = True
            For i = 1 To mrsSel.RecordCount
                If i = 1 Then
                    strItemID = mrsSel!id
                Else
                    If mrsSel!id <> strItemID Then blnSame = False: Exit For
                End If
                mrsSel.MoveNext
            Next
            mrsSel.MoveFirst
            If blnSame Then
                Screen.MousePointer = 0
                mblnOK = True: Unload Me: Exit Sub
            End If
        End If
    End If
    
    '确定名称字段
    strCode = "": strName = ""
    For i = 0 To mrsSel.Fields.count - 1
        If mrsSel.Fields(i).Name = "编码" Then strCode = "编码"
        If mrsSel.Fields(i).Name = "名称" Then
            strName = mrsSel.Fields(i).Name
        ElseIf mrsSel.Fields(i).Name = "姓名" And strName = "" Then
            strName = mrsSel.Fields(i).Name
        End If
    Next
    If strName = "" Then strName = "名称"
    
    '在填充数据之前设置CheckBox样式
    If mbytStyle <> 1 And mblnMulti Then
        lvw.Checkboxes = True
        cmdSelALL.Visible = True
        cmdClear.Visible = True
    End If
    
    '填充数据
    Select Case mbytStyle
        Case 0
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "病人ID") And mrsSel.Fields(i).Name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*价*" Or mrsSel.Fields(i).Name Like "*额*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.count).Alignment = lvwColumnRight
                    End If
                End If
            Next
            If mblnSearch Then lvw.ColumnHeaders.Add , "_行", "行", , 2
            Call gobjComLib.RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.count).Position = 1
            
            lvw.ListItems.Clear
            Call FillList
        Case 1
            '所有树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!上级ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级ID, 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!上级ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级ID, 4, "_" & mrsSel!id, mrsSel.Fields(strName).value, 1)
                        End If
                    End If
                    If objNode.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then
                        objNode.Selected = True
                        objNode.Parent.Expanded = True
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If tvw_s.SelectedItem.Index = 1 Then tvw_s.Nodes(1).Child.Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Case 2
            '非末级树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                mrsSel.Filter = "末级=0"
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!上级ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级ID, 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!上级ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级ID, 4, "_" & mrsSel!id, mrsSel.Fields(strName).value, 1)
                        End If
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "病人ID") And mrsSel.Fields(i).Name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*价*" Or mrsSel.Fields(i).Name Like "*额*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.count).Alignment = lvwColumnRight
                    End If
                End If
            Next
            If mblnSearch Then lvw.ColumnHeaders.Add , "_行", "行", , 2
            Call gobjComLib.RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.count).Position = 1
            
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End Select
    
    '设置控件可见性
    '---------------------------------------------------------------
    If mstrTitle <> "" Then
        Me.Caption = mstrTitle & "选择"
    End If
    If mstrNote <> "" Then
        lblInfo.Caption = mstrNote
    End If
    If mblnNoneWin Then
        pic.Width = 30
        pic.BackColor = vbBlack
        pic.ZOrder
        picInfo.Visible = False
        picCmd.Visible = False
        lvw.Appearance = ccFlat
        lvw.BorderStyle = ccFixedSingle
        tvw_s.Appearance = ccFlat
        tvw_s.BorderStyle = ccFixedSingle
    Else
        If mbytStyle <> 2 Then Me.Width = 4500 '缺省宽度
        Call gobjComLib.RestoreWinState(Me, App.ProductName, mstrSaveTag)
    End If
    Select Case mbytStyle
        Case 0
            lvw.Visible = True
            tvw_s.Visible = False
            pic.Visible = False
        Case 1
            lvw.Visible = False
            tvw_s.Visible = True
            pic.Visible = False
        Case 2
            lvw.Visible = True
            tvw_s.Visible = True
            pic.Visible = True
    End Select
    
    '调整窗体尺寸
    '---------------------------------------------------------------
    If mblnNoneWin Then
        Call gobjComLib.zlControl.FormSetCaption(Me, False, False)
        
        Me.Left = mlngX
        
        If mbytStyle = 1 Then
            Me.Width = 3100
        Else
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
            For i = 1 To lvw.ColumnHeaders.count
                lngColW = lngColW + lvw.ColumnHeaders(i).Width
            Next
            If mbytStyle = 2 Then lngColW = lngColW + tvw_s.Width
            
            If Me.Left + lngColW + lngScrW > Screen.Width Then
                lngColW = 0
                For i = 1 To lvw.ColumnHeaders.count
                    If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
                    lngColW = lngColW + lvw.ColumnHeaders(i).Width
                Next
                If Me.Left + lngColW + lngScrW > Screen.Width Then
                    Me.Width = Screen.Width - Me.Left
                Else
                    Me.Width = lngColW + lngScrW
                End If
            Else
                Me.Width = lngColW + lngScrW
            End If
        End If
        
        Me.Height = 3240
        lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
        If mlngY + mlngTxtH + Me.Height > lngScrH Then
            Me.Top = mlngY - Me.Height
        Else
            Me.Top = mlngY + mlngTxtH
        End If
        
        Call Form_Resize
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 1
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
            
            pic.Top = tvw_s.Top
            pic.Height = tvw_s.Height
            lvw.Top = tvw_s.Top
            lvw.Height = tvw_s.Height
            
            If mblnNoneWin Then
                pic.Left = tvw_s.Width - pic.Width / 2
                lvw.Left = tvw_s.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width
            Else
                pic.Left = tvw_s.Width
                lvw.Left = tvw_s.Width + pic.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
            End If
    End Select
    
    picBack.Left = lvw.Left
    picBack.Top = lvw.Top
    picBack.Width = lvw.Width
    picBack.Height = lvw.Height
    
    If Me.ScaleWidth - cmdCancel.Width * 1.3 - cmdOK.Width >= cmdClear.Left + cmdClear.Width * 1.3 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComLib.SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub lvw_DblClick()
    If cmdOK.Enabled And Not lvw.SelectedItem Is Nothing Then cmdOK_Click
End Sub

Private Function GetListFilter() As String
    Dim strFilter As String
    Dim i As Long
    
    '如果有勾选，以勾选的为准
    If mblnMulti Then
        For i = 1 To lvw.ListItems.count
            If lvw.ListItems(i).Checked Then
                If mrsSel.Fields("ID").type = adVarChar Then
                    If mbytStyle = 2 Then
                        strFilter = strFilter & " Or (ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And 末级=1)"
                    Else
                        strFilter = strFilter & " Or ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "'"
                    End If
                Else
                    If mbytStyle = 2 Then
                        strFilter = strFilter & " Or (ID=" & Split(lvw.ListItems(i).Key, "_")(1) & " And 末级=1)"
                    Else
                        strFilter = strFilter & " Or ID=" & Split(lvw.ListItems(i).Key, "_")(1)
                    End If
                End If
            End If
        Next
        strFilter = Mid(strFilter, 4)
    End If
    
    '如果没有勾选，则以当前项为准
    If strFilter = "" Then
        If mrsSel.Fields("ID").type = adVarChar Then
            strFilter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        Else
            strFilter = "ID=" & Split(lvw.SelectedItem.Key, "_")(1)
        End If
        If mbytStyle = 2 Then strFilter = strFilter & " And 末级=1"
    End If
    
    GetListFilter = strFilter
End Function

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mrsSel.Filter = GetListFilter
    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mrsSel.Filter = GetListFilter
    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If mblnSearch Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If lvw.ListItems.count >= CInt(strIdx) And CInt(strIdx) > 0 Then
                lvw.ListItems(CInt(strIdx)).Selected = True
                lvw.SelectedItem.EnsureVisible
                Call lvw_ItemClick(lvw.SelectedItem)
            End If
        End If
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        Me.Refresh
    End If
End Sub

Private Sub FillList()
'功能：装入ListView数据
    Dim i As Integer, j As Integer
    Dim objItem As ListItem
        
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To mrsSel.RecordCount
        For j = 0 To mrsSel.Fields.count - 1
            If (Not mrsSel.Fields(j).Name Like "*ID" Or mrsSel.Fields(j).Name = "病人ID") And mrsSel.Fields(j).Name <> "末级" Then
                If lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index = 1 Then
                    If mblnSearch Then '关键字加入行号
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsSel!id, IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsSel!id, IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index - 1) = IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value)
                End If
            End If
        Next
        If mblnSearch Then objItem.SubItems(lvw.ColumnHeaders("_行").Index - 1) = i
        mrsSel.MoveNext
    Next
    
    Call gobjComLib.zlControl.LvwSetColWidth(lvw)
    '20031013:限制最大宽度
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.count
            If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
        Next
    End If
    
    If Not lvw.SelectedItem Is Nothing Then
        cmdOK.Enabled = True
        
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        cmdOK.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub tvw_s_DblClick()
    If cmdOK.Enabled And mbytStyle = 1 Then cmdOK_Click
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim mstrKeys As String, i As Integer
    Dim strFilter As String
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If mrsSel.Fields("ID").type = adVarChar Then
                mrsSel.Filter = "ID='" & Mid(Node.Key, 2) & "'"
            Else
                mrsSel.Filter = "ID=" & Mid(Node.Key, 2)
            End If
            If mbln末级 Then
                cmdOK.Enabled = (mrsSel!末级 = 1)
            Else
                cmdOK.Enabled = True
            End If
        Else
            cmdOK.Enabled = False
        End If
    ElseIf mbytStyle = 2 Then
        lvw.ListItems.Clear
        If Node.Key = "Root" Then
            If mblnShowRoot Then
                mrsSel.Filter = "末级=1" '数据量大时很慢
            Else
                mrsSel.Filter = "末级=-1"
            End If
            If Visible Then lvw.SetFocus
        Else
            If mblnShowSub Then
                mstrKeys = GetSubTree(Node) '数据量大时很慢
            Else
                mstrKeys = Mid(Node.Key, 2)
            End If
            For i = 0 To UBound(Split(mstrKeys, ","))
                If mrsSel.Fields("上级ID").type = adVarChar Then
                    strFilter = strFilter & " Or (末级=1 And 上级ID='" & Split(mstrKeys, ",")(i) & "')"
                Else
                    strFilter = strFilter & " Or (末级=1 And 上级ID=" & Split(mstrKeys, ",")(i) & ")"
                End If
            Next
            strFilter = Mid(strFilter, 5)
            mrsSel.Filter = strFilter
            
'            If mrsSel.Fields("上级ID").Type = adVarChar Then
'                mrsSel.Filter = "末级=1 And 上级ID='" & Mid(Node.Key, 2) & "'"
'            Else
'                mrsSel.Filter = "末级=1 And 上级ID=" & Mid(Node.Key, 2)
'            End If
        End If
        If Not mrsSel.EOF Then Call FillList
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'功能：返回一个结点的子树结点的Key(含该结点)
    Dim mstrKeys As String
    Dim objTmp As Node
    
    mstrKeys = "," & Mid(objNode.Key, 2) & mstrKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            mstrKeys = "," & GetSubTree(objTmp) & mstrKeys
        Else
            mstrKeys = "," & Mid(objTmp.Key, 2) & mstrKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(mstrKeys, 2)
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If mblnSearch And ColumnHeader.Key = "_行" Then Exit Sub
    
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
        
    If mblnSearch Then
        For intIdx = 1 To lvw.ListItems.count
            lvw.ListItems(intIdx).SubItems(lvw.ColumnHeaders("_行").Index - 1) = intIdx
        Next
    End If
    intIdx = ColumnHeader.Index
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub
