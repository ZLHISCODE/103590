VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectDialog 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4710
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8040
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5340
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "所有下级"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6615
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4470
      Width           =   1050
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   4
         Top             =   45
         Width           =   225
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标题"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   360
      End
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   3180
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSelectDialog.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   1905
      Width           =   165
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   1395
      Left            =   75
      TabIndex        =   0
      Top             =   570
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   2461
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3225
      Top             =   585
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
            Picture         =   "frmSelectDialog.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1020
      Left            =   3165
      TabIndex        =   2
      Top             =   2580
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   1799
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   4395
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11430
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1860
      Left            =   2325
      MousePointer    =   9  'Size W E
      Top             =   465
      Width           =   30
   End
End
Attribute VB_Name = "frmSelectDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'********************************************************************************************************************************
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private mstrCaption As String
Private mblnStartUp As Boolean
Private mstrStatePath As String
Private mlngX As Long
Private mlngY As Long
Private mstrSvrKey As String
Private mlngSortColumn As Long
Private msglTxtH As Single
Private mstrSvrTag As String
Private mstrTitle As String
Private mblnLeftSelect As Boolean
Private mrsData As ADODB.Recordset
Private mblnOK As Boolean
Private mbytWinStyle As Byte
Private mInitSelectKey As String
Private mblnAllowResize As Boolean
Private mstrLvw As String
Private mblnMuliSelect As Boolean
Private mblnExpandNode As Boolean

Private gstrDBUser As String
Private gstrUserName As String

Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

'********************************************************************************************************************************
'

'--------------------------------------------------------------------------------------------------
'功能:获取任务栏的高度
'--------------------------------------------------------------------------------------------------
Private Function GetTrayHeight() As Long
    
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function

'--------------------------------------------------------------------------------------------------
'功能:对列表控件的列进行设置
'--------------------------------------------------------------------------------------------------
Private Function LvwSelectColumns(objSet As Object, ByVal strColumn As String) As Boolean

    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    
End Function

'--------------------------------------------------------------------------------------------------
'功能：保存当前选择器的状态
'--------------------------------------------------------------------------------------------------
Private Sub SaveFormState()
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath <> "" Then
    
        SaveSetting "ZLSOFT", mstrStatePath, "宽度", Me.Width
        SaveSetting "ZLSOFT", mstrStatePath, "高度", Me.Height
        SaveSetting "ZLSOFT", mstrStatePath, "分隔条", imgX.Left
        SaveSetting "ZLSOFT", mstrStatePath, "所有下级", chk.Value
        
'        For lngLoop = 1 To lvw.ColumnHeaders.Count
'            strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
'        Next
'        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
'
'        SaveSetting "ZLSOFT", mstrStatePath, "列宽", strTmp
        
        Call zl9ComLib.SaveListViewState(lvw, mstrCaption, "")
        
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'功能：保存当前选择器的状态
'--------------------------------------------------------------------------------------------------
Private Sub RestoreFormState()
    Dim lngLoop As Long
    Dim strTmp As String
    Dim blnDo As Boolean
        
    On Error Resume Next
        
    blnDo = (Val(zldatabase.GetPara("使用个性化风格")) = 1)
    
    If mstrStatePath <> "" And blnDo Then
        
        If mblnAllowResize Then
            Me.Width = GetSetting("ZLSOFT", mstrStatePath, "宽度", Me.Width)
            Me.Height = GetSetting("ZLSOFT", mstrStatePath, "高度", Me.Height)
        End If
        
        imgX.Left = GetSetting("ZLSOFT", mstrStatePath, "分隔条", imgX.Left)
        chk.Value = Val(GetSetting("ZLSOFT", mstrStatePath, "所有下级", 0))
        
        Call zl9ComLib.RestoreListViewState(lvw, mstrCaption, "")
        
'        For lngLoop = 1 To lvw.ColumnHeaders.Count
'            strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
'        Next
'
'        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
'
'        strTmp = GetSetting("ZLSOFT", mstrStatePath, "列宽", strTmp)
'
'        For lngLoop = 1 To lvw.ColumnHeaders.Count
'            lvw.ColumnHeaders(lngLoop).Width = Val(Split(strTmp, ";")(lngLoop - 1))
'        Next
'
    End If
    
    '检查是否超过屏幕高和宽度
    Dim lngTrayH As Long
    Dim lngH0 As Long
    Dim lngH1 As Long
    
    lngTrayH = GetTrayHeight
    
    If Me.Left + Me.Width > Screen.Width Then
        If (Screen.Width - Me.Width) >= 0 Then
            Me.Left = Screen.Width - Me.Width
        Else
            Me.Left = 0
            Me.Width = Screen.Width
        End If
    End If
    
    If Me.Top + Me.Height > (Screen.Height - lngTrayH) Then
        
        If (Me.Top - Me.Height - msglTxtH) >= 0 Then
            '放在输入框的上面
            Me.Top = Me.Top - Me.Height - msglTxtH
        Else
            
            '分别计算放置上面和放置下面的高度,取最大高度
            lngH0 = Me.Top - msglTxtH
            lngH1 = Screen.Height - lngTrayH - Me.Top
            
            If lngH0 > lngH1 Then
            
                '上面高
                Me.Top = 0
                Me.Height = lngH0
            Else
                Me.Height = Screen.Height - lngTrayH - Me.Top
            End If
        End If
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
'功能：装载数据
'参数：要装载数据的分类关键字
'--------------------------------------------------------------------------------------------------
Private Sub LoadData(ByVal strKey As String)
    Dim strKeys As String
    Dim objItem As ListItem
    Dim lngLoop As Long
    Dim strFilter As String
    
    On Error GoTo errHand
    
    mrsData.Filter = ""

    '显示所有下级
    If Not (tvw.SelectedItem Is Nothing) And chk.Value = 1 Then
        strKeys = GetDownKey(tvw, tvw.SelectedItem)
        strFilter = ""
        If strKeys <> "" Then
            For lngLoop = 0 To UBound(Split(strKeys, ","))
                strFilter = strFilter & " Or (末级=1 And 上级id=" & Split(strKeys, ",")(lngLoop) & ")"
            Next
            strFilter = Mid(strFilter, 5)
        End If
        
        mrsData.Filter = strFilter
    Else
        mrsData.Filter = "末级=1 AND 上级id=" & Mid(strKey, 2)
    End If
    
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    
    lvw.ListItems.Clear
    Do While Not mrsData.EOF
        Set objItem = lvw.ListItems.Add(, "K" & mrsData("ID").Value, mrsData(lvw.ColumnHeaders(1).Text).Value, 1, 1)
         
        For lngLoop = 2 To lvw.ColumnHeaders.Count
            objItem.SubItems(lngLoop - 1) = IIf(IsNull(mrsData(lvw.ColumnHeaders(lngLoop).Text)), "", mrsData(lvw.ColumnHeaders(lngLoop).Text))
        Next
        
        mrsData.MoveNext
    Loop
    
    Exit Sub
    
errHand:
        
End Sub

'--------------------------------------------------------------------------------------------------
'功能:
'--------------------------------------------------------------------------------------------------
Private Function GetDownKey(ByVal objTvw As TreeView, ByVal objNode As Node) As String
    Dim strTmp As String
    Dim objNodeChild As Node
                
    strTmp = ""
    Set objNodeChild = objNode.Child
    strTmp = strTmp & "," & Val(Mid(objNode.Key, 2))
    
    Do While Not (objNodeChild Is Nothing)
        strTmp = strTmp & "," & GetDownKey(objTvw, objNodeChild)
        Set objNodeChild = objNodeChild.Next
    Loop
    
    GetDownKey = IIf(strTmp <> "", Mid(strTmp, 2), "")
    
End Function

'--------------------------------------------------------------------------------------------------
'功能：
'--------------------------------------------------------------------------------------------------
Private Sub ReadTreeData()
    Dim objItem As Node
    Dim rs As New ADODB.Recordset
    
    mrsData.Filter = ""
    mrsData.Filter = "末级<>1"
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    
    Do While Not mrsData.EOF
        If IIf(IsNull(mrsData("上级id").Value), 0, mrsData("上级id").Value) <> 0 Then
            Set objItem = tvw.Nodes.Add("K" & mrsData("上级ID").Value, tvwChild, "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
            If mblnExpandNode Then objItem.Expanded = True
            
        Else
            Set objItem = tvw.Nodes.Add(, , "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
            If mblnExpandNode Then objItem.Expanded = True
            
        End If
        mrsData.MoveNext
    Loop
     
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).EnsureVisible
        tvw.Nodes(1).Expanded = True
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
'--------------------------------------------------------------------------------------------------
Private Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'--------------------------------------------------------------------------------------------------
'功能;
'--------------------------------------------------------------------------------------------------
Private Sub ReadListData()
    Dim lngLoop As Long
    Dim objItem As ListItem
    
    '装载数据
    With lvw
        Do While Not mrsData.EOF
            
            Set objItem = .ListItems.Add(, "K" & mrsData("ID").Value, mrsData(.ColumnHeaders(1).Text), 1, 1)
            For lngLoop = 1 To .ColumnHeaders.Count - 1
                objItem.SubItems(lngLoop) = NVL(mrsData(.ColumnHeaders(lngLoop + 1).Text))
            Next
            
            If mblnMuliSelect Then
                If Val(mrsData("选择")) = 1 Then objItem.Checked = True
            End If
            mrsData.MoveNext
        Loop
        .Refresh
    End With
End Sub

'--------------------------------------------------------------------------------------------------
'功能:显示查询选择器
'参数:
'返回:
'--------------------------------------------------------------------------------------------------
Public Function ShowSelect(ByVal frmMain As Form, _
                            ByVal WinStyle As Byte, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal LvwHead As String, _
                            ByVal Title As String, _
                            ByVal X As Single, _
                            ByVal Y As Single, _
                            Optional ByVal cx As Single = 7200, _
                            Optional ByVal cy As Single = 4500, _
                            Optional ByVal CtlHeight As Single = 300, _
                            Optional ByVal InitKey As String = "", _
                            Optional ByVal RegPath As String = "", _
                            Optional ByVal LeftSelect As Boolean = False, _
                            Optional ByVal AllowResize As Boolean = True, _
                            Optional ByVal MuliSelect As Boolean = False, _
                            Optional ByVal ExpandNode As Boolean = True, _
                            Optional ByVal CanSort As Boolean = True) As Boolean
   
    On Error GoTo errHand
    
    mblnStartUp = True
    
    If rsData.BOF Then Exit Function
    
    mstrCaption = frmMain.Caption
    
    Set mrsData = rsData
    mblnLeftSelect = LeftSelect
    mstrSvrKey = ""
    mblnOK = False
    mstrSvrTag = ""
    mlngSortColumn = 1
    msglTxtH = CtlHeight
    mstrTitle = Title
    mbytWinStyle = WinStyle              '1-TreeView;2-ListView;3-TreeView+ListView
    mInitSelectKey = InitKey
    mblnAllowResize = AllowResize
    mstrLvw = LvwHead
    mstrStatePath = RegPath
    mblnMuliSelect = MuliSelect
    mblnExpandNode = ExpandNode
    
    Me.Left = X
    Me.Top = Y
    Me.Width = cx
    Me.Height = cy
    lvw.Sorted = CanSort
    If InitData = False Then Exit Function
    
    mblnStartUp = False
    
    Me.Show 1, frmMain
    
    Set rsData = mrsData
    
    ShowSelect = mblnOK
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'功能:
'--------------------------------------------------------------------------------------------------
Private Function InitData() As Boolean
    
    Dim lngUpKey As Long
    Dim objNode As Node
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    Select Case mbytWinStyle
    Case 1
        
        mblnLeftSelect = True
        
        lvw.Visible = False
        chk.Visible = False
        imgX.Visible = False
        
    Case 2
    
        tvw.Visible = False
        chk.Visible = False
        imgX.Visible = False
        
    Case 4
    
        tvw.Visible = False
        chk.Visible = False
        imgX.Visible = False
        
        lvw.Checkboxes = True
    End Select
        
    If mblnAllowResize = False Then
        picDrag.Visible = False
        If stb.Panels.Count >= 3 Then stb.Panels.Remove 3
    End If
    
    If mbytWinStyle <> 3 And stb.Panels.Count >= 2 Then stb.Panels.Remove 2
    
    lblCaption.Caption = mstrTitle
    cmdOK.Visible = mblnMuliSelect
    
    lvw.ListItems.Clear
    
    LvwSelectColumns lvw, mstrLvw
    
    If mblnLeftSelect Then tvw.LineStyle = tvwTreeLines
    If mblnMuliSelect Then
        stb.Height = 480
    Else
        stb.Height = 315
    End If
    
    Call RestoreFormState
            
    Select Case mbytWinStyle
    Case 1
        
        tvw.Checkboxes = mblnMuliSelect
        Call ReadTreeData
        'tvw.SetFocus
        
    Case 2
        
        lvw.Checkboxes = mblnMuliSelect
        Call ReadListData
        
    Case 3
        
        lvw.Checkboxes = mblnMuliSelect
        Call ReadTreeData
        If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
        
        
    End Select
    
    '定位初始项InitSelectKey
    If mInitSelectKey <> "" Then
        
        On Error Resume Next
        
        If mbytWinStyle = 1 Or (mbytWinStyle = 3 And mblnLeftSelect) Then
            'tvw
            
            Set objNode = tvw.Nodes("K" & mInitSelectKey)
            
            If Not (objNode Is Nothing) Then
                objNode.Selected = True
                objNode.EnsureVisible
                
                Call tvw_NodeClick(objNode)
            End If
            
        Else
            'lvw
            If mbytWinStyle = 3 Then
                mrsData.Filter = ""
                mrsData.Filter = "ID=" & mInitSelectKey
                If mrsData.RecordCount > 0 Then
                    
                    lngUpKey = NVL(mrsData("上级id"), 0)
                    
                    Set objNode = tvw.Nodes("K" & lngUpKey)
                    
                    If Not (objNode Is Nothing) Then
                        objNode.Selected = True
                        objNode.EnsureVisible
                        
                        Call tvw_NodeClick(objNode)
                    End If
                    
                End If
                
                mrsData.Filter = ""
            End If
                        
            Set objItem = lvw.ListItems("K" & mInitSelectKey)
            
            If Not (objItem Is Nothing) Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
        End If
    End If
    
    InitData = True
    
    Exit Function
    
errHand:
'    Resume
End Function
'********************************************************************************************************************************
'

Private Sub chk_Click()
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrSvrKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
    
    On Error Resume Next
    
    lvw.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ReturnSelect()
    Dim lngLoop As Long
    
    If mblnMuliSelect Then
        For lngLoop = 1 To lvw.ListItems.Count
            mrsData.Filter = ""
            
            mrsData.Filter = "ID=" & Mid(lvw.ListItems(lngLoop).Key, 2)
            If lvw.ListItems(lngLoop).Checked Then
                mrsData("选择").Value = 1
            Else
                mrsData("选择").Value = 0
            End If
        Next
    Else
        If mblnLeftSelect = False Then
            If Not (lvw.SelectedItem Is Nothing) Then
                
                mrsData.Filter = ""
                mrsData.Filter = "末级=1 AND ID=" & Val(Mid(lvw.SelectedItem.Key, 2))
            End If
        Else
            If Not (tvw.SelectedItem Is Nothing) Then
                mrsData.Filter = ""
                mrsData.Filter = "末级=0 AND ID=" & Val(Mid(tvw.SelectedItem.Key, 2))
            End If
        End If
    End If
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If mblnMuliSelect = False Then Exit Sub
    
    Call ReturnSelect
        
End Sub

Private Sub Form_Activate()
         
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
     '置缺焦点
    DoEvents
    If lvw.Visible Then
        
        If mblnLeftSelect Then
            tvw.SetFocus
            If Not (tvw.SelectedItem Is Nothing) Then tvw.SelectedItem.EnsureVisible
        Else
            lvw.SetFocus
            If Not (lvw.SelectedItem Is Nothing) Then lvw.SelectedItem.EnsureVisible
        End If
    Else
        tvw.SetFocus
        If Not (tvw.SelectedItem Is Nothing) Then tvw.SelectedItem.EnsureVisible
    End If
    
    If lvw.Visible Then
        stb.Panels(1).Text = "共搜索到 " & lvw.ListItems.Count & " 条结果。"
    Else
        stb.Panels(1).Text = "共搜索到 " & tvw.Nodes.Count & " 条结果。"
    End If
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With picTitle
        .Left = -15
        .Top = -30
        .Width = Me.ScaleWidth + 30
    End With
    
    Select Case mbytWinStyle
    Case 1
    
        With tvw
            .Left = -15
            .Top = picTitle.Top + picTitle.Height
            .Height = Me.ScaleHeight - stb.Height - .Top
            .Width = Me.ScaleWidth - .Left
        End With
        
    Case 2, 4
    
        With lvw
            .Left = -15
            .Top = picTitle.Top + picTitle.Height
            .Height = Me.ScaleHeight - stb.Height - .Top
            .Width = Me.ScaleWidth - .Left
        End With
        
    Case 3
        With tvw
            .Left = -15
            .Top = picTitle.Top + picTitle.Height
            .Height = Me.ScaleHeight - stb.Height - .Top
            .Width = imgX.Left
        End With
        
        With lvw
            .Left = imgX.Left + imgX.Width
            .Top = tvw.Top
            .Width = Me.ScaleWidth - .Left
            .Height = tvw.Height
        End With
        
    End Select
    
    With imgX
        .Top = lvw.Top - 30
        .Height = lvw.Height + 60
    End With
    
    With picDrag
        .Left = Me.ScaleWidth - .Width - 30
        .Top = Me.ScaleHeight - .Height - 30
    End With
    
    With chk
        .Left = stb.Left + stb.Width - IIf(picDrag.Visible, picDrag.Width + 180, 0) - .Width
        .Top = stb.Top + 30 + (stb.Height - .Height) / 2
    End With
    
    With cmdOK
        .Left = IIf(chk.Visible, chk.Left - .Width - 300, stb.Left + stb.Width - .Width - 60)
        .Top = stb.Top + 30 + (stb.Height - .Height) / 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormState
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgX.Left = imgX.Left + X
    
    If imgX.Left < 1500 Then imgX.Left = 1500
    If Me.Width - imgX.Left - imgX.Width < 1000 Then imgX.Left = Me.Width - imgX.Width - 1000
    
    Form_Resize
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvw.SortKey = ColumnHeader.Index - 1 Then
        lvw.SortOrder = IIf(lvw.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvw.SortKey = ColumnHeader.Index - 1
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_DblClick()
    
    If mblnLeftSelect Then Exit Sub
    If mblnMuliSelect Then Exit Sub
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngLoop As Long
    
    If Shift = 2 And KeyCode = vbKeyA Then
        For lngLoop = 1 To lvw.ListItems.Count
            lvw.ListItems(lngLoop).Checked = True
        Next
    End If
    
    If Shift = 1 And KeyCode = vbKeyDelete Then
        For lngLoop = 1 To lvw.ListItems.Count
            lvw.ListItems(lngLoop).Checked = False
        Next
    End If
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Me.Width + X - mlngX < 3990 Then Exit Sub
        If Me.Height + Y - mlngY < 1995 Then Exit Sub
        
        Me.Width = Me.Width + X - mlngX
        Me.Height = Me.Height + Y - mlngY
        Call Form_Resize
    End If
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    
    With cmdClose
        .Left = picTitle.Width - .Width - 30
    End With
End Sub

Private Sub tvw_DblClick()
    If mblnLeftSelect = False And mbytWinStyle <> 1 Then Exit Sub
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '如果重复点击同一节点，则不再刷新数据
       
    If Node.Key <> mstrSvrKey Then
        mstrSvrKey = Node.Key
        
        '先清除数据再装载新数据
        lvw.ListItems.Clear
        Call LoadData(Node.Key)
        
        If tvw.SelectedItem Is Nothing Then
            stb.Panels(1).Text = "没有任何信息！"
        Else
            stb.Panels(1).Text = tvw.SelectedItem.Text & "下共有 " & lvw.ListItems.Count & " 条信息！"
        End If
    End If

End Sub


