VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelectExplorer 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3930
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6750
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "显示所有下级"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5025
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3705
      Width           =   1380
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   3630
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8731
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   527
            MinWidth        =   527
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   30
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
         TabIndex        =   6
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
      Left            =   4500
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSelectExplorer.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   2790
      Width           =   165
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1290
      Left            =   2670
      TabIndex        =   1
      Top             =   1140
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
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
      Left            =   4800
      Top             =   600
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
            Picture         =   "frmSelectExplorer.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgX 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1860
      Left            =   2280
      MousePointer    =   9  'Size W E
      Top             =   300
      Width           =   30
   End
End
Attribute VB_Name = "frmSelectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mstrStatePath As String
Private mlngX As Long
Private mlngY As Long
Private mstrSvrKey As String
Private mlngSortColumn As Long
Private msglTxtH As Single
Private mstrPrive As String
Private mstrSvrTag As String
Private mstrTitle As String
Private mblnLeftSelect As Boolean
Private mrsData As ADODB.Recordset
Private mblnOK As Boolean

Private Sub SaveFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", mstrStatePath, "宽度", Me.Width
    SaveSetting "ZLSOFT", mstrStatePath, "高度", Me.Height
    SaveSetting "ZLSOFT", mstrStatePath, "分隔条", imgX.Left
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    SaveSetting "ZLSOFT", mstrStatePath, "列宽", strTmp
    
End Sub

Private Sub RestoreFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    Me.Width = GetSetting("ZLSOFT", mstrStatePath, "宽度", Me.Width)
    Me.Height = GetSetting("ZLSOFT", mstrStatePath, "高度", Me.Height)
    imgX.Left = GetSetting("ZLSOFT", mstrStatePath, "分隔条", imgX.Left)
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    strTmp = GetSetting("ZLSOFT", mstrStatePath, "列宽", strTmp)
    
    On Error Resume Next
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        lvw.ColumnHeaders(lngLoop).Width = Val(Split(strTmp, ";")(lngLoop - 1))
    Next
    
    '检查是否超过屏幕高和宽度
    
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
    
    If Me.Top + Me.Height > Screen.Height Then Me.Top = Me.Top - Me.Height - msglTxtH
End Sub

Private Sub LoadData(ByVal strKey As String)

    '功能：装载数据
    '参数：要装载数据的分类关键字
    Dim strKeys As String
    Dim objItem As ListItem
    Dim lngLoop As Long
    Dim strFilter As String
    
    On Error GoTo ErrHand
    
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
    
ErrHand:
        
End Sub

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

Private Sub ReadTreeData()
    '功能：
    
    Dim objItem As Node
    Dim rs As New ADODB.Recordset
    
    mrsData.Filter = ""
    mrsData.Filter = "末级<>1"
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    
    Do While Not mrsData.EOF
        If IIf(IsNull(mrsData("上级id").Value), 0, mrsData("上级id").Value) <> 0 Then
            Set objItem = tvw.Nodes.Add("K" & mrsData("上级ID").Value, tvwChild, "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
        Else
            Set objItem = tvw.Nodes.Add(, , "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
        End If
        mrsData.MoveNext
    Loop
     
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).EnsureVisible
        tvw.Nodes(1).Expanded = True
    End If
End Sub

Public Function ShowSelect(ByVal frmMain As Form, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            ByVal sglTxtH As Single, _
                            Optional StatePath As String, _
                            Optional strLvw As String, _
                            Optional strTitle As String, _
                            Optional blnLeftSelect As Boolean = False, _
                            Optional BackColor As Long = &H80000005, _
                            Optional InitSelectKey As String = "") As Boolean
    
    '功能:显示查询选择器
    '参数:
    '返回:
    
    If rsData.BOF Then Exit Function
    
    Set mrsData = rsData
    
    mblnLeftSelect = blnLeftSelect
    mstrSvrKey = ""
    mblnOK = False
    mstrSvrTag = ""
    mlngSortColumn = 1
    msglTxtH = sglTxtH
    mstrPrive = strLvw
    mstrTitle = strTitle
    
    mstrStatePath = "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & StatePath
    
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
    lblCaption.Caption = strTitle
    
    lvw.ListItems.Clear
    
    zlControl.LvwSelectColumns lvw, strLvw, True
    
    
    If mblnLeftSelect Then tvw.LineStyle = tvwTreeLines
        
    Call RestoreFormState
    
    Call ReadTreeData
    
    If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
        
    Me.Show 1, frmMain
    
    Set rsData = mrsData
    
    ShowSelect = mblnOK
    
End Function

Private Sub chk_Click()
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrSvrKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ReturnSelect()
    
    If mblnLeftSelect = False Then
        If Not (lvw.SelectedItem Is Nothing) Then
            'If mrsData.RecordCount > 0 Then mrsData.MoveFirst
            'mrsData.Move lvw.SelectedItem.Index - 1
            
            mrsData.Filter = ""
            mrsData.Filter = "末级=1 AND ID=" & Mid(lvw.SelectedItem.Key, 2)
            
            mblnOK = True
        End If
    Else
        If Not (tvw.SelectedItem Is Nothing) Then
            mrsData.Filter = ""
            mrsData.Filter = "末级=0 AND ID=" & Mid(tvw.SelectedItem.Key, 2)
            mblnOK = True
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With picTitle
        .Left = -15
        .Top = -30
        .Width = Me.ScaleWidth + 30
    End With

    
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
    
    With imgX
        .Top = lvw.Top - 30
        .Height = lvw.Height + 60
    End With
    
    With picDrag
        .Left = Me.ScaleWidth - .Width - 30
        .Top = Me.ScaleHeight - .Height - 30
    End With
    
    With chk
        chk.Left = stb.Width + stb.Left - picDrag.Width - .Width - 180
        chk.Top = stb.Top + 75
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormState
End Sub

Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    imgX.Left = imgX.Left + x
    
    If imgX.Left < 1500 Then imgX.Left = 1500
    If Me.Width - imgX.Left - imgX.Width < 1000 Then imgX.Left = Me.Width - imgX.Width - 1000
    
    Form_Resize
End Sub

Private Sub lbl_Click()

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
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        mlngX = x
        mlngY = y
    End If
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Me.Width + x - mlngX < 3990 Then Exit Sub
        If Me.Height + y - mlngY < 1995 Then Exit Sub
        
        Me.Width = Me.Width + x - mlngX
        Me.Height = Me.Height + y - mlngY
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
    If mblnLeftSelect = False Then Exit Sub
    
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
