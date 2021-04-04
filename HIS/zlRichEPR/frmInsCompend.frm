VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInsCompend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插入提纲"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7410
   Icon            =   "frmInsCompend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picVBar_S 
      BackColor       =   &H8000000C&
      Height          =   4815
      Left            =   2565
      MouseIcon       =   "frmInsCompend.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4815
      ScaleWidth      =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   30
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   2655
      ScaleHeight     =   4260
      ScaleWidth      =   4575
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   15
      Width           =   4575
      Begin VB.ComboBox cbo级别 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1185
         Width           =   3555
      End
      Begin VB.CheckBox chk保留 
         Caption         =   "保留(&K),在书写时该提纲段不允许删除"
         Height          =   210
         Left            =   855
         TabIndex        =   8
         Top             =   2985
         Width           =   3585
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2250
         TabIndex        =   9
         Top             =   3765
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3360
         TabIndex        =   10
         Top             =   3765
         Width           =   1100
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   3600
         Width           =   4410
      End
      Begin VB.OptionButton opt预制 
         Caption         =   "插入自定义提纲(&2)"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Top             =   705
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton opt预制 
         Caption         =   "插入预制提纲(&1)"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   435
         Width           =   2775
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   855
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1590
         Width           =   3555
      End
      Begin VB.TextBox txt说明 
         Height          =   840
         Left            =   855
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2010
         Width           =   3555
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   1005
         Width           =   4410
      End
      Begin VB.Label lbl预制 
         AutoSize        =   -1  'True
         Caption         =   "注释:   对应预制提纲——###"
         Height          =   180
         Left            =   150
         TabIndex        =   18
         Top             =   3300
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         Caption         =   "命名(&N)"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   1620
         Width           =   630
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   60
         Picture         =   "frmInsCompend.frx":06DC
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lbl要素性质 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可自行设置或从已建立的预制提纲中选择插入。"
         Height          =   180
         Left            =   690
         TabIndex        =   15
         Top             =   135
         Width           =   3870
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         Caption         =   "说明(&S)"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   2010
         Width           =   630
      End
      Begin VB.Label lbl级别 
         AutoSize        =   -1  'True
         Caption         =   "级别(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   1245
         Width           =   630
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   4500
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   2370
      _cx             =   4180
      _cy             =   7937
      Appearance      =   2
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInsCompend.frx":0FA6
      ScrollTrack     =   -1  'True
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
   Begin VB.Label lblTitle 
      BackColor       =   &H80000003&
      Caption         =   " 可选预制提纲(&P)"
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2370
   End
End
Attribute VB_Name = "frmInsCompend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'################################################################################################################
'## 局部变量
'################################################################################################################

Private EditMode As EditModeEnum        '编辑方式（新增、修改）
Private frmParent As frmMain            '父窗体
Private Compends As New cEPRCompends    '父提纲集合
Private Compend As New cEPRCompend      '本地提纲对象
Private edtThis As Object               '编辑器

Private mblnOK As Boolean               '临时变量，表示是否保存修改结果。
Private mlngOldKey As Long              '临时变量，保存原始Key值。
Private mblnWithName As Boolean         '临时变量，是否提纲了名称字符串。（新增提纲时是否选中了一段文本作为名称）

'################################################################################################################
'## 功能：  显示提纲编辑窗体
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub ShowMe(ByRef oParentForm As frmMain, ByRef oedtThis As Object, _
    ByRef oCompends As cEPRCompends, _
    Optional ByVal oCompend As cEPRCompend)
    
    Dim lngCurCompKey As Long, i As Long, j As Long
    Dim ArrayKeys() As Long
    
    Set frmParent = oParentForm
    Set Compends = oCompends
    Set edtThis = oedtThis
    Call FillPreDefinedComps(frmParent.Document.EPRFileInfo.种类)
    
    mblnWithName = False
    If oCompend Is Nothing Then
        EditMode = cprEM_新增
        Me.Caption = "新增提纲"
        If edtThis.Selection.Text <> "" Then
            Me.txt名称.Text = MidB(edtThis.Selection.Text, 1, 40)      '限制名称长度为40。
            mblnWithName = True
        End If
        Compends.UpdateOrdersFromText edtThis   '更新提纲内部序号
        
        lngCurCompKey = frmParent.Document.GetCurCompendKey(frmParent.Editor1)
    Else
        EditMode = cprEM_修改
        Me.Caption = "修改提纲"
        Set Compend = oCompend.Clone(True)
        mlngOldKey = Compend.Key
        Me.txt名称.Text = Compend.名称
        txt说明.Text = Compend.说明
        Me.cbo级别.Tag = Compend.父Key
        If Compend.保留对象 Then chk保留.Value = vbChecked
        Me.lbl预制.Tag = Compend.预制提纲ID
        
        Compends.UpdateOrdersFromText edtThis   '更新提纲内部序号
        
        Dim lS As Long, lE As Long
        Compend.GetPosition frmParent.Editor1, lS, lE
        lngCurCompKey = frmParent.Document.GetCurCompendKey(frmParent.Editor1, lS)
    
    End If
    
    '填充上级列表（允许设置为上级的是当前位置前的提纲，且从最近一个一级提纲开始，到最近一个提纲终止）
    With Me.cbo级别
        .Clear
        .AddItem "1级": .ItemData(.NewIndex) = 0
        If lngCurCompKey > 0 Then
            ' '循环找出该提纲的根级提纲
            i = lngCurCompKey
            ReDim ArrayKeys(1 To 1) As Long
            ArrayKeys(1) = i
            Do While i > 0
                i = Compends.GetParentNodeKey(i)
                If i > 0 Then
                    ReDim Preserve ArrayKeys(1 To UBound(ArrayKeys) + 1) As Long
                    ArrayKeys(UBound(ArrayKeys)) = i
                End If
            Loop
            '装入已经定义的提纲级别
            For j = UBound(ArrayKeys) To 1 Step -1
                .AddItem .ListCount + 1 & "级(“" & Compends("K" & ArrayKeys(j)).名称 & "”的下级)"
                .ItemData(.NewIndex) = ArrayKeys(j)
                If .ItemData(.NewIndex) = Val(.Tag) Then .ListIndex = .NewIndex
                If .ListCount >= 8 Then Exit For
            Next
        End If
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    If Val(Me.lbl预制.Tag) <> 0 Then
        With Me.vfgThis
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) = Me.lbl预制.Tag Then
                    .Row = i
                    Call vfgThis_DblClick
                Else
                    If i = .Rows - 1 Then Me.lbl预制.Tag = 0: Me.lbl预制.ToolTipText = ""
                End If
            Next
        End With
    End If
    Me.Show vbModal, frmParent
End Sub

Private Sub cbo级别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk保留_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk保留_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long, lngStart As Long, lngEnd As Long, i As Long
    
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > 40 Then MsgBox "名称超长（最多40个字符或20个汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > 500 Then MsgBox "说明超长（最多500个字符或250个汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
'    '保证预制提纲不重复
'    If opt预制(0).Value Then
'        With frmParent.Document
'            For i = 1 To .Compends.Count
'                If .Compends(i).预制提纲ID <> 0 Then
'                    If Val(Me.lbl预制.Tag) = .Compends(i).预制提纲ID Then
'                        MsgBox "预制提纲不允许重复！", vbOKOnly + vbInformation, Me.Caption
'                        Exit Sub
'                    End If
'                End If
'            Next
'        End With
'    End If
    If Len(Trim(txt名称)) = 0 Then
        MsgBox "必须输入提纲名称，请重新输入！", vbOKOnly + vbInformation, Me.Caption
        Me.txt名称.SetFocus: Exit Sub
    End If
    If Me.cbo级别.ListIndex = 0 Then
        Compend.父ID = 0
        Compend.父Key = 0
        Compend.Level = 1
    Else
        Compend.父ID = 0
        Compend.父Key = Compends("K" & Me.cbo级别.ItemData(Me.cbo级别.ListIndex)).Key
        Compend.Level = Compends("K" & Me.cbo级别.ItemData(Me.cbo级别.ListIndex)).Level + 1
    End If
    Compend.名称 = Trim(txt名称)
    Compend.说明 = Trim(txt说明)
    If opt预制(0).Value Then
        Compend.预制提纲ID = Val(Me.lbl预制.Tag)
    Else
        Compend.预制提纲ID = 0
    End If
    Compend.保留对象 = (Me.chk保留.Value = vbChecked)
    
    If EditMode = cprEM_修改 Then
        Compends.Remove "K" & mlngOldKey
        Compend.Key = mlngOldKey
        lngKey = Compends.AddExistNode(Compend, True)
        mlngOldKey = lngKey
    Else
        lngKey = Compends.AddExistNode(Compend, False)
        mlngOldKey = lngKey
    End If
    
    '更新修改引起的其他提纲的父Key的关系变化
    Call UpdateParentKeys
    
    If EditMode = cprEM_新增 Then
        frmParent.InsertCompend frmParent.Editor1.Selection.StartPos, frmParent.Editor1.Selection.EndPos, Compends("K" & lngKey), True
        If mblnWithName Then
            lngEnd = lngEnd + 32
            edtThis.Range(lngEnd, lngEnd).Selected
        End If
    Else
        frmParent.ModifyCompend Compends("K" & lngKey)
    End If
    
    Compends.UpdateOrdersFromText edtThis
    Compends.FillTree frmParent.mfrmCompends.Tree, lngKey
    
    mblnOK = True
    Unload Me
End Sub

Private Sub UpdateParentKeys()
    '##################################################################################################################
    '处理插入的提纲 Comp0 所影响到的其他提纲（前一个同层次提纲 Comp1 与后面一个同层次提纲 Comp2 之间的所有提纲）
    '原则： 1、所有Comp1～Comp0之间的提纲，如果Comp0～Comp2之间的提纲有以范围1中提纲为父提纲的，清除父Key；
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngLevel As Long, lngCurPos As Long, lngOldPos As Long
    Dim i As Long, j As Long, blnFinded As Boolean, blnFirst As Boolean
    Dim ArrayPrevKeys() As Long, ArrayNextKeys() As Long, lngPrevCount As Long, lngNextCount As Long
    
    Compends.UpdateOrdersFromText edtThis       '更新顺序
    lngLevel = Compend.Level
    
    '-------------------------------
    
    lngCurPos = edtThis.Selection.StartPos
    lngOldPos = lngCurPos
    
    ReDim ArrayPrevKeys(1 To 1) As Long
    blnFinded = False
    blnFirst = True
    lngPrevCount = 0
    blnFinded = FindPrevKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Do While blnFinded
        If lKey <> Compend.Key Then
            If Compends("K" & lKey).Level < lngLevel Then
                Exit Do
            ElseIf Compends("K" & lKey).Level = lngLevel Then
                '找到前一个同层次的提纲Comp1
                If blnFirst Then
                    ArrayPrevKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayPrevKeys) + 1
                    ReDim Preserve ArrayPrevKeys(1 To i) As Long
                    ArrayPrevKeys(i) = lKey
                End If
                lngPrevCount = lngPrevCount + 1
                Exit Do
            Else
                If blnFirst Then
                    ArrayPrevKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayPrevKeys) + 1
                    ReDim Preserve ArrayPrevKeys(1 To i) As Long
                    ArrayPrevKeys(i) = lKey
                End If
                lngPrevCount = lngPrevCount + 1
            End If
        End If
        lngCurPos = lSS
        blnFinded = FindPrevKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Loop
    
    '-------------------------------
    
    ReDim ArrayNextKeys(1 To 1) As Long
    lngCurPos = lngOldPos
    
    blnFinded = False
    blnFirst = True
    lngNextCount = 0
    blnFinded = FindNextKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Do While blnFinded
        If lKey <> Compend.Key Then
            If Compends("K" & lKey).Level < lngLevel Then
                Exit Do
            ElseIf Compends("K" & lKey).Level = lngLevel Then
                '找到前一个同层次的提纲Comp1
                If blnFirst Then
                    ArrayNextKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayNextKeys) + 1
                    ReDim Preserve ArrayNextKeys(1 To i) As Long
                    ArrayNextKeys(i) = lKey
                End If
                lngNextCount = lngNextCount + 1
                Exit Do
            Else
                If blnFirst Then
                    ArrayNextKeys(1) = lKey
                    blnFirst = False
                Else
                    i = UBound(ArrayNextKeys) + 1
                    ReDim Preserve ArrayNextKeys(1 To i) As Long
                    ArrayNextKeys(i) = lKey
                End If
                lngNextCount = lngNextCount + 1
            End If
        End If
        lngCurPos = lEE
        blnFinded = FindNextKey(edtThis, lngCurPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
    Loop
    
    '-------------------------------

    If lngPrevCount > 0 And lngNextCount > 0 Then
        '处理关联关系
        For i = 1 To UBound(ArrayPrevKeys)
            For j = 1 To UBound(ArrayNextKeys)
                If Compends("K" & ArrayNextKeys(j)).父Key = ArrayPrevKeys(i) Then
                    Compends("K" & ArrayNextKeys(j)).父Key = 0
                    Compends("K" & ArrayNextKeys(j)).父ID = 0
                End If
            Next
        Next
    End If
    '##################################################################################################################
End Sub

Private Sub FillPreDefinedComps(ByVal lngKind As Long)
    '---------------------------------------------
    '功能：填写病历文件目录
    '---------------------------------------------
    Dim RS As New ADODB.Recordset
    Dim i As Long
    gstrSQL = "Select Id, To_Char(对象序号, '000') As 编号, 内容文本 As 名称, 对象属性 As 说明" & _
            " From 病历文件结构" & _
            " Where 文件id Is Null And Substr(使用时机, [1], 1) = '1'" & _
            " Order By 对象序号"
    Err = 0: On Error GoTo errHand
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKind)
    With Me.vfgThis
        .Clear
        Set .DataSource = RS
        .ColWidth(0) = 0
        For i = .FixedCols To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Me.picVBar_S.BackColor = Me.BackColor
End Sub

Private Sub Form_Resize()
    Dim lngHWarp As Long, lngWWarp As Long
    lngHWarp = Me.Height - Me.ScaleHeight
    lngWWarp = Me.Width - Me.ScaleWidth
    With Me.picVBar_S
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
        If .Left < 0 Then .Left = 0
        If .Left > 5000 Then .Left = 5000
    End With
    With Me.lblTitle
        .Top = Me.ScaleTop
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
    With Me.vfgThis
        .Top = Me.lblTitle.Top + Me.lblTitle.Height + 15: .Height = Me.ScaleHeight - .Top
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
    With Me.picBack
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = Me.ScaleTop
    End With
    Me.Width = Me.picBack.Left + Me.picBack.Width + lngWWarp
    Me.Height = Me.picBack.Height + lngHWarp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set Compends = Nothing
    Set Compend = Nothing
    Set edtThis = Nothing
     Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub opt预制_Click(Index As Integer)
    If Me.opt预制(0).Value And Val(Me.lbl预制.Tag) = 0 Then
        MsgBox "必须通过双击列表中的某一行来选择具体的预制提纲！", vbExclamation, gstrSysName
        Me.opt预制(0).Value = False: Me.opt预制(1).Value = True
        Exit Sub
    End If
    If Me.opt预制(0).Value Then
        Me.lbl预制.Visible = True
        If Val(Me.lbl预制.ToolTipText) = 1 Then
            Me.cbo级别.ListIndex = 0: Me.cbo级别.Enabled = False
            Me.chk保留.Value = vbChecked: Me.chk保留.Enabled = False
        Else
            Me.cbo级别.Enabled = True: Me.chk保留.Enabled = True
        End If
    Else
        Me.cbo级别.Enabled = True: Me.chk保留.Enabled = True: Me.lbl预制.Visible = False
    End If
End Sub

Private Sub opt预制_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub opt预制_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar_S.Left = Me.picVBar_S.Left + x: Me.picVBar_S.BackColor = RGB(192, 192, 192)
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picVBar_S.BackColor = Me.BackColor
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_Change()
    ValidControlText txt说明
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgThis_DblClick()
    If vfgThis.Row = 0 Then Exit Sub
    Me.lbl预制.Tag = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 0)
    Me.lbl预制.Caption = "注释:   对应预制提纲——" & Me.vfgThis.TextMatrix(Me.vfgThis.Row, 2)
    Me.lbl预制.ToolTipText = Val(Me.vfgThis.TextMatrix(Me.vfgThis.Row, 1))
    Me.opt预制(0).Value = True
    If EditMode <> cprEM_修改 Then
        Me.txt名称 = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 2)
        Me.txt说明 = Me.vfgThis.TextMatrix(Me.vfgThis.Row, 3)
    End If
    Call opt预制_Click(0)
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vfgThis_DblClick
        KeyAscii = 0
    End If
End Sub


