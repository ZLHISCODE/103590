VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceScheme 
   AutoRedraw      =   -1  'True
   Caption         =   "保存为成套医嘱"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "frmAdviceScheme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8955
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8940
      Begin VB.OptionButton opt个人 
         Caption         =   "公共(&U)"
         Height          =   180
         Index           =   1
         Left            =   6570
         TabIndex        =   19
         Top             =   1785
         Width           =   930
      End
      Begin VB.OptionButton opt个人 
         Caption         =   "私有(&P)"
         Height          =   180
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   1785
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.CommandButton cmd编码 
         Caption         =   "…"
         Height          =   255
         Left            =   3030
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "按 * 键选择已有方案"
         Top             =   690
         Width           =   285
      End
      Begin VB.Frame fraLine 
         Height          =   60
         Left            =   -60
         TabIndex        =   28
         Top             =   510
         Width           =   9510
      End
      Begin VB.CommandButton cmd分类 
         Caption         =   "…"
         Height          =   255
         Left            =   7125
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "按 * 键选择"
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox txt分类 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4350
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   4
         Top             =   660
         Width           =   3090
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Index           =   0
         Left            =   1095
         MaxLength       =   60
         TabIndex        =   7
         Top             =   1005
         Width           =   2250
      End
      Begin VB.TextBox txt拼音 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   4350
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1005
         Width           =   960
      End
      Begin VB.TextBox txt五笔 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   5970
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1005
         Width           =   960
      End
      Begin VB.TextBox txt五笔 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5970
         MaxLength       =   12
         TabIndex        =   15
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txt拼音 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   4350
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Index           =   1
         Left            =   1095
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1350
         Width           =   2250
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   1095
         MaxLength       =   60
         TabIndex        =   17
         Top             =   1695
         Width           =   4215
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   20
         TabIndex        =   1
         Top             =   660
         Width           =   2250
      End
      Begin VB.Label lbl分类 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分类(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3690
         TabIndex        =   3
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   0
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   6
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lbl简码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&S)           (拼音)            (五笔)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3690
         TabIndex        =   8
         Top             =   1065
         Width           =   3780
      End
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdviceScheme.frx":058A
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   1095
         TabIndex        =   29
         Top             =   75
         Width           =   6555
      End
      Begin VB.Label lbl简码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&M)           (拼音)            (五笔)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3690
         TabIndex        =   13
         Top             =   1410
         Width           =   3780
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "别名(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   11
         Top             =   1410
         Width           =   630
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   435
         Picture         =   "frmAdviceScheme.frx":061C
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   16
         Top             =   1740
         Width           =   630
      End
   End
   Begin VB.Frame fraCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      TabIndex        =   26
      Top             =   7005
      Width           =   9390
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5850
         TabIndex        =   21
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6960
         TabIndex        =   22
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   405
         TabIndex        =   25
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   1575
         TabIndex        =   23
         ToolTipText     =   "Ctrl+A"
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
         Height          =   350
         Left            =   2685
         TabIndex        =   24
         ToolTipText     =   "Ctrl+R"
         Top             =   135
         Width           =   1100
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4920
      Left            =   0
      TabIndex        =   20
      Top             =   2085
      Width           =   8955
      _cx             =   15796
      _cy             =   8678
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceScheme.frx":0EE6
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAdviceScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mint来源 As Integer 'IN:1-门诊,2-住院
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr挂号单 As String
Private mint婴儿 As Integer
Private mblnOK As Boolean
Private Enum COL成套方案
    col选择 = 0
    col期效 = 1
    col内容 = 2
    col总量 = 3
    col总量单位 = 4
    col单量 = 5
    col单量单位 = 6
    col频次 = 7
    col用法 = 8
    col嘱托 = 9
    col执行时间 = 10
    col执行科室 = 11
    col执行性质 = 12
    col序号 = 13
    col相关序号 = 14
    col诊疗类别 = 15
    col诊疗项目ID = 16
    col收费细目ID = 17
    col标本部位 = 18
    col频率次数 = 19
    col频率间隔 = 20
    col间隔单位 = 21
End Enum

Public Function ShowMe(ByVal strPrivs As String, ByVal int来源 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str挂号单 As String, ByVal int婴儿 As Integer, frmParent As Object) As Boolean
    
    mstrPrivs = strPrivs
    mint来源 = int来源
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr挂号单 = str挂号单
    mint婴儿 = int婴儿
    mblnOK = False
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdAll_Click()
    Call Form_KeyDown(vbKeyA, vbCtrlMask)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call Form_KeyDown(vbKeyR, vbCtrlMask)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL() As Variant
    Dim colSerial As New Collection, lng方案ID As Long
    Dim i As Long, j As Long
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt编码.Text) > txt编码.MaxLength Then
        MsgBox "编码的超长（最多" & txt编码.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Sub
    End If
    
    If Trim(Me.txt名称(0).Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称(0).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt名称(0).Text) > txt名称(0).MaxLength Then
        MsgBox "名称超长（" & txt名称(0).MaxLength & "个字符或" & txt名称(0).MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt名称(0).SetFocus: Exit Sub
    End If
    
    If Val(txt分类.Tag) = 0 Then
        MsgBox "请为该成套方案确定一个分类。", vbInformation, gstrSysName
        txt分类.SetFocus: Exit Sub
    End If
    
    If zlCommFun.ActualLen(txt名称(1).Text) > txt名称(1).MaxLength Then
        MsgBox "别名超长（" & txt名称(1).MaxLength & "个字符或" & txt名称(1).MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt名称(1).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt说明.Text) > txt说明.MaxLength Then
        MsgBox "说明超长（" & txt说明.MaxLength & "个字符或" & txt说明.MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt说明.SetFocus: Exit Sub
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, col诊疗项目ID)) = 0 Then
        MsgBox "没有可以保存为成套方案的医嘱！", vbInformation, gstrSysName
        vsAdvice.SetFocus: Exit Sub
    End If
    
    '数据保存
    If Val(txt编码.Tag) = 0 Then
        lng方案ID = zlDatabase.GetNextId("诊疗项目目录")
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text)) Then Exit Sub
    Else
        lng方案ID = Val(txt编码.Tag)
        If zlClinicCodeRepeat(Trim(Me.txt编码.Text), lng方案ID) Then Exit Sub
    End If
    
    arrSQL = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_成套方案项目_Update(" & _
        lng方案ID & "," & Val(Me.txt分类.Tag) & ",'" & Trim(Me.txt编码.Text) & "'," & _
        "'" & Trim(Me.txt名称(0).Text) & "','" & Trim(Me.txt拼音(0).Text) & "','" & Trim(Me.txt五笔(0).Text) & "'," & _
        "'" & Trim(Me.txt名称(1).Text) & "','" & Trim(Me.txt拼音(1).Text) & "','" & Trim(Me.txt五笔(1).Text) & "'," & _
        "'" & Trim(Me.txt说明.Text) & "'," & IIF(opt个人(0).Value, UserInfo.ID, "Null") & ")"
    With vsAdvice
        '记录原来的ID所关联的序号
        j = 1
        colSerial.Add 0, "_0"
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col诊疗项目ID)) <> 0 And Val(.TextMatrix(i, col选择)) <> 0 Then
                colSerial.Add j, "_" & Val(.TextMatrix(i, col序号))
                j = j + 1
            End If
        Next
        
        j = 1
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col诊疗项目ID)) <> 0 And Val(.TextMatrix(i, col选择)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_成套方案内容_Insert(" & _
                    lng方案ID & "," & j & "," & ZVal(colSerial("_" & Val(.TextMatrix(i, col相关序号)))) & "," & _
                    IIF(.TextMatrix(i, col期效) = "长嘱", 0, 1) & "," & Val(.TextMatrix(i, col诊疗项目ID)) & "," & _
                    ZVal(Val(.TextMatrix(i, col单量))) & "," & ZVal(Val(.TextMatrix(i, col总量))) & "," & _
                    ZVal(Val(.TextMatrix(i, col收费细目ID))) & ",'" & .TextMatrix(i, col标本部位) & "'," & _
                    "'" & .TextMatrix(i, col频次) & "'," & ZVal(.TextMatrix(i, col频率次数)) & "," & _
                    ZVal(.TextMatrix(i, col频率间隔)) & ",'" & .TextMatrix(i, col间隔单位) & "'," & _
                    "'" & .TextMatrix(i, col嘱托) & "'," & Val(.Cell(flexcpData, i, col执行性质)) & "," & _
                    ZVal(Val(.Cell(flexcpData, i, col执行科室))) & ",'" & .TextMatrix(i, col执行时间) & "')"
                j = j + 1
            End If
        Next
    End With

    If UBound(arrSQL) = 0 Then
        MsgBox "没有选择要保存为成套方案的医嘱！", vbInformation, gstrSysName
        vsAdvice.SetFocus: Exit Sub
    End If
    
    '提交SQL语句
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
        
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd编码_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim objTmp As Object
    
    strSQL = _
        " Select ID,上级ID,0 as 末级,编码,名称,NULL as 说明" & _
        " From 诊疗分类目录 Where 类型=6" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Union ALL " & _
        " Select ID,分类ID as 上级ID,1 as 末级,编码,名称,标本部位 as 说明" & _
        " From 诊疗项目目录 Where 类别='9'"
        If InStr(mstrPrivs, "保存成套方案") > 0 Then
            strSQL = strSQL & " And (人员ID is Null Or 人员ID=[1])"
        Else
            strSQL = strSQL & " And 人员ID=[1]"
        End If
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "成套方案", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, UserInfo.ID)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "当前还没有其他成套方案可以选择。", vbInformation, gstrSysName
        End If
        txt编码.SetFocus
    Else
        txt编码.Tag = rsTmp!ID
        txt编码.Text = rsTmp!编码
        txt名称(0).Text = rsTmp!名称
        
        On Error GoTo errH
        
        '分类及说明
        strSQL = "Select A.标本部位,A.分类ID,'['||B.编码||']'||B.名称 as 分类" & _
            " From 诊疗项目目录 A,诊疗分类目录 B Where A.分类ID=B.ID(+) And A.ID=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt编码.Tag))
        txt分类.Tag = Nvl(rsTmp!分类ID)
        txt分类.Text = Nvl(rsTmp!分类)
        txt说明.Text = Nvl(rsTmp!标本部位)
        
        '别名及简码
        strSQL = "Select 名称,性质,简码,码类 From 诊疗项目别名 Where 诊疗项目ID=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt编码.Tag))
        With rsTmp
            Do While Not .EOF
                If !性质 = 1 And !码类 = 1 Then Me.txt拼音(0).Text = !简码
                If !性质 = 1 And !码类 = 2 Then Me.txt五笔(0).Text = !简码
                If !性质 = 9 Then Me.txt名称(1).Text = !名称
                If !性质 = 9 And !码类 = 1 Then Me.txt拼音(1).Text = !简码
                If !性质 = 9 And !码类 = 2 Then Me.txt五笔(1).Text = !简码
                .MoveNext
            Loop
        End With
        
        '控件颜色标识
        For Each objTmp In Me.Controls
            If TypeName(objTmp) = "TextBox" Then
                objTmp.ForeColor = &HC00000
            End If
        Next
        
        vsAdvice.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd分类_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    strSQL = "Select ID,上级ID,编码,名称,简码" & _
        " From 诊疗分类目录 Where 类型=6" & _
        " Start With 上级ID is Null Connect by Prior ID=上级ID"
    vPoint = GetCoordPos(fraEdit.Hwnd, txt分类.Left, txt分类.Top)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 1, "成套分类", , txt分类.Text, , , , True, vPoint.x, vPoint.y, txt分类.Height, blnCancel)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有建立成套诊疗分类，请先到诊疗项目管理中建立。", vbInformation, gstrSysName
        End If
    Else
        txt分类.Tag = rsTmp!ID '记录分类ID
        txt分类.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        
        If gint诊疗编码 = 1 And Val(txt编码.Tag) = 0 Then
            Call GetMaxCode
        End If
    End If

    txt分类.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    Else
        With vsAdvice
            If KeyCode = vbKeyA And Shift = vbCtrlMask Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, col诊疗项目ID)) <> 0 Then
                        .TextMatrix(i, col选择) = -1
                    End If
                Next
            ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, col选择) = 0
                Next
            End If
        End With
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is vsAdvice Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    If InStr(mstrPrivs, "保存成套方案") = 0 Then
        opt个人(0).Enabled = False
        opt个人(1).Enabled = False
        opt个人(0).Value = True
    End If
    
    Call GetMaxCode
    Call LoadAdvice
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    fraEdit.Left = 0
    fraEdit.Top = 0
    fraEdit.Width = Me.ScaleWidth
    fraLine.Left = -15
    fraLine.Width = Me.ScaleWidth + 30
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraEdit.Top + fraEdit.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraEdit.Height - fraCommand.Height
    
    fraCommand.Left = 0
    fraCommand.Top = vsAdvice.Top + vsAdvice.Height
    fraCommand.Width = Me.ScaleWidth
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function LoadAdvice() As Boolean
'功能：读取当前病人指定的医嘱
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    '门诊只有在诊病人能够进入,在诊是未转出
    '住院病人选择时已限制了住院数据未转出
    strSQL = "Select Distinct A.ID,A.序号,A.相关ID,A.医嘱期效,A.诊疗项目ID,A.医嘱内容," & _
        " A.单次用量,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行性质," & _
        " Nvl(C.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'<院外执行>')) as 执行科室,A.执行时间方案," & _
        " A.执行科室ID,A.标本部位,B.类别,B.名称,B.计算单位,A.总给予量 as 总量,D.计算单位 as 总量单位,D.id as 收费细目ID" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D" & _
        " Where A.诊疗项目ID=B.ID And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+)" & _
        " And A.医嘱状态 Not IN(2,4) And A.开始执行时间 is Not Null And A.病人来源<>3 And Nvl(A.婴儿,0)=[2]" & _
        IIF(mlng主页ID = 0, " And A.病人ID+0=[1] And A.挂号单=[3]", " And A.病人ID=[1] And A.主页ID=[4]") & _
        " Order by A.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mint婴儿, mstr挂号单, mlng主页ID)
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col选择) = -1
                .TextMatrix(i, col期效) = IIF(Nvl(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .TextMatrix(i, col内容) = rsTmp!医嘱内容
                .TextMatrix(i, col标本部位) = Nvl(rsTmp!标本部位)  '检验标本
                .TextMatrix(i, col单量) = FormatEx(Nvl(rsTmp!单次用量), 4)
                If Not IsNull(rsTmp!单次用量) Then
                    .TextMatrix(i, col单量单位) = Nvl(rsTmp!计算单位)
                End If
                If .TextMatrix(i, col期效) = "临嘱" Then
                    If Not IsNull(rsTmp!总量) Then
                        .TextMatrix(i, col总量) = FormatEx(Nvl(rsTmp!总量), 4)
                        If Not IsNull(rsTmp!总量单位) Then
                            .TextMatrix(i, col总量单位) = Nvl(rsTmp!总量单位)
                        Else
                            .TextMatrix(i, col总量单位) = Nvl(rsTmp!计算单位)
                        End If
                    End If
                End If
                .TextMatrix(i, col频次) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, col频率次数) = Nvl(rsTmp!频率次数)
                .TextMatrix(i, col频率间隔) = Nvl(rsTmp!频率间隔)
                .TextMatrix(i, col间隔单位) = Nvl(rsTmp!间隔单位)
                .TextMatrix(i, col嘱托) = Nvl(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = Nvl(rsTmp!执行时间方案)
                .TextMatrix(i, col执行科室) = Nvl(rsTmp!执行科室)
                .Cell(flexcpData, i, col执行科室) = CLng(Nvl(rsTmp!执行科室ID, 0))
                .Cell(flexcpData, i, col执行性质) = Val(Nvl(rsTmp!执行性质, 0))
                .TextMatrix(i, col序号) = rsTmp!ID
                .TextMatrix(i, col相关序号) = Nvl(rsTmp!相关ID)
                .TextMatrix(i, col诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(i, col诊疗类别) = rsTmp!类别
                .TextMatrix(i, col收费细目ID) = zlCommFun.Nvl(rsTmp!收费细目ID)
                
                '处理行隐藏及用法显示
                If InStr(",C,D,F,G,E,", rsTmp!类别) > 0 And Not IsNull(rsTmp!相关ID) Then
                    .RowHidden(i) = True
                ElseIf rsTmp!类别 = "7" Then
                    .RowHidden(i) = True
                ElseIf rsTmp!类别 = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关序号)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '给药途径
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关序号)) = rsTmp!ID Then
                            .TextMatrix(j, col用法) = rsTmp!名称
                            
                            '显示成药执行性质
                            If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                                .TextMatrix(j, col执行性质) = "自备药"
                            ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                                .TextMatrix(j, col执行性质) = "离院带药"
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf rsTmp!类别 = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关序号)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '中药用法或检验采集方法
                    .TextMatrix(i, col用法) = rsTmp!名称
                    
                    '中药或检验的执行科室
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关序号)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col诊疗类别)) > 0 Then
                                .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '中药付数
                    If .TextMatrix(i - 1, col诊疗类别) <> "C" Then
                        .TextMatrix(i, col总量单位) = "付"
                        
                        '显示中药配方执行性质:以药品为准判断
                        j = .FindRow(CStr(rsTmp!ID), , col相关序号)
                        If Val(.Cell(flexcpData, j, col执行性质)) = 5 And Val(.Cell(flexcpData, i, col执行性质)) <> 5 Then
                            .TextMatrix(i, col执行性质) = "自备药"
                        ElseIf Val(.Cell(flexcpData, j, col执行性质)) <> 5 And Val(.Cell(flexcpData, i, col执行性质)) = 5 Then
                            .TextMatrix(i, col执行性质) = "离院带药"
                        End If
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt编码_GotFocus()
    Call zlControl.TxtSelAll(txt编码)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmd编码_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt分类_GotFocus()
    Call zlControl.TxtSelAll(txt分类)
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmd分类_Click
    End If
End Sub

Private Sub txt名称_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt名称(Index))
End Sub

Private Sub txt名称_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txt拼音(Index).Text = zlCommFun.zlGetSymbol(Me.txt名称(Index).Text, 0)
    Me.txt五笔(Index).Text = zlCommFun.zlGetSymbol(Me.txt名称(Index).Text, 1)
End Sub

Private Sub txt拼音_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt拼音(Index))
End Sub

Private Sub txt拼音_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Call zlControl.TxtSelAll(txt说明)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt五笔_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt五笔(Index))
End Sub

Private Sub txt五笔_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col选择 Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col选择 Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col诊疗项目ID)) <> 0 Then
                    .TextMatrix(.Row, col选择) = IIF(Val(.TextMatrix(.Row, col选择)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col选择 Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(vsAdvice.Row, col诊疗项目ID)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'功能：根据指定行(可能为任意行)的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col相关序号)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col相关序号)) _
                    Or Val(.TextMatrix(i, col序号)) = Val(.TextMatrix(lngRow, col相关序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关序号)) = Val(.TextMatrix(lngRow, col序号)) Then
                    .TextMatrix(i, col选择) = .TextMatrix(lngRow, col选择)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub GetMaxCode()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    If gint诊疗编码 = 1 And Val(txt分类.Tag) <> 0 Then
        '种类+分类+顺序编号
        strTmp = Mid(txt分类.Text, 2, InStr(1, txt分类.Text, "]") - 2)
        strSQL = "Select Nvl(Max(编码),'0000000') as 编码 From 诊疗项目目录 Where 类别='9' And 编码 Like [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "9" & strTmp & "%")
        On Error Resume Next
        txt编码.Text = "9" & strTmp & Right(String(10, "0") & Val(rsTmp!编码) + 1, Len(rsTmp!编码) - 1 - Len(strTmp))
    Else
        '顺序编号
        strSQL = "Select Nvl(Max(编码),'0000000') as 编码 From 诊疗项目目录 Where 类别='9'"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        txt编码.Text = Right(String(10, "0") & Val(rsTmp!编码) + 1, Len(rsTmp!编码))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
