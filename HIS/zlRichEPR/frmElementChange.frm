VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmElementChange 
   Caption         =   "病历要素联动设置"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14340
   Icon            =   "frmElementChange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7050
   ScaleMode       =   0  'User
   ScaleWidth      =   14453.51
   StartUpPosition =   1  '所有者中心
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   6975
      Left            =   4815
      TabIndex        =   28
      ToolTipText     =   "双击列表中数据行可以在左侧进行修改"
      Top             =   15
      Width           =   9525
      _Version        =   589884
      _ExtentX        =   16801
      _ExtentY        =   12303
      _StockProps     =   0
      BorderStyle     =   2
      ShowGroupBox    =   -1  'True
      ShowItemsInGroups=   -1  'True
   End
   Begin VB.Frame fraThis 
      Height          =   7065
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   4740
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3405
         TabIndex        =   2
         Top             =   6510
         Width           =   1080
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   2340
         TabIndex        =   1
         Top             =   6510
         Width           =   1080
      End
      Begin VB.Frame fraBecouse 
         Caption         =   "变动原因"
         Height          =   2310
         Left            =   210
         TabIndex        =   3
         Top             =   225
         Width           =   4290
         Begin VB.CommandButton cmdDisease 
            Caption         =   "…"
            Height          =   225
            Left            =   3765
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   1470
            Width           =   240
         End
         Begin VB.CommandButton cmdDiagnose 
            Caption         =   "…"
            Height          =   225
            Left            =   3765
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   1845
            Width           =   240
         End
         Begin VB.TextBox txtDiagnose 
            Height          =   270
            Left            =   1485
            TabIndex        =   12
            Top             =   1815
            Width           =   2550
         End
         Begin VB.TextBox txtDisease 
            Height          =   270
            Left            =   1485
            TabIndex        =   11
            Top             =   1440
            Width           =   2550
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "要素"
            Height          =   225
            Index           =   0
            Left            =   405
            TabIndex        =   10
            Top             =   405
            Value           =   -1  'True
            Width           =   690
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "病人疾病"
            Height          =   225
            Index           =   1
            Left            =   390
            TabIndex        =   9
            Top             =   1470
            Width           =   1035
         End
         Begin VB.ComboBox cboElName 
            Height          =   300
            Left            =   1710
            TabIndex        =   7
            Text            =   "cboElName"
            Top             =   360
            Width           =   2325
         End
         Begin VB.OptionButton optBecause 
            Caption         =   "病人诊断"
            Height          =   225
            Index           =   2
            Left            =   390
            TabIndex        =   5
            Top             =   1830
            Width           =   1035
         End
         Begin VB.ComboBox cboContent 
            Height          =   300
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   915
            Width           =   2325
         End
         Begin VB.Label lblElName 
            Caption         =   "名称"
            Height          =   225
            Left            =   1215
            TabIndex        =   14
            Top             =   405
            Width           =   435
         End
         Begin VB.Label lblElVal 
            Caption         =   "内容"
            Height          =   225
            Left            =   1215
            TabIndex        =   13
            Top             =   945
            Width           =   435
         End
      End
      Begin VB.Frame fraSo 
         Caption         =   "变动结果"
         Height          =   3555
         Left            =   210
         TabIndex        =   15
         Top             =   2715
         Width           =   4290
         Begin VB.ComboBox cboAddSentence 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   3180
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "追加词句"
            Height          =   225
            Index           =   4
            Left            =   210
            TabIndex        =   34
            Top             =   3218
            Width           =   1185
         End
         Begin VB.ComboBox cboSameElName 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   2715
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "相同要素同时变更"
            Height          =   360
            Index           =   3
            Left            =   210
            TabIndex        =   31
            Top             =   2685
            Width           =   1185
         End
         Begin VB.ComboBox cboDelElname 
            Height          =   300
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2235
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "删除要素"
            Height          =   225
            Index           =   2
            Left            =   225
            TabIndex        =   29
            Top             =   2280
            Width           =   1260
         End
         Begin VB.OptionButton optSo 
            Caption         =   "要素"
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   21
            Top             =   345
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.ComboBox cboSoElname 
            Height          =   300
            Left            =   1695
            TabIndex        =   20
            Text            =   "cboSoElname"
            Top             =   300
            Width           =   2325
         End
         Begin VB.OptionButton optSo 
            Caption         =   "词句"
            Height          =   225
            Index           =   1
            Left            =   225
            TabIndex        =   19
            Top             =   1200
            Width           =   705
         End
         Begin VB.TextBox txtSoElContent 
            Height          =   270
            Left            =   1695
            TabIndex        =   18
            ToolTipText     =   "以分号分隔"
            Top             =   780
            Width           =   2325
         End
         Begin VB.ComboBox cboSoStCompend 
            Height          =   300
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1425
            Width           =   2325
         End
         Begin VB.ComboBox cboSoSentence 
            Height          =   300
            Left            =   1695
            TabIndex        =   16
            Text            =   "cboSoSentence"
            Top             =   1755
            Width           =   2325
         End
         Begin VB.Label lblSoStCompend 
            Caption         =   "所在提纲"
            Height          =   225
            Left            =   825
            TabIndex        =   26
            Top             =   1470
            Width           =   810
         End
         Begin VB.Label lblSoElname 
            Caption         =   "名称"
            Height          =   225
            Left            =   1200
            TabIndex        =   25
            Top             =   345
            Width           =   435
         End
         Begin VB.Label lblSoElContent 
            Caption         =   "内容(以"";""分隔)"
            Height          =   225
            Left            =   285
            TabIndex        =   24
            Top             =   810
            Width           =   1350
         End
         Begin VB.Label Label5 
            Caption         =   "允许使用的词句"
            Height          =   225
            Left            =   1200
            TabIndex        =   23
            Top             =   1200
            Width           =   1350
         End
         Begin VB.Label lblSentence 
            Caption         =   "词句名称"
            Height          =   225
            Left            =   825
            TabIndex        =   22
            Top             =   1785
            Width           =   810
         End
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   1275
         TabIndex        =   27
         Top             =   6510
         Width           =   1080
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&D)"
         Height          =   350
         Left            =   210
         TabIndex        =   33
         Top             =   6510
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmElementChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng文件ID As Long
Private Enum mCol
    ID = 0
    原因类型
    原因
    结果类型
    结果
    原因要素ID
    原因要素名称
    原因要素内容
    原因病种ID
    原因病种名称
    结果提纲id
    结果提纲名称
    结果要素ID
    结果要素名称
    结果要素值域
    结果原始值域
    结果词句ID
    结果词句名称
    设置表述
End Enum
Public Sub ShowMe(ByVal objParent As Object, ByVal lng文件ID As Long)
    mlng文件ID = lng文件ID
    Me.Show vbModal, objParent
End Sub
Private Sub cboAddSentence_KeyPress(KeyAscii As Integer)
Call zlControl.CboSetIndex(cboAddSentence.hWnd, zlControl.CboMatchIndex(cboAddSentence.hWnd, KeyAscii))
End Sub

Private Sub cboDelElname_KeyPress(KeyAscii As Integer)
Call zlControl.CboSetIndex(cboDelElname.hWnd, zlControl.CboMatchIndex(cboDelElname.hWnd, KeyAscii))
End Sub

Private Sub cboElName_Click()
Dim i As Integer, strItems As String, strItem As String
    On Error Resume Next '不清除错误，因为在调用处会用到
    cboContent.Clear
    strItems = Split(lblElName.Tag, "|")(cboElName.ListIndex)
    For i = 0 To UBound(Split(strItems, ";"))
        strItem = Trim(Split(strItems, ";")(i))
        If strItem <> "自定义" Then
            cboContent.AddItem strItem
        End If
    Next
    If cboContent.ListCount > 0 Then cboContent.ListIndex = 0
End Sub

Private Sub cboElName_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboElName.hWnd, zlControl.CboMatchIndex(cboElName.hWnd, KeyAscii))
    If KeyAscii = vbKeyReturn Then cboElName_Click
End Sub

Private Sub cboSameElName_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSameElName.hWnd, zlControl.CboMatchIndex(cboSameElName.hWnd, KeyAscii))
End Sub

Private Sub cboSoElname_Click()
Dim i As Integer, strTmp As String
    On Error Resume Next '不清除错误，因为在调用处会用到
    txtSoElContent.Text = ""
    txtSoElContent.Tag = Split(lblSoElname.Tag, "|")(cboSoElname.ListIndex)
    For i = 0 To UBound(Split(txtSoElContent.Tag, ";"))
        strTmp = strTmp & ";" & Trim(Split(txtSoElContent.Tag, ";")(i))
    Next
    txtSoElContent.Text = Mid(strTmp, 2)
End Sub
Private Sub cboSoElname_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSoElname.hWnd, zlControl.CboMatchIndex(cboSoElname.hWnd, KeyAscii))
    If KeyAscii = vbKeyReturn Then cboSoElname_Click
End Sub

Private Sub cboSoSentence_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboSoSentence.hWnd, zlControl.CboMatchIndex(cboSoSentence.hWnd, KeyAscii))
End Sub

Private Sub cboSoStCompend_Click()
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select b.Id, b.名称,zlSpellCode(b.名称) 简码 From 病历提纲词句 A, 病历词句示范 B Where a.提纲id = [1] And a.词句分类id = b.分类id Order by B.名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取提纲词句", Val(Split(cboSoStCompend.Tag, ";")(cboSoStCompend.ListIndex)))
    cboSoSentence.Clear: cboSoSentence.Tag = ""
    Do Until rsTemp.EOF
        cboSoSentence.Tag = cboSoSentence.Tag & rsTemp!ID & ";"
        cboSoSentence.AddItem rsTemp!简码 & "-" & rsTemp!名称
        rsTemp.MoveNext
    Loop
End Sub

Private Function Validate() As Boolean
Dim i As Integer
    If optBecause(0).Value Then
        If cboElName.Text = "" Then
            MsgBox "变动原因的要素不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
        If cboContent.Text = "" Then
            MsgBox "变动原因的要素选项不能为空！", vbInformation, gstrSysName
            Exit Function
        End If

    ElseIf optBecause(1).Value Then
        If Trim(txtDisease.Text) = "" Or Val(txtDisease.Tag) = 0 Then
            MsgBox "变动原因疾病不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optBecause(2).Value Then
        If Trim(txtDiagnose.Text) = "" Or Val(txtDiagnose.Tag) = 0 Then
            MsgBox "变动原因诊断不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf fraBecouse.Enabled Then
        MsgBox "没有指定变动原因，请检查！", vbInformation, gstrSysName
        Exit Function
    End If

    If optSo(0).Value Then
        If cboSoElname.Text = "" Then
            MsgBox "变动结果的要素不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
        If txtSoElContent.Text = "" Then
            MsgBox "变动结果的要素选项不能为空！", vbInformation, gstrSysName
            Exit Function
        End If

        For i = 0 To UBound(Split(txtSoElContent.Text, ";"))
            If InStr(txtSoElContent.Tag, Trim(Split(txtSoElContent.Text, ";")(i))) < 1 Then
                MsgBox "变动结果要素选项不在原有选项中！", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    ElseIf optSo(1).Value Then
        If cboSoSentence.Text = "" Then
            MsgBox "变动结果词句不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(2).Value Then
        If cboDelElname.Text = "" Then
            MsgBox "变动结果为删除要素时，删除的要素不能为空！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(3).Value Then
        If cboSameElName.Text = "" Then
            MsgBox "变动结果为相同要素同时变更时，变更要素不能为空", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf optSo(4).Value Then
        If cboAddSentence.Text = "" Then
            MsgBox "变动结果为追加词句时，追加词句不能为空", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    If optBecause(0).Value And optSo(0).Value Then
        If CLng(Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))) = CLng(Val(Split(cboSoElname.Tag, ";")(cboSoElname.ListIndex))) _
            And zl9ComLib.zlStr.NeedName(cboElName.Text) = zl9ComLib.zlStr.NeedName(cboSoElname.Text) Then
            MsgBox "变动原因要素不能引起自身变动！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If optBecause(0).Value And optSo(2).Value Then
        If CLng(Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))) = CLng(Val(Split(cboDelElname.Tag, ";")(cboDelElname.ListIndex))) _
            And zl9ComLib.zlStr.NeedName(cboElName.Text) = zl9ComLib.zlStr.NeedName(cboDelElname.Text) Then
            MsgBox "变动原因要素不能引起自身被删除！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Validate = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    On Error GoTo errHand
    If MsgBox("确实要删除当前病历设置的所有联动关系条件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "Zl_病历变动情况_Delete(" & mlng文件ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "删除原因结果"
    
    Call InitList
    Call FillVfgList
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Exit Sub
    End If
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHand
    With rptList
        If .FocusedRow Is Nothing Then MsgBox "请先选择需要删除的行。", vbInformation, gstrSysName: Exit Sub
        If .FocusedRow.GroupRow Then MsgBox "请先选择需要删除的行。", vbInformation, gstrSysName: Exit Sub '分组行
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Sub
        gstrSQL = "Zl_病历变动情况_Delete(" & mlng文件ID & "," & .FocusedRow.Record.Item(mCol.ID).Value & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "删除原因结果"
    End With
    Call InitList
    Call FillVfgList
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Exit Sub
    End If
End Sub

Private Sub cmdDiagnose_Click()
    txtDiagnose.Text = ""
    Call txtDiagnose_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdDisease_Click()
    txtDisease.Text = ""
    Call txtDisease_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdOK_Click()
'不允许多个因素(病种/要素)控制同一要素
    On Error GoTo errHand
    If Not Validate Then Exit Sub
    If cmdOK.Caption = "修改(&M)" Then '双击列表之后表示修改,需要先删除原有记录
        If MsgBox("你确实要修改当前选中的联动设置条件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then cmdOK.Caption = "增加(&A)": Exit Sub
        gstrSQL = "Zl_病历变动情况_Delete(" & mlng文件ID & "," & rptList.FocusedRow.Record.Item(mCol.ID).Value & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "删除原因结果"
    End If
    
    Dim intBecause   As Integer, lngBecauseID As Long, strElname As String, strElContent As String
    Dim intSo As Integer, lngSoCompendID As Long, lngSoID As Long, strSoElname As String, strSoElContent As String, strSoElOldContent As String
        
    If Not fraBecouse.Enabled Then '相同要素同时变更
        intBecause = 4
    ElseIf optBecause(0).Value Then '要素
        intBecause = 1
    ElseIf optBecause(1).Value Then '疾疾
        intBecause = 2
    ElseIf optBecause(2).Value Then '诊断
        intBecause = 3
    End If
    
    Select Case intBecause
        Case 1
            lngBecauseID = Val(Split(cboElName.Tag, ";")(cboElName.ListIndex))
            strElname = zl9ComLib.zlStr.NeedName(cboElName.Text)
            strElContent = cboContent.Text
        Case 2
            lngBecauseID = Val(txtDisease.Tag)
            strElname = ""
            strElContent = ""
        Case 3
            lngBecauseID = Val(txtDiagnose.Tag)
            strElname = ""
            strElContent = ""
        Case 4
            lngBecauseID = Val(Split(cboSameElName.Tag, ";")(cboSameElName.ListIndex))
            strElname = zl9ComLib.zlStr.NeedName(cboSameElName.Text)
            strElContent = ""
    End Select
    
    If optSo(0).Value Then
        intSo = 1
    ElseIf optSo(1).Value Then
        intSo = 2
    ElseIf optSo(2).Value Then
        intSo = 3
    ElseIf optSo(3).Value Then
        intSo = 4
    ElseIf optSo(4).Value Then
        intSo = 5
    End If
    
    Select Case intSo
        Case 1
            lngSoCompendID = 0
            lngSoID = Val(Split(cboSoElname.Tag, ";")(cboSoElname.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboSoElname.Text)
            strSoElContent = txtSoElContent.Text
            strSoElOldContent = Split(lblSoElname.Tag, "|")(cboSoElname.ListIndex)
        Case 2
            lngSoCompendID = Val(Split(cboSoStCompend.Tag, ";")(cboSoStCompend.ListIndex))
            lngSoID = Val(Split(cboSoSentence.Tag, ";")(cboSoSentence.ListIndex))
            strSoElname = ""
            strSoElContent = ""
            strSoElOldContent = ""
        Case 3
            lngSoCompendID = 0
            lngSoID = Val(Split(cboDelElname.Tag, ";")(cboDelElname.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboDelElname.Text)
            strSoElContent = ""
            strSoElOldContent = ""
        Case 4
            lngSoCompendID = 0
            lngSoID = Val(Split(cboSameElName.Tag, ";")(cboSameElName.ListIndex))
            strSoElname = zl9ComLib.zlStr.NeedName(cboSameElName.Text)
            strSoElContent = ""
            strSoElOldContent = ""
        Case 5
            lngSoCompendID = 0
            lngSoID = Val(Split(cboAddSentence.Tag, ";")(cboAddSentence.ListIndex))
            strSoElname = ""
            strSoElContent = ""
            strSoElOldContent = ""
    End Select
        
    gstrSQL = "Zl_病历变动情况_Update(" & mlng文件ID & "," & intBecause & "," & lngBecauseID & ",'" & strElname & "','" & strElContent & "'," & _
                    intSo & "," & lngSoCompendID & "," & lngSoID & ",'" & strSoElname & "','" & strSoElContent & "','" & strSoElOldContent & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "保存原因结果"
    
    Call InitList
    Call FillVfgList
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitList
    Call FillVfgList
End Sub
Private Sub InitList()
Dim rptCol As ReportColumn

    With rptList
        .Columns.DeleteAll
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因类型, "原因类型", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因, "原因", 80, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Sortable = False: rptCol.Visible = True
        Set rptCol = .Columns.Add(mCol.结果类型, "结果类型", 0, False): rptCol.Editable = False: rptCol.Groupable = False:: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果, "结果", 80, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Sortable = True: rptCol.Visible = True
        Set rptCol = .Columns.Add(mCol.原因要素ID, "原因要素ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因要素名称, "原因要素", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因要素内容, "原因选项", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因病种ID, "原因病种ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.原因病种名称, "原因病种名称", 100, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.结果提纲id, "结果提纲id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果提纲名称, "结果提纲", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果要素ID, "结果要素ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果要素名称, "结果要素", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果要素值域, "结果值域", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果原始值域, "结果原始值域", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果词句ID, "结果词句ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.结果词句名称, "结果词句", 140, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.设置表述, "设置表述", 1280, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Visible = True
         
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = True
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = ""
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
End Sub
Private Sub FillVfgList()
Dim rsBecouse As ADODB.Recordset, rsSo As ADODB.Recordset
Dim rptRcd As ReportRecord, i As Integer, rptItem As ReportRecordItem, strShowF As String, strShowS As String

    rptList.Records.DeleteAll
    gstrSQL = "Select ID, 原因, 原因要素id, 原因要素名称, 原因要素内容, 原因病种id, 原因病种名称" & vbNewLine & _
                "From (Select a.Id, 1 原因, b.诊治要素id 原因要素id, b.要素名称 原因要素名称, a.原因内容 原因要素内容, 0 原因病种id, '' 原因病种名称" & vbNewLine & _
                "       From 病历变动原因 A, 病历文件结构 B" & vbNewLine & _
                "       Where a.病历文件id = [1] And a.变动原因 = 1 And a.病历文件id = b.文件id And a.原因要件id = Nvl(b.诊治要素id, 0) And a.原因要素 = b.要素名称" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 2 原因, 0 原因要素id, '' 原因要素名称, '' 原因要素内容, b.Id 原因病种id, b.名称 原因病种名称" & vbNewLine & _
                "       From 病历变动原因 A, 疾病编码目录 B" & vbNewLine & _
                "       Where a.病历文件id = [1] And a.变动原因 = 2 And a.原因要件id = b.Id" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 3 原因, 0 原因要素id, '' 原因要素名称, '' 原因要素内容, b.Id 原因病种id, b.名称 原因病种名称" & vbNewLine & _
                "       From 病历变动原因 A, 疾病诊断目录 B" & vbNewLine & _
                "       Where a.病历文件id = [1] And a.变动原因 = 3 And a.原因要件id = b.Id" & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, 4 原因, b.诊治要素id 原因要素id, b.要素名称 原因要素名称, a.原因内容 原因要素内容, 0 原因病种id, '' 原因病种名称" & vbNewLine & _
                "       From 病历变动原因 A, 病历文件结构 B" & vbNewLine & _
                "       Where a.病历文件id = [1] And a.变动原因 = 4 And a.病历文件id = b.文件id And a.原因要件id = Nvl(b.诊治要素id, 0) And a.原因要素 = b.要素名称)"
    Set rsBecouse = zlDatabase.OpenSQLRecord(gstrSQL, "提取变动原因", mlng文件ID)
    Do Until rsBecouse.EOF
        gstrSQL = "Select 结果, 结果提纲id, 结果提纲名称, 结果要素id, 结果要素名称, 结果要素值域,结果原始值域, 结果词句id, 结果词句名称" & vbNewLine & _
                    "From (Select 1 结果, 0 结果提纲id, '' 结果提纲名称, b.诊治要素id 结果要素id, b.要素名称 结果要素名称, a.结果值域 结果要素值域,a.原始值域 结果原始值域, 0 结果词句id, '' 结果词句名称" & vbNewLine & _
                    "       From 病历变动结果 A, 病历文件结构 B" & vbNewLine & _
                    "       Where a.变动原因id = [1] And a.变动结果 = 1 And b.文件id = [2] And a.结果要件id = Nvl(b.诊治要素id, 0) And a.结果要素 = b.要素名称" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 2 结果, b.Id 结果提纲id, b.内容文本 结果提纲名称, 0 结果要素id, '' 结果要素名称, '' 结果要素值域,'' 结果原始值域 , c.Id 结果词句id, c.名称 结果词句名称" & vbNewLine & _
                    "       From 病历变动结果 A, 病历文件结构 B, 病历词句示范 C" & vbNewLine & _
                    "       Where a.变动原因id = [1] And a.变动结果 = 2 And b.文件id = [2] And a.病历提纲id = b.Id And a.结果要件id = c.Id" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 3 结果, 0 结果提纲id, '' 结果提纲名称, b.诊治要素id 结果要素id, b.要素名称 结果要素名称, a.结果值域 结果要素值域,a.原始值域 结果原始值域, 0 结果词句id, '' 结果词句名称" & vbNewLine & _
                    "       From 病历变动结果 A, 病历文件结构 B" & vbNewLine & _
                    "       Where a.变动原因id = [1] And a.变动结果 = 3 And b.文件id = [2] And a.结果要件id = Nvl(b.诊治要素id, 0) And a.结果要素 = b.要素名称" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 4 结果, 0 结果提纲id, '' 结果提纲名称, b.诊治要素id 结果要素id, b.要素名称 结果要素名称, a.结果值域 结果要素值域,a.原始值域 结果原始值域, 0 结果词句id, '' 结果词句名称" & vbNewLine & _
                    "       From 病历变动结果 A, 病历文件结构 B" & vbNewLine & _
                    "       Where a.变动原因id = [1] And a.变动结果 = 4 And b.文件id = [2] And a.结果要件id = Nvl(b.诊治要素id, 0) And a.结果要素 = b.要素名称" & vbNewLine & _
                    "       Union" & vbNewLine & _
                    "       Select 5 结果, 0 结果提纲id, '' 结果提纲名称, 0 结果要素id, '' 结果要素名称, '' 结果要素值域,'' 结果原始值域 , c.Id 结果词句id, c.名称 结果词句名称" & vbNewLine & _
                    "       From 病历变动结果 A, 病历词句示范 C" & vbNewLine & _
                    "       Where a.变动原因id = [1] And a.变动结果 = 5 And a.结果要件id = c.Id)"
        Set rsSo = zlDatabase.OpenSQLRecord(gstrSQL, "提取变动结果", CLng(rsBecouse!ID), mlng文件ID)
        Do Until rsSo.EOF
            strShowF = "": strShowS = ""
            With rptList
                Set rptRcd = rptList.Records.Add()
                rptRcd.AddItem CStr(Val(rsBecouse!ID))
                rptRcd.AddItem CStr(Val(rsBecouse!原因))
                Select Case Val(rsBecouse!原因)
                    Case 1
                        Set rptItem = rptRcd.AddItem("要素变更")
                        rptItem.GroupCaption = CStr("因为要素选项更改引起的变化")
                        strShowF = "当要素<" & NVL(rsBecouse!原因要素名称) & ">" & " 选项:" & (NVL(rsBecouse!原因要素内容)) & " 被选中后,"
                    Case 2
                        Set rptItem = rptRcd.AddItem("符合疾病")
                        rptItem.GroupCaption = CStr("因为<最后诊断>符合《疾病标准》设置引起的变化")
                        strShowF = "当病人最后诊断为:<" & NVL(rsBecouse!原因病种名称) & ">时,"
                    Case 3
                        Set rptItem = rptRcd.AddItem("符合诊断")
                        rptItem.GroupCaption = CStr("因为<最后诊断>符合《诊断标准》设置引起的变化")
                        strShowF = "当病人最后诊断为:<" & NVL(rsBecouse!原因病种名称) & ">时,"
                    Case 4
                        Set rptItem = rptRcd.AddItem("相同要素")
                        rptItem.GroupCaption = CStr("因为符合<相同要素同时更新>设置引起的变化")
                        strShowF = "当要素<" & NVL(rsBecouse!原因要素名称) & ">" & "内容发生变化时，本病历内相同要素同时更新"
                    Case Else: rptRcd.AddItem CStr("未知设置")
                End Select
                rptRcd.AddItem CStr(Val(rsSo!结果))
                Select Case Val(rsSo!结果)
                    Case 1
                        rptRcd.AddItem CStr("要素变化")
                        strShowS = "要素<" & NVL(rsSo!结果要素名称) & ">的可选项为:(" & NVL(rsSo!结果要素值域) & ")"
                    Case 2
                        rptRcd.AddItem CStr("允许词句")
                        strShowS = "允许词句:<" & NVL(rsSo!结果词句名称) & ">在提纲:<" & NVL(rsSo!结果提纲名称) & ">中出现"
                    Case 3:  rptRcd.AddItem CStr("册除要素")
                        strShowS = "删除病历内的要素<" & NVL(rsSo!结果要素名称) & ">"
                    Case 4:  rptRcd.AddItem CStr("要素更新")
                        strShowS = ""
                    Case 5: rptRcd.AddItem CStr("追加词句")
                        strShowS = "自动追加词句<" & rsSo!结果词句名称 & ">在当前位置"
                    Case Else: rptRcd.AddItem CStr("未知设置")
                End Select

                rptRcd.AddItem CStr(NVL(rsBecouse("原因要素id"), 0))
                rptRcd.AddItem CStr(NVL(rsBecouse!原因要素名称))
                rptRcd.AddItem CStr(NVL(rsBecouse!原因要素内容))
                rptRcd.AddItem CStr(NVL(rsBecouse("原因病种ID"), 0))
                rptRcd.AddItem CStr(NVL(rsBecouse!原因病种名称))
                rptRcd.AddItem CStr(NVL(rsSo("结果提纲ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!结果提纲名称))
                rptRcd.AddItem CStr(NVL(rsSo("结果要素ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!结果要素名称))
                rptRcd.AddItem CStr(NVL(rsSo!结果要素值域))
                rptRcd.AddItem CStr(NVL(rsSo!结果原始值域))
                rptRcd.AddItem CStr(NVL(rsSo("结果词句ID"), 0))
                rptRcd.AddItem CStr(NVL(rsSo!结果词句名称))
                rptRcd.AddItem strShowF & strShowS
            End With
            rsSo.MoveNext
        Loop
        rsBecouse.MoveNext
    Loop
    
    rptList.GroupsOrder.Add rptList.Columns.Find(mCol.原因)
    rptList.GroupsOrder(0).SortAscending = True
    rptList.Populate
    
    If Me.Visible = False Then
        Call optBecause_Click(0)
        Call optSo_Click(0)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rptList.Move fraThis.Width + 50, 0, Me.ScaleWidth - fraThis.Width - 50, Me.ScaleHeight
    
End Sub

Private Sub optBecause_Click(Index As Integer)
    txtDisease.Enabled = False: cmdDisease.Enabled = False
    txtDiagnose.Enabled = False: cmdDiagnose.Enabled = False
    cboElName.Enabled = False: cboElName.Clear
    cboContent.Enabled = False: cboContent.Clear
    Select Case Index
        Case 0
            cboElName.Enabled = True: cboElName.Clear
            cboContent.Enabled = True: cboContent.Clear
            Dim rsTemp As ADODB.Recordset
            gstrSQL = "Select 诊治要素id,zlSpellCode(要素名称) 要素简码, 要素名称, 要素值域,对象属性" & vbNewLine & _
                        "From 病历文件结构" & vbNewLine & _
                        "Where 文件id = [1] And 对象类型 = 4 And 要素表示 In (2, 3)" & vbNewLine & _
                        "Order By 要素名称"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素", mlng文件ID)
            cboElName.Clear: cboElName.Tag = ""
            Do Until rsTemp.EOF
                cboElName.Tag = cboElName.Tag & NVL(rsTemp!诊治要素ID, 0) & ";" '记录诊治要素ID，以“;”分隔,维数与cboelName.listindex同步
                cboElName.AddItem Replace(rsTemp!要素简码, "-", "") & "-" & rsTemp!要素名称
                lblElName.Tag = lblElName.Tag & rsTemp!要素值域 & "|"   '以"|"分隔,维数与cboelName.listindex同步,每个值域以";"分隔
                rsTemp.MoveNext
            Loop
            If cboElName.ListCount > 0 Then cboElName.ListIndex = 0
        Case 1
            txtDisease.Enabled = True: cmdDisease.Enabled = True
        Case 2
            txtDiagnose.Enabled = True: cmdDiagnose.Enabled = True
    End Select
End Sub

Private Sub optSo_Click(Index As Integer)
Dim rsTemp As ADODB.Recordset
    fraBecouse.Enabled = True
    optBecause(1).Enabled = True
    optBecause(2).Enabled = True
    
    cboDelElname.Enabled = False: cboDelElname.Clear
    cboSameElName.Enabled = False: cboSameElName.Clear
    cboSoStCompend.Enabled = False: cboSoStCompend.Clear
    cboSoSentence.Enabled = False: cboSoSentence.Clear
    cboSoElname.Enabled = False: cboSoElname.Clear
    cboAddSentence.Enabled = False: cboAddSentence.Clear
    txtSoElContent.Enabled = False: txtSoElContent.Text = ""
    
    Select Case Index
        Case 0
            cboSoElname.Enabled = True: cboSoElname.Clear
            txtSoElContent.Enabled = True: txtSoElContent.Text = ""
            gstrSQL = "Select 诊治要素id,zlSpellCode(要素名称) 要素简码, 要素名称,要素值域,对象属性" & vbNewLine & _
                        "From 病历文件结构" & vbNewLine & _
                        "Where 文件id = [1] And 对象类型 = 4 And 要素表示 In (2, 3)" & vbNewLine & _
                        "Order By 要素名称"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素", mlng文件ID)
            cboSoElname.Clear: cboSoElname.Tag = "": lblSoElname.Tag = ""
            Do Until rsTemp.EOF
                If Val(Mid(NVL(rsTemp!对象属性, ""), 5, 1)) = 0 Then '动态域不能用作结果联动
                    cboSoElname.Tag = cboSoElname.Tag & NVL(rsTemp!诊治要素ID, 0) & ";" '记录诊治要素ID，以“;”分隔,维数与cboelName.listindex同步
                    cboSoElname.AddItem Replace(rsTemp!要素简码, "-", "") & "-" & rsTemp!要素名称
                    lblSoElname.Tag = lblSoElname.Tag & rsTemp!要素值域 & "|"   '以"|"分隔,维数与cboelName.listindex同步,每个值域以";"分隔
                End If
                rsTemp.MoveNext
            Loop
            If cboSoElname.ListCount > 0 Then cboSoElname.ListIndex = 0
        Case 1
            cboSoStCompend.Enabled = True: cboSoStCompend.Clear
            cboSoSentence.Enabled = True: cboSoSentence.Clear
            gstrSQL = "Select ID,内容文本 提纲名称" & vbNewLine & _
                        "From 病历文件结构" & vbNewLine & _
                        "   Where 文件id = [1] And 对象类型 = 1" & vbNewLine & _
                        "Order By ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取提纲", mlng文件ID)
            cboSoStCompend.Clear: cboSoStCompend.Tag = ""
            Do Until rsTemp.EOF
                cboSoStCompend.Tag = cboSoStCompend.Tag & rsTemp!ID & ";"
                cboSoStCompend.AddItem rsTemp!提纲名称
                rsTemp.MoveNext
            Loop
            If cboSoStCompend.ListCount > 0 Then cboSoStCompend.ListIndex = 0
        Case 2
            cboDelElname.Enabled = True: cboDelElname.Clear
            gstrSQL = "Select 诊治要素id,zlSpellCode(要素名称) 要素简码, 要素名称,要素值域,对象属性" & vbNewLine & _
                        "From 病历文件结构" & vbNewLine & _
                        "Where 文件id = [1] And 对象类型 = 4" & vbNewLine & _
                        "Order By 要素名称"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素", mlng文件ID)
            cboDelElname.Clear: cboDelElname.Tag = ""
            Do Until rsTemp.EOF
                    cboDelElname.Tag = cboDelElname.Tag & NVL(rsTemp!诊治要素ID, 0) & ";" '记录诊治要素ID，以“;”分隔,维数与cboelName.listindex同步
                    cboDelElname.AddItem Replace(rsTemp!要素简码, "-", "") & "-" & rsTemp!要素名称
                rsTemp.MoveNext
            Loop
            If cboDelElname.ListCount > 0 Then cboDelElname.ListIndex = 0
        Case 3
            fraBecouse.Enabled = False
            cboSameElName.Enabled = True: cboSameElName.Clear
            gstrSQL = "Select 诊治要素id,zlSpellCode(要素名称) 要素简码, 要素名称,要素值域,对象属性" & vbNewLine & _
                        "From 病历文件结构" & vbNewLine & _
                        "Where 文件id = [1] And 对象类型 = 4" & vbNewLine & _
                        "Order By 要素名称"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素", mlng文件ID)
            cboSameElName.Clear: cboSameElName.Tag = ""
            Do Until rsTemp.EOF
                cboSameElName.Tag = cboSameElName.Tag & NVL(rsTemp!诊治要素ID, 0) & ";" '记录诊治要素ID，以“;”分隔,维数与cboelName.listindex同步
                cboSameElName.AddItem Replace(rsTemp!要素简码, "-", "") & "-" & rsTemp!要素名称
                rsTemp.MoveNext
            Loop
            If cboSameElName.ListCount > 0 Then cboSameElName.ListIndex = 0
        Case 4
            optBecause(0).Enabled = True
            optBecause(1).Enabled = False
            optBecause(2).Enabled = False
            optBecause(0).Value = True
            cboAddSentence.Enabled = True: cboAddSentence.Clear
            gstrSQL = "Select c.Id,zlSpellCode(C.名称) 简码,C.名称" & vbNewLine & _
                        "From 病历文件结构 A, 病历提纲词句 B, 病历词句示范 C" & vbNewLine & _
                        "Where a.文件id = [1] And a.对象类型 = 1 And a.Id = b.提纲id And b.词句分类id =C.分类ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取可用词句", mlng文件ID)
            cboAddSentence.Clear: cboAddSentence.Tag = ""
            Do Until rsTemp.EOF
                cboAddSentence.Tag = cboAddSentence.Tag & NVL(rsTemp!ID, 0) & ";"
                cboAddSentence.AddItem Replace(rsTemp!简码, "-", "") & "-" & rsTemp!名称
                rsTemp.MoveNext
            Loop
            If cboAddSentence.ListCount Then cboAddSentence.ListIndex = 0
    End Select

End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    With rptList
        If .FocusedRow Is Nothing Or .FocusedRow.GroupRow Then Exit Sub '分组行
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Sub
        cmdOK.Caption = "修改(&M)"
    End With
End Sub
Private Function ShowItem() As Boolean
    On Error GoTo errHand
    With rptList
        If .FocusedRow Is Nothing Then Exit Function
        If .FocusedRow.GroupRow Then Exit Function '分组行
        If .FocusedRow.Record.Item(mCol.ID).Value = 0 Then Exit Function
        
        With .FocusedRow.Record
            fraBecouse.Enabled = True
            '处理原因显示
            Select Case .Item(mCol.原因类型).Value
                Case 1
                    optBecause(0).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboElName, .Item(mCol.原因要素名称).Value)
                    Call zl9ComLib.cbo.SeekIndex(cboContent, .Item(mCol.原因要素内容).Value)
                Case 2
                    optBecause(1).Value = True
                    txtDisease.Tag = .Item(mCol.原因病种ID).Value
                    txtDisease.Text = .Item(mCol.原因病种名称).Value
                Case 3
                    optBecause(2).Value = True
                    txtDiagnose.Tag = .Item(mCol.原因病种ID).Value
                    txtDiagnose.Text = .Item(mCol.原因病种名称).Value
                Case 4
                    optBecause(0).Value = False: optBecause(1).Value = False: optBecause(2).Value = False
                    fraBecouse.Enabled = False
            End Select
            
            '处理结果显示
            Select Case .Item(mCol.结果类型).Value
                Case 1
                    optSo(0).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboSoElname, .Item(mCol.结果要素名称).Value)
                    txtSoElContent.Text = .Item(mCol.结果要素值域).Value
                Case 2
                    optSo(1).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboSoStCompend, .Item(mCol.结果提纲名称).Value)
                    Call zl9ComLib.cbo.SeekIndex(cboSoSentence, .Item(mCol.结果词句名称).Value)
                Case 3
                    optSo(2).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboDelElname, .Item(mCol.结果要素名称).Value)
                Case 4
                    optSo(3).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboDelElname, .Item(mCol.结果要素名称).Value)
                Case 5
                    optSo(4).Value = True
                    Call zl9ComLib.cbo.SeekIndex(cboAddSentence, .Item(mCol.结果词句名称).Value)
            End Select
        End With
    End With
    If Err.Number <> 0 Then GoTo errHand
    Exit Function
errHand:
    MsgBox "当前病历文件内容发生变化，设置条件中原因或结果指定的项已不存在，将被自动删除。", vbInformation, gstrSysName
    Call cmdDel_Click
    Err.Number = 0: Err.Clear
End Function
Private Sub rptList_SelectionChanged()
    Call ShowItem
    If cmdOK.Caption = "修改(&M)" Then
        optBecause(0).Value = True
        optSo(0).Value = True
        cmdOK.Caption = "增加(&A)"
    End If
End Sub

Private Sub txtDiagnose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim objDiagnose As RECT, rsTemp As ADODB.Recordset
        objDiagnose = GetControlRect(txtDiagnose.hWnd)
        If Trim(txtDiagnose.Text) <> "" Then
            gstrSQL = "Select A.ID,A.编码,A.名称 From 疾病诊断目录 A,疾病诊断别名 B Where A.ID=B.诊断ID AND (A.编码=[1]" & _
                                                            " or " & ZLCommFun.GetLike("A", "名称", txtDiagnose.Text) & _
                                                            " or " & ZLCommFun.GetLike("B", "简码", txtDiagnose.Text) & ")" & _
                                                            " And (A.撤档时间 Is Null Or A.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Else
            gstrSQL = "Select ID,编码, 名称 From 疾病诊断目录 A Where 撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')"
        End If
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "选择疾病", True, txtDiagnose.Text, "", True, True, True, objDiagnose.Left, objDiagnose.Top, txtDiagnose.Height, True, False, True, CStr(txtDiagnose.Text))
        If Not rsTemp Is Nothing Then
            txtDiagnose.Tag = rsTemp!ID: txtDiagnose.Text = rsTemp!名称
        Else
            zlControl.TxtSelAll txtDisease
        End If
    End If
End Sub

Private Sub txtDisease_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim objDisease As RECT, rsTemp As ADODB.Recordset
        objDisease = GetControlRect(txtDisease.hWnd)
        If Trim(txtDisease.Text) <> "" Then
            gstrSQL = "Select ID,编码, 名称 From 疾病编码目录 A Where (编码=[1]" & _
                                                            " or " & ZLCommFun.GetLike("A", "名称", txtDisease.Text) & _
                                                            " or " & ZLCommFun.GetLike("A", "简码", txtDisease.Text) & _
                                                            " or " & ZLCommFun.GetLike("A", "五笔码", txtDisease.Text) & ")" & _
                                                            " And (撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Else
            gstrSQL = "Select ID,编码, 名称 From 疾病编码目录 A Where 撤档时间 Is Null Or 撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')"
        End If
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "选择疾病", True, txtDisease.Text, "", True, True, True, objDisease.Left, objDisease.Top, txtDisease.Height, True, False, True, CStr(txtDisease.Text))
        If Not rsTemp Is Nothing Then
            txtDisease.Tag = rsTemp!ID: txtDisease.Text = rsTemp!名称
        Else
            zlControl.TxtSelAll txtDisease
        End If
    End If
End Sub
Private Sub txtSoElContent_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "," Then
        KeyAscii = 0
    End If
End Sub
