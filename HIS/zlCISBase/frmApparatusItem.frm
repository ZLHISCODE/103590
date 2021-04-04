VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApparatusItem 
   BorderStyle     =   0  'None
   Caption         =   "仪器项目通道"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2490
      Left            =   135
      TabIndex        =   4
      Top             =   105
      Width           =   8145
      _cx             =   14367
      _cy             =   4392
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   135
      ScaleHeight     =   2505
      ScaleWidth      =   8145
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2640
      Width           =   8145
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   450
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "区分检验类型(&K)"
         Height          =   210
         Left            =   6060
         TabIndex        =   8
         Top             =   1935
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   6060
         TabIndex        =   1
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找…    "
         Height          =   350
         Left            =   6060
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧ 添加到仪器项目列表中"
         Height          =   350
         Index           =   0
         Left            =   15
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨ 从仪器项目列表中删除"
         Height          =   350
         Index           =   1
         Left            =   2610
         TabIndex        =   6
         Top             =   45
         Width           =   2535
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "区分大小写(&U)"
         Height          =   210
         Left            =   6060
         TabIndex        =   7
         Top             =   1605
         Width           =   1755
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查找内容:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6060
         TabIndex        =   0
         Top             =   495
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmApparatusItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long          '当前显示的仪器id
Private mstr类型 As String          '当前项目的检验类型

Private Enum mCol
    ID = 0: 序号: 编码: 中文名: 英文名: 类型: 通道码: 精度: 加算值: 换算比: 糖耐量项目
End Enum

Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 12: .FixedCols = 0
        End If
        .ColDataType(mCol.糖耐量项目) = flexDTBoolean
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.序号) = "序号": .TextMatrix(0, mCol.编码) = "编码"
        .TextMatrix(0, mCol.中文名) = "中文名": .TextMatrix(0, mCol.英文名) = "英文名": .TextMatrix(0, mCol.类型) = "类型"
        .TextMatrix(0, mCol.通道码) = "通道码": .TextMatrix(0, mCol.精度) = "精度"
        .TextMatrix(0, mCol.加算值) = "加算值": .TextMatrix(0, mCol.换算比) = "换算比": .TextMatrix(0, mCol.糖耐量项目) = "糖耐项目"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.序号) = 450: .ColWidth(mCol.编码) = 800
        .ColWidth(mCol.中文名) = 1900: .ColWidth(mCol.英文名) = 1400: .ColWidth(mCol.类型) = 0
        .ColWidth(mCol.通道码) = 720: .ColWidth(mCol.精度) = 510
        .ColWidth(mCol.加算值) = 630: .ColWidth(mCol.换算比) = 630: .ColWidth(mCol.糖耐量项目) = 800
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColAlignment(mCol.序号) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngAptId As Long) As Boolean
    '功能：根据仪器id刷新当前显示内容
    '参数：当前项目id
    Dim rsTemp As New ADODB.Recordset
    mLngAptId = lngAptId
    Me.txtFind.Text = ""
    Me.lvwItem.ListItems.Clear
        
    If lngAptId = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.诊治项目id As ID, Rownum As 序号, I.编码, I.名称 As 中文名, L.缩写 As 英文名, L.结果类型 As 类型," & vbNewLine & _
            "       C.通道编码 As 通道码, C.小数位数 As 精确度, C.加算值, C.换算比,C.糖耐量项目" & vbNewLine & _
            "From 检验仪器项目 C, 检验项目 L, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
            "Where C.项目id = L.诊治项目id And L.诊治项目id = R.报告项目id And R.诊疗项目id = I.ID And I.组合项目 <> 1 And" & vbNewLine & _
            "   (I.撤档时间>sysdate or I.撤档时间 is null) And    L.项目类别 <> 2 And C.仪器id = [1] order by i.编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgList.Rows > Me.vfgList.FixedRows Then Me.vfgList.Row = Me.vfgList.FixedRows
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    '参数： lngAptId-指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    mstr类型 = ""
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 仪器类型 From 检验仪器 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    If rsTemp.RecordCount > 0 Then mstr类型 = "" & rsTemp!仪器类型
    If mstr类型 = "" Then Me.chkKind.Value = vbUnchecked
        
    Me.Tag = "编辑": Call Form_Resize
    If Me.Visible Then Me.txtFind.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mLngAptId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strLists As String, strItems As String, dblValue As Double
    Dim strListsS As String
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.ID)) = 0 Then
                MsgBox "第" & lngCount & "行项目不确定！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If Trim(.TextMatrix(lngCount, mCol.通道码)) = "" Then
                MsgBox "第" & lngCount & "行“通道码”未填写！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(.TextMatrix(lngCount, mCol.通道码)), vbFromUnicode)) > 20 Then
                MsgBox "第" & lngCount & "行“通道码”超过长度(20个字符)！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.精度))
            If dblValue > 999999 Or Val(dblValue) - Int(Val(dblValue)) > 0 Then
                MsgBox "第" & lngCount & "行“精度”太大！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.加算值))
            If dblValue > 999999 Or Val(dblValue * 100000) - Int(Val(dblValue * 100000)) > 0 Then
                MsgBox "第" & lngCount & "行“加算值”太大或精度太高！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            dblValue = Val(.TextMatrix(lngCount, mCol.换算比))
            If dblValue > 999999 Or Val(dblValue * 100000) - Int(Val(dblValue * 100000)) > 0 Then
                MsgBox "第" & lngCount & "行“换算比”太大或精度太高！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            strItems = .TextMatrix(lngCount, mCol.ID)
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.通道码))
            If Val(.TextMatrix(lngCount, mCol.类型)) = 1 Or Val(.TextMatrix(lngCount, mCol.类型)) = 3 Then
                If Trim(.TextMatrix(lngCount, mCol.精度)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.精度))
                End If
                If Trim(.TextMatrix(lngCount, mCol.加算值)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.加算值))
                End If
                If Trim(.TextMatrix(lngCount, mCol.换算比)) = "" Then
                    strItems = strItems & ";"
                Else
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.换算比))
                End If
            Else
                strItems = strItems & ";;;"
            End If
            '糖耐项目
            strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.糖耐量项目))
              
            If LenB(strLists) < 3900 Then
                strLists = strLists & "|" & strItems
            Else
                strListsS = strListsS & "|" & strItems
            End If
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)


'    If LenB(gstrSql) > 4000 Then
'        MsgBox "仪器项目可能太多，不能保存！", vbInformation, gstrSysName
'        Me.vfgList.SetFocus: zlEditSave = 0: Exit Function
'    End If

    Err = 0: On Error GoTo ErrHand
   
    
    If strListsS <> "" Then
         '数据保存
         '如果字符数超过4000的字符，加入到strListsS中
        If strListsS <> "" Then strListsS = Mid(strListsS, 2)
        gstrSql = "Zl_检验仪器项目_Edit(" & mLngAptId & ",'" & strLists & "',0,'" & strListsS & "')"
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Else
        '数据保存
        gstrSql = "Zl_检验仪器项目_Edit(" & mLngAptId & ",'" & strLists & "')"
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    End If

    Me.Tag = "": Call Form_Resize
    zlEditSave = mLngAptId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0
End Function

Private Sub chkKind_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkUpper_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long, blnAdd As Boolean
    Dim strIDs As String '保存已添加的项目的ID
    Dim i As Long
    
    With Me.vfgList
        Select Case Index
        Case 0         '添加
            If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwItem.SelectedItem
            '将已存在项目的ID添加到变量中
            For i = 1 To .Rows
                strIDs = strIDs & "," & .TextMatrix(i - 1, mCol.ID) & ","
            Next
            '查找变量中是否已存在该项目
            If InStr(strIDs, "," & Mid(objItem.Key, 2) & ",") Then
                MsgBox """" & objItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) & """项目已存在", vbInformation
                Exit Sub
            End If
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mCol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Rows - 1, mCol.编码) = objItem.Text
            .TextMatrix(.Rows - 1, mCol.中文名) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1)
            .TextMatrix(.Rows - 1, mCol.英文名) = objItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1)
            .TextMatrix(.Rows - 1, mCol.类型) = Left(objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1), 1)
            If objItem.Tag <> "" Then
                aryTemp = Split(objItem.Tag, "|")
                .TextMatrix(.Rows - 1, mCol.通道码) = aryTemp(0)
                .TextMatrix(.Rows - 1, mCol.精度) = aryTemp(1)
                .TextMatrix(.Rows - 1, mCol.加算值) = aryTemp(2)
                .TextMatrix(.Rows - 1, mCol.换算比) = aryTemp(3)
            End If
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwItem.ListItems.Remove objItem.Key: Me.lvwItem.SetFocus
        Case 1          '删除
            If .Row < .FixedRows Then Exit Sub
            '--  10802 减少项目时，如查找出的项目列表中存在要减少的项目时报项目在集合中不唯一错
            '    检查仪器项目列表中是否已存在此项目,有则不加入选择列表
            blnAdd = True
            If lvwItem.ListItems.Count > 1 Then
                For Each objItem In lvwItem.ListItems
                    If Val(Mid(objItem.Key, 2)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        blnAdd = False
                        Exit For
                    End If
                Next
            End If

            If blnAdd Then
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.ID), .TextMatrix(.Row, mCol.编码))
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) = .TextMatrix(.Row, mCol.中文名)
                objItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1) = .TextMatrix(.Row, mCol.英文名)
                Select Case Val(.TextMatrix(.Row, mCol.类型))
                Case 1: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "1-定量"
                Case 2: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "2-定性"
                Case 3: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "3-半定量"
                End Select
                objItem.Tag = .TextMatrix(.Row, mCol.通道码) & "|" & .TextMatrix(.Row, mCol.精度)
                objItem.Tag = objItem.Tag & "|" & .TextMatrix(.Row, mCol.加算值) & "|" & .TextMatrix(.Row, mCol.换算比)
                
                objItem.Selected = True
            End If
            .RemoveItem .Row
        End Select
        
        For lngCount = .Row To .Rows - 1
            .TextMatrix(lngCount, mCol.序号) = lngCount
        Next
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String, strKind As String
    
    If Me.chkKind.Value = vbChecked Then
        strKind = "And I.操作类型 = '" & mstr类型 & "'"
    Else
        strKind = ""
    End If
    
    If Me.chkUpper.Value = 0 Then
        strFind = DelInvalidChar(Trim(UCase(Me.txtFind.Text)))
        gstrSql = "Select L.诊治项目id As ID, I.编码, I.名称 As 中文名, L.缩写 As 英文名, L.结果类型 As 类型" & vbNewLine & _
                "From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L" & vbNewLine & _
                "Where I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 <> 1 " & strKind & " And" & vbNewLine & _
                "   (I.撤档时间>sysdate or I.撤档时间 is null) And    (I.编码 Like '" & strFind & "%' Or Upper(I.名称) Like '" & gstrMatch & strFind & "%' Or Upper(L.缩写) Like '" & gstrMatch & strFind & "%')"
    Else
        strFind = DelInvalidChar(Trim(Me.txtFind.Text))
        gstrSql = "Select L.诊治项目id As ID, I.编码, I.名称 As 中文名, L.缩写 As 英文名, L.结果类型 As 类型" & vbNewLine & _
                "From 诊疗项目目录 I, 检验报告项目 R, 检验项目 L" & vbNewLine & _
                "Where I.ID = R.诊疗项目id And R.报告项目id = L.诊治项目id And I.组合项目 <> 1 " & strKind & " And" & vbNewLine & _
                "   (I.撤档时间>sysdate or I.撤档时间 is null) And     (I.编码 Like '" & strFind & "%' Or I.名称 Like '" & gstrMatch & strFind & "%' Or L.缩写 Like '" & gstrMatch & strFind & "%')"
    End If
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF

            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_中文名").Index - 1) = "" & !中文名
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_英文名").Index - 1) = "" & !英文名
            Select Case Val("" & !类型)
            Case 1: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "1-定量"
            Case 2: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "2-定性"
            Case 3: objItem.SubItems(Me.lvwItem.ColumnHeaders("_类型").Index - 1) = "3-半定量"
            End Select
            objItem.Tag = ""

            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
'    With Me.vfgList
'        For lngCount = .FixedRows To .Rows - 1
'            Me.lvwItem.ListItems.Remove "_" & .TextMatrix(lngCount, mcol.ID)
'        Next
'    End With
    
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "没有匹配的项目！", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgList.SetFocus
    End If
    Exit Sub

ErrHand:
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Resume
End Sub

Private Sub Form_Load()
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 900
        .Add , "_中文名", "中文名", 2300
        .Add , "_英文名", "英文名", 1500
        .Add , "_类型", "类型", 1000
    End With
    With Me.lvwItem
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgList.ZOrder 0
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 105
    If Me.Tag = "编辑" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 105
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItem
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwItem_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.lvwItem.Height = Me.picEdit.ScaleHeight - Me.lvwItem.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag <> "编辑" Then Exit Sub
    With Me.vfgList
        If .TextMatrix(.Row, mCol.通道码) = "" Then
            Call cmdEdit_Click(1)
        Else
            If MsgBox("该行已输入通道码，请确认是否要删除该行？", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call cmdEdit_Click(1)
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        Select Case Col
        Case mCol.通道码
            If InStr(1, "|;'", Chr(KeyAscii)) = 0 Then Exit Sub
        Case mCol.精度
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        Case mCol.加算值, mCol.换算比
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or Chr(KeyAscii) = "." Then Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.ID, mCol.序号, mCol.编码, mCol.中文名, mCol.英文名, mCol.类型: Cancel = True
    End Select
    If Row < Me.vfgList.FixedRows Then Cancel = True
End Sub



