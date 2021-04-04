VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmLabsUsage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "标本采集方式"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmLabsUsage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -15
      TabIndex        =   10
      Top             =   2745
      Width           =   6900
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -15
      TabIndex        =   9
      Top             =   585
      Width           =   6900
   End
   Begin ZL9BillEdit.BillEdit msfUsage 
      Height          =   1530
      Left            =   225
      TabIndex        =   4
      Top             =   1050
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   2699
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4335
      TabIndex        =   5
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   225
      Picture         =   "frmLabsUsage.frx":058A
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3585
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5445
      TabIndex        =   6
      Top             =   3585
      Width           =   1100
   End
   Begin VB.OptionButton optScope 
      Caption         =   "用于本项目"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Value           =   -1  'True
      Width           =   5610
   End
   Begin VB.OptionButton optScope 
      Caption         =   "用于本类别的项目"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   3195
      Width           =   5610
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5625
      Top             =   4830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabsUsage.frx":06D4
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabsUsage.frx":0C6E
            Key             =   "Method"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   240
      TabIndex        =   8
      Top             =   4785
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblUsage 
      AutoSize        =   -1  'True
      Caption         =   "常用采集方式(&U)"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   825
      Width           =   1350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1590
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "当前检验项目："
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    本程序可以指定检验项目常用的标本采集方式，目的在于辅助医生更加快速准确地开具检验申请。"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmLabsUsage.frx":1208
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmLabsUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序通过ShowMe函数传入
'   2、指定项目：由me.lblItem.tag保存，由上级程序通过ShowMe函数传入，可以传递，也可以不传递
'---------------------------------------------------
Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Public Sub ShowME(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng项目id As Long)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.cmdClose.Tag = IIf(blnEdit, "修改", "查阅")
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfUsage.Active = False
        Me.cmdSave.Visible = False
'        Me.cmdClear.Visible = False
'        Me.cmdRestore.Visible = False
    Else
        Me.msfUsage.Active = True
    End If
    Me.lblItem.Tag = lng项目id
    
    Err = 0: On Error GoTo ErrHand

    gstrSql = "Select I.ID,I.编码,I.名称,I.计算单位,I.分类id,nvl(I.服务对象,0) As 服务对象,K.编码 as 类别码,K.名称 as 类别名" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.id=[1] and I.类别=K.编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblItem.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then Unload Me: Exit Sub
        Me.lblItem.Tag = !ID
        Me.lblInfo.Caption = !名称
        Me.optScope(0).Tag = !ID: Me.optScope(0).Caption = "&1、应用于本项目(" & !编码 & "-" & !名称 & ")"
        Me.optScope(1).Tag = !类别码: Me.optScope(1).Caption = "&2、应用于所有“" & !类别名 & "”类项目"
        
        gstrSql = "select ID,编码,名称" & _
                " from 诊疗分类目录" & _
                " start with id=" & !分类ID & _
                " connect by prior 上级id=id" & _
                " order by level"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "ShowME")

        Do While Not rsTemp.EOF
            Load Me.optScope(rsTemp.AbsolutePosition + 1)
            Me.optScope(rsTemp.AbsolutePosition + 1).Tag = rsTemp!ID
            Me.optScope(rsTemp.AbsolutePosition + 1).Caption = "&" & rsTemp.AbsolutePosition + 2 & "、应用于“[" & rsTemp!编码 & "]" & rsTemp!名称 & "”类项目"
            Me.optScope(rsTemp.AbsolutePosition + 1).Left = Me.optScope(0).Left
            Me.optScope(rsTemp.AbsolutePosition + 1).Top = Me.optScope(rsTemp.AbsolutePosition).Top + Me.optScope(1).Top - Me.optScope(0).Top
            Me.optScope(rsTemp.AbsolutePosition + 1).Visible = True
            rsTemp.MoveNext
        Loop
        Me.optScope(0).Value = True
        
        Me.cmdHelp.Top = Me.optScope(Me.optScope.UBound).Top + 300
        Me.cmdSave.Top = Me.cmdHelp.Top: Me.cmdClose.Top = Me.cmdHelp.Top
        Me.Height = Me.cmdHelp.Top + Me.cmdHelp.Height + 500
    
        Call zlUsageRef(lng项目id)
    End With
    
    Me.Show 1, frmParent
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    strTemp = "": gstrSql = ""
    With Me.msfUsage
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行项目设置了重复的采集方式！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount) & "^^^^" & Trim(.TextMatrix(intCount, 2)) & "^"
            End If
        Next
    End With
    If gstrSql <> "" Then
        gstrSql = "'" & Mid(gstrSql, 2) & "'"
    Else
        gstrSql = "''"
    End If
    If Me.optScope(0).Value = True Then
        gstrSql = gstrSql & ",0,'" & Me.optScope(0).Tag & "'"
    ElseIf Me.optScope(1).Value = True Then
        gstrSql = gstrSql & ",1,'" & Me.optScope(1).Tag & "'"
    Else
        For i = 2 To Me.optScope.count - 1
            If Me.optScope(i).Value = True Then
                gstrSql = gstrSql & ",2,'" & Me.optScope(i).Tag & "'"
                Exit For
            End If
        Next
    End If
    gstrSql = "zl_用法用量_UPDATE(" & Val(Me.lblItem.Tag) & "," & _
            "'',0,0," & gstrSql & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
'    MsgBox Me.lblInfo.Caption & " 标本采集方式保存成功！", vbExclamation, gstrSysName
'    Me.msfUsage.SetFocus
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        Me.msfUsage.SetFocus
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfUsage
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 3
        
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "采集方式": .TextMatrix(0, 2) = "医生嘱托"
        
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 2200: .ColWidth(2) = 3350
        
        .ColAlignment(0) = 1: .ColAlignment(1) = 1: .ColAlignment(2) = 1
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Select Case .Tag
        Case "采集"
            Me.msfUsage.Text = .SelectedItem.Text
            Me.msfUsage.RowData(Me.msfUsage.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfUsage.TextMatrix(Me.msfUsage.Row, 1) = Me.msfUsage.Text
            Me.msfUsage.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End Select
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfUsage_AfterAddRow(Row As Long)
    With Me.msfUsage
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfUsage_AfterDeleteRow()
    With Me.msfUsage
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfUsage_CommandClick()
    If Me.msfUsage.Col = 1 Then
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 2000
            .Add , "编码", "编码", 900
        End With
        
        Err = 0: On Error GoTo ErrHand
        gstrSql = "select I.ID,I.编码,I.名称" & _
                " from 诊疗项目目录 I" & _
                " where I.类别='E' and I.操作类型='6'" & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "msfUsage_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF = 1 Then
                MsgBox "请建立标本采集项目后进行！", vbExclamation, gstrSysName: Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                objItem.Icon = "Method": objItem.SmallIcon = "Method"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
            .Tag = "采集"
            .Left = Me.msfUsage.Left + 250
            .Top = Me.msfUsage.Top
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfUsage_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfUsage.TextMatrix(Row, Col)
End Sub

Private Sub msfUsage_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfUsage_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfUsage
        If .Active = False Then Exit Sub
        Select Case .Col
        Case 2
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = Space(1)
            Else
                If Trim(.Text) = "" Then .Text = Space(1): .TextMatrix(.Row, .Col) = Space(1)
            End If
        End Select
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strInputed = strTemp Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfUsage.Col = 1 Then
        Me.lvwItems.ListItems.Clear
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 2000
            .Add , "编码", "编码", 900
        End With
        
        Err = 0: On Error GoTo ErrHand
        
        gstrSql = "select distinct I.ID,I.编码,I.名称" & _
                " from 诊疗项目目录 I,诊疗项目别名 N" & _
                " where I.ID=N.诊疗项目id and I.类别='E' and I.操作类型='6'" & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
        With rsTemp
            If .BOF Or .EOF = 1 Then
                MsgBox "未找到指定的采集方式，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            If .RecordCount = 1 Then
                Me.msfUsage.Text = !名称
                Me.msfUsage.TextMatrix(Me.msfUsage.Row, 1) = Me.msfUsage.Text
                Me.msfUsage.RowData(Me.msfUsage.Row) = !ID
                Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
                objItem.Icon = "Method": objItem.SmallIcon = "Method"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
            .Tag = "采集"
            .Left = Me.msfUsage.Left + 250
            .Top = Me.msfUsage.Top
            .Width = 3600
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlUsageRef(lngItemID As Long)
    '--------------------------------------------------------
    '功能：刷新显示药品用法用量
    '入参：lngItemId-指定的诊疗项目id（此处为成药）
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.名称 as 名称,R.医生嘱托 " & _
            " from 诊疗用法用量 R,诊疗项目目录 I" & _
            " where R.用法ID=I.ID and R.项目ID=[1] " & _
            " order by R.性质"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    With rsTemp
        Me.msfUsage.ClearBill
        Do While Not .EOF
            If Me.msfUsage.Rows - 1 < .AbsolutePosition Then Me.msfUsage.Rows = Me.msfUsage.Rows + 1
            Me.msfUsage.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfUsage.RowData(.AbsolutePosition) = !ID
            Me.msfUsage.TextMatrix(.AbsolutePosition, 1) = !名称
            Me.msfUsage.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!医生嘱托), "", !医生嘱托)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

