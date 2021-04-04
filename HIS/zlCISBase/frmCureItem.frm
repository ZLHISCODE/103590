VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCureItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参考目录编辑"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmCureItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   5025
      Width           =   8490
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   570
      TabIndex        =   20
      Top             =   330
      Width           =   8490
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   5175
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   285
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1815
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   0
      Top             =   450
      Width           =   3360
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Index           =   4
      Left            =   1815
      TabIndex        =   6
      Top             =   1920
      Width           =   3645
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Index           =   2
      Left            =   1815
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Index           =   1
      Left            =   1815
      TabIndex        =   3
      Top             =   1185
      Width           =   3645
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Index           =   0
      Left            =   1815
      TabIndex        =   2
      Top             =   825
      Width           =   1605
   End
   Begin VB.TextBox txtItem 
      Height          =   645
      Index           =   5
      Left            =   780
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4230
      Width           =   4695
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Index           =   3
      Left            =   3660
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4680
      TabIndex        =   10
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   9
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1100
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4380
      Left            =   -3480
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   2385
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7726
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -285
      Top             =   2835
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
            Picture         =   "frmCureItem.frx":000C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCureItem.frx":05A6
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msf别名 
      Height          =   1335
      Left            =   780
      TabIndex        =   7
      Top             =   2580
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2355
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
   Begin VB.Label lbl别名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "其他别名(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   22
      Top             =   2325
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "根据本类应用参考诊疗的特点，设置参考目录项目。"
      Height          =   180
      Left            =   780
      TabIndex        =   19
      Top             =   120
      Width           =   4140
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "参考分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   18
      Top             =   525
      Width           =   990
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "简要说明(&M)"
      Height          =   180
      Index           =   5
      Left            =   780
      TabIndex        =   17
      Top             =   4005
      Width           =   990
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "英文名称(&E)"
      Height          =   180
      Index           =   4
      Left            =   780
      TabIndex        =   16
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "名称简码(&S)              (拼音)               (五笔)"
      Height          =   180
      Index           =   2
      Left            =   780
      TabIndex        =   15
      Top             =   1635
      Width           =   4680
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "目录名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   780
      TabIndex        =   14
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "目录编码(&D)"
      Height          =   180
      Index           =   0
      Left            =   780
      TabIndex        =   13
      Top             =   885
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   135
      Picture         =   "frmCureItem.frx":0B40
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmCureItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer
Dim strTemp As String

Const con目录编码 As Integer = 0
Const con目录名称 As Integer = 1
Const con名称拼音 As Integer = 2
Const con名称五笔 As Integer = 3
Const con英文名称 As Integer = 4
Const con简要说明 As Integer = 5

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemId As Long, StrClass As String, strCollate As String
    
    If Trim(Me.txtItem(con目录编码).Text) = "" Then
        MsgBox "编码必须输入", vbExclamation, gstrSysName
        Me.txtItem(con目录编码).SetFocus
        Exit Sub
    End If
    If Trim(Me.txtItem(con目录名称).Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtItem(con目录名称).SetFocus
        Exit Sub
    End If
    For intCount = Me.txtItem.LBound To Me.txtItem.UBound
        Select Case intCount
        Case con目录名称, con简要说明
            If LenB(StrConv(Trim(Me.txtItem(intCount).Text), vbFromUnicode)) > Me.txtItem(intCount).MaxLength Then
                MsgBox Me.lblItem(intCount).Caption & "超过" & Me.txtItem(intCount).MaxLength & "的长度限制", vbExclamation, gstrSysName
                Me.txtItem(intCount).SetFocus
                Exit Sub
            End If
        End Select
    Next
    '别名检查
    strTemp = ";" & Trim(Me.txtItem(con目录名称).Text) & ";" & Trim(Me.txtItem(con英文名称).Text)
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "别名存在重复（包括目录名称和英文名称）！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    strTemp = ""
    With Me.msf别名
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    gstrSql = Val(Me.txt分类.Tag) & "," & _
        "'" & Trim(Me.txtItem(con目录编码).Text) & "'," & _
        "'" & Trim(Me.txtItem(con目录名称).Text) & "'," & _
        "'" & Trim(Me.txtItem(con名称拼音).Text) & "'," & _
        "'" & Trim(Me.txtItem(con名称五笔).Text) & "'," & _
        "'" & Trim(Me.txtItem(con英文名称).Text) & "'," & _
        "'" & Trim(Me.txtItem(con简要说明).Text) & "'"
    
    Err = 0: On Error GoTo ErrHand
    If Me.Tag = "增加" Then
        lngItemId = zlDatabase.GetNextId("诊疗参考目录")
        gstrSql = "zl_诊疗参考目录_Insert(" & lngItemId & "," & gstrSql & "," & Me.lblNote.Tag & ",'" & strTemp & "')"
    Else
        lngItemId = Me.Tag
        gstrSql = "zl_诊疗参考目录_Update(" & lngItemId & "," & gstrSql & ",'" & strTemp & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If Val(Me.lblNote.Tag) = 2 Or Val(Me.lblNote.Tag) = 3 Or Val(Me.lblNote.Tag) = 4 Then
        Me.lblItem(con英文名称).Visible = False: Me.txtItem(con英文名称).Visible = False
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    '分类选择树装入
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗参考分类" & _
            " Where 类型 = [1] " & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblNote.Tag))
    
    With rsTemp
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            If Val(Me.txt分类.Tag) = !ID Then
                objNode.Selected = True
                Me.txt分类.Text = objNode.Text
            End If
            .MoveNext
        Loop
    End With
    
    '名称等填写
    gstrSql = "select ID,编码,名称,说明" & _
            " From 诊疗参考目录" & _
            " Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
        
    Me.txtItem(con目录编码).MaxLength = rsTemp.Fields("编码").DefinedSize
    Me.txtItem(con目录名称).MaxLength = rsTemp.Fields("名称").DefinedSize
    Me.txtItem(con简要说明).MaxLength = rsTemp.Fields("说明").DefinedSize
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Me.txtItem(con目录编码).Text = rsTemp!编码
        Me.txtItem(con目录名称).Text = rsTemp!名称
        Me.txtItem(con简要说明).Text = IIf(IsNull(rsTemp!说明), "", rsTemp!说明)
    Else
        gstrSql = "select nvl(max(编码),'000000') as 编码" & _
                " From 诊疗参考目录" & _
                " Where 类型 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblNote.Tag))
        
        Me.txtItem(con目录编码).Text = Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
    End If
    
    '正名简码和英文名填写
    gstrSql = "select nvl(名称,'') as 名称, 性质, nvl(简码,'') as 简码, 码类" & _
            " From 诊疗参考别名" & _
            " Where 参考目录id=[1] And 性质 in (1,2)" & _
            " Order by 性质,码类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
    
    With rsTemp
        Me.txtItem(con英文名称).MaxLength = .Fields("名称").DefinedSize
        Me.txtItem(con名称拼音).MaxLength = .Fields("简码").DefinedSize
        Me.txtItem(con名称五笔).MaxLength = .Fields("简码").DefinedSize
        Do While Not .EOF
            If !性质 = 1 Then
                If !码类 = 2 Then
                    Me.txtItem(con名称五笔).Text = !简码
                Else
                    Me.txtItem(con名称拼音).Text = !简码
                End If
            Else
                Me.txtItem(con英文名称).Text = !名称
            End If
            .MoveNext
        Loop
    End With
    
    '其他别名填写
    gstrSql = "select N.名称,P.简码 as 拼音,W.简码 as 五笔" & _
            " from (select distinct 名称 from 诊疗参考别名 where 参考目录id=[1] and 性质=9) N," & _
            "      (select 名称,简码 from 诊疗参考别名 where 参考目录id=[1] and 性质=9 and 码类=1) P," & _
            "      (select 名称,简码 from 诊疗参考别名 where 参考目录id=[1] and 性质=9 and 码类=2) W" & _
            " where N.名称=P.名称(+) and N.名称=W.名称(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.Tag = "增加", -1, Me.Tag))
    
    With rsTemp
        Do While Not .EOF
            If Me.msf别名.Rows - 1 < .AbsolutePosition Then Me.msf别名.Rows = Me.msf别名.Rows + 1
            Me.msf别名.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msf别名.TextMatrix(.AbsolutePosition, 1) = !名称
            Me.msf别名.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!拼音), "", !拼音)
            Me.msf别名.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!五笔), "", !五笔)
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    '初始化设置表格编辑
    With Me.msf别名
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "别名": .TextMatrix(0, 2) = "拼音码": .TextMatrix(0, 3) = "五笔码"
        .ColData(0) = 5: .ColData(1) = 4: .ColData(2) = 4: .ColData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 2250: .ColWidth(2) = 950: .ColWidth(3) = 950
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False
        Cancel = True
    End If
End Sub

Private Sub lblItem_Click(Index As Integer)
    Me.txtItem(Index).SetFocus
End Sub

Private Sub msf别名_AfterAddRow(Row As Long)
    With Me.msf别名
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_AfterDeleteRow()
    With Me.msf别名
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf别名_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf别名
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            strTemp = Trim(.Text)
            If zlCommFun.StrIsValid(strTemp, 60) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If strTemp <> "" Then
                .TextMatrix(.Row, 1) = strTemp: .TextMatrix(.Row, 2) = zlStr.GetCodeByORCL(strTemp): .TextMatrix(.Row, 3) = zlStr.GetCodeByORCL(strTemp, True)
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
    End With
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Select Case Index
    Case con目录名称, con简要说明
        Call zlCommFun.OpenIme(True)
    End Select
    Me.txtItem(Index).SelStart = 0: Me.txtItem(Index).SelLength = 100
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case con目录编码
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
        KeyAscii = 0
    Case con目录名称
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case con英文名称, con名称拼音, con名称五笔
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii = vbKeySpace Then Exit Sub
        End Select
        KeyAscii = 0
    Case con简要说明
        If InStr("%_'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
End Sub

Private Sub txtItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case con目录名称
        Me.txtItem(con名称拼音).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), False)
        Me.txtItem(con名称五笔).Text = zlStr.GetCodeByORCL(Trim(Me.txtItem(Index).Text), True)
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
    Case con目录名称, con简要说明
        Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

