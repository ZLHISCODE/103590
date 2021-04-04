VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmServiceSectOffice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "存储库房设置"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmServiceSectOffice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDrug 
      Height          =   3300
      Left            =   2160
      ScaleHeight     =   3240
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2790
         Left            =   45
         TabIndex        =   19
         Top             =   405
         Visible         =   0   'False
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   4921
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
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "全选"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   83
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   21
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   50
         TabIndex        =   20
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   7080
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   548
      Width           =   285
   End
   Begin VB.Frame frame 
      Caption         =   "应用于同院区(&B)"
      Height          =   1245
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   8295
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本分类的所有药品(&6)"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   16
         Top             =   960
         Width           =   3165
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于同级的所有药品(&5)"
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   15
         Top             =   930
         Width           =   3255
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本品种下所有药品(&2)"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   3555
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“片剂”类药品(&4)"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   7
         Top             =   600
         Width           =   4545
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“西成药”(&3)"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   600
         Width           =   3345
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "仅应用于本规格药品(&1)"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2595
      Picture         =   "frmServiceSectOffice.frx":000C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1305
      Picture         =   "frmServiceSectOffice.frx":0156
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1290
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   2100
      MaxLength       =   50
      TabIndex        =   2
      Top             =   540
      Width           =   4980
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   6210
      TabIndex        =   8
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      Picture         =   "frmServiceSectOffice.frx":02A0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7320
      TabIndex        =   9
      Top             =   5625
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2925
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":03EA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":0984
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSectOffice.frx":0F1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfServiceSectOffice 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmServiceSectOffice.frx":2C28
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "规格：      生产商：       单位：瓶"
      Height          =   180
      Left            =   2130
      TabIndex        =   14
      Top             =   930
      Width           =   3150
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定药品(&M)"
      Height          =   180
      Left            =   1095
      TabIndex        =   1
      Top             =   600
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择规格药品后，设置该药品的存储库房以及药房的服务科室。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   750
      TabIndex        =   0
      Top             =   240
      Width           =   5580
   End
End
Attribute VB_Name = "frmServiceSectOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln编辑 As Boolean
Private mstr剂型 As String
Private mlng药品ID As Long                      '药品ID
Private mint材质分类 As Integer                 '5-西成药;6-中成药;7-中草药
Dim objItem As ListItem
Dim rsTemp As New ADODB.Recordset
Private mstrPrivs As String
Private mstrPara  As String         '当没有"所有库房"权限时记录当前药品其他存储库房情况
Private bln无药库药房性质部门 As Boolean
Private mstr其他库房ID As String
Private mstr全部库房ID As String
Private mstrStationNo As String
Private mrs科室 As ADODB.Recordset
Private mstr服务对象 As String
Private mintRow As Integer      '记录当前行
Private mintFind As Integer '用来记录查询到哪个位置了

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With lvwItems
        For i = 1 To .ListItems.Count
            If chkAllSelect.Value = 1 Then
                .ListItems(i).Checked = True
            Else
                .ListItems(i).Checked = False
            End If
        Next
    End With
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long, lngRows As Long
    With msfServiceSectOffice
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = ""
            .TextMatrix(lngRow, 3) = ""
            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMedi_Click()
    err = 0: On Error GoTo ErrHand
    Call AddColumnHeader
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mint材质分类)
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "尚未建立该类具体规格的药品！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If mint材质分类 <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "生产商：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.picDrug
        lblFind.Visible = False
        txtFind.Visible = False
        chkAllSelect.Visible = False
        lvwItems.Move 50, 0, .Width - 100, .Height
        
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        lvwItems.Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub cmdSave_Click()
    Dim strPara As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHand
    If mlng药品ID = 0 Then
        MsgBox "请先选择药品！", vbInformation, gstrSysName
        txtMedi.SetFocus
        Exit Sub
    End If
    If msfServiceSectOffice.Active = False Then
        MsgBox "没有找到任何药品库房，请在部门管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If opt应用于(0).Value = False Then
        For i = 1 To opt应用于.UBound
            If opt应用于(i).Value = True Then
                If MsgBox("该药品设置的存储库房应用范围为“" & opt应用于(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    '产生输入串并保存
    lngRows = msfServiceSectOffice.Rows - 1
    For lngRow = 1 To lngRows
        If msfServiceSectOffice.TextMatrix(lngRow, 1) <> "" Then
            strPara = strPara & "!!" & msfServiceSectOffice.RowData(lngRow)
            strPara = strPara & "|" & msfServiceSectOffice.TextMatrix(lngRow, 4)
        End If
    Next
    If strPara <> "" Then
        strPara = Mid(strPara, 3)
        
        If mstrPara <> "" Then
            strPara = strPara & mstrPara
        End If
    ElseIf mstrPara <> "" Then
        strPara = Mid(mstrPara, 3)
    End If
        
    gstrSql = "zl_药品存储库房_UPDATE(" & mlng药品ID & ",'" & strPara & "'"
    If opt应用于(0).Value Then
        gstrSql = gstrSql & ",1)"
    ElseIf opt应用于(1).Value Then
        gstrSql = gstrSql & ",2)"
    ElseIf opt应用于(2).Value Then
        gstrSql = gstrSql & ",3)"
    ElseIf opt应用于(3).Value Then
        gstrSql = gstrSql & ",4)"
    ElseIf opt应用于(4).Value Then
        gstrSql = gstrSql & ",5)"
    Else
        gstrSql = gstrSql & ",6)"
    End If
    Call zldatabase.ExecuteProcedure(gstrSql, "保存药品存储库房和服务科室")
    
    MsgBox "该药品的存储库房和服务科室保存成功！", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If picDrug.Visible = True Then
            picDrug.Visible = False
            With msfServiceSectOffice
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strRange As String
    Dim n As Integer
    
    Call InitFace
    Call ShowData
    
    mintFind = 1
    If bln无药库药房性质部门 = True Then
        MsgBox "请先设置具有药库药房性质的部门。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If mbln编辑 = False Then
        msfServiceSectOffice.Active = False
        frame.Enabled = False
        cmdSave.Visible = False
        cmdRestore.Visible = False
        cmdClear.Visible = False
        cmdClose.Caption = "退出(&X)"
    End If
    
    strRange = zldatabase.GetPara("应用范围", glngSys, 1023, False)
    For n = 0 To opt应用于.Count - 1
        opt应用于(n).Enabled = Mid(strRange, n + 1, 1)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs科室 = Nothing
    mintRow = 0
End Sub

Private Sub lvwItems_GotFocus()
    Dim j As Integer
    
    With lvwItems
        For j = 1 To .ListItems.Count
            .ListItems(j).ForeColor = vbBlack
        Next
    End With
End Sub

Private Sub msfServiceSectOffice_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    MsgBox "msfServiceSectOffice_BeforeDeleteRow"
    msfServiceSectOffice.TextMatrix(Row, 1) = ""
    msfServiceSectOffice.TextMatrix(Row, 3) = ""
    msfServiceSectOffice.TextMatrix(Row, 4) = ""
    Cancel = True
End Sub

Private Sub msfServiceSectOffice_CommandClick()
'    Dim str服务对象 As String
'    Dim objItem As ListItem
    Dim bln科室 As Boolean
    
    bln科室 = Check服务科室
    If bln科室 = True Then Exit Sub
    mintRow = msfServiceSectOffice.Row
    
    Call frmServiceSelect.ShowMe(frmServiceSectOffice, mintRow, mstr服务对象, 1)
End Sub

Private Function Check服务科室() As Boolean
    '功能：检查当前库房是不是药房或者是否设置临床科室
    '返回值 true 当前库房不是药房也没有设置临床科室,false 当前库房是药房或者或者设置了临床科室
    Dim str服务对象 As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    cmdSave.Enabled = True
    str服务对象 = ""
    gstrSql = "select distinct 服务对象 from 部门性质说明 where 部门ID=[1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取服务对象", msfServiceSectOffice.RowData(msfServiceSectOffice.Row))
    
    Do While Not rsTemp.EOF
        str服务对象 = str服务对象 & "," & rsTemp!服务对象
        rsTemp.MoveNext
    Loop
    If str服务对象 <> "" Then
        str服务对象 = Mid(str服务对象, 2)
        If InStr(1, str服务对象, 3) <> 0 Then
            str服务对象 = "0,1,2,3"
        ElseIf InStr(1, str服务对象, 1) <> 0 Or InStr(1, str服务对象, 2) <> 0 Then
            str服务对象 = str服务对象 & ",3"
        End If
    Else
        str服务对象 = "0"
    End If

    mstr服务对象 = str服务对象
    gstrSql = " Select ID,编码,名称,简码 From 部门表 A,部门性质说明 B " & _
              " Where A.ID=B.部门ID And B.工作性质 Like '临床%'" & _
              " And Instr([1], ',' || B.服务对象 || ',') > 0"
    gstrSql = gstrSql & " and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) "
    
    If mstrStationNo <> "" Then
            gstrSql = gstrSql & " And (A.站点 = '" & mstrStationNo & "' Or A.站点 is Null) "
        End If
    Set mrs科室 = zldatabase.OpenSQLRecord(gstrSql, "提取服务科室", "," & str服务对象 & ",")
    
    If mrs科室.RecordCount = 0 Then
        MsgBox "当前库房不是药房或者未设置临床科室！[部门管理]", vbInformation, gstrSysName
        msfServiceSectOffice.Text = ""
        msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Row, 4) = ""
        Check服务科室 = True
        Exit Function
    End If
    Check服务科室 = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub msfServiceSectOffice_EnterCell(Row As Long, Col As Long)
    If Col = 3 Then
        msfServiceSectOffice.TxtEnable = True
    End If
End Sub

Private Sub msfServiceSectOffice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim bln科室 As Boolean
    Dim objItem As ListItem
    Dim rsRecord As ADODB.Recordset
    Dim strKey As String
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        With msfServiceSectOffice
            If .Col = 3 Then
                strKey = Trim(UCase(.Text))

                If strKey = "" Then Exit Sub
                Debug.Print strKey
                mintRow = .Row
                
                bln科室 = Check服务科室
                If bln科室 = True Then
                    .TextMatrix(.Row, 3) = ""
                    .TextMatrix(.Row, 4) = ""
                Else
                    gstrSql = " Select distinct A.ID,A.编码,A.名称,A.简码 From 部门表 A,部门性质说明 B,部门性质分类 C " & _
                        " Where A.ID=B.部门ID And B.工作性质=C.名称 And Instr('3ABCDEF', C.编码) > 0 " & _
                        " And Instr([1], ',' || B.服务对象 || ',') > 0"
                        
                    gstrSql = gstrSql & " and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) and ( a.编码 like [2] or a.名称 like [2] or a.简码 like [2]) "
                    Set rsRecord = zldatabase.OpenSQLRecord(gstrSql, "科室", "," & mstr服务对象 & ",", strKey & "%")
                    
                    If rsRecord.RecordCount > 1 Then
                        mintRow = .Row
                        Call frmServiceSelect.ShowMe(frmServiceSectOffice, mintRow, mstr服务对象, 1, strKey)
                    ElseIf rsRecord.RecordCount = 1 Then
                        .Text = IIf(IsNull(rsRecord!名称), "", rsRecord!名称)
                        .TextMatrix(msfServiceSectOffice.Row, 4) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                        .TextMatrix(msfServiceSectOffice.Row, 3) = .Text
                        .SetFocus
                    ElseIf rsRecord.RecordCount = 0 Then
                        MsgBox "没有找到相应的部门！", vbInformation, gstrSysName
                        .Text = ""
                        .TextMatrix(msfServiceSectOffice.Row, 3) = msfServiceSectOffice.Text
                        .TextMatrix(msfServiceSectOffice.Row, 4) = ""
                        .TxtSetFocus
                        Cancel = True
                    End If
                End If
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
                    
Private Sub opt应用于_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To opt应用于.UBound
        If i = Index Then
            opt应用于(i).FontBold = True
        Else
            opt应用于(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim i As Integer
    Dim blnResult As Boolean
    Dim j As Integer
    Dim k As Integer
    
    blnResult = False
    With lvwItems
        If KeyCode = vbKeyReturn And Trim(txtFind.Text) <> "" Then
            strFind = UCase(Trim(txtFind.Text))
            If mintFind > .ListItems.Count Then
                mintFind = 1
            Else
                mintFind = mintFind + 1
                If mintFind > .ListItems.Count Then
                    mintFind = 1
                End If
            End If
            
            For i = mintFind To .ListItems.Count
                If IsNumeric(strFind) Then
                    If .ListItems(i).ListSubItems(.ColumnHeaders("编码").Index - 1).Text = strFind Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        
                        mintFind = i
                        Exit Sub
                    End If
                    
                    If i = .ListItems.Count Then
                        For k = 1 To mintFind
                            If .ListItems(k).ListSubItems(.ColumnHeaders("编码").Index - 1).Text = strFind Then
                                .ListItems(k).EnsureVisible
                                For j = 1 To .ListItems.Count
                                    .ListItems(j).ForeColor = vbBlack
                                Next
                                .ListItems(k).ForeColor = vbBlue
                                
                                mintFind = k
                                Exit Sub
                            End If
                        Next
                    End If
                Else
                    If .ListItems(i).ListSubItems(.ColumnHeaders("简码").Index - 1).Text Like "*" & strFind & "*" Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        mintFind = i
                        blnResult = True
                        Exit Sub
                    End If
                    
                    If i = .ListItems.Count Then
                        For k = 1 To mintFind
                            If .ListItems(k).ListSubItems(.ColumnHeaders("简码").Index - 1).Text Like "*" & strFind & "*" Then
                                .ListItems(k).EnsureVisible
                                For j = 1 To .ListItems.Count
                                    .ListItems(j).ForeColor = vbBlack
                                Next
                                .ListItems(k).ForeColor = vbBlue
                                
                                mintFind = k
                                blnResult = True
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            Next
            
            If blnResult = False Then
                For i = mintFind To .ListItems.Count
                    If .ListItems(i).Text Like "*" & strFind & "*" Then
                        .ListItems(i).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(i).ForeColor = vbBlue
                        mintFind = i
                        blnResult = True
                        Exit Sub
                    End If
                Next
                
                For k = 1 To mintFind
                    If .ListItems(k).Text Like "*" & strFind & "*" Then
                        .ListItems(k).EnsureVisible
                        For j = 1 To .ListItems.Count
                            .ListItems(j).ForeColor = vbBlack
                        Next
                        .ListItems(k).ForeColor = vbBlue
                        
                        mintFind = k
                        blnResult = True
                        Exit Sub
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N" & _
            " where I.ID=N.收费细目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mint材质分类, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定规格的药品，请重新指定！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If mint材质分类 <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "生产商：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Call AddColumnHeader
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDrug
        lblFind.Visible = False
        txtFind.Visible = False
        chkAllSelect.Visible = False
        lvwItems.Move 50, 0, .Width - 100, .Height
        
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        lvwItems.Visible = True
        .SetFocus
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub ShowData()
    '提取数据并显示出来
    Const str西药 As String = "'西药%'"
    Const str中药 As String = "'中药%'"
    Const str成药 As String = "'成药%'"
    Dim str库房ID As String, str科室 As String, str科室ID As String
    Dim intRow As Integer, intRows As Integer
    Dim blnSel As Boolean
    Dim lng诊疗项目ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsOther As New ADODB.Recordset
    Dim dbl所有库房 As Boolean
    Dim strTmp As String
    Dim strIdArr
    
    err = 0: On Error GoTo ErrHand
    
    mstrPara = ""
    If InStr(1, ";" & mstrPrivs & ";", ";所有库房;") > 0 Then dbl所有库房 = True
    
    Call cmdClear_Click
    If mblnFirst Then
        '提取药品信息
        If mlng药品ID <> 0 Then
            gstrSql = " Select A.药名ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位,B.药品剂型 剂型 " & _
                      " From 药品规格 A,药品特性 B,收费项目目录 I " & _
                      " Where A.药名ID=B.药名ID And A.药品ID=I.ID And A.药品ID=[1] "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取药品信息", mlng药品ID)
            
            lng诊疗项目ID = rsTemp!药名ID
            txtMedi.Tag = "[" & rsTemp!编码 & "]" & rsTemp!名称
            txtMedi.Text = txtMedi.Tag
            mstr剂型 = rsTemp!剂型
            If mint材质分类 <> "7" Then
                Me.lblSpec.Caption = "规格：" & IIf(IsNull(rsTemp!规格), "", rsTemp!规格) & _
                    "   生产商：" & IIf(IsNull(rsTemp!产地), "", rsTemp!产地) & _
                    "   单位：" & IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
            Else
                Me.lblSpec.Caption = "生产商：" & IIf(IsNull(rsTemp!产地), "", rsTemp!产地) & "   单位：" & IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
            End If
        End If
        
        '根据药品的用途分类提取所允许存储的库房
        gstrSql = " Select ID,编码,名称 From 部门表 " & _
                  " Where  (撤档时间 is null or to_char(撤档时间,'yyyy-mm-dd')='3000-01-01') and ID in (select distinct 部门id from 部门性质说明 where 工作性质 like "
        If mint材质分类 = "5" Then
            gstrSql = gstrSql & str西药
        ElseIf mint材质分类 = "6" Then
            gstrSql = gstrSql & str成药
        Else
            gstrSql = gstrSql & str中药
        End If
        gstrSql = gstrSql & " or 工作性质='制剂室')"
        
        gstrSql = gstrSql & " and (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) "
        
        If mstrStationNo <> "" Then
            gstrSql = gstrSql & " And (站点 = '" & mstrStationNo & "' Or 站点 is Null) "
        End If
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房(其他库房)")
        mstr全部库房ID = ""
        Do While Not rsTemp.EOF
            mstr全部库房ID = mstr全部库房ID & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
        If mstr全部库房ID <> "" Then
            mstr全部库房ID = Mid(mstr全部库房ID, 2)
            bln无药库药房性质部门 = False
        Else
            bln无药库药房性质部门 = True
            Exit Sub
        End If
        
        strTmp = gstrSql
        If Not dbl所有库房 Then
            '先取其他库房
            gstrSql = strTmp & " And Id not In(Select 部门ID From 部门人员 Where 人员id=[1])"
            Set rsOther = zldatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房(其他库房)", UserInfo.ID)
            
            mstr其他库房ID = ""
            Do While Not rsOther.EOF
                mstr其他库房ID = mstr其他库房ID & "," & rsOther!ID
                rsOther.MoveNext
            Loop
            If mstr其他库房ID <> "" Then
                mstr其他库房ID = Mid(mstr其他库房ID, 2)
            End If
                        
            '取当前用户所属库房
            gstrSql = strTmp & " And Id In(Select 部门ID From 部门人员 Where 人员id=[1])"
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "根据药品的用途分类提取所允许存储的库房", UserInfo.ID)
                
        Do While Not rsTemp.EOF
            msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Rows - 1, 2) = rsTemp!名称
            msfServiceSectOffice.RowData(msfServiceSectOffice.Rows - 1) = rsTemp!ID
            msfServiceSectOffice.Rows = msfServiceSectOffice.Rows + 1
            str库房ID = str库房ID & "," & rsTemp!ID
            rsTemp.MoveNext
        Loop
        If str库房ID <> "" Then
            str库房ID = Mid(str库房ID, 2)
            msfServiceSectOffice.Rows = msfServiceSectOffice.Rows - 1
            msfServiceSectOffice.Active = True
        Else
            msfServiceSectOffice.Active = False
        End If
    End If
    
    '取所有库房
    str库房ID = ""
    intRows = msfServiceSectOffice.Rows - 1
    For intRow = 1 To intRows
        str库房ID = str库房ID & "," & msfServiceSectOffice.RowData(intRow)
    Next
    If str库房ID <> "" Then str库房ID = Mid(str库房ID, 2)
    
   '将相应数据组织后装入单据控件
    gstrSql = " Select A.收费细目ID,A.开单科室ID,A.执行科室ID,B.名称 From 收费执行科室 A,部门表 B " & _
              " Where A.开单科室ID=B.ID(+) And A.收费细目ID=[1] And Instr([2],','||A.执行科室ID||',') > 0 " & _
              " Order by A.执行科室ID"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取已设置的收费执行科室数据", mlng药品ID, "," & mstr全部库房ID & ",")
    
    If rsTemp.RecordCount = 0 And mblnFirst And mlng药品ID <> 0 Then
        '提取同品种下其它规格药品的数据做为缺省数据
        gstrSql = " Select A.开单科室ID,A.执行科室ID,B.名称 From 收费执行科室 A,部门表 B," & _
                  "     (Select 药品ID From 药品规格 Where 药名ID=[1] And Rownum<2) C" & _
                  " Where A.开单科室ID=B.ID(+) And A.收费细目ID=C.药品ID And Instr([2],','||A.执行科室ID||',') > 0 " & _
                  " Order by A.执行科室ID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取相同材质分类下相同剂型的药品的数据做为缺省数据", lng诊疗项目ID, "," & mstr全部库房ID & ",")
                
        If rsTemp.RecordCount = 0 Then
            '提取相同材质分类下相同剂型的药品的数据做为缺省数据
            gstrSql = " Select A.开单科室ID,A.执行科室ID,B.名称 From 收费执行科室 A,部门表 B," & _
                      "     (Select C.药品ID From 诊疗项目目录 A,药品特性 B,药品规格 C,收费执行科室 D " & _
                      "     Where A.ID=B.药名ID And B.药名ID=C.药名ID And C.药品ID=D.收费细目ID And B.药品剂型=[1] And A.类别=[2] And Rownum<2) C" & _
                      " Where A.开单科室ID=B.ID(+) And A.收费细目ID=C.药品ID And Instr([3],','||A.执行科室ID||',') > 0 " & _
                      " Order by A.执行科室ID"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取相同材质分类下相同剂型的药品的数据做为缺省数据", mstr剂型, mint材质分类, "," & mstr全部库房ID & ",")
            
            If rsTemp.RecordCount <> 0 Then
                MsgBox "当前规格药品未设置存储库房，提取剂型相同的规格药品的存储库房做为缺省数据！", vbInformation, gstrSysName
            End If
        Else
            MsgBox "当前规格药品未设置存储库房，提取同品种下规格药品的存储库房做为缺省数据！", vbInformation, gstrSysName
        End If
    End If
    For intRow = 1 To intRows
        str科室 = "": str科室ID = ""
        rsTemp.Filter = "执行科室ID=" & msfServiceSectOffice.RowData(intRow)
        
        blnSel = False
        Do While Not rsTemp.EOF
            blnSel = True
            str科室 = str科室 & "," & Nvl(rsTemp!名称)
            str科室ID = str科室ID & "," & Nvl(rsTemp!开单科室ID, 0)
            rsTemp.MoveNext
        Loop
        If str科室 <> "" Then
            str科室 = Mid(str科室, 2)
            str科室ID = Mid(str科室ID, 2)
            If str科室ID = "0" Then str科室ID = ""
        End If
        msfServiceSectOffice.TextMatrix(intRow, 0) = intRow
        If blnSel Then msfServiceSectOffice.TextMatrix(intRow, 1) = "√"
        msfServiceSectOffice.TextMatrix(intRow, 3) = str科室
        msfServiceSectOffice.TextMatrix(intRow, 4) = str科室ID
    Next
    
    '取其他执行科室
    If Not dbl所有库房 And mstr其他库房ID <> "" Then
        gstrSql = " Select DISTINCT 开单科室ID,执行科室ID From 收费执行科室 " & _
              " Where 收费细目ID=[1] And Instr([2],','||执行科室ID||',') > 0 " & _
              " Order by 执行科室ID"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "提取已设置的收费执行科室数据", mlng药品ID, "," & mstr其他库房ID & ",")
                        
        strIdArr = Split(mstr其他库房ID, ",")
        For intRow = 0 To UBound(strIdArr)
            str科室ID = ""
            rsTemp.Filter = "执行科室ID=" & strIdArr(intRow)
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str科室ID = str科室ID & "," & Nvl(rsTemp!开单科室ID, 0)
                rsTemp.MoveNext
            Loop
            If str科室ID <> "" Then
                str科室ID = Mid(str科室ID, 2)
                If str科室ID = "0" Then str科室ID = ""
                mstrPara = mstrPara & "!!" & CStr(strIdArr(intRow))
                mstrPara = mstrPara & "|" & str科室ID
            End If
        Next
    End If
    
    '修改应用于信息
    opt应用于(2).Caption = "应用于所有“" & Switch(mint材质分类 = 5, "西成药", mint材质分类 = 6, "中成药", mint材质分类 = 7, "中草药") & "”(&3)"
    opt应用于(3).Caption = "应用于所有“" & mstr剂型 & "”剂型类药品(&4)"
    
    If mint材质分类 = 7 Then opt应用于(3).Enabled = False
    mblnFirst = False
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '初始化控件
    With msfServiceSectOffice
        .Rows = 2
        .Cols = 5
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "选择"
        .TextMatrix(0, 2) = "存储库房"
        .TextMatrix(0, 3) = "服务科室"
        .TextMatrix(0, 4) = "服务科室ID"
        .TextMatrix(1, 0) = "1"
        .colData(0) = 5
        .colData(1) = -1
        .colData(2) = 5
        .colData(3) = 1
        .colData(4) = 5
        .ColWidth(0) = 300
        .ColWidth(1) = 500
        .ColWidth(2) = 1500
        .ColWidth(3) = 5000
        .ColWidth(4) = 0
        
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub msfServiceSectOffice_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    MsgBox "msfServiceSectOffice_AfterAddRow"
    '修改行序号
    MsgBox "msfServiceSectOffice_AfterAddRow"
    With msfServiceSectOffice
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfServiceSectOffice_AfterDeleteRow()
    Dim lngCurRow As Long
    
    MsgBox "msfServiceSectOffice_AfterDeleteRow"
    '修改行序号
    With msfServiceSectOffice
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub AddColumnHeader(Optional ByVal bln药品 As Boolean = True)
    If bln药品 Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 2000
            .Add , "编码", "编码", 1000
            .Add , "规格", "规格", 800
            .Add , "产地", "生产商", 1500
            .Add , "单位", "单位", 800
        End With
        With Me.lvwItems
            .Checkboxes = False
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "1"
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 2000
            .Add , "编码", "编码", 1000
            .Add , "简码", "简码", 1000
        End With
        With Me.lvwItems
            .Checkboxes = True
            .ColumnHeaders("编码").Position = 1
            .SortKey = .ColumnHeaders("编码").Index - 2
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "2"
    End If
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
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str科室 As String, str科室ID As String
    
    If lvwItems.Tag = "1" Then
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        With Me.lvwItems
            If mlng药品ID <> Mid(.SelectedItem.Key, 2) Then
                mlng药品ID = Mid(.SelectedItem.Key, 2)
                Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.txtMedi.Text = Me.txtMedi.Tag
                mstr剂型 = .SelectedItem.SubItems(3)
                If mint材质分类 <> "7" Then
                    Me.lblSpec.Caption = "规格：" & .SelectedItem.SubItems(2) & _
                        "   生产商：" & .SelectedItem.SubItems(3) & _
                        "   单位：" & .SelectedItem.SubItems(4)
                Else
                    Me.lblSpec.Caption = "生产商：" & .SelectedItem.SubItems(3) & "   单位：" & .SelectedItem.SubItems(4)
                End If
                Call ShowData
            End If
            Me.txtMedi.SetFocus
            Call OS.PressKey(vbKeyTab)
        End With
        picDrug.Visible = False
    Else
        '循环提取用户所选择的科室
        lngRows = lvwItems.ListItems.Count
        For lngRow = 1 To lngRows
            If lvwItems.ListItems(lngRow).Checked Then
                str科室 = str科室 & "," & lvwItems.ListItems(lngRow).Text
                str科室ID = str科室ID & "," & Mid(lvwItems.ListItems(lngRow).Key, 2)
            End If
        Next
        If str科室 <> "" Then
            str科室 = Mid(str科室, 2)
            str科室ID = Mid(str科室ID, 2)
        End If
        msfServiceSectOffice.Visible = True
        lvwItems.Visible = False
        picDrug.Visible = False
        If str科室 <> "" Then msfServiceSectOffice.TextMatrix(msfServiceSectOffice.Row, 1) = "√"
        msfServiceSectOffice.Text = str科室
        msfServiceSectOffice.TextMatrix(mintRow, 3) = msfServiceSectOffice.Text
        msfServiceSectOffice.TextMatrix(mintRow, 4) = str科室ID
        If msfServiceSectOffice.Rows - 1 > msfServiceSectOffice.Row Then msfServiceSectOffice.Row = msfServiceSectOffice.Row + 1: msfServiceSectOffice.SetFocus
    End If
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

'Private Sub lvwItems_LostFocus()
'    Me.lvwItems.Visible = False
'End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal lng药品ID As Long, ByVal int用途分类 As Integer, ByVal bln编辑 As Boolean, ByVal strStationNo As String)
    On Error Resume Next
    mblnFirst = True
    mlng药品ID = lng药品ID
    mint材质分类 = int用途分类
    mbln编辑 = bln编辑
    'mstrPrivs = gstrPrivs
    mstrPrivs = frmParent.mstrPrivs
    mstrStationNo = strStationNo
    Me.Show 1, frmParent
End Sub
