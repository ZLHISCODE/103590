VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediSendType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品发药类型批量设置"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   Icon            =   "frmMediSendType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7560
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   7
      ToolTipText     =   "请输入编码，名称或简码！"
      Top             =   1620
      Width           =   2265
   End
   Begin VB.CheckBox chk分类 
      Caption         =   "中草药"
      Height          =   210
      Index           =   2
      Left            =   3360
      TabIndex        =   5
      Top             =   1265
      Width           =   1035
   End
   Begin VB.CheckBox chk分类 
      Caption         =   "中成药"
      Height          =   210
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1265
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk分类 
      Caption         =   "西药"
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1265
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4485
      Left            =   120
      TabIndex        =   8
      Tag             =   "1000"
      Top             =   1995
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7911
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   6240
      TabIndex        =   10
      Top             =   6840
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4920
      TabIndex        =   9
      Top             =   6840
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   6600
      Width           =   7380
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&S"
      Height          =   300
      Left            =   3270
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
   End
   Begin VB.ComboBox cbo发药类型 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Text            =   "cbo发药类型"
      Top             =   840
      Width           =   2265
   End
   Begin VB.ComboBox cbo分类 
      Height          =   300
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   885
      Width           =   2265
   End
   Begin VB.TextBox txtInput 
      Height          =   300
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   6
      ToolTipText     =   "请输入编码，名称或简码！"
      Top             =   1620
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   7500
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6000
      Top             =   0
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
            Picture         =   "frmMediSendType.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "查找"
      Height          =   180
      Left            =   4560
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "药品材质"
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label lbl发药类型 
      AutoSize        =   -1  'True
      Caption         =   "发药类型"
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lbl分类 
      AutoSize        =   -1  'True
      Caption         =   "过滤信息"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl分类方式 
      AutoSize        =   -1  'True
      Caption         =   "分类方式"
      Height          =   180
      Left            =   4200
      TabIndex        =   13
      Top             =   945
      Width           =   720
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmMediSendType.frx":2294
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    本程序可以批量设定药品的自定义发药类型，在部门发药时通过发药类型快速进行发药操作。"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   885
      TabIndex        =   0
      Top             =   150
      Width           =   4605
   End
End
Attribute VB_Name = "frmMediSendType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFindStyle As String
Private mrsTemp As ADODB.Recordset
Private mstrFind As String          '记录上次查询的语句

Private Enum MediType
    药品分类 = 0
    药品品种
    药品规格
    药品剂型
    给药途径
End Enum

Private Sub GetSelect(ByVal intType As Integer, ByVal strInput As String, Optional BlnFind As Boolean = False)
    Dim objNode As Node
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim strSql As String
    
    On Error GoTo errHandle
    If BlnFind = False Then
        tvwClass.Nodes.Clear
        Set objNode = tvwClass.Nodes.Add(, , "Root", "所有", 1)
    End If
    Set mrsTemp = Nothing
    Select Case intType
        Case MediType.药品分类
            gstrSql = "Select Level As 层, ID, 上级id, 编码, 名称, 简码, Decode(类型, 1, '西药', Decode(类型, 2, '中成药', '中草药')) As 类型 " & _
                " From 诊疗分类目录 " & _
                " Where 撤档时间 Is Null "
            If strInput <> "" Then
                gstrSql = gstrSql & " And (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2]) "
            End If
            
            If chk分类(0).Value = 1 And chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And 类型 In (1, 2, 3) "
            ElseIf chk分类(0).Value = 1 And chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And 类型 In (1, 2) "
            ElseIf chk分类(0).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And 类型 In (1, 3) "
            ElseIf chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And 类型 In (2, 3) "
            ElseIf chk分类(0).Value = 1 Then
                gstrSql = gstrSql & " And 类型 = 1 "
            ElseIf chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And 类型 = 2 "
            ElseIf chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And 类型 = 3 "
            End If
                
            gstrSql = gstrSql & " Start With 上级id Is Null Connect By Prior ID = 上级id " & _
                " Order By 诊疗分类目录.类型, Level, 编码 "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk分类(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_西药", "西药", 1)
            If chk分类(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中成药", "中成药", 1)
            If chk分类(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
    
            Do While Not rsTemp.EOF
                If rsTemp!层 = 1 Then
                    If InStr(1, "," & strID & ",", "," & rsTemp!ID & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & rsTemp!ID
                    End If
                    Set objNode = tvwClass.Nodes.Add("_" & rsTemp!类型, 4, "_" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, 1)
'                    objNode.Expanded = True
                Else
                    If InStr(1, "," & strID & ",", "," & rsTemp!上级ID & ",") > 0 Then
                        If InStr(1, "," & strID & ",", "," & rsTemp!ID & ",") = 0 Then
                            strID = IIf(strID = "", "", strID & ",") & rsTemp!ID
                        End If
                        Set objNode = tvwClass.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, 1)
                    Else
                        Set objNode = tvwClass.Nodes.Add("_" & rsTemp!类型, 4, "_" & rsTemp!ID, "[" & rsTemp!编码 & "]" & rsTemp!名称, 1)
                    End If
                End If
                rsTemp.MoveNext
            Loop
        Case MediType.药品品种
            gstrSql = "Select Distinct a.Id, a.编码, a.名称, a.类别, Decode(a.类别, '5', '西药', Decode(a.类别, '6', '中成药', '中草药')) As 类型 " & _
                " From 诊疗项目目录 A, 诊疗项目别名 B " & _
                " Where a.Id = b.诊疗项目id And a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') "
                      
            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.编码 Like [1] Or b.名称 Like [2] Or b.简码 Like [2]) "
            End If
            
            If chk分类(0).Value = 1 And chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '6', '7') "
            ElseIf chk分类(0).Value = 1 And chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '6') "
            ElseIf chk分类(0).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '7') "
            ElseIf chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('6', '7') "
            ElseIf chk分类(0).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '5' "
            ElseIf chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '6' "
            ElseIf chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '7' "
            End If
                
            gstrSql = gstrSql & " Order By a.类别, a.编码, a.Id, a.名称 "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk分类(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_西药", "西药", 1)
            If chk分类(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中成药", "中成药", 1)
            If chk分类(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("_" & rsTemp!类型, 4, "_" & rsTemp!ID, rsTemp!名称, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.药品规格
            gstrSql = "Select Distinct a.Id,  '[' || a.编码 || ']' || a.名称 || '(' || a.规格 || ')' As 名称, a.类别, Decode(a.类别, '5', '西药', Decode(a.类别, '6', '中成药', '中草药')) As 类型 " & _
                " From 收费项目目录 A, 收费项目别名 B " & _
                " Where a.Id = b.收费细目id And a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') "
            
            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.编码 Like [1] Or b.名称 Like [2] Or b.简码 Like [2]) "
            End If
            
            If chk分类(0).Value = 1 And chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '6', '7') "
            ElseIf chk分类(0).Value = 1 And chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '6') "
            ElseIf chk分类(0).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('5', '7') "
            ElseIf chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 In ('6', '7') "
            ElseIf chk分类(0).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '5' "
            ElseIf chk分类(1).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '6' "
            ElseIf chk分类(2).Value = 1 Then
                gstrSql = gstrSql & " And a.类别 = '7' "
            End If
            
            gstrSql = gstrSql & " Order By a.类别, '[' || a.编码 || ']' || a.名称 || '(' || a.规格 || ')', a.Id "
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk分类(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_西药", "西药", 1)
            If chk分类(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中成药", "中成药", 1)
            If chk分类(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("_" & rsTemp!类型, 4, "_" & rsTemp!ID, rsTemp!名称, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.药品剂型
            gstrSql = "Select 编码, 名称, 简码 From 药品剂型 "
            
            If strInput <> "" Then
                gstrSql = gstrSql & " Where (编码 Like [1] Or 名称 Like [2] Or 简码 Like [2]) "
            End If
            
            gstrSql = gstrSql & " Order By 编码 "
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
                            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("Root", 4, "_" & rsTemp!名称, rsTemp!名称, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.给药途径
            gstrSql = "Select Distinct a.Id, a.编码, a.名称 " & _
                " From 诊疗项目目录 A, 诊疗项目别名 B " & _
                " Where a.Id = b.诊疗项目id And a.类别 = 'E' And a.操作类型 = '2' And a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') "

            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.编码 Like [1] Or b.名称 Like [2] Or b.简码 Like [2]) "
            End If

            gstrSql = gstrSql & " Order By a.编码 "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("Root", 4, "_" & rsTemp!ID, rsTemp!名称, 1)

                rsTemp.MoveNext
            Loop
    End Select
    
    tvwClass.Nodes("Root").Selected = True
    tvwClass.Nodes("Root").Expanded = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo发药类型_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cbo分类_Click()
    With cbo分类
        If .ListIndex <> Val(.Tag) Then
            .Tag = .ListIndex
            
            txtInput.Text = ""
            txtInput.Tag = ""
            
            Call GetSelect(Val(cbo分类.Tag), Trim(txtInput.Text))
        End If
    End With
End Sub



Private Sub chk分类_Click(Index As Integer)
    If chk分类(0).Value = 0 And chk分类(1).Value = 0 And chk分类(2).Value = 0 Then
        chk分类(Index).Value = 1
        Exit Sub
    End If
    
    Call GetSelect(Val(cbo分类.Tag), Trim(txtInput.Text))
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    Dim str材质 As String
    Dim int分类 As Integer
    Dim str内容 As String
    Dim str发药类型 As String
    Dim lngCount As Long
    
    If chk分类(0).Value = 1 And chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
        str材质 = "5,6,7"
    ElseIf chk分类(0).Value = 1 And chk分类(1).Value = 1 Then
        str材质 = "5,6"
    ElseIf chk分类(0).Value = 1 And chk分类(2).Value = 1 Then
        str材质 = "5,7"
    ElseIf chk分类(1).Value = 1 And chk分类(2).Value = 1 Then
        str材质 = "6,7"
    ElseIf chk分类(0).Value = 1 Then
        str材质 = "5"
    ElseIf chk分类(1).Value = 1 Then
        str材质 = "6"
    ElseIf chk分类(2).Value = 1 Then
        str材质 = "7"
    End If
    
    int分类 = Val(cbo分类.Tag)
    
    For lngCount = 1 To tvwClass.Nodes.Count
        If tvwClass.Nodes(lngCount).Key <> "Root" And _
            tvwClass.Nodes(lngCount).Key <> "_中成药" And _
            tvwClass.Nodes(lngCount).Key <> "_中草药" And _
            tvwClass.Nodes(lngCount).Key <> "_西药" And _
            tvwClass.Nodes(lngCount).Checked Then
            str内容 = IIf(str内容 = "", "", str内容 & ",") & Mid(tvwClass.Nodes(lngCount).Key, 2)
        End If
    Next
    
    If Trim(str内容) = "" Then
        MsgBox "请在列表中选择具体分类。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    str发药类型 = Trim(cbo发药类型.Text)
    
    If str发药类型 = "" Then
        If MsgBox("你没有选择发药类型，将清除对应的药品的发药类型，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    gstrSql = "Zl_药品规格_发药类型("
    '药品材质
    gstrSql = gstrSql & "'" & str材质 & "'" & ","
    '分类方式
    gstrSql = gstrSql & int分类 & ","
    '分类内容
    gstrSql = gstrSql & "'" & str内容 & "'" & ","
    '发药类型
    gstrSql = gstrSql & IIf(str发药类型 = "", "Null", "'" & str发药类型 & "'")
    gstrSql = gstrSql & ")"
    
    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, "保存发药类型")
    
    MsgBox "保存成功！", vbExclamation, gstrSysName
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd分类_Click()
    Call GetSelect(Val(cbo分类.Tag), Trim(txtInput.Text))
End Sub
Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select 名称 From 发药类型 Order By 编码"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "取发药类型")
    
    With cbo发药类型
        .Clear
        Do While Not rsData.EOF
            cbo发药类型.AddItem rsData.Fields(0).Value
            rsData.MoveNext
        Loop
    End With
    
    With cbo分类
        .Clear
        .AddItem "0-药品分类"
        .AddItem "1-药品品种"
        .AddItem "2-药品规格"
        .AddItem "3-药品剂型"
        .AddItem "4-给药途径"
        
        .ListIndex = 0
        .Tag = 0
    End With
    
    mstrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    Call GetSelect(Val(cbo分类.Tag), Trim(txtInput.Text))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtFind.Text) <> "" Then
        Call GetSelect(Val(cbo分类.Tag), Trim(txtFind.Text), True)
        
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFind = ""
    Set mrsTemp = Nothing
End Sub

Private Sub tvwClass_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvwClass, Node, Node.Checked
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim objItem As Node
    
    If KeyAscii = vbKeyReturn And Trim(txtFind.Text) <> "" Then
        zlControl.TxtSelAll txtFind
        If mstrFind <> Trim(txtFind.Text) Then    '已经是最后了
            mstrFind = Trim(txtFind.Text)
            Call GetSelect(Val(cbo分类.Tag), UCase(Trim(txtFind.Text)), True)
            If mrsTemp.RecordCount > 0 Then
                For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                mrsTemp.MoveNext
            Else
                MsgBox "没有找到你想要得数据！", vbInformation, gstrSysName
                txtFind.SetFocus
                zlControl.TxtSelAll txtFind
            End If
        Else
            If Not mrsTemp.EOF Then
                mrsTemp.MoveNext
                If Not mrsTemp.EOF Then
                    For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                End If
            ElseIf mrsTemp.EOF Then
                mrsTemp.MoveFirst
                MsgBox "已查询到最后！", vbInformation, gstrSysName
                If Not mrsTemp.EOF Then
                    For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                End If
            End If
        End If
    End If
End Sub

Private Sub txtInput_GotFocus()
    zlControl.TxtSelAll txtInput
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call GetSelect(Val(cbo分类.Tag), Trim(txtInput.Text))
    End If
End Sub


