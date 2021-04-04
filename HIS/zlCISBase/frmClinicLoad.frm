VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "诊疗项目导入"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmClinicLoad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkExse 
      Caption         =   "选定项目的从属项目自动成为主项目的收费对照(&Z)"
      Height          =   180
      Left            =   1200
      TabIndex        =   19
      Top             =   2040
      Width           =   4530
   End
   Begin VB.CheckBox chkErrAsk 
      Caption         =   "导入遇错提醒(&E)"
      Height          =   240
      Left            =   3255
      TabIndex        =   2
      Top             =   1740
      Width           =   1725
   End
   Begin VB.CheckBox chkLeaves 
      Caption         =   "显示所有下级(&V)"
      Height          =   240
      Left            =   1200
      TabIndex        =   1
      Top             =   1740
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "返回(&X)"
      Height          =   350
      Left            =   5940
      TabIndex        =   5
      Top             =   2085
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5940
      Picture         =   "frmClinicLoad.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3465
      Width           =   1100
   End
   Begin VB.CommandButton cmdLoadIn 
      Caption         =   "导入(&L)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5940
      TabIndex        =   4
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&C)"
      Height          =   350
      Left            =   5940
      Picture         =   "frmClinicLoad.frx":06D4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2955
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全选(&S)"
      Height          =   350
      Left            =   5940
      Picture         =   "frmClinicLoad.frx":081E
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   5445
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      Width           =   285
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1185
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1305
      Width           =   4260
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -15
      TabIndex        =   10
      Top             =   1155
      Width           =   7290
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   5250
      Left            =   105
      TabIndex        =   3
      Top             =   2280
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9260
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4470
      Left            =   1200
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   2340
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   7885
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
      Left            =   5700
      Top             =   6900
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
            Picture         =   "frmClinicLoad.frx":0968
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLoad.frx":0F02
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicLoad.frx":149C
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "已导入:0条"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   5940
      TabIndex        =   18
      Top             =   7035
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lbl诊疗分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    当前诊疗分类目录："
      ForeColor       =   &H00000040&
      Height          =   180
      Left            =   675
      TabIndex        =   16
      Top             =   915
      Width           =   1980
   End
   Begin VB.Label lblCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "已选择:1234条"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   5940
      TabIndex        =   15
      Top             =   6735
      Width           =   1170
   End
   Begin VB.Label lblCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "项目数:1234条"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5940
      TabIndex        =   14
      Top             =   6435
      Width           =   1170
   End
   Begin VB.Label lbl项目 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费项目(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   13
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费分类(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   12
      Top             =   1380
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClinicLoad.frx":1A36
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   690
      TabIndex        =   9
      Top             =   60
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   105
      Picture         =   "frmClinicLoad.frx":1B0B
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：根据选择的收费项目，直接填写到诊疗项目中，同时建立收费对照
'调用：
'    If Me.tvwClass.SelectedItem Is Nothing Then MsgBox "尚未设置分类,不能增删项目！", vbExclamation, gstrSysName: Exit Sub
'    With frmClinicLoad
'        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
'        .Show 1, Me
'    End With
'---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Private Sub chkLeaves_Click()
    Call txt分类_Change
End Sub

Private Sub cmdAllCls_Click()
    For Each objItem In Me.lvwItems.ListItems
        objItem.Checked = False
    Next
    Me.lblCount(1).Caption = "已选择:0条": Me.lblCount(1).Tag = 0
    Me.cmdLoadIn.Enabled = False
End Sub

Private Sub cmdAllSel_Click()
    For Each objItem In Me.lvwItems.ListItems
        objItem.Checked = True
    Next
    Me.lblCount(1).Caption = "已选择:" & Me.lblCount(0).Tag & "条": Me.lblCount(1).Tag = Me.lblCount(0).Tag
    If Val(Me.lblCount(1).Tag) > 0 Then
        Me.cmdLoadIn.Enabled = True
    Else
        Me.cmdLoadIn.Enabled = False
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdLoadIn_Click()
    Dim intMode As Integer, blnError As Boolean
    
    intMode = Val(zlDatabase.GetPara(61, glngSys)) '诊疗项目编码递增模式
    
    Me.lblCount(2).Visible = True
    Me.Enabled = False
    intCount = 0
    blnError = False
    For Each objItem In Me.lvwItems.ListItems
        If objItem.Checked Then
            objItem.EnsureVisible
            
            gstrSql = "zl_诊疗项目_Load(" & Mid(objItem.Key, 2) & "," & Me.Tag & "," & intMode & "," & chkExse.Value & ")"
            Err = 0: On Error Resume Next
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            If Err <> 0 Then
                If Me.chkErrAsk.Value = 1 Then
                    strTemp = Err.Description
                    If InStr(1, strTemp, "ZLSOFT") > 0 Then
                        strTemp = Mid(strTemp, InStr(1, strTemp, "[ZLSOFT]") + 8)
                        strTemp = Mid(strTemp, 1, InStr(1, strTemp, "[ZLSOFT]") - 1)
                    End If
                    strTemp = "由于下面的错误，导入“" & objItem.Text & "”时失败：" & vbCrLf & strTemp & vbCrLf & vbCrLf & "可绕过继续导入后续项目，继续吗？"
                    If MsgBox(strTemp, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit For
                End If
                blnError = True
            Else
                objItem.Checked = False
            End If
            '补充导入过程
            intCount = intCount + 1
            Me.lblCount(2).Caption = "已导入:" & intCount & "条"
            DoEvents
        End If
    Next
    Call lvwItems_ItemCheck(Me.lvwItems.ListItems(1))
    Me.lblCount(2).Visible = False
    Me.Enabled = True
    If blnError = False Then
        MsgBox "本次导入完成，请检查调整导入项目的特性设置！", vbExclamation, gstrSysName
    ElseIf Me.chkErrAsk.Value = 0 Then
        MsgBox "导入结束，部分项目导入时发生错误(仍为选中状态的项目)，请检查！", vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdReturn_Click()
    Unload Me
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
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,编码,名称 From 诊疗分类目录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.Tag))
        
    With rsTemp
        Me.lbl诊疗分类.Caption = "    当前诊疗分类目录：" & "[" & !编码 & "]" & !名称
        strTemp = !编码
    End With
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 收费分类目录" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then MsgBox "尚未建立收费分类目录！", vbExclamation, gstrSysName: Unload Me: Exit Sub
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            If !编码 = strTemp Then objNode.Selected = True
            .MoveNext
        Loop
        If Me.tvwClass.SelectedItem Is Nothing Then
            Me.tvwClass.Nodes(1).Selected = True
        End If
        If Not (Me.tvwClass.SelectedItem Is Nothing) Then
            Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        End If
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txt分类.SetFocus: Exit Sub
    End If
    Call cmdReturn_Click
End Sub

Private Sub Form_Load()
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2500
        .Add , "_编码", "编码", 1000
        .Add , "_计算单位", "计算单位", 900
        .Add , "_服务对象", "服务对象", 1200
        .Add , "_类别", "类别", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.lblCount(2).Visible Then
        Cancel = 1
        Exit Sub
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

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    intCount = 0
    For Each objItem In Me.lvwItems.ListItems
        If objItem.Checked Then intCount = intCount + 1
    Next
    Me.lblCount(1).Caption = "已选择:" & intCount & "条": Me.lblCount(1).Tag = intCount
    If Val(Me.lblCount(1).Tag) > 0 Then
        Me.cmdLoadIn.Enabled = True
    Else
        Me.cmdLoadIn.Enabled = False
    End If
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

Private Sub txt分类_Change()
    Err = 0: On Error GoTo ErrHand
    
    If Me.chkLeaves.Value = 0 Then
        gstrSql = "Select I.ID,I.编码,I.名称,I.计算单位,I.服务对象,K.名称 As 类别" & _
                "   From 收费项目目录 I, 收费项目类别 K, 诊疗项目类别 N" & _
                " Where I.类别=K.编码 And I.类别=N.编码 And I.类别>='A' And 分类id=[1] " & _
                "       And (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Else
        gstrSql = "Select I.ID,I.编码,I.名称,I.计算单位,I.服务对象,K.名称 As 类别" & _
                "   From 收费项目目录 I, 收费项目类别 K, 诊疗项目类别 N" & _
                " Where I.类别=K.编码 And I.类别=N.编码 And I.类别>='A' And" & _
                "       分类id In (Select ID From 收费分类目录 Start With ID=[1] Connect by Prior ID=上级ID)" & _
                "       And (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.txt分类.Tag))
    
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "item": objItem.SmallIcon = "item"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            Select Case IIf(IsNull(!服务对象), 0, !服务对象)
            Case 1
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = "门诊"
            Case 2
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = "住院"
            Case 3
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_服务对象").Index - 1) = "门诊与住院"
            Case Else
            End Select
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_类别").Index - 1) = IIf(IsNull(!类别), "", !类别)
            objItem.Checked = True
            .MoveNext
        Loop
        Me.lblCount(0).Caption = "项目数:" & .RecordCount & "条": Me.lblCount(0).Tag = .RecordCount
        Me.lblCount(1).Caption = "已选择:" & .RecordCount & "条": Me.lblCount(1).Tag = .RecordCount
    End With
    If Val(Me.lblCount(1).Tag) > 0 Then
        Me.cmdLoadIn.Enabled = True
    Else
        Me.cmdLoadIn.Enabled = False
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
