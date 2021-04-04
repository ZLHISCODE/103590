VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm医保扩展对照 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置辅助编码"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frm医保扩展对照.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4230
      TabIndex        =   12
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmd删除 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   4230
      TabIndex        =   14
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmd修改 
      Caption         =   "修改(&M)"
      Height          =   350
      Left            =   2970
      TabIndex        =   13
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2970
      TabIndex        =   11
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmd新增 
      Caption         =   "新增(&N)"
      Height          =   350
      Left            =   1710
      TabIndex        =   10
      Top             =   4440
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   30
      TabIndex        =   17
      Top             =   2400
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Height          =   5205
      Left            =   5670
      TabIndex        =   16
      Top             =   -150
      Width           =   30
   End
   Begin VB.ComboBox cbo类别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   5940
      TabIndex        =   15
      Top             =   540
      Width           =   1100
   End
   Begin VB.TextBox txt说明 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   8
      Top             =   1500
      Width           =   3495
   End
   Begin VB.TextBox txt医保项目信息 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   3
      Top             =   690
      Width           =   3495
   End
   Begin VB.TextBox txt项目信息 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      TabIndex        =   1
      Top             =   300
      Width           =   3495
   End
   Begin MSComctlLib.ListView lvwAdvance 
      Height          =   1755
      Left            =   270
      TabIndex        =   9
      Top             =   2580
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3096
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "类别"
         Object.Width           =   1376
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "项目编码"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "项目名称"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "说明"
         Object.Width           =   3810
      EndProperty
   End
   Begin VB.CommandButton cmd医保项目信息 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5010
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lbl类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "类别(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label lbl说明 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   7
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lbl医保项目 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保项目信息(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label lbl项目信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HIS项目信息"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frm医保扩展对照"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mlng收费细目ID As Long
Private mrsTemp As New ADODB.Recordset

Private Sub cmd保存_Click()
    If Not IsValid Then Exit Sub
    If Not SaveData Then Exit Sub
    
    Call SetConsEnable(False)
    Call RefreshData
End Sub

Private Sub cmd取消_Click()
    Call SetConsEnable(False)
    If lvwAdvance.ListItems.Count <> 0 Then Call lvwAdvance_ItemClick(lvwAdvance.ListItems(1))
End Sub

Private Sub cmd删除_Click()
    On Error GoTo errHand
    If lvwAdvance.ListItems.Count = 0 Then Exit Sub
    If lvwAdvance.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("你确认要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "ZL_医保对照明细_Delete(" & mint险类 & "," & mlng收费细目ID & ",'" & lvwAdvance.SelectedItem.SubItems(1) & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd退出_Click()
    Unload Me
End Sub

Private Sub cmd新增_Click()
    Call SetConsEnable(True)
    
    txt医保项目信息.Text = ""
    cbo类别.ListIndex = 0
    txt说明.Text = ""
    txt医保项目信息.SetFocus
End Sub

Private Sub cmd修改_Click()
    Call SetConsEnable(True)
    
    txt医保项目信息.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    '提取HIS项目的编码与名称
    gstrSQL = "Select '['||编码||']'||名称 AS 项目信息 From 收费细目 Where ID=[1]"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取HIS项目的编码与名称", mlng收费细目ID)
    Me.txt项目信息.Text = mrsTemp!项目信息
    
    '提取医保对照类别
    gstrSQL = "Select 编码,名称 From 医保对照类别 Where 险类=[1] And Nvl(编码,0)<>0 Order by 编码"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保对照类别", mint险类)
    With mrsTemp
        Me.cbo类别.Clear
        Do While Not .EOF
            cbo类别.AddItem !名称
            cbo类别.ItemData(cbo类别.NewIndex) = !编码
            .MoveNext
        Loop
        cbo类别.ListIndex = 0
    End With
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ShowEditor(ByVal int险类 As Integer, ByVal lng收费细目ID As Long)
    mint险类 = int险类
    mlng收费细目ID = lng收费细目ID
    Me.Show 1
End Sub

Private Sub SetConsEnable(ByVal blnEnable As Boolean)
    cmd保存.Enabled = blnEnable
    cmd取消.Enabled = blnEnable
    txt医保项目信息.Enabled = blnEnable
    cmd医保项目信息.Enabled = blnEnable
    cbo类别.Enabled = blnEnable
    txt说明.Enabled = blnEnable
    
    cmd新增.Enabled = Not blnEnable
    cmd修改.Enabled = Not blnEnable
    cmd删除.Enabled = Not blnEnable
    lvwAdvance.Enabled = Not blnEnable
End Sub

Private Sub lvwAdvance_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intDO As Integer, intCOUNT As Integer
    With lvwAdvance
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        txt医保项目信息.Text = "[" & .SelectedItem.SubItems(1) & "]" & .SelectedItem.SubItems(2)
        txt医保项目信息.Tag = .SelectedItem.SubItems(1)
        txt说明.Text = .SelectedItem.SubItems(3)
        
        intCOUNT = cbo类别.ListCount
        For intDO = 1 To intCOUNT
            If Val(.SelectedItem.Tag) = cbo类别.ItemData(intDO - 1) Then
                cbo类别.ListIndex = intDO - 1
                Exit For
            End If
        Next
    End With
End Sub

Private Function IsValid() As Boolean
    If txt医保项目信息.Tag = "" Then
        MsgBox "请选择医保项目！", vbInformation, gstrSysName
        Exit Function
    End If
    
    IsValid = True
End Function

Private Function SaveData() As Boolean
    On Error GoTo errHand
    gstrSQL = "ZL_医保对照明细_Modify(" & mint险类 & "," & mlng收费细目ID & "," & Me.cbo类别.ItemData(Me.cbo类别.ListIndex) & ",'" & txt医保项目信息.Tag & "','" & txt说明.Text & "')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    SaveData = True
    
    Call RefreshData
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub RefreshData()
    Dim lvwItem As ListItem
    '提取已完成的对照信息
    gstrSQL = "Select A.类别 AS 类别编码,B.名称 AS 类别名称,A.收费细目ID,A.项目编码,C.名称 AS 项目名称,A.说明 " & _
        " From 医保对照明细 A,医保对照类别 B,保险项目 C" & _
        " Where A.险类=B.险类 And A.险类=[1] And A.收费细目ID=[2]" & _
        " And C.险类=A.险类 And C.编码=A.项目编码 And A.类别=B.编码 And B.编码<>0" & _
        " Order by A.类别,A.项目编码"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取已完成的对照信息", mint险类, mlng收费细目ID)
    With mrsTemp
        lvwAdvance.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvwAdvance.ListItems.Add(, "K_" & lvwAdvance.ListItems.Count, !类别名称)
            lvwItem.SubItems(1) = !项目编码
            lvwItem.SubItems(2) = !项目名称
            lvwItem.SubItems(3) = Nvl(!说明)
            lvwItem.Tag = !类别编码
            .MoveNext
        Loop
    End With
    
    If lvwAdvance.ListItems.Count <> 0 Then Call lvwAdvance_ItemClick(lvwAdvance.ListItems(1))
End Sub

Private Sub txt医保项目信息_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnReturn As Boolean
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    StrInput = UCase(Trim(txt医保项目信息.Text))
    If StrInput = "" Then
        MsgBox "请输入医保项目信息!", vbInformation, gstrSysName
        Exit Sub
    End If
    If Mid(StrInput, 1, 1) = "[" Then
        If InStr(2, StrInput, "]") <> 0 Then
            StrInput = Mid(StrInput, 2, InStr(2, StrInput, "]") - 2)
        Else
            StrInput = Mid(StrInput, 2)
        End If
    End If
    
    gstrSQL = "Select 编码,名称,简码,附注 From 保险项目 " & _
        " Where 险类=[1]" & _
        " And (编码 Like [2] || '%' Or 名称 Like [2] || '%' Or Upper(简码) Like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有项目", mint险类, StrInput)
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有找到匹配的医保项目,请重新输入!", vbInformation, gstrSysName
        Exit Sub
    End If
    If rsTemp.RecordCount > 1 Then
        blnReturn = frmListSel.ShowSelect(mint险类, rsTemp, "编码", "医保项目选择", "请选择对应的医保项目：")
    Else
        blnReturn = True
    End If
    If blnReturn Then
        txt医保项目信息.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
        txt医保项目信息.Tag = rsTemp!编码
    End If
End Sub
