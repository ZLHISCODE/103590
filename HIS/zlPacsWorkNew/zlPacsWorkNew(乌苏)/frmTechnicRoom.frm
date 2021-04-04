VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTechnicRoom 
   BorderStyle     =   0  'None
   Caption         =   "执行间设置"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "frmTechnicRoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTechniRoom 
      Height          =   6525
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      Begin VB.TextBox txtNoPrefix 
         Height          =   300
         Left            =   6435
         MaxLength       =   20
         TabIndex        =   11
         Top             =   5340
         Width           =   1170
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   3075
         Picture         =   "frmTechnicRoom.frx":058A
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5865
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   4215
         Picture         =   "frmTechnicRoom.frx":06D4
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   5865
         Width           =   1100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   6510
         TabIndex        =   8
         Top             =   5865
         Width           =   1100
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "恢复(&R)"
         Height          =   350
         Left            =   5370
         TabIndex        =   7
         Top             =   5865
         Width           =   1100
      End
      Begin VB.ComboBox cboDevice 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   5340
         Width           =   1830
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   840
         MaxLength       =   20
         TabIndex        =   3
         Top             =   5340
         Width           =   1635
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   4320
         Top             =   600
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
               Picture         =   "frmTechnicRoom.frx":081E
               Key             =   "Room"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwRoom 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   8281
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "号码前缀"
         Height          =   180
         Left            =   5610
         TabIndex        =   12
         Top             =   5400
         Width           =   720
      End
      Begin VB.Label lblDevice 
         Caption         =   "设备(&D)"
         Height          =   180
         Left            =   2805
         TabIndex        =   4
         Top             =   5400
         Width           =   630
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   5400
         Width           =   630
      End
      Begin VB.Label lblNote 
         Caption         =   "设置本科室的执行间后，才能有效进行执行的安排。"
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   4140
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmTechnicRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngdept As Long '科室


Private Sub cmdAdd_Click()
    Me.lblName.Tag = "": Me.txtName.Text = "": Me.txtName.Enabled = True
    Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True: cboDevice.Enabled = True: cboDevice.ListIndex = 0
    Me.txtName.SetFocus
End Sub


Private Sub cmdDel_Click()
    If MsgBoxD(Me, "真的删除执行间“" & Trim(Me.txtName.Text) & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "zl_医技执行房间_Delete(" & Val(mlngdept) & ",'" & Trim(Me.txtName.Text) & "')"
    err = 0: On Error GoTo errHand
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Call zlRoomRef
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call zlRoomRef
End Sub

Private Sub cmdSave_Click()
Dim blnExist As Boolean, i As Integer

    If Trim(Me.txtName.Text) = "" Then
        MsgBoxD Me, "名称必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBoxD Me, "名称超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvwRoom.ListItems.Count
        If txtName.Text = lvwRoom.ListItems(i).Text Then blnExist = True: Exit For '已经存在
    Next
    '-----------------------------------------
    err = 0: On Error GoTo errHand
    If Me.lblName.Tag = "" And Not blnExist Then
        gstrSQL = "zl_医技执行房间_Insert(" & Val(mlngdept) & ",'" & Trim(Me.txtName.Text) & "','" & NeedNo(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    Else
        gstrSQL = "zl_医技执行房间_Update(" & Val(mlngdept) & ",'" & Trim(Me.lblName.Tag) & "','" & Trim(Me.txtName.Text) & "','" & NeedNo(cboDevice.Text) & "','" & txtNoPrefix.Text & "')"
    End If
    
    err = 0: On Error GoTo errHand
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    MsgBoxD Me, "执行间保存成功！", vbInformation, gstrSysName
    Call zlRoomRef
    txtName.SetFocus
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    ElseIf InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim rsTemp As New ADODB.Recordset
    Me.lvwRoom.ListItems.Clear
    With Me.lvwRoom.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 3000
        .Add , "检查设备", "检查设备", 3000
        .Add , "号码前缀", "号码前缀", 2000
    End With
    
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 状态=1 and 类型=4"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, App.ProductName)
    cboDevice.Clear
    cboDevice.AddItem ""
    Do Until rsTemp.EOF
        cboDevice.AddItem rsTemp!设备号 & "-" & rsTemp!设备名
        rsTemp.MoveNext
    Loop
End Sub

Private Sub Form_Resize()
    fraTechniRoom.Left = (Me.ScaleWidth - fraTechniRoom.Width) / 2
End Sub




Private Sub lvwRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtName.Text = Item.Text
    Me.lblName.Tag = Me.txtName.Text
    Me.txtNoPrefix.Text = Item.SubItems(2)
    
    SeekIndexWithNo cboDevice, Item.SubItems(1), True
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Or KeyAscii = 95 Then KeyAscii = 0
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Public Sub zlRoomRef()
Dim objItem As ListItem
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    gstrSQL = "Select 执行间,检查设备,号码前缀 From 医技执行房间 where 科室id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Val(mlngdept)))
    Me.lvwRoom.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwRoom.ListItems.Add(, , !执行间, 1, 1)
            
            objItem.SubItems(1) = Nvl(!检查设备)
            objItem.SubItems(2) = Nvl(!号码前缀)
            .MoveNext
        Loop
    End With
    
    err = 0: On Error Resume Next
    If Me.lvwRoom.ListItems.Count > 0 Then
        Me.lvwRoom.ListItems(1).Selected = True
        Me.lvwRoom.SelectedItem.EnsureVisible
    End If
    
    err = 0: On Error GoTo 0
    If Me.lvwRoom.ListItems.Count > 0 Then
        Call lvwRoom_ItemClick(Me.lvwRoom.SelectedItem)
        Me.txtName.Enabled = True: cboDevice.Enabled = True
        Me.cmdDel.Enabled = True: Me.cmdSave.Enabled = True: Me.cmdRestore.Enabled = True
    Else
        Me.lblName.Tag = "": Me.txtName.Text = "": cboDevice.ListIndex = 0
        Me.txtName.Enabled = False: cboDevice.Enabled = False
        Me.cmdDel.Enabled = False: Me.cmdSave.Enabled = False: Me.cmdRestore.Enabled = False
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
