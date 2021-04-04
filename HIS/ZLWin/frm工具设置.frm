VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm工具设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "外部工具设置"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frm工具设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   5805
      TabIndex        =   11
      Top             =   570
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "应用(&A)"
      Height          =   350
      Left            =   5805
      TabIndex        =   10
      Top             =   165
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "工具设置"
      Height          =   4200
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   5655
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   4425
         TabIndex        =   8
         Top             =   1110
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   3330
         TabIndex        =   7
         Top             =   1110
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&N)"
         Height          =   350
         Left            =   2235
         TabIndex        =   6
         Top             =   1110
         Width           =   1100
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "…"
         Height          =   300
         Left            =   5190
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   690
         Width           =   4005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   2
         Top             =   300
         Width           =   1425
      End
      Begin MSComctlLib.ListView lvwTool 
         Height          =   2625
         Left            =   90
         TabIndex        =   9
         Top             =   1515
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   4630
         View            =   3
         Arrange         =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "ToolName"
            Object.Tag             =   "ToolName"
            Text            =   "工具名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "ToolFile"
            Object.Tag             =   "ToolFile"
            Text            =   "工具文件"
            Object.Width           =   6068
         EndProperty
      End
      Begin ComctlLib.ImageList ist16 
         Left            =   4830
         Top             =   1380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm工具设置.frx":1082
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "工具文件(&P)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "工具名称(&G)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   990
      End
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm工具设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mblnApply As Boolean '应用
Private mblnChange As Boolean

Private Sub cmdAdd_Click()
    Dim objItem As ListItem
    If IsValid = False Then Exit Sub
    With lvwTool
        Set objItem = .ListItems.Add(, "K" & .ListItems.Count + 1, txtEdit(0).Text)
        objItem.SubItems(1) = txtEdit(1).Text
        objItem.Selected = True
        objItem.EnsureVisible
    End With
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim lngIndex As Long
    If lvwTool.SelectedItem Is Nothing Then Exit Sub
    
    With lvwTool
        lngIndex = .SelectedItem.Index
        If .ListItems.Count = 1 Then
        ElseIf lngIndex < .ListItems.Count Then
            .ListItems(lngIndex + 1).Selected = True
        Else
            .ListItems(lngIndex - 1).Selected = True
        End If
        .ListItems.Remove lngIndex
    End With
    Call SetCtlEnable
    mblnChange = True
End Sub

Private Sub cmdModify_Click()
    Dim objItem As ListItem
    If IsValid = False Then Exit Sub
    If lvwTool.SelectedItem Is Nothing Then Exit Sub
    
    With lvwTool
        Set objItem = .SelectedItem
        objItem.Text = txtEdit(0).Text
        objItem.SubItems(1) = txtEdit(1).Text
    End With
    mblnChange = True
End Sub

Private Sub cmdPath_Click()
    Dim objFile As New FileSystemObject
    With cdgPub
        .DialogTitle = "选择外部工具文件"
        .Filter = "外部工具文件(*.EXE)|*.EXE"
        .Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        
        If objFile.FileExists(txtEdit(1).Text) Then
            .InitDir = objFile.GetParentFolderName(txtEdit(1).Text)
            .FileName = objFile.GetFileName(txtEdit(1).Text)
        Else
            .InitDir = "": .FileName = ""
        End If
        .CancelError = True
        On Error GoTo errH
        .ShowOpen
        On Error GoTo 0
        txtEdit(1).Text = .FileName
    End With
errH:
End Sub

Private Sub cmd确定_Click()
    If SaveToolsRegInfor = False Then Exit Sub
    mblnApply = True
    mblnChange = False
    Unload Me
End Sub

 

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadRegTool
    If Not lvwTool.SelectedItem Is Nothing Then
        Call lvwTool_ItemClick(lvwTool.SelectedItem)
    Else
        Call SetCtlEnable
    End If
    mblnChange = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "*^%$#@!|;,+?", Asc(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Function IsValid() As Boolean
    '------------------------------------------------------------------------
    '功能:检查输入的值是否合法
    '参数:
    '返回:输入值合法,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/22
    '------------------------------------------------------------------------
    If zlCommFun.ActualLen(txtEdit(0).Text) > 20 Then
        MsgBox "输入的工具名称长度不能大于20个字符或10个汉字,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Function
    End If
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "工具名称必需输入,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        Exit Function
    End If
        
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "必需选择工具文件,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    Dim objFile As New FileSystemObject
    If objFile.FileExists(txtEdit(1).Text) = False Then
        MsgBox "工具文件不存在,可能已经被删除,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("你已经对外部工具进行了改变,是否真的退出?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub lvwTool_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With Item
        txtEdit(0).Text = Item.Text
        txtEdit(1).Text = Item.SubItems(1)
    End With
    Call SetCtlEnable
End Sub

Private Sub SetCtlEnable()
    '------------------------------------------------------------------------
    '功能:设置控件的Enabled属性值
    '参数:
    '返回:
    '编制:刘兴宏
    '日期:2007/08/22
    '------------------------------------------------------------------------
    Dim blnSel As Boolean
    blnSel = Not Me.lvwTool.SelectedItem Is Nothing
    cmdModify.Enabled = blnSel
    cmdDelete.Enabled = blnSel
End Sub

Private Sub LoadRegTool()
    '------------------------------------------------------------------------
    '功能:加载注册中的相关外部工具
    '参数:
    '返回:
    '编制:刘兴宏
    '日期:2007/08/22
    '------------------------------------------------------------------------
    '注册表的存储规则为:
    '     1.HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ZLSOFT\公共全局:下建立TOOLS
    '     2.在TOOLS下建立:TOOLFILES,内容包含:工具名称,工具文件|工具名称1,工具文件1......
    Dim strReg As String, arrTemp As Variant, ArrTool As Variant, i As Long
    Dim objItem As ListItem
    strReg = GetSetting("ZLSOFT", "公共全局\TOOLS", "TOOLFILES", "")
    With lvwTool
        .ListItems.Clear
        If strReg = "" Then Exit Sub
        ArrTool = Split(strReg, "|")
        For i = 0 To UBound(ArrTool)
            arrTemp = Split(ArrTool(i) & ",", ",")
            If Trim(arrTemp(0)) <> "" And arrTemp(1) <> "" Then
                Set objItem = .ListItems.Add(, "K" & i, arrTemp(0))
                objItem.SubItems(1) = arrTemp(1)
                If .SelectedItem Is Nothing Then
                    objItem.Selected = True
                    objItem.EnsureVisible
                End If
            End If
        Next
    End With
End Sub

Private Function SaveToolsRegInfor() As Boolean
    '------------------------------------------------------------------------
    '功能:保存工具的信息到注册表
    '参数:
    '返回:保存成功,返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/08/22
    '------------------------------------------------------------------------
    Dim objItem As ListItem, strReg As String
    strReg = ""
    For Each objItem In lvwTool.ListItems
        strReg = strReg & "|" & objItem.Text & "," & objItem.SubItems(1)
    Next
    If strReg <> "" Then strReg = Mid(strReg, 2)
    
    Call SaveSetting("ZLSOFT", "公共全局\TOOLS", "TOOLFILES", strReg)
    
    SaveToolsRegInfor = True
End Function

Public Sub ShowEdit(ByVal frmMain As Object, Optional ByRef blnApply As Boolean = False)
    '-------------------------------------------------------------------------------------------------
    '功能:显示外部工具设置窗口
    '参数:frmMain-父窗口
    '出参:blnApply-应用
    '编制:刘兴宏
    '日期:2007/08/20
    '-------------------------------------------------------------------------------------------------
    On Error Resume Next
    mblnApply = False
    Me.Show 1, frmMain
    blnApply = mblnApply
End Sub
