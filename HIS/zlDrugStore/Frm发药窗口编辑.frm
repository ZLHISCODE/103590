VERSION 5.00
Begin VB.Form Frm发药窗口编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发药窗口编辑"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "Frm发药窗口编辑.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd保存 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   4
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   5
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   150
      TabIndex        =   6
      Top             =   0
      Width           =   3675
      Begin VB.ComboBox cboWindow 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1500
         Width           =   2085
      End
      Begin VB.CheckBox Chk专家 
         Caption         =   "专家"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox Txt编码 
         Height          =   300
         Left            =   1020
         MaxLength       =   1
         TabIndex        =   0
         Top             =   270
         Width           =   500
      End
      Begin VB.CommandButton Cmd药房 
         Caption         =   "…"
         Height          =   300
         Left            =   2820
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1050
         Width           =   285
      End
      Begin VB.TextBox Txt药房 
         Height          =   300
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox Txt名称 
         Height          =   300
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   1
         Top             =   660
         Width           =   2085
      End
      Begin VB.Label lblWindow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "叫号窗口"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Lbl编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Lbl药房 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药房"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   8
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
   End
End
Attribute VB_Name = "Frm发药窗口编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IntEditState As Integer '1:新增;2:修改
Public gStr编码 As String
Public gLng药房ID As Long
Private mstr窗口 As String




Private Sub Cmd保存_Click()
    Dim RecCheck As New ADODB.Recordset
    If CheckData = False Then Exit Sub
    
    
    On Error GoTo ErrHand
    If EditState = 1 Then
        gstrSQL = "Select Count(*) Records From 发药窗口 Where 药房ID=[1] And 名称=[2] "
        Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[保存检查]", Me.Txt药房.Tag, Me.Txt名称)
        
        With RecCheck
            If Not .EOF Then
                If !Records <> 0 Then
                    MsgBox "本药房的发药窗口[" & Txt名称 & "]已存在！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If mstr窗口 <> "" And mstr窗口 <> Txt名称.Text Then
        gstrSQL = " zl_发药窗口_业务调整 (" & Me.Txt药房.Tag & ",'" & mstr窗口 & "','" & Me.Txt名称.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新发药窗口")
    End If
    
    gcnOracle.BeginTrans
    If EditState = 1 Then
        gstrSQL = " zl_发药窗口_insert ('" & Txt编码 & "','" & Txt名称 & "',1," & Me.Txt药房.Tag & "," & Chk专家.Value & ",'" & Trim(Me.cboWindow.Text) & "')"
    Else
        gstrSQL = " zl_发药窗口_update ('" & Txt编码 & "','" & Txt名称 & "'," & Me.Txt药房.Tag & "," & Chk专家.Value & ",'" & gStr编码 & "'," & gLng药房ID & ",'" & Trim(Me.cboWindow.Text) & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-更新发药窗口")
    gcnOracle.CommitTrans
    
    If EditState = 2 Then
        Unload Me
        Exit Sub
    End If
    
    Me.Txt编码 = GetMaxCode(Txt药房.Tag)
    Me.Txt名称 = ""
    Me.Txt编码.SetFocus
    mstr窗口 = ""
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Public Property Get EditState() As Variant
    EditState = IntEditState
End Property

Public Property Let EditState(ByVal vNewValue As Variant)
    IntEditState = vNewValue
End Property

Private Function CheckData() As Boolean
    CheckData = False
    
    If Txt编码 = "" Then
        MsgBox "编码不能为空！", vbInformation, gstrSysName
        Me.Txt编码.SetFocus
        Exit Function
    End If
    If Not IsNumeric(Txt编码) Then
        MsgBox "编码应该为数字型！", vbInformation, gstrSysName
        Me.Txt编码.SetFocus
        Exit Function
    End If
    If Txt名称 = "" Then
        MsgBox "名称不能为空！", vbInformation, gstrSysName
        Me.Txt名称.SetFocus
        Exit Function
    End If
    If LenB(StrConv(Txt名称, vbFromUnicode)) > 10 Then
        MsgBox "名称长度超长（最多5个汉字或10个字符）！", vbInformation, gstrSysName
        Me.Txt名称.SetFocus
        Exit Function
    End If
    If Val(Me.Txt药房.Tag) = 0 Then
        MsgBox "请选择药房！", vbInformation, gstrSysName
        Me.Txt药房.SetFocus
        Exit Function
    End If
    
    CheckData = True
End Function

Private Sub Cmd药房_Click()
    Dim RecTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select ID,编码,名称 From 部门表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (" & _
          " Select distinct 部门ID From 部门性质说明" & _
          " Where 工作性质 Like '%药房')" & _
          " And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' order by 编码"
    Call zlDatabase.OpenRecordset(RecTmp, gstrSQL, "提取所有药房")
    
    With FrmNodeSelect
        Set .TreeRec = RecTmp.Clone
        .StrNode = "所有药房"
        .Show 1, Me
        If .BlnSuccess Then
            Me.Txt药房 = .CurrentName
            Me.Txt药房.Tag = .CurrentID
        Else
            Me.Txt药房 = ""
            Me.Txt药房.Tag = 0
        End If
        Unload FrmNodeSelect
    End With
    
    Call LoadWindow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    If EditState = 1 Then
        If frm发药窗口.Tree.SelectedItem.Key = "R" Then
            Txt药房 = ""
            Txt药房.Tag = 0
        Else
            Me.Txt药房 = Mid(frm发药窗口.Tree.SelectedItem, InStr(1, frm发药窗口.Tree.SelectedItem, "】") + 1)
            Me.Txt药房.Tag = Mid(frm发药窗口.Tree.SelectedItem.Key, 3)
            Call LoadWindow
        End If
        Txt编码 = GetMaxCode(Txt药房.Tag)
        Exit Sub
    End If
    
    Me.Txt编码 = frm发药窗口.Lvw.SelectedItem.SubItems(1)
    Me.Txt名称 = frm发药窗口.Lvw.SelectedItem
    Me.Chk专家 = IIf(frm发药窗口.Lvw.SelectedItem.SubItems(4) = "√", 1, 0)
    Me.Txt药房 = frm发药窗口.Lvw.SelectedItem.SubItems(3)
    Me.Txt药房.Tag = Mid(frm发药窗口.Lvw.SelectedItem.Key, 3, InStr(1, frm发药窗口.Lvw.SelectedItem.Key, ",") - 3)
    mstr窗口 = Txt名称.Text
    
    gStr编码 = Txt编码
    gLng药房ID = Me.Txt药房.Tag
    
    Call LoadWindow
    
End Sub

Private Sub LoadWindow()
    Dim rsTemp As Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim intRow As Integer
    
    On Error GoTo errHandle
    
    strSQL = "select 名称 from 发药窗口 where 药房id=[1] and 叫号窗口 is null and 名称<>[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取主叫窗口", Txt药房.Tag, IIf(Me.Txt名称.Text = "", " ", Me.Txt名称.Text))
    
    intRow = -1
    Me.cboWindow.Clear
    Me.cboWindow.AddItem " "
    Do While Not rsTemp.EOF
        i = i + 1
        Me.cboWindow.AddItem rsTemp!名称
        If frm发药窗口.Lvw.SelectedItem.SubItems(5) = rsTemp!名称 Then
            intRow = i
        End If
        rsTemp.MoveNext
    Loop
    
    If frm发药窗口.Lvw.SelectedItem Is Nothing Then Exit Sub
    
    If frm发药窗口.Lvw.SelectedItem.SubItems(5) <> "" Then
        If intRow >= 0 Then
            cboWindow.ListIndex = intRow
        Else
            Me.cboWindow.AddItem frm发药窗口.Lvw.SelectedItem.SubItems(5)
            cboWindow.ListIndex = i + 1
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt编码_GotFocus()
    GetFocus Txt编码
End Sub

Private Sub Txt名称_GotFocus()
    GetFocus Txt名称
End Sub

Private Sub Txt名称_KeyPress(KeyAscii As Integer)
    If InStr(1, "!@#$%^&*(){}[];:,.<>?/|\、《》，。｛｝【】；：？、￥%……&（）", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Txt药房_GotFocus()
    GetFocus Txt药房
End Sub

Private Sub Txt药房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CompareStr  As String, RecOpen As New ADODB.Recordset, StrBit As Byte
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药房) = "" Then
        Txt药房.Tag = 0
        Exit Sub
    End If
    
    CompareStr = UCase(Txt药房)
    If Mid(CompareStr, 1, 1) = "【" Then
        If InStr(2, CompareStr, "】") <> 0 Then
            CompareStr = Mid(CompareStr, 2, InStr(2, CompareStr, "】") - 2)
        Else
            CompareStr = Mid(CompareStr, 2)
        End If
    End If

    StrBit = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0")
    
    gstrSQL = " Select ID,编码,名称 From 部门表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And ID in (" & _
      " Select distinct 部门ID From 部门性质说明" & _
      " Where 工作性质 Like '%药房')" & _
      " And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01'" & _
      " And (简码 like [1] Or 编码 like [1] or 名称 like [1])"
    Set RecOpen = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取药房]", IIf(StrBit = "0", "%", "") & CompareStr & "%")
    
    With RecOpen
        If .EOF Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            Me.Txt药房 = ""
            Txt药房.Tag = 0
            KeyCode = 0
            Exit Sub
        End If
        If .RecordCount > 1 Then
            With FrmMutilSelect
                Set .gRecCommon = RecOpen.Clone
                .gStrHideCol = "000,1000,1500"
                .strCaption = "药房选择器"
                .FrmHeight = 3680
                .FrmWidth = 6000
                .Show 1, Me
                
                If .BlnSelect = False Then
                    Unload FrmMutilSelect
                    KeyCode = 0
                    Me.Txt药房 = ""
                    Txt药房.Tag = 0
                    Exit Sub
                Else
                    Me.Txt药房 = .gRecCommon!名称
                    Txt药房.Tag = .gRecCommon!Id
                    Unload FrmMutilSelect
                End If
            End With
        Else
            Me.Txt药房 = !名称
            Txt药房.Tag = !Id
        End If
        
    End With
    
    Call LoadWindow
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetMaxCode(ByVal lng药房ID As Long) As String
    Dim StrCode As String
    Dim RecCode As New ADODB.Recordset
    
    
'        If .State = 1 Then .Close
'        gstrSQL = "Select Max(编码) Code From 发药窗口 Where 药房ID=" & lng药房ID
'
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        .Open gstrSQL, gcnOracle
'        Call SQLTest
    On Error GoTo errHandle
    gstrSQL = "Select Max(编码) Code From 发药窗口 Where 药房ID=[1]"
    Set RecCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药房ID)
    
   With RecCode
        If .EOF Then
            GetMaxCode = 1
        Else
            If IsNull(!Code) Then
                GetMaxCode = 1
            Else
                GetMaxCode = !Code + 1
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
