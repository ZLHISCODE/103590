VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelectTree 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4110
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   3810
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSelectTree.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   4050
      Width           =   165
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7197
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   645
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   150
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   3
         Top             =   45
         Width           =   225
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标题"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   30
         TabIndex        =   5
         Top             =   45
         Width           =   360
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   1395
      Left            =   75
      TabIndex        =   0
      Top             =   570
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   2461
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3225
      Top             =   585
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
            Picture         =   "frmSelectTree.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelectTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mstrStatePath As String
Private mlngX As Long
Private mlngY As Long
Private mstrSvrKey As String
Private msglTxtH As Single
Private mstrPrive As String
Private mstrSvrTag As String
Private mstrTitle As String
Private mrsData As ADODB.Recordset
Private mblnOK As Boolean

Private Sub SaveFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", mstrStatePath, "宽度", Me.Width
    SaveSetting "ZLSOFT", mstrStatePath, "高度", Me.Height
    
End Sub

Private Sub RestoreFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    Me.Width = GetSetting("ZLSOFT", mstrStatePath, "宽度", Me.Width)
    Me.Height = GetSetting("ZLSOFT", mstrStatePath, "高度", Me.Height)
    
    On Error Resume Next
    
    '检查是否超过屏幕高和宽度
    
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
    
    If Me.Top + Me.Height > Screen.Height Then Me.Top = Me.Top - Me.Height - msglTxtH
End Sub

Private Sub ReadTreeData()
    '功能：
    
    Dim objItem As Node
    Dim rs As New ADODB.Recordset
        
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    
    Do While Not mrsData.EOF
        
        On Error Resume Next
        
        If IIf(IsNull(mrsData("上级id").Value), 0, mrsData("上级id").Value) <> 0 Then
            Set objItem = tvw.Nodes.Add("K" & mrsData("上级ID").Value, tvwChild, "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
        Else
            Set objItem = tvw.Nodes.Add(, , "K" & mrsData("ID").Value, mrsData("名称").Value, 1, 1)
        End If
        
        objItem.Expanded = True
        
        mrsData.MoveNext
    Loop
     
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).EnsureVisible
        tvw.Nodes(1).Expanded = True
    End If
End Sub

Public Function ShowSelect(ByVal frmMain As Form, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            ByVal sglTxtH As Single, _
                            Optional strSelectItem As String, _
                            Optional StatePath As String, _
                            Optional strTitle As String, _
                            Optional BackColor As Long = &H80000005, _
                            Optional InitSelectKey As String = "") As Boolean
    
    '功能:显示查询选择器
    '参数:
    '返回:
    
    If rsData.BOF Then Exit Function
    
    Set mrsData = rsData
        
    mstrSvrKey = ""
    mblnOK = False
    mstrSvrTag = ""
    msglTxtH = sglTxtH
    mstrTitle = strTitle
    
    mstrStatePath = "私有模块\" & gstrUserName & "\" & App.ProductName & "\" & StatePath
    
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
    lblCaption.Caption = strTitle
    
    Call RestoreFormState
    
    Call ReadTreeData
    
    If strSelectItem <> "" Then
        On Error Resume Next
        tvw.Nodes("K" & strSelectItem).Selected = True
        tvw.Nodes("K" & strSelectItem).EnsureVisible
        On Error GoTo 0
    End If
    
    If Not (tvw.SelectedItem Is Nothing) Then Call tvw_NodeClick(tvw.SelectedItem)
        
    Me.Show 1, frmMain
    
    Set rsData = mrsData
    
    ShowSelect = mblnOK
    
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ReturnSelect()
    
    If Not (tvw.SelectedItem Is Nothing) Then
        mrsData.Filter = ""
        mrsData.Filter = "ID=" & Mid(tvw.SelectedItem.Key, 2)
        mblnOK = True
    End If
    
    Unload Me
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With picTitle
        .Left = -15
        .Top = -30
        .Width = Me.ScaleWidth + 30
    End With

    
    With tvw
        .Left = -15
        .Top = picTitle.Top + picTitle.Height
        .Height = Me.ScaleHeight - stb.Height - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picDrag
        .Left = Me.ScaleWidth - .Width - 30
        .Top = Me.ScaleHeight - .Height - 30
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormState
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        mlngX = x
        mlngY = y
    End If
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Me.Width + x - mlngX < 1200 Then Exit Sub
        If Me.Height + y - mlngY < 1995 Then Exit Sub
        
        Me.Width = Me.Width + x - mlngX
        Me.Height = Me.Height + y - mlngY
        Call Form_Resize
    End If
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    
    With cmdClose
        .Left = picTitle.Width - .Width - 30
    End With
End Sub

Private Sub tvw_DblClick()
    If tvw.SelectedItem Is Nothing Then Exit Sub
    Call ReturnSelect
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    '如果重复点击同一节点，则不再刷新数据
       
    If Node.Key <> mstrSvrKey Then
        mstrSvrKey = Node.Key
            
        If tvw.SelectedItem Is Nothing Then
            stb.Panels(1).Text = "没有任何信息！"
        Else
            stb.Panels(1).Text = "共有 " & tvw.Nodes.Count & " 条信息！"
        End If
    End If
End Sub
