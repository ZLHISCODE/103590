VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelectList 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   2730
   ClientTop       =   3435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5775
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   3555
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9578
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   5115
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSelectList.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   2805
      Width           =   165
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2730
      Left            =   225
      TabIndex        =   0
      Top             =   615
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00C66300&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   0
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
         TabIndex        =   4
         Top             =   45
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3450
      Top             =   3105
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
            Picture         =   "frmSelectList.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   3165
      TabIndex        =   1
      Top             =   255
      Width           =   1100
   End
End
Attribute VB_Name = "frmSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mlngIndex As Long
Private mblnOK As Boolean
Private mstrStatePath As String
Private mlngSortColumn As Long
Private mstrTitle As String
Private mlngX As Long
Private mlngY As Long

Private Sub SaveFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", mstrStatePath, "宽度", Me.Width
    SaveSetting "ZLSOFT", mstrStatePath, "高度", Me.Height
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    SaveSetting "ZLSOFT", mstrStatePath, "列宽", strTmp
    
End Sub

Private Sub RestoreFormState()
    
    '功能：保存当前选择器的状态
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    Me.Width = GetSetting("ZLSOFT", mstrStatePath, "宽度", Me.Width)
    Me.Height = GetSetting("ZLSOFT", mstrStatePath, "高度", Me.Height)
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    strTmp = GetSetting("ZLSOFT", mstrStatePath, "列宽", strTmp)
    
    On Error Resume Next
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        lvw.ColumnHeaders(lngLoop).Width = Val(Split(strTmp, ";")(lngLoop - 1))
    Next
    
End Sub

Public Function ShowSelect(ByVal frmMain As Object, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal strLvw As String, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            Optional ByVal StatePath As String, _
                            Optional strTitle As String, _
                            Optional BackColor As Long = &H80000005) As Boolean
    
    '功能:显示查询选择器
    '参数:
    '返回:
    
    Dim lngLoop As Long
    Dim objItem As ListItem
    
    If rsData.BOF Then Exit Function
    
    rsData.MoveFirst
    
    mblnStartUp = True
    mblnOK = False
    mstrStatePath = "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & StatePath
    mlngSortColumn = 1
    mstrTitle = strTitle
        
    lvw.BackColor = BackColor
'    zlControl.LvwFlatColumnHeader lvw
    lvw.ListItems.Clear
    zlControl.LvwSelectColumns lvw, strLvw, True
        
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
                    
    '恢复窗体大小及列表宽度
    Call RestoreFormState
    
    '装载数据
    With lvw
        Do While Not rsData.EOF
            Set objItem = .ListItems.Add(, , rsData(.ColumnHeaders(1).Text), 1, 1)
            For lngLoop = 1 To .ColumnHeaders.Count - 1
                objItem.SubItems(lngLoop) = zlCommFun.Nvl(rsData(.ColumnHeaders(lngLoop + 1).Text))
            Next
            rsData.MoveNext
        Loop
    End With
    
    stb.Panels(1).Text = "共搜索到 " & lvw.ListItems.Count & " 条结果。"
    
    Me.Show 1
    
    rsData.MoveFirst
    rsData.Move mlngIndex - 1
    
    ShowSelect = mblnOK
    
End Function


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlngIndex = lvw.SelectedItem.Index
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With picTitle
        .Left = -15
        .Top = -30
        .Width = Me.ScaleWidth + 30
    End With
    
    With lvw
        .Left = -15
        .Top = picTitle.Top + picTitle.Height
        .Width = Me.ScaleWidth + 30
        .Height = Me.ScaleHeight - stb.Height - .Top
    End With
    
    With picDrag
        .Left = Me.ScaleWidth - .Width - 30
        .Top = Me.ScaleHeight - .Height - 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormState
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    zlControl.LvwSortColumn lvw, mlngSortColumn
End Sub

Private Sub lvw_DblClick()
    Call cmdOK_Click
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.Width = Me.Width + X - mlngX
        Me.Height = Me.Height + Y - mlngY
        Call Form_Resize
    End If
End Sub

Private Sub picTitle_Paint()
    zlControl.PicShowFlat picTitle, 1, mstrTitle, taLeftAlign
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    
    With cmdClose
        .Left = picTitle.Width - .Width - 30
    End With
End Sub
