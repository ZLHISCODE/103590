VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectMuli 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3525
   ClientLeft      =   2730
   ClientTop       =   3435
   ClientWidth     =   5580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5580
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "??"
         BeginProperty Font 
            Name            =   "????"
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
         TabStop         =   0   'False
         Top             =   45
         Width           =   225
      End
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
      Picture         =   "frmSelectMuli.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   2805
      Width           =   165
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9790
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2730
      Left            =   225
      TabIndex        =   2
      Top             =   615
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   4710
      Top             =   735
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
            Picture         =   "frmSelectMuli.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelectMuli"
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
Private mrsData As New ADODB.Recordset
Private msglTxtH As Single

Private Sub SaveFormState()
    
    '???ܣ????浱ǰѡ??????״̬
    
    Dim lngloop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    zlDatabase.SetPara "frmSelectMuli_????", Me.Width, 100, 1208
    zlDatabase.SetPara "frmSelectMuli_?߶?", Me.Height, 100, 1208
    
    For lngloop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngloop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    zlDatabase.SetPara "frmSelectMuli_?п?", strTmp, 100, 1208
    
End Sub

Private Sub RestoreFormState()
    
    '???ܣ????浱ǰѡ??????״̬
    
    Dim lngloop As Long
    Dim strTmp As String
    
    On Error Resume Next
    
    If mstrStatePath = "" Then Exit Sub
    

    Me.Width = zlDatabase.GetPara("frmSelectMuli_????", 100, 1208, Me.Width)
    Me.Height = zlDatabase.GetPara("frmSelectMuli_?߶?", 100, 1208, Me.Height)
    
    If Me.Width < 100 Then
        Me.Width = 5000
    End If
    
    If Me.Height < 100 Then
        Me.Height = 5000
    End If
    
    For lngloop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngloop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    strTmp = zlDatabase.GetPara("frmSelectMuli_?п?", 100, 1208, strTmp)
    
    
    For lngloop = 1 To lvw.ColumnHeaders.Count
        lvw.ColumnHeaders(lngloop).Width = Val(Split(strTmp, ";")(lngloop - 1))
    Next

    
    '?????Ƿ񳬹???Ļ?ߺͿ???
    Dim lngTrayH As Long
    
    lngTrayH = GetTrayHeight
    
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
    
    If Me.Top + Me.Height > (Screen.Height - lngTrayH) Then Me.Top = Me.Top - Me.Height - msglTxtH
    
End Sub

Public Function ShowSelect(ByVal frmMain As Form, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal strLvw As String, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            Optional ByVal StatePath As String, _
                            Optional strTitle As String, _
                            Optional BackColor As Long = &H80000005, _
                            Optional InitSelectKey As String = "", _
                            Optional ByVal sglTxtH As Single = 0) As Boolean
    
    '????:??ʾ??ѯѡ????
    '????:
    '????:
    
    Dim lngloop As Long
    Dim objItem As ListItem
    
    If rsData.BOF Then Exit Function
    
    Set mrsData = rsData
    
    mrsData.MoveFirst
    
    mblnStartUp = True
    mblnOK = False
    mstrStatePath = "˽??ģ??\" & gstrDBUser & "\" & App.ProductName & "\" & StatePath
    mlngSortColumn = 1
    mstrTitle = strTitle
    msglTxtH = sglTxtH
        
    lvw.BackColor = BackColor
'    zlControl.LvwFlatColumnHeader lvw
    lvw.ListItems.Clear
    zlControl.LvwSelectColumns lvw, strLvw, True
        
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
                    
    '?ָ???????С???б?????
    Call RestoreFormState
    
    'װ??????
    With lvw
        Do While Not mrsData.EOF
            
            Set objItem = .ListItems.Add(, "K" & mrsData("ID").Value, mrsData(.ColumnHeaders(1).Text), 1, 1)
            For lngloop = 1 To .ColumnHeaders.Count - 1
                objItem.SubItems(lngloop) = zlCommFun.Nvl(mrsData(.ColumnHeaders(lngloop + 1).Text))
            Next
            
            If mrsData("ѡ??").Value = 1 Then objItem.Checked = True
                        
            mrsData.MoveNext
        Loop
    End With
    
    stb.Panels(1).Text = "???????? " & lvw.ListItems.Count & " ????????"
    
    On Error Resume Next
    If InitSelectKey <> "" Then
        
        lvw.ListItems("K" & InitSelectKey).Selected = True
        lvw.ListItems("K" & InitSelectKey).EnsureVisible
        
    End If
    On Error GoTo 0
    
    If Me.Visible = False Then
        Me.Show 1, frmMain
    End If
    
'    rsData.MoveFirst
'    rsData.Move mlngIndex - 1
    Set rsData = mrsData
    
    ShowSelect = mblnOK
    
End Function

Public Function ShowSelectSP(ByVal frmMain As Form, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal strLvw As String, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            Optional ByVal StatePath As String, _
                            Optional strTitle As String, _
                            Optional BackColor As Long = &H80000005, _
                            Optional InitSelectKey As String = "", _
                            Optional ByVal sglTxtH As Single = 0) As Boolean
    
    '????:??ʾ??ѯѡ????
    '????:
    '????:
    
    Dim lngloop As Long
    Dim objItem As ListItem
    Dim strData As String
    If rsData.BOF Then Exit Function
    
    Set mrsData = rsData
    
    mrsData.MoveFirst
    
    mblnStartUp = True
    mblnOK = False
    mstrStatePath = "˽??ģ??\" & gstrDBUser & "\" & App.ProductName & "\" & StatePath
    mlngSortColumn = 1
    mstrTitle = strTitle
    msglTxtH = sglTxtH
        
    lvw.BackColor = BackColor
'    zlControl.LvwFlatColumnHeader lvw
    lvw.ListItems.Clear
    zlControl.LvwSelectColumns lvw, strLvw, True
        
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
                    
    '?ָ???????С???б?????
    Call RestoreFormState
    
    'װ??????
    With lvw
        Do While Not mrsData.EOF
            If InStr(1, strData, mrsData("ID").Value) <= 0 Then
                Set objItem = .ListItems.Add(, "K" & mrsData("ID").Value, mrsData(.ColumnHeaders(1).Text), 1, 1)
            End If
            strData = strData & mrsData("ID").Value
            For lngloop = 1 To .ColumnHeaders.Count - 1
                objItem.SubItems(lngloop) = zlCommFun.Nvl(mrsData(.ColumnHeaders(lngloop + 1).Text))
            Next
            
            If mrsData("ѡ??").Value = 1 Then objItem.Checked = True
                        
            mrsData.MoveNext
        Loop
    End With
    
    stb.Panels(1).Text = "???????? " & lvw.ListItems.Count & " ????????"
    
    On Error Resume Next
    If InitSelectKey <> "" Then
        
        lvw.ListItems("K" & InitSelectKey).Selected = True
        lvw.ListItems("K" & InitSelectKey).EnsureVisible
        
    End If
    On Error GoTo 0
    
    If Me.Visible = False Then
        Me.Show 1, frmMain
    End If
    
'    rsData.MoveFirst
'    rsData.Move mlngIndex - 1
    Set rsData = mrsData
    
    ShowSelectSP = mblnOK
    
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ReturnSelect()
    Dim lngloop As Long
    Dim strFilter As String
    
'    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    strFilter = "ID=-1"
    For lngloop = 1 To lvw.ListItems.Count
        'mrsData.Filter = ""
        
        'mrsData.Filter = "ID=" & Mid(lvw.ListItems(lngLoop).Key, 2)
        'If lvw.ListItems(lngLoop).Checked Then
        '    mrsData("ѡ??").Value = 1
        'Else
        '    mrsData("ѡ??").Value = 0
        'End If
        
        If lvw.ListItems.Count = 1 Then
            lvw.ListItems(lngloop).Checked = True
        End If
        
        If lvw.ListItems(lngloop).Checked Then strFilter = strFilter & " Or ID=" & Mid(lvw.ListItems(lngloop).Key, 2)
    Next
    
    mrsData.filter = strFilter
    
    'mlngIndex = lvw.SelectedItem.Index
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    lvw.SetFocus
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
    Call ReturnSelect
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


